using ClosedXML.Excel;
using ExcelTransfer.Core.Data;
using ExcelTransfer.Core.Transfer;

// Builds the example workbooks in /examples and runs the example transfer on them.
// Usage: dotnet run --project tools/ExampleGenerator [-- <output folder>]
var outputDir = args.Length > 0 ? Path.GetFullPath(args[0]) : FindExamplesFolder();
Directory.CreateDirectory(outputDir);
var sourcePath = Path.Combine(outputDir, "source.xlsx");
var destinationPath = Path.Combine(outputDir, "destination.xlsx");
var resultPath = Path.Combine(outputDir, "destination - updated.xlsx");

BuildSource(sourcePath);
BuildDestination(destinationPath);
Console.WriteLine($"Wrote {sourcePath}");
Console.WriteLine($"Wrote {destinationPath}");

var spec = new TransferSpec
{
    SourcePath = sourcePath,
    Source = new DataSetSelection { SheetName = "Data 2026", HeaderRow = 1 },
    DestinationPath = destinationPath,
    Destination = new DataSetSelection { SheetName = "Plan", TableName = "Products" },
    OutputPath = resultPath,
    KeyMappings = new[] { new ColumnMapping(1, 1), new ColumnMapping(2, 2) },          // ProductCode -> ProductCode, Region -> Region
    PeriodMappings = new[] { new ColumnMapping(3, 4), new ColumnMapping(4, 5) },       // 2026 Q1 -> Q1 Actual, 2026 Q2 -> Q2 Actual
    Options = new TransferOptions { OverwriteExistingOutput = true },
};

var plan = TransferPlanner.Plan(spec);
Console.WriteLine();
Console.WriteLine("Validation:");
foreach (var issue in plan.Issues)
{
    Console.WriteLine($"  [{issue.Severity}] {issue.Message}");
    foreach (var d in issue.Details.Take(10)) Console.WriteLine($"      - {d}");
}
Console.WriteLine($"  matched={plan.Matches.Count} unmatchedSource={plan.UnmatchedSource.Count} unmatchedDestination={plan.UnmatchedDestination.Count}");
Console.WriteLine($"  columns to clear: {string.Join(", ", plan.ColumnsToClear.Select(c => c.Column.Label))} ({plan.TotalCellsToClear} cells, {plan.NonBlankCellsToClear} currently non-blank)");
Console.WriteLine($"  values to write: {plan.ValuesToWrite}");
if (!plan.CanTransfer)
{
    Console.Error.WriteLine("The example plan has blocking errors; aborting.");
    return 1;
}

var result = TransferExecutor.Execute(plan);
Console.WriteLine();
Console.WriteLine($"Wrote {result.OutputPath}: matched {result.RecordsMatched}, cleared {result.CellsCleared} cells, wrote {result.ValuesWritten} values in {result.Duration.TotalMilliseconds:F0} ms");
Console.WriteLine();
PrintTable("Destination before", destinationPath);
PrintTable("Destination after", resultPath);
return 0;

static string FindExamplesFolder()
{
    var dir = new DirectoryInfo(Directory.GetCurrentDirectory());
    while (dir is not null && !File.Exists(Path.Combine(dir.FullName, "ExcelTransfer.slnx"))) dir = dir.Parent;
    if (dir is null) throw new InvalidOperationException("Run from inside the repository or pass an output folder.");
    return Path.Combine(dir.FullName, "examples");
}

static void BuildSource(string path)
{
    using var wb = new XLWorkbook();
    var ws = wb.AddWorksheet("Data 2026");
    string[] headers = { "ProductCode", "Region", "2026 Q1", "2026 Q2", "Comment" };
    for (var c = 0; c < headers.Length; c++) ws.Cell(1, c + 1).Value = headers[c];
    ws.Row(1).Style.Font.Bold = true;
    object?[][] rows =
    {
        new object?[] { "P-003", "South", 330.0, 331.5, "shuffled order on purpose" },
        new object?[] { "P-001", "North", 110.0, 111.0, null },
        new object?[] { "P-002", "North", 220.0, null, "Q2 not reported yet" },
        new object?[] { "P-001", "South", 115.0, 116.0, null },
        new object?[] { "P-005", "East", 55.0, 56.0, "no matching destination row -> reported and skipped" },
        new object?[] { "p-002 ", "south", "=200+25", 226.0, "key differs only by case/spaces; Q1 is a formula" },
    };
    for (var r = 0; r < rows.Length; r++)
    {
        for (var c = 0; c < rows[r].Length; c++)
        {
            var cell = ws.Cell(r + 2, c + 1);
            switch (rows[r][c])
            {
                case null: break;
                case string s when s.StartsWith('='): cell.FormulaA1 = s[1..]; break;
                case string s: cell.Value = s; break;
                case double d: cell.Value = d; break;
            }
        }
    }
    ws.Columns().AdjustToContents();
    wb.SaveAs(path, new SaveOptions { EvaluateFormulasBeforeSaving = true });
}

static void BuildDestination(string path)
{
    using var wb = new XLWorkbook();
    var ws = wb.AddWorksheet("Plan");
    ws.Cell("A1").Value = "Product plan 2026";
    ws.Cell("A1").Style.Font.Bold = true;
    ws.Cell("A1").Style.Font.FontSize = 14;
    string[] headers = { "ProductCode", "Region", "Description", "Q1 Actual", "Q2 Actual", "Q3 Actual", "Total" };
    for (var c = 0; c < headers.Length; c++) ws.Cell(3, c + 1).Value = headers[c];
    object?[][] rows =
    {
        new object?[] { "P-001", "North", "Widget (north)", 1.0, 2.0, 3.0 },
        new object?[] { "P-001", "South", "Widget (south)", 4.0, 5.0, 6.0 },
        new object?[] { "P-002", "North", "Gadget", 7.0, 8.0, 9.0 },
        new object?[] { "P-002", "South", "Gadget (south)", 7.5, 8.5, 9.5 },
        new object?[] { "P-003", "South", "Gizmo", 10.0, 11.0, 12.0 },
        new object?[] { "P-004", "East", "Orphan: no source row, Q1/Q2 will be cleared", 13.0, 14.0, 15.0 },
    };
    for (var r = 0; r < rows.Length; r++)
    {
        var row = r + 4;
        for (var c = 0; c < rows[r].Length; c++)
        {
            var cell = ws.Cell(row, c + 1);
            switch (rows[r][c])
            {
                case string s: cell.Value = s; break;
                case double d: cell.Value = d; break;
            }
        }
        ws.Cell(row, 7).FormulaA1 = $"SUM(D{row}:F{row})";
    }
    var lastRow = 3 + rows.Length;
    ws.Range(4, 4, lastRow, 7).Style.NumberFormat.Format = "#,##0.00";
    ws.Range(4, 6, lastRow, 6).Style.Fill.BackgroundColor = XLColor.LightCyan; // Q3: unselected, must stay intact
    var table = ws.Range(3, 1, lastRow, 7).CreateTable("Products");
    table.Theme = XLTableTheme.TableStyleMedium2;
    table.ShowTotalsRow = true;
    table.Field("Q1 Actual").TotalsRowFunction = XLTotalsRowFunction.Sum;
    table.Field("Q2 Actual").TotalsRowFunction = XLTotalsRowFunction.Sum;
    table.Field("Q3 Actual").TotalsRowFunction = XLTotalsRowFunction.Sum;
    table.Field("Total").TotalsRowFunction = XLTotalsRowFunction.Sum;
    ws.Cell(lastRow + 3, 1).Value = "Cells outside the table, like this note, are never touched.";
    ws.Columns().AdjustToContents();

    var notes = wb.AddWorksheet("Notes");
    notes.Cell("A1").Value = "This second sheet is preserved byte for byte.";
    notes.Cell("A3").Value = "Q1 total from Plan:";
    notes.Cell("B3").FormulaA1 = "SUM(Products[Q1 Actual])";
    notes.Columns().AdjustToContents();
    wb.SaveAs(path, new SaveOptions { EvaluateFormulasBeforeSaving = true });
}

static void PrintTable(string title, string path)
{
    Console.WriteLine(title + ":");
    using var wb = new XLWorkbook(path);
    var ws = wb.Worksheet("Plan");
    Console.WriteLine("  Row | ProductCode | Region | Q1 Actual | Q2 Actual | Q3 Actual");
    for (var r = 4; r <= 9; r++)
    {
        string V(string col) => ws.Cell(col + r).IsEmpty() ? "(blank)" : ws.Cell(col + r).GetFormattedString();
        Console.WriteLine($"  {r,3} | {V("A"),-11} | {V("B"),-6} | {V("D"),9} | {V("E"),9} | {V("F"),9}");
    }
}
