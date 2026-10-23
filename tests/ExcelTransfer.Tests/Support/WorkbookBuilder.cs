using ClosedXML.Excel;
using ExcelTransfer.Core.Data;
using ExcelTransfer.Core.Transfer;

namespace ExcelTransfer.Tests.Support;

public sealed record SourceRow(string? Code, string? Region, object? Q1, object? Q2, object? Q3 = null);

public sealed record DestRow(string? Code, string? Region, string Description, object? Q1, object? Q2, object? Q3);

/// <summary>Builds test workbooks with ClosedXML (used only by the tests; the product reads and writes files with its own engines).</summary>
public static class WorkbookBuilder
{
    public const string SourceSheet = "Data 2026";
    public const string DestinationSheet = "Plan";
    public const string DestinationTable = "Products";

    // Source columns: A ProductCode, B Region, C 2026 Q1, D 2026 Q2, E 2026 Q3
    public const int SrcCode = 1, SrcRegion = 2, SrcQ1 = 3, SrcQ2 = 4, SrcQ3 = 5;
    // Destination columns: A ProductCode, B Region, C Description, D Q1 Actual, E Q2 Actual, F Q3 Actual, G Total
    public const int DstCode = 1, DstRegion = 2, DstDescription = 3, DstQ1 = 4, DstQ2 = 5, DstQ3 = 6, DstTotal = 7;

    public static readonly SourceRow[] DefaultSource =
    {
        new("P-003", "South", 330.0, 331.5),
        new("P-001", "North", 110.0, 111.0),
        new("P-002", "North", 220.0, null),
        new("P-001", "South", 115.0, 116.0),
        new("P-999", "West", 1.0, 2.0),   // no destination match
    };

    public static readonly DestRow[] DefaultDestination =
    {
        new("P-001", "North", "Widget north", 1.0, 2.0, 3.0),
        new("P-001", "South", "Widget south", 4.0, 5.0, 6.0),
        new("P-002", "North", "Gadget", 7.0, 8.0, 9.0),
        new("P-003", "South", "Gizmo", 10.0, 11.0, 12.0),
        new("P-004", "East", "Orphan (no source)", 13.0, 14.0, 15.0),
    };

    public static string CreateSource(string path, IEnumerable<SourceRow>? rows = null, Action<XLWorkbook, IXLWorksheet>? customize = null, bool evaluateFormulas = true)
    {
        rows ??= DefaultSource;
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(SourceSheet);
        ws.Cell(1, SrcCode).Value = "ProductCode";
        ws.Cell(1, SrcRegion).Value = "Region";
        ws.Cell(1, SrcQ1).Value = "2026 Q1";
        ws.Cell(1, SrcQ2).Value = "2026 Q2";
        ws.Cell(1, SrcQ3).Value = "2026 Q3";
        ws.Row(1).Style.Font.Bold = true;
        var r = 2;
        foreach (var row in rows)
        {
            SetCell(ws.Cell(r, SrcCode), row.Code);
            SetCell(ws.Cell(r, SrcRegion), row.Region);
            SetCell(ws.Cell(r, SrcQ1), row.Q1);
            SetCell(ws.Cell(r, SrcQ2), row.Q2);
            SetCell(ws.Cell(r, SrcQ3), row.Q3);
            r++;
        }
        customize?.Invoke(wb, ws);
        wb.SaveAs(path, new SaveOptions { EvaluateFormulasBeforeSaving = evaluateFormulas });
        return path;
    }

    public static string CreateDestination(string path, IEnumerable<DestRow>? rows = null, bool asTable = true, bool totalsRow = true, Action<XLWorkbook, IXLWorksheet>? customize = null, bool evaluateFormulas = true)
    {
        rows ??= DefaultDestination;
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(DestinationSheet);
        ws.Cell(1, DstCode).Value = "ProductCode";
        ws.Cell(1, DstRegion).Value = "Region";
        ws.Cell(1, DstDescription).Value = "Description";
        ws.Cell(1, DstQ1).Value = "Q1 Actual";
        ws.Cell(1, DstQ2).Value = "Q2 Actual";
        ws.Cell(1, DstQ3).Value = "Q3 Actual";
        ws.Cell(1, DstTotal).Value = "Total";
        var r = 2;
        foreach (var row in rows)
        {
            SetCell(ws.Cell(r, DstCode), row.Code);
            SetCell(ws.Cell(r, DstRegion), row.Region);
            SetCell(ws.Cell(r, DstDescription), row.Description);
            SetCell(ws.Cell(r, DstQ1), row.Q1);
            SetCell(ws.Cell(r, DstQ2), row.Q2);
            SetCell(ws.Cell(r, DstQ3), row.Q3);
            ws.Cell(r, DstTotal).FormulaA1 = $"SUM(D{r}:F{r})";
            r++;
        }
        var lastRow = r - 1;
        ws.Range(2, DstQ1, lastRow, DstTotal).Style.NumberFormat.Format = "#,##0.00";
        ws.Range(1, 1, 1, DstTotal).Style.Font.Bold = true;
        ws.Cell(2, DstQ1).Style.Fill.BackgroundColor = XLColor.LightYellow;
        if (asTable)
        {
            var table = ws.Range(1, 1, lastRow, DstTotal).CreateTable(DestinationTable);
            table.Theme = XLTableTheme.TableStyleMedium2;
            if (totalsRow)
            {
                table.ShowTotalsRow = true;
                table.Field("Q1 Actual").TotalsRowFunction = XLTotalsRowFunction.Sum;
                table.Field("Q2 Actual").TotalsRowFunction = XLTotalsRowFunction.Sum;
                table.Field("Q3 Actual").TotalsRowFunction = XLTotalsRowFunction.Sum;
            }
        }
        ws.Cell(lastRow + 3, 1).Value = "Note below the data";

        var notes = wb.AddWorksheet("Notes");
        notes.Cell("A1").Value = "This sheet must stay untouched.";
        notes.Cell("A2").FormulaA1 = "Plan!D2*10";
        notes.Cell("B1").Value = new DateTime(2026, 1, 31);
        notes.Range("C1:D1").Merge();
        notes.Cell("C1").Value = "merged title";

        customize?.Invoke(wb, ws);
        wb.SaveAs(path, new SaveOptions { EvaluateFormulasBeforeSaving = evaluateFormulas });
        return path;
    }

    public static void SetCell(IXLCell cell, object? value)
    {
        switch (value)
        {
            case null: break;
            case string s: cell.Value = s; break;
            case double d: cell.Value = d; break;
            case int i: cell.Value = i; break;
            case bool b: cell.Value = b; break;
            case DateTime dt: cell.Value = dt; break;
            default: cell.Value = value.ToString(); break;
        }
    }

    public static DataSetSelection SourceSelection(string sheet = SourceSheet, int headerRow = 1, int? lastDataRow = null, bool includeHidden = true) =>
        new() { SheetName = sheet, HeaderRow = headerRow, LastDataRow = lastDataRow, IncludeHiddenRows = includeHidden };

    public static DataSetSelection DestinationTableSelection(bool includeHidden = true) =>
        new() { SheetName = DestinationSheet, TableName = DestinationTable, IncludeHiddenRows = includeHidden };

    public static DataSetSelection DestinationRangeSelection(int headerRow = 1, int? lastDataRow = null, bool includeHidden = true) =>
        new() { SheetName = DestinationSheet, HeaderRow = headerRow, LastDataRow = lastDataRow, IncludeHiddenRows = includeHidden };

    public static IReadOnlyList<ColumnMapping> CompositeKeys => new[] { new ColumnMapping(SrcCode, DstCode), new ColumnMapping(SrcRegion, DstRegion) };
    public static IReadOnlyList<ColumnMapping> CodeKeyOnly => new[] { new ColumnMapping(SrcCode, DstCode) };
    public static IReadOnlyList<ColumnMapping> Q1Q2Periods => new[] { new ColumnMapping(SrcQ1, DstQ1), new ColumnMapping(SrcQ2, DstQ2) };

    public static TransferSpec Spec(string source, string destination, string output,
        DataSetSelection? sourceSelection = null, DataSetSelection? destinationSelection = null,
        IReadOnlyList<ColumnMapping>? keys = null, IReadOnlyList<ColumnMapping>? periods = null, TransferOptions? options = null) =>
        new()
        {
            SourcePath = source,
            Source = sourceSelection ?? SourceSelection(),
            DestinationPath = destination,
            Destination = destinationSelection ?? DestinationTableSelection(),
            OutputPath = output,
            KeyMappings = keys ?? CompositeKeys,
            PeriodMappings = periods ?? Q1Q2Periods,
            Options = options ?? new TransferOptions(),
        };

    public static TransferResult Run(TransferSpec spec)
    {
        var plan = TransferPlanner.Plan(spec);
        if (!plan.CanTransfer)
            throw new InvalidOperationException("Plan has errors: " + string.Join(" | ", plan.Errors.Select(e => e.Message + " " + string.Join("; ", e.Details))));
        return TransferExecutor.Execute(plan);
    }
}
