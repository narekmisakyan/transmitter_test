using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using ExcelTransfer.Core.Cells;
using ExcelTransfer.Core.Data;
using ExcelTransfer.Core.Transfer;
using ExcelTransfer.Core.Workbooks;
using ExcelTransfer.Tests.Support;
using static ExcelTransfer.Tests.Support.WorkbookBuilder;

namespace ExcelTransfer.Tests;

public class TransferScenarioTests
{
    private static double? Num(IXLWorksheet ws, string a1) => ws.Cell(a1).IsEmpty() ? null : ws.Cell(a1).GetDouble();

    [Fact]
    public void ProductCode_and_Region_keys_transfer_Q1_Q2_clear_unmatched_rows_and_leave_Q3_alone()
    {
        using var dir = new TempDir();
        var source = CreateSource(dir.File("source.xlsx"));
        var destination = CreateDestination(dir.File("destination.xlsx"));
        var output = dir.File("output.xlsx");
        var sourceHash = OpenXmlFixture.Sha256(source);
        var destinationHash = OpenXmlFixture.Sha256(destination);

        var plan = TransferPlanner.Plan(Spec(source, destination, output));
        Assert.True(plan.CanTransfer, string.Join(" | ", plan.Errors.Select(e => e.Message)));
        var result = TransferExecutor.Execute(plan);

        Assert.Equal(output, result.OutputPath);
        Assert.Equal(4, result.RecordsMatched);
        Assert.Equal(10, result.CellsCleared);
        Assert.Equal(7, result.ValuesWritten);
        Assert.True(File.Exists(output));
        Assert.Equal(sourceHash, OpenXmlFixture.Sha256(source));
        Assert.Equal(destinationHash, OpenXmlFixture.Sha256(destination));

        using var wb = new XLWorkbook(output);
        var ws = wb.Worksheet("Plan");
        // Matched rows (rows in destination order; source order was shuffled).
        Assert.Equal(110, Num(ws, "D2")); Assert.Equal(111, Num(ws, "E2")); Assert.Equal(3, Num(ws, "F2"));
        Assert.Equal(115, Num(ws, "D3")); Assert.Equal(116, Num(ws, "E3")); Assert.Equal(6, Num(ws, "F3"));
        Assert.Equal(220, Num(ws, "D4")); Assert.Null(Num(ws, "E4"));      Assert.Equal(9, Num(ws, "F4"));
        Assert.Equal(330, Num(ws, "D5")); Assert.Equal(331.5, Num(ws, "E5")); Assert.Equal(12, Num(ws, "F5"));
        // Unmatched destination row: Q1/Q2 cleared, Q3 untouched.
        Assert.Null(Num(ws, "D6")); Assert.Null(Num(ws, "E6")); Assert.Equal(15, Num(ws, "F6"));
        // Keys, description, header, formulas and the note below the data are intact.
        Assert.Equal("P-004", ws.Cell("A6").GetString());
        Assert.Equal("Orphan (no source)", ws.Cell("C6").GetString());
        Assert.Equal("Q1 Actual", ws.Cell("D1").GetString());
        Assert.Equal("SUM(D2:F2)", ws.Cell("G2").FormulaA1);
        Assert.Equal("Note below the data", ws.Cell("A9").GetString());
        // Formatting and the table (with totals row) survive.
        Assert.Equal(XLColor.LightYellow.Color.ToArgb(), ws.Cell("D2").Style.Fill.BackgroundColor.Color.ToArgb());
        Assert.Equal("#,##0.00", ws.Cell("D3").Style.NumberFormat.Format);
        var table = ws.Table("Products");
        Assert.True(table.ShowTotalsRow);
        Assert.Equal("A1:G7", table.RangeAddress.ToStringRelative());
        Assert.True(ws.Cell("D7").HasFormula);
        // Other sheet untouched byte for byte; calc chain removed; recalculation requested.
        OpenXmlFixture.AssertSameEntry(destination, output, "xl/worksheets/sheet2.xml");
        OpenXmlFixture.AssertSameEntry(destination, output, "docProps/app.xml");
        Assert.False(OpenXmlFixture.HasEntry(output, "xl/calcChain.xml"));
        using var doc = SpreadsheetDocument.Open(output, false);
        Assert.True(doc.WorkbookPart!.Workbook.CalculationProperties!.FullCalculationOnLoad!.Value);
    }

    [Fact]
    public void Single_key_column_works_when_codes_are_unique()
    {
        using var dir = new TempDir();
        var source = CreateSource(dir.File("s.xlsx"), new[] { new SourceRow("P-002", "x", 20.0, 21.0), new SourceRow("P-003", "y", 30.0, 31.0) });
        var destination = CreateDestination(dir.File("d.xlsx"), new[]
        {
            new DestRow("P-003", "South", "Gizmo", 1.0, 1.0, 1.0),
            new DestRow("P-002", "North", "Gadget", 2.0, 2.0, 2.0),
        });
        Run(Spec(source, destination, dir.File("o.xlsx"), keys: CodeKeyOnly));
        using var wb = new XLWorkbook(dir.File("o.xlsx"));
        var ws = wb.Worksheet("Plan");
        Assert.Equal(30, Num(ws, "D2")); Assert.Equal(31, Num(ws, "E2"));
        Assert.Equal(20, Num(ws, "D3")); Assert.Equal(21, Num(ws, "E3"));
    }

    [Fact]
    public void Composite_key_collision_AB_C_versus_A_BC_is_kept_apart()
    {
        using var dir = new TempDir();
        var source = CreateSource(dir.File("s.xlsx"), new[] { new SourceRow("AB", "C", 1.0, 10.0), new SourceRow("A", "BC", 2.0, 20.0) });
        var destination = CreateDestination(dir.File("d.xlsx"), new[]
        {
            new DestRow("A", "BC", "a-bc", 0.0, 0.0, 0.0),
            new DestRow("AB", "C", "ab-c", 0.0, 0.0, 0.0),
        });
        var plan = TransferPlanner.Plan(Spec(source, destination, dir.File("o.xlsx")));
        Assert.True(plan.CanTransfer);
        Assert.Equal(2, plan.Matches.Count);
        TransferExecutor.Execute(plan);
        using var wb = new XLWorkbook(dir.File("o.xlsx"));
        var ws = wb.Worksheet("Plan");
        Assert.Equal(2, Num(ws, "D2")); Assert.Equal(20, Num(ws, "E2"));
        Assert.Equal(1, Num(ws, "D3")); Assert.Equal(10, Num(ws, "E3"));
    }

    [Fact]
    public void Shuffled_and_whitespace_or_case_variant_keys_still_match()
    {
        using var dir = new TempDir();
        var source = CreateSource(dir.File("s.xlsx"), new[]
        {
            new SourceRow(" p-003 ", "SOUTH", 3.0, 3.5),
            new SourceRow("P-001", "north ", 1.0, 1.5),
            new SourceRow("p-002", "North", 2.0, 2.5),
        });
        var destination = CreateDestination(dir.File("d.xlsx"), new[]
        {
            new DestRow("P-001", "North", "", 9.0, 9.0, 9.0),
            new DestRow("P-002", "North", "", 9.0, 9.0, 9.0),
            new DestRow("P-003", "South", "", 9.0, 9.0, 9.0),
        });
        Run(Spec(source, destination, dir.File("o.xlsx")));
        using var wb = new XLWorkbook(dir.File("o.xlsx"));
        var ws = wb.Worksheet("Plan");
        Assert.Equal(1, Num(ws, "D2")); Assert.Equal(2, Num(ws, "D3")); Assert.Equal(3, Num(ws, "D4"));
        Assert.Equal(1.5, Num(ws, "E2")); Assert.Equal(2.5, Num(ws, "E3")); Assert.Equal(3.5, Num(ws, "E4"));
        Assert.Equal("P-001", ws.Cell("A2").GetString()); // keys untouched
    }

    [Fact]
    public void Formula_results_are_copied_as_values_and_missing_results_block()
    {
        using var dir = new TempDir();
        var source = CreateSource(dir.File("s.xlsx"), customize: (_, ws) => ws.Cell(3, SrcQ1).FormulaA1 = "100*2");
        var destination = CreateDestination(dir.File("d.xlsx"));
        var plan = TransferPlanner.Plan(Spec(source, destination, dir.File("o.xlsx")));
        Assert.True(plan.CanTransfer);
        Assert.Equal(1, plan.FormulaResultsUsed);
        TransferExecutor.Execute(plan);
        using (var wb = new XLWorkbook(dir.File("o.xlsx")))
        {
            var cell = wb.Worksheet("Plan").Cell("D2"); // P-001/North is source row 3
            Assert.False(cell.HasFormula);
            Assert.Equal(200, cell.GetDouble());
        }

        var stale = CreateSource(dir.File("stale.xlsx"), customize: (_, ws) => ws.Cell(3, SrcQ1).FormulaA1 = "100*2");
        OpenXmlFixture.RemoveCachedValue(stale, SourceSheet, "C3");
        var blocked = TransferPlanner.Plan(Spec(stale, destination, dir.File("o2.xlsx")));
        var issue = Assert.Single(blocked.Errors.Where(e => e.Code == "formula-without-result"));
        Assert.Contains(issue.Details, d => d.Contains("Row 3") && d.Contains("100*2"));
    }

    [Fact]
    public void Destination_formulas_in_mapped_columns_are_replaced_with_a_warning()
    {
        using var dir = new TempDir();
        var source = CreateSource(dir.File("s.xlsx"));
        var destination = CreateDestination(dir.File("d.xlsx"), customize: (_, ws) => ws.Cell(2, DstQ1).FormulaA1 = "F2*2");
        var plan = TransferPlanner.Plan(Spec(source, destination, dir.File("o.xlsx")));
        Assert.True(plan.CanTransfer);
        Assert.Contains(plan.Warnings, w => w.Code == "formulas-replaced");
        Assert.Equal(1, plan.FormulaCellsToClear);
        TransferExecutor.Execute(plan);
        using var wb = new XLWorkbook(dir.File("o.xlsx"));
        var cell = wb.Worksheet("Plan").Cell("D2");
        Assert.False(cell.HasFormula);
        Assert.Equal(110, cell.GetDouble());
    }

    [Fact]
    public void Unrelated_parts_pictures_comments_validations_and_second_sheet_are_byte_identical()
    {
        using var dir = new TempDir();
        var source = CreateSource(dir.File("s.xlsx"));
        var destination = CreateDestination(dir.File("d.xlsx"), customize: (wb, ws) =>
        {
            ws.Cell("C2").CreateComment().AddText("Check this value");
            ws.Range("D2:D6").CreateDataValidation().WholeNumber.Between(0, 100000);
            ws.Range("D2:D6").AddConditionalFormat().WhenGreaterThan(100).Fill.SetBackgroundColor(XLColor.LightGreen);
            using var png = new MemoryStream(OpenXmlFixture.TinyPng);
            ws.AddPicture(png).MoveTo(ws.Cell("J2"));
            wb.Worksheet("Notes").Cell("A5").Value = "unchanged";
        });
        Run(Spec(source, destination, dir.File("o.xlsx")));
        var output = dir.File("o.xlsx");
        var names = OpenXmlFixture.EntryNames(destination);
        foreach (var name in names.Where(n => n.Contains("/media/") || n.Contains("/drawings/") || n.Contains("comments") || n.Contains("sheet2.xml") || n.StartsWith("docProps") || n.Contains("theme")))
            OpenXmlFixture.AssertSameEntry(destination, output, name);
        using var wb2 = new XLWorkbook(output);
        var ws2 = wb2.Worksheet("Plan");
        Assert.True(ws2.Cell("C2").HasComment);
        Assert.Single(ws2.Pictures);
        Assert.NotEmpty(ws2.ConditionalFormats);
        Assert.Equal(110, Num(ws2, "D2"));
    }

    [Fact]
    public void Hidden_rows_are_included_by_default_and_can_be_excluded()
    {
        using var dir = new TempDir();
        var source = CreateSource(dir.File("s.xlsx"));
        var destination = CreateDestination(dir.File("d.xlsx"), customize: (_, ws) => ws.Row(4).Hide()); // P-002/North
        Run(Spec(source, destination, dir.File("included.xlsx")));
        using (var wb = new XLWorkbook(dir.File("included.xlsx")))
        {
            var ws = wb.Worksheet("Plan");
            Assert.Equal(220, Num(ws, "D4"));
            Assert.True(ws.Row(4).IsHidden);
        }
        var plan = TransferPlanner.Plan(Spec(source, destination, dir.File("excluded.xlsx"), destinationSelection: DestinationTableSelection(includeHidden: false)));
        Assert.True(plan.CanTransfer);
        Assert.Equal(1, plan.DestinationHiddenRowsSkipped);
        Assert.Contains(plan.UnmatchedSource, u => u.Key.Display() == "P-002 | North");
        TransferExecutor.Execute(plan);
        using (var wb = new XLWorkbook(dir.File("excluded.xlsx")))
        {
            var ws = wb.Worksheet("Plan");
            Assert.Equal(7, Num(ws, "D4")); // untouched
            Assert.Equal(8, Num(ws, "E4"));
        }
    }

    [Fact]
    public void Dates_get_a_date_format_when_the_destination_cell_has_none()
    {
        using var dir = new TempDir();
        var source = CreateSource(dir.File("s.xlsx"), new[] { new SourceRow("P-001", "North", new DateTime(2026, 3, 31), new DateTime(2026, 6, 30, 8, 30, 0)) });
        var destination = CreateDestination(dir.File("d.xlsx"), new[] { new DestRow("P-001", "North", "d", 1.0, 2.0, 3.0) });
        var plan = TransferPlanner.Plan(Spec(source, destination, dir.File("o.xlsx")));
        Assert.True(plan.CanTransfer);
        Assert.Equal(2, plan.DateValuesNeedingDateFormat);
        var result = TransferExecutor.Execute(plan);
        Assert.Equal(2, result.DateStylesApplied);
        using var wb = new XLWorkbook(dir.File("o.xlsx"));
        var ws = wb.Worksheet("Plan");
        Assert.Equal(XLDataType.DateTime, ws.Cell("D2").DataType);
        Assert.Equal(new DateTime(2026, 3, 31), ws.Cell("D2").GetDateTime());
        Assert.Equal(new DateTime(2026, 6, 30, 8, 30, 0), ws.Cell("E2").GetDateTime());
        Assert.Equal(3, Num(ws, "F2"));
    }

    [Fact]
    public void Dates_from_a_1904_workbook_keep_their_calendar_value()
    {
        using var dir = new TempDir();
        var source = CreateSource(dir.File("s.xlsx"), new[] { new SourceRow("P-001", "North", new DateTime(2026, 3, 31), 5.0) }, customize: (wb, _) => wb.Use1904DateSystem = true);
        var destination = CreateDestination(dir.File("d.xlsx"), new[] { new DestRow("P-001", "North", "d", 1.0, 2.0, 3.0) });
        Run(Spec(source, destination, dir.File("o.xlsx")));
        using var wb = new XLWorkbook(dir.File("o.xlsx"));
        Assert.Equal(new DateTime(2026, 3, 31), wb.Worksheet("Plan").Cell("D2").GetDateTime());
    }

    [Fact]
    public void Ordinary_range_destination_needs_an_explicit_last_row_when_notes_follow_the_data()
    {
        using var dir = new TempDir();
        var source = CreateSource(dir.File("s.xlsx"));
        var destination = CreateDestination(dir.File("d.xlsx"), asTable: false);
        var guessed = TransferPlanner.Plan(Spec(source, destination, dir.File("o.xlsx"), destinationSelection: DestinationRangeSelection()));
        Assert.Contains(guessed.Errors, e => e.Code == "missing-keys"); // the "Note below the data" row has no Region
        var explicitRows = TransferPlanner.Plan(Spec(source, destination, dir.File("o.xlsx"), destinationSelection: DestinationRangeSelection(lastDataRow: 6)));
        Assert.True(explicitRows.CanTransfer, string.Join("|", explicitRows.Errors.Select(e => e.Message)));
        TransferExecutor.Execute(explicitRows);
        using var wb = new XLWorkbook(dir.File("o.xlsx"));
        var ws = wb.Worksheet("Plan");
        Assert.Equal(110, Num(ws, "D2"));
        Assert.Null(Num(ws, "D6"));
        Assert.Equal("Note below the data", ws.Cell("A9").GetString());
    }

    [Fact]
    public void Protected_sheets_merged_cells_partial_shared_formulas_and_calculated_columns_block()
    {
        using var dir = new TempDir();
        var source = CreateSource(dir.File("s.xlsx"));

        var protectedDest = CreateDestination(dir.File("protected.xlsx"), customize: (_, ws) => ws.Protect());
        Assert.Contains(TransferPlanner.Plan(Spec(source, protectedDest, dir.File("o1.xlsx"))).Errors, e => e.Code == "sheet-protected");

        var mergedDest = CreateDestination(dir.File("merged.xlsx"), customize: (_, ws) => ws.Range("D3:E3").Merge());
        var mergedIssue = Assert.Single(TransferPlanner.Plan(Spec(source, mergedDest, dir.File("o2.xlsx"))).Errors.Where(e => e.Code == "merged-cells"));
        Assert.Contains("D3:E3", mergedIssue.Details);

        var sharedDest = CreateDestination(dir.File("shared.xlsx"));
        OpenXmlFixture.MakeSharedFormulaGroup(sharedDest, DestinationSheet, "D2", "D2:E3", "F2*2", "E2", "D3", "E3");
        var partial = TransferPlanner.Plan(Spec(source, sharedDest, dir.File("o3.xlsx"), periods: new[] { new ColumnMapping(SrcQ1, DstQ1) }));
        Assert.Contains(partial.Errors, e => e.Code == "shared-formula-partial");
        var whole = TransferPlanner.Plan(Spec(source, sharedDest, dir.File("o3.xlsx")));
        Assert.DoesNotContain(whole.Errors, e => e.Code == "shared-formula-partial");
        Assert.True(whole.CanTransfer);
        TransferExecutor.Execute(whole);
        using (var wb = new XLWorkbook(dir.File("o3.xlsx")))
        {
            var ws = wb.Worksheet("Plan");
            Assert.False(ws.Cell("D2").HasFormula);
            Assert.False(ws.Cell("E3").HasFormula);
            Assert.Equal(110, Num(ws, "D2"));
        }

        var calcDest = CreateDestination(dir.File("calc.xlsx"));
        OpenXmlFixture.SetCalculatedColumn(calcDest, DestinationSheet, "Q1 Actual", "[@Description]");
        Assert.Contains(TransferPlanner.Plan(Spec(source, calcDest, dir.File("o4.xlsx"))).Errors, e => e.Code == "calculated-column");
    }

    [Fact]
    public void Manual_calculation_mode_in_the_source_is_flagged()
    {
        using var dir = new TempDir();
        var source = CreateSource(dir.File("s.xlsx"), customize: (wb, _) => wb.CalculateMode = XLCalculateMode.Manual);
        var destination = CreateDestination(dir.File("d.xlsx"));
        var plan = TransferPlanner.Plan(Spec(source, destination, dir.File("o.xlsx")));
        Assert.Contains(plan.Warnings, w => w.Code == "manual-calculation");
    }

    [Fact]
    public void Xlsm_destination_keeps_its_vba_project_and_output_must_stay_xlsm()
    {
        using var dir = new TempDir();
        var source = CreateSource(dir.File("s.xlsx"));
        var xlsx = CreateDestination(dir.File("d.xlsx"));
        var vba = OpenXmlFixture.ReadZipEntry(Fixtures.PathOf("xlsm/poi-SimpleMacro.xlsm"), "xl/vbaProject.bin");
        var destination = OpenXmlFixture.ConvertToXlsm(xlsx, dir.File("d.xlsm"), vba);
        Assert.Equal(WorkbookFormat.Xlsm, WorkbookFileCheck.Check(destination).Format);

        var wrongExtension = TransferPlanner.Plan(Spec(source, destination, dir.File("o.xlsx")));
        Assert.Contains(wrongExtension.Errors, e => e.Code == "output-extension");

        var result = Run(Spec(source, destination, dir.File("o.xlsm")));
        Assert.Equal(WorkbookFormat.Xlsm, result.Format);
        OpenXmlFixture.AssertSameEntry(destination, dir.File("o.xlsm"), "xl/vbaProject.bin");
        Assert.Equal(WorkbookFormat.Xlsm, WorkbookFileCheck.Check(dir.File("o.xlsm")).Format);
        using var wb = new XLWorkbook(dir.File("o.xlsm"));
        Assert.Equal(110, Num(wb.Worksheet("Plan"), "D2"));
    }

    [Fact]
    public void Real_xlsm_fixture_can_be_used_as_destination()
    {
        using var dir = new TempDir();
        var destination = Fixtures.CopyTo(dir, "xlsm/poi-SimpleMacro.xlsm");
        var summary = WorkbookInspector.Inspect(destination);
        var sheet = summary.Worksheets.First();
        var data = SheetReader.Load(destination, new DataSetSelection { SheetName = sheet.Name });
        Assert.NotNull(data);
        Assert.True(OpenXmlFixture.HasEntry(destination, "xl/vbaProject.bin"));
    }

    [Fact]
    public void Failed_verification_leaves_no_output_behind()
    {
        using var dir = new TempDir();
        var source = CreateSource(dir.File("s.xlsx"));
        var destination = CreateDestination(dir.File("d.xlsx"));
        var plan = TransferPlanner.Plan(Spec(source, destination, dir.File("o.xlsx")));
        // Sabotage: an impossible write far outside the data range is caught by verification.
        var sabotaged = new TransferPlan
        {
            Spec = plan.Spec, Issues = plan.Issues, Source = plan.Source, Destination = plan.Destination, DestinationFormat = plan.DestinationFormat,
            Matches = plan.Matches, UnmatchedSource = plan.UnmatchedSource, UnmatchedDestination = plan.UnmatchedDestination,
            ColumnsToClear = plan.ColumnsToClear, DestinationRowsInScope = plan.DestinationRowsInScope,
            Writes = plan.Writes.Concat(new[] { new CellWrite(2, DstDescription, CellValue.OfText("tampered")) }).ToList(),
        };
        var ex = Assert.Throws<TransferException>(() => TransferExecutor.Execute(sabotaged));
        Assert.Contains("Verification", ex.Message);
        Assert.False(File.Exists(dir.File("o.xlsx")));
        Assert.Empty(Directory.GetFiles(dir.Path, "~*"));
    }
}
