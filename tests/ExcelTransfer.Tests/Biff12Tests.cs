using System.Globalization;
using ExcelTransfer.Core.Cells;
using ExcelTransfer.Core.Data;
using ExcelTransfer.Core.Transfer;
using ExcelTransfer.Core.Workbooks;
using ExcelTransfer.Core.Workbooks.Biff12;
using ExcelTransfer.Tests.Support;
using static ExcelTransfer.Tests.Support.WorkbookBuilder;

namespace ExcelTransfer.Tests;

public class Biff12Tests
{
    public static IEnumerable<object[]> FixtureNames() => Fixtures.XlsbFiles.Select(p => new object[] { Path.GetFileName(p) });

    [Theory]
    [MemberData(nameof(FixtureNames))]
    public void Reader_matches_independent_reference_values(string fileName)
    {
        var path = Fixtures.PathOf("xlsb/" + fileName);
        var expectedPath = Fixtures.PathOf("xlsb/expected/" + Path.GetFileNameWithoutExtension(fileName) + ".cells.txt");
        var expected = File.ReadAllLines(expectedPath).Where(l => l.Length > 0 && !l.StartsWith('#')).Select(l => l.Split('\t')).ToDictionary(p => p[0], p => (Kind: p[1], Value: p[2]));

        using var source = WorkbookSources.Open(path);
        var actual = new Dictionary<string, (string Kind, string Value)>();
        foreach (var sheet in source.Sheets.Where(s => !s.IsChartSheet))
        {
            var snap = source.LoadSheet(sheet.Name);
            foreach (var row in snap.Rows)
                for (var i = 0; i < row.Columns.Length; i++)
                {
                    var v = row.Values[i];
                    if (v.IsBlank) continue;
                    var number = v.Kind == CellKind.Date && source.Date1904 ? v.Number - ExcelDates.Offset1904 : v.Number;
                    var (kind, value) = v.Kind switch
                    {
                        CellKind.Number or CellKind.Date => ("n", number.ToString("R", CultureInfo.InvariantCulture)),
                        CellKind.Boolean => ("b", v.BooleanValue ? "TRUE" : "FALSE"),
                        CellKind.Error => ("e", v.TextOrEmpty),
                        _ => ("s", v.TextOrEmpty),
                    };
                    actual[sheet.Name + "!" + CellRef.ToA1(row.Row, row.Columns[i])] = (kind, value);
                }
        }
        // The pre-release 2007 Beta file confuses the reference parser's sheet-name handling; compare by address only there.
        if (fileName.Contains("Simple"))
        {
            expected = expected.ToDictionary(kv => kv.Key[(kv.Key.LastIndexOf('!') + 1)..], kv => kv.Value);
            actual = actual.ToDictionary(kv => kv.Key[(kv.Key.LastIndexOf('!') + 1)..], kv => kv.Value);
        }
        Assert.Equal(expected.Count, actual.Count);
        foreach (var (key, exp) in expected)
        {
            Assert.True(actual.TryGetValue(key, out var act), $"missing {key}");
            Assert.Equal(exp.Kind, act.Kind);
            if (exp.Kind == "n")
                Assert.Equal(double.Parse(exp.Value, CultureInfo.InvariantCulture), double.Parse(act.Value, CultureInfo.InvariantCulture), 6);
            else
                Assert.Equal(exp.Value, act.Value);
        }
    }

    [Theory]
    [MemberData(nameof(FixtureNames))]
    public void Record_streams_round_trip_byte_for_byte(string fileName)
    {
        using var package = OpcZipPackage.Open(Fixtures.PathOf("xlsb/" + fileName));
        var binParts = package.PartNames.Where(n => n.EndsWith(".bin", StringComparison.OrdinalIgnoreCase) && !n.Contains("vbaProject") && !n.Contains("printerSettings")).ToList();
        Assert.NotEmpty(binParts);
        foreach (var part in binParts)
        {
            var data = package.ReadPart(part);
            var records = Biff12Stream.Parse(data, part);
            using var ms = new MemoryStream(data.Length);
            foreach (var rec in records) Biff12Stream.WriteRecord(ms, rec.Type, rec.Payload(data));
            Assert.True(ms.ToArray().AsSpan().SequenceEqual(data), $"{fileName}:{part} did not round-trip");
        }
    }

    [Fact]
    public void Reader_exposes_sheet_states_dates_and_formula_flags()
    {
        using var issues = WorkbookSources.Open(Fixtures.PathOf("xlsb/calamine-issues.xlsb"));
        var datatypes = issues.LoadSheet("datatypes");
        var a3 = datatypes.Rows.First(r => r.Row == 3).Values[0];
        Assert.Equal(CellKind.Text, a3.Kind);
        Assert.True(a3.FromFormula);
        Assert.Equal(CellKind.Date, datatypes.Rows.First(r => r.Row == 6).Values[0].Kind);

        using var dates1904 = WorkbookSources.Open(Fixtures.PathOf("xlsb/calamine-date_1904.xlsb"));
        Assert.True(dates1904.Date1904);
        var a1 = dates1904.LoadSheet("Sheet1").Rows[0].Values[0];
        Assert.Equal(new DateTime(2021, 1, 1), a1.ToDateTime());

        var summary = WorkbookInspector.Inspect(Fixtures.PathOf("xlsb/poi-testVarious.xlsb"));
        Assert.Equal(WorkbookFormat.Xlsb, summary.Format);
        Assert.Contains(summary.Sheets, s => s.Name == "mySheet1");
    }

    [Fact]
    public void Editor_edits_cells_and_leaves_vba_charts_and_comments_untouched()
    {
        using var dir = new TempDir();
        var seed = Fixtures.CopyTo(dir, "xlsb/poi-testVarious.xlsb");
        var output = dir.File("edited.xlsb");
        IReadOnlySet<string> changed;
        using (var editor = WorkbookEditors.CreateWorkingCopy(seed, output, WorkbookFormat.Xlsb))
        {
            var stats = editor.EditSheet("mySheet1",
                new[] { (2, 2), (3, 2), (50, 50) },
                new[]
                {
                    new CellWrite(2, 2, CellValue.OfNumber(42)),
                    new CellWrite(2, 5, CellValue.OfText("new text")),
                    new CellWrite(3, 2, CellValue.OfDate(ExcelDates.FromDateTime(new DateTime(2026, 3, 31)))),
                    new CellWrite(60, 1, CellValue.OfBoolean(true)),
                    new CellWrite(60, 1200, CellValue.OfError("#N/A")),
                }, null, default);
            Assert.Equal(2, stats.CellsCleared);
            Assert.Equal(5, stats.ValuesWritten);
            Assert.Equal(1, stats.DateStylesApplied);
            changed = editor.Finish();
        }
        OutputVerifier.VerifyUntouchedParts(seed, output, changed);
        foreach (var name in new[] { "xl/vbaProject.bin", "xl/charts/chart1.xml", "xl/drawings/drawing1.xml", "xl/comments1.bin", "xl/printerSettings/printerSettings1.bin", "docProps/core.xml" })
            OpenXmlFixture.AssertSameEntry(seed, output, name);
        Assert.False(OpenXmlFixture.HasEntry(output, "xl/calcChain.bin"));
        Assert.False(OpenXmlFixture.HasEntry(output, "xl/worksheets/binaryIndex1.bin"));
        var rels = System.Text.Encoding.UTF8.GetString(OpenXmlFixture.ReadZipEntry(output, "xl/worksheets/_rels/sheet1.bin.rels"));
        Assert.DoesNotContain("binaryIndex", rels);
        var contentTypes = System.Text.Encoding.UTF8.GetString(OpenXmlFixture.ReadZipEntry(output, "[Content_Types].xml"));
        Assert.DoesNotContain("binaryIndex1.bin", contentTypes);
        Assert.DoesNotContain("calcChain.bin", contentTypes);

        using var reopened = WorkbookSources.Open(output);
        Assert.True(reopened.FullCalcOnLoad);
        var sheet = reopened.LoadSheet("mySheet1");
        CellValue At(int r, int c) { var row = sheet.Rows.First(x => x.Row == r); var i = Array.IndexOf(row.Columns, c); return i < 0 ? CellValue.Blank : row.Values[i]; }
        Assert.Equal(42, At(2, 2).Number);
        Assert.Equal("new text", At(2, 5).Text);
        Assert.Equal(new DateTime(2026, 3, 31), At(3, 2).ToDateTime());
        Assert.True(At(60, 1).BooleanValue);
        Assert.Equal("#N/A", At(60, 1200).Text);
        Assert.Equal("This is a string", At(1, 2).Text); // untouched neighbour
        Assert.True(At(13, 2).FromFormula);              // formulas elsewhere survive

        // Shared string and style tables stay internally consistent.
        using var pkg = OpcZipPackage.Open(output);
        var sst = pkg.ReadPart("xl/sharedStrings.bin");
        var sstRecords = Biff12Stream.Parse(sst, "sst");
        var begin = sstRecords.First(r => r.Type == Rt.SSTItem - 0x13 + 0x9F);
        var unique = BitConverter.ToInt32(sst, begin.DataOffset + 4);
        Assert.Equal(sstRecords.Count(r => r.Type == Rt.SSTItem), unique);
        var styles = pkg.ReadPart("xl/styles.bin");
        var styleRecords = Biff12Stream.Parse(styles, "styles");
        var beginXfs = styleRecords.First(r => r.Type == Rt.BeginCellXFs);
        var cxfs = BitConverter.ToInt32(styles, beginXfs.DataOffset);
        var xfCount = styleRecords.SkipWhile(r => r.Type != Rt.BeginCellXFs).Skip(1).TakeWhile(r => r.Type != Rt.EndCellXFs).Count(r => r.Type == Rt.XF);
        Assert.Equal(cxfs, xfCount);
    }

    private static string BuildXlsbFromSeed(TempDir dir, string name, IEnumerable<CellWrite> writes, IEnumerable<(int, int)>? clears = null)
    {
        var seed = Fixtures.CopyTo(dir, "xlsb/calamine-date.xlsb", "seed-" + name);
        var temp = dir.File(name + ".building.xlsb");
        using (var editor = WorkbookEditors.CreateWorkingCopy(seed, temp, WorkbookFormat.Xlsb))
        {
            editor.EditSheet("Sheet1", (clears ?? new[] { (1, 1), (1, 2), (2, 1), (2, 2), (3, 1), (3, 2) }).ToList(), writes.ToList(), null, default);
            editor.Finish();
        }
        var final = dir.File(name);
        File.Move(temp, final, true);
        return final;
    }

    private static IEnumerable<CellWrite> DestinationLayout()
    {
        var headers = new[] { "ProductCode", "Region", "Description", "Q1 Actual", "Q2 Actual", "Q3 Actual" };
        for (var c = 0; c < headers.Length; c++) yield return new CellWrite(1, c + 1, CellValue.OfText(headers[c]));
        var r = 2;
        foreach (var row in DefaultDestination)
        {
            yield return new CellWrite(r, 1, CellValue.OfText(row.Code!));
            yield return new CellWrite(r, 2, CellValue.OfText(row.Region!));
            yield return new CellWrite(r, 3, CellValue.OfText(row.Description));
            yield return new CellWrite(r, 4, CellValue.OfNumber((double)row.Q1!));
            yield return new CellWrite(r, 5, CellValue.OfNumber((double)row.Q2!));
            yield return new CellWrite(r, 6, CellValue.OfNumber((double)row.Q3!));
            r++;
        }
    }

    private static CellValue Cell(LoadedDataSet data, int row, int col) => data.Records.First(x => x.RowNumber == row)[col];

    [Fact]
    public void Transfer_from_xlsx_source_into_xlsb_destination()
    {
        using var dir = new TempDir();
        var source = CreateSource(dir.File("source.xlsx"));
        var destination = BuildXlsbFromSeed(dir, "destination.xlsb", DestinationLayout());
        var destSelection = new DataSetSelection { SheetName = "Sheet1", HeaderRow = 1 };
        var before = SheetReader.Load(destination, destSelection);
        Assert.Equal(5, before.Records.Count);
        Assert.Equal("Q1 Actual", before.Columns[3].Header);

        var wrongExt = TransferPlanner.Plan(Spec(source, destination, dir.File("o.xlsx"), destinationSelection: destSelection));
        Assert.Contains(wrongExt.Errors, e => e.Code == "output-extension");

        var output = dir.File("output.xlsb");
        var plan = TransferPlanner.Plan(Spec(source, destination, output, destinationSelection: destSelection));
        Assert.True(plan.CanTransfer, string.Join("|", plan.Errors.Select(e => e.Message)));
        Assert.Equal(4, plan.Matches.Count);
        var result = TransferExecutor.Execute(plan);
        Assert.Equal(WorkbookFormat.Xlsb, result.Format);
        Assert.Equal(7, result.ValuesWritten);

        var after = SheetReader.Load(output, destSelection);
        Assert.Equal(110, Cell(after, 2, 4).Number); Assert.Equal(111, Cell(after, 2, 5).Number); Assert.Equal(3, Cell(after, 2, 6).Number);
        Assert.Equal(115, Cell(after, 3, 4).Number); Assert.Equal(116, Cell(after, 3, 5).Number);
        Assert.Equal(220, Cell(after, 4, 4).Number); Assert.True(Cell(after, 4, 5).IsBlank); Assert.Equal(9, Cell(after, 4, 6).Number);
        Assert.Equal(330, Cell(after, 5, 4).Number); Assert.Equal(331.5, Cell(after, 5, 5).Number);
        Assert.True(Cell(after, 6, 4).IsBlank); Assert.True(Cell(after, 6, 5).IsBlank); Assert.Equal(15, Cell(after, 6, 6).Number);
        Assert.Equal("Orphan (no source)", Cell(after, 6, 3).Text);
        OpenXmlFixture.AssertSameEntry(destination, output, "xl/theme/theme1.xml");
        OpenXmlFixture.AssertSameEntry(destination, output, "docProps/app.xml");
    }

    [Fact]
    public void Transfer_from_xlsb_source_into_xlsx_destination()
    {
        using var dir = new TempDir();
        var writes = new List<CellWrite>
        {
            new(1, 1, CellValue.OfText("ProductCode")), new(1, 2, CellValue.OfText("Region")), new(1, 3, CellValue.OfText("2026 Q1")), new(1, 4, CellValue.OfText("2026 Q2")),
        };
        var r = 2;
        foreach (var row in DefaultSource)
        {
            writes.Add(new CellWrite(r, 1, CellValue.OfText(row.Code!)));
            writes.Add(new CellWrite(r, 2, CellValue.OfText(row.Region!)));
            if (row.Q1 is double q1) writes.Add(new CellWrite(r, 3, CellValue.OfNumber(q1)));
            if (row.Q2 is double q2) writes.Add(new CellWrite(r, 4, CellValue.OfNumber(q2)));
            r++;
        }
        var source = BuildXlsbFromSeed(dir, "source.xlsb", writes);
        var destination = CreateDestination(dir.File("destination.xlsx"));
        var plan = TransferPlanner.Plan(Spec(source, destination, dir.File("o.xlsx"), sourceSelection: new DataSetSelection { SheetName = "Sheet1", HeaderRow = 1 }));
        Assert.True(plan.CanTransfer, string.Join("|", plan.Errors.Select(e => e.Message)));
        TransferExecutor.Execute(plan);
        var after = SheetReader.Load(dir.File("o.xlsx"), DestinationTableSelection());
        Assert.Equal(110, Cell(after, 2, DstQ1).Number);
        Assert.Equal(331.5, Cell(after, 5, DstQ2).Number);
        Assert.True(Cell(after, 6, DstQ1).IsBlank);
    }

    [Fact]
    public void Transfer_into_an_excel_made_xlsb_with_formulas_keeps_the_formula_cells_outside_the_range()
    {
        using var dir = new TempDir();
        var destination = Fixtures.CopyTo(dir, "xlsb/poi-sample.xlsb");
        var destSelection = new DataSetSelection { SheetName = "Sheet1", HeaderRow = 1, LastDataRow = 9 }; // row 10 holds a SUM formula
        var source = CreateSource(dir.File("s.xlsx"), new[] { new SourceRow("ipsum", "x", 2000.0, 1.0), new SourceRow("amet", "x", 5000.0, 1.0), new SourceRow("nobody", "x", 1.0, 1.0) });
        var spec = Spec(source, destination, dir.File("o.xlsb"), sourceSelection: SourceSelection(), destinationSelection: destSelection,
            keys: new[] { new ColumnMapping(SrcCode, 1) }, periods: new[] { new ColumnMapping(SrcQ1, 2) });
        var plan = TransferPlanner.Plan(spec);
        Assert.True(plan.CanTransfer, string.Join("|", plan.Errors.Select(e => e.Message)));
        Assert.Equal(2, plan.Matches.Count);
        Assert.Equal(8, plan.TotalCellsToClear);
        TransferExecutor.Execute(plan);

        var after = SheetReader.Load(dir.File("o.xlsb"), new DataSetSelection { SheetName = "Sheet1", HeaderRow = 1, LastDataRow = 10 });
        Assert.Equal(2000, Cell(after, 2, 2).Number);
        Assert.True(Cell(after, 3, 2).IsBlank);
        Assert.Equal(5000, Cell(after, 5, 2).Number);
        var total = Cell(after, 10, 2);
        Assert.True(total.FromFormula);
        Assert.Equal(4995, total.Number); // cached value kept; Excel recalculates on open
        OpenXmlFixture.AssertSameEntry(destination, dir.File("o.xlsb"), "xl/worksheets/sheet2.bin");
        Assert.False(OpenXmlFixture.HasEntry(dir.File("o.xlsb"), "xl/calcChain.bin"));
    }
}
