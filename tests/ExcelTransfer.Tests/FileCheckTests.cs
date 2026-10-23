using System.IO.Compression;
using ExcelTransfer.Core.Workbooks;
using ExcelTransfer.Tests.Support;

namespace ExcelTransfer.Tests;

public class FileCheckTests
{
    [Fact]
    public void Missing_file_and_empty_path_are_reported()
    {
        Assert.Equal(FileProblem.Missing, WorkbookFileCheck.Check(null).Problem);
        Assert.Equal(FileProblem.Missing, WorkbookFileCheck.Check(Path.Combine(Path.GetTempPath(), "does-not-exist-" + Guid.NewGuid() + ".xlsx")).Problem);
    }

    [Theory]
    [InlineData("legacy.xls", FileProblem.LegacyFormat)]
    [InlineData("data.csv", FileProblem.UnsupportedExtension)]
    [InlineData("sheet.ods", FileProblem.UnsupportedExtension)]
    [InlineData("tmpl.xltx", FileProblem.UnsupportedExtension)]
    [InlineData("notes.txt", FileProblem.UnsupportedExtension)]
    public void Unsupported_extensions_are_rejected_with_a_message(string name, FileProblem expected)
    {
        using var dir = new TempDir();
        var path = dir.File(name);
        File.WriteAllText(path, "whatever");
        var result = WorkbookFileCheck.Check(path);
        Assert.False(result.Ok);
        Assert.Equal(expected, result.Problem);
        Assert.False(string.IsNullOrWhiteSpace(result.Message));
    }

    [Fact]
    public void Encrypted_or_ole_files_are_recognised()
    {
        using var dir = new TempDir();
        var path = dir.File("secret.xlsx");
        var bytes = new byte[64];
        new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 }.CopyTo(bytes, 0);
        File.WriteAllBytes(path, bytes);
        var result = WorkbookFileCheck.Check(path);
        Assert.Equal(FileProblem.Encrypted, result.Problem);
        Assert.Contains("password", result.Message!, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Random_bytes_and_non_workbook_zips_are_invalid()
    {
        using var dir = new TempDir();
        var junk = dir.File("junk.xlsx");
        File.WriteAllBytes(junk, new byte[] { 1, 2, 3, 4, 5, 6, 7, 8, 9 });
        Assert.Equal(FileProblem.Invalid, WorkbookFileCheck.Check(junk).Problem);

        var zip = dir.File("zip.xlsx");
        using (var archive = ZipFile.Open(zip, ZipArchiveMode.Create))
        {
            var entry = archive.CreateEntry("readme.txt");
            using var w = new StreamWriter(entry.Open());
            w.Write("not a workbook");
        }
        Assert.Equal(FileProblem.Invalid, WorkbookFileCheck.Check(zip).Problem);
    }

    [Fact]
    public void Valid_workbooks_report_their_format()
    {
        using var dir = new TempDir();
        var xlsx = WorkbookBuilder.CreateSource(dir.File("s.xlsx"));
        var xlsb = Fixtures.CopyTo(dir, "xlsb/poi-sample.xlsb");
        var xlsm = Fixtures.CopyTo(dir, "xlsm/poi-SimpleMacro.xlsm");
        Assert.Equal(WorkbookFormat.Xlsx, WorkbookFileCheck.Check(xlsx).Format);
        Assert.Equal(WorkbookFormat.Xlsb, WorkbookFileCheck.Check(xlsb).Format);
        Assert.Equal(WorkbookFormat.Xlsm, WorkbookFileCheck.Check(xlsm).Format);
    }

    [Fact]
    public void Extension_that_does_not_match_the_content_is_rejected()
    {
        using var dir = new TempDir();
        var renamedXlsb = Fixtures.CopyTo(dir, "xlsb/poi-sample.xlsb", "pretend.xlsx");
        var result = WorkbookFileCheck.Check(renamedXlsb);
        Assert.Equal(FileProblem.FormatMismatch, result.Problem);
        Assert.Contains(".xlsb", result.Hint!);

        var renamedXlsm = Fixtures.CopyTo(dir, "xlsm/poi-SimpleMacro.xlsm", "pretend2.xlsx");
        Assert.Equal(FileProblem.FormatMismatch, WorkbookFileCheck.Check(renamedXlsm).Problem);
    }

    [Fact]
    public void Inspector_lists_sheets_and_tables()
    {
        using var dir = new TempDir();
        var dest = WorkbookBuilder.CreateDestination(dir.File("d.xlsx"));
        var summary = WorkbookInspector.Inspect(dest);
        Assert.Equal(WorkbookFormat.Xlsx, summary.Format);
        Assert.Equal(new[] { "Plan", "Notes" }, summary.Sheets.Select(s => s.Name));
        var plan = summary.FindSheet("plan")!;
        var table = Assert.Single(plan.Tables);
        Assert.Equal("Products", table.DisplayName);
        Assert.Equal(1, table.HeaderRowCount);
        Assert.Equal(1, table.TotalsRowCount);
        Assert.Equal(7, table.Columns.Count);
        Assert.Equal(2, table.FirstDataRow);
        Assert.Equal(6, table.LastDataRow);
    }
}
