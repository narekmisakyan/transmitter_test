using System.IO.Compression;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelTransfer.Tests.Support;

/// <summary>Low-level manipulations of test workbooks that ClosedXML cannot express (shared formulas, missing cached values, VBA).</summary>
public static class OpenXmlFixture
{
    public static void EditWorksheet(string path, string sheetName, Action<WorkbookPart, WorksheetPart> edit)
    {
        using var doc = SpreadsheetDocument.Open(path, true);
        var wb = doc.WorkbookPart!;
        var sheet = wb.Workbook.Sheets!.Elements<Sheet>().First(s => string.Equals(s.Name!.Value, sheetName, StringComparison.OrdinalIgnoreCase));
        var part = (WorksheetPart)wb.GetPartById(sheet.Id!.Value!);
        edit(wb, part);
        part.Worksheet.Save();
        doc.Save();
    }

    public static Cell GetCell(WorksheetPart part, string a1)
    {
        var cells = part.Worksheet.GetFirstChild<SheetData>()!.Descendants<Cell>();
        return cells.First(c => string.Equals(c.CellReference?.Value, a1, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>Turns the given cells into one shared formula group (Excel's "fill" representation).</summary>
    public static void MakeSharedFormulaGroup(string path, string sheetName, string masterA1, string range, string formula, params string[] dependents)
    {
        EditWorksheet(path, sheetName, (_, part) =>
        {
            var master = GetCell(part, masterA1);
            master.CellFormula = new CellFormula(formula) { FormulaType = CellFormulaValues.Shared, Reference = range, SharedIndex = 0 };
            foreach (var d in dependents)
            {
                var cell = GetCell(part, d);
                cell.CellFormula = new CellFormula { FormulaType = CellFormulaValues.Shared, SharedIndex = 0 };
            }
        });
    }

    public static void RemoveCachedValue(string path, string sheetName, string a1)
    {
        EditWorksheet(path, sheetName, (_, part) =>
        {
            var cell = GetCell(part, a1);
            cell.CellValue?.Remove();
            cell.DataType = null;
        });
    }

    public static void SetCalculatedColumn(string path, string sheetName, string columnName, string formula)
    {
        EditWorksheet(path, sheetName, (_, part) =>
        {
            var table = part.TableDefinitionParts.First().Table;
            var column = table.TableColumns!.Elements<TableColumn>().First(c => c.Name!.Value == columnName);
            column.CalculatedColumnFormula = new CalculatedColumnFormula(formula);
            part.TableDefinitionParts.First().Table.Save();
        });
    }

    /// <summary>Copies an .xlsx to an .xlsm path, switches the content type and embeds the given VBA project bytes.</summary>
    public static string ConvertToXlsm(string xlsxPath, string xlsmPath, byte[] vbaProject)
    {
        File.Copy(xlsxPath, xlsmPath, overwrite: true);
        using var doc = SpreadsheetDocument.Open(xlsmPath, true);
        doc.ChangeDocumentType(SpreadsheetDocumentType.MacroEnabledWorkbook);
        var part = doc.WorkbookPart!.AddNewPart<VbaProjectPart>();
        using (var ms = new MemoryStream(vbaProject)) part.FeedData(ms);
        doc.Save();
        return xlsmPath;
    }

    public static byte[] ReadZipEntry(string path, string entryName)
    {
        using var zip = ZipFile.OpenRead(path);
        var entry = zip.Entries.First(e => string.Equals(e.FullName, entryName, StringComparison.OrdinalIgnoreCase));
        using var s = entry.Open();
        using var ms = new MemoryStream();
        s.CopyTo(ms);
        return ms.ToArray();
    }

    public static IReadOnlyList<string> EntryNames(string path)
    {
        using var zip = ZipFile.OpenRead(path);
        return zip.Entries.Select(e => e.FullName).Where(n => !n.EndsWith('/')).ToList();
    }

    public static bool HasEntry(string path, string entryName) => EntryNames(path).Any(n => string.Equals(n, entryName, StringComparison.OrdinalIgnoreCase));

    public static void AssertSameEntry(string a, string b, string entryName)
    {
        Assert.True(HasEntry(a, entryName), $"{entryName} missing in {Path.GetFileName(a)}");
        Assert.True(HasEntry(b, entryName), $"{entryName} missing in {Path.GetFileName(b)}");
        Assert.True(ReadZipEntry(a, entryName).AsSpan().SequenceEqual(ReadZipEntry(b, entryName)), $"{entryName} differs");
    }

    public static string Sha256(string path)
    {
        using var s = File.OpenRead(path);
        return Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(s));
    }

    // 1x1 transparent PNG.
    public static readonly byte[] TinyPng = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==");
}
