using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelTransfer.Core.Cells;
using ExcelTransfer.Core.Transfer;

namespace ExcelTransfer.Core.Workbooks;

/// <summary>
/// Read-only access to an .xlsx workbook through the Open XML SDK, with the shared string table and the
/// date-format flags of the style table cached. The file is opened with shared read/write access so a
/// workbook that is currently open in Excel can still be read.
/// </summary>
public sealed class OpenXmlWorkbook : IDisposable
{
    private readonly FileStream _stream;
    private readonly SpreadsheetDocument _document;
    private List<string>? _sharedStrings;
    private bool[]? _dateStyles;

    public string Path { get; }
    public WorkbookPart WorkbookPart { get; }
    public bool Date1904 { get; }
    public bool ManualCalculation { get; }
    public bool FullCalcOnLoad { get; }

    private OpenXmlWorkbook(string path, FileStream stream, SpreadsheetDocument document)
    {
        Path = path;
        _stream = stream;
        _document = document;
        WorkbookPart = document.WorkbookPart ?? throw new TransferException($"\"{System.IO.Path.GetFileName(path)}\" does not contain a workbook.");
        var workbook = WorkbookPart.Workbook ?? throw new TransferException($"\"{System.IO.Path.GetFileName(path)}\" does not contain a workbook.");
        Date1904 = workbook.WorkbookProperties?.Date1904?.Value == true;
        var calc = workbook.CalculationProperties;
        ManualCalculation = calc?.CalculationMode?.Value == CalculateModeValues.Manual;
        FullCalcOnLoad = calc?.FullCalculationOnLoad?.Value == true;
    }

    public static OpenXmlWorkbook OpenRead(string path)
    {
        var fullPath = System.IO.Path.GetFullPath(path);
        var name = System.IO.Path.GetFileName(fullPath);
        FileStream stream;
        try
        {
            stream = new FileStream(fullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        }
        catch (FileNotFoundException ex)
        {
            throw new TransferException($"The file \"{fullPath}\" does not exist.", "Check the path and that the drive or network share is available.", ex);
        }
        catch (DirectoryNotFoundException ex)
        {
            throw new TransferException($"The folder of \"{fullPath}\" does not exist.", "Check the path and that the drive or network share is available.", ex);
        }
        catch (UnauthorizedAccessException ex)
        {
            throw new TransferException($"You do not have permission to read \"{fullPath}\".", "Check the file and folder permissions.", ex);
        }
        catch (IOException ex)
        {
            throw new TransferException($"\"{name}\" could not be opened: {ex.Message}", "If the file is open in another program, close it and try again.", ex);
        }

        SpreadsheetDocument document;
        try
        {
            document = SpreadsheetDocument.Open(stream, false);
        }
        catch (Exception ex) when (ex is not OutOfMemoryException)
        {
            stream.Dispose();
            throw new TransferException($"\"{name}\" could not be read as an .xlsx workbook: {ex.Message}", "Open it in Excel and save it again as .xlsx.", ex);
        }

        try
        {
            return new OpenXmlWorkbook(fullPath, stream, document);
        }
        catch
        {
            document.Dispose();
            stream.Dispose();
            throw;
        }
    }

    /// <summary>All sheets in workbook order (worksheets and chart sheets).</summary>
    public IEnumerable<(Sheet Sheet, OpenXmlPart? Part)> Sheets
    {
        get
        {
            var sheets = WorkbookPart.Workbook?.Sheets;
            if (sheets is null) yield break;
            foreach (var sheet in sheets.Elements<Sheet>())
            {
                OpenXmlPart? part = null;
                var id = sheet.Id?.Value;
                if (!string.IsNullOrEmpty(id))
                {
                    try { part = WorkbookPart.GetPartById(id); }
                    catch (ArgumentOutOfRangeException) { part = null; }
                    catch (KeyNotFoundException) { part = null; }
                }
                yield return (sheet, part);
            }
        }
    }

    /// <summary>Finds a worksheet by name. Excel sheet names are not case-sensitive.</summary>
    public (Sheet Sheet, WorksheetPart Part) RequireWorksheet(string sheetName)
    {
        foreach (var (sheet, part) in Sheets)
        {
            if (!string.Equals(sheet.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase)) continue;
            if (part is WorksheetPart wsPart) return (sheet, wsPart);
            throw new TransferException($"Sheet \"{sheetName}\" in \"{System.IO.Path.GetFileName(Path)}\" is a chart sheet and contains no cells.", "Select a worksheet instead.");
        }
        throw new TransferException($"Sheet \"{sheetName}\" was not found in \"{System.IO.Path.GetFileName(Path)}\".", "The workbook may have been changed since it was selected. Select the sheet again.");
    }

    public IReadOnlyList<string> SharedStrings => _sharedStrings ??= LoadSharedStrings();

    private List<string> LoadSharedStrings()
    {
        var list = new List<string>();
        var table = WorkbookPart.SharedStringTablePart?.SharedStringTable;
        if (table is null) return list;
        foreach (var item in table.Elements<SharedStringItem>())
            list.Add(FlattenText(item));
        return list;
    }

    /// <summary>Concatenates the plain text of a shared string item or inline string, skipping phonetic (furigana) runs.</summary>
    public static string FlattenText(OpenXmlElement item)
    {
        var sb = new StringBuilder();
        foreach (var child in item.ChildElements)
        {
            switch (child)
            {
                case Text t:
                    sb.Append(t.Text);
                    break;
                case Run r:
                    foreach (var t in r.Elements<Text>()) sb.Append(t.Text);
                    break;
                // PhoneticRun and PhoneticProperties are intentionally skipped.
            }
        }
        return sb.ToString();
    }

    public bool IsDateStyle(uint? styleIndex)
    {
        if (styleIndex is null) return false;
        var flags = _dateStyles ??= LoadDateStyles();
        return styleIndex.Value < flags.Length && flags[styleIndex.Value];
    }

    private bool[] LoadDateStyles()
    {
        var stylesheet = WorkbookPart.WorkbookStylesPart?.Stylesheet;
        var cellFormats = stylesheet?.CellFormats;
        if (cellFormats is null) return Array.Empty<bool>();

        var customFormats = new Dictionary<uint, string>();
        var numberingFormats = stylesheet!.NumberingFormats;
        if (numberingFormats is not null)
        {
            foreach (var nf in numberingFormats.Elements<NumberingFormat>())
            {
                if (nf.NumberFormatId is not null && nf.FormatCode?.Value is not null)
                    customFormats[nf.NumberFormatId.Value] = nf.FormatCode.Value;
            }
        }

        var list = new List<bool>();
        foreach (var xf in cellFormats.Elements<CellFormat>())
        {
            var id = xf.NumberFormatId?.Value ?? 0;
            customFormats.TryGetValue(id, out var code);
            list.Add(NumberFormats.IsDateFormat((int)id, code));
        }
        return list.ToArray();
    }

    public void Dispose()
    {
        _document.Dispose();
        _stream.Dispose();
    }
}
