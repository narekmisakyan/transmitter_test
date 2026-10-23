using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelTransfer.Core.Cells;
using ExcelTransfer.Core.Transfer;
using CellValue = ExcelTransfer.Core.Cells.CellValue;
using XCellValue = DocumentFormat.OpenXml.Spreadsheet.CellValue;

namespace ExcelTransfer.Core.Workbooks;

/// <summary>
/// Edits .xlsx/.xlsm working copies with the Open XML SDK. Only the target worksheet part, the shared string
/// table, the style sheet (when a date format must be added) and the workbook part (calculation flag) are
/// touched; all other parts (other sheets, charts, pivot caches, VBA, custom XML, ...) are copied as they are.
/// </summary>
internal sealed class OpenXmlEditor : IWorkbookEditor
{
    private readonly string _path;
    private readonly SpreadsheetDocument _document;
    private readonly WorkbookPart _workbookPart;
    private readonly bool _date1904;
    private readonly HashSet<string> _changedParts = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<uint, uint> _dateStyleCache = new();
    private Dictionary<string, int>? _sharedStringIndex;
    private SharedStringTable? _sharedStrings;
    private int _sharedStringRefsAdded;
    private List<bool>? _dateStyles;
    private bool _disposed;

    private OpenXmlEditor(string path, SpreadsheetDocument document, WorkbookFormat format)
    {
        _path = path;
        _document = document;
        Format = format;
        _workbookPart = document.WorkbookPart ?? throw new TransferException("The working copy has no workbook part.");
        _date1904 = _workbookPart.Workbook?.WorkbookProperties?.Date1904?.Value == true;
    }

    public WorkbookFormat Format { get; }

    public static OpenXmlEditor Create(string sourcePath, string workingCopyPath, WorkbookFormat format)
    {
        File.Copy(sourcePath, workingCopyPath, overwrite: true);
        File.SetAttributes(workingCopyPath, FileAttributes.Normal);
        SpreadsheetDocument document;
        try
        {
            document = SpreadsheetDocument.Open(workingCopyPath, true);
        }
        catch (Exception ex) when (ex is not OutOfMemoryException)
        {
            throw new TransferException($"The destination workbook could not be opened for editing: {ex.Message}", "Open it in Excel and save it again.", ex);
        }
        return new OpenXmlEditor(workingCopyPath, document, format);
    }

    public SheetEditStats EditSheet(string sheetName, IReadOnlyCollection<(int Row, int Column)> cellsToClear, IReadOnlyList<CellWrite> writes, IProgress<ProgressReport>? progress, CancellationToken ct)
    {
        var part = FindWorksheetPart(sheetName);
        _changedParts.Add(part.Uri.OriginalString);
        var worksheet = part.Worksheet ?? throw new TransferException($"Sheet \"{sheetName}\" has no worksheet content.");
        var sheetData = worksheet.GetFirstChild<SheetData>();
        if (sheetData is null)
        {
            sheetData = new SheetData();
            worksheet.AppendChild(sheetData);
        }

        var rows = new SortedDictionary<uint, Row>();
        uint counter = 0;
        foreach (var row in sheetData.Elements<Row>())
        {
            if (row.RowIndex is null) row.RowIndex = counter + 1;
            counter = row.RowIndex.Value;
            rows[counter] = row;
        }
        var cellMaps = new Dictionary<uint, Dictionary<int, Cell>>();
        var columnStyles = ReadColumnStyles(worksheet);

        var cleared = 0;
        var total = cellsToClear.Count + writes.Count;
        var done = 0;
        foreach (var (r, c) in cellsToClear)
        {
            ct.ThrowIfCancellationRequested();
            if (rows.TryGetValue((uint)r, out var row))
            {
                var cells = CellMap(row, cellMaps);
                if (cells.TryGetValue(c, out var cell) && ClearCell(cell, cells, c)) cleared++;
            }
            if (++done % 5000 == 0) progress?.Report(new ProgressReport("Clearing cells", 0.1 + 0.5 * done / Math.Max(1, total)));
        }

        var written = 0;
        var dateStylesApplied = 0;
        int maxRow = 0, maxCol = 0;
        foreach (var write in writes)
        {
            ct.ThrowIfCancellationRequested();
            var row = GetOrCreateRow(sheetData, rows, (uint)write.Row);
            var cells = CellMap(row, cellMaps);
            if (write.Value.IsBlank)
            {
                if (cells.TryGetValue(write.Column, out var existing) && ClearCell(existing, cells, write.Column)) cleared++;
                continue;
            }
            var cell = GetOrCreateCell(row, cells, write.Row, write.Column, columnStyles);
            if (SetValue(cell, write.Value)) dateStylesApplied++;
            written++;
            if (write.Row > maxRow) maxRow = write.Row;
            if (write.Column > maxCol) maxCol = write.Column;
            if (++done % 5000 == 0) progress?.Report(new ProgressReport("Writing values", 0.1 + 0.5 * done / Math.Max(1, total)));
        }

        ExtendDimension(worksheet, maxRow, maxCol);
        return new SheetEditStats(cleared, written, dateStylesApplied);
    }

    private WorksheetPart FindWorksheetPart(string sheetName)
    {
        var sheets = _workbookPart.Workbook?.Sheets?.Elements<Sheet>() ?? Enumerable.Empty<Sheet>();
        foreach (var sheet in sheets)
        {
            if (!string.Equals(sheet.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase)) continue;
            if (sheet.Id?.Value is string id && _workbookPart.GetPartById(id) is WorksheetPart wsPart) return wsPart;
            throw new TransferException($"Sheet \"{sheetName}\" is not a worksheet.");
        }
        throw new TransferException($"Sheet \"{sheetName}\" was not found in the destination workbook.");
    }

    private static List<(uint Min, uint Max, uint Style)> ReadColumnStyles(Worksheet worksheet)
    {
        var list = new List<(uint, uint, uint)>();
        var columns = worksheet.GetFirstChild<Columns>();
        if (columns is null) return list;
        foreach (var col in columns.Elements<Column>())
        {
            if (col.Style?.Value is uint style && style != 0 && col.Min is not null && col.Max is not null)
                list.Add((col.Min.Value, col.Max.Value, style));
        }
        return list;
    }

    private static Dictionary<int, Cell> CellMap(Row row, Dictionary<uint, Dictionary<int, Cell>> cache)
    {
        var key = row.RowIndex!.Value;
        if (cache.TryGetValue(key, out var map)) return map;
        map = new Dictionary<int, Cell>();
        var colPos = 0;
        foreach (var cell in row.Elements<Cell>())
        {
            int col;
            if (CellRef.TryParse(cell.CellReference?.Value, out _, out var parsed)) col = parsed;
            else
            {
                col = colPos + 1;
                cell.CellReference = CellRef.ToA1((int)key, col);
            }
            colPos = col;
            map[col] = cell;
        }
        cache[key] = map;
        return map;
    }

    private static Row GetOrCreateRow(SheetData sheetData, SortedDictionary<uint, Row> rows, uint rowIndex)
    {
        if (rows.TryGetValue(rowIndex, out var row)) return row;
        row = new Row { RowIndex = rowIndex };
        Row? next = null;
        foreach (var kv in rows)
        {
            if (kv.Key > rowIndex) { next = kv.Value; break; }
        }
        if (next is null) sheetData.AppendChild(row);
        else sheetData.InsertBefore(row, next);
        rows[rowIndex] = row;
        return row;
    }

    private static Cell GetOrCreateCell(Row row, Dictionary<int, Cell> cells, int rowIndex, int column, List<(uint Min, uint Max, uint Style)> columnStyles)
    {
        if (cells.TryGetValue(column, out var cell)) return cell;
        cell = new Cell { CellReference = CellRef.ToA1(rowIndex, column) };
        uint? style = null;
        if (row.CustomFormat?.Value == true && row.StyleIndex is not null) style = row.StyleIndex.Value;
        else
        {
            foreach (var (min, max, s) in columnStyles)
            {
                if (column >= min && column <= max) { style = s; break; }
            }
        }
        if (style is uint st && st != 0) cell.StyleIndex = st;

        Cell? next = null;
        var nextCol = int.MaxValue;
        foreach (var kv in cells)
        {
            if (kv.Key > column && kv.Key < nextCol) { nextCol = kv.Key; next = kv.Value; }
        }
        if (next is null) row.AppendChild(cell);
        else row.InsertBefore(cell, next);
        cells[column] = cell;
        return cell;
    }

    /// <summary>Removes value, formula and type from a cell, keeping its style. Returns true when the cell had content.</summary>
    private static bool ClearCell(Cell cell, Dictionary<int, Cell> cells, int column)
    {
        var hadContent = cell.CellFormula is not null || cell.InlineString is not null || !string.IsNullOrEmpty(cell.CellValue?.Text);
        cell.CellFormula?.Remove();
        cell.CellValue?.Remove();
        cell.InlineString?.Remove();
        cell.DataType = null;
        if (cell.StyleIndex is null || cell.StyleIndex.Value == 0)
        {
            cell.Remove();
            cells.Remove(column);
        }
        return hadContent;
    }

    /// <summary>Writes a value. Returns true when a date number format had to be applied to the cell.</summary>
    private bool SetValue(Cell cell, CellValue value)
    {
        cell.CellFormula?.Remove();
        cell.InlineString?.Remove();
        cell.CellValue?.Remove();
        var appliedDateStyle = false;
        switch (value.Kind)
        {
            case CellKind.Text:
                cell.DataType = CellValues.SharedString;
                cell.CellValue = new XCellValue(GetSharedStringIndex(value.TextOrEmpty).ToString(CultureInfo.InvariantCulture));
                break;
            case CellKind.Number:
                cell.DataType = null;
                cell.CellValue = new XCellValue(value.Number.ToString("R", CultureInfo.InvariantCulture));
                break;
            case CellKind.Date:
            {
                cell.DataType = null;
                var serial = ExcelDates.FromSerial1900(value.Number, _date1904);
                cell.CellValue = new XCellValue(serial.ToString("R", CultureInfo.InvariantCulture));
                var baseStyle = cell.StyleIndex?.Value ?? 0;
                if (!IsDateStyle(baseStyle))
                {
                    cell.StyleIndex = GetOrCreateDateStyle(baseStyle, Math.Abs(serial - Math.Floor(serial)) > 1e-9);
                    appliedDateStyle = true;
                }
                break;
            }
            case CellKind.Boolean:
                cell.DataType = CellValues.Boolean;
                cell.CellValue = new XCellValue(value.BooleanValue ? "1" : "0");
                break;
            case CellKind.Error:
                cell.DataType = CellValues.Error;
                cell.CellValue = new XCellValue(value.TextOrEmpty.Length == 0 ? "#N/A" : value.TextOrEmpty);
                break;
            default:
                cell.DataType = null;
                break;
        }
        return appliedDateStyle;
    }

    private SharedStringTable SharedStrings
    {
        get
        {
            if (_sharedStrings is not null) return _sharedStrings;
            var part = _workbookPart.SharedStringTablePart;
            if (part is null)
            {
                part = _workbookPart.AddNewPart<SharedStringTablePart>();
                part.SharedStringTable = new SharedStringTable();
            }
            part.SharedStringTable ??= new SharedStringTable();
            _changedParts.Add(part.Uri.OriginalString);
            _sharedStrings = part.SharedStringTable;
            return _sharedStrings;
        }
    }

    private int GetSharedStringIndex(string text)
    {
        var table = SharedStrings;
        if (_sharedStringIndex is null)
        {
            _sharedStringIndex = new Dictionary<string, int>(StringComparer.Ordinal);
            var i = 0;
            foreach (var item in table.Elements<SharedStringItem>())
            {
                // Only plain items are reusable; rich-text items would carry their formatting into the new cell.
                if (item.ChildElements.Count == 1 && item.Text is not null)
                    _sharedStringIndex.TryAdd(item.Text.Text, i);
                i++;
            }
            _sharedStringIndex["\0count"] = i; // sentinel storing the item count
        }
        _sharedStringRefsAdded++;
        if (_sharedStringIndex.TryGetValue(text, out var index) && !ReferenceEquals(text, "\0count")) return index;
        var count = _sharedStringIndex["\0count"];
        var newItem = new SharedStringItem(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        table.AppendChild(newItem);
        _sharedStringIndex[text] = count;
        _sharedStringIndex["\0count"] = count + 1;
        return count;
    }

    private Stylesheet StylesheetForEdit
    {
        get
        {
            var part = _workbookPart.WorkbookStylesPart;
            if (part is null)
            {
                part = _workbookPart.AddNewPart<WorkbookStylesPart>();
                part.Stylesheet = new Stylesheet();
            }
            part.Stylesheet ??= new Stylesheet();
            _changedParts.Add(part.Uri.OriginalString);
            return part.Stylesheet;
        }
    }

    private List<bool> DateStyles
    {
        get
        {
            if (_dateStyles is not null) return _dateStyles;
            _dateStyles = new List<bool>();
            var stylesheet = _workbookPart.WorkbookStylesPart?.Stylesheet;
            if (stylesheet is not null)
            {
                _changedParts.Add(_workbookPart.WorkbookStylesPart!.Uri.OriginalString);
                var custom = new Dictionary<uint, string>();
                if (stylesheet.NumberingFormats is not null)
                    foreach (var nf in stylesheet.NumberingFormats.Elements<NumberingFormat>())
                        if (nf.NumberFormatId is not null && nf.FormatCode?.Value is not null) custom[nf.NumberFormatId.Value] = nf.FormatCode.Value;
                if (stylesheet.CellFormats is not null)
                    foreach (var xf in stylesheet.CellFormats.Elements<CellFormat>())
                    {
                        var id = xf.NumberFormatId?.Value ?? 0;
                        custom.TryGetValue(id, out var code);
                        _dateStyles.Add(NumberFormats.IsDateFormat((int)id, code));
                    }
            }
            return _dateStyles;
        }
    }

    private bool IsDateStyle(uint styleIndex) => styleIndex < DateStyles.Count && DateStyles[(int)styleIndex];

    /// <summary>Clones the cell's format with a built-in date (14) or date-time (22) number format so the value displays as a date.</summary>
    private uint GetOrCreateDateStyle(uint baseStyle, bool withTime)
    {
        var key = baseStyle * 2 + (withTime ? 1u : 0u);
        if (_dateStyleCache.TryGetValue(key, out var cached)) return cached;
        var stylesheet = StylesheetForEdit;
        stylesheet.CellFormats ??= new CellFormats();
        var formats = stylesheet.CellFormats.Elements<CellFormat>().ToList();
        var clone = baseStyle < formats.Count ? (CellFormat)formats[(int)baseStyle].CloneNode(true) : new CellFormat();
        clone.NumberFormatId = withTime ? 22u : 14u;
        clone.ApplyNumberFormat = true;
        stylesheet.CellFormats.AppendChild(clone);
        var newIndex = (uint)formats.Count;
        stylesheet.CellFormats.Count = newIndex + 1;
        DateStyles.Add(true);
        _dateStyleCache[key] = newIndex;
        return newIndex;
    }

    private static void ExtendDimension(Worksheet worksheet, int maxRow, int maxCol)
    {
        if (maxRow == 0 || maxCol == 0) return;
        var dim = worksheet.SheetDimension;
        if (dim is null || !CellRange.TryParse(dim.Reference?.Value, out var range)) return;
        if (range.LastRow >= maxRow && range.LastColumn >= maxCol) return;
        var extended = new CellRange(range.FirstRow, range.FirstColumn, Math.Max(range.LastRow, maxRow), Math.Max(range.LastColumn, maxCol));
        dim.Reference = extended.ToString();
    }

    public IReadOnlySet<string> Finish()
    {
        var workbook = _workbookPart.Workbook ?? throw new TransferException("The working copy has no workbook content.");
        var calcChain = _workbookPart.CalculationChainPart;
        if (calcChain is not null)
        {
            _changedParts.Add(calcChain.Uri.OriginalString);
            _workbookPart.DeletePart(calcChain);
        }
        var calc = workbook.CalculationProperties;
        if (calc is null)
        {
            calc = new CalculationProperties();
            workbook.CalculationProperties = calc;
        }
        calc.FullCalculationOnLoad = true;
        _changedParts.Add(_workbookPart.Uri.OriginalString);

        if (_sharedStrings is not null)
        {
            var unique = (uint)_sharedStrings.Elements<SharedStringItem>().Count();
            _sharedStrings.UniqueCount = unique;
            var previous = _sharedStrings.Count?.Value ?? unique;
            _sharedStrings.Count = previous + (uint)_sharedStringRefsAdded;
        }

        _changedParts.Add("/[Content_Types].xml");
        _changedParts.Add("/" + Biff12.OpcZipPackage.RelationshipsPartName(_workbookPart.Uri.OriginalString.TrimStart('/')));
        try
        {
            _document.Save();
            _document.Dispose();
            _disposed = true;
        }
        catch (Exception ex) when (ex is not OutOfMemoryException)
        {
            throw new TransferException($"The output workbook could not be saved: {ex.Message}", null, ex);
        }
        return _changedParts;
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        _document.Dispose();
    }
}
