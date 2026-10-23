using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelTransfer.Core.Cells;
using ExcelTransfer.Core.Data;
using CellValue = ExcelTransfer.Core.Cells.CellValue;

namespace ExcelTransfer.Core.Workbooks;

/// <summary><see cref="IWorkbookSource"/> for .xlsx and .xlsm files, built on the Open XML SDK.</summary>
public sealed class OpenXmlSource : IWorkbookSource
{
    private readonly OpenXmlWorkbook _workbook;
    private IReadOnlyList<SheetInfo>? _sheets;

    private OpenXmlSource(OpenXmlWorkbook workbook, WorkbookFormat format)
    {
        _workbook = workbook;
        Format = format;
    }

    public static OpenXmlSource Open(string path, WorkbookFormat format) => new(OpenXmlWorkbook.OpenRead(path), format);

    public string Path => _workbook.Path;
    public WorkbookFormat Format { get; }
    public bool Date1904 => _workbook.Date1904;
    public bool ManualCalculation => _workbook.ManualCalculation;
    public bool FullCalcOnLoad => _workbook.FullCalcOnLoad;

    public IReadOnlyList<SheetInfo> Sheets => _sheets ??= ReadSheets();

    private List<SheetInfo> ReadSheets()
    {
        var sheets = new List<SheetInfo>();
        var index = 0;
        foreach (var (sheet, part) in _workbook.Sheets)
        {
            var name = sheet.Name?.Value ?? $"Sheet{index + 1}";
            var hidden = sheet.State?.Value is SheetStateValues state && (state == SheetStateValues.Hidden || state == SheetStateValues.VeryHidden);
            var tables = part is WorksheetPart wsPart ? ReadTables(wsPart) : new List<TableInfo>();
            sheets.Add(new SheetInfo(name, index, hidden, part is not WorksheetPart, tables));
            index++;
        }
        return sheets;
    }

    private static List<TableInfo> ReadTables(WorksheetPart part)
    {
        var tables = new List<TableInfo>();
        foreach (var tablePart in part.TableDefinitionParts)
        {
            var info = ReadTable(tablePart.Table);
            if (info is not null) tables.Add(info);
        }
        return tables;
    }

    public static TableInfo? ReadTable(Table? table)
    {
        if (table is null) return null;
        if (!CellRange.TryParse(table.Reference?.Value, out var range)) return null;
        var name = table.Name?.Value ?? table.DisplayName?.Value ?? "Table";
        var displayName = table.DisplayName?.Value ?? name;
        var headerRows = (int)(table.HeaderRowCount?.Value ?? 1);
        var totalsRows = (int)(table.TotalsRowCount?.Value ?? 0);
        var columns = new List<TableColumnInfo>();
        if (table.TableColumns is not null)
        {
            foreach (var column in table.TableColumns.Elements<TableColumn>())
            {
                columns.Add(new TableColumnInfo(
                    column.Name?.Value ?? string.Empty,
                    column.CalculatedColumnFormula is not null,
                    column.TotalsRowFunction is not null || column.TotalsRowFormula is not null));
            }
        }
        return new TableInfo(name, displayName, range, headerRows, totalsRows, columns);
    }

    public SheetSnapshot LoadSheet(string sheetName)
    {
        var (sheet, part) = _workbook.RequireWorksheet(sheetName);
        var worksheet = part.Worksheet ?? throw new Transfer.TransferException($"Sheet \"{sheetName}\" has no worksheet content.");
        var sheetData = worksheet.GetFirstChild<SheetData>();

        var rows = new List<RowSnapshot>();
        var sharedFormulas = new List<SharedFormulaGroup>();
        var arrayFormulas = new List<CellRange>();
        var dataTableFormulas = new List<CellRange>();
        var columnsBuffer = new List<int>();
        var valuesBuffer = new List<CellValue>();

        if (sheetData is not null)
        {
            var rowCounter = 0;
            foreach (var row in sheetData.Elements<Row>())
            {
                var r = (int)(row.RowIndex?.Value ?? (uint)(rowCounter + 1));
                rowCounter = r;
                columnsBuffer.Clear();
                valuesBuffer.Clear();
                var colPos = 0;
                foreach (var cell in row.Elements<Cell>())
                {
                    var c = ResolveColumn(cell, colPos);
                    colPos = c;
                    var formula = cell.CellFormula;
                    if (formula is not null)
                    {
                        var type = formula.FormulaType?.Value;
                        if (type == CellFormulaValues.Shared)
                        {
                            if (formula.Reference?.Value is string sharedRef && CellRange.TryParse(sharedRef, out var sharedRange))
                                sharedFormulas.Add(new SharedFormulaGroup(formula.SharedIndex?.Value ?? 0, sharedRange));
                        }
                        else if (type == CellFormulaValues.Array)
                        {
                            if (formula.Reference?.Value is string arrayRef && CellRange.TryParse(arrayRef, out var arrayRange))
                                arrayFormulas.Add(arrayRange);
                            else
                                arrayFormulas.Add(new CellRange(r, c, r, c));
                        }
                        else if (type == CellFormulaValues.DataTable)
                        {
                            if (formula.Reference?.Value is string dtRef && CellRange.TryParse(dtRef, out var dtRange)) dataTableFormulas.Add(dtRange);
                            else dataTableFormulas.Add(new CellRange(r, c, r, c));
                        }
                    }
                    columnsBuffer.Add(c);
                    valuesBuffer.Add(ParseCell(_workbook, cell));
                }
                rows.Add(new RowSnapshot(r, row.Hidden?.Value == true, columnsBuffer.ToArray(), valuesBuffer.ToArray()));
            }
        }
        rows.Sort((a, b) => a.Row.CompareTo(b.Row));

        var merged = new List<CellRange>();
        var mergeCells = worksheet.GetFirstChild<MergeCells>();
        if (mergeCells is not null)
        {
            foreach (var m in mergeCells.Elements<MergeCell>())
                if (CellRange.TryParse(m.Reference?.Value, out var range)) merged.Add(range);
        }
        var protection = worksheet.GetFirstChild<SheetProtection>();
        var isProtected = protection is not null && protection.Sheet?.Value != false;

        return new SheetSnapshot
        {
            SheetName = sheet.Name?.Value ?? sheetName,
            PartUri = part.Uri.OriginalString,
            Rows = rows,
            Tables = ReadTables(part),
            Protected = isProtected,
            MergedRanges = merged,
            SharedFormulas = sharedFormulas,
            ArrayFormulas = arrayFormulas,
            DataTableFormulas = dataTableFormulas,
        };
    }

    /// <summary>Column index of a cell from its reference, or the next position when the reference is missing.</summary>
    internal static int ResolveColumn(Cell cell, int previousColumn)
    {
        if (CellRef.TryParse(cell.CellReference?.Value, out _, out var col)) return col;
        return previousColumn + 1;
    }

    /// <summary>Interprets a cell element as a typed value, using the cached result when the cell holds a formula.</summary>
    public static CellValue ParseCell(OpenXmlWorkbook wb, Cell cell)
    {
        var flags = CellFlags.None;
        var hasFormula = cell.CellFormula is not null;
        if (hasFormula) flags |= CellFlags.FromFormula;
        var dateStyled = wb.IsDateStyle(cell.StyleIndex?.Value);
        if (dateStyled) flags |= CellFlags.DateStyled;

        var type = cell.DataType?.Value;
        var text = cell.CellValue?.Text;

        if (type == CellValues.SharedString)
        {
            if (string.IsNullOrEmpty(text) || !int.TryParse(text, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out var index) || index < 0 || index >= wb.SharedStrings.Count)
                return CellValue.BlankWithFlags(flags);
            var s = wb.SharedStrings[index];
            return s.Length == 0 ? CellValue.BlankWithFlags(flags) : CellValue.OfText(s, flags);
        }
        if (type == CellValues.InlineString)
        {
            var s = cell.InlineString is null ? string.Empty : OpenXmlWorkbook.FlattenText(cell.InlineString);
            return s.Length == 0 ? CellValue.BlankWithFlags(flags) : CellValue.OfText(s, flags);
        }
        if (type == CellValues.String)
            return string.IsNullOrEmpty(text) ? CellValue.BlankWithFlags(flags) : CellValue.OfText(text, flags);
        if (type == CellValues.Boolean)
        {
            if (string.IsNullOrEmpty(text)) return hasFormula ? CellValue.FormulaWithoutResult(cell.CellFormula!.Text, flags) : CellValue.BlankWithFlags(flags);
            return CellValue.OfBoolean(text == "1" || text.Equals("true", StringComparison.OrdinalIgnoreCase), flags);
        }
        if (type == CellValues.Error)
            return CellValue.OfError(string.IsNullOrEmpty(text) ? "#N/A" : text, flags);
        if (type == CellValues.Date)
        {
            if (string.IsNullOrEmpty(text)) return hasFormula ? CellValue.FormulaWithoutResult(cell.CellFormula!.Text, flags) : CellValue.BlankWithFlags(flags);
            if (DateTime.TryParse(text, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.RoundtripKind, out var dt))
                return CellValue.OfDate(ExcelDates.FromDateTime(dt), flags);
            return CellValue.OfText(text, flags);
        }

        if (text is null)
            return hasFormula ? CellValue.FormulaWithoutResult(cell.CellFormula!.Text, flags) : CellValue.BlankWithFlags(flags);
        if (text.Length == 0)
            return CellValue.BlankWithFlags(flags);
        if (!double.TryParse(text, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var number))
            return CellValue.OfText(text, flags);
        if (double.IsNaN(number) || double.IsInfinity(number))
            return CellValue.OfError("#NUM!", flags);
        return dateStyled
            ? CellValue.OfDate(ExcelDates.ToSerial1900(number, wb.Date1904), flags)
            : CellValue.OfNumber(number, flags);
    }

    public void Dispose() => _workbook.Dispose();
}
