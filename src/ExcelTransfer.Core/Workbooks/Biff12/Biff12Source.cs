using ExcelTransfer.Core.Cells;
using ExcelTransfer.Core.Data;
using ExcelTransfer.Core.Transfer;
using CellValue = ExcelTransfer.Core.Cells.CellValue;

namespace ExcelTransfer.Core.Workbooks.Biff12;

/// <summary><see cref="IWorkbookSource"/> for .xlsb files: BIFF12 record streams inside an OPC package.</summary>
public sealed class Biff12Source : IWorkbookSource
{
    internal const string WorksheetContentType = "application/vnd.ms-excel.worksheet";
    internal const string SharedStringsRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
    internal const string StylesRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
    internal const string TableRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";
    internal const string CalcChainRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain";

    private readonly OpcZipPackage _package;
    private readonly List<(SheetInfo Info, string PartName)> _sheets = new();
    private IReadOnlyList<string>? _sharedStrings;
    private bool[]? _dateStyles;

    public string Path => _package.Path;
    public WorkbookFormat Format => WorkbookFormat.Xlsb;
    public bool Date1904 { get; private set; }
    public bool ManualCalculation { get; private set; }
    public bool FullCalcOnLoad { get; private set; }
    public IReadOnlyList<SheetInfo> Sheets => _sheets.Select(s => s.Info).ToList();

    internal OpcZipPackage Package => _package;
    internal string WorkbookPartName { get; }
    internal string? SharedStringsPartName { get; }
    internal string? StylesPartName { get; }
    internal string? CalcChainPartName { get; }

    private Biff12Source(OpcZipPackage package)
    {
        _package = package;
        var name = System.IO.Path.GetFileName(package.Path);
        WorkbookPartName = package.MainPartName ?? throw new TransferException($"\"{name}\" has no main workbook part.", "Open it in Excel and save it again.");
        if (!WorkbookPartName.EndsWith(".bin", StringComparison.OrdinalIgnoreCase))
            throw new TransferException($"\"{name}\" is not an Excel binary workbook.", "Check the file extension; the content looks like a different Excel format.");

        var rels = package.ReadRelationships(WorkbookPartName);
        SharedStringsPartName = rels.FirstOrDefault(r => r.Type == SharedStringsRelType && !r.External)?.TargetPart;
        StylesPartName = rels.FirstOrDefault(r => r.Type == StylesRelType && !r.External)?.TargetPart;
        CalcChainPartName = rels.FirstOrDefault(r => r.Type == CalcChainRelType && !r.External)?.TargetPart;

        var wbData = package.ReadPart(WorkbookPartName);
        var index = 0;
        foreach (var rec in Biff12Stream.Parse(wbData, WorkbookPartName))
        {
            switch (rec.Type)
            {
                case Rt.WbProp:
                {
                    var r = new ByteReader(wbData, rec.DataOffset, rec.DataLength);
                    var flags = r.ReadUInt32();
                    Date1904 = (flags & 0x01) != 0;
                    break;
                }
                case Rt.CalcProp:
                {
                    var r = new ByteReader(wbData, rec.DataOffset, rec.DataLength);
                    _ = r.ReadInt32(); // recalcID
                    var calcMode = r.ReadInt32(); // 0 manual, 1 automatic, 2 automatic except tables
                    ManualCalculation = calcMode == 0;
                    if (rec.DataLength >= 26)
                    {
                        r.Skip(4 + 8 + 4);
                        var flags = r.ReadUInt16();
                        FullCalcOnLoad = (flags & 0x0001) != 0;
                    }
                    break;
                }
                case Rt.BundleSh:
                {
                    var (state, relId, sheetName) = ReadBundleSh(wbData, rec);
                    var rel = relId is null ? null : rels.FirstOrDefault(x => x.Id == relId && !x.External);
                    var partName = rel?.TargetPart ?? string.Empty;
                    var contentType = partName.Length == 0 ? null : package.ContentTypeOf(partName);
                    var isWorksheet = partName.Length > 0 && package.HasPart(partName) && (contentType is null || contentType == WorksheetContentType);
                    var tables = isWorksheet ? ReadTablesOfSheet(partName) : new List<TableInfo>();
                    _sheets.Add((new SheetInfo(sheetName, index, state != 0, !isWorksheet, tables), partName));
                    index++;
                    break;
                }
            }
        }
    }

    /// <summary>BrtBundleSh: hsState, iTabID, strRelID, strName. Pre-release Excel 2007 files carry one extra 4-byte field.</summary>
    private static (uint State, string? RelId, string Name) ReadBundleSh(byte[] data, Biff12Record rec)
    {
        try
        {
            var r = new ByteReader(data, rec.DataOffset, rec.DataLength);
            var state = r.ReadUInt32();
            _ = r.ReadUInt32();
            var relId = r.ReadXLNullableWideString();
            var name = r.ReadXLWideString();
            if (r.Remaining == 0) return (state, relId, name);
        }
        catch (TransferException)
        {
        }
        var r2 = new ByteReader(data, rec.DataOffset, rec.DataLength);
        var state2 = r2.ReadUInt32();
        r2.Skip(8);
        var relId2 = r2.ReadXLNullableWideString();
        var name2 = r2.ReadXLWideString();
        return (state2, relId2, name2);
    }

    public static Biff12Source Open(string path)
    {
        var package = OpcZipPackage.Open(path);
        try
        {
            return new Biff12Source(package);
        }
        catch
        {
            package.Dispose();
            throw;
        }
    }

    internal (SheetInfo Info, string PartName) RequireWorksheet(string sheetName)
    {
        foreach (var s in _sheets)
        {
            if (!string.Equals(s.Info.Name, sheetName, StringComparison.OrdinalIgnoreCase)) continue;
            if (s.Info.IsChartSheet)
                throw new TransferException($"Sheet \"{sheetName}\" in \"{System.IO.Path.GetFileName(Path)}\" is not a worksheet and contains no cells.", "Select a worksheet instead.");
            return s;
        }
        throw new TransferException($"Sheet \"{sheetName}\" was not found in \"{System.IO.Path.GetFileName(Path)}\".", "The workbook may have been changed since it was selected. Select the sheet again.");
    }

    internal IReadOnlyList<string> SharedStrings => _sharedStrings ??= ReadSharedStrings(_package, SharedStringsPartName);

    internal static List<string> ReadSharedStrings(OpcZipPackage package, string? partName)
    {
        var list = new List<string>();
        if (partName is null || !package.HasPart(partName)) return list;
        var data = package.ReadPart(partName);
        foreach (var rec in Biff12Stream.Parse(data, partName))
        {
            if (rec.Type != Rt.SSTItem) continue;
            list.Add(new ByteReader(data, rec.DataOffset, rec.DataLength).ReadRichStrText());
        }
        return list;
    }

    internal bool IsDateStyle(uint styleIndex)
    {
        var flags = _dateStyles ??= ReadDateStyles(_package, StylesPartName);
        return styleIndex < flags.Length && flags[styleIndex];
    }

    /// <summary>For every cell format (cellXfs entry), whether its number format is a date/time format.</summary>
    internal static bool[] ReadDateStyles(OpcZipPackage package, string? partName)
    {
        if (partName is null || !package.HasPart(partName)) return Array.Empty<bool>();
        var data = package.ReadPart(partName);
        var customFormats = new Dictionary<int, string>();
        var cellXfFormatIds = new List<int>();
        var inCellXfs = false;
        foreach (var rec in Biff12Stream.Parse(data, partName))
        {
            switch (rec.Type)
            {
                case Rt.Fmt:
                {
                    var r = new ByteReader(data, rec.DataOffset, rec.DataLength);
                    var id = r.ReadUInt16();
                    customFormats[id] = r.ReadXLWideString();
                    break;
                }
                case Rt.BeginCellXFs:
                    inCellXfs = true;
                    break;
                case Rt.EndCellXFs:
                    inCellXfs = false;
                    break;
                case Rt.XF when inCellXfs:
                {
                    var r = new ByteReader(data, rec.DataOffset, rec.DataLength);
                    _ = r.ReadUInt16(); // ixfeParent
                    cellXfFormatIds.Add(r.ReadUInt16());
                    break;
                }
            }
        }
        var flags = new bool[cellXfFormatIds.Count];
        for (var i = 0; i < flags.Length; i++)
        {
            customFormats.TryGetValue(cellXfFormatIds[i], out var code);
            flags[i] = NumberFormats.IsDateFormat(cellXfFormatIds[i], code);
        }
        return flags;
    }

    private List<TableInfo> ReadTablesOfSheet(string sheetPartName)
    {
        var tables = new List<TableInfo>();
        if (!_package.HasPart(sheetPartName)) return tables;
        var rels = _package.ReadRelationships(sheetPartName);
        var tableRels = rels.Where(r => r.Type == TableRelType && !r.External).ToList();
        if (tableRels.Count == 0) return tables;
        foreach (var rel in tableRels)
        {
            if (!_package.HasPart(rel.TargetPart)) continue;
            var info = ReadTable(_package.ReadPart(rel.TargetPart), rel.TargetPart);
            if (info is not null) tables.Add(info);
        }
        return tables;
    }

    /// <summary>Parses a table part (BrtBeginList, BrtBeginListCol...) into a <see cref="TableInfo"/>.</summary>
    internal static TableInfo? ReadTable(byte[] data, string partName)
    {
        CellRange? range = null;
        var headerRows = 1;
        var totalsRows = 0;
        var name = string.Empty;
        var displayName = string.Empty;
        var columns = new List<TableColumnInfo>();
        string? currentColumnName = null;
        var currentHasCalc = false;
        var currentHasTotals = false;
        var inColumn = false;
        foreach (var rec in Biff12Stream.Parse(data, partName))
        {
            switch (rec.Type)
            {
                case Rt.BeginList:
                {
                    var r = new ByteReader(data, rec.DataOffset, rec.DataLength);
                    range = r.ReadRfX();
                    _ = r.ReadInt32(); // lt (table type)
                    _ = r.ReadInt32(); // idList
                    headerRows = r.ReadInt32();
                    totalsRows = r.ReadInt32();
                    r.Skip(32);
                    name = r.ReadXLWideString();
                    displayName = r.ReadXLWideString();
                    break;
                }
                case Rt.BeginListCol:
                {
                    inColumn = true;
                    currentHasCalc = false;
                    currentHasTotals = false;
                    currentColumnName = null;
                    try
                    {
                        var r = new ByteReader(data, rec.DataOffset, rec.DataLength);
                        _ = r.ReadInt32(); // idField
                        var ilta = r.ReadInt32(); // totals row function
                        r.Skip(16); // nDxfHdr, nDxfInsertRow, nDxfAgg, idqsif
                        _ = r.ReadXLNullableWideString(); // stName
                        currentColumnName = r.ReadXLWideString(); // stCaption
                        currentHasTotals = ilta != 0;
                    }
                    catch (TransferException)
                    {
                        currentColumnName = null;
                    }
                    break;
                }
                case Rt.ListCCFmla when inColumn:
                    currentHasCalc = true;
                    break;
                case Rt.ListTrFmla when inColumn:
                    currentHasTotals = true;
                    break;
                case Rt.EndListCol:
                    if (inColumn)
                        columns.Add(new TableColumnInfo(currentColumnName ?? string.Empty, currentHasCalc, currentHasTotals));
                    inColumn = false;
                    break;
            }
        }
        if (range is null) return null;
        if (displayName.Length == 0) displayName = name;
        if (name.Length == 0) name = displayName;
        return new TableInfo(name, displayName, range.Value, headerRows, totalsRows, columns);
    }

    public SheetSnapshot LoadSheet(string sheetName)
    {
        var (info, partName) = RequireWorksheet(sheetName);
        var data = _package.ReadPart(partName);
        var records = Biff12Stream.Parse(data, partName);
        return BuildSnapshot(info.Name, partName, data, records, info.Tables);
    }

    internal SheetSnapshot BuildSnapshot(string sheetName, string partName, byte[] data, List<Biff12Record> records, IReadOnlyList<TableInfo> tables)
    {
        var rows = new List<RowSnapshot>();
        var merged = new List<CellRange>();
        var shared = new List<SharedFormulaGroup>();
        var arrays = new List<CellRange>();
        var dataTables = new List<CellRange>();
        var isProtected = false;

        var currentRow = 0;
        var currentHidden = false;
        var columns = new List<int>();
        var values = new List<CellValue>();
        var sharedIndex = 0u;

        void FlushRow()
        {
            if (currentRow > 0)
                rows.Add(new RowSnapshot(currentRow, currentHidden, columns.ToArray(), values.ToArray()));
            columns.Clear();
            values.Clear();
        }

        foreach (var rec in records)
        {
            var type = rec.Type;
            if (type == Rt.RowHdr)
            {
                FlushRow();
                var r = new ByteReader(data, rec.DataOffset, rec.DataLength);
                currentRow = r.ReadInt32() + 1;
                _ = r.ReadInt32(); // ixfe
                _ = r.ReadUInt16(); // miyRw
                var flags1 = r.ReadUInt16();
                currentHidden = (flags1 & 0x1000) != 0;
                continue;
            }
            if (Rt.IsCell(type))
            {
                var r = new ByteReader(data, rec.DataOffset, rec.DataLength);
                var col = r.ReadInt32() + 1;
                var styleAndFlags = r.ReadUInt32();
                var style = styleAndFlags & 0x00FFFFFF;
                var value = ReadCellValue(type, r, style);
                columns.Add(col);
                values.Add(value);
                continue;
            }
            switch (type)
            {
                case Rt.ShrFmla:
                {
                    var r = new ByteReader(data, rec.DataOffset, rec.DataLength);
                    shared.Add(new SharedFormulaGroup(sharedIndex++, r.ReadRfX()));
                    break;
                }
                case Rt.ArrFmla:
                {
                    var r = new ByteReader(data, rec.DataOffset, rec.DataLength);
                    arrays.Add(r.ReadRfX());
                    break;
                }
                case Rt.DataTable:
                {
                    var r = new ByteReader(data, rec.DataOffset, rec.DataLength);
                    dataTables.Add(r.ReadRfX());
                    break;
                }
                case Rt.MergeCell:
                {
                    var r = new ByteReader(data, rec.DataOffset, rec.DataLength);
                    merged.Add(r.ReadRfX());
                    break;
                }
                case Rt.SheetProtection:
                {
                    var r = new ByteReader(data, rec.DataOffset, rec.DataLength);
                    _ = r.ReadUInt16(); // password hash
                    isProtected |= r.ReadInt32() != 0;
                    break;
                }
                case Rt.SheetProtectionIso:
                    isProtected = true;
                    break;
                case Rt.EndSheetData:
                    FlushRow();
                    currentRow = 0;
                    break;
            }
        }
        FlushRow();
        rows.Sort((a, b) => a.Row.CompareTo(b.Row));

        return new SheetSnapshot
        {
            SheetName = sheetName,
            PartUri = "/" + partName,
            Rows = rows,
            Tables = tables,
            Protected = isProtected,
            MergedRanges = merged,
            SharedFormulas = shared,
            ArrayFormulas = arrays,
            DataTableFormulas = dataTables,
        };
    }

    private CellValue ReadCellValue(int type, ByteReader r, uint style)
    {
        var flags = CellFlags.None;
        var dateStyled = IsDateStyle(style);
        if (dateStyled) flags |= CellFlags.DateStyled;
        if (Rt.IsFormulaCell(type)) flags |= CellFlags.FromFormula;

        switch (type)
        {
            case Rt.CellBlank:
                return CellValue.BlankWithFlags(flags);
            case Rt.CellRk:
                return Numeric(r.ReadRk(), dateStyled, flags);
            case Rt.CellReal:
            case Rt.FmlaNum:
                return Numeric(r.ReadDouble(), dateStyled, flags);
            case Rt.CellBool:
            case Rt.FmlaBool:
                return CellValue.OfBoolean(r.ReadByte() != 0, flags);
            case Rt.CellError:
            case Rt.FmlaError:
                return CellValue.OfError(BErr.ToText(r.ReadByte()), flags);
            case Rt.CellSt:
            case Rt.FmlaString:
            {
                var s = r.ReadXLWideString();
                return s.Length == 0 ? CellValue.BlankWithFlags(flags) : CellValue.OfText(s, flags);
            }
            case Rt.CellRString:
            {
                var s = r.ReadRichStrText();
                return s.Length == 0 ? CellValue.BlankWithFlags(flags) : CellValue.OfText(s, flags);
            }
            case Rt.CellIsst:
            {
                var index = r.ReadInt32();
                var table = SharedStrings;
                if (index < 0 || index >= table.Count) return CellValue.BlankWithFlags(flags);
                var s = table[index];
                return s.Length == 0 ? CellValue.BlankWithFlags(flags) : CellValue.OfText(s, flags);
            }
            default:
                return CellValue.BlankWithFlags(flags);
        }
    }

    private CellValue Numeric(double number, bool dateStyled, CellFlags flags)
    {
        if (double.IsNaN(number) || double.IsInfinity(number)) return CellValue.OfError("#NUM!", flags);
        return dateStyled
            ? CellValue.OfDate(ExcelDates.ToSerial1900(number, Date1904), flags)
            : CellValue.OfNumber(number, flags);
    }

    public void Dispose() => _package.Dispose();
}
