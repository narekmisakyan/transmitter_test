using System.Buffers.Binary;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using ExcelTransfer.Core.Cells;
using ExcelTransfer.Core.Transfer;
using CellValue = ExcelTransfer.Core.Cells.CellValue;

namespace ExcelTransfer.Core.Workbooks.Biff12;

/// <summary>
/// <see cref="IWorkbookEditor"/> for .xlsb files. The destination package is read, the edited worksheet
/// stream is rewritten record by record (only the rows that change are re-encoded), the shared string table,
/// style sheet and workbook calculation properties are patched where needed, and a new package is written.
/// Every other part is copied byte for byte.
/// </summary>
internal sealed class Biff12Editor : IWorkbookEditor
{
    private const string BinaryIndexRelType = "http://schemas.microsoft.com/office/2006/relationships/xlBinaryIndex";
    private const string SharedStringsContentType = "application/vnd.ms-excel.sharedStrings";
    private const ushort BuiltInShortDate = 14;
    private const ushort BuiltInDateTime = 22;

    private readonly string _workingCopyPath;
    private readonly Biff12Source _source;
    private readonly OpcZipPackage _package;
    private readonly Dictionary<string, byte[]> _replacedParts = new(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<string> _removedParts = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, byte[]> _newParts = new(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<string> _changedParts = new(StringComparer.OrdinalIgnoreCase);

    // Shared strings.
    private Dictionary<string, int>? _sstIndex;
    private int _sstOriginalCount;
    private readonly List<string> _sstNew = new();
    private int _sstRefsAdded;

    // Styles.
    private bool[]? _dateStyles;
    private List<byte[]>? _cellXfPayloads;
    private readonly List<byte[]> _newXfPayloads = new();
    private readonly Dictionary<(uint, bool), uint> _dateStyleCache = new();

    private bool _finished;

    private Biff12Editor(Biff12Source source, string workingCopyPath)
    {
        _source = source;
        _package = source.Package;
        _workingCopyPath = workingCopyPath;
    }

    public WorkbookFormat Format => WorkbookFormat.Xlsb;

    public static Biff12Editor Create(string sourcePath, string workingCopyPath)
    {
        var source = Biff12Source.Open(sourcePath);
        return new Biff12Editor(source, workingCopyPath);
    }

    // ---------------------------------------------------------------- sheet editing

    private sealed class CellEntry
    {
        public int Column;
        public int Type;
        public uint Style;
        public byte[] Record = Array.Empty<byte>();
        /// <summary>Records that precede the cell and belong to it (BrtCellMeta).</summary>
        public List<byte[]> Before = new();
        /// <summary>Records that follow the cell and belong to it (BrtShrFmla, BrtArrFmla, BrtTable, ...).</summary>
        public List<byte[]> After = new();
        public bool HasContent => Type != Rt.CellBlank;
    }

    private const int BrtCellMeta = 0x0031;
    private const int BrtValueMeta = 0x0032;

    public SheetEditStats EditSheet(string sheetName, IReadOnlyCollection<(int Row, int Column)> cellsToClear, IReadOnlyList<CellWrite> writes, IProgress<ProgressReport>? progress, CancellationToken ct)
    {
        var (_, partName) = _source.RequireWorksheet(sheetName);
        var data = _package.ReadPart(partName);
        var records = Biff12Stream.Parse(data, partName);

        var clearByRow = new Dictionary<int, HashSet<int>>();
        foreach (var (r, c) in cellsToClear)
        {
            if (!clearByRow.TryGetValue(r, out var set)) clearByRow[r] = set = new HashSet<int>();
            set.Add(c);
        }
        var writeByRow = new Dictionary<int, Dictionary<int, CellValue>>();
        foreach (var w in writes)
        {
            if (!writeByRow.TryGetValue(w.Row, out var map)) writeByRow[w.Row] = map = new Dictionary<int, CellValue>();
            map[w.Column] = w.Value;
        }
        var editRows = new SortedSet<int>(clearByRow.Keys);
        editRows.UnionWith(writeByRow.Keys);
        var pendingWriteRows = new SortedSet<int>(writeByRow.Keys);

        var columnStyles = ReadColumnStyles(data, records);
        var output = new MemoryStream(data.Length + 4096);
        var cleared = 0;
        var written = 0;
        var dateStyles = 0;
        int maxRow = 0, maxCol = 0;
        long dimPayloadOffset = -1;
        var total = Math.Max(1, editRows.Count);
        var processed = 0;

        var i = 0;
        // Copy everything up to and including BrtBeginSheetData.
        while (i < records.Count)
        {
            var rec = records[i];
            if (rec.Type == Rt.WsDim) dimPayloadOffset = output.Length + (rec.DataOffset - rec.HeaderOffset);
            output.Write(data, rec.HeaderOffset, rec.End - rec.HeaderOffset);
            i++;
            if (rec.Type == Rt.BeginSheetData) break;
        }

        void EmitSynthesizedRowsBefore(int boundaryRow)
        {
            while (pendingWriteRows.Count > 0 && pendingWriteRows.Min < boundaryRow)
            {
                var r = pendingWriteRows.Min;
                pendingWriteRows.Remove(r);
                var entries = new List<CellEntry>();
                ApplyWrites(entries, r, writeByRow[r], 0, false, columnStyles, ref written, ref dateStyles, ref maxRow, ref maxCol);
                if (entries.Count == 0) continue;
                WriteRow(output, r, null, entries);
            }
        }

        // Row blocks.
        while (i < records.Count)
        {
            ct.ThrowIfCancellationRequested();
            var rec = records[i];
            if (rec.Type == Rt.EndSheetData)
            {
                EmitSynthesizedRowsBefore(int.MaxValue);
                break;
            }
            if (rec.Type != Rt.RowHdr)
            {
                // Unexpected record outside a row block: copy it.
                output.Write(data, rec.HeaderOffset, rec.End - rec.HeaderOffset);
                i++;
                continue;
            }

            var rowNumber = BinaryPrimitives.ReadInt32LittleEndian(data.AsSpan(rec.DataOffset)) + 1;
            var blockStart = i;
            var j = i + 1;
            while (j < records.Count && records[j].Type != Rt.RowHdr && records[j].Type != Rt.EndSheetData) j++;
            var blockEnd = j; // exclusive

            EmitSynthesizedRowsBefore(rowNumber);
            pendingWriteRows.Remove(rowNumber);

            if (!editRows.Contains(rowNumber))
            {
                output.Write(data, rec.HeaderOffset, records[blockEnd - 1].End - rec.HeaderOffset);
                i = blockEnd;
                continue;
            }

            // Parse the row block into cell entries.
            var header = data.AsSpan(rec.DataOffset, rec.DataLength).ToArray();
            var entries = new List<CellEntry>();
            var preRecords = new List<byte[]>();
            var pendingBefore = new List<byte[]>();
            for (var k = blockStart + 1; k < blockEnd; k++)
            {
                var r2 = records[k];
                var bytes = data.AsSpan(r2.HeaderOffset, r2.End - r2.HeaderOffset).ToArray();
                if (Rt.IsCell(r2.Type))
                {
                    var entry = new CellEntry
                    {
                        Column = BinaryPrimitives.ReadInt32LittleEndian(data.AsSpan(r2.DataOffset)) + 1,
                        Type = r2.Type,
                        Style = BinaryPrimitives.ReadUInt32LittleEndian(data.AsSpan(r2.DataOffset + 4)) & 0x00FFFFFF,
                        Record = bytes,
                    };
                    entry.Before.AddRange(pendingBefore);
                    pendingBefore.Clear();
                    entries.Add(entry);
                }
                else if (r2.Type == BrtCellMeta || r2.Type == BrtValueMeta)
                {
                    pendingBefore.Add(bytes);
                }
                else if (entries.Count > 0)
                {
                    entries[^1].After.Add(bytes);
                }
                else
                {
                    preRecords.Add(bytes);
                }
            }
            if (pendingBefore.Count > 0) preRecords.AddRange(pendingBefore);

            var flags1 = BinaryPrimitives.ReadUInt16LittleEndian(header.AsSpan(10));
            var rowStyle = (flags1 & 0x4000) != 0 ? BinaryPrimitives.ReadUInt32LittleEndian(header.AsSpan(4)) & 0x00FFFFFF : 0u;

            if (clearByRow.TryGetValue(rowNumber, out var clears))
            {
                foreach (var col in clears)
                {
                    var idx = entries.FindIndex(e => e.Column == col);
                    if (idx < 0) continue;
                    var entry = entries[idx];
                    if (entry.HasContent) cleared++;
                    if (entry.Style != 0)
                        entries[idx] = new CellEntry { Column = col, Type = Rt.CellBlank, Style = entry.Style, Record = BuildCellRecord(Rt.CellBlank, col, entry.Style, ReadOnlySpan<byte>.Empty) };
                    else
                        entries.RemoveAt(idx);
                }
            }
            if (writeByRow.TryGetValue(rowNumber, out var rowWrites))
                ApplyWrites(entries, rowNumber, rowWrites, rowStyle, true, columnStyles, ref written, ref dateStyles, ref maxRow, ref maxCol);

            WriteRow(output, rowNumber, header, entries, preRecords);
            i = blockEnd;
            if (++processed % 2000 == 0) progress?.Report(new ProgressReport("Rewriting rows", 0.1 + 0.5 * processed / total));
        }

        // Copy BrtEndSheetData and everything after it.
        while (i < records.Count)
        {
            var rec = records[i];
            output.Write(data, rec.HeaderOffset, rec.End - rec.HeaderOffset);
            i++;
        }

        var outBytes = output.ToArray();
        if (dimPayloadOffset >= 0 && maxRow > 0)
        {
            var span = outBytes.AsSpan((int)dimPayloadOffset, 16);
            var rwLast = BinaryPrimitives.ReadInt32LittleEndian(span.Slice(4));
            var colLast = BinaryPrimitives.ReadInt32LittleEndian(span.Slice(12));
            if (maxRow - 1 > rwLast) BinaryPrimitives.WriteInt32LittleEndian(span.Slice(4), maxRow - 1);
            if (maxCol - 1 > colLast) BinaryPrimitives.WriteInt32LittleEndian(span.Slice(12), maxCol - 1);
        }
        _replacedParts[partName] = outBytes;
        _changedParts.Add("/" + partName);
        RemoveBinaryIndex(partName);
        return new SheetEditStats(cleared, written, dateStyles);
    }

    private void ApplyWrites(List<CellEntry> entries, int rowNumber, Dictionary<int, CellValue> rowWrites, uint rowStyle, bool rowStyleKnown,
        List<(int First, int Last, uint Style)> columnStyles, ref int written, ref int dateStyles, ref int maxRow, ref int maxCol)
    {
        foreach (var (col, value) in rowWrites)
        {
            var idx = entries.FindIndex(e => e.Column == col);
            if (value.IsBlank)
            {
                if (idx >= 0)
                {
                    var existing = entries[idx];
                    if (existing.Style != 0)
                        entries[idx] = new CellEntry { Column = col, Type = Rt.CellBlank, Style = existing.Style, Record = BuildCellRecord(Rt.CellBlank, col, existing.Style, ReadOnlySpan<byte>.Empty) };
                    else
                        entries.RemoveAt(idx);
                }
                continue;
            }
            uint style;
            if (idx >= 0) style = entries[idx].Style;
            else if (rowStyleKnown && rowStyle != 0) style = rowStyle;
            else style = ColumnStyle(columnStyles, col);

            byte[] record;
            switch (value.Kind)
            {
                case CellKind.Text:
                {
                    var payload = new ByteWriter().Int32(GetSharedStringIndex(value.TextOrEmpty)).ToArray();
                    record = BuildCellRecord(Rt.CellIsst, col, style, payload);
                    break;
                }
                case CellKind.Number:
                    record = BuildCellRecord(Rt.CellReal, col, style, new ByteWriter().Double(value.Number).ToArray());
                    break;
                case CellKind.Date:
                {
                    var serial = ExcelDates.FromSerial1900(value.Number, _source.Date1904);
                    if (!IsDateStyle(style))
                    {
                        style = GetOrCreateDateStyle(style, Math.Abs(serial - Math.Floor(serial)) > 1e-9);
                        dateStyles++;
                    }
                    record = BuildCellRecord(Rt.CellReal, col, style, new ByteWriter().Double(serial).ToArray());
                    break;
                }
                case CellKind.Boolean:
                    record = BuildCellRecord(Rt.CellBool, col, style, new[] { value.BooleanValue ? (byte)1 : (byte)0 });
                    break;
                case CellKind.Error:
                    record = BuildCellRecord(Rt.CellError, col, style, new[] { BErr.FromText(value.TextOrEmpty) });
                    break;
                default:
                    continue;
            }
            var entry = new CellEntry { Column = col, Type = Rt.CellReal, Style = style, Record = record };
            if (idx >= 0) entries[idx] = entry;
            else entries.Add(entry);
            written++;
            if (rowNumber > maxRow) maxRow = rowNumber;
            if (col > maxCol) maxCol = col;
        }
    }

    private static uint ColumnStyle(List<(int First, int Last, uint Style)> columnStyles, int col)
    {
        foreach (var (first, last, style) in columnStyles)
            if (col >= first && col <= last) return style;
        return 0;
    }

    private static List<(int First, int Last, uint Style)> ReadColumnStyles(byte[] data, List<Biff12Record> records)
    {
        var list = new List<(int, int, uint)>();
        foreach (var rec in records)
        {
            if (rec.Type == Rt.BeginSheetData) break;
            if (rec.Type != Rt.ColInfo || rec.DataLength < 16) continue;
            var r = new ByteReader(data, rec.DataOffset, rec.DataLength);
            var first = r.ReadInt32() + 1;
            var last = r.ReadInt32() + 1;
            _ = r.ReadInt32(); // width
            var style = r.ReadUInt32() & 0x00FFFFFF;
            if (style != 0) list.Add((first, last, style));
        }
        return list;
    }

    private static byte[] BuildCellRecord(int type, int column, uint style, ReadOnlySpan<byte> payload)
    {
        var body = new ByteWriter().CellHeader(column, style).Bytes(payload).ToArray();
        using var ms = new MemoryStream(body.Length + 4);
        Biff12Stream.WriteRecord(ms, type, body);
        return ms.ToArray();
    }

    /// <summary>Writes a row header (original header bytes when available) with recomputed column spans, then the cells.</summary>
    private static void WriteRow(Stream output, int rowNumber, byte[]? originalHeader, List<CellEntry> entries, List<byte[]>? preRecords = null)
    {
        entries.Sort((a, b) => a.Column.CompareTo(b.Column));
        var header = new ByteWriter();
        if (originalHeader is not null && originalHeader.Length >= 17)
        {
            header.Bytes(originalHeader.AsSpan(0, 13));
        }
        else
        {
            header.Int32(rowNumber - 1); // rw
            header.Int32(0);             // ixfe
            header.UInt16(300);          // miyRw (15 pt), not custom
            header.UInt16(0);            // flags
            header.Byte(0);              // phonetic
        }
        // Column spans: one (colMic, colLast) pair per 1024-column block that contains cells.
        var spans = new List<(int First, int Last)>();
        foreach (var e in entries)
        {
            var col0 = e.Column - 1;
            var block = col0 >> 10;
            if (spans.Count > 0 && (spans[^1].First >> 10) == block)
                spans[^1] = (spans[^1].First, col0);
            else
                spans.Add((col0, col0));
        }
        header.Int32(spans.Count);
        foreach (var (first, last) in spans)
        {
            header.Int32(first);
            header.Int32(last);
        }
        Biff12Stream.WriteRecord(output, Rt.RowHdr, header.ToArray());
        if (preRecords is not null)
            foreach (var p in preRecords) output.Write(p);
        foreach (var e in entries)
        {
            foreach (var b in e.Before) output.Write(b);
            output.Write(e.Record);
            foreach (var a in e.After) output.Write(a);
        }
    }

    // ---------------------------------------------------------------- shared strings

    private int GetSharedStringIndex(string text)
    {
        if (_sstIndex is null)
        {
            _sstIndex = new Dictionary<string, int>(StringComparer.Ordinal);
            var existing = _source.SharedStrings;
            for (var i = 0; i < existing.Count; i++) _sstIndex.TryAdd(existing[i], i);
            _sstOriginalCount = existing.Count;
        }
        _sstRefsAdded++;
        if (_sstIndex.TryGetValue(text, out var index)) return index;
        index = _sstOriginalCount + _sstNew.Count;
        _sstNew.Add(text);
        _sstIndex[text] = index;
        return index;
    }

    private void WriteSharedStrings()
    {
        if (_sstNew.Count == 0 && _sstRefsAdded == 0) return;
        var partName = _source.SharedStringsPartName;
        var output = new MemoryStream();
        if (partName is not null && _package.HasPart(partName))
        {
            var data = _package.ReadPart(partName);
            var records = Biff12Stream.Parse(data, partName);
            foreach (var rec in records)
            {
                if (rec.Type == Rt.BeginSst && rec.DataLength >= 8)
                {
                    var total = BinaryPrimitives.ReadInt32LittleEndian(data.AsSpan(rec.DataOffset));
                    var unique = BinaryPrimitives.ReadInt32LittleEndian(data.AsSpan(rec.DataOffset + 4));
                    var payload = new ByteWriter().Int32(total + _sstRefsAdded).Int32(unique + _sstNew.Count).Bytes(data.AsSpan(rec.DataOffset + 8, rec.DataLength - 8)).ToArray();
                    Biff12Stream.WriteRecord(output, Rt.BeginSst, payload);
                    continue;
                }
                if (rec.Type == Rt.EndSst)
                {
                    foreach (var s in _sstNew)
                        Biff12Stream.WriteRecord(output, Rt.SSTItem, new ByteWriter().RichStr(s).ToArray());
                }
                output.Write(data, rec.HeaderOffset, rec.End - rec.HeaderOffset);
            }
            _replacedParts[partName] = output.ToArray();
            _changedParts.Add("/" + partName);
            return;
        }

        // No shared string part yet: create one and wire it into the package.
        partName = "xl/sharedStrings.bin";
        Biff12Stream.WriteRecord(output, Rt.BeginSst, new ByteWriter().Int32(_sstRefsAdded).Int32(_sstNew.Count).ToArray());
        foreach (var s in _sstNew)
            Biff12Stream.WriteRecord(output, Rt.SSTItem, new ByteWriter().RichStr(s).ToArray());
        Biff12Stream.WriteRecord(output, Rt.EndSst, ReadOnlySpan<byte>.Empty);
        _newParts[partName] = output.ToArray();
        _changedParts.Add("/" + partName);
        AddRelationship(_source.WorkbookPartName, Biff12Source.SharedStringsRelType, "sharedStrings.bin");
        AddContentTypeOverride("/" + partName, SharedStringsContentType);
    }

    // ---------------------------------------------------------------- styles

    private void EnsureStylesLoaded()
    {
        if (_dateStyles is not null) return;
        _dateStyles = Biff12Source.ReadDateStyles(_package, _source.StylesPartName);
        _cellXfPayloads = new List<byte[]>();
        var partName = _source.StylesPartName;
        if (partName is null || !_package.HasPart(partName)) return;
        var data = _package.ReadPart(partName);
        var inCellXfs = false;
        foreach (var rec in Biff12Stream.Parse(data, partName))
        {
            if (rec.Type == Rt.BeginCellXFs) inCellXfs = true;
            else if (rec.Type == Rt.EndCellXFs) inCellXfs = false;
            else if (rec.Type == Rt.XF && inCellXfs) _cellXfPayloads.Add(data.AsSpan(rec.DataOffset, rec.DataLength).ToArray());
        }
    }

    private bool IsDateStyle(uint style)
    {
        EnsureStylesLoaded();
        var count = _dateStyles!.Length + _newXfPayloads.Count;
        if (style >= count) return false;
        if (style < _dateStyles.Length) return _dateStyles[style];
        return true; // every XF this editor adds is a date format
    }

    private uint GetOrCreateDateStyle(uint baseStyle, bool withTime)
    {
        EnsureStylesLoaded();
        if (_dateStyleCache.TryGetValue((baseStyle, withTime), out var cached)) return cached;
        if (_source.StylesPartName is null)
            throw new TransferException("The destination workbook has no style sheet part, so a date format cannot be added. Open and save it in Excel first.");
        byte[] payload;
        if (baseStyle < _cellXfPayloads!.Count && _cellXfPayloads[(int)baseStyle].Length >= 16)
            payload = (byte[])_cellXfPayloads[(int)baseStyle].Clone();
        else
        {
            payload = new byte[16];
        }
        BinaryPrimitives.WriteUInt16LittleEndian(payload.AsSpan(2), withTime ? BuiltInDateTime : BuiltInShortDate);
        var used = BinaryPrimitives.ReadUInt16LittleEndian(payload.AsSpan(14));
        BinaryPrimitives.WriteUInt16LittleEndian(payload.AsSpan(14), (ushort)(used | 0x0001)); // number format applied
        var index = (uint)(_dateStyles!.Length + _newXfPayloads.Count);
        _newXfPayloads.Add(payload);
        _dateStyleCache[(baseStyle, withTime)] = index;
        return index;
    }

    private void WriteStyles()
    {
        if (_newXfPayloads.Count == 0) return;
        var partName = _source.StylesPartName!;
        var data = _package.ReadPart(partName);
        var output = new MemoryStream(data.Length + 64);
        foreach (var rec in Biff12Stream.Parse(data, partName))
        {
            if (rec.Type == Rt.BeginCellXFs && rec.DataLength >= 4)
            {
                var count = BinaryPrimitives.ReadInt32LittleEndian(data.AsSpan(rec.DataOffset));
                var payload = new ByteWriter().Int32(count + _newXfPayloads.Count).Bytes(data.AsSpan(rec.DataOffset + 4, rec.DataLength - 4)).ToArray();
                Biff12Stream.WriteRecord(output, Rt.BeginCellXFs, payload);
                continue;
            }
            if (rec.Type == Rt.EndCellXFs)
            {
                foreach (var xf in _newXfPayloads)
                    Biff12Stream.WriteRecord(output, Rt.XF, xf);
            }
            output.Write(data, rec.HeaderOffset, rec.End - rec.HeaderOffset);
        }
        _replacedParts[partName] = output.ToArray();
        _changedParts.Add("/" + partName);
    }

    // ---------------------------------------------------------------- workbook part, calc chain, index parts

    private void WriteWorkbookPart()
    {
        var partName = _source.WorkbookPartName;
        var data = _package.ReadPart(partName);
        var output = new MemoryStream(data.Length);
        var found = false;
        foreach (var rec in Biff12Stream.Parse(data, partName))
        {
            if (rec.Type == Rt.CalcProp && rec.DataLength >= 25)
            {
                found = true;
                var payload = data.AsSpan(rec.DataOffset, rec.DataLength).ToArray();
                // recalcID = 0 forces a full recalculation in every Excel version; flag bit 0 = fFullCalcOnLoad.
                BinaryPrimitives.WriteInt32LittleEndian(payload.AsSpan(0), 0);
                payload[24] |= 0x01;
                Biff12Stream.WriteRecord(output, Rt.CalcProp, payload);
                continue;
            }
            output.Write(data, rec.HeaderOffset, rec.End - rec.HeaderOffset);
        }
        if (!found) return;
        _replacedParts[partName] = output.ToArray();
        _changedParts.Add("/" + partName);
    }

    private void RemoveCalcChain()
    {
        var calcChain = _source.CalcChainPartName;
        if (calcChain is null) return;
        RemovePartWithRelationship(_source.WorkbookPartName, calcChain);
    }

    private void RemoveBinaryIndex(string sheetPartName)
    {
        foreach (var rel in _package.ReadRelationships(sheetPartName))
        {
            if (rel.External || !rel.Type.EndsWith("/xlBinaryIndex", StringComparison.OrdinalIgnoreCase)) continue;
            RemovePartWithRelationship(sheetPartName, rel.TargetPart);
        }
    }

    private void RemovePartWithRelationship(string ownerPartName, string targetPartName)
    {
        if (_package.HasPart(targetPartName))
        {
            _removedParts.Add(targetPartName);
            _changedParts.Add("/" + targetPartName);
        }
        var relsName = OpcZipPackage.RelationshipsPartName(ownerPartName);
        if (!_package.HasPart(relsName)) return;
        var doc = LoadXml(relsName);
        XNamespace ns = OpcZipPackage.RelationshipsNamespace;
        var baseDir = OpcZipPackage.GetDirectory(ownerPartName);
        var removed = false;
        foreach (var rel in doc.Root!.Elements(ns + "Relationship").ToList())
        {
            var target = (string?)rel.Attribute("Target") ?? string.Empty;
            if (string.Equals((string?)rel.Attribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase)) continue;
            if (string.Equals(OpcZipPackage.ResolveTarget(baseDir, target), targetPartName, StringComparison.OrdinalIgnoreCase))
            {
                rel.Remove();
                removed = true;
            }
        }
        if (removed) StoreXml(relsName, doc);
        RemoveContentTypeOverride("/" + targetPartName);
    }

    private void AddRelationship(string ownerPartName, string type, string target)
    {
        var relsName = OpcZipPackage.RelationshipsPartName(ownerPartName);
        XNamespace ns = OpcZipPackage.RelationshipsNamespace;
        var doc = _package.HasPart(relsName) || _replacedParts.ContainsKey(relsName)
            ? LoadXml(relsName)
            : new XDocument(new XElement(ns + "Relationships"));
        var existingIds = doc.Root!.Elements(ns + "Relationship").Select(e => (string?)e.Attribute("Id")).Where(x => x is not null).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var n = 1;
        string id;
        do { id = "rId" + n++; } while (existingIds.Contains(id));
        doc.Root.Add(new XElement(ns + "Relationship", new XAttribute("Id", id), new XAttribute("Type", type), new XAttribute("Target", target)));
        StoreXml(relsName, doc);
    }

    private void AddContentTypeOverride(string partName, string contentType)
    {
        const string ctName = "[Content_Types].xml";
        var doc = LoadXml(ctName);
        XNamespace ns = OpcZipPackage.ContentTypesNamespace;
        if (!doc.Root!.Elements(ns + "Override").Any(o => string.Equals((string?)o.Attribute("PartName"), partName, StringComparison.OrdinalIgnoreCase)))
            doc.Root.Add(new XElement(ns + "Override", new XAttribute("PartName", partName), new XAttribute("ContentType", contentType)));
        StoreXml(ctName, doc);
    }

    private void RemoveContentTypeOverride(string partName)
    {
        const string ctName = "[Content_Types].xml";
        if (!_package.HasPart(ctName)) return;
        var doc = LoadXml(ctName);
        XNamespace ns = OpcZipPackage.ContentTypesNamespace;
        var removed = false;
        foreach (var o in doc.Root!.Elements(ns + "Override").ToList())
        {
            if (string.Equals((string?)o.Attribute("PartName"), partName, StringComparison.OrdinalIgnoreCase))
            {
                o.Remove();
                removed = true;
            }
        }
        if (removed) StoreXml(ctName, doc);
    }

    private XDocument LoadXml(string partName)
    {
        var bytes = _replacedParts.TryGetValue(partName, out var replaced) ? replaced : _package.ReadPart(partName);
        using var ms = new MemoryStream(bytes);
        return XDocument.Load(ms);
    }

    private void StoreXml(string partName, XDocument doc)
    {
        using var ms = new MemoryStream();
        using (var writer = new StreamWriter(ms, new UTF8Encoding(false), leaveOpen: true))
        {
            doc.Declaration ??= new XDeclaration("1.0", "UTF-8", "yes");
            doc.Save(writer, SaveOptions.DisableFormatting);
        }
        _replacedParts[partName] = ms.ToArray();
        _changedParts.Add("/" + partName);
    }

    // ---------------------------------------------------------------- finish

    public IReadOnlySet<string> Finish()
    {
        if (_finished) throw new InvalidOperationException("Finish was already called.");
        _finished = true;
        WriteSharedStrings();
        WriteStyles();
        WriteWorkbookPart();
        RemoveCalcChain();

        try
        {
            using var fs = new FileStream(_workingCopyPath, FileMode.Create, FileAccess.ReadWrite, FileShare.None);
            using var zip = new ZipArchive(fs, ZipArchiveMode.Create, leaveOpen: false);
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var entry in _package.Entries)
            {
                var name = OpcZipPackage.Normalize(entry.FullName);
                if (_removedParts.Contains(name)) continue;
                if (!seen.Add(name)) continue;
                var bytes = _replacedParts.TryGetValue(name, out var replaced) ? replaced : _package.ReadEntry(entry);
                var newEntry = zip.CreateEntry(name, CompressionLevel.Optimal);
                newEntry.LastWriteTime = entry.LastWriteTime;
                using var s = newEntry.Open();
                s.Write(bytes);
            }
            foreach (var (name, bytes) in _newParts)
            {
                if (!seen.Add(name)) continue;
                var newEntry = zip.CreateEntry(name, CompressionLevel.Optimal);
                using var s = newEntry.Open();
                s.Write(bytes);
            }
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            throw new TransferException($"The output workbook could not be written: {ex.Message}", "Check that the output folder is writable.", ex);
        }
        _changedParts.Add("/[Content_Types].xml");
        _changedParts.Add("/" + OpcZipPackage.RelationshipsPartName(_source.WorkbookPartName));
        return _changedParts;
    }

    public void Dispose() => _source.Dispose();
}
