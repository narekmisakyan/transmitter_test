using System.Buffers.Binary;
using System.Text;
using ExcelTransfer.Core.Transfer;

namespace ExcelTransfer.Core.Workbooks.Biff12;

/// <summary>One record in a BIFF12 stream: a variable-length type, a variable-length size, then the payload.</summary>
internal readonly record struct Biff12Record(int Type, int HeaderOffset, int DataOffset, int DataLength)
{
    public int End => DataOffset + DataLength;
    public ReadOnlySpan<byte> Payload(byte[] data) => data.AsSpan(DataOffset, DataLength);
}

/// <summary>Framing of BIFF12 record streams ([MS-XLSB] 2.1.4).</summary>
internal static class Biff12Stream
{
    public static List<Biff12Record> Parse(byte[] data, string partName)
    {
        var records = new List<Biff12Record>(Math.Max(16, data.Length / 24));
        var pos = 0;
        while (pos < data.Length)
        {
            var start = pos;
            if (!TryReadHeader(data, ref pos, out var type, out var length) || pos + length > data.Length)
                throw new TransferException($"The binary part \"{partName}\" is damaged (bad record at offset {start}).", "Open the workbook in Excel and save it again.");
            records.Add(new Biff12Record(type, start, pos, length));
            pos += length;
        }
        return records;
    }

    public static bool TryReadHeader(byte[] data, ref int pos, out int type, out int length)
    {
        type = 0;
        length = 0;
        if (pos >= data.Length) return false;
        var b0 = data[pos++];
        type = b0 & 0x7F;
        if ((b0 & 0x80) != 0)
        {
            if (pos >= data.Length) return false;
            var b1 = data[pos++];
            type |= (b1 & 0x7F) << 7;
        }
        var shift = 0;
        for (var i = 0; i < 4; i++)
        {
            if (pos >= data.Length) return false;
            var b = data[pos++];
            length |= (b & 0x7F) << shift;
            if ((b & 0x80) == 0) return true;
            shift += 7;
        }
        return false;
    }

    public static void WriteRecord(Stream output, int type, ReadOnlySpan<byte> payload)
    {
        WriteHeader(output, type, payload.Length);
        output.Write(payload);
    }

    public static void WriteHeader(Stream output, int type, int length)
    {
        if (type < 0x80)
        {
            output.WriteByte((byte)type);
        }
        else
        {
            output.WriteByte((byte)((type & 0x7F) | 0x80));
            output.WriteByte((byte)((type >> 7) & 0x7F));
        }
        var len = (uint)length;
        do
        {
            var b = (byte)(len & 0x7F);
            len >>= 7;
            if (len != 0) b |= 0x80;
            output.WriteByte(b);
        } while (len != 0);
    }
}

/// <summary>Sequential reader for BIFF12 record payloads. All integers are little-endian.</summary>
internal sealed class ByteReader
{
    private readonly byte[] _data;
    private readonly int _end;

    public ByteReader(byte[] data, int offset, int length)
    {
        _data = data;
        Position = offset;
        _end = offset + length;
    }

    public ByteReader(byte[] data) : this(data, 0, data.Length) { }

    public int Position { get; set; }
    public int Remaining => _end - Position;

    private void Need(int count)
    {
        if (Position + count > _end) throw new TransferException("A binary workbook record is shorter than expected (the file may be damaged).", "Open the workbook in Excel and save it again.");
    }

    public byte ReadByte() { Need(1); return _data[Position++]; }
    public ushort ReadUInt16() { Need(2); var v = BinaryPrimitives.ReadUInt16LittleEndian(_data.AsSpan(Position)); Position += 2; return v; }
    public int ReadInt32() { Need(4); var v = BinaryPrimitives.ReadInt32LittleEndian(_data.AsSpan(Position)); Position += 4; return v; }
    public uint ReadUInt32() { Need(4); var v = BinaryPrimitives.ReadUInt32LittleEndian(_data.AsSpan(Position)); Position += 4; return v; }
    public double ReadDouble() { Need(8); var v = BinaryPrimitives.ReadDoubleLittleEndian(_data.AsSpan(Position)); Position += 8; return v; }
    public void Skip(int count) { Need(count); Position += count; }

    /// <summary>XLWideString: 4-byte character count followed by UTF-16LE text.</summary>
    public string ReadXLWideString()
    {
        var cch = ReadInt32();
        if (cch < 0) throw new TransferException("A binary workbook string has an invalid length (the file may be damaged).");
        Need(cch * 2);
        var s = Encoding.Unicode.GetString(_data, Position, cch * 2);
        Position += cch * 2;
        return s;
    }

    /// <summary>XLNullableWideString: like XLWideString, but a count of 0xFFFFFFFF means null.</summary>
    public string? ReadXLNullableWideString()
    {
        var cch = ReadUInt32();
        if (cch == 0xFFFFFFFF) return null;
        if (cch > int.MaxValue / 2) throw new TransferException("A binary workbook string has an invalid length (the file may be damaged).");
        Need((int)cch * 2);
        var s = Encoding.Unicode.GetString(_data, Position, (int)cch * 2);
        Position += (int)cch * 2;
        return s;
    }

    /// <summary>RichStr: flags byte, base string, then optional formatting runs and phonetic data (ignored).</summary>
    public string ReadRichStrText()
    {
        _ = ReadByte();
        return ReadXLWideString();
    }

    /// <summary>RfX / UncheckedRfX: rwFirst, rwLast, colFirst, colLast (all 0-based, 4 bytes each). Returned 1-based.</summary>
    public Cells.CellRange ReadRfX()
    {
        var rwFirst = ReadInt32();
        var rwLast = ReadInt32();
        var colFirst = ReadInt32();
        var colLast = ReadInt32();
        return new Cells.CellRange(rwFirst + 1, colFirst + 1, rwLast + 1, colLast + 1);
    }

    /// <summary>RkNumber: a 30-bit integer or truncated double, optionally divided by 100.</summary>
    public double ReadRk()
    {
        var rk = ReadUInt32();
        return DecodeRk(rk);
    }

    public static double DecodeRk(uint rk)
    {
        var x100 = (rk & 1) != 0;
        var isInt = (rk & 2) != 0;
        double value;
        if (isInt)
        {
            value = (int)rk >> 2;
        }
        else
        {
            var bits = (ulong)(rk & 0xFFFFFFFCu) << 32;
            value = BitConverter.Int64BitsToDouble((long)bits);
        }
        return x100 ? value / 100 : value;
    }
}

/// <summary>Builds BIFF12 record payloads.</summary>
internal sealed class ByteWriter
{
    private readonly MemoryStream _stream = new();

    public int Length => (int)_stream.Length;

    public ByteWriter Byte(byte value) { _stream.WriteByte(value); return this; }
    public ByteWriter UInt16(ushort value) { Span<byte> b = stackalloc byte[2]; BinaryPrimitives.WriteUInt16LittleEndian(b, value); _stream.Write(b); return this; }
    public ByteWriter Int32(int value) { Span<byte> b = stackalloc byte[4]; BinaryPrimitives.WriteInt32LittleEndian(b, value); _stream.Write(b); return this; }
    public ByteWriter UInt32(uint value) { Span<byte> b = stackalloc byte[4]; BinaryPrimitives.WriteUInt32LittleEndian(b, value); _stream.Write(b); return this; }
    public ByteWriter Double(double value) { Span<byte> b = stackalloc byte[8]; BinaryPrimitives.WriteDoubleLittleEndian(b, value); _stream.Write(b); return this; }
    public ByteWriter Bytes(ReadOnlySpan<byte> bytes) { _stream.Write(bytes); return this; }

    public ByteWriter XLWideString(string value)
    {
        Int32(value.Length);
        if (value.Length > 0) _stream.Write(Encoding.Unicode.GetBytes(value));
        return this;
    }

    public ByteWriter XLNullableWideString(string? value)
    {
        if (value is null) return UInt32(0xFFFFFFFF);
        return XLWideString(value);
    }

    /// <summary>A plain (unformatted) RichStr.</summary>
    public ByteWriter RichStr(string value) => Byte(0).XLWideString(value);

    /// <summary>Cell header: 0-based column, 24-bit style index, no phonetic flag.</summary>
    public ByteWriter CellHeader(int column1Based, uint styleIndex)
    {
        Int32(column1Based - 1);
        UInt32(styleIndex & 0x00FFFFFF);
        return this;
    }

    public byte[] ToArray() => _stream.ToArray();
}
