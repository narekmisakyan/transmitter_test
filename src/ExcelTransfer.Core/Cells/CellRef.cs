using System.Text;

namespace ExcelTransfer.Core.Cells;

/// <summary>A1-style reference helpers. Rows and columns are 1-based.</summary>
public static class CellRef
{
    public const int MaxColumns = 16384;
    public const int MaxRows = 1048576;

    public static string ColumnLetter(int columnIndex)
    {
        if (columnIndex < 1) throw new ArgumentOutOfRangeException(nameof(columnIndex));
        var sb = new StringBuilder();
        var n = columnIndex;
        while (n > 0)
        {
            n--;
            sb.Insert(0, (char)('A' + n % 26));
            n /= 26;
        }
        return sb.ToString();
    }

    public static bool TryParseColumn(string letters, out int columnIndex)
    {
        columnIndex = 0;
        if (string.IsNullOrEmpty(letters)) return false;
        var c = 0;
        foreach (var ch in letters)
        {
            if (!char.IsAsciiLetter(ch)) return false;
            c = c * 26 + (char.ToUpperInvariant(ch) - 'A' + 1);
            if (c > MaxColumns) return false;
        }
        columnIndex = c;
        return c > 0;
    }

    public static string ToA1(int row, int column) => ColumnLetter(column) + row.ToString(System.Globalization.CultureInfo.InvariantCulture);

    public static bool TryParse(string? a1, out int row, out int column)
    {
        row = 0;
        column = 0;
        if (string.IsNullOrEmpty(a1)) return false;
        var i = 0;
        while (i < a1.Length && a1[i] == '$') i++;
        var c = 0;
        while (i < a1.Length && char.IsAsciiLetter(a1[i]))
        {
            c = c * 26 + (char.ToUpperInvariant(a1[i]) - 'A' + 1);
            if (c > MaxColumns) return false;
            i++;
        }
        while (i < a1.Length && a1[i] == '$') i++;
        var r = 0;
        var anyDigit = false;
        while (i < a1.Length && char.IsAsciiDigit(a1[i]))
        {
            r = r * 10 + (a1[i] - '0');
            if (r > MaxRows) return false;
            i++;
            anyDigit = true;
        }
        if (c == 0 || !anyDigit || r == 0 || i != a1.Length) return false;
        row = r;
        column = c;
        return true;
    }
}

/// <summary>An inclusive rectangular range of cells. Rows and columns are 1-based.</summary>
public readonly record struct CellRange(int FirstRow, int FirstColumn, int LastRow, int LastColumn)
{
    public int RowCount => LastRow - FirstRow + 1;
    public int ColumnCount => LastColumn - FirstColumn + 1;

    public bool Contains(int row, int column) => row >= FirstRow && row <= LastRow && column >= FirstColumn && column <= LastColumn;

    public bool Intersects(CellRange other) =>
        other.FirstRow <= LastRow && other.LastRow >= FirstRow && other.FirstColumn <= LastColumn && other.LastColumn >= FirstColumn;

    public static bool TryParse(string? text, out CellRange range)
    {
        range = default;
        if (string.IsNullOrWhiteSpace(text)) return false;
        var parts = text.Trim().Split(':');
        if (parts.Length == 1)
        {
            if (!CellRef.TryParse(parts[0], out var r, out var c)) return false;
            range = new CellRange(r, c, r, c);
            return true;
        }
        if (parts.Length != 2) return false;
        if (!CellRef.TryParse(parts[0], out var r1, out var c1) || !CellRef.TryParse(parts[1], out var r2, out var c2)) return false;
        range = new CellRange(Math.Min(r1, r2), Math.Min(c1, c2), Math.Max(r1, r2), Math.Max(c1, c2));
        return true;
    }

    public override string ToString() => FirstRow == LastRow && FirstColumn == LastColumn
        ? CellRef.ToA1(FirstRow, FirstColumn)
        : CellRef.ToA1(FirstRow, FirstColumn) + ":" + CellRef.ToA1(LastRow, LastColumn);
}
