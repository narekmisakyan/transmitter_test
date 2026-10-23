using ExcelTransfer.Core.Cells;

namespace ExcelTransfer.Core.Data;

/// <summary>One data row of a loaded data set.</summary>
public sealed class DataRecord
{
    private readonly CellValue[] _values;
    private readonly int _firstColumn;

    public DataRecord(int rowNumber, bool hidden, int firstColumn, CellValue[] values)
    {
        RowNumber = rowNumber;
        Hidden = hidden;
        _firstColumn = firstColumn;
        _values = values;
    }

    /// <summary>1-based worksheet row number.</summary>
    public int RowNumber { get; }

    /// <summary>True when the row is hidden (manually or by a filter).</summary>
    public bool Hidden { get; }

    public IReadOnlyList<CellValue> Values => _values;

    /// <summary>Value in the given 1-based worksheet column, or Blank when the column is outside the data set.</summary>
    public CellValue this[int columnIndex]
    {
        get
        {
            var i = columnIndex - _firstColumn;
            return i >= 0 && i < _values.Length ? _values[i] : CellValue.Blank;
        }
    }

    public bool IsCompletelyEmpty
    {
        get
        {
            foreach (var v in _values)
                if (!v.IsBlank) return false;
            return true;
        }
    }
}
