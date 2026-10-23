using ExcelTransfer.Core.Cells;
using ExcelTransfer.Core.Data;

namespace ExcelTransfer.Core.Workbooks;

/// <summary>The cells of one worksheet row that exist in the file, in ascending column order.</summary>
public sealed class RowSnapshot
{
    public RowSnapshot(int row, bool hidden, int[] columns, CellValue[] values)
    {
        if (columns.Length != values.Length) throw new ArgumentException("Column and value arrays must have the same length.");
        Row = row;
        Hidden = hidden;
        Columns = columns;
        Values = values;
    }

    /// <summary>1-based row number.</summary>
    public int Row { get; }
    public bool Hidden { get; }
    /// <summary>1-based column indexes, ascending.</summary>
    public int[] Columns { get; }
    public CellValue[] Values { get; }

    public bool HasContent
    {
        get
        {
            foreach (var v in Values)
                if (!v.IsBlank) return true;
            return false;
        }
    }
}

/// <summary>
/// A format-independent, read-only view of one worksheet: every stored cell as a typed value plus the
/// structural facts needed to validate edits. Produced by <see cref="IWorkbookSource.LoadSheet"/>.
/// </summary>
public sealed class SheetSnapshot
{
    public required string SheetName { get; init; }
    /// <summary>Package part name of the worksheet, e.g. "/xl/worksheets/sheet1.xml" or "/xl/worksheets/sheet1.bin".</summary>
    public required string PartUri { get; init; }
    /// <summary>Rows that exist in the file, ascending by row number.</summary>
    public required IReadOnlyList<RowSnapshot> Rows { get; init; }
    public IReadOnlyList<TableInfo> Tables { get; init; } = Array.Empty<TableInfo>();
    public bool Protected { get; init; }
    public IReadOnlyList<CellRange> MergedRanges { get; init; } = Array.Empty<CellRange>();
    public IReadOnlyList<SharedFormulaGroup> SharedFormulas { get; init; } = Array.Empty<SharedFormulaGroup>();
    public IReadOnlyList<CellRange> ArrayFormulas { get; init; } = Array.Empty<CellRange>();
    public IReadOnlyList<CellRange> DataTableFormulas { get; init; } = Array.Empty<CellRange>();

    public SheetMetadata ToMetadata() => new()
    {
        Protected = Protected,
        MergedRanges = MergedRanges,
        SharedFormulas = SharedFormulas,
        ArrayFormulas = ArrayFormulas,
        DataTableFormulas = DataTableFormulas,
        PartUri = PartUri,
    };
}
