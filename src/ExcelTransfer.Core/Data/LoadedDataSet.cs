using ExcelTransfer.Core.Cells;
using ExcelTransfer.Core.Workbooks;

namespace ExcelTransfer.Core.Data;

/// <summary>The records and columns of a selected table or range, plus the facts needed to validate a transfer.</summary>
public sealed class LoadedDataSet
{
    public required string WorkbookPath { get; init; }
    public required WorkbookFormat Format { get; init; }
    public required DataSetSelection Selection { get; init; }
    /// <summary>Sheet name with the casing used in the workbook.</summary>
    public required string SheetName { get; init; }
    public TableInfo? Table { get; init; }
    public int? HeaderRow { get; init; }
    public required int FirstDataRow { get; init; }
    public required int LastDataRow { get; init; }
    public required int FirstColumn { get; init; }
    public required int LastColumn { get; init; }
    public required IReadOnlyList<ColumnInfo> Columns { get; init; }
    /// <summary>All data rows in worksheet order, hidden rows included (see <see cref="DataRecord.Hidden"/>).</summary>
    public required IReadOnlyList<DataRecord> Records { get; init; }
    public required SheetMetadata Metadata { get; init; }
    public bool Date1904 { get; init; }
    public bool ManualCalculation { get; init; }
    public bool FullCalcOnLoad { get; init; }

    public CellRange DataRange => new(FirstDataRow, FirstColumn, LastDataRow, LastColumn);

    public int DataRowCount => Math.Max(0, LastDataRow - FirstDataRow + 1);

    public ColumnInfo? FindColumn(int columnIndex) => Columns.FirstOrDefault(c => c.ColumnIndex == columnIndex);

    public string Describe() => Table is not null
        ? $"table \"{Table.DisplayName}\" on sheet \"{SheetName}\" ({Table.Range})"
        : $"sheet \"{SheetName}\", header row {HeaderRow}, data rows {FirstDataRow}-{LastDataRow}, columns {CellRef.ColumnLetter(FirstColumn)}-{CellRef.ColumnLetter(LastColumn)}";
}
