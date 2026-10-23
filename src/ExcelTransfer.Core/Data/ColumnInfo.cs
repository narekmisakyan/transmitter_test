using ExcelTransfer.Core.Cells;

namespace ExcelTransfer.Core.Data;

/// <summary>A column of a loaded data set, identified by its 1-based worksheet column index and its header text.</summary>
public sealed record ColumnInfo(int ColumnIndex, string Header, bool HeaderIsBlank, bool HeaderIsDuplicate)
{
    public string Letter => CellRef.ColumnLetter(ColumnIndex);

    /// <summary>Header plus column letter, e.g. "ProductCode (A)". The letter disambiguates duplicate or blank headers.</summary>
    public string Label => (HeaderIsBlank ? "(no header)" : Header) + " (" + Letter + ")";

    public override string ToString() => Label;
}
