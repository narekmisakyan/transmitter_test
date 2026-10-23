using ExcelTransfer.Core.Cells;

namespace ExcelTransfer.Core.Data;

/// <summary>A shared-formula group: the master cell's formula is reused by every cell in <see cref="Range"/>.</summary>
public sealed record SharedFormulaGroup(uint Index, CellRange Range);

/// <summary>Structural facts about a worksheet that affect whether cells can be cleared or written safely.</summary>
public sealed class SheetMetadata
{
    public bool Protected { get; init; }
    public IReadOnlyList<CellRange> MergedRanges { get; init; } = Array.Empty<CellRange>();
    public IReadOnlyList<SharedFormulaGroup> SharedFormulas { get; init; } = Array.Empty<SharedFormulaGroup>();
    public IReadOnlyList<CellRange> ArrayFormulas { get; init; } = Array.Empty<CellRange>();
    public IReadOnlyList<CellRange> DataTableFormulas { get; init; } = Array.Empty<CellRange>();

    /// <summary>Package URI of the worksheet part, e.g. "/xl/worksheets/sheet1.xml".</summary>
    public string PartUri { get; init; } = string.Empty;
}
