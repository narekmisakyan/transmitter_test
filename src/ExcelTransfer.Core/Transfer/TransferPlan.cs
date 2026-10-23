using ExcelTransfer.Core.Cells;
using ExcelTransfer.Core.Data;
using ExcelTransfer.Core.Matching;
using ExcelTransfer.Core.Workbooks;

namespace ExcelTransfer.Core.Transfer;

public sealed record MatchedPair(int SourceRow, int DestinationRow, CompositeKey Key);

public sealed record UnmatchedRecord(int Row, CompositeKey Key);

/// <summary>A destination column whose data cells will be cleared before values are written.</summary>
public sealed record ClearColumnPlan(ColumnInfo Column, int CellCount, int NonBlankCount, int FormulaCount);


/// <summary>The result of validating a <see cref="TransferSpec"/>: issues, the match report and the exact edits.</summary>
public sealed class TransferPlan
{
    public required TransferSpec Spec { get; init; }
    public LoadedDataSet? Source { get; init; }
    public LoadedDataSet? Destination { get; init; }
    public WorkbookFormat? DestinationFormat { get; init; }
    public required IReadOnlyList<ValidationIssue> Issues { get; init; }

    public IReadOnlyList<MatchedPair> Matches { get; init; } = Array.Empty<MatchedPair>();
    public IReadOnlyList<UnmatchedRecord> UnmatchedSource { get; init; } = Array.Empty<UnmatchedRecord>();
    public IReadOnlyList<UnmatchedRecord> UnmatchedDestination { get; init; } = Array.Empty<UnmatchedRecord>();

    public IReadOnlyList<ClearColumnPlan> ColumnsToClear { get; init; } = Array.Empty<ClearColumnPlan>();
    /// <summary>Destination data rows (1-based) that are in scope for clearing and writing.</summary>
    public IReadOnlyList<int> DestinationRowsInScope { get; init; } = Array.Empty<int>();
    public IReadOnlyList<CellWrite> Writes { get; init; } = Array.Empty<CellWrite>();

    public int SourceRecordsConsidered { get; init; }
    public int DestinationRecordsConsidered { get; init; }
    public int SourceEmptyRecordsSkipped { get; init; }
    public int DestinationEmptyRecordsSkipped { get; init; }
    public int SourceHiddenRowsSkipped { get; init; }
    public int DestinationHiddenRowsSkipped { get; init; }
    public int SourceRecordsWithMissingKeys { get; init; }
    public int DestinationRecordsWithMissingKeys { get; init; }
    public int BlankSourceValues { get; init; }
    public int FormulaResultsUsed { get; init; }
    public int DateValuesNeedingDateFormat { get; init; }

    public IReadOnlyList<string> MatchingRules { get; } = Matching.MatchingRules.Describe();

    public bool CanTransfer => !Issues.Any(i => i.IsError) && Source is not null && Destination is not null;
    public IEnumerable<ValidationIssue> Errors => Issues.Where(i => i.Severity == IssueSeverity.Error);
    public IEnumerable<ValidationIssue> Warnings => Issues.Where(i => i.Severity == IssueSeverity.Warning);
    public IEnumerable<ValidationIssue> Infos => Issues.Where(i => i.Severity == IssueSeverity.Info);

    public int TotalCellsToClear => ColumnsToClear.Sum(c => c.CellCount);
    public int NonBlankCellsToClear => ColumnsToClear.Sum(c => c.NonBlankCount);
    public int FormulaCellsToClear => ColumnsToClear.Sum(c => c.FormulaCount);
    public int ValuesToWrite => Writes.Count;

    /// <summary>Every destination cell that will be cleared (rows in scope x mapped destination columns).</summary>
    public IEnumerable<(int Row, int Column)> CellsToClear()
    {
        foreach (var row in DestinationRowsInScope)
            foreach (var c in ColumnsToClear)
                yield return (row, c.Column.ColumnIndex);
    }
}
