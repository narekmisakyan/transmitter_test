using ExcelTransfer.Core.Data;

namespace ExcelTransfer.Core.Transfer;

/// <summary>A source column paired with the destination column it maps to (1-based worksheet column indexes).</summary>
public sealed record ColumnMapping(int SourceColumnIndex, int DestinationColumnIndex);

public sealed record TransferOptions
{
    /// <summary>When true, records with missing key values are skipped (reported as warnings) instead of blocking the transfer.</summary>
    public bool AllowMissingKeys { get; init; }

    /// <summary>When true, an existing output file (that is not an input file) may be replaced.</summary>
    public bool OverwriteExistingOutput { get; init; }
}

/// <summary>Everything needed to validate and run one transfer.</summary>
public sealed record TransferSpec
{
    public required string SourcePath { get; init; }
    public required DataSetSelection Source { get; init; }
    public required string DestinationPath { get; init; }
    public required DataSetSelection Destination { get; init; }
    public required string OutputPath { get; init; }
    public required IReadOnlyList<ColumnMapping> KeyMappings { get; init; }
    public required IReadOnlyList<ColumnMapping> PeriodMappings { get; init; }
    public TransferOptions Options { get; init; } = new();
}
