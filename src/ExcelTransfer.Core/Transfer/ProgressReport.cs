namespace ExcelTransfer.Core.Transfer;

/// <summary>Progress information reported by long-running operations.</summary>
/// <param name="Stage">Short description of the current stage.</param>
/// <param name="Fraction">Overall completion between 0 and 1, when known.</param>
public sealed record ProgressReport(string Stage, double? Fraction = null);
