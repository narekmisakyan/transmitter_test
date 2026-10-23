namespace ExcelTransfer.Core.Transfer;

public enum IssueSeverity
{
    /// <summary>Blocks the transfer.</summary>
    Error,
    /// <summary>Does not block, but the user should know.</summary>
    Warning,
    /// <summary>Informational.</summary>
    Info,
}

public enum DataSide
{
    General,
    Source,
    Destination,
}

/// <summary>A validation finding. <see cref="Details"/> lists affected records or cells (already capped for display).</summary>
public sealed record ValidationIssue(IssueSeverity Severity, DataSide Side, string Code, string Message, IReadOnlyList<string> Details, IReadOnlyList<int> Rows)
{
    public const int MaxDetails = 50;

    public static ValidationIssue Create(IssueSeverity severity, DataSide side, string code, string message, IEnumerable<string>? details = null, IEnumerable<int>? rows = null)
    {
        var d = details is null ? new List<string>() : details.ToList();
        var r = rows is null ? new List<int>() : rows.ToList();
        if (d.Count > MaxDetails)
        {
            var extra = d.Count - MaxDetails;
            d = d.Take(MaxDetails).ToList();
            d.Add($"... and {extra} more");
        }
        return new ValidationIssue(severity, side, code, message, d, r);
    }

    public bool IsError => Severity == IssueSeverity.Error;

    public override string ToString() => $"{Severity}: {Message}";
}
