namespace ExcelTransfer.Core.Data;

/// <summary>Identifies the records a user selected in a workbook: either an Excel Table or a header row plus data rows.</summary>
public sealed record DataSetSelection
{
    public required string SheetName { get; init; }

    /// <summary>Name of an Excel Table on the sheet. When set, the header/data row properties are ignored.</summary>
    public string? TableName { get; init; }

    /// <summary>1-based row that holds the column headers (ordinary ranges only).</summary>
    public int? HeaderRow { get; init; }

    /// <summary>1-based first data row (ordinary ranges only). Defaults to the row after the header.</summary>
    public int? FirstDataRow { get; init; }

    /// <summary>1-based last data row (ordinary ranges only). Defaults to the last row with any content.</summary>
    public int? LastDataRow { get; init; }

    /// <summary>When false, rows hidden by filters or manually are excluded from matching and clearing.</summary>
    public bool IncludeHiddenRows { get; init; } = true;

    public bool IsTable => !string.IsNullOrEmpty(TableName);

    public string Describe() => IsTable
        ? $"table \"{TableName}\" on sheet \"{SheetName}\""
        : $"sheet \"{SheetName}\" (header row {HeaderRow?.ToString() ?? "auto"}, data rows {FirstDataRow?.ToString() ?? "auto"}-{LastDataRow?.ToString() ?? "auto"})";
}
