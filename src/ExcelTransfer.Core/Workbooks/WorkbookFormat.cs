namespace ExcelTransfer.Core.Workbooks;

/// <summary>Workbook container formats this application understands.</summary>
public enum WorkbookFormat
{
    /// <summary>Office Open XML workbook (.xlsx).</summary>
    Xlsx,
    /// <summary>Office Open XML macro-enabled workbook (.xlsm). Same structure as .xlsx plus a VBA project part.</summary>
    Xlsm,
    /// <summary>Excel binary workbook (.xlsb): the same package layout, but parts use the BIFF12 binary record format.</summary>
    Xlsb,
}

public static class WorkbookFormats
{
    public const string XlsxMainContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
    public const string XlsmMainContentType = "application/vnd.ms-excel.sheet.macroEnabled.main+xml";
    public const string XlsbMainContentType = "application/vnd.ms-excel.sheet.binary.macroEnabled.main";

    public static readonly IReadOnlyList<string> SupportedExtensions = new[] { ".xlsx", ".xlsm", ".xlsb" };

    public static WorkbookFormat? FromPath(string? path)
    {
        if (string.IsNullOrEmpty(path)) return null;
        return Path.GetExtension(path).ToLowerInvariant() switch
        {
            ".xlsx" => WorkbookFormat.Xlsx,
            ".xlsm" => WorkbookFormat.Xlsm,
            ".xlsb" => WorkbookFormat.Xlsb,
            _ => null,
        };
    }

    public static WorkbookFormat? FromMainContentType(string? contentType) => contentType switch
    {
        XlsxMainContentType => WorkbookFormat.Xlsx,
        XlsmMainContentType => WorkbookFormat.Xlsm,
        XlsbMainContentType => WorkbookFormat.Xlsb,
        _ => null,
    };

    public static string Extension(this WorkbookFormat format) => format switch
    {
        WorkbookFormat.Xlsx => ".xlsx",
        WorkbookFormat.Xlsm => ".xlsm",
        WorkbookFormat.Xlsb => ".xlsb",
        _ => throw new ArgumentOutOfRangeException(nameof(format)),
    };

    public static string Describe(this WorkbookFormat format) => format switch
    {
        WorkbookFormat.Xlsx => "Excel workbook (.xlsx)",
        WorkbookFormat.Xlsm => "Excel macro-enabled workbook (.xlsm)",
        WorkbookFormat.Xlsb => "Excel binary workbook (.xlsb)",
        _ => format.ToString(),
    };

    public static bool IsOpenXml(this WorkbookFormat format) => format != WorkbookFormat.Xlsb;
}
