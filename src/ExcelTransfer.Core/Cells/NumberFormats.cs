namespace ExcelTransfer.Core.Cells;

/// <summary>Decides whether an Excel number format displays a date and/or time.</summary>
public static class NumberFormats
{
    // Built-in formats that Excel defines as date/time formats (ECMA-376 Part 1, 18.8.30).
    private static readonly HashSet<int> BuiltInDateFormats = new()
    {
        14, 15, 16, 17, 18, 19, 20, 21, 22,
        27, 28, 29, 30, 31, 32, 33, 34, 35, 36,
        45, 46, 47,
        50, 51, 52, 53, 54, 55, 56, 57, 58,
    };

    public static bool IsBuiltInDateFormat(int numberFormatId) => BuiltInDateFormats.Contains(numberFormatId);

    /// <summary>
    /// True when the format (built-in id, or custom format code when the id is custom) renders a date/time.
    /// Custom codes count as dates when they contain a day/month/year/hour/second token outside quoted
    /// literals, bracketed sections and escaped characters.
    /// </summary>
    public static bool IsDateFormat(int numberFormatId, string? formatCode)
    {
        if (numberFormatId == 0) return false; // General
        if (IsBuiltInDateFormat(numberFormatId)) return true;
        if (numberFormatId < 164 && formatCode is null) return false; // other built-ins are numeric/text
        return formatCode is not null && IsDateFormatCode(formatCode);
    }

    public static bool IsDateFormatCode(string formatCode)
    {
        if (string.IsNullOrWhiteSpace(formatCode)) return false;
        if (formatCode.Equals("General", StringComparison.OrdinalIgnoreCase)) return false;
        var i = 0;
        while (i < formatCode.Length)
        {
            var ch = formatCode[i];
            switch (ch)
            {
                case '"':
                    i++;
                    while (i < formatCode.Length && formatCode[i] != '"') i++;
                    i++;
                    continue;
                case '[':
                    while (i < formatCode.Length && formatCode[i] != ']') i++;
                    i++;
                    continue;
                case '\\':
                case '_':
                case '*':
                    i += 2;
                    continue;
                case 'd': case 'D':
                case 'm': case 'M':
                case 'y': case 'Y':
                case 'h': case 'H':
                case 's': case 'S':
                    return true;
                default:
                    i++;
                    continue;
            }
        }
        return false;
    }
}
