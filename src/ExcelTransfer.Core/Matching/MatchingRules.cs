namespace ExcelTransfer.Core.Matching;

/// <summary>The matching rules in plain language, shown to users before a transfer.</summary>
public static class MatchingRules
{
    public static IReadOnlyList<string> Describe() => new[]
    {
        "Records match when every key column pair agrees, compared part by part in the selected order. Parts are never joined, so (\"AB\",\"C\") and (\"A\",\"BC\") are different keys.",
        "Text: spaces at the start and end are ignored, runs of spaces/tabs/line breaks inside the text count as one space, non-breaking spaces count as spaces, invisible zero-width characters are ignored, and letter case is ignored (\"abc\" = \"ABC\"). Accented characters are compared in their composed (NFC) form.",
        "Leading zeros are kept: \"007\" is different from \"7\" and from the number 7.",
        "Numbers are compared by value after rounding to 15 significant digits (Excel's precision): 1 = 1.0 = 1.00, and 0 = -0.",
        "Text matches a number only when the text is exactly how the number would be written: \"123\" = 123, but \"123.0\", \"1,234\" and \"1e3\" stay text and do not match numbers.",
        "Dates are compared by their underlying Excel serial number (converted from the 1904 date system when a workbook uses it), so the same date matches regardless of display format, and a date also equals the plain number with the same serial. A date never matches text such as \"2026-01-31\".",
        "TRUE/FALSE values compare as booleans; they do not match the texts \"TRUE\"/\"FALSE\" or the numbers 1/0.",
        "Blank cells and cells containing only spaces are missing key values. Error values (#N/A, #REF!, ...) and formulas without a stored result are invalid key values. Records with missing or invalid keys block the transfer unless you choose to skip them.",
        "Records whose key columns and transferred columns are all blank are ignored entirely.",
        "Duplicate keys on either side block the transfer: the first or last duplicate is never chosen silently, and values are never aggregated.",
    };
}
