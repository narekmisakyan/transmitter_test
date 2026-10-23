using System.Globalization;

namespace ExcelTransfer.Core.Cells;

/// <summary>Kind of value stored in a worksheet cell, as this application interprets it.</summary>
public enum CellKind
{
    /// <summary>No value (the cell is empty or does not exist).</summary>
    Blank,
    Text,
    Number,
    /// <summary>A number whose cell number format is a date/time format. Stored as an Excel 1900-system serial.</summary>
    Date,
    Boolean,
    /// <summary>An Excel error value such as #N/A or #REF!.</summary>
    Error,
    /// <summary>The cell contains a formula but the workbook stores no cached result for it.</summary>
    FormulaWithoutResult,
}

[Flags]
public enum CellFlags
{
    None = 0,
    /// <summary>The value was produced by a formula (we read the cached result stored in the file).</summary>
    FromFormula = 1,
    /// <summary>The cell's number format is a date/time format.</summary>
    DateStyled = 2,
}

/// <summary>
/// A cell value as read from a worksheet. Immutable. Dates are kept as Excel serial numbers in the
/// 1900 date system (values from 1904-system workbooks are converted on load).
/// </summary>
public readonly record struct CellValue(CellKind Kind, string? Text, double Number, CellFlags Flags)
{
    public static readonly CellValue Blank = new(CellKind.Blank, null, 0, CellFlags.None);

    public static CellValue OfText(string text, CellFlags flags = CellFlags.None) => new(CellKind.Text, text, 0, flags);
    public static CellValue OfNumber(double number, CellFlags flags = CellFlags.None) => new(CellKind.Number, null, number, flags);
    public static CellValue OfDate(double serial1900, CellFlags flags = CellFlags.None) => new(CellKind.Date, null, serial1900, flags | CellFlags.DateStyled);
    public static CellValue OfBoolean(bool value, CellFlags flags = CellFlags.None) => new(CellKind.Boolean, null, value ? 1 : 0, flags);
    public static CellValue OfError(string code, CellFlags flags = CellFlags.None) => new(CellKind.Error, code, 0, flags);
    public static CellValue FormulaWithoutResult(string formula, CellFlags flags = CellFlags.None) => new(CellKind.FormulaWithoutResult, formula, 0, flags | CellFlags.FromFormula);
    public static CellValue BlankWithFlags(CellFlags flags) => new(CellKind.Blank, null, 0, flags);

    public bool IsBlank => Kind == CellKind.Blank;
    public bool FromFormula => (Flags & CellFlags.FromFormula) != 0;
    public bool DateStyled => (Flags & CellFlags.DateStyled) != 0;
    public bool BooleanValue => Number != 0;
    public string TextOrEmpty => Text ?? string.Empty;

    /// <summary>The value as a calendar date/time, when the kind is Date and the serial is representable.</summary>
    public DateTime? ToDateTime() => Kind == CellKind.Date ? ExcelDates.ToDateTime(Number) : null;

    /// <summary>True when the stored content is the same, ignoring provenance flags.</summary>
    public bool ValueEquals(CellValue other)
    {
        if (Kind != other.Kind) return false;
        return Kind switch
        {
            CellKind.Blank => true,
            CellKind.Text or CellKind.Error or CellKind.FormulaWithoutResult => string.Equals(TextOrEmpty, other.TextOrEmpty, StringComparison.Ordinal),
            CellKind.Number or CellKind.Date => Number.Equals(other.Number) || Math.Abs(Number - other.Number) <= 1e-9 * Math.Max(1, Math.Abs(Number)),
            CellKind.Boolean => BooleanValue == other.BooleanValue,
            _ => false,
        };
    }

    /// <summary>Human-readable rendering used in previews and reports.</summary>
    public string Display()
    {
        switch (Kind)
        {
            case CellKind.Blank: return string.Empty;
            case CellKind.Text: return Text ?? string.Empty;
            case CellKind.Number: return FormatNumber(Number);
            case CellKind.Date:
                var dt = ExcelDates.ToDateTime(Number);
                if (dt is null) return FormatNumber(Number) + " (invalid date)";
                return dt.Value.TimeOfDay == TimeSpan.Zero
                    ? dt.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
                    : dt.Value.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
            case CellKind.Boolean: return BooleanValue ? "TRUE" : "FALSE";
            case CellKind.Error: return Text ?? "#N/A";
            case CellKind.FormulaWithoutResult: return "(formula without stored result: =" + Text + ")";
            default: return string.Empty;
        }
    }

    public static string FormatNumber(double n) => n.ToString("G15", CultureInfo.InvariantCulture);

    public override string ToString() => Display();
}
