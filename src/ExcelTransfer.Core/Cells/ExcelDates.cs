namespace ExcelTransfer.Core.Cells;

/// <summary>
/// Conversions between Excel date serial numbers and <see cref="DateTime"/>.
/// Excel's 1900 date system treats 1900 as a leap year (serial 60 = the non-existent 29 Feb 1900),
/// so serials below 61 are shifted by one day. The 1904 system counts days from 1 Jan 1904.
/// </summary>
public static class ExcelDates
{
    private static readonly DateTime Epoch1900 = new(1899, 12, 30);
    private static readonly DateTime Epoch1900Early = new(1899, 12, 31);
    private static readonly DateTime Epoch1904 = new(1904, 1, 1);

    /// <summary>Days to add to a 1904-system serial to express the same instant in the 1900 system.</summary>
    public const double Offset1904 = 1462;

    /// <summary>Converts a serial read from a workbook into a 1900-system serial.</summary>
    public static double ToSerial1900(double serial, bool workbookUses1904) => workbookUses1904 ? serial + Offset1904 : serial;

    /// <summary>Converts a 1900-system serial into the serial a workbook with the given date system expects.</summary>
    public static double FromSerial1900(double serial1900, bool workbookUses1904) => workbookUses1904 ? serial1900 - Offset1904 : serial1900;

    /// <summary>Returns the calendar value of a 1900-system serial, or null when it has no calendar meaning (serial 0 or 60).</summary>
    public static DateTime? ToDateTime(double serial1900)
    {
        if (double.IsNaN(serial1900) || double.IsInfinity(serial1900)) return null;
        if (serial1900 < 1 || serial1900 >= 2958466) return null; // 2958465 = 31 Dec 9999
        var whole = Math.Floor(serial1900);
        if (whole == 60) return null;
        var baseDate = whole < 60 ? Epoch1900Early : Epoch1900;
        try
        {
            var ticks = (long)Math.Round(serial1900 * TimeSpan.TicksPerDay);
            return baseDate.AddTicks(ticks);
        }
        catch (ArgumentOutOfRangeException)
        {
            return null;
        }
    }

    /// <summary>Converts a calendar value to a 1900-system serial.</summary>
    public static double FromDateTime(DateTime value)
    {
        var baseDate = value < new DateTime(1900, 3, 1) ? Epoch1900Early : Epoch1900;
        return (value - baseDate).Ticks / (double)TimeSpan.TicksPerDay;
    }

    public static DateTime From1904Serial(double serial) => Epoch1904.AddTicks((long)Math.Round(serial * TimeSpan.TicksPerDay));
}
