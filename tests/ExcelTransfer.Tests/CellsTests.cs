using ExcelTransfer.Core.Cells;

namespace ExcelTransfer.Tests;

public class CellsTests
{
    [Theory]
    [InlineData(1, "A")]
    [InlineData(26, "Z")]
    [InlineData(27, "AA")]
    [InlineData(702, "ZZ")]
    [InlineData(703, "AAA")]
    [InlineData(16384, "XFD")]
    public void Column_letters_round_trip(int index, string letters)
    {
        Assert.Equal(letters, CellRef.ColumnLetter(index));
        Assert.True(CellRef.TryParseColumn(letters, out var parsed));
        Assert.Equal(index, parsed);
    }

    [Theory]
    [InlineData("A1", 1, 1)]
    [InlineData("$B$12", 12, 2)]
    [InlineData("xfd1048576", 1048576, 16384)]
    public void Cell_references_parse(string text, int row, int col)
    {
        Assert.True(CellRef.TryParse(text, out var r, out var c));
        Assert.Equal(row, r);
        Assert.Equal(col, c);
        Assert.Equal(text.Replace("$", "").ToUpperInvariant(), CellRef.ToA1(r, c));
    }

    [Theory]
    [InlineData("")]
    [InlineData("1A")]
    [InlineData("A")]
    [InlineData("A0")]
    [InlineData("XFE1")]
    public void Invalid_cell_references_are_rejected(string text)
    {
        Assert.False(CellRef.TryParse(text, out _, out _));
    }

    [Fact]
    public void Ranges_parse_and_test_containment()
    {
        Assert.True(CellRange.TryParse("B2:D10", out var range));
        Assert.Equal(new CellRange(2, 2, 10, 4), range);
        Assert.True(range.Contains(5, 3));
        Assert.False(range.Contains(11, 3));
        Assert.True(range.Intersects(new CellRange(10, 4, 20, 20)));
        Assert.False(range.Intersects(new CellRange(11, 1, 11, 10)));
        Assert.Equal("B2:D10", range.ToString());
        Assert.True(CellRange.TryParse("D10:B2", out var reversed));
        Assert.Equal(range, reversed);
        Assert.True(CellRange.TryParse("C3", out var single));
        Assert.Equal("C3", single.ToString());
        Assert.False(CellRange.TryParse("A:B", out _));
    }

    [Theory]
    [InlineData(14, null, true)]
    [InlineData(22, null, true)]
    [InlineData(47, null, true)]
    [InlineData(0, null, false)]
    [InlineData(2, null, false)]
    [InlineData(164, "dd/mm/yyyy", true)]
    [InlineData(164, "[$-409]mmm-yy;@", true)]
    [InlineData(164, "[h]:mm:ss", true)]
    [InlineData(164, "0.00", false)]
    [InlineData(164, "#,##0 \"days\"", false)]
    [InlineData(164, "\"Year\" 0", false)]
    [InlineData(164, "General", false)]
    [InlineData(164, "0.00E+00", false)]
    [InlineData(164, "@", false)]
    [InlineData(164, "[Red]0.0;[Blue]-0.0", false)]
    public void Date_formats_are_recognised(int id, string? code, bool expected)
    {
        Assert.Equal(expected, NumberFormats.IsDateFormat(id, code));
    }

    [Fact]
    public void Excel_serials_convert_to_dates_including_the_1900_leap_year_bug()
    {
        Assert.Equal(new DateTime(1900, 1, 1), ExcelDates.ToDateTime(1));
        Assert.Equal(new DateTime(1900, 2, 28), ExcelDates.ToDateTime(59));
        Assert.Null(ExcelDates.ToDateTime(60));
        Assert.Equal(new DateTime(1900, 3, 1), ExcelDates.ToDateTime(61));
        Assert.Equal(new DateTime(2026, 3, 31), ExcelDates.ToDateTime(46112));
        Assert.Equal(46112, ExcelDates.FromDateTime(new DateTime(2026, 3, 31)));
        Assert.Equal(new DateTime(2026, 3, 31, 12, 0, 0), ExcelDates.ToDateTime(46112.5));
        Assert.Null(ExcelDates.ToDateTime(0));
        // 1904 system: serial 0 = 1 Jan 1904
        Assert.Equal(ExcelDates.FromDateTime(new DateTime(1904, 1, 1)), ExcelDates.ToSerial1900(0, workbookUses1904: true));
        Assert.Equal(0, ExcelDates.FromSerial1900(ExcelDates.FromDateTime(new DateTime(1904, 1, 1)), workbookUses1904: true));
    }

    [Fact]
    public void Cell_values_display_sensibly()
    {
        Assert.Equal("", CellValue.Blank.Display());
        Assert.Equal("1234.5", CellValue.OfNumber(1234.5).Display());
        Assert.Equal("2026-03-31", CellValue.OfDate(46112).Display());
        Assert.Equal("2026-03-31 06:00:00", CellValue.OfDate(46112.25).Display());
        Assert.Equal("TRUE", CellValue.OfBoolean(true).Display());
        Assert.Equal("#N/A", CellValue.OfError("#N/A").Display());
        Assert.Contains("A1*2", CellValue.FormulaWithoutResult("A1*2").Display());
        Assert.True(CellValue.FormulaWithoutResult("A1").FromFormula);
        Assert.True(CellValue.OfNumber(1, CellFlags.FromFormula).ValueEquals(CellValue.OfNumber(1)));
    }
}
