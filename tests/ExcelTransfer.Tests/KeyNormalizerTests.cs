using ExcelTransfer.Core.Cells;
using ExcelTransfer.Core.Matching;

namespace ExcelTransfer.Tests;

public class KeyNormalizerTests
{
    private static KeyPart Part(CellValue v)
    {
        Assert.Equal(KeyProblem.None, KeyNormalizer.TryNormalize(v, out var part));
        return part;
    }

    [Theory]
    [InlineData("abc", "ABC")]
    [InlineData("  P-001  ", "P-001")]
    [InlineData("north east", "NORTH EAST")]
    [InlineData("a   b\t\nc", "A B C")]
    [InlineData("zero​width", "ZEROWIDTH")]
    [InlineData("007", "007")]
    public void Text_is_normalised(string input, string expected)
    {
        Assert.Equal(expected, KeyNormalizer.NormalizeText(input));
    }

    [Fact]
    public void Composed_and_decomposed_accents_compare_equal()
    {
        var composed = Part(CellValue.OfText("café"));
        var decomposed = Part(CellValue.OfText("café"));
        Assert.Equal(composed, decomposed);
    }

    [Fact]
    public void Leading_zeros_are_preserved_and_do_not_match_numbers()
    {
        var text = Part(CellValue.OfText("007"));
        var number = Part(CellValue.OfNumber(7));
        Assert.Equal(KeyPartKind.Text, text.Kind);
        Assert.NotEqual(text, number);
    }

    [Theory]
    [InlineData("123", 123, true)]
    [InlineData(" 123 ", 123, true)]
    [InlineData("-5", -5, true)]
    [InlineData("1.5", 1.5, true)]
    [InlineData("123.0", 123, false)]
    [InlineData("1,234", 1234, false)]
    [InlineData("1e3", 1000, false)]
    [InlineData(".5", 0.5, false)]
    [InlineData("+7", 7, false)]
    public void Text_matches_a_number_only_in_canonical_form(string text, double number, bool shouldMatch)
    {
        var t = Part(CellValue.OfText(text));
        var n = Part(CellValue.OfNumber(number));
        Assert.Equal(shouldMatch, t == n);
    }

    [Fact]
    public void Numbers_compare_at_excel_precision()
    {
        Assert.Equal(Part(CellValue.OfNumber(0.1 + 0.2)), Part(CellValue.OfNumber(0.3)));
        Assert.Equal(Part(CellValue.OfNumber(-0.0)), Part(CellValue.OfNumber(0)));
        Assert.Equal(Part(CellValue.OfNumber(1.0)), Part(CellValue.OfNumber(1)));
        Assert.NotEqual(Part(CellValue.OfNumber(1.0000001)), Part(CellValue.OfNumber(1)));
    }

    [Fact]
    public void Dates_compare_by_serial_and_equal_the_plain_number()
    {
        var date = Part(CellValue.OfDate(46112));
        Assert.Equal(date, Part(CellValue.OfNumber(46112)));
        Assert.NotEqual(date, Part(CellValue.OfText("2026-03-31")));
    }

    [Fact]
    public void Booleans_match_only_booleans()
    {
        var b = Part(CellValue.OfBoolean(true));
        Assert.NotEqual(b, Part(CellValue.OfText("TRUE")));
        Assert.NotEqual(b, Part(CellValue.OfNumber(1)));
        Assert.Equal(b, Part(CellValue.OfBoolean(true)));
    }

    [Fact]
    public void Blank_whitespace_error_and_unresolved_formula_are_problems()
    {
        Assert.Equal(KeyProblem.Missing, KeyNormalizer.TryNormalize(CellValue.Blank, out _));
        Assert.Equal(KeyProblem.Missing, KeyNormalizer.TryNormalize(CellValue.OfText("   "), out _));
        Assert.Equal(KeyProblem.ErrorValue, KeyNormalizer.TryNormalize(CellValue.OfError("#N/A"), out _));
        Assert.Equal(KeyProblem.FormulaWithoutResult, KeyNormalizer.TryNormalize(CellValue.FormulaWithoutResult("A1"), out _));
    }

    [Fact]
    public void Composite_keys_are_compared_part_by_part_not_by_concatenation()
    {
        var ab_c = KeyNormalizer.BuildKey(new[] { CellValue.OfText("AB"), CellValue.OfText("C") }).Key!;
        var a_bc = KeyNormalizer.BuildKey(new[] { CellValue.OfText("A"), CellValue.OfText("BC") }).Key!;
        Assert.NotEqual(ab_c, a_bc);
        Assert.Equal("AB | C", ab_c.Display());

        var same = KeyNormalizer.BuildKey(new[] { CellValue.OfText(" ab "), CellValue.OfText("c") }).Key!;
        Assert.Equal(ab_c, same);
        Assert.Equal(ab_c.GetHashCode(), same.GetHashCode());

        var dict = new Dictionary<CompositeKey, int> { [ab_c] = 1 };
        Assert.True(dict.ContainsKey(same));
        Assert.False(dict.ContainsKey(a_bc));
    }

    [Fact]
    public void Build_key_reports_every_problem_component()
    {
        var result = KeyNormalizer.BuildKey(new[] { CellValue.OfText("x"), CellValue.Blank, CellValue.OfError("#REF!") });
        Assert.False(result.Ok);
        Assert.Equal(2, result.Problems.Count);
        Assert.Contains(result.Problems, p => p.ComponentIndex == 1 && p.Problem == KeyProblem.Missing);
        Assert.Contains(result.Problems, p => p.ComponentIndex == 2 && p.Problem == KeyProblem.ErrorValue);
    }

    [Fact]
    public void Matching_rules_are_documented()
    {
        var rules = MatchingRules.Describe();
        Assert.True(rules.Count >= 8);
        Assert.Contains(rules, r => r.Contains("Leading zeros"));
        Assert.Contains(rules, r => r.Contains("Duplicate keys"));
    }
}
