using System.Globalization;
using System.Text;
using ExcelTransfer.Core.Cells;

namespace ExcelTransfer.Core.Matching;

public enum KeyProblem
{
    None,
    /// <summary>The cell is blank or contains only whitespace.</summary>
    Missing,
    /// <summary>The cell contains an Excel error value such as #N/A.</summary>
    ErrorValue,
    /// <summary>The cell holds a formula whose result is not stored in the file.</summary>
    FormulaWithoutResult,
}

public sealed record KeyBuildResult(CompositeKey? Key, IReadOnlyList<(int ComponentIndex, KeyProblem Problem)> Problems)
{
    public bool Ok => Key is not null;
}

/// <summary>Turns cell values into comparable key parts. The rules are listed in <see cref="MatchingRules"/>.</summary>
public static class KeyNormalizer
{
    public static KeyBuildResult BuildKey(IReadOnlyList<CellValue> components)
    {
        var parts = new KeyPart[components.Count];
        var display = new string[components.Count];
        List<(int, KeyProblem)>? problems = null;
        for (var i = 0; i < components.Count; i++)
        {
            var problem = TryNormalize(components[i], out var part);
            if (problem != KeyProblem.None)
            {
                problems ??= new List<(int, KeyProblem)>();
                problems.Add((i, problem));
                continue;
            }
            parts[i] = part;
            display[i] = DisplayValue(components[i]);
        }
        if (problems is not null) return new KeyBuildResult(null, problems);
        return new KeyBuildResult(new CompositeKey(parts, display), Array.Empty<(int, KeyProblem)>());
    }

    public static KeyProblem TryNormalize(CellValue value, out KeyPart part)
    {
        part = default;
        switch (value.Kind)
        {
            case CellKind.Blank:
                return KeyProblem.Missing;
            case CellKind.Text:
            {
                var normalized = NormalizeText(value.TextOrEmpty);
                if (normalized.Length == 0) return KeyProblem.Missing;
                if (TryCanonicalNumber(value.TextOrEmpty, out var canonical))
                    part = new KeyPart(KeyPartKind.Number, canonical);
                else
                    part = new KeyPart(KeyPartKind.Text, normalized);
                return KeyProblem.None;
            }
            case CellKind.Number:
            case CellKind.Date:
                part = new KeyPart(KeyPartKind.Number, CanonicalNumber(value.Number));
                return KeyProblem.None;
            case CellKind.Boolean:
                part = new KeyPart(KeyPartKind.Boolean, value.BooleanValue ? "TRUE" : "FALSE");
                return KeyProblem.None;
            case CellKind.Error:
                return KeyProblem.ErrorValue;
            case CellKind.FormulaWithoutResult:
                return KeyProblem.FormulaWithoutResult;
            default:
                return KeyProblem.Missing;
        }
    }

    /// <summary>
    /// Text normalization: Unicode NFC, zero-width characters removed, every whitespace character (including
    /// non-breaking spaces) treated as a space, trimmed, internal runs collapsed to one space, upper-cased
    /// with the invariant culture. Leading zeros and punctuation are preserved.
    /// </summary>
    public static string NormalizeText(string text)
    {
        if (text.Length == 0) return string.Empty;
        var s = text.IsNormalized(NormalizationForm.FormC) ? text : text.Normalize(NormalizationForm.FormC);
        var sb = new StringBuilder(s.Length);
        var pendingSpace = false;
        foreach (var ch in s)
        {
            if (IsZeroWidth(ch)) continue;
            if (char.IsWhiteSpace(ch) || ch == ' ' || ch == ' ' || ch == ' ')
            {
                pendingSpace = sb.Length > 0;
                continue;
            }
            if (pendingSpace)
            {
                sb.Append(' ');
                pendingSpace = false;
            }
            sb.Append(ch);
        }
        return sb.ToString().ToUpperInvariant();
    }

    private static bool IsZeroWidth(char ch) => ch is '​' or '‌' or '‍' or '⁠' or '﻿';

    /// <summary>Numbers are compared after rounding to 15 significant digits, which is Excel's precision; -0 equals 0.</summary>
    public static string CanonicalNumber(double value)
    {
        if (value == 0 || double.IsNaN(value)) return "0";
        var rounded = double.Parse(value.ToString("G15", CultureInfo.InvariantCulture), NumberStyles.Float, CultureInfo.InvariantCulture);
        if (rounded == 0) return "0";
        return rounded.ToString("R", CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// A text matches a number only when the text is exactly the canonical form of that number, so "123"
    /// equals 123 but "007", "123.0", "1,234" and " 1e3" stay text. This never changes the meaning of an identifier.
    /// </summary>
    public static bool TryCanonicalNumber(string text, out string canonical)
    {
        canonical = string.Empty;
        var trimmed = text.Trim();
        if (trimmed.Length == 0 || trimmed.Length > 40) return false;
        if (!double.TryParse(trimmed, NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint | NumberStyles.AllowExponent, CultureInfo.InvariantCulture, out var d)) return false;
        if (double.IsNaN(d) || double.IsInfinity(d)) return false;
        var c = CanonicalNumber(d);
        if (!string.Equals(c, trimmed, StringComparison.Ordinal)) return false;
        canonical = c;
        return true;
    }

    public static string DisplayValue(CellValue value) => value.Kind == CellKind.Text ? value.TextOrEmpty.Trim() : value.Display();
}
