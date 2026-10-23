namespace ExcelTransfer.Core.Matching;

/// <summary>
/// An ordered tuple of normalized key parts. Equality is structural over the parts (never over a joined
/// string), so ("AB","C") and ("A","BC") are different keys even though their concatenations agree.
/// </summary>
public sealed class CompositeKey : IEquatable<CompositeKey>
{
    private readonly KeyPart[] _parts;
    private readonly string[] _display;
    private readonly int _hash;

    public CompositeKey(KeyPart[] parts, string[] displayParts)
    {
        if (parts.Length == 0) throw new ArgumentException("A key needs at least one part.", nameof(parts));
        if (parts.Length != displayParts.Length) throw new ArgumentException("Display parts must match key parts.", nameof(displayParts));
        _parts = parts;
        _display = displayParts;
        var h = new HashCode();
        h.Add(parts.Length);
        foreach (var p in parts)
        {
            h.Add((int)p.Kind);
            h.Add(p.Canonical, StringComparer.Ordinal);
        }
        _hash = h.ToHashCode();
    }

    public IReadOnlyList<KeyPart> Parts => _parts;

    /// <summary>The original cell values as displayed, for reports.</summary>
    public IReadOnlyList<string> DisplayParts => _display;

    public bool Equals(CompositeKey? other)
    {
        if (other is null) return false;
        if (ReferenceEquals(this, other)) return true;
        if (_parts.Length != other._parts.Length) return false;
        for (var i = 0; i < _parts.Length; i++)
        {
            if (_parts[i].Kind != other._parts[i].Kind) return false;
            if (!string.Equals(_parts[i].Canonical, other._parts[i].Canonical, StringComparison.Ordinal)) return false;
        }
        return true;
    }

    public override bool Equals(object? obj) => Equals(obj as CompositeKey);
    public override int GetHashCode() => _hash;

    /// <summary>Human-readable form, e.g. "P-001 | North". For display only; never used for comparison.</summary>
    public string Display() => string.Join(" | ", _display);
    public override string ToString() => Display();
}
