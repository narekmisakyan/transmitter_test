namespace ExcelTransfer.Core.Matching;

public enum KeyPartKind
{
    Text,
    Number,
    Boolean,
}

/// <summary>One normalized component of a matching key. Two parts are equal only when kind and canonical form agree.</summary>
public readonly record struct KeyPart(KeyPartKind Kind, string Canonical);
