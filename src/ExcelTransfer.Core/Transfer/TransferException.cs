namespace ExcelTransfer.Core.Transfer;

/// <summary>An error with a message written for end users. <see cref="Hint"/> suggests what to do next.</summary>
public sealed class TransferException : Exception
{
    public string? Hint { get; }

    public TransferException(string message, string? hint = null, Exception? inner = null) : base(message, inner)
    {
        Hint = hint;
    }

    public string FullMessage => Hint is null ? Message : Message + " " + Hint;
}
