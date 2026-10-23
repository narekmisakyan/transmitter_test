namespace ExcelTransfer.Tests.Support;

/// <summary>A scratch folder that is deleted when disposed.</summary>
public sealed class TempDir : IDisposable
{
    public TempDir()
    {
        Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "ExcelTransferTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(Path);
    }

    public string Path { get; }

    public string File(string name) => System.IO.Path.Combine(Path, name);

    public void Dispose()
    {
        try { Directory.Delete(Path, recursive: true); }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
    }
}

public static class Fixtures
{
    public static string PathOf(string relative) => System.IO.Path.Combine(AppContext.BaseDirectory, "Fixtures", relative.Replace('/', System.IO.Path.DirectorySeparatorChar));

    public static string CopyTo(TempDir dir, string relative, string? newName = null)
    {
        var target = dir.File(newName ?? System.IO.Path.GetFileName(relative));
        System.IO.File.Copy(PathOf(relative), target, overwrite: true);
        return target;
    }

    public static IEnumerable<string> XlsbFiles => Directory.GetFiles(PathOf("xlsb"), "*.xlsb").OrderBy(p => p, StringComparer.Ordinal);
}
