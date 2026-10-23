using ExcelTransfer.Core.Workbooks;

namespace ExcelTransfer.Core.Transfer;

/// <summary>Rules for where the result may be written.</summary>
public static class OutputPathRules
{
    public static StringComparison PathComparison => OperatingSystem.IsLinux() ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

    public static bool SamePath(string a, string b)
    {
        try
        {
            return string.Equals(Path.GetFullPath(a).TrimEnd(Path.DirectorySeparatorChar), Path.GetFullPath(b).TrimEnd(Path.DirectorySeparatorChar), PathComparison);
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
        {
            return false;
        }
    }

    /// <summary>Default output name: "&lt;destination name&gt; - updated.&lt;same extension&gt;", made unique if that file exists.</summary>
    public static string SuggestOutputPath(string destinationPath)
    {
        var dir = Path.GetDirectoryName(Path.GetFullPath(destinationPath)) ?? string.Empty;
        var name = Path.GetFileNameWithoutExtension(destinationPath);
        var ext = Path.GetExtension(destinationPath);
        var candidate = Path.Combine(dir, $"{name} - updated{ext}");
        var n = 2;
        while (File.Exists(candidate) && n < 1000)
        {
            candidate = Path.Combine(dir, $"{name} - updated ({n}){ext}");
            n++;
        }
        return candidate;
    }

    /// <summary>Validates the output path. Returns issues (errors block the transfer).</summary>
    public static List<ValidationIssue> Validate(string? outputPath, string sourcePath, string destinationPath, WorkbookFormat destinationFormat, bool overwriteAllowed)
    {
        var issues = new List<ValidationIssue>();
        if (string.IsNullOrWhiteSpace(outputPath))
        {
            issues.Add(ValidationIssue.Create(IssueSeverity.Error, DataSide.General, "output-missing", "Choose where to save the output workbook."));
            return issues;
        }
        string fullPath;
        try
        {
            fullPath = Path.GetFullPath(outputPath);
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
        {
            issues.Add(ValidationIssue.Create(IssueSeverity.Error, DataSide.General, "output-invalid", $"The output path \"{outputPath}\" is not valid."));
            return issues;
        }

        var expectedExt = destinationFormat.Extension();
        if (!string.Equals(Path.GetExtension(fullPath), expectedExt, StringComparison.OrdinalIgnoreCase))
        {
            issues.Add(ValidationIssue.Create(IssueSeverity.Error, DataSide.General, "output-extension",
                $"The output file must use the destination's format and end with {expectedExt}: the destination workbook is copied and edited in place, so its format cannot change."));
        }
        if (SamePath(fullPath, sourcePath) || SamePath(fullPath, destinationPath))
        {
            issues.Add(ValidationIssue.Create(IssueSeverity.Error, DataSide.General, "output-overwrites-input",
                "The output file must not overwrite the source or destination workbook. Choose a different name."));
        }
        var dir = Path.GetDirectoryName(fullPath);
        if (string.IsNullOrEmpty(dir) || !Directory.Exists(dir))
        {
            issues.Add(ValidationIssue.Create(IssueSeverity.Error, DataSide.General, "output-folder-missing", $"The output folder \"{dir}\" does not exist."));
            return issues;
        }
        if (File.Exists(fullPath))
        {
            if (!overwriteAllowed)
                issues.Add(ValidationIssue.Create(IssueSeverity.Error, DataSide.General, "output-exists", $"The output file \"{Path.GetFileName(fullPath)}\" already exists. Choose another name or confirm that it may be replaced."));
            else if (!CanWriteExisting(fullPath))
                issues.Add(ValidationIssue.Create(IssueSeverity.Error, DataSide.General, "output-locked", $"The output file \"{Path.GetFileName(fullPath)}\" is read-only or open in another program. Close it or choose another name."));
        }
        if (!CanCreateIn(dir))
            issues.Add(ValidationIssue.Create(IssueSeverity.Error, DataSide.General, "output-folder-readonly", $"Files cannot be created in \"{dir}\". Check permissions or choose another folder."));
        return issues;
    }

    private static bool CanWriteExisting(string path)
    {
        try
        {
            if ((File.GetAttributes(path) & FileAttributes.ReadOnly) != 0) return false;
            using var s = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            return true;
        }
        catch (IOException) { return false; }
        catch (UnauthorizedAccessException) { return false; }
    }

    private static bool CanCreateIn(string dir)
    {
        var probe = Path.Combine(dir, ".excel-transfer-probe-" + Guid.NewGuid().ToString("N"));
        try
        {
            using (new FileStream(probe, FileMode.CreateNew, FileAccess.Write, FileShare.None, 1, FileOptions.DeleteOnClose)) { }
            return true;
        }
        catch (IOException) { return false; }
        catch (UnauthorizedAccessException) { return false; }
    }
}
