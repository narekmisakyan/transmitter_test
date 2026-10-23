using DocumentFormat.OpenXml.Packaging;
using ExcelTransfer.Core.Transfer;
using ExcelTransfer.Core.Workbooks.Biff12;

namespace ExcelTransfer.Core.Workbooks;

public enum FileProblem
{
    None,
    Missing,
    UnsupportedExtension,
    Locked,
    AccessDenied,
    Encrypted,
    LegacyFormat,
    FormatMismatch,
    Invalid,
}

public sealed record FileCheckResult(bool Ok, FileProblem Problem, string? Message, string? Hint, WorkbookFormat? Format)
{
    public static FileCheckResult Success(WorkbookFormat format) => new(true, FileProblem.None, null, null, format);
    public string FullMessage => Message is null ? string.Empty : Hint is null ? Message : Message + " " + Hint;
}

/// <summary>Checks that a path points to a readable .xlsx, .xlsm or .xlsb workbook and explains clearly when it does not.</summary>
public static class WorkbookFileCheck
{
    private static readonly byte[] ZipSignature = { 0x50, 0x4B, 0x03, 0x04 };
    private static readonly byte[] OleSignature = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };
    private const string SaveAsHint = "Save a copy as .xlsx, .xlsm or .xlsb in Excel (File > Save As), then select that copy.";

    public static FileCheckResult Check(string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
            return Fail(FileProblem.Missing, "No file selected.", "Choose an .xlsx, .xlsm or .xlsb workbook.");

        string fullPath;
        try
        {
            fullPath = Path.GetFullPath(path);
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
        {
            return Fail(FileProblem.Invalid, $"The path \"{path}\" is not valid.", null);
        }

        if (!File.Exists(fullPath))
            return Fail(FileProblem.Missing, $"The file \"{fullPath}\" does not exist.", "Check the path and that the drive or network share is available.");

        var name = Path.GetFileName(fullPath);
        var ext = Path.GetExtension(fullPath).ToLowerInvariant();
        var extFormat = WorkbookFormats.FromPath(fullPath);
        if (extFormat is null)
        {
            return ext switch
            {
                ".xls" => Fail(FileProblem.LegacyFormat, "Legacy Excel 97-2003 workbooks (.xls) are not supported.", SaveAsHint),
                ".csv" or ".txt" => Fail(FileProblem.UnsupportedExtension, "Text files (.csv/.txt) are not supported; they have no worksheets or tables.", "Open the file in Excel and save it as .xlsx."),
                ".ods" => Fail(FileProblem.UnsupportedExtension, "OpenDocument spreadsheets (.ods) are not supported.", SaveAsHint),
                ".xltx" or ".xltm" => Fail(FileProblem.UnsupportedExtension, "Excel templates are not supported.", "Create a workbook from the template and save it as .xlsx, .xlsm or .xlsb."),
                ".xlam" or ".xla" => Fail(FileProblem.UnsupportedExtension, "Excel add-ins are not workbooks and cannot be used.", null),
                _ => Fail(FileProblem.UnsupportedExtension, $"Files of type \"{ext}\" are not supported. Only .xlsx, .xlsm and .xlsb workbooks can be used.", null),
            };
        }

        byte[] header;
        try
        {
            using var stream = new FileStream(fullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            header = new byte[8];
            var read = 0;
            while (read < header.Length)
            {
                var n = stream.Read(header, read, header.Length - read);
                if (n <= 0) break;
                read += n;
            }
            if (read < 4)
                return Fail(FileProblem.Invalid, $"\"{name}\" is empty or truncated and is not a valid workbook.", null);
        }
        catch (UnauthorizedAccessException)
        {
            return Fail(FileProblem.AccessDenied, $"You do not have permission to read \"{fullPath}\".", "Check the file and folder permissions.");
        }
        catch (IOException ex)
        {
            return Fail(FileProblem.Locked, $"\"{name}\" could not be opened: {ex.Message}", "If the file is open in another program, close it and try again.");
        }

        if (header.AsSpan(0, 8).SequenceEqual(OleSignature))
            return Fail(FileProblem.Encrypted, $"\"{name}\" is password-protected (encrypted) or is a legacy .xls file saved with the wrong extension.", "Remove the password in Excel (File > Info > Protect Workbook) and save, then try again.");
        if (!header.AsSpan(0, 4).SequenceEqual(ZipSignature))
            return Fail(FileProblem.Invalid, $"\"{name}\" is not a valid Excel workbook.", "Open it in Excel and save it again.");

        // Look inside the package to confirm the real format matches the extension.
        WorkbookFormat? contentFormat;
        try
        {
            using var package = OpcZipPackage.Open(fullPath);
            var main = package.MainPartName;
            if (main is null)
                return Fail(FileProblem.Invalid, $"\"{name}\" does not contain a workbook.", "Open it in Excel and save it again.");
            var contentType = package.ContentTypeOf(main);
            contentFormat = WorkbookFormats.FromMainContentType(contentType);
            if (contentFormat is null)
            {
                if (contentType is not null && contentType.Contains("template", StringComparison.OrdinalIgnoreCase))
                    return Fail(FileProblem.UnsupportedExtension, $"\"{name}\" is an Excel template, not a workbook.", "Create a workbook from the template and save it as .xlsx, .xlsm or .xlsb.");
                if (contentType is not null && contentType.Contains("addin", StringComparison.OrdinalIgnoreCase))
                    return Fail(FileProblem.UnsupportedExtension, $"\"{name}\" is an Excel add-in, not a workbook.", null);
                if (contentType is not null && contentType.Contains("wordprocessingml", StringComparison.OrdinalIgnoreCase))
                    return Fail(FileProblem.Invalid, $"\"{name}\" is a Word document, not an Excel workbook.", null);
                if (contentType is not null && contentType.Contains("presentationml", StringComparison.OrdinalIgnoreCase))
                    return Fail(FileProblem.Invalid, $"\"{name}\" is a PowerPoint presentation, not an Excel workbook.", null);
                return Fail(FileProblem.Invalid, $"\"{name}\" is not a supported Excel workbook (content type {contentType ?? "unknown"}).", "Open it in Excel and save it as .xlsx, .xlsm or .xlsb.");
            }
        }
        catch (TransferException ex)
        {
            return Fail(FileProblem.Invalid, ex.Message, ex.Hint ?? "Open it in Excel and save it again.");
        }

        if (contentFormat != extFormat)
            return Fail(FileProblem.FormatMismatch,
                $"\"{name}\" has the extension {ext} but actually contains an {contentFormat.Value.Describe()}.",
                $"Rename the file to end with {contentFormat.Value.Extension()}, or open it in Excel and save it in the format you need.");

        try
        {
            if (contentFormat == WorkbookFormat.Xlsb)
            {
                using var source = Biff12Source.Open(fullPath);
                _ = source.Sheets;
            }
            else
            {
                using var stream = new FileStream(fullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using var doc = SpreadsheetDocument.Open(stream, false);
                if (doc.WorkbookPart?.Workbook is null)
                    return Fail(FileProblem.Invalid, $"\"{name}\" does not contain a workbook.", null);
            }
        }
        catch (TransferException ex)
        {
            return Fail(FileProblem.Invalid, ex.Message, ex.Hint);
        }
        catch (Exception ex) when (ex is not OutOfMemoryException)
        {
            return Fail(FileProblem.Invalid, $"\"{name}\" could not be read as an Excel workbook: {ex.Message}", "Open it in Excel and save it again.");
        }

        return FileCheckResult.Success(contentFormat.Value);
    }

    private static FileCheckResult Fail(FileProblem problem, string message, string? hint) => new(false, problem, message, hint, null);
}
