using System.IO.Compression;
using System.Xml.Linq;
using ExcelTransfer.Core.Transfer;

namespace ExcelTransfer.Core.Workbooks.Biff12;

internal sealed record OpcRelationship(string Id, string Type, string TargetPart, bool External);

/// <summary>
/// Minimal Open Packaging Conventions reader used for .xlsb files, which the Open XML SDK cannot open.
/// Part names are stored without the leading slash, e.g. "xl/workbook.bin".
/// </summary>
internal sealed class OpcZipPackage : IDisposable
{
    public const string RelationshipsNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
    public const string ContentTypesNamespace = "http://schemas.openxmlformats.org/package/2006/content-types";
    public const string OfficeDocumentRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

    private readonly FileStream _stream;
    private readonly ZipArchive _zip;
    private readonly Dictionary<string, ZipArchiveEntry> _entries = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, string> _overrides = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, string> _defaults = new(StringComparer.OrdinalIgnoreCase);

    private OpcZipPackage(string path, FileStream stream, ZipArchive zip)
    {
        Path = path;
        _stream = stream;
        _zip = zip;
        foreach (var entry in zip.Entries)
        {
            var name = Normalize(entry.FullName);
            if (name.Length == 0 || name.EndsWith('/')) continue;
            _entries[name] = entry;
        }
        if (_entries.TryGetValue("[Content_Types].xml", out var ct))
        {
            var doc = LoadXml(ct);
            XNamespace ns = ContentTypesNamespace;
            foreach (var d in doc.Root!.Elements(ns + "Default"))
            {
                var ext = (string?)d.Attribute("Extension");
                var type = (string?)d.Attribute("ContentType");
                if (ext is not null && type is not null) _defaults[ext.TrimStart('.')] = type;
            }
            foreach (var o in doc.Root.Elements(ns + "Override"))
            {
                var part = (string?)o.Attribute("PartName");
                var type = (string?)o.Attribute("ContentType");
                if (part is not null && type is not null) _overrides[Normalize(part)] = type;
            }
        }
    }

    public string Path { get; }

    public static OpcZipPackage Open(string path)
    {
        var fullPath = System.IO.Path.GetFullPath(path);
        var name = System.IO.Path.GetFileName(fullPath);
        FileStream stream;
        try
        {
            stream = new FileStream(fullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        }
        catch (FileNotFoundException ex)
        {
            throw new TransferException($"The file \"{fullPath}\" does not exist.", "Check the path and that the drive or network share is available.", ex);
        }
        catch (UnauthorizedAccessException ex)
        {
            throw new TransferException($"You do not have permission to read \"{fullPath}\".", "Check the file and folder permissions.", ex);
        }
        catch (IOException ex)
        {
            throw new TransferException($"\"{name}\" could not be opened: {ex.Message}", "If the file is open in another program, close it and try again.", ex);
        }
        try
        {
            var zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
            return new OpcZipPackage(fullPath, stream, zip);
        }
        catch (TransferException) { stream.Dispose(); throw; }
        catch (Exception ex) when (ex is InvalidDataException or System.Xml.XmlException or IOException)
        {
            stream.Dispose();
            throw new TransferException($"\"{name}\" is not a valid workbook package: {ex.Message}", "Open it in Excel and save it again.", ex);
        }
    }

    public IEnumerable<string> PartNames => _entries.Keys;

    /// <summary>Zip entries in their original order (directory entries excluded).</summary>
    public IEnumerable<ZipArchiveEntry> Entries => _zip.Entries.Where(e => !e.FullName.EndsWith('/') && e.FullName.Length > 0);

    public byte[] ReadEntry(ZipArchiveEntry entry)
    {
        using var s = entry.Open();
        using var ms = new MemoryStream(entry.Length > 0 && entry.Length < int.MaxValue ? (int)entry.Length : 0);
        s.CopyTo(ms);
        return ms.ToArray();
    }

    public bool HasPart(string partName) => _entries.ContainsKey(Normalize(partName));

    public byte[] ReadPart(string partName)
    {
        if (!_entries.TryGetValue(Normalize(partName), out var entry))
            throw new TransferException($"The workbook is missing the part \"{partName}\".", "Open it in Excel and save it again.");
        using var s = entry.Open();
        using var ms = new MemoryStream(entry.Length > 0 && entry.Length < int.MaxValue ? (int)entry.Length : 0);
        s.CopyTo(ms);
        return ms.ToArray();
    }

    public string? ContentTypeOf(string partName)
    {
        var name = Normalize(partName);
        if (_overrides.TryGetValue(name, out var ct)) return ct;
        var ext = System.IO.Path.GetExtension(name).TrimStart('.');
        return _defaults.TryGetValue(ext, out var dct) ? dct : null;
    }

    /// <summary>Content type of the main document part, resolved through the package relationships.</summary>
    public string? MainPartName
    {
        get
        {
            var rels = ReadRelationships(string.Empty);
            var main = rels.FirstOrDefault(r => r.Type == OfficeDocumentRelType && !r.External);
            return main?.TargetPart;
        }
    }

    /// <summary>Relationships of a part (empty string for the package itself).</summary>
    public List<OpcRelationship> ReadRelationships(string partName)
    {
        var relsName = RelationshipsPartName(partName);
        var list = new List<OpcRelationship>();
        if (!_entries.TryGetValue(relsName, out var entry)) return list;
        var doc = LoadXml(entry);
        XNamespace ns = RelationshipsNamespace;
        var baseDir = partName.Length == 0 ? string.Empty : GetDirectory(partName);
        foreach (var rel in doc.Root!.Elements(ns + "Relationship"))
        {
            var id = (string?)rel.Attribute("Id") ?? string.Empty;
            var type = (string?)rel.Attribute("Type") ?? string.Empty;
            var target = (string?)rel.Attribute("Target") ?? string.Empty;
            var external = string.Equals((string?)rel.Attribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase);
            var resolved = external ? target : ResolveTarget(baseDir, target);
            list.Add(new OpcRelationship(id, type, resolved, external));
        }
        return list;
    }

    public static string RelationshipsPartName(string partName)
    {
        if (partName.Length == 0) return "_rels/.rels";
        var dir = GetDirectory(partName);
        var file = System.IO.Path.GetFileName(partName);
        return (dir.Length == 0 ? string.Empty : dir + "/") + "_rels/" + file + ".rels";
    }

    public static string GetDirectory(string partName)
    {
        var idx = partName.LastIndexOf('/');
        return idx < 0 ? string.Empty : partName.Substring(0, idx);
    }

    public static string ResolveTarget(string baseDir, string target)
    {
        if (target.StartsWith('/')) return Normalize(target);
        var segments = new List<string>();
        if (baseDir.Length > 0) segments.AddRange(baseDir.Split('/'));
        foreach (var seg in target.Split('/'))
        {
            if (seg == "..") { if (segments.Count > 0) segments.RemoveAt(segments.Count - 1); }
            else if (seg == "." || seg.Length == 0) { }
            else segments.Add(seg);
        }
        return string.Join('/', segments);
    }

    public static string Normalize(string partName) => Uri.UnescapeDataString(partName.Replace('\\', '/').TrimStart('/'));

    private static XDocument LoadXml(ZipArchiveEntry entry)
    {
        using var s = entry.Open();
        return XDocument.Load(s);
    }

    public void Dispose()
    {
        _zip.Dispose();
        _stream.Dispose();
    }
}
