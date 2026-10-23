namespace ExcelTransfer.Core.Workbooks;

/// <summary>Read-only access to a workbook, independent of its container format.</summary>
public interface IWorkbookSource : IDisposable
{
    string Path { get; }
    WorkbookFormat Format { get; }
    bool Date1904 { get; }
    bool ManualCalculation { get; }
    bool FullCalcOnLoad { get; }
    /// <summary>All sheets in workbook order, including hidden sheets and chart sheets.</summary>
    IReadOnlyList<SheetInfo> Sheets { get; }
    /// <summary>Reads every cell of the named worksheet. Sheet names are matched without regard to case.</summary>
    SheetSnapshot LoadSheet(string sheetName);
}

/// <summary>Opens the right reader for a workbook file.</summary>
public static class WorkbookSources
{
    public static IWorkbookSource Open(string path)
    {
        var format = WorkbookFormats.FromPath(path)
            ?? throw new Transfer.TransferException($"\"{System.IO.Path.GetFileName(path)}\" is not a supported workbook type. Use .xlsx, .xlsm or .xlsb files.");
        return format == WorkbookFormat.Xlsb
            ? Biff12.Biff12Source.Open(path)
            : OpenXmlSource.Open(path, format);
    }
}
