using ExcelTransfer.Core.Cells;
using ExcelTransfer.Core.Transfer;

namespace ExcelTransfer.Core.Workbooks;

public sealed record SheetEditStats(int CellsCleared, int ValuesWritten, int DateStylesApplied);

/// <summary>
/// Applies cell edits to a working copy of a workbook. Only the cells named in the edits change; every other
/// cell, sheet and workbook part is carried over untouched. Implementations exist per container format.
/// </summary>
public interface IWorkbookEditor : IDisposable
{
    WorkbookFormat Format { get; }

    /// <summary>
    /// Clears the content of <paramref name="cellsToClear"/> (formatting is kept), then writes <paramref name="writes"/>.
    /// Cells on the same sheet may appear in both lists; writes win.
    /// </summary>
    SheetEditStats EditSheet(string sheetName, IReadOnlyCollection<(int Row, int Column)> cellsToClear, IReadOnlyList<CellWrite> writes, IProgress<ProgressReport>? progress, CancellationToken ct);

    /// <summary>
    /// Saves the working copy: removes the stored calculation chain and flags a full recalculation on next open,
    /// so formulas that depend on the written cells update when the workbook is opened in Excel.
    /// Returns the package part names (e.g. "/xl/worksheets/sheet1.xml") whose bytes may legitimately differ from the original.
    /// </summary>
    IReadOnlySet<string> Finish();
}

public static class WorkbookEditors
{
    /// <summary>Copies <paramref name="sourcePath"/> to <paramref name="workingCopyPath"/> and returns an editor for the copy.</summary>
    public static IWorkbookEditor CreateWorkingCopy(string sourcePath, string workingCopyPath, WorkbookFormat format)
    {
        return format == WorkbookFormat.Xlsb
            ? Biff12.Biff12Editor.Create(sourcePath, workingCopyPath)
            : OpenXmlEditor.Create(sourcePath, workingCopyPath, format);
    }
}
