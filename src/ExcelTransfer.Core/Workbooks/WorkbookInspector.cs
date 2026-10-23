using ExcelTransfer.Core.Cells;

namespace ExcelTransfer.Core.Workbooks;

public sealed record TableColumnInfo(string Name, bool HasCalculatedColumnFormula, bool HasTotalsRowFunction);

/// <summary>An Excel Table (ListObject) on a worksheet.</summary>
public sealed record TableInfo(string Name, string DisplayName, CellRange Range, int HeaderRowCount, int TotalsRowCount, IReadOnlyList<TableColumnInfo> Columns)
{
    public int FirstDataRow => Range.FirstRow + HeaderRowCount;
    public int LastDataRow => Range.LastRow - TotalsRowCount;
    public int DataRowCount => Math.Max(0, LastDataRow - FirstDataRow + 1);
    public override string ToString() => $"{DisplayName} ({Range})";
}

public sealed record SheetInfo(string Name, int Index, bool Hidden, bool IsChartSheet, IReadOnlyList<TableInfo> Tables)
{
    public override string ToString() => Hidden ? Name + " (hidden)" : Name;
}

public sealed record WorkbookSummary(string Path, WorkbookFormat Format, IReadOnlyList<SheetInfo> Sheets, bool Date1904, bool ManualCalculation, bool FullCalcOnLoad)
{
    public IReadOnlyList<SheetInfo> Worksheets => Sheets.Where(s => !s.IsChartSheet).ToList();
    public SheetInfo? FindSheet(string name) => Sheets.FirstOrDefault(s => string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase));
}

/// <summary>Lists the worksheets and Excel Tables of a workbook.</summary>
public static class WorkbookInspector
{
    public static WorkbookSummary Inspect(string path)
    {
        using var source = WorkbookSources.Open(path);
        return new WorkbookSummary(source.Path, source.Format, source.Sheets, source.Date1904, source.ManualCalculation, source.FullCalcOnLoad);
    }
}
