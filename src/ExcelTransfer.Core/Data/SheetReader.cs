using ExcelTransfer.Core.Cells;
using ExcelTransfer.Core.Transfer;
using ExcelTransfer.Core.Workbooks;

namespace ExcelTransfer.Core.Data;

/// <summary>Suggested header/data rows for a sheet, based on where the sheet has content.</summary>
public sealed record RangeGuess(int HeaderRow, int FirstDataRow, int LastDataRow, int FirstColumn, int LastColumn)
{
    public bool HasData => LastDataRow >= FirstDataRow && LastColumn >= FirstColumn;
}

/// <summary>Builds a <see cref="LoadedDataSet"/> from a sheet snapshot according to a <see cref="DataSetSelection"/>.</summary>
public static class SheetReader
{
    public const int MaxColumnsPerDataSet = 2000;

    public static RangeGuess GuessRange(string path, string sheetName)
    {
        using var source = WorkbookSources.Open(path);
        return GuessRange(source.LoadSheet(sheetName));
    }

    public static RangeGuess GuessRange(SheetSnapshot sheet)
    {
        int firstRow = 0, lastRow = 0, firstCol = 0, lastCol = 0;
        foreach (var row in sheet.Rows)
        {
            var any = false;
            for (var i = 0; i < row.Values.Length; i++)
            {
                if (row.Values[i].IsBlank) continue;
                any = true;
                var c = row.Columns[i];
                if (firstCol == 0 || c < firstCol) firstCol = c;
                if (c > lastCol) lastCol = c;
            }
            if (!any) continue;
            if (firstRow == 0) firstRow = row.Row;
            lastRow = row.Row;
        }
        if (firstRow == 0) return new RangeGuess(1, 2, 1, 1, 0);
        return new RangeGuess(firstRow, firstRow + 1, lastRow, firstCol, lastCol);
    }

    public static LoadedDataSet Load(string path, DataSetSelection selection)
    {
        using var source = WorkbookSources.Open(path);
        return Load(source, selection);
    }

    public static LoadedDataSet Load(IWorkbookSource source, DataSetSelection selection)
    {
        var sheet = source.LoadSheet(selection.SheetName);
        return Build(source, sheet, selection);
    }

    public static LoadedDataSet Build(IWorkbookSource source, SheetSnapshot sheet, DataSetSelection selection)
    {
        var fileName = Path.GetFileName(source.Path);
        var sheetName = sheet.SheetName;

        TableInfo? table = null;
        int? headerRow;
        int firstDataRow, lastDataRow, firstColumn, lastColumn;

        if (selection.IsTable)
        {
            table = sheet.Tables.FirstOrDefault(t =>
                         string.Equals(t.Name, selection.TableName, StringComparison.OrdinalIgnoreCase) ||
                         string.Equals(t.DisplayName, selection.TableName, StringComparison.OrdinalIgnoreCase))
                    ?? throw new TransferException($"Table \"{selection.TableName}\" was not found on sheet \"{sheetName}\" of \"{fileName}\".", "Select the table again.");
            headerRow = table.HeaderRowCount > 0 ? table.Range.FirstRow : null;
            firstDataRow = table.FirstDataRow;
            lastDataRow = table.LastDataRow;
            firstColumn = table.Range.FirstColumn;
            lastColumn = table.Range.LastColumn;
        }
        else
        {
            var guess = GuessRange(sheet);
            headerRow = selection.HeaderRow ?? guess.HeaderRow;
            if (headerRow < 1 || headerRow > CellRef.MaxRows)
                throw new TransferException($"Header row {headerRow} is not a valid row number.");
            firstDataRow = selection.FirstDataRow ?? headerRow.Value + 1;
            lastDataRow = selection.LastDataRow ?? Math.Max(guess.LastDataRow, firstDataRow - 1);
            if (firstDataRow <= headerRow)
                throw new TransferException($"The first data row ({firstDataRow}) must be below the header row ({headerRow}).");
            if (lastDataRow < firstDataRow - 1)
                throw new TransferException($"The last data row ({lastDataRow}) must not be above the first data row ({firstDataRow}).");
            if (lastDataRow > CellRef.MaxRows)
                throw new TransferException($"The last data row ({lastDataRow}) exceeds the worksheet limit of {CellRef.MaxRows} rows.");
            firstColumn = 0;
            lastColumn = 0;
        }

        var rowsByIndex = new Dictionary<int, RowSnapshot>(sheet.Rows.Count);
        foreach (var row in sheet.Rows) rowsByIndex[row.Row] = row;

        RowSnapshot? header = headerRow is int hr && rowsByIndex.TryGetValue(hr, out var h) ? h : null;

        if (!selection.IsTable)
        {
            var headerContentColumns = new List<int>();
            if (header is not null)
            {
                for (var i = 0; i < header.Values.Length; i++)
                    if (!header.Values[i].IsBlank) headerContentColumns.Add(header.Columns[i]);
            }
            if (headerContentColumns.Count == 0)
                throw new TransferException($"Row {headerRow} on sheet \"{sheetName}\" of \"{fileName}\" has no values, so it cannot be the header row.", "Choose the row that contains the column names.");
            firstColumn = headerContentColumns.Min();
            lastColumn = headerContentColumns.Max();
            foreach (var row in sheet.Rows)
            {
                if (row.Row < firstDataRow || row.Row > lastDataRow) continue;
                for (var i = row.Values.Length - 1; i >= 0; i--)
                {
                    if (row.Values[i].IsBlank) continue;
                    if (row.Columns[i] > lastColumn) lastColumn = row.Columns[i];
                    break;
                }
            }
        }

        if (lastColumn - firstColumn + 1 > MaxColumnsPerDataSet)
            throw new TransferException($"The selected data spans {lastColumn - firstColumn + 1} columns; at most {MaxColumnsPerDataSet} are supported.", "Narrow the data range.");

        var width = lastColumn - firstColumn + 1;

        var headers = new string[width];
        var headerBlank = new bool[width];
        for (var i = 0; i < width; i++)
        {
            var col = firstColumn + i;
            var text = string.Empty;
            if (header is not null)
            {
                var idx = Array.BinarySearch(header.Columns, col);
                if (idx >= 0) text = header.Values[idx].Display().Trim();
            }
            if (text.Length == 0 && table is not null && i < table.Columns.Count) text = table.Columns[i].Name.Trim();
            headers[i] = text;
            headerBlank[i] = text.Length == 0;
        }
        var duplicates = headers.Where(h => h.Length > 0)
            .GroupBy(h => h, StringComparer.OrdinalIgnoreCase)
            .Where(g => g.Count() > 1)
            .Select(g => g.Key)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        var columns = new List<ColumnInfo>(width);
        for (var i = 0; i < width; i++)
            columns.Add(new ColumnInfo(firstColumn + i, headers[i], headerBlank[i], duplicates.Contains(headers[i])));

        var records = new List<DataRecord>(Math.Max(0, lastDataRow - firstDataRow + 1));
        for (var r = firstDataRow; r <= lastDataRow; r++)
        {
            var values = new CellValue[width];
            var hidden = false;
            if (rowsByIndex.TryGetValue(r, out var row))
            {
                hidden = row.Hidden;
                for (var i = 0; i < row.Columns.Length; i++)
                {
                    var c = row.Columns[i];
                    if (c < firstColumn || c > lastColumn) continue;
                    values[c - firstColumn] = row.Values[i];
                }
            }
            records.Add(new DataRecord(r, hidden, firstColumn, values));
        }

        return new LoadedDataSet
        {
            WorkbookPath = source.Path,
            Format = source.Format,
            Selection = selection,
            SheetName = sheetName,
            Table = table,
            HeaderRow = headerRow,
            FirstDataRow = firstDataRow,
            LastDataRow = lastDataRow,
            FirstColumn = firstColumn,
            LastColumn = lastColumn,
            Columns = columns,
            Records = records,
            Metadata = sheet.ToMetadata(),
            Date1904 = source.Date1904,
            ManualCalculation = source.ManualCalculation,
            FullCalcOnLoad = source.FullCalcOnLoad,
        };
    }
}
