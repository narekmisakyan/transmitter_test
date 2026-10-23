using Avalonia.Controls;
using Avalonia.Threading;
using ExcelTransfer.Core.Data;
using ExcelTransfer.Core.Transfer;
using ExcelTransfer.Core.Workbooks;

namespace ExcelTransfer.App;

/// <summary>
/// Drives one "select sheet / table / rows" step: fills the sheet and table lists, suggests the header and data
/// rows for ordinary ranges, loads the selected data in the background and shows a preview.
/// </summary>
internal sealed class DataSetStep
{
    private const string RangeOption = "(No table: use header row + data rows)";

    private readonly string _role;
    private readonly ComboBox _sheetBox;
    private readonly ComboBox _tableBox;
    private readonly NumericUpDown _headerRowBox;
    private readonly NumericUpDown _firstRowBox;
    private readonly NumericUpDown _lastRowBox;
    private readonly CheckBox _includeHiddenBox;
    private readonly TextBlock _info;
    private readonly Grid _previewGrid;
    private readonly Action _onDataChanged;
    private bool _suppressEvents;
    private int _version;
    private Task _pending = Task.CompletedTask;

    public DataSetStep(string role, ComboBox sheetBox, ComboBox tableBox, NumericUpDown headerRowBox, NumericUpDown firstRowBox, NumericUpDown lastRowBox,
        CheckBox includeHiddenBox, TextBlock info, Grid previewGrid, Action onDataChanged)
    {
        _role = role;
        _sheetBox = sheetBox;
        _tableBox = tableBox;
        _headerRowBox = headerRowBox;
        _firstRowBox = firstRowBox;
        _lastRowBox = lastRowBox;
        _includeHiddenBox = includeHiddenBox;
        _info = info;
        _previewGrid = previewGrid;
        _onDataChanged = onDataChanged;

        _sheetBox.SelectionChanged += (_, _) => { if (!_suppressEvents) OnSheetChanged(); };
        _tableBox.SelectionChanged += (_, _) => { if (!_suppressEvents) Track(ReloadAsync(guessRange: true)); };
        _headerRowBox.ValueChanged += (_, _) => OnRangeEdited(headerChanged: true);
        _firstRowBox.ValueChanged += (_, _) => OnRangeEdited(headerChanged: false);
        _lastRowBox.ValueChanged += (_, _) => OnRangeEdited(headerChanged: false);
        _includeHiddenBox.IsCheckedChanged += (_, _) => { if (!_suppressEvents) Track(ReloadAsync(guessRange: false)); };
    }

    public string? Path { get; private set; }
    public WorkbookSummary? Summary { get; private set; }
    public LoadedDataSet? Data { get; private set; }
    public DataSetSelection? Selection { get; private set; }
    public string? LoadError { get; private set; }
    public bool IsLoading => !_pending.IsCompleted;
    public Task Pending => _pending;

    private void Track(Task task) => _pending = task;

    /// <summary>Loads the workbook summary when the path changed since the last load; keeps the current selection otherwise.</summary>
    public async Task EnsureLoadedAsync(string path)
    {
        var stamp = File.GetLastWriteTimeUtc(path);
        if (Summary is not null && string.Equals(Path, path, StringComparison.OrdinalIgnoreCase) && _loadedStamp == stamp) return;
        Path = path;
        _loadedStamp = stamp;
        Data = null;
        Selection = null;
        LoadError = null;
        PreviewRenderer.Render(_previewGrid, null);
        _info.Text = "Reading workbook...";
        try
        {
            Summary = await Task.Run(() => WorkbookInspector.Inspect(path));
        }
        catch (TransferException ex)
        {
            Summary = null;
            LoadError = ex.FullMessage;
            _info.Text = LoadError;
            return;
        }
        _suppressEvents = true;
        _sheetBox.ItemsSource = Summary.Worksheets;
        _sheetBox.SelectedIndex = -1;
        _suppressEvents = false;
        if (Summary.Worksheets.Count == 0)
        {
            _info.Text = "This workbook has no worksheets.";
            return;
        }
        var preferred = Summary.Worksheets.FirstOrDefault(s => !s.Hidden) ?? Summary.Worksheets[0];
        _suppressEvents = true;
        _sheetBox.SelectedItem = preferred;
        _suppressEvents = false;
        PopulateTables(preferred);
        await ReloadAsync(guessRange: true);
    }

    private DateTime _loadedStamp;

    private void OnSheetChanged()
    {
        if (_sheetBox.SelectedItem is not SheetInfo sheet) return;
        PopulateTables(sheet);
        Track(ReloadAsync(guessRange: true));
    }

    private void PopulateTables(SheetInfo sheet)
    {
        _suppressEvents = true;
        var items = new List<object> { RangeOption };
        items.AddRange(sheet.Tables);
        _tableBox.ItemsSource = items;
        _tableBox.SelectedItem = sheet.Tables.Count > 0 ? sheet.Tables[0] : RangeOption;
        _suppressEvents = false;
    }

    private void OnRangeEdited(bool headerChanged)
    {
        if (_suppressEvents || _tableBox.SelectedItem is TableInfo) return;
        if (headerChanged && _headerRowBox.Value is decimal h && _firstRowBox.Value is decimal f && f <= h)
        {
            _suppressEvents = true;
            _firstRowBox.Value = h + 1;
            _suppressEvents = false;
        }
        Track(ReloadAsync(guessRange: false));
    }

    public bool UsesTable => _tableBox.SelectedItem is TableInfo;

    private async Task ReloadAsync(bool guessRange)
    {
        var version = ++_version;
        var path = Path;
        if (path is null || _sheetBox.SelectedItem is not SheetInfo sheet) return;
        var table = _tableBox.SelectedItem as TableInfo;
        var includeHidden = _includeHiddenBox.IsChecked != false;
        SetRowBoxesEnabled(table is null);
        _info.Text = "Loading data...";
        LoadError = null;

        DataSetSelection selection;
        if (table is not null)
        {
            selection = new DataSetSelection { SheetName = sheet.Name, TableName = table.DisplayName, IncludeHiddenRows = includeHidden };
        }
        else
        {
            if (guessRange)
            {
                RangeGuess guess;
                try
                {
                    guess = await Task.Run(() => SheetReader.GuessRange(path, sheet.Name));
                }
                catch (TransferException ex)
                {
                    if (version != _version) return;
                    Fail(ex.FullMessage);
                    return;
                }
                if (version != _version) return;
                _suppressEvents = true;
                _headerRowBox.Value = guess.HeaderRow;
                _firstRowBox.Value = guess.FirstDataRow;
                _lastRowBox.Value = Math.Max(guess.LastDataRow, guess.FirstDataRow - 1);
                _suppressEvents = false;
            }
            selection = new DataSetSelection
            {
                SheetName = sheet.Name,
                HeaderRow = (int)(_headerRowBox.Value ?? 1),
                FirstDataRow = (int)(_firstRowBox.Value ?? 2),
                LastDataRow = (int)(_lastRowBox.Value ?? 1),
                IncludeHiddenRows = includeHidden,
            };
        }

        LoadedDataSet data;
        try
        {
            data = await Task.Run(() => SheetReader.Load(path, selection));
        }
        catch (TransferException ex)
        {
            if (version != _version) return;
            Fail(ex.FullMessage);
            return;
        }
        if (version != _version) return;

        Data = data;
        Selection = selection;
        if (table is not null)
        {
            _suppressEvents = true;
            _headerRowBox.Value = data.HeaderRow ?? data.FirstDataRow - 1;
            _firstRowBox.Value = data.FirstDataRow;
            _lastRowBox.Value = data.LastDataRow;
            _suppressEvents = false;
        }
        PreviewRenderer.Render(_previewGrid, data);
        var hidden = data.Records.Count(r => r.Hidden);
        var empty = data.Records.Count(r => r.IsCompletelyEmpty);
        var info = $"{UiText.Plural(data.Records.Count, "data row")} and {UiText.Plural(data.Columns.Count, "column")} in {data.Describe()}.";
        if (hidden > 0) info += $" {hidden} row(s) are hidden.";
        if (empty > 0) info += $" {empty} row(s) are completely empty and will be ignored.";
        if (data.Columns.Any(c => c.HeaderIsDuplicate)) info += " Some headers repeat; use the column letter to tell them apart.";
        if (data.Columns.Any(c => c.HeaderIsBlank)) info += " Some columns have no header.";
        info += data.Records.Count > 8 ? " Showing the first 8 rows." : string.Empty;
        _info.Text = info;
        _onDataChanged();
    }

    private void Fail(string message)
    {
        Data = null;
        Selection = null;
        LoadError = message;
        PreviewRenderer.Render(_previewGrid, null);
        _info.Text = message;
        _onDataChanged();
    }

    private void SetRowBoxesEnabled(bool enabled)
    {
        _headerRowBox.IsEnabled = enabled;
        _firstRowBox.IsEnabled = enabled;
        _lastRowBox.IsEnabled = enabled;
    }

    public bool Validate(out string? error)
    {
        if (IsLoading) { error = $"The {_role} data is still loading. Wait a moment and try again."; return false; }
        if (LoadError is not null) { error = LoadError; return false; }
        if (Data is null || Selection is null) { error = $"Select the {_role} worksheet and data first."; return false; }
        if (Data.Columns.Count == 0) { error = $"The selected {_role} data has no columns."; return false; }
        if (Data.Records.Count == 0) { error = $"The selected {_role} data has no data rows. Check the table or the first/last data row."; return false; }
        error = null;
        return true;
    }

    public void Reset()
    {
        _version++;
        Path = null;
        Summary = null;
        Data = null;
        Selection = null;
        LoadError = null;
        _suppressEvents = true;
        _sheetBox.ItemsSource = null;
        _tableBox.ItemsSource = null;
        _suppressEvents = false;
        PreviewRenderer.Render(_previewGrid, null);
        _info.Text = string.Empty;
    }
}
