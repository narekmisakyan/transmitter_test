using System.Diagnostics;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Platform.Storage;
using ExcelTransfer.Core.Data;
using ExcelTransfer.Core.Matching;
using ExcelTransfer.Core.Transfer;
using ExcelTransfer.Core.Workbooks;

namespace ExcelTransfer.App;

/// <summary>
/// The wizard: files -> source data -> destination data -> keys -> period columns -> validate/preview -> transfer.
/// All workbook work runs on a background thread; this class only orchestrates and renders.
/// </summary>
public partial class MainWindow : Window
{
    internal const int StepFiles = 0, StepSource = 1, StepDestination = 2, StepKeys = 3, StepPeriods = 4, StepPreview = 5, StepTransfer = 6;
    private static readonly string[] StepTitles = { "Files", "Source data", "Destination data", "Match keys", "Period columns", "Validate & preview", "Transfer" };
    private static readonly IBrush ErrorBrush = new SolidColorBrush(Color.Parse("#C62828"));
    private static readonly IBrush WarningBrush = new SolidColorBrush(Color.Parse("#B26A00"));
    private static readonly IBrush OkBrush = new SolidColorBrush(Color.Parse("#2E7D32"));
    private static readonly IBrush InfoBrush = new SolidColorBrush(Color.Parse("#1565C0"));
    private static readonly FilePickerFileType ExcelFilter = new("Excel workbooks (*.xlsx, *.xlsm, *.xlsb)") { Patterns = new[] { "*.xlsx", "*.xlsm", "*.xlsb" } };

    private readonly DataSetStep _source;
    private readonly DataSetStep _destination;
    private readonly List<MappingRow> _keyRows = new();
    private readonly List<MappingRow> _periodRows = new();
    private LoadedDataSet? _mappingSource;
    private LoadedDataSet? _mappingDestination;
    private WorkbookFormat? _destinationFormat;
    private TransferPlan? _plan;
    private TransferResult? _result;
    private bool _busy;
    private Task _pending = Task.CompletedTask;

    public MainWindow()
    {
        InitializeComponent();
        _source = new DataSetStep("source", SourceSheetBox, SourceTableBox, SourceHeaderRowBox, SourceFirstRowBox, SourceLastRowBox, SourceIncludeHiddenBox, SourceDataInfo, SourcePreviewGrid, OnDataSetChanged);
        _destination = new DataSetStep("destination", DestinationSheetBox, DestinationTableBox, DestinationHeaderRowBox, DestinationFirstRowBox, DestinationLastRowBox, DestinationIncludeHiddenBox, DestinationDataInfo, DestinationPreviewGrid, OnDataSetChanged);

        BackButton.Click += (_, _) => Track(GoBackAsync());
        NextButton.Click += (_, _) => Track(GoNextAsync());
        BrowseSourceButton.Click += (_, _) => Track(BrowseOpenAsync(SourcePathBox, "Select the source workbook"));
        BrowseDestinationButton.Click += (_, _) => Track(BrowseOpenAsync(DestinationPathBox, "Select the destination workbook"));
        BrowseOutputButton.Click += (_, _) => Track(BrowseSaveAsync());
        DestinationPathBox.TextChanged += (_, _) => { OutputInfo.Text = string.Empty; DestinationFileInfo.Text = string.Empty; };
        SourcePathBox.TextChanged += (_, _) => SourceFileInfo.Text = string.Empty;
        AddKeyButton.Click += (_, _) => AddMappingRow(_keyRows, KeyRows, isKey: true);
        AddPeriodButton.Click += (_, _) => AddMappingRow(_periodRows, PeriodRows, isKey: false);
        ValidateButton.Click += (_, _) => Track(ValidateAsync());
        ConfirmBox.IsCheckedChanged += (_, _) => UpdateButtons();
        AllowMissingKeysBox.IsCheckedChanged += (_, _) => InvalidatePlan();
        OpenFolderButton.Click += (_, _) => { if (_result is not null) RevealInFolder(_result.OutputPath); };
        RestartButton.Click += (_, _) => Restart();
        RulesList.ItemsSource = MatchingRules.Describe().Select((r, i) => $"{i + 1}. {r}").ToList();
        ShowStep(StepFiles);
    }

    // ---- state exposed for tests and automation
    internal int CurrentStep => Steps.SelectedIndex;
    internal TransferPlan? CurrentPlan => _plan;
    internal TransferResult? LastResult => _result;
    internal IReadOnlyList<MappingRow> KeyMappingRows => _keyRows;
    internal IReadOnlyList<MappingRow> PeriodMappingRows => _periodRows;
    internal LoadedDataSet? SourceData => _source.Data;
    internal LoadedDataSet? DestinationData => _destination.Data;

    /// <summary>Waits until background work started by the UI has settled (used by tests).</summary>
    internal async Task WaitForIdleAsync()
    {
        for (var i = 0; i < 200; i++)
        {
            var tasks = new[] { _pending, _source.Pending, _destination.Pending };
            if (tasks.All(t => t.IsCompleted)) return;
            try { await Task.WhenAll(tasks); }
            catch { /* reported in the UI */ }
        }
    }

    private void Track(Task task) => _pending = task;

    // ---- navigation

    private void ShowStep(int step)
    {
        for (var i = 0; i < Steps.ItemCount; i++)
            if (Steps.Items[i] is TabItem item && i == step) item.IsEnabled = true;
        Steps.SelectedIndex = step;
        for (var i = 0; i < Steps.ItemCount; i++)
            if (Steps.Items[i] is TabItem item && i != step) item.IsEnabled = false;
        StepIndicator.Text = string.Join("   ›   ", StepTitles.Select((t, i) => i == step ? $"▶ {i + 1}. {t}" : $"{i + 1}. {t}"));
        SetStatus(null);
        UpdateButtons();
    }

    private void UpdateButtons()
    {
        var step = Steps.SelectedIndex;
        BackButton.IsEnabled = !_busy && step > StepFiles && !(step == StepTransfer && _result is not null);
        NextButton.IsVisible = step != StepTransfer;
        NextButton.Content = step == StepPreview ? "Transfer" : "Next";
        NextButton.IsEnabled = !_busy && (step != StepPreview || (_plan is { CanTransfer: true } && ConfirmBox.IsChecked == true));
        ValidateButton.IsEnabled = !_busy;
    }

    private void SetStatus(string? text, bool error = false)
    {
        StatusText.Text = text ?? string.Empty;
        StatusText.Foreground = error ? ErrorBrush : Brushes.Gray;
    }

    private void SetBusy(bool busy, string? status)
    {
        _busy = busy;
        if (status is not null) SetStatus(status);
        UpdateButtons();
    }

    private Task GoBackAsync()
    {
        var step = Steps.SelectedIndex;
        if (_busy || step == StepFiles) return Task.CompletedTask;
        if (step == StepTransfer && _result is not null) return Task.CompletedTask;
        ShowStep(step - 1);
        return Task.CompletedTask;
    }

    private async Task GoNextAsync()
    {
        if (_busy) return;
        try
        {
            SetStatus(null);
            switch (Steps.SelectedIndex)
            {
                case StepFiles:
                    if (!await CompleteFilesStepAsync()) return;
                    ShowStep(StepSource);
                    await _source.EnsureLoadedAsync(SourcePathBox.Text!.Trim());
                    break;
                case StepSource:
                    if (!_source.Validate(out var sourceError)) { SetStatus(sourceError, error: true); return; }
                    ShowStep(StepDestination);
                    await _destination.EnsureLoadedAsync(DestinationPathBox.Text!.Trim());
                    break;
                case StepDestination:
                    if (!_destination.Validate(out var destinationError)) { SetStatus(destinationError, error: true); return; }
                    PrepareMappingRows();
                    ShowStep(StepKeys);
                    break;
                case StepKeys:
                    if (!ValidateKeys(out var keyError)) { SetStatus(keyError, error: true); return; }
                    RefreshPeriodColumnChoices();
                    ShowStep(StepPeriods);
                    break;
                case StepPeriods:
                    if (!ValidatePeriods(out var periodError)) { SetStatus(periodError, error: true); return; }
                    ShowStep(StepPreview);
                    await ValidateAsync();
                    break;
                case StepPreview:
                    if (_plan is not { CanTransfer: true } || ConfirmBox.IsChecked != true) return;
                    ShowStep(StepTransfer);
                    await RunTransferAsync();
                    break;
            }
        }
        catch (Exception ex)
        {
            SetStatus(UserMessage(ex), error: true);
        }
        finally
        {
            UpdateButtons();
        }
    }

    private static string UserMessage(Exception ex) => ex switch
    {
        TransferException te => te.FullMessage,
        OperationCanceledException => "The operation was cancelled.",
        _ => "Unexpected error: " + ex.Message,
    };

    // ---- step 1: files

    private async Task<bool> CompleteFilesStepAsync()
    {
        var sourcePath = (SourcePathBox.Text ?? string.Empty).Trim();
        var destinationPath = (DestinationPathBox.Text ?? string.Empty).Trim();
        if (sourcePath.Length == 0 || destinationPath.Length == 0)
        {
            SetStatus("Select both the source and the destination workbook.", error: true);
            return false;
        }
        SetBusy(true, "Checking files...");
        FileCheckResult sourceCheck, destinationCheck;
        try
        {
            (sourceCheck, destinationCheck) = await Task.Run(() => (WorkbookFileCheck.Check(sourcePath), WorkbookFileCheck.Check(destinationPath)));
        }
        finally
        {
            SetBusy(false, null);
        }
        ShowFileCheck(SourceFileInfo, sourceCheck);
        ShowFileCheck(DestinationFileInfo, destinationCheck);
        if (!sourceCheck.Ok || !destinationCheck.Ok)
        {
            SetStatus("Fix the file problems shown above, then continue.", error: true);
            return false;
        }
        _destinationFormat = destinationCheck.Format;
        if (string.IsNullOrWhiteSpace(OutputPathBox.Text))
            OutputPathBox.Text = OutputPathRules.SuggestOutputPath(destinationPath);
        var outputPath = OutputPathBox.Text!.Trim();
        var outputIssue = OutputPathRules.Validate(outputPath, sourcePath, destinationPath, _destinationFormat!.Value, OverwriteOutputBox.IsChecked == true).FirstOrDefault(i => i.IsError);
        if (outputIssue is not null)
        {
            OutputInfo.Text = outputIssue.Message;
            OutputInfo.Foreground = ErrorBrush;
            SetStatus(outputIssue.Message, error: true);
            return false;
        }
        OutputInfo.Text = $"The result will be saved as \"{Path.GetFileName(outputPath)}\" ({_destinationFormat.Value.Describe()}). The inputs stay unchanged.";
        OutputInfo.Foreground = OkBrush;
        return true;
    }

    private static void ShowFileCheck(TextBlock target, FileCheckResult check)
    {
        target.Text = check.Ok ? $"OK: {check.Format!.Value.Describe()}" : check.FullMessage;
        target.Foreground = check.Ok ? OkBrush : ErrorBrush;
    }

    private async Task BrowseOpenAsync(TextBox target, string title)
    {
        var files = await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = title,
            AllowMultiple = false,
            FileTypeFilter = new[] { ExcelFilter, FilePickerFileTypes.All },
        });
        var path = files.Count > 0 ? files[0].TryGetLocalPath() : null;
        if (path is not null) target.Text = path;
    }

    private async Task BrowseSaveAsync()
    {
        var destinationPath = (DestinationPathBox.Text ?? string.Empty).Trim();
        var format = _destinationFormat ?? WorkbookFormats.FromPath(destinationPath) ?? WorkbookFormat.Xlsx;
        var suggested = destinationPath.Length > 0 && File.Exists(destinationPath) ? OutputPathRules.SuggestOutputPath(destinationPath) : "transfer result" + format.Extension();
        var folder = Path.GetDirectoryName(suggested);
        var options = new FilePickerSaveOptions
        {
            Title = "Save the result as",
            SuggestedFileName = Path.GetFileName(suggested),
            DefaultExtension = format.Extension().TrimStart('.'),
            ShowOverwritePrompt = true,
            FileTypeChoices = new[] { new FilePickerFileType(format.Describe()) { Patterns = new[] { "*" + format.Extension() } } },
        };
        if (!string.IsNullOrEmpty(folder) && Directory.Exists(folder))
            options.SuggestedStartLocation = await StorageProvider.TryGetFolderFromPathAsync(folder);
        var file = await StorageProvider.SaveFilePickerAsync(options);
        var path = file?.TryGetLocalPath();
        if (path is null) return;
        if (File.Exists(path)) OverwriteOutputBox.IsChecked = true;
        OutputPathBox.Text = path;
    }

    // ---- steps 2-3: data sets

    private void OnDataSetChanged()
    {
        if (!ReferenceEquals(_mappingSource, _source.Data) || !ReferenceEquals(_mappingDestination, _destination.Data))
            InvalidatePlan();
    }

    // ---- steps 4-5: mappings

    private void PrepareMappingRows()
    {
        var source = _source.Data!;
        var destination = _destination.Data!;
        if (ReferenceEquals(_mappingSource, source) && ReferenceEquals(_mappingDestination, destination) && _keyRows.Count > 0) return;
        _mappingSource = source;
        _mappingDestination = destination;
        ClearRows(_keyRows, KeyRows);
        ClearRows(_periodRows, PeriodRows);
        InvalidatePlan();

        var suggestion = source.Columns.Where(s => !s.HeaderIsBlank)
            .Select(s => (Source: s, Destination: destination.Columns.FirstOrDefault(d => string.Equals(d.Header, s.Header, StringComparison.OrdinalIgnoreCase))))
            .FirstOrDefault(p => p.Destination is not null);
        if (suggestion.Destination is not null) AddMappingRow(_keyRows, KeyRows, isKey: true, suggestion.Source, suggestion.Destination);
        else AddMappingRow(_keyRows, KeyRows, isKey: true);
        AddMappingRow(_periodRows, PeriodRows, isKey: false);
    }

    private static void ClearRows(List<MappingRow> rows, StackPanel host)
    {
        rows.Clear();
        host.Children.Clear();
    }

    internal MappingRow AddMappingRow(List<MappingRow> rows, StackPanel host, bool isKey, ColumnInfo? source = null, ColumnInfo? destination = null)
    {
        var sourceColumns = _source.Data?.Columns ?? Array.Empty<ColumnInfo>();
        var destinationColumns = isKey ? _destination.Data?.Columns ?? Array.Empty<ColumnInfo>() : PeriodDestinationChoices();
        MappingRow? row = null;
        row = new MappingRow(sourceColumns, destinationColumns,
            onRemove: r => { rows.Remove(r); host.Children.Remove(r.Panel); if (isKey) RefreshPeriodColumnChoices(); InvalidatePlan(); },
            onChanged: () => { if (isKey) RefreshPeriodColumnChoices(); InvalidatePlan(); });
        if (source is not null) row.SourceBox.SelectedItem = source;
        if (destination is not null) row.DestinationBox.SelectedItem = destination;
        rows.Add(row);
        host.Children.Add(row.Panel);
        return row;
    }

    internal MappingRow AddKeyRow(ColumnInfo? source = null, ColumnInfo? destination = null) => AddMappingRow(_keyRows, KeyRows, true, source, destination);
    internal MappingRow AddPeriodRow(ColumnInfo? source = null, ColumnInfo? destination = null) => AddMappingRow(_periodRows, PeriodRows, false, source, destination);

    private IReadOnlyList<ColumnInfo> PeriodDestinationChoices()
    {
        var keyColumns = _keyRows.Select(r => r.Destination?.ColumnIndex).Where(i => i is not null).ToHashSet();
        return (_destination.Data?.Columns ?? Array.Empty<ColumnInfo>()).Where(c => !keyColumns.Contains(c.ColumnIndex)).ToList();
    }

    private void RefreshPeriodColumnChoices()
    {
        var choices = PeriodDestinationChoices();
        foreach (var row in _periodRows) row.SetDestinationColumns(choices);
    }

    private bool ValidateKeys(out string? error)
    {
        var complete = _keyRows.Where(r => r.IsComplete).ToList();
        if (complete.Count == 0) { error = "Choose at least one key column pair (a source column and the destination column that holds the same identifier)."; return false; }
        if (_keyRows.Any(r => !r.IsComplete)) { error = "Every key row needs both a source and a destination column. Remove unused rows."; return false; }
        if (complete.Select(r => r.Source!.ColumnIndex).Distinct().Count() != complete.Count) { error = "Each source column can be used only once as a key."; return false; }
        if (complete.Select(r => r.Destination!.ColumnIndex).Distinct().Count() != complete.Count) { error = "Each destination column can be used only once as a key."; return false; }
        error = null;
        return true;
    }

    private bool ValidatePeriods(out string? error)
    {
        var complete = _periodRows.Where(r => r.IsComplete).ToList();
        if (complete.Count == 0) { error = "Choose at least one period column pair to transfer."; return false; }
        if (_periodRows.Any(r => !r.IsComplete)) { error = "Every period row needs both a source and a destination column. Remove unused rows."; return false; }
        if (complete.Select(r => r.Destination!.ColumnIndex).Distinct().Count() != complete.Count) { error = "Two source columns cannot write to the same destination column."; return false; }
        var keyDestinations = _keyRows.Where(r => r.IsComplete).Select(r => r.Destination!.ColumnIndex).ToHashSet();
        if (complete.Any(r => keyDestinations.Contains(r.Destination!.ColumnIndex))) { error = "A destination key column cannot be used as a period column."; return false; }
        error = null;
        return true;
    }

    // ---- step 6: validation & preview

    private void InvalidatePlan()
    {
        _plan = null;
        ConfirmBox.IsChecked = false;
        ConfirmBox.IsEnabled = false;
        PreviewPanel.Children.Clear();
        UpdateButtons();
    }

    internal TransferSpec BuildSpec() => new()
    {
        SourcePath = (SourcePathBox.Text ?? string.Empty).Trim(),
        Source = _source.Selection ?? throw new TransferException("Select the source data first."),
        DestinationPath = (DestinationPathBox.Text ?? string.Empty).Trim(),
        Destination = _destination.Selection ?? throw new TransferException("Select the destination data first."),
        OutputPath = (OutputPathBox.Text ?? string.Empty).Trim(),
        KeyMappings = _keyRows.Where(r => r.IsComplete).Select(r => new ColumnMapping(r.Source!.ColumnIndex, r.Destination!.ColumnIndex)).ToList(),
        PeriodMappings = _periodRows.Where(r => r.IsComplete).Select(r => new ColumnMapping(r.Source!.ColumnIndex, r.Destination!.ColumnIndex)).ToList(),
        Options = new TransferOptions { AllowMissingKeys = AllowMissingKeysBox.IsChecked == true, OverwriteExistingOutput = OverwriteOutputBox.IsChecked == true },
    };

    private async Task ValidateAsync()
    {
        if (_busy) return;
        InvalidatePlan();
        TransferSpec spec;
        try { spec = BuildSpec(); }
        catch (TransferException ex) { SetStatus(ex.FullMessage, error: true); return; }
        SetBusy(true, "Validating...");
        ValidateProgress.IsVisible = true;
        ValidateProgress.Value = 0;
        var progress = new Progress<ProgressReport>(r =>
        {
            ValidateProgress.Value = (r.Fraction ?? 0) * 100;
            ValidateStage.Text = r.Stage;
        });
        try
        {
            _plan = await Task.Run(() => TransferPlanner.Plan(spec, progress));
            RenderPlan(_plan);
            ConfirmBox.IsEnabled = _plan.CanTransfer;
            SetStatus(_plan.CanTransfer
                ? "Validation passed. Review the preview, tick the confirmation, then press Transfer."
                : "The transfer is blocked. Resolve the errors listed above and validate again.", error: !_plan.CanTransfer);
        }
        catch (Exception ex)
        {
            PreviewPanel.Children.Add(Card("Validation failed", Line(UserMessage(ex), ErrorBrush)));
            SetStatus(UserMessage(ex), error: true);
        }
        finally
        {
            ValidateStage.Text = string.Empty;
            ValidateProgress.IsVisible = false;
            SetBusy(false, null);
        }
    }

    private void RenderPlan(TransferPlan plan)
    {
        PreviewPanel.Children.Clear();
        var spec = plan.Spec;
        var summary = new List<Control>
        {
            Line($"Source: {UiText.FileLabel(spec.SourcePath)} — {plan.Source?.Describe() ?? spec.Source.Describe()}"),
            Line($"Destination: {UiText.FileLabel(spec.DestinationPath)} — {plan.Destination?.Describe() ?? spec.Destination.Describe()}"),
            Line($"Output: {spec.OutputPath}"),
            Line($"Keys: {string.Join(", ", spec.KeyMappings.Select(m => $"{LabelOf(plan.Source, m.SourceColumnIndex)} → {LabelOf(plan.Destination, m.DestinationColumnIndex)}"))}"),
            Line($"Periods: {string.Join(", ", spec.PeriodMappings.Select(m => $"{LabelOf(plan.Source, m.SourceColumnIndex)} → {LabelOf(plan.Destination, m.DestinationColumnIndex)}"))}"),
        };
        if (plan.Source is not null && plan.Destination is not null)
        {
            summary.Add(Line($"Records: {plan.SourceRecordsConsidered} source and {plan.DestinationRecordsConsidered} destination records considered " +
                             $"({plan.SourceEmptyRecordsSkipped + plan.DestinationEmptyRecordsSkipped} completely empty rows ignored)."));
            summary.Add(Line($"Matched: {plan.Matches.Count} record pairs.   Unmatched source: {plan.UnmatchedSource.Count} (skipped, never appended).   Unmatched destination: {plan.UnmatchedDestination.Count} (period cells left blank).", bold: true));
            summary.Add(Line($"Cells to clear: {plan.TotalCellsToClear} in column(s) {UiText.ColumnList(plan.ColumnsToClear.Select(c => c.Column))} across {plan.DestinationRowsInScope.Count} data rows " +
                             $"({plan.NonBlankCellsToClear} currently contain values, {plan.FormulaCellsToClear} contain formulas).", bold: true));
            summary.Add(Line($"Values to write: {plan.ValuesToWrite}" + (plan.BlankSourceValues > 0 ? $" ({plan.BlankSourceValues} matched source cells are blank and leave the destination cell empty)" : string.Empty) +
                             (plan.FormulaResultsUsed > 0 ? $"; {plan.FormulaResultsUsed} come from stored formula results." : ".")));
            summary.Add(Line("Everything else (keys, other columns, headers, totals rows, formatting, other sheets) stays as it is.", InfoBrush));
        }
        PreviewPanel.Children.Add(Card(plan.CanTransfer ? "Ready to transfer" : "Transfer blocked", summary.ToArray()));

        var issueControls = new List<Control>();
        foreach (var issue in plan.Issues.OrderBy(i => i.Severity))
        {
            var brush = issue.Severity switch { IssueSeverity.Error => ErrorBrush, IssueSeverity.Warning => WarningBrush, _ => InfoBrush };
            issueControls.Add(Line($"[{issue.Severity}] {issue.Message}", brush, bold: issue.Severity == IssueSeverity.Error));
            foreach (var detail in issue.Details.Take(15))
                issueControls.Add(Line("      " + detail, Brushes.Gray));
            if (issue.Details.Count > 15) issueControls.Add(Line($"      ... {issue.Details.Count - 15} more", Brushes.Gray));
        }
        if (issueControls.Count == 0) issueControls.Add(Line("No issues found.", OkBrush));
        PreviewPanel.Children.Add(Card($"Validation results ({plan.Errors.Count()} errors, {plan.Warnings.Count()} warnings, {plan.Infos.Count()} notes)", issueControls.ToArray()));

        if (plan.Matches.Count > 0)
        {
            var lines = plan.Matches.Take(20).Select(m => Line($"source row {m.SourceRow} → destination row {m.DestinationRow}:  {m.Key.Display()}")).Cast<Control>().ToList();
            if (plan.Matches.Count > 20) lines.Add(Line($"... {plan.Matches.Count - 20} more", Brushes.Gray));
            PreviewPanel.Children.Add(Card($"Matched records ({plan.Matches.Count})", lines.ToArray()));
        }
        if (plan.UnmatchedSource.Count > 0)
            PreviewPanel.Children.Add(Card($"Source records without a destination match ({plan.UnmatchedSource.Count}) — skipped",
                plan.UnmatchedSource.Take(20).Select(u => Line($"row {u.Row}:  {u.Key.Display()}")).Cast<Control>().ToArray()));
        if (plan.UnmatchedDestination.Count > 0)
            PreviewPanel.Children.Add(Card($"Destination records without a source match ({plan.UnmatchedDestination.Count}) — period cells will be blank",
                plan.UnmatchedDestination.Take(20).Select(u => Line($"row {u.Row}:  {u.Key.Display()}")).Cast<Control>().ToArray()));

        var rules = new Expander { Header = "Matching rules used", Content = new ItemsControl { ItemsSource = plan.MatchingRules.Select((r, i) => $"{i + 1}. {r}").ToList(), Margin = new Thickness(8) } };
        PreviewPanel.Children.Add(rules);

        ConfirmBox.Content = plan.CanTransfer
            ? $"I have reviewed the preview: {plan.TotalCellsToClear} cells in {UiText.ColumnList(plan.ColumnsToClear.Select(c => c.Column))} will be cleared and {plan.ValuesToWrite} values written to \"{Path.GetFileName(spec.OutputPath)}\"."
            : "The transfer cannot run until the errors above are resolved.";
    }

    private static string LabelOf(LoadedDataSet? data, int columnIndex) => data?.FindColumn(columnIndex)?.Label ?? Core.Cells.CellRef.ColumnLetter(columnIndex);

    private static Border Card(string title, params Control[] content)
    {
        var panel = new StackPanel { Spacing = 3 };
        panel.Children.Add(new TextBlock { Text = title, FontWeight = FontWeight.SemiBold, FontSize = 15, Margin = new Thickness(0, 0, 0, 6) });
        foreach (var c in content) panel.Children.Add(c);
        var border = new Border { Child = panel };
        border.Classes.Add("card");
        return border;
    }

    private static TextBlock Line(string text, IBrush? brush = null, bool bold = false)
    {
        var block = new TextBlock { Text = text, TextWrapping = TextWrapping.Wrap };
        if (brush is not null) block.Foreground = brush;
        if (bold) block.FontWeight = FontWeight.SemiBold;
        return block;
    }

    // ---- step 7: transfer

    private async Task RunTransferAsync()
    {
        if (_plan is null) return;
        _result = null;
        ResultCard.IsVisible = false;
        TransferError.IsVisible = false;
        TransferProgress.Value = 0;
        TransferStage.Text = "Starting...";
        SetBusy(true, "Transferring...");
        var progress = new Progress<ProgressReport>(r =>
        {
            TransferProgress.Value = (r.Fraction ?? 0) * 100;
            TransferStage.Text = r.Stage;
        });
        try
        {
            var plan = _plan;
            _result = await Task.Run(() => TransferExecutor.Execute(plan, progress));
            TransferStage.Text = "Finished.";
            ResultTitle.Text = "Transfer complete";
            ResultDetails.Text = $"Saved \"{_result.OutputPath}\" ({_result.Format.Describe()}).\n" +
                                 $"Matched {UiText.Plural(_result.RecordsMatched, "record")}, cleared {UiText.Plural(_result.CellsCleared, "cell")} that had content, wrote {UiText.Plural(_result.ValuesWritten, "value")} in {_result.Duration.TotalSeconds:F1} s.";
            ResultNotes.ItemsSource = _result.Notes.Select(n => "• " + n).ToList();
            ResultCard.IsVisible = true;
            SetStatus("Done. The source and destination workbooks were left unchanged.");
        }
        catch (Exception ex)
        {
            TransferStage.Text = "Failed.";
            TransferError.Text = UserMessage(ex) + "\n\nNothing was written: the output file was discarded and the original workbooks are untouched. Press Back to review the preview or fix the problem and try again.";
            TransferError.IsVisible = true;
            SetStatus("The transfer failed.", error: true);
        }
        finally
        {
            SetBusy(false, null);
        }
    }

    private void Restart()
    {
        _result = null;
        _plan = null;
        _mappingSource = null;
        _mappingDestination = null;
        ClearRows(_keyRows, KeyRows);
        ClearRows(_periodRows, PeriodRows);
        InvalidatePlan();
        _source.Reset();
        _destination.Reset();
        ResultCard.IsVisible = false;
        TransferError.IsVisible = false;
        TransferProgress.Value = 0;
        TransferStage.Text = string.Empty;
        OutputPathBox.Text = string.Empty;
        OverwriteOutputBox.IsChecked = false;
        ShowStep(StepFiles);
    }

    private static void RevealInFolder(string path)
    {
        try
        {
            if (OperatingSystem.IsWindows())
                Process.Start(new ProcessStartInfo("explorer.exe", $"/select,\"{path}\"") { UseShellExecute = true });
            else if (OperatingSystem.IsMacOS())
                Process.Start("open", new[] { "-R", path });
            else
                Process.Start("xdg-open", new[] { Path.GetDirectoryName(path) ?? "." });
        }
        catch (Exception)
        {
            // Opening a file browser is a convenience only.
        }
    }
}
