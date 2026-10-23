using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Headless.XUnit;
using Avalonia.Interactivity;
using Avalonia.Threading;
using ExcelTransfer.App;
using ExcelTransfer.Core.Data;
using ExcelTransfer.Core.Workbooks;
using ExcelTransfer.Tests.Support;
using static ExcelTransfer.Tests.Support.WorkbookBuilder;

[assembly: AvaloniaTestApplication(typeof(ExcelTransfer.Tests.TestAppBuilder))]

namespace ExcelTransfer.Tests;

public static class TestAppBuilder
{
    public static AppBuilder BuildAvaloniaApp() => AppBuilder.Configure<ExcelTransfer.App.App>()
        .UseSkia()
        .UseHarfBuzz()
        .UseHeadless(new AvaloniaHeadlessPlatformOptions { UseHeadlessDrawing = false });
}

public class UiTests
{
    private static readonly string? ScreenshotDir = Environment.GetEnvironmentVariable("EXCELTRANSFER_SCREENSHOT_DIR");

    private static async Task ClickAsync(MainWindow window, Button button)
    {
        button.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
        await window.WaitForIdleAsync();
        Dispatcher.UIThread.RunJobs();
        await window.WaitForIdleAsync();
        Dispatcher.UIThread.RunJobs();
    }

    private static void Screenshot(MainWindow window, string name)
    {
        if (string.IsNullOrEmpty(ScreenshotDir)) return;
        Directory.CreateDirectory(ScreenshotDir);
        Dispatcher.UIThread.RunJobs();
        using var frame = window.CaptureRenderedFrame();
        frame?.Save(Path.Combine(ScreenshotDir, name + ".png"));
    }

    [AvaloniaFact]
    public async Task Wizard_runs_the_ProductCode_Region_example_end_to_end()
    {
        using var dir = new TempDir();
        var source = CreateSource(dir.File("source.xlsx"));
        var destination = CreateDestination(dir.File("destination.xlsx"));
        var output = dir.File("result.xlsx");

        var window = new MainWindow { Width = 1120, Height = 780 };
        window.Show();
        window.SourcePathBox.Text = source;
        window.DestinationPathBox.Text = destination;
        window.OutputPathBox.Text = output;
        Screenshot(window, "01-files");

        await ClickAsync(window, window.NextButton);
        Assert.Equal(MainWindow.StepSource, window.CurrentStep);
        Assert.Contains("OK", window.SourceFileInfo.Text);
        Assert.NotNull(window.SourceData);
        Assert.Equal(5, window.SourceData!.Records.Count);
        Assert.Contains("5 data rows", window.SourceDataInfo.Text);
        Screenshot(window, "02-source");

        await ClickAsync(window, window.NextButton);
        Assert.Equal(MainWindow.StepDestination, window.CurrentStep);
        Assert.IsType<TableInfo>(window.DestinationTableBox.SelectedItem);
        Assert.False(window.DestinationHeaderRowBox.IsEnabled);
        Assert.NotNull(window.DestinationData);
        Assert.Equal(5, window.DestinationData!.Records.Count);
        Screenshot(window, "03-destination");

        await ClickAsync(window, window.NextButton);
        Assert.Equal(MainWindow.StepKeys, window.CurrentStep);
        var keyRow = Assert.Single(window.KeyMappingRows);
        Assert.Equal("ProductCode", keyRow.Source?.Header);
        Assert.Equal("ProductCode", keyRow.Destination?.Header);
        var srcCols = window.SourceData.Columns;
        var dstCols = window.DestinationData.Columns;
        window.AddKeyRow(srcCols.First(c => c.Header == "Region"), dstCols.First(c => c.Header == "Region"));
        Screenshot(window, "04-keys");

        await ClickAsync(window, window.NextButton);
        Assert.Equal(MainWindow.StepPeriods, window.CurrentStep);
        var periodRow = Assert.Single(window.PeriodMappingRows);
        Assert.DoesNotContain(periodRow.DestinationBox.ItemsSource!.Cast<ColumnInfo>(), c => c.Header == "ProductCode"); // key columns are not offered
        periodRow.SourceBox.SelectedItem = srcCols.First(c => c.Header == "2026 Q1");
        periodRow.DestinationBox.SelectedItem = periodRow.DestinationBox.ItemsSource!.Cast<ColumnInfo>().First(c => c.Header == "Q1 Actual");
        var second = window.AddPeriodRow();
        second.SourceBox.SelectedItem = srcCols.First(c => c.Header == "2026 Q2");
        second.DestinationBox.SelectedItem = second.DestinationBox.ItemsSource!.Cast<ColumnInfo>().First(c => c.Header == "Q2 Actual");
        Screenshot(window, "05-periods");

        await ClickAsync(window, window.NextButton);
        Assert.Equal(MainWindow.StepPreview, window.CurrentStep);
        Assert.NotNull(window.CurrentPlan);
        Assert.True(window.CurrentPlan!.CanTransfer);
        Assert.Equal(4, window.CurrentPlan.Matches.Count);
        Assert.True(window.ConfirmBox.IsEnabled);
        Assert.False(window.NextButton.IsEnabled); // confirmation required
        Screenshot(window, "06-preview");
        window.ConfirmBox.IsChecked = true;
        Assert.True(window.NextButton.IsEnabled);

        await ClickAsync(window, window.NextButton);
        Assert.Equal(MainWindow.StepTransfer, window.CurrentStep);
        Assert.NotNull(window.LastResult);
        Assert.True(window.ResultCard.IsVisible);
        Assert.True(File.Exists(output));
        Assert.Equal(7, window.LastResult!.ValuesWritten);
        Screenshot(window, "07-done");

        var after = SheetReader.Load(output, DestinationTableSelection());
        Assert.Equal(110, after.Records.First(r => r.RowNumber == 2)[DstQ1].Number);
        Assert.True(after.Records.First(r => r.RowNumber == 6)[DstQ1].IsBlank);
        Assert.Equal(15, after.Records.First(r => r.RowNumber == 6)[DstQ3].Number);
    }

    [AvaloniaFact]
    public async Task Wizard_rejects_unsupported_files_and_bad_output_paths()
    {
        using var dir = new TempDir();
        var csv = dir.File("data.csv");
        File.WriteAllText(csv, "a,b");
        var destination = CreateDestination(dir.File("destination.xlsx"));

        var window = new MainWindow();
        window.Show();
        window.SourcePathBox.Text = csv;
        window.DestinationPathBox.Text = destination;
        await ClickAsync(window, window.NextButton);
        Assert.Equal(MainWindow.StepFiles, window.CurrentStep);
        Assert.Contains("not supported", window.SourceFileInfo.Text);

        window.SourcePathBox.Text = CreateSource(dir.File("source.xlsx"));
        window.OutputPathBox.Text = destination; // would overwrite an input
        await ClickAsync(window, window.NextButton);
        Assert.Equal(MainWindow.StepFiles, window.CurrentStep);
        Assert.Contains("must not overwrite", window.OutputInfo.Text);

        window.OutputPathBox.Text = string.Empty; // auto-suggest
        await ClickAsync(window, window.NextButton);
        Assert.Equal(MainWindow.StepSource, window.CurrentStep);
        Assert.EndsWith("destination - updated.xlsx", window.OutputPathBox.Text);
    }

    [AvaloniaFact]
    public async Task Wizard_blocks_duplicate_keys_in_the_preview()
    {
        using var dir = new TempDir();
        var source = CreateSource(dir.File("source.xlsx"), DefaultSource.Concat(new[] { new SourceRow("P-001", "North", 1.0, 1.0) }));
        var destination = CreateDestination(dir.File("destination.xlsx"));
        var window = new MainWindow();
        window.Show();
        window.SourcePathBox.Text = source;
        window.DestinationPathBox.Text = destination;
        window.OutputPathBox.Text = dir.File("result.xlsx");
        await ClickAsync(window, window.NextButton); // source
        await ClickAsync(window, window.NextButton); // destination
        await ClickAsync(window, window.NextButton); // keys
        window.AddKeyRow(window.SourceData!.Columns.First(c => c.Header == "Region"), window.DestinationData!.Columns.First(c => c.Header == "Region"));
        await ClickAsync(window, window.NextButton); // periods
        var row = window.PeriodMappingRows[0];
        row.SourceBox.SelectedItem = window.SourceData.Columns.First(c => c.Header == "2026 Q1");
        row.DestinationBox.SelectedItem = row.DestinationBox.ItemsSource!.Cast<ColumnInfo>().First(c => c.Header == "Q1 Actual");
        await ClickAsync(window, window.NextButton); // preview
        Assert.Equal(MainWindow.StepPreview, window.CurrentStep);
        Assert.False(window.CurrentPlan!.CanTransfer);
        Assert.Contains(window.CurrentPlan.Errors, e => e.Code == "duplicate-keys");
        Assert.False(window.ConfirmBox.IsEnabled);
        Assert.False(window.NextButton.IsEnabled);
        Screenshot(window, "08-blocked");
    }
}
