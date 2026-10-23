using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using ExcelTransfer.Core.Cells;
using ExcelTransfer.Core.Data;

namespace ExcelTransfer.App;

/// <summary>One "source column -> destination column" editor line.</summary>
internal sealed class MappingRow
{
    public MappingRow(IReadOnlyList<ColumnInfo> sourceColumns, IReadOnlyList<ColumnInfo> destinationColumns, Action<MappingRow> onRemove, Action onChanged)
    {
        SourceBox = new ComboBox { ItemsSource = sourceColumns, MinWidth = 260, HorizontalAlignment = HorizontalAlignment.Stretch, PlaceholderText = "Source column" };
        DestinationBox = new ComboBox { ItemsSource = destinationColumns, MinWidth = 260, HorizontalAlignment = HorizontalAlignment.Stretch, PlaceholderText = "Destination column" };
        RemoveButton = new Button { Content = "Remove", Margin = new Thickness(8, 0, 0, 0) };
        RemoveButton.Click += (_, _) => onRemove(this);
        SourceBox.SelectionChanged += (_, _) => onChanged();
        DestinationBox.SelectionChanged += (_, _) => onChanged();
        Panel = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto,*,Auto") };
        var arrow = new TextBlock { Text = "→", Margin = new Thickness(12, 0), VerticalAlignment = VerticalAlignment.Center, FontSize = 18 };
        Grid.SetColumn(arrow, 1);
        Grid.SetColumn(DestinationBox, 2);
        Grid.SetColumn(RemoveButton, 3);
        Panel.Children.Add(SourceBox);
        Panel.Children.Add(arrow);
        Panel.Children.Add(DestinationBox);
        Panel.Children.Add(RemoveButton);
    }

    public Grid Panel { get; }
    public ComboBox SourceBox { get; }
    public ComboBox DestinationBox { get; }
    public Button RemoveButton { get; }
    public ColumnInfo? Source => SourceBox.SelectedItem as ColumnInfo;
    public ColumnInfo? Destination => DestinationBox.SelectedItem as ColumnInfo;
    public bool IsComplete => Source is not null && Destination is not null;

    /// <summary>Replaces the destination choices, keeping the current selection when it is still offered.</summary>
    public void SetDestinationColumns(IReadOnlyList<ColumnInfo> columns)
    {
        var current = Destination;
        DestinationBox.ItemsSource = columns;
        DestinationBox.SelectedItem = current is not null ? columns.FirstOrDefault(c => c.ColumnIndex == current.ColumnIndex) : null;
    }
}

/// <summary>Renders the first rows of a data set into a grid of text cells.</summary>
internal static class PreviewRenderer
{
    private static readonly IBrush HeaderBrush = new SolidColorBrush(Color.Parse("#22808080"));
    private static readonly IBrush StripeBrush = new SolidColorBrush(Color.Parse("#0C808080"));
    private static readonly IBrush LineBrush = new SolidColorBrush(Color.Parse("#40808080"));

    public static void Render(Grid grid, LoadedDataSet? data, int maxRows = 8, int maxColumns = 30)
    {
        grid.Children.Clear();
        grid.RowDefinitions.Clear();
        grid.ColumnDefinitions.Clear();
        if (data is null || data.Columns.Count == 0) return;

        var columns = data.Columns.Take(maxColumns).ToList();
        var records = data.Records.Take(maxRows).ToList();
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        foreach (var _ in columns) grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        for (var r = 0; r <= records.Count; r++) grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        AddCell(grid, 0, 0, data.HeaderRow?.ToString() ?? "", bold: true, background: HeaderBrush);
        for (var c = 0; c < columns.Count; c++)
            AddCell(grid, 0, c + 1, columns[c].Label, bold: true, background: HeaderBrush);
        for (var r = 0; r < records.Count; r++)
        {
            var record = records[r];
            var stripe = r % 2 == 1 ? StripeBrush : null;
            AddCell(grid, r + 1, 0, record.RowNumber + (record.Hidden ? " (hidden)" : string.Empty), bold: false, background: HeaderBrush);
            for (var c = 0; c < columns.Count; c++)
            {
                var text = record[columns[c].ColumnIndex].Display();
                if (text.Length > 40) text = text[..37] + "...";
                AddCell(grid, r + 1, c + 1, text, bold: false, background: stripe);
            }
        }
    }

    private static void AddCell(Grid grid, int row, int column, string text, bool bold, IBrush? background)
    {
        var block = new TextBlock { Text = text, FontWeight = bold ? FontWeight.SemiBold : FontWeight.Normal, VerticalAlignment = VerticalAlignment.Center };
        var border = new Border
        {
            Child = block,
            Padding = new Thickness(8, 4),
            BorderBrush = LineBrush,
            BorderThickness = new Thickness(0, 0, 1, 1),
            Background = background ?? Brushes.Transparent,
            MinWidth = 60,
        };
        Grid.SetRow(border, row);
        Grid.SetColumn(border, column);
        grid.Children.Add(border);
    }
}

internal static class UiText
{
    public static string Plural(int count, string singular, string? plural = null) => count == 1 ? $"{count} {singular}" : $"{count} {plural ?? singular + "s"}";

    public static string ColumnList(IEnumerable<ColumnInfo> columns) => string.Join(", ", columns.Select(c => c.Label));

    public static string FileLabel(string path) => $"{Path.GetFileName(path)}  ({Path.GetDirectoryName(path)})";
}
