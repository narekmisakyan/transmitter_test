using Avalonia;

namespace ExcelTransfer.App;

internal static class Program
{
    // Avalonia configuration; also used by the previewer and by the headless UI tests.
    [STAThread]
    public static void Main(string[] args) => BuildAvaloniaApp().StartWithClassicDesktopLifetime(args);

    public static AppBuilder BuildAvaloniaApp() => AppBuilder.Configure<App>()
        .UsePlatformDetect()
        .WithInterFont()
        .LogToTrace();
}
