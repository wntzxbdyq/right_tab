using UnifiedContextMenu.Infrastructure.Windows;

namespace UnifiedContextMenu.App.WinForms;

internal static class Program
{
    [STAThread]
    private static void Main(string[] args)
    {
        ApplicationConfiguration.Initialize();

        var contextMenuProvider = new RegistryContextMenuProvider();
        var fluentModeService = new WindowsFluentModeService();
        var explorerService = new WindowsExplorerService();
        var openWithService = new WindowsOpenWithService();
        var sendToService = new WindowsSendToService();
        var winXService = new WindowsWinXService();

        var runInTray = args.Any(x => string.Equals(x, "--tray", StringComparison.OrdinalIgnoreCase));
        using var form = new MainForm(
            contextMenuProvider,
            fluentModeService,
            explorerService,
            openWithService,
            sendToService,
            winXService,
            runInTray);
        Application.Run(form);
    }
}
