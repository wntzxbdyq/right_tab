using System.Diagnostics;
using UnifiedContextMenu.Core;

namespace UnifiedContextMenu.Infrastructure.Windows;

public sealed class WindowsExplorerService : IExplorerService
{
    public void RestartExplorer()
    {
        var processes = Process.GetProcessesByName("explorer");
        foreach (var process in processes)
        {
            process.Kill(entireProcessTree: true);
        }

        Thread.Sleep(700);
        Process.Start(new ProcessStartInfo
        {
            FileName = "explorer.exe",
            UseShellExecute = true
        });
    }
}
