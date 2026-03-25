using Microsoft.Win32;
using UnifiedContextMenu.Core;

namespace UnifiedContextMenu.Infrastructure.Windows;

public sealed class WindowsFluentModeService : IFluentModeService
{
    private const string ClassicMenuClsid = "{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}";
    private const string RunKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Run";
    private const string AutoStartName = "UnifiedContextMenu";

    public FluentModeStatus GetStatus()
    {
        var classicMenuEnabled = IsClassicContextMenuEnabled();
        var autoStartEnabled = IsAutoStartEnabled();
        return new FluentModeStatus
        {
            ClassicContextMenuEnabled = classicMenuEnabled,
            AutoStartEnabled = autoStartEnabled
        };
    }

    public void SetClassicContextMenu(bool enabled)
    {
        var baseKeyPath = $@"Software\Classes\CLSID\{ClassicMenuClsid}";
        var inProcKeyPath = $@"{baseKeyPath}\InprocServer32";
        if (enabled)
        {
            using var key = Registry.CurrentUser.CreateSubKey(inProcKeyPath, writable: true);
            key?.SetValue(string.Empty, string.Empty, RegistryValueKind.String);
        }
        else
        {
            Registry.CurrentUser.DeleteSubKeyTree(baseKeyPath, throwOnMissingSubKey: false);
        }
    }

    public void SetAutoStart(bool enabled, string appPath)
    {
        using var runKey = Registry.CurrentUser.CreateSubKey(RunKeyPath, writable: true);
        if (runKey is null)
        {
            throw new InvalidOperationException("无法访问开机启动注册表项。");
        }

        if (enabled)
        {
            var value = $"\"{appPath}\" --tray";
            runKey.SetValue(AutoStartName, value, RegistryValueKind.String);
        }
        else
        {
            runKey.DeleteValue(AutoStartName, throwOnMissingValue: false);
        }
    }

    private static bool IsClassicContextMenuEnabled()
    {
        var path = $@"Software\Classes\CLSID\{ClassicMenuClsid}\InprocServer32";
        using var key = Registry.CurrentUser.OpenSubKey(path);
        return key is not null;
    }

    private static bool IsAutoStartEnabled()
    {
        using var key = Registry.CurrentUser.OpenSubKey(RunKeyPath);
        var value = key?.GetValue(AutoStartName) as string;
        return !string.IsNullOrWhiteSpace(value);
    }
}
