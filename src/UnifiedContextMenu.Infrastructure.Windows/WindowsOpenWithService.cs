using Microsoft.Win32;
using UnifiedContextMenu.Core;

namespace UnifiedContextMenu.Infrastructure.Windows;

public sealed class WindowsOpenWithService : IOpenWithService
{
    private const string ApplicationsPath = @"Applications";

    public IReadOnlyList<OpenWithAppItem> GetItems()
    {
        using var appRoot = Registry.ClassesRoot.OpenSubKey(ApplicationsPath);
        if (appRoot is null)
        {
            return Array.Empty<OpenWithAppItem>();
        }

        var items = new List<OpenWithAppItem>();
        foreach (var appKeyName in appRoot.GetSubKeyNames())
        {
            if (!appKeyName.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var commandPath = $@"{ApplicationsPath}\{appKeyName}\shell\open\command";
            using var cmdKey = Registry.ClassesRoot.OpenSubKey(commandPath);
            var command = cmdKey?.GetValue(string.Empty)?.ToString();
            if (string.IsNullOrWhiteSpace(command))
            {
                continue;
            }

            var executablePath = ExtractExecutablePath(command);
            if (string.IsNullOrWhiteSpace(executablePath))
            {
                continue;
            }

            var appPath = $@"{ApplicationsPath}\{appKeyName}";
            using var appKey = Registry.ClassesRoot.OpenSubKey(appPath);
            var friendlyName = appKey?.GetValue("FriendlyAppName")?.ToString();
            var visible = appKey?.GetValue("NoOpenWith") is null;
            var displayName = string.IsNullOrWhiteSpace(friendlyName)
                ? Path.GetFileNameWithoutExtension(appKeyName)
                : friendlyName;

            items.Add(new OpenWithAppItem
            {
                AppKeyName = appKeyName,
                DisplayName = displayName,
                ExecutablePath = executablePath,
                CommandRegistryPath = $@"HKEY_CLASSES_ROOT\{commandPath}",
                Visible = visible
            });
        }

        return items.OrderBy(x => x.DisplayName, StringComparer.OrdinalIgnoreCase).ToList();
    }

    public void Add(string executablePath, string displayName)
    {
        if (!File.Exists(executablePath))
        {
            throw new FileNotFoundException("可执行文件不存在。", executablePath);
        }
        var appKeyName = Path.GetFileName(executablePath);
        if (string.IsNullOrWhiteSpace(appKeyName) || !appKeyName.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException("仅支持 exe 程序。");
        }

        var appPath = $@"{ApplicationsPath}\{appKeyName}";
        using (var appKey = Registry.ClassesRoot.CreateSubKey(appPath, writable: true))
        {
            appKey?.SetValue("FriendlyAppName", displayName, RegistryValueKind.String);
        }
        using (var commandKey = Registry.ClassesRoot.CreateSubKey($@"{appPath}\shell\open\command", writable: true))
        {
            var command = $"\"{executablePath}\" \"%1\"";
            commandKey?.SetValue(string.Empty, command, RegistryValueKind.String);
        }
    }

    public void Rename(OpenWithAppItem item, string newDisplayName)
    {
        var appPath = $@"{ApplicationsPath}\{item.AppKeyName}";
        using var appKey = Registry.ClassesRoot.CreateSubKey(appPath, writable: true);
        appKey?.SetValue("FriendlyAppName", newDisplayName, RegistryValueKind.String);
    }

    public void SetVisible(OpenWithAppItem item, bool visible)
    {
        var appPath = $@"{ApplicationsPath}\{item.AppKeyName}";
        using var appKey = Registry.ClassesRoot.CreateSubKey(appPath, writable: true);
        if (appKey is null)
        {
            throw new InvalidOperationException("无法访问 OpenWith 注册表项。");
        }
        if (visible)
        {
            appKey.DeleteValue("NoOpenWith", false);
        }
        else
        {
            appKey.SetValue("NoOpenWith", string.Empty, RegistryValueKind.String);
        }
    }

    public void Delete(OpenWithAppItem item)
    {
        var commandPath = $@"{ApplicationsPath}\{item.AppKeyName}\shell\open\command";
        Registry.ClassesRoot.DeleteSubKeyTree(commandPath, throwOnMissingSubKey: false);

        using var shellKey = Registry.ClassesRoot.OpenSubKey($@"{ApplicationsPath}\{item.AppKeyName}\shell");
        if (shellKey is null || shellKey.SubKeyCount == 0)
        {
            Registry.ClassesRoot.DeleteSubKeyTree($@"{ApplicationsPath}\{item.AppKeyName}", throwOnMissingSubKey: false);
        }
    }

    private static string? ExtractExecutablePath(string command)
    {
        var text = command.Trim();
        if (text.StartsWith('"'))
        {
            var end = text.IndexOf('"', 1);
            if (end > 1)
            {
                return text[1..end];
            }
        }
        var index = text.IndexOf(".exe", StringComparison.OrdinalIgnoreCase);
        if (index <= 0)
        {
            return null;
        }
        return text[..(index + 4)];
    }
}
