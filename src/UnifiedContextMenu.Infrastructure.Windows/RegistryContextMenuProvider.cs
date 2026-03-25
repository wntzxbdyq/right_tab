using Microsoft.Win32;
using UnifiedContextMenu.Core;

namespace UnifiedContextMenu.Infrastructure.Windows;

public sealed class RegistryContextMenuProvider : IContextMenuProvider
{
    private const string LegacyDisable = "LegacyDisable";

    private static readonly IReadOnlyDictionary<ContextMenuScene, string[]> SceneRegistryPaths =
        new Dictionary<ContextMenuScene, string[]>
        {
            [ContextMenuScene.File] = new[] { @"*\shell", @"*\shellex\ContextMenuHandlers" },
            [ContextMenuScene.Folder] = new[] { @"Folder\shell", @"Folder\shellex\ContextMenuHandlers" },
            [ContextMenuScene.Directory] = new[] { @"Directory\shell", @"Directory\shellex\ContextMenuHandlers" },
            [ContextMenuScene.Background] = new[] { @"Directory\Background\shell", @"Directory\Background\shellex\ContextMenuHandlers" },
            [ContextMenuScene.Desktop] = new[] { @"DesktopBackground\shell", @"DesktopBackground\shellex\ContextMenuHandlers" },
            [ContextMenuScene.Drive] = new[] { @"Drive\shell", @"Drive\shellex\ContextMenuHandlers" },
            [ContextMenuScene.AllObjects] = new[] { @"AllFilesystemObjects\shell", @"AllFilesystemObjects\shellex\ContextMenuHandlers" }
        };

    public IReadOnlyList<ContextMenuItem> GetItems(ContextMenuScene scene)
    {
        if (!SceneRegistryPaths.TryGetValue(scene, out var roots))
        {
            return Array.Empty<ContextMenuItem>();
        }

        var items = new List<ContextMenuItem>();
        foreach (var root in roots)
        {
            EnumerateBranch(Registry.CurrentUser, $@"Software\Classes\{root}", items);
            EnumerateBranch(Registry.LocalMachine, $@"Software\Classes\{root}", items);
        }

        return items
            .GroupBy(x => x.RegistryPath, StringComparer.OrdinalIgnoreCase)
            .Select(g => g.First())
            .OrderBy(x => x.Name, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    public void SetEnabled(ContextMenuItem item, bool enabled)
    {
        var hiveAndPath = ResolveRoot(item.RegistryPath);
        if (hiveAndPath is null)
        {
            throw new InvalidOperationException($"无法识别注册表路径: {item.RegistryPath}");
        }

        var (root, subKeyPath) = hiveAndPath.Value;
        using var key = root.OpenSubKey(subKeyPath, writable: true);
        if (key is null)
        {
            throw new InvalidOperationException($"注册表项不存在或不可写: {item.RegistryPath}");
        }

        if (enabled)
        {
            if (key.GetValue(LegacyDisable) is not null)
            {
                key.DeleteValue(LegacyDisable, throwOnMissingValue: false);
            }
        }
        else
        {
            key.SetValue(LegacyDisable, string.Empty, RegistryValueKind.String);
        }
    }

    private static void EnumerateBranch(RegistryKey root, string path, ICollection<ContextMenuItem> items)
    {
        using var branch = root.OpenSubKey(path);
        if (branch is null)
        {
            return;
        }

        foreach (var subName in branch.GetSubKeyNames())
        {
            using var itemKey = branch.OpenSubKey(subName);
            if (itemKey is null)
            {
                continue;
            }

            var fullPath = $"{ToHiveName(root)}\\{path}\\{subName}";
            var disabled = itemKey.GetValue(LegacyDisable) is not null;
            var displayName = BuildDisplayName(subName, itemKey);
            items.Add(new ContextMenuItem
            {
                Name = displayName,
                RegistryPath = fullPath,
                Enabled = !disabled
            });
        }
    }

    private static string BuildDisplayName(string subName, RegistryKey key)
    {
        var muiVerb = key.GetValue("MUIVerb") as string;
        var defaultValue = key.GetValue(string.Empty) as string;
        if (!string.IsNullOrWhiteSpace(muiVerb))
        {
            return muiVerb;
        }
        if (!string.IsNullOrWhiteSpace(defaultValue))
        {
            return defaultValue;
        }
        return subName;
    }

    private static (RegistryKey root, string subKeyPath)? ResolveRoot(string fullPath)
    {
        const string hkcu = "HKEY_CURRENT_USER\\";
        const string hklm = "HKEY_LOCAL_MACHINE\\";

        if (fullPath.StartsWith(hkcu, StringComparison.OrdinalIgnoreCase))
        {
            return (Registry.CurrentUser, fullPath[hkcu.Length..]);
        }
        if (fullPath.StartsWith(hklm, StringComparison.OrdinalIgnoreCase))
        {
            return (Registry.LocalMachine, fullPath[hklm.Length..]);
        }
        return null;
    }

    private static string ToHiveName(RegistryKey key)
    {
        if (ReferenceEquals(key, Registry.CurrentUser))
        {
            return "HKEY_CURRENT_USER";
        }
        if (ReferenceEquals(key, Registry.LocalMachine))
        {
            return "HKEY_LOCAL_MACHINE";
        }
        throw new NotSupportedException("不支持的注册表根节点。");
    }
}
