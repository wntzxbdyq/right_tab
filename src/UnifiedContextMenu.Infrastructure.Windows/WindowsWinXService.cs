using UnifiedContextMenu.Core;

namespace UnifiedContextMenu.Infrastructure.Windows;

public sealed class WindowsWinXService : IWinXService
{
    private readonly string _winXPath = Environment.ExpandEnvironmentVariables(@"%LocalAppData%\Microsoft\Windows\WinX");

    public IReadOnlyList<string> GetGroups()
    {
        if (!Directory.Exists(_winXPath))
        {
            return Array.Empty<string>();
        }

        return Directory.GetDirectories(_winXPath)
            .Select(Path.GetFileName)
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .Cast<string>()
            .OrderByDescending(x => x, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    public IReadOnlyList<WinXEntryModel> GetEntries()
    {
        if (!Directory.Exists(_winXPath))
        {
            return Array.Empty<WinXEntryModel>();
        }

        var entries = new List<WinXEntryModel>();
        foreach (var group in GetGroups())
        {
            var groupPath = Path.Combine(_winXPath, group);
            foreach (var lnkPath in Directory.GetFiles(groupPath, "*.lnk").OrderByDescending(Path.GetFileName))
            {
                var link = ShellLinkHelper.Read(lnkPath);
                var attrs = File.GetAttributes(lnkPath);
                var visible = (attrs & FileAttributes.Hidden) != FileAttributes.Hidden;
                var displayName = string.IsNullOrWhiteSpace(link.Description)
                    ? Path.GetFileNameWithoutExtension(lnkPath)
                    : link.Description;
                entries.Add(new WinXEntryModel
                {
                    GroupName = group,
                    FilePath = lnkPath,
                    Name = displayName,
                    TargetPath = string.IsNullOrWhiteSpace(link.TargetPath) ? lnkPath : link.TargetPath,
                    Visible = visible
                });
            }
        }

        return entries;
    }

    public string CreateGroup()
    {
        Directory.CreateDirectory(_winXPath);
        var index = 1;
        while (Directory.Exists(Path.Combine(_winXPath, $"Group{index}")))
        {
            index++;
        }

        var dirPath = Path.Combine(_winXPath, $"Group{index}");
        Directory.CreateDirectory(dirPath);
        var iniPath = Path.Combine(dirPath, "desktop.ini");
        File.WriteAllText(iniPath, string.Empty);
        File.SetAttributes(dirPath, File.GetAttributes(dirPath) | FileAttributes.ReadOnly);
        File.SetAttributes(iniPath, File.GetAttributes(iniPath) | FileAttributes.Hidden | FileAttributes.System);
        return Path.GetFileName(dirPath);
    }

    public void AddEntry(string groupName, string title, string targetPath, string arguments)
    {
        var groupPath = Path.Combine(_winXPath, groupName);
        if (!Directory.Exists(groupPath))
        {
            throw new DirectoryNotFoundException($"WinX 分组不存在: {groupName}");
        }

        var count = Directory.GetFiles(groupPath, "*.lnk").Length;
        var prefix = (count + 1).ToString("D2");
        var baseName = Path.GetFileNameWithoutExtension(targetPath);
        var fileName = $"{prefix} - {SanitizeFileName(baseName)}.lnk";
        var shortcutPath = GetUniquePath(Path.Combine(groupPath, fileName));

        ShellLinkHelper.Create(shortcutPath, targetPath, arguments, title);
        WinXHasher.HashShortcut(shortcutPath);
    }

    public void Rename(WinXEntryModel item, string newName)
    {
        ShellLinkHelper.SetDescription(item.FilePath, newName);
        WinXHasher.HashShortcut(item.FilePath);
    }

    public void SetVisible(WinXEntryModel item, bool visible)
    {
        var attrs = File.GetAttributes(item.FilePath);
        if (visible)
        {
            attrs &= ~FileAttributes.Hidden;
        }
        else
        {
            attrs |= FileAttributes.Hidden;
        }
        File.SetAttributes(item.FilePath, attrs);
    }

    public void Delete(WinXEntryModel item)
    {
        File.Delete(item.FilePath);
    }

    public void MoveWithinGroup(WinXEntryModel item, bool moveUp)
    {
        var groupPath = Path.Combine(_winXPath, item.GroupName);
        var files = Directory.GetFiles(groupPath, "*.lnk")
            .OrderBy(Path.GetFileName, StringComparer.OrdinalIgnoreCase)
            .ToList();

        var index = files.FindIndex(x => string.Equals(x, item.FilePath, StringComparison.OrdinalIgnoreCase));
        if (index < 0)
        {
            return;
        }

        var targetIndex = moveUp ? index - 1 : index + 1;
        if (targetIndex < 0 || targetIndex >= files.Count)
        {
            return;
        }

        (files[index], files[targetIndex]) = (files[targetIndex], files[index]);
        RenameGroupOrder(groupPath, files);
    }

    private static void RenameGroupOrder(string groupPath, IReadOnlyList<string> orderedPaths)
    {
        var tempPaths = new List<string>(orderedPaths.Count);
        foreach (var path in orderedPaths)
        {
            var tmp = Path.Combine(groupPath, Guid.NewGuid() + ".tmp");
            File.Move(path, tmp);
            tempPaths.Add(tmp);
        }

        for (var i = 0; i < tempPaths.Count; i++)
        {
            var originalName = Path.GetFileNameWithoutExtension(orderedPaths[i]);
            var dash = originalName.IndexOf(" - ", StringComparison.Ordinal);
            var suffix = dash >= 0 ? originalName[(dash + 3)..] : originalName;
            var finalName = $"{i + 1:D2} - {suffix}.lnk";
            var finalPath = Path.Combine(groupPath, finalName);
            File.Move(tempPaths[i], finalPath);
            WinXHasher.HashShortcut(finalPath);
        }
    }

    private static string SanitizeFileName(string input)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var chars = input.Select(c => invalid.Contains(c) ? '_' : c).ToArray();
        var result = new string(chars).Trim();
        return string.IsNullOrWhiteSpace(result) ? "Item" : result;
    }

    private static string GetUniquePath(string path)
    {
        if (!File.Exists(path))
        {
            return path;
        }

        var dir = Path.GetDirectoryName(path)!;
        var name = Path.GetFileNameWithoutExtension(path);
        var ext = Path.GetExtension(path);
        var index = 1;
        while (true)
        {
            var candidate = Path.Combine(dir, $"{name} ({index}){ext}");
            if (!File.Exists(candidate))
            {
                return candidate;
            }
            index++;
        }
    }
}
