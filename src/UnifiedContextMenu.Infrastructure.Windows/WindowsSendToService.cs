using UnifiedContextMenu.Core;

namespace UnifiedContextMenu.Infrastructure.Windows;

public sealed class WindowsSendToService : ISendToService
{
    public string SendToDirectory { get; } = Environment.ExpandEnvironmentVariables(
        @"%AppData%\Microsoft\Windows\SendTo");

    public IReadOnlyList<SendToItemModel> GetItems()
    {
        if (!Directory.Exists(SendToDirectory))
        {
            return Array.Empty<SendToItemModel>();
        }

        var items = new List<SendToItemModel>();
        foreach (var path in Directory.GetFileSystemEntries(SendToDirectory))
        {
            var fileName = Path.GetFileName(path);
            if (string.Equals(fileName, "desktop.ini", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var isShortcut = string.Equals(Path.GetExtension(path), ".lnk", StringComparison.OrdinalIgnoreCase);
            var targetPath = path;
            if (isShortcut)
            {
                var link = ShellLinkHelper.Read(path);
                if (!string.IsNullOrWhiteSpace(link.TargetPath))
                {
                    targetPath = link.TargetPath;
                }
            }

            var attrs = File.GetAttributes(path);
            var visible = (attrs & FileAttributes.Hidden) != FileAttributes.Hidden;
            items.Add(new SendToItemModel
            {
                Name = Path.GetFileNameWithoutExtension(path),
                FilePath = path,
                TargetPath = targetPath,
                IsShortcut = isShortcut,
                Visible = visible
            });
        }

        return items.OrderBy(x => x.Name, StringComparer.OrdinalIgnoreCase).ToList();
    }

    public void AddShortcut(string name, string targetPath, string arguments)
    {
        if (!File.Exists(targetPath) && !Directory.Exists(targetPath))
        {
            throw new FileNotFoundException("目标路径不存在。", targetPath);
        }

        Directory.CreateDirectory(SendToDirectory);
        var rawName = SanitizeFileName(name);
        var shortcutPath = GetUniquePath(Path.Combine(SendToDirectory, $"{rawName}.lnk"));
        ShellLinkHelper.Create(shortcutPath, targetPath, arguments, name);
    }

    public void Rename(SendToItemModel item, string newName)
    {
        var extension = Path.GetExtension(item.FilePath);
        var newPath = Path.Combine(Path.GetDirectoryName(item.FilePath)!, $"{SanitizeFileName(newName)}{extension}");
        newPath = GetUniquePath(newPath);
        File.Move(item.FilePath, newPath);
    }

    public void SetVisible(SendToItemModel item, bool visible)
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

    public void Delete(SendToItemModel item)
    {
        File.Delete(item.FilePath);
    }

    private static string SanitizeFileName(string input)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var chars = input.Select(c => invalid.Contains(c) ? '_' : c).ToArray();
        var result = new string(chars).Trim();
        return string.IsNullOrWhiteSpace(result) ? "NewItem" : result;
    }

    private static string GetUniquePath(string path)
    {
        if (!File.Exists(path) && !Directory.Exists(path))
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
            if (!File.Exists(candidate) && !Directory.Exists(candidate))
            {
                return candidate;
            }
            index++;
        }
    }
}
