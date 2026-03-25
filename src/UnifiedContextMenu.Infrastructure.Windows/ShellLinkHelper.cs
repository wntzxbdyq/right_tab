using System.Reflection;

namespace UnifiedContextMenu.Infrastructure.Windows;

internal sealed class ShellLinkInfo
{
    public required string TargetPath { get; init; }
    public required string Arguments { get; init; }
    public required string Description { get; init; }
}

internal static class ShellLinkHelper
{
    public static ShellLinkInfo Read(string shortcutPath)
    {
        dynamic shortcut = CreateShortcut(shortcutPath);
        return new ShellLinkInfo
        {
            TargetPath = shortcut.TargetPath as string ?? string.Empty,
            Arguments = shortcut.Arguments as string ?? string.Empty,
            Description = shortcut.Description as string ?? string.Empty
        };
    }

    public static void Create(
        string shortcutPath,
        string targetPath,
        string arguments,
        string description,
        string? workingDirectory = null)
    {
        dynamic shortcut = CreateShortcut(shortcutPath);
        shortcut.TargetPath = targetPath;
        shortcut.Arguments = arguments;
        shortcut.Description = description;
        shortcut.WorkingDirectory = workingDirectory ?? Path.GetDirectoryName(targetPath) ?? string.Empty;
        shortcut.Save();
    }

    public static void SetDescription(string shortcutPath, string description)
    {
        dynamic shortcut = CreateShortcut(shortcutPath);
        shortcut.Description = description;
        shortcut.Save();
    }

    private static dynamic CreateShortcut(string shortcutPath)
    {
        var shellType = Type.GetTypeFromProgID("WScript.Shell")
            ?? throw new InvalidOperationException("WScript.Shell 不可用。");
        var shell = Activator.CreateInstance(shellType)
            ?? throw new InvalidOperationException("无法创建 WScript.Shell 实例。");
        return shellType.InvokeMember(
            "CreateShortcut",
            BindingFlags.InvokeMethod,
            binder: null,
            target: shell,
            args: new object[] { shortcutPath });
    }
}
