namespace UnifiedContextMenu.Core;

public sealed class OpenWithAppItem
{
    public required string AppKeyName { get; init; }
    public required string DisplayName { get; init; }
    public required string ExecutablePath { get; init; }
    public required string CommandRegistryPath { get; init; }
    public required bool Visible { get; init; }
}
