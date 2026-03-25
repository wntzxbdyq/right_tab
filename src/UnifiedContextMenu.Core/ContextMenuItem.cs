namespace UnifiedContextMenu.Core;

public sealed class ContextMenuItem
{
    public required string Name { get; init; }
    public required string RegistryPath { get; init; }
    public required bool Enabled { get; init; }
}
