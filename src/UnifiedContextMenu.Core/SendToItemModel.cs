namespace UnifiedContextMenu.Core;

public sealed class SendToItemModel
{
    public required string Name { get; init; }
    public required string FilePath { get; init; }
    public required string TargetPath { get; init; }
    public required bool IsShortcut { get; init; }
    public required bool Visible { get; init; }
}
