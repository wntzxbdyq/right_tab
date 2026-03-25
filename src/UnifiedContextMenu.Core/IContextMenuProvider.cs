namespace UnifiedContextMenu.Core;

public interface IContextMenuProvider
{
    IReadOnlyList<ContextMenuItem> GetItems(ContextMenuScene scene);
    void SetEnabled(ContextMenuItem item, bool enabled);
}
