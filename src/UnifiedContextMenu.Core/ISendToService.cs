namespace UnifiedContextMenu.Core;

public interface ISendToService
{
    string SendToDirectory { get; }
    IReadOnlyList<SendToItemModel> GetItems();
    void AddShortcut(string name, string targetPath, string arguments);
    void Rename(SendToItemModel item, string newName);
    void SetVisible(SendToItemModel item, bool visible);
    void Delete(SendToItemModel item);
}
