namespace UnifiedContextMenu.Core;

public interface IOpenWithService
{
    IReadOnlyList<OpenWithAppItem> GetItems();
    void Add(string executablePath, string displayName);
    void Rename(OpenWithAppItem item, string newDisplayName);
    void SetVisible(OpenWithAppItem item, bool visible);
    void Delete(OpenWithAppItem item);
}
