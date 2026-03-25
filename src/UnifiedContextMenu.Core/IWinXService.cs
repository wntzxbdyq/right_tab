namespace UnifiedContextMenu.Core;

public interface IWinXService
{
    IReadOnlyList<string> GetGroups();
    IReadOnlyList<WinXEntryModel> GetEntries();
    string CreateGroup();
    void AddEntry(string groupName, string title, string targetPath, string arguments);
    void Rename(WinXEntryModel item, string newName);
    void SetVisible(WinXEntryModel item, bool visible);
    void Delete(WinXEntryModel item);
    void MoveWithinGroup(WinXEntryModel item, bool moveUp);
}
