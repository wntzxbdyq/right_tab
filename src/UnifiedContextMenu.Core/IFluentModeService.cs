namespace UnifiedContextMenu.Core;

public interface IFluentModeService
{
    FluentModeStatus GetStatus();
    void SetClassicContextMenu(bool enabled);
    void SetAutoStart(bool enabled, string appPath);
}
