using UnifiedContextMenu.Core;

namespace UnifiedContextMenu.App.WinForms;

public sealed class MainForm : Form
{
    private readonly IContextMenuProvider _contextMenuProvider;
    private readonly IFluentModeService _fluentModeService;
    private readonly IExplorerService _explorerService;
    private readonly IOpenWithService _openWithService;
    private readonly ISendToService _sendToService;
    private readonly IWinXService _winXService;

    private readonly ComboBox _sceneCombo = new();
    private readonly ListView _itemList = new();
    private readonly CheckBox _classicMenuCheck = new();
    private readonly CheckBox _autoStartCheck = new();
    private readonly Button _applyFluentButton = new();
    private readonly NotifyIcon _notifyIcon = new();
    private readonly ContextMenuStrip _trayMenu = new();
    private readonly Label _hintLabel = new();
    private readonly ListView _openWithList = new();
    private readonly ListView _sendToList = new();
    private readonly ListView _winXList = new();

    public MainForm(
        IContextMenuProvider contextMenuProvider,
        IFluentModeService fluentModeService,
        IExplorerService explorerService,
        IOpenWithService openWithService,
        ISendToService sendToService,
        IWinXService winXService,
        bool startInTray)
    {
        _contextMenuProvider = contextMenuProvider;
        _fluentModeService = fluentModeService;
        _explorerService = explorerService;
        _openWithService = openWithService;
        _sendToService = sendToService;
        _winXService = winXService;

        InitializeComponent();
        LoadScenes();
        LoadFluentStatus();
        ReloadOpenWith();
        ReloadSendTo();
        ReloadWinX();

        if (startInTray)
        {
            BeginInvoke(() =>
            {
                Hide();
                ShowInTaskbar = false;
                _notifyIcon.Visible = true;
            });
        }
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _notifyIcon.Dispose();
            _trayMenu.Dispose();
        }
        base.Dispose(disposing);
    }

    private void InitializeComponent()
    {
        Text = "Unified Context Menu";
        Width = 1100;
        Height = 700;
        StartPosition = FormStartPosition.CenterScreen;
        MinimizeBox = true;
        MaximizeBox = true;

        var tabControl = new TabControl
        {
            Dock = DockStyle.Fill
        };

        var managerPage = new TabPage("菜单管理");
        var fluentPage = new TabPage("Win11增强");
        var openWithPage = new TabPage("OpenWith");
        var sendToPage = new TabPage("SendTo");
        var winXPage = new TabPage("WinX");

        tabControl.TabPages.Add(managerPage);
        tabControl.TabPages.Add(fluentPage);
        tabControl.TabPages.Add(openWithPage);
        tabControl.TabPages.Add(sendToPage);
        tabControl.TabPages.Add(winXPage);
        Controls.Add(tabControl);

        BuildManagerPage(managerPage);
        BuildFluentPage(fluentPage);
        BuildOpenWithPage(openWithPage);
        BuildSendToPage(sendToPage);
        BuildWinXPage(winXPage);
        BuildTrayMenu();
    }

    private void BuildManagerPage(TabPage page)
    {
        var topPanel = new Panel { Dock = DockStyle.Top, Height = 48 };
        var contentPanel = new Panel { Dock = DockStyle.Fill };
        page.Controls.Add(contentPanel);
        page.Controls.Add(topPanel);

        var sceneLabel = new Label
        {
            Text = "场景",
            AutoSize = true,
            Left = 12,
            Top = 16
        };
        topPanel.Controls.Add(sceneLabel);

        _sceneCombo.Left = 56;
        _sceneCombo.Top = 10;
        _sceneCombo.Width = 220;
        _sceneCombo.DropDownStyle = ComboBoxStyle.DropDownList;
        _sceneCombo.SelectedIndexChanged += (_, _) => ReloadSceneItems();
        topPanel.Controls.Add(_sceneCombo);

        var refreshButton = new Button
        {
            Text = "刷新",
            Left = 288,
            Top = 9,
            Width = 90
        };
        refreshButton.Click += (_, _) => ReloadSceneItems();
        topPanel.Controls.Add(refreshButton);

        var restartExplorerButton = new Button
        {
            Text = "重启Explorer",
            Left = 388,
            Top = 9,
            Width = 130
        };
        restartExplorerButton.Click += (_, _) =>
        {
            _explorerService.RestartExplorer();
            MessageBox.Show(this, "Explorer 已重启。", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
        };
        topPanel.Controls.Add(restartExplorerButton);

        _itemList.Dock = DockStyle.Fill;
        _itemList.View = View.Details;
        _itemList.FullRowSelect = true;
        _itemList.CheckBoxes = true;
        _itemList.Columns.Add("菜单名", 260);
        _itemList.Columns.Add("注册表路径", 620);
        _itemList.ItemChecked += OnItemChecked;
        contentPanel.Controls.Add(_itemList);
    }

    private void BuildFluentPage(TabPage page)
    {
        var panel = new Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(20)
        };
        page.Controls.Add(panel);

        _classicMenuCheck.Text = "关闭系统 WinUI 菜单（启用经典菜单）";
        _classicMenuCheck.AutoSize = true;
        _classicMenuCheck.Top = 20;
        _classicMenuCheck.Left = 24;
        panel.Controls.Add(_classicMenuCheck);

        _autoStartCheck.Text = "开机自启动（托盘模式）";
        _autoStartCheck.AutoSize = true;
        _autoStartCheck.Top = 56;
        _autoStartCheck.Left = 24;
        panel.Controls.Add(_autoStartCheck);

        _applyFluentButton.Text = "应用并重启 Explorer";
        _applyFluentButton.Top = 96;
        _applyFluentButton.Left = 24;
        _applyFluentButton.Width = 220;
        _applyFluentButton.Click += (_, _) => ApplyFluentSettings();
        panel.Controls.Add(_applyFluentButton);

        _hintLabel.Text =
            "说明：\r\n" +
            "- 经典菜单开关来自 FluentContextMenu 的核心行为。\r\n" +
            "- 右键菜单项管理来自 ContextMenuManager 的场景化思路。\r\n" +
            "- 统一在一个程序中维护。";
        _hintLabel.AutoSize = true;
        _hintLabel.Top = 150;
        _hintLabel.Left = 24;
        panel.Controls.Add(_hintLabel);
    }

    private void BuildOpenWithPage(TabPage page)
    {
        var topPanel = new Panel { Dock = DockStyle.Top, Height = 48 };
        var contentPanel = new Panel { Dock = DockStyle.Fill };
        page.Controls.Add(contentPanel);
        page.Controls.Add(topPanel);

        var addButton = new Button { Text = "添加程序", Left = 12, Top = 9, Width = 100 };
        addButton.Click += (_, _) => AddOpenWith();
        topPanel.Controls.Add(addButton);

        var renameButton = new Button { Text = "重命名", Left = 122, Top = 9, Width = 90 };
        renameButton.Click += (_, _) => RenameOpenWith();
        topPanel.Controls.Add(renameButton);

        var deleteButton = new Button { Text = "删除", Left = 222, Top = 9, Width = 90 };
        deleteButton.Click += (_, _) => DeleteOpenWith();
        topPanel.Controls.Add(deleteButton);

        var refreshButton = new Button { Text = "刷新", Left = 322, Top = 9, Width = 90 };
        refreshButton.Click += (_, _) => ReloadOpenWith();
        topPanel.Controls.Add(refreshButton);

        _openWithList.Dock = DockStyle.Fill;
        _openWithList.View = View.Details;
        _openWithList.FullRowSelect = true;
        _openWithList.CheckBoxes = true;
        _openWithList.Columns.Add("显示名", 220);
        _openWithList.Columns.Add("程序路径", 500);
        _openWithList.Columns.Add("注册表", 360);
        _openWithList.ItemChecked += OnOpenWithChecked;
        contentPanel.Controls.Add(_openWithList);
    }

    private void BuildSendToPage(TabPage page)
    {
        var topPanel = new Panel { Dock = DockStyle.Top, Height = 48 };
        var contentPanel = new Panel { Dock = DockStyle.Fill };
        page.Controls.Add(contentPanel);
        page.Controls.Add(topPanel);

        var addButton = new Button { Text = "添加快捷方式", Left = 12, Top = 9, Width = 120 };
        addButton.Click += (_, _) => AddSendTo();
        topPanel.Controls.Add(addButton);

        var renameButton = new Button { Text = "重命名", Left = 142, Top = 9, Width = 90 };
        renameButton.Click += (_, _) => RenameSendTo();
        topPanel.Controls.Add(renameButton);

        var deleteButton = new Button { Text = "删除", Left = 242, Top = 9, Width = 90 };
        deleteButton.Click += (_, _) => DeleteSendTo();
        topPanel.Controls.Add(deleteButton);

        var folderButton = new Button { Text = "打开目录", Left = 342, Top = 9, Width = 90 };
        folderButton.Click += (_, _) => OpenSendToDirectory();
        topPanel.Controls.Add(folderButton);

        var refreshButton = new Button { Text = "刷新", Left = 442, Top = 9, Width = 90 };
        refreshButton.Click += (_, _) => ReloadSendTo();
        topPanel.Controls.Add(refreshButton);

        _sendToList.Dock = DockStyle.Fill;
        _sendToList.View = View.Details;
        _sendToList.FullRowSelect = true;
        _sendToList.CheckBoxes = true;
        _sendToList.Columns.Add("名称", 240);
        _sendToList.Columns.Add("目标", 500);
        _sendToList.Columns.Add("文件", 340);
        _sendToList.ItemChecked += OnSendToChecked;
        contentPanel.Controls.Add(_sendToList);
    }

    private void BuildWinXPage(TabPage page)
    {
        var topPanel = new Panel { Dock = DockStyle.Top, Height = 48 };
        var contentPanel = new Panel { Dock = DockStyle.Fill };
        page.Controls.Add(contentPanel);
        page.Controls.Add(topPanel);

        var addGroupButton = new Button { Text = "新建分组", Left = 12, Top = 9, Width = 90 };
        addGroupButton.Click += (_, _) => AddWinXGroup();
        topPanel.Controls.Add(addGroupButton);

        var addButton = new Button { Text = "添加项目", Left = 112, Top = 9, Width = 90 };
        addButton.Click += (_, _) => AddWinXItem();
        topPanel.Controls.Add(addButton);

        var renameButton = new Button { Text = "重命名", Left = 212, Top = 9, Width = 90 };
        renameButton.Click += (_, _) => RenameWinXItem();
        topPanel.Controls.Add(renameButton);

        var moveUpButton = new Button { Text = "上移", Left = 312, Top = 9, Width = 70 };
        moveUpButton.Click += (_, _) => MoveWinXItem(true);
        topPanel.Controls.Add(moveUpButton);

        var moveDownButton = new Button { Text = "下移", Left = 392, Top = 9, Width = 70 };
        moveDownButton.Click += (_, _) => MoveWinXItem(false);
        topPanel.Controls.Add(moveDownButton);

        var deleteButton = new Button { Text = "删除", Left = 472, Top = 9, Width = 80 };
        deleteButton.Click += (_, _) => DeleteWinXItem();
        topPanel.Controls.Add(deleteButton);

        var refreshButton = new Button { Text = "刷新", Left = 562, Top = 9, Width = 80 };
        refreshButton.Click += (_, _) => ReloadWinX();
        topPanel.Controls.Add(refreshButton);

        var restartButton = new Button { Text = "重启Explorer", Left = 652, Top = 9, Width = 120 };
        restartButton.Click += (_, _) =>
        {
            _explorerService.RestartExplorer();
            MessageBox.Show(this, "Explorer 已重启。", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
        };
        topPanel.Controls.Add(restartButton);

        _winXList.Dock = DockStyle.Fill;
        _winXList.View = View.Details;
        _winXList.FullRowSelect = true;
        _winXList.CheckBoxes = true;
        _winXList.Columns.Add("分组", 120);
        _winXList.Columns.Add("名称", 220);
        _winXList.Columns.Add("目标", 420);
        _winXList.Columns.Add("文件", 320);
        _winXList.ItemChecked += OnWinXChecked;
        contentPanel.Controls.Add(_winXList);
    }

    private void BuildTrayMenu()
    {
        _notifyIcon.Text = "Unified Context Menu";
        _notifyIcon.Icon = SystemIcons.Application;
        _notifyIcon.DoubleClick += (_, _) => RestoreFromTray();

        var openMenu = new ToolStripMenuItem("打开主窗口");
        openMenu.Click += (_, _) => RestoreFromTray();
        _trayMenu.Items.Add(openMenu);

        var restartMenu = new ToolStripMenuItem("重启 Explorer");
        restartMenu.Click += (_, _) => _explorerService.RestartExplorer();
        _trayMenu.Items.Add(restartMenu);

        _trayMenu.Items.Add(new ToolStripSeparator());
        var exitMenu = new ToolStripMenuItem("退出");
        exitMenu.Click += (_, _) =>
        {
            _notifyIcon.Visible = false;
            Application.Exit();
        };
        _trayMenu.Items.Add(exitMenu);

        _notifyIcon.ContextMenuStrip = _trayMenu;
        Resize += (_, _) =>
        {
            if (WindowState == FormWindowState.Minimized)
            {
                Hide();
                ShowInTaskbar = false;
                _notifyIcon.Visible = true;
            }
        };
    }

    private void RestoreFromTray()
    {
        Show();
        ShowInTaskbar = true;
        WindowState = FormWindowState.Normal;
        Activate();
        _notifyIcon.Visible = false;
    }

    private void LoadScenes()
    {
        _sceneCombo.DataSource = Enum.GetValues<ContextMenuScene>();
        _sceneCombo.SelectedItem = ContextMenuScene.File;
    }

    private void ReloadSceneItems()
    {
        if (_sceneCombo.SelectedItem is not ContextMenuScene scene)
        {
            return;
        }

        _itemList.ItemChecked -= OnItemChecked;
        try
        {
            _itemList.Items.Clear();
            var items = _contextMenuProvider.GetItems(scene);
            foreach (var item in items)
            {
                var row = new ListViewItem(item.Name)
                {
                    Checked = item.Enabled,
                    Tag = item
                };
                row.SubItems.Add(item.RegistryPath);
                _itemList.Items.Add(row);
            }
        }
        finally
        {
            _itemList.ItemChecked += OnItemChecked;
        }
    }

    private void OnItemChecked(object? sender, ItemCheckedEventArgs e)
    {
        if (e.Item.Tag is not ContextMenuItem item)
        {
            return;
        }

        try
        {
            _contextMenuProvider.SetEnabled(item, e.Item.Checked);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "设置失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void LoadFluentStatus()
    {
        var status = _fluentModeService.GetStatus();
        _classicMenuCheck.Checked = status.ClassicContextMenuEnabled;
        _autoStartCheck.Checked = status.AutoStartEnabled;
    }

    private void ApplyFluentSettings()
    {
        _fluentModeService.SetClassicContextMenu(_classicMenuCheck.Checked);
        _fluentModeService.SetAutoStart(_autoStartCheck.Checked, Application.ExecutablePath);
        _explorerService.RestartExplorer();
        LoadFluentStatus();
        MessageBox.Show(this, "设置已应用。", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private void ReloadOpenWith()
    {
        _openWithList.ItemChecked -= OnOpenWithChecked;
        try
        {
            _openWithList.Items.Clear();
            foreach (var item in _openWithService.GetItems())
            {
                var row = new ListViewItem(item.DisplayName) { Tag = item, Checked = item.Visible };
                row.SubItems.Add(item.ExecutablePath);
                row.SubItems.Add(item.CommandRegistryPath);
                _openWithList.Items.Add(row);
            }
        }
        finally
        {
            _openWithList.ItemChecked += OnOpenWithChecked;
        }
    }

    private void OnOpenWithChecked(object? sender, ItemCheckedEventArgs e)
    {
        if (e.Item.Tag is not OpenWithAppItem item)
        {
            return;
        }

        try
        {
            _openWithService.SetVisible(item, e.Item.Checked);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "OpenWith 设置失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void AddOpenWith()
    {
        using var dialog = new OpenFileDialog
        {
            Filter = "Program|*.exe"
        };
        if (dialog.ShowDialog(this) != DialogResult.OK)
        {
            return;
        }

        var defaultName = Path.GetFileNameWithoutExtension(dialog.FileName);
        var name = PromptDialog.Show(this, "OpenWith 显示名", defaultName);
        if (string.IsNullOrWhiteSpace(name))
        {
            return;
        }

        try
        {
            _openWithService.Add(dialog.FileName, name);
            ReloadOpenWith();
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "添加失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void RenameOpenWith()
    {
        if (_openWithList.SelectedItems.Count == 0 || _openWithList.SelectedItems[0].Tag is not OpenWithAppItem item)
        {
            return;
        }
        var newName = PromptDialog.Show(this, "新显示名", item.DisplayName);
        if (string.IsNullOrWhiteSpace(newName))
        {
            return;
        }
        try
        {
            _openWithService.Rename(item, newName);
            ReloadOpenWith();
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "重命名失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void DeleteOpenWith()
    {
        if (_openWithList.SelectedItems.Count == 0 || _openWithList.SelectedItems[0].Tag is not OpenWithAppItem item)
        {
            return;
        }
        var confirm = MessageBox.Show(this, $"确认删除 {item.DisplayName}？", "删除确认", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
        if (confirm != DialogResult.OK)
        {
            return;
        }
        try
        {
            _openWithService.Delete(item);
            ReloadOpenWith();
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "删除失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void ReloadSendTo()
    {
        _sendToList.ItemChecked -= OnSendToChecked;
        try
        {
            _sendToList.Items.Clear();
            foreach (var item in _sendToService.GetItems())
            {
                var row = new ListViewItem(item.Name) { Tag = item, Checked = item.Visible };
                row.SubItems.Add(item.TargetPath);
                row.SubItems.Add(item.FilePath);
                _sendToList.Items.Add(row);
            }
        }
        finally
        {
            _sendToList.ItemChecked += OnSendToChecked;
        }
    }

    private void OnSendToChecked(object? sender, ItemCheckedEventArgs e)
    {
        if (e.Item.Tag is not SendToItemModel item)
        {
            return;
        }

        try
        {
            _sendToService.SetVisible(item, e.Item.Checked);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "SendTo 设置失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void AddSendTo()
    {
        using var dialog = new OpenFileDialog
        {
            Filter = "Program|*.exe;*.bat;*.cmd;*.ps1|All Files|*.*"
        };
        if (dialog.ShowDialog(this) != DialogResult.OK)
        {
            return;
        }

        var name = PromptDialog.Show(this, "SendTo 名称", Path.GetFileNameWithoutExtension(dialog.FileName));
        if (string.IsNullOrWhiteSpace(name))
        {
            return;
        }
        var args = PromptDialog.Show(this, "启动参数（可空）", string.Empty) ?? string.Empty;

        try
        {
            _sendToService.AddShortcut(name, dialog.FileName, args);
            ReloadSendTo();
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "添加失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void RenameSendTo()
    {
        if (_sendToList.SelectedItems.Count == 0 || _sendToList.SelectedItems[0].Tag is not SendToItemModel item)
        {
            return;
        }
        var newName = PromptDialog.Show(this, "新名称", item.Name);
        if (string.IsNullOrWhiteSpace(newName))
        {
            return;
        }
        try
        {
            _sendToService.Rename(item, newName);
            ReloadSendTo();
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "重命名失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void DeleteSendTo()
    {
        if (_sendToList.SelectedItems.Count == 0 || _sendToList.SelectedItems[0].Tag is not SendToItemModel item)
        {
            return;
        }
        var confirm = MessageBox.Show(this, $"确认删除 {item.Name}？", "删除确认", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
        if (confirm != DialogResult.OK)
        {
            return;
        }
        try
        {
            _sendToService.Delete(item);
            ReloadSendTo();
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "删除失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void OpenSendToDirectory()
    {
        var info = new System.Diagnostics.ProcessStartInfo
        {
            FileName = _sendToService.SendToDirectory,
            UseShellExecute = true
        };
        System.Diagnostics.Process.Start(info);
    }

    private void ReloadWinX()
    {
        _winXList.ItemChecked -= OnWinXChecked;
        try
        {
            _winXList.Items.Clear();
            foreach (var item in _winXService.GetEntries())
            {
                var row = new ListViewItem(item.GroupName) { Tag = item, Checked = item.Visible };
                row.SubItems.Add(item.Name);
                row.SubItems.Add(item.TargetPath);
                row.SubItems.Add(item.FilePath);
                _winXList.Items.Add(row);
            }
        }
        finally
        {
            _winXList.ItemChecked += OnWinXChecked;
        }
    }

    private void OnWinXChecked(object? sender, ItemCheckedEventArgs e)
    {
        if (e.Item.Tag is not WinXEntryModel item)
        {
            return;
        }

        try
        {
            _winXService.SetVisible(item, e.Item.Checked);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "WinX 设置失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void AddWinXGroup()
    {
        try
        {
            _winXService.CreateGroup();
            ReloadWinX();
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "创建分组失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void AddWinXItem()
    {
        var groups = _winXService.GetGroups();
        if (groups.Count == 0)
        {
            MessageBox.Show(this, "请先创建 WinX 分组。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        using var dialog = new OpenFileDialog
        {
            Filter = "Program|*.exe;*.lnk|All Files|*.*"
        };
        if (dialog.ShowDialog(this) != DialogResult.OK)
        {
            return;
        }

        var group = PromptDialog.Show(this, $"输入分组名（可选：{string.Join(", ", groups)}）", groups[0]);
        if (string.IsNullOrWhiteSpace(group))
        {
            return;
        }
        var title = PromptDialog.Show(this, "显示名称", Path.GetFileNameWithoutExtension(dialog.FileName));
        if (string.IsNullOrWhiteSpace(title))
        {
            return;
        }
        var args = PromptDialog.Show(this, "启动参数（可空）", string.Empty) ?? string.Empty;

        try
        {
            _winXService.AddEntry(group, title, dialog.FileName, args);
            ReloadWinX();
            _explorerService.RestartExplorer();
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "添加失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void RenameWinXItem()
    {
        if (_winXList.SelectedItems.Count == 0 || _winXList.SelectedItems[0].Tag is not WinXEntryModel item)
        {
            return;
        }
        var newName = PromptDialog.Show(this, "新名称", item.Name);
        if (string.IsNullOrWhiteSpace(newName))
        {
            return;
        }
        try
        {
            _winXService.Rename(item, newName);
            ReloadWinX();
            _explorerService.RestartExplorer();
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "重命名失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void MoveWinXItem(bool moveUp)
    {
        if (_winXList.SelectedItems.Count == 0 || _winXList.SelectedItems[0].Tag is not WinXEntryModel item)
        {
            return;
        }
        try
        {
            _winXService.MoveWithinGroup(item, moveUp);
            ReloadWinX();
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "移动失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void DeleteWinXItem()
    {
        if (_winXList.SelectedItems.Count == 0 || _winXList.SelectedItems[0].Tag is not WinXEntryModel item)
        {
            return;
        }
        var confirm = MessageBox.Show(this, $"确认删除 {item.Name}？", "删除确认", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
        if (confirm != DialogResult.OK)
        {
            return;
        }
        try
        {
            _winXService.Delete(item);
            ReloadWinX();
            _explorerService.RestartExplorer();
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "删除失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
