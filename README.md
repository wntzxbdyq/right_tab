# UnifiedContextMenu（源码深度合并版）

将这两个项目进行源码级融合，落地为一个统一代码库：

- [ContextMenuManager](https://github.com/BluePointLilac/ContextMenuManager)
- [FluentContextMenu](https://winmoes.com/tools/17626.html)

## 合并思路

- `ContextMenuManager`：提炼其“按场景管理右键菜单项”的核心设计（文件、文件夹、背景、驱动器等）。
- `FluentContextMenu`：实现其“关闭系统 WinUI 菜单/恢复经典菜单 + 托盘运行 + 开机自启 + 重启 Explorer”的核心行为。
- 最终统一为一个应用：一个窗口、两类能力、统一服务层。

## 源码结构

```text
src/
  UnifiedContextMenu.sln
  UnifiedContextMenu.Core/
    ContextMenuScene.cs
    ContextMenuItem.cs
    IContextMenuProvider.cs
    IFluentModeService.cs
    IExplorerService.cs
  UnifiedContextMenu.Infrastructure.Windows/
    RegistryContextMenuProvider.cs
    WindowsFluentModeService.cs
    WindowsExplorerService.cs
  UnifiedContextMenu.App.WinForms/
    Program.cs
    ApplicationConfiguration.cs
    MainForm.cs
```

## 功能映射（合并结果）

- 菜单管理（来自 ContextMenuManager 思路）
  - 场景切换
  - 枚举场景下菜单项
  - 勾选启用/禁用菜单项（`LegacyDisable`）
- Win11 增强（来自 FluentContextMenu 行为）
  - 关闭系统 WinUI 菜单（启用经典菜单）
  - 开机自启（托盘模式）
  - 最小化到托盘
  - 应用设置后自动重启 Explorer
- 第二阶段深合并（来自 ContextMenuManager 的专项能力）
  - `OpenWith`：枚举、添加、重命名、显示/隐藏、删除
  - `SendTo`：枚举、添加快捷方式、重命名、显示/隐藏、删除、打开目录
  - `WinX`：分组枚举、新建分组、添加项目、重命名、显示/隐藏、删除、同组上移/下移（含 WinX Hash）

## 构建运行

1. 安装 .NET 8 SDK
2. 打开 `src/UnifiedContextMenu.sln`
3. 启动 `UnifiedContextMenu.App.WinForms`

托盘启动参数：

```powershell
UnifiedContextMenu.App.WinForms.exe --tray
```

## 兼容说明

- 目标系统：Windows 10/11
- 注册表写入项涉及：
  - `HKCU\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}`
  - `HKCU\Software\Microsoft\Windows\CurrentVersion\Run`
- 若需写入 `HKLM` 项，需管理员权限

## 附加说明

- 根目录的 `UnifiedContextMenu.ps1` 是第一阶段的过渡入口，保留用于快速操作。
- 当前主实现已是源码级合并架构，后续可以继续并入更多 `ContextMenuManager` 原生编辑能力（如 OpenWith、SendTo、WinX 分组编辑）。
