[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ConfigPath = Join-Path $ScriptRoot "unified-contextmenu.config.json"

function Write-Title {
    Clear-Host
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host "        Unified Context Menu Hub          " -ForegroundColor Cyan
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host ""
}

function Get-DefaultConfig {
    [ordered]@{
        ContextMenuManagerExe = (Join-Path $ScriptRoot "tools/ContextMenuManager/ContextMenuManager.exe")
        FluentContextMenuExe  = (Join-Path $ScriptRoot "tools/FluentContextMenu/FluentContextMenuLoader.exe")
    }
}

function Load-Config {
    if (Test-Path $ConfigPath) {
        try {
            return Get-Content -Raw -Path $ConfigPath | ConvertFrom-Json
        }
        catch {
            Write-Warning "配置文件读取失败，将回退默认配置。"
        }
    }
    return [pscustomobject](Get-DefaultConfig)
}

function Save-Config {
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$Config
    )
    $Config | ConvertTo-Json -Depth 5 | Set-Content -Path $ConfigPath -Encoding UTF8
}

function Test-ClassicMenuEnabled {
    $keyPath = "HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32"
    return Test-Path $keyPath
}

function Enable-ClassicMenu {
    $baseKey = "HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}"
    $target = Join-Path $baseKey "InprocServer32"
    if (-not (Test-Path $target)) {
        New-Item -Path $target -Force | Out-Null
    }
    Set-ItemProperty -Path $target -Name "(default)" -Value "" -Force
    Write-Host "已启用 Win11 经典右键菜单。" -ForegroundColor Green
}

function Disable-ClassicMenu {
    $baseKey = "HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}"
    if (Test-Path $baseKey) {
        Remove-Item -Path $baseKey -Recurse -Force
    }
    Write-Host "已关闭 Win11 经典右键菜单（恢复系统默认）。" -ForegroundColor Yellow
}

function Restart-ExplorerProcess {
    Write-Host "正在重启 Explorer..." -ForegroundColor DarkCyan
    Get-Process explorer -ErrorAction SilentlyContinue | Stop-Process -Force
    Start-Sleep -Milliseconds 700
    Start-Process explorer.exe
    Write-Host "Explorer 已重启。" -ForegroundColor Green
}

function Start-ExternalTool {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ToolPath,
        [Parameter(Mandatory = $true)]
        [string]$ToolName
    )

    if (-not [string]::IsNullOrWhiteSpace($ToolPath) -and (Test-Path $ToolPath)) {
        Start-Process -FilePath $ToolPath | Out-Null
        Write-Host "已启动: $ToolName" -ForegroundColor Green
        return
    }

    Write-Warning "$ToolName 路径无效: $ToolPath"
}

function Set-ToolPath {
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$Config,
        [Parameter(Mandatory = $true)]
        [ValidateSet("ContextMenuManagerExe", "FluentContextMenuExe")]
        [string]$PropertyName,
        [Parameter(Mandatory = $true)]
        [string]$DisplayName
    )

    $inputPath = Read-Host "请输入 $DisplayName 的 exe 完整路径"
    if ([string]::IsNullOrWhiteSpace($inputPath)) {
        Write-Warning "未输入路径，已取消。"
        return
    }

    if (-not (Test-Path $inputPath)) {
        Write-Warning "路径不存在: $inputPath"
        return
    }

    $Config.$PropertyName = $inputPath
    Save-Config -Config $Config
    Write-Host "已保存 $DisplayName 路径。" -ForegroundColor Green
}

function Show-Status {
    param([pscustomobject]$Config)
    $classic = if (Test-ClassicMenuEnabled) { "已启用" } else { "未启用" }
    Write-Host "当前状态：" -ForegroundColor Cyan
    Write-Host "  - Win11 经典菜单: $classic"
    Write-Host "  - ContextMenuManager: $($Config.ContextMenuManagerExe)"
    Write-Host "  - FluentContextMenu:  $($Config.FluentContextMenuExe)"
    Write-Host ""
}

$config = Load-Config

while ($true) {
    Write-Title
    Show-Status -Config $config

    Write-Host "请选择操作：" -ForegroundColor White
    Write-Host "  1. 启用经典右键菜单"
    Write-Host "  2. 关闭经典右键菜单（恢复默认）"
    Write-Host "  3. 重启 Explorer"
    Write-Host "  4. 启动 ContextMenuManager"
    Write-Host "  5. 启动 FluentContextMenuLoader"
    Write-Host "  6. 设置 ContextMenuManager 路径"
    Write-Host "  7. 设置 FluentContextMenu 路径"
    Write-Host "  8. 一键组合：启用经典菜单 + 重启 Explorer + 启动两个工具"
    Write-Host "  0. 退出"
    Write-Host ""

    $choice = Read-Host "输入编号"
    switch ($choice) {
        "1" {
            Enable-ClassicMenu
        }
        "2" {
            Disable-ClassicMenu
        }
        "3" {
            Restart-ExplorerProcess
        }
        "4" {
            Start-ExternalTool -ToolPath $config.ContextMenuManagerExe -ToolName "ContextMenuManager"
        }
        "5" {
            Start-ExternalTool -ToolPath $config.FluentContextMenuExe -ToolName "FluentContextMenuLoader"
        }
        "6" {
            Set-ToolPath -Config $config -PropertyName "ContextMenuManagerExe" -DisplayName "ContextMenuManager"
        }
        "7" {
            Set-ToolPath -Config $config -PropertyName "FluentContextMenuExe" -DisplayName "FluentContextMenuLoader"
        }
        "8" {
            Enable-ClassicMenu
            Restart-ExplorerProcess
            Start-ExternalTool -ToolPath $config.ContextMenuManagerExe -ToolName "ContextMenuManager"
            Start-ExternalTool -ToolPath $config.FluentContextMenuExe -ToolName "FluentContextMenuLoader"
        }
        "0" {
            break
        }
        default {
            Write-Warning "无效输入，请重试。"
        }
    }

    Write-Host ""
    Read-Host "按回车继续"
}
