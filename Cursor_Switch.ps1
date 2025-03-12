# Cursor 多版本管理腳本

# 環境變數
$cursorVersionsDir = "D:\Cursor_Versions"
$cursorProgramPath = "$env:LOCALAPPDATA\Programs\cursor"
$cursorConfigPath = "$env:APPDATA\Cursor"

# 檢查是否以管理員權限運行
function Test-Administrator {
    $user = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal $user
    return $principal.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}

if (-not (Test-Administrator)) {
    Write-Host "請以管理員權限運行此腳本" -ForegroundColor Red
    exit
}

# 確保版本目錄存在
if (-not (Test-Path $cursorVersionsDir)) {
    New-Item -Path $cursorVersionsDir -ItemType Directory | Out-Null
}

# 檢測當前安裝的 Cursor 版本
function Get-CurrentCursorVersion {
    if (Test-Path $cursorProgramPath) {
        try {
            if ((Get-Item $cursorProgramPath -ErrorAction SilentlyContinue).LinkType -eq "SymbolicLink") {
                $target = (Get-Item $cursorProgramPath).Target
                if ($target -match "D:\\Cursor_Versions\\(.+?)\\Program") {
                    return $matches[1]
                }
            }
            return "預設安裝 (非版本管理)"
        } catch {
            return "預設安裝 (非版本管理)"
        }
    }
    return "未安裝"
}

# 列出所有已安裝的版本
function Get-AllCursorVersions {
    if (Test-Path $cursorVersionsDir) {
        $versions = Get-ChildItem -Path $cursorVersionsDir -Directory | Select-Object -ExpandProperty Name
        return $versions
    }
    return @()
}

# 建立新版本
function New-CursorVersion {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Version
    )
    
    # 驗證版本名稱不為空
    if ([string]::IsNullOrWhiteSpace($Version)) {
        Write-Host "版本名稱不能為空！" -ForegroundColor Red
        return
    }
    
    $versionDir = Join-Path $cursorVersionsDir $Version
    $programDir = Join-Path $versionDir "Program"
    $configDir = Join-Path $versionDir "Config"
    
    # 檢查版本是否已存在
    if (Test-Path $versionDir) {
        Write-Host "版本 $Version 已存在！" -ForegroundColor Yellow
        return
    }
    
    # 檢測當前版本
    $currentVersion = Get-CurrentCursorVersion
    
    # 創建版本目錄
    New-Item -Path $versionDir -ItemType Directory | Out-Null
    
    # 根據當前狀態選擇不同的動作
    if ($currentVersion -ne "預設安裝 (非版本管理)" -and $currentVersion -ne "未安裝") {
        # 已在版本管理下
        $currentVersionDir = Join-Path $cursorVersionsDir $currentVersion
        $currentProgramDir = Join-Path $currentVersionDir "Program"
        $currentConfigDir = Join-Path $currentVersionDir "Config"
        
        $choice = Read-Host "當前正使用版本 $currentVersion，是否要基於此版本創建新版本 $Version？(Y/N)"
        
        if ($choice -eq 'Y' -or $choice -eq 'y') {
            # 從當前版本複製設定，但使用符號連結
            try {
                # 創建符號連結，指向當前版本的配置
                New-Item -Path $configDir -ItemType SymbolicLink -Value $currentConfigDir -ErrorAction Stop | Out-Null
                # 創建新版本的程序目錄（空目錄，準備安裝新版）
                New-Item -Path $programDir -ItemType Directory | Out-Null
                
                Write-Host "已基於版本 $currentVersion 創建新版本 $Version" -ForegroundColor Green
                Write-Host "配置將與版本 $currentVersion 共享，但程序將是全新安裝" -ForegroundColor Green
            } catch {
                Write-Host "錯誤: 無法創建版本" -ForegroundColor Red
                Write-Host $_.Exception.Message
                return
            }
        } else {
            # 創建全新版本
            New-Item -Path $programDir -ItemType Directory | Out-Null
            New-Item -Path $configDir -ItemType Directory | Out-Null
            Write-Host "已創建全新版本 $Version" -ForegroundColor Green
        }
    } else {
        # 處理預設安裝或未安裝的情況
        $backupProgram = $null
        $backupConfig = $null
        
        if (Test-Path $cursorProgramPath) {
            $backupProgram = $true
            try {
                if ((Get-Item $cursorProgramPath -ErrorAction SilentlyContinue).LinkType -eq "SymbolicLink") {
                    Remove-Item $cursorProgramPath -Force
                } else {
                    Rename-Item $cursorProgramPath "${cursorProgramPath}_backup" -Force
                }
            } catch {
                Write-Host "警告: 無法處理程式安裝路徑 $cursorProgramPath" -ForegroundColor Yellow
            }
        }
        
        if (Test-Path $cursorConfigPath) {
            $backupConfig = $true
            try {
                if ((Get-Item $cursorConfigPath -ErrorAction SilentlyContinue).LinkType -eq "SymbolicLink") {
                    Remove-Item $cursorConfigPath -Force
                } else {
                    Rename-Item $cursorConfigPath "${cursorConfigPath}_backup" -Force
                }
            } catch {
                Write-Host "警告: 無法處理配置路徑 $cursorConfigPath" -ForegroundColor Yellow
            }
        }
        
        # 如果有備份，詢問是否要使用已安裝版本的資料
        if ($backupProgram -or $backupConfig) {
            $choice = Read-Host "檢測到預設安裝的 Cursor。要使用現有安裝的檔案嗎？(Y/N)"
            if ($choice -eq 'Y' -or $choice -eq 'y') {
                # 直接創建符號連結，從版本目錄指向原始備份
                if ($backupProgram) {
                    try {
                        New-Item -Path $programDir -ItemType SymbolicLink -Value "${cursorProgramPath}_backup" -ErrorAction Stop | Out-Null
                    } catch {
                        Write-Host "錯誤: 無法創建程式目錄連結" -ForegroundColor Red
                        Write-Host $_.Exception.Message
                    }
                }
                
                if ($backupConfig) {
                    try {
                        New-Item -Path $configDir -ItemType SymbolicLink -Value "${cursorConfigPath}_backup" -ErrorAction Stop | Out-Null
                    } catch {
                        Write-Host "錯誤: 無法創建配置目錄連結" -ForegroundColor Red
                        Write-Host $_.Exception.Message
                    }
                }
                
                Write-Host "已將預設安裝轉換為版本 $Version" -ForegroundColor Green
            } else {
                # 創建全新版本
                New-Item -Path $programDir -ItemType Directory | Out-Null
                New-Item -Path $configDir -ItemType Directory | Out-Null
                Write-Host "已創建全新版本 $Version" -ForegroundColor Green
            }
        } else {
            # 創建全新版本
            New-Item -Path $programDir -ItemType Directory | Out-Null
            New-Item -Path $configDir -ItemType Directory | Out-Null
            Write-Host "已創建全新版本 $Version" -ForegroundColor Green
        }
    }
    
    # 切換到新版本
    try {
        # 移除現有連結
        if (Test-Path $cursorProgramPath) {
            Remove-Item $cursorProgramPath -Force -ErrorAction Stop
        }
        if (Test-Path $cursorConfigPath) {
            Remove-Item $cursorConfigPath -Force -ErrorAction Stop
        }
        
        # 創建新連結
        New-Item -Path $cursorProgramPath -ItemType SymbolicLink -Value $programDir -ErrorAction Stop | Out-Null
        New-Item -Path $cursorConfigPath -ItemType SymbolicLink -Value $configDir -ErrorAction Stop | Out-Null
        
        Write-Host "已切換到版本 $Version，請安裝或更新 Cursor" -ForegroundColor Green
    } catch {
        Write-Host "錯誤: 無法切換到新版本" -ForegroundColor Red
        Write-Host $_.Exception.Message
    }
}

# 切換到特定版本
function Switch-CursorVersion {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Version
    )
    
    $versionDir = Join-Path $cursorVersionsDir $Version
    $programDir = Join-Path $versionDir "Program"
    $configDir = Join-Path $versionDir "Config"
    
    # 檢查版本是否存在
    if (-not (Test-Path $versionDir)) {
        Write-Host "版本 $Version 不存在！" -ForegroundColor Red
        return
    }
    
    # 移除現有連結
    if (Test-Path $cursorProgramPath) {
        try {
            Remove-Item $cursorProgramPath -Force -ErrorAction Stop
        } catch {
            Write-Host "錯誤: 無法移除程式目錄連結" -ForegroundColor Red
            Write-Host $_.Exception.Message
            return
        }
    }
    if (Test-Path $cursorConfigPath) {
        try {
            Remove-Item $cursorConfigPath -Force -ErrorAction Stop
        } catch {
            Write-Host "錯誤: 無法移除配置目錄連結" -ForegroundColor Red
            Write-Host $_.Exception.Message
            return
        }
    }
    
    # 創建新連結
    try {
        New-Item -Path $cursorProgramPath -ItemType SymbolicLink -Value $programDir -ErrorAction Stop | Out-Null
        New-Item -Path $cursorConfigPath -ItemType SymbolicLink -Value $configDir -ErrorAction Stop | Out-Null
        Write-Host "已切換到 Cursor 版本 $Version" -ForegroundColor Green
    } catch {
        Write-Host "錯誤: 無法創建符號連結" -ForegroundColor Red
        Write-Host $_.Exception.Message
    }
}

# 更新特定版本
function Update-CursorVersion {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Version
    )
    
    $versionDir = Join-Path $cursorVersionsDir $Version
    
    # 檢查版本是否存在
    if (-not (Test-Path $versionDir)) {
        Write-Host "版本 $Version 不存在！" -ForegroundColor Red
        return
    }
    
    # 切換到該版本
    Switch-CursorVersion -Version $Version
    
    Write-Host "已切換到 Cursor 版本 $Version，請運行 Cursor 更新程序" -ForegroundColor Green
    Write-Host "更新完成後，版本 $Version 將保持更新狀態" -ForegroundColor Green
}

# 還原到預設路徑
function Restore-CursorDefaultPath {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Version
    )
    
    $versionDir = Join-Path $cursorVersionsDir $Version
    $programDir = Join-Path $versionDir "Program"
    $configDir = Join-Path $versionDir "Config"
    
    # 檢查版本是否存在
    if (-not (Test-Path $versionDir)) {
        Write-Host "版本 $Version 不存在！" -ForegroundColor Red
        return
    }
    
    # 移除現有連結
    if (Test-Path $cursorProgramPath) {
        try {
            Remove-Item $cursorProgramPath -Force -ErrorAction Stop
        } catch {
            Write-Host "錯誤: 無法移除程式目錄連結" -ForegroundColor Red
            return
        }
    }
    if (Test-Path $cursorConfigPath) {
        try {
            Remove-Item $cursorConfigPath -Force -ErrorAction Stop
        } catch {
            Write-Host "錯誤: 無法移除配置目錄連結" -ForegroundColor Red
            return
        }
    }
    
    # 直接創建連結到版本目錄
    try {
        New-Item -Path $cursorProgramPath -ItemType SymbolicLink -Value $programDir -ErrorAction Stop | Out-Null
        New-Item -Path $cursorConfigPath -ItemType SymbolicLink -Value $configDir -ErrorAction Stop | Out-Null
        Write-Host "已將 Cursor 版本 $Version 還原到預設路徑" -ForegroundColor Green
    } catch {
        Write-Host "錯誤: 無法創建符號連結" -ForegroundColor Red
        Write-Host $_.Exception.Message
    }
}

# 檢測系統上已安裝的版本並引導初次設置
function Initialize-CursorVersionManagement {
    $currentState = Get-CurrentCursorVersion
    
    if ($currentState -eq "未安裝") {
        Write-Host "未檢測到 Cursor 安裝。請先安裝 Cursor 或選擇建立新版本。" -ForegroundColor Yellow
        return
    }
    
    if ($currentState -eq "預設安裝 (非版本管理)") {
        $choice = Read-Host "檢測到預設安裝的 Cursor。要將其轉換為版本管理嗎？(Y/N)"
        if ($choice -eq 'Y' -or $choice -eq 'y') {
            $version = Read-Host "請輸入版本名稱 (例如: v0.45)"
            
            # 檢查版本名稱不為空
            if ([string]::IsNullOrWhiteSpace($version)) {
                Write-Host "版本名稱不能為空！" -ForegroundColor Red
                return
            }
            
            $versionDir = Join-Path $cursorVersionsDir $version
            
            # 檢查版本是否已存在
            if (Test-Path $versionDir) {
                Write-Host "版本 $version 已存在！請選擇其他版本名稱。" -ForegroundColor Red
                return
            }
            
            # 創建版本目錄
            New-Item -Path $versionDir -ItemType Directory -Force | Out-Null
            $programDir = Join-Path $versionDir "Program"
            $configDir = Join-Path $versionDir "Config"
            
            # 備份原目錄位置
            $programBackupPath = "${cursorProgramPath}_original"
            $configBackupPath = "${cursorConfigPath}_original"
            
            # 移動原目錄（重命名）
            try {
                Rename-Item $cursorProgramPath $programBackupPath -Force -ErrorAction Stop
                Rename-Item $cursorConfigPath $configBackupPath -Force -ErrorAction Stop
            } catch {
                Write-Host "錯誤: 無法重命名原始目錄" -ForegroundColor Red
                Write-Host $_.Exception.Message
                return
            }
            
            # 創建符號連結
            try {
                New-Item -Path $programDir -ItemType SymbolicLink -Value $programBackupPath -ErrorAction Stop | Out-Null
                New-Item -Path $configDir -ItemType SymbolicLink -Value $configBackupPath -ErrorAction Stop | Out-Null
            
                # 原始路徑指向版本目錄
                New-Item -Path $cursorProgramPath -ItemType SymbolicLink -Value $programDir -ErrorAction Stop | Out-Null
                New-Item -Path $cursorConfigPath -ItemType SymbolicLink -Value $configDir -ErrorAction Stop | Out-Null
            
                Write-Host "已將當前安裝轉換為版本 $version" -ForegroundColor Green
            } catch {
                Write-Host "錯誤: 無法創建符號連結" -ForegroundColor Red
                Write-Host $_.Exception.Message
                
                # 嘗試還原原始目錄
                try {
                    Rename-Item $programBackupPath $cursorProgramPath -Force
                    Rename-Item $configBackupPath $cursorConfigPath -Force
                } catch {
                    Write-Host "警告: 無法還原原始目錄，請手動檢查" -ForegroundColor Red
                }
            }
        }
    }
}

# 顯示菜單
function Show-Menu {
    Clear-Host
    $currentVersion = Get-CurrentCursorVersion
    $allVersions = Get-AllCursorVersions
    
    Write-Host "===== Cursor 多版本管理工具 =====" -ForegroundColor Cyan
    Write-Host
    Write-Host "當前版本: $currentVersion" -ForegroundColor Green
    Write-Host
    Write-Host "可用版本:" -ForegroundColor Cyan
    if ($allVersions.Count -eq 0) {
        Write-Host "  尚未創建任何版本" -ForegroundColor Yellow
    } else {
        foreach ($version in $allVersions) {
            Write-Host "  $version" -ForegroundColor White
        }
    }
    
    Write-Host
    Write-Host "1. 建立新版本" -ForegroundColor Yellow
    Write-Host "2. 切換版本" -ForegroundColor Yellow
    Write-Host "3. 更新特定版本" -ForegroundColor Yellow
    Write-Host "4. 還原到預設路徑" -ForegroundColor Yellow
    Write-Host "5. 初始化版本管理" -ForegroundColor Yellow
    Write-Host "6. 創建版本啟動捷徑" -ForegroundColor Green  # 新增選項
    Write-Host "7. 創建版本啟動批處理檔" -ForegroundColor Green  # 新增選項
    Write-Host "8. 直接啟動特定版本" -ForegroundColor Green  # 新增選項
    Write-Host "9. 為所有版本創建啟動器" -ForegroundColor Green  # 新增選項
    Write-Host "0. 退出" -ForegroundColor Yellow
    Write-Host
    
    $choice = Read-Host "請選擇操作"
    
    switch ($choice) {
        "1" {
            $version = Read-Host "請輸入新版本名稱 (例如: v0.45)"
            if ([string]::IsNullOrWhiteSpace($version)) {
                Write-Host "版本名稱不能為空！" -ForegroundColor Red
            } else {
                New-CursorVersion -Version $version
            }
            pause
            Show-Menu
        }
        "2" {
            if ($allVersions.Count -eq 0) {
                Write-Host "尚未創建任何版本！" -ForegroundColor Red
                pause
                Show-Menu
                return
            }
            
            Write-Host "可用版本:" -ForegroundColor Cyan
            for ($i = 0; $i -lt $allVersions.Count; $i++) {
                Write-Host "$($i+1). $($allVersions[$i])" -ForegroundColor White
            }
            
            $versionChoice = Read-Host "請選擇版本編號"
            if ([string]::IsNullOrWhiteSpace($versionChoice)) {
                Write-Host "未選擇任何版本！" -ForegroundColor Red
            } else {
                try {
                    $versionIdx = [int]$versionChoice - 1
                    if ($versionIdx -ge 0 -and $versionIdx -lt $allVersions.Count) {
                        Switch-CursorVersion -Version $allVersions[$versionIdx]
                    } else {
                        Write-Host "無效的選擇！" -ForegroundColor Red
                    }
                } catch {
                    Write-Host "請輸入有效的數字！" -ForegroundColor Red
                }
            }
            pause
            Show-Menu
        }
        "3" {
            if ($allVersions.Count -eq 0) {
                Write-Host "尚未創建任何版本！" -ForegroundColor Red
                pause
                Show-Menu
                return
            }
            
            Write-Host "可用版本:" -ForegroundColor Cyan
            for ($i = 0; $i -lt $allVersions.Count; $i++) {
                Write-Host "$($i+1). $($allVersions[$i])" -ForegroundColor White
            }
            
            $versionChoice = Read-Host "請選擇要更新的版本編號"
            if ([string]::IsNullOrWhiteSpace($versionChoice)) {
                Write-Host "未選擇任何版本！" -ForegroundColor Red
            } else {
                try {
                    $versionIdx = [int]$versionChoice - 1
                    if ($versionIdx -ge 0 -and $versionIdx -lt $allVersions.Count) {
                        Update-CursorVersion -Version $allVersions[$versionIdx]
                    } else {
                        Write-Host "無效的選擇！" -ForegroundColor Red
                    }
                } catch {
                    Write-Host "請輸入有效的數字！" -ForegroundColor Red
                }
            }
            pause
            Show-Menu
        }
        "4" {
            if ($allVersions.Count -eq 0) {
                Write-Host "尚未創建任何版本！" -ForegroundColor Red
                pause
                Show-Menu
                return
            }
            
            Write-Host "可用版本:" -ForegroundColor Cyan
            for ($i = 0; $i -lt $allVersions.Count; $i++) {
                Write-Host "$($i+1). $($allVersions[$i])" -ForegroundColor White
            }
            
            $versionChoice = Read-Host "請選擇要還原到預設路徑的版本編號"
            if ([string]::IsNullOrWhiteSpace($versionChoice)) {
                Write-Host "未選擇任何版本！" -ForegroundColor Red
            } else {
                try {
                    $versionIdx = [int]$versionChoice - 1
                    if ($versionIdx -ge 0 -and $versionIdx -lt $allVersions.Count) {
                        Restore-CursorDefaultPath -Version $allVersions[$versionIdx]
                    } else {
                        Write-Host "無效的選擇！" -ForegroundColor Red
                    }
                } catch {
                    Write-Host "請輸入有效的數字！" -ForegroundColor Red
                }
            }
            pause
            Show-Menu
        }
        "5" {
            Initialize-CursorVersionManagement
            pause
            Show-Menu
        }
        "6" {
            if ($allVersions.Count -eq 0) {
                Write-Host "尚未創建任何版本！" -ForegroundColor Red
                pause
                Show-Menu
                return
            }
            
            Write-Host "可用版本:" -ForegroundColor Cyan
            for ($i = 0; $i -lt $allVersions.Count; $i++) {
                Write-Host "$($i+1). $($allVersions[$i])" -ForegroundColor White
            }
            
            $versionChoice = Read-Host "請選擇版本編號"
            if ([string]::IsNullOrWhiteSpace($versionChoice)) {
                Write-Host "未選擇任何版本！" -ForegroundColor Red
            } else {
                try {
                    $versionIdx = [int]$versionChoice - 1
                    if ($versionIdx -ge 0 -and $versionIdx -lt $allVersions.Count) {
                        $locChoice = Read-Host "創建在桌面(D)還是開始選單(S)？(D/S)"
                        if ($locChoice -eq 'S' -or $locChoice -eq 's') {
                            Create-CursorShortcut -Version $allVersions[$versionIdx] -ShortcutLocation "StartMenu"
                        } else {
                            Create-CursorShortcut -Version $allVersions[$versionIdx] -ShortcutLocation "Desktop"
                        }
                    } else {
                        Write-Host "無效的選擇！" -ForegroundColor Red
                    }
                } catch {
                    Write-Host "請輸入有效的數字！" -ForegroundColor Red
                }
            }
            pause
            Show-Menu
        }
        "7" {
            if ($allVersions.Count -eq 0) {
                Write-Host "尚未創建任何版本！" -ForegroundColor Red
                pause
                Show-Menu
                return
            }
            
            Write-Host "可用版本:" -ForegroundColor Cyan
            for ($i = 0; $i -lt $allVersions.Count; $i++) {
                Write-Host "$($i+1). $($allVersions[$i])" -ForegroundColor White
            }
            
            $versionChoice = Read-Host "請選擇版本編號"
            if ([string]::IsNullOrWhiteSpace($versionChoice)) {
                Write-Host "未選擇任何版本！" -ForegroundColor Red
            } else {
                try {
                    $versionIdx = [int]$versionChoice - 1
                    if ($versionIdx -ge 0 -and $versionIdx -lt $allVersions.Count) {
                        Create-CursorBatchFile -Version $allVersions[$versionIdx]
                    } else {
                        Write-Host "無效的選擇！" -ForegroundColor Red
                    }
                } catch {
                    Write-Host "請輸入有效的數字！" -ForegroundColor Red
                }
            }
            pause
            Show-Menu
        }
        "8" {
            if ($allVersions.Count -eq 0) {
                Write-Host "尚未創建任何版本！" -ForegroundColor Red
                pause
                Show-Menu
                return
            }
            
            Write-Host "可用版本:" -ForegroundColor Cyan
            for ($i = 0; $i -lt $allVersions.Count; $i++) {
                Write-Host "$($i+1). $($allVersions[$i])" -ForegroundColor White
            }
            
            $versionChoice = Read-Host "請選擇要啟動的版本編號"
            if ([string]::IsNullOrWhiteSpace($versionChoice)) {
                Write-Host "未選擇任何版本！" -ForegroundColor Red
            } else {
                try {
                    $versionIdx = [int]$versionChoice - 1
                    if ($versionIdx -ge 0 -and $versionIdx -lt $allVersions.Count) {
                        Start-CursorVersion -Version $allVersions[$versionIdx]
                    } else {
                        Write-Host "無效的選擇！" -ForegroundColor Red
                    }
                } catch {
                    Write-Host "請輸入有效的數字！" -ForegroundColor Red
                }
            }
            pause
            Show-Menu
        }
        "9" {
            Create-AllVersionLaunchers
            pause
            Show-Menu
        }
        "0" {
            return
        }
        default {
            Write-Host "無效的選擇！" -ForegroundColor Red
            pause
            Show-Menu
        }
    }
}

# 為指定版本創建啟動捷徑
function Create-CursorShortcut {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Version,
        
        [Parameter(Mandatory=$false)]
        [string]$ShortcutLocation = "Desktop" # 可選：Desktop 或 StartMenu
    )
    
    $versionDir = Join-Path $cursorVersionsDir $Version
    $programDir = Join-Path $versionDir "Program"
    $configDir = Join-Path $versionDir "Config"
    $exePath = Join-Path $programDir "Cursor.exe"
    
    # 檢查版本是否存在
    if (-not (Test-Path $versionDir)) {
        Write-Host "版本 $Version 不存在！" -ForegroundColor Red
        return
    }
    
    # 檢查可執行檔是否存在
    if (-not (Test-Path $exePath)) {
        Write-Host "找不到 $Version 版本的可執行檔，請先安裝此版本！" -ForegroundColor Red
        return
    }
    
    # 決定捷徑位置
    $targetPath = ""
    if ($ShortcutLocation -eq "Desktop") {
        $targetPath = [Environment]::GetFolderPath("Desktop")
    } elseif ($ShortcutLocation -eq "StartMenu") {
        $targetPath = [Environment]::GetFolderPath("StartMenu")
    } else {
        $targetPath = [Environment]::GetFolderPath("Desktop")
    }
    
    $shortcutPath = Join-Path $targetPath "Cursor_$Version.lnk"
    
    # 創建捷徑
    try {
        $WshShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut($shortcutPath)
        $Shortcut.TargetPath = $exePath
        $Shortcut.Arguments = "--user-data-dir=""$configDir"""
        $Shortcut.WorkingDirectory = $programDir
        $Shortcut.Description = "Cursor 編輯器 - $Version 版本"
        # 如果有圖示，可以設置圖示
        if (Test-Path (Join-Path $programDir "resources\app\static\icons\win\app.ico")) {
            $Shortcut.IconLocation = Join-Path $programDir "resources\app\static\icons\win\app.ico"
        }
        $Shortcut.Save()
        
        Write-Host "已成功創建 Cursor $Version 版本的桌面捷徑" -ForegroundColor Green
    } catch {
        Write-Host "建立捷徑時發生錯誤: $_" -ForegroundColor Red
    }
}

# 創建獨立的啟動批處理檔
function Create-CursorBatchFile {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Version,
        
        [Parameter(Mandatory=$false)]
        [string]$OutputPath = "" # 如為空，將存儲在版本目錄下
    )
    
    $versionDir = Join-Path $cursorVersionsDir $Version
    $programDir = Join-Path $versionDir "Program"
    $configDir = Join-Path $versionDir "Config"
    $exePath = Join-Path $programDir "Cursor.exe"
    
    # 檢查版本是否存在
    if (-not (Test-Path $versionDir)) {
        Write-Host "版本 $Version 不存在！" -ForegroundColor Red
        return
    }
    
    # 檢查可執行檔是否存在
    if (-not (Test-Path $exePath)) {
        Write-Host "找不到 $Version 版本的可執行檔，請先安裝此版本！" -ForegroundColor Red
        return
    }
    
    # 決定輸出路徑
    if ([string]::IsNullOrWhiteSpace($OutputPath)) {
        $batchPath = Join-Path $versionDir "Launch_Cursor_$Version.bat"
    } else {
        $batchPath = Join-Path $OutputPath "Launch_Cursor_$Version.bat"
    }
    
    # 創建批處理檔
    $batchContent = @"
@echo off
echo 正在啟動 Cursor $Version 版本...
start "" "$exePath" --user-data-dir="$configDir"
"@
    
    try {
        Set-Content -Path $batchPath -Value $batchContent -Encoding ASCII
        Write-Host "已成功創建 Cursor $Version 版本的啟動批處理檔: $batchPath" -ForegroundColor Green
    } catch {
        Write-Host "建立啟動批處理檔時發生錯誤: $_" -ForegroundColor Red
    }
}

# 為所有版本創建啟動器
function Create-AllVersionLaunchers {
    $allVersions = Get-AllCursorVersions
    
    if ($allVersions.Count -eq 0) {
        Write-Host "尚未建立任何版本！" -ForegroundColor Red
        return
    }
    
    foreach ($version in $allVersions) {
        $choice = Read-Host "是否為 $version 版本創建啟動捷徑？(Y/N)"
        if ($choice -eq 'Y' -or $choice -eq 'y') {
            $locChoice = Read-Host "創建在桌面(D)還是開始選單(S)？(D/S)"
            if ($locChoice -eq 'S' -or $locChoice -eq 's') {
                Create-CursorShortcut -Version $version -ShortcutLocation "StartMenu"
            } else {
                Create-CursorShortcut -Version $version -ShortcutLocation "Desktop"
            }
            
            $batchChoice = Read-Host "是否也創建啟動批處理檔？(Y/N)"
            if ($batchChoice -eq 'Y' -or $batchChoice -eq 'y') {
                Create-CursorBatchFile -Version $version
            }
        }
    }
}

# 直接啟動指定版本的 Cursor（不切換系統路徑）
function Start-CursorVersion {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Version
    )
    
    $versionDir = Join-Path $cursorVersionsDir $Version
    $programDir = Join-Path $versionDir "Program"
    $configDir = Join-Path $versionDir "Config"
    $exePath = Join-Path $programDir "Cursor.exe"
    
    # 檢查版本是否存在
    if (-not (Test-Path $versionDir)) {
        Write-Host "版本 $Version 不存在！" -ForegroundColor Red
        return
    }
    
    # 檢查可執行檔是否存在
    if (-not (Test-Path $exePath)) {
        Write-Host "找不到 $Version 版本的可執行檔，請先安裝此版本！" -ForegroundColor Red
        return
    }
    
    # 啟動 Cursor
    try {
        Start-Process $exePath -ArgumentList "--user-data-dir=""$configDir"""
        Write-Host "已啟動 Cursor $Version 版本" -ForegroundColor Green
    } catch {
        Write-Host "啟動 Cursor 時發生錯誤: $_" -ForegroundColor Red
    }
}

# 啟動菜單
Show-Menu