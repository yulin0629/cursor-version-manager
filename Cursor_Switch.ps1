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
            # 檢查是否為符號連結
            $programItem = Get-Item $cursorProgramPath -ErrorAction SilentlyContinue
            if ($programItem -and $programItem.LinkType -eq "SymbolicLink") {
                $target = $programItem.Target
                # 從目標路徑中提取版本名稱
                if ($target -match [regex]::Escape($cursorVersionsDir) + "\\(.+?)\\Program") {
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
        $backupProgram = $false # 用布林值更清晰
        $backupConfig = $false
        
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
                if ($backupProgram -and (Test-Path "${cursorProgramPath}_backup")) {
                    try {
                        New-Item -Path $programDir -ItemType SymbolicLink -Value "${cursorProgramPath}_backup" -ErrorAction Stop | Out-Null
                    } catch {
                        Write-Host "錯誤: 無法創建程式目錄連結" -ForegroundColor Red
                        Write-Host $_.Exception.Message
                    }
                } elseif ($backupProgram) {
                     New-Item -Path $programDir -ItemType Directory | Out-Null # 確保目錄存在
                }
                
                if ($backupConfig -and (Test-Path "${cursorConfigPath}_backup")) {
                    try {
                        New-Item -Path $configDir -ItemType SymbolicLink -Value "${cursorConfigPath}_backup" -ErrorAction Stop | Out-Null
                    } catch {
                        Write-Host "錯誤: 無法創建配置目錄連結" -ForegroundColor Red
                        Write-Host $_.Exception.Message
                    }
                } elseif ($backupConfig) {
                    New-Item -Path $configDir -ItemType Directory | Out-Null # 確保目錄存在
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
    } else {
        Write-Host "Cursor 已在版本管理模式下 ($currentState)。" -ForegroundColor Green
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
        # 開始選單的程式捷徑通常放在 AppData\Roaming\Microsoft\Windows\Start Menu\Programs
        $targetPath = Join-Path ([Environment]::GetFolderPath("ApplicationData")) "Microsoft\Windows\Start Menu\Programs"
        if (-not (Test-Path $targetPath)) {
            New-Item -Path $targetPath -ItemType Directory -Force | Out-Null
        }
    } else {
        $targetPath = [Environment]::GetFolderPath("Desktop") # 預設為桌面
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
        $iconPath = Join-Path $programDir "resources\app\static\icons\win\app.ico" # Cursor 的標準圖示路徑
        if (Test-Path $iconPath) {
            $Shortcut.IconLocation = $iconPath
        }
        $Shortcut.Save()
        
        Write-Host "已成功創建 Cursor $Version 版本的捷徑於 $ShortcutLocation ($shortcutPath)" -ForegroundColor Green
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
    $finalBatchPath = ""
    if ([string]::IsNullOrWhiteSpace($OutputPath)) {
        $finalBatchPath = Join-Path $versionDir "Launch_Cursor_$Version.bat"
    } else {
        # 確保輸出目錄存在
        if (-not (Test-Path $OutputPath)) {
            try {
                New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
            } catch {
                Write-Host "建立輸出目錄 $OutputPath 時發生錯誤: $_" -ForegroundColor Red
                return
            }
        }
        $finalBatchPath = Join-Path $OutputPath "Launch_Cursor_$Version.bat"
    }
    
    # 創建批處理檔內容
    $batchContent = @"
@echo off
echo 正在啟動 Cursor $Version 版本...
start "" "$exePath" --user-data-dir="$configDir"
"@
    
    try {
        Set-Content -Path $finalBatchPath -Value $batchContent -Encoding UTF8 # 使用 UTF8 以支援可能的特殊字元
        Write-Host "已成功創建 Cursor $Version 版本的啟動批處理檔: $finalBatchPath" -ForegroundColor Green
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
        $exePath = Join-Path $cursorVersionsDir "$version\Program\Cursor.exe"
        if (-not (Test-Path $exePath)) {
            Write-Host "版本 $version 的 Cursor.exe 不存在，跳過創建啟動器。" -ForegroundColor Yellow
            continue
        }

        $choiceShortcut = Read-Host "是否為 $version 版本創建啟動捷徑？(Y/N)"
        if ($choiceShortcut -match '^[Yy]$') {
            $locChoice = Read-Host "創建在桌面(D)還是開始選單(S)？(D/S)"
            if ($locChoice -match '^[Ss]$') {
                Create-CursorShortcut -Version $version -ShortcutLocation "StartMenu"
            } else {
                Create-CursorShortcut -Version $version -ShortcutLocation "Desktop"
            }
        }
            
        $choiceBatch = Read-Host "是否也為 $version 版本創建啟動批處理檔？(Y/N)"
        if ($choiceBatch -match '^[Yy]$') {
            Create-CursorBatchFile -Version $version # 預設儲存在版本目錄內
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
        Write-Host "找不到 $Version 版本的可執行檔 ($exePath)，請先安裝此版本！" -ForegroundColor Red
        return
    }

    # 啟動 Cursor
    try {
        Write-Host "正在啟動 Cursor $Version 版本從 $exePath ..." -ForegroundColor Cyan
        Start-Process $exePath -ArgumentList "--user-data-dir=""$configDir"""
        Write-Host "已啟動 Cursor $Version 版本" -ForegroundColor Green
    } catch {
        Write-Host "啟動 Cursor 時發生錯誤: $_" -ForegroundColor Red
    }
}

# 重命名版本
function Rename-CursorVersion {
    param (
        [Parameter(Mandatory=$true)]
        [string]$OldVersion,

        [Parameter(Mandatory=$true)]
        [string]$NewVersion
    )

    # 驗證參數
    if ([string]::IsNullOrWhiteSpace($OldVersion) -or [string]::IsNullOrWhiteSpace($NewVersion)) {
        Write-Host "版本名稱不能為空！" -ForegroundColor Red
        return
    }
    if ($OldVersion -eq $NewVersion) {
        Write-Host "新舊版本名稱相同，無需重命名。" -ForegroundColor Yellow
        return
    }


    $oldVersionDir = Join-Path $cursorVersionsDir $OldVersion
    $newVersionDir = Join-Path $cursorVersionsDir $NewVersion

    # 檢查原版本是否存在
    if (-not (Test-Path $oldVersionDir)) {
        Write-Host "版本 $OldVersion 不存在！" -ForegroundColor Red
        return
    }

    # 檢查新版本名稱是否已被使用
    if (Test-Path $newVersionDir) {
        Write-Host "版本名稱 $NewVersion 已存在！請使用其他名稱。" -ForegroundColor Red
        return
    }

    # 檢查當前版本
    $currentActiveVersion = Get-CurrentCursorVersion
    $isCurrentVersionBeingRenamed = ($currentActiveVersion -eq $OldVersion)

    # 準備路徑
    $oldProgramDir = Join-Path $oldVersionDir "Program"
    $oldConfigDir = Join-Path $oldVersionDir "Config"

    # 臨時變數，用於恢復符號連結
    $tempBackupProgramLinkTarget = $null
    $tempBackupConfigLinkTarget = $null

    # 如果是當前版本，需要暫時解除連結
    if ($isCurrentVersionBeingRenamed) {
        Write-Host "檢測到 $OldVersion 是當前使用中的版本，將暫時解除系統連結..." -ForegroundColor Yellow

        try {
            if (Test-Path $cursorProgramPath -and (Get-Item $cursorProgramPath -ErrorAction SilentlyContinue).LinkType -eq "SymbolicLink") {
                $tempBackupProgramLinkTarget = (Get-Item $cursorProgramPath).Target
                Remove-Item $cursorProgramPath -Force -ErrorAction Stop
            }

            if (Test-Path $cursorConfigPath -and (Get-Item $cursorConfigPath -ErrorAction SilentlyContinue).LinkType -eq "SymbolicLink") {
                $tempBackupConfigLinkTarget = (Get-Item $cursorConfigPath).Target
                Remove-Item $cursorConfigPath -Force -ErrorAction Stop
            }
        } catch {
            Write-Host "解除系統連結時發生錯誤: $_" -ForegroundColor Red
            return # 關鍵操作失敗，終止重命名
        }
    }

    # 執行重命名資料夾
    Write-Host "正在重命名版本資料夾從 '$OldVersion' 到 '$NewVersion'..." -ForegroundColor Cyan
    try {
        Rename-Item -Path $oldVersionDir -NewName $NewVersion -ErrorAction Stop
        Write-Host "版本資料夾已成功重命名為 $NewVersion" -ForegroundColor Green
    } catch {
        Write-Host "重命名版本資料夾時發生錯誤: $_" -ForegroundColor Red

        # 嘗試恢復系統連結（如果之前解除了）
        if ($isCurrentVersionBeingRenamed) {
            Write-Host "嘗試恢復系統連結至原版本 $OldVersion..." -ForegroundColor Yellow
            try {
                if ($tempBackupProgramLinkTarget) {
                    New-Item -Path $cursorProgramPath -ItemType SymbolicLink -Value $tempBackupProgramLinkTarget -ErrorAction Stop | Out-Null
                }
                if ($tempBackupConfigLinkTarget) {
                    New-Item -Path $cursorConfigPath -ItemType SymbolicLink -Value $tempBackupConfigLinkTarget -ErrorAction Stop | Out-Null
                }
                Write-Host "已恢復系統連結至 $OldVersion" -ForegroundColor Green
            } catch {
                Write-Host "警告: 無法完全恢復系統連結，您可能需要手動檢查 $cursorProgramPath 和 $cursorConfigPath" -ForegroundColor Red
            }
        }
        return # 資料夾重命名失敗，終止
    }

    # 如果成功重命名，且原版本是作用中版本，則重建系統連結到新名稱
    if ($isCurrentVersionBeingRenamed) {
        $newProgramDirForLink = Join-Path $newVersionDir "Program"
        $newConfigDirForLink = Join-Path $newVersionDir "Config"
        
        Write-Host "正在重新建立系統符號連結至新版本路徑 ($NewVersion)..." -ForegroundColor Cyan
        try {
            New-Item -Path $cursorProgramPath -ItemType SymbolicLink -Value $newProgramDirForLink -ErrorAction Stop | Out-Null
            New-Item -Path $cursorConfigPath -ItemType SymbolicLink -Value $newConfigDirForLink -ErrorAction Stop | Out-Null
            Write-Host "已成功重新建立系統符號連結至新版本 $NewVersion" -ForegroundColor Green
        } catch {
            Write-Host "重建系統符號連結至 $NewVersion 時發生錯誤: $_" -ForegroundColor Red
            Write-Host "您可能需要手動將 $cursorProgramPath 連結至 $newProgramDirForLink" -ForegroundColor Yellow
            Write-Host "以及將 $cursorConfigPath 連結至 $newConfigDirForLink" -ForegroundColor Yellow
        }
    }

    # 更新相關的捷徑和批次檔
    Write-Host "正在更新相關的捷徑與批次檔名稱..." -ForegroundColor Cyan
    # 更新桌面捷徑
    $desktopPath = [Environment]::GetFolderPath("Desktop")
    $oldDesktopShortcut = Join-Path $desktopPath "Cursor_$OldVersion.lnk"
    if (Test-Path $oldDesktopShortcut) {
        try {
            Create-CursorShortcut -Version $NewVersion -ShortcutLocation "Desktop" # 創建新的
            Remove-Item $oldDesktopShortcut -Force -ErrorAction SilentlyContinue # 刪除舊的
            Write-Host "已更新桌面捷徑指向 $NewVersion" -ForegroundColor Green
        } catch {
            Write-Host "更新桌面捷徑時發生錯誤，您可能需要手動調整或刪除舊捷徑 '$oldDesktopShortcut'" -ForegroundColor Yellow
        }
    }

    # 更新開始選單捷徑
    $startMenuProgramsPath = Join-Path ([Environment]::GetFolderPath("ApplicationData")) "Microsoft\Windows\Start Menu\Programs"
    $oldStartMenuShortcut = Join-Path $startMenuProgramsPath "Cursor_$OldVersion.lnk"
    if (Test-Path $oldStartMenuShortcut) {
        try {
            Create-CursorShortcut -Version $NewVersion -ShortcutLocation "StartMenu" # 創建新的
            Remove-Item $oldStartMenuShortcut -Force -ErrorAction SilentlyContinue # 刪除舊的
            Write-Host "已更新開始選單捷徑指向 $NewVersion" -ForegroundColor Green
        } catch {
            Write-Host "更新開始選單捷徑時發生錯誤，您可能需要手動調整或刪除舊捷徑 '$oldStartMenuShortcut'" -ForegroundColor Yellow
        }
    }

    # 更新版本目錄內的啟動批處理檔 (如果存在)
    $oldBatchFile = Join-Path $newVersionDir "Launch_Cursor_$OldVersion.bat" # 注意：此時 $oldVersionDir 已被重命名為 $newVersionDir
    $newBatchFile = Join-Path $newVersionDir "Launch_Cursor_$NewVersion.bat"
    if (Test-Path $oldBatchFile) {
        try {
            Rename-Item -Path $oldBatchFile -NewName "Launch_Cursor_$NewVersion.bat" -ErrorAction Stop
            # 更新批次檔內容中的版本名 (可選，但更完善)
            $batchContent = Get-Content $newBatchFile -Raw
            $batchContent = $batchContent -replace [regex]::Escape("Cursor $OldVersion 版本"), "Cursor $NewVersion 版本"
            $batchContent = $batchContent -replace [regex]::Escape("start """" ""$((Join-Path $newVersionDir "Program").Replace('\','\\'))\\Cursor.exe"" --user-data-dir=""$((Join-Path $newVersionDir "Config").Replace('\','\\'))"""), "start """" ""$((Join-Path $newVersionDir "Program").Replace('\','\\'))\\Cursor.exe"" --user-data-dir=""$((Join-Path $newVersionDir "Config").Replace('\','\\'))"""
            Set-Content -Path $newBatchFile -Value $batchContent -Encoding UTF8
            Write-Host "已更新版本目錄內的啟動批處理檔名稱與內容。" -ForegroundColor Green
        } catch {
            Write-Host "更新版本目錄內的啟動批處理檔 '$oldBatchFile' 時發生錯誤: $_" -ForegroundColor Yellow
            Write-Host "您可能需要手動重命名或修改該批次檔。" -ForegroundColor Yellow
        }
    } elseif (-not (Test-Path $newBatchFile)) { # 如果舊的不存在，檢查是否需要為新名稱創建一個
         # 如果原本就沒有批次檔，可以考慮是否要自動創建一個，或保持原樣
         # Create-CursorBatchFile -Version $NewVersion # 根據需求決定是否自動創建
    }


    Write-Host "版本 $OldVersion 已成功重命名為 $NewVersion，並已嘗試更新相關連結與檔案。" -ForegroundColor Green
}

# 刪除版本 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< NEW FUNCTION
function Delete-CursorVersion {
    param (
        [Parameter(Mandatory=$true)]
        [string]$VersionToDelete
    )

    # 1. 參數驗證與路徑準備
    if ([string]::IsNullOrWhiteSpace($VersionToDelete)) {
        Write-Host "要刪除的版本名稱不能為空！" -ForegroundColor Red
        return
    }
    $versionDirToDelete = Join-Path $cursorVersionsDir $VersionToDelete

    # 2. 版本存在性檢查
    if (-not (Test-Path $versionDirToDelete)) {
        Write-Host "版本 '$VersionToDelete' 不存在於 '$cursorVersionsDir'！" -ForegroundColor Red
        return
    }

    # 3. 使用者最終確認
    Write-Host "警告：此操作將永久刪除 Cursor 版本 '$VersionToDelete'。" -ForegroundColor Yellow
    Write-Host "這包括位於 '$versionDirToDelete' 的所有程式檔案和設定。" -ForegroundColor Yellow
    $confirmation = Read-Host "您確定要繼續嗎？此操作無法復原。(Y/N)"
    if ($confirmation -ne 'Y' -and $confirmation -ne 'y') {
        Write-Host "刪除操作已取消。" -ForegroundColor Green
        return
    }

    # 4. 處理作用中版本的刪除
    $currentActiveVersion = Get-CurrentCursorVersion
    if ($currentActiveVersion -eq $VersionToDelete) {
        Write-Host "版本 '$VersionToDelete' 是目前系統連結的作用中版本。" -ForegroundColor Yellow
        Write-Host "正在移除系統層級的 Cursor 程式與設定符號連結..." -ForegroundColor Cyan
        try {
            if (Test-Path $cursorProgramPath) {
                $linkItem = Get-Item $cursorProgramPath -ErrorAction SilentlyContinue
                if ($linkItem -and $linkItem.LinkType -eq "SymbolicLink") {
                    Remove-Item $cursorProgramPath -Force -ErrorAction Stop
                    Write-Host "已移除程式路徑符號連結: $cursorProgramPath" -ForegroundColor Green
                } elseif ($linkItem) {
                     Write-Host "警告: $cursorProgramPath 不是預期的符號連結，未移除。" -ForegroundColor Yellow
                }
            }
            if (Test-Path $cursorConfigPath) {
                $linkItem = Get-Item $cursorConfigPath -ErrorAction SilentlyContinue
                if ($linkItem -and $linkItem.LinkType -eq "SymbolicLink") {
                    Remove-Item $cursorConfigPath -Force -ErrorAction Stop
                    Write-Host "已移除設定路徑符號連結: $cursorConfigPath" -ForegroundColor Green
                } elseif ($linkItem) {
                    Write-Host "警告: $cursorConfigPath 不是預期的符號連結，未移除。" -ForegroundColor Yellow
                }
            }
            Write-Host "系統預設的 Cursor 路徑已被清除。建議切換到其他版本或重新初始化。" -ForegroundColor Green
        } catch {
            Write-Host "移除系統符號連結時發生錯誤: $_" -ForegroundColor Red
            Write-Host "您可能需要手動檢查並移除 $cursorProgramPath 和 $cursorConfigPath" -ForegroundColor Yellow
            # 即使符號連結移除失敗，也應詢問是否繼續刪除資料夾
            $continueDespiteError = Read-Host "移除系統連結時出錯。是否仍要繼續刪除版本資料夾 '$VersionToDelete'？ (Y/N)"
            if ($continueDespiteError -ne 'Y' -and $continueDespiteError -ne 'y') {
                Write-Host "刪除操作已中止。" -ForegroundColor Green
                return
            }
        }
    }

    # 5. 刪除版本主資料夾
    Write-Host "正在刪除版本 '$VersionToDelete' 的資料夾內容: '$versionDirToDelete'..." -ForegroundColor Cyan
    try {
        Remove-Item -Path $versionDirToDelete -Recurse -Force -ErrorAction Stop
        Write-Host "版本 '$VersionToDelete' 的資料夾已成功刪除。" -ForegroundColor Green
    } catch {
        Write-Host "錯誤：刪除版本 '$VersionToDelete' 的資料夾 ('$versionDirToDelete') 時發生問題：" -ForegroundColor Red
        Write-Host $_.Exception.Message
        Write-Host "提示：請檢查是否有 Cursor 相關程序正在執行，或檔案是否被鎖定。" -ForegroundColor Yellow
        Write-Host "您可能需要手動清理 '$versionDirToDelete'。" -ForegroundColor Yellow
        # 即使資料夾刪除失敗，也應繼續嘗試清理捷徑等
    }

    # 6. 清理相關的捷徑與批次檔
    Write-Host "正在清理 '$VersionToDelete' 的相關捷徑..." -ForegroundColor Cyan
    # 桌面捷徑
    $desktopShortcutPath = Join-Path ([Environment]::GetFolderPath("Desktop")) "Cursor_$VersionToDelete.lnk"
    if (Test-Path $desktopShortcutPath) {
        try {
            Remove-Item $desktopShortcutPath -Force -ErrorAction Stop
            Write-Host "已刪除桌面捷徑: $desktopShortcutPath" -ForegroundColor Green
        } catch {
            Write-Host "刪除桌面捷徑 '$desktopShortcutPath' 時發生錯誤: $_" -ForegroundColor Yellow
        }
    }

    # 開始選單捷徑
    $startMenuProgramsPath = Join-Path ([Environment]::GetFolderPath("ApplicationData")) "Microsoft\Windows\Start Menu\Programs"
    $startMenuShortcutPath = Join-Path $startMenuProgramsPath "Cursor_$VersionToDelete.lnk"
    if (Test-Path $startMenuShortcutPath) {
        try {
            Remove-Item $startMenuShortcutPath -Force -ErrorAction Stop
            Write-Host "已刪除開始選單捷徑: $startMenuShortcutPath" -ForegroundColor Green
        } catch {
            Write-Host "刪除開始選單捷徑 '$startMenuShortcutPath' 時發生錯誤: $_" -ForegroundColor Yellow
        }
    }
    
    # 批次啟動檔 (通常在版本目錄內，已隨資料夾刪除。此處為防萬一有外部批次檔的檢查，但目前腳本邏輯不會產生)
    # 如果 Create-CursorBatchFile 曾使用 $OutputPath 參數將批次檔建立在版本目錄之外，則需要額外的邏輯。
    # 目前假設批次檔在版本目錄內，已被步驟 5 處理。

    # 7. 最終訊息
    Write-Host "版本 '$VersionToDelete' 已成功從版本管理器中移除 (或嘗試移除)。" -ForegroundColor Green
    Write-Host "請檢查上述訊息確認所有操作是否成功。" -ForegroundColor Cyan
}


# 顯示菜單
function Show-Menu {
    Clear-Host
    $currentVersion = Get-CurrentCursorVersion
    $allVersions = Get-AllCursorVersions
    
    Write-Host "===== Cursor 多版本管理工具 =====" -ForegroundColor Cyan
    Write-Host
    Write-Host "當前系統連結版本: $currentVersion" -ForegroundColor Green
    Write-Host
    Write-Host "可用版本 (位於 $cursorVersionsDir):" -ForegroundColor Cyan
    if ($allVersions.Count -eq 0) {
        Write-Host "  尚未創建任何版本" -ForegroundColor Yellow
    } else {
        foreach ($version in $allVersions) {
            Write-Host "  $version" -ForegroundColor White
        }
    }
    
    Write-Host
    Write-Host "1. 建立新版本" -ForegroundColor Yellow
    Write-Host "2. 切換系統連結版本" -ForegroundColor Yellow
    Write-Host "3. 更新特定版本 (會先切換)" -ForegroundColor Yellow
    Write-Host "4. 還原特定版本到預設路徑 (會先切換)" -ForegroundColor Yellow
    Write-Host "5. 初始化現有 Cursor 安裝為版本管理" -ForegroundColor Yellow
    Write-Host "6. 創建版本啟動捷徑" -ForegroundColor Green
    Write-Host "7. 創建版本啟動批處理檔" -ForegroundColor Green
    Write-Host "8. 直接啟動特定版本 (不切換系統連結)" -ForegroundColor Green
    Write-Host "9. 為所有版本創建啟動器 (捷徑與批次檔)" -ForegroundColor Green
    Write-Host "10. 重命名版本" -ForegroundColor Green
    Write-Host "11. 刪除版本" -ForegroundColor Red # <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< NEW MENU OPTION
    Write-Host "0. 退出" -ForegroundColor Yellow
    Write-Host
    
    $choice = Read-Host "請選擇操作"
    
    switch ($choice) {
        "1" {
            $versionName = Read-Host "請輸入新版本名稱 (例如: 0.45.0 或 custom_build)"
            if ([string]::IsNullOrWhiteSpace($versionName)) {
                Write-Host "版本名稱不能為空！" -ForegroundColor Red
            } else {
                New-CursorVersion -Version $versionName
            }
        }
        "2" {
            if ($allVersions.Count -eq 0) {
                Write-Host "尚未創建任何版本！無法切換。" -ForegroundColor Red
            } else {
                Write-Host "可用版本:" -ForegroundColor Cyan
                for ($i = 0; $i -lt $allVersions.Count; $i++) {
                    Write-Host "$($i+1). $($allVersions[$i])" -ForegroundColor White
                }
                $versionChoice = Read-Host "請選擇要切換到的版本編號"
                try {
                    $versionIdx = [int]$versionChoice - 1
                    if ($versionIdx -ge 0 -and $versionIdx -lt $allVersions.Count) {
                        Switch-CursorVersion -Version $allVersions[$versionIdx]
                    } else {
                        Write-Host "無效的選擇！" -ForegroundColor Red
                    }
                } catch { Write-Host "請輸入有效的數字！" -ForegroundColor Red }
            }
        }
        "3" {
            if ($allVersions.Count -eq 0) {
                Write-Host "尚未創建任何版本！無法更新。" -ForegroundColor Red
            } else {
                Write-Host "可用版本:" -ForegroundColor Cyan
                for ($i = 0; $i -lt $allVersions.Count; $i++) {
                    Write-Host "$($i+1). $($allVersions[$i])" -ForegroundColor White
                }
                $versionChoice = Read-Host "請選擇要更新的版本編號"
                try {
                    $versionIdx = [int]$versionChoice - 1
                    if ($versionIdx -ge 0 -and $versionIdx -lt $allVersions.Count) {
                        Update-CursorVersion -Version $allVersions[$versionIdx]
                    } else {
                        Write-Host "無效的選擇！" -ForegroundColor Red
                    }
                } catch { Write-Host "請輸入有效的數字！" -ForegroundColor Red }
            }
        }
        "4" {
            if ($allVersions.Count -eq 0) {
                Write-Host "尚未創建任何版本！無法還原。" -ForegroundColor Red
            } else {
                Write-Host "可用版本:" -ForegroundColor Cyan
                for ($i = 0; $i -lt $allVersions.Count; $i++) {
                    Write-Host "$($i+1). $($allVersions[$i])" -ForegroundColor White
                }
                $versionChoice = Read-Host "請選擇要還原到預設路徑的版本編號"
                try {
                    $versionIdx = [int]$versionChoice - 1
                    if ($versionIdx -ge 0 -and $versionIdx -lt $allVersions.Count) {
                        Restore-CursorDefaultPath -Version $allVersions[$versionIdx]
                    } else {
                        Write-Host "無效的選擇！" -ForegroundColor Red
                    }
                } catch { Write-Host "請輸入有效的數字！" -ForegroundColor Red }
            }
        }
        "5" {
            Initialize-CursorVersionManagement
        }
        "6" {
            if ($allVersions.Count -eq 0) {
                Write-Host "尚未創建任何版本！無法創建捷徑。" -ForegroundColor Red
            } else {
                Write-Host "可用版本:" -ForegroundColor Cyan
                for ($i = 0; $i -lt $allVersions.Count; $i++) {
                    Write-Host "$($i+1). $($allVersions[$i])" -ForegroundColor White
                }
                $versionChoice = Read-Host "請選擇要為其創建捷徑的版本編號"
                try {
                    $versionIdx = [int]$versionChoice - 1
                    if ($versionIdx -ge 0 -and $versionIdx -lt $allVersions.Count) {
                        $selectedVersion = $allVersions[$versionIdx]
                        $exePath = Join-Path $cursorVersionsDir "$selectedVersion\Program\Cursor.exe"
                        if (-not (Test-Path $exePath)) {
                             Write-Host "版本 $selectedVersion 的 Cursor.exe 不存在，無法創建捷徑。" -ForegroundColor Red
                        } else {
                            $locChoice = Read-Host "創建在桌面(D)還是開始選單(S)？(D/S)"
                            if ($locChoice -match '^[Ss]$') {
                                Create-CursorShortcut -Version $selectedVersion -ShortcutLocation "StartMenu"
                            } else {
                                Create-CursorShortcut -Version $selectedVersion -ShortcutLocation "Desktop"
                            }
                        }
                    } else {
                        Write-Host "無效的選擇！" -ForegroundColor Red
                    }
                } catch { Write-Host "請輸入有效的數字！" -ForegroundColor Red }
            }
        }
        "7" {
            if ($allVersions.Count -eq 0) {
                Write-Host "尚未創建任何版本！無法創建批處理檔。" -ForegroundColor Red
            } else {
                Write-Host "可用版本:" -ForegroundColor Cyan
                for ($i = 0; $i -lt $allVersions.Count; $i++) {
                    Write-Host "$($i+1). $($allVersions[$i])" -ForegroundColor White
                }
                $versionChoice = Read-Host "請選擇要為其創建批處理檔的版本編號"
                try {
                    $versionIdx = [int]$versionChoice - 1
                    if ($versionIdx -ge 0 -and $versionIdx -lt $allVersions.Count) {
                         $selectedVersion = $allVersions[$versionIdx]
                         $exePath = Join-Path $cursorVersionsDir "$selectedVersion\Program\Cursor.exe"
                         if (-not (Test-Path $exePath)) {
                             Write-Host "版本 $selectedVersion 的 Cursor.exe 不存在，無法創建批處理檔。" -ForegroundColor Red
                         } else {
                            Create-CursorBatchFile -Version $selectedVersion
                         }
                    } else {
                        Write-Host "無效的選擇！" -ForegroundColor Red
                    }
                } catch { Write-Host "請輸入有效的數字！" -ForegroundColor Red }
            }
        }
        "8" {
            if ($allVersions.Count -eq 0) {
                Write-Host "尚未創建任何版本！無法啟動。" -ForegroundColor Red
            } else {
                Write-Host "可用版本:" -ForegroundColor Cyan
                for ($i = 0; $i -lt $allVersions.Count; $i++) {
                    Write-Host "$($i+1). $($allVersions[$i])" -ForegroundColor White
                }
                $versionChoice = Read-Host "請選擇要直接啟動的版本編號"
                try {
                    $versionIdx = [int]$versionChoice - 1
                    if ($versionIdx -ge 0 -and $versionIdx -lt $allVersions.Count) {
                        Start-CursorVersion -Version $allVersions[$versionIdx]
                    } else {
                        Write-Host "無效的選擇！" -ForegroundColor Red
                    }
                } catch { Write-Host "請輸入有效的數字！" -ForegroundColor Red }
            }
        }
        "9" {
            Create-AllVersionLaunchers
        }
        "10" {
            if ($allVersions.Count -eq 0) {
                Write-Host "尚未創建任何版本！無法重命名。" -ForegroundColor Red
            } else {
                Write-Host "可用版本:" -ForegroundColor Cyan
                for ($i = 0; $i -lt $allVersions.Count; $i++) {
                    Write-Host "$($i+1). $($allVersions[$i])" -ForegroundColor White
                }
                $versionChoice = Read-Host "請選擇要重命名的版本編號"
                try {
                    $versionIdx = [int]$versionChoice - 1
                    if ($versionIdx -ge 0 -and $versionIdx -lt $allVersions.Count) {
                        $oldVersion = $allVersions[$versionIdx]
                        $newVersion = Read-Host "請輸入 '$oldVersion' 的新版本名稱"
                        if ([string]::IsNullOrWhiteSpace($newVersion)) {
                            Write-Host "新版本名稱不能為空！" -ForegroundColor Red
                        } else {
                            Rename-CursorVersion -OldVersion $oldVersion -NewVersion $newVersion
                        }
                    } else {
                        Write-Host "無效的選擇！" -ForegroundColor Red
                    }
                } catch { Write-Host "請輸入有效的數字！" -ForegroundColor Red }
            }
        }
        "11" { # <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< NEW CASE FOR DELETE
            if ($allVersions.Count -eq 0) {
                Write-Host "目前沒有已建立的版本可供刪除。" -ForegroundColor Yellow
            } else {
                Write-Host "可用版本 (選擇要刪除的版本):" -ForegroundColor Cyan
                for ($i = 0; $i -lt $allVersions.Count; $i++) {
                    Write-Host "$($i+1). $($allVersions[$i])" -ForegroundColor White
                }
                $versionChoiceToDelete = Read-Host "請輸入要刪除的版本編號"
                try {
                    $versionIdxToDelete = [int]$versionChoiceToDelete - 1
                    if ($versionIdxToDelete -ge 0 -and $versionIdxToDelete -lt $allVersions.Count) {
                        $versionToDelete = $allVersions[$versionIdxToDelete]
                        Delete-CursorVersion -VersionToDelete $versionToDelete
                    } else {
                        Write-Host "無效的選擇！沒有此版本編號。" -ForegroundColor Red
                    }
                } catch {
                    Write-Host "請輸入有效的數字作為版本編號！" -ForegroundColor Red
                }
            }
        }
        "0" {
            Write-Host "正在退出腳本..." -ForegroundColor Green
            return
        }
        default {
            Write-Host "無效的選擇！請重新輸入。" -ForegroundColor Red
        }
    }
    pause
    Show-Menu # 遞迴調用以保持菜單顯示，直到選擇退出
}

# 啟動菜單
Show-Menu
