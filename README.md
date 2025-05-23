# Cursor 多版本管理工具 [![Ask DeepWiki](https://deepwiki.com/badge.svg)](https://deepwiki.com/yulin0629/cursor-version-manager)

這是一個用於 Windows 環境的 Cursor 編輯器多版本管理腳本，允許用戶在多個 Cursor 版本之間輕鬆切換。

## 功能特色

- 創建及管理多個 Cursor 版本的隔離環境
- 通過符號連結實現版本快速切換
- 提供桌面捷徑和啟動批處理檔
- 允許直接啟動特定版本而不切換系統路徑
- 多版本配置共享或隔離選項

## 適用場景

- 需要同時使用或測試多個 Cursor 版本的用戶
- 想要保留舊版本功能同時嘗試新版本的用戶
- 團隊成員需要使用不同版本確保兼容性的環境

## 資料夾結構

```
Cursor_Versions/
├── 版本1/
│   ├── Config/ - 配置資料夾
│   └── Program/ - 程式資料夾
├── 版本2/
│   ├── Config/
│   └── Program/
└── Cursor_Switch.ps1 - 主腳本
```

## 使用方法

1. 以管理員身份運行 PowerShell
2. 執行腳本：
```powershell
powershell -ExecutionPolicy Bypass -File D:\Cursor_Versions\Cursor_Switch.ps1
```

## 主要功能

1. **建立新版本**：創建新的 Cursor 版本隔離環境
2. **切換版本**：在不同的 Cursor 版本之間切換
3. **更新版本**：更新特定版本的 Cursor
4. **還原到預設路徑**：將特定版本還原到系統預設路徑
5. **初始化版本管理**：將現有 Cursor 安裝轉換為版本管理
6. **創建捷徑**：為特定版本創建桌面或開始選單捷徑
7. **創建批處理檔**：為特定版本創建啟動批處理檔
8. **直接啟動特定版本**：不切換系統路徑直接啟動版本
9. **為所有版本創建啟動器**：批量創建捷徑和啟動器
10. **重命名版本**：修改已建立版本的名稱

## 系統要求

- Windows 10/11
- PowerShell 5.0+
- 管理員權限（用於創建符號連結）