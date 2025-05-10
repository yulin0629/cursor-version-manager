# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 專案概述

這是一個 Cursor 編輯器的多版本管理腳本 (Cursor_Switch.ps1)，允許用戶在多個 Cursor 版本之間輕鬆切換，適合同時使用或測試多個 Cursor 版本的用戶。

## 架構說明

腳本使用 PowerShell 實現，主要功能包括：

1. 創建並管理不同版本的 Cursor
2. 通過符號連結 (Symbolic Links) 實現版本切換
3. 提供桌面捷徑和啟動批處理檔
4. 允許直接啟動特定版本而不切換系統路徑

## 資料夾結構

- `/mnt/d/Cursor_Versions/` - 主目錄
  - `0.50/` 等 - 版本資料夾
    - `Config/` - 配置資料夾
    - `Program/` - 程式資料夾
  - `Cursor_Switch.ps1` - 主腳本

## 常用命令

在 Windows 系統上以管理員權限執行此腳本：

```powershell
# 以管理員身份執行 PowerShell
powershell -ExecutionPolicy Bypass -File D:\Cursor_Versions\Cursor_Switch.ps1
```

## 操作指南

1. **建立新版本**：創建新的 Cursor 版本隔離環境
2. **切換版本**：在不同的 Cursor 版本之間切換
3. **更新版本**：更新特定版本的 Cursor
4. **還原到預設路徑**：將特定版本還原到系統預設路徑
5. **初始化版本管理**：將現有 Cursor 安裝轉換為版本管理
6. **創建捷徑**：為特定版本創建桌面或開始選單捷徑
7. **創建批處理檔**：為特定版本創建啟動批處理檔
8. **直接啟動特定版本**：不切換系統路徑直接啟動版本
9. **為所有版本創建啟動器**：批量創建捷徑和啟動器