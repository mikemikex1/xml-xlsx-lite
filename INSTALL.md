# 安裝指南

本指南說明如何安裝和設定 xml-xlsx-lite 專案。

## 🔧 系統需求

- Node.js 16.x 或更高版本
- npm 8.x 或更高版本
- Git

## 📥 安裝 Node.js

### Windows

1. 前往 [Node.js 官網](https://nodejs.org/)
2. 下載 LTS 版本（推薦）
3. 執行安裝程式，按照指示完成安裝
4. 重新開啟命令提示字元或 PowerShell

### macOS

使用 Homebrew：
```bash
brew install node
```

或從官網下載安裝程式。

### Linux (Ubuntu/Debian)

```bash
curl -fsSL https://deb.nodesource.com/setup_lts.x | sudo -E bash -
sudo apt-get install -y nodejs
```

## ✅ 驗證安裝

安裝完成後，驗證是否成功：

```bash
node --version
npm --version
```

應該會顯示版本號，例如：
```
v18.17.0
9.6.7
```

## 🚀 設定專案

### 1. 安裝依賴

```bash
npm install
```

### 2. 建置專案

```bash
npm run build
```

### 3. 執行測試

```bash
npm test
```

### 4. 瀏覽器測試

```bash
npm run test:browser
```

## 🌐 瀏覽器測試

瀏覽器測試會自動生成 HTML 檔案：

1. 執行 `npm run test:browser`
2. 前往 `test/browser/` 目錄
3. 開啟 `test.html` 檔案
4. 點擊測試按鈕執行測試

## 🔍 常見問題

### Q: 命令 'npm' 無法辨識

**A:** 需要安裝 Node.js，或重新開啟終端機

### Q: 建置失敗

**A:** 檢查 Node.js 版本是否為 16.x 或更高

### Q: 測試失敗

**A:** 確保已執行 `npm install` 安裝依賴

## 📚 下一步

安裝完成後，您可以：

1. 查看 [README.md](README.md) 了解使用方法
2. 查看 [PUBLISHING.md](PUBLISHING.md) 了解如何發佈
3. 開始開發新功能

## 🤝 需要幫助？

如果遇到安裝問題，請：

1. 檢查 [Issues](https://github.com/mikemikex1/xml-xlsx-lite/issues)
2. 建立新的 Issue
3. 聯繫專案維護者
