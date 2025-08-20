# 專案設定檢查清單

在發佈 xml-xlsx-lite 到 GitHub 和 npm 之前，請完成以下設定：

## 🔧 基本設定

### 1. 更新 package.json

- [ ] 修改 `author` 欄位為您的姓名
- [ ] 更新 `repository.url` 為您的 GitHub 專案 URL
- [ ] 更新 `bugs.url` 和 `homepage` 為您的 GitHub 專案 URL

### 2. 更新 README.md

- [ ] 修改所有 GitHub 連結為您的專案 URL
- [ ] 檢查使用說明是否正確
- [ ] 確認 API 文件與程式碼一致

### 3. 更新發佈檔案

- [ ] 修改 `PUBLISHING.md` 中的 GitHub 專案連結
- [ ] 檢查發佈步驟是否正確

## 🚀 GitHub 設定

### 1. 建立專案

- [ ] 在 GitHub 建立新的 repository
- [ ] 命名為 `xml-xlsx-lite`
- [ ] 設為 Public
- [ ] 不要初始化 README（我們已經有了）

### 2. 推送程式碼

```bash
git init
git add .
git commit -m "Initial commit: xml-xlsx-lite project"
git branch -M main
git remote add origin https://github.com/mikemikex1/xml-xlsx-lite.git
git push -u origin main
```

### 3. 設定 GitHub Pages（可選）

- [ ] 前往 Settings > Pages
- [ ] Source 選擇 "Deploy from a branch"
- [ ] Branch 選擇 "main" 和 "/docs"

## 📦 npm 設定

### 1. 登入 npm

```bash
npm login
```

### 2. 檢查專案名稱可用性

```bash
npm search xml-xlsx-lite
```

如果名稱已被使用，請在 `package.json` 中修改 `name` 欄位。

### 3. 設定 npm 發佈

- [ ] 確認 `package.json` 中的 `files` 欄位包含必要檔案
- [ ] 檢查 `main`、`module`、`types` 欄位是否正確

## 🧪 測試設定

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
npm run test:browser
```

### 4. 檢查建置輸出

- [ ] `dist/` 目錄已建立
- [ ] 包含 `index.js`、`index.esm.js`、`index.d.ts`
- [ ] 檔案大小合理

## 🔒 安全性設定

### 1. 檢查依賴

```bash
npm audit
```

- [ ] 沒有高風險漏洞
- [ ] 依賴版本是最新的

### 2. 設定 .gitignore

- [ ] 排除 `node_modules/`
- [ ] 排除 `dist/`
- [ ] 排除測試輸出檔案

## 📋 發佈前檢查

### 1. 程式碼品質

- [ ] 沒有 console.log 或 debug 程式碼
- [ ] 錯誤處理完善
- [ ] 型別定義完整

### 2. 文件完整性

- [ ] README.md 完整且正確
- [ ] API 文件與程式碼一致
- [ ] 使用範例可執行

### 3. 測試覆蓋

- [ ] 基本功能測試通過
- [ ] 瀏覽器測試通過
- [ ] 錯誤情況測試通過

## 🚀 發佈步驟

### 1. 建立標籤

```bash
git tag v1.0.0
git push origin v1.0.0
```

### 2. 發佈到 npm

```bash
npm publish
```

### 3. 建立 GitHub Release

- [ ] 前往 GitHub Releases 頁面
- [ ] 建立新的 Release
- [ ] 選擇剛才的標籤
- [ ] 填寫 Release 說明

## 🔍 發佈後檢查

### 1. npm 檢查

- [ ] 前往 https://www.npmjs.com/package/xml-xlsx-lite
- [ ] 確認版本號正確
- [ ] 確認檔案大小合理

### 2. GitHub 檢查

- [ ] Release 已建立
- [ ] 標籤已推送
- [ ] 程式碼已同步

### 3. 功能驗證

- [ ] 從 npm 安裝測試
- [ ] 基本功能正常
- [ ] 瀏覽器環境正常

## 📞 需要幫助？

如果遇到問題：

1. 檢查 [GitHub Issues](https://github.com/mikemikex1/xml-xlsx-lite/issues)
2. 查看 [npm 發佈指南](https://docs.npmjs.com/packages-and-modules/contributing-packages-to-the-registry)
3. 聯繫專案維護者

---

**完成所有檢查項目後，您的專案就可以正式發佈了！** 🎉
