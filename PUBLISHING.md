# 發佈指南

本指南說明如何將 xml-xlsx-lite 發佈到 npm 和 GitHub。

## 📦 發佈到 npm

### 1. 準備工作

確保您已經：
- 有 npm 帳號
- 已登入 npm CLI
- 有發佈權限

```bash
# 登入 npm
npm login

# 檢查登入狀態
npm whoami
```

### 2. 更新版本號

使用以下命令之一更新版本號：

```bash
# 修補版本 (1.0.0 -> 1.0.1)
npm version patch

# 次要版本 (1.0.0 -> 1.1.0)
npm version minor

# 主要版本 (1.0.0 -> 2.0.0)
npm version major
```

或者手動編輯 `package.json` 中的版本號。

### 3. 建置專案

```bash
npm run build
```

### 4. 執行測試

```bash
npm test
npm run test:browser
```

### 5. 發佈

```bash
npm publish
```

### 6. 驗證發佈

檢查 npm 網站：https://www.npmjs.com/package/xml-xlsx-lite

## 🚀 發佈到 GitHub

### 1. 建立標籤

```bash
# 建立標籤
git tag v1.0.0

# 推送標籤
git push origin v1.0.0
```

### 2. 建立 Release

1. 前往 GitHub 專案頁面
2. 點擊 "Releases"
3. 點擊 "Create a new release"
4. 選擇剛才建立的標籤
5. 填寫標題和描述
6. 點擊 "Publish release"

### 3. 自動化發佈

本專案已設定 GitHub Actions 自動化發佈：

- 推送到 `main` 分支時會自動：
  - 執行測試
  - 建置專案
  - 發佈到 npm（如果設定了 NPM_TOKEN）
  - 建立 GitHub Release

## 🔧 設定 CI/CD

### 1. 設定 NPM_TOKEN

在 GitHub 專案設定中新增 Secret：

1. 前往 Settings > Secrets and variables > Actions
2. 點擊 "New repository secret"
3. 名稱：`NPM_TOKEN`
4. 值：您的 npm 存取權杖

### 2. 設定 GITHUB_TOKEN

`GITHUB_TOKEN` 會自動提供，無需手動設定。

## 📋 發佈檢查清單

發佈前請確認：

- [ ] 所有測試都通過
- [ ] 程式碼已建置完成
- [ ] 版本號已更新
- [ ] CHANGELOG 已更新（如果有）
- [ ] README 是最新版本
- [ ] 已登入 npm CLI
- [ ] 有發佈權限

## 🚨 緊急回滾

如果發佈後發現問題，可以：

### 1. 從 npm 移除版本

```bash
npm unpublish xml-xlsx-lite@1.0.0
```

**注意**：npm 有 72 小時的移除限制。

### 2. 從 GitHub 移除標籤

```bash
git tag -d v1.0.0
git push origin :refs/tags/v1.0.0
```

### 3. 建立修復版本

```bash
npm version patch
npm run build
npm test
npm publish
```

## 📚 相關資源

- [npm 發佈指南](https://docs.npmjs.com/packages-and-modules/contributing-packages-to-the-registry)
- [GitHub Releases](https://docs.github.com/en/repositories/releasing-projects-on-github)
- [GitHub Actions](https://docs.github.com/en/actions)

## 🤝 需要幫助？

如果遇到發佈問題，請：

1. 檢查 [Issues](https://github.com/mikemikex1/xml-xlsx-lite/issues)
2. 建立新的 Issue
3. 聯繫專案維護者
