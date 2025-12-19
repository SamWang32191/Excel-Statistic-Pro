# GitHub Pages 部署指引

本文件說明如何將此專案部署到 GitHub Pages。

## 1. 設定 GitHub Secrets (必要)

由於本專案使用 Google Apps Script 作為後端，您需要將 API URL 儲存為 GitHub Secret，以便在構建時注入。

1. 前往 GitHub 儲存庫頁面。
2. 點擊 **Settings** (設定)。
3. 在左側選單中找到 **Secrets and variables** > **Actions**。
4. 點擊 **New repository secret**。
5. **Name**: `VITE_GAS_URL`
6. **Value**: 輸入您的 Google Apps Script 執行 URL。
7. 點擊 **Add secret**。

## 2. 自動化部署流程

本專案已配置 GitHub Actions (`.github/workflows/deploy.yml`)。

*   **觸發條件**: 每次推送 (Push) 到 `main` 分支時。
*   **動作**: 
    1. 自動執行 `npm install`。
    2. 使用 `VITE_GAS_URL` 安全地執行 `npm run build`。
    3. 將產出的檔案推送到 `gh-pages` 分支。

## 3. GitHub Pages 設定

第一次部署完成後，您需要確認 GitHub Pages 的來源設定：

1. 前往 **Settings** > **Pages**。
2. 在 **Build and deployment** > **Branch** 部分。
3. 選擇 `gh-pages` 分支，並選擇 `/(root)` 資料夾。
4. 點擊 **Save**。

您的網站將在幾分鐘後上線。

## 4. 本地開發環境變數

在本地開發時，請在根目錄建立 `.env` 檔案並加入：
```env
VITE_GAS_URL=您的_GAS_API_URL
```

---

> [!NOTE]
> 如果您修改了儲存庫名稱，請記得更新 `vite.config.js` 中的 `base` 路徑。
