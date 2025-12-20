# 技術棧 (Tech Stack)

Excel Statistic Pro 採用輕量化的前端架構結合 Google 生態系，旨在提供無須複雜後端部署的行政工具解決方案。

## 1. 前端開發 (Frontend)
- **核心語言**：JavaScript (ES6+), CSS3, HTML5
- **開發伺服器與建置**：[Vite](https://vitejs.dev/)
    - 提供快速的熱更新 (HMR) 與最佳化的生產環境建置。
- **UI 樣式**：純 CSS3 (Vanilla CSS)
    - 採用毛玻璃效果 (Glassmorphism) 與動態漸層背景。
- **數據解析**：[SheetJS (xlsx)](https://sheetjs.com/)
    - 用於在瀏覽器端解析複雜的 Excel 檔案格式。

## 2. 後端整合 (Backend & Integration)
- **平台**：[Google Apps Script (GAS)](https://developers.google.com/apps-script)
    - 作為中間層 API，負責與 Google Sheets 進行數據讀寫。
- **資料儲存**：Google Sheets
    - 使用試算表作為終端資料庫，方便行政人員直接查閱與編輯。

## 3. 環境與部署 (DevOps)
- **套件管理**：[pnpm](https://pnpm.io/) (取代 npm 以獲得更好的效能與嚴格的依賴管理)
- **環境變數**：`.env` (經由 Vite `import.meta.env` 管理)
    - 儲存 GAS 部署網址等敏感資訊。
- **版本控制**：Git
- **部署平台**：GitHub Pages (預計)
