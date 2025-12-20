# Excel Statistic Pro 📊

這是一個專為個案管理與數據統計設計的現代化 Web 工具。使用者可以上傳 XLSX 檔案，自動統計 **個案名單** 或 **服務清冊**，並將結果一鍵同步或累加至 Google Sheets。

## ✨ 特色功能

-   **🎨 Premium UI**: 採用毛玻璃效果 (Glassmorphism)、動態漸層背景與直觀的拖放式上傳介面。
-   **🔄 雙重統計模式**: 支援「個案名單」（性別、年齡、居住狀況、CMS等級）與「服務清冊」（行政區、類別、補助比率）兩種切換模式。
-   **🖱️ 自定義列與標頭**: 支援手動輸入 Excel 列代號，並可指定「服務清冊」的標頭行數，數據讀取更靈活。
-   **📅 縣市月份管理**: 支援依「縣市 + 月份」自動建立或切換 Google Sheet 工作表。
-   **➕ 累加儲存邏輯**: GAS 後端支援累加模式，多次同步時會自動在舊資料下方增加空白行並寫入新結果，不覆蓋舊有數據。
-   **📶 智能排序**: 
    -   行政區依台北/新北特定順序排列。
    -   補助比率固定依 `100% > 95% > 84%` 排序。
-   **🔐 環境變數保護**: GAS API URL 透過 `.env` 檔案管理，避免機敏資訊外洩。
-   **🏙️ 台北老福統計**: 
    - 針對老福個案自動拆分 65+ 與 50-64 歲統計。
    - 自動彙整「身心障礙者」統計資訊，符合最新行政報告需求。

## 🚀 快速開始

### 1. 本地環境設定
確保您的電腦已安裝 [Node.js](https://nodejs.org/)。

```bash
# 安裝依賴套件
pnpm install

# 啟動開發伺服器
pnpm run dev

# 執行單元測試
pnpm test
```

### 2. 環境變數設定
在專案根目錄建立 `.env` 檔案（或修改現有的 `.env`），並填入您的 Google Apps Script 部署網址：
```env
VITE_GAS_URL=https://script.google.com/macros/s/您的部署ID/exec
```

### 3. Google Apps Script 部署
1. 建立一個新的 Google 試算表。
2. 點擊 `擴充功能 > Apps Script`。
3. 將 [`gas/code.gs`](gas/code.gs) 的內容複製並貼上到平台上。
4. 點擊 `部署 > 新增部署`。
5. 類型選擇 **「網頁應用程式」**。
    - **執行身分**：本人
    - **誰可以存取**：所有人 (Anyone)
6. 複製產生的網址並填入上述的 `.env` 檔案中。

## 📖 使用指南

1. **基本設定**:
    - **縣市**: 選擇統計所屬的縣市。
    - **月份**: 輸入 2 位數月份（例如 `03`）。
2. **切換模式**:
    - **個案名單**: 設定性別、年齡、居住狀況、CMS等級所在列。
    - **服務清冊**: 設定標頭行數（預設 5）及行政區、類別、補助比率所在列。
3. **上傳檔案**: 拖放您的 XLSX 檔案至虛線區域（切換模式會自動清除先前上傳的檔案）。
4. **預覽結果**: 系統會依據預設或自定義排序規則顯示統計人數。
    - 台北老福模式會顯示三段式統計（65歲以上、50-64歲、身心障礙者總計）。
5. **同步雲端**: 點擊「確認並同步至 Google Sheet」，資料將累加至雲端試算表的工作表中。

## 🛠️ 技術架構

-   **前端**: Vite, Vanilla JavaScript, CSS3
-   **測試**: Vitest
-   **模組化設計 (位於 `src/js/`)**:
    - `taipei.js`: 台北專屬統計與渲染邏輯。
    - `newtaipei.js`: 新北專屬統計與渲染邏輯。
    - `logic.js`: 通用統計工具。
    - `script.js`: 主控制器。
-   **解析庫**: [SheetJS (xlsx)](https://github.com/SheetJS/sheetjs)
-   **後端**: Google Apps Script (GAS)

## 📝 專案結構

```text
├── src/
│   ├── js/              # 前端邏輯核心
│   │   ├── script.js    # 主控制器 (Orchestrator)
│   │   ├── logic.js     # 共用工具 (Common Utils)
│   │   ├── taipei.js    # 台北統計邏輯
│   │   └── newtaipei.js # 新北統計邏輯
│   ├── css/
│   │   └── style.css    # 視覺樣式
│   └── tests/           # 單元測試 (Vitest)
├── gas/
│   └── code.gs          # GAS 後端程式碼 (支援累加寫入)
├── index.html           # 主頁面結構
├── .env                 # 環境變數 (GAS URL)
├── package.json         # 專案設定與腳本
└── README.md            # 專案說明文件
```

---
*Developed with ❤️ for efficient data management.*
