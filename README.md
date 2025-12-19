# Excel Statistic Pro 📊

這是一個專為個案管理與數據統計設計的現代化 Web 工具。使用者可以上傳 XLSX 檔案，自動統計 **性別分佈** 與 **五大年齡級距**，並將結果一鍵同步至 Google Sheets。

## ✨ 特色功能

-   **🎨 Premium UI**: 採用毛玻璃效果 (Glassmorphism)、動態漸層背景與直觀的拖放式上傳介面。
-   **🖱️ 自定義列索引**: 支援手動輸入 Excel 列代號（如 A, B, G...），資料處理更精確。
-   **🔐 環境變數保護**: GAS API URL 透過 `.env` 檔案管理，避免機敏資訊外洩。
-   **📅 月份分頁管理**: 支援依「民國年月」自動建立或切換 Google Sheet 工作表。
-   **📈 垂直資料輸出**: 統計結果以垂直格式（標題在 A 欄，數值在 B 欄）寫入雲端，方便閱讀與彙整。

## 🚀 快速開始

### 1. 本地環境設定
確保您的電腦已安裝 [Node.js](https://nodejs.org/)。

```bash
# 安裝依賴套件
npm install

# 啟動開發伺服器
npm run dev
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

1. **輸入設定**:
    - **工作表月份**: 輸入 5 位數民國年月（例如 `11403`）。
    - **性別列**: 輸入 Excel 中性別欄位的代號（如 `G`）。
    - **年齡列**: 輸入 Excel 中年齡欄位的代號（如 `F`）。
2. **上傳檔案**: 拖放您的 XLSX 檔案至虛線區域，或點擊中心點進行選擇。
3. **預覽結果**: 系統會即時顯示統計的人數。
4. **同步雲端**: 點擊「確認並同步至 Google Sheet」，等待片刻即可在雲端看到結果。

## 🛠️ 技術架構

-   **前端**: Vite, Vanilla JavaScript, CSS3
-   **解析庫**: [SheetJS (xlsx)](https://github.com/SheetJS/sheetjs)
-   **後端**: Google Apps Script (GAS)
-   **統計邏輯**: 
    - 年齡級距: `<= 49`, `50-64`, `65-74`, `75-84`, `>= 85`
    - 性別: 僅計數「男」與「女」

## 📝 專案結構

```text
├── gas/
│   └── code.gs          # GAS 後端程式碼
├── index.html           # 主頁面結構
├── script.js            # 核心邏輯處理
├── style.css            # 視覺樣式
├── .env                 # 環境變數 (機敏資訊，不進入 git)
├── package.json         # Vite 專案設定
└── README.md            # 專案說明文件
```

---
*Developed with ❤️ for efficient data management.*
