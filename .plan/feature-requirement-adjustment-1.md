---
goal: 調整 Excel 工具需求：移除年份輸入、新增縣市選單、修改工作表命名規則
version: 1.1
date_created: 2025-12-19
status: 'Completed'
tags: ['feature', 'refactor']
---

# Introduction (介紹)

![Status: Completed](https://img.shields.io/badge/status-Completed-green)

本計畫旨在調整現有 Excel 統計工具的輸入介面與處理邏輯，以符合新的業務需求。主要變更包含簡化日期輸入（僅保留月份）、新增作業縣市選擇功能，以及自動化 Google Sheet 工作表命名規則。

## 1. Requirements & Constraints (需求與限制)

- **REQ-001**: 移除「民國年」輸入欄位，僅保留「月份」輸入。
  - 輸入格式應為 2 位數月份 (例如: 03, 12)。
- **REQ-002**: 新增「縣市」下拉選單。
  - 選項包含：「台北」、「新北」。
  - 預設值為空或提示選擇。
  - **強制檢核**：必須有選取縣市才能進行同步。
- **REQ-003**: Google Sheet 工作表名稱格式須為「縣市月份」。
  - 範例：若選擇「台北」且月份輸入「03」，則名稱為「台北03」。
- **CON-001**: 必須維持現有的 UI 設計風格 (CSS)。
- **CON-002**: 必須確保送出至 GAS (Google Apps Script) 的資料結構正確，僅 `sheetName` 改變，`data` 內容維持統計結果。

## 2. Implementation Steps (實作步驟)

### Implementation Phase 1 (實作階段 1) - UI 調整

- GOAL-001: 修改 HTML 結構以符合新的輸入需求。

| Task (任務) | Description (描述) | Completed (已完成) | Date (日期) |
| ----------- | ------------------ | ------------------ | ----------- |
| TASK-001    | 修改 `index.html` 中的日期輸入欄位 | Yes | 2025-12-19 |
|             | - 將 `id="rocDate"` 的標籤改為「月份」 | Yes | 2025-12-19 |
|             | - `placeholder` 改為 "例如: 03" | Yes | 2025-12-19 |
|             | - `maxlength` 改為 "2" | Yes | 2025-12-19 |
| TASK-002    | 在 `index.html` 中新增縣市下拉選單 | Yes | 2025-12-19 |
|             | - 在 `.config-section` 中新增 `select` 元素 | Yes | 2025-12-19 |
|             | - `id="citySelect"` | Yes | 2025-12-19 |
|             | - 選項：`<option value="" disabled selected>請選擇縣市</option>`, `<option value="台北">台北</option>`, `<option value="新北">新北</option>` | Yes | 2025-12-19 |

### Implementation Phase 2 (實作階段 2) - 邏輯更新

- GOAL-002: 更新 JavaScript 邏輯以處理新的輸入驗證與資料格式。

| Task (任務) | Description (描述) | Completed (已完成) | Date (日期) |
| ----------- | ------------------ | ------------------ | ----------- |
| TASK-003    | 更新 `script.js` 中的 DOM 元素參照 | Yes | 2025-12-19 |
|             | - 新增 `const citySelect = document.getElementById('citySelect');` | Yes | 2025-12-19 |
| TASK-004    | 更新 `checkReady` 驗證邏輯 | Yes | 2025-12-19 |
|             | - 驗證 `citySelect.value` 是否有效（非空）。 | Yes | 2025-12-19 |
|             | - 調整日期驗證邏輯（長度改為 >= 1 或 2，視需求而定，建議 2）。 | Yes | 2025-12-19 |
|             | - 當驗證失敗時，確保按鈕為 `disabled`。 | Yes | 2025-12-19 |
| TASK-005    | 更新 `uploadBtn.onclick` 事件處理 | Yes | 2025-12-19 |
|             | - **防禦性檢核**：在送出前再次檢查 `citySelect.value` 是否為空。若為空則 `alert` 或 `showStatus` 提示並中止。 | Yes | 2025-12-19 |
|             | - 組合 `sheetName`: `const sheetName = citySelect.value + rocDateInput.value;` | Yes | 2025-12-19 |

## 3. Alternatives (替代方案)

- **ALT-001**: 使用原本的民國年欄位由程式自動擷取月份。
  - 原因不採用：使用者明確要求「不輸入民國年」，直接輸入月份較直覺且減少輸入錯誤。
- **ALT-002**: 縣市欄位使用文字輸入框。
  - 原因不採用：下拉選單可避免錯字（如：臺北 vs 台北），確保命名一致性。

## 4. Dependencies (依賴項)

- 無新增外部依賴。既有的 `style.css` 需支援新的 `select` 元素樣式（假設既有 CSS 可通用，若無須微調）。

## 5. Files (檔案)

- **FILE-001**: `index.html` - 修改輸入介面。
- **FILE-002**: `script.js` - 修改邏輯與驗證。

## 6. Testing (測試)

- **TEST-001**: 介面顯示測試
  - 確認「月份」欄位 `placeholder` 正確。
  - 確認「縣市」下拉選單包含「台北」與「新北」，且預設為「請選擇縣市」。
- **TEST-002**: 輸入驗證與按鈕狀態測試
  - 狀態 1：未選縣市，按鈕 Disabled。
  - 狀態 2：已選縣市、未輸月份，按鈕 Disabled。
  - 狀態 3：皆已完成，按鈕 Enabled。
- **TEST-003**: 點擊同步強制檢核測試
  - (Edge Case) 嘗試移除 `disabled` 屬性後點擊按鈕，應觸發 `TASK-005` 的防禦性檢核，不進行 API 呼叫。
- **TEST-004**: 功能測試
  - 模擬上傳與送出，攔截 `fetch` 請求，確認送出的 payload 中 `sheetName` 是否為「縣市月份」格式（例如「台北03」）。

## 7. Risks & Assumptions (風險與假設)

- **ASSUMPTION-001**: GAS 端不需要修改即可接收新的 `sheetName` 格式。
- **RISK-001**: 若使用者輸入非數字的月份（如 "三月"），目前邏輯僅做字串串接。依賴人工輸入正確性。

## 8. Related Specifications / Further Reading (相關規格 / 延伸閱讀)

- 無
