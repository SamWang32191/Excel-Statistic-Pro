# Implementation Plan - Case and Service List Adjustment

本計畫旨在根據 `spec.md` 實作「個案清單」與「服務清單」的功能調整，包含多工作表讀取、台北專屬過濾邏輯、欄位映射優化以及 UI 呈現調整。

## Phase 1: Infrastructure & Data Logic (TDD)
本階段專注於核心邏輯的調整，確保能正確解析 Excel 多工作表並套用正確的欄位映射與過濾規則。

- [ ] Task: 建立 Excel 測試資料 (台北/新北多工作表) 用於單元測試
- [ ] Task: 實作多工作表識別邏輯 (搜尋 "台北"、"新北" 關鍵字)
- [ ] Task: 實作台北資料過濾邏輯 (僅保留「老福」與「身障」)
- [ ] Task: 更新欄位映射邏輯，支援台北與新北不同的預設代號
- [ ] Task: 重構數據匯總結構，支援按縣市與子類別 (身障/老福) 拆分數據
- [ ] Task: Conductor - User Manual Verification 'Phase 1: Data Logic' (Protocol in workflow.md)

## Phase 2: User Interface Updates (UI)
本階段將調整 HTML/CSS 與 JS 控制邏輯，以符合新的 UI 需求。

- [ ] Task: 更新「個案清單」欄位設定 UI，新增「類別列」輸入框並動態更新預設值
- [ ] Task: 實作「服務清單」台北子類別選擇器 (Radio Buttons)
- [ ] Task: 調整「個案清單」預覽區塊，支援「台北身障」與「台北老福」分組顯示
- [ ] Task: 整合月份輸入值與工作表名稱生成邏輯 (`{行政區}{類別}{月份}`)
- [ ] Task: Conductor - User Manual Verification 'Phase 2: UI Updates' (Protocol in workflow.md)

## Phase 3: Integration & End-to-End Testing
本階段進行前後端整合測試，確保資料能正確送往 GAS。

- [ ] Task: 驗證送到 GAS 的 Payload 結構，確保工作表名稱符合規格
- [ ] Task: 執行端到端測試 (上傳 Excel -> 預覽 -> 模擬同步至 GAS)
- [ ] Task: 最終程式碼清理與風格檢查 (Linting)
- [ ] Task: Conductor - User Manual Verification 'Phase 3: Final Integration' (Protocol in workflow.md)

[checkpoint: pending]
