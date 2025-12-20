# Implementation Plan: Case and Service List Adjustment

## Phase 1: Global UI Layout Refactoring
- [x] Task: 重構功能切換按鈕佈局。將「個案名單」與「服務清冊」移至頁面頂端。
- [x] Task: 實作縣市選單的條件顯示邏輯（僅在服務清冊模式顯示）。
- [ ] Task: Conductor - User Manual Verification 'Phase 1' (Protocol in workflow.md)

## Phase 2: Case List UI & Settings Enhancement
- [x] Task: 在個案名單模式下，同時顯示「台北」與「新北」的欄位設定區塊。
- [x] Task: 更新欄位設定的預設值（新北: G/F/O/J；台北: B/H/G/P/K）。
- [ ] Task: Conductor - User Manual Verification 'Phase 2' (Protocol in workflow.md)
- [ ] Task: Conductor - User Manual Verification 'Phase 2' (Protocol in workflow.md)

## Phase 3: Multi-Sheet Loading & Filtering Logic (Case List)
- [x] Task: [TDD] 撰寫多工作表搜尋與識別邏輯的測試。
- [x] Task: 實作工作表名稱搜尋（台北/新北）與邏輯分流。
- [x] Task: [TDD] 撰寫台北資料過濾邏輯測試（類別 B 欄位篩選「老福/身障」）。
- [x] Task: 實作台北專屬的過濾與分類邏輯。
- [ ] Task: Conductor - User Manual Verification 'Phase 3' (Protocol in workflow.md)

## Phase 4: Preview UI & Missing Data Handling
- [x] Task: [TDD] 撰寫預覽資料分組（台北身障/台北老福/新北）的單元測試。
- [~] Task: 實作預覽畫面的分組顯示邏輯，包含月份標題。
- [ ] Task: 實作「無資料」狀態的顯示邏輯（針對缺失工作表的縣市）。
- [ ] Task: Conductor - User Manual Verification 'Phase 4' (Protocol in workflow.md)

## Phase 5: Service List Adjustments & GAS Integration
- [x] Task: 在服務清冊模式選擇台北時，新增身障/老福的 Radio Button 選擇器。
- [x] Task: 更新 GAS 傳送邏輯，僅傳送有資料的縣市，並套用新的工作表命名規則。
- [ ] Task: Conductor - User Manual Verification 'Phase 5' (Protocol in workflow.md)