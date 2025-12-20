# Implementation Plan: Migrate Project to pnpm

本計畫將引導專案從 npm 遷移至 pnpm，並確保 CI/CD 流程同步更新。

## Phase 1: Cleanup & Initialization [checkpoint: 8ba235b]
- [x] Task: 刪除舊有的套件管理檔案 (node_modules, package-lock.json) <!-- b9f2725 -->
- [x] Task: 使用 pnpm 初始化依賴並生成 pnpm-lock.yaml <!-- 5d47fb9 -->
- [x] Task: 驗證本地開發伺服器啟動是否正常 (`pnpm run dev`) <!-- 66e54a8 -->
- [x] Task: Conductor - User Manual Verification 'Phase 1: Cleanup & Initialization' (Protocol in workflow.md) <!-- 8ba235b -->

## Phase 2: Project Configuration & Enforcement [checkpoint: 7ab107a]
- [x] Task: 更新 package.json 加入 engines 限制與 preinstall 腳本 <!-- c4da4d7 -->
- [x] Task: 更新 README.md 將 npm 指令全面替換為 pnpm 指令 <!-- 6b3765a -->
- [x] Task: 測試使用 npm 安裝是否會被成功阻擋 <!-- d01b9d6 -->
- [x] Task: Conductor - User Manual Verification 'Phase 2: Project Configuration & Enforcement' (Protocol in workflow.md) <!-- 7ab107a -->

## Phase 3: CI/CD Migration
- [x] Task: 修改 .github/workflows/deploy.yml 以支援 pnpm action-setup <!-- 77d1868 -->
- [ ] Task: 提交變更並驗證 GitHub Actions 是否能正確安裝並部署
- [ ] Task: Conductor - User Manual Verification 'Phase 3: CI/CD Migration' (Protocol in workflow.md)
