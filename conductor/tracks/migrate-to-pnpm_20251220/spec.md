# Specification: Migrate Project to pnpm

## 1. Overview
本軌跡的目標是將專案的套件管理工具從 `npm` 遷移至 `pnpm`。這將提升安裝速度、節省磁碟空間，並確保依賴管理的嚴格性。

## 2. Scope
- **In Scope:**
    - 移除 `node_modules` 與 `package-lock.json`。
    - 初始化 `pnpm` 並生成 `pnpm-lock.yaml`。
    - 更新 `package.json` 以強制限定使用 `pnpm`。
    - 更新專案文件 (`README.md`) 中的安裝指令。
    - 更新 GitHub Actions workflow (`.github/workflows/deploy.yml`) 以支援 pnpm。
- **Out of Scope:**
    - 升級專案依賴套件的版本 (僅做遷移，不升級)。

## 3. Implementation Details
- **Cleanup:**
    - Delete `node_modules/`
    - Delete `package-lock.json`
- **Installation:**
    - Run `pnpm install` to generate `pnpm-lock.yaml`.
- **Configuration:**
    - Add `"engines": { "pnpm": ">=9" }` to `package.json`.
    - Add `"scripts": { "preinstall": "npx only-allow pnpm" }` to `package.json`.
- **CI/CD:**
    - Update `.github/workflows/deploy.yml` to use `pnpm/action-setup` and run `pnpm install`.

## 4. Success Criteria
- 專案依賴能透過 `pnpm install` 成功安裝。
- `npm install` 會被阻擋或顯示警告。
- 本地開發伺服器 (`pnpm run dev`) 能正常啟動。
- GitHub Actions CI 流程能成功執行並通過。
