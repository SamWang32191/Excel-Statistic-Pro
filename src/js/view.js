import { identifyWorksheets, colLetterToIndex, processServiceStats } from './logic.js';
import { processTaipeiCaseStats, formatTaipeiCaseOutput, renderTaipeiCasePreview } from './taipei.js';
import { processNewTaipeiCaseStats, formatNewTaipeiCaseOutput, renderNewTaipeiCasePreview } from './newtaipei.js';

document.addEventListener('DOMContentLoaded', () => {
    // 取得 DOM 元素
    const dropZone = document.getElementById('dropZone');
    const fileInput = document.getElementById('fileInput');
    const cityButtons = document.querySelectorAll('.btn-city');
    const rocDateInput = document.getElementById('rocDate');

    // 台北輸入欄位
    const tpCategoryColInput = document.getElementById('tpCategoryCol');
    const tpGenderColInput = document.getElementById('tpGenderCol');
    const tpAgeColInput = document.getElementById('tpAgeCol');
    const tpLivingColInput = document.getElementById('tpLivingCol');
    const tpCmsColInput = document.getElementById('tpCmsCol');

    // 新北輸入欄位
    const ntGenderColInput = document.getElementById('ntGenderCol');
    const ntAgeColInput = document.getElementById('ntAgeCol');
    const ntLivingColInput = document.getElementById('ntLivingCol');
    const ntCmsColInput = document.getElementById('ntCmsCol');

    const previewSection = document.getElementById('previewSection');
    const previewContainer = document.getElementById('previewContainer');
    const uploadBtn = document.getElementById('uploadBtn');
    const statusMessage = document.getElementById('statusMessage');
    const fileInfo = document.getElementById('fileInfo');
    const removeFileBtn = document.getElementById('removeFile');

    const caseListCols = document.getElementById('caseListCols');
    const serviceRegisterCols = document.getElementById('serviceRegisterCols');
    const cityInputGroup = document.getElementById('cityInputGroup');
    const tpServiceCategoryGroup = document.getElementById('tpServiceCategoryGroup');
    const headerRowIdxInput = document.getElementById('headerRowIdx');
    const districtColInput = document.getElementById('districtCol');
    const categoryColInput = document.getElementById('categoryCol');
    const subsidyColInput = document.getElementById('subsidyCol');

    // 服務清冊統計容器
    const serviceStatsGrid = document.getElementById('serviceStatsGrid');
    const districtStatsDiv = document.getElementById('districtStats');
    const categoryStatsDiv = document.getElementById('categoryStats');
    const subsidyStatsDiv = document.getElementById('subsidyStats');

    const modeButtons = document.querySelectorAll('.btn-switch');

    /**
     * 更新台北服務項目類別選單的顯示狀態
     */
    const updateTpServiceCategoryVisibility = () => {
        const show = currentMode === 'service' && currentCity === '台北';
        tpServiceCategoryGroup.classList.toggle('hidden', !show);
    };

    // 從環境變數取得 GAS URL (Vite 前綴)
    const GAS_URL = import.meta.env.VITE_GAS_URL;

    let currentFile = null;
    let currentMode = 'case'; // 'case' (個案) 或 'service' (服務)
    let currentCity = '台北'; // 預設縣市

    // 初始狀態設置
    cityInputGroup.classList.toggle('hidden', currentMode === 'case');
    updateTpServiceCategoryVisibility();

    // --- 模式切換邏輯 (個案清單 / 服務清冊) ---
    modeButtons.forEach(btn => {
        btn.onclick = () => {
            if (btn.dataset.mode === currentMode) return;

            // 更新 UI 狀態
            modeButtons.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            currentMode = btn.dataset.mode;

            // 切換輸入欄位顯示
            caseListCols.classList.toggle('hidden', currentMode !== 'case');
            serviceRegisterCols.classList.toggle('hidden', currentMode !== 'service');
            cityInputGroup.classList.toggle('hidden', currentMode === 'case');
            updateTpServiceCategoryVisibility();

            // 重置上傳檔案與預覽
            resetUpload();
        };
    });

    // --- 縣市切換邏輯 (僅限服務模式) ---
    cityButtons.forEach(btn => {
        btn.onclick = () => {
            if (btn.dataset.city === currentCity) return;

            cityButtons.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            currentCity = btn.dataset.city;

            updateTpServiceCategoryVisibility();
            resetUpload(); // 切換縣市時清空檔案與預覽
            checkReady();
        };
    });

    // --- 服務項目類別切換邏輯 (僅限台北) ---
    const tpServiceCategoryRadios = document.querySelectorAll('input[name="tpServiceCategory"]');
    tpServiceCategoryRadios.forEach(radio => {
        radio.addEventListener('change', () => {
            resetUpload(); // 切換類別時清空檔案與預覽
        });
    });

    // --- 檔案上傳與拖放邏輯 ---
    dropZone.onclick = () => fileInput.click();

    dropZone.ondragover = (e) => {
        e.preventDefault();
        dropZone.classList.add('dragover');
    };

    dropZone.ondragleave = () => dropZone.classList.remove('dragover');

    dropZone.ondrop = (e) => {
        e.preventDefault();
        dropZone.classList.remove('dragover');
        if (e.dataTransfer.files.length) {
            handleFile(e.dataTransfer.files[0]);
        }
    };

    fileInput.onchange = (e) => {
        if (e.target.files.length) {
            handleFile(e.target.files[0]);
        }
    };

    removeFileBtn.onclick = (e) => {
        e.stopPropagation();
        resetUpload();
    };

    /**
     * 處理上傳的檔案
     * @param {File} file - 上傳的檔案物件
     */
    function handleFile(file) {
        if (!file.name.match(/\.(xlsx|xls)$/)) {
            showStatus('不支援的檔案格式', 'error');
            return;
        }
        currentFile = file;
        triggerProcess();
    }

    /**
     * 觸發檔案處理流程
     */
    function triggerProcess() {
        if (!currentFile) return;

        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });

                if (currentMode === 'case') {
                    const { taipeiSheetName, newTaipeiSheetName } = identifyWorksheets(workbook);
                    const result = {
                        taipei: taipeiSheetName ? XLSX.utils.sheet_to_json(workbook.Sheets[taipeiSheetName], { header: 1 }) : null,
                        newTaipei: newTaipeiSheetName ? XLSX.utils.sheet_to_json(workbook.Sheets[newTaipeiSheetName], { header: 1 }) : null
                    };
                    processCaseList(result, currentFile.name, taipeiSheetName, newTaipeiSheetName);
                } else {
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                    processServiceRegister(jsonData, currentFile.name);
                }
            } catch (err) {
                console.error(err);
                showStatus('讀取檔案失敗: ' + err.message, 'error');
            }
        };
        reader.readAsArrayBuffer(currentFile);
    }

    /**
     * 重置上傳狀態與 UI
     */
    function resetUpload() {
        fileInput.value = '';
        currentFile = null;
        fileInfo.classList.add('hidden');
        document.querySelector('.upload-content').classList.remove('hidden');
        previewSection.classList.add('hidden');
        uploadBtn.disabled = true;
        showStatus('請選擇縣市、輸入月份並上傳檔案', '');
    }

    // --- 統計邏輯區塊 ---

    /**
     * 處理個案清單統計
     * @param {Object} data - 工作表資料 { taipei, newTaipei }
     * @param {string} fileName - 檔案名稱
     * @param {string} taipeiSheetName - 台北工作表名稱
     * @param {string} newTaipeiSheetName - 新北工作表名稱
     */
    function processCaseList(data, fileName, taipeiSheetName, newTaipeiSheetName) {
        previewContainer.innerHTML = '';
        if (serviceStatsGrid) serviceStatsGrid.classList.add('hidden');
        previewContainer.classList.remove('hidden');

        const allPayloads = [];
        const month = rocDateInput.value;

        // 1. 處理新北
        if (data.newTaipei) {
            const ntStats = processNewTaipeiCaseStats(data.newTaipei, {
                gender: ntGenderColInput.value,
                age: ntAgeColInput.value,
                living: ntLivingColInput.value,
                cms: ntCmsColInput.value
            });
            const sheetName = `新北${month}`;
            previewContainer.appendChild(renderNewTaipeiCasePreview(ntStats, sheetName));
            allPayloads.push({ sheetName, rows: formatNewTaipeiCaseOutput(ntStats) });
        }

        // 2. 處理台北
        if (data.taipei) {
            const tpConfig = {
                category: tpCategoryColInput.value,
                gender: tpGenderColInput.value,
                age: tpAgeColInput.value,
                living: tpLivingColInput.value,
                cms: tpCmsColInput.value
            };

            const elderlyTotal = processTaipeiCaseStats(data.taipei, tpConfig, '老福');
            const elderly65Plus = processTaipeiCaseStats(data.taipei, tpConfig, '老福', { min: 65 });
            const elderly5064 = processTaipeiCaseStats(data.taipei, tpConfig, '老福', { min: 50, max: 64 });
            const disabled = processTaipeiCaseStats(data.taipei, tpConfig, '身障');

            const sheetNameElderly = `台北老福${month}`;
            const sheetNameDisabled = `台北身障${month}`;

            previewContainer.appendChild(renderTaipeiCasePreview(elderlyTotal, elderly65Plus, elderly5064, disabled, sheetNameElderly, sheetNameDisabled));

            const tpOutputs = formatTaipeiCaseOutput(elderlyTotal, elderly65Plus, elderly5064, disabled);
            allPayloads.push({ sheetName: sheetNameElderly, rows: tpOutputs.elderlyRows });
            allPayloads.push({ sheetName: sheetNameDisabled, rows: tpOutputs.disabledRows });
        }

        uploadBtn.dataset.outputPayload = JSON.stringify(allPayloads);
        finalizePreview(fileName);
    }

    /**
     * 處理服務清冊統計
     * @param {Array} rows - 工作表資料列
     * @param {string} fileName - 檔案名稱
     */
    function processServiceRegister(rows, fileName) {
        const config = {
            headerRowIdx: parseInt(headerRowIdxInput.value || '5'),
            district: districtColInput.value || 'U',
            category: categoryColInput.value || 'G',
            subsidy: subsidyColInput.value || 'O'
        };

        const { districtStats, categoryStats, subsidyStats } = processServiceStats(rows, config);

        renderServiceStats(districtStats, categoryStats, subsidyStats);
        finalizePreview(fileName);
    }

    /**
     * 完成統計預覽並更新 UI
     * @param {string} fileName - 檔案名稱
     */
    function finalizePreview(fileName) {
        fileInfo.classList.remove('hidden');
        fileInfo.querySelector('.file-name').textContent = fileName;
        document.querySelector('.upload-content').classList.add('hidden');
        previewSection.classList.remove('hidden');
        showStatus('檔案已就緒，請確認統計結果', 'success');
        checkReady();
    }

    /**
     * 渲染服務統計摘要
     * @param {Object} districtStats - 行政區統計
     * @param {Object} categoryStats - 類別統計
     * @param {Object} subsidyStats - 補助統計
     */
    function renderServiceStats(districtStats, categoryStats, subsidyStats) {
        previewContainer.classList.add('hidden');
        serviceStatsGrid.classList.remove('hidden');

        // 排序行政區順序
        const orderNT = ['三重區', '蘆洲區', '五股區', '板橋區', '土城區', '新莊區', '中和區', '永和區', '樹林區', '泰山區', '新店區'];
        const orderTP = ['松山區', '信義區', '大安區', '中山區', '中正區', '大同區', '萬華區', '文山區', '南港區', '內湖區', '士林區', '北投區'];
        const baseOrder = currentCity === '台北' ? orderTP : orderNT;

        const sortedDistricts = Object.keys(districtStats).sort((a, b) => {
            const idxA = baseOrder.indexOf(a);
            const idxB = baseOrder.indexOf(b);
            if (idxA !== -1 && idxB !== -1) return idxA - idxB;
            if (idxA !== -1) return -1;
            if (idxB !== -1) return 1;
            return a.localeCompare(b, 'zh-hant');
        });

        districtStatsDiv.innerHTML = sortedDistricts
            .map(label => `<div class="stat-item"><span>${label}</span> <span class="stat-value">${districtStats[label]}</span></div>`)
            .join('') || '<div class="stat-item"><span>無資料</span></div>';

        // 排序類別順序
        const sortedCategories = Object.keys(categoryStats).sort((a, b) => a.localeCompare(b, 'zh-hant'));
        categoryStatsDiv.innerHTML = sortedCategories
            .map(label => `<div class="stat-item"><span>${label}</span> <span class="stat-value">${categoryStats[label]}</span></div>`)
            .join('') || '<div class="stat-item"><span>無資料</span></div>';

        // 排序補助比率順序
        const subsidyOrder = ['100%', '95%', '84%'];
        const sortedSubsidies = Object.keys(subsidyStats).sort((a, b) => {
            const idxA = subsidyOrder.indexOf(a);
            const idxB = subsidyOrder.indexOf(b);
            if (idxA !== -1 && idxB !== -1) return idxA - idxB;
            if (idxA !== -1) return -1;
            if (idxB !== -1) return 1;
            return a.localeCompare(b);
        });
        subsidyStatsDiv.innerHTML = sortedSubsidies
            .map(label => `<div class="stat-item"><span>${label}</span> <span class="stat-value">${subsidyStats[label]}</span></div>`)
            .join('') || '<div class="stat-item"><span>無資料</span></div>';

        const outputRows = buildServiceOutputRows(districtStats, sortedDistricts, categoryStats, sortedCategories, subsidyStats, sortedSubsidies);
        uploadBtn.dataset.outputRows = JSON.stringify(outputRows);
    }

    /**
     * 建立服務統計輸出資料列
     * @returns {Array} 格式化後的資料列陣列
     */
    function buildServiceOutputRows(districtStats, sortedDistricts, categoryStats, sortedCategories, subsidyStats, sortedSubsidies) {
        const rows = [['【行政區統計】', '']];
        sortedDistricts.forEach(label => rows.push([label, districtStats[label]]));
        rows.push(['', ''], ['【服務項目類別】', '']);
        sortedCategories.forEach(label => rows.push([label, categoryStats[label]]));
        rows.push(['', ''], ['【補助比率】', '']);
        sortedSubsidies.forEach(label => rows.push([label, subsidyStats[label]]));
        return rows;
    }

    // --- 驗證邏輯區塊 ---

    /**
     * 檢查必要資訊是否齊全以啟用同步按鈕
     */
    function checkReady() {
        const hasCity = currentMode === 'case' || currentCity !== '';
        const hasDate = rocDateInput.value.length >= 2;
        const hasFile = !!currentFile;

        let hasCols = false;
        if (currentMode === 'case') {
            const hasTpCols = tpCategoryColInput.value.trim() && tpGenderColInput.value.trim() && tpAgeColInput.value.trim() && tpLivingColInput.value.trim() && tpCmsColInput.value.trim();
            const hasNtCols = ntGenderColInput.value.trim() && ntAgeColInput.value.trim() && ntLivingColInput.value.trim() && ntCmsColInput.value.trim();
            hasCols = hasTpCols && hasNtCols;
        } else {
            hasCols = districtColInput.value.trim() && categoryColInput.value.trim() && subsidyColInput.value.trim();
        }

        uploadBtn.disabled = !(hasCity && hasDate && hasCols && hasFile);
    }

    // 監聽各個輸入項目的變更
    [rocDateInput, tpCategoryColInput, tpGenderColInput, tpAgeColInput, tpLivingColInput, tpCmsColInput, ntGenderColInput, ntAgeColInput, ntLivingColInput, ntCmsColInput, districtColInput, categoryColInput, subsidyColInput, headerRowIdxInput].forEach(el => {
        el.addEventListener('input', () => {
            checkReady();
            if (currentFile) triggerProcess();
        });
        el.addEventListener('change', () => {
            checkReady();
            if (currentFile) triggerProcess();
        });
    });

    // --- 同步至 GAS (Google Apps Script) 邏輯 ---
    uploadBtn.onclick = async () => {
        if (currentMode === 'service' && !currentCity) {
            showStatus('請選擇縣市', 'error');
            return;
        }
        if (rocDateInput.value.length < 2) {
            showStatus('請輸入月份', 'error');
            return;
        }

        if (!GAS_URL) {
            showStatus('錯誤: 未設定 GAS URL 環境變數', 'error');
            return;
        }

        setLoading(true);
        showStatus('正在同步至 Google Sheet...', 'success');

        try {
            if (currentMode === 'case') {
                const payloads = JSON.parse(uploadBtn.dataset.outputPayload || '[]');
                for (const payload of payloads) {
                    await fetch(GAS_URL, {
                        method: 'POST',
                        mode: 'no-cors',
                        cache: 'no-cache',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify(payload)
                    });
                }
            } else {
                const outputRows = JSON.parse(uploadBtn.dataset.outputRows || '[]');
                let sheetNamePrefix = currentCity;
                if (currentCity === '台北') {
                    const selectedCategory = document.querySelector('input[name="tpServiceCategory"]:checked').value;
                    sheetNamePrefix = `台北${selectedCategory}`;
                }
                const sheetName = sheetNamePrefix + rocDateInput.value;
                await fetch(GAS_URL, {
                    method: 'POST',
                    mode: 'no-cors',
                    cache: 'no-cache',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ sheetName, rows: outputRows })
                });
            }
            resetUpload();
            if (currentMode === 'case') {
                rocDateInput.value = '';
            }
            showStatus('同步完成！請檢查您的 Google Sheet', 'success');
        } catch (error) {
            console.error(error);
            showStatus('同步失敗: ' + error.message, 'error');
        } finally {
            setLoading(false);
        }
    };

    /**
     * 設置按鈕的載入狀態
     * @param {boolean} isLoading - 是否為載入中
     */
    function setLoading(isLoading) {
        uploadBtn.disabled = isLoading;
        uploadBtn.querySelector('.btn-text').classList.toggle('hidden', isLoading);
        uploadBtn.querySelector('.loader').classList.toggle('hidden', !isLoading);
    }

    /**
     * 顯示狀態訊息
     * @param {string} msg - 訊息內容
     * @param {string} type - 訊息類型 ('success' 或 'error')
     */
    function showStatus(msg, type) {
        statusMessage.textContent = msg;
        statusMessage.className = 'status-message ' + type;
    }
});
