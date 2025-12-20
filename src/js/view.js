import { identifyWorksheets, colLetterToIndex, processServiceStats } from './logic.js';
import { processTaipeiCaseStats, formatTaipeiCaseOutput, renderTaipeiCasePreview } from './taipei.js';
import { processNewTaipeiCaseStats, formatNewTaipeiCaseOutput, renderNewTaipeiCasePreview } from './newtaipei.js';

document.addEventListener('DOMContentLoaded', () => {
    const dropZone = document.getElementById('dropZone');
    const fileInput = document.getElementById('fileInput');
    const cityButtons = document.querySelectorAll('.btn-city');
    const rocDateInput = document.getElementById('rocDate');

    // Taipei Inputs
    const tpCategoryColInput = document.getElementById('tpCategoryCol');
    const tpGenderColInput = document.getElementById('tpGenderCol');
    const tpAgeColInput = document.getElementById('tpAgeCol');
    const tpLivingColInput = document.getElementById('tpLivingCol');
    const tpCmsColInput = document.getElementById('tpCmsCol');

    // New Taipei Inputs
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

    // Service Stats Containers
    const serviceStatsGrid = document.getElementById('serviceStatsGrid');
    const districtStatsDiv = document.getElementById('districtStats');
    const categoryStatsDiv = document.getElementById('categoryStats');
    const subsidyStatsDiv = document.getElementById('subsidyStats');

    const modeButtons = document.querySelectorAll('.btn-switch');

    const updateTpServiceCategoryVisibility = () => {
        const show = currentMode === 'service' && currentCity === '台北';
        tpServiceCategoryGroup.classList.toggle('hidden', !show);
    };

    // Get GAS URL from environment variable (Vite prefix)
    const GAS_URL = import.meta.env.VITE_GAS_URL;

    let currentFile = null;
    let currentMode = 'case'; // 'case' or 'service'
    let currentCity = '台北'; // default city

    // Initial state setup
    cityInputGroup.classList.toggle('hidden', currentMode === 'case');
    updateTpServiceCategoryVisibility();

    // --- Mode Switching ---
    modeButtons.forEach(btn => {
        btn.onclick = () => {
            if (btn.dataset.mode === currentMode) return;

            // Update UI State
            modeButtons.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            currentMode = btn.dataset.mode;

            // Toggle Input Visibility
            caseListCols.classList.toggle('hidden', currentMode !== 'case');
            serviceRegisterCols.classList.toggle('hidden', currentMode !== 'service');
            cityInputGroup.classList.toggle('hidden', currentMode === 'case');
            updateTpServiceCategoryVisibility();

            // Reset File and Preview
            resetUpload();
        };
    });

    // --- City Button Switching (Service Mode) ---
    cityButtons.forEach(btn => {
        btn.onclick = () => {
            if (btn.dataset.city === currentCity) return;

            cityButtons.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            currentCity = btn.dataset.city;

            updateTpServiceCategoryVisibility();
            resetUpload(); // Clear file and preview on city change
            checkReady();
        };
    });

    // --- Service Category Change (Taipei) ---
    const tpServiceCategoryRadios = document.querySelectorAll('input[name="tpServiceCategory"]');
    tpServiceCategoryRadios.forEach(radio => {
        radio.addEventListener('change', () => {
            resetUpload(); // Clear file and preview on category change
        });
    });

    // --- Upload Logic ---
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

    function handleFile(file) {
        if (!file.name.match(/\.(xlsx|xls)$/)) {
            showStatus('不支援的檔案格式', 'error');
            return;
        }
        currentFile = file;
        triggerProcess();
    }

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

    function resetUpload() {
        fileInput.value = '';
        currentFile = null;
        fileInfo.classList.add('hidden');
        document.querySelector('.upload-content').classList.remove('hidden');
        previewSection.classList.add('hidden');
        uploadBtn.disabled = true;
        showStatus('請選擇縣市、輸入月份並上傳檔案', '');
    }

    // --- Statistics Logic ---
    function processCaseList(data, fileName, taipeiSheetName, newTaipeiSheetName) {
        previewContainer.innerHTML = '';
        if (serviceStatsGrid) serviceStatsGrid.classList.add('hidden');
        previewContainer.classList.remove('hidden');

        const allPayloads = [];
        const month = rocDateInput.value;

        // 1. Process New Taipei
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

        // 2. Process Taipei
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

    function finalizePreview(fileName) {
        fileInfo.classList.remove('hidden');
        fileInfo.querySelector('.file-name').textContent = fileName;
        document.querySelector('.upload-content').classList.add('hidden');
        previewSection.classList.remove('hidden');
        showStatus('檔案已就緒，請確認統計結果', 'success');
        checkReady();
    }

    function renderServiceStats(districtStats, categoryStats, subsidyStats) {
        previewContainer.classList.add('hidden');
        serviceStatsGrid.classList.remove('hidden');

        // Sorting Administrative Districts
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

        // Sorting Categories
        const sortedCategories = Object.keys(categoryStats).sort((a, b) => a.localeCompare(b, 'zh-hant'));
        categoryStatsDiv.innerHTML = sortedCategories
            .map(label => `<div class="stat-item"><span>${label}</span> <span class="stat-value">${categoryStats[label]}</span></div>`)
            .join('') || '<div class="stat-item"><span>無資料</span></div>';

        // Sorting Subsidies
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

    function buildServiceOutputRows(districtStats, sortedDistricts, categoryStats, sortedCategories, subsidyStats, sortedSubsidies) {
        const rows = [['【行政區統計】', '']];
        sortedDistricts.forEach(label => rows.push([label, districtStats[label]]));
        rows.push(['', ''], ['【服務項目類別】', '']);
        sortedCategories.forEach(label => rows.push([label, categoryStats[label]]));
        rows.push(['', ''], ['【補助比率】', '']);
        sortedSubsidies.forEach(label => rows.push([label, subsidyStats[label]]));
        return rows;
    }

    // --- Validation Logic ---
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

    // --- Upload to GAS ---
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

    function setLoading(isLoading) {
        uploadBtn.disabled = isLoading;
        uploadBtn.querySelector('.btn-text').classList.toggle('hidden', isLoading);
        uploadBtn.querySelector('.loader').classList.toggle('hidden', !isLoading);
    }

    function showStatus(msg, type) {
        statusMessage.textContent = msg;
        statusMessage.className = 'status-message ' + type;
    }
});
