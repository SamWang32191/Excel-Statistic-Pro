import { identifyWorksheets, processCityStats, colLetterToIndex } from './logic.js';

document.addEventListener('DOMContentLoaded', () => {
    const dropZone = document.getElementById('dropZone');
    const fileInput = document.getElementById('fileInput');
    const citySelect = document.getElementById('citySelect');
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

    // Helper: Toggle Taipei Service Category visibility
    const updateTpServiceCategoryVisibility = () => {
        const show = currentMode === 'service' && citySelect.value === '台北';
        tpServiceCategoryGroup.classList.toggle('hidden', !show);
    };

    // Get GAS URL from environment variable (Vite prefix)
    const GAS_URL = import.meta.env.VITE_GAS_URL;

    let currentFile = null;
    let currentMode = 'case'; // 'case' or 'service'

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

    citySelect.addEventListener('change', () => {
        updateTpServiceCategoryVisibility();
        checkReady();
        if (currentFile) triggerProcess();
    });

    // --- Utility: Column Letter to Index ---
    // colLetterToIndex removed, using imported version implicitly inside processCityStats logic (but script.js still uses it for Service Mode?)
    // Wait, Service Mode still uses colLetterToIndex directly!
    // I need to import it explicitly or keep a local copy if I don't import it.
    // I should import it.


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
                    processCaseList(result, currentFile.name);
                } else {
                    // Service List Mode: Keep original behavior (single sheet, or maybe default to first sheet?)
                    // Spec says: "選擇台北時，新增子類別...". It implies service list might also be city-specific.
                    // But for now, let's assume Service List still works on the FIRST sheet or user selects?
                    // The spec mainly focuses on Case List for multi-sheet. 
                    // Let's stick to the first sheet for Service List for now unless specified otherwise.
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
    // processAndPreview removed as logic moved to triggerProcess

    function processCaseList(data, fileName) {
        // data structure: { taipei: [...], newTaipei: [...] }
        const groups = [];

        // Helper to create empty stats
        const createEmptyStats = () => ({
            gender: { '男': 0, '女': 0 },
            age: { '<= 49': 0, '50 >= && <= 64': 0, '65 >= && <= 74': 0, '75 >= && <= 84': 0, '85 >=': 0 },
            living: {},
            cms: {}
        });

        // 1. Process New Taipei
        if (data.newTaipei && data.newTaipei.length > 0) {
            const stats = processCityStats(data.newTaipei, {
                gender: ntGenderColInput.value,
                age: ntAgeColInput.value,
                living: ntLivingColInput.value,
                cms: ntCmsColInput.value
            });
            groups.push({
                title: '新北',
                sheetName: `新北${rocDateInput.value}`,
                stats,
                hasData: true
            });
        } else {
            groups.push({
                title: '新北',
                sheetName: `新北${rocDateInput.value}`,
                stats: createEmptyStats(),
                hasData: false
            });
        }

        // 2. Process Taipei
        if (data.taipei && data.taipei.length > 0) {
            // Taipei - Elderly (老福)
            const statsElderly = processCityStats(data.taipei, {
                category: tpCategoryColInput.value,
                gender: tpGenderColInput.value,
                age: tpAgeColInput.value,
                living: tpLivingColInput.value,
                cms: tpCmsColInput.value
            }, true, '老福');

            groups.push({
                title: '台北老福',
                sheetName: `台北老福${rocDateInput.value}`,
                stats: statsElderly,
                hasData: statsElderly.gender['男'] + statsElderly.gender['女'] > 0
            });

            // Taipei - Disabled (身障)
            const statsDisabled = processCityStats(data.taipei, {
                category: tpCategoryColInput.value,
                gender: tpGenderColInput.value,
                age: tpAgeColInput.value,
                living: tpLivingColInput.value,
                cms: tpCmsColInput.value
            }, true, '身障');

            groups.push({
                title: '台北身障',
                sheetName: `台北身障${rocDateInput.value}`,
                stats: statsDisabled,
                hasData: statsDisabled.gender['男'] + statsDisabled.gender['女'] > 0
            });
        } else {
            // No Taipei Data found
            groups.push({
                title: '台北老福',
                sheetName: `台北老福${rocDateInput.value}`,
                stats: createEmptyStats(),
                hasData: false
            });
            groups.push({
                title: '台北身障',
                sheetName: `台北身障${rocDateInput.value}`,
                stats: createEmptyStats(),
                hasData: false
            });
        }

        renderAllStats(groups);
        finalizePreview(fileName);
    }

    function processServiceRegister(rows, fileName) {
        const headerRowIdx = parseInt(headerRowIdxInput.value || '5');
        const districtColIdx = colLetterToIndex(districtColInput.value || 'U');
        const categoryColIdx = colLetterToIndex(categoryColInput.value || 'G');
        const subsidyColIdx = colLetterToIndex(subsidyColInput.value || 'O');

        const districtStats = {};
        const categoryStats = {};
        const subsidyStats = {};

        // 從標頭行的下一行開始 (Excel 1-based, rows is 0-based)
        // 例如標頭是第 5 行，其索引是 4，資料從第 6 行開始 (索引 5)
        for (let i = headerRowIdx; i < rows.length; i++) {
            const row = rows[i];
            if (districtColIdx !== -1 && row[districtColIdx] !== undefined) {
                const val = row[districtColIdx].toString().trim();
                if (val) districtStats[val] = (districtStats[val] || 0) + 1;
            }
            if (categoryColIdx !== -1 && row[categoryColIdx] !== undefined) {
                const val = row[categoryColIdx].toString().trim();
                if (val) categoryStats[val] = (categoryStats[val] || 0) + 1;
            }
            if (subsidyColIdx !== -1 && row[subsidyColIdx] !== undefined) {
                const val = row[subsidyColIdx].toString().trim();
                if (val) subsidyStats[val] = (subsidyStats[val] || 0) + 1;
            }
        }

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

    function renderAllStats(groups) {
        previewContainer.innerHTML = ''; // Clear previous content
        if (serviceStatsGrid) serviceStatsGrid.classList.add('hidden');
        previewContainer.classList.remove('hidden');

        // Note: Even if groups exist, they might all be empty. That's fine, we show them as empty.
        
        const allPayloads = [];

        groups.forEach(group => {
            const { title, sheetName, stats, hasData } = group;
            
            // Build HTML for this group
            const groupHtml = document.createElement('div');
            groupHtml.className = 'preview-group';
            
            if (!hasData) {
                groupHtml.innerHTML = `
                    <div class="preview-group-header">
                        <span class="preview-group-title">${title}</span>
                        <span class="preview-group-badge" style="background: rgba(100,116,139, 0.2); color: #94a3b8;">無資料</span>
                    </div>
                    <div class="stats-grid" style="opacity: 0.5;">
                        <div class="stat-card"><div class="stat-label">狀態</div><div class="stat-values">無資料</div></div>
                    </div>
                `;
            } else {
                const livingOrder = ['獨居', '獨居(兩老)', '與家人或其他人同住', '與朋友同住', '其他'];
                const livingHtml = livingOrder
                    .filter(label => stats.living[label] !== undefined)
                    .map(label => `<div class="stat-item"><span>${label}</span> <span class="stat-value">${stats.living[label]}</span></div>`)
                    .join('') || '<div class="stat-item"><span>無資料</span></div>';

                const cmsOrder = ['1', '2', '3', '4', '5', '6', '7', '8'];
                const otherCms = Object.keys(stats.cms).filter(level => !cmsOrder.includes(level)).sort();
                const combinedCmsOrder = cmsOrder.concat(otherCms);
                const cmsHtml = combinedCmsOrder
                    .filter(label => stats.cms[label] !== undefined)
                    .map(label => `<div class="stat-item"><span>${label}</span> <span class="stat-value">${stats.cms[label]}</span></div>`)
                    .join('') || '<div class="stat-item"><span>無資料</span></div>';

                groupHtml.innerHTML = `
                    <div class="preview-group-header">
                        <span class="preview-group-title">${title}</span>
                        <span class="preview-group-badge">${sheetName}</span>
                    </div>
                    <div class="stats-grid">
                        <div class="stat-card">
                            <div class="stat-label">性別統計</div>
                            <div class="stat-values">
                                <div class="stat-item"><span>男</span> <span class="stat-value">${stats.gender['男']}</span></div>
                                <div class="stat-item"><span>女</span> <span class="stat-value">${stats.gender['女']}</span></div>
                            </div>
                        </div>
                        <div class="stat-card">
                            <div class="stat-label">年齡分佈</div>
                            <div class="stat-values">
                                <div class="stat-item"><span><= 49</span> <span class="stat-value">${stats.age['<= 49']}</span></div>
                                <div class="stat-item"><span>50 - 64</span> <span class="stat-value">${stats.age['50 >= && <= 64']}</span></div>
                                <div class="stat-item"><span>65 - 74</span> <span class="stat-value">${stats.age['65 >= && <= 74']}</span></div>
                                <div class="stat-item"><span>75 - 84</span> <span class="stat-value">${stats.age['75 >= && <= 84']}</span></div>
                                <div class="stat-item"><span>>= 85</span> <span class="stat-value">${stats.age['85 >=']}</span></div>
                            </div>
                        </div>
                        <div class="stat-card">
                            <div class="stat-label">居住狀況</div>
                            <div class="stat-values">${livingHtml}</div>
                        </div>
                        <div class="stat-card">
                            <div class="stat-label">CMS 等級</div>
                            <div class="stat-values">${cmsHtml}</div>
                        </div>
                    </div>
                `;
                
                // Prepare Payload for GAS ONLY if data exists
                const rows = buildCaseOutputRows(stats.gender, stats.living, stats.cms, livingOrder, combinedCmsOrder, stats.age);
                allPayloads.push({ sheetName, rows });
            }
            
            previewContainer.appendChild(groupHtml);
        });

        // Store payload (only contains valid data)
        uploadBtn.dataset.outputPayload = JSON.stringify(allPayloads);
        
        // Enable upload even if some parts are missing, as long as *something* is there?
        // Or if user wants to upload partial data.
        // Currently validation is done in 'checkReady'.
        // We should warn if ALL are no-data.
        if (allPayloads.length === 0) {
            showStatus('警告：所有縣市皆無資料', 'error');
            uploadBtn.disabled = true;
        }
    }

    function renderServiceStats(districtStats, categoryStats, subsidyStats) {
        previewContainer.classList.add('hidden');
        serviceStatsGrid.classList.remove('hidden');

        // Sorting Administrative Districts
        const orderNT = ['三重區', '蘆洲區', '五股區', '板橋區', '土城區', '新莊區', '中和區', '永和區', '樹林區', '泰山區', '新店區'];
        const orderTP = ['松山區', '信義區', '大安區', '中山區', '中正區', '大同區', '萬華區', '文山區', '南港區', '內湖區', '士林區', '北投區'];
        const baseOrder = citySelect.value === '台北' ? orderTP : orderNT;
        
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

    function buildCaseOutputRows(stats, livingStats, cmsStats, livingOrder, cmsOrder, ageStats) {
        // Use ageStats if provided, otherwise fallback to stats
        const finalAgeStats = ageStats || stats;
        
        const rows = [['【性別統計】', ''], ['男', stats['男'] || 0], ['女', stats['女'] || 0], ['', '']];
        rows.push(['【年齡分佈】', ''], ['<= 49', finalAgeStats['<= 49'] || 0], ['50 >= && <= 64', finalAgeStats['50 >= && <= 64'] || 0], ['65 >= && <= 74', finalAgeStats['65 >= && <= 74'] || 0], ['75 >= && <= 84', finalAgeStats['75 >= && <= 84'] || 0], ['85 >=', finalAgeStats['85 >='] || 0], ['', '']);
        rows.push(['【居住狀況】', '']);
        livingOrder.forEach(label => { if (livingStats[label] !== undefined) rows.push([label, livingStats[label]]); });
        rows.push(['', ''], ['【CMS 等級】', '']);
        cmsOrder.forEach(label => { if (cmsStats[label] !== undefined) rows.push([label, cmsStats[label]]); });
        return rows;
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
        const hasCity = currentMode === 'case' || citySelect.value !== '';
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

    [citySelect, rocDateInput, tpCategoryColInput, tpGenderColInput, tpAgeColInput, tpLivingColInput, tpCmsColInput, ntGenderColInput, ntAgeColInput, ntLivingColInput, ntCmsColInput, districtColInput, categoryColInput, subsidyColInput, headerRowIdxInput].forEach(el => {
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
        // 防禦性檢核：再次確認必要欄位
        if (currentMode === 'service' && !citySelect.value) {
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
                // Multi-sheet Upload
                const payloads = JSON.parse(uploadBtn.dataset.outputPayload || '[]');
                
                // Upload sequentially to avoid race conditions or hitting rate limits too hard
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
                // Single sheet Upload (Service Mode)
                const outputRows = JSON.parse(uploadBtn.dataset.outputRows || '[]');
                let sheetNamePrefix = citySelect.value;
                
                if (citySelect.value === '台北') {
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
