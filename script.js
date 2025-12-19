document.addEventListener('DOMContentLoaded', () => {
    const dropZone = document.getElementById('dropZone');
    const fileInput = document.getElementById('fileInput');
    const citySelect = document.getElementById('citySelect');
    const rocDateInput = document.getElementById('rocDate');
    const genderColInput = document.getElementById('genderCol');
    const ageColInput = document.getElementById('ageCol');
    const livingColInput = document.getElementById('livingCol');
    const cmsColInput = document.getElementById('cmsCol');
    const previewSection = document.getElementById('previewSection');
    const uploadBtn = document.getElementById('uploadBtn');
    const statusMessage = document.getElementById('statusMessage');
    const fileInfo = document.getElementById('fileInfo');
    const removeFileBtn = document.getElementById('removeFile');
    const genderStatsDiv = document.getElementById('genderStats');
    const ageStatsDiv = document.getElementById('ageStats');
    const livingStatsDiv = document.getElementById('livingStats');
    const cmsStatsDiv = document.getElementById('cmsStats');

    // Get GAS URL from environment variable (Vite prefix)
    const caseListCols = document.getElementById('caseListCols');
    const serviceRegisterCols = document.getElementById('serviceRegisterCols');
    const headerRowIdxInput = document.getElementById('headerRowIdx');
    const districtColInput = document.getElementById('districtCol');
    const categoryColInput = document.getElementById('categoryCol');
    const subsidyColInput = document.getElementById('subsidyCol');
    const modeButtons = document.querySelectorAll('.btn-switch');
    const caseStatsGrid = document.getElementById('caseStatsGrid');
    const serviceStatsGrid = document.getElementById('serviceStatsGrid');
    const districtStatsDiv = document.getElementById('districtStats');
    const categoryStatsDiv = document.getElementById('categoryStats');
    const subsidyStatsDiv = document.getElementById('subsidyStats');

    // Get GAS URL from environment variable (Vite prefix)
    const GAS_URL = import.meta.env.VITE_GAS_URL;

    let currentFile = null;
    let currentMode = 'case'; // 'case' or 'service'

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

            // Reset File and Preview
            resetUpload();
        };
    });

    // --- Utility: Column Letter to Index ---
    function colLetterToIndex(letter) {
        if (!letter) return -1;
        let index = 0;
        const s = letter.toUpperCase();
        for (let i = 0; i < s.length; i++) {
            index = index * 26 + s.charCodeAt(i) - 64;
        }
        return index - 1;
    }

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
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                
                // Read as Array of Arrays to support manual column indexing
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                
                processAndPreview(jsonData, currentFile.name);
            } catch (err) {
                console.error(err);
                showStatus('讀取檔案失敗', 'error');
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
    function processAndPreview(rows, fileName) {
        if (currentMode === 'case') {
            processCaseList(rows, fileName);
        } else {
            processServiceRegister(rows, fileName);
        }
    }

    function processCaseList(rows, fileName) {
        const genderColIdx = colLetterToIndex(genderColInput.value || 'G');
        const ageColIdx = colLetterToIndex(ageColInput.value || 'F');
        const livingColIdx = colLetterToIndex(livingColInput.value || 'O');
        const cmsColIdx = colLetterToIndex(cmsColInput.value || 'J');

        const stats = {
            '男': 0, '女': 0,
            '<= 49': 0, '50 >= && <= 64': 0, '65 >= && <= 74': 0, '75 >= && <= 84': 0, '85 >=': 0
        };
        const livingStats = {};
        const cmsStats = {};

        for (let i = 1; i < rows.length; i++) {
            const row = rows[i];
            if (genderColIdx !== -1 && row[genderColIdx] !== undefined) {
                const gender = row[genderColIdx].toString().trim();
                if (gender === '男') stats['男']++;
                else if (gender === '女') stats['女']++;
            }
            if (ageColIdx !== -1 && row[ageColIdx] !== undefined) {
                const age = parseInt(row[ageColIdx]);
                if (!isNaN(age)) {
                    if (age <= 49) stats['<= 49']++;
                    else if (age >= 50 && age <= 64) stats['50 >= && <= 64']++;
                    else if (age >= 65 && age <= 74) stats['65 >= && <= 74']++;
                    else if (age >= 75 && age <= 84) stats['75 >= && <= 84']++;
                    else if (age >= 85) stats['85 >=']++;
                }
            }
            if (livingColIdx !== -1 && row[livingColIdx] !== undefined) {
                let living = row[livingColIdx].toString().trim();
                if (living) {
                    if (living === '其他類型居住狀況') living = '其他';
                    livingStats[living] = (livingStats[living] || 0) + 1;
                }
            }
            if (cmsColIdx !== -1 && row[cmsColIdx] !== undefined) {
                const cms = row[cmsColIdx].toString().trim();
                if (cms) cmsStats[cms] = (cmsStats[cms] || 0) + 1;
            }
        }

        renderCaseStats(stats, livingStats, cmsStats);
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

    function renderCaseStats(stats, livingStats, cmsStats) {
        caseStatsGrid.classList.remove('hidden');
        serviceStatsGrid.classList.add('hidden');

        genderStatsDiv.innerHTML = `
            <div class="stat-item"><span>男</span> <span class="stat-value">${stats['男']}</span></div>
            <div class="stat-item"><span>女</span> <span class="stat-value">${stats['女']}</span></div>
        `;
        
        ageStatsDiv.innerHTML = `
            <div class="stat-item"><span><= 49</span> <span class="stat-value">${stats['<= 49']}</span></div>
            <div class="stat-item"><span>50 - 64</span> <span class="stat-value">${stats['50 >= && <= 64']}</span></div>
            <div class="stat-item"><span>65 - 74</span> <span class="stat-value">${stats['65 >= && <= 74']}</span></div>
            <div class="stat-item"><span>75 - 84</span> <span class="stat-value">${stats['75 >= && <= 84']}</span></div>
            <div class="stat-item"><span>>= 85</span> <span class="stat-value">${stats['85 >=']}</span></div>
        `;

        const livingOrder = ['獨居', '獨居(兩老)', '與家人或其他人同住', '與朋友同住', '其他'];
        livingStatsDiv.innerHTML = livingOrder
            .filter(label => livingStats[label] !== undefined)
            .map(label => `<div class="stat-item"><span>${label}</span> <span class="stat-value">${livingStats[label]}</span></div>`)
            .join('') || '<div class="stat-item"><span>無資料</span></div>';

        const cmsOrder = ['1', '2', '3', '4', '5', '6', '7', '8'];
        const otherCms = Object.keys(cmsStats).filter(level => !cmsOrder.includes(level)).sort();
        const combinedCmsOrder = cmsOrder.concat(otherCms);
        cmsStatsDiv.innerHTML = combinedCmsOrder
            .filter(label => cmsStats[label] !== undefined)
            .map(label => `<div class="stat-item"><span>${label}</span> <span class="stat-value">${cmsStats[label]}</span></div>`)
            .join('') || '<div class="stat-item"><span>無資料</span></div>';

        const outputRows = buildCaseOutputRows(stats, livingStats, cmsStats, livingOrder, combinedCmsOrder);
        uploadBtn.dataset.outputRows = JSON.stringify(outputRows);
    }

    function renderServiceStats(districtStats, categoryStats, subsidyStats) {
        caseStatsGrid.classList.add('hidden');
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

    function buildCaseOutputRows(stats, livingStats, cmsStats, livingOrder, cmsOrder) {
        const rows = [['【性別統計】', ''], ['男', stats['男'] || 0], ['女', stats['女'] || 0], ['', '']];
        rows.push(['【年齡分佈】', ''], ['<= 49', stats['<= 49'] || 0], ['50 >= && <= 64', stats['50 >= && <= 64'] || 0], ['65 >= && <= 74', stats['65 >= && <= 74'] || 0], ['75 >= && <= 84', stats['75 >= && <= 84'] || 0], ['85 >=', stats['85 >='] || 0], ['', '']);
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
        const hasCity = citySelect.value !== '';
        const hasDate = rocDateInput.value.length >= 2;
        const hasFile = !!currentFile;

        let hasCols = false;
        if (currentMode === 'case') {
            hasCols = genderColInput.value.trim() && ageColInput.value.trim() && livingColInput.value.trim() && cmsColInput.value.trim();
        } else {
            hasCols = districtColInput.value.trim() && categoryColInput.value.trim() && subsidyColInput.value.trim();
        }

        uploadBtn.disabled = !(hasCity && hasDate && hasCols && hasFile);
    }

    [citySelect, rocDateInput, genderColInput, ageColInput, livingColInput, cmsColInput, districtColInput, categoryColInput, subsidyColInput, headerRowIdxInput].forEach(el => {
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
        if (!citySelect.value) {
            showStatus('請選擇縣市', 'error');
            return;
        }
        if (rocDateInput.value.length < 2) {
            showStatus('請輸入月份', 'error');
            return;
        }

        // 取得已建構好的輸出資料
        const outputRows = JSON.parse(uploadBtn.dataset.outputRows || '[]');
        // 組合 sheetName 為「縣市月份」格式
        const sheetName = citySelect.value + rocDateInput.value;

        if (!GAS_URL) {
            showStatus('錯誤: 未設定 GAS URL 環境變數', 'error');
            return;
        }

        setLoading(true);
        showStatus('正在同步至 Google Sheet...', 'success');

        try {
            // 直接傳送已格式化的二維陣列，GAS 只需寫入即可
            await fetch(GAS_URL, {
                method: 'POST',
                mode: 'no-cors',
                cache: 'no-cache',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ sheetName, rows: outputRows })
            });

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
