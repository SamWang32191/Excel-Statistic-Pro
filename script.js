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
    const GAS_URL = import.meta.env.VITE_GAS_URL;

    let currentFile = null;

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
        const genderColIdx = colLetterToIndex(genderColInput.value || 'G');
        const ageColIdx = colLetterToIndex(ageColInput.value || 'F');
        const livingColIdx = colLetterToIndex(livingColInput.value || 'O');
        const cmsColIdx = colLetterToIndex(cmsColInput.value || 'J');

        const stats = {
            '男': 0,
            '女': 0,
            '<= 49': 0,
            '50 >= && <= 64': 0,
            '65 >= && <= 74': 0,
            '75 >= && <= 84': 0,
            '85 >=': 0
        };

        // 居住狀況統計（動態收集所有不同的值）
        const livingStats = {};
        // CMS 等級統計
        const cmsStats = {};

        // Skip header row (index 0)
        for (let i = 1; i < rows.length; i++) {
            const row = rows[i];
            
            // Gender
            if (genderColIdx !== -1 && row[genderColIdx] !== undefined) {
                const gender = row[genderColIdx].toString().trim();
                if (gender === '男') stats['男']++;
                else if (gender === '女') stats['女']++;
            }

            // Age
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

            // Living Status (居住狀況)
            if (livingColIdx !== -1 && row[livingColIdx] !== undefined) {
                let living = row[livingColIdx].toString().trim();
                if (living) {
                    // 將「其他類型居住狀況」合併為「其他」
                    if (living === '其他類型居住狀況') {
                        living = '其他';
                    }
                    livingStats[living] = (livingStats[living] || 0) + 1;
                }
            }

            // CMS Level (CMS等級)
            if (cmsColIdx !== -1 && row[cmsColIdx] !== undefined) {
                const cms = row[cmsColIdx].toString().trim();
                if (cms) {
                    cmsStats[cms] = (cmsStats[cms] || 0) + 1;
                }
            }
        }

        renderStats(stats, livingStats, cmsStats);
        
        fileInfo.classList.remove('hidden');
        fileInfo.querySelector('.file-name').textContent = fileName;
        document.querySelector('.upload-content').classList.add('hidden');
        previewSection.classList.remove('hidden');
        
        showStatus('檔案已就緒，請確認統計結果', 'success');
        checkReady();
    }

    function renderStats(stats, livingStats, cmsStats) {
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

        // 渲染居住狀況統計（依指定順序）
        const livingOrder = ['獨居', '獨居(兩老)', '與家人或其他人同住', '與朋友同住', '其他'];
        const livingHtml = livingOrder
            .filter(label => livingStats[label] !== undefined)
            .map(label => `<div class="stat-item"><span>${label}</span> <span class="stat-value">${livingStats[label]}</span></div>`)
            .join('');
        livingStatsDiv.innerHTML = livingHtml || '<div class="stat-item"><span>無資料</span></div>';

        // 渲染 CMS 等級統計（依 1-8 順序）
        const cmsOrder = ['1', '2', '3', '4', '5', '6', '7', '8'];
        // 找出不在預定義順序中的其他等級
        const otherCms = Object.keys(cmsStats).filter(level => cmsOrder.indexOf(level) === -1).sort();
        const combinedCmsOrder = cmsOrder.concat(otherCms);
        
        const cmsHtml = combinedCmsOrder
            .filter(label => cmsStats[label] !== undefined)
            .map(label => `<div class="stat-item"><span>${label} 級</span> <span class="stat-value">${cmsStats[label]}</span></div>`)
            .join('');
        cmsStatsDiv.innerHTML = cmsHtml || '<div class="stat-item"><span>無資料</span></div>';

        // 建構輸出用的二維陣列（Google Sheet 格式）
        const outputRows = buildOutputRows(stats, livingStats, cmsStats, livingOrder, combinedCmsOrder);
        uploadBtn.dataset.outputRows = JSON.stringify(outputRows);
    }

    // 建構 Google Sheet 輸出格式的二維陣列
    function buildOutputRows(stats, livingStats, cmsStats, livingOrder, cmsOrder) {
        const rows = [];
        
        // 性別統計
        rows.push(['【性別統計】', '']);
        rows.push(['男', stats['男'] || 0]);
        rows.push(['女', stats['女'] || 0]);
        rows.push(['', '']);
        
        // 年齡分佈
        rows.push(['【年齡分佈】', '']);
        rows.push(['<= 49', stats['<= 49'] || 0]);
        rows.push(['50 >= && <= 64', stats['50 >= && <= 64'] || 0]);
        rows.push(['65 >= && <= 74', stats['65 >= && <= 74'] || 0]);
        rows.push(['75 >= && <= 84', stats['75 >= && <= 84'] || 0]);
        rows.push(['85 >=', stats['85 >='] || 0]);
        rows.push(['', '']);
        
        // 居住狀況
        rows.push(['【居住狀況】', '']);
        livingOrder.forEach(label => {
            if (livingStats[label] !== undefined) {
                rows.push([label, livingStats[label]]);
            }
        });
        rows.push(['', '']);
        
        // CMS 等級
        rows.push(['【CMS 等級】', '']);
        cmsOrder.forEach(label => {
            if (cmsStats[label] !== undefined) {
                rows.push([label, cmsStats[label]]);
            }
        });
        
        return rows;
    }

    // --- Validation Logic ---
    function checkReady() {
        const hasCity = citySelect.value !== '';
        const hasDate = rocDateInput.value.length >= 2;
        const hasGenderCol = genderColInput.value.trim().length > 0;
        const hasAgeCol = ageColInput.value.trim().length > 0;
        const hasLivingCol = livingColInput.value.trim().length > 0;
        const hasCmsCol = cmsColInput.value.trim().length > 0;
        const hasFile = !!currentFile;
        uploadBtn.disabled = !(hasCity && hasDate && hasGenderCol && hasAgeCol && hasLivingCol && hasCmsCol && hasFile);
    }

    [citySelect, rocDateInput, genderColInput, ageColInput, livingColInput, cmsColInput].forEach(el => {
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
