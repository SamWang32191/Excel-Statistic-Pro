document.addEventListener('DOMContentLoaded', () => {
    const dropZone = document.getElementById('dropZone');
    const fileInput = document.getElementById('fileInput');
    const rocDateInput = document.getElementById('rocDate');
    const genderColInput = document.getElementById('genderCol');
    const ageColInput = document.getElementById('ageCol');
    const previewSection = document.getElementById('previewSection');
    const uploadBtn = document.getElementById('uploadBtn');
    const statusMessage = document.getElementById('statusMessage');
    const fileInfo = document.getElementById('fileInfo');
    const removeFileBtn = document.getElementById('removeFile');
    const genderStatsDiv = document.getElementById('genderStats');
    const ageStatsDiv = document.getElementById('ageStats');

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
        showStatus('請輸入設定並上傳檔案', '');
    }

    // --- Statistics Logic ---
    function processAndPreview(rows, fileName) {
        const genderColIdx = colLetterToIndex(genderColInput.value || 'G');
        const ageColIdx = colLetterToIndex(ageColInput.value || 'F');

        const stats = {
            '男': 0,
            '女': 0,
            '<= 49': 0,
            '50 >= && <= 64': 0,
            '65 >= && <= 74': 0,
            '75 >= && <= 84': 0,
            '85 >=': 0
        };

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
        }

        renderStats(stats);
        
        fileInfo.classList.remove('hidden');
        fileInfo.querySelector('.file-name').textContent = fileName;
        document.querySelector('.upload-content').classList.add('hidden');
        previewSection.classList.remove('hidden');
        
        showStatus('檔案已就緒，請確認統計結果', 'success');
        checkReady();
    }

    function renderStats(stats) {
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

        uploadBtn.dataset.stats = JSON.stringify(stats);
    }

    // --- Validation Logic ---
    function checkReady() {
        const hasDate = rocDateInput.value.length >= 5;
        const hasGenderCol = genderColInput.value.trim().length > 0;
        const hasAgeCol = ageColInput.value.trim().length > 0;
        const hasFile = !!currentFile;
        uploadBtn.disabled = !(hasDate && hasGenderCol && hasAgeCol && hasFile);
    }

    [rocDateInput, genderColInput, ageColInput].forEach(el => {
        el.oninput = () => {
            checkReady();
            if (currentFile) triggerProcess();
        };
    });

    // --- Upload to GAS ---
    uploadBtn.onclick = async () => {
        const stats = JSON.parse(uploadBtn.dataset.stats);
        const sheetName = rocDateInput.value;

        if (!GAS_URL) {
            showStatus('錯誤: 未設定 GAS URL 環境變數', 'error');
            return;
        }

        setLoading(true);
        showStatus('正在同步至 Google Sheet...', 'success');

        try {
            await fetch(GAS_URL, {
                method: 'POST',
                mode: 'no-cors',
                cache: 'no-cache',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ sheetName, data: stats })
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
