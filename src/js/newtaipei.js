import { colLetterToIndex } from './logic.js';

/**
 * 處理新北個案統計
 * @param {Array} rows - 工作表資料列陣列
 * @param {Object} config - 欄位配置資訊 { gender, age, living, cms }
 * @returns {Object} 包含性別、年齡、居住狀況及 CMS 等級統計結果的物件
 */
export function processNewTaipeiCaseStats(rows, config) {
    const stats = {
        gender: { '男': 0, '女': 0 },
        age: { '<= 49': 0, '50 >= && <= 64': 0, '65 >= && <= 74': 0, '75 >= && <= 84': 0, '85 >=': 0 },
        living: {},
        cms: {}
    };

    if (!rows || rows.length === 0) return stats;

    const genderColIdx = colLetterToIndex(config.gender || '');
    const ageColIdx = colLetterToIndex(config.age || '');
    const livingColIdx = colLetterToIndex(config.living || '');
    const cmsColIdx = colLetterToIndex(config.cms || '');

    for (let i = 1; i < rows.length; i++) {
        const row = rows[i];

        // 性別統計
        if (genderColIdx !== -1 && row[genderColIdx] !== undefined) {
            const gender = row[genderColIdx].toString().trim();
            if (gender === '男') stats.gender['男']++;
            else if (gender === '女') stats.gender['女']++;
        }

        // 年齡統計
        if (ageColIdx !== -1 && row[ageColIdx] !== undefined) {
            const age = parseInt(row[ageColIdx]);
            if (!isNaN(age)) {
                if (age <= 49) stats.age['<= 49']++;
                else if (age >= 50 && age <= 64) stats.age['50 >= && <= 64']++;
                else if (age >= 65 && age <= 74) stats.age['65 >= && <= 74']++;
                else if (age >= 75 && age <= 84) stats.age['75 >= && <= 84']++;
                else if (age >= 85) stats.age['85 >=']++;
            }
        }

        // 居住狀況統計
        if (livingColIdx !== -1 && row[livingColIdx] !== undefined) {
            let living = row[livingColIdx].toString().trim();
            if (living) {
                if (living === '其他類型居住狀況') living = '其他';
                stats.living[living] = (stats.living[living] || 0) + 1;
            }
        }

        // CMS 等級統計
        if (cmsColIdx !== -1 && row[cmsColIdx] !== undefined) {
            const cms = row[cmsColIdx].toString().trim();
            if (cms) stats.cms[cms] = (stats.cms[cms] || 0) + 1;
        }
    }

    return stats;
}

/**
 * 格式化新北輸出內容 (用於寫入 Google Sheets 的格式)
 * @param {Object} stats - 統計資料物件
 * @returns {Array} 二維陣列，代表要寫入 Excel 的各個儲存格內容
 */
export function formatNewTaipeiCaseOutput(stats) {
    const livingOrder = ['獨居', '獨居(兩老)', '與家人或其他人同住', '與朋友同住', '其他'];
    const cmsOrder = ['1', '2', '3', '4', '5', '6', '7', '8'];
    const otherCms = Object.keys(stats.cms).filter(level => !cmsOrder.includes(level)).sort();
    const combinedCmsOrder = cmsOrder.concat(otherCms);

    const rows = [['【性別統計】', ''], ['男', stats.gender['男'] || 0], ['女', stats.gender['女'] || 0], ['', '']];
    rows.push(['【年齡分佈】', ''], ['<= 49', stats.age['<= 49'] || 0], ['50 >= && <= 64', stats.age['50 >= && <= 64'] || 0], ['65 >= && <= 74', stats.age['65 >= && <= 74'] || 0], ['75 >= && <= 84', stats.age['75 >= && <= 84'] || 0], ['85 >=', stats.age['85 >='] || 0], ['', '']);
    rows.push(['【居住狀況】', '']);
    livingOrder.forEach(label => { if (stats.living[label] !== undefined) rows.push([label, stats.living[label]]); });
    rows.push(['', ''], ['【CMS 等級】', '']);
    combinedCmsOrder.forEach(label => { if (stats.cms[label] !== undefined) rows.push([label, stats.cms[label]]); });

    return rows;
}

/**
 * 渲染新北預覽 HTML 元素
 * @param {Object} stats - 統計資料物件
 * @param {string} sheetName - 工作表名稱
 * @returns {HTMLElement} 包含統計摘要預覽的 HTML div 元素
 */
export function renderNewTaipeiCasePreview(stats, sheetName) {
    const livingOrder = ['獨居', '獨居(兩老)', '與家人或其他人同住', '與朋友同住', '其他'];
    const cmsOrder = ['1', '2', '3', '4', '5', '6', '7', '8'];
    const otherCms = Object.keys(stats.cms).filter(level => !cmsOrder.includes(level)).sort();
    const combinedCmsOrder = cmsOrder.concat(otherCms);

    const groupHtml = document.createElement('div');
    groupHtml.className = 'preview-group';
    groupHtml.innerHTML = `
        <div class="preview-group-header">
            <span class="preview-group-title">新北</span>
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
                <div class="stat-values">
                    ${livingOrder.filter(label => stats.living[label] !== undefined).map(label => `<div class="stat-item"><span>${label}</span> <span class="stat-value">${stats.living[label]}</span></div>`).join('')}
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-label">CMS 等級</div>
                <div class="stat-values">
                    ${combinedCmsOrder.filter(label => stats.cms[label] !== undefined).map(label => `<div class="stat-item"><span>${label}</span> <span class="stat-value">${stats.cms[label]}</span></div>`).join('')}
                </div>
            </div>
        </div>
    `;
    return groupHtml;
}
