import { colLetterToIndex } from './logic.js';

/**
 * 處理台北個案統計
 * @param {Array} rows - 工作表資料列陣列
 * @param {Object} config - 欄位配置資訊
 * @param {string} targetCategory - 目標類別，例如 '老福' 或 '身障'
 * @param {Object} [ageFilter=null] - 年齡篩選條件，例如 { min: 65 } 或 { min: 50, max: 64 }
 * @returns {Object} 包含性別、年齡、居住狀況及 CMS 等級統計結果的物件
 */
export function processTaipeiCaseStats(rows, config, targetCategory, ageFilter = null) {
    const stats = {
        gender: { '男': 0, '女': 0 },
        age: { '<= 49': 0, '50 >= && <= 64': 0, '65 >= && <= 74': 0, '75 >= && <= 84': 0, '85 >=': 0 },
        living: {},
        cms: {}
    };

    if (!rows || rows.length === 0) return stats;

    const categoryColIdx = colLetterToIndex(config.category || '');
    const genderColIdx = colLetterToIndex(config.gender || '');
    const ageColIdx = colLetterToIndex(config.age || '');
    const livingColIdx = colLetterToIndex(config.living || '');
    const cmsColIdx = colLetterToIndex(config.cms || '');

    for (let i = 1; i < rows.length; i++) {
        const row = rows[i];

        // 類別篩選
        const category = (row[categoryColIdx] || '').toString().trim();
        if (category !== targetCategory) continue;

        // 年齡篩選 (新需求)
        const age = parseInt(row[ageColIdx]);
        if (ageFilter) {
            if (isNaN(age)) continue;
            if (ageFilter.min !== undefined && age < ageFilter.min) continue;
            if (ageFilter.max !== undefined && age > ageFilter.max) continue;
        }

        // 性別統計
        if (genderColIdx !== -1 && row[genderColIdx] !== undefined) {
            const gender = row[genderColIdx].toString().trim();
            if (gender === '男') stats.gender['男']++;
            else if (gender === '女') stats.gender['女']++;
        }

        // 年齡統計
        if (!isNaN(age)) {
            if (age <= 49) stats.age['<= 49']++;
            else if (age >= 50 && age <= 64) stats.age['50 >= && <= 64']++;
            else if (age >= 65 && age <= 74) stats.age['65 >= && <= 74']++;
            else if (age >= 75 && age <= 84) stats.age['75 >= && <= 84']++;
            else if (age >= 85) stats.age['85 >=']++;
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
 * 格式化台北輸出內容 (用於寫入 Google Sheets 的格式)
 * @param {Object} elderlyTotal - 整體老人統計
 * @param {Object} elderly65Plus - 65歲以上老人統計
 * @param {Object} elderly50to64 - 50-64歲身障/組員統計
 * @param {Object} disabled - 身心障礙者整體統計
 * @returns {Object} 包含 elderlyRows 與 disabledRows 二維陣列的物件
 */
export function formatTaipeiCaseOutput(elderlyTotal, elderly65Plus, elderly50to64, disabled) {
    const livingOrder = ['獨居', '獨居(兩老)', '與家人或其他人同住', '與朋友住', '其他'];

    // 標準統計格式化工具
    const formatBasic = (stats, title) => {
        const rows = [[`【${title}】`, ''], ['男', stats.gender['男'] || 0], ['女', stats.gender['女'] || 0], ['', '']];
        rows.push(['【年齡分佈】', ''], ['<= 49', stats.age['<= 49'] || 0], ['50 >= && <= 64', stats.age['50 >= && <= 64'] || 0], ['65 >= && <= 74', stats.age['65 >= && <= 74'] || 0], ['75 >= && <= 84', stats.age['75 >= && <= 84'] || 0], ['85 >=', stats.age['85 >='] || 0], ['', '']);
        return rows;
    };

    const formatCms = (stats) => {
        const cmsOrder = ['1', '2', '3', '4', '5', '6', '7', '8'];
        const otherCms = Object.keys(stats.cms).filter(level => !cmsOrder.includes(level)).sort();
        const combinedCmsOrder = cmsOrder.concat(otherCms);
        const rows = [['', ''], ['【CMS 等級】', '']];
        combinedCmsOrder.forEach(label => { if (stats.cms[label] !== undefined) rows.push([label, stats.cms[label]]); });
        return rows;
    };

    // 老福資料特殊格式化 (結合標準統計)
    const formatElderly = () => {
        // 1. 性別與年齡 (標準格式)
        let rows = formatBasic(elderlyTotal, '性別統計');

        // 2. 特殊化居住狀況區塊
        rows.push(['(65歲以上 老人)', '']);
        livingOrder.forEach(label => { rows.push([label === '與家人或其他人同住' ? '與家人住' : label, elderly65Plus.living[label] || 0]); });

        rows.push(['', ''], ['(50~64 身障)', '']);
        const livingOrderShort = ['獨居', '僅配偶2人', '與家人住', '與朋友住', '其他'];
        livingOrderShort.forEach(label => {
            let count = 0;
            if (label === '僅配偶2人') count = elderly50to64.living['獨居(兩老)'] || 0;
            else if (label === '與家人住') count = elderly50to64.living['與家人或其他人同住'] || 0;
            else count = elderly50to64.living[label] || 0;
            rows.push([label, count]);
        });

        rows.push(['', ''], ['身心障礙者', '']);
        const total50to64 = Object.values(elderly50to64.gender).reduce((a, b) => a + b, 0);
        const total65Plus = Object.values(elderly65Plus.gender).reduce((a, b) => a + b, 0);
        rows.push(['50-64歲', total50to64], ['65歲以上', total65Plus]);

        // 3. CMS (標準格式)
        rows = rows.concat(formatCms(elderlyTotal));

        return rows;
    };

    // 身障資料特殊格式化 (標準格式)
    const formatDisabledWhole = () => {
        let rows = formatBasic(disabled, '性別統計');
        rows.push(['【居住狀況】', '']);
        livingOrder.forEach(label => { if (disabled.living[label] !== undefined) rows.push([label, disabled.living[label]]); });
        rows = rows.concat(formatCms(disabled));
        return rows;
    };

    return {
        elderlyRows: formatElderly(),
        disabledRows: formatDisabledWhole()
    };
}

/**
 * 渲染台北預覽 HTML 元素
 * @param {Object} elderlyTotal - 整體老人統計
 * @param {Object} elderly65Plus - 65歲以上老人統計
 * @param {Object} elderly50to64 - 50-64歲老人統計
 * @param {Object} disabled - 身心障礙者統計
 * @param {string} sheetNameElderly - 老福工作表名稱
 * @param {string} sheetNameDisabled - 身障工作表名稱
 * @returns {HTMLElement} 包含台北統計預覽的 HTML div 元素
 */
export function renderTaipeiCasePreview(elderlyTotal, elderly65Plus, elderly50to64, disabled, sheetNameElderly, sheetNameDisabled) {
    const livingOrder = ['獨居', '獨居(兩老)', '與家人或其他人同住', '與朋友住', '其他'];

    const total50to64 = Object.values(elderly50to64.gender).reduce((a, b) => a + b, 0);
    const total65Plus = Object.values(elderly65Plus.gender).reduce((a, b) => a + b, 0);

    const groupHtml = document.createElement('div');
    groupHtml.className = 'preview-group';
    groupHtml.innerHTML = `
        <div class="preview-group-header">
            <span class="preview-group-title">台北老福</span>
            <span class="preview-group-badge">${sheetNameElderly}</span>
        </div>
        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-label">性別統計</div>
                <div class="stat-values">
                    <div class="stat-item"><span>男</span> <span class="stat-value">${elderlyTotal.gender['男']}</span></div>
                    <div class="stat-item"><span>女</span> <span class="stat-value">${elderlyTotal.gender['女']}</span></div>
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-label">年齡分佈</div>
                <div class="stat-values">
                    <div class="stat-item"><span><= 49</span> <span class="stat-value">${elderlyTotal.age['<= 49']}</span></div>
                    <div class="stat-item"><span>50 - 64</span> <span class="stat-value">${elderlyTotal.age['50 >= && <= 64']}</span></div>
                    <div class="stat-item"><span>65 - 74</span> <span class="stat-value">${elderlyTotal.age['65 >= && <= 74']}</span></div>
                    <div class="stat-item"><span>75 - 84</span> <span class="stat-value">${elderlyTotal.age['75 >= && <= 84']}</span></div>
                    <div class="stat-item"><span>>= 85</span> <span class="stat-value">${elderlyTotal.age['85 >=']}</span></div>
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-label">CMS 等級</div>
                <div class="stat-values">
                    ${Object.keys(elderlyTotal.cms).sort().map(l => `<div class="stat-item"><span>${l}</span> <span class="stat-value">${elderlyTotal.cms[l]}</span></div>`).join('') || '<div class="stat-item"><span>無資料</span></div>'}
                </div>
            </div>
        </div>
        <div class="stats-grid triples mt-4">
            <div class="stat-card">
                <div class="stat-label">(65歲以上 老人)</div>
                <div class="stat-values">
                    ${livingOrder.map(l => `<div class="stat-item"><span>${l === '與家人或其他人同住' ? '與家人住' : l}</span> <span class="stat-value">${elderly65Plus.living[l] || 0}</span></div>`).join('')}
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-label">(50~64 身障)</div>
                <div class="stat-values">
                    ${['獨居', '僅配偶2人', '與家人住', '與朋友住', '其他'].map(l => {
        let count = 0;
        if (l === '僅配偶2人') count = elderly50to64.living['獨居(兩老)'] || 0;
        else if (l === '與家人住') count = elderly50to64.living['與家人或其他人同住'] || 0;
        else count = elderly50to64.living[l] || 0;
        return `<div class="stat-item"><span>${l}</span> <span class="stat-value">${count}</span></div>`;
    }).join('')}
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-label">身心障礙者</div>
                <div class="stat-values">
                    <div class="stat-item"><span>50-64歲</span> <span class="stat-value">${total50to64}</span></div>
                    <div class="stat-item"><span>65歲以上</span> <span class="stat-value">${total65Plus}</span></div>
                </div>
            </div>
        </div>

        <div class="preview-group-header mt-8">
            <span class="preview-group-title">台北身障</span>
            <span class="preview-group-badge">${sheetNameDisabled}</span>
        </div>
        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-label">性別統計</div>
                <div class="stat-values">
                    <div class="stat-item"><span>男</span> <span class="stat-value">${disabled.gender['男']}</span></div>
                    <div class="stat-item"><span>女</span> <span class="stat-value">${disabled.gender['女']}</span></div>
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-label">年齡分佈</div>
                <div class="stat-values">
                    <div class="stat-item"><span><= 49</span> <span class="stat-value">${disabled.age['<= 49']}</span></div>
                    <div class="stat-item"><span>50 - 64</span> <span class="stat-value">${disabled.age['50 >= && <= 64']}</span></div>
                    <div class="stat-item"><span>65 - 74</span> <span class="stat-value">${disabled.age['65 >= && <= 74']}</span></div>
                    <div class="stat-item"><span>75 - 84</span> <span class="stat-value">${disabled.age['75 >= && <= 84']}</span></div>
                    <div class="stat-item"><span>>= 85</span> <span class="stat-value">${disabled.age['85 >=']}</span></div>
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-label">居住狀況</div>
                <div class="stat-values">
                    ${livingOrder.filter(l => disabled.living[l] !== undefined).map(l => `<div class="stat-item"><span>${l}</span> <span class="stat-value">${disabled.living[l]}</span></div>`).join('') || '<div class="stat-item"><span>無資料</span></div>'}
                </div>
            </div>
            <div class="stat-card">
                <div class="stat-label">CMS 等級</div>
                <div class="stat-values">
                    ${Object.keys(disabled.cms).sort().map(l => `<div class="stat-item"><span>${l}</span> <span class="stat-value">${disabled.cms[l]}</span></div>`).join('') || '<div class="stat-item"><span>無資料</span></div>'}
                </div>
            </div>
        </div>
    `;
    return groupHtml;
}
