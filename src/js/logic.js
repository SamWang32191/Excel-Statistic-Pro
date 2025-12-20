/**
 * 識別 Excel 工作表名稱
 * @param {Object} workbook - Excel 活頁簿物件
 * @returns {Object} 包含台北與新北工作表名稱的物件
 */
export function identifyWorksheets(workbook) {
    let taipeiSheetName = null;
    let newTaipeiSheetName = null;

    workbook.SheetNames.forEach(name => {
        if (name.includes("台北")) {
            taipeiSheetName = name;
        } else if (name.includes("新北")) {
            newTaipeiSheetName = name;
        }
    });

    return {
        taipeiSheetName,
        newTaipeiSheetName
    };
}

/**
 * Excel 欄位字母轉索引 (例如：A -> 0, B -> 1, ...)
 * @param {string} letter - 欄位字母
 * @returns {number} 0 型索引，若無效則回傳 -1
 */
export function colLetterToIndex(letter) {
    if (!letter) return -1;
    let index = 0;
    const s = letter.toUpperCase();
    for (let i = 0; i < s.length; i++) {
        index = index * 26 + s.charCodeAt(i) - 64;
    }
    return index - 1;
}

/**
 * 建立通用統計初始化模板
 * @returns {Object} 初始化的統計資料結構
 */
export function createEmptyStats() {
    return {
        gender: { '男': 0, '女': 0 },
        age: { '<= 49': 0, '50 >= && <= 64': 0, '65 >= && <= 74': 0, '75 >= && <= 84': 0, '85 >=': 0 },
        living: {},
        cms: {}
    };
}

/**
 * 處理服務清冊統計
 * @param {Array} rows - 工作表資料列陣列
 * @param {Object} config - 欄位配置資訊 { headerRowIdx, district, category, subsidy }
 * @returns {Object} 包含行政區、類別及補助統計結果的物件
 */
export function processServiceStats(rows, config) {
    const headerRowIdx = config.headerRowIdx || 5;
    const districtColIdx = colLetterToIndex(config.district || '');
    const categoryColIdx = colLetterToIndex(config.category || '');
    const subsidyColIdx = colLetterToIndex(config.subsidy || '');

    const districtStats = {};
    const categoryStats = {};
    const subsidyStats = {};

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

    return {
        districtStats,
        categoryStats,
        subsidyStats
    };
}