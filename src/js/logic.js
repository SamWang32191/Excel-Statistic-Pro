/**
 * 識別 Excel 工作表名稱
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
 * Excel 欄位字母轉索引 (A -> 0, B -> 1, ...)
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

// 通用統計初始化模板
export function createEmptyStats() {
    return {
        gender: { '男': 0, '女': 0 },
        age: { '<= 49': 0, '50 >= && <= 64': 0, '65 >= && <= 74': 0, '75 >= && <= 84': 0, '85 >=': 0 },
        living: {},
        cms: {}
    };
}