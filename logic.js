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

// Helper: Column Letter to Index
export function colLetterToIndex(letter) {
    if (!letter) return -1;
    let index = 0;
    const s = letter.toUpperCase();
    for (let i = 0; i < s.length; i++) {
        index = index * 26 + s.charCodeAt(i) - 64;
    }
    return index - 1;
}

export function processCityStats(rows, config, isTaipei = false, targetCategory = null) {
    // Initialize Stats Containers
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
    // Taipei only
    const categoryColIdx = isTaipei ? colLetterToIndex(config.category || '') : -1;

    // Start from index 1 (skip header)
    for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        
        // Taipei Filter Logic
        if (isTaipei && categoryColIdx !== -1) {
            const category = (row[categoryColIdx] || '').toString().trim();
            
            if (targetCategory) {
                // Precise filtering
                if (category !== targetCategory) continue;
            } else {
                // Default Taipei filtering (Union)
                if (category !== '老福' && category !== '身障') continue; 
            }
        }

        // Gender
        if (genderColIdx !== -1 && row[genderColIdx] !== undefined) {
            const gender = row[genderColIdx].toString().trim();
            if (gender === '男') stats.gender['男']++;
            else if (gender === '女') stats.gender['女']++;
        }
        
        // Age
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

        // Living
        if (livingColIdx !== -1 && row[livingColIdx] !== undefined) {
            let living = row[livingColIdx].toString().trim();
            if (living) {
                if (living === '其他類型居住狀況') living = '其他';
                stats.living[living] = (stats.living[living] || 0) + 1;
            }
        }

        // CMS
        if (cmsColIdx !== -1 && row[cmsColIdx] !== undefined) {
            const cms = row[cmsColIdx].toString().trim();
            if (cms) stats.cms[cms] = (stats.cms[cms] || 0) + 1;
        }
    }

    return stats;
}