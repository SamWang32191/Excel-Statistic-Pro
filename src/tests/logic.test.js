import { describe, it, expect } from 'vitest';
import { identifyWorksheets, colLetterToIndex } from '../js/logic';
import * as XLSX from 'xlsx';

describe('多工作表識別', () => {
    it('當台北和新北工作表都存在時，應該正確識別', () => {
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, [], "台北01");
        XLSX.utils.book_append_sheet(workbook, [], "新北01");

        const result = identifyWorksheets(workbook);
        
        expect(result.taipeiSheetName).toBe("台北01");
        expect(result.newTaipeiSheetName).toBe("新北01");
    });

    it('當只有台北工作表時，應該正確識別', () => {
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, [], "台北01");

        const result = identifyWorksheets(workbook);
        
        expect(result.taipeiSheetName).toBe("台北01");
        expect(result.newTaipeiSheetName).toBeNull();
    });

    it('應該正確比對部分名稱', () => {
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, [], "2025台北個案");
        XLSX.utils.book_append_sheet(workbook, [], "2025新北個案");

        const result = identifyWorksheets(workbook);
        
        expect(result.taipeiSheetName).toBe("2025台北個案");
        expect(result.newTaipeiSheetName).toBe("2025新北個案");
    });
    
    it('當找不到符合的工作表時應回傳 null', () => {
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, [], "Sheet1");

        const result = identifyWorksheets(workbook);
        
        expect(result.taipeiSheetName).toBeNull();
        expect(result.newTaipeiSheetName).toBeNull();
    });
});

describe('欄位字母轉索引', () => {
    it('應該正確轉換單一字母', () => {
        expect(colLetterToIndex('A')).toBe(0);
        expect(colLetterToIndex('B')).toBe(1);
        expect(colLetterToIndex('Z')).toBe(25);
    });

    it('應該正確轉換雙字母', () => {
        expect(colLetterToIndex('AA')).toBe(26);
        expect(colLetterToIndex('AB')).toBe(27);
    });

    it('空白或無效輸入應回傳 -1', () => {
        expect(colLetterToIndex('')).toBe(-1);
        expect(colLetterToIndex(null)).toBe(-1);
    });
});
