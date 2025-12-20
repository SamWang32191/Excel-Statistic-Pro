import { describe, it, expect } from 'vitest';
import { identifyWorksheets, colLetterToIndex } from '../js/logic';
import * as XLSX from 'xlsx';

describe('Multi-Sheet Identification', () => {
    it('should identify Taipei and New Taipei sheets when both exist', () => {
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, [], "台北01");
        XLSX.utils.book_append_sheet(workbook, [], "新北01");

        const result = identifyWorksheets(workbook);
        
        expect(result.taipeiSheetName).toBe("台北01");
        expect(result.newTaipeiSheetName).toBe("新北01");
    });

    it('should identify only Taipei sheet when New Taipei is missing', () => {
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, [], "台北01");

        const result = identifyWorksheets(workbook);
        
        expect(result.taipeiSheetName).toBe("台北01");
        expect(result.newTaipeiSheetName).toBeNull();
    });

    it('should match partial names correctly', () => {
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, [], "2025台北個案");
        XLSX.utils.book_append_sheet(workbook, [], "2025新北個案");

        const result = identifyWorksheets(workbook);
        
        expect(result.taipeiSheetName).toBe("2025台北個案");
        expect(result.newTaipeiSheetName).toBe("2025新北個案");
    });
    
    it('should return nulls when no matching sheets found', () => {
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, [], "Sheet1");

        const result = identifyWorksheets(workbook);
        
        expect(result.taipeiSheetName).toBeNull();
        expect(result.newTaipeiSheetName).toBeNull();
    });
});

describe('Column Letter to Index', () => {
    it('should convert single letters correctly', () => {
        expect(colLetterToIndex('A')).toBe(0);
        expect(colLetterToIndex('B')).toBe(1);
        expect(colLetterToIndex('Z')).toBe(25);
    });

    it('should convert double letters correctly', () => {
        expect(colLetterToIndex('AA')).toBe(26);
        expect(colLetterToIndex('AB')).toBe(27);
    });

    it('should return -1 for empty or invalid input', () => {
        expect(colLetterToIndex('')).toBe(-1);
        expect(colLetterToIndex(null)).toBe(-1);
    });
});
