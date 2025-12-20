import { describe, it, expect } from 'vitest';
import { identifyWorksheets, processCityStats, colLetterToIndex } from './logic';
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

describe('Taipei Data Filtering', () => {
    // Mock Helper for Column Index
    // const colToIndex = (char) => char.charCodeAt(0) - 65; // A=0, B=1... (Already imported)

    it('should filter out rows where Category is not 老福 or 身障', () => {
        const config = { category: 'B', gender: 'H' }; // Category at index 1, Gender at index 7
        const rows = [
            [], // Header (skipped)
            [null, '老福', null, null, null, null, null, '男'], // Valid
            [null, '身障', null, null, null, null, null, '女'], // Valid
            [null, '其它', null, null, null, null, null, '男'], // Invalid
            [null, '', null, null, null, null, null, '女'],     // Invalid
        ];

        const stats = processCityStats(rows, config, true); // isTaipei = true

        expect(stats.gender['男']).toBe(1);
        expect(stats.gender['女']).toBe(1);
    });

    it('should NOT filter rows for New Taipei', () => {
        const config = { gender: 'G' }; // Gender at index 6
        const rows = [
            [],
            [null, '其它', null, null, null, null, '男'], // Valid (No filter)
            [null, '', null, null, null, null, '女'],     // Valid (No filter)
        ];

        const stats = processCityStats(rows, config, false); // isTaipei = false

        expect(stats.gender['男']).toBe(1);
        expect(stats.gender['女']).toBe(1);
    });

    it('should allow filtering by specific category value', () => {
        const config = { category: 'B', gender: 'H' };
        const rows = [
            [],
            [null, '老福', null, null, null, null, null, '男'],
            [null, '身障', null, null, null, null, null, '女'],
        ];

        // Test Specific Filter: '老福'
        const statsElderly = processCityStats(rows, config, true, '老福');
        expect(statsElderly.gender['男']).toBe(1);
        expect(statsElderly.gender['女']).toBe(0);

        // Test Specific Filter: '身障'
        const statsDisabled = processCityStats(rows, config, true, '身障');
        expect(statsDisabled.gender['男']).toBe(0);
        expect(statsDisabled.gender['女']).toBe(1);
    });
});
