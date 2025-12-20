import { describe, it, expect } from 'vitest';
import { identifyWorksheets, colLetterToIndex } from '../js/logic';
import { processTaipeiCaseStats } from '../js/taipei';
import { processNewTaipeiCaseStats } from '../js/newtaipei';
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
    it('should filter out rows where Category is not the target category', () => {
        const config = { category: 'B', gender: 'H' };
        const rows = [
            [], // Header (skipped)
            [null, '老福', null, null, null, null, null, '男'], // Valid for '老福'
            [null, '身障', null, null, null, null, null, '女'], // Invalid for '老福'
        ];

        const stats = processTaipeiCaseStats(rows, config, '老福');

        expect(stats.gender['男']).toBe(1);
        expect(stats.gender['女']).toBe(0);
    });

    it('should handle age filtering for Taipei', () => {
        const config = { category: 'B', gender: 'H', age: 'G' };
        const rows = [
            [],
            [null, '老福', null, null, null, null, 65, '男'],
            [null, '老福', null, null, null, null, 50, '女'],
        ];

        const stats65Plus = processTaipeiCaseStats(rows, config, '老福', { min: 65 });
        expect(stats65Plus.gender['男']).toBe(1);
        expect(stats65Plus.gender['女']).toBe(0);

        const stats50to64 = processTaipeiCaseStats(rows, config, '老福', { min: 50, max: 64 });
        expect(stats50to64.gender['男']).toBe(0);
        expect(stats50to64.gender['女']).toBe(1);
    });
});

describe('New Taipei Data Processing', () => {
    it('should NOT filter rows by category for New Taipei', () => {
        const config = { gender: 'G' };
        const rows = [
            [],
            [null, '其它', null, null, null, null, '男'],
            [null, '', null, null, null, null, '女'],
        ];

        const stats = processNewTaipeiCaseStats(rows, config);

        expect(stats.gender['男']).toBe(1);
        expect(stats.gender['女']).toBe(1);
    });
});
