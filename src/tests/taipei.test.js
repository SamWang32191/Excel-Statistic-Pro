import { describe, it, expect } from 'vitest';
import { processTaipeiCaseStats } from '../js/taipei';

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

    it('should correctly categorize age into groups', () => {
        const config = { category: 'B', age: 'G' };
        const rows = [
            [],
            [null, '老福', null, null, null, null, 40],
            [null, '老福', null, null, null, null, 50],
            [null, '老福', null, null, null, null, 65],
            [null, '老福', null, null, null, null, 75],
            [null, '老福', null, null, null, null, 85],
        ];

        const stats = processTaipeiCaseStats(rows, config, '老福');

        expect(stats.age['<= 49']).toBe(1);
        expect(stats.age['50 >= && <= 64']).toBe(1);
        expect(stats.age['65 >= && <= 74']).toBe(1);
        expect(stats.age['75 >= && <= 84']).toBe(1);
        expect(stats.age['85 >=']).toBe(1);
    });

    it('should map living status correctly', () => {
        const config = { category: 'B', living: 'P' };
        const rows = [
            [],
            [null, '老福', null, null, null, null, null, null, null, null, null, null, null, null, null, '獨居'],
            [null, '老福', null, null, null, null, null, null, null, null, null, null, null, null, null, '其他類型居住狀況'],
        ];

        const stats = processTaipeiCaseStats(rows, config, '老福');

        expect(stats.living['獨居']).toBe(1);
        expect(stats.living['其他']).toBe(1);
    });
});
