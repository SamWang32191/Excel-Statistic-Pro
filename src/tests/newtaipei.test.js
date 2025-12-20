import { describe, it, expect } from 'vitest';
import { processNewTaipeiCaseStats } from '../js/newtaipei';

describe('新北資料處理', () => {
    it('新北不應該依類別過濾資料列', () => {
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

    it('應該正確將新北的年齡分類到各年齡區間', () => {
        const config = { age: 'F' };
        const rows = [
            [],
            [null, null, null, null, null, 40],
            [null, null, null, null, null, 50],
            [null, null, null, null, null, 65],
            [null, null, null, null, null, 75],
            [null, null, null, null, null, 85],
        ];

        const stats = processNewTaipeiCaseStats(rows, config);

        expect(stats.age['<= 49']).toBe(1);
        expect(stats.age['50 >= && <= 64']).toBe(1);
        expect(stats.age['65 >= && <= 74']).toBe(1);
        expect(stats.age['75 >= && <= 84']).toBe(1);
        expect(stats.age['85 >=']).toBe(1);
    });

    it('應該正確對應新北的居住狀況', () => {
        const config = { living: 'O' };
        const rows = [
            [],
            [null, null, null, null, null, null, null, null, null, null, null, null, null, null, '獨居'],
            [null, null, null, null, null, null, null, null, null, null, null, null, null, null, '其他類型居住狀況'],
        ];

        const stats = processNewTaipeiCaseStats(rows, config);

        expect(stats.living['獨居']).toBe(1);
        expect(stats.living['其他']).toBe(1);
    });
});
