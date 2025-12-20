import { describe, it, expect } from 'vitest';
import * as XLSX from 'xlsx';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';
import { processTaipeiCaseStats } from '../js/taipei';
import { processNewTaipeiCaseStats } from '../js/newtaipei';
import { processServiceStats, identifyWorksheets } from '../js/logic';

// 在 ESM 中取得 __dirname
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// 從 src/tests 向上兩層到專案根目錄
const EXAMPLE_DIR = path.join(__dirname, '../../example-file');

function readExcel(filename) {
    const filePath = path.join(EXAMPLE_DIR, filename);
    const fileBuffer = fs.readFileSync(filePath);
    return XLSX.read(fileBuffer, { type: 'buffer' });
}

describe('整合測試 - 真實 Excel 檔案', () => {
    it('個案名單統計（台北 & 新北）', () => {
        const workbook = readExcel('個案名單(測試).xlsx');
        const { taipeiSheetName, newTaipeiSheetName } = identifyWorksheets(workbook);
        
        expect(taipeiSheetName).toBeTruthy();
        expect(newTaipeiSheetName).toBeTruthy();

        // 測試台北老福
        const taipeiRows = XLSX.utils.sheet_to_json(workbook.Sheets[taipeiSheetName], { header: 1 });
        const tpConfig = { category: 'B', gender: 'H', age: 'G', living: 'P', cms: 'K' };
        
        const tpElderly = processTaipeiCaseStats(taipeiRows, tpConfig, '老福');
        const totalTpElderly = tpElderly.gender['男'] + tpElderly.gender['女'];
        expect(totalTpElderly).toBe(621);

        // 測試台北身障
        const tpDisabled = processTaipeiCaseStats(taipeiRows, tpConfig, '身障');
        const totalTpDisabled = tpDisabled.gender['男'] + tpDisabled.gender['女'];
        expect(totalTpDisabled).toBe(30);

        // 測試新北
        const newTaipeiRows = XLSX.utils.sheet_to_json(workbook.Sheets[newTaipeiSheetName], { header: 1 });
        const ntConfig = { gender: 'G', age: 'F', living: 'O', cms: 'J' };
        
        const ntStats = processNewTaipeiCaseStats(newTaipeiRows, ntConfig);
        const totalNt = ntStats.gender['男'] + ntStats.gender['女'];
        expect(totalNt).toBe(804);
    });

    it('服務清冊統計（新北）', () => {
        const workbook = readExcel('服務清冊-新北(測試).xlsx');
        const firstSheetName = workbook.SheetNames[0];
        const rows = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheetName], { header: 1 });
        
        const config = { headerRowIdx: 5, district: 'U', category: 'G', subsidy: 'O' };
        const stats = processServiceStats(rows, config);
        
        // 驗證行政區統計
        expect(stats.districtStats['三重區']).toBe(628);
        expect(stats.districtStats['蘆洲區']).toBe(52);
        expect(stats.districtStats['板橋區']).toBe(510);
        expect(stats.districtStats['新莊區']).toBe(553);
        
        // 驗證服務類別統計
        expect(stats.categoryStats['BA07 協助沐浴及洗頭']).toBe(364);
        expect(stats.categoryStats['BA20 陪伴服務']).toBe(385);
        expect(stats.categoryStats['BA13 陪同外出']).toBe(345);
        expect(stats.categoryStats['BA15-1 家務協助(自用)']).toBe(226);
        
        // 驗證補助比率統計
        expect(stats.subsidyStats['84%']).toBe(1580);
        expect(stats.subsidyStats['100%']).toBe(337);
        expect(stats.subsidyStats['95%']).toBe(252);
        
        // 驗證總筆數
        const totalRecords = Object.values(stats.districtStats).reduce((a, b) => a + b, 0);
        expect(totalRecords).toBe(2169);
    });

    it('服務清冊統計（台北老福）', () => {
        const workbook = readExcel('服務清冊-台北老福(測試).xlsx');
        const firstSheetName = workbook.SheetNames[0];
        const rows = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheetName], { header: 1 });
        
        const config = { headerRowIdx: 5, district: 'U', category: 'G', subsidy: 'O' };
        const stats = processServiceStats(rows, config);
        
        // 驗證行政區統計
        expect(stats.districtStats['文山區']).toBe(554);
        expect(stats.districtStats['萬華區']).toBe(370);
        expect(stats.districtStats['中山區']).toBe(344);
        expect(stats.districtStats['中正區']).toBe(301);
        
        // 驗證服務類別統計
        expect(stats.categoryStats['BA13 陪同外出']).toBe(374);
        expect(stats.categoryStats['BA20 陪伴服務']).toBe(359);
        expect(stats.categoryStats['BA15-1 家務協助(自用)']).toBe(343);
        expect(stats.categoryStats['BA07 協助沐浴及洗頭']).toBe(324);
        
        // 驗證補助比率統計
        expect(stats.subsidyStats['84%']).toBe(1909);
        expect(stats.subsidyStats['100%']).toBe(472);
        expect(stats.subsidyStats['95%']).toBe(74);
        
        // 驗證總筆數
        const totalRecords = Object.values(stats.districtStats).reduce((a, b) => a + b, 0);
        expect(totalRecords).toBe(2455);
    });

    it('服務清冊統計（台北身障）', () => {
        const workbook = readExcel('服務清冊-台北身障(測試).xlsx');
        const firstSheetName = workbook.SheetNames[0];
        const rows = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheetName], { header: 1 });
        
        const config = { headerRowIdx: 5, district: 'U', category: 'G', subsidy: 'O' };
        const stats = processServiceStats(rows, config);
        
        // 驗證行政區統計
        expect(stats.districtStats['文山區']).toBe(34);
        expect(stats.districtStats['萬華區']).toBe(12);
        expect(stats.districtStats['大同區']).toBe(11);
        expect(stats.districtStats['中正區']).toBe(10);
        
        // 驗證服務類別統計
        expect(stats.categoryStats['BA13 陪同外出']).toBe(23);
        expect(stats.categoryStats['BA02 基本日常照顧']).toBe(14);
        expect(stats.categoryStats['BA20 陪伴服務']).toBe(11);
        expect(stats.categoryStats['BA15-1 家務協助(自用)']).toBe(10);
        
        // 驗證補助比率統計
        expect(stats.subsidyStats['84%']).toBe(42);
        expect(stats.subsidyStats['100%']).toBe(27);
        expect(stats.subsidyStats['95%']).toBe(24);
        
        // 驗證總筆數
        const totalRecords = Object.values(stats.districtStats).reduce((a, b) => a + b, 0);
        expect(totalRecords).toBe(93);
    });
});
