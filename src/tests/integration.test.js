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

        // 測試台北
        const taipeiRows = XLSX.utils.sheet_to_json(workbook.Sheets[taipeiSheetName], { header: 1 });
        // 使用 index.html 預設的欄位配置
        const tpConfig = { category: 'B', gender: 'H', age: 'G', living: 'P', cms: 'K' };
        
        const tpElderly = processTaipeiCaseStats(taipeiRows, tpConfig, '老福');
        // 基本驗證：確保有解析到資料
        const totalTpElderly = tpElderly.gender['男'] + tpElderly.gender['女'];
        expect(totalTpElderly).toBeGreaterThan(0);
        console.log(`Taipei Elderly Count: ${totalTpElderly}`);

        const tpDisabled = processTaipeiCaseStats(taipeiRows, tpConfig, '身障');
        const totalTpDisabled = tpDisabled.gender['男'] + tpDisabled.gender['女'];
        expect(totalTpDisabled).toBeGreaterThan(0);
        console.log(`Taipei Disabled Count: ${totalTpDisabled}`);

        // 測試新北
        const newTaipeiRows = XLSX.utils.sheet_to_json(workbook.Sheets[newTaipeiSheetName], { header: 1 });
        const ntConfig = { gender: 'G', age: 'F', living: 'O', cms: 'J' }; // 使用 index.html 預設的欄位配置
        
        const ntStats = processNewTaipeiCaseStats(newTaipeiRows, ntConfig);
        const totalNt = ntStats.gender['男'] + ntStats.gender['女'];
        expect(totalNt).toBeGreaterThan(0);
        console.log(`New Taipei Count: ${totalNt}`);
    });

    it('服務清冊統計（新北）', () => {
        const workbook = readExcel('服務清冊-新北(測試).xlsx');
        const firstSheetName = workbook.SheetNames[0];
        const rows = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheetName], { header: 1 });
        
        const config = { headerRowIdx: 5, district: 'U', category: 'G', subsidy: 'O' };
        const stats = processServiceStats(rows, config);
        
        expect(Object.keys(stats.districtStats).length).toBeGreaterThan(0);
        expect(Object.keys(stats.categoryStats).length).toBeGreaterThan(0);
        expect(Object.keys(stats.subsidyStats).length).toBeGreaterThan(0);
        
        console.log('New Taipei Service Stats:', JSON.stringify(stats.districtStats, null, 2));
    });

    it('服務清冊統計（台北老福）', () => {
        const workbook = readExcel('服務清冊-台北老福(測試).xlsx');
        const firstSheetName = workbook.SheetNames[0];
        const rows = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheetName], { header: 1 });
        
        const config = { headerRowIdx: 5, district: 'U', category: 'G', subsidy: 'O' };
        const stats = processServiceStats(rows, config);
        
        expect(Object.keys(stats.districtStats).length).toBeGreaterThan(0);
        console.log('Taipei Elderly Service Stats:', JSON.stringify(stats.districtStats, null, 2));
    });

    it('服務清冊統計（台北身障）', () => {
        const workbook = readExcel('服務清冊-台北身障(測試).xlsx');
        const firstSheetName = workbook.SheetNames[0];
        const rows = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheetName], { header: 1 });
        
        const config = { headerRowIdx: 5, district: 'U', category: 'G', subsidy: 'O' };
        const stats = processServiceStats(rows, config);
        
        expect(Object.keys(stats.districtStats).length).toBeGreaterThan(0);
        console.log('Taipei Disabled Service Stats:', JSON.stringify(stats.districtStats, null, 2));
    });
});
