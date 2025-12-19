/**
 * 處理 POST 請求
 * 接收格式: { 
 *   "sheetName": "台北03", 
 *   "rows": [["【性別統計】", ""], ["男", 10], ["女", 20], ...]
 * }
 * 前端已處理好所有統計邏輯與格式，GAS 只負責寫入
 */
function doPost(e) {
  try {
    var params = JSON.parse(e.postData.contents);
    var sheetName = params.sheetName;
    var rows = params.rows;
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    
    // 如果工作表不存在，則建立它
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    
    // 找出目前最後一行，準備從其下方開始寫入 (如果是空表則從第一行開始)
    var lastRow = sheet.getLastRow();
    var startRow = lastRow > 0 ? lastRow + 2 : 1; // 若有舊資料，空一行後再寫入
    
    // 直接寫入前端已格式化好的資料
    if (rows && rows.length > 0) {
      sheet.getRange(startRow, 1, rows.length, 2).setValues(rows);
    }
    
    return ContentService.createTextOutput(JSON.stringify({ "status": "success", "message": "資料已成功儲存至 " + sheetName }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ "status": "error", "message": error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
