/**
 * 處理 POST 請求
 * 接收格式: { 
 *   "sheetName": "台北03", 
 *   "data": { "男": 10, "女": 20, "<= 49": 5, ..., "獨居": 5, "與家人同住": 15, ... }
 * }
 * 所有統計資料（性別、年齡、居住狀況）都合併在 data 物件中
 */
function doPost(e) {
  try {
    var params = JSON.parse(e.postData.contents);
    var sheetName = params.sheetName;
    var data = params.data;
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    
    // 如果工作表不存在，則建立它
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    } else {
      sheet.clear(); // 清除舊資料
    }
    
    // 預定義的性別與年齡欄位 key
    var genderKeys = ["男", "女"];
    var ageKeys = ["<= 49", "50 >= && <= 64", "65 >= && <= 74", "75 >= && <= 84", "85 >="];
    var predefinedKeys = genderKeys.concat(ageKeys);
    
    // 性別與年齡統計資料
    var outputData = [
      ["【性別統計】", ""],
      ["男", data["男"] || 0],
      ["女", data["女"] || 0],
      ["", ""],
      ["【年齡分佈】", ""],
      ["<= 49", data["<= 49"] || 0],
      ["50 >= && <= 64", data["50 >= && <= 64"] || 0],
      ["65 >= && <= 74", data["65 >= && <= 74"] || 0],
      ["75 >= && <= 84", data["75 >= && <= 84"] || 0],
      ["85 >=", data["85 >="] || 0],
      ["", ""],
      ["【居住狀況】", ""]
    ];
    
    // 依指定順序輸出居住狀況資料
    var livingOrder = ['獨居', '獨居(兩老)', '與家人或其他人同住', '與朋友同住', '其他'];
    
    for (var i = 0; i < livingOrder.length; i++) {
      var key = livingOrder[i];
      if (data[key] !== undefined) {
        outputData.push([key, data[key]]);
      }
    }
    
    sheet.getRange(1, 1, outputData.length, 2).setValues(outputData);
    
    return ContentService.createTextOutput(JSON.stringify({ "status": "success", "message": "資料已成功儲存至 " + sheetName }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ "status": "error", "message": error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
