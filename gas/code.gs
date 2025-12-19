/**
 * 處理 POST 請求
 * 接收格式: { "sheetName": "11403", "data": { "男": 10, "女": 20, "<= 49": 5, ... } }
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
    
    var outputData = [
      ["男", data["男"] || 0],
      ["女", data["女"] || 0],
      ["<= 49", data["<= 49"] || 0],
      ["50 >= && <= 64", data["50 >= && <= 64"] || 0],
      ["65 >= && <= 74", data["65 >= && <= 74"] || 0],
      ["75 >= && <= 84", data["75 >= && <= 84"] || 0],
      ["85 >=", data["85 >="] || 0]
    ];
    
    sheet.getRange(1, 1, outputData.length, 2).setValues(outputData);
    
    return ContentService.createTextOutput(JSON.stringify({ "status": "success", "message": "資料已成功儲存至 " + sheetName }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ "status": "error", "message": error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
