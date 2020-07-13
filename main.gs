const STACKEXCHANGE_API_KEY = "";
const baseUrl = "https://api.stackexchange.com/2.2/";
const sheetID = "SSID";
const sheet = SpreadsheetApp.openById(SSID).getSheetByName("Sheet1");

/*
Example params object:
var params = {
  "pagesize": 100,
  "fromdate": 1220227200,
  "todate": 1594598400,
  "order": "desc",
  "sort": "activity",
  "tagged": "google-api",
  "site": "stackoverflow"
}
*/

function get(params) {  
  var endpoint = "questions";
  var count = PropertiesService.getScriptProperties().getProperty("count") > 0 ? PropertiesService.getScriptProperties().getProperty("count") : 1;
  
  var rowsToRemove = sheet.getDataRange().getNumRows() % 100;
  var numRows = sheet.getDataRange().getNumRows();
  
  if (count > 1) {
    sheet.getRange((numRows - (rowsToRemove - 1)) + ":" + numRows)
         .deleteCells(SpreadsheetApp.Dimension.ROWS);
  }
  
  do {
    try {      
      var response = JSON.parse(UrlFetchApp.fetch(baseUrl + urlParameterfy(params) + "key=" + STACKEXCHANGE_API_KEY + "&page=" + parseInt(count))
                         .getContentText();
      PropertiesService.getScriptProperties().setProperty("count", count);
      
      for (var i = 0; i < response["items"].length; i++) {
        sheet.appendRow([response["items"][i]["title"], response["items"][i]["link"], JSON.stringify(response["items"][i]["tags"])]);
      }
      count++
    }
    catch (e) {
      console.log(e);
      Utilities.sleep(100 * count);
      count++;
    }
  } while (response["has_more"] == true)
    
  if response["has_more"] === false resetProperties();
}

function resetProperties() {
 PropertiesService.getScriptProperties().deleteAllProperties(); 
}

function urlParameterfy(params) {
  var keys = Object.keys(params);
  var returnStr = "?";
  
  keys.forEach(function(key) {
    returnStr+= key + "=" + params[key] + "&";
  });
  
  return returnStr;
}
