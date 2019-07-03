// For Weekly Meeting Speaker Sign Up
function myFunction() {
  // get spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('工作表1');

  var masterCal = CalendarApp.getCalendarById("");

  var dataRange = sheet.getDataRange();

  var data = sheet.getRange(sheet.getLastRow(),1,1,4).getValues();
  var date = data[0][3];
  var host = data[0][2];
  var title = data[0][4]?data[0][4]:'TBD';
  masterCal.createAllDayEvent("DSP組 Meeting 講者: "+host,date,{description: '主題: '+ title}); 
}

