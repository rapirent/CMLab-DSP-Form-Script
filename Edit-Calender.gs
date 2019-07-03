function myFunction() {
    // get spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('工作表1');

  var masterCal = CalendarApp.getCalendarById("");

  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();

  for (var i = 1; i < values.length; i++) {
    if (values[i][3] == "X") {
      var date = values[i][0];
      var host = values[i][1];
      masterCal.createAllDayEvent("DSP組 Meeting 講者:"+host,date); 
      sheet.getRange(i+1, 4).setValue("O"); 
    }   
  } 
}
