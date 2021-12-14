var app = SpreadsheetApp;                                                                                              //Setting up oftenly used variables
var ss = app.getActiveSpreadsheet();
var budget = ss.getSheetByName("Budget");
var reports = ss.getSheetByName("Reports");
var allAccounts = ss.getSheetByName("All Accounts");
var storage = ss.getSheetByName("Storage");
var stocks = ss.getSheetByName("Stocks")
var activeSheet = ss.getActiveSheet();
var ui = SpreadsheetApp.getUi();
var scriptProperties = PropertiesService.getScriptProperties(); 
var date

function addStock(){
  stocks.getRange('E4:N4').activate();
  stocks.getRange('E4:N4').insertCells(SpreadsheetApp.Dimension.ROWS);
  stocks.getRange('F4:M4').activate();
  stocks.getActiveRangeList().setBorder(null, null, null, null, true, null, '#ffffff', SpreadsheetApp.BorderStyle.SOLID);
  stocks.getRange('L5').activate();
  stocks.getActiveRange().autoFill(stocks.getRange('L4:L5'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  stocks.getRange('L4:L5').activate();
  stocks.setCurrentCell(stocks.getRange('L5'));
  var date = Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy");                                                              //variable for current date
  stocks.getRange('F4').setValue(date);
  stocks.getRange('I4').activate();
  stock.getActiveRangeList().setHorizontalAlignment('right')
  .setNumberFormat('[$$]#,##0.00');
  stocks.getRange('E4').activate();
}