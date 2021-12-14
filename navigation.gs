var app = SpreadsheetApp;                                                                                                          //Setting up oftenly used variables
var ss = app.getActiveSpreadsheet();
var budget = ss.getSheetByName("Budget");
var reports = ss.getSheetByName("Reports");
var allAccounts = ss.getSheetByName("All Accounts");
var storage = ss.getSheetByName("Storage");
var stocks = ss.getSheetByName("Stocks");
function goBudget() {
  budget.getRange('A1').activate();
}
function goReports() {
  reports.getRange('A1').activate();
}
function goAllAccounts() {
  allAccounts.getRange('A1').activate();
}
function goStocks() {
  stocks.getRange('A1').activate();
}