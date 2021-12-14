var app = SpreadsheetApp;                                                                                              //Setting up oftenly used variables
var ss = app.getActiveSpreadsheet();
var budget = ss.getSheetByName("Budget");
var reports = ss.getSheetByName("Reports");
var allAccounts = ss.getSheetByName("All Accounts");
var storage = ss.getSheetByName("Storage");
var activeSheet = ss.getActiveSheet();
var ui = SpreadsheetApp.getUi();
var scriptProperties = PropertiesService.getScriptProperties(); 
/*--------------------------------------------------------------------------*/ 

