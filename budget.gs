//Matthias Benaets

var app = SpreadsheetApp;                                                                                              //Setting up oftenly used variables
var ss = app.getActiveSpreadsheet();
var budget = ss.getSheetByName("Budget");
var reports = ss.getSheetByName("Reports");
var allAccounts = ss.getSheetByName("All Accounts");
var storage = ss.getSheetByName("Storage");
var activeSheet = ss.getActiveSheet();
var ui = SpreadsheetApp.getUi();
var scriptProperties = PropertiesService.getScriptProperties();                                                        //Needed to save values from prompt

var a1, a2, a3, a4;
/*--------------------------------------------------------------------------*/ 
function onLoad(e){
  var date = Utilities.formatDate(new Date(), "GMT+1", "MMM")
  if (storage.getRange('B9').getValue() ==  "" ){
    storage.getRange('B9').setValue(date);   
  }else if (date != storage.getRange('B9').getValue()){
    monthTrigger();
    storage.getRange('B9').setValue(date); 
  }else{return;}
}
/*--------------------------------------------------------------------------*/ 
function monthTrigger(){
  var lastMonthRef = new Date();
  lastMonthRef.setDate(1);
  lastMonthRef.setMonth(lastMonthRef.getMonth()-1);

  var spreadsheet = SpreadsheetApp.getActive();                                                                        //Copy all over to new sheet.
  var curDate = Utilities.formatDate(lastMonthRef, "GMT+1", "MMM yyyy")
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Budget'), true);
  SpreadsheetApp.getActiveSpreadsheet().insertSheet(curDate);
  spreadsheet.getRange('Budget!A1:Z1000').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_COLUMN_WIDTHS, false)
  spreadsheet.getRange('Budget!A1:Z1000').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getRange('Budget!A1:Z1000').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  spreadsheet.getRange('E1').setValue(curDate);
  spreadsheet.getRange('I13').setValue(storage.getRange('B8').getValue());
  spreadsheet.getRange('1:1000').activate();
  spreadsheet.getActiveSheet().setRowHeights(1, 1000, 30);
  spreadsheet.getRange('1:1').activate();
  spreadsheet.getActiveSheet().setRowHeight(1, 100);
  spreadsheet.moveActiveSheet(7);
  
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Budget'), true);                                             //save old values in storage
  spreadsheet.getRange('H4:H').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Storage'), true);
  spreadsheet.getRange('B41').activate();
  spreadsheet.getRange('Budget!H4:H').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  budget.getRange('A1').activate();
  
  var i = 41;                                                                                                          //Code below is to reset budgeted and budgeted
  var j = 4;
  do{
  var checker = storage.getRange("A"+ i).getValue();
  if (checker == false){
    budget.getRange("G"+ j).setValue("0");
    budget.getRange("F"+j).setValue("0");
  }
  i++
  j++
  }while(storage.getRange("A"+ i).getValue().toString() != "");
  storage.hideSheet();
}
/*--------------------------------------------------------------------------*/ 
function addAccount(){                                                                                                 //ADD ACCOUNT TO DOCUMENT
  var i = storage.getRange("B2").getValue();                                                                           //Check value on K2 so script knows where to place new data
  
  do{                                                                                                                  //Do until input is not empty
    var responseAcc = ui.prompt('Account Name?', ui.ButtonSet.OK);                                                     //Prompt inputbox
    var button = responseAcc.getSelectedButton();                                                                      //Button variable used for cancel and close
    scriptProperties.setProperty('account',responseAcc.getResponseText());                                             //Save input as property
    if (button == ui.Button.CANCEL || button == ui.Button.CLOSE) {return;};                                            //Stop function is using cancel or close
  }
  while (scriptProperties.getProperty('account') == "" );
  
  do{                                                                                                                  //Ditto do while loop above
    var responseCap = ui.prompt('Start Capital?', ui.ButtonSet.OK);
    var button = responseCap.getSelectedButton();
    scriptProperties.setProperty('capital',responseCap.getResponseText());
    if (button == ui.Button.CANCEL || button == ui.Button.CLOSE) {return;};
  }
  while (isNaN(scriptProperties.getProperty('capital')) === true || scriptProperties.getProperty('capital') == "");    //Repeat loop is input is not a number of empty
  
  budget.getRange(8+i,2).setValue(scriptProperties.getProperty('account'));                                            //Print value from prompt in correct location depending on value on I2   
  budget.getRange(8+i,3).setValue(scriptProperties.getProperty('capital'));                                            //Print start capital
  reports.getRange(8+i,2).setValue('=\'Budget\'!'+ budget.getRange(8+i,2).getA1Notation());                            //Print all account links from budget to other sheets
  reports.getRange(8+i,3).setValue('=\'Budget\'!'+ budget.getRange(8+i,3).getA1Notation());
  allAccounts.getRange(8+i,2).setValue('=\'Budget\'!'+ budget.getRange(8+i,2).getA1Notation());   
  allAccounts.getRange(8+i,3).setValue('=\'Budget\'!'+ budget.getRange(8+i,3).getA1Notation());
  storage.getRange("B2").setValue(i+1);                                                                                //Add +1 to K2
  storage.getRange(20+i,2).setValue(scriptProperties.getProperty('account'));                                          //Save name in storage
  
  allAccounts.getRange('E4').setValue(scriptProperties.getProperty('account'));
  var date = Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy");                                                  //variable for current date
  allAccounts.getRange('F4').setValue(date);
  allAccounts.getRange('G4').setValue("Starting Balance");
  allAccounts.getRange('H4').setValue("To Be Budgeted");
  allAccounts.getRange('I4').setValue("Subcategory not needed");
  allAccounts.getRange('J4').setValue("Starting Balance");
  allAccounts.getRange('K4').setValue("0");
  allAccounts.getRange('L4').setValue(scriptProperties.getProperty('capital'));
  allAccounts.getRange("M4").setFormula('=TEXT(F4;"mmm")');                                                            //Print month in hidden cells to use for pivot table   

  allAccounts.getRange('E4:M4').activate();
  allAccounts.getRange('E4:M4').insertCells(SpreadsheetApp.Dimension.ROWS);                                            //All code above to change design of cells
  allAccounts.getActiveRangeList().setBorder(null, null, null, null, true, true, '#ffffff', SpreadsheetApp.BorderStyle.SOLID);     
  allAccounts.getRange('E5:M5').activate();                                                                            //Clear all data validation from last transaction
  allAccounts.getRange('E5:M5').clearDataValidations();
  
  allAccounts.getRange('\'All Accounts\'!E4').setDataValidation(SpreadsheetApp.newDataValidation()                     //Data validation. Choose between all categories in storage.
  .setAllowInvalid(true)
  .requireValueInRange(allAccounts.getRange('Storage!$B$20:$B$39'), true)                                              //Max 20 Accounts
  .build());
  
  allAccounts.getRange('\'All Accounts\'!F4').setDataValidation(SpreadsheetApp.newDataValidation()                     //Add calendar and date validation=valid
  .setAllowInvalid(true)
  .requireDate()
  .build());
  var date = Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy");                                                  //variable for current date
  allAccounts.getRange('F4').setValue(date)                                                                            //print date
  
  allAccounts.getRange('\'All Accounts\'!H4').setDataValidation(SpreadsheetApp.newDataValidation()                     //Data validation. Choose between all categories in storage.
  .setAllowInvalid(true)
  .requireValueInRange(allAccounts.getRange('Storage!$D$1:$Z$1'), true)
  .build());
  
  budget.getRange('A1').activate();
  
  app.flush();
}
/*--------------------------------------------------------------------------*/ 
function addCategory(){                                                                                                //ADD CATEGORY TO DOCUMENT
  var l = storage.getRange("B3").getValue() + storage.getRange("B4").getValue();
  
  do{                                                                                                                  //Do until input is not empty
    var responseCat = ui.prompt('Category Name?', ui.ButtonSet.OK);                                                    //Prompt inputbox
    var button = responseCat.getSelectedButton();                                                                      //Button variable used for cancel and close
    scriptProperties.setProperty('category',responseCat.getResponseText());                                            //Save input as property
    if (button == ui.Button.CANCEL || button == ui.Button.CLOSE) {return;};                                            //Stop fuction is using cancel or close
  }
  while (scriptProperties.getProperty('category') == "" );
  
  budget.getRange(4+l,5,1,4).activate();                                                                               //Select cell and all 4 column next to it.
  activeSheet.getActiveRangeList().setBackground('#e6f5fa')                                                            //Change design of cells
  .setBorder(null, null, null, null, true, true, '#e6f5fa', SpreadsheetApp.BorderStyle.SOLID);
  var budgeted = activeSheet.getRange(4+l,6).setValue("0").getA1Notation();                                            //Adding default 0 to budgeted and activities. Also save address to use in next lines of code.
  var activity = activeSheet.getRange(4+l,7).setValue("0").getA1Notation();
  var lastMonth = allAccounts.getRange(41+l,2).getA1Notation();                                                        //Get address of storage location
  budget.getRange(4+l,8).setFormula('='+ budgeted + '+' + activity+ '+\'Storage\'!' + lastMonth);
  
  budget.getRange(4+l, 5).setValue(scriptProperties.getProperty('category'));                                          //Print value on correct location
  storage.getRange("B3").setValue(l+1-storage.getRange("B4").getValue());                                              //Save amount in storage  (-storage.getrange... is for compensating since i is sum of 2 variables)
  storage.getRange(1,6+l-storage.getRange("B4").getValue()).setValue(scriptProperties.getProperty('category'));        //Save category in storage
  storage.getRange(41+l,1).setValue("true");
  
  addSubCategory();                                                                                                    //Run addSubCategory
  
  budget.getRange(budgeted).setFormula('=SUM('+ a1 + ':' + a3+ ')');                                                   //This crappy code is used to print the formula with the address variables
  budget.getRange(activity).setFormula('=SUM('+ a2 + ':' + a4 + ')');
}
/*--------------------------------------------------------------------------*/ 
function addSubCategory(){
  var j = 2;                                                                                                           //Variable for storage
  var k = 1;
  do{
    var i = storage.getRange("B3").getValue() + storage.getRange("B4").getValue();                                     //Ditto addCategory
    var responseSub = ui.prompt('Subcategory Name?', ui.ButtonSet.OK);
    var button = responseSub.getSelectedButton();
    scriptProperties.setProperty('subcategory',responseSub.getResponseText());
    if (button == ui.Button.CANCEL || button == ui.Button.CLOSE) {return;}; 
    budget.getRange(4+i,5).setValue(scriptProperties.getProperty('subcategory'));
    
    if (k == 1){                                                                                                       //This crappy code is used to get the start address for the category formulas
    a1 = budget.getRange(4+i,6).getA1Notation();
    a2 = budget.getRange(4+i,7).getA1Notation();
    }
    
    var budgeted = budget.getRange(4+i,6).setValue("0").getA1Notation();                                               //Adding default 0 to budgeted and activities. Also save address to use in next lines of code.
    var activity = budget.getRange(4+i,7).setValue("0").getA1Notation();
    var available = budget.getRange(4+i,8).getA1Notation();
    var lastMonth = allAccounts.getRange(40+i,2).getA1Notation();                                                      //Get address of storage location
    budget.getRange(4+i,8).setFormula('='+ budgeted + '+' + activity + '+\'Storage\'!' + lastMonth);                   //Print formula to add budgeted to activities.
    
    var spreadsheet = SpreadsheetApp.getActive();                                                                      //All this code is for conditional formatting (colour green or red)
    var conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
    conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
    conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
    .setRanges([spreadsheet.getRange(available)])
    .whenNumberGreaterThan(0.001)
    .setBackground('#B7E1CD')
    .build());
    spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
    conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
    conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
    .setRanges([spreadsheet.getRange(available)])
    .whenNumberLessThan(-0.001)
    .setBackground('#E6B8AF')
    .build());
    spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
    
    var yn = ui.alert('Add another subcategory?', ui.ButtonSet.YES_NO);                                                //Ask if user want to add another subcategory
    
    storage.getRange("B4").setValue(i+1-storage.getRange("B3").getValue());
    storage.getRange(j,5+storage.getRange("B3").getValue()).setValue(scriptProperties.getProperty('subcategory'));     //Save subcategory in storage
    storage.getRange(41+i,1).setValue("false");
    
    a3 = budget.getRange(4+i,6).getA1Notation();                                                                       //This crappy code is used to get the end address for the category formulas
    a4 = budget.getRange(4+i,7).getA1Notation();
    
    j++
    k++
    app.flush();                                                                                                       //Ensure variable is printed (not optimization problems)
  }
  while (yn == ui.Button.YES);                                                                                         //Do until answer is NO
}