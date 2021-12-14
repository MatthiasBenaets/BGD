var app = SpreadsheetApp;                                                                                                          //Setting up oftenly used variables
var ss = app.getActiveSpreadsheet();
var budget = ss.getSheetByName("Budget");
var reports = ss.getSheetByName("Reports");
var allAccounts = ss.getSheetByName("All Accounts");
var storage = ss.getSheetByName("Storage");
var activeSheet = ss.getActiveSheet();
var ui = SpreadsheetApp.getUi();
var scriptProperties = PropertiesService.getScriptProperties(); 
/*--------------------------------------------------------------------------*/ 
function addTransaction() {
  if (allAccounts.getRange('H4').getValue() == "To Be Budgeted"){
    var selectedValue = allAccounts.getRange('L4').getValue();
    var selectedAccount = allAccounts.getRange('E4').getValue();
    var i = 0;
    do{                                                                                                                            //Used to add value to budget table
      var search = budget.getRange(8+i,2).getValue();
      i++
    }
    while(search != selectedAccount);
    var j = budget.getRange(7+i,3).getValue();
    budget.getRange(7+i,3).setValue(j+selectedValue);
  }else if(allAccounts.getRange('H4').getValue() == "Transfer"){
    var referenceAccount = allAccounts.getRange('G4').getValue();
    var count = 0
    if (allAccounts.getRange('K4').getValue() > "0"){
      do{
        budget.getRange(8+count,2).getValue();
        count++
      }while(budget.getRange(7+count,2).getValue() !=referenceAccount)
      var ref = budget.getRange(7+count,3).getValue()
      var selectedValue = allAccounts.getRange('K4').getValue();
      budget.getRange(7+count,3).setValue(ref + selectedValue);
      count = 0;
      var selectedAccount = allAccounts.getRange('E4').getValue();
      do{
        budget.getRange(8+count,2).getValue();
        count++
      }while(budget.getRange(7+count,2).getValue()!=selectedAccount)
      var ref = budget.getRange(7+count,3).getValue()
      budget.getRange(7+count,3).setValue(ref - selectedValue);
    }else if (allAccounts.getRange('L4').getValue() > "0"){
      do{
        budget.getRange(8+count,2).getValue();
        count++
      }while(budget.getRange(7+count,2).getValue()!=referenceAccount)
      var ref = budget.getRange(7+count,3).getValue()
      budget.getRange(7+count,3).setValue(ref - selectedValue);
      count = 0;
      var selectedAccount = allAccounts.getRange('E4').getValue();
      do{
        budget.getRange(8+count,2).getValue();
        count++
      }while(budget.getRange(7+count,2).getValue()!=selectedAccount)
      var ref = budget.getRange(7+count,3).getValue()
      budget.getRange(7+count,3).setValue(ref + selectedValue);
    };
  }else{saveTransaction();};
  
  allAccounts.getRange("M4").setFormula('=TEXT(F4;"yyyy/mm")');                                                                    //Print month in hidden cells to use for pivot table   

  allAccounts.getRange('E4:M4').activate();
  allAccounts.getRange('E4:M4').insertCells(SpreadsheetApp.Dimension.ROWS);
  allAccounts.getActiveRangeList().setBorder(null, null, null, null, true, true, '#ffffff', SpreadsheetApp.BorderStyle.SOLID);     //All code above to change design of cells
  allAccounts.getRange('E5:M5').activate();                                                                                        //Clear all data validation from last transaction
  allAccounts.getRange('E5:M5').clearDataValidations();
  
  allAccounts.getRange('\'All Accounts\'!E4').setDataValidation(SpreadsheetApp.newDataValidation()                                 //Data validation. Choose between all categories in storage.
  .setAllowInvalid(true)
  .requireValueInRange(allAccounts.getRange('Storage!$B$20:$B$39'), true)                                                          //Max 20 Accounts
  .build());
  
  allAccounts.getRange('\'All Accounts\'!F4').setDataValidation(SpreadsheetApp.newDataValidation()                                 //Add calendar and date validation=valid
  .setAllowInvalid(true)
  .requireDate()
  .build());
  var date = Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy");                                                              //variable for current date
  allAccounts.getRange('F4').setValue(date)                                                                                        //print date
  
  allAccounts.getRange('\'All Accounts\'!H4').setDataValidation(SpreadsheetApp.newDataValidation()                                 //Data validation. Choose between all categories in storage.
  .setAllowInvalid(true)
  .requireValueInRange(allAccounts.getRange('Storage!$D$1:$Z$1'), true)
  .build());
  
  allAccounts.getRange('\'All Accounts\'!G4').setDataValidation(SpreadsheetApp.newDataValidation()                                 //Data validation. Choose between all categories in storage.
  .setAllowInvalid(true)
  .requireValueInRange(allAccounts.getRange('Storage!$B$20:$B$39'), true)
  .build());
  
  //For I4 - see onEdit
}
/*--------------------------------------------------------------------------*/  
function automationOnEdit(){
  var i = 0;
  if (allAccounts.getRange('H4').getValue() != ""){                                                                                //Check if H4 is filled in to auto populate I4
    var selectedCat = allAccounts.getRange('H4').getValue();
    do{                                                                                                                        //Loop to check if selected cat = search. Counting i for cell address.
      var search = storage.getRange(1,4+i).getValue();
      i++
    }
    while(search != selectedCat);
  }
  var cell = storage.getRange(1, 3+i).getA1Notation();                                                                             //Cell address
  var column = cell.charAt(0);                                                                                                     //First letter of cell address
  
  allAccounts.getRange('\'All Accounts\'!I4').setDataValidation(SpreadsheetApp.newDataValidation()                                 //Data validation. Choose between all categories in storage.
  .setAllowInvalid(true)
  .requireValueInRange(activeSheet.getRange('Storage!'+column+'2'+':'+column), true)                                               //Range in storage for validation
  .build());
  
  if (allAccounts.getRange('H4').getValue() == "To Be Budgeted"){
    allAccounts.getRange('I4').setValue("Subcategory not needed");
    //allAccounts.getRange('K4').setValue("0");
  }

  /*if (allAccounts.getRange('K4').getValue() > "0"){                                                                                //Automated values if income or outcome is given
    allAccounts.getRange('L4').setValue("0");
  }else if (allAccounts.getRange('L4').getValue() > "0"){
    allAccounts.getRange('K4').getValue("0");
  };*/  
}
/*--------------------------------------------------------------------------*/ 
function saveTransaction(){
  app.flush();
  var i = 0;                                                                                                                       //Check if all cells are filled in for transaction
  if (allAccounts.getRange('E4').getValue() != "" 
  && allAccounts.getRange('F4').getValue() != "" 
  && allAccounts.getRange('G4').getValue() != "" 
  && allAccounts.getRange('H4').getValue() != "" 
  && allAccounts.getRange('I4').getValue() != ""){
    var referenceSubcategory  = allAccounts.getRange('I4').getValue();
    var referenceAccount = allAccounts.getRange('E4').getValue();
    if (allAccounts.getRange('K4').getValue() > "0"){                                                                //Check what is filled in and respond accordingly. Added to budget or transaction
      var referenceValue = allAccounts.getRange('K4').getValue();                                                                  //TRANSACTION CODE BELOW
      do{                                                                                                                          //Used to add value to budget table
      var search = budget.getRange(4+i, 5).getValue();
      i++
      }
      while(search != referenceSubcategory);
      var existingValue = budget.getRange(3+i, 7).getValue();
      budget.getRange(3+i,7).setValue(existingValue-referenceValue);
      
      i = 0;
      do{                                                                                                                          //Search for given account in budget. Subtract value of total
        var existingValue = budget.getRange(8+i,2).getValue();
        i++
      }while(existingValue != referenceAccount);
      existingValue = budget.getRange(7+i,3).getValue();
      var search = budget.getRange(7+i,3).getA1Notation();
      var referenceValue = allAccounts.getRange('K4').getValue();
      budget.getRange(search).setValue(existingValue-referenceValue);
      
    }else if (allAccounts.getRange('L4').getValue() > "0"){                                                                        //INCOME CODE BELOW
      var referenceSubcategory  = allAccounts.getRange('I4').getValue();
      var referenceValue = allAccounts.getRange('L4').getValue();
      do{                                                                                                                          //Search for given account in budget. Print added value to total
        var search = budget.getRange(4+i,5).getValue();
        i++
      }while(search != referenceSubcategory);
      var existingValue = budget.getRange(3+i,7).getValue();
      budget.getRange(3+i,7).setValue(existingValue+referenceValue);
      
      i=0;
      do{                                                                                                                          //Search for given account in budget. Add value of total
        var existingValue = budget.getRange(8+i,2).getValue();
        i++
      }while(existingValue != referenceAccount);
      existingValue = budget.getRange(7+i,3).getValue();
      var search = budget.getRange(7+i,3).getA1Notation();
      var referenceValue = allAccounts.getRange('L4').getValue();
      budget.getRange(search).setValue(existingValue+referenceValue);
    }else{return;}
  }else{return;};
  app.flush();
}