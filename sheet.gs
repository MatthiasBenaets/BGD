var ssID ="REDACTED";
var formID = "REDACTED";

var wsData = SpreadsheetApp.openById(ssID).getSheetByName("storage");
var form = FormApp.openById(formID);

var app = SpreadsheetApp;                                                                                              //Setting up oftenly used variables
var ss = app.getActiveSpreadsheet();
var budget = ss.getSheetByName("Budget");
var reports = ss.getSheetByName("Reports");
var allAccounts = ss.getSheetByName("All Accounts");
var storage = ss.getSheetByName("Storage");
var forms = ss.getSheetByName("Forms");
var activeSheet = ss.getActiveSheet();

var stoAcc = storage.getRange("B20:B39").getValues().map(function(element){return element.toString();});
var filStoAcc = stoAcc.filter(String);
var stoCat = storage.getRange("C4:C20").getValues().map(function(element){return element.toString();});
var filStoCat = stoCat.filter(String);

var stoCat1 = storage.getRange("F2:F20").getValues().map(function(element){return element.toString();});
var filStoCat1 = stoCat1.filter(String);
var stoCat2 = storage.getRange("G2:G20").getValues().map(function(element){return element.toString();});
var filStoCat2 = stoCat2.filter(String);
var stoCat3 = storage.getRange("H2:H20").getValues().map(function(element){return element.toString();});
var filStoCat3 = stoCat3.filter(String);
var stoCat4 = storage.getRange("I2:I20").getValues().map(function(element){return element.toString();});
var filStoCat4 = stoCat4.filter(String);
var stoCat5 = storage.getRange("J2:J20").getValues().map(function(element){return element.toString();});
var filStoCat5 = stoCat5.filter(String);
var stoCat6 = storage.getRange("K2:K20").getValues().map(function(element){return element.toString();});
var filStoCat6 = stoCat6.filter(String);

function populate(){
  var acc= form.getItems(FormApp.ItemType.LIST).filter(function(acc){return acc.getTitle()==='Account';})[0].asListItem();
  acc.setChoiceValues(filStoAcc);
  var payee= form.getItems(FormApp.ItemType.MULTIPLE_CHOICE).filter(function(payee){return payee.getTitle()==='Payee';})[0].asMultipleChoiceItem();
  payee.setChoiceValues(filStoAcc);
  var category= form.getItems(FormApp.ItemType.LIST).filter(function(category){return category.getTitle()==='Category';})[0].asListItem();
  category.setChoiceValues(filStoCat);

  var category1= form.getItems(FormApp.ItemType.LIST).filter(function(category1){return category1.getTitle()==='Persoonlijk';})[0].asListItem();
  category1.setChoiceValues(filStoCat1);
  var category2= form.getItems(FormApp.ItemType.LIST).filter(function(category2){return category2.getTitle()==='Voeding';})[0].asListItem();
  category2.setChoiceValues(filStoCat2);
  var category3= form.getItems(FormApp.ItemType.LIST).filter(function(category3){return category3.getTitle()==='Transport';})[0].asListItem();
  category3.setChoiceValues(filStoCat3);
  var category4= form.getItems(FormApp.ItemType.LIST).filter(function(category4){return category4.getTitle()==='Hobby';})[0].asListItem();
  category4.setChoiceValues(filStoCat4);
  var category5= form.getItems(FormApp.ItemType.LIST).filter(function(category5){return category5.getTitle()==='Abonnement';})[0].asListItem();
  category5.setChoiceValues(filStoCat5);
  var category6= form.getItems(FormApp.ItemType.LIST).filter(function(category6){return category6.getTitle()==='Spaargeld';})[0].asListItem();
  category6.setChoiceValues(filStoCat6);
}

function onSubmit(){
  var i = storage.getRange("B5").getValue();
  if (forms.getRange("A"+ i).getValue() != ""){
    var trDate = forms.getRange("A" + i).getValue();
    var trAcc = forms.getRange("B"+ i).getValue();
    var trPay = forms.getRange("C"+ i).getValue();
    var trCat = forms.getRange("D"+ i).getValue();
    if (forms.getRange("E"+ i).getValue() != ""){
      var trSub = forms.getRange("E"+ i).getValue();
    }
    if (forms.getRange("I"+ i).getValue() != ""){
      var trSub = forms.getRange("I"+ i).getValue();
    }
    if (forms.getRange("J"+ i).getValue() != ""){
      var trSub = forms.getRange("J"+ i).getValue();
    }
    if (forms.getRange("K"+ i).getValue() != ""){
      var trSub = forms.getRange("K"+ i).getValue();
    }
    if (forms.getRange("L"+ i).getValue() != ""){
      var trSub = forms.getRange("L"+ i).getValue();
    }
    if (forms.getRange("M"+ i).getValue() !== ""){
      var trSub = forms.getRange("M"+ i).getValue();
    }
    if (forms.getRange("D"+i).getValue() == "Transfer"){
      var trCat = "Transfer";
      var trSub = "Subcategory not needed";
      var trMemo = "Transfer";
    }
    if (forms.getRange("D"+i).getValue() == "To Be Budgeted"){
      var trCat = "To Be Budgeted";
      var trSub = "Subcategory not needed";
    }
    var trMemo = forms.getRange("F"+ i).getValue();
    var trInf = forms.getRange("H"+ i).getValue();
    var trOutf = forms.getRange("G"+ i).getValue();

    allAccounts.getRange("E4").setValue(trAcc);
    temp = new Date(trDate);
    temp = Utilities.formatDate(temp, "GMT", "yyyy-MM-dd");                                                          //variable for current date
    allAccounts.getRange('F4').setValue(temp)
    allAccounts.getRange("G4").setValue(trPay);
    allAccounts.getRange("H4").setValue(trCat);
    allAccounts.getRange("I4").setValue(trSub);
    allAccounts.getRange("H4").setValue(trCat);
    allAccounts.getRange("J4").setValue(trMemo);
    allAccounts.getRange("K4").setValue(trOutf);
    allAccounts.getRange("L4").setValue(trInf);

    storage.getRange("B5").setValue(i+1);
  }else{ui.alert("No more recent transactions")}
}