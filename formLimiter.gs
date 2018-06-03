
function onOpen() {
  var ui=SpreadsheetApp.getUi();
  ui.createMenu('Restrictions')
  .addItem('Restrict responses','showPromptResponse')
  .addItem('Restrict date', 'showPromptDate')
  .addToUi();
}

function showPromptDate() {
  var ui=SpreadsheetApp.getUi();
  var html = HtmlService.createHtmlOutputFromFile('dateLimitBox')
  .setWidth(200);
  ui.showModalDialog(html, 'Restrictions'); 
}

function showPromptResponse() {
  var ui=SpreadsheetApp.getUi();
  var html = HtmlService.createHtmlOutputFromFile('responseLimitBox')
  .setHeight(100)
  .setWidth(300);
  ui.showModalDialog(html, 'Restrictions');  
}

function getDateLimit(date, time){
  var maxDate=date+ "," + time;
  var sheet = SpreadsheetApp.getActiveSheet();
  var lrow = sheet.getLastRow();
  var lcol = sheet.getLastColumn();
  var lastcolumn=lcol;
  var delimeter = '=';
  var limitVar = sheet.getRange(1, lastcolumn).getValue();
  var responseLimit = limitVar.split(delimeter);
  var delimeter2=','
  if(responseLimit[0]=='Date Limit'){
    lastcolumn=lcol;
  }
  else{
    lastcolumn=lcol+1;
  }
  SpreadsheetApp.getActiveSheet().getRange(1,lastcolumn).setValue("Date Limit="+ maxDate);
}

function formLimiterDate() { 
  var sheet = SpreadsheetApp.getActiveSheet();
  var lrow = sheet.getLastRow();
  var lcol = sheet.getLastColumn();
  var lastcolumn=lcol;
  var delimeter = '=';
  var limitVar = sheet.getRange(1, lastcolumn).getValue();
  var responseLimit = limitVar.split(delimeter);
  var delimeter2=','
  if(responseLimit[0]=='Date Limit'){
    var viewLimit = responseLimit[1].split(delimeter2);
    var date = viewLimit[0];
    var time = viewLimit[1];
    var currentdate = new Date(); 
    var datetime = (currentdate.getMonth()+1)  + "/" 
                + currentdate.getDate() + "/"
                + currentdate.getFullYear() + " "  
                + currentdate.getHours() + ":"  
                + currentdate.getMinutes() + ":" 
                + currentdate.getSeconds();
                

   var isLarger = new Date(datetime) < new Date(date + " " + time);
   var formId= (FormApp.openByUrl(SpreadsheetApp.getActiveSpreadsheet().getFormUrl())).getId() ;
   var form = FormApp.openById(formId);
   if(isLarger==false){
     form.setAcceptingResponses(false);
    }
  }
}

function getViewLimit(form){
  var maxLimit=form;
  var sheet = SpreadsheetApp.getActiveSheet(); 
  var lrow = sheet.getLastRow();
  var lcol = sheet.getLastColumn();
  var lastcolumn=lcol;
  var delimeter = '=';
  var limitVar = sheet.getRange(1, lastcolumn).getValue();
  var responseLimit = limitVar.split(delimeter);  
  if(responseLimit[0]=='Response Limit '){
    lastcolumn=lcol;
    SpreadsheetApp.getActiveSheet().getRange(1,lastcolumn).setValue("Response Limit = "+maxLimit);
  }
  else if(responseLimit[0]=='Date Limit'){
    var prevColumn = lastcolumn-1;
    var prevLastColumn =sheet.getRange(1, prevColumn).getValue();
    var prevLimit=prevLastColumn.split(delimeter);
    var datcol;
    if(prevLimit[0] == 'Response Limit '){
      datcol = lastcolumn;
      lastcolumn=prevColumn;
    }
    else{
      lastcolumn=lcol;
      datcol=lastcolumn+1;
    }
    SpreadsheetApp.getActiveSheet().getRange(1,datcol).setValue(limitVar);
    SpreadsheetApp.getActiveSheet().getRange(1,lastcolumn).setValue("Response Limit = "+maxLimit);
    if(lrow>=2){
      var startRow = 2;  
      var dataRange = sheet.getRange(startRow, 1, lrow-1, lastcolumn);
      var data = dataRange.getValues();
      var data = dataRange.getValues();
      for (var i = 0; i < data.length; ++i) {
        var row = data[i];
        sheet.getRange(startRow + i, lastcolumn).setValue(i+1);
      }
    }
    
  }
  else{
    lastcolumn=lcol+1;
    SpreadsheetApp.getActiveSheet().getRange(1,lastcolumn).setValue("Response Limit = "+maxLimit);
    if(lrow>=2){
    var startRow = 2;  
    var dataRange = sheet.getRange(startRow, 1, lrow-1, lastcolumn);
    var data = dataRange.getValues();
    var data = dataRange.getValues();
    for (var i = 0; i < data.length; ++i) {
      var row = data[i];
      sheet.getRange(startRow + i, lastcolumn).setValue(i+1);
     }
   }
 }
}

function formLimiterViews(){
  var sheet = SpreadsheetApp.getActiveSheet();
  var lrow = sheet.getLastRow();
  var lcol = sheet.getLastColumn();
  var lastcolumn=lcol;
  var startRow = 2;  
  var dataRange = sheet.getRange(startRow, 1, lrow-1, lastcolumn);
  var data = dataRange.getValues();
  var firstView = 1;
  var viewNumPrev;
  var viewNumCurr;
  var delimeter = '=';
  var getColVal = sheet.getRange(1, lastcolumn).getValue();
  var delVal= getColVal.split(delimeter);
  if(delVal[0]=='Date Limit'){
    lastcolumn=lastcolumn-1;
    if(sheet.getRange(2, lastcolumn).getValue()==''){
      sheet.getRange(2, lastcolumn).setValue(firstView);
    }
    else{
      viewNumPrev= sheet.getRange(lrow-1, lastcolumn).getValue();
      viewNumCurr = Number(viewNumPrev) + 1;
      sheet.getRange(lrow, lastcolumn).setValue(viewNumCurr);
      var limitVar = sheet.getRange(1, lastcolumn).getValue();
      var responseLimit = limitVar.split(delimeter);
      var formId= (FormApp.openByUrl(SpreadsheetApp.getActiveSpreadsheet().getFormUrl())).getId() ;
      var form = FormApp.openById(formId);
      if(viewNumCurr>=Number(responseLimit[1])){
        form.setAcceptingResponses(false);    
      }
    }   
  }
  else if(delVal=='Response Limit '){
    lastcolumn=lastcolumn;
    if(sheet.getRange(2, lastcolumn).getValue()==''){
      sheet.getRange(2, lastcolumn).setValue(firstView);
    }
    else{
      viewNumPrev= sheet.getRange(lrow-1, lastcolumn).getValue();
      viewNumCurr = Number(viewNumPrev) + 1;
      sheet.getRange(lrow, lastcolumn).setValue(viewNumCurr);
      var limitVar = sheet.getRange(1, lastcolumn).getValue();
      var responseLimit = limitVar.split(delimeter);
      var formId= (FormApp.openByUrl(SpreadsheetApp.getActiveSpreadsheet().getFormUrl())).getId() ;
      var form = FormApp.openById(formId);
      if(viewNumCurr>=Number(responseLimit[1])){
        form.setAcceptingResponses(false);
      }
    } 
  }
}
