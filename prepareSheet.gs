function onOpen(e) {
  var html = HtmlService(); 
  var ui = SpreadsheetApp.getUi()
      .createMenu('Billing')
      .addItem('Set Invoice Number', 'setNumber')
      .addItem('Prepare Sheet', 'menuItem1')
      .addItem('Import Transaction Data', 'menuItem2')
      .addItem('Begin Approvals', 'menuItem3')
      .addToUi();
}
function getInvoiceNumber(invoicenumber){
  var beginningnumber = invoicenumber;
  var numberrange = SpreadsheetApp.getActive().getSheetByName('Invoice Numbers').getRange(2,2,1,1);
  numberrange.setValue(beginningnumber);
}

function setNumber(){
  var modal = HtmlService.createHtmlOutputFromFile('getInvoiceNumber')
      .setWidth(350)
      .setHeight(85);
  var dialog = ui.showModalDialog(modal, 'Enter Starting Invoice Number');
}

// Add row banding for visibility,four columns to recieve purchase and return data, and validation column in last column for clearing function
function menuItem1() {
  var active = SpreadsheetApp.getActive();
  var sheet = active.getActiveSheet();
  var lastrow = sheet.getLastRow();
  var stylehelper = ('A1:N' + lastrow)
  var stylerange = active.getRange(stylehelper);
  var approvalhelper = ('N2:N' + lastrow);
  var approvalrange = active.getRange(approvalhelper);
  var approvalvalidation = SpreadsheetApp.newDataValidation().requireValueInList(['Approved', 'Rejected'], true).build();
    stylerange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
    approvalrange.setDataValidation(approvalvalidation);
// Add column headers for purchase and return data and add column header for approvals
  var volumeheaders = sheet.getRange(1,8,1,7);
  var headervalues = [["Invoice Number", "Purchase Count", "Purchase Amount", "Return Count", "Return Amount", "Review","Approvals"]];
    volumeheaders.setValues(headervalues).setFontWeight('Bold');
// Format Client Number column to six digits with leading zeroes
  var acctnumhelper = ('A2:A' + lastrow);
  var acctnumcolumn = sheet.getRange(acctnumhelper);
    acctnumcolumn.setNumberFormat("000000");
// Formats the column with Next Billing Cycle as text    
  var billcyclehelper = ('F2:F'+lastrow);
  var billcyclecolumn = sheet.getRange(billcyclehelper);
    billcyclecolumn.setNumberFormat('@STRING@');
// Assigns invoice number for each line item
  var invoicenumbertab = active.getSheetByName('Invoice Numbers');
  var nextinvoicenumber = invoicenumbertab.getRange(2,1,1,1).getValue();
  var newinvoicenumber = sheet.getRange(lastrow,8,1,1).getValue();
    sheet.getRange(2,8,1,1).setValue(nextinvoicenumber);
// Creates Unique ID for each line item to merge Sales and Returns data
  var lineitemdata = sheet.getDataRange();
  var lastcolumn = lineitemdata.getLastColumn();
  var searchrange = sheet.getRange(2,1,lastrow, lastcolumn);
  var sourcetab = active.getSheetByName('Invoice Numbers');
  var targettab = active.getSheetByName('Clearing Sheet');
  var targettabdata = targettab.getDataRange();
  var invoicenumlastrow = targettab.getLastRow();
  var invoicenumlastcolumn = targettab.getLastColumn();
  var nextnumber = sourcetab.getRange(2,2,1,1).getValue();
  var newlastnumber = targettab.getRange(targettab.getLastRow(),8,1,1).getValue();
  var workingformula = ('=IF(R[-1]C[-6] <> RC[-6], R[-1]C+1, R[-1]C)');
      targettab.getRange(2,8,1,1).setValue(nextnumber)
            for ( var i = 0; i < invoicenumlastrow-2; i++){
              targettab.getRange(i+3,invoicenumlastcolumn-6,1,1).setValue(workingformula);
  }            
      logTransactionMonth();        
}

function itemApproval() {
  var source = SpreadsheetApp.getActive();
  var sourcesheet = source.getSheetByName('Clearing Sheet');
  var approvalsheet = source.getSheetByName('Approved Line Items');
  var approvalrange = approvalsheet.getRange(approvalsheet.getLastRow()+1,1,1,13);
  var rejectsheet = source.getSheetByName('Rejected Line Items');
  var rejectrange = rejectsheet.getRange(rejectsheet.getLastRow()+1, 1, 1, 13);
  var values = sourcesheet.getRange(source.getSelection().getCurrentCell().getRow(),1,1,13).getValues();
  var reviewcheck = sourcesheet.getRange(source.getSelection().getCurrentCell().getRow(),13,1,1).getValue();
  var statuscheck = sourcesheet.getRange(source.getSelection().getCurrentCell().getRow(),14,1,1).getValue();
  var activerow = sourcesheet.getSelection().getCurrentCell().getRow();
  var approvedrange = sourcesheet.getRange(activerow,1,1,13).getA1Notation();
  var approvedvalues = sourcesheet.getRange(activerow,1,1,13).getValues();      
      if (statuscheck == "Approved"){
          var docid = retrieveDocId();
          var targetSheet = SpreadsheetApp.openById(docid);
          var targetTab = targetSheet.getSheetByName('Purchase Data');
          var lastRow = targetTab.getLastRow();
          var targetRow = lastRow+1;
          var targetRange = targetTab.getRange(targetRow,1,1,13);
          var url = "https://docs.google.com/spreadsheets/d/"+docid;
          var html = "<script>window.open('" + url + "');google.script.host.close();</script>";
          var userInterface = HtmlService.createHtmlOutput(html);
            targetTab.getRange(approvedrange).setValues(values);    
            approvalrange.setValues(approvedvalues);
            SpreadsheetApp.getUi().showModalDialog(userInterface, "Opening Sheet");
       //Reject condition moves the row to a rejected line items page where a new process begins
       if (statuscheck == "Rejected") {
        var rejectvalues =sourcesheet.getRange(sourcesheet.getSelection().getCurrentCell().getRow(),1,1,12).getValues();
        rejectrange.setValues(rejectvalues);
        }
   }
}
function menuItem2() {
  var modal = HtmlService.createHtmlOutputFromFile('getPurchaseData')
      .setWidth(400)
      .setHeight(115);    
  SpreadsheetApp.getUi().showModalDialog(modal, 'Purchase and Returns Month');
}  
function cleanRows(){
var workingSheet = SpreadsheetApp.getActive();
var clearingSheet = workingSheet.getSheetByName('Clearing Sheet');
var clearingTest = clearingSheet.getRange(clearingSheet.getSelection().getCurrentCell().getRow(),14,1,1).getValue();
var clearingTest2 = clearingSheet.getRange(clearingSheet.getSelection().getCurrentCell().getRow(),3,1,1).getValue();
var deleteRow = clearingSheet.getSelection().getCurrentCell().getRow();
var rejectionSheet = workingSheet.getSheetByName('Rejected Line Items');
var rejectionTest = rejectionSheet.getRange(rejectionSheet.getLastRow(),3,1,1).getValue();
var approvalSheet = workingSheet.getSheetByName('Approved Line Items');
var approvalTest = approvalSheet.getRange(approvalSheet.getLastRow(),3,1,1).getValue();
  if (clearingTest == 'Approved' && clearingTest2 == approvalTest){
  clearingSheet.deleteRow(deleteRow);
  }
  else if (clearingTest == 'Rejected' && clearingTest2 == rejectionTest){
  clearingSheet.deleteRow(deleteRow);
  }
}
function logTransactionMonth(){
  var transactionmonth = SpreadsheetApp.getActive().getSheetByName('Invoice Numbers').getRange(3,2,1,1).getValue();
  var monthpurchases = (transactionmonth + ' Purchases');
  var transactionssheet = SpreadsheetApp.openById('1n0oFePjP3SGpE9fK2j_IptZ93RySbGbZB12T7wz9Bhg')
  var salestab = transactionssheet.getSheetByName(monthpurchases);
  var salesrange = salestab.getDataRange();
  var salesrangeA1 = salesrange.getA1Notation();
  var salesdata = salesrange.getValues()
  var targetsheet = SpreadsheetApp.getActive();
  var salestargettab = targetsheet.getSheetByName('SalesData');
  var monthreturns = (transactionmonth + ' returns');
  var returnstab = transactionssheet.getSheetByName(monthreturns);
  var returnsrange = returnstab.getDataRange();
  var returnsrangeA1 = returnsrange.getA1Notation();
  var returnsdata = returnsrange.getValues()
  var returnstargettab = targetsheet.getSheetByName('ReturnsData');    
      salestargettab.clear({contentsOnly: true});
      salestargettab.getRange(salesrangeA1).setValues(salesdata);
      returnstargettab.clear({contentsOnly: true});
      returnstargettab.getRange(returnsrangeA1).setValues(returnsdata);
  var salessheetlastrow = salestargettab.getLastRow();
  var returnssheetlastrow = returnstargettab.getLastRow();
  var salesclientrange = salestargettab.getRange('A2:A' + salessheetlastrow);    
  var returnsclientrange = returnstargettab.getRange('A2:A' + salessheetlastrow);
      salesclientrange.setNumberFormat("000000");
      returnsclientrange.setNumberFormat("000000");
      mergeTransactionData();
}

function mergeTransactionData(){
  var itemsheet = SpreadsheetApp.getActive();
  var salestab = itemsheet.getSheetByName('SalesData');
  var returnstab = itemsheet.getSheetByName('ReturnsData');
  var salesdata = salestab.getDataRange();
  var returnsdata = returnstab.getDataRange();
  var saleslastrow = salesdata.getLastRow();
  var saleslastcolumn = salesdata.getLastColumn();
  var salessearchrange = salestab.getRange(1,1,saleslastrow, saleslastcolumn);
  var returnslastrow = returnsdata.getLastRow();
  var returnslastcolumn = returnsdata.getLastColumn();
  var returnssearchrange = returnstab.getRange(1,1,returnslastrow, returnslastcolumn);   
  var clearingtab = itemsheet.getSheetByName('Clearing Sheet');
  var clearingdata = clearingtab.getDataRange();
  var clearinglastrow = clearingdata.getLastRow();
  var clearinglastcolumn = clearingdata.getLastColumn();
      for ( var i = 0; i < clearinglastrow-1; i++){    
          var salescount = ("=iferror(vlookup(RC[-6],SalesData!B:G,3,FALSE),0)");
          var salesamount = ("=iferror(vlookup(RC[-7],SalesData!B:G,4,FALSE),0)");
          var returnscount = ("=iferror(vlookup(RC[-8],ReturnsData!B:F,3,FALSE),0)");
          var returnsamount = ("=iferror(vlookup(RC[-9],ReturnsData!B:F,4,FALSE),0)");
          var reviewflag = ("=iferror(vlookup(RC[-10],SalesData!B:G,6,FALSE),0)");
              clearingtab.getRange(i+2,13,1,1).setValue(reviewflag);
              clearingtab.getRange(i+2,9,1,1).setValue(salescount);
              clearingtab.getRange(i+2,10,1,1).setValue(salesamount);
              clearingtab.getRange(i+2,11,1,1).setValue(returnscount);
              clearingtab.getRange(i+2,12,1,1).setValue(returnsamount);
  }   
  
  var destinvoicenumber = itemsheet.getSheetByName('Invoice Numbers');
  var destlastnumber = destinvoicenumber.getRange(1,2,1,1);
  var lastinvoicenumber = clearingtab.getRange(clearingtab.getLastRow(),8,1,1).getValue();
    destlastnumber.setValue(lastinvoicenumber);
    menuItem3();
}

function menuItem3(){
  var finalsheet = SpreadsheetApp.getActive();
  var finaltab = finalsheet.getSheetByName('Clearing Sheet');
  var finaldata = finaltab.getDataRange();
  var finalvalues = finaldata.getValues();
    finaldata.copyTo(finaltab.getRange(1,1,finaltab.getLastRow(),finaltab.getLastColumn()), {contentsOnly: true});
}

// Calls name of company being edited on the Clearing Sheet
function ActiveCompanyNumber() {
    var sheet = SpreadsheetApp.getActive().getActiveSheet();
    var activerow = sheet.getRange(sheet.getSelection().getCurrentCell().getRow(),1,1,1);
    var companynumber = activerow.getValue();
    //Logger.log('Company name is ' + companyname);
    return companynumber;  
} 

// Calls the row number of the Master List that contains the company name from above
 function rowOfDocId(){ 
  var sourcesheet = SpreadsheetApp.openById('1WQBEVDTyK8XvTG5BkMJMbqWMyKTf3aYuFjCQPuc23GI').getSheetByName('Client Master List');
  var sourcevalues = sourcesheet.getDataRange().getValues();
  var companynumber = ActiveCompanyNumber();
   
     for (var i = 0; i < sourcevalues.length;i++){
     for (var j = 0; j < sourcevalues[i].length; j++){
        if(sourcevalues[i][j] == companynumber){

    return i+1;
      }
    }
  }
} 
  
// Retrieves the ID of the invoicing document related to the active company from above and sends it to onEdit function on prepareSheet.gs  
function retrieveDocId(){ 
  var masterrownum = rowOfDocId();
  var mastersheet = SpreadsheetApp.openById('1WQBEVDTyK8XvTG5BkMJMbqWMyKTf3aYuFjCQPuc23GI');
  var mastertab = mastersheet.getSheetByName('Client Master List');
  var masterrange = mastertab.getRange(masterrownum, 6,1,1);
  var billingDocId = masterrange.getValue();
  var routingId = billingDocId;
  return routingId;
}