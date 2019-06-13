function onOpen(e) {
//Add custom menu to the active sheet
  var menu = SpreadsheetApp.getUi();
  var modal = HtmlService.createHtmlOutputFromFile('getInvoiceNumber')
      .setWidth(350)
      .setHeight(85);
  var dialog = menu.showModalDialog(modal, 'Enter Starting Invoice Number');    
      menu.createMenu('Billing')
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
  var volumeheaders = sheet.getRange(1,7,1,8);
  var headervalues = [[ "Invoice Number", "Unique ID", "Purchase Count", "Purchase Amount", "Return Count", "Return Amount", "Review","Approvals"]];
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
  var nextinvoicenumber = invoicenumbertab.getRange(2,2,1,1).getValue();
  var newinvoicenumber = sheet.getRange(lastrow,7,1,1).getValue();
    sheet.getRange(2,7,1,1).setValue(nextinvoicenumber);
// Creates Unique ID for each line item to merge Sales and Returns data
  var lineitemdata = sheet.getDataRange();
  var lastcolumn = lineitemdata.getLastColumn();
  var searchrange = sheet.getRange(2,1,lastrow, lastcolumn);
  var lineitemvalues = searchrange.getValues();
      for ( var i = 0; i < lastrow-1; i++){
      for ( var j = 0 ; j < lastcolumn; j++){
          var uniqueid = ("=concatenate(RC[-7], RC[-5], RC[-3])");
              sheet.getRange(i+2,lastcolumn-5,1,1).setValue(uniqueid);
    }   
  }
  var sourcetab = active.getSheetByName('Invoice Numbers');
  var targettab = active.getSheetByName('Clearing Sheet');
  var targettabdata = targettab.getDataRange();
  var invoicenumlastrow = targettab.getLastRow();
  var invoicenumlastcolumn = targettab.getLastColumn();
  var nextnumber = sourcetab.getRange(2,2,1,1).getValue();
  var newlastnumber = targettab.getRange(targettab.getLastRow(),7,1,1).getValue();
  var workingformula = ('=IF(R[-1]C[-6] <> RC[-6], R[-1]C+1, R[-1]C)');
      targettab.getRange(2,7,1,1).setValue(nextnumber)
            for ( var i = 0; i < invoicenumlastrow-2; i++){
              targettab.getRange(i+3,invoicenumlastcolumn-6,1,1).setValue(workingformula);
  }            
      logTransactionMonth();        
}

function onEdit(e) {
  var docid = retrieveDocId();
  var targetSheet = SpreadsheetApp.openById(docid);
  var targetTab = targetSheet.getSheetByName('Purchase Data');
  var lastRow = targetTab.getLastRow();
  var targetRow = lastRow+1;
  var targetRange = targetTab.getRange(targetRow,1,1,13);
  var deleteCount = 0;
  var source = SpreadsheetApp.getActive();
  var sourcesheet = source.getSheetByName('Clearing Sheet');
  var approvalsheet = source.getSheetByName('Approved Line Items');
  var approvalrange = approvalsheet.getRange(approvalsheet.getLastRow()+1,1,1,13);
  var rejectsheet = source.getSheetByName('Rejected Line Items');
  var rejectrange = rejectsheet.getRange(rejectsheet.getLastRow()+1, 1, 1, 13);
  var values = sourcesheet.getRange(source.getSelection().getCurrentCell().getRow(),1,1,13).getValues();
  var reviewcheck = sourcesheet.getRange(source.getSelection().getCurrentCell().getRow(),13,1,1).getValue();
  var statuscheck = sourcesheet.getSelection().getCurrentCell().getValue();
  var activerow = sourcesheet.getActiveCell.getRow();
      // Approved condition moves active row to corresponding invoicinging sheet. IMPORTANT A1Notation moves row to corresponding row of target sheet
      if (statuscheck == "Approved" && reviewcheck == "false") {
        deleteCount++;
        var approvedrange = sourcesheet.getRange(activerow,1,1,13).getA1Notation();
        targetTab.getRange(approvedrange).setValues(values);
        var approvedvalues = sourcesheet.getRange(activerow,1,1,13).getValues();
        approvalrange.setValues(approvedvalues);

        var ui = SpreadsheetApp.getUi();
        var dialog = ui.alert('Click OK to open the invoice file', ui.ButtonSet.YES_NO);
          if (dialog == ui.Button.YES){
              var url = "https://docs.google.com/spreadsheets/d/"+docid;
              var html = "<script>window.open('" + url + "');google.script.host.close();</script>";
              var userInterface = HtmlService.createHtmlOutput(html);
                     SpreadsheetApp.getUi().showModalDialog(userInterface, "Opening Sheet");
            }
            else {
            ui.alert("Don't forget to generate the invoice.");
            }
    
       if (statuscheck == "Approved" && reviewcheck == "true") {
        deleteCount++;
        var approvedrange = sourcesheet.getRange(sourcesheet.getSelection().getCurrentCell().getRow(),1,1,13).getA1Notation();
        targetTab.getRange(approvedrange).setValues(values);
        var approvedvalues = sourcesheet.getRange(sourcesheet.getSelection().getCurrentCell().getRow(),1,1,13).getValues();
        approvalrange.setValues(approvedvalues);
        //Logger.log('Called from approve if statement');
         /*if (deleteCount < 2) {sourcesheet.deleteRow(sourcesheet.getSelection().getCurrentCell().getRow());}
        Logger.log('Row number is ' + sourcesheet.getSelection().getCurrentCell().getRow());*/
        var ui = SpreadsheetApp.getUi();
        var dialog = ui.alert('Flagged for review.','After creating this invoice it must first be sent to Tom and Sebastian for approval. Click OK to open the invoice document.', ui.ButtonSet.YES_NO);
          if (dialog == ui.Button.YES){
              var url = "https://docs.google.com/spreadsheets/d/"+docid;
              var html = "<script>window.open('" + url + "');google.script.host.close();</script>";
              var userInterface = HtmlService.createHtmlOutput(html);
                     SpreadsheetApp.getUi().showModalDialog(userInterface, "Opening Sheet");
            }
            else {
            ui.alert("Don't forget to generate the invoice.");
            }
    }
       //Reject condition moves the row to a rejected line items page where a new process begins
       if (statuscheck == "Rejected") {
        var rejectvalues =sourcesheet.getRange(sourcesheet.getSelection().getCurrentCell().getRow(),1,1,12).getValues();
        rejectrange.setValues(rejectvalues);
        //sourcesheet.deleteRow(sourcesheet.getSelection().getCurrentCell().getRow());
        }
    }
}
function menuItem2() {
  var modal = HtmlService.createHtmlOutputFromFile('getPurchaseData')
      .setWidth(400)
      .setHeight(115);    
  SpreadsheetApp.getUi().showModalDialog(modal, 'Purchase and Returns Month');
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
      salestab.insertColumnBefore(4);
      returnstab.insertColumnBefore(4);
      for ( var i = 0; i < saleslastrow; i++){
          var salesuniqueid = ("=concatenate(RC[-3], RC[-2], RC[3])");
              salestab.getRange(i+2,saleslastcolumn-2,1,1).setValue(salesuniqueid);  
  }
      for ( var i = 0; i < returnslastrow; i++){
          var returnsuniqueid = ("=concatenate(RC[-3], RC[-2], RC[3])");
              returnstab.getRange(i+2,returnslastcolumn-2,1,1).setValue(returnsuniqueid);  
  }
  var clearingtab = itemsheet.getSheetByName('Clearing Sheet');
  var clearingdata = clearingtab.getDataRange();
  var clearinglastrow = clearingdata.getLastRow();
  var clearinglastcolumn = clearingdata.getLastColumn();
      for ( var i = 0; i < clearinglastrow-1; i++){    
          var salescount = ("=iferror(vlookup(RC[-1],SalesData!D:F,2,FALSE),0)");
          var salesamount = ("=iferror(vlookup(RC[-2],SalesData!D:F,3,FALSE),0)");
          var returnscount = ("=iferror(vlookup(RC[-3],ReturnsData!D:F,2,FALSE),0)");
          var returnsamount = ("=iferror(vlookup(RC[-4],ReturnsData!D:F,3,FALSE),0)");
          var reviewflag = ("=iferror(vlookup(RC[-5],SalesData!D:G,5,FALSE),0)");
              clearingtab.getRange(i+2,clearinglastcolumn-5,1,1).setValue(reviewflag);
              clearingtab.getRange(i+2,clearinglastcolumn-4,1,1).setValue(salescount);
              clearingtab.getRange(i+2,clearinglastcolumn-3,1,1).setValue(salesamount);
              clearingtab.getRange(i+2,clearinglastcolumn-2,1,1).setValue(returnscount);
              clearingtab.getRange(i+2,clearinglastcolumn-1,1,1).setValue(returnsamount);
  }   
  
  var destinvoicenumber = itemsheet.getSheetByName('Invoice Numbers');
  var destlastnumber = destinvoicenumber.getRange(1,2,1,1);
  var lastinvoicenumber = clearingtab.getRange(clearingtab.getLastRow(),7,1,1).getValue();
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
Logger.log(i+1);
} 
  
// Retrieves the ID of the invoicing document related to the active company from above and sends it to onEdit function on prepareSheet.gs  
function retrieveDocId(){ 
  var masterrownum = rowOfDocId();
  var mastersheet = SpreadsheetApp.openById('1WQBEVDTyK8XvTG5BkMJMbqWMyKTf3aYuFjCQPuc23GI');
  var mastertab = mastersheet.getSheetByName('Client Master List');
  var masterrange = mastertab.getRange(masterrownum, 6,1,1);
  var billingDocId = masterrange.getValue();
  var routingId = billingDocId;
  //Logger.log('billingDocId is ' + routingId);
  return routingId;
}