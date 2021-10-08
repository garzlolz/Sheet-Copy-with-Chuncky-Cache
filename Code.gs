/*function Disabled_onEdit(e) {

  var range = e.range;
  var Column = range.getColumn();
  var source = e.source;
  var Sheet_Name = source.getSheetName();
  var value = e.value;
  
  if ( Sheet_Name == 'Vendor' && Column == 16 && value == '999') 
  {
    var row = range.getRow();
    var Lookup_Value = source.getSheetByName(Sheet_Name).getRange( row, 23 ).getValue(); // Column = Rim Type
    source.getSheetByName('VPO Lookup').getRange("A2").setValue(Lookup_Value);
    range.clearContent();
    range.activate();
  }
  
}
*/

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Biggio Menu')
      .addItem('1. Update PI', 'run_pi_update')  
      .addItem('2. Force Reset', 'reset')
      .addSeparator()
      .addItem('CSV Upload & Send Notification', 'CSV_import_And_Notice')
      .addItem('CSV Upload', 'run_CSV_import')  
      .addSeparator()
      .addItem('臨時更新用，帶入新單', 'TenMinsCheck')  
      .addSeparator()
      .addItem('分配','distribute_Alert')
      
      .addToUi();
}

function CSV_import_And_Notice()
{
  run_CSV_import();
  sendNotifEmail();
}

function UpdateFormat() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Vendor');
  var Id = ss.getId();
  var SheetId = sheet.getSheetId();
  
  sheet.getRange(2,1,sheet.getMaxRows()-1, sheet.getMaxColumns() ).setFontColor(null).setBackground(null);
  sheet.getRange("P:R").setBackground("yellow");
  sheet.getRange("L:M").setNumberFormat("mm/dd/yyyy");
  sheet.getRange("P:P").setNumberFormat("mm/dd/yyyy");  
  sheet.getRange("AD:AD").setNumberFormat("mm/dd/yyyy");  
  //SpreadsheetApp.flush()
  var Formula = '=AND(D:D<>"",M:M="",P:P="")';
  
  //Logger.log(Id);
  
  var requests = [
    {
      "deleteConditionalFormatRule" : {
        "index": 0,
        "sheetId": SheetId,
      },
    },
    {
      "deleteConditionalFormatRule" : {
        "index": 0,
        "sheetId": SheetId,
      }
    }
  ];
  
  //Sheets.Spreadsheets.batchUpdate({'requests': requests}, Id);
  try { 
    Sheets.Spreadsheets.batchUpdate({'requests': requests}, Id);
  } catch(e) {
    //Catch any error here. Example below is just sending an email with the error.
    Logger.log("Error, No Format");
  }
  
  
  
  var requests =  [
    {
      "addConditionalFormatRule": {
        "rule": {
          "ranges": {
            "sheetId": SheetId,
            //    "startRowIndex": number,
            //    "endRowIndex": number,
            "startColumnIndex": 3,
            "endColumnIndex": 4,
          },
          "booleanRule": {
            "condition": {
              "type": "CUSTOM_FORMULA",
              "values": [
                {
                  "userEnteredValue": '=AND(D:D<>"",M:M="",P:P="")'
                  
                }
              ]
            },
            "format": {
              "backgroundColor": {
                "red": 1,
                "green": 0,
                "blue": 0,
                "alpha": 0,
              },
              
              "textFormat": {
                "foregroundColor": {
                  "red": 1,
                  "green": 1,
                  "blue": 1,
                  "alpha": 0,
                },
                "bold": false,
                "italic": true
              }
            }
          }
        },
        "index": 0
      }, },
    {
      "addConditionalFormatRule": {   
        "rule": {
          "ranges": {
            "sheetId": SheetId,
            //    "startRowIndex": number,
            //    "endRowIndex": number,
            "startColumnIndex": 19,
            "endColumnIndex":20,
          },
          "booleanRule": {
            "condition": {
              "type": "CUSTOM_FORMULA",
              "values": [
                {
                  "userEnteredValue": '=AND(D:D<>"",M:M="",P:P="",S:S="",T:T>0,V:V="")'
                }
              ]
            },
            "format": {
              "backgroundColor": {
                "red": 0,
                "green": 1,
                "blue": 0,
                "alpha": 0,
              },
              
              "textFormat": {
                "foregroundColor": {
                  "red": 0,
                  "green": 0,
                  "blue": 0,
                  "alpha": 0,
                },
                "bold": false,
                "italic": true
              }
            }
          }
        },
        "index": 0
      }
    }
  ];
  
  var response = Sheets.Spreadsheets.batchUpdate({'requests': requests}, Id);
  //Logger.log(response);
  
  
}

function run_pi_update() {
  CopyRange2('1Tnvgth_BtyR_tzRQv4eVFu06FVvoTF0aRcFnrcTV1vI','get_PI_list','A:Z','PI_List', 1, 1, 1 )
}

function reset() {
  unlock('Vendor');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = ss.getSheetByName("Config");
  sheet1.getRange("G2").setValue("1");
  //SpreadsheetApp.flush()
  //var plock = LockService.getScriptLock();
  //var success = plock.tryLock(100000);
  //if (!success) {
  //  Logger.log('Could not obtain lock after 10 seconds.');
  //}
  //var ss0 = SpreadsheetApp.openById("1Tnvgth_BtyR_tzRQv4eVFu06FVvoTF0aRcFnrcTV1vI"); // VPO CDD Source
  //var sheet0 = ss0.getSheetByName("Config");
  //var range0 = sheet0.getRange("B2");
  //range0.setValue("1").setNumberFormat("@");
  
  
  //SpreadsheetApp.flush()
  UpdateNewVPO3(0,24);
  Logger.log('3');  
  UpdateFormat();
  Logger.log('1');
  CopyRange2('1hUlRAJbgj3DxEuNd0C_p4KrJwQ4LzZ2N553dh-bK8K4','Work_SCM-1', 'A:C', '報表A區',1,1,1);
  CopyRange2("1hUlRAJbgj3DxEuNd0C_p4KrJwQ4LzZ2N553dh-bK8K4","Work_SCM", "A:I", "Work4",1,1,1);
  Logger.log('2');  
  //plock.releaseLock();
  lock('Vendor');
}


function TenMinsCheck() {
  if (!LockService.getScriptLock().hasLock()) {
  
  var nowtime = new Date();
  var nowhour = nowtime.getHours();
  if ( nowhour <= 20 && nowhour >= 9 ) { 
    UpdateNewVPO3(9,20);
    UpdateFormat();
    CopyRange2('1hUlRAJbgj3DxEuNd0C_p4KrJwQ4LzZ2N553dh-bK8K4','Work_SCM-1', 'A:C', '報表A區',1,1,1);
    CopyRange2("1hUlRAJbgj3DxEuNd0C_p4KrJwQ4LzZ2N553dh-bK8K4","Work_SCM", "A:I", "Work4",1,1,1);
  }

  }
}

function test() {
CopyRange2('1Tnvgth_BtyR_tzRQv4eVFu06FVvoTF0aRcFnrcTV1vI','get_PI_list','A:Z','Sheet47', 1 ,1,1)
}



function UpdateNewVPO3(starttime, endtime) {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet0 = ss.getSheetByName("Vendor");
  var sheet1 = ss.getSheetByName("Work1");
  var sheet2 = ss.getSheetByName("Config");
  
  var Reset = sheet2.getRange("G2").getValue();  // Reset field
  //Logger.log(Reset);
  
  if (Reset == "1")
  {
    clearFilter("Vendor");
    sheet0.setFrozenRows(0);
    var maxRow = sheet0.getMaxRows();
    Logger.log(maxRow);
    if(maxRow>1) 
    { 
      //sheet0.getRange(2,1,sheet0.getLastRow()-1,sheet0.getLastColumn()).clearContent();
      sheet0.deleteRows(2,maxRow-1);
    } 
    
    sheet0.insertRowsAfter(1, 1);
    sheet0.getRange("2:" + sheet0.getMaxRows() ).clearFormat();
    
    
    //var range0 = sheet0.getRange(2,1,lastRow, sheet0.getLastColumn());
    //range0.clearContent();
    
    
    sheet2.getRange("G2").setValue("0"); 
    starttime = 0;
    endtime = 24;
    //Utilities.sleep(10000);

  }
  
  if (starttime== undefined) var starttime=0;
  if (endtime== undefined) var endtime=24;
  var nowtime = new Date();
  var nowhour = nowtime.getHours();
  if ( nowhour <= endtime && nowhour >= starttime ) {
    
    //var sheet0 = ss.getSheets()[0];
    //var sheet1 = ss.getSheets()[2];
    
    var lastRow0=sheet0.getLastRow();
    var lastCol0=sheet0.getLastColumn();
    var maxRow = sheet0.getMaxRows();
    
    
    
    if ( maxRow < lastRow0+1 ) 
    {
      sheet0.insertRowsAfter(maxRow, 1);
      //sheet0.getRange("2:" + lastRow0+1).clearFormat();
      Logger.log('Blank Row inserted');
      
    }
    
    
    var lastRow1=sheet1.getLastRow();
    var lastCol1=sheet1.getLastColumn();
    
    
    //var lastRow1= 30;
    //var lastCol1 = 2;
    //Logger.log(lastRow0);
    //Logger.log(lastCol0); 
    //Logger.log(maxRow);
    //Logger.log(lastRow1);
    //Logger.log(lastCol1);
    var A2Value = sheet1.getRange("A2").getValue();
    if (lastRow1>1 && A2Value >0 ) {
      if (lastRow0==2) { lastRow0=1; }
      var range1 = sheet1.getRange(2,1,lastRow1-1, lastCol1);

      range1.copyValuesToRange(sheet0, 1, lastCol1, lastRow0+1, lastRow0+lastRow1-1);
      
    }

      sheet0.setFrozenRows(1);
      sheet0.getRange("2:" + sheet0.getMaxRows() ).clearFormat();
      
      //SpreadsheetApp.flush()
      
      
  }
}
  
function clearFilter(sheet) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  /*
  var ssId = ss.getId();
  
  var sheetId = ss.getSheetByName(sheet).getSheetId();
  var requests = [{
  "clearBasicFilter": {
  "sheetId": sheetId
  }
  }];
  Sheets.Spreadsheets.batchUpdate({'requests': requests}, ssId);
  */

  var filter = ss.getSheetByName(sheet).getFilter();
  if (filter !==null) {
    filter.remove();
  }
  
}


function CopyRange(sheetkey, sheet1, cell_range, sheet2, sheet_row, sheet_column) {
  
  var values = SpreadsheetApp.openById(sheetkey).getSheetByName(sheet1).getRange(cell_range).getValues();
  if ( values.length > 50 ) {
    var ss = SpreadsheetApp.getActive().getSheetByName(sheet2);
    ss.getRange(sheet_row || 1,sheet_column || 1, ss.getMaxRows(), values[0].length ).clear();
    ss.getRange(sheet_row || 1,sheet_column || 1, values.length, values[0].length ).setValues(values);  
  } 
}






function sendNotifEmail() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setActiveSheet(ss.getSheetByName("Notification"));
  var sheet = SpreadsheetApp.getActiveSheet();
  
  var lastRow = sheet.getLastRow();
  
  if ( lastRow ==1 ) { return; }
  
  //var dataRange = sheet.getRange(1, 1, sheet.getLastRow() , 4 ); // A:D
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();
  
  if ( data[0][1] == "" ) { return; }
  
  var today = new Date();
  
  var subject = 'Internal Supply Chain 異動清單 (' + getFormattedDate(today) +')';
  var message = '<p> * 以下為 Internal Supply Chain 異動清單 (' + getFormattedDate(today) +') <p>';
  
  var table = '<table style="border-style:solid; border-collapse:collapse;">';
  
  for (var i=0; i<data.length; i++) {
    var rowData = data[i];
    if ( rowData[0] != "")
    {

      var docDate = getFormattedDate(rowData[0]);
      var DocumentNumber = rowData[1];
      var Customer = rowData[2];
      var PO = rowData[3];
      var Item = rowData[4];
      var Qty = rowData[5];
      var Open = rowData[6];
      var XDD = getFormattedDate(rowData[7]);
      var CDD = getFormattedDate(rowData[8]);
      var Note = rowData[9];
      var CDDChange = getFormattedDate(rowData[10]);
      var NOTEChange = rowData[11];
      var CDDDiff = rowData[12];
      
      table += '<tr><td style="border: 1px solid #999; padding: 0.5rem;">'+ docDate + '</td> \
                    <td style="border: 1px solid #999; padding: 0.5rem;">'+ DocumentNumber + '</td> \
                    <td style="border: 1px solid #999; padding: 0.5rem;">'+ Customer +'</td> \
                    <td style="border: 1px solid #999; padding: 0.5rem;">'+ PO +'</td> \
                    <td style="border: 1px solid #999; padding: 0.5rem;">'+ Item +'</td> \
                    <td style="border: 1px solid #999; padding: 0.5rem;">'+ Qty +'</td> \
                    <td style="border: 1px solid #999; padding: 0.5rem;">'+ Open +'</td> \
                    <td style="border: 1px solid #999; padding: 0.5rem;">'+ XDD +'</td> \
                    <td style="border: 1px solid #999; padding: 0.5rem;">'+ CDD +'</td> \
                    <td style="border: 1px solid #999; padding: 0.5rem;" bgcolor="yellow">'+ CDDChange +'</td> \
                    <td style="border: 1px solid #999; padding: 0.5rem;">'+ Note +'</td> \
                    <td style="border: 1px solid #999; padding: 0.5rem;" bgcolor="yellow">'+ NOTEChange +'</td> \
                    <td style="border: 1px solid #999; padding: 0.5rem;" bgclolor="red">'+ CDDDiff +'</td> \
                </tr>';
    } 
    
    
    
  }
  table = table +'</table>' ;
  message = message + table;

  
  MailApp.sendEmail('', subject, "" , { htmlBody: message, cc: 'biggiol@ethirteen.com, taichung.sales.assist@bythehive.com' } );
  //MailApp.sendEmail('', subject, "" , { htmlBody: message, cc: 'biggio.lai@bythehive.com' } );
}

function run_CSV_import() {

  function getService() {
    return OAuth1.createService('netsuite')
    
    .setConsumerKey('7d351208b61283ec86a4588041ba5646ca046c36d4a152a07f1ef633a6714dac')
    .setConsumerSecret('9c1f7e7263277e814148577bd90f22b873823e71c024befc41458ee8cdab510c')
    .setAccessToken('78dd6056fd1eb277f0c73aeff00e5e3b5f1925661a9fc7744190f6c9e551ca22','2130228ee7e7446248038a54f7a128256e83268ca3786de2b8425cb039c63724')
    
    .setRealm('1334755')
    .setOAuthVersion('1.0')    
  }
  
  var url = 'https://1334755.restlets.api.netsuite.com/app/site/hosting/restlet.nl?script=616&deploy=1';
  var service = getService();
  var params = {
    contentType : 'application/json'
  };

  var response = service.fetch(url, params);
  
  //sendNotifEmail();
}




function getFormattedDate(s_date) {
  var date = new Date(s_date);
  var year = date.getFullYear().toString();
  var month = (1 + date.getMonth()).toString();
  month = month.length > 1 ? month : '0' + month;

  var day = date.getDate().toString();
  day = day.length > 1 ? day : '0' + day;
  
  if ( year >0) {
    return month + '/' + day + '/' + year;
  } else 
  { 
    return s_date;
  }
  
}
