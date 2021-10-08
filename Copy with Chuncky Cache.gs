function testFun(){
  CopyRange2("sheet key here","sheet1 Name","cell range","sheet name 2","sheet row","sheet_column","check=1/0")
}


/******************************  CopyRange by chunky cache   ******************************** */
function CopyRange2(sheetkey, sheet1, cell_range, sheet2, sheet_row, sheet_column,check){
  var x=0;
  var chunky = ChunkyCache(CacheService.getDocumentCache(), 1024*90);
  var CacheName = `${sheet1}_${sheet2}_${cell_range}`;
  var values= chunky.get(CacheName);
  var cacheTime = 5; //分鐘
  var PropertyName = `${sheet1}_${sheet2}_${cell_range}_length`;
  var Property = PropertiesService.getDocumentProperties();
  var Pre_length = Property.getProperty(PropertyName);
  var formatter = new Intl.DateTimeFormat();
  function parseISOString(s) {
  var b = s.split(/\D+/);

  return new Date(Date.UTC(b[0], --b[1], b[2], b[3], b[4], b[5], b[6]));
}
   Logger.log('一一一一一一一一 Ｃｏｐｙ Ｒａｎｇｅ２ 一一一一一一一一一') 
if(values == undefined || typeof(values.length) == undefined) {
     
      Logger.log(`pull resource & put on cache -> ${CacheName}`)
      x++;    
      values = Sheets.Spreadsheets.Values.get(sheetkey,sheet1+'!'+cell_range).values;
      
      chunky.put(CacheName, values, cacheTime*60);
      var values= chunky.get(CacheName);
  }
  
      if(check==1){
        if(Pre_length==null){
          Logger.log(`Cache is null ,set new Cache:${CacheName}`)
          set(); 
          }         
        else if(Pre_length/2>values.length){
          Logger.log('Error,New Data length is less 50% than last time length');
          }
        else{
          Logger.log(`get values on Cache : ${CacheName}`)
          set();
          }
      }
      else{
        Logger.log('ignore,update anyway')
        set();
      };
      function set(){
        for(var i=0;i<values.length;i++){
          for(var j=0;j<values[i].length;j++){
            var value = `${values[i][j]}`;
            if(value.indexOf('T16:00:00.000Z')>0){
                  values[i][j] = formatter.format(parseISOString(value)); 
            }
          }
        }
        
      
        var ss = SpreadsheetApp.getActive().getSheetByName(sheet2);
        var Range = ss.getRange(sheet_row || 1,sheet_column || 1, values.length , values[0].length);
        ss.getRange(sheet_row || 1,sheet_column || 1, ss.getMaxRows(), values[0].length ).clear();
        SpreadsheetApp.flush();
        
        var ranges = sheet2+`!`+ Range.getA1Notation();
        Logger.log('Range:'+ranges);   

        var val = values;
        console.log('values length:'+val.length)
        var spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
        Sheets.Spreadsheets.Values.update({values: val}, spreadsheetId, ranges, {valueInputOption: "USER_ENTERED"});
        Property.deleteProperty(PropertyName);
        Property.setProperty((PropertyName),values.length)
        
        if(x==1){
          SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet2).getRange(sheet_row,sheet_column).setNote('Last pull is at:' +Utilities.formatDate(new Date(), "GMT+8", "dd/MM/yyyy HH:mm:ss"));
        }
        console.log(`Done,Cache will Only live in ${cacheTime} minutes`); 
    }
}

/*****************************  刪除頁保護/將保護加回去 (6/22) ***************************/
//                      使用方法 
//     將表單設為所有人都可以編輯 =>  GoEdit('表單名稱')
//     將表單設為原本的保護      =>  NoEdit('表單名稱')
//     
//     測試function => Unprotected_backup_setup

//測試用function(將保護設為警告/將保護加回去)
    function Unprotected_backup_setup(){
      unlock('LockerTest');
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getSheetByName('LockerTest');
        sheet.deleteRow(1)
      lock('LockerTest');
    }
//保護設為WarningOnly
    function unlock(sheetName){
     UrlFetchApp.fetch(`https://script.google.com/macros/s/AKfycbyraxH94LRL2GoVIOsfTTVhVSj_oOhUE4D8lydF9EvpnJMoq1XmUH-SmhjY8a6KkkA5/exec?page=${sheetName}&status=off`);
    }
    function bp(sheetName){
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName(sheetName);
      var protection =sheet.protect();
      var unprotect = protection.getUnprotectedRanges()
      var backupRange=[]
      var backupEditor = protection.getEditors().join('|')
      for(var i=0;i<unprotect.length;i++){
        backupRange.push(unprotect[i].getA1Notation());
      }
      console.log(backupRange)
      backupRange = backupRange.join('|');
         
      var Property = PropertiesService.getDocumentProperties();
      Property.setProperty('Range',backupRange);
      Property.setProperty('Editor',backupEditor);
      protection.remove();
    }
//保護並addEditors
function lock(sheetName){
     UrlFetchApp.fetch(`https://script.google.com/macros/s/AKfycbyraxH94LRL2GoVIOsfTTVhVSj_oOhUE4D8lydF9EvpnJMoq1XmUH-SmhjY8a6KkkA5/exec?page=${sheetName}&status=on`);
    }
    function ap(sheetName){
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName(sheetName);
      var Property = PropertiesService.getDocumentProperties();
      var backupEditor = Property.getProperty('Editor').split('|');
      var backupRange = Property.getProperty('Range').split('|');

      var protection =sheet.protect();
      protection.removeEditors(protection.getEditors());
      protection.addEditors(backupEditor);
      var unprotect = protection.getUnprotectedRanges()
      for(var i=0;i<backupRange.length;i++){
        unprotect.push(sheet.getRange(backupRange[i]))
      }
      protection.setUnprotectedRanges(unprotect)
      Property.deleteAllProperties();
    }

/********************************* doGet [部署 CSV URL] *****************************************/
 function doGet(e){
   //parameter給定page
    var page = e.parameter.page;
    var status = e.parameter.status;                                       
    //定義一個空値，以及googlesheet的陣列値
    var csv = "";                                                 
    var v = SpreadsheetApp
            .getActiveSpreadsheet()
            .getSheetByName('CDD_Import')
            .getDataRange()
            .getValues();

    //分別將每一個値賦予分號
    for(var row = 0 ; row <v.length;row++){
        for(var col = 0; col<v[row].length;col++){
          v[row][col] = isDate(v[row][col]);      // Format, if date
          if(v[row][col].toString().indexOf(",")!=-1){
            v[row][col] = `\"${v[row][col]}\"`
          }
        }
    }
  //為字串賦予ＣＳＶ格式並儲存至變數中
   v.forEach(function(e) {
      csv += e.join(',')+'\n'
    })
  //設定回傳値url: https://script.google.com/macros/s/AKfycbyZSgcxqci_rpyNDcZFe_by3KTtAd90OMF4BEo12s497Ra_9JZwLz7e5ymKLkWwqGlr/exec?page=1





    if(page == 1){
          return ContentService.createTextOutput(csv).downloadAsFile("CDD_Import.csv").setMimeType(ContentService.MimeType.CSV);
      }
      else if(status == 'off'){
           bp(page)
      }
      else if(status == 'on'){
           ap(page)
      }
       else if(page !='',range == off){
          lockRange(page,range);
      }
       else if(page !='',range == on){
          lockRange(page,range);
      }
      else{
        return ContentService.createTextOutput(
        ` 
          Parameter Error.
        `
        )

      }
}            

// Returns 'true' if variable d is a date object.
function isValidDate(d) {
  if ( Object.prototype.toString.call(d) !== "[object Date]" )
    return false;
  return !isNaN(d.getTime());
}

// Test if value is a date and if so format
// otherwise, reflect input variable back as-is. 
function isDate(sDate) {
  if (isValidDate(sDate)) {
    sDate = Utilities.formatDate(new Date(sDate), 'GMT+8', "dd-MMM-yyyy");
  }
  return sDate;
}





/**** ChunkyCache by pilbot https://gist.github.com/pilbot/9d0567ef1daf556449fb ********/ 


function ChunkyCache(cache, chunkSize){
  return {
    put: function (key, value, timeout) {
      var json = JSON.stringify(value);
      var cSize = Math.floor(chunkSize / 2);
      var chunks = [];
      var index = 0;
      while (index < json.length){
        cKey = key + "_" + index;
        chunks.push(cKey);
        cache.put(cKey, json.substr(index, cSize), timeout+5);
        index += cSize;
      }
      
      var superBlk = {
        chunkSize: chunkSize,
        chunks: chunks,
        length: json.length
      };
      cache.put(key, JSON.stringify(superBlk), timeout);
    },
    get: function (key) {
      var superBlkCache = cache.get(key);
      if (superBlkCache != null) {
        var superBlk = JSON.parse(superBlkCache);
        chunks = superBlk.chunks.map(function (cKey){
          return cache.get(cKey);
        });
        if (chunks.every(function (c) { return c != null; })){
          return JSON.parse(chunks.join(''));
        }
      }
    }
  };
};

function testGetCacheFrom(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');
  var data = sheet.getDataRange().getValues();
  var chunky = ChunkyCache(CacheService.getDocumentCache(), 1024*90);
  chunky.put('Data', data, 120);
  var check = chunky.get('Data');
  var sheetPut = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Out');
  for (c in check) {
     sheetPut.appendRow(check[c]);
  }
}




