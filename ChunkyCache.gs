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