function uploadSFTP() {
  var sa = SpreadsheetApp.getActive();
  var sheets = sa.getSheets();
  var folders = DriveApp.getFileById(sa.getId()).getParents();
  var folder = folders.next();
  var TXTfiles = folder.getFilesByType(MimeType.PLAIN_TEXT);
  
  var n = 0;
  var sheetNum = 200;
  while(n < sheets.length) {
    var curSheet = sheets[n].getName();
    if(curSheet.search("CRM") == 0) {
      sheetNum = n;
      //break;
    }
    n++;
  }
  
  Logger.log("");
  while(TXTfiles.hasNext()) {
    var TXTfile = TXTfiles.next();
    var name = TXTfile.getName();
    if(name.search("ToFrontRush") > 0 && name.search("PARSED") < 0) {
      var blob = TXTfile.getBlob().getDataAsString().split('\n');
      var i = 0;
      var rows = [];
      i = 0;
      while(i < blob.length) {
        blob[i] = blob[i].toString().split("\"").join("");
        rows.push(blob[i].split('\t'));
        i++;
      }
      
      if(sheetNum == 200) {
        var sheetDay = 1;
      } else {
        var sheetDay = sheets[sheetNum].getSheetName();
        sheetDay = parseInt(sheetDay.substr(sheetDay.length - 1, 1)) + 1;
        sheetNum = sheets.length + 1;
      }
      
      var sheetName = "CRM - Day " + sheetDay.toString();
      var newSheet = sa.insertSheet(sheetName, sheets.length+1);
      rows.pop();
      newSheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
      var interim = name.substring(0, name.length - 4);
      TXTfile.setName("PARSED-" + interim + ".txt");
    } 
  }
}
