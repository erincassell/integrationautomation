function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var csvMenuEntries = [{name: "Run Compare", functionName: "compareSheets"}, {name: "Upload CSV file", functionName: "doGet"}];
  ss.addMenu("Live Monitoring", csvMenuEntries);
}

function createDay() {
  var ss = SpreadsheetApp.getActive();
  var sheetCount = ss.getSheets().length;
  var sheet = ss.getSheetByName("Upload");
  var newsheet = sheet.copyTo(ss);
  newsheet.setName("Day " + sheetCount);
  sheet.clearContents();
}

function uploadSheet() {
  var sa = SpreadsheetApp.getActive();
  var sheets = sa.getSheets();
  var folders = DriveApp.getFileById(sa.getId()).getParents();
  var folder = folders.next();
  var CSVfiles = folder.getFilesByType(MimeType.CSV);
  
  while(CSVfiles.hasNext()) {
    var CSVfile = CSVfiles.next();
    var name = CSVfile.getName();
    if(name.search("fr-_all_admissions") == 0 && name.search("updated") > 0) {
      var updateFile = CSVfile;
    } else {
      if(name.search("fr-_all_admissions") == 0) {
        var newFile = CSVfile;
      }
    }
  }
  
  if(sheets.length == 1) {
    var data = Utilities.parseCsv(newFile.getBlob().getDataAsString());
    newFile.setName("Imported-" + newFile.getName());
  } else {
    var data = Utilities.parseCsv(updateFile.getBlob().getDataAsString());
    updateFile.setName("Imported-" + updateFile.getName());
  }
  
  try {
    var today = new Date().toDateString();
    var newSheet = sa.insertSheet(today + " - Day " + sa.getSheets().length);
    newSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    newSheet.activate();
  } catch(err) {
    Logger.log(JSON.stringify(err));
  }
}
