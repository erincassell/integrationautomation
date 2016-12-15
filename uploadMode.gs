function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var csvMenuEntries = [{name: "Compare Data", functionName: "compareSheets"}, 
                        {name: "Upload Mode Report", functionName: "uploadMode"},
                        {name: "Upload CRM File", functionName: "uploadSFTP"}];
  ss.addMenu("Live Monitoring", csvMenuEntries);
}

function uploadMode() {
  var sa = SpreadsheetApp.getActive();
  var sheets = sa.getSheets();
  var folders = DriveApp.getFileById(sa.getId()).getParents();
  var folder = folders.next();
  var CSVfiles = folder.getFilesByType(MimeType.CSV);
  
  var n = 0;
  var sheetNum = 200;
  while(n < sheets.length) {
    var curSheet = sheets[n].getName();
    if(curSheet.search("CSV") == 0) {
      sheetNum = n;
      break;
    }
    n++;
  }
  
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
  
  var data = Utilities.parseCsv(updateFile.getBlob().getDataAsString());
  updateFile.setName("Imported-" + updateFile.getName());
  folder.removeFile(newFile);

  if(sheetNum == 200) {
    var sheetDay = 2;
  } else {
    var sheetDay = sheets[sheetNum].getSheetName();
    sheetDay = parseInt(sheetDay.substr(sheetDay.length - 1, 1)) + 1;
    sheetNum = sheets.length + 2;
  }
  
  try {
    var newSheet = sa.insertSheet("FR - Day " + sheetDay.toString());
    newSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    newSheet.activate();
  } catch(err) {
    Logger.log(JSON.stringify(err));
  }
}
