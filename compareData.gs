function compareSheets() {
  var sa = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = sa.getSheets();
  var i = 0;
  var day = 0;
  
  //Find the largest number of days compared
  while(i < sheets.length) {
    var sheetName = sheets[i].getSheetName();
    var value = sheetName.substring(0,4);
    if(sheetName.substring(0,4) == "CRM ") {
      var newday = parseInt(sheetName.substr(sheetName.length-1, 1));
      if(newday > day) {
        day = newday;
      }
    }
    i++;
  }
  
  //Get the sheets that will be used
  var previous = day - 1;
  var SFTP1 = sa.getSheetByName("CRM - Day " + previous.toString());
  var SFTP2 = sa.getSheetByName("CRM - Day " + day.toString());
  var CSV = sa.getSheetByName("FR - Day " + day.toString());
  var inputs = sa.getSheetByName("Instructions");
  
  //Get the data from the sheets and the number of columns
  var data1 = SFTP1.getDataRange().getValues();
  var data2 = SFTP2.getDataRange().getValues();
  var csv = CSV.getDataRange().getValues();
  var headerlen = data1[0].length;
  var updateValue = inputs.getRange(4, 5).getValue();
  
  Logger.log("");
  //Set the concatenation column on the first sheet
  if(day == 2) {
    if(updateValue == 0) {
      var concat = "=TRIM(CONCATENATE(RC[-" + headerlen.toString() + "]:RC[-1]))";
    } else {
      var stopCol = headerlen - updateValue + 2;
      var startCol = stopCol - 2;
      var concat = "=TRIM(CONCATENATE(RC[-" + headerlen.toString() + "]:RC[-" + stopCol.toString() + "],RC[-" + startCol.toString() + "]:RC[-1]))";
    }
    if(data1[0][headerlen] != "Concatenation") {
      i = 1;
      data1[0].push("Concatenation");
      while(i < data1.length) {
        data1[i].push(concat);
        i++;
      }
    }
    //Write the data to the first SFTP sheet
    SFTP1.getRange(1, 1, data1.length, data1[0].length).setValues(data1);
    var concatCol = SFTP1.getLastColumn()-1;
  } else {
    var headers = data1[0];
    var concatenation = headers.indexOf("Concatenation");
    var concatCol = concatenation;
    headerlen = headerlen - 6;
    if(updateValue == 0) {
      var concat = "=TRIM(CONCATENATE(RC[-" + headerlen.toString() + "]:RC[-1]))";
    } else {
      var stopCol = headerlen - updateValue + 2;
      var startCol = stopCol - 2;
      var concat = "=TRIM(CONCATENATE(RC[-" + headerlen.toString() + "]:RC[-" + stopCol.toString() + "],RC[-" + startCol.toString() + "]:RC[-1]))";
    }
  }
  
  var formulas = [];
  formulas.push(concat);
  var formulaname = SFTP1.getSheetName();
  var lookupRange = SFTP1.getRange(2, concatCol + 1, SFTP1.getLastRow()-1, 1).getA1Notation();
  formulas.push("=IFERROR(VLOOKUP(RC[-1],\'" + formulaname + "\'! " + lookupRange + ", 1, FALSE), \"NO MATCH\")");
  formulas.push("=IF(RC[-2]<>RC[-1], \"NO MATCH\", \"MATCH\")");
  var idCol = inputs.getRange(3, 5).getValue();
  var idColDiff = SFTP2.getLastColumn()+4-idCol;
  lookupRange = SFTP2.getRange(2, idCol, SFTP2.getLastRow()-1, 1).getA1Notation();
  formulas.push("=IF(RC[-1]=\"NO MATCH\", IFERROR(VLOOKUP(RC[-" + idColDiff.toString() + "],\'" + formulaname + "\'!" + lookupRange + ", 1, FALSE), \"NEW\"), RC[-1])");
  formulas.push("=IF(RC[-1]=\"MATCH\", \"UPDATE\", IF(RC[-1]=\"NEW\", \"NEW\", \"CHANGED UPDATE\"))");
  formulaname = CSV.getSheetName();
  formulas.push("=IF(RC[-1]=\"CHANGED UPDATE\", IFERROR(VLOOKUP(RC[-2], \'" + formulaname + "\'!A:A, 1, FALSE), \"NOT UPDATED\"),\"\")");
  var newheaders = ["Concatenation","Lookup","Check","Check2","Final","Change Check"];
  
  i = 0;
  j = 0;
  headerlen = data2[0].length;
  while(i < newheaders.length) {
    data2[0][i+headerlen] = newheaders[i];
    i++;
  }
  
  i=1;
  while(i < data2.length) {
    j = 0;
    while(j < formulas.length){
      data2[i][j+headerlen]=formulas[j];
      j++;
    }
    i++;
  }
  
  SFTP2.getRange(1, 1, data2.length, data2[0].length).setValues(data2);
  data2 = SFTP2.getRange(1, 1, SFTP2.getLastRow(), SFTP2.getLastColumn()).getDisplayValues();
  data1 = SFTP1.getRange(1, 1, SFTP1.getLastRow(), SFTP1.getLastColumn()).getDisplayValues();
  
  i = 1;
  var headers = data2[0];
  var final = headers.indexOf("Final");
  var IDcol = SFTP1.getRange(2, idCol, data1.length-1, 1).getValues().toString();
  IDcol = IDcol.split(",");
  while(i < data2.length) {
    if(data2[i][final].trim() == "CHANGED UPDATE") {
      var rowFound = IDcol.indexOf(data2[i][idCol-1]) + 1;
      var data1Row = data1[rowFound];
      var data2Row = data2[i];
      var j = 1;
      while(j < data1Row.length) {
        if(data1Row[j] != data2Row[j]) {
          SFTP2.getRange(i+1, j+1).setBackground("yellow");
        }
        j++;
      }
    }
    i++;
  }
}