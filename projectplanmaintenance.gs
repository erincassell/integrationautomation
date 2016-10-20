function mergeTaskColumns() {
  var ss, sa, lastRow, helper;
  var rngCols, valCols, i;
  
  //Merges columns B, C, D in the Project Plan sheet
  sa = SpreadsheetApp.getActive();
  ss = sa.getSheetByName("Project Plan");
  
  lastRow = ss.getLastRow();
  valCols = ss.getRange(1, 2, lastRow, 3).getValues();
  helper = valCols.length;
  i = 0;
  while(i < valCols.length) {
    helper = "entering";
    
    if(valCols[i][0] != "") {
    rngCols = ss.getRange(i+1, 2, 1, 3);
    rngCols.mergeAcross();
    Logger.log("stop");
    }
    i++;
  }
}

function mergeCategoryColumns() {
  var ss, sa, lastRow, helper;
  var rngCols, valCols, i;
  
  //Merges columns A, B, C, D in the Project Plan sheet
  sa = SpreadsheetApp.getActive();
  ss = sa.getSheetByName("Project Plan");
  
  lastRow = ss.getLastRow();
  valCols = ss.getRange(2, 1, lastRow, 1).getValues();
  
  i = 0;
  while(i < valCols.length) {    
    if(valCols[i][0] != "") {
    rngCols = ss.getRange(i+2, 1, 1, 4);
    rngCols.mergeAcross();
    }
    i++;
  }
}