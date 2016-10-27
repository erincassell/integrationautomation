// http://stackoverflow.com/questions/11273268/script-import-local-csv-in-google-spreadsheet
// Modified by ecassell@frontrush to accommodate our needs

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

function doGet(e) {
  var app = UiApp.createApplication().setTitle("Upload CSV to Sheet");
  var formContent = app.createVerticalPanel();
  formContent.add(app.createFileUpload().setName('thefile'));
  formContent.add(app.createSubmitButton('Start Upload'));
  var form = app.createFormPanel();
  form.add(formContent);
  app.add(form);
  SpreadsheetApp.getActiveSpreadsheet().show(app);
}

function doPost(e) {
  var helper, sheet;
  // data returned is a blob for FileUpload widget
  var fileBlob = e.parameter.thefile;

  // parse the data to fill values, a two dimensional array of rows
  // Assuming newlines separate rows and commas separate columns, then:
  var values = []
  var rows = fileBlob.contents.split('\n');
  for(var r=0, max_r=rows.length; r<max_r; ++r)
    values.push( rows[r].split(',') );  // rows must have the same number of columns

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  sheet = ss.getActiveSheet();
  for (var i = 0; i < values.length; i++) {
    sheet.getRange(i+1, 1, 1, values[i].length).setValues(new Array(values[i]));
  }
  createDay();
}