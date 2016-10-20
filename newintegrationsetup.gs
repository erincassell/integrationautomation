function schoolSetup() {
  var sa, ss, lastCol, lastRow, headers, i, header;
  var phaseCol, folderCol, schools, helper, phase, folder;
  var columns = [4], schoolfolder, templatefolder;
  var integrationfolders, integrationfolder, message;
  
  //Default values
  sa = SpreadsheetApp.getActive();
  ss = sa.getSheetByName("Full List of Integrations");
  lastCol = ss.getLastColumn();
  lastRow = ss.getLastRow();
  headers = ss.getRange(1, 1, 1, lastCol).getValues();
  integrationfolders = DriveApp.getFoldersByName("*In Process Integrations");
  if(integrationfolders.hasNext()) {
    integrationfolder = integrationfolders.next();
  }
  
  //Find the column number for the Implemenation Phase and Folder columns
  i = 0;
  while(i < lastCol) {
    header = headers[0][i].trim();
    switch(header) {
      case "School":
        columns.push(i);
        break;
      case "CRM/SIS":
        columns.push(i);
        break;
      case "Overall Implementation Status":
        columns.push(i);
        break;
      case "Folder":
        columns.push(i);
        break;
    }
    i++;
  }
  columns.reverse();
  columns.pop();
  columns.reverse();
  
  //Get all of the data in the sheet
  schools = ss.getRange(2, 1, lastRow, lastCol).getValues();

  try {
    //Loop through and find those that need a folder
    i = 0;
    while(i < schools.length) {
      phase = schools[i][columns[2]].trim();
      folder = schools[i][columns[3]].trim();
      if(phase == 'Signed Contract' && folder == "") {
        ss.getRange(i+2, columns[3]+2).setFormulaR1C1("=CONCATENATE(IFERROR(VLOOKUP(R[0]C[-26], Inputs!R2C7:R20C9, 3, FALSE), \"\"), \" and \", IFERROR(VLOOKUP(R[0]C[-25], Inputs!R2C7:R20C9, 3, FALSE)))");
        ss.getRange(i+2, columns[3]+3).setFormulaR1C1("=IF(TRIM(R[0]C[-1])=\"and\", \"\", IF(RIGHT(TRIM(R[0]C[-1]),3)=\"and\", LEFT(R[0]C[-1], LEN(R[0]C[-1])-5), R[0]C[-1]))");
        ss.getRange(i+2, columns[3]+4).setFormulaR1C1("=IFERROR(VLOOKUP(R[0]C[-28], Inputs!R2C7:R20C9, 2, FALSE), \"\")");
        ss.getRange(i+2, columns[3]+5).setFormulaR1C1("=IFERROR(VLOOKUP(R[0]C[-28], Inputs!R2C7:R20C9, 2, FALSE), \"\")");
        ss.activate();
        schoolfolder = createFolder(schools, i, columns[0], columns[1]);
        templatefolder = getMappingDoc(schoolfolder, schools, i, columns[0], columns[1]);
        getS2SDoc(schoolfolder, templatefolder, schools, i, columns[0], columns[1]);
        getKickoff(schoolfolder, templatefolder, schools, i, columns[0], columns[1]);
        moveFolder(schoolfolder, integrationfolder);
        //copyFormulas(i, columns[3]);
        sendSetupEmail(schools, i, columns[0], columns[1], columns[3]);
        ss.getRange(i+2, columns[3]+1).setValue("X");
      }
      i++;
    }
  } catch(e) {
    
    message = "Message: " + e.message + "\n";
    message += "File: " + e.fileName + "\n";
    message += "Line: " + e.lineNumber + "\n";
    MailApp.sendEmail("ecassell@frontrush.com", "Error in New Integration Setup", message);
  }
}

function createFolder(schools, row, schoolCol, crmCol) {
  var foldername, newschools, helper, i;
  var newFolder, insidefolders, insidefolder;
  
  //Name of the folders that need to go inside
  insidefolders = ["0 Communication and Meetings", "1 Pre Integration", "2 Integration and Testing", "3 Go Live", "4 Support"];
  
  //Create main folder
  foldername = schools[row][schoolCol] + " (" + schools[row][crmCol] + ")";
  newFolder = DriveApp.createFolder(foldername);
  
  //Add in the child folders
  i = 0;
  while(i < insidefolders.length) {
    insidefolder = DriveApp.createFolder(insidefolders[i]);
    newFolder.addFolder(insidefolder);
    DriveApp.removeFolder(insidefolder);
    i++;
  }
  return newFolder;
}

function getMappingDoc(newFolder, schools, row, schoolCol, crmCol) {
  var mappings, mapping, mappingname, mappingtemplate, searchvalue;
  var templatecopy, templatecopyname, helper;
  var templatefolders, templatefolder, parents, parent, template;
                       
  //Get Templates folder
  templatefolders = DriveApp.getFoldersByName("Templates");
  while(templatefolders.hasNext()){
    templatefolder = templatefolders.next();
  }
  
  helper = templatefolder.getName();
  
  //Locate the correct template
  mappings = templatefolder.getFilesByType(MimeType.GOOGLE_SHEETS);
  //mappings = templatefolder.getFiles();
  while(mappings.hasNext()) {
    mapping = mappings.next();
    searchvalue = schools[row][crmCol] + " ";
    mappingname = mapping.getName();
    if(mappingname.search(searchvalue) == 0) { //Only want it if it starts with the word
      mappingtemplate = mapping;
    }
  }
  
  //Change the name of the mapping document
  templatecopy = mappingtemplate.makeCopy(newFolder);
  templatecopyname = templatecopy.getName();
  templatecopyname = templatecopyname.slice(8, templatecopyname.length); 
  templatecopy.setName(schools[row][schoolCol] + " " + templatecopyname);
  
  return templatefolder;
}

function getS2SDoc(newFolder, templatefolder, schools, row, schoolCol, crmCol) {
  var documents, document, documentname, handoff, handoffcopy, handoffcopyname;
  
  //Locate the S2S document in the Templates folder
  documents = templatefolder.getFilesByType(MimeType.GOOGLE_DOCS);
  while(documents.hasNext()) {
    document = documents.next();
    documentname = document.getName();
    if(documentname == "Sales to Services Handoff") {
      handoff = document;
    }
  }
  
  //Change the name of the mapping document
  handoffcopy = handoff.makeCopy(newFolder);
  handoffcopyname = handoffcopy.getName();
  handoffcopyname = handoffcopyname.slice(8, handoffcopyname.length); 
  handoffcopy.setName(schools[row][schoolCol] + " " + handoffcopyname);
}

function getKickoff(newFolder, templatefolder, schools, row, schoolCol, crmCol) {
  var documents, document, documentname, handoff, handoffcopy, handoffcopyname, searchvalue;
  
  //Locate the S2S document in the Templates folder
  documents = templatefolder.getFilesByType(MimeType.GOOGLE_SLIDES);
  while(documents.hasNext()) {
    document = documents.next();
    searchvalue = schools[row][crmCol];
    documentname = document.getName();
    if(documentname.search(searchvalue)==0) {
      handoff = document;
    }
  }
  
  //Change the name of the mapping document
  handoffcopy = handoff.makeCopy(newFolder);
  handoffcopyname = handoffcopy.getName();
  handoffcopyname = handoffcopyname.slice(8, handoffcopyname.length); 
  handoffcopy.setName(schools[row][schoolCol] + " " + handoffcopyname);
}

function moveFolder(schoolfolder, integrationfolder) {
  integrationfolder.addFolder(schoolfolder);
  DriveApp.removeFolder(schoolfolder);
}

function copyFormulas(row, folderCol) {
  var sa, ss, helper;
  var copyRng, toRng;
  
  sa = SpreadsheetApp.getActiveSpreadsheet();
  ss = sa.getActiveSheet();
  helper = ss.getName();
  
  copyRng = ss.getRange(row+1, folderCol+2, 1, 4).getFormulasR1C1();

  //toRng = ss.getRange(row+2, folderCol+2, 1, 4).setValues(copyRng);
  //ss.getRange(row+2, folderCol+2, 1, 4).activate();
}

function sendSetupEmail(schools, row, schoolCol, crmCol, folderCol){
  var email, cc, assignees, institution, crm, columns;
  columns = schools[0].length;
  
  email = schools[row][folderCol+3];
  cc = schools[row][folderCol+4];
  assignees = schools[row][folderCol+2];
  institution = schools[row][schoolCol];
  crm = schools[row][crmCol];
  
  var subject = "New Integration: " + institution + " (" + crm + ")";
  
  var message = assignees + ",\n\n";
  message += "Congratulations!\nYou've just been assigned a new integration.\n\n";
  message += "    Institution: " + institution + "\n";
  message += "    CRM: " + crm + "\n\n";
  message += "The folder has been set-up, and you can schedule the sales to services handoff for the school.\n\n";
  message += "Thank you!";
  
  MailApp.sendEmail(email, subject, message, {cc: cc});
}
