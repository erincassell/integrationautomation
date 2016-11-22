function schoolSetup() {
  
  //Default values
  var sa = SpreadsheetApp.getActive();
  var ss = sa.getSheetByName("Full List of Integrations");
  var sc = sa.getSheetByName("College Contacts");
  var integrations = ss.getDataRange().getValues();
  var contacts = sc.getDataRange().getValues();
  var headers = integrations[0];
  var headerValues = new Object();
  
  var integrationfolders = DriveApp.getFoldersByName("*In Process Integrations");
  var integrationfolder = integrationfolders.next();
  
  //Find the column number for the Implemenation Phase and Folder columns
  var i = 0;
  while(i < headers.length) {
    header = headers[i].trim();
    switch(header) {
      case "School":
        headerValues['school'] = i;
        break;
      case "CRM/SIS":
        headerValues['CRM'] = i;
        break;
      case "Integration Status":
        headerValues['status'] = i;
        break;
      case "Integration Manager":
        headerValues['im'] = i;
        break;
      case "Integration Specialist":
        headerValues['is'] = i;
        break;
      case "Folder":
        headerValues['folder'] = i;
        break;
      case "Scheduling Link":
        headerValues['link'] = i;
        break;
      case "Primary Email":
        headerValues['primary'] = i;
        break;
      case "Secondary Email":
        headerValues['secondary'] = i;
        break;
      case "Assignees":
        headerValues['assignees'] = i;
        break;
    }
    i++;
  }
  
  i = 0;
  while(i < integrations.length) {
    if(integrations[i][0].trim() == "") {
      var holder = integrations.splice(i);
    }
    i++;
  }

  try {
    //Loop through and find those that need a folder
    i = integrations.length-1;

    while(i > 0) {
      var statusCol = integrations[i][Number(headerValues['status']).valueOf()];
      var folderCol = integrations[i][Number(headerValues['folder']).valueOf()];
      if(statusCol == 'Signed Contract' && folderCol == "") {
        var school = integrations[i].slice(0);
        var schoolContacts = [];
        var j = 0;
        while(j < contacts.length) {
          if(contacts[j][1] == school[headerValues['school']]) {
            schoolContacts.push(contacts[j]);
          }
          j++;
        }
        var schoolfolder = createFolder(school, headerValues);
        var templatefolder = getMappingDoc(school, schoolfolder, headerValues, schoolContacts);
        createWelcome(school, schoolfolder, templatefolder, headerValues, schoolContacts);
        copyFile(school, schoolfolder, MimeType.GOOGLE_DOCS, "Sales to Services Handoff", templatefolder, headerValues);
        copyFile(school, schoolfolder, MimeType.GOOGLE_SHEETS, "Sample Data", templatefolder, headerValues);
        getKickoff(school, schoolfolder, templatefolder, headerValues);
        copyFile(school, schoolfolder, MimeType.GOOGLE_DOCS, "Credential Information", templatefolder, headerValues);
        copyFile(school, schoolfolder, MimeType.GOOGLE_SHEETS, "Integration Checklists", templatefolder, headerValues);
        copyFile(school, schoolfolder, MimeType.GOOGLE_SHEETS, "Live Monitoring", templatefolder, headerValues);
        createSupport(school, schoolfolder, templatefolder, headerValues, schoolContacts);
        moveAgreement(school, schoolfolder, headerValues);
        moveFolder(schoolfolder, integrationfolder);
        sendSetupEmail(school, headerValues);
        ss.getRange(i+1, headerValues['folder']+1).setValue("X");
      } else {
        if(school[headerValues['folder']] == "X") {
          i = 0;
        }
      }
      i--;
    }
  } catch(e) {
    
    message = "Message: " + e.message + "\n";
    message += "File: " + e.fileName + "\n";
    message += "Line: " + e.lineNumber + "\n";
    MailApp.sendEmail("ecassell@frontrush.com", "Error in New Integration Setup", message);
  }
}

function createFolder(school, headerValues) {  
  //Name of the folders that need to go inside
  var insidefolders = ["0 Communication and Meetings", "1 Pre Integration", "2 Integration and Testing", "3 Go Live", "4 Support"];
  
  //Create main folder
  var foldername = school[headerValues['school']] + " (" + school[headerValues['CRM']] + ")";
  var newFolder = DriveApp.createFolder(foldername);
  
  //Add in the child folders
  i = 0;
  while(i < insidefolders.length) {
    var insidefolder = DriveApp.createFolder(insidefolders[i]);
    newFolder.addFolder(insidefolder);
    DriveApp.removeFolder(insidefolder);
    i++;
  }
  return newFolder;
}

function getMappingDoc(school, newFolder, headerValues, contacts) {                       
  //Get Templates folder
  var templatefolders = DriveApp.getFoldersByName("Templates");
  var templatefolder = templatefolders.next();
  
  //Locate the correct template
  var mappings = templatefolder.getFilesByType(MimeType.GOOGLE_SHEETS);
  
  while(mappings.hasNext()) {
    var mapping = mappings.next();
    var helper = mapping.getName();
    var searchvalue = school[headerValues['CRM']] + " ";
    if(mapping.getName().search(searchvalue) == 0) { //Only want it if it starts with the word
      var mappingtemplate = mapping;
    }
  }
  
  //Make a copy of the template to the school folder with the correct name
  var templatecopy = mappingtemplate.makeCopy(school[headerValues['school']] + " " + mappingtemplate.getName(), newFolder);
  var ss = SpreadsheetApp.open(templatecopy);
  
  if(contacts.length != 0) {
    var contactSheet = ss.insertSheet("Contacts", 1);
    contactSheet.getRange(1, 1, contacts.length, contacts[0].length).setValues(contacts);
    contactSheet.deleteColumns(1, 2);
  }
  return templatefolder;
}

function createWelcome(school, schoolfolder, templatefolder, headerValues, contacts) {
  var documents = templatefolder.getFilesByName("Welcome Message");
  var document = documents.next();
  
  //Find the contact information
  var i = 0;
  while(i < contacts.length) {
    if(contacts[i][1] == school[headerValues['school'].valueOf()] && contacts[i][5] == "X") {
      var first = contacts[i][6];
      var email = contacts[i][9];
    }
    i++;
  }
  
  var welcomeFile = document.makeCopy(school[headerValues['school']] + " " + document.getName(), schoolfolder);
  var welcome = DocumentApp.openById(welcomeFile.getId());
  
  var welcomeBody = welcome.getBody();
  welcomeBody.replaceText("<<Institution>>", school[headerValues['school']]);
  welcomeBody.replaceText("<<CRM>>", school[headerValues['CRM']]);
  welcomeBody.replaceText("<<First>>", first);
  welcomeBody.replaceText("<<POC>>", email);
  welcomeBody.replaceText("<<Primary>>", school[headerValues['im']]);
  welcomeBody.replaceText("<<Backup>>", school[headerValues['is']]);
  welcomeBody.replaceText("<<link>>", school[headerValues['link']]);
  welcome.saveAndClose();
}

function copyFile(school, newFolder, fileType, fileName, templateFolder, headerValues) {
  var files = templateFolder.getFilesByType(fileType);
  while(files.hasNext()) {
    var file = files.next();
    if(file.getName() == fileName) {
      var foundFile = file;
      var helper = foundFile.getName();
    }
  }
  
  var fileCopy = foundFile.makeCopy(school[headerValues['school']] + " " + foundFile.getName(), newFolder);  
}

function getKickoff(school, newFolder, templatefolder, headerValues) {
  
  //Locate the S2S document in the Templates folder
  var documents = templatefolder.getFilesByType(MimeType.GOOGLE_SLIDES);
  while(documents.hasNext()) {
    var document = documents.next();
    var searchvalue = school[headerValues['CRM']];

    if(document.getName().search(searchvalue)==0) {
      var handoff = document;
    }
  }
  
  //Change the name of the mapping document
  var copy = handoff.makeCopy(school[headerValues['school']] + " " + handoff.getName(), newFolder);
}

function moveAgreement(school, schoolfolder, headerValues) {
  var folders = DriveApp.getFoldersByName("Executed Contracts");
  var folder = folders.next();
  var searchvalue = school[headerValues['school']];
  
  var files = folder.getFilesByType(MimeType.PDF);
  while(files.hasNext()) {
    var file = files.next();
    var test = file.getName();
    var helper = 0;
    
    if(file.getName().search(searchvalue)==0){
      var agreement = file;
      agreement.makeCopy(school[headerValues['school']] + " " + agreement.getName(), schoolfolder);
      folder.removeFile(agreement);
    }
  }
}

function createSupport(school, schoolfolder, templatefolder, headerValues, contacts) {
  var documents = templatefolder.getFilesByName("Move to Support Email");
  var document = documents.next();
  
  var supportFile = document.makeCopy(school[headerValues['school']] + " " + document.getName(), schoolfolder);
  var support = DocumentApp.openById(supportFile.getId());
  
  var supportBody = support.getBody();
  supportBody.replaceText("<<Institution>>", school[headerValues['school']]);
  supportBody.replaceText("<<CRM>>", school[headerValues['CRM']]);
  supportBody.replaceText("<<Integration Manager>>", school[headerValues['im']]);
  support.saveAndClose();
}

function moveFolder(schoolfolder, integrationfolder) {
  integrationfolder.addFolder(schoolfolder);
  DriveApp.removeFolder(schoolfolder);
}

function sendSetupEmail(school, headerValues){
  var email, cc, assignees, institution, crm, columns;
  columns = school.length;
  
  email = school[headerValues['primary']];
  cc = school[headerValues['secondary']];
  assignees = school[headerValues['assignees']];
  institution = school[headerValues['school']];
  crm = school[headerValues['CRM']];
  
  var subject = "New Integration: " + institution + " (" + crm + ")";
  
  var message = assignees + ",\n\n";
  message += "Congratulations!\nYou've just been assigned a new integration.\n\n";
  message += "    Institution: " + institution + "\n";
  message += "    CRM: " + crm + "\n\n";
  message += "The folder has been set-up, and you can schedule the sales to services handoff for the school.\n\n";
  message += "Thank you!";
  
  MailApp.sendEmail(email, subject, message, {cc: cc});
}