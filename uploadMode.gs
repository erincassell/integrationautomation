function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Tasks')
      .addSubMenu(ui.createMenu('Management')
          .addItem('Move Completed', 'moveComplete')
          .addItem('Move Daily', 'moveDaily')
          .addItem('Move Inbox', 'moveInbox'))
      .addSeparator()
      .addSubMenu(ui.createMenu('Projects')
          .addItem('Move to Projects', 'moveProjects')
          .addItem('Move NAs', 'moveNA'))
      .addSeparator()
      .addSubMenu(ui.createMenu('Sorting')
          .addItem('By Due Date', 'sortDue')
          .addItem('By Priority', 'sortPriority'))
      .addToUi();
  sortDue();
}

function earlyMorning() {
  moveComplete();
  reprioritize();
  moveDaily();
}

function reprioritize() {
  var sa = SpreadsheetApp.getActiveSpreadsheet();
  var ss = sa.getSheetByName("Current");
  var allTasks = ss.getDataRange().getValues();
  for(var i = 1; i < allTasks.length; i++) {
    if(allTasks[i][0] > 0) {
      allTasks[i][0] = 4;
      allTasks[i][1] = "";
    }
  }
  
  ss.getRange(1, 1, allTasks.length, allTasks[0].length).setValues(allTasks);
}

function moveComplete() {
  var sa = SpreadsheetApp.getActiveSpreadsheet();
  var ss = sa.getSheetByName("Current");
  var allTasks = ss.getDataRange().getValues();
  var header = allTasks[0];
  var move = [];
  var keep = [];
  var toDelete = [];
  var completeCol = header.indexOf("Complete");
  var today = new Date();
  var dd = today.getDate();
  var mm = today.getMonth()+1; //January is 0!
  var yyyy = today.getFullYear();

  if(dd<10) {
    dd='0'+dd;
  } 

  if(mm<10) {
    mm='0'+mm;
  } 

  today = mm+'/'+dd+'/'+yyyy;
  
  for(var i = 0; i < allTasks.length; i ++) {
    if(allTasks[i][completeCol].toUpperCase() == 'X') {
      allTasks[i].push(today);
      move.push(allTasks[i]);
    } else if(allTasks[i][completeCol].toUpperCase() == 'D') {
      allTasks[i].push(today);
      toDelete.push(allTasks[i]);
    } else {
      keep.push(allTasks[i]);
    }
  }
  
  if(move.length > 0) {
    var completeSS = sa.getSheetByName("Complete");
    var completed = completeSS.getDataRange().getValues();
    completeSS.getRange(completed.length+1, 1, move.length, move[0].length).setValues(move);
  }
  
  if(toDelete.length > 0) {
    var deleteSS = sa.getSheetByName("Deleted");
    var deleted = deleteSS.getDataRange().getValues();
    deleteSS.getRange(deleted.length+1, 1, toDelete.length, toDelete[0].length).setValues(toDelete);
  }
  
  ss.clearContents();
  ss.getRange(1, 1, keep.length, keep[0].length).setValues(keep);
  
  sortPriority();
}

function moveInbox() {
  var sa = SpreadsheetApp.getActiveSpreadsheet();
  var inbox = sa.getSheetByName("Inbox");
  var current = sa.getSheetByName("Current");
  var inboxTasks = inbox.getDataRange().getValues();
  inboxTasks.reverse();
  inboxTasks.pop();
  inboxTasks.reverse();
  
  var currentTasks = current.getDataRange().getValues();
  current.getRange(currentTasks.length+1, 1, inboxTasks.length, inboxTasks[0].length).setValues(inboxTasks);
  
  sortDue();
  
  inbox.getRange(2, 1, inboxTasks.length, inboxTasks[0].length).clearContent();
  sa.setActiveSheet(current);
}

function sortDue() {
  var sa = SpreadsheetApp.getActiveSpreadsheet();
  var ss = sa.getSheetByName("Current");
  
  var data = ss.getDataRange().getValues();
  var header = data[0];
  var dueCol = header.indexOf("Due");
  
  var sortRange = ss.getRange(2, 1, data.length - 1, data[0].length);
  sortRange.sort(dueCol + 1);
}

function sortPriority() {
  var sa = SpreadsheetApp.getActiveSpreadsheet();
  var ss = sa.getSheetByName("Current");
  
  var data = ss.getDataRange().getValues();
  var header = data[0];
  var pCol = [parseInt(header.indexOf("Priority"))+1, parseInt(header.indexOf("Priority 2"))+1];
  
  var sortRange = ss.getRange(2, 1, data.length - 1, data[0].length);
  sortRange.sort(pCol);
}

function moveDaily() {
  var sa = SpreadsheetApp.getActiveSpreadsheet();
  var current = sa.getSheetByName("Current");
  var daily = sa.getSheetByName("Daily");
  
  var currentData = current.getDataRange().getValues();
  var dailyData = daily.getDataRange().getValues();
  
  dailyData.reverse();
  dailyData.pop();
  dailyData.reverse();
  
  current.getRange(currentData.length + 1, 1, dailyData.length, dailyData[0].length).setValues(dailyData);
  sortPriority();
}

function moveProjects() {
  var sa = SpreadsheetApp.getActiveSpreadsheet();
  var current = sa.getSheetByName("Current");
  var projects = sa.getSheetByName("Projects");
  
  var currentData = current.getDataRange().getValues();
  var projectData = projects.getDataRange().getValues();
  
  var move = [];
  var keep = [];
  var projectCol = currentData[0].indexOf("Project");
  var i = 1;
  keep.push(currentData[0]);
  while(i < currentData.length) {
    if(currentData[i][projectCol] !== "") {
      move.push(currentData[i]);
    } else {
      keep.push(currentData[i]);
    }
    i++;
  }
  
  if(move.length > 0) {
    projects.getRange(projectData.length+1, 1, move.length, move[0].length).setValues(move);
  }

  current.clearContents();
  current.getRange(1, 1, keep.length, keep[0].length).setValues(keep);
}

function moveNA() {
  var sa = SpreadsheetApp.getActiveSpreadsheet();
  var projects = sa.getSheetByName("Projects");
  var current = sa.getSheetByName("Current");
  
  var currentData = current.getDataRange().getValues();
  var projectData = projects.getDataRange().getValues();
  
  var move = [];
  var keep = [];
  var rankCol = projectData[0].indexOf("Project Rank");
  var waitingCol = projectData[0].indexOf("Waiting For");
  keep.push(projectData[0]);
  var i = 1;
  while(i < projectData.length) {
    if(projectData[i][rankCol] == "NA" && projectData[i][waitingCol] =="") {
      move.push(projectData[i]);
    } else {
      keep.push(projectData[i]);
    }
    i++;
  }
  
  if(move.length > 0) {
    current.getRange(currentData.length+1, 1, move.length, move[0].length).setValues(move);
  }

  projects.clearContents();
  projects.getRange(1, 1, keep.length, keep[0].length).setValues(keep);
}

function futureToProjects() {
  var sa = SpreadsheetApp.getActiveSpreadsheet();
  var future = sa.getSheetByName("Future");
  var projects = sa.getSheetByName("Projects");
  
  var futureData = future.getDataRange().getValues();
  var projectsData = projects.getDataRange().getValues();
  
  var move = [];
  var keep = [];
  var projectCol = futureData[0].indexOf("Project");
  var futureCol = futureData[0].indexOf("Future");
  var nameCol = projectsData[0].indexOf("Name") + 1;
  keep.push(futureData[0]);
  var i = 1;
  while(i < futureData.length) {
    if(futureData[i][futureCol] == "" && futureData[i][projectCol] != "") {
      move.push(futureData[i]);
    } else {
      keep.push(futureData[i]);
    }
    i++;
  }
  
  if(move.length > 0) {
    projects.getRange(projectsData.length+1, nameCol, move.length, move[0].length).setValues(move);
  }
  
  //projects.sort({column: 10, ascending: true}, {column:11, ascending: true});
  
  future.clearContents();
  future.getRange(1, 1, keep.length, keep[0].length).setValues(keep);
}