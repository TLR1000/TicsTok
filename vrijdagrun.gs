/*
runFridayBatch draait iedere vrijdagmiddag (om ca. 1600).
Maakt en verstuurt de ticsfiles.
*/

function runFridayBatch() {
  //loop door users en kijk of file gesubmit moet worden
  Logger.log('runFridayBatch started');
  var scriptProperties = PropertiesService.getScriptProperties();
  var sheetId = scriptProperties.getProperty('sheetHandle');
  var sheet = SpreadsheetApp.openById(sheetId).getSheetByName("users");
  var numRows = sheet.getLastRow();
  var todayDate = getTodayDate();
  for (var i = 2; i <= numRows; i++) {
    var user = sheet.getRange(i, 1).getValue();
    var email = sheet.getRange(i, 2).getValue();
    Logger.log('runFridayBatch: processing user %s, email %s', user.toFixed(0), email);
    var submissionDate = sheet.getRange(i, 3).getValue();
    if (submissionDate != todayDate) {
      //Logger.log('runFridayBatch submissionDate != todayDate (not submitted yet so submitting)');
      //Logger.log('submissionDate = %s, todayDate =%s', submissionDate, todayDate);
      submitTics(user, email);
    } else {
      Logger.log('skip submission, already sent');
    }
    Logger.log('runFridayBatch: done processing user %s, email %s', user.toFixed(0), email);
  }
  Logger.log('runFridayBatch ended');
}


function submitTics(user, email) {
  //file maken, opsturen en bijhouden dat hij gestuurd is
  var scriptProperties = PropertiesService.getScriptProperties();
  var recipient = scriptProperties.getProperty('ticsreceiveraddress');
  Logger.log('submitTics started for user: %s, email: %s, recipient: %s', user.toFixed(0), email, recipient);
  var ticsWorkFolder = DriveApp.getFoldersByName("ticstok_work").next();
  var fileName = padNumber(user, 8) + '.W' + padNumber(getCurrentWeekNumber(), 2);
  // Check if file with same name already exists and delete it
  var existingFile = ticsWorkFolder.getFilesByName(fileName);
  if (existingFile.hasNext()) {
    var fileToRemove = existingFile.next();
    ticsWorkFolder.removeFile(fileToRemove);
  }
  //create file
  var headerRecord = createHeaderRecord(user);
  var timeRecords = createTimeRecords(user);
  var footerRecord = createFooterRecord(user);
  var fileContent = headerRecord + timeRecords + footerRecord + "\n";
  var ticsFile = ticsWorkFolder.createFile(fileName, fileContent, "text/plain");
  GmailApp.sendEmail(
    recipient,
    "Tics file submission",
    "Tics file submission",
    {
      attachments: [ticsFile],
      cc: email
    }
  );
  //file sent, now update timestamp
  var scriptProperties = PropertiesService.getScriptProperties();
  var sheetId = scriptProperties.getProperty('sheetHandle');
  var sheet = SpreadsheetApp.openById(sheetId).getSheetByName("users");
  var numRows = sheet.getLastRow();
  for (var i = 2; i <= numRows; i++) {
    if (sheet.getRange(i, 2).getValue() == email) {
      sheet.getRange(i, 3).setValue(getTodayDate());
      //Logger.log('submitTics: timestamp in sheet where email = %s, value set is %s', email, getTodayDate());
      break;
    }
  }
}

function createTicsFileForUser(user) {
  //ticsfile maken
  Logger.log('createTicsFileForUser started for %s', user.toFixed(0));
  var folder = DriveApp.getFoldersByName("ticstok_work").next();
  var fileName = padNumber(user, 8) + '.W' + padNumber(getCurrentWeekNumber(), 2);
  var headerRecord = createHeaderRecord(user);
  var timeRecords = createTimeRecords(user);
  var footerRecord = createFooterRecord(user);
  var fileContent = headerRecord + timeRecords + footerRecord;
  var file = folder.createFile(fileName, fileContent, "text/plain");
  return file;
}

function createTimeRecords(user) {
  /*
  Temporary solution:
  copy all timerecords from latest known ticsfile and adjust them to fit this week.
  */
  Logger.log("createTimeRecords started for user %s", user.toFixed(0))
  var tbSet = "";//set of tb recors to be written in output file as a formatted block of data
  //create array with latest known timerecords
  var tbRecords = getTBRecords(user);
  //loop through the array and update records in the array
  for (var i = 0; i < tbRecords.length; i++) {
    var tbRecord = tbRecords[i];
    //modify
    //read date, determine what weekday that was en find the date for that weekday this week
    var dateString = tbRecords[i].substring(13, 21);
    var dayOfWeek = getDayOfWeek(dateString);
    //this week's date for that weekday (closest occurence)
    var newDateString = getDateForWeekday(dayOfWeek);
    tbRecord = tbRecord.substring(0, 13) + newDateString + tbRecord.substring(21);
    //add to tbSet
    tbSet = tbSet + tbRecord + "\n";
  }
  return tbSet;
}

function getTBRecords(user) {
  Logger.log("getTBRecords for user %s", user.toFixed(0))
  var newestFile = null;
  var highestNumber = null;
  var folder = DriveApp.getFoldersByName("ticstok_received").next();
  var query = "title contains '" + padNumber(user, 8) + ".W' and '" + folder.getId() + "' in parents";
  var files = folder.searchFiles(query);
  while (files.hasNext()) {
    var file = files.next();
    //Logger.log('getTBRecords: file ' + file.getName());
    var fileNameParts = file.getName().split(".");
    if (fileNameParts.length !== 2 || fileNameParts[1].length !== 3) {
      continue; // Skip files with invalid names
    }
    var fileNumber = parseInt(fileNameParts[1].substr(1), 10);
    if (isNaN(fileNumber)) {
      continue; // Skip files with invalid numbers
    }
    if (highestNumber === null || fileNumber > highestNumber) {
      //Logger.log('getTBRecords: (highestNumber === null || fileNumber > highestNumber) fileNumber = %s, highestNumber = %s', fileNumber, highestNumber);
      newestFile = file;
      highestNumber = fileNumber;
      //Logger.log('getTBRecords: newestFile = %s, highestNumber = %s', newestFile, highestNumber);
    }
  }
  if (newestFile) {
    Logger.log('getTBRecords: newest file used: ' + newestFile.getName());
    var content = newestFile.getBlob().getDataAsString();
    var lines = content.split("\n");
    var tbRecords = [];
    for (var i = 0; i < lines.length; i++) {
      var line = lines[i];
      if (line.substring(0, 2) === 'TB') {
        tbRecords.push(line);
      }
    }
    return tbRecords;
  }
  return "";
}

function createHeaderRecord(user) {
  Logger.log("createHeaderRecord for user %s", user.toFixed(0))
  //get the base record
  var record = readNewestTicsFileHeaderForUser(user);
  //Logger.log("record: " + record);
  //make it fit for this week
  record = record.substring(0, 14) + getSixDaysAgoDate() + record.substring(22);
  record = record.substring(0, 22) + getTodayDate() + record.substring(30);
  record = record.substring(0, 105) + getRunMoment() + record.substring(119);
  //Logger.log("record: " + record);
  record = record + "\n";
  return record;
}

function createFooterRecord(user) {
  Logger.log("createFooterRecord for user %s", user.toFixed(0))
  //get the base record
  var record = readNewestTicsFileFooterForUser(user);
  //Logger.log("record: " + record);
  //make it fit for this week
  record = record.substring(0, 14) + getSixDaysAgoDate() + record.substring(22);
  record = record.substring(0, 22) + getTodayDate() + record.substring(30);
  record = record.substring(0, 105) + getRunMoment() + record.substring(119);
  //Logger.log("record: " + record);
  record = record + "\n";
  record = record + "\n";//laatste regel is empty record
  return record;
}

function readNewestTicsFileHeaderForUser(user) {
  Logger.log("readNewestTicsFileHeaderForUser for user %s", user.toFixed(0))
  var newestFile = null;
  var highestNumber = null;
  var ticstokReceivedFolder = DriveApp.getFoldersByName("ticstok_received").next();
  var query = "title contains '" + padNumber(user, 8) + ".W'";
  var files = ticstokReceivedFolder.searchFiles(query);
  while (files.hasNext()) {
    var file = files.next();
    //Logger.log('readNewestTicsFileHeaderForUser: file found ' + file.getName());
    var fileNameParts = file.getName().split(".");
    if (fileNameParts.length !== 2 || fileNameParts[1].length !== 3) {
      continue; // Skip files with invalid names
    }
    var fileNumber = parseInt(fileNameParts[1].substr(1), 10);
    //Logger.log('readNewestTicsFileHeaderForUser: fileNumber:%s found, highestNumber is %s', fileNumber, highestNumber);
    if (isNaN(fileNumber)) {
      continue; // Skip files with invalid numbers
    }
    if (highestNumber === null || fileNumber > highestNumber) {
      //Logger.log('readNewestTicsFileHeaderForUser: (highestNumber === null || fileNumber > highestNumber) fileNumber = %s, highestNumber = %s', fileNumber, highestNumber);
      newestFile = file;
      highestNumber = fileNumber;
      //Logger.log('readNewestTicsFileHeaderForUser: newestFile = %s, highestNumber = %s', newestFile, highestNumber);
    }
  }
  if (newestFile) {// if found
    Logger.log('readNewestTicsFileHeaderForUser: newestFile used: %s', newestFile.getName());
    var content = newestFile.getBlob().getDataAsString();
    var firstLine = content.split("\n")[0];
    return firstLine;
  }
  return "";
}


function readNewestTicsFileFooterForUser(user) {
  Logger.log("readNewestTicsFileFooterForUser for user %s", user.toFixed(0))
  var newestFile = null;
  var highestNumber = null;
  var query = "title contains '" + padNumber(user, 8) + ".W'";
  var folder = DriveApp.getFoldersByName("ticstok_received").next();
  var files = folder.searchFiles(query);
  while (files.hasNext()) {
    var file = files.next();
    //Logger.log('readNewestTicsFileFooterForUser: file found ' + file.getName());
    var fileNameParts = file.getName().split(".");
    if (fileNameParts.length !== 2 || fileNameParts[1].length !== 3) {
      //Logger.log('readNewestTicsFileFooterForUser: invalid name');
      continue; // Skip files with invalid names
    }
    var fileNumber = parseInt(fileNameParts[1].substr(1), 10);
    if (isNaN(fileNumber)) {
      continue; // Skip files with invalid numbers
    }
    if (highestNumber === null || fileNumber > highestNumber) {
      newestFile = file;
      highestNumber = fileNumber;
      //Logger.log('readNewestTicsFileFooterForUser: newestFile = %s, highestNumber = %s', newestFile, highestNumber);
    }
  }
  if (newestFile) {
    Logger.log('readNewestTicsFileFooterForUser: newest file used: ' + newestFile.getName());
    var content = newestFile.getBlob().getDataAsString();
    var lines = content.split("\n");
    var lastLine = null;
    for (var i = 0; i < lines.length; i++) {
      if (lines[i].startsWith('EK')) {
        lastLine = lines[i];
        break;
      }
    }
    return lastLine;
  } else {
    Logger.log("Error: No newest file found");
  }
  return "";
}





