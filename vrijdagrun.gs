/*
runFridayBatch draait iedere vrijdagmiddag (om ca. 1600).
Maakt en verstuurt de ticsfiles.
*/

function runFridayBatch() {
  //loop door users en kijk of file gesubmit moet worden
  Logger.log('runFridayBatch');
  var sheet = SpreadsheetApp.openById({sheet-id}).getSheetByName("users");
  var numRows = sheet.getLastRow();
  var todayDate = getTodayDate();
  for (var i = 2; i <= numRows; i++) {
    var user = sheet.getRange(i, 1).getValue();
    var email = sheet.getRange(i, 2).getValue();
    Logger.log('user %s, email %s', user, email);
    var submissionDate = sheet.getRange(i, 3).getValue();
    if (submissionDate != todayDate) {
      Logger.log('submissionDate != todayDate');
      submitTics(user, email);
    } else {
      Logger.log('skip, already sent');
    }
  }
}


function submitTics(user, email) {
  //file maken, opsturen en bijhouden dat hij gestuurd is
  Logger.log('submitTics for %s, %s', user, email);
  var ticsFile = createTicsFileForUser(user);
  GmailApp.sendEmail(
    "{tics-file-receiver-address}",
    "Tics file submission",
    "Tics file submission",
    {
      attachments: [ticsFile],
      cc: email
    }
  );
  //file sent, now update timestamp
  var sheet = SpreadsheetApp.openById({sheet-id}).getSheetByName("users");
  var numRows = sheet.getLastRow();
  for (var i = 2; i <= numRows; i++) {
    if (sheet.getRange(i, 2).getValue() == email) {
      sheet.getRange(i, 3).setValue(getTodayDate());
      break;
    }
  }
}


function createTicsFileForUser(user) {
  //ticsfile maken
  Logger.log('createTicsFileForUser for %s', user);
  var rootFolder = DriveApp.getRootFolder();
  var fileName = padNumber(user, 8) + '.W' + padNumber(getCurrentWeekNumber(), 2);
  var headerRecord = createHeaderRecord(user);
  var timeRecords = createTimeRecords(user);
  var footerRecord = createFooterRecord(user);
  var fileContent = headerRecord + timeRecords + footerRecord;
  var file = rootFolder.createFile(fileName, fileContent, "text/plain");
  return file;
}


function createTimeRecords(user) {
  /*
  Temporary solution:
  copy all timerecords from latest known ticsfile and adjust them to fit this week.
  */
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
  var newestFile = null;
  var highestNumber = null;
  var query = "title contains '" + padNumber(user, 8) + ".W'";
  var files = DriveApp.searchFiles(query);
  while (files.hasNext()) {
    var file = files.next();
    Logger.log('file ' + file.getName());
    var fileNameParts = file.getName().split(".");
    if (fileNameParts.length !== 2 || fileNameParts[1].length !== 3) {
      Logger.log('invalid name');
      continue; // Skip files with invalid names
    }
    var fileNumber = parseInt(fileNameParts[1].substr(1), 10);
    if (isNaN(fileNumber)) {
      continue; // Skip files with invalid numbers
    }
    if (highestNumber === null || fileNumber > highestNumber) {
      newestFile = file;
      highestNumber = fileNumber;
    }
  }
  Logger.log('no more files');
  if (newestFile) {
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
  var newestFile = null;
  var highestNumber = null;
  var query = "title contains '" + padNumber(user, 8) + ".W'";
  var files = DriveApp.searchFiles(query);
  while (files.hasNext()) {
    var file = files.next();
    //Logger.log('file ' + file.getName());
    var fileNameParts = file.getName().split(".");
    if (fileNameParts.length !== 2 || fileNameParts[1].length !== 3) {
      continue; // Skip files with invalid names
    }
    var fileNumber = parseInt(fileNameParts[1].substr(1), 10);
    if (isNaN(fileNumber)) {
      continue; // Skip files with invalid numbers
    }
    if (highestNumber === null || fileNumber > highestNumber) {
      newestFile = file;
      highestNumber = fileNumber;
    }
  }
  //Logger.log('no more files');
  if (newestFile) {
    //Logger.log('newest file: '+newestFile.getName());
    var content = newestFile.getBlob().getDataAsString();
    var firstLine = content.split("\n")[0];
    return firstLine;
  }
  return "";
}

function readNewestTicsFileFooterForUser(user) {
  var newestFile = null;
  var highestNumber = null;
  var query = "title contains '" + padNumber(user, 8) + ".W'";
  var files = DriveApp.searchFiles(query);
  while (files.hasNext()) {
    var file = files.next();
    Logger.log('file ' + file.getName());
    var fileNameParts = file.getName().split(".");
    if (fileNameParts.length !== 2 || fileNameParts[1].length !== 3) {
      Logger.log('invalid name');
      continue; // Skip files with invalid names
    }
    var fileNumber = parseInt(fileNameParts[1].substr(1), 10);
    if (isNaN(fileNumber)) {
      continue; // Skip files with invalid numbers
    }
    if (highestNumber === null || fileNumber > highestNumber) {
      newestFile = file;
      highestNumber = fileNumber;
    }
  }
  Logger.log('no more files');
  if (newestFile) {
    Logger.log('newest file: ' + newestFile.getName());
    Logger.log('newest file');
    var content = newestFile.getBlob().getDataAsString();
    var lines = content.split("\n");
    var lastLine = lines[lines.length - 2];
    return lastLine;
  }
  return "";
}






