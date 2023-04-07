function runFridayReportBatch() {
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
    Logger.log('runFridayReportBatch: processing user %s, email %s', user.toFixed(0), email);
    var submissionDate = sheet.getRange(i, 3).getValue();
    if (submissionDate == todayDate) {
      reportTicsSubmission(user, email);
    }
    Logger.log('runFridayReportBatch: done processing user %s, email %s', user.toFixed(0), email);
  }
  Logger.log('runFridayReportBatch ended');
}

function reportTicsSubmission(user, email) {
  Logger.log("reportTicsSubmission for user %s", user.toFixed(0))
  var newestFile = null;
  var highestNumber = null;
  var query = "title contains '" + padNumber(user, 8) + ".W'";
  var folder = DriveApp.getFoldersByName("ticstok_work").next();
  var files = folder.searchFiles(query);
  var prevDOW = "";
  while (files.hasNext()) {
    var file = files.next();
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
      //Logger.log('readNewestTicsFileFooterForUser: newestFile = %s, highestNumber = %s', newestFile, highestNumber);
    }
  }
  var reportText = "";
  var weekdays = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
  if (newestFile) {
    Logger.log('newest file used: ' + newestFile.getName());
    var content = newestFile.getBlob().getDataAsString();
    var lines = content.split("\n");
    //loop through lines to build report text
    for (var i = 0; i < lines.length; i++) {
      var line = lines[i];
      if (line.substring(0, 2) === 'TB') {
        //we have an activity time record
        //read date, determine what weekday that was en find the date for that weekday this week
        var dateString = line.substring(13, 21);
        var dayOfWeek = getDayOfWeek(dateString);
        if (dayOfWeek !== prevDOW) { reportText += "\n"; }
        reportText += line.substring(19, 21) + "-" + line.substring(17, 19) + "-" + line.substring(13, 17);
        reportText += " " + weekdays[dayOfWeek];
        reportText += " " + line.substring(40, 48);//projectcode
        reportText += " " + line.substring(36, 38);//uren
        reportText += " hours" + line.substring(103, 131);//comments
        reportText += "\n";
        prevDOW = dayOfWeek;
      }
    }
    Logger.log(reportText);
    // all done, now mail the report
    var mailFooter = "\n\nTo stop using this service, simply reply or send an email containing 'stop tics' in the message body to abigfatsmiley@gmail.com.\nTo update your existing time allocation preferences, simply send a newer ticsfile to the same address.\n\nNever miss the deadline!\nTo start using this service simply send an email to abigfatsmiley@gmail.com and attach a ticsfile and I will use that as a template. I will sumbit a plausible ticsfile for you every friday-afternoon, based on the tics file(s) you have sent me.\n";
    GmailApp.sendEmail(
      email,
      "Tics file submitted today",
      "Today a ticsfile was submitted containing the following time allocations:\n\n" + reportText + mailFooter
    );
  } else {
    Logger.log('Error: no ticsfile found in ticstok_work');
  }
}
