/*
processUnreadEmails draait met time trigger (1 minuut).
Verwerkt emails met ticsfile als attachment.
Slaat de attachments op en voegt gebruiker toe aan sheet, users.
*/

function processUnreadEmails() {
  var threads = GmailApp.search("is:unread");
  Logger.log("Processing %s unread mails", threads.length.toFixed(0));
  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];
      if (message.isUnread()) {
        Logger.log("processUnreadEmails: index:" + j);
        processEmail(message);
        message.markRead();
      }
    }
  }
}

function processEmail(message) {
  var recognized = false;

  // Check if message contains 'stop tics' and call processUnsubscriber
  var body = message.getPlainBody();
  if (body.toLowerCase().indexOf('stop tics') !== -1 || message.getSubject().toLowerCase().indexOf('stop tics') !== -1) {
    recognized = true;
    Logger.log("recognized: stop tics");
    processUnsubscriber(message);
    return;
  }

  // Check if message contains 'f126' and call reportBiweeklyF126Totals
  if (body.toLowerCase().indexOf('f126') !== -1 || message.getSubject().toLowerCase().indexOf('f126') !== -1) {
    recognized = true;
    Logger.log("recognized: f126");
    reportBiweeklyF126Totals(message);
    return;
  }

  // Check if we have an error report for our submission
  if (message.getSubject().indexOf("Automatically reply from TICS") !== -1) {
    recognized = true;
    Logger.log("recognized: Automatically reply from TICS");
    emptySubmissionHandler(message);
    return;
  }

  var attachments = message.getAttachments();
  for (var i = 0; i < attachments.length; i++) {
    var attachment = attachments[i];
    if (attachment.getName().match(/\.W\d{2}$/i)) {
      Logger.log("recognized: tics file attached");
      recognized = true;
      processSubscriber(message, attachment);
      return;
    }
  }

  if (recognized !== true) {
    Logger.log("Not recognized.");
    processUnrecognized(message);
  }

  message.markUnread();
}

function processUnsubscriber(message) {
  var sender = extractEmail(message.getFrom());
  var scriptProperties = PropertiesService.getScriptProperties();
  var sheetId = scriptProperties.getProperty('sheetHandle');
  var sheet = SpreadsheetApp.openById(sheetId).getSheetByName("users");

  // Check if the sender email is in the sheet
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === sender) {
      //remove files
      deleteTicsFilesForUser(data[i][0]);
      // Remove the row from the sheet
      sheet.deleteRow(i + 1);
      break;
    }
  }

  // Send confirmation email
  var body = "You have terminated the TicsTok tics submission service.\n"
  body += "All your data was erased. You have ceased to exist in this system.\n\n"
  body += "To start using it again, simply send an email with a ticsfile attached to abigfatsmiley@gmail.com.";
  MailApp.sendEmail(sender, "TicsTok service termination confirmation", body);
}

function deleteTicsFilesForUser(user) {
  var folder = DriveApp.getFoldersByName("ticstok_received").next();
  var query = "title contains '" + padNumber(user, 8) + ".W'";
  var files = folder.searchFiles(query);
  while (files.hasNext()) {
    var file = files.next();
    file.setTrashed(true);
  }
  var folder = DriveApp.getFoldersByName("ticstok_work").next();
  var query = "title contains '" + padNumber(user, 8) + ".W'";
  var files = folder.searchFiles(query);
  while (files.hasNext()) {
    var file = files.next();
    file.setTrashed(true);
  }
}

function processSubscriber(message, attachment) {
  var fileId = saveAttachmentToDrive(attachment, "ticstok_received");
  var sender = message.getFrom();
  var fileName = attachment.getName().split(".")[0];
  var scriptProperties = PropertiesService.getScriptProperties();
  var sheetId = scriptProperties.getProperty('sheetHandle');
  var sheet = SpreadsheetApp.openById(sheetId).getSheetByName("users");
  var lastRow = sheet.getLastRow();
  var emailColumn = 2;
  var nameColumn = 1;
  var email = extractEmail(sender);
  var name = fileName;
  var replyMessage = '';
  var emailFound = false;
  for (var i = 2; i <= lastRow; i++) {
    if (sheet.getRange(i, emailColumn).getValue().toLowerCase() == email) {
      name = sheet.getRange(i, nameColumn).getValue();
      emailFound = true;
      replyMessage = 'Thanks for the update.\n\n';
      break;
    } else {
      replyMessage = 'Welcome to the TicsTok tics submission service.\nNever miss the deadline!\n\n';
    }
  }
  if (!emailFound) {
    sheet.appendRow([name, email]);
  }
  message.markRead();
  replyMessage += "I will submit a valid ticsfile you every fridayafternoon from now on, based on the tics file(s) you have sent me.\n\n"
  replyMessage += "To stop this service, reply or send an email containing 'stop tics' in the body or subject to abigfatsmiley@gmail.com.\n"
  replyMessage += "To update your existing time allocation preferences, simply send a newer ticsfile to the same address.";
  GmailApp.sendEmail(email, "Tics Submission Scheduled", replyMessage);
}

function saveAttachmentToDrive(attachment) {
  //remove if exists otherwise we end up with 2 versions of the same file in google drive
  var folderName = 'ticstok_received';
  var folder = DriveApp.getFoldersByName(folderName).next();
  var fileName = attachment.getName();
  var existingFiles = folder.getFilesByName(fileName);
  while (existingFiles.hasNext()) {
    var file = existingFiles.next();
    folder.removeFile(file); //make sure it never appears in searches
  }
  var file = folder.createFile(attachment);
  return file.getId();
}

function addToTicsTokSheet(sender, fileName) {
  var scriptProperties = PropertiesService.getScriptProperties();
  var sheetId = scriptProperties.getProperty('sheetHandle');
  var sheet = SpreadsheetApp.openById(sheetId).getSheetByName("users");
  var row = [String(fileName), extractEmail(sender)];
  sheet.appendRow(row);
}

function processUnrecognized(message) {
  //unrecognized, forwarden dan maar
  Logger.log("Unrecognized email");
  var adminEmailAddress = PropertiesService.getScriptProperties().getProperty('adminEmailAddress');
  Logger.log("adminEmailAddress: " + adminEmailAddress);
  var forwardedMessage = message.forward(adminEmailAddress, {subject: "Fwd: " + message.getSubject()});
  Logger.log("sending...");
  GmailApp.sendEmail(adminEmailAddress, forwardedMessage.getSubject(), forwardedMessage.getBody());
}


