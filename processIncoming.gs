/*
processUnreadEmails draait met time trigger (1 minuut).
Verwerkt emails met ticsfile als attachment.
Slaat de attachments op en voegt gebruiker toe aan sheet, users.
*/

function processUnreadEmails() {
  var threads = GmailApp.search("is:unread");
  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];
      if (message.isUnread()) {
        processEmail(message);
        message.markRead();
      }
    }
  }
}

function processEmail(message) {
  // Check if message contains 'stop tics' and call processUnsubscriber
  var body = message.getPlainBody();
  if (body.toLowerCase().indexOf('stop tics') !== -1) {
    processUnsubscriber(message);
    return;
  }

  var attachments = message.getAttachments();
  for (var i = 0; i < attachments.length; i++) {
    var attachment = attachments[i];
    if (attachment.getName().match(/\.W\d{2}$/i)) {
      processSubscriber(message, attachment);
      return;
    }
  }
  message.markUnread();
}

function processUnsubscriber(message) {
  var sender = extractEmail(message.getFrom());
  var sheet = SpreadsheetApp.openById({sheet-id}).getSheetByName("users");

  // Check if the sender email is in the sheet
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === sender) {
      // Remove the row from the sheet
      sheet.deleteRow(i + 1);
      break;
    }
  }

  // Send confirmation email
  var body = "You have terminated the TicsTok tics submission service.\nTo start using it again, simply send an email with a ticsfile attached to abigfatsmiley@gmail.com.";
  MailApp.sendEmail(sender, "TicsTok service termination confirmation", body);
}


function processSubscriber(message, attachment) {
  var fileId = saveAttachmentToDrive(attachment);
  var sender = message.getFrom();
  var fileName = attachment.getName().split(".")[0];
  var sheet = SpreadsheetApp.openById({sheet-id}).getSheetByName("users");
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
      replyMessage = 'Thanks for the update.\n';
      break;
    } else {
      replyMessage = 'Welcome to the TicsTok tics submission service.\nNever miss the deadline!\n\n';
    }
  }
  if (!emailFound) {
    sheet.appendRow([name, email]);
  }
  message.markRead();
  replyMessage += "\nI will schedule a tics submission for you every fridayafternoon from now on, based on the tics file(s) you have sent me.\nTo stop using this service, simply reply or send an email containing 'stop tics' in the message body to abigfatsmiley@gmail.com.\nTo update your existing time allocation preferences, simply send a newer ticsfile to the same address.";
  GmailApp.sendEmail(email, "Tics Submission Scheduled", replyMessage);
}

function saveAttachmentToDrive(attachment) {
  var file = DriveApp.createFile(attachment);
  return file.getId();
}


function addToTicsTokSheet(sender, fileName) {
  var sheet = SpreadsheetApp.openById({sheet-id}).getSheetByName("users");
  var row = [String(fileName), extractEmail(sender)];
  sheet.appendRow(row);
}
