function emptySubmissionHandler(message) {
  //probeer uit te vogelen voor wie dit is
  var body = message.getBody().toLowerCase();
  var lines = body.split("\n");
  var cleanedName = "";
  for (var i = 0; i < lines.length; i++) {
    if (lines[i].indexOf("@cscportal.onmicrosoft.com") !== -1) {
      var nameLine = lines[i];
      cleanedName = nameLine.replace(/<[^>]+>/g, "").replace(/ /g, "").replace(/cc:/, "").replace(/\d+/g, '');
      cleanedName = cleanedName.split("@")[0];
      break;
    }
  }
  var sheetId = PropertiesService.getScriptProperties().getProperty('sheetHandle');
  var sheet = SpreadsheetApp.openById(sheetId).getSheetByName("users");
  var data = sheet.getDataRange().getValues();
  var nameSearchString = cleanedName.substring(3); //3 zou genoeg moeten zijn
  for (var i = 0; i < data.length; i++) {
    if (data[i][1].indexOf(nameSearchString) !== -1) {
      var userId = data[i][0];
      var email = data[i][1];
      break;
    }
  }
  Logger.log("cleanedName: " + cleanedName);
  Logger.log("userId: " + userId);
  Logger.log("email: " + email);

  //rerun creation en submission.
  submitTics(user, email);
}
