//generic tics helper functions

function getCurrentWeekNumber() {
  var date = new Date();
  var weekNumber = Utilities.formatDate(date, "GMT", "w");
  //Logger.log(weekNumber);
  return weekNumber;
}

function getTodayDate() {
  var today = new Date();
  //Logger.log(today);
  var todayFormatted = Utilities.formatDate(today, "GMT+1", "yyyyMMdd");
  //Logger.log(todayFormatted);
  return todayFormatted;
}

function getSixDaysAgoDate() {
  var today = new Date();
  var sixDaysAgo = new Date(today.getTime() - 6 * 24 * 60 * 60 * 1000);
  var sixDaysAgoFormatted = Utilities.formatDate(sixDaysAgo, "GMT+1", "yyyyMMdd");
  //Logger.log(sixDaysAgoFormatted);
  return sixDaysAgoFormatted;
}

function padNumber(num, length) {
  var numString = num.toString();
  while (numString.length < length) {
    numString = "0" + numString;
  }
  return numString;
}

function getRunMoment() {
  var now = new Date();
  var year = now.getFullYear();
  var month = padNumber(now.getMonth() + 1, 2);
  var day = padNumber(now.getDate(), 2);
  var hours = padNumber(now.getHours(), 2);
  var minutes = padNumber(now.getMinutes(), 2);
  var seconds = padNumber(now.getSeconds(), 2);
  return year + month + day + hours + minutes + seconds;
}

function getDateOfLastWeekday(weekday) {
  var today = new Date();
  var daysSinceLastWeekday = today.getDay() - weekday;
  if (daysSinceLastWeekday < 0) {
    daysSinceLastWeekday += 7;
  }
  var lastWeekday = new Date(today.getTime() - daysSinceLastWeekday * 24 * 60 * 60 * 1000);
  var date = lastWeekday.getFullYear() + padNumber(lastWeekday.getMonth() + 1, 2) + padNumber(lastWeekday.getDate(), 2);
  return date;
}

function getDayOfWeek(dateString) {
  var year = dateString.substr(0, 4);
  var month = parseInt(dateString.substr(4, 2), 10) - 1;
  var day = dateString.substr(6, 2);
  var date = new Date(year, month, day);
  return date.getDay();
}

function getDateForWeekday(weekdayNum) {
  //most recent occurence of that weekday
  var today = new Date();
  var daysSinceLastWeekday = today.getDay() - weekdayNum;
  if (daysSinceLastWeekday < 0) {
    daysSinceLastWeekday += 7;
  }
  var targetDate = new Date(today.getTime() - daysSinceLastWeekday * 24 * 60 * 60 * 1000);
  var year = targetDate.getFullYear();
  var month = targetDate.getMonth() + 1;
  var day = targetDate.getDate();
  return padNumber(year, 4) + padNumber(month, 2) + padNumber(day, 2);
}

function extractEmail(address) {
  var regex = /[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}/i;
  var matches = address.match(regex);
  if (matches && matches.length > 0) {
    return matches[0].toLowerCase();
  }
  return '';
}
