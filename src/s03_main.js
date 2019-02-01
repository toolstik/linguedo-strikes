var DASHBOARD_SHEET_NAME = 'dashboard';

var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
var dashBoardSheet = new SheetProxy(DASHBOARD_SHEET_NAME);
var csvSheet = new SheetProxy('points_journal');
var holidaysSheet = new SheetProxy('global_holidays');
var vacationsSheet = new SheetProxy('student_holidays');

var params = new function () {
  var paramsSheet = new SheetProxy('params');
  var valuesCache = null;

  var values = function () {
    if (valuesCache)
      return valuesCache;

    console.time('params.read');
    valuesCache = paramsSheet.getValues()
      .reduce(function (res, row, i) {
        var param = {
          rowIndex: i + 2,
          value: row[1]
        };
        res[row[0]] = param;
        return res;
      }, {});
    console.timeEnd('params.read');
    return valuesCache;
  }

  this.get = function (name) {
    var item = values()[name];
    return item ? item.value : null;
  }

  this.set = function (name, value) {
    var item = values()[name];
    if (!item) return null;
    item.value = value;
    paramsSheet.sheet.getRange(item.rowIndex, 2).setValue(value);
    paramsSheet.resetValues();
    return value;
  }
}

var dashBoardProxy = new function () {
  var data = null;

  this.get = function () {
    return data || (data = getDashBoard());
  }

  this.set = function (newState) {
    saveDashBoard(newState);
    data = newState;
    dashBoardSheet.resetValues();
  }

  function getDashBoard() {
    console.time('getDashBoard');
    var result = dashBoardSheet
      .getValues()
      .map(function (values, i) {
        return {
          rowIndex: i + 1,
          firstName: values[1],
          lastName: values[2],
          email: values[3],
          username: values[4],
          memriseStrike: parseInt(values[5], 10) || 0,
          audioStrike: parseInt(values[6], 10) || 0,
          quizStrike: parseInt(values[7], 10) || 0,
          deductedStrikes: parseInt(values[9], 10) || 0,
          deductedManually: parseInt(values[10], 10) || 0,
          startDate: values[11] || null,
          endDate: values[12] || null,
          vacationsTaken: parseInt(values[14], 10) || 0,
          lastNotified: values[16],
          lastStrikesModified: values[17]
        };
      });
    console.timeEnd('getDashBoard');
    return result;
  }

  function saveDashBoard(dashboard) {
    console.time('saveDashBoard');

    var vacationLimit = getParamInt('vacation-limit-days');
    var rangeValues = [[], [], [], [], [], [], []];
    dashboard
      .forEach(function (r) {
        rangeValues[0].push([
          r.memriseStrike,
          r.audioStrike,
          r.quizStrike
        ]);

        rangeValues[1].push([
          '=RC[-3]+RC[-2]+RC[-1]'
        ]);

        rangeValues[2].push([
          r.deductedStrikes,
          r.deductedManually,
          r.startDate,
          r.endDate
        ]);

        rangeValues[3].push([
          '=RC[-5]+RC[-4]+RC[-3]'
        ]);

        rangeValues[4].push([
          r.vacationsTaken
        ]);

        rangeValues[5].push([
          '=' + vacationLimit + '-RC[-1]'
        ]);

        rangeValues[6].push([
          r.lastNotified,
          r.lastStrikesModified
        ]);
      });

    var dashboardRange = dashBoardSheet.getDataRange();
    var dashboardRangeNumRows = dashboardRange.getNumRows();
    dashboardRange.offset(0, 5, dashboardRangeNumRows, 3).setValues(rangeValues[0]);
    dashboardRange.offset(0, 8, dashboardRangeNumRows, 1).setFormulasR1C1(rangeValues[1]);
    dashboardRange.offset(0, 9, dashboardRangeNumRows, 4).setValues(rangeValues[2]);
    dashboardRange.offset(0, 13, dashboardRangeNumRows, 1).setFormulasR1C1(rangeValues[3]);
    dashboardRange.offset(0, 14, dashboardRangeNumRows, 1).setValues(rangeValues[4]);
    dashboardRange.offset(0, 15, dashboardRangeNumRows, 1).setFormulasR1C1(rangeValues[5]);
    dashboardRange.offset(0, 16, dashboardRangeNumRows, 2).setValues(rangeValues[6]);

    console.timeEnd('saveDashBoard');
  }
}

function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu('Linguedo')
    .addItem('Deduct selected', 'deductSelected')
    .addSeparator()
    .addItem('Load new CSV', 'loadNewCsv')
    .addItem('Calculate new Strikes', 'calcMemriseStrikes')
    .addItem('Deduct all Strikes', 'deductAllStrikes')
    .addItem('Send emails', 'sendAllEmails')
    .addSeparator()
    .addItem('Do everything needed', 'doAuto')
    .addToUi();
}

function onEdit(e) {
  updateStrikeDate(e);
}

function doAuto() {
  console.time('doAuto');
  loadNewCsv();

  updateUsedVacations();

  var today = new Date();

  if (today.getDay() == 1)
    calcMemriseStrikes();

  if (today.getDate() == 1)
    deductAllStrikes();

  // this function won't send mail if three is nothing to
  sendAllEmails();

  console.timeEnd('doAuto');
}

function updateStrikeDate(e) {
  var row = e.range.getRow();
  var column = e.range.getColumn();
  var sheet = e.range.getSheet();

  if (sheet.getName() != DASHBOARD_SHEET_NAME)
    return;

  if (column < 6 || column > 8)
    return;

  sheet.getRange(row, 18).setValue(new Date());
}

function deductSelected() {
  var cell = spreadSheet.getSelection().getCurrentCell();
  var row = cell.getRow();

  if (row < 2)
    return;

  var dashboard = dashBoardProxy.get();

  if (row > dashboard.length)
    return;

  var item = dashboard[row - 2];

  if (!item.memriseStrike)
    return;

  var ui = SpreadsheetApp.getUi();

  var result = ui.alert(
    'Are you sure you want to deduct Memrise strike for ' + item.firstName + ' ' + item.lastName + '?',
    ui.ButtonSet.YES_NO);

  if (result != ui.Button.YES)
    return;

  item.memriseStrike--;
  item.deductedManually++;

  dashBoardProxy.set(dashboard);
}

function getParam(name) {
  return params.get(name);
}

function getParamInt(name) {
  return parseInt(getParam(name), 10);
}

function setParam(name, value) {
  return params.set(name, value);
}

function loadNewCsv() {
  console.time('loadNewCsv');
  var inputFolderId = getParam('memrise-input-folder-id');
  var csvFileName = getParam('memrise-file-name');
  var lastRunDate = getParam('memrise-file-last-date');
  var pointsThreshold = getParam('memrise-failure-threshold');
  var memriseFailureDepth = getParamInt('memrise-failure-depth');
  var totalsMode = getParam('memrise-file-totals-mode');
  var weekTotalsMode = totalsMode == 'week';

  var snapshot = null;
  var newValues = [];

  // move to local func, collect files into array and save at once
  function readCsv(date, file) {
    console.time('readCsv');

    var body = file.getBlob().getDataAsString();
    var data = Utilities.parseCsv(body, ',').slice(1);

    snapshot = snapshot || getPointTotals(date.addDays(-memriseFailureDepth), date);

    var result = data
      .map(function (i) {
        var username = i[1];
        var totalPoints = parseInt(i[2], 10);
        var uid = i[3];
        var currentPoints;

        if (weekTotalsMode && date.getDay() == 1) {
          currentPoints = totalPoints;
        } else {
          currentPoints = snapshot[username] ? totalPoints - snapshot[username].latest : null;
        }

        var failure = currentPoints != null &&
          currentPoints >= 0 && (currentPoints < pointsThreshold)
          ? 1
          : 0;

        return [date, username, totalPoints, uid, currentPoints, failure];
      });

    snapshot = mergePointTotals(snapshot, result);
    newValues = newValues.concat(result);

    console.timeEnd('readCsv');
  }

  var inputFolder = DriveApp.getFolderById(inputFolderId);
  var dateFolders = inputFolder.getFolders();

  var array = [];

  while (dateFolders.hasNext()) {
    var dateFolder = dateFolders.next();
    var date = parseFolderDate(dateFolder.getName());

    if (!lastRunDate || date > lastRunDate) {
      array.push({
        date: date,
        folder: dateFolder
      });
    }
  }

  var lastDate = null;

  array
    .sort(function (x, y) {
      if (x.date == y.date)
        return 0;

      return x.date < y.date ? -1 : 1;
    })
    .forEach(function (i) {
      var files = i.folder.getFilesByName(csvFileName);
      if (files.hasNext()) {
        var file = files.next();
        readCsv(i.date, file);
        lastDate = i.date;
      }

      onFolderProcessed(i.folder);
    });

  if (lastDate && newValues && newValues.length) {
    csvSheet.append(newValues);
    setParam('memrise-file-last-date', lastDate);
  }

  console.timeEnd('loadNewCsv');
}

function parseFolderDate(folderName) {
  var parts = folderName.split('-').map(function (i) { return parseInt(i, 10) });
  return new Date(parts[2], parts[1] - 1, parts[0]);
}

function onFolderProcessed(folder) {
  // delete?
}

function getInitPointTotals() {
  console.time('getInitPointTotals');
  var result = dashBoardProxy.get()
    .reduce(function (res, next) {
      if (next.username) {
        res[next.username] = {
          date: null,
          latest: null,
          failures: 0,
          startDate: next.startDate,
          endDate: next.endDate
        };
      }
      return res;
    }, {});

  console.timeEnd('getInitPointTotals');
  return result;
}

function mergePointTotals(current, values) {
  console.time('mergePointTotals');
  current = current || getInitPointTotals();

  if (!values || !values.length)
    return current;

  var result = values
    .reduce(function (state, row) {
      var date = row[0];

      if (!date)
        return state;

      var username = row[1];

      var userSnapshot = state[username];

      if (!userSnapshot)
        return state;

      var totalPoints = parseInt(row[2], 10);

      var dateFilter =
        ((userSnapshot.startDate || date) <= date && (userSnapshot.endDate || date) >= date);
      var failure = parseInt(row[5], 10) && dateFilter ? 1 : 0;

      if (userSnapshot.date == null || date > userSnapshot.date) {
        userSnapshot.latest = totalPoints;
        userSnapshot.date = date;
      }

      userSnapshot.failures = userSnapshot.failures + failure;

      return state;
    }, current);

  console.timeEnd('mergePointTotals');

  return result;
}

function getPointTotals(fromDate, toDate) {

  var init = getInitPointTotals();

  var values = csvSheet.getValues()
    .filter(function (row) {
      var date = row[0];
      return !(!date || date < fromDate || date >= toDate)
    });

  return mergePointTotals(init, values);
}

function calcMemriseStrikes() {
  console.time('calcMemriseStrikes');

  var extraDaysOffCount = getParamInt('extra-days-off-per-week') || 0;
  var lastRun = getParam('memrise-strike-last-date');
  var today = new Date().removeTime();

  if (lastRun && lastRun >= getSunday(today).addDays(-7))
    throw new Error("It is forbidden to run strike calculation more than once a week. " +
      "To force this action edit value of parameter 'memrise-strike-last-date' on 'params' tab");

  //set lastRun to Sunday
  if (!lastRun)
    lastRun = getSunday(today).addDays(-14);
  else
    lastRun = getSunday(lastRun);

  var dashboard = dashBoardProxy.get();

  while (lastRun < today) {
    var weekStart = lastRun.addDays(1);
    var weekEnd = weekStart.addDays(7);

    var daysOff = getDaysOff(weekStart, weekEnd);
    var holidays = getHolidays(weekStart, weekEnd);
    daysOff = daysOff.concat(holidays).distinct();
    var vacations = getVacations(weekStart, weekEnd);
    var points = getPointTotals(weekStart, weekEnd);

    for (var i = 0; i < dashboard.length; i++) {
      var user = dashboard[i];
      var userDaysOff = daysOff;

      if (vacations[user.email])
        userDaysOff = userDaysOff.concat(vacations[user.email]).distinct();

      var daysOffCount = userDaysOff.length;

      if (points[user.username]) {
        var strikeCount = points[user.username].failures - daysOffCount - extraDaysOffCount;

        if (strikeCount > 0) {
          user.memriseStrike += strikeCount;
          user.lastStrikesModified = new Date();
        }
      }
    }

    dashBoardProxy.set(dashboard);
    setParam('memrise-strike-last-date', lastRun);

    lastRun = lastRun.addDays(7);
  }

  console.timeEnd('calcMemriseStrikes');
}

function getDaysOff(fromDate, toDate) {
  if (toDate <= fromDate)
    return [];
  var result = [];
  for (var d = fromDate; d < toDate; d = d.addDays(1)) {
    if (d.getDay() == 0 || d.getDay() == 6)
      result.push(d);
  }
  return result;
}

function getHolidays(fromDate, toDate) {
  console.time('getHolidays');
  if (toDate <= fromDate)
    return [];

  var range = holidaysSheet.getValues();
  var result = range.map(function (i) { return i[0]; })
    .filter(function (i) { return i >= fromDate && i < toDate; })
    .distinct();
  console.time('getHolidays');
  return result;
}

function getVacations(fromDate, toDate) {
  console.time('getVacations');

  if (toDate <= fromDate)
    return [];

  var vacationLimit = getParamInt('vacation-limit-days');

  var getAll = function (from, to) {
    var daysOff = getDaysOff(from, to);
    var holidays = getHolidays(from, to);
    daysOff = daysOff.concat(holidays).distinct();
    daysOff = daysOff.map(function (i) { return i.getTime(); });

    var result = {};

    var rangeValues = vacationsSheet.getValues();

    for (var i = 0; i < rangeValues.length; i++) {
      var values = rangeValues[i];
      var email = values[1];
      var date = values[2];
      var days = parseInt(values[3], 10);

      if (!date || date >= to || date.addDays(days) < from) continue;

      for (var d = 0; d < days; d++) {
        var vacDate = date.addDays(d);

        if (vacDate >= from && vacDate < to && daysOff.indexOf(vacDate.getTime()) < 0) {
          if (!result[email])
            result[email] = [];

          result[email].push(vacDate);
        }
      }
    }

    console.timeEnd('getVacations');
    return result;
  }

  var monthVacations = function (d) {
    var monthStart = d.addDays(1 - d.getDate());
    var monthEnd = monthStart.addMonths(1);

    var all = getAll(monthStart, monthEnd);

    for (var user in all) {
      var dates = all[user];

      if (dates.length > vacationLimit) {
        all[user] = dates.distinct().sort(function (x, y) { return x - y; }).slice(0, vacationLimit);
      }
    }


    return all;
  }

  var months = [];

  for (var curMon = fromDate.addDays(1 - fromDate.getDate()); curMon < toDate; curMon = curMon.addMonths(1)) {
    months.push(monthVacations(curMon));
  }

  var significantVacations = months.reduce(function (x, y) {
    for (user in y) {
      if (x[user] == undefined)
        x[user] = y[user];
      else
        x[user] = x[user].concat(y[user]);
    }
    return x;
  });

  var result = {};

  for (var user in significantVacations) {
    var dates = significantVacations[user].filter(function (i) { return i >= fromDate && i < toDate });

    if (dates.length > 0)
      result[user] = dates;
  }

  return result;
}

// returns sunday of the same week
function getSunday(date) {
  return date.addDays(6 - (date.getDay() + 6) % 7);
}

function monthDiff(d1, d2) {
  var months;
  months = (d2.getFullYear() - d1.getFullYear()) * 12;
  months += d2.getMonth() - d1.getMonth();
  return months;
}

function deductAllStrikes() {
  console.time('deductAllStrikes');

  var lastRunDate = getParam('strike-deduction-last-date');
  var currentDate = new Date();

  if (lastRunDate && monthDiff(lastRunDate, currentDate) < 1)
    throw new Error("It is forbidden to run strike deduction more than once a month. " +
      "To force this action edit value of parameter 'strike-deduction-last-date' on 'params' tab");

  var dashboard = dashBoardProxy.get();

  for (i = 0; i < dashboard.length; i++) {
    var strike = dashboard[i];
    deductStrike(strike);
  }

  dashBoardProxy.set(dashboard);
  setParam('strike-deduction-last-date', currentDate);

  console.timeEnd('deductAllStrikes');
}

function deductStrike(strike) {
  var total = 0;

  if (strike.memriseStrike > 0) {
    strike.memriseStrike--;
    total++;
  }

  if (strike.audioStrike > 0) {
    strike.audioStrike--;
    total++;
  }

  if (strike.quizStrike > 0) {
    strike.quizStrike--;
    total++;
  }

  strike.deductedStrikes += total;

  return strike;
}

function evaluateTemplate(template, obj) {
  var regex = /\$\{(.+?)\}/g;

  var result = template.replace(regex, function (match, name, offset, string) {
    if (obj[name] == undefined)
      return match;
    return obj[name];
  });

  return result;
}

function sendEmail(template, user) {
  var body = evaluateTemplate(template, user);

  var message = {
    to: user.email,
    subject: 'Linguedo weekly statistics',
    htmlBody: body,
  };

  var quota = MailApp.getRemainingDailyQuota();

  if (quota < 1)
    return false;

  MailApp.sendEmail(message);
  return true;
}

function updateUsedVacations() {
  console.time('updateUsedVacations');

  var dashboard = dashBoardProxy.get();
  var today = new Date().removeTime();
  var monthVacations = getVacations(today.addDays(1 - today.getDate()), today.addDays(1 - today.getDate()).addMonths(1));
  for (var i = 0; i < dashboard.length; i++) {
    var user = dashboard[i];
    user.vacationsTaken = monthVacations[user.email] ? monthVacations[user.email].filter(function (i) { return i < today; }).length : 0;
  }

  dashBoardProxy.set(dashboard);

  console.timeEnd('updateUsedVacations');

  return dashboard;
}

function sendAllEmails() {
  console.time('sendAllEmails');

  var enabled = getParamInt('email-enabled');

  if (!enabled)
    return false;

  var today = new Date().removeTime();
  var weekEnd = getSunday(today).addDays(-6);
  var weekStart = weekEnd.addDays(-7);

  var dashboard = dashBoardProxy.get();
  var weekVacations = getVacations(weekStart, weekEnd);
  var vacationLimit = getParamInt('vacation-limit-days');

  var templateId = getParam('email-template-file-id');
  var template = DriveApp.getFileById(templateId).getBlob().getDataAsString();

  for (var i = 0; i < dashboard.length; i++) {
    var user = dashboard[i];
    if (!user.email) continue;

    if (user.lastNotified && getSunday(user.lastNotified) >= getSunday(today)) continue;

    user['daysOffUsedLastWeek'] = weekVacations[user.email] ? weekVacations[user.email].length : 0;
    user['daysOffLeft'] = vacationLimit - user.vacationsTaken;

    var strikesModified = !user.lastStrikesModified ||
      (user.lastStrikesModified >= weekStart && user.lastStrikesModified < weekEnd);
    var vacationsAquire = user.daysOffUsedLastWeek > 0;

    if (!strikesModified && !vacationsAquire) {
      continue;
    }

    var sent = sendEmail(template, user);
    if (!sent) break;

    user.lastNotified = today;
  }

  dashBoardProxy.set(dashboard);

  console.timeEnd('sendAllEmails');

  return dashboard;
}
