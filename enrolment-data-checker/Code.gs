/**
 * Enrolment data checker
 * ----------------------
 * Checks an enrolment sheet before a course starts and highlights problems:
 *   - required fields that are empty
 *   - invalid email addresses
 *   - duplicate learners (same email in the same course)
 * Problems are coloured in the sheet and listed in a "Checks" sheet,
 * so that the team can fix them before the course goes live.
 *
 * Expected sheet "Enrolments" (row 1 = headers):
 *   Name | Email | Phone | Course | Start date
 *
 * Adds a menu "Data checks" when the spreadsheet opens.
 */

const CHECK = {
  SHEET_NAME: 'Enrolments',
  REPORT_SHEET: 'Checks',
  REQUIRED: ['Name', 'Email', 'Course', 'Start date'],
  COLOUR: '#fde2e1'
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Data checks')
    .addItem('Run checks', 'runChecks')
    .addItem('Clear highlights', 'clearHighlights')
    .addToUi();
}

function runChecks() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(CHECK.SHEET_NAME);
  if (!sheet) throw new Error('Sheet not found: ' + CHECK.SHEET_NAME);
  clearHighlights();

  const values = sheet.getDataRange().getValues();
  const header = values[0].map(String);
  const idx = name => header.indexOf(name);
  CHECK.REQUIRED.forEach(n => { if (idx(n) === -1) throw new Error('Missing column: ' + n); });

  const issues = [];
  const seen = {};
  const emailRe = /^[^@\s]+@[^@\s]+\.[^@\s]+$/;

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (row.every(v => v === '' || v === null)) continue; // skip empty rows

    CHECK.REQUIRED.forEach(n => {
      if (String(row[idx(n)]).trim() === '') issues.push([r + 1, n, 'Required field is empty']);
    });

    const email = String(row[idx('Email')]).trim().toLowerCase();
    if (email && !emailRe.test(email)) issues.push([r + 1, 'Email', 'Invalid email: ' + email]);

    const key = email + '|' + String(row[idx('Course')]).trim().toLowerCase();
    if (email) {
      if (seen[key]) issues.push([r + 1, 'Email', 'Duplicate of row ' + seen[key]]);
      else seen[key] = r + 1;
    }
  }

  issues.forEach(([row, field]) => sheet.getRange(row, idx(field) + 1).setBackground(CHECK.COLOUR));
  writeReport_(ss, issues);
  ss.toast(issues.length ? issues.length + ' issue(s) found. See the "Checks" sheet.' : 'No issues found.',
           'Data checks', 6);
}

function clearHighlights() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CHECK.SHEET_NAME);
  if (sheet) sheet.getDataRange().offset(1, 0).setBackground(null);
}

function writeReport_(ss, issues) {
  const report = ss.getSheetByName(CHECK.REPORT_SHEET) || ss.insertSheet(CHECK.REPORT_SHEET);
  report.clear();
  const rows = [['Row', 'Field', 'Issue']].concat(issues);
  report.getRange(1, 1, rows.length, 3).setValues(rows);
  report.getRange(1, 1, 1, 3).setFontWeight('bold');
  report.autoResizeColumns(1, 3);
}
