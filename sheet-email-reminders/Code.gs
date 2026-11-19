/**
 * Sheet email reminders
 * ---------------------
 * Sends a reminder email to every learner whose session starts in N days,
 * using a Google Sheet as the source. Each row is marked as "sent" so that
 * nobody receives the same reminder twice.
 *
 * Expected sheet "Sessions" (row 1 = headers):
 *   Name | Email | Course | Session date | Reminder sent
 *
 * Setup: Extensions > Apps Script, paste this file, set CONFIG, run
 * `sendReminders` once to authorise, then add a daily time-driven trigger
 * with `createDailyTrigger`.
 */

const CONFIG = {
  SHEET_NAME: 'Sessions',
  DAYS_BEFORE: 2,            // send the reminder 2 days before the session
  SENDER_NAME: 'Training Team',
  DRY_RUN: true              // true = only log what would be sent
};

function sendReminders() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error('Sheet not found: ' + CONFIG.SHEET_NAME);

  const values = sheet.getDataRange().getValues();
  const header = values[0].map(String);
  const col = name => {
    const i = header.indexOf(name);
    if (i === -1) throw new Error('Missing column: ' + name);
    return i;
  };
  const C = {
    name: col('Name'), email: col('Email'), course: col('Course'),
    date: col('Session date'), sent: col('Reminder sent')
  };

  const tz = Session.getScriptTimeZone();
  const target = new Date();
  target.setDate(target.getDate() + CONFIG.DAYS_BEFORE);
  const targetKey = Utilities.formatDate(target, tz, 'yyyy-MM-dd');

  let count = 0;
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const date = row[C.date];
    if (!(date instanceof Date) || row[C.sent]) continue;
    if (Utilities.formatDate(date, tz, 'yyyy-MM-dd') !== targetKey) continue;

    const email = String(row[C.email]).trim();
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) {
      Logger.log('Row %s: invalid email "%s", skipped', r + 1, email);
      continue;
    }

    const when = Utilities.formatDate(date, tz, 'dd/MM/yyyy HH:mm');
    const subject = 'Reminder: ' + row[C.course] + ' on ' + when;
    const body =
      'Hello ' + row[C.name] + ',\n\n' +
      'This is a reminder that your session "' + row[C.course] + '" starts on ' + when + '.\n\n' +
      'See you there!\n' + CONFIG.SENDER_NAME;

    if (CONFIG.DRY_RUN) {
      Logger.log('[DRY RUN] Would send to %s: %s', email, subject);
    } else {
      MailApp.sendEmail({ to: email, subject: subject, body: body, name: CONFIG.SENDER_NAME });
      sheet.getRange(r + 1, C.sent + 1).setValue(new Date());
    }
    count++;
  }
  Logger.log('%s reminder(s) %s.', count, CONFIG.DRY_RUN ? 'prepared' : 'sent');
}

/** Run once: creates a trigger that runs sendReminders every day at about 08:00. */
function createDailyTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'sendReminders')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('sendReminders').timeBased().everyDays(1).atHour(8).create();
}
