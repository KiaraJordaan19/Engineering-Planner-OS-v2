/**
 * Attendance_v1.2.0_Migration.gs — OPTIONAL, ONE-TIME, MANUAL migration.
 *
 * Do not run this automatically. Run it yourself from the Apps Script
 * editor (runAttendanceV120Migration) once, before using the new "Today's
 * non-negotiable classes" Attended/Missed tracking on the Attendance
 * screen.
 *
 * What it does: adds three new, purely additive columns to "19 Attendance"
 * — "Attendance ID" (stable ID, prefix ATT), "Session key" (a composite
 * Module|SessionType|DateISO key used ONLY to prevent creating two records
 * for the same real timetable session — never shown to the user), and
 * "Status" (Not recorded / Attended / Missed). The existing "Attended"
 * TRUE/FALSE column is left completely alone and keeps working exactly as
 * before — api_setSessionAttendance() (v1.2.0) keeps it in sync (TRUE only
 * when Status = "Attended") purely so the existing attendance-percentage
 * aggregation already used by the Dashboard/Attendance screen keeps working
 * unchanged, without needing to know the new Status column exists.
 *
 * Idempotent: safe to run more than once.
 */
function runAttendanceV120Migration() {
  var sheet = SpreadsheetApp.getActive().getSheetByName(ATTENDANCE_SHEET);
  if (!sheet) {
    migrationReport_('Attendance v1.2.0 migration', ['"' + ATTENDANCE_SHEET + '" sheet not found — nothing to do.']);
    return;
  }
  var headerRow = HEADER_ROW;
  var report = [];

  var idResult = ensureColumn_(sheet, headerRow, 'Attendance ID');
  report.push((idResult.added ? 'Added column: "Attendance ID"' : 'Already present: "Attendance ID"'));

  ['Session key', 'Status'].forEach(function (name) {
    var result = ensureColumn_(sheet, headerRow, name);
    report.push((result.added ? 'Added column: "' : 'Already present: "') + name + '"');
  });

  // Backfill IDs and session keys for any rows already logged by hand.
  var assigned = assignStableIds_(ATTENDANCE_SHEET, 'Attendance ID', 'Module', 'ATT');
  report.push('Attendance ID: ' + assigned + ' existing row(s) assigned a new stable ID.');

  var map = getColMap_(sheet);
  var lastRow = sheet.getLastRow();
  if (map['Session key'] && map['Module'] && map['Session type'] && map['Date'] && lastRow > headerRow) {
    var modules = sheet.getRange(headerRow + 1, map['Module'], lastRow - headerRow, 1).getValues();
    var types = sheet.getRange(headerRow + 1, map['Session type'], lastRow - headerRow, 1).getValues();
    var dates = sheet.getRange(headerRow + 1, map['Date'], lastRow - headerRow, 1).getValues();
    var keys = sheet.getRange(headerRow + 1, map['Session key'], lastRow - headerRow, 1).getValues();
    var backfilled = 0;
    for (var i = 0; i < modules.length; i++) {
      if (keys[i][0] || !modules[i][0]) continue;
      var dateStr = (dates[i][0] instanceof Date) ? Utilities.formatDate(dates[i][0], TIMEZONE, 'yyyy-MM-dd') : String(dates[i][0] || '');
      var key = modules[i][0] + '|' + (types[i][0] || '') + '|' + dateStr;
      sheet.getRange(headerRow + 1 + i, map['Session key']).setValue(key);
      backfilled++;
    }
    report.push('Session key: ' + backfilled + ' existing row(s) backfilled from Module + Session type + Date.');
  }
  if (map['Status'] && map['Attended'] && lastRow > headerRow) {
    var attendedVals = sheet.getRange(headerRow + 1, map['Attended'], lastRow - headerRow, 1).getValues();
    var statusVals = sheet.getRange(headerRow + 1, map['Status'], lastRow - headerRow, 1).getValues();
    var statusBackfilled = 0;
    for (var j = 0; j < attendedVals.length; j++) {
      if (statusVals[j][0]) continue;
      if (attendedVals[j][0] === true) { sheet.getRange(headerRow + 1 + j, map['Status']).setValue('Attended'); statusBackfilled++; }
    }
    report.push('Status: ' + statusBackfilled + ' existing "Attended" row(s) backfilled to Status = "Attended".');
  }

  report.push('');
  report.push('The existing "Attended" checkbox column is untouched and stays in sync going forward.');
  report.push('Safe to run again at any time.');
  migrationReport_('Attendance v1.2.0 migration', report);
}
