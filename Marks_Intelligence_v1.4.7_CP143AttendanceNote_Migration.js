/**
 * Marks_Intelligence_v1.4.7_CP143AttendanceNote_Migration.gs — OPTIONAL,
 * ONE-TIME, MANUAL migration.
 *
 * Do not run this automatically. Run it yourself from the Apps Script editor
 * (runCp143AttendanceNoteV147Migration) once.
 *
 * Adds a visible reminder to Computer Programming 143's (30317-143) "04
 * Module Rules" row about its practical attendance requirement, per the
 * user's own instruction (2026-09-21): "can't miss more than 2" practicals.
 *
 * Disclosed discrepancy: the current (2026) module framework itself, as
 * relayed by the user, says missing MORE THAN 3 practical tests results in
 * an INCOMPLETE -- i.e. missing 1-3 is survivable, the 4th is the problem.
 * The note below uses the user's own stricter "not more than 2" wording as
 * their personal safety margin, not a correction of the sourced framework
 * figure. If this was meant to literally replace the "more than 3" figure,
 * tell Claude and this note (and this comment) should be corrected.
 *
 * This is purely informational text -- no calculation, weighting, or the
 * Marks Intelligence engine is affected by it.
 *
 * Idempotent: appends to "General notes" only if this exact reminder isn't
 * already present in the cell; never overwrites or removes anything else
 * already written there.
 */
function runCp143AttendanceNoteV147Migration() {
  var report = [];
  var ss = SpreadsheetApp.getActive();
  var targetCode = "30317-143";
  var noteText = "Attendance: cannot miss more than 2 practicals (per user's own note, 2026 — the module framework itself specifies missing MORE THAN 3 practical tests causes an INCOMPLETE).";

  var mrSheet = ss.getSheetByName(MODULE_RULES_SHEET);
  if (!mrSheet) {
    migrationReport_('CP143 attendance note v1.4.7 migration', ['"' + MODULE_RULES_SHEET + '" sheet not found — nothing to do.']);
    return;
  }
  var mrMap = getColMap_(mrSheet, HEADER_ROW);
  var mrLastRow = mrSheet.getLastRow();
  var codeCol = mrMap['Module code'], notesCol = mrMap['General notes'];
  if (!codeCol || !notesCol) {
    migrationReport_('CP143 attendance note v1.4.7 migration', ['"Module code" or "General notes" column not found on "' + MODULE_RULES_SHEET + '" — nothing to do.']);
    return;
  }

  var found = false;
  for (var row = HEADER_ROW + 1; row <= mrLastRow; row++) {
    if (mrSheet.getRange(row, codeCol).getValue() !== targetCode) continue;
    found = true;
    var cell = mrSheet.getRange(row, notesCol);
    var current = (cell.getValue() || "").toString();
    if (current.indexOf("cannot miss more than 2 practicals") !== -1) {
      report.push('Note already present on ' + targetCode + ' (Computer Programming 143) — left unchanged.');
    } else if (!current) {
      cell.setValue(noteText);
      report.push('Set "General notes" for ' + targetCode + ' (Computer Programming 143).');
    } else {
      cell.setValue(current + "\n\n" + noteText);
      report.push('Appended to existing "General notes" for ' + targetCode + ' (Computer Programming 143).');
    }
    break;
  }
  if (!found) report.push('No "04 Module Rules" row found for ' + targetCode + ' (Computer Programming 143) — nothing to do.');

  report.push('Safe to run again at any time.');
  migrationReport_('CP143 attendance note v1.4.7 migration', report);
}
