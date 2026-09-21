/**
 * Marks_Intelligence_v1.4.7_CP143AttendanceNote_Migration.gs — OPTIONAL,
 * ONE-TIME, MANUAL migration.
 *
 * Do not run this automatically. Run it yourself from the Apps Script editor
 * (runCp143AttendanceNoteV147Migration) once.
 *
 * Adds a visible reminder to Computer Programming 143's (30317-143) "04
 * Module Rules" row about its practical attendance requirement.
 *
 * v1.4.7.1 UPDATE: originally written from the user's own recollection
 * ("can't miss more than 2"), with a disclosed mismatch against a dated
 * (2023) third-party copy of the framework that said "more than 3". The
 * user's own current (2026) module framework PDF, read directly, settles
 * it in the user's favour: "Each student will be allowed to miss a maximum
 * of two practical tests for any reason... If more than two tests were
 * missed, the student will receive an INCOMPLETE... and WILL NOT BE ALLOWED
 * TO WRITE THE MAIN ASSESSMENTS." The 2023 copy was simply out of date.
 *
 * Also directly confirms the drop-lowest-2 rule for this module (already
 * implemented, Marks_Intelligence_v1.4.6_DropLowest2_Migration.gs): "The
 * two worst practical test marks (including missed practical tests) will
 * be ignored when calculating the semester mark."
 *
 * This is purely informational text -- no calculation, weighting, or the
 * Marks Intelligence engine is affected by it.
 *
 * Idempotent: replaces the earlier hedged note (from v1.4.7) with the now-
 * confirmed version if that exact earlier text is still present; otherwise
 * appends (if some other note already exists) or sets (if blank). Never
 * touches a "General notes" value that doesn't match the earlier note and
 * isn't blank -- that looks like something you wrote yourself.
 */
function runCp143AttendanceNoteV147Migration() {
  var report = [];
  var ss = SpreadsheetApp.getActive();
  var targetCode = "30317-143";
  var oldNoteMarker = "cannot miss more than 2 practicals (per user's own note";
  var noteText = "Attendance (confirmed, 2026 module framework): may miss at most 2 practical tests (0 awarded for each) -- missing a 3rd results in an INCOMPLETE and no admission to A1/A2/A3. The semester mark already ignores your 2 worst practical test marks (including missed ones) when averaging, per the same framework.";

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
    if (current.indexOf("Attendance (confirmed, 2026 module framework)") !== -1) {
      report.push('Confirmed note already present on ' + targetCode + ' (Computer Programming 143) — left unchanged.');
    } else if (current.indexOf(oldNoteMarker) !== -1) {
      cell.setValue(noteText);
      report.push('Replaced the earlier hedged note with the now-confirmed version for ' + targetCode + ' (Computer Programming 143).');
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
