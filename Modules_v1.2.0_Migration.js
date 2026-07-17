/**
 * Modules_v1.2.0_Migration.gs — OPTIONAL, ONE-TIME, MANUAL migration.
 *
 * Do not run this automatically. Do not attach it to a trigger. Run it
 * yourself from the Apps Script editor (select runModulesV120Migration from
 * the function dropdown, click Run) once, after installing the v1.2.0
 * files, before using the new Module Overview contacts/configuration UI.
 *
 * What it does: adds 13 new, purely additive columns to "03 Modules" —
 * colour, difficulty/workload/target-mark/study-hours configuration,
 * attendance-sensitivity flags, and minimal per-module contact fields
 * (coordinator/lecturer name+email, one notes field). Nothing existing on
 * "03 Modules" is renamed, reordered, or removed, and no formula anywhere
 * in the workbook is touched. Requires Migrations_Shared_v1.2.0.gs to be
 * present in the same project (for ensureColumn_/migrationReport_).
 *
 * Idempotent: run it as many times as you like. Every column is checked by
 * exact header name before being added — a column that's already there
 * (from a previous run, or because you added it by hand) is left exactly
 * as it is.
 */
function runModulesV120Migration() {
  var sheet = SpreadsheetApp.getActive().getSheetByName(MODULES_SHEET);
  if (!sheet) {
    migrationReport_('Modules v1.2.0 migration', ['"' + MODULES_SHEET + '" sheet not found — nothing to do.']);
    return;
  }
  var headerRow = HEADER_ROW;
  var columns = [
    'Colour (hex)',
    'Conceptual difficulty (1-5)',
    'Workload intensity (1-5)',
    'Repeated module (TRUE/FALSE)',
    'Target mark (%)',
    'Min weekly study hours',
    'Attendance-sensitive lecture (TRUE/FALSE)',
    'Compulsory practical/lab (TRUE/FALSE)',
    'Coordinator name',
    'Coordinator email',
    'Lecturer name',
    'Lecturer email',
    'Contact notes'
  ];
  var report = [];
  columns.forEach(function (name) {
    var result = ensureColumn_(sheet, headerRow, name);
    report.push((result.added ? 'Added column: "' : 'Already present: "') + name + '"');
  });

  // Convert the three boolean fields to native checkboxes (matches the
  // existing v1.1.x convention for every other TRUE/FALSE column in the
  // workbook — see insertCheckboxes_() in Calendar_Sync_v1.2.0.gs).
  ensureCheckboxColumn_(sheet, headerRow, 'Repeated module (TRUE/FALSE)');
  ensureCheckboxColumn_(sheet, headerRow, 'Attendance-sensitive lecture (TRUE/FALSE)');
  ensureCheckboxColumn_(sheet, headerRow, 'Compulsory practical/lab (TRUE/FALSE)');

  report.push('');
  report.push('No existing "03 Modules" column was renamed, reordered, or removed.');
  report.push('Safe to run again at any time.');
  migrationReport_('Modules v1.2.0 migration', report);
}
