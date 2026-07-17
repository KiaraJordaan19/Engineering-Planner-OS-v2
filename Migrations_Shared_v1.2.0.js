/**
 * Migrations_Shared_v1.2.0.gs — Engineering Planner OS
 *
 * Small, reusable, idempotent helpers shared by every v1.2.0 migration file
 * (Modules_v1.2.0_Migration.gs, Resources_v1.2.0_Migration.gs,
 * Study_Tasks_v1.2.0_Migration.gs, Revision_v1.2.0_Migration.gs,
 * Attendance_v1.2.0_Migration.gs). Add this file ALONGSIDE those five files
 * in the same Apps Script project — it must be present exactly once (these
 * helpers are deliberately factored out here, rather than copy-pasted into
 * each migration file, precisely so five separate files can safely coexist
 * in one project's shared global scope without redeclaring the same
 * function name five times).
 *
 * Nothing in this file runs automatically. Nothing here is called by the
 * web app, by any trigger, or by any other production file. Every migration
 * is a manually-run, one-time, idempotent action — run it once from the
 * Apps Script editor's function dropdown, read the report, done. Running
 * any migration a second time is always safe: every helper below checks
 * "does this already exist?" before creating anything, and never touches a
 * column or sheet that already exists.
 *
 * Depends on getColMap_(), HEADER_ROW, and assignStableIds_() already
 * declared in Calendar_Sync_v1.2.0.gs — this file assumes it is loaded into
 * the same project as the production files, not run standalone.
 */

/**
 * Adds one new header (at the next free column, in the given header row) to
 * an existing sheet, ONLY if a column with that exact header name doesn't
 * already exist. Never moves, renames, or removes any existing column —
 * purely additive, at the right-hand edge of the sheet's current data.
 * Returns { added: boolean, column: number } so callers can build a report.
 */
function ensureColumn_(sheet, headerRow, headerName) {
  var map = getColMap_(sheet);
  if (map[headerName]) return { added: false, column: map[headerName] };
  var newCol = sheet.getLastColumn() + 1;
  sheet.getRange(headerRow, newCol).setValue(headerName);
  return { added: true, column: newCol };
}

/**
 * Ensures a sheet exists (creating it, blank, if it doesn't) and that every
 * header in `headers` is present somewhere in the given header row — again,
 * purely additive: existing headers (in any order) are left completely
 * alone, and only genuinely missing ones are appended.
 * Returns { created: boolean, addedHeaders: string[] }.
 */
function ensureSheetExists_(sheetName, headerRow, headers) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(sheetName);
  var created = false;
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    created = true;
  }
  var addedHeaders = [];
  headers.forEach(function (h) {
    var result = ensureColumn_(sheet, headerRow, h);
    if (result.added) addedHeaders.push(h);
  });
  return { created: created, addedHeaders: addedHeaders, sheet: sheet };
}

/**
 * Converts a (possibly newly-added) TRUE/FALSE column to native Sheets
 * checkboxes over a generous row range, the same way insertCheckboxes_() in
 * Calendar_Sync_v1.2.0.gs already does for the v1.1.x checkbox columns. Safe
 * to call on a column that already has checkboxes — re-inserting is a no-op
 * beyond the visual formatting call itself.
 */
function ensureCheckboxColumn_(sheet, headerRow, headerName, extraRows) {
  var map = getColMap_(sheet);
  if (!map[headerName]) return false;
  var lastRow = Math.max(sheet.getLastRow(), headerRow + (extraRows || 60));
  sheet.getRange(headerRow + 1, map[headerName], lastRow - headerRow, 1).insertCheckboxes();
  return true;
}

/**
 * Shows a Sheets UI alert (when run from the Sheets-bound Apps Script editor
 * with a UI context available) AND logs the same report to the Apps Script
 * execution log (Logger/console), so the migration's outcome is visible
 * either way — including when run in a context where getUi() throws.
 */
function migrationReport_(title, lines) {
  var msg = lines.join('\n');
  try {
    SpreadsheetApp.getUi().alert(title, msg, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    // No UI context (e.g. run directly from the editor without the sheet
    // open) — the console log below is still the source of truth.
  }
  console.log(title + '\n\n' + msg);
  return msg;
}
