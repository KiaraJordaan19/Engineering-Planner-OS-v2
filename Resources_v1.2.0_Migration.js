/**
 * Resources_v1.2.0_Migration.gs — OPTIONAL, ONE-TIME, MANUAL migration.
 *
 * Do not run this automatically. Run it yourself from the Apps Script
 * editor (runResourcesV120Migration) once, before using the new Resources
 * add/edit/delete/progress UI.
 *
 * What it does:
 *  1. Adds a stable "Resource ID" column (prefix RES) and backfills it for
 *     every existing row, using the same assignStableIds_() helper the
 *     v1.1.x workbook already uses for Assignments/AF Components/Study
 *     Planner/Academic Inbox — never re-derived from row position later.
 *  2. Adds "Status" (Not started / In progress / Completed / Reviewed /
 *     Needs redo) plus four generic "Progress Flag A–D" columns and a
 *     "Score" column and a "Flashcards created" checkbox. The four generic
 *     flags are deliberately generic rather than one column per resource
 *     type — the frontend maps each resource type to the right on-screen
 *     labels for A–D (e.g. Past paper: Attempted/Marked/Mistakes
 *     reviewed/Redone; Lecture: Watched/Understood/Practice
 *     done/Needs revisit), so this stays "lightweight" per the v1.2.0 spec
 *     rather than becoming a bespoke column per resource type.
 *
 * The existing "Reviewed" checkbox column and api_toggleResourceReviewed()
 * are completely untouched — Status supersedes Reviewed for new UI, but the
 * old column and function keep working exactly as before.
 *
 * Idempotent: safe to run more than once.
 */
function runResourcesV120Migration() {
  var sheet = SpreadsheetApp.getActive().getSheetByName(RESOURCES_SHEET);
  if (!sheet) {
    migrationReport_('Resources v1.2.0 migration', ['"' + RESOURCES_SHEET + '" sheet not found — nothing to do.']);
    return;
  }
  var headerRow = HEADER_ROW;
  var report = [];

  var idResult = ensureColumn_(sheet, headerRow, 'Resource ID');
  report.push((idResult.added ? 'Added column: "Resource ID"' : 'Already present: "Resource ID"'));

  ['Status', 'Progress Flag A', 'Progress Flag B', 'Progress Flag C', 'Progress Flag D',
    'Flashcards created (TRUE/FALSE)', 'Score'].forEach(function (name) {
      var result = ensureColumn_(sheet, headerRow, name);
      report.push((result.added ? 'Added column: "' : 'Already present: "') + name + '"');
    });

  // Backfill stable IDs for every existing row that has a Title but no
  // Resource ID yet — same pattern as assignStableIds_() already uses for
  // every other sheet's stable-ID column.
  var assigned = assignStableIds_(RESOURCES_SHEET, 'Resource ID', 'Title', 'RES');
  report.push('Resource ID: ' + assigned + ' existing row(s) assigned a new stable ID.');

  ensureCheckboxColumn_(sheet, headerRow, 'Flashcards created (TRUE/FALSE)');

  report.push('');
  report.push('The existing "Reviewed" column and api_toggleResourceReviewed() are untouched.');
  report.push('Safe to run again at any time.');
  migrationReport_('Resources v1.2.0 migration', report);
}
