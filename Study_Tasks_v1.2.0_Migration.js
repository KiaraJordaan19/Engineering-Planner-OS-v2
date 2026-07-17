/**
 * Study_Tasks_v1.2.0_Migration.gs — OPTIONAL, ONE-TIME, MANUAL migration.
 *
 * Do not run this automatically. Run it yourself from the Apps Script
 * editor (runStudyTasksV120Migration) once, before using the task-based
 * Study Planner (add/edit session, tick tasks, log independent study,
 * generate/accept weekly suggestions).
 *
 * What it does:
 *  1. Creates a brand-new, thin sheet, "21 Study Tasks" (only if it doesn't
 *     already exist), using the SAME header-row convention as every other
 *     sheet in this workbook (header row 4, per the shared HEADER_ROW
 *     constant) so every existing shared helper (readSheetRows_,
 *     getColMap_, firstBlankRow_, assignStableIds_) works against it
 *     without any special-casing. Headers: Task ID, Study Session ID,
 *     Module, Topic, Task type, Task text, Resource ID, Completed, Score,
 *     Notes, Sort order, Last updated.
 *  2. Adds four new columns to the existing "12 Study Planner" sheet:
 *     "Status" (Planned/Completed/Partly completed/Skipped/Rescheduled —
 *     additive to, and independent from, the existing Planned/Completed
 *     checkboxes, which are left exactly as they are for any code that
 *     still reads them), "Objective", "Actual duration (min)", and
 *     "Confidence after". Also ensures the "Origin" column exists (some
 *     v1.1.x workbooks already have it from generateStudySuggestions();
 *     this migration adds it only if genuinely missing) — v1.2.0 reuses
 *     "Origin" with an expanded set of values ("Manual", "Generated",
 *     "Independent", "Catch-up") rather than adding a new column for the
 *     same concept.
 *
 * Idempotent: safe to run more than once. No existing "12 Study Planner"
 * column is renamed or removed, and re-running never re-creates "21 Study
 * Tasks" or duplicates its headers.
 */
function runStudyTasksV120Migration() {
  var report = [];

  var sheetResult = ensureSheetExists_('21 Study Tasks', HEADER_ROW, [
    'Task ID', 'Study Session ID', 'Module', 'Topic', 'Task type', 'Task text',
    'Resource ID', 'Completed', 'Score', 'Notes', 'Sort order', 'Last updated'
  ]);
  report.push(sheetResult.created ? 'Created new sheet: "21 Study Tasks".' : '"21 Study Tasks" already existed.');
  report.push(sheetResult.addedHeaders.length
    ? ('Headers added: ' + sheetResult.addedHeaders.join(', '))
    : 'All expected headers were already present.');
  ensureCheckboxColumn_(sheetResult.sheet, HEADER_ROW, 'Completed');

  var studySheet = SpreadsheetApp.getActive().getSheetByName(STUDY_SHEET);
  if (!studySheet) {
    report.push('"' + STUDY_SHEET + '" sheet not found — skipped the Study Planner column additions.');
  } else {
    ['Status', 'Objective', 'Actual duration (min)', 'Confidence after', 'Origin'].forEach(function (name) {
      var result = ensureColumn_(studySheet, HEADER_ROW, name);
      report.push((result.added ? 'Added column to "12 Study Planner": "' : 'Already present on "12 Study Planner": "') + name + '"');
    });
  }

  report.push('');
  report.push('Existing Planned/Completed checkboxes on "12 Study Planner" are untouched — the new');
  report.push('"Status" column is additive and used by the new UI alongside them, not instead of them.');
  report.push('Safe to run again at any time.');
  migrationReport_('Study Tasks v1.2.0 migration', report);
}
