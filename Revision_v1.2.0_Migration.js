/**
 * Revision_v1.2.0_Migration.gs — OPTIONAL, ONE-TIME, MANUAL migration.
 *
 * Do not run this automatically. Run it yourself from the Apps Script
 * editor (runRevisionV120Migration) once, before using the auto-updating
 * Revision Tracker.
 *
 * What it does: adds 9 new, purely additive columns to "13 Revision
 * Tracker" — aggregate fields that the new api_completeStudySession /
 * api_logIndependentStudy / api_toggleStudyTask flows keep up to date
 * automatically (Total actual study minutes, Session count, Last revised,
 * Completed task count, Score, Weak topic flag, Planned-study count,
 * Independent-study count, Source). The existing Module / Topic /
 * Confidence / Priority / Next revision date / Completed / Notes columns
 * are untouched — Confidence in particular remains a manual field (per the
 * v1.2.0 spec: completing a task never silently sets confidence).
 *
 * Idempotent: safe to run more than once.
 */
function runRevisionV120Migration() {
  var sheet = SpreadsheetApp.getActive().getSheetByName(REVISION_SHEET);
  if (!sheet) {
    migrationReport_('Revision Tracker v1.2.0 migration', ['"' + REVISION_SHEET + '" sheet not found — nothing to do.']);
    return;
  }
  var headerRow = HEADER_ROW;
  var report = [];
  [
    'Total actual study minutes', 'Session count', 'Last revised', 'Completed task count',
    'Score', 'Weak topic (TRUE/FALSE)', 'Planned-study count', 'Independent-study count', 'Source'
  ].forEach(function (name) {
    var result = ensureColumn_(sheet, headerRow, name);
    report.push((result.added ? 'Added column: "' : 'Already present: "') + name + '"');
  });

  ensureCheckboxColumn_(sheet, headerRow, 'Weak topic (TRUE/FALSE)');

  report.push('');
  report.push('Existing Module/Topic/Confidence/Priority/Next revision date/Completed/Notes columns are untouched.');
  report.push('Confidence stays a manual field — nothing here infers it automatically.');
  report.push('Safe to run again at any time.');
  migrationReport_('Revision Tracker v1.2.0 migration', report);
}
