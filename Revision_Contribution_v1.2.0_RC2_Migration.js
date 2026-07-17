/**
 * Revision_Contribution_v1.2.0_RC2_Migration.gs — OPTIONAL, ONE-TIME, MANUAL
 * migration, part of the RC2 correction pass.
 *
 * Do not run this automatically. Run it yourself from the Apps Script
 * editor (runRevisionContributionV120RC2Migration) once, before relying on
 * idempotent Revision Tracker aggregation (RC2 correction #1).
 *
 * Background: in the original v1.2.0 RC, completing a Study Planner session
 * called updateRevisionFromStudy_() every time the completion form was
 * saved, with no record of what a session had already contributed. Saving
 * the same form twice, editing the actual duration afterward, or flipping
 * between "Completed" and "Partly completed" could each add the same
 * session's minutes/session-count/source contribution more than once.
 *
 * What this migration does: adds five new, purely additive columns to "12
 * Study Planner" that let each session remember exactly what it last
 * contributed to Revision Tracker, so it can be reversed before a new
 * contribution is applied:
 *   - "Revision contribution applied" (TRUE/FALSE checkbox)
 *   - "Revision contribution module"
 *   - "Revision contribution topic"
 *   - "Revision contribution minutes"
 *   - "Revision contribution source" (Planned/Independent)
 *
 * These columns are written and read only by the RC2 correction code in
 * API_v1.2.0.gs (reconcileSessionRevisionContribution_,
 * reverseRevisionContribution_, readStoredContribution_,
 * writeStoredContribution_). Every existing "12 Study Planner" column —
 * including the ones added by Study_Tasks_v1.2.0_Migration.gs — is left
 * completely untouched. No existing Revision Tracker aggregate is
 * recalculated retroactively by this migration; it only prepares the ledger
 * columns going forward. If your Revision Tracker totals already
 * double-counted something from the original RC, see
 * V1.2.0_MIGRATION_GUIDE.md (RC2 section) for how to reset a topic's
 * aggregates by hand if you want a clean starting point — this is optional
 * and does not affect Confidence, Priority, or any other manually-set field.
 *
 * Idempotent: safe to run more than once. Existing values in these columns
 * (if this migration has already been run) are never reset by a second run.
 */
function runRevisionContributionV120RC2Migration() {
  var report = [];
  var sheet = SpreadsheetApp.getActive().getSheetByName(STUDY_SHEET);
  if (!sheet) {
    migrationReport_('Revision Contribution v1.2.0 RC2 migration', ['"' + STUDY_SHEET + '" sheet not found — nothing to do.']);
    return;
  }

  [
    'Revision contribution applied', 'Revision contribution module', 'Revision contribution topic',
    'Revision contribution minutes', 'Revision contribution source'
  ].forEach(function (name) {
    var result = ensureColumn_(sheet, HEADER_ROW, name);
    report.push((result.added ? 'Added column: "' : 'Already present: "') + name + '"');
  });

  ensureCheckboxColumn_(sheet, HEADER_ROW, 'Revision contribution applied');

  report.push('');
  report.push('These columns are additive only -- no existing "12 Study Planner" column was changed.');
  report.push('Existing Revision Tracker aggregates are not recalculated retroactively by this migration.');
  report.push('Safe to run again at any time.');
  migrationReport_('Revision Contribution v1.2.0 RC2 migration', report);
}
