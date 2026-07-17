/**
 * Semester_Lifecycle_v1.4.0_Migration.gs — OPTIONAL, ONE-TIME, MANUAL migration.
 *
 * Do not run this automatically. Run it yourself from the Apps Script editor
 * (runSemesterLifecycleV140Migration) once, before using any v1.4.0 Semester
 * Lifecycle & Analytics feature (Semester Report, Semester Archive, Backup,
 * Delete Semester, Reset Active Semester, New Semester Wizard, Module
 * Templates, or Historical Analytics).
 *
 * This is the ONLY schema change v1.4.0 requires. It is purely ADDITIVE —
 * four new sheets and a handful of new Settings rows — and does not rename,
 * reorder, remove, or repurpose any existing sheet, column, row, or formula
 * anywhere in the workbook, including the v1.1-era semester tools already in
 * Calendar_Sync (archiveSemester / startNewSemester / updateSemesterDates /
 * duplicateSemester) and "17 Archive", which are left completely alone and
 * keep working exactly as before. v1.4.0's Semester screen is a new,
 * separate, considerably more complete system that sits alongside them —
 * see V1.4.0_MIGRATION_GUIDE.md for how the two relate.
 *
 * What this migration adds:
 *
 *  1. New sheet "23 Semester Archive" — one row per COMPLETED, archived
 *     semester. This is the read-only historical record: once a row is
 *     written here (by semCreateArchive_ in Semester_Archive_v1.4.0.gs) it is
 *     never modified or deleted by any v1.4.0 code path — see "Archive
 *     immutability" in V1.4.0_ARCHITECTURE_NOTES.md. Headers: Semester ID,
 *     Semester Name, Academic Year, Start Date, End Date, Completion Date,
 *     Created Timestamp, Planner Version, Status, Immutable (TRUE/FALSE),
 *     Report JSON, Charts JSON, Reflection JSON.
 *  2. New sheet "24 Semester Backups" — one row per backup ever created
 *     (semCreateBackup_ in Semester_Backup_v1.4.0.gs), each pointing at one
 *     Drive file containing the full deterministic JSON backup. Headers:
 *     Backup ID, Semester ID, Timestamp, Planner Version, Drive File ID,
 *     Drive File URL, Sheets Included, Notes.
 *  3. New sheet "25 Module Templates" — one row per module, holding exactly
 *     the fields the spec requires to survive every semester (colour,
 *     contacts, weightings, attendance rules, study preferences, assessment
 *     framework). Refreshed (never silently discarded) every time a semester
 *     is archived, and read by the New Semester Wizard's "reuse previous"
 *     options. Headers: Module Code, Module Name, Colour (hex), Contacts
 *     JSON, Weightings JSON, Attendance Rules JSON, Study Preferences JSON,
 *     Assessment Framework JSON, Last Updated, Source Semester ID.
 *  4. New sheet "26 Semester History" — one flattened row per archived
 *     semester, purely so Historical Analytics can compare semesters without
 *     parsing every archive's full Report JSON. Deterministic rollups only —
 *     no predictive analytics. Headers: Semester ID, Semester Name, Academic
 *     Year, Completion Date, Study Hours, Avg Mark, Highest Mark, Lowest
 *     Mark, Distinction Count, Pass Count, Attendance Rate (%), Revision
 *     Completion Rate (%), Resource Completion Rate (%).
 *  5. New Settings rows in "02 Settings" (only if the exact label doesn't
 *     already exist):
 *       - "Academic year" (blank default) — e.g. "2026".
 *       - "Semester ID" (blank default) — the CURRENT active semester's
 *         unique ID; blank until the New Semester Wizard is completed once,
 *         or Generate Report is used for the first time (which assigns one
 *         lazily if still blank, so Report/Archive/Backup always have a
 *         stable ID to key off even on a workbook that pre-dates the
 *         Wizard).
 *       - "Semester status" (default "Active") — "Active" or "Complete".
 *         Only ever set to "Complete" by the Delete Semester workflow or by
 *         the manual "Finish Semester" action — see
 *         V1.4.0_SEMESTER_REPORT.md §1 (Completion Detection). Never set
 *         automatically just because assessments finished; that only makes
 *         the "Finish Semester" affordance appear.
 *       - "Semester manually finished (TRUE/FALSE)" (default FALSE) — set by
 *         the explicit "Finish Semester" button; distinct from automatic
 *         completion DETECTION (all assessments finished), which never sets
 *         this on its own.
 *       - "Planner version" (default "1.4.0") — stamped onto every archive
 *         and backup created from this workbook; update it by hand only if
 *         you install a later release over this one without running that
 *         release's own migration.
 *
 * Nothing here recomputes, backfills, or guesses a historical value — every
 * addition starts blank/default and is only ever populated by a real,
 * forward-going action (running a report, creating a backup, completing the
 * wizard, etc).
 *
 * Idempotent: safe to run more than once. Depends on
 * Migrations_Shared_v1.2.0.gs (ensureSheetExists_ / ensureColumn_ /
 * migrationReport_) and ensureSettingRow_ (from
 * Academic_Intelligence_v1.3.1_Migration.gs) already being present in the
 * same project.
 */
function runSemesterLifecycleV140Migration() {
  var report = [];

  // ---- 1. "23 Semester Archive" ----
  var archiveResult = ensureSheetExists_('23 Semester Archive', HEADER_ROW, [
    'Semester ID', 'Semester Name', 'Academic Year', 'Start Date', 'End Date', 'Completion Date',
    'Created Timestamp', 'Planner Version', 'Status', 'Immutable', 'Report JSON', 'Charts JSON', 'Reflection JSON'
  ]);
  report.push(archiveResult.created ? 'Created new sheet: "23 Semester Archive".' : '"23 Semester Archive" already existed.');
  report.push(archiveResult.addedHeaders.length ? ('Headers added: ' + archiveResult.addedHeaders.join(', ')) : 'All expected headers were already present.');

  // ---- 2. "24 Semester Backups" ----
  var backupsResult = ensureSheetExists_('24 Semester Backups', HEADER_ROW, [
    'Backup ID', 'Semester ID', 'Timestamp', 'Planner Version', 'Drive File ID', 'Drive File URL', 'Sheets Included', 'Notes'
  ]);
  report.push(backupsResult.created ? 'Created new sheet: "24 Semester Backups".' : '"24 Semester Backups" already existed.');
  report.push(backupsResult.addedHeaders.length ? ('Headers added: ' + backupsResult.addedHeaders.join(', ')) : 'All expected headers were already present.');

  // ---- 3. "25 Module Templates" ----
  var templatesResult = ensureSheetExists_('25 Module Templates', HEADER_ROW, [
    'Module Code', 'Module Name', 'Colour (hex)', 'Contacts JSON', 'Weightings JSON',
    'Attendance Rules JSON', 'Study Preferences JSON', 'Assessment Framework JSON', 'Last Updated', 'Source Semester ID'
  ]);
  report.push(templatesResult.created ? 'Created new sheet: "25 Module Templates".' : '"25 Module Templates" already existed.');
  report.push(templatesResult.addedHeaders.length ? ('Headers added: ' + templatesResult.addedHeaders.join(', ')) : 'All expected headers were already present.');

  // ---- 4. "26 Semester History" ----
  var historyResult = ensureSheetExists_('26 Semester History', HEADER_ROW, [
    'Semester ID', 'Semester Name', 'Academic Year', 'Completion Date', 'Study Hours', 'Avg Mark',
    'Highest Mark', 'Lowest Mark', 'Distinction Count', 'Pass Count',
    'Attendance Rate (%)', 'Revision Completion Rate (%)', 'Resource Completion Rate (%)'
  ]);
  report.push(historyResult.created ? 'Created new sheet: "26 Semester History".' : '"26 Semester History" already existed.');
  report.push(historyResult.addedHeaders.length ? ('Headers added: ' + historyResult.addedHeaders.join(', ')) : 'All expected headers were already present.');

  // ---- 5. "27 Semester Workflows" (single-use destructive-operation authorizations) ----
  var workflowResult = ensureSheetExists_('27 Semester Workflows', HEADER_ROW, [
    'Workflow Token', 'Semester ID', 'Backup ID', 'Archive Semester ID', 'Prepared At',
    'Expires At', 'Status', 'Consumed At', 'Planner Version', 'Last Stage', 'Last Error'
  ]);
  report.push(workflowResult.created ? 'Created new sheet: "27 Semester Workflows".' : '"27 Semester Workflows" already existed.');
  report.push(workflowResult.addedHeaders.length ? ('Headers added: ' + workflowResult.addedHeaders.join(', ')) : 'All expected headers were already present.');

  // ---- 6. "02 Settings" ----
  var settingsSheet = SpreadsheetApp.getActive().getSheetByName('02 Settings');
  if (!settingsSheet) {
    report.push('"02 Settings" sheet not found — skipped every Settings row below.');
  } else {
    var settingsAdds = [
      ['Academic year', ''],
      ['Semester ID', ''],
      ['Semester status', 'Active'],
      ['Semester manually finished (TRUE/FALSE)', false],
      ['Planner version', '1.4.0']
    ];
    settingsAdds.forEach(function (pair) {
      var added = ensureSettingRow_(settingsSheet, pair[0], pair[1]);
      report.push(added
        ? ('Added new Settings row: "' + pair[0] + '" (default ' + JSON.stringify(pair[1]) + ').')
        : ('"' + pair[0] + '" already exists in 02 Settings — left exactly as set.'));
    });
  }

  report.push('');
  report.push('No existing column, row, formula, or value was renamed, reordered, or overwritten.');
  report.push('The v1.1-era menu tools (Archive Semester / Start New Semester / Update Semester');
  report.push('Dates / Duplicate Semester) and "17 Archive" are untouched and continue to work.');
  report.push('Safe to run again at any time.');
  migrationReport_('Semester Lifecycle v1.4.0 migration', report);
}
