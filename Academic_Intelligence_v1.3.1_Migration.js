/**
 * Academic_Intelligence_v1.3.0_Migration.gs — OPTIONAL, ONE-TIME, MANUAL migration.
 *
 * Do not run this automatically. Run it yourself from the Apps Script editor
 * (runAcademicIntelligenceV130Migration) once, before using the Marks Priority
 * Engine, Weekly Study Allocation, "What should I do next?", Exam Focus Mode,
 * Resource-aware rules, or Improvement Tracker introduced in v1.3.0.
 *
 * This is the ONLY schema change v1.3.0 requires. Every v1.3.0 feature listed
 * above computes entirely from data v1.2.0 already stores (Marks Tracker,
 * Module Rules, Module configuration, Assessments, Study Planner, Study
 * Tasks, Revision Tracker, Resources, Automation Log) EXCEPT for the four
 * additions below, which this migration adds:
 *
 *  1. A new, thin sheet, "22 Exam Focus Log" (only if it doesn't already
 *     exist) — the only place Exam Focus Mode's activation state lives.
 *     Headers: Module, Trigger reasons, Activated date, Deactivated date,
 *     Status. A module is "currently in Exam Focus" if it has a row with
 *     Status = "Active" and a blank Deactivated date. Activating/exiting
 *     Exam Focus never touches Study Planner rows except through the normal
 *     explicit-confirmation flow described in V1.3.0_ARCHITECTURE_NOTES.md —
 *     this sheet only ever records *that* a decision was made and when.
 *  2. One new column on "15 Resources": "Score updated" (a date). Set
 *     automatically by api_updateResourceProgress() whenever a Score is
 *     saved, going forward. Existing Score values (if any) are left with a
 *     blank "Score updated" — the Improvement Tracker's past-paper score
 *     trend only uses resources with a real date, so pre-migration scores
 *     are correctly treated as "not enough dated evidence yet" rather than
 *     guessed at.
 *  3. One new column on "13 Revision Tracker": "Weak topic since" (a date).
 *     Set automatically by api_updateRevisionTopic() the moment a topic's
 *     "Weak topic (TRUE/FALSE)" flips from false/blank to true, and cleared
 *     the moment it flips back to false. Used only to detect "repeated weak
 *     topic" (an Exam Focus Mode trigger rule) and to explain priority
 *     reasons — never used to compute marks or aggregates.
 *  4. One new row in "02 Settings" (only if the exact label doesn't already
 *     exist): "Weekly study capacity (hours)", defaulted to 14. This is the
 *     single input the Weekly Study Allocation engine treats as the
 *     student's available weekly study time; editable in Settings like any
 *     other value, exactly the same read/write pattern as every other
 *     Settings row (getSetting_ / api_saveSetting).
 *  5. Four new, OPTIONAL columns on "04 Module Rules": "Alt route name",
 *     "Alt AF weighting", "Alt A1 weighting", "Alt A2 weighting". Added
 *     entirely blank. These exist solely so the A3 Recovery Calculator can
 *     evaluate a genuine second faculty route (e.g. an exam-only or
 *     best-of-two rule some faculties publish) for a module WHERE ONE
 *     ACTUALLY EXISTS. If left blank for a module, the calculator only ever
 *     evaluates that module's one (standard AF/A1/A2/A3) route — a second
 *     route is never invented or guessed. See V1.3.0_RULE_CATALOGUE.md.
 *
 * Nothing here recomputes, backfills, or guesses a historical value for any
 * of the five additions — every one of them starts blank/default and is only
 * ever populated by real, forward-going user actions. No existing column,
 * row, or formula anywhere in the workbook is touched.
 *
 * Idempotent: safe to run more than once. Depends on
 * Migrations_Shared_v1.2.0.gs (ensureColumn_ / ensureSheetExists_ /
 * migrationReport_) being present in the same project — re-added alongside
 * this file if you already deleted your v1.2.0 migration files.
 */
function runAcademicIntelligenceV130Migration() {
  var report = [];

  // ---- 1. "22 Exam Focus Log" ----
  var examFocusResult = ensureSheetExists_('22 Exam Focus Log', HEADER_ROW, [
    'Module', 'Trigger reasons', 'Activated date', 'Deactivated date', 'Status'
  ]);
  report.push(examFocusResult.created ? 'Created new sheet: "22 Exam Focus Log".' : '"22 Exam Focus Log" already existed.');
  report.push(examFocusResult.addedHeaders.length
    ? ('Headers added: ' + examFocusResult.addedHeaders.join(', '))
    : 'All expected headers were already present.');

  // ---- 2. "15 Resources" — Score updated ----
  var resourcesSheet = SpreadsheetApp.getActive().getSheetByName('15 Resources');
  if (!resourcesSheet) {
    report.push('"15 Resources" sheet not found — skipped the "Score updated" column (run Resources_v1.2.0_Migration.gs first if this is unexpected).');
  } else {
    var scoreDateResult = ensureColumn_(resourcesSheet, HEADER_ROW, 'Score updated');
    report.push(scoreDateResult.added ? 'Added column to "15 Resources": "Score updated".' : 'Already present on "15 Resources": "Score updated".');
  }

  // ---- 3. "13 Revision Tracker" — Weak topic since ----
  var revisionSheet = SpreadsheetApp.getActive().getSheetByName('13 Revision Tracker');
  if (!revisionSheet) {
    report.push('"13 Revision Tracker" sheet not found — skipped the "Weak topic since" column.');
  } else {
    var weakSinceResult = ensureColumn_(revisionSheet, HEADER_ROW, 'Weak topic since');
    report.push(weakSinceResult.added ? 'Added column to "13 Revision Tracker": "Weak topic since".' : 'Already present on "13 Revision Tracker": "Weak topic since".');
  }

  // ---- 4. "02 Settings" — Weekly study capacity (hours) ----
  var settingsSheet = SpreadsheetApp.getActive().getSheetByName('02 Settings');
  if (!settingsSheet) {
    report.push('"02 Settings" sheet not found — skipped the "Weekly study capacity (hours)" setting.');
  } else {
    var added = ensureSettingRow_(settingsSheet, 'Weekly study capacity (hours)', 14);
    report.push(added ? 'Added new Settings row: "Weekly study capacity (hours)" (default 14).' : '"Weekly study capacity (hours)" already exists in 02 Settings — left exactly as set.');
  }

  // ---- 5. "04 Module Rules" — optional alternative faculty route ----
  var moduleRulesSheet = SpreadsheetApp.getActive().getSheetByName('04 Module Rules');
  if (!moduleRulesSheet) {
    report.push('"04 Module Rules" sheet not found — skipped the alternative faculty route columns.');
  } else {
    var altCols = ['Alt route name', 'Alt AF weighting', 'Alt A1 weighting', 'Alt A2 weighting'];
    var altAdded = [];
    altCols.forEach(function (col) {
      var r = ensureColumn_(moduleRulesSheet, HEADER_ROW, col);
      if (r.added) altAdded.push(col);
    });
    report.push(altAdded.length
      ? ('Added columns to "04 Module Rules": ' + altAdded.join(', ') + ' (left blank — fill in only for a module whose faculty genuinely publishes a second calculation route).')
      : 'Alternative faculty route columns already present on "04 Module Rules".');
  }

  report.push('');
  report.push('No existing column, row, formula, or value was renamed, reordered, or overwritten.');
  report.push('Safe to run again at any time.');
  migrationReport_('Academic Intelligence v1.3.0 migration', report);
}

/**
 * Self-contained helper (not added to Migrations_Shared_v1.2.0.gs, so that
 * shared file stays byte-identical to v1.2.0 — verifiable by checksum diff).
 * Appends one new label/value row to a Settings-shaped sheet ONLY if a row
 * with that exact label doesn't already exist anywhere in the sheet.
 * Returns true if a row was added, false if the label was already present.
 * Mirrors the exact column layout every other 02 Settings row already uses:
 * label in column A, value in column C.
 */
function ensureSettingRow_(settingsSheet, label, defaultValue) {
  var finder = settingsSheet.createTextFinder(label).matchEntireCell(true);
  var existing = finder.findNext();
  if (existing) return false;
  var lastRow = settingsSheet.getLastRow();
  var newRow = lastRow + 1;
  settingsSheet.getRange(newRow, 1).setValue(label);
  settingsSheet.getRange(newRow, 3).setValue(defaultValue);
  return true;
}
