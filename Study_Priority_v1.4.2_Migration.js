/**
 * Study_Priority_v1.4.2_Migration.gs — OPTIONAL, ONE-TIME, MANUAL migration.
 *
 * Do not run this automatically. Run it yourself from the Apps Script
 * editor (runStudyPriorityV142Migration) once.
 *
 * BACKGROUND: the Marks Priority Engine (ai_computePriority_) requires a
 * module's Rule status to be "Verified" before it will compute anything —
 * correct for modules that genuinely have an AF/A1/A2/A3 framework, but it
 * permanently locks out a module like Industrial Engineering, which has no
 * such framework at all (tutorial-only) and can never reach "Verified."
 * That module was therefore silently invisible to "Generate study plan,"
 * with no way to include it.
 *
 * Two new, purely additive columns on "03 Modules":
 *
 *  1. "Include in study plan (TRUE/FALSE)" — a persistent per-module
 *     switch for the Study Planner/suggestion engine specifically (does
 *     NOT affect "Active status", Attendance, Calendar sync, or anything
 *     else — those are unrelated). Explicitly set to TRUE for every
 *     existing module by this migration (opt-out, not opt-in — nothing
 *     silently disappears from your study plan just from running this).
 *     A blank cell (e.g. a module added after this migration ran) is ALSO
 *     treated as TRUE by the code that reads it — see api_getPlannerData —
 *     so this is safe even before you've explicitly set it either way.
 *
 *  2. "Manual priority override" — one of Critical / Very High / High /
 *     Maintain / Low, left BLANK by default (never invented). When set,
 *     it wins outright over whatever ai_computePriority_ would otherwise
 *     compute, for ANY module — including one with a fully Verified
 *     framework, if you disagree with the computed number. This is the
 *     ONLY way a module without a framework (Rule status never
 *     "Verified") can appear in the Marks Priority Engine, the weekly
 *     study allocation, or "Generate study plan" at all — set it once for
 *     Industrial Engineering (e.g. "Maintain") and it's included from
 *     then on.
 *
 * Also, as of this release, the priority engine factors in the next
 * scheduled Tutorial session for a module (from "05 Timetable Import"),
 * not just the next A1/A2/A3 — because a tutorial is when your AF tests
 * actually get written. This needs no new column; it's computed live from
 * data that already exists.
 *
 * Idempotent: safe to run more than once. "Include in study plan" is only
 * ever set to TRUE for a module currently blank (never re-sets one you've
 * since changed to FALSE); "Manual priority override" is never touched
 * once a value exists.
 */
var VALID_PRIORITY_OVERRIDES = ["Critical", "Very High", "High", "Maintain", "Low"];

function runStudyPriorityV142Migration() {
  var report = [];
  var sheet = SpreadsheetApp.getActive().getSheetByName(MODULES_SHEET);
  if (!sheet) {
    migrationReport_('Study Priority v1.4.2 migration', ['"' + MODULES_SHEET + '" sheet not found — nothing to do.']);
    return;
  }

  var includeResult = ensureColumn_(sheet, HEADER_ROW, 'Include in study plan (TRUE/FALSE)');
  report.push(includeResult.added ? 'Added column: "Include in study plan (TRUE/FALSE)".' : 'Already present: "Include in study plan (TRUE/FALSE)".');
  if (includeResult.added) ensureCheckboxColumn_(sheet, HEADER_ROW, 'Include in study plan (TRUE/FALSE)');

  var overrideResult = ensureColumn_(sheet, HEADER_ROW, 'Manual priority override');
  report.push(overrideResult.added ? 'Added column: "Manual priority override" (left blank).' : 'Already present: "Manual priority override".');

  var map = getColMap_(sheet, HEADER_ROW);
  var lastRow = sheet.getLastRow();
  var includeCol = map['Include in study plan (TRUE/FALSE)'];
  var nameCol = map['Module name'];
  var setCount = 0;
  if (includeCol) {
    for (var row = HEADER_ROW + 1; row <= lastRow; row++) {
      var moduleName = nameCol ? sheet.getRange(row, nameCol).getValue() : "";
      if (!moduleName) continue;
      var current = sheet.getRange(row, includeCol).getValue();
      if (current !== true && current !== false) {
        sheet.getRange(row, includeCol).setValue(true);
        setCount++;
      }
    }
  }
  report.push(setCount + ' module(s) explicitly set to Include in study plan = TRUE (every module currently blank there — none excluded by running this).');

  report.push('');
  report.push('To include Industrial Engineering (or any module without a framework), set "Manual priority override" for it in "' + MODULES_SHEET + '" — Critical/Very High/High/Maintain/Low.');
  report.push('Safe to run again at any time.');
  migrationReport_('Study Priority v1.4.2 migration', report);
}
