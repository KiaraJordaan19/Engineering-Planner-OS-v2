/**
 * Marks_Intelligence_v1.3.1_Migration.gs — OPTIONAL, ONE-TIME, MANUAL migration.
 *
 * Do not run this automatically. Run it yourself from the Apps Script editor
 * (runMarksIntelligenceV131Migration) once, before using the Faculty Rule
 * Engine (FM1/FM2/FM3/FMp), Faculty Override Rules, DCA, the Assessment
 * Simulator, or the new "Published Final (Before A3)" field introduced in
 * v1.3.1 — Marks Intelligence Completion.
 *
 * This is a FOCUSED completion release: it does not touch Study Planner,
 * Resources, Attendance, Dashboard, or Semester Lifecycle schema at all.
 * Every addition below is scoped to "11 Marks Tracker" and "04 Module
 * Rules" — the same two sheets v1.3.0's Marks Intelligence work already
 * read from — plus one workbook-wide Settings default. Nothing here
 * recomputes, backfills, or guesses a historical value; every addition
 * starts blank/default and is only ever populated by a real, forward-going
 * user action or explicit configuration choice.
 *
 * v1.3.1 shipped in two passes. The 2nd pass is a CRITICAL correction to
 * the module-weighting model itself (the Engineering Faculty's real model
 * is route-based -- FM1=AF+A1+A2, FM2=AF+A1+A3, FM3=AF+A2+A3, each
 * independently renormalized -- not "A3 = 100% minus AF/A1/A2", which
 * silently produced A3=0% for every confirmed module). This single
 * migration function reflects BOTH passes together, since this is one
 * v1.3.1 release, not two versions. If you already ran an earlier build of
 * this migration (1st pass), re-running it now is safe: every step below
 * checks for its target before adding it, and the two items the 1st pass
 * added that the correction retires (the "Alt2 ..." FM3 columns and "Rule 1
 * disabled (TRUE/FALSE)") are left in place, inert, rather than deleted --
 * see "What changed from the 1st pass" below.
 *
 *  1. One new column on "11 Marks Tracker": "Published Final (Before A3)".
 *     This is DELIBERATELY SEPARATE from the pre-existing "Provisional
 *     final (after A2)" column (which is this sheet's own FORMULA output,
 *     computed only once a real A2 is on record) and from "Manual combined
 *     AF entry (%) — reference only". "Published Final (Before A3)" exists
 *     specifically for a module that does NOT publish A2 directly and
 *     instead publishes a single combined pre-A3 percentage — the student
 *     types that published figure in here, and the Hidden A2 inference
 *     engine reads it from here, never from the formula-computed
 *     "Provisional final (after A2)" column. Writing to this column NEVER
 *     touches A2, A3, "Provisional final (after A2)", or "Final /
 *     provisional (after A3)" — those remain exactly as v1.2.0/v1.3.0 left
 *     them.
 *  2. One new, REQUIRED-TO-VERIFY column on "04 Module Rules": "A3 weight
 *     (%) (verified)". This is the 2nd-pass correction's central field: A3
 *     no longer derives its weight as a remainder of AF/A1/A2 — it has its
 *     own, separately-verified weight, used only in FM2 (AF+A1+A3) and FM3
 *     (AF+A2+A3). Left entirely blank by default. Until a real value is
 *     entered here for a module, every FM2/FM3/Required-A3 calculation for
 *     that module reports "A3 rules incomplete — required-A3 and FM2/FM3
 *     calculations unavailable" rather than guessing, assuming A3's weight
 *     equals A2's, or splitting a remaining share evenly. A module MAY be
 *     configured with A3 weight equal to A2's weight, but only as an
 *     explicit verified entry — never a global assumption.
 *  3. One new, OPTIONAL column on "04 Module Rules": "A2/A3 subminimum
 *     mode". Resolves a genuine ambiguity in Faculty Override Rule 1's
 *     wording (does "A2 or A3 below 40% caps FMp at 45%" mean EACH written
 *     assessment must independently reach 40%, or does it mean the cap only
 *     applies if NEITHER reaches 40%? These give different results — e.g.
 *     A2=35%, A3=70% is capped under the first reading, not capped under
 *     the second). Left blank by default, in which case Rule 1 is NOT
 *     evaluated for that module at all (disclosed in the planner) rather
 *     than guessing between the two readings. Accepted values:
 *       - "EACH_WRITTEN_ASSESSMENT_MUST_REACH_40"
 *       - "AT_LEAST_ONE_OF_A2_OR_A3_MUST_REACH_40"
 *  4. Three new columns on "11 Marks Tracker": "A1 status", "A2 status",
 *     "A3 status". Free-text, one of Written / Not written / Excused /
 *     Deferred (case-insensitive), left blank by default. A blank status
 *     cell does NOT mean "not written" — if the cell is blank but a real
 *     mark is present, the planner safely infers "written" from the mark
 *     itself; only a blank status AND a blank mark together produce
 *     "unknown" (never "not written" by default). Explicit status always
 *     takes priority over inference once entered.
 *  5. One new column on "11 Marks Tracker": "DCA mark (%)". The 2nd-pass
 *     correction to Faculty Rule 3 (DCA): DCA is a genuine, separately
 *     awarded substitute mark, entered here — NOT automatically "the higher
 *     of A2/A3" (that was a disclosed 1st-pass judgment call, now retired
 *     as too strong an assumption). When DCA is enabled for a module (see
 *     item 6) and both A2 and A3 are known, the LOWER of the two is
 *     replaced with this entered DCA mark before any FM calculation; the
 *     original mark is preserved for audit and shown alongside the
 *     substitution everywhere it's displayed.
 *  6. Three new, OPTIONAL Faculty Override Rule / DCA columns on "04 Module
 *     Rules", each a TRUE/FALSE toggle, all defaulted to FALSE:
 *       - "Rule 2 exception (TRUE/FALSE)" — exempts this module from the
 *         "A1, A2 and A3 all written caps FMp at 50%" override.
 *       - "DCA enabled (TRUE/FALSE)" — enables Deemed Continuous Assessment
 *         substitution for this module (see item 5 above for the exact,
 *         corrected mechanics).
 *       - "A2 rounding tolerance override (%)" — lets a specific module use
 *         a tighter or looser tolerance for the Hidden A2 inference
 *         engine's "verify by recalculation" check than the workbook
 *         default (item 7 below).
 *  7. One new row in "02 Settings" (only if the exact label doesn't already
 *     exist): "Default A2 recalculation match tolerance (%)", defaulted to
 *     0.15 — the workbook-wide default consumed by the tolerance-override
 *     column above. Editable in Settings exactly like every other Settings
 *     row.
 *
 * What changed from the 1st pass (informational only — nothing here is
 * auto-deleted, since this migration never removes a column):
 *   - "Rule 1 disabled (TRUE/FALSE)" is RETIRED. Rule 1's application is now
 *     entirely governed by "A2/A3 subminimum mode" (item 3): unset = Rule 1
 *     not evaluated at all, exactly like "disabled" used to behave. If your
 *     workbook already has a "Rule 1 disabled" column with TRUE values set,
 *     those cells are simply no longer read by v1.3.1 code — replace them
 *     by leaving "A2/A3 subminimum mode" blank for that module (equivalent
 *     effect), or set a subminimum mode to opt back into Rule 1.
 *   - The four "Alt2 ..." columns (Alt2 route name / AF / A1 / A2 weighting)
 *     are RETIRED. FM3 is no longer a module-specific "optional alternative
 *     route" bolted on separately — every module now has FM3 = AF+A2+A3
 *     defined identically by the Faculty model itself, gated only on A3's
 *     weight being verified (item 2). These four columns are simply no
 *     longer read; you may delete them manually if you like, or leave them
 *     — see V1.3.1_ROLLBACK.md.
 *
 * Idempotent: safe to run more than once. Depends on
 * Migrations_Shared_v1.2.0.gs (ensureColumn_ / migrationReport_) and the
 * ensureSettingRow_ helper (from Academic_Intelligence_v1.3.0_Migration.gs
 * — both must be present in the same project since this file reuses that
 * helper rather than redefining it a second time).
 */
function runMarksIntelligenceV131Migration() {
  var report = [];

  // ---- 1. "11 Marks Tracker" — Published Final (Before A3), status cols, DCA mark ----
  var marksTrackerSheet = SpreadsheetApp.getActive().getSheetByName('11 Marks Tracker');
  if (!marksTrackerSheet) {
    report.push('"11 Marks Tracker" sheet not found — skipped every "11 Marks Tracker" item below.');
  } else {
    var pubFinalResult = ensureColumn_(marksTrackerSheet, MARKS_TRACKER_HEADER_ROW, 'Published Final (Before A3)');
    report.push(pubFinalResult.added
      ? 'Added column to "11 Marks Tracker": "Published Final (Before A3)" (left blank — enter only for a module that publishes a combined pre-A3 mark instead of a raw A2).'
      : 'Already present on "11 Marks Tracker": "Published Final (Before A3)".');

    var statusCols = ['A1 status', 'A2 status', 'A3 status'];
    var statusAdded = [];
    statusCols.forEach(function (col) {
      var r = ensureColumn_(marksTrackerSheet, MARKS_TRACKER_HEADER_ROW, col);
      if (r.added) statusAdded.push(col);
    });
    report.push(statusAdded.length
      ? ('Added columns to "11 Marks Tracker": ' + statusAdded.join(', ') + ' (left blank — a blank status with a real mark present is still safely treated as "written"; a blank status with a blank mark is treated as "unknown", never assumed "not written").')
      : 'Assessment status columns (A1/A2/A3 status) already present on "11 Marks Tracker".');

    var dcaMarkResult = ensureColumn_(marksTrackerSheet, MARKS_TRACKER_HEADER_ROW, 'DCA mark (%)');
    report.push(dcaMarkResult.added
      ? 'Added column to "11 Marks Tracker": "DCA mark (%)" (left blank — enter the genuine Deemed Continuous Assessment mark here; it substitutes for whichever of A2/A3 is lower, only when DCA is enabled for the module).'
      : 'Already present on "11 Marks Tracker": "DCA mark (%)".');
  }

  // ---- 2-6. "04 Module Rules" ----
  var moduleRulesSheet = SpreadsheetApp.getActive().getSheetByName('04 Module Rules');
  if (!moduleRulesSheet) {
    report.push('"04 Module Rules" sheet not found — skipped every "04 Module Rules" item below.');
  } else {
    var a3WeightResult = ensureColumn_(moduleRulesSheet, HEADER_ROW, 'A3 weight (%) (verified)');
    report.push(a3WeightResult.added
      ? 'Added column to "04 Module Rules": "A3 weight (%) (verified)" (left blank — REQUIRED before this module\'s FM2/FM3/Required-A3 calculations become available; A3\'s weight is never derived, never assumed equal to A2\'s, never split evenly).'
      : 'Already present on "04 Module Rules": "A3 weight (%) (verified)".');

    var subminResult = ensureColumn_(moduleRulesSheet, HEADER_ROW, 'A2/A3 subminimum mode');
    report.push(subminResult.added
      ? 'Added column to "04 Module Rules": "A2/A3 subminimum mode" (left blank — Faculty Override Rule 1 is not evaluated for a module until this is explicitly set to EACH_WRITTEN_ASSESSMENT_MUST_REACH_40 or AT_LEAST_ONE_OF_A2_OR_A3_MUST_REACH_40).'
      : 'Already present on "04 Module Rules": "A2/A3 subminimum mode".');

    var ruleCols = ['Rule 2 exception (TRUE/FALSE)', 'DCA enabled (TRUE/FALSE)'];
    var ruleAdded = [];
    ruleCols.forEach(function (col) {
      var r = ensureColumn_(moduleRulesSheet, HEADER_ROW, col);
      if (r.added) ruleAdded.push(col);
    });
    if (ruleAdded.length) {
      report.push('Added Faculty Override Rule columns to "04 Module Rules": ' + ruleAdded.join(', ') + '.');
      report.push('Both default to blank/FALSE for every existing module — Rule 2 remains ACTIVE and DCA remains OFF for every module until explicitly configured otherwise.');
    } else {
      report.push('Faculty Override Rule columns (Rule 2 exception, DCA enabled) already present on "04 Module Rules".');
    }

    var toleranceResult = ensureColumn_(moduleRulesSheet, HEADER_ROW, 'A2 rounding tolerance override (%)');
    report.push(toleranceResult.added
      ? 'Added column to "04 Module Rules": "A2 rounding tolerance override (%)" (left blank — the workbook-wide Settings default is used until a module-specific override is entered).'
      : 'Already present on "04 Module Rules": "A2 rounding tolerance override (%)".');

    report.push('Note: "Alt2 route name/AF/A1/A2 weighting" and "Rule 1 disabled (TRUE/FALSE)" (if present from an earlier build) are retired and no longer read by v1.3.1 code — see the doc comment above for what replaces them. Not deleted automatically.');
  }

  // ---- 7. "02 Settings" — default A2 recalculation match tolerance ----
  var settingsSheet = SpreadsheetApp.getActive().getSheetByName('02 Settings');
  if (!settingsSheet) {
    report.push('"02 Settings" sheet not found — skipped the default A2 recalculation tolerance setting.');
  } else {
    var addedTolerance = ensureSettingRow_(settingsSheet, 'Default A2 recalculation match tolerance (%)', 0.15);
    report.push(addedTolerance
      ? 'Added new Settings row: "Default A2 recalculation match tolerance (%)" (default 0.15).'
      : '"Default A2 recalculation match tolerance (%)" already exists in 02 Settings — left exactly as set.');
  }

  report.push('');
  report.push('No existing column, row, formula, or value was renamed, reordered, or overwritten.');
  report.push('No Study Planner, Resources, Attendance, Dashboard, or Semester Lifecycle schema was touched.');
  report.push('Safe to run again at any time.');
  migrationReport_('Marks Intelligence v1.3.1 migration', report);
}
