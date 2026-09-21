/**
 * Marks_Intelligence_v1.4.5_SubminimumMode_Migration.gs — OPTIONAL,
 * ONE-TIME, MANUAL migration.
 *
 * Do not run this automatically. Run it yourself from the Apps Script editor
 * (runSubminimumModeV145Migration) once.
 *
 * BACKGROUND: mi_evaluateOverrideCaps_ (API.js) has always supported two
 * readings of Faculty Rule 1 ("A2<40 OR A3<40" vs "max(A2,A3)<40"), gated on
 * a per-module "A2/A3 subminimum mode" column that was left blank for every
 * module -- Rule 1 was therefore never actually evaluated for anyone, on
 * purpose, until the correct reading could be confirmed from source.
 *
 * The Faculty of Engineering Assessment Rules document (2022-09-22 edition,
 * §4.3.2.1) settles it: "If a student did not achieve at least 40 in A2 or
 * A3, FMp may not exceed 45" -- i.e. EITHER written assessment below 40
 * triggers the cap, matching the "EACH_WRITTEN_ASSESSMENT_MUST_REACH_40"
 * reading. §4.3 applies uniformly to every standard semester module with
 * A1/A2/A3 -- it is not a per-module choice, only Appendix A/B exceptions
 * deviate from it, and Appendix A/B (checked in full) lists no exception for
 * any of the 5 modules below.
 *
 * Sets "A2/A3 subminimum mode" = "EACH_WRITTEN_ASSESSMENT_MUST_REACH_40" for:
 *   Applied Mathematics B154, Electrotechnique 143, Engineering Mathematics
 *   145, Computer Programming 143, Strength of Materials 143.
 * Industrial Engineering 152 is excluded on purpose -- it has no A1/A2/A3 at
 * all (WAF=100%, confirmed in the same document's Appendix A.1), so Rule 1
 * (which is entirely about A2/A3) does not apply to it.
 *
 * Idempotent: only sets the cell if currently blank -- never overwrites a
 * value you (or a future correction) already set.
 */
function runSubminimumModeV145Migration() {
  var report = [];
  var ss = SpreadsheetApp.getActive();
  var codes = ["20753-154", "12599-143", "38571-145", "30317-143", "19712-143"];

  var mrSheet = ss.getSheetByName(MODULE_RULES_SHEET);
  if (!mrSheet) {
    migrationReport_('Subminimum mode v1.4.5 migration', ['"' + MODULE_RULES_SHEET + '" sheet not found — nothing to do.']);
    return;
  }
  var mrMap = getColMap_(mrSheet, HEADER_ROW);
  var mrLastRow = mrSheet.getLastRow();
  var codeCol = mrMap["Module code"], modeCol = mrMap["A2/A3 subminimum mode"];
  if (!codeCol || !modeCol) {
    migrationReport_('Subminimum mode v1.4.5 migration', ['"Module code" or "A2/A3 subminimum mode" column not found on "' + MODULE_RULES_SHEET + '" — run Marks_Intelligence_v1.3.1_Migration.gs first (it added this column).']);
    return;
  }

  codes.forEach(function (code) {
    var found = false;
    for (var row = HEADER_ROW + 1; row <= mrLastRow; row++) {
      if (mrSheet.getRange(row, codeCol).getValue() !== code) continue;
      found = true;
      var cell = mrSheet.getRange(row, modeCol);
      var current = cell.getValue();
      if (!current) {
        cell.setValue("EACH_WRITTEN_ASSESSMENT_MUST_REACH_40");
        report.push(code + ': set "A2/A3 subminimum mode" = EACH_WRITTEN_ASSESSMENT_MUST_REACH_40.');
      } else {
        report.push(code + ': "A2/A3 subminimum mode" already set to "' + current + '" — left unchanged.');
      }
      break;
    }
    if (!found) report.push(code + ': no "04 Module Rules" row found — skipped.');
  });

  report.push('');
  report.push('Rule 1 (the 45% cap) now actually evaluates for these 5 modules — previously it was configured-but-inactive since neither reading had been confirmed. Whether it ever triggers depends on your actual A2/A3 marks; this migration only turns on the correct rule, it does not change any recorded mark.');
  report.push('Industrial Engineering 152 intentionally left untouched — it has no A2/A3, so Rule 1 never applies to it.');
  report.push('Safe to run again at any time.');
  migrationReport_('Subminimum mode v1.4.5 migration', report);
}
