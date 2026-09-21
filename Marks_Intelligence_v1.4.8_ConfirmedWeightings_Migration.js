/**
 * Marks_Intelligence_v1.4.8_ConfirmedWeightings_Migration.gs — OPTIONAL,
 * ONE-TIME, MANUAL migration.
 *
 * Do not run this automatically. Run it yourself from the Apps Script editor
 * (runConfirmedWeightingsV148Migration) once.
 *
 * BACKGROUND: Applied Mathematics B154, Engineering Mathematics 145 and
 * Strength of Materials 143's AF/A1/A2 weighting values were entered into
 * "04 Module Rules" before this project's migration history began (Rule
 * status was already "Verified" the first time any of this project's
 * scripts looked at them) -- so their exact figures were never visible in
 * any file Claude could read, only in the live sheet.
 *
 * The user's own current (2026) module framework PDF for all six modules
 * (uploaded 2026-09-21) settles this with direct quotes from each module's
 * own "Calculation of final marks" section:
 *   Applied Mathematics B154 (20753-154): wAF=0.15, wA1=0.35, wA2=wA3=0.5
 *     "For this module, the weighting factors are as follows:
 *      wAF = 0.15, wA1 = 0.35, wA3 = wA2 = 0.5."
 *   Engineering Mathematics 145 (38571-145): wAF=0.1, wA1=0.35, wA2=0.55
 *     "Calculations of the final mark... WAF = 0.1  WA1 = 0.35  WA2 = 0.55"
 *   Strength of Materials 143 (19712-143): wAF=0.15, wA1=0.35, wA2=0.5
 *     "FM = 0.15WAF + 0.35WA1 + 0.50WA2"
 *
 * A3 WEIGHT: only Applied Mathematics B154 states it explicitly --
 * "wA3 = wA2 = 0.5" (its own FMp3a formula literally uses wsum = wAF+wA2+wA3
 * = 1.15, i.e. 0.15+0.5+0.5). Electrotechnique 143, Engineering Mathematics
 * 145, Computer Programming 143 and Strength of Materials 143 all confirm A3
 * EXISTS for them (exam dates in their own schedules), but none of their
 * "Calculation of final marks" sections ever states its weight -- only
 * wAF/wA1/wA2 are given. Per this project's existing design (API.js,
 * mi_getModuleWeights_): A3's weight is deliberately never assumed equal to
 * A2's or derived any other way without its own source, so it is NOT set
 * for those four here -- their Required A3 Calculator stays unavailable
 * until a real number is found (ask the module coordinator, or check
 * SUNLearn for a fuller version of each framework).
 *
 * Also confirmed directly from the same document (already implemented,
 * unaffected by this migration):
 *   - Applied Mathematics B154's own AF rule matches
 *     Marks_Intelligence_v1.4.6_DropLowest2_Migration.gs's "drop lowest 2"
 *     exactly -- "the two lowest weekly test marks will be disregarded" --
 *     EXCEPT if the student writes the optional, non-compulsory Week-1
 *     revision test, in which case it's the three lowest (including the
 *     revision test itself) that get disregarded. This migration does NOT
 *     implement that 2-vs-3 conditional -- v1.4.6 always drops exactly 2.
 *     Flagged as a known gap; tell Claude if you plan to write that
 *     optional revision test and this should be built properly.
 *   - Engineering Mathematics 145's own rule is phrased as "average of your
 *     10 best tutorial test marks" rather than "drop lowest 2" -- but its
 *     own schedule gives exactly 12 weekly tutorial tests over the
 *     semester, so "best 10 of 12" and "drop lowest 2 of 12" are the same
 *     set of marks. This only holds if all 12 get an entry (a missed test
 *     logged as 0%, not left out of "10 AF Components" entirely) -- same
 *     "no excuses, 0 awarded" baseline as every other flagged module here.
 *   - Strength of Materials 143's own document doesn't restate a drop rule
 *     (just "AF = Average (Class/Tutorial tests)"), but its own schedule
 *     confirms exactly 12 Class/Tutorial Tests -- comfortably over the
 *     Faculty's 10+ threshold (Assessment Rules §10.6), so the blanket
 *     Faculty rule applies to it even though the module's own framework
 *     doesn't spell it out.
 *
 * Idempotent: each weighting cell is only ever set if currently blank. If a
 * cell already holds a DIFFERENT number than the confirmed value above,
 * this migration does NOT overwrite it (that could be a legitimate value
 * verified against something Claude hasn't seen) -- it reports the mismatch
 * instead so you can check it.
 */
function runConfirmedWeightingsV148Migration() {
  var report = [];
  var ss = SpreadsheetApp.getActive();

  var confirmed = [
    { code: '20753-154', name: 'Applied Mathematics B154', af: 0.15, a1: 0.35, a2: 0.5, a3: 0.5 },
    { code: '38571-145', name: 'Engineering Mathematics 145', af: 0.1, a1: 0.35, a2: 0.55 },
    { code: '19712-143', name: 'Strength of Materials 143', af: 0.15, a1: 0.35, a2: 0.5 }
  ];

  var mrSheet = ss.getSheetByName(MODULE_RULES_SHEET);
  if (!mrSheet) {
    migrationReport_('Confirmed weightings v1.4.8 migration', ['"' + MODULE_RULES_SHEET + '" sheet not found — nothing to do.']);
    return;
  }
  var mrMap = getColMap_(mrSheet, HEADER_ROW);
  var mrLastRow = mrSheet.getLastRow();
  var codeCol = mrMap['Module code'];
  var afCol = mrMap['AF weighting'], a1Col = mrMap['A1 weighting'], a2Col = mrMap['A2 weighting'], a3Col = mrMap['A3 weight (%) (verified)'], statusCol = mrMap['Rule status'];
  if (!codeCol || !afCol || !a1Col || !a2Col) {
    migrationReport_('Confirmed weightings v1.4.8 migration', ['Required column(s) not found on "' + MODULE_RULES_SHEET + '" — nothing to do.']);
    return;
  }

  function setIfBlankElseCompare(cell, confirmedValue, label) {
    var current = cell.getValue();
    if (current === '' || current === null || typeof current === 'undefined') {
      cell.setValue(confirmedValue);
      report.push('  Set ' + label + ' = ' + confirmedValue + '.');
    } else if (current === confirmedValue) {
      report.push('  ' + label + ' already = ' + confirmedValue + ' — matches the confirmed value, left as-is.');
    } else {
      report.push('  ⚠ ' + label + ' is currently ' + current + ', but the module framework says ' + confirmedValue + ' — NOT changed automatically, please check this cell yourself.');
    }
  }

  confirmed.forEach(function (m) {
    var targetRow = -1;
    for (var row = HEADER_ROW + 1; row <= mrLastRow; row++) {
      if (mrSheet.getRange(row, codeCol).getValue() === m.code) { targetRow = row; break; }
    }
    if (targetRow === -1) {
      report.push(m.name + ' (' + m.code + '): no "04 Module Rules" row found — skipped.');
      return;
    }
    report.push(m.name + ' (' + m.code + '):');
    setIfBlankElseCompare(mrSheet.getRange(targetRow, afCol), m.af, 'AF weighting');
    setIfBlankElseCompare(mrSheet.getRange(targetRow, a1Col), m.a1, 'A1 weighting');
    setIfBlankElseCompare(mrSheet.getRange(targetRow, a2Col), m.a2, 'A2 weighting');
    if (typeof m.a3 === 'number') {
      if (a3Col) {
        setIfBlankElseCompare(mrSheet.getRange(targetRow, a3Col), m.a3, 'A3 weight (%) (verified)');
      } else {
        report.push('  Column "A3 weight (%) (verified)" not found — skipped (run Marks_Intelligence_v1.3.1_Migration.gs first, it added this column).');
      }
    }
    if (statusCol) {
      var statusCell = mrSheet.getRange(targetRow, statusCol);
      if (!statusCell.getValue()) {
        statusCell.setValue('Verified');
        report.push('  Set Rule status = Verified.');
      }
    }
  });

  report.push('');
  report.push('Applied Mathematics B154 now has a verified A3 weight (0.5) — its Required A3 Calculator can compute targets. Electrotechnique 143, Computer Programming 143, Engineering Mathematics 145 and Strength of Materials 143 do NOT have a confirmed A3 weight anywhere in the uploaded module frameworks — their Required A3 Calculators stay unavailable ("A3 rules incomplete") until a real number is found for them.');
  report.push('Applied Mathematics B154: remember the "drop lowest 2" rule becomes "drop lowest 3" for a student who also writes the optional Week-1 revision test — not implemented, flag to Claude if relevant.');
  report.push('Safe to run again at any time.');
  migrationReport_('Confirmed weightings v1.4.8 migration', report);
}
