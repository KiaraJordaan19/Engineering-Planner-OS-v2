/**
 * Marks_Intelligence_v1.4.6_DropLowest2_Migration.gs — OPTIONAL, ONE-TIME,
 * MANUAL migration.
 *
 * Do not run this automatically. Run it yourself from the Apps Script editor
 * (runDropLowest2V146Migration) once, after AF_Components_v1.4.1_Migration.gs
 * (needs "10 AF Components"' Weight/Item type columns and the AF % formula
 * it introduced).
 *
 * BACKGROUND: the Faculty of Engineering Assessment Rules document (§10.6,
 * "AF" section) states: for a semester module that gives typically 10 or
 * more tutorial-linked assessments, (a) a mark of 0 is recorded for any one
 * not done (no excuses), and (b) EACH STUDENT'S TWO LOWEST MARKS among those
 * assessments are omitted from the semester-mark calculation. This is a
 * standing Faculty rule, not a module-specific option.
 *
 * Eligibility (10+ tutorial-linked assessments per semester) was first
 * estimated from the yearbook's teaching-load numbers (L/P/T per week) times
 * its own confirmed 12-week semester -- that gave a CONTACT-PERIOD count of
 * 24-36 per module, which is only a ceiling, since not every tutorial/
 * practical period necessarily ends in a graded test. The user's own
 * first-hand count (~12 actual tutorial tests per module in a typical
 * semester) is the better source and supersedes that estimate here. Either
 * way, all five modules below clear the "10 or more" threshold:
 *   Applied Mathematics B154 (20753-154), Electrotechnique 143 (12599-143),
 *   Engineering Mathematics 145 (38571-145), Computer Programming 143
 *   (30317-143), Strength of Materials 143 (19712-143).
 * Industrial Engineering 152 is excluded on purpose -- it is project+quiz
 * based (WAF=100%, confirmed in the Faculty document's Appendix A.1), it has
 * no tutorial-test AF component for this rule to apply to.
 *
 * WHAT THIS ADDS:
 *  1. "04 Module Rules": "Drop lowest 2 tutorial marks (TRUE/FALSE)" -- set
 *     TRUE for the 5 modules above.
 *  2. "10 AF Components": "Dropped (lowest 2 rule)" -- a formula column,
 *     TRUE for exactly the 2 lowest-scoring Included-in-AF rows of a module
 *     that has the flag above set, per student, recomputed live as marks
 *     come in. Ties (two rows with the identical percentage) are broken by
 *     sheet row order -- in the rare case of an exact tie at the boundary,
 *     the earlier-entered row is treated as "lower"; this is a disclosed
 *     simplification, not expected to matter in practice since percentages
 *     rarely land on the exact same value.
 *  3. "11 Marks Tracker": "AF %" formula rewritten to also exclude any row
 *     marked "Dropped (lowest 2 rule)" -- supersedes the v1.4.1 formula,
 *     same as v1.4.1 superseded v1.4.0's; expected and safe.
 *
 * Idempotent: columns are only ever added if missing; the module-level flag
 * is only ever set if currently blank; the "Dropped" and "AF %" formulas are
 * unconditionally re-applied every run (so a future correction to this
 * migration is picked up by re-running), same convention as v1.4.1/v1.4.0.
 */
function runDropLowest2V146Migration() {
  var report = [];
  var ss = SpreadsheetApp.getActive();
  var flaggedCodes = ["20753-154", "12599-143", "38571-145", "30317-143", "19712-143"];

  // ---- 1. "04 Module Rules": the per-module flag ----
  var mrSheet = ss.getSheetByName(MODULE_RULES_SHEET);
  if (!mrSheet) {
    migrationReport_('Drop lowest 2 v1.4.6 migration', ['"' + MODULE_RULES_SHEET + '" sheet not found — nothing to do.']);
    return;
  }
  var flagResult = ensureColumn_(mrSheet, HEADER_ROW, 'Drop lowest 2 tutorial marks (TRUE/FALSE)');
  report.push(flagResult.added ? 'Added column to "' + MODULE_RULES_SHEET + '": "Drop lowest 2 tutorial marks (TRUE/FALSE)".' : 'Already present on "' + MODULE_RULES_SHEET + '": "Drop lowest 2 tutorial marks (TRUE/FALSE)".');
  if (flagResult.added) ensureCheckboxColumn_(mrSheet, HEADER_ROW, 'Drop lowest 2 tutorial marks (TRUE/FALSE)');

  var mrMap = getColMap_(mrSheet, HEADER_ROW);
  var mrLastRow = mrSheet.getLastRow();
  var mrCodeCol = mrMap['Module code'], mrFlagCol = mrMap['Drop lowest 2 tutorial marks (TRUE/FALSE)'], mrNameCol = mrMap['Module'];
  if (mrCodeCol && mrFlagCol) {
    flaggedCodes.forEach(function (code) {
      for (var row = HEADER_ROW + 1; row <= mrLastRow; row++) {
        if (mrSheet.getRange(row, mrCodeCol).getValue() !== code) continue;
        var current = mrSheet.getRange(row, mrFlagCol).getValue();
        if (current !== true && current !== false) {
          mrSheet.getRange(row, mrFlagCol).setValue(true);
          report.push(code + ': set "Drop lowest 2 tutorial marks" = TRUE.');
        } else {
          report.push(code + ': "Drop lowest 2 tutorial marks" already set to ' + current + ' — left unchanged.');
        }
        return;
      }
      report.push(code + ': no "04 Module Rules" row found — skipped.');
    });
  }

  // ---- 2. "10 AF Components": the "Dropped (lowest 2 rule)" helper column ----
  var afSheet = ss.getSheetByName(AF_COMPONENTS_SHEET);
  if (!afSheet) {
    migrationReport_('Drop lowest 2 v1.4.6 migration', report.concat(['"' + AF_COMPONENTS_SHEET + '" sheet not found — skipped the rest.']));
    return;
  }
  var droppedResult = ensureColumn_(afSheet, HEADER_ROW, 'Dropped (lowest 2 rule)');
  report.push(droppedResult.added ? 'Added column to "' + AF_COMPONENTS_SHEET + '": "Dropped (lowest 2 rule)".' : 'Already present on "' + AF_COMPONENTS_SHEET + '": "Dropped (lowest 2 rule)".');

  var afMap = getColMap_(afSheet, HEADER_ROW);
  var afRequired = ['Module', 'Percentage', 'Included in AF', 'Dropped (lowest 2 rule)'];
  var afMissing = afRequired.filter(function (h) { return !afMap[h]; });
  if (afMissing.length) {
    report.push('Could not write the "Dropped (lowest 2 rule)" formula — missing column(s): ' + afMissing.join(', ') + '.');
    report.push('Safe to run again at any time.');
    migrationReport_('Drop lowest 2 v1.4.6 migration', report);
    return;
  }
  if (!mrNameCol || !mrFlagCol) {
    report.push('Could not write the "Dropped (lowest 2 rule)" formula — "' + MODULE_RULES_SHEET + '" is missing its "Module" or flag column.');
    report.push('Safe to run again at any time.');
    migrationReport_('Drop lowest 2 v1.4.6 migration', report);
    return;
  }

  var afModuleCol = colLetter_(afMap['Module']);
  var afPctCol = colLetter_(afMap['Percentage']);
  var afIncludedCol = colLetter_(afMap['Included in AF']);
  var afFirstRow = HEADER_ROW + 1;
  var afLastRow = Math.max(afSheet.getLastRow(), afFirstRow + 200);
  var afModuleRange = "$" + afModuleCol + "$" + afFirstRow + ":$" + afModuleCol + "$" + afLastRow;
  var afIncludedRange = "$" + afIncludedCol + "$" + afFirstRow + ":$" + afIncludedCol + "$" + afLastRow;
  var afPctRange = "$" + afPctCol + "$" + afFirstRow + ":$" + afPctCol + "$" + afLastRow;

  var mrModuleRange = "'" + MODULE_RULES_SHEET + "'!$" + colLetter_(mrNameCol) + "$" + (HEADER_ROW + 1) + ":$" + colLetter_(mrNameCol) + "$" + mrLastRow;
  var mrFlagRange = "'" + MODULE_RULES_SHEET + "'!$" + colLetter_(mrFlagCol) + "$" + (HEADER_ROW + 1) + ":$" + colLetter_(mrFlagCol) + "$" + mrLastRow;

  var written = 0;
  for (var r = afFirstRow; r <= afLastRow; r++) {
    var moduleRef = "$" + afModuleCol + r;
    var pctRef = "$" + afPctCol + r;
    var includedRef = "$" + afIncludedCol + r;
    // INDEX/MATCH rather than VLOOKUP -- doesn't require the flag column to
    // sit to the right of the Module-name column on "04 Module Rules" (it
    // does today, since ensureColumn_ always appends new columns at the
    // end, but this doesn't depend on that holding true forever).
    var eligibleExpr = "IFERROR(INDEX(" + mrFlagRange + ",MATCH(" + moduleRef + "," + mrModuleRange + ",0)),FALSE)=TRUE";

    var priorRange = (r > afFirstRow)
      ? ("$" + afModuleCol + "$" + afFirstRow + ":$" + afModuleCol + "$" + (r - 1) + "," + moduleRef + "," + "$" + afIncludedCol + "$" + afFirstRow + ":$" + afIncludedCol + "$" + (r - 1) + ",TRUE,$" + afPctCol + "$" + afFirstRow + ":$" + afPctCol + "$" + (r - 1) + "," + pctRef)
      : null;

    var rankExpr = "COUNTIFS(" + afModuleRange + "," + moduleRef + "," + afIncludedRange + ",TRUE," + afPctRange + ",\"<\"&" + pctRef + ")"
      + (priorRange ? "+COUNTIFS(" + priorRange + ")" : "")
      + "+1";
    // Only start dropping once there are MORE than 2 included marks for
    // this module -- with only 1 or 2 marks in so far (early semester),
    // both would rank <=2 and everything would be "dropped", leaving
    // nothing to average (a #DIV/0! in the AF % cell). Recomputes live as
    // more marks come in, same as everything else here.
    var totalIncludedExpr = "COUNTIFS(" + afModuleRange + "," + moduleRef + "," + afIncludedRange + ",TRUE)>2";

    var formula = "=IF(AND(" + includedRef + "=TRUE,ISNUMBER(" + pctRef + ")," + eligibleExpr + "," + totalIncludedExpr + ")," + rankExpr + "<=2,FALSE)";
    afSheet.getRange(r, afMap['Dropped (lowest 2 rule)']).setFormula(formula);
    written++;
  }
  report.push('Wrote the "Dropped (lowest 2 rule)" formula on ' + written + ' row(s) of "' + AF_COMPONENTS_SHEET + '".');

  // ---- 3. "11 Marks Tracker": rewrite "AF %" to also exclude dropped rows ----
  var mtSheet = ss.getSheetByName(MARKS_TRACKER_SHEET);
  if (!mtSheet) {
    report.push('"' + MARKS_TRACKER_SHEET + '" sheet not found — skipped the AF % formula rewrite.');
  } else {
    var mtMap = getColMap_(mtSheet, MARKS_TRACKER_HEADER_ROW);
    var required = ['Module', 'AF %'];
    var missing = required.filter(function (h) { return !mtMap[h]; });
    var afMap2 = getColMap_(afSheet, HEADER_ROW);
    var afRequired2 = ['Module', 'Percentage', 'Included in AF', 'Weight', 'Dropped (lowest 2 rule)'];
    var afMissing2 = afRequired2.filter(function (h) { return !afMap2[h]; });
    if (missing.length || afMissing2.length) {
      report.push('Could not rewrite the AF % formula — missing column(s): ' + missing.concat(afMissing2).join(', ') + '.');
    } else {
      var wModuleCol = colLetter_(afMap2['Module']);
      var wPctCol = colLetter_(afMap2['Percentage']);
      var wIncludedCol = colLetter_(afMap2['Included in AF']);
      var wWeightCol = colLetter_(afMap2['Weight']);
      var wDroppedCol = colLetter_(afMap2['Dropped (lowest 2 rule)']);
      var wFirstRow = HEADER_ROW + 1;
      var wLastRow = Math.max(afSheet.getLastRow(), wFirstRow + 200);
      var wModuleRange = "'" + AF_COMPONENTS_SHEET + "'!$" + wModuleCol + "$" + wFirstRow + ":$" + wModuleCol + "$" + wLastRow;
      var wPctRange = "'" + AF_COMPONENTS_SHEET + "'!$" + wPctCol + "$" + wFirstRow + ":$" + wPctCol + "$" + wLastRow;
      var wIncludedRange = "'" + AF_COMPONENTS_SHEET + "'!$" + wIncludedCol + "$" + wFirstRow + ":$" + wIncludedCol + "$" + wLastRow;
      var wWeightRange = "'" + AF_COMPONENTS_SHEET + "'!$" + wWeightCol + "$" + wFirstRow + ":$" + wWeightCol + "$" + wLastRow;
      var wDroppedRange = "'" + AF_COMPONENTS_SHEET + "'!$" + wDroppedCol + "$" + wFirstRow + ":$" + wDroppedCol + "$" + wLastRow;

      var mtModuleCol = colLetter_(mtMap['Module']);
      var mtLastRow = mtSheet.getLastRow();
      var rewritten = 0;
      for (var row2 = MARKS_TRACKER_HEADER_ROW + 1; row2 <= mtLastRow; row2++) {
        var moduleName = mtSheet.getRange(row2, mtMap['Module']).getValue();
        if (!moduleName) continue;
        var moduleRef2 = "$" + mtModuleCol + row2;
        var notDroppedExpr = "(" + wDroppedRange + "<>TRUE)";
        var countExpr = "COUNTIFS(" + wModuleRange + "," + moduleRef2 + "," + wIncludedRange + ",TRUE)";
        var weightExpr = "IF(" + wWeightRange + "=\"\",1," + wWeightRange + ")";
        var numeratorExpr = "SUMPRODUCT((" + wModuleRange + "=" + moduleRef2 + ")*(" + wIncludedRange + "=TRUE)*" + notDroppedExpr + "*" + weightExpr + "*" + wPctRange + ")";
        var denominatorExpr = "SUMPRODUCT((" + wModuleRange + "=" + moduleRef2 + ")*(" + wIncludedRange + "=TRUE)*" + notDroppedExpr + "*" + weightExpr + ")";
        var formula2 = "=IF(" + countExpr + "=0,\"No AF marks yet\"," + numeratorExpr + "/" + denominatorExpr + ")";
        mtSheet.getRange(row2, mtMap['AF %']).setFormula(formula2);
        rewritten++;
      }
      report.push('Rewrote the "AF %" formula on ' + rewritten + ' module row(s) in "' + MARKS_TRACKER_SHEET + '": now also excludes any row marked "Dropped (lowest 2 rule)" — supersedes the v1.4.1 formula.');
    }
  }

  report.push('');
  report.push('Only affects modules with "Drop lowest 2 tutorial marks" = TRUE (the 5 above). Everything else (Industrial Engineering, and any future module you have not flagged) is completely unaffected.');
  report.push('Once you have 3+ marks logged for a flagged module, check the "Dropped (lowest 2 rule)" column: exactly 2 rows should read TRUE (your 2 lowest), the rest FALSE. If that ever looks wrong, tell me and I will re-check the formula.');
  report.push('Safe to run again at any time.');
  migrationReport_('Drop lowest 2 v1.4.6 migration', report);
}
