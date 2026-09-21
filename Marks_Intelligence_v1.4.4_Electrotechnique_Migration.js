/**
 * Marks_Intelligence_v1.4.4_Electrotechnique_Migration.gs — OPTIONAL,
 * ONE-TIME, MANUAL migration.
 *
 * Do not run this automatically. Run it yourself from the Apps Script editor
 * (runElectrotechniqueV144Migration) once.
 *
 * BACKGROUND: AF_Components_v1.4.1_Migration.gs auto-set Electrotechnique
 * 143 (12599-143)'s "AF item type" (04 Module Rules) to "Tutorial tests",
 * locking its "+ Add mark" dropdown to that single option. The official
 * yearbook entry for this module lists a teaching load of 3.5L; 1P; 2T --
 * it has a real weekly practical period, not just tutorials. Per user
 * confirmation (2026-09-21 audit): the practical and tutorial-test sessions
 * alternate week to week, and marks from both genuinely count toward AF --
 * neither is a separate, independently-weighted final-mark component.
 *
 * Since "Item type" is purely a label for the "+ Add mark" form (it plays
 * no part in the AF % SUMPRODUCT formula -- only "Weight" and "Included in
 * AF" do), the correct, lowest-risk fix is to UNLOCK it: leave the
 * module-level "AF item type" blank, exactly like Industrial Engineering's
 * Project/Quiz mix, so each AF Components row can be logged as whichever of
 * "Tutorial test" or "Practical" it actually was. The AF % calculation is
 * unaffected either way.
 *
 * This migration also removes "12599-143" from AF_Components_v1.4.1_
 * Migration.gs's knownItemType map, so re-running that older migration
 * won't re-lock this module back to "Tutorial tests" only.
 *
 * Idempotent: only clears the cell if it is still exactly "Tutorial tests"
 * (the value the older migration itself set) -- never touches a different
 * value, including one you've since set deliberately.
 */
function runElectrotechniqueV144Migration() {
  var report = [];
  var ss = SpreadsheetApp.getActive();
  var targetCode = "12599-143";

  var mrSheet = ss.getSheetByName(MODULE_RULES_SHEET);
  if (!mrSheet) {
    migrationReport_('Electrotechnique item type v1.4.4 migration', ['"' + MODULE_RULES_SHEET + '" sheet not found — nothing to do.']);
    return;
  }
  var mrMap = getColMap_(mrSheet, HEADER_ROW);
  var mrLastRow = mrSheet.getLastRow();
  var codeCol = mrMap["Module code"], itemTypeCol = mrMap["AF item type"];
  if (!codeCol || !itemTypeCol) {
    migrationReport_('Electrotechnique item type v1.4.4 migration', ['"Module code" or "AF item type" column not found on "' + MODULE_RULES_SHEET + '" — run AF_Components_v1.4.1_Migration.gs first.']);
    return;
  }

  var found = false;
  for (var row = HEADER_ROW + 1; row <= mrLastRow; row++) {
    if (mrSheet.getRange(row, codeCol).getValue() !== targetCode) continue;
    found = true;
    var cell = mrSheet.getRange(row, itemTypeCol);
    var current = cell.getValue();
    if (current === "Tutorial tests") {
      cell.setValue("");
      report.push('Cleared "AF item type" for ' + targetCode + ' (Electrotechnique) — was "Tutorial tests", now blank/unlocked.');
    } else if (!current) {
      report.push('"AF item type" for ' + targetCode + ' is already blank — nothing to change.');
    } else {
      report.push('"AF item type" for ' + targetCode + ' is set to "' + current + '" (not the auto-set "Tutorial tests") — left unchanged, this looks like a deliberate choice.');
    }
    break;
  }
  if (!found) {
    report.push('No "04 Module Rules" row found for ' + targetCode + ' (Electrotechnique) — nothing to do.');
  } else {
    report.push('');
    report.push('Electrotechnique\'s "+ Add mark" Item type dropdown now offers both "Tutorial test" and "Practical" — log each AF entry as whichever it actually was. Both count toward AF % the same way (via Weight), so this only affects labeling, not the calculation.');
  }
  report.push('Safe to run again at any time.');
  migrationReport_('Electrotechnique item type v1.4.4 migration', report);
}
