/**
 * Marks_Intelligence_v1.4.3_IE152_Migration.gs — OPTIONAL, ONE-TIME, MANUAL
 * migration.
 *
 * Do not run this automatically. Run it yourself from the Apps Script editor
 * (runIe152NoExamV143Migration) once, after AF_Components_v1.4.1_Migration.gs
 * (needs the "Item type"/"Weight" columns it adds).
 *
 * BACKGROUND: every module the Marks Intelligence engine previously
 * understood has a real A1/A2/A3 exam sitting, feeding into the Faculty's
 * FM1=AF+A1+A2 / FM2=AF+A1+A3 / FM3=AF+A2+A3 routes. Industrial Engineering
 * (IE 152) has NO exam at all — per its own module framework slide, the
 * Final Mark is:
 *
 *   FM = 0.8 × (0.30×Project1 + 0.30×Project2 + 0.30×Project3 + 0.10×Project4)
 *      + 0.2 × (0.5×BestQuiz + 0.5×SecondBestQuiz)     [best 2 of 3 quizzes]
 *
 * Four projects (group work), three quizzes (individual, SUNLearn) — no
 * other assessments. FM1/FM2/FM3 can never resolve for a module like this
 * (they structurally require A1 AND A2 to be marked "written"), so without
 * this migration IE 152 would sit permanently at "Insufficient data" on
 * Marks Tracker no matter how many projects/quizzes you enter.
 *
 * This migration marks IE 152 in "04 Module Rules" as a pure-AF module (AF
 * weighting = 100%, A1/A2 weighting = 0%, Rule status = Verified). API.js
 * (mi_computeFmpAndOfficial_) recognises that exact combination and treats
 * AF % as the Official Final Mark directly — live, with no manual "Published
 * Final" re-entry needed every time it changes.
 *
 * The flattened FM formula above collapses to a single weighted average —
 * exactly what the "10 AF Components" weighted-AF-% engine
 * (AF_Components_v1.4.1_Migration.gs) already computes — using these Weight
 * values on each item (SUMPRODUCT-normalized, so only the weights of
 * INCLUDED rows matter, not a fixed total):
 *
 *   Project 1  -- Weight 24   Project 2 -- Weight 24   Project 3 -- Weight 24
 *   Project 4  -- Weight 8
 *   Quiz 1/2/3 -- Weight 10 each (but only the BEST 2 should ever be
 *                 "Included in AF" — see below)
 *
 * This migration does NOT create those 7 rows for you — exactly like every
 * other module, "10 AF Components" rows are created as you actually enter
 * each mark via "+ Add mark" on the AF screen, not pre-seeded for
 * assessments that haven't happened yet. When you add each one, set Item
 * type to "Project" or "Quiz" (added to the item-type list by this same
 * release, see AF_Components_v1.4.1_Migration.js) and Weight to the value
 * above. IE 152's own "AF item type" (04 Module Rules) is deliberately left
 * BLANK by this migration — unlike the other 5 modules, it genuinely mixes
 * two item types, so the "+ Add mark" dropdown stays unlocked and offers
 * both.
 *
 * BEST 2 OF 3 QUIZZES: the AF weighting engine averages whatever you mark
 * "Included in AF" — it has no automatic "drop the lowest" logic. Once all
 * three quizzes are written, manually untick "Included in AF" on whichever
 * one scored lowest, so only the best two count. (A deliberate, disclosed
 * choice — see the conversation this shipped from — over building
 * module-specific "best N of M" logic into the shared AF engine.)
 *
 * Idempotent: safe to run more than once. Every "04 Module Rules" cell is
 * only ever set if currently blank — a value you've already entered
 * yourself (including having set this up differently) is never overwritten.
 * If no module with "Industrial Engineering" in its name exists yet in
 * "03 Modules", this migration does nothing (safe to re-run once it does).
 */
function runIe152NoExamV143Migration() {
  var report = [];
  var ss = SpreadsheetApp.getActive();

  var modSheet = ss.getSheetByName(MODULES_SHEET);
  if (!modSheet) {
    migrationReport_('IE 152 (no-exam module) v1.4.3 migration', ['"' + MODULES_SHEET + '" sheet not found — nothing to do.']);
    return;
  }
  var modMap = getColMap_(modSheet, HEADER_ROW);
  var modLastRow = modSheet.getLastRow();
  var nameCol = modMap['Module name'], codeCol = modMap['Module code'];
  var targetName = null, targetCode = null;
  if (nameCol && codeCol && modLastRow > HEADER_ROW) {
    var modVals = modSheet.getRange(HEADER_ROW + 1, 1, modLastRow - HEADER_ROW, modSheet.getLastColumn()).getValues();
    for (var i = 0; i < modVals.length; i++) {
      var nm = (modVals[i][nameCol - 1] || '').toString();
      if (nm.toLowerCase().indexOf('industrial engineering') !== -1) { targetName = nm; targetCode = modVals[i][codeCol - 1]; break; }
    }
  }
  if (!targetName) {
    migrationReport_('IE 152 (no-exam module) v1.4.3 migration', [
      'No module with "Industrial Engineering" in its name was found in "' + MODULES_SHEET + '" — nothing changed.',
      'Add the module there first (Module name containing "Industrial Engineering"), then re-run this migration.'
    ]);
    return;
  }
  report.push('Found module: "' + targetName + '" (' + targetCode + ').');

  var mrSheet = ss.getSheetByName(MODULE_RULES_SHEET);
  if (!mrSheet) {
    migrationReport_('IE 152 (no-exam module) v1.4.3 migration', report.concat(['"' + MODULE_RULES_SHEET + '" sheet not found — skipped.']));
    return;
  }
  var mrMap = getColMap_(mrSheet, HEADER_ROW);
  var mrLastRow = mrSheet.getLastRow();
  var mrCodeCol = mrMap['Module code'];
  if (!mrCodeCol) {
    migrationReport_('IE 152 (no-exam module) v1.4.3 migration', report.concat(['"Module code" column not found on "' + MODULE_RULES_SHEET + '" — skipped.']));
    return;
  }
  var mrRow = -1;
  if (mrLastRow > HEADER_ROW) {
    for (var row = HEADER_ROW + 1; row <= mrLastRow; row++) {
      if (mrSheet.getRange(row, mrCodeCol).getValue() === targetCode) { mrRow = row; break; }
    }
  }
  if (mrRow === -1) {
    mrRow = firstBlankRow_(mrSheet, mrCodeCol);
    mrSheet.getRange(mrRow, mrCodeCol).setValue(targetCode);
    if (mrMap['Module']) mrSheet.getRange(mrRow, mrMap['Module']).setValue(targetName);
    report.push('Added a new row to "' + MODULE_RULES_SHEET + '" for ' + targetName + '.');
  }

  function setIfBlank(header, value, label) {
    if (!mrMap[header]) { report.push('Column "' + header + '" not found on "' + MODULE_RULES_SHEET + '" — skipped ' + label + '.'); return; }
    var cell = mrSheet.getRange(mrRow, mrMap[header]);
    var current = cell.getValue();
    if (current === '' || current === null || typeof current === 'undefined') {
      cell.setValue(value);
      report.push('Set ' + label + ' = ' + value + '.');
    } else {
      report.push(label + ' already set to "' + current + '" — left unchanged.');
    }
  }

  setIfBlank('AF weighting', 100, 'AF weighting');
  setIfBlank('A1 weighting', 0, 'A1 weighting');
  setIfBlank('A2 weighting', 0, 'A2 weighting');
  setIfBlank('Rule status', 'Verified', 'Rule status');
  setIfBlank('AF calculation method',
    'No exam — Final Mark = weighted average of 4 Projects (Weight 24/24/24/8 in 10 AF Components) + best 2 of 3 Quizzes (Weight 10 each). AF % IS the Official Final Mark.',
    'AF calculation method');

  report.push('');
  report.push(targetName + ' now has no A1/A2/A3 exam configured — Official Final Mark equals AF % directly (API.js: mi_computeFmpAndOfficial_), no more "Insufficient data" once AF % has a value.');
  report.push('"AF item type" (04 Module Rules) is deliberately left BLANK for this module — it uses both "Project" and "Quiz", not one fixed type, so the "+ Add mark" dropdown stays unlocked.');
  report.push('This migration does not add any rows to "10 AF Components" — add each Project/Quiz mark yourself via "+ Add mark" as it happens, using Weight 24/24/24/8 for Projects 1-4 and Weight 10 for each Quiz. Once all 3 quizzes are written, untick "Included in AF" on the lowest-scoring one so only the best 2 count.');
  report.push('Safe to run again at any time.');
  migrationReport_('IE 152 (no-exam module) v1.4.3 migration', report);
}
