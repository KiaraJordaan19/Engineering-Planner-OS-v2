/**
 * AF_Components_v1.4.1_Migration.gs — OPTIONAL, ONE-TIME, MANUAL migration.
 *
 * Do not run this automatically. Run it yourself from the Apps Script
 * editor (runAfComponentsV141Migration) once. Depends on
 * Marks_Intelligence_v1.4.0_Migration.gs having already been run (this
 * migration rewrites the same "AF %" formula that one introduced, so run
 * that one first if you haven't).
 *
 * Two related additions, both scoped to "10 AF Components" and "04 Module
 * Rules":
 *
 *  1. "AF item type" on "04 Module Rules": Tutorial tests / Practicals.
 *     Auto-populated for the 5 modules already covered by
 *     Marks_Intelligence_v1.4.0_Migration.gs, read straight off each
 *     module's own "AF calculation method" text (never guessed for any
 *     other module — left blank, meaning "not yet chosen," until you set
 *     it yourself). Only ever set if the cell is currently blank, so a
 *     value you've already entered by hand is never overwritten.
 *     Once set, the "+ Add Assessment" form's Item type choice for that
 *     module is locked to the matching option — a module you've marked
 *     "Practicals" only ever offers "Practical," never "Tutorial test."
 *
 *  2. "Item type" and "Weight" on "10 AF Components" — both purely
 *     additive, both optional. "Item type" mirrors the module-level
 *     setting above, per row (Tutorial test / Practical). "Weight" is a
 *     plain number, left BLANK by default: a blank-weight test counts as
 *     weight 1 (an ordinary, equal-contribution test) in the AF average,
 *     mixed freely with any test you DO give an explicit weight — give one
 *     test a weight of 2 and it counts twice as much as a blank/weight-1
 *     test, entirely relative, not a percentage that has to add up to
 *     anything. Never setting a single weight behaves exactly like a plain
 *     average, same as v1.4.0.
 *
 * The "AF %" formula on "11 Marks Tracker" is rewritten to a weighted
 * average (SUMPRODUCT) over every "Included in AF" = TRUE row for that
 * module, using each row's Weight (blank treated as 1) — this SUPERSEDES
 * the plain-average version v1.4.0 wrote; running this migration after
 * v1.4.0 is expected and safe.
 *
 * Idempotent: safe to run more than once. The "AF item type" / "Item type"
 * / "Weight" columns are only ever added if missing, never re-added; the
 * module-level "AF item type" default is only ever set if currently blank;
 * the "AF %" formula is unconditionally re-applied every run (so a further
 * correction to this migration in the future is picked up by re-running),
 * same convention Marks_Intelligence_v1.4.0_Migration.gs already uses.
 */
// v1.4.3 -- "Project" and "Quiz" added for modules like Industrial
// Engineering (IE 152) that are assessed entirely by group projects and
// individual quizzes, no tutorial tests or practicals at all. See
// Marks_Intelligence_v1.4.3_IE152_Migration.gs.
var VALID_AF_ITEM_TYPES = ["Tutorial test", "Practical", "Project", "Quiz"];
var VALID_MODULE_AF_ITEM_TYPES = ["Tutorial tests", "Practicals"];

function runAfComponentsV141Migration() {
  var report = [];
  var ss = SpreadsheetApp.getActive();

  // ---- 1. "04 Module Rules": AF item type ----
  var mrSheet = ss.getSheetByName(MODULE_RULES_SHEET);
  if (!mrSheet) {
    report.push('"' + MODULE_RULES_SHEET + '" sheet not found — skipped "AF item type".');
  } else {
    var itemTypeCol = ensureColumn_(mrSheet, HEADER_ROW, "AF item type");
    report.push(itemTypeCol.added ? 'Added column to "' + MODULE_RULES_SHEET + '": "AF item type".' : 'Already present on "' + MODULE_RULES_SHEET + '": "AF item type".');

    var mrMap = getColMap_(mrSheet, HEADER_ROW);
    var mrLastRow = mrSheet.getLastRow();
    var codeCol = mrMap["Module code"];
    // Read straight from each module's own already-entered "AF calculation
    // method" text (set by Marks_Intelligence_v1.4.0_Migration.gs / your
    // module frameworks) -- never invented independently of it.
    var knownItemType = {
      "20753-154": "Tutorial tests",   // Applied Mathematics B154
      "12599-143": "Tutorial tests",   // Electrotechniques 143
      "38571-145": "Tutorial tests",   // Engineering Mathematics 145
      "30317-143": "Practicals",       // Computer Programming 143 -- "Average of practical test marks"
      "19712-143": "Tutorial tests"    // Strength of Materials 143
    };
    if (codeCol && itemTypeCol) {
      Object.keys(knownItemType).forEach(function (code) {
        for (var row = HEADER_ROW + 1; row <= mrLastRow; row++) {
          if (mrSheet.getRange(row, codeCol).getValue() !== code) continue;
          var current = mrSheet.getRange(row, mrMap["AF item type"]).getValue();
          if (!current) {
            mrSheet.getRange(row, mrMap["AF item type"]).setValue(knownItemType[code]);
            report.push(code + ': set AF item type = "' + knownItemType[code] + '".');
          } else {
            report.push(code + ': AF item type already set to "' + current + '" — left unchanged.');
          }
          break;
        }
      });
    }
    report.push('Every other module\'s "AF item type" is left blank — set it yourself in "' + MODULE_RULES_SHEET + '" (Tutorial tests / Practicals) once you know which applies.');
  }

  // ---- 2. "10 AF Components": Item type, Weight ----
  var afSheet = ss.getSheetByName(AF_COMPONENTS_SHEET);
  if (!afSheet) {
    report.push('"' + AF_COMPONENTS_SHEET + '" sheet not found — skipped "Item type"/"Weight" and the AF % formula rewrite.');
  } else {
    ['Item type', 'Weight'].forEach(function (name) {
      var result = ensureColumn_(afSheet, HEADER_ROW, name);
      report.push((result.added ? 'Added column to "' : 'Already present on "') + AF_COMPONENTS_SHEET + '": "' + name + '".');
    });

    // ---- 3. Rewrite "AF %" on "11 Marks Tracker" as a weighted average ----
    var mtSheet = ss.getSheetByName(MARKS_TRACKER_SHEET);
    if (!mtSheet) {
      report.push('"' + MARKS_TRACKER_SHEET + '" sheet not found — skipped the AF % formula rewrite.');
    } else {
      var mtMap = getColMap_(mtSheet, MARKS_TRACKER_HEADER_ROW);
      var afMap = getColMap_(afSheet, HEADER_ROW);
      var required = ['Module', 'AF %'];
      var missing = required.filter(function (h) { return !mtMap[h]; });
      var afRequired = ['Module', 'Percentage', 'Included in AF', 'Weight'];
      var afMissing = afRequired.filter(function (h) { return !afMap[h]; });
      if (missing.length || afMissing.length) {
        report.push('Could not rewrite the AF % formula — missing column(s): ' + missing.concat(afMissing).join(', ') + '.');
      } else {
        var afModuleCol = colLetter_(afMap['Module']);
        var afPctCol = colLetter_(afMap['Percentage']);
        var afIncludedCol = colLetter_(afMap['Included in AF']);
        var afWeightCol = colLetter_(afMap['Weight']);
        var afFirstRow = HEADER_ROW + 1;
        var afLastRow = Math.max(afSheet.getLastRow(), afFirstRow + 200);
        var afModuleRange = "'" + AF_COMPONENTS_SHEET + "'!$" + afModuleCol + "$" + afFirstRow + ":$" + afModuleCol + "$" + afLastRow;
        var afPctRange = "'" + AF_COMPONENTS_SHEET + "'!$" + afPctCol + "$" + afFirstRow + ":$" + afPctCol + "$" + afLastRow;
        var afIncludedRange = "'" + AF_COMPONENTS_SHEET + "'!$" + afIncludedCol + "$" + afFirstRow + ":$" + afIncludedCol + "$" + afLastRow;
        var afWeightRange = "'" + AF_COMPONENTS_SHEET + "'!$" + afWeightCol + "$" + afFirstRow + ":$" + afWeightCol + "$" + afLastRow;

        var mtModuleCol = colLetter_(mtMap['Module']);
        var mtLastRow = mtSheet.getLastRow();
        var rewritten = 0;
        for (var row = MARKS_TRACKER_HEADER_ROW + 1; row <= mtLastRow; row++) {
          var moduleName = mtSheet.getRange(row, mtMap['Module']).getValue();
          if (!moduleName) continue;
          var moduleRef = "$" + mtModuleCol + row;
          var countExpr = "COUNTIFS(" + afModuleRange + "," + moduleRef + "," + afIncludedRange + ",TRUE())";
          var weightExpr = "IF(" + afWeightRange + "=\"\",1," + afWeightRange + ")";
          var numeratorExpr = "SUMPRODUCT((" + afModuleRange + "=" + moduleRef + ")*(" + afIncludedRange + "=TRUE)*" + weightExpr + "*" + afPctRange + ")";
          var denominatorExpr = "SUMPRODUCT((" + afModuleRange + "=" + moduleRef + ")*(" + afIncludedRange + "=TRUE)*" + weightExpr + ")";
          var formula = "=IF(" + countExpr + "=0,\"No AF marks yet\"," + numeratorExpr + "/" + denominatorExpr + ")";
          mtSheet.getRange(row, mtMap['AF %']).setFormula(formula);
          rewritten++;
        }
        report.push('Rewrote the "AF %" formula on ' + rewritten + ' module row(s) in "' + MARKS_TRACKER_SHEET +
          '": now a weighted average (blank Weight = 1) of "' + AF_COMPONENTS_SHEET + '" entries marked Included in AF = TRUE — supersedes the plain-average version from Marks_Intelligence_v1.4.0_Migration.gs.');
      }
    }
  }

  report.push('');
  report.push('Nothing here changes what counts as "Included in AF" -- exclusion is still the Written/Excused checkboxes, exactly as v1.4.0 left it.');
  report.push('Safe to run again at any time.');
  migrationReport_('AF Components v1.4.1 migration', report);
}
