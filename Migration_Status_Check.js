/**
 * Migration_Status_Check.gs — READ-ONLY. Never writes to the spreadsheet.
 *
 * Answers one question: "which of this project's optional migrations have
 * already been run on THIS spreadsheet?" Every migration in this project
 * (Modules/Resources/Revision/Study Tasks/Attendance v1.2.0, Revision
 * Contribution v1.2.0 RC2, Academic Intelligence v1.3.1, Marks Intelligence
 * v1.3.1, Semester Lifecycle v1.4.0) is additive and idempotent — running
 * one that's already applied is always safe — but there was previously no
 * way to see current status without opening the Apps Script editor and
 * reading each migration file's doc comment by hand.
 *
 * Checks one representative marker column/sheet per migration rather than
 * every column each one adds — enough to say "not run" / "already run",
 * not a full diff. Run from the "Academic Planner" menu ("Check Migration
 * Status") or directly as runMigrationStatusCheck().
 */
function runMigrationStatusCheck() {
  var ss = SpreadsheetApp.getActive();
  var lines = [];

  function checkColumn(label, sheetName, headerRow, columnName, runFn) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) { lines.push("? " + label + " — sheet \"" + sheetName + "\" not found."); return; }
    var map = getColMap_(sheet);
    var applied = !!map[columnName];
    lines.push((applied ? "✓ " : "○ ") + label + (applied ? " — already applied." : " — NOT applied yet. Run " + runFn + "()."));
  }

  function checkSheet(label, sheetName, runFn) {
    var applied = !!ss.getSheetByName(sheetName);
    lines.push((applied ? "✓ " : "○ ") + label + (applied ? " — already applied." : " — NOT applied yet. Run " + runFn + "()."));
  }

  checkColumn("Modules v1.2.0", MODULES_SHEET, HEADER_ROW, "Repeated module (TRUE/FALSE)", "runModulesV120Migration");
  checkColumn("Resources v1.2.0", RESOURCES_SHEET, HEADER_ROW, "Resource ID", "runResourcesV120Migration");
  checkColumn("Revision v1.2.0", REVISION_SHEET, HEADER_ROW, "Weak topic (TRUE/FALSE)", "runRevisionV120Migration");
  checkColumn("Revision Contribution v1.2.0 RC2", STUDY_SHEET, HEADER_ROW, "Revision contribution applied", "runRevisionContributionV120RC2Migration");
  checkSheet("Study Tasks v1.2.0", "21 Study Tasks", "runStudyTasksV120Migration");
  checkColumn("Attendance v1.2.0", ATTENDANCE_SHEET, HEADER_ROW, "Attendance ID", "runAttendanceV120Migration");
  checkSheet("Academic Intelligence v1.3.1", EXAM_FOCUS_SHEET, "runAcademicIntelligenceV130Migration");
  checkColumn("Marks Intelligence v1.3.1", MARKS_TRACKER_SHEET, MARKS_TRACKER_HEADER_ROW, "Published Final (Before A3)", "runMarksIntelligenceV131Migration");
  checkColumn("Marks Intelligence v1.4.0 (AF engine + Published Final After A3)", MARKS_TRACKER_SHEET, MARKS_TRACKER_HEADER_ROW, "Published Final (After A3)", "runMarksIntelligenceV140Migration");
  checkSheet("Semester Lifecycle v1.4.0", "23 Semester Archive", "runSemesterLifecycleV140Migration");
  checkColumn("Drive Resources v1.4.0", RESOURCES_SHEET, HEADER_ROW, "Drive category", "runDriveResourcesV140Migration");
  checkColumn("Reminders Feed v1.4.0", ASSIGNMENTS_SHEET, HEADER_ROW, "Synced to Reminders", "runRemindersFeedV140Migration");
  checkColumn("AF Components v1.4.1 (item type + weighting)", AF_COMPONENTS_SHEET, HEADER_ROW, "Weight", "runAfComponentsV141Migration");
  checkColumn("Study Priority v1.4.2 (include toggle + manual priority)", MODULES_SHEET, HEADER_ROW, "Manual priority override", "runStudyPriorityV142Migration");
  checkIe152NoExamStatus_(ss, lines);

  lines.push("");
  lines.push("Legend: ✓ applied · ○ not applied. Every migration above is additive and idempotent — running an already-applied one again is safe and changes nothing.");

  migrationReport_("Migration status", lines);
}

/** IE 152's marker isn't a new column -- it's AF weighting=1/A1=0/A2=0 on
 *  its own "04 Module Rules" row (weightings are stored as fractions of 1,
 *  same scale as every other module) -- so this checks that combination
 *  directly rather than via checkColumn/checkSheet above. */
function checkIe152NoExamStatus_(ss, lines) {
  var label = "IE 152 no-exam module v1.4.3";
  var modSheet = ss.getSheetByName(MODULES_SHEET);
  var mrSheet = ss.getSheetByName(MODULE_RULES_SHEET);
  if (!modSheet || !mrSheet) { lines.push("? " + label + " — required sheet(s) not found."); return; }
  var modMap = getColMap_(modSheet, HEADER_ROW);
  var modLastRow = modSheet.getLastRow();
  var nameCol = modMap["Module name"], codeCol = modMap["Module code"];
  var targetCode = null;
  if (nameCol && codeCol && modLastRow > HEADER_ROW) {
    var modVals = modSheet.getRange(HEADER_ROW + 1, 1, modLastRow - HEADER_ROW, modSheet.getLastColumn()).getValues();
    for (var i = 0; i < modVals.length; i++) {
      if ((modVals[i][nameCol - 1] || "").toString().toLowerCase().indexOf("industrial engineering") !== -1) { targetCode = modVals[i][codeCol - 1]; break; }
    }
  }
  if (!targetCode) { lines.push("○ " + label + " — no module with \"Industrial Engineering\" in its name found yet. Run runIe152NoExamV143Migration() once it's added."); return; }
  var mrMap = getColMap_(mrSheet, HEADER_ROW);
  var mrLastRow = mrSheet.getLastRow();
  var mrCodeCol = mrMap["Module code"];
  var applied = false;
  if (mrCodeCol && mrLastRow > HEADER_ROW) {
    for (var row = HEADER_ROW + 1; row <= mrLastRow; row++) {
      if (mrSheet.getRange(row, mrCodeCol).getValue() !== targetCode) continue;
      applied = mrMap["AF weighting"] && mrSheet.getRange(row, mrMap["AF weighting"]).getValue() === 1
        && mrMap["A1 weighting"] && mrSheet.getRange(row, mrMap["A1 weighting"]).getValue() === 0
        && mrMap["A2 weighting"] && mrSheet.getRange(row, mrMap["A2 weighting"]).getValue() === 0;
      break;
    }
  }
  lines.push((applied ? "✓ " : "○ ") + label + (applied ? " — already applied." : " — NOT applied yet. Run runIe152NoExamV143Migration()."));
}
