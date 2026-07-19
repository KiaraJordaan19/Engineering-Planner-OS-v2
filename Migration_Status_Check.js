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

  lines.push("");
  lines.push("Legend: ✓ applied · ○ not applied. Every migration above is additive and idempotent — running an already-applied one again is safe and changes nothing.");

  migrationReport_("Migration status", lines);
}
