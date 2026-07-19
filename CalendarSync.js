/**
 * Calendar_Sync_v1.1.gs — Engineering Academic Planner, Google Sheets Edition
 * Hardening pass v1.1. Replaces Calendar_Sync.gs.
 *
 * One-way sync: Google Sheets -> Google Calendar -> (already visible in Apple Calendar).
 * The workbook is the source of truth. Never reads calendar edits back in, never
 * auto-deletes events, never creates duplicates on repeated runs.
 *
 * SETUP
 * 1. Extensions > Apps Script > paste this file in (replace Code.gs, or add as new file).
 * 2. Reload the spreadsheet. An "Academic Planner" menu appears (built by onOpen).
 * 3. Academic Planner > Validate Workbook — run this FIRST. It assigns stable IDs to any
 *    row missing one, and converts the TRUE/FALSE checkbox columns into real clickable
 *    Google Sheets checkboxes (an xlsx import cannot create native checkboxes by itself).
 * 4. Academic Planner > Sync All Calendar Items — run once to authorize Calendar access.
 * 5. In 02 Settings, "Planner Calendar ID (read/write)" accepts either an exact Calendar ID
 *    (recommended — Calendar > Settings > [calendar] > Integrate calendar > Calendar ID)
 *    or a calendar name (used only as a fallback; if more than one calendar shares that
 *    name, sync stops and reports the ambiguity rather than guessing).
 *
 * v1.2.0 note: everything above and below this note, up to the "v1.2.0 ACADEMIC
 * WORKFLOW RELEASE" section near the end, is byte-for-byte the same v1.1.12
 * code — nothing was removed, renamed, or restructured. v1.2.0's Study
 * Planner / Revision Tracker / Resources / Attendance / Module Config work is
 * additive, appended at the end of this file.
 */

var TIMEZONE = "Africa/Johannesburg";
var LOCK_WAIT_MS = 15000;

var SETTINGS_SHEET = "02 Settings";
var ASSIGNMENTS_SHEET = "08 Assignments";
var ASSESSMENTS_SHEET = "09 Assessments";
var INBOX_SHEET = "07 Academic Inbox";
var STUDY_SHEET = "12 Study Planner";
var TIMETABLE_SHEET = "05 Timetable Import";
// v1.2.0 -- new sheet constant for the thin Study Tasks checklist store (see
// Study_Tasks_v1.2.0_Migration.gs). Declared here, alongside the other
// sheet-name constants already in this file, rather than in API.gs, purely
// to keep every "var X_SHEET = ..." declaration in one place for this half
// of the split; API.gs's own sheet constants (MODULES_SHEET, RESOURCES_SHEET,
// etc.) are unaffected and unchanged.
var STUDY_TASKS_SHEET = "21 Study Tasks";

// ============================================================
// MENU
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi().createMenu("Academic Planner")
    .addItem("Process Academic Inbox", "processAcademicInbox")
    .addItem("Sync Selected Row", "syncSelectedRow")
    .addItem("Sync All Calendar Items", "syncAllToCalendar")
    .addSeparator()
    .addItem("Generate Study Suggestions", "generateStudySuggestions")
    .addSeparator()
    .addItem("Enable Automatic Sync (Inbox + Calendar + Notifications)", "enableAutomation")
    .addItem("Disable Automatic Sync", "disableAutomation")
    .addSeparator()
    .addItem("Archive Semester", "archiveSemester")
    .addItem("Start New Semester", "startNewSemester")
    .addItem("Update Semester Dates", "updateSemesterDates")
    .addItem("Duplicate Semester (new file copy)", "duplicateSemester")
    .addSeparator()
    .addItem("Validate Workbook", "validateWorkbook")
    .addItem("Check Migration Status", "runMigrationStatusCheck")
    .addItem("Delete Calendar Event for Selected Row", "deleteCalendarEventForSelectedRow")
    .addItem("Open Setup Guide", "openSetupGuide")
    .addToUi();
}

// ============================================================
// AUTOMATION TRIGGERS  ("only touch one place" workflow)
// ============================================================
var MANAGED_FUNCTIONS = ["processAcademicInbox", "syncAllToCalendar", "sendSmartNotifications"];

function enableAutomation() {
  disableAutomation(); // clear any existing managed triggers first, so this is always idempotent
  ScriptApp.newTrigger("processAcademicInbox").timeBased().everyMinutes(15).create();
  ScriptApp.newTrigger("syncAllToCalendar").timeBased().everyMinutes(30).create();
  ScriptApp.newTrigger("sendSmartNotifications").timeBased().everyDays(1).atHour(7).inTimezone(TIMEZONE).create();
  SpreadsheetApp.getUi().alert("Automatic sync enabled",
    "Academic Inbox processes every 15 min, Calendar sync runs every 30 min, and a daily 07:00 email summary is scheduled. " +
    "From now on you should only need to type things into the workbook (or add events to your personal Google Calendar) — " +
    "everything else keeps itself up to date. Disable any time from this menu.",
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function disableAutomation() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  triggers.forEach(function (t) {
    if (MANAGED_FUNCTIONS.indexOf(t.getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  if (removed > 0) SpreadsheetApp.getActive().toast(removed + " automation trigger(s) removed.", "Academic Planner", 5);
}

// ============================================================
// SMART NOTIFICATIONS  (daily email summary)
// ============================================================
function sendSmartNotifications() {
  var ss = SpreadsheetApp.getActive();
  var email = Session.getActiveUser().getEmail();
  if (!email) return; // no identifiable user email (e.g. some trigger contexts) — skip silently

  var lines = [];
  var today = new Date();
  var todayStr = Utilities.formatDate(today, TIMEZONE, "yyyy-MM-dd");

  // Assessments due in 7 days or tomorrow
  try {
    var assess = ss.getSheetByName(ASSESSMENTS_SHEET);
    var aMap = getColMap_(assess);
    var lastRow = assess.getLastRow();
    var dates = assess.getRange(HEADER_ROW + 1, col_(aMap, "Date"), lastRow - HEADER_ROW, 1).getValues();
    var modules = assess.getRange(HEADER_ROW + 1, col_(aMap, "Module"), lastRow - HEADER_ROW, 1).getValues();
    var types = assess.getRange(HEADER_ROW + 1, col_(aMap, "Assessment type"), lastRow - HEADER_ROW, 1).getValues();
    for (var i = 0; i < dates.length; i++) {
      var d = dates[i][0];
      if (!(d instanceof Date)) continue;
      var days = Math.round((d - today) / 86400000);
      if (days === 1) lines.push("⚠ Tomorrow: " + modules[i][0] + " " + types[i][0]);
      else if (days === 7) lines.push("📅 In 7 days: " + modules[i][0] + " " + types[i][0]);
    }
  } catch (e) { lines.push("(Could not check assessments: " + e.message + ")"); }

  // Overdue assignments
  try {
    var asg = ss.getSheetByName(ASSIGNMENTS_SHEET);
    var sMap = getColMap_(asg);
    var lastRow2 = asg.getLastRow();
    var dueDates = asg.getRange(HEADER_ROW + 1, col_(sMap, "Due date"), lastRow2 - HEADER_ROW, 1).getValues();
    var statuses = asg.getRange(HEADER_ROW + 1, col_(sMap, "Status"), lastRow2 - HEADER_ROW, 1).getValues();
    var titles = asg.getRange(HEADER_ROW + 1, col_(sMap, "Title"), lastRow2 - HEADER_ROW, 1).getValues();
    for (var j = 0; j < dueDates.length; j++) {
      var dd = dueDates[j][0];
      if (!(dd instanceof Date)) continue;
      var st = statuses[j][0];
      if (dd < today && st !== "Submitted" && st !== "Graded" && st !== "Cancelled") {
        lines.push("🔴 Overdue: " + titles[j][0]);
      }
    }
  } catch (e) { lines.push("(Could not check assignments: " + e.message + ")"); }

  // Today's tutorials (from the real timetable)
  try {
    var tt = ss.getSheetByName(TIMETABLE_SHEET);
    var tMap = getColMap_(tt);
    var lastRow3 = tt.getLastRow();
    var ttDates = tt.getRange(HEADER_ROW + 1, col_(tMap, "Start date"), lastRow3 - HEADER_ROW, 1).getValues();
    var ttTypes = tt.getRange(HEADER_ROW + 1, col_(tMap, "Session type"), lastRow3 - HEADER_ROW, 1).getValues();
    var ttModules = tt.getRange(HEADER_ROW + 1, col_(tMap, "Module"), lastRow3 - HEADER_ROW, 1).getValues();
    for (var k = 0; k < ttDates.length; k++) {
      var td = ttDates[k][0];
      if (!(td instanceof Date)) continue;
      if (Utilities.formatDate(td, TIMEZONE, "yyyy-MM-dd") === todayStr && ttTypes[k][0] === "Tutorial") {
        lines.push("📘 Tutorial today: " + ttModules[k][0]);
      }
    }
  } catch (e) { lines.push("(Could not check timetable: " + e.message + ")"); }

  if (lines.length === 0) return; // nothing worth emailing today — stay quiet, don't spam
  var body = "Academic Planner daily summary (" + todayStr + "):\n\n" + lines.join("\n") +
    "\n\nOpen the workbook for full details. This email is sent by your own Calendar_Sync automation.";
  MailApp.sendEmail(email, "Academic Planner — daily summary", body);
}


// ============================================================
// HEADER-NAME COLUMN MAPPING  (never hardcode column numbers)
// ============================================================
var HEADER_ROW = 4;

// `headerRow` is optional and defaults to the workbook-wide HEADER_ROW (4) --
// every sheet uses that except "11 Marks Tracker", whose real header row is
// MARKS_TRACKER_HEADER_ROW (6, extra title/note rows above it). Callers
// working with Marks Tracker MUST pass MARKS_TRACKER_HEADER_ROW explicitly;
// omitting it there silently maps almost nothing (row 4 has one unrelated
// note in it, not real headers) and every col_() lookup after it throws
// "column not found" -- this was happening in api_setPublishedFinalBeforeA3,
// api_setAssessmentStatus, api_setDcaMark, and both archiveSemester
// functions until this fix.
function getColMap_(sheet, headerRow) {  var lastColumn = sheet.getLastColumn();
  if (lastColumn < 1) {
    return {};
  }
  var lastCol = sheet.getLastColumn();
  var row = headerRow || HEADER_ROW;
  var headers = sheet.getRange(row, 1, 1, lastCol).getValues()[0];
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    var h = (headers[i] || "").toString().trim();
    if (h) map[h] = i + 1; // 1-based column index
  }
  return map;
}

function col_(map, name) {
  if (!map[name]) throw new Error('Expected column "' + name + '" not found — has a header been renamed?');
  return map[name];
}

// ============================================================
// SETTINGS HELPERS
// ============================================================
function getSetting_(label) {
  var ss = SpreadsheetApp.getActive();
  var settings = ss.getSheetByName(SETTINGS_SHEET);
  var finder = settings.createTextFinder(label).matchEntireCell(false);
  var cell = finder.findNext();
  if (!cell) return null;
  return settings.getRange(cell.getRow(), 3).getValue(); // Value column = C
}

function syncEnabled_() {
  var v = getSetting_("Calendar sync enabled");
  return v === true;
}

function getTargetCalendar_() {
  var idOrName = getSetting_("Planner Calendar ID (read/write)");
  if (!idOrName) throw new Error("No Google Calendar name/ID set in 02 Settings.");

  // Prefer treating the value as an exact Calendar ID first.
  var byId = null;
  try { byId = CalendarApp.getCalendarById(idOrName); } catch (e) { byId = null; }
  if (byId) return byId;

  // Fall back to name lookup, but refuse to guess if ambiguous.
  var byName = CalendarApp.getCalendarsByName(idOrName);
  if (byName.length === 0) {
    throw new Error("No calendar found for '" + idOrName + "'. Create it in Google Calendar, " +
      "or use its exact Calendar ID in 02 Settings. Nothing was guessed.");
  }
  if (byName.length > 1) {
    throw new Error("Ambiguous: " + byName.length + " calendars are named '" + idOrName +
      "'. Use the exact Calendar ID in 02 Settings instead (Calendar Settings > Integrate calendar).");
  }
  return byName[0];
}

// ============================================================
// v1.4.0 -- GOOGLE CALENDAR PULL-SYNC (monthly view only)
// ============================================================
// Every sync elsewhere in this file is one-way, Sheets -> Google Calendar
// (see the file header). This is the one read path: it lets the monthly
// calendar screen show events that exist ON the configured Google Calendar
// but were never created by this workbook -- e.g. something added directly
// in Google Calendar, or via Academic Inbox -> "Send to Calendar" outside
// the normal Assignments/Assessments flow. Events already represented by an
// Assignments/Assessments/Study Planner row (tracked via that row's own
// "Calendar Event ID") are excluded so nothing is shown twice.
function getTrackedCalendarEventIds_() {
  var ids = {};
  [ASSIGNMENTS_SHEET, ASSESSMENTS_SHEET, STUDY_SHEET].forEach(function (sheetName) {
    var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    if (!sheet) return;
    var map = getColMap_(sheet);
    if (!map["Calendar Event ID"]) return;
    var lastRow = sheet.getLastRow();
    if (lastRow <= HEADER_ROW) return;
    var vals = sheet.getRange(HEADER_ROW + 1, map["Calendar Event ID"], lastRow - HEADER_ROW, 1).getValues();
    vals.forEach(function (r) { if (r[0]) ids[r[0]] = true; });
  });
  return ids;
}

/**
 * Read-only. `monthStart`/`monthEndExclusive` bound one calendar month.
 * Returns { events: [{id,title,date,allDay}], reason } -- `reason` is a
 * human-readable explanation (never a thrown error) when nothing could be
 * fetched, e.g. no calendar configured yet, so the monthly view can show a
 * quiet inline message instead of failing the whole screen load.
 */
function buildExternalCalendarEvents_(monthStart, monthEndExclusive) {
  var idOrName = getSetting_("Planner Calendar ID (read/write)");
  if (!idOrName) return { events: [], reason: "No Google Calendar configured yet (\"Planner Calendar ID (read/write)\" in 02 Settings)." };
  var calendar;
  try { calendar = getTargetCalendar_(); } catch (e) { return { events: [], reason: e.message }; }

  var tracked = getTrackedCalendarEventIds_();
  var calEvents;
  try { calEvents = calendar.getEvents(monthStart, monthEndExclusive); } catch (e) { return { events: [], reason: "Could not read from Google Calendar: " + e.message }; }

  var out = [];
  calEvents.forEach(function (ev) {
    var id = ev.getId();
    if (tracked[id]) return; // already shown via its own Assignments/Assessments/Study Planner row
    out.push({
      id: id, title: ev.getTitle(),
      date: Utilities.formatDate(ev.getStartTime(), TIMEZONE, "yyyy-MM-dd"),
      allDay: ev.isAllDayEvent()
    });
  });
  return { events: out, reason: null };
}

function checkTimezones_() {
  var ss = SpreadsheetApp.getActive();
  var ssTz = ss.getSpreadsheetTimeZone();
  var scriptTz = Session.getScriptTimeZone();
  var warnings = [];
  if (ssTz !== TIMEZONE) warnings.push("Spreadsheet timezone is " + ssTz + ", expected " + TIMEZONE + ".");
  if (scriptTz !== TIMEZONE) warnings.push("Script timezone is " + scriptTz + ", expected " + TIMEZONE + ".");
  return warnings;
}

// ============================================================
// STABLE IDs  (generated once, never re-derived from row position)
// ============================================================
function nextId_(prefix, existingIds) {
  var max = 0;
  existingIds.forEach(function (id) {
    var m = /^-?(\d+)$/.exec((id || "").toString().replace(prefix + "-", ""));
    if (m) max = Math.max(max, parseInt(m[1], 10));
  });
  return prefix + "-" + Utilities.formatString("%03d", max + 1);
}

function assignStableIds_(sheetName, idHeader, keyHeader, prefix) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sheet) return 0;
  var map = getColMap_(sheet);
  var idCol = col_(map, idHeader);
  var keyCol = col_(map, keyHeader);
  var lastRow = sheet.getLastRow();
  var ids = sheet.getRange(HEADER_ROW + 1, idCol, lastRow - HEADER_ROW, 1).getValues().map(function (r) { return r[0]; });
  var keys = sheet.getRange(HEADER_ROW + 1, keyCol, lastRow - HEADER_ROW, 1).getValues().map(function (r) { return r[0]; });
  var assigned = 0;
  for (var i = 0; i < ids.length; i++) {
    if (!ids[i] && keys[i]) {
      var newId = nextId_(prefix, ids);
      sheet.getRange(HEADER_ROW + 1 + i, idCol).setValue(newId);
      ids[i] = newId;
      assigned++;
    }
  }
  return assigned;
}

// ============================================================
// CHECKBOXES  (xlsx import cannot create native Sheets checkboxes — this does)
// ============================================================
var CHECKBOX_TARGETS = [
  { sheet: "00 Start Here", cols: null, range: "C9:C15" },
  { sheet: ASSIGNMENTS_SHEET, headers: ["Sync to Calendar"] },
  { sheet: ASSESSMENTS_SHEET, headers: ["Sync to Calendar", "Revision started"] },
  { sheet: INBOX_SHEET, headers: ["Processed", "Send to Calendar"] },
  { sheet: STUDY_SHEET, headers: ["Planned", "Completed", "Sync to Calendar"] },
  { sheet: "13 Revision Tracker", headers: ["Completed"] },
  { sheet: "15 Resources", headers: ["Reviewed"] },
  { sheet: "10 AF Components", headers: ["Written", "Excused or excluded"] },
  { sheet: "19 Attendance", headers: ["Attended"] }
];

function insertCheckboxes_() {
  var ss = SpreadsheetApp.getActive();
  CHECKBOX_TARGETS.forEach(function (t) {
    var sheet = ss.getSheetByName(t.sheet);
    if (!sheet) return;
    var lastRow = Math.max(sheet.getLastRow(), HEADER_ROW + 60);
    if (t.range) {
      sheet.getRange(t.range).insertCheckboxes();
      return;
    }
    var map = getColMap_(sheet);
    t.headers.forEach(function (h) {
      if (!map[h]) return;
      var c = map[h];
      sheet.getRange(HEADER_ROW + 1, c, lastRow - HEADER_ROW, 1).insertCheckboxes();
    });
  });
}

// ============================================================
// VALIDATE WORKBOOK  (menu entry point — run this first)
// ============================================================
function validateWorkbook() {
  var report = [];
  try {
    var n1 = assignStableIds_(ASSIGNMENTS_SHEET, "Assignment ID", "Title", "ASG");
    report.push("Assignments: " + n1 + " new ID(s) assigned.");
  } catch (e) { report.push("Assignments ID check skipped: " + e.message); }
  try {
    var n2 = assignStableIds_("10 AF Components", "AF Item ID", "Module", "AF");
    report.push("AF Components: " + n2 + " new ID(s) assigned.");
  } catch (e) { report.push("AF Components ID check skipped: " + e.message); }
  try {
    var n3 = assignStableIds_(STUDY_SHEET, "Study Session ID", "Module", "SS");
    report.push("Study Planner: " + n3 + " new ID(s) assigned.");
  } catch (e) { report.push("Study Planner ID check skipped: " + e.message); }
  try {
    var n4 = assignStableIds_(INBOX_SHEET, "Inbox ID", "Title", "INB");
    report.push("Academic Inbox: " + n4 + " new ID(s) assigned.");
  } catch (e) { report.push("Inbox ID check skipped: " + e.message); }

  try {
    insertCheckboxes_();
    report.push("Checkboxes: converted to native Sheets checkboxes.");
  } catch (e) { report.push("Checkbox conversion issue: " + e.message); }

  var tzWarnings = checkTimezones_();
  if (tzWarnings.length) report.push("Timezone warning: " + tzWarnings.join(" "));
  else report.push("Timezone: spreadsheet and script both Africa/Johannesburg.");

  var dupWarnings = findDuplicateIds_();
  if (dupWarnings.length) report.push("Duplicate ID warning: " + dupWarnings.join("; "));
  else report.push("No duplicate stable IDs found.");

  SpreadsheetApp.getUi().alert("Validate Workbook", report.join("\n"), SpreadsheetApp.getUi().ButtonSet.OK);
}

function findDuplicateIds_() {
  var checks = [
    [ASSIGNMENTS_SHEET, "Assignment ID"], ["10 AF Components", "AF Item ID"],
    [STUDY_SHEET, "Study Session ID"], [INBOX_SHEET, "Inbox ID"], [ASSESSMENTS_SHEET, "Assessment ID"]
  ];
  var warnings = [];
  var ss = SpreadsheetApp.getActive();
  checks.forEach(function (c) {
    var sheet = ss.getSheetByName(c[0]);
    if (!sheet) return;
    var map = getColMap_(sheet);
    if (!map[c[1]]) return;
    var col = map[c[1]];
    var lastRow = sheet.getLastRow();
    var vals = sheet.getRange(HEADER_ROW + 1, col, lastRow - HEADER_ROW, 1).getValues()
      .map(function (r) { return r[0]; }).filter(function (v) { return v; });
    var seen = {}, dupes = {};
    vals.forEach(function (v) { seen[v] = (seen[v] || 0) + 1; if (seen[v] > 1) dupes[v] = true; });
    var dupeKeys = Object.keys(dupes);
    if (dupeKeys.length) warnings.push(c[0] + ": " + dupeKeys.join(", "));
  });
  return warnings;
}

function openSetupGuide() {
  SpreadsheetApp.getUi().alert("Setup Guide",
    "See the '18 User Guide' sheet, and Setup_Guide_and_Formula_Reference.md alongside this workbook.",
    SpreadsheetApp.getUi().ButtonSet.OK);
}

// ============================================================
// ACADEMIC INBOX PROCESSING
// ============================================================
/**
 * Classifies and files ONE Academic Inbox row (already confirmed not yet
 * processed) into Assignments or Assessments, marks it Processed, and
 * records where it landed. This is the SAME classification/row-creation
 * logic previously inlined in processAcademicInbox()'s loop -- extracted
 * here, unchanged, so both the bulk trigger below and the new targeted
 * per-item processor (processOneInboxItemById_, added in v1.1.12) share
 * one single source of truth for "what type goes where." Never call this
 * on a row that hasn't already been checked for Processed=true / missing
 * Title-or-Module -- callers do that check themselves so this function
 * always either successfully files the row or returns null.
 *
 * Returns null if the row is already processed or missing Title/Module
 * (nothing to do). Otherwise returns
 * { destSheetName, destRow, destId, filedType, sendToCalendar }.
 */
function processOneInboxRow_(inbox, map, row) {
  var processedCell = inbox.getRange(row, col_(map, "Processed"));
  if (processedCell.getValue() === true) return null; // never double-process

  var title = inbox.getRange(row, col_(map, "Title")).getValue();
  var module = inbox.getRange(row, col_(map, "Module")).getValue();
  if (!title || !module) return null; // nothing to process yet

  var itemType = inbox.getRange(row, col_(map, "Item type")).getValue();
  var dueDate = inbox.getRange(row, col_(map, "Due date")).getValue();
  var dueTime = inbox.getRange(row, col_(map, "Due time")).getValue();
  var venueLink = inbox.getRange(row, col_(map, "Venue or link")).getValue();
  var notes = inbox.getRange(row, col_(map, "Notes")).getValue();
  var priority = inbox.getRange(row, col_(map, "Priority")).getValue();
  var sendToCal = inbox.getRange(row, col_(map, "Send to Calendar")).getValue();
  var inboxId = inbox.getRange(row, col_(map, "Inbox ID")).getValue();

  var destSheetName, destRow, destIdHeader, filedType;
  if (itemType === "Assignment" || itemType === "Project" || itemType === "Presentation") {
    destRow = appendAssignmentFromInbox_(module, title, notes, dueDate, dueTime, priority, sendToCal, inboxId);
    destSheetName = ASSIGNMENTS_SHEET;
    destIdHeader = "Assignment ID";
    filedType = "Assignment";
  } else {
    destRow = appendAssessmentFromInbox_(module, itemType, title, dueDate, dueTime, venueLink, priority, sendToCal, inboxId);
    destSheetName = ASSESSMENTS_SHEET;
    destIdHeader = "Assessment ID";
    filedType = "Assessment";
  }

  processedCell.setValue(true);
  inbox.getRange(row, col_(map, "Processed into (sheet!row)")).setValue(destSheetName + "!" + destRow);

  var destSheet = SpreadsheetApp.getActive().getSheetByName(destSheetName);
  var destMap = getColMap_(destSheet);
  var destId = destSheet.getRange(destRow, col_(destMap, destIdHeader)).getValue();

  return {
    destSheetName: destSheetName,
    destRow: destRow,
    destId: destId,
    filedType: filedType,
    sendToCalendar: sendToCal === true
  };
}

function processAcademicInbox() {
  var ss = SpreadsheetApp.getActive();
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(LOCK_WAIT_MS)) {
    SpreadsheetApp.getUi().alert("Another sync/process run is in progress. Try again shortly.");
    return;
  }
  try {
    assignStableIds_(INBOX_SHEET, "Inbox ID", "Title", "INB");
    var inbox = ss.getSheetByName(INBOX_SHEET);
    var map = getColMap_(inbox);
    var lastRow = inbox.getLastRow();
    var processedCount = 0;

    for (var row = HEADER_ROW + 1; row <= lastRow; row++) {
      var filed = processOneInboxRow_(inbox, map, row);
      if (filed) processedCount++;
    }
    SpreadsheetApp.getActive().toast(processedCount + " inbox item(s) processed.", "Academic Inbox", 6);
  } finally {
    lock.releaseLock();
  }
}

/**
 * v1.1.12 -- targeted per-item processing, added for the Academic Inbox
 * "process immediately on add" usability feature. Finds exactly one row by
 * its stable Inbox ID (never by position), reuses processOneInboxRow_()
 * for the actual filing (no duplicated classification logic), and -- only
 * if that item requested Calendar sync -- syncs only that one newly filed
 * row via the existing syncRow_(), never syncAllToCalendar(). Idempotent:
 * if the row is already marked Processed (e.g. the scheduled trigger beat
 * this call to it), returns the existing filed location instead of firing
 * appendAssignmentFromInbox_/appendAssessmentFromInbox_ a second time.
 *
 * Throws on hard failures (locking, sheet/row not found) -- the calling
 * api_processInboxItem() wrapper in API.gs catches and reports these via
 * the normal {ok:false, error} contract. Returns a plain result object on
 * every other path, including the idempotent "already processed" case.
 */
function processOneInboxItemById_(inboxId) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(LOCK_WAIT_MS)) throw new Error("Workbook is busy — try again in a moment.");
  try {
    var inbox = SpreadsheetApp.getActive().getSheetByName(INBOX_SHEET);
    if (!inbox) throw new Error("Sheet not found: " + INBOX_SHEET);
    var map = getColMap_(inbox);
    var idCol = col_(map, "Inbox ID");
    var lastRow = inbox.getLastRow();
    var ids = inbox.getRange(HEADER_ROW + 1, idCol, Math.max(lastRow - HEADER_ROW, 0), 1).getValues();
    var row = -1;
    for (var i = 0; i < ids.length; i++) { if (ids[i][0] === inboxId) { row = HEADER_ROW + 1 + i; break; } }
    if (row === -1) throw new Error("Inbox item not found for Inbox ID = " + inboxId);

    var processedCell = inbox.getRange(row, col_(map, "Processed"));
    if (processedCell.getValue() === true) {
      // Idempotent path -- something else (the bulk trigger, or an earlier
      // call) already filed this row. Report where, rather than filing it
      // again.
      var existingLocation = inbox.getRange(row, col_(map, "Processed into (sheet!row)")).getValue() || "";
      var parts = String(existingLocation).split("!");
      var existingSheetName = parts[0] || "";
      var existingRow = parseInt(parts[1], 10);
      var filedType = (existingSheetName === ASSIGNMENTS_SHEET) ? "Assignment" :
        (existingSheetName === ASSESSMENTS_SHEET) ? "Assessment" : "Unknown";
      var filedId = "";
      if (existingSheetName && existingRow) {
        try {
          var esheet = SpreadsheetApp.getActive().getSheetByName(existingSheetName);
          var emap = getColMap_(esheet);
          var idHeader = (filedType === "Assignment") ? "Assignment ID" : "Assessment ID";
          filedId = esheet.getRange(existingRow, col_(emap, idHeader)).getValue();
        } catch (e) { /* best-effort only -- idempotent response still returns without it */ }
      }
      // v1.1.12 master -- Feature 3 refinement: calendarStatus is always one
      // of Created/Updated/Skipped/Not requested/Error. Nothing was actually
      // attempted on this call (the item was already filed earlier), so
      // "Skipped" with a detail explaining why is the honest value here --
      // there is no separate "Unchanged" state in the standardized enum.
      return { inboxId: inboxId, alreadyProcessed: true, filedType: filedType, filedId: filedId, calendarStatus: "Skipped", calendarDetail: "Already processed previously; calendar sync (if any) was not re-run." };
    }

    var filed = processOneInboxRow_(inbox, map, row);
    if (!filed) {
      throw new Error("This Inbox item is missing a Title or Module and cannot be filed yet.");
    }

    logAutomation_("Inbox item processed (targeted)", inboxId, "Success", filed.filedType + " " + filed.destId);

    // v1.1.12 master -- Feature 3 refinement: standardize the targeted-sync
    // result to exactly one of Created/Updated/Skipped/Not requested/Error,
    // with any extra detail (the underlying syncRow_ error text, or why sync
    // was skipped) carried separately in calendarDetail rather than folded
    // into the status string itself.
    var calendarStatus = "Not requested";
    var calendarDetail = "";
    if (filed.sendToCalendar) {
      if (!syncEnabled_()) {
        calendarStatus = "Skipped";
        calendarDetail = "Calendar sync is disabled in Settings.";
      } else {
        try {
          var calendar = getTargetCalendar_();
          var destSheet = SpreadsheetApp.getActive().getSheetByName(filed.destSheetName);
          var destMap = getColMap_(destSheet);
          var isAssessment = (filed.destSheetName === ASSESSMENTS_SHEET);
          var rawResult = syncRow_(destSheet, destMap, filed.destRow, calendar, isAssessment);
          if (rawResult === "Created" || rawResult === "Updated") {
            calendarStatus = rawResult;
          } else {
            calendarStatus = "Error";
            calendarDetail = rawResult;
          }
        } catch (calErr) {
          calendarStatus = "Error";
          calendarDetail = calErr && calErr.message ? calErr.message : String(calErr);
        }
      }
    }

    return { inboxId: inboxId, alreadyProcessed: false, filedType: filed.filedType, filedId: filed.destId, calendarStatus: calendarStatus, calendarDetail: calendarDetail };
  } finally {
    lock.releaseLock();
  }
}

function appendAssignmentFromInbox_(module, title, notes, dueDate, dueTime, priority, sendToCal, inboxId) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(ASSIGNMENTS_SHEET);
  var map = getColMap_(sheet);
  var row = sheet.getLastRow() + 1;
  // find first fully-blank row instead of always appending, to keep within the pre-formatted block
  row = firstBlankRow_(sheet, col_(map, "Title"));
  var newId = nextId_("ASG", getColumnValues_(sheet, col_(map, "Assignment ID")));
  sheet.getRange(row, col_(map, "Assignment ID")).setValue(newId);
  sheet.getRange(row, col_(map, "Module")).setValue(module);
  sheet.getRange(row, col_(map, "Title")).setValue(title);
  if (notes) sheet.getRange(row, col_(map, "Notes")).setValue(notes);
  if (dueDate) sheet.getRange(row, col_(map, "Due date")).setValue(dueDate);
  if (dueTime) sheet.getRange(row, col_(map, "Due time")).setValue(dueTime);
  if (priority) sheet.getRange(row, col_(map, "Priority")).setValue(priority);
  sheet.getRange(row, col_(map, "Status")).setValue("Not started");
  if (sendToCal === true) sheet.getRange(row, col_(map, "Sync to Calendar")).setValue(true);
  return row;
}

function appendAssessmentFromInbox_(module, itemType, title, date, startTime, venue, priority, sendToCal, inboxId) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(ASSESSMENTS_SHEET);
  var map = getColMap_(sheet);
  var row = firstBlankRow_(sheet, col_(map, "Title"));
  var newId = nextId_("ASM", getColumnValues_(sheet, col_(map, "Assessment ID")));
  sheet.getRange(row, col_(map, "Assessment ID")).setValue(newId);
  var code = lookupModuleCode_(module);
  if (code) sheet.getRange(row, col_(map, "Module code")).setValue(code);
  sheet.getRange(row, col_(map, "Module")).setValue(module);
  sheet.getRange(row, col_(map, "Assessment type")).setValue(itemType || "Other");
  sheet.getRange(row, col_(map, "Title")).setValue(title);
  if (date) sheet.getRange(row, col_(map, "Date")).setValue(date);
  if (startTime) sheet.getRange(row, col_(map, "Start time")).setValue(startTime);
  if (venue) sheet.getRange(row, col_(map, "Venue")).setValue(venue);
  sheet.getRange(row, col_(map, "Status")).setValue("Scheduled");
  if (sendToCal === true) sheet.getRange(row, col_(map, "Sync to Calendar")).setValue(true);
  return row;
}

function lookupModuleCode_(moduleName) {
  var modules = SpreadsheetApp.getActive().getSheetByName("03 Modules");
  var map = getColMap_(modules);
  var lastRow = modules.getLastRow();
  var names = modules.getRange(HEADER_ROW + 1, col_(map, "Module name"), lastRow - HEADER_ROW, 1).getValues();
  var codes = modules.getRange(HEADER_ROW + 1, col_(map, "Module code"), lastRow - HEADER_ROW, 1).getValues();
  for (var i = 0; i < names.length; i++) {
    if (names[i][0] === moduleName) return codes[i][0];
  }
  return null;
}

function firstBlankRow_(sheet, keyCol) {
  var lastRow = Math.max(sheet.getLastRow(), HEADER_ROW + 1);
  var vals = sheet.getRange(HEADER_ROW + 1, keyCol, lastRow - HEADER_ROW, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (!vals[i][0]) return HEADER_ROW + 1 + i;
  }
  return lastRow + 1; // table is full — append below it
}

function getColumnValues_(sheet, col) {
  var lastRow = sheet.getLastRow();
  return sheet.getRange(HEADER_ROW + 1, col, lastRow - HEADER_ROW, 1).getValues().map(function (r) { return r[0]; });
}

// ============================================================
// CALENDAR SYNC
// ============================================================
function syncAllToCalendar() {
  if (!syncEnabled_()) {
    SpreadsheetApp.getUi().alert('Calendar sync is disabled. Set "Calendar sync enabled" to TRUE in 02 Settings first.');
    return;
  }
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(LOCK_WAIT_MS)) {
    SpreadsheetApp.getActive().toast("Another sync is already running — try again shortly.", "Calendar sync", 6);
    return;
  }
  try {
    var calendar;
    try {
      calendar = getTargetCalendar_();
    } catch (e) {
      SpreadsheetApp.getUi().alert("Calendar sync error", e.message, SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    var a = syncSheet_(ASSIGNMENTS_SHEET, calendar, false);
    var b = syncSheet_(ASSESSMENTS_SHEET, calendar, true);
    var c = syncStudySessions_(calendar);
    SpreadsheetApp.getActive().toast(
      "Assignments: " + a.created + " created, " + a.updated + " updated, " + a.errors + " error(s). " +
      "Assessments: " + b.created + " created, " + b.updated + " updated, " + b.errors + " error(s). " +
      "Study sessions: " + c.created + " created, " + c.updated + " updated, " + c.errors + " error(s).",
      "Calendar sync complete", 8);
  } finally {
    lock.releaseLock();
  }
}

/** Sync only the row the cursor is currently on (in Assignments or Assessments). */
function syncSelectedRow() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var name = sheet.getName();
  if (name !== ASSIGNMENTS_SHEET && name !== ASSESSMENTS_SHEET) {
    SpreadsheetApp.getUi().alert("Select a row in 08 Assignments or 09 Assessments first.");
    return;
  }
  if (!syncEnabled_()) {
    SpreadsheetApp.getUi().alert('Calendar sync is disabled. Set "Calendar sync enabled" to TRUE in 02 Settings first.');
    return;
  }
  var row = sheet.getActiveCell().getRow();
  if (row <= HEADER_ROW) {
    SpreadsheetApp.getUi().alert("Select a data row (not the header).");
    return;
  }
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(LOCK_WAIT_MS)) {
    SpreadsheetApp.getActive().toast("Another sync is already running — try again shortly.", "Calendar sync", 6);
    return;
  }
  try {
    var calendar;
    try { calendar = getTargetCalendar_(); }
    catch (e) { SpreadsheetApp.getUi().alert("Calendar sync error", e.message, SpreadsheetApp.getUi().ButtonSet.OK); return; }
    var isAssessment = (name === ASSESSMENTS_SHEET);
    var result = syncRow_(sheet, getColMap_(sheet), row, calendar, isAssessment);
    SpreadsheetApp.getActive().toast("Row " + row + ": " + result, "Calendar sync", 6);
  } finally {
    lock.releaseLock();
  }
}

function syncSheet_(sheetName, calendar, isAssessment) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  var counts = { created: 0, updated: 0, errors: 0 };
  if (!sheet) return counts;
  var map = getColMap_(sheet);
  var lastRow = sheet.getLastRow();
  for (var row = HEADER_ROW + 1; row <= lastRow; row++) {
    var syncFlag = sheet.getRange(row, col_(map, "Sync to Calendar")).getValue();
    if (syncFlag !== true) continue;
    var result = syncRow_(sheet, map, row, calendar, isAssessment);
    if (result === "Created") counts.created++;
    else if (result === "Updated") counts.updated++;
    else counts.errors++;
  }
  return counts;
}

/** Returns "Created", "Updated", or an "Error: ..." string. Never throws. */
function syncRow_(sheet, map, row, calendar, isAssessment) {
  var statusCell = sheet.getRange(row, col_(map, "Sync status"));
  var errorCellRange = sheet.getRange(row, col_(map, isAssessment ? "Title" : "Title")); // restrained highlight target
  try {
    var moduleName = sheet.getRange(row, col_(map, "Module")).getValue();
    var title = isAssessment
      ? sheet.getRange(row, col_(map, "Assessment type")).getValue() + " - " + sheet.getRange(row, col_(map, "Title")).getValue()
      : sheet.getRange(row, col_(map, "Title")).getValue();
    var dateVal = sheet.getRange(row, col_(map, isAssessment ? "Date" : "Due date")).getValue();

    if (!moduleName || !title || !dateVal) throw new Error("Missing required field (module, title, or date).");

    var startTimeHeader = isAssessment ? "Start time" : "Due time";
    var startTimeVal = sheet.getRange(row, col_(map, startTimeHeader)).getValue();
    var hasTime = startTimeVal instanceof Date && (startTimeVal.getHours() !== 0 || startTimeVal.getMinutes() !== 0);

    var startDateTime = combineDateAndTime_(dateVal, hasTime ? startTimeVal : null);
    var endDateTime;
    if (isAssessment && map["End time"]) {
      var endTimeVal = sheet.getRange(row, col_(map, "End time")).getValue();
      if (endTimeVal instanceof Date) {
        endDateTime = combineDateAndTime_(dateVal, endTimeVal);
        if (endDateTime <= startDateTime) throw new Error("End time is not after start time.");
      } else {
        endDateTime = new Date(startDateTime.getTime() + 60 * 60 * 1000);
      }
    } else {
      endDateTime = new Date(startDateTime.getTime() + 60 * 60 * 1000);
    }

    var descriptionParts = [];
    var dateLabel = formatLocalDateOnly_(dateVal);
    if (dateLabel) descriptionParts.push("Date: " + dateLabel);
    if (isAssessment && map["Venue"]) {
      var venue = sheet.getRange(row, col_(map, "Venue")).getValue();
      if (venue) descriptionParts.push("Venue: " + venue);
    } else if (map["Notes"]) {
      var notes = sheet.getRange(row, col_(map, "Notes")).getValue();
      if (notes) descriptionParts.push("Notes: " + notes);
    }
    if (map["Submission link"]) {
      var link = sheet.getRange(row, col_(map, "Submission link")).getValue();
      if (link) descriptionParts.push("Link: " + link);
    }
    var stableIdHeader = isAssessment ? "Assessment ID" : "Assignment ID";
    var stableId = map[stableIdHeader] ? sheet.getRange(row, col_(map, stableIdHeader)).getValue() : "";
    descriptionParts.push("Source: " + sheet.getName() + " " + stableId + " (row " + row + ")");
    var description = descriptionParts.join("\n");

    var eventTitle = "[" + moduleName + "] " + title;
    var idCol = col_(map, "Calendar Event ID");
    var existingId = sheet.getRange(row, idCol).getValue();
    var event = null;
    var wasRecreated = false;
    if (existingId) {
      try { event = calendar.getEventById(existingId); } catch (e) { event = null; }
      if (!event) wasRecreated = true; // stored ID is stale — will recreate once below
    }

    if (event) {
      event.setTitle(eventTitle);
      event.setDescription(description);
      if (hasTime) event.setTime(startDateTime, endDateTime);
      else event.setAllDayDate(dateVal);
    } else if (hasTime) {
      event = calendar.createEvent(eventTitle, startDateTime, endDateTime, { description: description });
    } else {
      event = calendar.createAllDayEvent(eventTitle, dateVal, { description: description });
    }

    applyReminders_(event, sheet, map, row);

    sheet.getRange(row, idCol).setValue(event.getId());
    var statusText = wasRecreated ? "Recreated (previous event no longer existed)" : (existingId ? "Updated" : "Created");
    statusCell.setValue(statusText + " " + Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd HH:mm"));
    if (map["Last synced"]) sheet.getRange(row, col_(map, "Last synced")).setValue(new Date());
    clearErrorHighlight_(sheet, row, map);
    return wasRecreated ? "Updated" : (existingId ? "Updated" : "Created");

  } catch (err) {
    statusCell.setValue("Error: " + err.message);
    if (map["Last synced"]) sheet.getRange(row, col_(map, "Last synced")).setValue(new Date());
    applyErrorHighlight_(sheet, row, map);
    return "Error: " + err.message;
  }
}

// Restrained highlight: colour only the Sync status cell red, not the whole row.
function applyErrorHighlight_(sheet, row, map) {
  sheet.getRange(row, col_(map, "Sync status")).setBackground("#FFC7CE");
}
function clearErrorHighlight_(sheet, row, map) {
  sheet.getRange(row, col_(map, "Sync status")).setBackground("#C6EFCE");
}

function combineDateAndTime_(dateVal, timeVal) {
  var d = new Date(dateVal);
  if (timeVal) d.setHours(timeVal.getHours(), timeVal.getMinutes(), 0, 0);
  else d.setHours(0, 0, 0, 0);
  return d;
}

// v1.1.12 master -- Feature 8 (clear date labels), backend half. Used only
// when building Calendar event descriptions below -- the event's own
// date/time fields (which Google Calendar renders itself) are untouched;
// this only makes the human-readable "Date: ..." line inside the event
// description clearer than a bare yyyy-MM-dd would be. Reuses the existing
// TIMEZONE constant, same as every other date formatting call in this file.
function formatLocalDateOnly_(dateVal) {
  if (!(dateVal instanceof Date)) return "";
  return Utilities.formatDate(dateVal, TIMEZONE, "EEEE, d MMMM yyyy");
}

// ============================================================
// REMINDERS  (read Settings defaults, allow no row-level override columns to still work)
// ============================================================
function applyReminders_(event, sheet, map, row) {
  var r1 = getSetting_("Default reminder 1 (days before)");
  var r2 = getSetting_("Default reminder 2 (days before)");
  var minutesList = [];
  [r1, r2].forEach(function (days) {
    if (typeof days === "number" && days > 0 && days <= 28) {
      minutesList.push(days * 24 * 60);
    }
  });
  // de-duplicate and cap at 2 reminders to avoid over-notifying
  minutesList = minutesList.filter(function (v, i, a) { return a.indexOf(v) === i; }).slice(0, 2);
  try {
    event.removeAllReminders();
    minutesList.forEach(function (m) { event.addPopupReminder(m); });
  } catch (e) {
    // reminders are best-effort; a failure here should not fail the whole sync
  }
}

// ============================================================
// EXPLICIT DELETE (never automatic)
// ============================================================
// v1.1.12 master -- shared by api_deleteAssignment/api_deleteAssessment/
// api_deleteStudySession in API.gs. Adapted from
// deleteCalendarEventForSelectedRow's core logic just below, but callable
// from the web app (no getUi()/getActiveCell()) and operating on an explicit
// event ID rather than "whatever row the spreadsheet cursor happens to be
// on." Never invoked automatically -- the caller (API.gs) only calls this
// when the user explicitly chose "Workbook + Calendar event" in the delete
// confirmation.
function deleteCalendarEventByStoredId_(calendarEventId) {
  try {
    var calendar = getTargetCalendar_();
    var event = null;
    try { event = calendar.getEventById(calendarEventId); } catch (e) { event = null; }
    if (!event) {
      // Already gone from Calendar (manually deleted, wrong calendar
      // configured, etc.) -- nothing left to delete, which from the
      // caller's point of view is success, not failure.
      return { ok: true, alreadyGone: true };
    }
    event.deleteEvent();
    return { ok: true, alreadyGone: false };
  } catch (e) {
    return { ok: false, error: e && e.message ? e.message : String(e) };
  }
}

function deleteCalendarEventForSelectedRow() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var name = sheet.getName();
  if (name !== ASSIGNMENTS_SHEET && name !== ASSESSMENTS_SHEET) {
    SpreadsheetApp.getUi().alert("Select a row in 08 Assignments or 09 Assessments first.");
    return;
  }
  var row = sheet.getActiveCell().getRow();
  var map = getColMap_(sheet);
  var idCol = col_(map, "Calendar Event ID");
  var eventId = sheet.getRange(row, idCol).getValue();
  if (!eventId) {
    SpreadsheetApp.getUi().alert("This row has no synced Calendar Event ID.");
    return;
  }
  var ui = SpreadsheetApp.getUi();
  var confirm = ui.alert("Delete calendar event for row " + row + "? This cannot be undone.", ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;
  try {
    var calendar = getTargetCalendar_();
    var event = calendar.getEventById(eventId);
    if (event) event.deleteEvent();
    sheet.getRange(row, idCol).setValue("");
    sheet.getRange(row, col_(map, "Sync to Calendar")).setValue(false);
    sheet.getRange(row, col_(map, "Sync status")).setValue("Deleted " + Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd HH:mm"));
  } catch (e) {
    ui.alert("Delete failed: " + e.message);
  }
}

// Study sessions are optional and structurally different (Date/Start/End time rather than
// Due date), so they get their own small sync pass instead of overloading syncRow_.
function syncStudySessions_(calendar) {
  var counts = { created: 0, updated: 0, errors: 0 };
  var sheet = SpreadsheetApp.getActive().getSheetByName(STUDY_SHEET);
  if (!sheet) return counts;
  var map = getColMap_(sheet);
  if (!map["Sync to Calendar"]) return counts; // older workbook version without this column
  var lastRow = sheet.getLastRow();
  for (var row = HEADER_ROW + 1; row <= lastRow; row++) {
    if (sheet.getRange(row, col_(map, "Sync to Calendar")).getValue() !== true) continue;
    var statusCell = sheet.getRange(row, col_(map, "Sync status"));
    try {
      var moduleName = sheet.getRange(row, col_(map, "Module")).getValue();
      var studyType = sheet.getRange(row, col_(map, "Study type")).getValue();
      var date = sheet.getRange(row, col_(map, "Date")).getValue();
      var startTime = sheet.getRange(row, col_(map, "Start time")).getValue();
      var endTime = sheet.getRange(row, col_(map, "End time")).getValue();
      if (!moduleName || !date || !(startTime instanceof Date)) throw new Error("Missing module, date, or start time.");
      var start = combineDateAndTime_(date, startTime);
      var end = (endTime instanceof Date) ? combineDateAndTime_(date, endTime) : new Date(start.getTime() + 60 * 60000);
      if (end <= start) throw new Error("End time is not after start time.");
      var title = "[Study] " + moduleName + (studyType ? " — " + studyType : "");
      var idCol = col_(map, "Calendar Event ID");
      var existingId = sheet.getRange(row, idCol).getValue();
      var event = existingId ? (function () { try { return calendar.getEventById(existingId); } catch (e) { return null; } })() : null;
      var isNew = !event;
      if (event) { event.setTitle(title); event.setTime(start, end); }
      else event = calendar.createEvent(title, start, end, { description: "Date: " + formatLocalDateOnly_(date) + "\nSource: " + STUDY_SHEET + " row " + row });
      sheet.getRange(row, idCol).setValue(event.getId());
      statusCell.setValue((isNew ? "Created " : "Updated ") + Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd HH:mm"));
      statusCell.setBackground("#C6EFCE");
      if (map["Last synced"]) sheet.getRange(row, col_(map, "Last synced")).setValue(new Date());
      isNew ? counts.created++ : counts.updated++;
    } catch (err) {
      statusCell.setValue("Error: " + err.message);
      statusCell.setBackground("#FFC7CE");
      counts.errors++;
    }
  }
  return counts;
}

// ============================================================
// SEMESTER MANAGEMENT  (reusable every semester, without breaking formulas)
// ============================================================
function updateSemesterDates() {
  var ui = SpreadsheetApp.getUi();
  var startResp = ui.prompt("Update Semester Dates", "New semester start date (YYYY-MM-DD):", ui.ButtonSet.OK_CANCEL);
  if (startResp.getSelectedButton() !== ui.Button.OK) return;
  var endResp = ui.prompt("Update Semester Dates", "New semester end date (YYYY-MM-DD):", ui.ButtonSet.OK_CANCEL);
  if (endResp.getSelectedButton() !== ui.Button.OK) return;
  var settings = SpreadsheetApp.getActive().getSheetByName(SETTINGS_SHEET);
  var startCell = settings.createTextFinder("Semester start date").matchEntireCell(false).findNext();
  var endCell = settings.createTextFinder("Semester end date").matchEntireCell(false).findNext();
  if (!startCell || !endCell) { ui.alert("Could not find the semester date rows in 02 Settings."); return; }
  settings.getRange(startCell.getRow(), 3).setValue(new Date(startResp.getResponseText()));
  settings.getRange(endCell.getRow(), 3).setValue(new Date(endResp.getResponseText()));
  ui.alert("Semester dates updated. Note: the Dashboard's phase logic and the 06 Weekly Timetable / 14 Semester Calendar " +
    "date tables use fixed week boundaries built for THIS semester's actual dates — for a genuinely new semester, use " +
    "'Start New Semester' instead so those get rebuilt correctly rather than just re-pointed.");
}

function archiveSemester() {
  var ui = SpreadsheetApp.getUi();
  var confirm = ui.alert("Archive Semester",
    "This copies current Marks Tracker results into 17 Archive, then clears working data in Assignments, Assessments, " +
    "AF Components, Study Planner, Academic Inbox and Attendance (Module Rules, Modules, and Settings are kept). " +
    "This cannot be undone from within the sheet. Continue?", ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;

  var ss = SpreadsheetApp.getActive();
  var mt = ss.getSheetByName("11 Marks Tracker");
  var archive = ss.getSheetByName("17 Archive");
  var semesterName = getSetting_("Semester name") || "Unknown semester";
  var mtMap = getColMap_(mt, MARKS_TRACKER_HEADER_ROW);
  var lastRow = mt.getLastRow();
  var archiveRow = archive.getLastRow() + 1;
  if (archiveRow < 8) archiveRow = 8; // below the archive sheet's own header block

  for (var row = MARKS_TRACKER_HEADER_ROW + 1; row <= lastRow; row++) {
    var moduleName = mt.getRange(row, col_(mtMap, "Module")).getValue();
    if (!moduleName) continue;
    var finalMark = mt.getRange(row, col_(mtMap, "Final / provisional (after A3)")).getValue();
    if (!finalMark) finalMark = mt.getRange(row, col_(mtMap, "Provisional final (after A2)")).getValue();
    var status = mt.getRange(row, col_(mtMap, "Pass status")).getValue();
    archive.getRange(archiveRow, 2).setValue(semesterName);
    archive.getRange(archiveRow, 3).setValue(moduleName);
    archive.getRange(archiveRow, 4).setValue(finalMark);
    archive.getRange(archiveRow, 5).setValue(status);
    archiveRow++;
  }

  var sheetsToClear = [ASSIGNMENTS_SHEET, ASSESSMENTS_SHEET, "10 AF Components", STUDY_SHEET, INBOX_SHEET, "19 Attendance"];
  sheetsToClear.forEach(function (name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) return;
    var lr = sheet.getLastRow();
    var lc = sheet.getLastColumn();
    if (lr > HEADER_ROW) sheet.getRange(HEADER_ROW + 1, 1, lr - HEADER_ROW, lc).clearContent();
  });

  ui.alert("Semester archived. Working sheets are cleared; formulas and structure are untouched.");
}

function startNewSemester() {
  var ui = SpreadsheetApp.getUi();
  var confirm = ui.alert("Start New Semester",
    "This runs Archive Semester first, then prompts you to update the semester name and dates. " +
    "For a genuinely different teaching-week/recess/exam calendar, ask for a rebuilt workbook instead — " +
    "this only clears data and updates dates/name, it does not regenerate 06 Weekly Timetable or 14 Semester Calendar's " +
    "week boundaries. Continue?", ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;
  archiveSemester();
  var nameResp = ui.prompt("New semester name (e.g. 'Semester 1, 2027'):", ui.ButtonSet.OK_CANCEL);
  if (nameResp.getSelectedButton() === ui.Button.OK) {
    var settings = SpreadsheetApp.getActive().getSheetByName(SETTINGS_SHEET);
    var nameCell = settings.createTextFinder("Semester name").matchEntireCell(false).findNext();
    if (nameCell) settings.getRange(nameCell.getRow(), 3).setValue(nameResp.getResponseText());
  }
  updateSemesterDates();
}

function duplicateSemester() {
  var ui = SpreadsheetApp.getUi();
  var nameResp = ui.prompt("Duplicate Semester", "Name for the new copy:", ui.ButtonSet.OK_CANCEL);
  if (nameResp.getSelectedButton() !== ui.Button.OK) return;
  var file = DriveApp.getFileById(SpreadsheetApp.getActive().getId());
  var copy = file.makeCopy(nameResp.getResponseText());
  ui.alert("Duplicated", "Created: " + copy.getUrl() + "\n\nAutomation triggers are NOT copied — run " +
    "Academic Planner > Enable Automatic Sync inside the new copy if you want it there too.", ui.ButtonSet.OK);
}


// ============================================================
// STUDY SUGGESTIONS  (controlled generation — never auto-overwrites manual sessions)
// ============================================================
/**
 * v1.2.0 -- extracted, UNCHANGED in behaviour, from what used to be the top
 * half of generateStudySuggestions()'s try block. Computes which
 * date+hour slots in the coming Monday–Friday are already occupied by a
 * real timetable class, so both the legacy menu action below and the new
 * web-app preview flow (buildStudySuggestions_, further down) compute
 * "busy" the exact same way -- one source of truth, not two.
 */
function computeBusyTimetableSlots_(mondayDate) {
  var tt = SpreadsheetApp.getActive().getSheetByName(TIMETABLE_SHEET);
  var busy = {};
  if (!tt) return busy;
  var ttMap = getColMap_(tt);
  var ttLastRow = tt.getLastRow();
  if (ttLastRow <= HEADER_ROW) return busy;
  var startDates = tt.getRange(HEADER_ROW + 1, col_(ttMap, "Start date"), ttLastRow - HEADER_ROW, 1).getValues();
  var startTimes = tt.getRange(HEADER_ROW + 1, col_(ttMap, "Start time"), ttLastRow - HEADER_ROW, 1).getValues();
  var endTimes = tt.getRange(HEADER_ROW + 1, col_(ttMap, "End time"), ttLastRow - HEADER_ROW, 1).getValues();
  for (var i = 0; i < startDates.length; i++) {
    var d = startDates[i][0];
    if (!(d instanceof Date)) continue;
    var sH = startTimes[i][0] instanceof Date ? startTimes[i][0].getHours() : null;
    var eH = endTimes[i][0] instanceof Date ? endTimes[i][0].getHours() : null;
    if (sH === null || eH === null) continue;
    for (var h = sH; h < eH; h++) {
      busy[Utilities.formatDate(d, TIMEZONE, "yyyy-MM-dd") + "|" + h] = true;
    }
  }
  return busy;
}

/** Same fixed two-slot-per-weekday suggestion shape the legacy menu action
 *  has always used (17:00 Tutorial preparation, 19:00 Practice questions) —
 *  kept as its own small function purely so it's named and documented once
 *  instead of appearing as an anonymous inline array in two places. */
function defaultSuggestionSlots_() {
  return [
    { hour: 17, duration: 90, type: "Tutorial preparation" },
    { hour: 19, duration: 60, type: "Practice questions" }
  ];
}

function generateStudySuggestions() {
  var ss = SpreadsheetApp.getActive();
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(LOCK_WAIT_MS)) {
    SpreadsheetApp.getActive().toast("Busy — try again shortly.", "Study suggestions", 5);
    return;
  }
  try {
    var today = new Date();
    var monday = new Date(today);
    monday.setDate(today.getDate() + ((1 - today.getDay() + 7) % 7));
    var busy = computeBusyTimetableSlots_(monday);

    var study = ss.getSheetByName(STUDY_SHEET);
    var stMap = getColMap_(study);
    var lastRow = Math.max(study.getLastRow(), HEADER_ROW + 1);
    var existingDates = study.getRange(HEADER_ROW + 1, col_(stMap, "Date"), lastRow - HEADER_ROW, 1).getValues();
    var existingStart = study.getRange(HEADER_ROW + 1, col_(stMap, "Start time"), lastRow - HEADER_ROW, 1).getValues();

    // find next Monday..Friday with no existing session (manual or generated) in a given slot
    var suggestions = defaultSuggestionSlots_();
    var created = 0;
    var writeRow = firstBlankRow_(study, col_(stMap, "Module"));

    for (var day = 0; day < 5 && created < 5; day++) {
      var date = new Date(monday); date.setDate(monday.getDate() + day);
      var dateStr = Utilities.formatDate(date, TIMEZONE, "yyyy-MM-dd");
      for (var s = 0; s < suggestions.length && created < 5; s++) {
        var slot = suggestions[s];
        if (busy[dateStr + "|" + slot.hour]) continue; // avoid class overlap

        var alreadyExists = false;
        for (var k = 0; k < existingDates.length; k++) {
          var ed = existingDates[k][0];
          if (ed instanceof Date && Utilities.formatDate(ed, TIMEZONE, "yyyy-MM-dd") === dateStr) {
            var est = existingStart[k][0];
            if (est instanceof Date && est.getHours() === slot.hour) { alreadyExists = true; break; }
          }
        }
        if (alreadyExists) continue; // never overwrite manual (or previously generated) sessions

        var startT = new Date(date); startT.setHours(slot.hour, 0, 0, 0);
        var endT = new Date(startT.getTime() + slot.duration * 60000);
        study.getRange(writeRow, col_(stMap, "Date")).setValue(date);
        study.getRange(writeRow, col_(stMap, "Start time")).setValue(startT);
        study.getRange(writeRow, col_(stMap, "End time")).setValue(endT);
        study.getRange(writeRow, col_(stMap, "Study type")).setValue(slot.type);
        study.getRange(writeRow, col_(stMap, "Priority")).setValue("Medium");
        study.getRange(writeRow, col_(stMap, "Planned")).setValue(true);
        study.getRange(writeRow, col_(stMap, "Completed")).setValue(false);
        study.getRange(writeRow, col_(stMap, "Notes")).setValue("Auto-generated suggestion — edit or delete freely.");
        if (stMap["Origin"]) study.getRange(writeRow, col_(stMap, "Origin")).setValue("Generated");
        writeRow = firstBlankRow_(study, col_(stMap, "Module"));
        created++;
      }
    }
    assignStableIds_(STUDY_SHEET, "Study Session ID", "Study type", "SS");
    SpreadsheetApp.getActive().toast(created + " study suggestion(s) added for next week (no class overlaps, nothing overwritten).", "Study Planner", 6);
  } finally {
    lock.releaseLock();
  }
}

// ============================================================================================
// v1.2.0 ACADEMIC WORKFLOW RELEASE — everything below this line is new.
// Nothing above this line was removed, renamed, or behaviourally changed.
// ============================================================================================

// ------------------------------------------------------------------
// SHARED ROW-LOOKUP HELPER (v1.2.0)
// ------------------------------------------------------------------
/**
 * Generic "find this stable ID's row number on this sheet" lookup, used by
 * every new v1.2.0 mutation (Study Tasks, Resources, Attendance) instead of
 * re-writing the same linear-scan loop each api_delete_/api_update_ function
 * in the v1.1.x baseline already inlines individually. Returns -1 if not
 * found. Never used for anything security-sensitive beyond "which row" —
 * every caller still validates the row exists before acting on it.
 */
function findRowByStableId_(sheetName, idHeader, id) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sheet) return -1;
  var map = getColMap_(sheet);
  if (!map[idHeader]) return -1;
  var idCol = col_(map, idHeader);
  var lastRow = sheet.getLastRow();
  if (lastRow <= HEADER_ROW) return -1;
  var ids = sheet.getRange(HEADER_ROW + 1, idCol, lastRow - HEADER_ROW, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return HEADER_ROW + 1 + i;
  }
  return -1;
}

// ------------------------------------------------------------------
// STUDY SESSION ROW WRITER (v1.2.0) — shared by api_addStudySession,
// api_logIndependentStudy, and api_acceptStudySuggestions in API.gs, so
// "how a Study Planner row gets written" exists in exactly one place, the
// same way appendAssignmentFromInbox_/appendAssessmentFromInbox_ already do
// for Assignments/Assessments above.
// ------------------------------------------------------------------
function appendStudySessionRow_(fields) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(STUDY_SHEET);
  var map = getColMap_(sheet);
  var row = firstBlankRow_(sheet, col_(map, "Module"));
  var newId = nextId_("SS", getColumnValues_(sheet, col_(map, "Study Session ID")));
  sheet.getRange(row, col_(map, "Study Session ID")).setValue(newId);
  sheet.getRange(row, col_(map, "Module")).setValue(fields.module);
  sheet.getRange(row, col_(map, "Date")).setValue(fields.date);
  if (fields.startTime) sheet.getRange(row, col_(map, "Start time")).setValue(fields.startTime);
  if (fields.endTime) sheet.getRange(row, col_(map, "End time")).setValue(fields.endTime);
  if (fields.duration != null && map["Duration (min)"]) sheet.getRange(row, col_(map, "Duration (min)")).setValue(fields.duration);
  sheet.getRange(row, col_(map, "Study type")).setValue(fields.studyType || "Practice");
  if (fields.topic) sheet.getRange(row, col_(map, "Topic")).setValue(fields.topic);
  sheet.getRange(row, col_(map, "Priority")).setValue(fields.priority || "Medium");
  if (map["Objective"] && fields.objective) sheet.getRange(row, col_(map, "Objective")).setValue(fields.objective);
  sheet.getRange(row, col_(map, "Planned")).setValue(fields.planned !== false);
  sheet.getRange(row, col_(map, "Completed")).setValue(!!fields.completed);
  if (map["Status"]) sheet.getRange(row, col_(map, "Status")).setValue(fields.status || (fields.completed ? "Completed" : "Planned"));
  if (fields.notes) sheet.getRange(row, col_(map, "Notes")).setValue(fields.notes);
  if (map["Origin"]) sheet.getRange(row, col_(map, "Origin")).setValue(fields.origin || "Manual");
  if (fields.syncToCalendar && map["Sync to Calendar"]) sheet.getRange(row, col_(map, "Sync to Calendar")).setValue(true);
  return { row: row, id: newId };
}

// ------------------------------------------------------------------
// WEEKLY STUDY SUGGESTIONS — PREVIEW/ACCEPT FLOW (v1.2.0, Feature: Study
// Planner §4.6). Deliberately separate from generateStudySuggestions()
// above (which stays exactly as it was — a menu action that writes
// immediately): the web app instead computes candidates WITHOUT writing
// anything, lets the user Accept/Edit/Reject each one client-side, and only
// api_acceptStudySuggestions (in API.gs) actually writes the ones the user
// picked, via appendStudySessionRow_ above. Both paths share
// computeBusyTimetableSlots_/defaultSuggestionSlots_, so "which slots are
// free" is never computed two different ways.
// ------------------------------------------------------------------
function buildStudySuggestions_(weekAnchorDate) {
  var dow = (weekAnchorDate.getDay() + 6) % 7;
  var monday = new Date(weekAnchorDate.getFullYear(), weekAnchorDate.getMonth(), weekAnchorDate.getDate() - dow);
  var busy = computeBusyTimetableSlots_(monday);

  var study = SpreadsheetApp.getActive().getSheetByName(STUDY_SHEET);
  var stMap = getColMap_(study);
  var lastRow = Math.max(study.getLastRow(), HEADER_ROW + 1);
  var existingDates = lastRow > HEADER_ROW ? study.getRange(HEADER_ROW + 1, col_(stMap, "Date"), lastRow - HEADER_ROW, 1).getValues() : [];
  var existingStart = lastRow > HEADER_ROW ? study.getRange(HEADER_ROW + 1, col_(stMap, "Start time"), lastRow - HEADER_ROW, 1).getValues() : [];

  var suggestions = defaultSuggestionSlots_();
  var out = [];
  for (var day = 0; day < 5; day++) {
    var date = new Date(monday); date.setDate(monday.getDate() + day);
    var dateStr = Utilities.formatDate(date, TIMEZONE, "yyyy-MM-dd");
    for (var s = 0; s < suggestions.length; s++) {
      var slot = suggestions[s];
      if (busy[dateStr + "|" + slot.hour]) continue;
      var alreadyExists = false;
      for (var k = 0; k < existingDates.length; k++) {
        var ed = existingDates[k][0];
        if (ed instanceof Date && Utilities.formatDate(ed, TIMEZONE, "yyyy-MM-dd") === dateStr) {
          var est = existingStart[k][0];
          if (est instanceof Date && est.getHours() === slot.hour) { alreadyExists = true; break; }
        }
      }
      if (alreadyExists) continue;
      out.push({
        date: dateStr, hour: slot.hour, duration: slot.duration, studyType: slot.type,
        startTimeIso: Utilities.formatString("%02d:00", slot.hour)
      });
    }
  }
  return { mondayIso: Utilities.formatDate(monday, TIMEZONE, "yyyy-MM-dd"), suggestions: out };
}

// ============================================================================================
// v1.3.0 -- IMPROVED WEEKLY STUDY SUGGESTIONS (spec §7). Additive only:
// buildStudySuggestions_ above (and the menu-driven generateStudySuggestions()
// higher in this file) are byte-for-byte unchanged and still fully
// functional -- this is a second, richer builder used only by the web app's
// own preview/accept flow (api_generateStudySuggestions in API_v1.3.0.gs),
// which was already explicitly designed for exactly this kind of iteration
// (see the v1.2.0 architecture notes: "the web app instead computes
// candidates WITHOUT writing anything"). Same busy-slot/no-duplicate
// mechanics as buildStudySuggestions_ (never overlaps a real timetable class,
// never re-suggests an existing session) -- the only thing that changed is
// WHICH module/type/topic gets suggested for each free slot.
// ============================================================================================
function buildPrioritizedStudySuggestions_(weekAnchorDate, priorities, revisionTracker, resources, studyTasks) {
  var dow = (weekAnchorDate.getDay() + 6) % 7;
  var monday = new Date(weekAnchorDate.getFullYear(), weekAnchorDate.getMonth(), weekAnchorDate.getDate() - dow);
  var busy = computeBusyTimetableSlots_(monday);

  var study = SpreadsheetApp.getActive().getSheetByName(STUDY_SHEET);
  var stMap = getColMap_(study);
  var lastRow = Math.max(study.getLastRow(), HEADER_ROW + 1);
  var existingDates = lastRow > HEADER_ROW ? study.getRange(HEADER_ROW + 1, col_(stMap, "Date"), lastRow - HEADER_ROW, 1).getValues() : [];
  var existingStart = lastRow > HEADER_ROW ? study.getRange(HEADER_ROW + 1, col_(stMap, "Start time"), lastRow - HEADER_ROW, 1).getValues() : [];

  var slots = defaultSuggestionSlots_(); // same fixed two daily slots as the legacy generator -- only WHAT fills them changes
  var queue = buildSuggestionCandidateQueue_(priorities, revisionTracker, resources, studyTasks);

  var out = [];
  var qi = 0;
  for (var day = 0; day < 5 && qi < queue.length; day++) {
    var date = new Date(monday); date.setDate(monday.getDate() + day);
    var dateStr = Utilities.formatDate(date, TIMEZONE, "yyyy-MM-dd");
    for (var s = 0; s < slots.length && qi < queue.length; s++) {
      var slot = slots[s];
      if (busy[dateStr + "|" + slot.hour]) continue;
      var alreadyExists = false;
      for (var k = 0; k < existingDates.length; k++) {
        var ed = existingDates[k][0];
        if (ed instanceof Date && Utilities.formatDate(ed, TIMEZONE, "yyyy-MM-dd") === dateStr) {
          var est = existingStart[k][0];
          if (est instanceof Date && est.getHours() === slot.hour) { alreadyExists = true; break; }
        }
      }
      if (alreadyExists) continue;
      var cand = queue[qi]; qi++;
      out.push({
        date: dateStr, hour: slot.hour, duration: slot.duration, studyType: cand.studyType,
        module: cand.module, topic: cand.topic, reason: cand.reason,
        startTimeIso: Utilities.formatString("%02d:00", slot.hour)
      });
    }
  }
  // Distinguish WHY nothing came back -- "no eligible module" (every module
  // is Rules Incomplete / Insufficient data, so the candidate queue itself
  // was empty) vs. "eligible modules exist but this week's slots are all
  // busy or already covered" -- rather than one generic empty state. This
  // is exactly the gap that made "Generate this week's plan" look silently
  // broken when every module still had unverified rules.
  var emptyReason = null;
  if (!out.length) {
    emptyReason = queue.length
      ? "No free slots found for that week — fully booked by classes, or already covered by existing sessions."
      : "No eligible module for suggestions yet — every module's Rule status is still Missing/Partially verified, or there isn't enough AF/A1/A2/A3 data yet. Verify a module's rules in \"04 Module Rules\" to start getting suggestions for it.";
  }
  return { mondayIso: Utilities.formatDate(monday, TIMEZONE, "yyyy-MM-dd"), suggestions: out, emptyReason: emptyReason };
}

/**
 * Pure (no SpreadsheetApp calls -- unit-tested directly in Node, see
 * V1.3.0_STATIC_VALIDATION.md). Ranks modules by priority category/score,
 * then proposes one concrete candidate per module per pass of a fixed
 * weak-topic -> formula -> practice -> incomplete-task cycle, so higher
 * priority modules get more of the week's free slots without any single
 * module hogging every one, and every proposal names a real Study type
 * (never a vague "Study Module" entry -- spec §7).
 */
function buildSuggestionCandidateQueue_(priorities, revisionTracker, resources, studyTasks) {
  var eligible = (priorities || []).filter(function (p) { return p.category !== "Rules Incomplete" && p.category !== "Insufficient data"; });
  var order = { "Critical": 6, "Very High": 5, "High": 4, "Maintain": 3, "Low": 2 };
  var sorted = eligible.slice().sort(function (a, b) { return (order[b.category] || 0) - (order[a.category] || 0) || (b.score || 0) - (a.score || 0); });
  if (!sorted.length) return [];

  var weakTopicsByModule = {};
  (revisionTracker || []).forEach(function (r) { if (r.weakTopic) { (weakTopicsByModule[r.module] = weakTopicsByModule[r.module] || []).push(r.topic); } });
  var hasFormulaByModule = {}, hasPastPaperByModule = {};
  (resources || []).forEach(function (r) {
    if (r.type === "Formula bible") hasFormulaByModule[r.module] = true;
    if (r.type === "Past paper" || r.type === "Practice set") hasPastPaperByModule[r.module] = true;
  });
  var incompleteByModule = {};
  (studyTasks || []).forEach(function (t) { if (!t.completed) { (incompleteByModule[t.module] = incompleteByModule[t.module] || []).push(t); } });

  var cyclePattern = ["weak", "formula", "practice", "task"];
  var queue = [];
  var maxPerModule = 3, maxTotal = 8;
  sorted.forEach(function (m) {
    var given = 0;
    for (var i = 0; i < cyclePattern.length && given < maxPerModule && queue.length < maxTotal; i++) {
      var kind = cyclePattern[i];
      if (kind === "weak" && (weakTopicsByModule[m.name] || []).length) {
        queue.push({ module: m.name, studyType: "Practice", topic: weakTopicsByModule[m.name][0], reason: "Weak-topic practice (Revision Tracker)." });
        given++;
      } else if (kind === "formula" && hasFormulaByModule[m.name]) {
        queue.push({ module: m.name, studyType: "Formula", topic: "", reason: "Formula recall -- a Formula bible resource exists for this module." });
        given++;
      } else if (kind === "practice" && hasPastPaperByModule[m.name]) {
        queue.push({ module: m.name, studyType: "Exam", topic: "", reason: "Timed past-paper/practice-set practice -- a matching resource exists." });
        given++;
      } else if (kind === "task" && (incompleteByModule[m.name] || []).length) {
        queue.push({ module: m.name, studyType: "Practice", topic: incompleteByModule[m.name][0].topic || "", reason: "Incomplete Study Task(s) outstanding for this module." });
        given++;
      }
    }
    if (given === 0 && queue.length < maxTotal) {
      // Every eligible module still gets at least one concrete suggestion,
      // even with no weak topics/resources/incomplete tasks recorded yet.
      queue.push({ module: m.name, studyType: "Learn", topic: "", reason: "Priority \"" + m.category + "\" -- no weak topics or outstanding tasks recorded yet." });
    }
  });
  return queue;
}
