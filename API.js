/**
 * API.gs — Engineering Planner OS, Version 1.0
 * Web-app + data-bridge layer. Add this file alongside Calendar_Sync_v2.0.gs in the SAME
 * Apps Script project (bound to the Engineering Planner workbook). Do not rename the
 * global constants already declared in Calendar_Sync_v2.0.gs (TIMEZONE, HEADER_ROW,
 * LOCK_WAIT_MS, SETTINGS_SHEET, ASSIGNMENTS_SHEET, ASSESSMENTS_SHEET, INBOX_SHEET,
 * STUDY_SHEET, TIMETABLE_SHEET) — this file reuses them.
 *
 * Responsibility split:
 *  - Calendar_Sync_v2.0.gs: menu, triggers, calendar sync, inbox processing, semester tools.
 *  - API.gs (this file): serves the web-app UI (doGet) and exposes every read/write the
 *    UI needs as plain functions callable from the browser via google.script.run.
 *
 * Every function in this file is defensive: it never throws a raw error back to the
 * client silently — all api_* functions return { ok: true, data } or { ok: false, error }
 * so the front end can show a real message instead of a blank screen.
 *
 * v1.2.0 note: everything above and below this note, up to the "v1.2.0 ACADEMIC
 * WORKFLOW RELEASE" section near the end, is the same v1.1.12 code, with a small
 * number of narrowly-scoped, in-place edits called out inline with a "v1.2.0 --"
 * comment where a function's INTERNAL activity-log text now uses a human label
 * instead of a bare stable ID (Feature: human-readable activity history). No
 * function was renamed, no signature changed, no {ok,data}/{ok:false,error}
 * contract changed, and no other behaviour changed in any of those edits.
 */

// ------------------------------------------------------------------
// SHEET NAME CONSTANTS not already declared in Calendar_Sync_v2.0.gs
// ------------------------------------------------------------------
var MODULES_SHEET = "03 Modules";
var MODULE_RULES_SHEET = "04 Module Rules";
var AF_COMPONENTS_SHEET = "10 AF Components";
var MARKS_TRACKER_SHEET = "11 Marks Tracker";
var REVISION_SHEET = "13 Revision Tracker";
var RESOURCES_SHEET = "15 Resources";
var NOTES_SHEET = "16 Notes";
var ARCHIVE_SHEET = "17 Archive";
var ATTENDANCE_SHEET = "19 Attendance";
var AUTOMATION_LOG_SHEET = "20 Automation Log";
// v1.3.0 -- Exam Focus Mode's only persisted state (see
// Academic_Intelligence_v1.3.0_Migration.gs). Declared here alongside every
// other sheet-name constant in this file, same convention as STUDY_TASKS_SHEET
// in Calendar_Sync_v1.3.0.gs.
var EXAM_FOCUS_SHEET = "22 Exam Focus Log";
var MARKS_TRACKER_HEADER_ROW = 6; // this sheet has extra title/note rows above its header

// Design-system metadata for modules — colours/short-ids are the APPROVED VISUAL DESIGN
// (from the signed-off HTML prototype). These stay as the DEFAULT/fallback for any module
// that hasn't set its own "Colour (hex)" in 03 Modules yet (v1.2.0 Feature: module
// configuration) — a blank sheet cell always falls back to this table, so nothing changes
// for the existing six modules unless someone deliberately fills in a colour override.
var MODULE_UI_META = {
  "20753-154": { id: "appmaths", short: "Applied Mathematics B", color: "#A78BC9" },
  "38571-145": { id: "engmaths", short: "Engineering Mathematics", color: "#2E8F6E" },
  "12599-143": { id: "electro", short: "Electro-Techniques", color: "#5B2C3D" },
  "30317-143": { id: "compprog", short: "Computer Programming", color: "#C23E7D" },
  "19712-143": { id: "som", short: "Strength of Materials", color: "#4A9FDE" },
  "N/A": { id: "indeng", short: "Industrial Engineering", color: "#8B8B87" }
};

// ============================================================
// RECOVERY RC2 — MANUAL, READ-ONLY DIAGNOSTIC
// Not called by the frontend. Not invoked by any trigger. Not invoked
// automatically. Run it yourself from the Apps Script editor: select
// runRecoveryDiagnostics_ from the function dropdown, click Run, then
// check the Executions log (or the Logger/console output) for the
// JSON report. Read-only: makes zero writes to the workbook, zero
// Calendar API calls, and does not change api_getPlannerData()'s (or
// any other api_* function's) return shape in any way — it only calls
// api_getPlannerData() once, internally, purely to measure it.
// ============================================================
function runRecoveryDiagnostics_() {
  var report = {
    timestamp: new Date().toISOString(),
    spreadsheetAccess: false,
    requiredSheetsPresent: {},
    calendarSettingsLabelPresent: false,
    getPlannerDataDurationMs: null,
    getPlannerDataOk: null,
    getPlannerDataSerializable: false,
    payloadBytes: null,
    errors: []
  };

  var ss;
  try {
    ss = SpreadsheetApp.getActive();
    report.spreadsheetAccess = !!ss;
  } catch (e) {
    report.errors.push('Spreadsheet access failed: ' + (e && e.message ? e.message : String(e)));
    console.log(JSON.stringify(report));
    return report;
  }

  var requiredSheets = [
    SETTINGS_SHEET, MODULES_SHEET, MODULE_RULES_SHEET, MARKS_TRACKER_SHEET, AF_COMPONENTS_SHEET,
    INBOX_SHEET, ASSIGNMENTS_SHEET, ASSESSMENTS_SHEET, STUDY_SHEET, REVISION_SHEET,
    ATTENDANCE_SHEET, RESOURCES_SHEET, NOTES_SHEET, ARCHIVE_SHEET, AUTOMATION_LOG_SHEET, TIMETABLE_SHEET,
    STUDY_TASKS_SHEET // v1.2.0 -- optional until Study_Tasks_v1.2.0_Migration.gs is run; reported, not required
  ];
  requiredSheets.forEach(function (name) {
    try {
      report.requiredSheetsPresent[name] = !!ss.getSheetByName(name);
    } catch (e) {
      report.requiredSheetsPresent[name] = false;
    }
  });

  try {
    var settingsRows = readSheetRows_(SETTINGS_SHEET, HEADER_ROW, "Setting");
    report.calendarSettingsLabelPresent = settingsRows.some(function (r) {
      return r["Setting"] === "Planner Calendar ID (read/write)";
    });
  } catch (e) {
    report.errors.push('Settings label check failed: ' + (e && e.message ? e.message : String(e)));
  }

  try {
    var t0 = Date.now();
    var result = api_getPlannerData(); // read-only call, purely to measure — result is discarded, not returned to any caller
    report.getPlannerDataDurationMs = Date.now() - t0;
    report.getPlannerDataOk = !!(result && result.ok);
    if (result && result.ok === false) {
      report.errors.push('api_getPlannerData() returned ok:false — ' + (result.error || 'no message given'));
    }
    try {
      var serialized = JSON.stringify(result);
      report.getPlannerDataSerializable = true;
      report.payloadBytes = serialized.length;
    } catch (serErr) {
      report.getPlannerDataSerializable = false;
      report.errors.push('JSON.stringify(result) failed: ' + (serErr && serErr.message ? serErr.message : String(serErr)));
    }
  } catch (e) {
    report.errors.push('api_getPlannerData() threw: ' + (e && e.message ? e.message : String(e)));
  }

  console.log(JSON.stringify(report));
  return report;
}

// ============================================================
// WEB APP ENTRY POINT
// ============================================================
// v1.2.0 RC2 -- REMOVAL: the optional Apple Reminders HTTP bridge (doGet's
// ?action=pending-reminders branch, doPost's ?action=confirm-reminder
// branch, and every helper function that only existed to support them) has
// been removed. This is the deliberate removal of an unused, optional,
// legacy integration per the RC2 correction brief -- it is NOT a change to
// core planner behavior. Apple Reminders will be handled externally via
// Apple Shortcuts + Calendar going forward, reading directly from Google
// Calendar events this planner already creates. Ordinary planner reminders
// and Calendar event reminders (both handled entirely by Calendar_Sync.gs
// and Google Calendar's own native reminder settings) are completely
// unaffected -- nothing about those was ever part of this bridge. doGet now
// simply serves Index for every request, exactly as it did before v1.1.12
// master introduced the bridge. doPost has been removed outright: it had no
// other production purpose.
function doGet(e) {
  // v1.4.0 -- Reminders feed (see Reminders_Feed_v1.4.0.js): a request
  // carrying feed/markSynced query params is a Shortcut/automation call,
  // never a browser loading the app -- handled entirely separately, always
  // returns JSON, and never falls through to serving the app page below.
  var params = (e && e.parameter) || {};
  if (params.feed === "assignments" || params.markSynced) {
    return remindersFeedRequest_(params);
  }
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("Engineering Planner OS")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================
// SMALL SHARED HELPERS
// ============================================================
function ok_(data) { return { ok: true, data: data }; }
function fail_(err) {
  var msg = (err && err.message) ? err.message : String(err);
  logAutomation_("API error", "-", "Error", msg);
  return { ok: false, error: msg };
}
function isoDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TIMEZONE, "yyyy-MM-dd");
  return v || "";
}
function isoTime_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TIMEZONE, "HH:mm");
  return v || "";
}
function num_(v) { return (v === "" || v === null || typeof v === "undefined") ? null : Number(v); }

// v1.2.0 -- small shared helper: combines a date-only value with an "HH:MM"
// string into one Date, the same way the client already builds date+time
// pairs. Used only by the new Study Planner mutation endpoints below; does
// not replace or alter combineDateAndTime_() in Calendar_Sync_v1.2.0.gs
// (that one combines a Date value with a Date *time* value, not a string --
// kept separate rather than overloading one function for two input shapes).
function parseTimeOfDay_(hhmm, dateVal) {
  if (!hhmm || !dateVal) return null;
  var parts = String(hhmm).split(":");
  if (parts.length < 2) return null;
  var h = Number(parts[0]), m = Number(parts[1]);
  if (!isFinite(h) || !isFinite(m)) return null;
  var d = new Date(dateVal);
  d.setHours(h, m, 0, 0);
  return d;
}

// ------------------------------------------------------------------
// VALIDATION & SANITIZATION — every mutation in this file runs input through
// these before it touches a sheet. The client validates too (for instant
// feedback), but the client can never be trusted: a malformed or malicious
// request can always be sent directly to the Apps Script endpoint, bypassing
// the UI entirely, so the server-side check here is the one that actually
// protects the workbook.
// ------------------------------------------------------------------
var MAX_TEXT_LEN = 500;
/** Single-line fields (titles, names, venues): collapse newlines, trim, cap length. */
function cleanText_(v, maxLen) {
  if (v === null || typeof v === "undefined") return "";
  var s = String(v).replace(/[\r\n]+/g, " ").trim();
  var cap = maxLen || MAX_TEXT_LEN;
  return s.length > cap ? s.slice(0, cap) : s;
}
/** Multi-line fields (Notes, free text): keep newlines, strip other control
 *  characters, cap length. Google Sheets cell values are always plain text —
 *  there is no HTML/script injection risk here, but a very long paste
 *  shouldn't be able to blow out a cell or the daily payload size. */
function cleanMultiline_(v, maxLen) {
  if (v === null || typeof v === "undefined") return "";
  // eslint-disable-next-line no-control-regex
  var s = String(v).replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, "");
  var cap = maxLen || 4000;
  return s.length > cap ? s.slice(0, cap) : s;
}
function requireText_(v, fieldName, maxLen) {
  var s = cleanText_(v, maxLen);
  if (!s) throw new Error(fieldName + " is required.");
  return s;
}
function requireModule_(moduleName) {
  var s = requireText_(moduleName, "Module");
  var modules = readSheetRows_(MODULES_SHEET, HEADER_ROW, "Module name");
  var known = modules.some(function (m) { return m["Module name"] === s; });
  if (!known) throw new Error('Unknown module: "' + s + '". Refusing to write a row against a module that is not in 03 Modules.');
  return s;
}
function requireFiniteNumber_(v, fieldName, min, max) {
  // Number("") is 0 in JS, not NaN — without this explicit check, a blank
  // field sent straight to the Apps Script endpoint (bypassing the client's
  // own blank check) would silently be saved as a real 0 instead of being
  // rejected. This is the server-side half of "intentional zero is
  // preserved, blank stays blank": blank must fail here, not become 0.
  if (v === "" || v === null || typeof v === "undefined") throw new Error(fieldName + " must be a number.");
  var n = Number(v);
  if (!isFinite(n)) throw new Error(fieldName + " must be a number.");
  if (typeof min === "number" && n < min) throw new Error(fieldName + " must be at least " + min + ".");
  if (typeof max === "number" && n > max) throw new Error(fieldName + " must be at most " + max + ".");
  return n;
}
function parseValidDate_(v, fieldName) {
  if (!v) return null;
  var d = new Date(v);
  if (isNaN(d.getTime())) throw new Error(fieldName + " is not a valid date.");
  return d;
}

/** Reads an entire data sheet into an array of plain objects keyed by header name.
 *  Blank key-column rows are skipped (they are unused pre-formatted rows). */
function readSheetRows_(sheetName, headerRow, keyHeader) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow <= headerRow) return [];
  var headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  var values = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, lastCol).getValues();
  var keyIdx = keyHeader ? headers.indexOf(keyHeader) : -1;
  var out = [];
  for (var r = 0; r < values.length; r++) {
    if (keyIdx >= 0 && !values[r][keyIdx]) continue; // unused pre-formatted row
    var obj = { __row: headerRow + 1 + r };
    for (var c = 0; c < headers.length; c++) {
      var h = (headers[c] || "").toString().trim();
      if (!h) continue;
      var v = values[r][c];
      obj[h] = (v instanceof Date) ? v : v;
    }
    out.push(obj);
  }
  return out;
}

function logAutomation_(action, item, result, detail) {
  try {
    var sheet = SpreadsheetApp.getActive().getSheetByName(AUTOMATION_LOG_SHEET);
    if (!sheet) return;
    var row = sheet.getLastRow() + 1;
    if (row < 5) row = 5;
    sheet.getRange(row, 1, 1, 5).setValues([[new Date(), action, item || "-", result, detail || ""]]);
  } catch (e) { /* logging must never break the caller */ }
}

// ============================================================
// api_getPlannerData — one batched read powering every screen
// ============================================================
function api_getPlannerData() {
  try {
    var ss = SpreadsheetApp.getActive();

    // ---- Settings ----
    var settingsSheet = ss.getSheetByName(SETTINGS_SHEET);
    var settingsRows = readSheetRows_(SETTINGS_SHEET, HEADER_ROW, "Setting");
    var settingsMap = {};
    settingsRows.forEach(function (r) { settingsMap[r["Setting"]] = r["Value"]; });
    var settings = {
      semesterName: settingsMap["Semester name"] || "",
      semesterStart: isoDate_(settingsMap["Semester start date"]),
      semesterEnd: isoDate_(settingsMap["Semester end date"]),
      calendarId: settingsMap["Planner Calendar ID (read/write)"] || settingsMap["Google Calendar name or ID"] || "",
      syncEnabled: settingsMap["Calendar sync enabled (TRUE/FALSE)"] === true,
      reminder1: num_(settingsMap["Default reminder 1 (days before)"]),
      reminder2: num_(settingsMap["Default reminder 2 (days before)"]),
      targetMark: num_(settingsMap["Target final mark (%)"]),
      passMark: num_(settingsMap["Pass mark (%)"]),
      distinctionMark: num_(settingsMap["Distinction mark (%)"]),
      // v1.3.0 -- the single input the Weekly Study Allocation engine treats
      // as available weekly study time; added by
      // Academic_Intelligence_v1.3.0_Migration.gs. null (not a fabricated
      // default) until that migration has run.
      weeklyStudyCapacity: num_(settingsMap["Weekly study capacity (hours)"]),
      // v1.3.1 -- workbook-wide default for the Hidden A2 inference engine's
      // "verify by recalculation" match tolerance; added by
      // Marks_Intelligence_v1.3.1_Migration.gs. A module can override this
      // via "A2 rounding tolerance override (%)" on 04 Module Rules (see
      // modules.a2ToleranceOverride below). null until that migration has
      // run -- mi_inferHiddenA2_ falls back to a fixed constant only then.
      defaultA2Tolerance: num_(settingsMap["Default A2 recalculation match tolerance (%)"])
    };

    // ---- Modules + Module Rules merged ----
    var moduleRows = readSheetRows_(MODULES_SHEET, HEADER_ROW, "Module name");
    var ruleRows = readSheetRows_(MODULE_RULES_SHEET, HEADER_ROW, "Module code");
    var rulesByCode = {};
    ruleRows.forEach(function (r) { rulesByCode[r["Module code"]] = r; });
    var modules = moduleRows.map(function (m) {
      var code = m["Module code"];
      var meta = MODULE_UI_META[code] || { id: code, short: m["Short name"] || m["Module name"], color: "#8B8B87" };
      var rule = rulesByCode[code];
      // v1.2.0 Feature: module configuration -- "Colour (hex)" on 03 Modules
      // (added by Modules_v1.2.0_Migration.gs) overrides the hardcoded
      // MODULE_UI_META colour when present; blank means "use the existing
      // default", so nothing changes for a workbook that hasn't run the
      // migration or hasn't filled the column in yet.
      var colourOverride = m["Colour (hex)"];
      return {
        code: code,
        name: m["Module name"],
        id: meta.id,
        short: meta.short,
        color: colourOverride || meta.color,
        active: m["Active status"] === "Active",
        ruleStatus: rule ? rule["Rule status"] : "Missing rules",
        afWeight: rule ? num_(rule["AF weighting"]) : null,
        a1Weight: rule ? num_(rule["A1 weighting"]) : null,
        a2Weight: rule ? num_(rule["A2 weighting"]) : null,
        // v1.3.1 (module-weighting correction, 2nd pass) -- A3's weight is
        // now a genuinely SEPARATE, explicitly-verified field, NEVER derived
        // as "100 minus AF/A1/A2". The Faculty's real model is route-based
        // (FM1=AF+A1+A2, FM2=AF+A1+A3, FM3=AF+A2+A3), not additive-to-100 --
        // see mi_getModuleWeights_ and V1.3.1_RULE_CATALOGUE.md §0-3 for the
        // full correction. null/undefined until a real, faculty-confirmed
        // value is entered in "04 Module Rules" -- FM2/FM3 are simply never
        // computed until then (never assumed equal to A2's weight, never
        // split evenly, never derived).
        a3Weight: rule ? num_(rule["A3 weight (%) (verified)"]) : null,
        // v1.3.1 (2nd pass) -- which reading of Faculty Override Rule 1 this
        // module's regulations actually use. Blank/unset means Rule 1 is NOT
        // evaluated at all for this module (never guess between the two
        // readings -- see V1.3.1_RULE_CATALOGUE.md §5).
        subminimumMode: rule ? (rule["A2/A3 subminimum mode"] || "") : "",
        passMark: rule ? num_(rule["Pass mark"]) : 50,
        distinctionMark: rule ? num_(rule["Distinction mark"]) : 75,
        rulesVerified: !!(rule && rule["Rule status"] === "Verified"),
        afMethod: rule ? rule["AF calculation method"] : "UNKNOWN",
        // v1.4.1 -- "Tutorial tests" / "Practicals" / "" (not yet chosen).
        // The "+ Add Assessment" form's Item type choice locks to whichever
        // is set here -- see AF_Components_v1.4.1_Migration.gs.
        afItemType: rule ? (rule["AF item type"] || "") : "",
        generalNotes: rule ? rule["General notes"] : "No assessment framework supplied for this module yet.",
        // v1.3.1 -- Faculty Override Rule 2 exception + DCA enabled toggles.
        // Both default to false (Rule 2 ACTIVE, DCA OFF) for a module that
        // hasn't explicitly set the column -- see
        // Marks_Intelligence_v1.3.1_Migration.gs and V1.3.1_RULE_CATALOGUE.md.
        // ("Rule 1 disabled" was retired in the 2nd correction pass in favor
        // of subminimumMode above, which is unset/inactive by the same
        // default-safe convention -- one configuration surface, not two.)
        rule2Exception: rule ? (rule["Rule 2 exception (TRUE/FALSE)"] === true) : false,
        dcaEnabled: rule ? (rule["DCA enabled (TRUE/FALSE)"] === true) : false,
        // v1.3.1 -- optional per-module override of the workbook-wide
        // "Default A2 recalculation match tolerance (%)" Settings value.
        a2ToleranceOverride: rule ? num_(rule["A2 rounding tolerance override (%)"]) : null,
        // v1.2.0 -- module configuration fields (Feature: Module difficulty,
        // workload and repeated-module configuration + Minimal contacts).
        // Every one of these is undefined/blank until Modules_v1.2.0_Migration.gs
        // has been run AND a value has been typed in -- num_()/=== true/|| ""
        // all degrade to a harmless null/false/empty default until then.
        difficulty: num_(m["Conceptual difficulty (1-5)"]),
        workload: num_(m["Workload intensity (1-5)"]),
        repeated: m["Repeated module (TRUE/FALSE)"] === true,
        targetMark: num_(m["Target mark (%)"]),
        minWeeklyHours: num_(m["Min weekly study hours"]),
        attendanceSensitive: m["Attendance-sensitive lecture (TRUE/FALSE)"] === true,
        compulsoryPractical: m["Compulsory practical/lab (TRUE/FALSE)"] === true,
        coordinatorName: m["Coordinator name"] || "",
        coordinatorEmail: m["Coordinator email"] || "",
        lecturerName: m["Lecturer name"] || "",
        lecturerEmail: m["Lecturer email"] || "",
        // v1.4.2 -- Study Planner-specific include switch (independent of
        // "Active status", which drives everything else). A blank cell
        // (module added before/without the migration having set it) is
        // treated as TRUE, never silently excluded -- only an explicit
        // FALSE excludes.
        includeInStudyPlan: m["Include in study plan (TRUE/FALSE)"] !== false,
        // v1.4.2's persistent "Manual priority override" read is retired as
        // of v1.4.3 (see ai_computePriority_) -- the column may still exist
        // on a workbook that ran that migration, but nothing reads it
        // anymore; priority for a module with no verified framework is now
        // entered fresh per generation instead (api_generateStudySuggestions).
        contactNotes: m["Contact notes"] || ""
      };
    });

    // ---- Marks Tracker (read the computed columns as-is — never recompute client-side) ----
    var marksRows = readSheetRows_(MARKS_TRACKER_SHEET, MARKS_TRACKER_HEADER_ROW, "Module");
    var marksTracker = marksRows.map(function (r) {
      return {
        code: r["Module code"], module: r["Module"], ruleStatus: r["Rule status"],
        af: num_(r["AF %"]), a1: num_(r["A1 %"]), a2: num_(r["A2 %"]), a3: num_(r["A3 %"]),
        // v1.3.1 (2nd pass) -- this sheet's own MTD/pass-status/A3-risk
        // formula columns are kept as read-only REFERENCE values only
        // (renamed sheetMtd/sheetPassStatus/sheetA3Risk so nothing confuses
        // them with the new engine's own, written-status-aware, subminimum-
        // mode-configurable calculations below) -- they are never treated as
        // authoritative by mi_computeMtd_ / mi_computeFmpAndOfficial_.
        sheetMtd: r["MTD"], provisionalAfterA2: r["Provisional final (after A2)"],
        finalOrProvisional: r["Final / provisional (after A3)"],
        req50: r["Required A2 for 50%"], req60: r["Required A2 for 60%"],
        req75: r["Required A2 for 75%"], req80: r["Required A2 for 80%"],
        a2Submin: r["A2 subminimum"], sheetPassStatus: r["Pass status"], sheetA3Risk: r["A3 risk"],
        notes: r["Notes"], manualAf: num_(r["Manual combined AF entry (%) — reference only"]),
        // v1.3.1 -- deliberately SEPARATE from provisionalAfterA2 above (that
        // column is this sheet's own FORMULA output, computed only once a
        // real A2 is on record). "Published Final (Before A3)" is a manual
        // entry for a module that does NOT publish A2 directly and instead
        // publishes one combined pre-A3 percentage -- the Hidden A2
        // inference engine reads ONLY this field now, never
        // provisionalAfterA2 (a correction from the earlier v1.3.0 pass --
        // see V1.3.1_CHANGE_SUMMARY.md). null until
        // Marks_Intelligence_v1.3.1_Migration.gs has run and a value has
        // been entered.
        publishedFinalBeforeA3: r["Published Final (Before A3)"],
        // v1.4.0 -- mirror of the field above for the AFTER-A3 case: a
        // module that does not publish a raw A3 mark, only a new combined
        // final percentage once A3 is marked. Read by
        // mi_computeMarksIntelligence_ only (never by the Hidden A2
        // inference engine, which is a before-A3 concern) -- see there for
        // how it's folded into the Official Final Mark.
        publishedFinalAfterA3: r["Published Final (After A3)"],
        // v1.3.1 (2nd pass) -- explicit, independent written/not-written/
        // excused/deferred status per assessment. A BLANK MARK MUST NOT BE
        // ASSUMED "not written" (spec correction) -- this is why a separate
        // status field exists rather than inferring status from the mark
        // cell. See mi_statusOf_.
        a1Status: r["A1 status"] || "", a2Status: r["A2 status"] || "", a3Status: r["A3 status"] || "",
        // v1.3.1 (2nd pass) -- the actual substitute mark awarded under DCA
        // (Deemed Continuous Assessment), when applicable. This is NOT "the
        // higher of A2/A3" (that was a disclosed guess in the first
        // correction pass, now retired) -- it is the module's own examiner-
        // awarded DCA figure, entered here, which replaces whichever of
        // A2/A3 is lower. See mi_applyDca_.
        dcaMark: num_(r["DCA mark (%)"])
      };
    });

    // ---- AF Components (checklist state: which named items have been written/excused) ----
    var afRows = readSheetRows_(AF_COMPONENTS_SHEET, HEADER_ROW, "Module");
    var afComponents = afRows.map(function (r) {
      return {
        id: r["AF Item ID"], module: r["Module"], itemName: r["Item name"],
        mark: num_(r["Mark"]), max: num_(r["Maximum mark"]), pct: num_(r["Percentage"]),
        date: isoDate_(r["Date"]),
        written: r["Written"] === true, excused: r["Excused or excluded"] === true,
        included: r["Included in AF"] === true, notes: r["Notes"],
        // v1.4.1 -- blank until AF_Components_v1.4.1_Migration.gs has run.
        itemType: r["Item type"] || "", weight: num_(r["Weight"])
      };
    });

    // ---- Academic Inbox ----
    var inbox = readSheetRows_(INBOX_SHEET, HEADER_ROW, "Title").map(function (r) {
      // v1.1.12: derive filedType from the existing "Processed into (sheet!row)"
      // column (already written by processAcademicInbox()/processOneInboxItemById_,
      // not a new column) so the frontend can show "Filed as Assignment/Assessment"
      // accurately even after a full page reload, without guessing.
      var processedInto = r["Processed into (sheet!row)"] || "";
      // v1.4.0 -- now goes through the same resolveFiledRecord_ helper
      // processOneInboxItemById_ uses, so filedId (the filed record's own
      // stable Assignment/Assessment ID) comes along too -- needed for the
      // Inbox "delete this AND its filed record" option.
      var resolvedFiled = resolveFiledRecord_(processedInto);
      return {
        id: r["Inbox ID"], module: r["Module"], type: r["Item type"], title: r["Title"],
        dueDate: isoDate_(r["Due date"]), dueTime: isoTime_(r["Due time"]), venue: r["Venue or link"],
        notes: r["Notes"], priority: r["Priority"], processed: r["Processed"] === true,
        sendToCalendar: r["Send to Calendar"] === true, filedType: resolvedFiled.filedType, filedId: resolvedFiled.filedId
      };
    });

    // ---- Assignments ----
    var assignments = readSheetRows_(ASSIGNMENTS_SHEET, HEADER_ROW, "Title").map(function (r) {
      return {
        id: r["Assignment ID"], module: r["Module"], title: r["Title"], description: r["Description"],
        dueDate: isoDate_(r["Due date"]), dueTime: isoTime_(r["Due time"]), daysRemaining: r["Days remaining"],
        priority: r["Priority"], estHours: num_(r["Estimated hours"]), hoursCompleted: num_(r["Hours completed"]),
        status: r["Status"], completion: num_(r["Completion %"]), notes: r["Notes"],
        // v1.1.12 master -- additive fields only, needed so the frontend can
        // offer a Calendar-aware delete confirmation without a second
        // round trip. Nothing existing reads these two keys today, so no
        // other consumer is affected by their presence.
        calendarEventId: r["Calendar Event ID"] || "", syncToCalendar: r["Sync to Calendar"] === true
      };
    });

    // ---- Assessments ----
    var assessments = readSheetRows_(ASSESSMENTS_SHEET, HEADER_ROW, "Title").map(function (r) {
      return {
        id: r["Assessment ID"], module: r["Module"], type: r["Assessment type"], title: r["Title"],
        date: isoDate_(r["Date"]), startTime: isoTime_(r["Start time"]), venue: r["Venue"],
        status: r["Status"], daysRemaining: r["Days remaining"], notes: r["Notes"],
        calendarEventId: r["Calendar Event ID"] || "", syncToCalendar: r["Sync to Calendar"] === true
      };
    });

    // ---- Study Planner ----
    var studySessions = readSheetRows_(STUDY_SHEET, HEADER_ROW, "Module").map(function (r) {
      return {
        id: r["Study Session ID"], date: isoDate_(r["Date"]), startTime: isoTime_(r["Start time"]),
        endTime: isoTime_(r["End time"]), duration: r["Duration (min)"], module: r["Module"],
        type: r["Study type"], topic: r["Topic"], priority: r["Priority"],
        planned: r["Planned"] === true, completed: r["Completed"] === true, notes: r["Notes"],
        calendarEventId: r["Calendar Event ID"] || "",
        // v1.2.0 -- task-based Study Planner additive fields. status falls
        // back to a value derived from the existing Completed checkbox when
        // the new "Status" column is blank/not-yet-migrated, so old rows
        // never show as blank.
        status: r["Status"] || (r["Completed"] === true ? "Completed" : "Planned"),
        objective: r["Objective"] || "",
        actualDuration: num_(r["Actual duration (min)"]),
        confidenceAfter: num_(r["Confidence after"]),
        origin: r["Origin"] || "Manual"
      };
    });

    // ---- Study Tasks (v1.2.0 -- new thin checklist sheet; [] if the
    // migration hasn't been run yet, since readSheetRows_ already returns []
    // for a sheet that doesn't exist) ----
    var studyTasks = readSheetRows_(STUDY_TASKS_SHEET, HEADER_ROW, "Task ID").map(function (t) {
      return {
        id: t["Task ID"], sessionId: t["Study Session ID"], module: t["Module"], topic: t["Topic"],
        taskType: t["Task type"], text: t["Task text"], resourceId: t["Resource ID"] || "",
        completed: t["Completed"] === true, score: num_(t["Score"]), notes: t["Notes"] || "",
        sortOrder: num_(t["Sort order"]) || 0,
        // v1.3.0 -- read-only addition: feeds the Improvement Tracker's
        // "completed Study Tasks" weekly trend. This column already existed
        // in v1.2.0 (api_addStudyTask/api_updateStudyTask already write it);
        // v1.3.0 is simply the first feature to also read it back.
        lastUpdated: isoDate_(t["Last updated"])
      };
    }).sort(function (a, b) { return a.sortOrder - b.sortOrder; });

    // ---- Revision Tracker ----
    var revisionTracker = readSheetRows_(REVISION_SHEET, HEADER_ROW, "Module").map(function (r) {
      return {
        module: r["Module"], topic: r["Topic"], confidence: num_(r["Confidence"]),
        priority: r["Priority"], nextRevision: isoDate_(r["Next revision date"]),
        completed: r["Completed"] === true, notes: r["Notes"],
        // v1.2.0 -- auto-updating aggregate fields (Feature: Revision
        // Tracker integration). All additive; blank/undefined until the
        // migration has run and at least one study session has updated them.
        totalActualMinutes: num_(r["Total actual study minutes"]),
        sessionCount: num_(r["Session count"]),
        lastRevised: isoDate_(r["Last revised"]),
        completedTaskCount: num_(r["Completed task count"]),
        score: num_(r["Score"]),
        weakTopic: r["Weak topic (TRUE/FALSE)"] === true,
        plannedCount: num_(r["Planned-study count"]),
        independentCount: num_(r["Independent-study count"]),
        source: r["Source"] || "",
        // v1.3.0 -- set/cleared by api_updateRevisionTopic; see
        // Academic_Intelligence_v1.3.0_Migration.gs. Blank until that
        // migration has run and this topic has been flagged weak at least
        // once since.
        weakTopicSince: isoDate_(r["Weak topic since"])
      };
    });

    // ---- Attendance ----
    var attendance = readSheetRows_(ATTENDANCE_SHEET, HEADER_ROW, "Module").map(function (r) {
      return {
        date: isoDate_(r["Date"]), module: r["Module"], type: r["Session type"], attended: r["Attended"] === true, notes: r["Notes"],
        // v1.2.0 -- simple attendance-sensitive tracking (Feature 7). status
        // falls back to the existing boolean when blank; sessionKey lets the
        // frontend match a real timetable session to its recorded status
        // without a second lookup round trip.
        attendanceId: r["Attendance ID"] || "",
        status: r["Status"] || (r["Attended"] === true ? "Attended" : "Not recorded"),
        sessionKey: r["Session key"] || ""
      };
    });

    // ---- Resources ----
    var resources = readSheetRows_(RESOURCES_SHEET, HEADER_ROW, "Title").map(function (r) {
      return {
        module: r["Module"], type: r["Resource type"], title: r["Title"], link: r["Link or file reference"],
        topic: r["Topic"], description: r["Description"], priority: r["Priority"],
        reviewed: r["Reviewed"] === true, notes: r["Notes"],
        // v1.2.0 -- lightweight Resources CRUD + type-specific progress
        // (Feature 6). All additive; blank until Resources_v1.2.0_Migration.gs
        // has run.
        resourceId: r["Resource ID"] || "",
        status: r["Status"] || (r["Reviewed"] === true ? "Reviewed" : "Not started"),
        flagA: r["Progress Flag A"] === true, flagB: r["Progress Flag B"] === true,
        flagC: r["Progress Flag C"] === true, flagD: r["Progress Flag D"] === true,
        flashcardsCreated: r["Flashcards created (TRUE/FALSE)"] === true,
        score: num_(r["Score"]),
        // v1.3.0 -- set by api_updateResourceProgress whenever a Score is
        // saved; see Academic_Intelligence_v1.3.0_Migration.gs. Blank for any
        // score saved before that migration ran, or before v1.3.0 was
        // installed -- the Improvement Tracker correctly treats those as "no
        // dated evidence" rather than guessing an order.
        scoreUpdated: isoDate_(r["Score updated"])
      };
    });

    // ---- Notes (quick-capture block added at the bottom of 16 Notes) ----
    var notesSheet = ss.getSheetByName(NOTES_SHEET);
    var notes = { today: "", parking: "", questions: "", remember: "" };
    if (notesSheet) {
      notes.today = notesSheet.getRange("C61").getValue() || "";
      notes.parking = notesSheet.getRange("C62").getValue() || "";
      notes.questions = notesSheet.getRange("C63").getValue() || "";
      notes.remember = notesSheet.getRange("C64").getValue() || "";
    }

    // ---- Archive ---- (headers: Semester, Module, Final mark, Result, Notes — 17 Archive row 12)
    var archiveRows = readSheetRows_(ARCHIVE_SHEET, 12, "Module").map(function (r) {
      return { semester: r["Semester"], module: r["Module"], mark: r["Final mark"], status: r["Result"] };
    });

    // ---- Automation log + trigger status ----
    // v1.2.0 -- widened from the last 15 to the last 60 rows so the new
    // client-side category filters (All/Inbox/Calendar/Study/Resources/
    // Marks/Settings/Errors) on the Automation screen have enough history to
    // actually be useful once a filter narrows things down. Still a bounded,
    // read-only slice of an existing sheet -- no new sheet, no unbounded
    // growth, no behavior change to logAutomation_() itself.
    var logRows = readSheetRows_(AUTOMATION_LOG_SHEET, 4, "Action");
    var syncLog = logRows.slice(-60).reverse().map(function (r) {
      return { timestamp: isoDate_(r["Timestamp"]) + " " + isoTime_(r["Timestamp"]), action: r["Action"], item: r["Item"], result: r["Result"], detail: r["Detail"] };
    });
    var triggers = ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); });
    var automationStatus = {
      enabled: triggers.indexOf("processAcademicInbox") !== -1 && triggers.indexOf("syncAllToCalendar") !== -1,
      handlers: triggers
    };

    // ---- Timetable: current month grid + current week ----
    var calendarDays = buildMonthGrid_();
    var weekDays = buildCurrentWeek_();
    var guideEntries = readUserGuide_();
    var currentPhaseIndex = currentSemesterPhaseIndex_();

    // v1.4.2 -- flat {module,type,date} list straight off "05 Timetable
    // Import", used only to find each module's next Tutorial session (see
    // computeAcademicIntelligence_'s nextTutorialByModule) -- a tutorial is
    // when AF tests actually get written, so it's its own priority signal
    // alongside the next A1/A2/A3. [] on a workbook with no timetable data,
    // never fabricated.
    var timetableData = getSheetRowsRaw_(TIMETABLE_SHEET, HEADER_ROW);
    var timetableSessions = [];
    if (timetableData) {
      var ttDateIdx = timetableData.headers.indexOf("Start date"), ttModuleIdx = timetableData.headers.indexOf("Module"), ttTypeIdx = timetableData.headers.indexOf("Session type");
      timetableData.rows.forEach(function (r) {
        if (!r[ttModuleIdx]) return;
        timetableSessions.push({ module: r[ttModuleIdx], type: r[ttTypeIdx], date: isoDate_(r[ttDateIdx]) });
      });
    }

    // v1.3.0 -- Deterministic Academic Intelligence. Reuses every array
    // already built above (no extra sheet reads except "22 Exam Focus Log",
    // which v1.2.0 never read at all). See computeAcademicIntelligence_ for
    // the full breakdown; on a workbook that hasn't run
    // Academic_Intelligence_v1.3.0_Migration.gs yet, this still runs safely
    // -- every field it depends on that the migration adds simply reads back
    // blank/null, which every ai_... function already treats as "not enough
    // data" rather than fabricating a value.
    var academicIntelligence = computeAcademicIntelligence_({
      modules: modules, marksTracker: marksTracker, assessments: assessments, studySessions: studySessions,
      studyTasks: studyTasks, revisionTracker: revisionTracker, resources: resources, afComponents: afComponents,
      settings: settings, logRows: syncLog, todayIso: isoDate_(new Date()), timetableSessions: timetableSessions
    });

    return ok_({
      settings: settings, modules: modules, marksTracker: marksTracker, afComponents: afComponents,
      inbox: inbox, assignments: assignments, assessments: assessments, studySessions: studySessions,
      studyTasks: studyTasks, revisionTracker: revisionTracker, attendance: attendance, resources: resources, notes: notes,
      archiveRows: archiveRows, syncLog: syncLog, automationStatus: automationStatus,
      calendarDays: calendarDays, weekDays: weekDays, guideEntries: guideEntries,
      currentPhaseIndex: currentPhaseIndex,
      // v1.3.0 additions -- additive only, same "one batched read" contract
      // every prior release's new fields have used (see V1.3.0_ARCHITECTURE_NOTES.md).
      priorities: academicIntelligence.priorities,
      weeklyAllocation: academicIntelligence.weeklyAllocation,
      nextAction: academicIntelligence.nextAction,
      resourceInsights: academicIntelligence.resourceInsights,
      examFocus: academicIntelligence.examFocus,
      improvementTracker: academicIntelligence.improvementTracker,
      // v1.4.0 final integration fix: the engine already computes this array;
      // expose the same result to the browser so Marks Intelligence cards render.
      marksIntelligence: academicIntelligence.marksIntelligence,
      generatedAt: new Date().toISOString()
    });
  } catch (e) {
    return fail_(e);
  }
}

// ============================================================
// TIMETABLE / CALENDAR DERIVATION  (real dated events only — never a repeating guess)
// ============================================================
function buildModuleColorIndex_() {
  var idx = {};
  var moduleRows = readSheetRows_(MODULES_SHEET, HEADER_ROW, "Module name");
  moduleRows.forEach(function (m) {
    var meta = MODULE_UI_META[m["Module code"]];
    // v1.2.0 -- respect a per-module colour override from "Colour (hex)" the
    // same way api_getPlannerData's modules array now does, so the Calendar
    // grid and the Dashboard/Assignments/Assessments cards never disagree
    // about which colour a module is.
    idx[m["Module name"]] = m["Colour (hex)"] || (meta ? meta.color : "#8B8B87");
  });
  return idx;
}

/** Shared by buildMonthGrid_ and buildCurrentWeek_ — both need the same
 *  "headers + raw data rows" shape from a sheet; factored out so a future
 *  change to how header rows are located only has to happen in one place. */
function getSheetRowsRaw_(sheetName, headerRow) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sheet) return null;
  var lastRow = sheet.getLastRow();
  if (lastRow <= headerRow) return { headers: [], rows: [] };
  var headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  var rows = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, sheet.getLastColumn()).getValues();
  return { headers: headers, rows: rows };
}

function buildMonthGrid_() {
  var tz = TIMEZONE;
  var today = new Date();
  var year = Number(Utilities.formatDate(today, tz, "yyyy"));
  var month = Number(Utilities.formatDate(today, tz, "M")); // 1-12
  var first = new Date(year, month - 1, 1);
  var startWeekday = (first.getDay() + 6) % 7; // 0 = Monday
  var daysInMonth = new Date(year, month, 0).getDate();

  var colorIdx = buildModuleColorIndex_();
  var eventsByDay = {}; // day-of-month -> [{label,color}]

  function pushEvent(dateVal, label, color) {
    if (!(dateVal instanceof Date)) return;
    if (dateVal.getFullYear() !== year || dateVal.getMonth() !== month - 1) return;
    var d = dateVal.getDate();
    if (!eventsByDay[d]) eventsByDay[d] = [];
    eventsByDay[d].push({ label: label, color: color });
  }

  // Timetable Import (real per-date class/tutorial events)
  var ttData = getSheetRowsRaw_(TIMETABLE_SHEET, HEADER_ROW);
  if (ttData) {
    var iDate = ttData.headers.indexOf("Start date"), iModule = ttData.headers.indexOf("Module"), iType = ttData.headers.indexOf("Session type");
    ttData.rows.forEach(function (r) {
      if (!r[iModule]) return;
      pushEvent(r[iDate], r[iModule] + " " + r[iType], colorIdx[r[iModule]] || "#8B8B87");
    });
  }
  // Assessments due dates
  var asData = getSheetRowsRaw_(ASSESSMENTS_SHEET, HEADER_ROW);
  if (asData) {
    var iaDate = asData.headers.indexOf("Date"), iaModule = asData.headers.indexOf("Module"), iaType = asData.headers.indexOf("Assessment type");
    asData.rows.forEach(function (r) { if (r[iaModule]) pushEvent(r[iaDate], r[iaModule] + " " + r[iaType], colorIdx[r[iaModule]] || "#8B8B87"); });
  }
  // Assignment due dates
  var agData = getSheetRowsRaw_(ASSIGNMENTS_SHEET, HEADER_ROW);
  if (agData) {
    var igDate = agData.headers.indexOf("Due date"), igModule = agData.headers.indexOf("Module"), igTitle = agData.headers.indexOf("Title");
    agData.rows.forEach(function (r) { if (r[igModule]) pushEvent(r[igDate], r[igTitle] || (r[igModule] + " due"), colorIdx[r[igModule]] || "#8B8B87"); });
  }

  var todayNum = Number(Utilities.formatDate(today, tz, "d"));
  var days = [];
  for (var i = 0; i < startWeekday; i++) days.push({ num: "", events: [], isToday: "#F3E3D2", dim: "opacity:.4;" });
  for (var d = 1; d <= daysInMonth; d++) {
    days.push({ num: d, events: eventsByDay[d] || [], isToday: d === todayNum ? "#F7CBAE" : "#FDF6F0", dim: "" });
  }
  var trailing = (7 - (days.length % 7)) % 7;
  for (var j = 0; j < trailing; j++) days.push({ num: "", events: [], isToday: "#F3E3D2", dim: "opacity:.4;" });
  return days;
}

// v1.4.0 -- Google Calendar pull-sync for the monthly view, lazy-loaded by
// the frontend when the Calendar screen opens (same pattern as
// api_getWeekTimetable / api_semDashboard) rather than folded into every
// api_getPlannerData call, since it makes an external Calendar API request.
// Read-only: never writes anything. See buildExternalCalendarEvents_ for
// what "external" means here (excludes events this workbook itself created).
function api_getExternalCalendarEvents() {
  try {
    var tz = TIMEZONE;
    var today = new Date();
    var year = Number(Utilities.formatDate(today, tz, "yyyy"));
    var month = Number(Utilities.formatDate(today, tz, "M")); // 1-12
    var monthStart = new Date(year, month - 1, 1);
    var monthEndExclusive = new Date(year, month, 1);
    return ok_(buildExternalCalendarEvents_(monthStart, monthEndExclusive));
  } catch (e) { return fail_(e); }
}

// v1.1.12 master -- Feature 10 (Weekly Timetable preview). buildCurrentWeek_
// used to hardcode "today" internally; it is now a one-line wrapper around
// buildWeekForDate_(anchorDate) so the exact same grouping/sorting logic
// (unchanged below) can be reused for any week, not just the current one.
// mondayOfWeek_ is factored out too so api_getWeekTimetable can compute the
// same Monday a caller asked for without duplicating the "(getDay()+6)%7"
// trick in two places.
function mondayOfWeek_(anchorDate) {
  var dow = (anchorDate.getDay() + 6) % 7; // 0 = Monday
  return new Date(anchorDate.getFullYear(), anchorDate.getMonth(), anchorDate.getDate() - dow);
}

function buildCurrentWeek_() {
  return buildWeekForDate_(new Date());
}

function buildWeekForDate_(anchorDate) {
  var tz = TIMEZONE;
  var monday = mondayOfWeek_(anchorDate);
  var friday = new Date(monday.getFullYear(), monday.getMonth(), monday.getDate() + 4);

  var colorIdx = buildModuleColorIndex_();
  var labels = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"];
  var byDay = { Monday: [], Tuesday: [], Wednesday: [], Thursday: [], Friday: [] };

  var ttData = getSheetRowsRaw_(TIMETABLE_SHEET, HEADER_ROW);
  if (ttData) {
    var iDate = ttData.headers.indexOf("Start date"), iModule = ttData.headers.indexOf("Module"),
      iType = ttData.headers.indexOf("Session type"), iStart = ttData.headers.indexOf("Start time"), iEnd = ttData.headers.indexOf("End time"),
      iDay = ttData.headers.indexOf("Day"), iVenue = ttData.headers.indexOf("Venue");
    ttData.rows.forEach(function (r) {
      var d = r[iDate];
      if (!(d instanceof Date)) return;
      if (d < monday || d > friday) return;
      var dayLabel = r[iDay] || labels[(d.getDay() + 6) % 7];
      if (!byDay[dayLabel]) return;
      var startMin = r[iStart] instanceof Date ? r[iStart].getHours() * 60 + r[iStart].getMinutes() : null;
      var endMin = r[iEnd] instanceof Date ? r[iEnd].getHours() * 60 + r[iEnd].getMinutes() : null;
      byDay[dayLabel].push({
        module: r[iModule],
        time: (r[iStart] instanceof Date ? Utilities.formatDate(r[iStart], tz, "HH:mm") : "") + "–" + (r[iEnd] instanceof Date ? Utilities.formatDate(r[iEnd], tz, "HH:mm") : ""),
        type: r[iType], venue: iVenue >= 0 ? r[iVenue] : "",
        color: colorIdx[r[iModule]] || "#8B8B87",
        // v1.4.0 -- kept through to the client (previously computed only to
        // sort here, then deleted) so the Weekly Timetable can position each
        // class by its real clock time instead of just stacking cards in
        // order -- a 08:00 class and a 10:00 class on different days used to
        // render as each column's first card, vertically aligned with each
        // other despite being 2 hours apart.
        startMin: startMin, endMin: endMin,
        // v1.2.0 -- carry the exact class date through so the frontend can
        // build a Module|SessionType|DateISO attendance session key without
        // a second lookup (Feature 7).
        dateIso: Utilities.formatDate(d, tz, "yyyy-MM-dd")
      });
    });
  }
  return labels.map(function (label) {
    var sessions = byDay[label].sort(function (a, b) { return (a.startMin || 0) - (b.startMin || 0); });
    return { label: label, sessions: sessions };
  });
}

// v1.1.12 master -- Feature 10. Read-only: no writes, no Calendar calls,
// same TIMEZONE-aware grouping as the boot-time current week. Validates its
// one input strictly and fails closed (fail_()) rather than guessing on a
// malformed date, since this endpoint can be called far more often than a
// normal page load (every Prev/Next click).
function api_getWeekTimetable(anchorDateIso) {
  try {
    if (!anchorDateIso || typeof anchorDateIso !== "string" || !/^\d{4}-\d{2}-\d{2}$/.test(anchorDateIso)) {
      return fail_(new Error("Invalid date — expected YYYY-MM-DD."));
    }
    var parts = anchorDateIso.split("-");
    var anchor = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
    if (isNaN(anchor.getTime())) return fail_(new Error("Invalid date — expected YYYY-MM-DD."));
    var monday = mondayOfWeek_(anchor);
    var friday = new Date(monday.getFullYear(), monday.getMonth(), monday.getDate() + 4);
    var weekDays = buildWeekForDate_(anchor);
    var currentMonday = mondayOfWeek_(new Date());
    return ok_({
      weekDays: weekDays,
      mondayIso: isoDate_(monday),
      fridayIso: isoDate_(friday),
      isCurrentWeek: isoDate_(currentMonday) === isoDate_(monday)
    });
  } catch (e) {
    return fail_(e);
  }
}

// ============================================================
// MUTATIONS — Academic Inbox
// ============================================================
var VALID_INBOX_TYPES = ["Assignment", "Project", "Presentation", "Test", "Exam", "Other"];
var VALID_PRIORITIES = ["Low", "Medium", "High"];

function api_addInboxItem(form) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(LOCK_WAIT_MS)) return fail_(new Error("Workbook is busy — try again in a moment."));
    if (!form) return fail_(new Error("No form data received."));
    var title = requireText_(form.title, "Title", 200);
    var module = requireModule_(form.module);
    var type = VALID_INBOX_TYPES.indexOf(form.type) !== -1 ? form.type : "Assignment";
    var priority = VALID_PRIORITIES.indexOf(form.priority) !== -1 ? form.priority : "Medium";
    var dueDate = form.dueDate ? parseValidDate_(form.dueDate, "Due date") : null;
    var notes = cleanMultiline_(form.notes, 1000);
    var venue = cleanText_(form.venue, 200);

    var sheet = SpreadsheetApp.getActive().getSheetByName(INBOX_SHEET);
    var map = getColMap_(sheet);
    var row = firstBlankRow_(sheet, col_(map, "Title"));
    var newId = nextId_("INB", getColumnValues_(sheet, col_(map, "Inbox ID")));
    sheet.getRange(row, col_(map, "Inbox ID")).setValue(newId);
    sheet.getRange(row, col_(map, "Date captured")).setValue(new Date());
    sheet.getRange(row, col_(map, "Module")).setValue(module);
    sheet.getRange(row, col_(map, "Item type")).setValue(type);
    sheet.getRange(row, col_(map, "Title")).setValue(title);
    if (dueDate) sheet.getRange(row, col_(map, "Due date")).setValue(dueDate);
    if (form.dueTime) sheet.getRange(row, col_(map, "Due time")).setValue(cleanText_(form.dueTime, 10));
    if (venue) sheet.getRange(row, col_(map, "Venue or link")).setValue(venue);
    if (notes) sheet.getRange(row, col_(map, "Notes")).setValue(notes);
    sheet.getRange(row, col_(map, "Priority")).setValue(priority);
    sheet.getRange(row, col_(map, "Processed")).setValue(false);
    sheet.getRange(row, col_(map, "Send to Calendar")).setValue(!!form.sendToCalendar);
    logAutomation_("Inbox item captured", title, "Created", newId);
    return ok_({ id: newId });
  } catch (e) { return fail_(e); } finally { lock.releaseLock(); }
}

function api_toggleInboxProcessed(inboxId, processed) {
  return withRowByKey_(INBOX_SHEET, "Inbox ID", inboxId, function (sheet, map, row) {
    // v1.2.0 -- capture the human title before acting, purely so the
    // activity-log entry below reads as a title instead of a bare Inbox ID
    // (Feature: human-readable activity history). No other behaviour here
    // changed.
    var title = map["Title"] ? sheet.getRange(row, col_(map, "Title")).getValue() : inboxId;
    if (processed) {
      // Route through the real, existing inbox processor so Assignments/Assessments
      // stay the single source of truth — never hand-roll a second code path here.
      processAcademicInbox();
    } else {
      sheet.getRange(row, col_(map, "Processed")).setValue(false);
    }
    logAutomation_("Inbox item toggled", title || inboxId, "Updated", "processed=" + processed + " (ID " + inboxId + ")");
  });
}

/**
 * v1.1.12 -- delete one Academic Inbox capture by its stable Inbox ID (never
 * by displayed row position). Deletes only the Inbox row itself -- never
 * touches any Assignment/Assessment it may already have been filed as, and
 * never touches Calendar. The frontend confirmation modal is responsible for
 * explaining that distinction to the user before this is ever called.
 * Core logic factored into deleteInboxRowCore_ (no locking of its own) so
 * api_deleteInboxItemAndFiled below can run it under a single outer lock
 * alongside the filed-record delete, instead of nesting two independent
 * LockService acquisitions in one execution.
 */
function deleteInboxRowCore_(inboxId) {
  var id = requireText_(inboxId, "Inbox ID", 40);
  var sheet = SpreadsheetApp.getActive().getSheetByName(INBOX_SHEET);
  if (!sheet) return fail_(new Error("Sheet not found: " + INBOX_SHEET));
  var map = getColMap_(sheet);
  var lastRow = sheet.getLastRow();
  var idCol = col_(map, "Inbox ID");
  var ids = sheet.getRange(HEADER_ROW + 1, idCol, Math.max(lastRow - HEADER_ROW, 0), 1).getValues();
  var targetRow = -1;
  for (var i = 0; i < ids.length; i++) { if (ids[i][0] === id) { targetRow = HEADER_ROW + 1 + i; break; } }
  if (targetRow === -1) return fail_(new Error("Could not find Inbox item " + id + " — it may already have been deleted."));
  var wasProcessed = sheet.getRange(targetRow, col_(map, "Processed")).getValue() === true;
  // v1.2.0 -- capture the human title before deleting, so the activity log
  // shows the title rather than a bare Inbox ID (Feature: human-readable
  // activity history). The ID is kept in the detail text.
  var title = map["Title"] ? sheet.getRange(targetRow, col_(map, "Title")).getValue() : id;
  sheet.deleteRow(targetRow);
  logAutomation_("Inbox item deleted", title || id, "Deleted", "ID " + id + "; row " + targetRow + (wasProcessed ? " (was processed)" : ""));
  return ok_({ inboxId: id, wasProcessed: wasProcessed });
}
function api_deleteInboxItem(inboxId) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(LOCK_WAIT_MS)) return fail_(new Error("Workbook is busy — try again in a moment."));
    return deleteInboxRowCore_(inboxId);
  } catch (e) { return fail_(e); } finally { lock.releaseLock(); }
}

/**
 * v1.4.0 -- companion to api_deleteInboxItem for the case the delete
 * confirmation modal now explicitly offers: also delete the Assignment/
 * Assessment this Inbox item was already filed as (and its Calendar event,
 * if `deleteCalendarEvent` is true), not just the Inbox capture. `filedType`
 * / `filedId` come from the Inbox item's own filedType/filedId (resolved by
 * resolveFiledRecord_ in api_getPlannerData) -- never re-derived by title/
 * module matching, which could hit the wrong row.
 * The filed record is deleted FIRST; if that fails (or if a requested
 * Calendar deletion fails), the Inbox row is left untouched too, so a
 * partial failure never silently leaves things half-deleted.
 */
function api_deleteInboxItemAndFiled(inboxId, filedType, filedId, deleteCalendarEvent) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(LOCK_WAIT_MS)) return fail_(new Error("Workbook is busy — try again in a moment."));
    var filedResult = null;
    if (filedType && filedId) {
      var sheetName = filedType === "Assignment" ? ASSIGNMENTS_SHEET : filedType === "Assessment" ? ASSESSMENTS_SHEET : null;
      var idHeader = filedType === "Assignment" ? "Assignment ID" : "Assessment ID";
      if (!sheetName) return fail_(new Error("Unknown filed type: " + filedType));
      filedResult = deleteEntityRowCore_(sheetName, idHeader, filedId, !!deleteCalendarEvent, filedType);
      if (!filedResult.ok) return filedResult;
    }
    var inboxResult = deleteInboxRowCore_(inboxId);
    if (!inboxResult.ok) return inboxResult;
    return ok_({ inboxId: inboxId, filed: filedResult ? filedResult.data : null });
  } catch (e) { return fail_(e); } finally { lock.releaseLock(); }
}

/**
 * v1.1.12 -- thin web-app wrapper around processOneInboxItemById_() in
 * Calendar_Sync_v2.0.gs, mirroring the existing api_processInbox() /
 * api_syncCalendar() pattern (validate here, do the real work in the
 * Calendar_Sync file). Files exactly the one requested Inbox item and, if
 * it requested Calendar sync, syncs only that newly filed row -- never a
 * full processAcademicInbox()/syncAllToCalendar() sweep. Idempotent: safe
 * to call more than once for the same Inbox ID.
 */
function api_processInboxItem(inboxId) {
  try {
    var id = requireText_(inboxId, "Inbox ID", 40);
    var result = processOneInboxItemById_(id);
    return ok_(result);
  } catch (e) {
    return fail_(e);
  }
}

// ============================================================
// MUTATIONS — AF Components (written / excused checklist)
// ============================================================
function api_setAfFlag(moduleName, itemName, weightLabel, field, value) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(LOCK_WAIT_MS)) return fail_(new Error("Workbook is busy — try again in a moment."));
    var sheet = SpreadsheetApp.getActive().getSheetByName(AF_COMPONENTS_SHEET);
    var map = getColMap_(sheet);
    var lastRow = sheet.getLastRow();
    var moduleCol = col_(map, "Module"), nameCol = col_(map, "Item name");
    var foundRow = null;
    if (lastRow > HEADER_ROW) {
      var vals = sheet.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, sheet.getLastColumn()).getValues();
      for (var i = 0; i < vals.length; i++) {
        if (vals[i][moduleCol - 1] === moduleName && vals[i][nameCol - 1] === itemName) { foundRow = HEADER_ROW + 1 + i; break; }
      }
    }
    if (!foundRow) {
      foundRow = firstBlankRow_(sheet, moduleCol);
      var newId = nextId_("AF", getColumnValues_(sheet, col_(map, "AF Item ID")));
      sheet.getRange(foundRow, col_(map, "AF Item ID")).setValue(newId);
      sheet.getRange(foundRow, moduleCol).setValue(moduleName);
      sheet.getRange(foundRow, nameCol).setValue(itemName);
    }
    var targetHeader = field === "written" ? "Written" : "Excused or excluded";
    sheet.getRange(foundRow, col_(map, targetHeader)).setValue(!!value);
    logAutomation_("AF flag updated", moduleName + " — " + itemName, "Updated", field + "=" + value);
    return ok_({ row: foundRow });
  } catch (e) { return fail_(e); } finally { lock.releaseLock(); }
}

// ============================================================
// MUTATIONS — Marks Entry (the in-app replacement for typing numbers into Sheets)
// ============================================================
var VALID_MARK_TYPES = ["AF", "A1", "A2", "A3"];
var MARKS_TRACKER_PCT_COL = { A1: 5, A2: 6, A3: 7 }; // columns E, F, G

/**
 * Records one real assessment mark.
 *  - type "AF": finds-or-creates the matching row in 10 AF Components (same
 *    matching rule as api_setAfFlag, so a checklist tick and a typed mark for
 *    the same named item always land in the same row, never a duplicate).
 *  - type "A1"/"A2"/"A3": writes the computed percentage into the exact cell
 *    11 Marks Tracker already expects manual entry in (see 18 User Guide,
 *    "Entering marks"). Nothing here recomputes AF/A1/A2/A3 weighting —
 *    that stays 100% inside the sheet's own formulas.
 * Every input is validated server-side; this function is safe to call even
 * if the client-side validation in Component.js were somehow bypassed.
 */
function api_addMarkEntry(moduleCode, moduleName, form) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(LOCK_WAIT_MS)) return fail_(new Error("Workbook is busy — try again in a moment."));
    if (!form) return fail_(new Error("No form data received."));

    var modules = readSheetRows_(MODULES_SHEET, HEADER_ROW, "Module name");
    var moduleRow = modules.filter(function (m) { return m["Module code"] === moduleCode; })[0];
    if (!moduleRow) throw new Error("Unknown module code: " + moduleCode);
    var verifiedModuleName = moduleRow["Module name"];

    var type = VALID_MARK_TYPES.indexOf(form.type) !== -1 ? form.type : "AF";
    var mark = requireFiniteNumber_(form.mark, "Mark", 0, null);
    var max = requireFiniteNumber_(form.max, "Maximum mark", 0.01, null);
    if (mark > max * 1.05) throw new Error("Mark (" + mark + ") cannot exceed the maximum (" + max + ").");
    var dateVal = form.date ? parseValidDate_(form.date, "Date") : new Date();

    if (type === "AF") {
      var name = requireText_(form.name, "Assessment name", 150);
      var sheet = SpreadsheetApp.getActive().getSheetByName(AF_COMPONENTS_SHEET);
      var map = getColMap_(sheet);
      var moduleCol = col_(map, "Module"), nameCol = col_(map, "Item name");
      var lastRow = sheet.getLastRow();
      var foundRow = null;
      if (lastRow > HEADER_ROW) {
        var vals = sheet.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, sheet.getLastColumn()).getValues();
        for (var i = 0; i < vals.length; i++) {
          if (vals[i][moduleCol - 1] === verifiedModuleName && vals[i][nameCol - 1] === name) { foundRow = HEADER_ROW + 1 + i; break; }
        }
      }
      if (!foundRow) {
        foundRow = firstBlankRow_(sheet, moduleCol);
        var newId = nextId_("AF", getColumnValues_(sheet, col_(map, "AF Item ID")));
        sheet.getRange(foundRow, col_(map, "AF Item ID")).setValue(newId);
        sheet.getRange(foundRow, moduleCol).setValue(verifiedModuleName);
        sheet.getRange(foundRow, nameCol).setValue(name);
      }
      var writtenVal = (typeof form.written === "boolean") ? form.written : true;
      var excusedVal = !!form.excused;
      var notesVal = cleanMultiline_(form.notes, 500);
      sheet.getRange(foundRow, col_(map, "Mark")).setValue(mark);
      sheet.getRange(foundRow, col_(map, "Maximum mark")).setValue(max);
      sheet.getRange(foundRow, col_(map, "Date")).setValue(dateVal);
      sheet.getRange(foundRow, col_(map, "Written")).setValue(writtenVal);
      sheet.getRange(foundRow, col_(map, "Excused or excluded")).setValue(excusedVal);
      sheet.getRange(foundRow, col_(map, "Included in AF")).setValue(!!form.includeInAf);
      if (map["Notes"]) sheet.getRange(foundRow, col_(map, "Notes")).setValue(notesVal);
      // v1.4.1 -- Item type / Weight, both optional, both no-ops if the
      // AF_Components_v1.4.1_Migration.gs columns haven't been added yet
      // (never throws just because that migration hasn't run).
      if (map["Item type"] && VALID_AF_ITEM_TYPES.indexOf(form.itemType) !== -1) {
        sheet.getRange(foundRow, col_(map, "Item type")).setValue(form.itemType);
      }
      if (map["Weight"]) {
        if (form.weight === "" || form.weight === null || typeof form.weight === "undefined") {
          sheet.getRange(foundRow, col_(map, "Weight")).setValue("");
        } else {
          sheet.getRange(foundRow, col_(map, "Weight")).setValue(requireFiniteNumber_(form.weight, "Weight", 0, null));
        }
      }
      // Percentage column is a pre-existing formula (=Mark/Maximum*100) — never overwritten here.
      logAutomation_("Mark entered (AF)", verifiedModuleName + " — " + name, "Saved", mark + "/" + max);
      var afItemId = sheet.getRange(foundRow, col_(map, "AF Item ID")).getValue();
      return ok_({ row: foundRow, id: afItemId, pct: max ? Math.round((mark / max) * 1000) / 10 : null });
    }

    // A1 / A2 / A3 — direct percentage entry into 11 Marks Tracker's manual input cell.
    var mtRow = findMarksTrackerRow_(moduleCode);
    if (mtRow === -1) throw new Error("Module not found in Marks Tracker: " + moduleCode);
    var pct = Math.round((mark / max) * 100); // "Standard (nearest whole %)" rounding, matching 04 Module Rules
    var mtSheet = SpreadsheetApp.getActive().getSheetByName(MARKS_TRACKER_SHEET);
    mtSheet.getRange(mtRow, MARKS_TRACKER_PCT_COL[type]).setValue(pct);
    logAutomation_("Mark entered (" + type + ")", verifiedModuleName, "Saved", pct + "%");
    return ok_({ row: mtRow, pct: pct });

  } catch (e) { return fail_(e); } finally { lock.releaseLock(); }
}

function findMarksTrackerRow_(moduleCode) {
  var mtSheet = SpreadsheetApp.getActive().getSheetByName(MARKS_TRACKER_SHEET);
  var mtLastRow = mtSheet.getLastRow();
  var mtVals = mtSheet.getRange(MARKS_TRACKER_HEADER_ROW + 1, 1, Math.max(mtLastRow - MARKS_TRACKER_HEADER_ROW, 0), 1).getValues();
  for (var j = 0; j < mtVals.length; j++) { if (mtVals[j][0] === moduleCode) return MARKS_TRACKER_HEADER_ROW + 1 + j; }
  return -1;
}

/** Clears one A1/A2/A3 manual-entry cell back to blank. Used by the "Clear"
 *  link next to that mark type in the Add Assessment form — the one place a
 *  mistaken A1/A2/A3 entry can be undone without going into Sheets. Only
 *  touches the single cell for this module + type; every other module's row
 *  and every other percentage cell in the same row is untouched. */
function api_clearMarksTrackerPct(moduleCode, type) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(LOCK_WAIT_MS)) return fail_(new Error("Workbook is busy — try again in a moment."));
    if (VALID_MARK_TYPES.indexOf(type) === -1 || type === "AF") return fail_(new Error("Invalid type to clear: " + type));
    var mtRow = findMarksTrackerRow_(moduleCode);
    if (mtRow === -1) return fail_(new Error("Module not found in Marks Tracker: " + moduleCode));
    var mtSheet = SpreadsheetApp.getActive().getSheetByName(MARKS_TRACKER_SHEET);
    mtSheet.getRange(mtRow, MARKS_TRACKER_PCT_COL[type]).setValue("");
    logAutomation_("Mark cleared (" + type + ")", moduleCode, "Cleared", "");
    return ok_({});
  } catch (e) { return fail_(e); } finally { lock.releaseLock(); }
}

// ------------------------------------------------------------------
// v1.3.1 -- "Published Final (Before A3)". A DEDICATED endpoint, entirely
// separate from api_addMarkEntry/api_clearMarksTrackerPct above (which are
// unmodified by this release), because this value is not an A1/A2/A3
// percentage entry -- it is the module's own published combined pre-A3
// figure, used ONLY by Hidden A2 inference. Writing here never touches AF,
// A1, A2, A3, "Provisional final (after A2)", or "Final / provisional
// (after A3)" -- confirmed by reading the body below: it locates the
// "Published Final (Before A3)" column via col_()/getColMap_() (a genuinely
// new, migration-appended column) and sets exactly that one cell.
// ------------------------------------------------------------------
function api_setPublishedFinalBeforeA3(moduleCode, value) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(LOCK_WAIT_MS)) return fail_(new Error("Workbook is busy — try again in a moment."));
    var mtRow = findMarksTrackerRow_(moduleCode);
    if (mtRow === -1) throw new Error("Module not found in Marks Tracker: " + moduleCode);
    var mtSheet = SpreadsheetApp.getActive().getSheetByName(MARKS_TRACKER_SHEET);
    var map = getColMap_(mtSheet, MARKS_TRACKER_HEADER_ROW);
    if (!map["Published Final (Before A3)"]) {
      throw new Error('"Published Final (Before A3)" column not found -- run Marks_Intelligence_v1.3.1_Migration.gs first.');
    }
    if (value === "" || value === null || typeof value === "undefined") {
      mtSheet.getRange(mtRow, col_(map, "Published Final (Before A3)")).setValue("");
      logAutomation_("Published Final (Before A3) cleared", moduleCode, "Cleared", "");
      return ok_({ row: mtRow, pct: null });
    }
    var pct = requireFiniteNumber_(value, "Published Final (Before A3)", 0, 100);
    mtSheet.getRange(mtRow, col_(map, "Published Final (Before A3)")).setValue(pct);
    logAutomation_("Published Final (Before A3) entered", moduleCode, "Saved", pct + "%");
    return ok_({ row: mtRow, pct: pct });
  } catch (e) { return fail_(e); } finally { lock.releaseLock(); }
}

// ------------------------------------------------------------------
// v1.4.0 -- "Published Final (After A3)". Mirrors
// api_setPublishedFinalBeforeA3 exactly, one column over: for a module that
// does not publish a raw A3 mark, only a new combined final percentage once
// A3 is marked. Read by mi_computeMarksIntelligence_ as an alternative
// candidate for the Official Final Mark (the higher of the computed value
// and this one wins) -- never by Hidden A2 inference, which only reads the
// Before-A3 field. Writing here never touches AF, A1, A2, A3, or any other
// column.
// ------------------------------------------------------------------
function api_setPublishedFinalAfterA3(moduleCode, value) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(LOCK_WAIT_MS)) return fail_(new Error("Workbook is busy — try again in a moment."));
    var mtRow = findMarksTrackerRow_(moduleCode);
    if (mtRow === -1) throw new Error("Module not found in Marks Tracker: " + moduleCode);
    var mtSheet = SpreadsheetApp.getActive().getSheetByName(MARKS_TRACKER_SHEET);
    var map = getColMap_(mtSheet, MARKS_TRACKER_HEADER_ROW);
    if (!map["Published Final (After A3)"]) {
      throw new Error('"Published Final (After A3)" column not found -- run Marks_Intelligence_v1.4.0_Migration.gs first.');
    }
    if (value === "" || value === null || typeof value === "undefined") {
      mtSheet.getRange(mtRow, col_(map, "Published Final (After A3)")).setValue("");
      logAutomation_("Published Final (After A3) cleared", moduleCode, "Cleared", "");
      return ok_({ row: mtRow, pct: null });
    }
    var pctAfter = requireFiniteNumber_(value, "Published Final (After A3)", 0, 100);
    mtSheet.getRange(mtRow, col_(map, "Published Final (After A3)")).setValue(pctAfter);
    logAutomation_("Published Final (After A3) entered", moduleCode, "Saved", pctAfter + "%");
    return ok_({ row: mtRow, pct: pctAfter });
  } catch (e) { return fail_(e); } finally { lock.releaseLock(); }
}

var MI_VALID_STATUS_VALUES_ = ["written", "not written", "excused", "deferred"];

/**
 * v1.3.1 2nd correction pass: writes the explicit assessment status
 * ("Written" / "Not written" / "Excused" / "Deferred") for A1, A2, or A3 on
 * "11 Marks Tracker". An explicit status here always takes priority over
 * mi_statusOf_'s mark-based inference -- this is the only way to correctly
 * record "not written" (a blank mark alone is never assumed to mean that).
 * `assessment` must be exactly "a1", "a2", or "a3". `status` must be one of
 * MI_VALID_STATUS_VALUES_ (case-insensitive), or "" / null to clear back to
 * "unknown"/inferred.
 */
function api_setAssessmentStatus(moduleCode, assessment, status) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(LOCK_WAIT_MS)) return fail_(new Error("Workbook is busy — try again in a moment."));
    var a = (assessment || "").toString().trim().toLowerCase();
    if (["a1", "a2", "a3"].indexOf(a) === -1) throw new Error('assessment must be "a1", "a2", or "a3" — got "' + assessment + '".');
    var header = a.toUpperCase() + " status";

    var mtRow = findMarksTrackerRow_(moduleCode);
    if (mtRow === -1) throw new Error("Module not found in Marks Tracker: " + moduleCode);
    var mtSheet = SpreadsheetApp.getActive().getSheetByName(MARKS_TRACKER_SHEET);
    var map = getColMap_(mtSheet, MARKS_TRACKER_HEADER_ROW);
    if (!map[header]) throw new Error('"' + header + '" column not found -- run Marks_Intelligence_v1.3.1_Migration.gs first.');

    if (status === "" || status === null || typeof status === "undefined") {
      mtSheet.getRange(mtRow, col_(map, header)).setValue("");
      logAutomation_(header + " cleared", moduleCode, "Cleared", "");
      return ok_({ row: mtRow, assessment: a.toUpperCase(), status: null });
    }
    var s = status.toString().trim().toLowerCase();
    if (MI_VALID_STATUS_VALUES_.indexOf(s) === -1) {
      throw new Error('status must be one of: ' + MI_VALID_STATUS_VALUES_.join(", ") + ' (or blank to clear) — got "' + status + '".');
    }
    // Store in Title Case for a readable sheet cell; mi_statusOf_ reads case-insensitively.
    var stored = s.charAt(0).toUpperCase() + s.slice(1);
    mtSheet.getRange(mtRow, col_(map, header)).setValue(stored);
    logAutomation_(header + " set", moduleCode, "Saved", stored);
    return ok_({ row: mtRow, assessment: a.toUpperCase(), status: stored });
  } catch (e) { return fail_(e); } finally { lock.releaseLock(); }
}

/**
 * v1.3.1 2nd correction pass: writes the genuine, separately-entered DCA
 * (Deemed Continuous Assessment) mark for a module to "11 Marks Tracker"'s
 * "DCA mark (%)" column. This is NOT a computed value -- it's the real
 * substitute mark the Faculty actually awarded. mi_applyDca_ uses it to
 * replace whichever of A2/A3 is lower, only when DCA is enabled for the
 * module (see "DCA enabled (TRUE/FALSE)" on 04 Module Rules) and both A2
 * and A3 are known. Pass "" / null to clear.
 */
function api_setDcaMark(moduleCode, value) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(LOCK_WAIT_MS)) return fail_(new Error("Workbook is busy — try again in a moment."));
    var mtRow = findMarksTrackerRow_(moduleCode);
    if (mtRow === -1) throw new Error("Module not found in Marks Tracker: " + moduleCode);
    var mtSheet = SpreadsheetApp.getActive().getSheetByName(MARKS_TRACKER_SHEET);
    var map = getColMap_(mtSheet, MARKS_TRACKER_HEADER_ROW);
    if (!map["DCA mark (%)"]) throw new Error('"DCA mark (%)" column not found -- run Marks_Intelligence_v1.3.1_Migration.gs first.');

    if (value === "" || value === null || typeof value === "undefined") {
      mtSheet.getRange(mtRow, col_(map, "DCA mark (%)")).setValue("");
      logAutomation_("DCA mark (%) cleared", moduleCode, "Cleared", "");
      return ok_({ row: mtRow, pct: null });
    }
    var pct = requireFiniteNumber_(value, "DCA mark (%)", 0, 100);
    mtSheet.getRange(mtRow, col_(map, "DCA mark (%)")).setValue(pct);
    logAutomation_("DCA mark (%) entered", moduleCode, "Saved", pct + "%");
    return ok_({ row: mtRow, pct: pct });
  } catch (e) { return fail_(e); } finally { lock.releaseLock(); }
}

/** Deletes one AF Components entry by its stable AF Item ID — clears every
 *  column on that exact row back to blank (never touches any other row), so
 *  the slot is reusable and no other assessment's history is affected. This
 *  is the "undo a mistaken mark" path for AF items; A1/A2/A3 use
 *  api_clearMarksTrackerPct instead, since those live in a single cell rather
 *  than a row. */
function api_deleteAfEntry(afItemId) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(LOCK_WAIT_MS)) return fail_(new Error("Workbook is busy — try again in a moment."));
    if (!afItemId) return fail_(new Error("No AF Item ID given."));
    var sheet = SpreadsheetApp.getActive().getSheetByName(AF_COMPONENTS_SHEET);
    var map = getColMap_(sheet);
    var lastRow = sheet.getLastRow();
    var idCol = col_(map, "AF Item ID");
    var ids = sheet.getRange(HEADER_ROW + 1, idCol, Math.max(lastRow - HEADER_ROW, 0), 1).getValues();
    var targetRow = -1;
    for (var i = 0; i < ids.length; i++) { if (ids[i][0] === afItemId) { targetRow = HEADER_ROW + 1 + i; break; } }
    if (targetRow === -1) return fail_(new Error("Could not find AF entry " + afItemId + " — it may already have been deleted."));
    // v1.2.0 -- capture a human label ("Module — Item name") before clearing
    // the row, purely so the activity log reads as something recognisable
    // instead of a bare AF Item ID (Feature: human-readable activity
    // history). The ID stays in the detail text.
    var humanLabel = afItemId;
    if (map["Module"] && map["Item name"]) {
      var modVal = sheet.getRange(targetRow, col_(map, "Module")).getValue();
      var itemVal = sheet.getRange(targetRow, col_(map, "Item name")).getValue();
      if (modVal || itemVal) humanLabel = (modVal || "") + (itemVal ? " — " + itemVal : "");
    }
    // Clear every data column on this row only. Percentage (a formula) clears
    // itself once Mark/Maximum are blank, so it is not set directly.
    ["AF Item ID", "Module", "Item number", "Item name", "Date", "Mark", "Maximum mark",
      "Written", "Excused or excluded", "Included in AF", "Notes"].forEach(function (h) {
        if (map[h]) sheet.getRange(targetRow, map[h]).clearContent();
      });
    logAutomation_("AF entry deleted", humanLabel, "Deleted", "ID " + afItemId + "; row " + targetRow);
    return ok_({});
  } catch (e) { return fail_(e); } finally { lock.releaseLock(); }
}

// ============================================================
// MUTATIONS — Assignment / Assessment / Study Session deletion (v1.1.12 master)
// Same find-row-by-stable-ID pattern as api_deleteAfEntry/api_deleteInboxItem
// above (never by row position). Each one accepts an explicit
// deleteCalendarEvent boolean the frontend gets from the user via a two-choice
// modal -- Calendar is NEVER touched unless the caller asked for it, and if a
// requested Calendar deletion fails, the workbook row is NOT deleted either,
// so nothing is ever silently left out of sync between the two.
// ============================================================
// Core logic with no locking of its own -- deleteEntityRow_ below wraps it
// with a single LockService acquisition for a normal single-record delete;
// api_deleteInboxItemAndFiled calls it directly under ITS OWN outer lock
// instead, so one execution never nests two independent script-lock
// acquisitions.
function deleteEntityRowCore_(sheetName, idHeader, entityId, deleteCalendarEvent, logLabel) {
  if (!entityId) return fail_(new Error("No " + idHeader + " given."));
  var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sheet) return fail_(new Error("Sheet not found: " + sheetName));
  var map = getColMap_(sheet);
  var lastRow = sheet.getLastRow();
  var idCol = col_(map, idHeader);
  var ids = sheet.getRange(HEADER_ROW + 1, idCol, Math.max(lastRow - HEADER_ROW, 0), 1).getValues();
  var targetRow = -1;
  for (var i = 0; i < ids.length; i++) { if (ids[i][0] === entityId) { targetRow = HEADER_ROW + 1 + i; break; } }
  if (targetRow === -1) return fail_(new Error("Could not find " + logLabel + " " + entityId + " — it may already have been deleted."));

  // v1.2.0 -- capture a human label before deleting, so the activity log
  // reads as a title (Assignments/Assessments) or "Module — Topic/Study
  // type" (Study Planner, which has no Title column) instead of a bare
  // stable ID (Feature: human-readable activity history). Purely a
  // logging-text change -- the delete logic and return contract below are
  // unchanged.
  var humanLabel = entityId;
  if (map["Title"]) {
    var titleVal = sheet.getRange(targetRow, col_(map, "Title")).getValue();
    if (titleVal) humanLabel = titleVal;
  } else if (map["Module"]) {
    var modVal = sheet.getRange(targetRow, col_(map, "Module")).getValue();
    var topicVal = map["Topic"] ? sheet.getRange(targetRow, col_(map, "Topic")).getValue()
      : (map["Study type"] ? sheet.getRange(targetRow, col_(map, "Study type")).getValue() : "");
    humanLabel = topicVal ? (modVal + " — " + topicVal) : modVal;
  }

  var calendarEventId = map["Calendar Event ID"] ? sheet.getRange(targetRow, col_(map, "Calendar Event ID")).getValue() : "";
  var calendarResult = "Not requested";
  if (calendarEventId) {
    if (deleteCalendarEvent) {
      var delResult = deleteCalendarEventByStoredId_(calendarEventId);
      if (!delResult.ok) {
        // Do not pretend success and do not delete the workbook row either
        // -- leaving both sides intact (row + stale-but-real Calendar
        // event) is recoverable; deleting the row while the event survives
        // untracked would not be.
        return fail_(new Error("Calendar event could not be deleted (" + delResult.error + "). The " + logLabel + " row was NOT deleted either — try again."));
      }
      calendarResult = delResult.alreadyGone ? "Already gone from Calendar" : "Deleted";
    } else {
      calendarResult = "Left in Calendar (event " + calendarEventId + ")";
    }
  }

  sheet.deleteRow(targetRow);
  logAutomation_(logLabel + " deleted", humanLabel, "Deleted", "ID " + entityId + "; row " + targetRow + "; calendar: " + calendarResult);
  return ok_({ id: entityId, calendarEventId: calendarEventId || null, calendarResult: calendarResult });
}
function deleteEntityRow_(sheetName, idHeader, entityId, deleteCalendarEvent, logLabel) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(LOCK_WAIT_MS)) return fail_(new Error("Workbook is busy — try again in a moment."));
    return deleteEntityRowCore_(sheetName, idHeader, entityId, deleteCalendarEvent, logLabel);
  } catch (e) { return fail_(e); } finally { lock.releaseLock(); }
}

function api_deleteAssignment(assignmentId, deleteCalendarEvent) {
  return deleteEntityRow_(ASSIGNMENTS_SHEET, "Assignment ID", assignmentId, !!deleteCalendarEvent, "Assignment");
}

function api_deleteAssessment(assessmentId, deleteCalendarEvent) {
  return deleteEntityRow_(ASSESSMENTS_SHEET, "Assessment ID", assessmentId, !!deleteCalendarEvent, "Assessment");
}

function api_deleteStudySession(studySessionId, deleteCalendarEvent) {
  // v1.2.0 RC2 -- correction #7: reverse this session's Revision Tracker
  // contribution (if any) BEFORE the row disappears, so the aggregate never
  // retains a dangling contribution nothing can find or edit afterward. If
  // reversing this drops a topic's aggregates to zero, the topic row itself
  // is left in place -- reverseRevisionContribution_ only ever decrements
  // fields on the existing row, it never deletes a Revision Tracker row.
  reverseSessionRevisionContributionById_(studySessionId);
  var result = deleteEntityRow_(STUDY_SHEET, "Study Session ID", studySessionId, !!deleteCalendarEvent, "Study session");
  // v1.2.0 -- cascade to this session's own Study Tasks only (never any
  // other session's tasks). Best-effort and only attempted after a
  // successful session delete; a failure here does not roll back the
  // already-completed delete above, and an orphaned task row is harmless
  // (every screen groups tasks by session ID, and the session is gone).
  if (result && result.ok) {
    deleteStudyTasksForSession_(studySessionId);
  }
  return result;
}

function deleteStudyTasksForSession_(sessionId) {
  try {
    var sheet = SpreadsheetApp.getActive().getSheetByName(STUDY_TASKS_SHEET);
    if (!sheet) return;
    var map = getColMap_(sheet);
    if (!map["Study Session ID"]) return;
    var lastRow = sheet.getLastRow();
    if (lastRow <= HEADER_ROW) return;
    var ids = sheet.getRange(HEADER_ROW + 1, col_(map, "Study Session ID"), lastRow - HEADER_ROW, 1).getValues();
    for (var i = ids.length - 1; i >= 0; i--) {
      if (ids[i][0] === sessionId) sheet.deleteRow(HEADER_ROW + 1 + i);
    }
  } catch (e) { /* best-effort cleanup only -- never blocks the session delete result above */ }
}

// ============================================================
// MUTATIONS — generic checkbox toggles (Study Planner, Resources, Revision, Attendance)
// ============================================================
function withRowByKey_(sheetName, keyHeader, keyValue, fn) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(LOCK_WAIT_MS)) return fail_(new Error("Workbook is busy — try again in a moment."));
    var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    if (!sheet) return fail_(new Error("Sheet not found: " + sheetName));
    var map = getColMap_(sheet);
    var keyCol = col_(map, keyHeader);
    var lastRow = sheet.getLastRow();
    var keys = sheet.getRange(HEADER_ROW + 1, keyCol, Math.max(lastRow - HEADER_ROW, 0), 1).getValues();
    var row = -1;
    for (var i = 0; i < keys.length; i++) { if (keys[i][0] === keyValue) { row = HEADER_ROW + 1 + i; break; } }
    if (row === -1) return fail_(new Error("Row not found for " + keyHeader + " = " + keyValue));
    fn(sheet, map, row);
    return ok_({ row: row });
  } catch (e) { return fail_(e); } finally { lock.releaseLock(); }
}

function api_toggleStudyDone(sessionId, completed) {
  return withRowByKey_(STUDY_SHEET, "Study Session ID", sessionId, function (sheet, map, row) {
    sheet.getRange(row, col_(map, "Completed")).setValue(!!completed);
  });
}

function api_toggleResourceReviewed(module, title, reviewed) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(LOCK_WAIT_MS)) return fail_(new Error("Workbook is busy — try again in a moment."));
    var sheet = SpreadsheetApp.getActive().getSheetByName(RESOURCES_SHEET);
    var map = getColMap_(sheet);
    var lastRow = sheet.getLastRow();
    var vals = sheet.getRange(HEADER_ROW + 1, 1, Math.max(lastRow - HEADER_ROW, 0), sheet.getLastColumn()).getValues();
    var mCol = col_(map, "Module") - 1, tCol = col_(map, "Title") - 1;
    for (var i = 0; i < vals.length; i++) {
      if (vals[i][mCol] === module && vals[i][tCol] === title) {
        sheet.getRange(HEADER_ROW + 1 + i, col_(map, "Reviewed")).setValue(!!reviewed);
        return ok_({});
      }
    }
    return fail_(new Error("Resource not found."));
  } catch (e) { return fail_(e); } finally { lock.releaseLock(); }
}

// ============================================================
// MUTATIONS — Notes (quick-capture block, debounced from the client)
// ============================================================
function api_saveNotes(field, value) {
  try {
    var sheet = SpreadsheetApp.getActive().getSheetByName(NOTES_SHEET);
    var cellByField = { today: "C61", parking: "C62", questions: "C63", remember: "C64" };
    if (!cellByField[field]) return fail_(new Error("Unknown notes field: " + field));
    sheet.getRange(cellByField[field]).setValue(cleanMultiline_(value, 4000));
    return ok_({});
  } catch (e) { return fail_(e); }
}

// ============================================================
// MUTATIONS — Settings
// ============================================================
function api_saveSetting(label, value) {
  try {
    var settings = SpreadsheetApp.getActive().getSheetByName(SETTINGS_SHEET);
    var finder = settings.createTextFinder(label).matchEntireCell(false);
    var cell = finder.findNext();
    if (!cell) return fail_(new Error("Setting not found: " + label));
    settings.getRange(cell.getRow(), 3).setValue(value);
    logAutomation_("Setting changed", label, "Updated", String(value));
    return ok_({});
  } catch (e) { return fail_(e); }
}

function api_setManualAf(moduleCode, value) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(LOCK_WAIT_MS)) return fail_(new Error("Workbook is busy — try again in a moment."));
    var sheet = SpreadsheetApp.getActive().getSheetByName(MARKS_TRACKER_SHEET);
    var lastRow = sheet.getLastRow();
    var codeCol = 1, targetCol = 19; // column S — "Manual combined AF entry (%) — reference only"
    var cleanVal = "";
    if (value !== "" && value !== null && typeof value !== "undefined") {
      cleanVal = requireFiniteNumber_(value, "Manual AF entry", 0, 100);
    }
    var vals = sheet.getRange(MARKS_TRACKER_HEADER_ROW + 1, codeCol, Math.max(lastRow - MARKS_TRACKER_HEADER_ROW, 0), 1).getValues();
    for (var i = 0; i < vals.length; i++) {
      if (vals[i][0] === moduleCode) {
        sheet.getRange(MARKS_TRACKER_HEADER_ROW + 1 + i, targetCol).setValue(cleanVal);
        return ok_({});
      }
    }
    return fail_(new Error("Module not found in Marks Tracker: " + moduleCode));
  } catch (e) { return fail_(e); } finally { lock.releaseLock(); }
}

// ============================================================
// AUTOMATION ACTIONS  (wrap the real Calendar_Sync functions, log the outcome)
// ============================================================
function api_processInbox() {
  try {
    var before = readSheetRows_(INBOX_SHEET, HEADER_ROW, "Title").filter(function (r) { return !r["Processed"]; }).length;
    processAcademicInbox();
    var after = readSheetRows_(INBOX_SHEET, HEADER_ROW, "Title").filter(function (r) { return !r["Processed"]; }).length;
    var count = Math.max(0, before - after);
    logAutomation_("Process Academic Inbox", "-", "Success", count + " item(s) processed");
    return ok_({ processed: count });
  } catch (e) { return fail_(e); }
}

function api_syncCalendar() {
  try {
    syncAllToCalendar();
    logAutomation_("Sync All Calendar Items", "-", "Success", "Sync completed");
    return ok_({});
  } catch (e) { return fail_(e); }
}

// ============================================================
// SEMESTER LIFECYCLE — web-app-safe reimplementations of archiveSemester() /
// duplicateSemester() from Calendar_Sync_v2.0.gs.
//
// IMPORTANT: the originals call SpreadsheetApp.getUi() to show confirmation
// dialogs and prompts. getUi() only works when Apps Script is invoked from
// within the Sheets UI itself (a custom menu click) — it throws when called
// from a deployed web app / google.script.run context, which is exactly how
// the web app's "Archive Semester" button would call it. Wiring the button
// straight to the existing function would look fine in review and then throw
// on the very first real click — the kind of defect that only shows up at
// point of use, which is why this was caught here rather than shipped.
//
// The fix is not to touch Calendar_Sync_v2.0.gs (menu-driven use keeps
// working exactly as before) but to give the web app its own safe entry
// points: confirmation happens client-side (Component.js), and "new semester
// name/dates" reuses the Settings screen the app already has, instead of a
// second prompt() flow.
// ============================================================
function api_archiveSemester() {
  try {
    var ss = SpreadsheetApp.getActive();
    var mt = ss.getSheetByName(MARKS_TRACKER_SHEET);
    var archive = ss.getSheetByName(ARCHIVE_SHEET);
    var semesterName = getSetting_("Semester name") || "Unknown semester";
    var mtMap = getColMap_(mt, MARKS_TRACKER_HEADER_ROW);
    var lastRow = mt.getLastRow();
    var archiveRow = Math.max(archive.getLastRow() + 1, 13); // below 17 Archive's header block (row 12)
    var count = 0;
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
      archiveRow++; count++;
    }
    var sheetsToClear = [ASSIGNMENTS_SHEET, ASSESSMENTS_SHEET, AF_COMPONENTS_SHEET, STUDY_SHEET, INBOX_SHEET, ATTENDANCE_SHEET];
    sheetsToClear.forEach(function (name) {
      var sheet = ss.getSheetByName(name);
      if (!sheet) return;
      var lr = sheet.getLastRow(), lc = sheet.getLastColumn();
      if (lr > HEADER_ROW) sheet.getRange(HEADER_ROW + 1, 1, lr - HEADER_ROW, lc).clearContent();
    });
    logAutomation_("Archive Semester", semesterName, "Success", count + " module row(s) archived, working sheets cleared");
    return ok_({ archived: count });
  } catch (e) { return fail_(e); }
}

function api_duplicateSemester(newName) {
  try {
    var name = requireText_(newName, "New file name");
    var file = DriveApp.getFileById(SpreadsheetApp.getActive().getId());
    var copy = file.makeCopy(name);
    logAutomation_("Duplicate Semester", name, "Success", copy.getUrl());
    return ok_({ url: copy.getUrl() });
  } catch (e) { return fail_(e); }
}

function api_toggleAutomation(enable) {
  try {
    if (enable) { enableAutomation(); logAutomation_("Automation", "-", "Enabled", "Triggers created"); }
    else { disableAutomation(); logAutomation_("Automation", "-", "Disabled", "Triggers removed"); }
    return ok_({});
  } catch (e) { return fail_(e); }
}

// ============================================================
// USER GUIDE — read straight from 18 User Guide so the web UI never
// duplicates guide text that could drift out of sync with the sheet.
// ============================================================
function readUserGuide_() {
  var sheet = SpreadsheetApp.getActive().getSheetByName("18 User Guide");
  if (!sheet) return [];
  var entries = [];
  var lastRow = sheet.getLastRow();
  var r = HEADER_ROW;
  while (r <= lastRow) {
    var title = sheet.getRange(r, 2).getValue();
    if (title) {
      var body = sheet.getRange(r + 1, 2).getValue();
      entries.push({ title: title, body: body || "" });
      r += 3;
    } else {
      r += 1;
    }
  }
  return entries;
}

// ============================================================
// SEMESTER PHASE  — boundaries mirror the exact IF() logic already in
// 01 Dashboard ("Current academic week" formula) so the phase shown in
// the UI can never disagree with the Dashboard sheet.
// ============================================================
function currentSemesterPhaseIndex_() {
  var tz = TIMEZONE;
  var today = new Date();
  var boundaries = [
    new Date(2026, 6, 20), new Date(2026, 7, 9),   // 0 Foundation & Semester Start (20 Jul - 9 Aug)
    new Date(2026, 7, 10), new Date(2026, 7, 30),  // 1 A1 Build-up (10-30 Aug)
    new Date(2026, 7, 31), new Date(2026, 8, 6),   // 2 A1 Final Week (31 Aug - 6 Sep) [sheet: A1 Assessment Week ends Sep 6]
    new Date(2026, 8, 7), new Date(2026, 9, 11),   // 3 Teaching Block 2 & Recovery (7 Sep - 11 Oct) [covers recess 7-13 Sep + teaching 14 Sep-23 Oct + pre-A2 24-28 Oct, collapsed to match phase list below]
    new Date(2026, 9, 12), new Date(2026, 10, 1),  // 4 A2 Build-up (12 Oct - 1 Nov)
    new Date(2026, 10, 2), new Date(2026, 10, 19), // 5 A2 Test Window (2-19 Nov)
    new Date(2026, 10, 20), new Date(2026, 11, 2)  // 6 A3 Test Window (20 Nov - 2 Dec)
  ];
  for (var i = 0; i < boundaries.length; i += 2) {
    if (today >= boundaries[i] && today <= new Date(boundaries[i + 1].getFullYear(), boundaries[i + 1].getMonth(), boundaries[i + 1].getDate(), 23, 59, 59)) {
      return i / 2;
    }
  }
  if (today < boundaries[0]) return -1; // before semester
  return 6; // after last boundary — treat as final phase (A3 window) until dates are updated for next semester
}
function api_getPlannerDataJson() {
  try {
    return JSON.stringify(api_getPlannerData());
  } catch (error) {
    return JSON.stringify({
      ok: false,
      error: error && error.message
        ? error.message
        : String(error)
    });
  }
}

// ============================================================================================
// v1.2.0 ACADEMIC WORKFLOW RELEASE — everything below this line is new.
// Nothing above this line was removed, renamed, or had its return contract
// changed; the only in-place edits above are the small logAutomation_()
// call-site improvements marked "v1.2.0 --" (human-readable activity
// history), which do not alter behaviour, signatures, or return shapes.
// ============================================================================================

var VALID_STUDY_TYPES = ["Preview", "Learn", "Practice", "Mistake", "Recall", "Formula", "Visualization", "Programming", "Synthesis", "Exam", "Weekly Review"];
var VALID_SESSION_STATUSES = ["Planned", "Completed", "Partly completed", "Skipped", "Rescheduled"];
var VALID_TASK_TYPES = ["Watch lecture", "Review notes", "Learn concept", "Complete tutorial", "Practice questions",
  "Attempt past paper", "Mark work", "Review mistakes", "Redo incorrect questions", "Formula review",
  "Flashcards", "Coding practice", "Weekly review", "Custom"];
var VALID_RESOURCE_TYPES = ["Lecture", "Lecture video", "Tutorial", "Past paper", "Formula bible", "Textbook chapter",
  "Practice set", "Flashcards", "Summary notes", "Mistake review", "Other"];
var VALID_RESOURCE_STATUSES = ["Not started", "In progress", "Completed", "Reviewed", "Needs redo"];
var VALID_ATTENDANCE_STATUSES = ["Not recorded", "Attended", "Missed"];

// ------------------------------------------------------------------
// STUDY PLANNER — create / edit / reschedule / complete (Feature: Task-based
// Study Planner, §4.1–§4.2)
// ------------------------------------------------------------------
function api_addStudySession(form) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(LOCK_WAIT_MS)) return fail_(new Error("Workbook is busy — try again in a moment."));
    if (!form) return fail_(new Error("No form data received."));
    var module = requireModule_(form.module);
    var dateVal = parseValidDate_(form.date, "Date");
    if (!dateVal) throw new Error("Date is required.");
    var studyType = VALID_STUDY_TYPES.indexOf(form.studyType) !== -1 ? form.studyType : "Practice";
    var priority = VALID_PRIORITIES.indexOf(form.priority) !== -1 ? form.priority : "Medium";
    var topic = cleanText_(form.topic, 200);
    var objective = cleanText_(form.objective, 300);
    var notes = cleanMultiline_(form.notes, 1000);
    var duration = (form.duration !== "" && form.duration != null) ? requireFiniteNumber_(form.duration, "Duration", 5, 600) : null;
    var startTime = form.startTime ? parseTimeOfDay_(form.startTime, dateVal) : null;
    var endTime = null;
    if (startTime && duration) endTime = new Date(startTime.getTime() + duration * 60000);
    else if (startTime && form.endTime) endTime = parseTimeOfDay_(form.endTime, dateVal);

    var result = appendStudySessionRow_({
      module: module, date: dateVal, startTime: startTime, endTime: endTime, duration: duration,
      studyType: studyType, topic: topic, priority: priority, objective: objective, notes: notes,
      planned: true, completed: false, status: "Planned", origin: "Manual",
      syncToCalendar: !!form.syncToCalendar
    });
    logAutomation_("Study session added", module + " — " + (topic || studyType), "Created", result.id);
    return ok_({ id: result.id });
  } catch (e) { return fail_(e); } finally { lock.releaseLock(); }
}

function api_updateStudySession(sessionId, form) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(LOCK_WAIT_MS)) return fail_(new Error("Workbook is busy — try again in a moment."));
    if (!form) return fail_(new Error("No form data received."));
    var row = findRowByStableId_(STUDY_SHEET, "Study Session ID", sessionId);
    if (row === -1) return fail_(new Error("Could not find study session " + sessionId + " — it may have been deleted."));
    var sheet = SpreadsheetApp.getActive().getSheetByName(STUDY_SHEET);
    var map = getColMap_(sheet);

    if (form.module) sheet.getRange(row, col_(map, "Module")).setValue(requireModule_(form.module));
    if (form.date) {
      var d = parseValidDate_(form.date, "Date");
      if (d) sheet.getRange(row, col_(map, "Date")).setValue(d);
    }
    var dateVal = sheet.getRange(row, col_(map, "Date")).getValue();
    if (form.startTime) sheet.getRange(row, col_(map, "Start time")).setValue(parseTimeOfDay_(form.startTime, dateVal));
    if (form.duration !== undefined && form.duration !== "") {
      var dur = requireFiniteNumber_(form.duration, "Duration", 5, 600);
      sheet.getRange(row, col_(map, "Duration (min)")).setValue(dur);
      var st = sheet.getRange(row, col_(map, "Start time")).getValue();
      if (st instanceof Date) sheet.getRange(row, col_(map, "End time")).setValue(new Date(st.getTime() + dur * 60000));
    }
    if (form.studyType) sheet.getRange(row, col_(map, "Study type")).setValue(VALID_STUDY_TYPES.indexOf(form.studyType) !== -1 ? form.studyType : "Practice");
    if (form.topic !== undefined) sheet.getRange(row, col_(map, "Topic")).setValue(cleanText_(form.topic, 200));
    if (form.priority) sheet.getRange(row, col_(map, "Priority")).setValue(VALID_PRIORITIES.indexOf(form.priority) !== -1 ? form.priority : "Medium");
    if (form.objective !== undefined && map["Objective"]) sheet.getRange(row, col_(map, "Objective")).setValue(cleanText_(form.objective, 300));
    if (form.notes !== undefined) sheet.getRange(row, col_(map, "Notes")).setValue(cleanMultiline_(form.notes, 1000));
    if (typeof form.syncToCalendar === "boolean" && map["Sync to Calendar"]) sheet.getRange(row, col_(map, "Sync to Calendar")).setValue(form.syncToCalendar);

    // v1.2.0 RC2 -- correction #1: if this session already has an applied
    // Revision Tracker contribution (i.e. it was previously completed) and
    // editing this form just changed its module and/or topic, move that
    // exact contribution across to the new module/topic rather than leaving
    // it stranded against the old one -- reconcileSessionRevisionContribution_
    // is only ever called from the completion/independent-study flows, so an
    // edit-only module/topic change needs its own equivalent handling here.
    // This only ever relocates the already-recorded minutes/session-count/
    // source -- it never adds a new contribution or changes the total.
    if (form.module !== undefined || form.topic !== undefined) {
      var existingContribution = readStoredContribution_(sheet, map, row);
      if (existingContribution.applied) {
        var newModuleName = sheet.getRange(row, col_(map, "Module")).getValue();
        var newTopic = map["Topic"] ? sheet.getRange(row, col_(map, "Topic")).getValue() : "";
        if (newModuleName !== existingContribution.module || newTopic !== existingContribution.topic) {
          reverseRevisionContribution_(existingContribution);
          if (newModuleName && newTopic) {
            updateRevisionFromStudy_(newModuleName, newTopic, { minutes: existingContribution.minutes, isIndependent: existingContribution.source === "Independent" });
            writeStoredContribution_(sheet, map, row, { applied: true, module: newModuleName, topic: newTopic, minutes: existingContribution.minutes, source: existingContribution.source });
          } else {
            writeStoredContribution_(sheet, map, row, { applied: false });
          }
        }
      }
    }

    logAutomation_("Study session updated", sessionId, "Updated", "");
    return ok_({});
  } catch (e) { return fail_(e); } finally { lock.releaseLock(); }
}

/** Distinct, named "reschedule" action (per spec §4.1), reusing the same
 *  row-lookup as api_updateStudySession rather than duplicating it — moves
 *  the date/time and stamps Status = "Rescheduled". */
function api_rescheduleStudySession(sessionId, newDateIso, newStartTime) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(LOCK_WAIT_MS)) return fail_(new Error("Workbook is busy — try again in a moment."));
    var row = findRowByStableId_(STUDY_SHEET, "Study Session ID", sessionId);
    if (row === -1) return fail_(new Error("Could not find study session " + sessionId + " — it may have been deleted."));
    var d = parseValidDate_(newDateIso, "Date");
    if (!d) throw new Error("A valid new date is required.");
    var sheet = SpreadsheetApp.getActive().getSheetByName(STUDY_SHEET);
    var map = getColMap_(sheet);
    sheet.getRange(row, col_(map, "Date")).setValue(d);
    if (newStartTime) sheet.getRange(row, col_(map, "Start time")).setValue(parseTimeOfDay_(newStartTime, d));
    if (map["Status"]) sheet.getRange(row, col_(map, "Status")).setValue("Rescheduled");
    logAutomation_("Study session rescheduled", sessionId, "Updated", newDateIso);
    return ok_({});
  } catch (e) { return fail_(e); } finally { lock.releaseLock(); }
}

/**
 * Closes out a session: short form only (status, actual duration,
 * confidence after, optional note — per spec §4.4, never a long form).
 * When the session is Completed or Partly completed AND has a Module+Topic,
 * this also rolls the actual study time into the matching Revision Tracker
 * row (creating it if this is the very first time that Module+Topic has
 * been studied) via updateRevisionFromStudy_ in Calendar_Sync_v1.2.0.gs --
 * never inferring or overwriting Confidence, which stays a manual field.
 */
function api_completeStudySession(sessionId, form) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(LOCK_WAIT_MS)) return fail_(new Error("Workbook is busy — try again in a moment."));
    if (!form) return fail_(new Error("No form data received."));
    var status = VALID_SESSION_STATUSES.indexOf(form.status) !== -1 ? form.status : "Completed";
    var row = findRowByStableId_(STUDY_SHEET, "Study Session ID", sessionId);
    if (row === -1) return fail_(new Error("Could not find study session " + sessionId + " — it may have been deleted."));
    var sheet = SpreadsheetApp.getActive().getSheetByName(STUDY_SHEET);
    var map = getColMap_(sheet);

    var actualDuration = (form.actualDuration !== "" && form.actualDuration != null) ? requireFiniteNumber_(form.actualDuration, "Actual duration", 0, 600) : null;
    var confidenceAfter = (form.confidenceAfter !== "" && form.confidenceAfter != null) ? requireFiniteNumber_(form.confidenceAfter, "Confidence", 0, 100) : null;
    var note = cleanMultiline_(form.note, 1000);

    if (map["Status"]) sheet.getRange(row, col_(map, "Status")).setValue(status);
    sheet.getRange(row, col_(map, "Completed")).setValue(status === "Completed" || status === "Partly completed");
    if (actualDuration != null && map["Actual duration (min)"]) sheet.getRange(row, col_(map, "Actual duration (min)")).setValue(actualDuration);
    if (confidenceAfter != null && map["Confidence after"]) sheet.getRange(row, col_(map, "Confidence after")).setValue(confidenceAfter);
    if (note && map["Notes"]) {
      var existingNotes = sheet.getRange(row, col_(map, "Notes")).getValue();
      sheet.getRange(row, col_(map, "Notes")).setValue(existingNotes ? existingNotes + "\n" + note : note);
    }

    var moduleName = sheet.getRange(row, col_(map, "Module")).getValue();
    var topic = map["Topic"] ? sheet.getRange(row, col_(map, "Topic")).getValue() : "";
    var origin = map["Origin"] ? sheet.getRange(row, col_(map, "Origin")).getValue() : "Manual";

    // v1.2.0 RC2 -- idempotent aggregation (correction #1). Every save of
    // this form -- first completion, resubmission with no changes, editing
    // the actual duration afterward, flipping between Completed and Partly
    // completed, or changing the module/topic before re-saving -- goes
    // through reconcileSessionRevisionContribution_, which always reverses
    // this exact session's PREVIOUSLY stored contribution (if any, against
    // whatever module/topic it was previously stored against) before
    // applying its current one. That guarantees this session contributes to
    // Revision Tracker exactly once, at its current values, no matter how
    // many times this form is saved.
    var minutes = actualDuration != null ? actualDuration : (num_(sheet.getRange(row, col_(map, "Duration (min)")).getValue()) || 0);
    var shouldContribute = (status === "Completed" || status === "Partly completed");
    var revisionResult = reconcileSessionRevisionContribution_(sheet, map, row, moduleName, topic, {
      shouldContribute: shouldContribute, minutes: minutes, isIndependent: (origin === "Independent"), confidence: confidenceAfter
    });

    logAutomation_("Study session " + status.toLowerCase(), moduleName + " — " + (topic || "session"), "Updated", sessionId);
    return ok_({ id: sessionId, status: status, revisionUpdated: revisionResult.updated });
  } catch (e) { return fail_(e); } finally { lock.releaseLock(); }
}

/**
 * "Log independent study" (spec §4.5) — creates a new Study Planner row
 * labelled Origin = "Independent" (so the frontend can always distinguish
 * it from a planned session) and immediately rolls it into Revision
 * Tracker, exactly like completing a planned session does.
 */
function api_logIndependentStudy(form) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(LOCK_WAIT_MS)) return fail_(new Error("Workbook is busy — try again in a moment."));
    if (!form) return fail_(new Error("No form data received."));
    var module = requireModule_(form.module);
    var topic = requireText_(form.topic, "Topic", 200);
    var dateVal = form.date ? parseValidDate_(form.date, "Date") : new Date();
    var duration = requireFiniteNumber_(form.duration, "Duration", 1, 600);
    var studyType = VALID_STUDY_TYPES.indexOf(form.taskType) !== -1 ? form.taskType : "Practice";
    var score = (form.score !== "" && form.score != null) ? requireFiniteNumber_(form.score, "Score", 0, 100) : null;
    var confidence = (form.confidence !== "" && form.confidence != null) ? requireFiniteNumber_(form.confidence, "Confidence", 0, 100) : null;
    var activityName = cleanText_(form.activityName, 200);
    var notes = cleanMultiline_(form.notes, 1000);
    var combinedNotes = activityName ? ("Activity/resource: " + activityName + (notes ? "\n" + notes : "")) : notes;

    var result = appendStudySessionRow_({
      module: module, date: dateVal, duration: duration, studyType: studyType, topic: topic,
      priority: "Medium", notes: combinedNotes, planned: false, completed: true,
      status: "Completed", origin: "Independent"
    });

    // v1.2.0 RC2 -- route through the same reconcile helper api_completeStudySession
    // uses (rather than calling updateRevisionFromStudy_ directly) so this
    // brand-new session's contribution is also recorded in its own ledger
    // columns on "12 Study Planner". That makes it just as safe to edit or
    // delete later as a planned session that was completed through the
    // normal flow -- nothing here is exempt from the idempotency guarantee.
    var newRow = findRowByStableId_(STUDY_SHEET, "Study Session ID", result.id);
    var revisionResult = { updated: false };
    if (newRow !== -1) {
      var sheet = SpreadsheetApp.getActive().getSheetByName(STUDY_SHEET);
      var map = getColMap_(sheet);
      revisionResult = reconcileSessionRevisionContribution_(sheet, map, newRow, module, topic, {
        shouldContribute: true, minutes: duration, isIndependent: true, confidence: confidence, score: score
      });
    }

    logAutomation_("Independent study logged", module + " — " + topic, "Created", result.id);
    return ok_({ id: result.id, revisionUpdated: revisionResult.updated });
  } catch (e) { return fail_(e); } finally { lock.releaseLock(); }
}

// ------------------------------------------------------------------
// STUDY TASKS — open-ended checklist per session (Feature: Task-based Study
// Planner, §4.3), stored in the new "21 Study Tasks" sheet.
// ------------------------------------------------------------------
function api_addStudyTask(sessionId, form) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(LOCK_WAIT_MS)) return fail_(new Error("Workbook is busy — try again in a moment."));
    if (!form) return fail_(new Error("No form data received."));
    var sessionRow = findRowByStableId_(STUDY_SHEET, "Study Session ID", sessionId);
    if (sessionRow === -1) return fail_(new Error("Study session " + sessionId + " not found."));
    var studySheet = SpreadsheetApp.getActive().getSheetByName(STUDY_SHEET);
    var studyMap = getColMap_(studySheet);
    var module = studySheet.getRange(sessionRow, col_(studyMap, "Module")).getValue();
    var topic = studyMap["Topic"] ? studySheet.getRange(sessionRow, col_(studyMap, "Topic")).getValue() : "";

    var taskText = requireText_(form.text, "Task text", 300);
    var taskType = VALID_TASK_TYPES.indexOf(form.taskType) !== -1 ? form.taskType : "Custom";
    var resourceId = cleanText_(form.resourceId, 40);

    var sheet = SpreadsheetApp.getActive().getSheetByName(STUDY_TASKS_SHEET);
    if (!sheet) return fail_(new Error('Sheet not found: "' + STUDY_TASKS_SHEET + '" — run Study_Tasks_v1.2.0_Migration.gs first.'));
    var map = getColMap_(sheet);
    var row = firstBlankRow_(sheet, col_(map, "Study Session ID"));
    var newId = nextId_("TSK", getColumnValues_(sheet, col_(map, "Task ID")));
    var existingCount = readSheetRows_(STUDY_TASKS_SHEET, HEADER_ROW, "Study Session ID")
      .filter(function (t) { return t["Study Session ID"] === sessionId; }).length;

    sheet.getRange(row, col_(map, "Task ID")).setValue(newId);
    sheet.getRange(row, col_(map, "Study Session ID")).setValue(sessionId);
    sheet.getRange(row, col_(map, "Module")).setValue(module);
    if (topic) sheet.getRange(row, col_(map, "Topic")).setValue(topic);
    sheet.getRange(row, col_(map, "Task type")).setValue(taskType);
    sheet.getRange(row, col_(map, "Task text")).setValue(taskText);
    if (resourceId) sheet.getRange(row, col_(map, "Resource ID")).setValue(resourceId);
    sheet.getRange(row, col_(map, "Completed")).setValue(false);
    sheet.getRange(row, col_(map, "Sort order")).setValue(existingCount + 1);
    sheet.getRange(row, col_(map, "Last updated")).setValue(new Date());
    logAutomation_("Study task added", module + " — " + taskText, "Created", newId);
    return ok_({ id: newId });
  } catch (e) { return fail_(e); } finally { lock.releaseLock(); }
}

function api_updateStudyTask(taskId, form) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(LOCK_WAIT_MS)) return fail_(new Error("Workbook is busy — try again in a moment."));
    if (!form) return fail_(new Error("No form data received."));
    var row = findRowByStableId_(STUDY_TASKS_SHEET, "Task ID", taskId);
    if (row === -1) return fail_(new Error("Could not find study task " + taskId + " — it may already have been deleted."));
    var sheet = SpreadsheetApp.getActive().getSheetByName(STUDY_TASKS_SHEET);
    var map = getColMap_(sheet);
    if (form.text !== undefined) sheet.getRange(row, col_(map, "Task text")).setValue(requireText_(form.text, "Task text", 300));
    if (form.taskType) sheet.getRange(row, col_(map, "Task type")).setValue(VALID_TASK_TYPES.indexOf(form.taskType) !== -1 ? form.taskType : "Custom");
    if (form.resourceId !== undefined) sheet.getRange(row, col_(map, "Resource ID")).setValue(cleanText_(form.resourceId, 40));
    if (form.score !== undefined && form.score !== "") sheet.getRange(row, col_(map, "Score")).setValue(requireFiniteNumber_(form.score, "Score", 0, 100));
    if (form.notes !== undefined) sheet.getRange(row, col_(map, "Notes")).setValue(cleanMultiline_(form.notes, 1000));
    sheet.getRange(row, col_(map, "Last updated")).setValue(new Date());
    logAutomation_("Study task updated", taskId, "Updated", "");
    return ok_({});
  } catch (e) { return fail_(e); } finally { lock.releaseLock(); }
}

function api_toggleStudyTask(taskId, completed) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(LOCK_WAIT_MS)) return fail_(new Error("Workbook is busy — try again in a moment."));
    var row = findRowByStableId_(STUDY_TASKS_SHEET, "Task ID", taskId);
    if (row === -1) return fail_(new Error("Could not find study task " + taskId + " — it may already have been deleted."));
    var sheet = SpreadsheetApp.getActive().getSheetByName(STUDY_TASKS_SHEET);
    var map = getColMap_(sheet);
    var wasCompleted = sheet.getRange(row, col_(map, "Completed")).getValue() === true;
    var nowCompleted = !!completed;
    sheet.getRange(row, col_(map, "Completed")).setValue(nowCompleted);
    sheet.getRange(row, col_(map, "Last updated")).setValue(new Date());
    var taskText = sheet.getRange(row, col_(map, "Task text")).getValue();

    // v1.2.0 RC2 -- correction #2: only touch Revision Tracker's "Completed
    // task count" on an actual false->true or true->false transition, so a
    // toggle that doesn't actually change the stored value (e.g. resending
    // the same state) never increments or decrements twice for one real
    // change.
    if (wasCompleted !== nowCompleted) {
      var taskModule = sheet.getRange(row, col_(map, "Module")).getValue();
      var taskTopic = map["Topic"] ? sheet.getRange(row, col_(map, "Topic")).getValue() : "";
      if (taskModule && taskTopic) {
        adjustRevisionCompletedTaskCount_(taskModule, taskTopic, nowCompleted ? 1 : -1);
      }
      // No valid Module/Topic on this task -- per spec, Revision Tracker is
      // deliberately left untouched rather than guessed at.
    }

    logAutomation_("Study task " + (nowCompleted ? "completed" : "reopened"), taskText || taskId, "Updated", taskId);
    return ok_({});
  } catch (e) { return fail_(e); } finally { lock.releaseLock(); }
}

/**
 * Correction #2: adjusts "Completed task count" on the exact Module+Topic
 * Revision Tracker row by +1 or -1. Never creates a new Revision Tracker
 * row -- a task's own session (or independent-study entry) is what creates
 * that row in the first place; ticking a task can only adjust an existing
 * row, never conjure one into existence. Clamped so the count can never
 * drop below zero.
 */
function adjustRevisionCompletedTaskCount_(moduleName, topic, delta) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(REVISION_SHEET);
  if (!sheet) return;
  var map = getColMap_(sheet);
  if (!map["Completed task count"]) return;
  var lastRow = sheet.getLastRow();
  if (lastRow <= HEADER_ROW) return;
  var vals = sheet.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, sheet.getLastColumn()).getValues();
  var moduleCol = col_(map, "Module") - 1, topicCol = col_(map, "Topic") - 1;
  for (var i = 0; i < vals.length; i++) {
    if (vals[i][moduleCol] === moduleName && vals[i][topicCol] === topic) {
      var targetRow = HEADER_ROW + 1 + i;
      var current = num_(sheet.getRange(targetRow, col_(map, "Completed task count")).getValue()) || 0;
      sheet.getRange(targetRow, col_(map, "Completed task count")).setValue(Math.max(0, current + delta));
      return;
    }
  }
  // No matching Revision Tracker row exists yet -- intentionally not created here.
}

function api_deleteStudyTask(taskId) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(LOCK_WAIT_MS)) return fail_(new Error("Workbook is busy — try again in a moment."));
    var row = findRowByStableId_(STUDY_TASKS_SHEET, "Task ID", taskId);
    if (row === -1) return fail_(new Error("Could not find study task " + taskId + " — it may already have been deleted."));
    var sheet = SpreadsheetApp.getActive().getSheetByName(STUDY_TASKS_SHEET);
    var map = getColMap_(sheet);
    var taskText = sheet.getRange(row, col_(map, "Task text")).getValue();
    // v1.2.0 RC2 -- correction #2 consistency: a completed task that's about
    // to be deleted should not leave Revision Tracker's Completed task count
    // permanently inflated by a task that no longer exists. Same exact
    // Module+Topic matching, same "never below zero" clamp, same "never
    // create a new row" rule as the toggle path above.
    var wasCompleted = sheet.getRange(row, col_(map, "Completed")).getValue() === true;
    if (wasCompleted) {
      var taskModule = sheet.getRange(row, col_(map, "Module")).getValue();
      var taskTopic = map["Topic"] ? sheet.getRange(row, col_(map, "Topic")).getValue() : "";
      if (taskModule && taskTopic) adjustRevisionCompletedTaskCount_(taskModule, taskTopic, -1);
    }
    sheet.deleteRow(row);
    logAutomation_("Study task deleted", taskText || taskId, "Deleted", "ID " + taskId + "; row " + row);
    return ok_({});
  } catch (e) { return fail_(e); } finally { lock.releaseLock(); }
}

// ------------------------------------------------------------------
// WEEKLY STUDY SUGGESTIONS — preview + accept (Feature: §4.6, improved by
// v1.3.0 spec §7). Computation lives in buildPrioritizedStudySuggestions_
// (Calendar_Sync_v1.3.0.gs), which folds in priority/assessment
// proximity/difficulty/workload/repeated-status/min-hours/incomplete
// tasks/weak topics/resource status -- see V1.3.0_RULE_CATALOGUE.md. This is
// still just the validated web-app entry point plus the "write only what
// was accepted" half of the flow -- api_acceptStudySuggestions below is
// completely unmodified from v1.2.0/RC2 and never runs automatically.
// ------------------------------------------------------------------
/**
 * `manualPriorities` (v1.4.3) is an OPTIONAL {moduleCode: category} map --
 * category one of VALID_PRIORITY_OVERRIDES -- entered fresh in the Study
 * Planner's "Which modules?" checklist immediately before this call and
 * NEVER saved anywhere (replaces v1.4.2's persistent "Manual priority
 * override" column, which kept silently overriding the computed category
 * indefinitely once set -- see ai_computePriority_). Applies ONLY to this
 * one generation: a module with no computed priority (e.g. Industrial
 * Engineering, "Rules Incomplete") needs an entry here to be eligible at
 * all; a module WITH a computed priority can still have its category
 * overridden for just this call if you disagree with it on the day.
 */
function api_generateStudySuggestions(weekAnchorIso, manualPriorities) {
  try {
    var anchor = new Date();
    if (weekAnchorIso && /^\d{4}-\d{2}-\d{2}$/.test(weekAnchorIso)) {
      var parts = weekAnchorIso.split("-");
      anchor = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
    }
    // v1.3.0 -- reuses the exact same batched read api_getPlannerData()
    // already performs (including the priority engine) rather than
    // re-implementing "what is each module's priority" a second time here.
    var plannerData = api_getPlannerData();
    if (!plannerData.ok) return plannerData;
    // v1.4.3 -- apply the ephemeral overrides to a COPY of the computed
    // priorities, never to plannerData.data.priorities itself -- nothing
    // here is written back to any sheet.
    var priorities = plannerData.data.priorities.map(function (p) {
      var override = manualPriorities && manualPriorities[p.code];
      if (!override || VALID_PRIORITY_OVERRIDES.indexOf(override) === -1) return p;
      return Object.assign({}, p, {
        category: override, score: null,
        reasons: ["Priority entered for this generation only: \"" + override + "\" -- not saved."]
      });
    });
    var built = buildPrioritizedStudySuggestions_(anchor, priorities, plannerData.data.revisionTracker, plannerData.data.resources, plannerData.data.studyTasks);
    return ok_(built);
  } catch (e) { return fail_(e); }
}

/**
 * Writes ONLY the suggestions the user explicitly accepted (each item must
 * have already been assigned a real module by the user during the
 * Accept/Edit/Reject preview step in the frontend — never invented here).
 * Reuses appendStudySessionRow_, so this is not a second "how a study
 * session gets written" implementation.
 */
// v1.2.0 RC2 -- correction #3: deterministic (never AI-generated) starter
// checklists for sessions created by accepting a weekly study suggestion,
// keyed only by Study type. Every list below is fixed, hand-authored, and
// reviewed as part of this release -- there is no model call, no external
// service, and no randomness involved anywhere in this feature. The user
// can edit or delete any of these tasks exactly like a manually-added one;
// nothing afterward treats them any differently.
var STARTER_TASKS_BY_STUDY_TYPE = {
  "Practice": ["Complete targeted practice questions", "Mark answers", "Review mistakes", "Redo incorrect questions"],
  "Learn": ["Review the selected concept", "Complete a small set of checking questions", "Record unclear points"],
  "Formula": ["Review formulas", "Test recall without notes", "Apply formulas in questions"],
  "Programming": ["Complete coding exercises", "Test the code", "Fix errors and record difficult concepts"],
  "Exam": ["Attempt timed questions or a paper", "Mark the attempt", "Record score", "Review mistakes", "Redo failed questions"],
  "Mistake": ["Review recorded mistakes", "Identify the cause", "Redo the affected questions", "Add a prevention note or flashcard where useful"]
};

/**
 * Writes the fixed starter checklist for `studyType` (if one is defined --
 * Preview/Recall/Visualization/Synthesis/Weekly Review simply get no
 * starter tasks, exactly like a manually created session would) onto the
 * given session, using the same "21 Study Tasks" schema and stable
 * TSK-prefixed IDs api_addStudyTask uses. Silently does nothing if the
 * Study Tasks migration hasn't been run yet (isolated, not an error).
 */
function addStarterTasksForSession_(sessionId, moduleName, topic, studyType) {
  var starters = STARTER_TASKS_BY_STUDY_TYPE[studyType];
  if (!starters || !starters.length) return;
  var sheet = SpreadsheetApp.getActive().getSheetByName(STUDY_TASKS_SHEET);
  if (!sheet) return;
  var map = getColMap_(sheet);
  starters.forEach(function (text, i) {
    var row = firstBlankRow_(sheet, col_(map, "Study Session ID"));
    var newId = nextId_("TSK", getColumnValues_(sheet, col_(map, "Task ID")));
    sheet.getRange(row, col_(map, "Task ID")).setValue(newId);
    sheet.getRange(row, col_(map, "Study Session ID")).setValue(sessionId);
    sheet.getRange(row, col_(map, "Module")).setValue(moduleName);
    if (topic) sheet.getRange(row, col_(map, "Topic")).setValue(topic);
    sheet.getRange(row, col_(map, "Task type")).setValue("Custom");
    sheet.getRange(row, col_(map, "Task text")).setValue(text);
    sheet.getRange(row, col_(map, "Completed")).setValue(false);
    sheet.getRange(row, col_(map, "Sort order")).setValue(i + 1);
    sheet.getRange(row, col_(map, "Last updated")).setValue(new Date());
  });
}

function api_acceptStudySuggestions(items) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(LOCK_WAIT_MS)) return fail_(new Error("Workbook is busy — try again in a moment."));
    if (!items || !items.length) return fail_(new Error("No suggestions selected."));
    var created = [];
    items.forEach(function (item) {
      var module = requireModule_(item.module);
      var dateVal = parseValidDate_(item.date, "Date");
      if (!dateVal) throw new Error("Invalid suggestion date.");
      var startTime = parseTimeOfDay_(item.startTimeIso || Utilities.formatString("%02d:00", item.hour || 0), dateVal);
      var duration = num_(item.duration) || 60;
      var endTime = startTime ? new Date(startTime.getTime() + duration * 60000) : null;
      var studyType = VALID_STUDY_TYPES.indexOf(item.studyType) !== -1 ? item.studyType : "Practice";
      var topic = cleanText_(item.topic, 200);
      var result = appendStudySessionRow_({
        module: module, date: dateVal, startTime: startTime, endTime: endTime, duration: duration,
        studyType: studyType, topic: topic, priority: "Medium",
        notes: "Accepted from weekly suggestions.", planned: true, completed: false,
        status: "Planned", origin: "Generated"
      });
      // v1.2.0 RC2 -- correction #3: give every accepted suggestion its
      // deterministic starter checklist right away, so a generated session
      // never lands in the Study Planner as an empty shell.
      addStarterTasksForSession_(result.id, module, topic, studyType);
      created.push(result.id);
    });
    logAutomation_("Weekly study suggestions accepted", created.length + " session(s)", "Created", created.join(", "));
    return ok_({ created: created });
  } catch (e) { return fail_(e); } finally { lock.releaseLock(); }
}

// ------------------------------------------------------------------
// REVISION TRACKER — auto-update from completed study, plus manual topic
// add/edit for confidence (Feature: Revision Tracker integration, §5).
// ------------------------------------------------------------------
function findOrCreateRevisionRow_(moduleName, topic) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(REVISION_SHEET);
  var map = getColMap_(sheet);
  var lastRow = sheet.getLastRow();
  if (lastRow > HEADER_ROW) {
    var vals = sheet.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, sheet.getLastColumn()).getValues();
    var moduleCol = col_(map, "Module") - 1, topicCol = col_(map, "Topic") - 1;
    for (var i = 0; i < vals.length; i++) {
      if (vals[i][moduleCol] === moduleName && vals[i][topicCol] === topic) return HEADER_ROW + 1 + i;
    }
  }
  var row = firstBlankRow_(sheet, col_(map, "Module"));
  sheet.getRange(row, col_(map, "Module")).setValue(moduleName);
  sheet.getRange(row, col_(map, "Topic")).setValue(topic);
  sheet.getRange(row, col_(map, "Priority")).setValue("Medium");
  sheet.getRange(row, col_(map, "Completed")).setValue(false);
  return row;
}

/**
 * Called by api_completeStudySession / api_logIndependentStudy whenever a
 * Module+Topic pair is studied. Matches EXACTLY on Module + Topic (never
 * fuzzy) -- creates the row on first use, otherwise accumulates onto the
 * existing one. Never sets Confidence -- that stays the one field a person
 * sets deliberately, per spec §5's explicit "do not infer confidence
 * automatically" rule.
 */
function updateRevisionFromStudy_(moduleName, topic, opts) {
  if (!moduleName || !topic) return { updated: false };
  var sheet = SpreadsheetApp.getActive().getSheetByName(REVISION_SHEET);
  if (!sheet) return { updated: false };
  var map = getColMap_(sheet);
  var row = findOrCreateRevisionRow_(moduleName, topic);

  var minutes = num_(opts.minutes) || 0;
  if (map["Total actual study minutes"]) {
    var prevMinutes = num_(sheet.getRange(row, col_(map, "Total actual study minutes")).getValue()) || 0;
    sheet.getRange(row, col_(map, "Total actual study minutes")).setValue(prevMinutes + minutes);
  }
  if (map["Session count"]) {
    var prevCount = num_(sheet.getRange(row, col_(map, "Session count")).getValue()) || 0;
    sheet.getRange(row, col_(map, "Session count")).setValue(prevCount + 1);
  }
  if (map["Last revised"]) sheet.getRange(row, col_(map, "Last revised")).setValue(new Date());
  if (opts.score != null && map["Score"]) sheet.getRange(row, col_(map, "Score")).setValue(opts.score);

  if (map["Planned-study count"] || map["Independent-study count"]) {
    var plannedCount = map["Planned-study count"] ? (num_(sheet.getRange(row, col_(map, "Planned-study count")).getValue()) || 0) : 0;
    var independentCount = map["Independent-study count"] ? (num_(sheet.getRange(row, col_(map, "Independent-study count")).getValue()) || 0) : 0;
    if (opts.isIndependent) independentCount++; else plannedCount++;
    if (map["Planned-study count"]) sheet.getRange(row, col_(map, "Planned-study count")).setValue(plannedCount);
    if (map["Independent-study count"]) sheet.getRange(row, col_(map, "Independent-study count")).setValue(independentCount);
    if (map["Source"]) {
      var source = (plannedCount > 0 && independentCount > 0) ? "Both" : (independentCount > 0 ? "Independent" : "Planned");
      sheet.getRange(row, col_(map, "Source")).setValue(source);
    }
  }
  return { updated: true, row: row };
}

// ------------------------------------------------------------------
// REVISION TRACKER CONTRIBUTION LEDGER (v1.2.0 RC2 -- correction #1: make
// Study Planner's contribution to Revision Tracker idempotent).
//
// Each Study Planner session stores exactly what it last contributed
// (module/topic/minutes/source) in five additive columns on "12 Study
// Planner" -- see Revision_Contribution_v1.2.0_RC2_Migration.gs. Every time
// a session's contribution might change (completing it, editing its actual
// duration, changing its module/topic, resubmitting the same completion
// form, changing its status between Completed and Partly completed, or
// deleting it), the OLD stored contribution -- reversed against whatever
// module/topic it was ORIGINALLY recorded against, which may differ from
// the session's current module/topic -- is undone first, then the new
// contribution (if the session currently qualifies at all) is applied, then
// the ledger columns are updated to reflect exactly what was just applied.
// This guarantees a session contributes exactly once, at its current
// values, no matter how many times its completion form is saved, and never
// leaves a stale contribution behind if its module/topic/status changes.
//
// If the migration hasn't been run yet, every ledger read degrades to
// "nothing stored" and every ledger write is silently skipped -- contributions
// still apply (exactly like v1.2.0 RC behaved), just without the new
// idempotency guarantee until the migration has run.
// ------------------------------------------------------------------
var REVISION_CONTRIBUTION_COLS = {
  applied: "Revision contribution applied",
  module: "Revision contribution module",
  topic: "Revision contribution topic",
  minutes: "Revision contribution minutes",
  source: "Revision contribution source"
};

function readStoredContribution_(sheet, map, row) {
  var c = REVISION_CONTRIBUTION_COLS;
  if (!map[c.applied]) return { applied: false };
  var applied = sheet.getRange(row, col_(map, c.applied)).getValue() === true;
  if (!applied) return { applied: false };
  return {
    applied: true,
    module: sheet.getRange(row, col_(map, c.module)).getValue(),
    topic: sheet.getRange(row, col_(map, c.topic)).getValue(),
    minutes: num_(sheet.getRange(row, col_(map, c.minutes)).getValue()) || 0,
    source: sheet.getRange(row, col_(map, c.source)).getValue()
  };
}

function writeStoredContribution_(sheet, map, row, contribution) {
  var c = REVISION_CONTRIBUTION_COLS;
  if (!map[c.applied]) return;
  sheet.getRange(row, col_(map, c.applied)).setValue(!!contribution.applied);
  sheet.getRange(row, col_(map, c.module)).setValue(contribution.applied ? contribution.module : "");
  sheet.getRange(row, col_(map, c.topic)).setValue(contribution.applied ? contribution.topic : "");
  sheet.getRange(row, col_(map, c.minutes)).setValue(contribution.applied ? (contribution.minutes || 0) : 0);
  sheet.getRange(row, col_(map, c.source)).setValue(contribution.applied ? contribution.source : "");
}

/**
 * Reverses a previously-applied contribution against its ORIGINAL
 * module/topic (never the session's possibly-since-changed current
 * module/topic -- that is exactly what makes changing a session's
 * module/topic safe to reverse correctly). Never creates a Revision
 * Tracker row -- if the topic row is already gone there is nothing to
 * reverse against. Never lets any aggregate drop below zero. Recomputes
 * Source from the post-decrement Planned/Independent counts using the same
 * rule updateRevisionFromStudy_ uses when applying, so Source never lags
 * behind after a reversal. Deliberately does not touch Confidence, Priority,
 * Next revision date, Completed, Notes, or Weak topic -- reversal only ever
 * undoes the specific aggregate fields a contribution added.
 */
function reverseRevisionContribution_(contribution) {
  if (!contribution || !contribution.applied || !contribution.module || !contribution.topic) return;
  var sheet = SpreadsheetApp.getActive().getSheetByName(REVISION_SHEET);
  if (!sheet) return;
  var map = getColMap_(sheet);
  var lastRow = sheet.getLastRow();
  var targetRow = -1;
  if (lastRow > HEADER_ROW) {
    var vals = sheet.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, sheet.getLastColumn()).getValues();
    var moduleCol = col_(map, "Module") - 1, topicCol = col_(map, "Topic") - 1;
    for (var i = 0; i < vals.length; i++) {
      if (vals[i][moduleCol] === contribution.module && vals[i][topicCol] === contribution.topic) { targetRow = HEADER_ROW + 1 + i; break; }
    }
  }
  if (targetRow === -1) return; // topic row no longer exists -- nothing to reverse against; never recreated here

  function dec(header, amount) {
    if (!map[header]) return;
    var current = num_(sheet.getRange(targetRow, col_(map, header)).getValue()) || 0;
    sheet.getRange(targetRow, col_(map, header)).setValue(Math.max(0, current - amount));
  }
  dec("Total actual study minutes", contribution.minutes || 0);
  dec("Session count", 1);
  if (contribution.source === "Independent") dec("Independent-study count", 1);
  else dec("Planned-study count", 1);

  if (map["Source"] && map["Planned-study count"] && map["Independent-study count"]) {
    var plannedCount = num_(sheet.getRange(targetRow, col_(map, "Planned-study count")).getValue()) || 0;
    var independentCount = num_(sheet.getRange(targetRow, col_(map, "Independent-study count")).getValue()) || 0;
    var source = (plannedCount > 0 && independentCount > 0) ? "Both" : (independentCount > 0 ? "Independent" : (plannedCount > 0 ? "Planned" : ""));
    sheet.getRange(targetRow, col_(map, "Source")).setValue(source);
  }
}

/**
 * The one entry point api_completeStudySession and api_logIndependentStudy
 * both use to keep a single Study Planner session's Revision Tracker
 * contribution correct and single-counted: reverse whatever this exact
 * session previously contributed (if the ledger says it did), then apply
 * its current contribution (if opts.shouldContribute), then record the new
 * ledger state. Both steps run under the caller's already-held lock -- this
 * function takes no lock of its own.
 */
function reconcileSessionRevisionContribution_(sheet, map, row, moduleName, topic, opts) {
  var previous = readStoredContribution_(sheet, map, row);
  reverseRevisionContribution_(previous);

  var shouldContribute = !!(opts.shouldContribute && moduleName && topic);
  var result = { updated: false };
  if (shouldContribute) {
    result = updateRevisionFromStudy_(moduleName, topic, opts);
    writeStoredContribution_(sheet, map, row, {
      applied: true, module: moduleName, topic: topic, minutes: num_(opts.minutes) || 0,
      source: opts.isIndependent ? "Independent" : "Planned"
    });
  } else {
    writeStoredContribution_(sheet, map, row, { applied: false });
  }
  return result;
}

/**
 * Used by api_deleteStudySession (correction #7) to reverse a session's
 * contribution before its row disappears. Reads the ledger by stable ID
 * under its own lock/release cycle -- deleteEntityRow_ then re-acquires the
 * lock to perform the actual delete. Both steps are inherently sequential
 * within a single web-app request, so there is no window where another
 * request could observe a half-reversed state. Best-effort: a failure here
 * (e.g. workbook briefly busy) never blocks the session delete itself --
 * worst case is a small aggregate drift that a later edit/recalculation
 * would correct, never a broken delete.
 */
function reverseSessionRevisionContributionById_(sessionId) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(LOCK_WAIT_MS)) return;
    var row = findRowByStableId_(STUDY_SHEET, "Study Session ID", sessionId);
    if (row === -1) return;
    var sheet = SpreadsheetApp.getActive().getSheetByName(STUDY_SHEET);
    var map = getColMap_(sheet);
    var previous = readStoredContribution_(sheet, map, row);
    reverseRevisionContribution_(previous);
    writeStoredContribution_(sheet, map, row, { applied: false });
  } catch (e) { /* best-effort only -- never blocks the session delete */ } finally { lock.releaseLock(); }
}

function api_addRevisionTopic(form) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(LOCK_WAIT_MS)) return fail_(new Error("Workbook is busy — try again in a moment."));
    if (!form) return fail_(new Error("No form data received."));
    var module = requireModule_(form.module);
    var topic = requireText_(form.topic, "Topic", 200);
    var sheet = SpreadsheetApp.getActive().getSheetByName(REVISION_SHEET);
    var map = getColMap_(sheet);
    var lastRow = sheet.getLastRow();
    if (lastRow > HEADER_ROW) {
      var vals = sheet.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, sheet.getLastColumn()).getValues();
      var moduleCol = col_(map, "Module") - 1, topicCol = col_(map, "Topic") - 1;
      for (var i = 0; i < vals.length; i++) {
        if (vals[i][moduleCol] === module && vals[i][topicCol] === topic) {
          return fail_(new Error('A revision topic for "' + topic + '" already exists for ' + module + '. Edit it instead of adding a duplicate.'));
        }
      }
    }
    var row = firstBlankRow_(sheet, col_(map, "Module"));
    sheet.getRange(row, col_(map, "Module")).setValue(module);
    sheet.getRange(row, col_(map, "Topic")).setValue(topic);
    sheet.getRange(row, col_(map, "Priority")).setValue(VALID_PRIORITIES.indexOf(form.priority) !== -1 ? form.priority : "Medium");
    if (form.confidence !== "" && form.confidence != null) sheet.getRange(row, col_(map, "Confidence")).setValue(requireFiniteNumber_(form.confidence, "Confidence", 0, 100));
    sheet.getRange(row, col_(map, "Completed")).setValue(false);
    if (form.notes) sheet.getRange(row, col_(map, "Notes")).setValue(cleanMultiline_(form.notes, 1000));
    logAutomation_("Revision topic added", module + " — " + topic, "Created", "");
    return ok_({});
  } catch (e) { return fail_(e); } finally { lock.releaseLock(); }
}

/** Manual edits -- primarily Confidence, which nothing else in this release
 *  ever sets automatically (per spec §5). */
function api_updateRevisionTopic(module, topic, form) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(LOCK_WAIT_MS)) return fail_(new Error("Workbook is busy — try again in a moment."));
    if (!form) return fail_(new Error("No form data received."));
    var sheet = SpreadsheetApp.getActive().getSheetByName(REVISION_SHEET);
    var map = getColMap_(sheet);
    var lastRow = sheet.getLastRow();
    var targetRow = -1;
    if (lastRow > HEADER_ROW) {
      var vals = sheet.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, sheet.getLastColumn()).getValues();
      var moduleCol = col_(map, "Module") - 1, topicCol = col_(map, "Topic") - 1;
      for (var i = 0; i < vals.length; i++) {
        if (vals[i][moduleCol] === module && vals[i][topicCol] === topic) { targetRow = HEADER_ROW + 1 + i; break; }
      }
    }
    if (targetRow === -1) return fail_(new Error('No revision topic found for "' + topic + '" in ' + module + '.'));
    if (form.confidence !== undefined && form.confidence !== "") sheet.getRange(targetRow, col_(map, "Confidence")).setValue(requireFiniteNumber_(form.confidence, "Confidence", 0, 100));
    if (form.priority) sheet.getRange(targetRow, col_(map, "Priority")).setValue(VALID_PRIORITIES.indexOf(form.priority) !== -1 ? form.priority : "Medium");
    if (typeof form.weakTopic === "boolean" && map["Weak topic (TRUE/FALSE)"]) {
      sheet.getRange(targetRow, col_(map, "Weak topic (TRUE/FALSE)")).setValue(form.weakTopic);
      // v1.3.0 -- stamps/clears the moment a topic flips weak, into the
      // column added by Academic_Intelligence_v1.3.0_Migration.gs. Read by
      // the Exam Focus Mode "repeated weak topic" trigger rule and shown in
      // priority reasons -- never used for marks/aggregates. No-op on a
      // workbook that hasn't run that migration yet.
      if (map["Weak topic since"]) {
        sheet.getRange(targetRow, col_(map, "Weak topic since")).setValue(form.weakTopic ? new Date() : "");
      }
    }
    if (form.nextRevisionDate) { var nd = parseValidDate_(form.nextRevisionDate, "Next revision date"); if (nd) sheet.getRange(targetRow, col_(map, "Next revision date")).setValue(nd); }
    if (typeof form.completed === "boolean") sheet.getRange(targetRow, col_(map, "Completed")).setValue(form.completed);
    if (form.notes !== undefined) sheet.getRange(targetRow, col_(map, "Notes")).setValue(cleanMultiline_(form.notes, 1000));
    logAutomation_("Revision topic updated", module + " — " + topic, "Updated", "");
    return ok_({});
  } catch (e) { return fail_(e); } finally { lock.releaseLock(); }
}

// ------------------------------------------------------------------
// RESOURCES — lightweight named-materials CRUD + type-specific progress
// (Feature 6). Never uploads or stores a file -- Link/Notes stay plain text.
// ------------------------------------------------------------------
function api_addResource(form) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(LOCK_WAIT_MS)) return fail_(new Error("Workbook is busy — try again in a moment."));
    if (!form) return fail_(new Error("No form data received."));
    var module = requireModule_(form.module);
    var title = requireText_(form.title, "Resource name", 200);
    var type = VALID_RESOURCE_TYPES.indexOf(form.type) !== -1 ? form.type : "Other";
    var topic = cleanText_(form.topic, 200);
    var link = cleanText_(form.link, 300);
    var notes = cleanMultiline_(form.notes, 1000);
    var priority = VALID_PRIORITIES.indexOf(form.priority) !== -1 ? form.priority : "Medium";

    var sheet = SpreadsheetApp.getActive().getSheetByName(RESOURCES_SHEET);
    var map = getColMap_(sheet);
    if (!map["Resource ID"]) throw new Error('"Resource ID" column not found — run Resources_v1.2.0_Migration.gs first.');
    var row = firstBlankRow_(sheet, col_(map, "Title"));
    var newId = nextId_("RES", getColumnValues_(sheet, col_(map, "Resource ID")));
    sheet.getRange(row, col_(map, "Resource ID")).setValue(newId);
    sheet.getRange(row, col_(map, "Module")).setValue(module);
    sheet.getRange(row, col_(map, "Resource type")).setValue(type);
    sheet.getRange(row, col_(map, "Title")).setValue(title);
    if (topic) sheet.getRange(row, col_(map, "Topic")).setValue(topic);
    if (link) sheet.getRange(row, col_(map, "Link or file reference")).setValue(link);
    if (notes) sheet.getRange(row, col_(map, "Notes")).setValue(notes);
    sheet.getRange(row, col_(map, "Priority")).setValue(priority);
    sheet.getRange(row, col_(map, "Status")).setValue("Not started");
    sheet.getRange(row, col_(map, "Reviewed")).setValue(false);
    logAutomation_("Resource added", module + " — " + title, "Created", newId);
    return ok_({ id: newId });
  } catch (e) { return fail_(e); } finally { lock.releaseLock(); }
}

function api_updateResource(resourceId, form) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(LOCK_WAIT_MS)) return fail_(new Error("Workbook is busy — try again in a moment."));
    if (!form) return fail_(new Error("No form data received."));
    var row = findRowByStableId_(RESOURCES_SHEET, "Resource ID", resourceId);
    if (row === -1) return fail_(new Error("Could not find resource " + resourceId + " — it may already have been deleted."));
    var sheet = SpreadsheetApp.getActive().getSheetByName(RESOURCES_SHEET);
    var map = getColMap_(sheet);
    if (form.title) sheet.getRange(row, col_(map, "Title")).setValue(requireText_(form.title, "Resource name", 200));
    if (form.type) sheet.getRange(row, col_(map, "Resource type")).setValue(VALID_RESOURCE_TYPES.indexOf(form.type) !== -1 ? form.type : "Other");
    if (form.topic !== undefined) sheet.getRange(row, col_(map, "Topic")).setValue(cleanText_(form.topic, 200));
    if (form.link !== undefined) sheet.getRange(row, col_(map, "Link or file reference")).setValue(cleanText_(form.link, 300));
    if (form.priority) sheet.getRange(row, col_(map, "Priority")).setValue(VALID_PRIORITIES.indexOf(form.priority) !== -1 ? form.priority : "Medium");
    if (form.notes !== undefined) sheet.getRange(row, col_(map, "Notes")).setValue(cleanMultiline_(form.notes, 1000));
    logAutomation_("Resource updated", resourceId, "Updated", "");
    return ok_({});
  } catch (e) { return fail_(e); } finally { lock.releaseLock(); }
}

function api_deleteResource(resourceId) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(LOCK_WAIT_MS)) return fail_(new Error("Workbook is busy — try again in a moment."));
    var row = findRowByStableId_(RESOURCES_SHEET, "Resource ID", resourceId);
    if (row === -1) return fail_(new Error("Could not find resource " + resourceId + " — it may already have been deleted."));
    var sheet = SpreadsheetApp.getActive().getSheetByName(RESOURCES_SHEET);
    var map = getColMap_(sheet);
    var title = sheet.getRange(row, col_(map, "Title")).getValue();
    sheet.deleteRow(row);
    logAutomation_("Resource deleted", title || resourceId, "Deleted", "ID " + resourceId + "; row " + row);
    return ok_({});
  } catch (e) { return fail_(e); } finally { lock.releaseLock(); }
}

/** Type-specific progress: exactly four generic flags (labelled differently
 *  per resource type by the frontend, the same "one small lookup table"
 *  technique already used by afTemplateByCode in Index.html) plus an
 *  optional score and a flashcards-created flag -- deliberately NOT a
 *  bespoke column per resource type, per spec §6's "keep this lightweight"
 *  instruction. */
function api_updateResourceProgress(resourceId, form) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(LOCK_WAIT_MS)) return fail_(new Error("Workbook is busy — try again in a moment."));
    if (!form) return fail_(new Error("No form data received."));
    var row = findRowByStableId_(RESOURCES_SHEET, "Resource ID", resourceId);
    if (row === -1) return fail_(new Error("Could not find resource " + resourceId + " — it may already have been deleted."));
    var sheet = SpreadsheetApp.getActive().getSheetByName(RESOURCES_SHEET);
    var map = getColMap_(sheet);
    if (form.status) {
      var status = VALID_RESOURCE_STATUSES.indexOf(form.status) !== -1 ? form.status : "Not started";
      sheet.getRange(row, col_(map, "Status")).setValue(status);
      if (map["Reviewed"]) sheet.getRange(row, col_(map, "Reviewed")).setValue(status === "Reviewed" || status === "Completed");
    }
    ["flagA", "flagB", "flagC", "flagD"].forEach(function (key, i) {
      var header = "Progress Flag " + String.fromCharCode(65 + i);
      if (typeof form[key] === "boolean" && map[header]) sheet.getRange(row, col_(map, header)).setValue(form[key]);
    });
    if (typeof form.flashcardsCreated === "boolean" && map["Flashcards created (TRUE/FALSE)"]) {
      sheet.getRange(row, col_(map, "Flashcards created (TRUE/FALSE)")).setValue(form.flashcardsCreated);
    }
    if (form.score !== undefined && form.score !== "" && map["Score"]) {
      sheet.getRange(row, col_(map, "Score")).setValue(requireFiniteNumber_(form.score, "Score", 0, 100));
      // v1.3.0 -- stamps the real moment a score is saved, into the column
      // added by Academic_Intelligence_v1.3.0_Migration.gs. This is the only
      // thing the Improvement Tracker's past-paper/practice score trend uses
      // to order scores chronologically -- silently a no-op on a workbook
      // that hasn't run that migration yet (map["Score updated"] is then
      // undefined), so this never breaks api_updateResourceProgress on an
      // un-migrated v1.2.0 workbook.
      if (map["Score updated"]) sheet.getRange(row, col_(map, "Score updated")).setValue(new Date());
    }
    logAutomation_("Resource progress updated", resourceId, "Updated", "");
    return ok_({});
  } catch (e) { return fail_(e); } finally { lock.releaseLock(); }
}

/** "Attach to Study Planner session" / "create a Study Planner task from a
 *  resource" (spec §6 actions) -- both are the same underlying action: file
 *  a new Study Task against the given session, referencing this Resource ID,
 *  with a task type inferred from the resource type. */
function api_createStudyTaskFromResource(resourceId, sessionId) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(LOCK_WAIT_MS)) return fail_(new Error("Workbook is busy — try again in a moment."));
    var resRow = findRowByStableId_(RESOURCES_SHEET, "Resource ID", resourceId);
    if (resRow === -1) return fail_(new Error("Could not find resource " + resourceId + "."));
    var sessionRow = findRowByStableId_(STUDY_SHEET, "Study Session ID", sessionId);
    if (sessionRow === -1) return fail_(new Error("Could not find study session " + sessionId + "."));

    var resSheet = SpreadsheetApp.getActive().getSheetByName(RESOURCES_SHEET);
    var resMap = getColMap_(resSheet);
    var resTitle = resSheet.getRange(resRow, col_(resMap, "Title")).getValue();
    var resType = resSheet.getRange(resRow, col_(resMap, "Resource type")).getValue();

    var typeToTask = {
      "Lecture": "Watch lecture", "Lecture video": "Watch lecture", "Tutorial": "Complete tutorial",
      "Past paper": "Attempt past paper", "Formula bible": "Formula review", "Textbook chapter": "Review notes",
      "Practice set": "Practice questions", "Flashcards": "Flashcards", "Summary notes": "Review notes",
      "Mistake review": "Review mistakes", "Other": "Custom"
    };

    var studySheet = SpreadsheetApp.getActive().getSheetByName(STUDY_SHEET);
    var studyMap = getColMap_(studySheet);
    var module = studySheet.getRange(sessionRow, col_(studyMap, "Module")).getValue();
    var topic = studyMap["Topic"] ? studySheet.getRange(sessionRow, col_(studyMap, "Topic")).getValue() : "";

    var sheet = SpreadsheetApp.getActive().getSheetByName(STUDY_TASKS_SHEET);
    if (!sheet) return fail_(new Error('Sheet not found: "' + STUDY_TASKS_SHEET + '" — run Study_Tasks_v1.2.0_Migration.gs first.'));
    var map = getColMap_(sheet);
    var row = firstBlankRow_(sheet, col_(map, "Study Session ID"));
    var newId = nextId_("TSK", getColumnValues_(sheet, col_(map, "Task ID")));
    var existingCount = readSheetRows_(STUDY_TASKS_SHEET, HEADER_ROW, "Study Session ID")
      .filter(function (t) { return t["Study Session ID"] === sessionId; }).length;
    sheet.getRange(row, col_(map, "Task ID")).setValue(newId);
    sheet.getRange(row, col_(map, "Study Session ID")).setValue(sessionId);
    sheet.getRange(row, col_(map, "Module")).setValue(module);
    if (topic) sheet.getRange(row, col_(map, "Topic")).setValue(topic);
    sheet.getRange(row, col_(map, "Task type")).setValue(typeToTask[resType] || "Custom");
    sheet.getRange(row, col_(map, "Task text")).setValue(resTitle);
    sheet.getRange(row, col_(map, "Resource ID")).setValue(resourceId);
    sheet.getRange(row, col_(map, "Completed")).setValue(false);
    sheet.getRange(row, col_(map, "Sort order")).setValue(existingCount + 1);
    sheet.getRange(row, col_(map, "Last updated")).setValue(new Date());
    logAutomation_("Study task created from resource", resTitle, "Created", newId);
    return ok_({ id: newId });
  } catch (e) { return fail_(e); } finally { lock.releaseLock(); }
}

// ------------------------------------------------------------------
// ATTENDANCE — simple Attended/Missed tracking for mandatory sessions only
// (Feature 7). No percentages, no analytics beyond what the existing
// Attendance screen aggregate already computes from the "Attended" boolean.
// ------------------------------------------------------------------
/**
 * Finds-or-creates the one Attendance row for this exact real timetable
 * session (Module + Session type + Date, packed into timetableSessionKey by
 * the frontend as "Module|SessionType|DateISO") and sets its Status --
 * duplicate-proof because the lookup is keyed on that composite, never on
 * row position or a fresh append. Keeps the pre-existing "Attended" boolean
 * column in sync (TRUE only when Status === "Attended") purely so the
 * Dashboard/Attendance screen's existing aggregate keeps working unchanged.
 */
function api_setSessionAttendance(timetableSessionKey, status) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(LOCK_WAIT_MS)) return fail_(new Error("Workbook is busy — try again in a moment."));
    if (VALID_ATTENDANCE_STATUSES.indexOf(status) === -1) return fail_(new Error("Invalid attendance status: " + status));
    var parts = String(timetableSessionKey || "").split("|");
    if (parts.length !== 3) return fail_(new Error("Invalid session reference."));
    var module = parts[0], sessionType = parts[1], dateIso = parts[2];
    var dateVal = parseValidDate_(dateIso, "Date");
    if (!dateVal) return fail_(new Error("Invalid session date."));

    var sheet = SpreadsheetApp.getActive().getSheetByName(ATTENDANCE_SHEET);
    if (!sheet) return fail_(new Error("Sheet not found: " + ATTENDANCE_SHEET));
    var map = getColMap_(sheet);
    if (!map["Session key"]) throw new Error('"Session key" column not found — run Attendance_v1.2.0_Migration.gs first.');

    var lastRow = sheet.getLastRow();
    var targetRow = -1;
    if (lastRow > HEADER_ROW) {
      var keys = sheet.getRange(HEADER_ROW + 1, col_(map, "Session key"), lastRow - HEADER_ROW, 1).getValues();
      for (var i = 0; i < keys.length; i++) {
        if (keys[i][0] === timetableSessionKey) { targetRow = HEADER_ROW + 1 + i; break; }
      }
    }
    if (targetRow === -1) {
      targetRow = firstBlankRow_(sheet, col_(map, "Module"));
      var newId = nextId_("ATT", getColumnValues_(sheet, col_(map, "Attendance ID")));
      sheet.getRange(targetRow, col_(map, "Attendance ID")).setValue(newId);
      sheet.getRange(targetRow, col_(map, "Module")).setValue(module);
      sheet.getRange(targetRow, col_(map, "Session type")).setValue(sessionType);
      sheet.getRange(targetRow, col_(map, "Date")).setValue(dateVal);
      sheet.getRange(targetRow, col_(map, "Session key")).setValue(timetableSessionKey);
    }
    sheet.getRange(targetRow, col_(map, "Status")).setValue(status);
    sheet.getRange(targetRow, col_(map, "Attended")).setValue(status === "Attended");
    logAutomation_("Attendance recorded", module + " " + sessionType, status, dateIso);
    return ok_({ status: status });
  } catch (e) { return fail_(e); } finally { lock.releaseLock(); }
}

/**
 * "Create catch-up study task" (spec §7) -- shown only when a mandatory
 * session was missed. Creates one small ad-hoc Study Planner session
 * (Origin = "Catch-up") plus exactly one linked Study Task, rather than an
 * empty session with nothing to actually do in it.
 */
/**
 * v1.2.0 RC2 -- correction #5: the frontend now always passes the actual
 * missed date (read straight from the timetable session's own session
 * key), and this endpoint now records that date in a dedicated,
 * greppable marker inside Notes ("[Catch-up:Module|SessionType|DateISO]")
 * so the missed date is never lost, and so a second call for the exact
 * same Module + Session type + missed date is recognized as a likely
 * duplicate and refused unless the caller explicitly passes confirmed=true
 * (the frontend does this only after the user picks "Create another
 * anyway" in a confirmation prompt). The `confirmed` parameter is new and
 * additive -- omitting it (as any pre-RC2 caller would) simply means "run
 * the normal duplicate check," which is the same protective default this
 * correction is meant to add.
 */
function api_createCatchUpTask(module, sessionType, dateIso, confirmed) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(LOCK_WAIT_MS)) return fail_(new Error("Workbook is busy — try again in a moment."));
    var moduleName = requireModule_(module);
    var sessType = cleanText_(sessionType, 60);
    var cleanDateIso = cleanText_(dateIso, 20);
    var topic = "Catch-up: " + sessType;
    var marker = cleanDateIso ? ("[Catch-up:" + moduleName + "|" + sessType + "|" + cleanDateIso + "]") : "";

    if (marker && !confirmed) {
      var existingSessions = readSheetRows_(STUDY_SHEET, HEADER_ROW, "Module");
      var isDuplicate = existingSessions.some(function (s) {
        return s["Module"] === moduleName && s["Origin"] === "Catch-up" && String(s["Notes"] || "").indexOf(marker) !== -1;
      });
      if (isDuplicate) {
        return fail_(new Error("DUPLICATE_CATCHUP: A catch-up task for this missed " + sessType + " on " + cleanDateIso + " already exists."));
      }
    }

    var missedLabel = cleanDateIso ? ("Missed " + sessType + " on " + cleanDateIso) : ("Missed " + sessType);
    var notes = marker ? (marker + " " + missedLabel) : missedLabel;
    var sessionResult = appendStudySessionRow_({
      module: moduleName, date: new Date(), studyType: "Practice", topic: topic,
      priority: "High", notes: notes,
      planned: true, completed: false, status: "Planned", origin: "Catch-up"
    });
    var tasksSheet = SpreadsheetApp.getActive().getSheetByName(STUDY_TASKS_SHEET);
    if (tasksSheet) {
      var map = getColMap_(tasksSheet);
      var row = firstBlankRow_(tasksSheet, col_(map, "Study Session ID"));
      var newId = nextId_("TSK", getColumnValues_(tasksSheet, col_(map, "Task ID")));
      tasksSheet.getRange(row, col_(map, "Task ID")).setValue(newId);
      tasksSheet.getRange(row, col_(map, "Study Session ID")).setValue(sessionResult.id);
      tasksSheet.getRange(row, col_(map, "Module")).setValue(moduleName);
      tasksSheet.getRange(row, col_(map, "Topic")).setValue(topic);
      tasksSheet.getRange(row, col_(map, "Task type")).setValue("Custom");
      tasksSheet.getRange(row, col_(map, "Task text")).setValue("Catch up on missed " + sessType + (cleanDateIso ? " (" + cleanDateIso + ")" : ""));
      tasksSheet.getRange(row, col_(map, "Completed")).setValue(false);
      tasksSheet.getRange(row, col_(map, "Sort order")).setValue(1);
      tasksSheet.getRange(row, col_(map, "Last updated")).setValue(new Date());
    }
    logAutomation_("Catch-up task created", moduleName + " — " + sessType + (cleanDateIso ? " (missed " + cleanDateIso + ")" : ""), "Created", sessionResult.id);
    return ok_({ studySessionId: sessionResult.id });
  } catch (e) { return fail_(e); } finally { lock.releaseLock(); }
}

// ------------------------------------------------------------------
// MODULE CONFIGURATION + MINIMAL CONTACTS (Features 9 & 10) — one combined
// save endpoint per spec §12's own suggested API list, writing straight
// onto "03 Modules" (no new sheet, no join). Stored and displayed only in
// this release -- no Marks Priority Engine / Exam Focus Mode logic reads
// these fields yet.
// ------------------------------------------------------------------
function api_saveModuleAcademicConfig(moduleCode, form) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(LOCK_WAIT_MS)) return fail_(new Error("Workbook is busy — try again in a moment."));
    if (!form) return fail_(new Error("No form data received."));
    var sheet = SpreadsheetApp.getActive().getSheetByName(MODULES_SHEET);
    var map = getColMap_(sheet);
    var lastRow = sheet.getLastRow();
    var codes = sheet.getRange(HEADER_ROW + 1, col_(map, "Module code"), lastRow - HEADER_ROW, 1).getValues();
    var targetRow = -1;
    for (var i = 0; i < codes.length; i++) { if (codes[i][0] === moduleCode) { targetRow = HEADER_ROW + 1 + i; break; } }
    if (targetRow === -1) return fail_(new Error("Unknown module code: " + moduleCode));

    var textFieldMap = {
      colour: 'Colour (hex)', coordinatorName: 'Coordinator name', lecturerName: 'Lecturer name'
    };
    var emailFieldMap = { coordinatorEmail: 'Coordinator email', lecturerEmail: 'Lecturer email' };

    Object.keys(textFieldMap).forEach(function (key) {
      var header = textFieldMap[key];
      if (form[key] !== undefined && map[header]) sheet.getRange(targetRow, col_(map, header)).setValue(cleanText_(form[key], 200));
    });
    Object.keys(emailFieldMap).forEach(function (key) {
      var header = emailFieldMap[key];
      if (form[key] !== undefined && map[header]) sheet.getRange(targetRow, col_(map, header)).setValue(cleanText_(form[key], 200));
    });
    if (form.contactNotes !== undefined && map['Contact notes']) sheet.getRange(targetRow, col_(map, 'Contact notes')).setValue(cleanMultiline_(form.contactNotes, 1000));

    if (form.difficulty !== undefined && form.difficulty !== "" && map['Conceptual difficulty (1-5)']) {
      sheet.getRange(targetRow, col_(map, 'Conceptual difficulty (1-5)')).setValue(requireFiniteNumber_(form.difficulty, "Difficulty", 1, 5));
    }
    if (form.workload !== undefined && form.workload !== "" && map['Workload intensity (1-5)']) {
      sheet.getRange(targetRow, col_(map, 'Workload intensity (1-5)')).setValue(requireFiniteNumber_(form.workload, "Workload", 1, 5));
    }
    if (form.targetMark !== undefined && form.targetMark !== "" && map['Target mark (%)']) {
      sheet.getRange(targetRow, col_(map, 'Target mark (%)')).setValue(requireFiniteNumber_(form.targetMark, "Target mark", 0, 100));
    }
    if (form.minWeeklyHours !== undefined && form.minWeeklyHours !== "" && map['Min weekly study hours']) {
      sheet.getRange(targetRow, col_(map, 'Min weekly study hours')).setValue(requireFiniteNumber_(form.minWeeklyHours, "Min weekly study hours", 0, 80));
    }
    if (typeof form.repeated === "boolean" && map['Repeated module (TRUE/FALSE)']) {
      sheet.getRange(targetRow, col_(map, 'Repeated module (TRUE/FALSE)')).setValue(form.repeated);
    }
    if (typeof form.attendanceSensitive === "boolean" && map['Attendance-sensitive lecture (TRUE/FALSE)']) {
      sheet.getRange(targetRow, col_(map, 'Attendance-sensitive lecture (TRUE/FALSE)')).setValue(form.attendanceSensitive);
    }
    if (typeof form.compulsoryPractical === "boolean" && map['Compulsory practical/lab (TRUE/FALSE)']) {
      sheet.getRange(targetRow, col_(map, 'Compulsory practical/lab (TRUE/FALSE)')).setValue(form.compulsoryPractical);
    }
    // v1.4.2 -- Study Planner-specific include switch. (Manual priority
    // override used to be settable here too -- retired in v1.4.3, see
    // ai_computePriority_; priority is now entered fresh per generation.)
    if (typeof form.includeInStudyPlan === "boolean" && map['Include in study plan (TRUE/FALSE)']) {
      sheet.getRange(targetRow, col_(map, 'Include in study plan (TRUE/FALSE)')).setValue(form.includeInStudyPlan);
    }

    logAutomation_("Module configuration saved", moduleCode, "Updated", "");
    return ok_({});
  } catch (e) { return fail_(e); } finally { lock.releaseLock(); }
}

// ============================================================================================
// v1.3.0 DETERMINISTIC ACADEMIC INTELLIGENCE RELEASE -- everything below this
// line is new. Nothing above this line was removed, renamed, or
// behaviourally changed (the two "Weak topic since"/"Score updated" stamps
// added just above to api_updateRevisionTopic/api_updateResourceProgress are
// the only edits to pre-existing functions in this whole file).
//
// Every function below is a PLAIN, DETERMINISTIC function of the data
// already loaded by api_getPlannerData -- no AI, no model call, no external
// HTTP request, no randomness, anywhere in this section. Every "engine" is
// split into a pure computation half (named ai_..._, takes and returns plain
// objects/arrays only, touches no SpreadsheetApp/CalendarApp API -- these are
// unit-tested directly in Node; see V1.3.0_STATIC_VALIDATION.md) and, where a
// read or write is actually needed, a thin Apps-Script-facing wrapper. See
// V1.3.0_RULE_CATALOGUE.md for the full list of rules in plain English and
// V1.3.0_SCORING_METHODOLOGY.md for exactly how every number below is
// computed.
// ============================================================================================

// ------------------------------------------------------------------
// SHARED NUMBER HELPERS
// ------------------------------------------------------------------
/** Rounds to 1 decimal place, avoiding float-drift accumulation across the
 *  many small additions the engines below do. */
function round1_(n) { return Math.round(n * 100) / 100 === Math.round(n) ? Math.round(n) : Math.round(n * 10) / 10; }

/**
 * Normalizes a percentage-shaped value that may already be a plain 0-100
 * number, a 0-1 fraction (Sheets' native "percent format" storage), or a
 * formatted string like "72%" or "72" -- into a plain 0-100 number, or null
 * if it isn't a usable percentage at all. This exists because
 * "Final / provisional (after A3)" and similar Marks Tracker formula outputs
 * are read as-is from the sheet (exactly like every other v1.2.0 read -- see
 * V1.3.0_ARCHITECTURE_NOTES.md) rather than re-derived, and this repo's own
 * code never normalizes their raw type. Never fabricates a value: returns
 * null (never 0) for anything it cannot confidently parse, and every caller
 * below treats null as "insufficient data," not "0%."
 */
function pctToNumber_(v) {
  if (typeof v === "number" && isFinite(v)) return (v > 0 && v <= 1) ? round1_(v * 100) : round1_(v);
  if (typeof v === "string" && v.trim() !== "") {
    var n = parseFloat(v.replace("%", "").trim());
    if (isFinite(n)) return (n > 0 && n <= 1) ? round1_(n * 100) : round1_(n);
  }
  return null;
}

/** Whole days between two ISO ("yyyy-MM-dd") date strings, b - a. Pure --
 *  used by both the live code and the Node unit tests, so "how many days
 *  until an assessment" is computed exactly one way everywhere. */
function daysBetween_(aIso, bIso) {
  var a = new Date(aIso + "T00:00:00");
  var b = new Date(bIso + "T00:00:00");
  if (isNaN(a.getTime()) || isNaN(b.getTime())) return null;
  return Math.round((b.getTime() - a.getTime()) / 86400000);
}

/** Monday-of-the-week key ("yyyy-MM-dd") for a given ISO date string -- used
 *  to bucket sessions/tasks/log entries into weeks for the Improvement
 *  Tracker's weekly trends. Pure JS Date only (no Utilities.formatDate), so
 *  it runs identically in Node and in Apps Script. */
function isoWeekKey_(dateIso) {
  var d = new Date(dateIso + "T00:00:00");
  if (isNaN(d.getTime())) return null;
  var day = (d.getDay() + 6) % 7; // 0 = Monday
  d.setDate(d.getDate() - day);
  var y = d.getFullYear(), m = ("0" + (d.getMonth() + 1)).slice(-2), dd = ("0" + d.getDate()).slice(-2);
  return y + "-" + m + "-" + dd;
}

// ------------------------------------------------------------------
// 1. MARKS PRIORITY ENGINE (spec §1)
// ------------------------------------------------------------------
/**
 * Pure. Input `m` is a plain object describing one module's already-loaded
 * v1.2.0 data (see computeAcademicIntelligence_ below for exactly how it's
 * assembled). Returns { code, name, category, score, reasons }. `score` is
 * null whenever `category` is "Rules Incomplete" or "Insufficient data" --
 * those two categories are the explicit "do not fabricate a projection"
 * escape hatches spec §1 requires, and nothing downstream (allocation, next
 * action, exam focus) ever treats a null score as zero.
 *
 * `m.finalOrProvisional` may be either 11 Marks Tracker's own recorded
 * figure OR (per computeAcademicIntelligence_'s marksIntelligenceByCode
 * fallback) the Marks Intelligence Engine's own currentMark for a module
 * whose A3 hasn't happened yet -- `m.currentMarkIsEstimatedFallback` says
 * which, purely so the reason string below can disclose it honestly.
 */
function ai_computePriority_(m) {
  // v1.4.3 -- the v1.4.2 persistent "Manual priority override" (03 Modules)
  // was retired here: it silently kept overriding the computed category
  // indefinitely once set, with no visible reminder, which produced
  // confusing "why is this module's priority stuck?" results as a course
  // progressed. The Marks Priority Engine goes back to reporting "Rules
  // Incomplete" honestly for a module with no verified framework -- see
  // api_generateStudySuggestions/buildPrioritizedStudySuggestions_ for
  // where a module like that can still get study suggestions: a priority
  // entered fresh each time you generate a plan, never saved.
  if (m.ruleStatus !== "Verified") {
    return { code: m.code, name: m.name, category: "Rules Incomplete", score: null,
      reasons: ["Module Rules for " + m.name + " are not marked \"Verified\" in 04 Module Rules -- no priority is computed until rules are verified."] };
  }
  var current = pctToNumber_(m.finalOrProvisional);
  if (current === null) {
    return { code: m.code, name: m.name, category: "Insufficient data", score: null,
      reasons: ["No final/provisional mark has been computed yet for " + m.name + " in 11 Marks Tracker, and no AF/A1/A2 component data is available to estimate one either -- nothing to project from."] };
  }
  var target = (typeof m.targetMarkOverride === "number") ? m.targetMarkOverride
    : (typeof m.settingsTargetMark === "number") ? m.settingsTargetMark : null;
  if (target === null) {
    return { code: m.code, name: m.name, category: "Insufficient data", score: null,
      reasons: ["No target mark is set for " + m.name + " (neither a per-module target in 03 Modules nor a workbook-wide \"Target final mark (%)\" in 02 Settings) -- cannot measure progress toward a target."] };
  }

  var reasons = [];
  if (m.currentMarkIsEstimatedFallback) {
    reasons.push("No \"Final / provisional (after A3)\" mark is recorded yet in 11 Marks Tracker -- this uses the Marks Intelligence Engine's current mark (which incorporates a verified Estimated A2, if applicable) instead.");
  }
  var score = 0;
  var gap = round1_(target - current);
  if (gap > 15) { score += 40; reasons.push("Currently " + current + "%, " + gap + " percentage points below the " + target + "% target."); }
  else if (gap > 5) { score += 25; reasons.push("Currently " + current + "%, " + gap + " percentage points below the " + target + "% target."); }
  else if (gap > 0) { score += 10; reasons.push("Currently " + current + "%, " + gap + " percentage points below the " + target + "% target -- close."); }
  else { reasons.push("Currently " + current + "% -- already at or above the " + target + "% target."); }

  if (typeof m.daysToNextAssessment === "number" && m.daysToNextAssessment >= 0) {
    if (m.daysToNextAssessment <= 7) { score += 30; reasons.push("Next assessment (" + (m.nextAssessmentLabel || "upcoming") + ") is in " + m.daysToNextAssessment + " day(s)."); }
    else if (m.daysToNextAssessment <= 14) { score += 15; reasons.push("Next assessment (" + (m.nextAssessmentLabel || "upcoming") + ") is in " + m.daysToNextAssessment + " days."); }
  }
  // v1.4.2 -- a tutorial is when this module's AF tests actually get
  // written, so it deserves its own (smaller, since tutorials are frequent
  // and lower-stakes individually than a formal A1/A2/A3) proximity boost,
  // separate from and additive to the exam-proximity one above.
  if (typeof m.daysToNextTutorial === "number" && m.daysToNextTutorial >= 0) {
    if (m.daysToNextTutorial <= 3) { score += 20; reasons.push("Next tutorial (" + (m.nextTutorialLabel || "upcoming") + ") is in " + m.daysToNextTutorial + " day(s) -- this is when this module's AF tests are written."); }
    else if (m.daysToNextTutorial <= 7) { score += 10; reasons.push("Next tutorial (" + (m.nextTutorialLabel || "upcoming") + ") is in " + m.daysToNextTutorial + " days."); }
  }
  if (m.repeated) { score += 15; reasons.push("Repeated module -- elevated monitoring applies regardless of current mark."); }
  if (typeof m.difficulty === "number" && m.difficulty >= 4) { score += 10; reasons.push("High conceptual difficulty (" + m.difficulty + "/5)."); }
  if (typeof m.workload === "number" && m.workload >= 4) { score += 5; reasons.push("High workload intensity (" + m.workload + "/5)."); }
  if (m.incompleteTaskCount > 0) { score += Math.min(m.incompleteTaskCount * 3, 15); reasons.push(m.incompleteTaskCount + " incomplete Study Task(s) for this module."); }
  if (m.weakTopicCount > 0) { score += Math.min(m.weakTopicCount * 5, 20); reasons.push(m.weakTopicCount + " weak Revision Tracker topic(s) for this module."); }
  if (Array.isArray(m.recentPracticeScores) && m.recentPracticeScores.length) {
    var avg = round1_(m.recentPracticeScores.reduce(function (s, v) { return s + v; }, 0) / m.recentPracticeScores.length);
    if (avg < target) { score += 10; reasons.push("Recorded practice/past-paper score(s) average " + avg + "%, below the " + target + "% target."); }
  }
  if (typeof m.requiredFutureMark === "number" && (m.requiredFutureMark - current) > 20) {
    score += 15; reasons.push("Required future mark (" + round1_(m.requiredFutureMark) + "%) is substantially above current performance.");
  }

  var category;
  if (score >= 70) category = "Critical";
  else if (score >= 50) category = "Very High";
  else if (score >= 30) category = "High";
  else if (score >= 15) category = "Maintain";
  else category = "Low";

  if (!reasons.length) reasons.push("No risk factors detected.");
  return { code: m.code, name: m.name, category: category, score: score, reasons: reasons };
}

// ------------------------------------------------------------------
// 2. RECOMMENDED WEEKLY STUDY ALLOCATION (spec §2)
// ------------------------------------------------------------------
/**
 * Pure. `prioritized` is the array ai_computePriority_ produced (one entry
 * per module, each also carrying minWeeklyHours/daysToNextAssessment copied
 * over by the caller -- see computeAcademicIntelligence_). Returns
 * { capacityHours, totalAllocated, allocations:[{code,name,hours,reasons}] }.
 * totalAllocated can never exceed capacityHours -- every step below only
 * ever gives up to `remaining`, so this is a structural guarantee, not just
 * a post-hoc check (also verified by a Node unit test -- see
 * V1.3.0_STATIC_VALIDATION.md).
 */
function ai_computeAllocation_(prioritized, capacityHours) {
  var cap = (typeof capacityHours === "number" && capacityHours > 0) ? capacityHours : 0;
  var reasons = {}, alloc = {};

  var incomplete = prioritized.filter(function (m) { return m.category === "Rules Incomplete" || m.category === "Insufficient data"; });
  incomplete.forEach(function (m) {
    alloc[m.code] = 0;
    reasons[m.code] = [m.category === "Rules Incomplete"
      ? "Module Rules are not Verified -- no allocation computed until rules are verified."
      : "Not enough recorded marks data to compute a priority-based allocation yet."];
  });

  var eligible = prioritized.filter(function (m) { return m.category !== "Rules Incomplete" && m.category !== "Insufficient data"; });
  var sorted = eligible.slice().sort(function (a, b) { return (b.score || 0) - (a.score || 0); });
  var remaining = cap;

  // Pass 1: each module's own configured minimum weekly hours, highest
  // priority first, never exceeding remaining capacity.
  sorted.forEach(function (m) {
    var floor = (typeof m.minWeeklyHours === "number" && m.minWeeklyHours > 0) ? m.minWeeklyHours : 0;
    var give = round1_(Math.max(0, Math.min(floor, remaining)));
    alloc[m.code] = give;
    reasons[m.code] = [];
    if (floor > 0) {
      reasons[m.code].push(give >= floor
        ? ("Configured minimum: " + floor + " hour(s)/week.")
        : ("Only " + give + " of its " + floor + "-hour configured minimum could be allocated -- weekly capacity is already committed to higher-priority modules."));
    }
    remaining = round1_(remaining - give);
  });

  // Pass 2: a module with an assessment inside 7 days can never be left at
  // zero hours without an explicit reason (spec §2's third requirement).
  sorted.forEach(function (m) {
    if (typeof m.daysToNextAssessment === "number" && m.daysToNextAssessment >= 0 && m.daysToNextAssessment <= 7 && alloc[m.code] === 0) {
      if (remaining >= 1) {
        alloc[m.code] = round1_(alloc[m.code] + 1);
        remaining = round1_(remaining - 1);
        reasons[m.code].push("Given 1 hour despite no configured minimum, because its next assessment is in " + m.daysToNextAssessment + " day(s).");
      } else {
        reasons[m.code].push("Kept at 0 hours even though its next assessment is in " + m.daysToNextAssessment + " day(s) -- weekly capacity is fully committed to higher-priority modules. Consider raising \"Weekly study capacity (hours)\" in Settings.");
      }
    }
  });

  // Pass 3: distribute whatever capacity remains, weighted by priority
  // score, in 0.5-hour steps, each module capped so no single module can
  // absorb the whole remaining budget (repeated-module status already
  // contributes to score via ai_computePriority_ -- it does NOT grant a
  // separate, unlimited hours boost here, per spec §2's second requirement).
  var ceilingPerModule = round1_(Math.max(3, cap * 0.4));
  var stepsAvailable = Math.round(remaining / 0.5);
  for (var i = 0; i < stepsAvailable; i++) {
    var candidates = sorted.filter(function (m) { return alloc[m.code] < ceilingPerModule; });
    if (!candidates.length) break;
    var best = null, bestRatio = -1;
    candidates.forEach(function (m) {
      var ratio = Math.max(m.score || 1, 1) / (alloc[m.code] + 0.25);
      if (ratio > bestRatio) { bestRatio = ratio; best = m; }
    });
    if (!best) break;
    alloc[best.code] = round1_(alloc[best.code] + 0.5);
    remaining = round1_(remaining - 0.5);
    reasons[best.code].push("+0.5 hour from remaining weekly capacity, allocated by priority score (" + best.score + ").");
  }

  sorted.forEach(function (m) {
    if (!reasons[m.code].length) reasons[m.code].push("Baseline only -- no configured minimum and not currently elevated priority.");
  });

  var allocations = prioritized.map(function (m) {
    return { code: m.code, name: m.name, hours: alloc[m.code] || 0, category: m.category, reasons: reasons[m.code] || [] };
  });
  var totalAllocated = round1_(allocations.reduce(function (s, a) { return s + a.hours; }, 0));
  return { capacityHours: cap, totalAllocated: totalAllocated, allocations: allocations };
}

// ------------------------------------------------------------------
// 3. RESOURCE-AWARE RULES (spec §5) -- feeds both the Resource insights
// panel and the "What should I do next?" engine below.
// ------------------------------------------------------------------
var AI_RESOURCE_FLAG_MEANING = {
  "Past paper": { attempted: "A", marked: "B", reviewed: "C", redone: "D" },
  "Tutorial": { attempted: "A", marked: "B", reviewed: "C", redone: "D" },
  "Practice set": { attempted: "A", marked: "B", reviewed: "C", redone: "D" },
  "Formula bible": { reviewed: "A", recallTested: "C" },
  "Lecture": { watched: "A", understood: "B", practiceCompleted: "C" },
  "Lecture video": { watched: "A", understood: "B", practiceCompleted: "C" }
};

/** Pure. `resources` and `passMarkByModule` are plain arrays/maps already
 *  loaded by api_getPlannerData. Returns an array of
 *  { module, resourceId, title, rule, recommendation, reasons }. */
function ai_computeResourceInsights_(resources, passMarkByModule) {
  var out = [];
  var byModTopicScored = {};

  resources.forEach(function (r) {
    var flag = AI_RESOURCE_FLAG_MEANING[r.type];
    if (flag && flag.attempted) {
      var attempted = r["flag" + flag.attempted];
      var marked = flag.marked ? r["flag" + flag.marked] : null;
      var reviewed = flag.reviewed ? r["flag" + flag.reviewed] : null;
      var redone = flag.redone ? r["flag" + flag.redone] : null;
      if (attempted && marked === false) {
        out.push({ module: r.module, resourceId: r.resourceId, title: r.title, rule: "attempted_not_marked",
          recommendation: "Mark and review the " + r.module + " \"" + r.title + "\" " + r.type.toLowerCase() + ".",
          reasons: ["\"" + r.title + "\" has been attempted but not yet marked."] });
      } else if (marked && reviewed === false) {
        out.push({ module: r.module, resourceId: r.resourceId, title: r.title, rule: "marked_not_reviewed",
          recommendation: "Review mistakes from the " + r.module + " \"" + r.title + "\" " + r.type.toLowerCase() + ".",
          reasons: ["\"" + r.title + "\" has been marked but mistakes have not been reviewed yet."] });
      } else if (reviewed && redone === false) {
        out.push({ module: r.module, resourceId: r.resourceId, title: r.title, rule: "reviewed_not_redone",
          recommendation: "Redo the failed questions from the " + r.module + " \"" + r.title + "\" " + r.type.toLowerCase() + ".",
          reasons: ["Mistakes have been reviewed on \"" + r.title + "\" but the failed questions have not been redone yet."] });
      }
    }
    if (r.type === "Formula bible" && r.flagA && r.flagC === false) {
      out.push({ module: r.module, resourceId: r.resourceId, title: r.title, rule: "formula_not_recall_tested",
        recommendation: "Recall-test the " + r.module + " formula bible (\"" + r.title + "\") without notes.",
        reasons: ["\"" + r.title + "\" has been reviewed but not recall-tested."] });
    }
    if ((r.type === "Past paper" || r.type === "Practice set") && typeof r.score === "number") {
      var key = r.module + "||" + (r.topic || "");
      (byModTopicScored[key] = byModTopicScored[key] || []).push(r);
    }
  });

  // Lecture-completion-high / practice-low, per module.
  var byModule = {};
  resources.forEach(function (r) { (byModule[r.module] = byModule[r.module] || []).push(r); });
  Object.keys(byModule).forEach(function (mod) {
    var list = byModule[mod];
    var lectures = list.filter(function (r) { return r.type === "Lecture" || r.type === "Lecture video"; });
    var lecturesUnderstood = lectures.filter(function (r) { return r.flagB; }).length;
    var practiceItems = list.filter(function (r) { return r.type === "Practice set" || r.type === "Past paper"; });
    var practiceAttempted = practiceItems.filter(function (r) { return r.flagA; }).length;
    if (lectures.length >= 2 && lecturesUnderstood >= Math.ceil(lectures.length * 0.6) && practiceItems.length > 0 && practiceAttempted === 0) {
      out.push({ module: mod, resourceId: "", title: "", rule: "lecture_high_practice_low",
        recommendation: "Start practice questions for " + mod + " -- lecture material is well covered but no practice has been attempted yet.",
        reasons: [lecturesUnderstood + " of " + lectures.length + " " + mod + " lecture resource(s) marked understood, but 0 of " + practiceItems.length + " practice/past-paper resource(s) attempted."] });
    }
  });

  // Repeated low scores on one module+topic.
  Object.keys(byModTopicScored).forEach(function (key) {
    var list = byModTopicScored[key];
    if (list.length < 2) return;
    var parts = key.split("||");
    var passMark = (passMarkByModule && typeof passMarkByModule[parts[0]] === "number") ? passMarkByModule[parts[0]] : 50;
    var lowCount = list.filter(function (r) { return r.score < passMark; }).length;
    if (lowCount >= 2) {
      out.push({ module: parts[0], resourceId: "", title: "", rule: "repeated_low_scores",
        recommendation: "Give targeted practice to " + parts[0] + (parts[1] ? " -- " + parts[1] : "") + " -- repeated low scores recorded.",
        reasons: [lowCount + " of " + list.length + " scored " + (parts[1] || "resource") + " item(s) for " + parts[0] + " are below the " + passMark + "% pass mark."] });
    }
  });

  return out;
}

// ------------------------------------------------------------------
// 4. "WHAT SHOULD I DO NEXT?" (spec §3)
// ------------------------------------------------------------------
/**
 * Pure. Picks exactly one highest-impact action, always tied to a real
 * Study Task, Resource, Revision topic, or Assessment -- never a vague
 * "study more" message. Returns { text, why:[...], ref:{type,id}|null }.
 */
function ai_computeNextAction_(ctx) {
  var byModuleName = {};
  (ctx.prioritized || []).forEach(function (p) { byModuleName[p.name] = p; });
  var rank = { "Critical": 6, "Very High": 5, "High": 4, "Maintain": 3, "Low": 2, "Insufficient data": 1, "Rules Incomplete": 0 };
  function priorityRank(moduleName) { var p = byModuleName[moduleName]; return p ? (rank[p.category] || 0) : -1; }

  var urgentInsight = (ctx.resourceInsights || [])
    .filter(function (r) { return r.rule === "attempted_not_marked" || r.rule === "marked_not_reviewed"; })
    .filter(function (r) { var p = byModuleName[r.module]; return p && (p.category === "Critical" || p.category === "Very High"); })
    .sort(function (a, b) { return priorityRank(b.module) - priorityRank(a.module); })[0];
  if (urgentInsight) return { text: urgentInsight.recommendation, why: urgentInsight.reasons.concat(["Selected first: tied to a Critical/Very High priority module."]), ref: { type: "resource", id: urgentInsight.resourceId } };

  var weakTopicKeys = {};
  (ctx.revisionTracker || []).forEach(function (r) { if (r.weakTopic) weakTopicKeys[r.module + "||" + r.topic] = true; });
  var weakTask = (ctx.studyTasks || [])
    .filter(function (t) { return !t.completed && weakTopicKeys[t.module + "||" + t.topic]; })
    .sort(function (a, b) { return priorityRank(b.module) - priorityRank(a.module); })[0];
  if (weakTask) {
    return { text: "Complete \"" + weakTask.text + "\" (" + weakTask.module + (weakTask.topic ? " -- " + weakTask.topic : "") + ").",
      why: ["This task is linked to a topic flagged Weak in Revision Tracker.", weakTask.module + " is currently priority \"" + ((byModuleName[weakTask.module] || {}).category || "unknown") + "\"."],
      ref: { type: "studyTask", id: weakTask.id } };
  }

  var anyUnmarked = (ctx.resourceInsights || []).filter(function (r) { return r.rule === "attempted_not_marked"; })[0];
  if (anyUnmarked) return { text: anyUnmarked.recommendation, why: anyUnmarked.reasons, ref: { type: "resource", id: anyUnmarked.resourceId } };

  var anyUnreviewed = (ctx.resourceInsights || []).filter(function (r) { return r.rule === "marked_not_reviewed"; })[0];
  if (anyUnreviewed) return { text: anyUnreviewed.recommendation, why: anyUnreviewed.reasons, ref: { type: "resource", id: anyUnreviewed.resourceId } };

  var recallDue = (ctx.resourceInsights || []).filter(function (r) { return r.rule === "formula_not_recall_tested"; })
    .sort(function (a, b) { return priorityRank(b.module) - priorityRank(a.module); })[0];
  if (recallDue) return { text: recallDue.recommendation, why: recallDue.reasons, ref: { type: "resource", id: recallDue.resourceId } };

  var soonAssessment = (ctx.prioritized || []).filter(function (p) { return typeof p.daysToNextAssessment === "number" && p.daysToNextAssessment >= 0 && p.daysToNextAssessment <= 3; })
    .sort(function (a, b) { return (a.daysToNextAssessment || 0) - (b.daysToNextAssessment || 0); })[0];
  if (soonAssessment) {
    return { text: "Review formulas and do timed practice for the " + soonAssessment.name + " assessment in " + soonAssessment.daysToNextAssessment + " day(s) (" + (soonAssessment.nextAssessmentLabel || "upcoming") + ").",
      why: ["An assessment for " + soonAssessment.name + " is in " + soonAssessment.daysToNextAssessment + " day(s) -- the nearest of any module."],
      ref: { type: "assessment", id: soonAssessment.nextAssessmentId || null } };
  }

  var anyTask = (ctx.studyTasks || []).filter(function (t) { return !t.completed; })
    .sort(function (a, b) { return priorityRank(b.module) - priorityRank(a.module); })[0];
  if (anyTask) {
    return { text: "Complete \"" + anyTask.text + "\" (" + anyTask.module + ").",
      why: [anyTask.module + " is currently priority \"" + ((byModuleName[anyTask.module] || {}).category || "unknown") + "\"."],
      ref: { type: "studyTask", id: anyTask.id } };
  }

  return { text: "Nothing urgent right now -- no incomplete Study Tasks, no unmarked or unreviewed resources, and no assessment within 3 days.",
    why: ["This reflects the current state of Study Tasks, Resources, and Assessments -- there is nothing outstanding to act on."],
    ref: null };
}

// ------------------------------------------------------------------
// 5. EXAM FOCUS MODE (spec §4) -- trigger detection is pure; activation is
// the only part of this whole release that writes anything, and only ever
// after an explicit, separate confirmation call from the frontend.
// ------------------------------------------------------------------
/** Pure. Returns { eligible:boolean, reasons:[...] , ...context } for one
 *  module. Never activates anything -- this only decides whether the Exam
 *  Focus card should be OFFERED, per spec §4's "never activate
 *  automatically." */
function ai_computeExamFocusTriggers_(m) {
  if (typeof m.daysToNextAssessment !== "number" || m.daysToNextAssessment < 0 || m.daysToNextAssessment > 7) {
    return { eligible: false, module: m.name, reasons: [] };
  }
  var reasons = [];
  if (typeof m.latestRelevantScore === "number" && typeof m.targetMark === "number" && m.latestRelevantScore < m.targetMark) {
    reasons.push("Latest relevant practice/past-paper score (" + round1_(m.latestRelevantScore) + "%) is below the " + round1_(m.targetMark) + "% target.");
  }
  if (m.weakTopicCount > 0 && typeof m.weakTopicSinceDays === "number" && m.weakTopicSinceDays >= 7) {
    reasons.push(m.weakTopicCount + " weak Revision Tracker topic(s), flagged weak for " + m.weakTopicSinceDays + "+ day(s) (repeated weak topic).");
  }
  if (typeof m.requiredFutureMark === "number" && typeof m.currentMark === "number" && (m.requiredFutureMark - m.currentMark) > 20) {
    reasons.push("Required future mark (" + round1_(m.requiredFutureMark) + "%) is substantially above current performance (" + round1_(m.currentMark) + "%).");
  }
  if (m.passiveCount >= 2 && m.practiceCount === 0) {
    reasons.push(m.passiveCount + " passive review session(s) logged in the last 7 days with 0 practice/timed-question sessions, this close to the assessment.");
  }
  if (!reasons.length) return { eligible: false, module: m.name, reasons: [] };
  return { eligible: true, module: m.name, assessment: m.nextAssessmentLabel, assessmentDate: m.nextAssessmentDate,
    daysRemaining: m.daysToNextAssessment, latestScore: m.latestRelevantScore, target: m.targetMark, reasons: reasons };
}

/** Fixed, hand-authored, deterministic Exam Focus study-type priority order
 *  (spec §4's "prioritise formula recall; weak-topic practice; timed
 *  questions; past papers; marking; mistake review; redoing failed
 *  questions" -- reusing the exact Study types/Task types this workbook
 *  already has, never inventing new ones). */
var EXAM_FOCUS_STUDY_TYPES = ["Formula", "Practice", "Exam", "Mistake"];

function readExamFocusLog_() {
  return readSheetRows_(EXAM_FOCUS_SHEET, HEADER_ROW, "Module");
}

function findActiveExamFocusRow_(sheet, moduleName) {
  var map = getColMap_(sheet);
  var lastRow = sheet.getLastRow();
  if (lastRow <= HEADER_ROW) return -1;
  var vals = sheet.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, sheet.getLastColumn()).getValues();
  for (var i = 0; i < vals.length; i++) {
    var row = vals[i];
    if (row[col_(map, "Module") - 1] === moduleName && row[col_(map, "Status") - 1] === "Active" && !row[col_(map, "Deactivated date") - 1]) {
      return HEADER_ROW + 1 + i;
    }
  }
  return -1;
}

/**
 * Builds the proposed plan for review -- writes nothing. `existingSessions`
 * are this module's own Planned sessions in the next 7 days (never another
 * module's). Passive types (Preview/Learn/Visualization/Synthesis/Weekly
 * Review) are proposed as "skip candidates"; everything else is left alone.
 * New sessions are proposed from EXAM_FOCUS_STUDY_TYPES, skipping any type
 * that isn't relevant given this module's actual Revision Tracker/Resource
 * state (e.g. no Formula-recall proposal if there is no Formula bible
 * resource for this module at all).
 */
function api_previewExamFocusPlan(moduleName) {
  try {
    if (!moduleName) return fail_(new Error("No module given."));
    var studySheet = SpreadsheetApp.getActive().getSheetByName(STUDY_SHEET);
    var studyRows = readSheetRows_(STUDY_SHEET, HEADER_ROW, "Module");
    var today = new Date(); today.setHours(0, 0, 0, 0);
    var horizon = new Date(today.getTime() + 7 * 86400000);

    var candidateSkips = studyRows.filter(function (r) {
      if (r["Module"] !== moduleName) return false;
      var status = r["Status"] || (r["Completed"] === true ? "Completed" : "Planned");
      if (status !== "Planned") return false;
      var d = r["Date"];
      if (!(d instanceof Date) || d < today || d > horizon) return false;
      var passiveTypes = ["Preview", "Learn", "Visualization", "Synthesis", "Weekly Review"];
      return passiveTypes.indexOf(r["Study type"]) !== -1;
    }).map(function (r) {
      return { sessionId: r["Study Session ID"], label: (r["Study type"] || "Session") + (r["Topic"] ? " -- " + r["Topic"] : "") + " on " + isoDate_(r["Date"]) };
    });

    var keptSessions = studyRows.filter(function (r) {
      if (r["Module"] !== moduleName) return false;
      var status = r["Status"] || (r["Completed"] === true ? "Completed" : "Planned");
      if (status !== "Planned") return false;
      var d = r["Date"];
      if (!(d instanceof Date) || d < today || d > horizon) return false;
      return ["Preview", "Learn", "Visualization", "Synthesis", "Weekly Review"].indexOf(r["Study type"]) === -1;
    }).map(function (r) { return { sessionId: r["Study Session ID"], label: (r["Study type"] || "Session") + (r["Topic"] ? " -- " + r["Topic"] : "") + " on " + isoDate_(r["Date"]) }; });

    var revisionRows = readSheetRows_(REVISION_SHEET, HEADER_ROW, "Module").filter(function (r) { return r["Module"] === moduleName; });
    var weakTopics = revisionRows.filter(function (r) { return r["Weak topic (TRUE/FALSE)"] === true; }).map(function (r) { return r["Topic"]; });
    var resourceRows = readSheetRows_(RESOURCES_SHEET, HEADER_ROW, "Title").filter(function (r) { return r["Module"] === moduleName; });
    var hasFormulaBible = resourceRows.some(function (r) { return r["Resource type"] === "Formula bible"; });
    var hasPastPaper = resourceRows.some(function (r) { return r["Resource type"] === "Past paper"; });

    var proposedNewSessions = [];
    if (weakTopics.length) {
      proposedNewSessions.push({ studyType: "Practice", topic: weakTopics[0], reason: "Weak-topic practice for \"" + weakTopics[0] + "\" (flagged Weak in Revision Tracker)." });
    }
    if (hasFormulaBible) {
      proposedNewSessions.push({ studyType: "Formula", topic: "", reason: "Formula recall -- a Formula bible resource exists for " + moduleName + "." });
    }
    if (hasPastPaper) {
      proposedNewSessions.push({ studyType: "Exam", topic: "", reason: "Timed past-paper practice -- a Past paper resource exists for " + moduleName + "." });
    }
    proposedNewSessions.push({ studyType: "Mistake", topic: weakTopics[0] || "", reason: "Mistake review and redo of failed questions." });

    return ok_({
      module: moduleName,
      keptSessions: keptSessions,
      candidateSkipSessions: candidateSkips,
      proposedNewSessions: proposedNewSessions
    });
  } catch (e) { return fail_(e); }
}

/**
 * The ONLY function in this entire release that changes anything for Exam
 * Focus Mode, and only ever in response to an explicit user click on
 * "Activate" in the frontend's preview modal (spec §4: "never activate
 * automatically... require explicit user confirmation"). `approvedSkipIds`
 * is the exact list of Study Session IDs the user ticked to skip in the
 * preview -- every one of them is set to Status = "Skipped" (never
 * deleted -- spec §4: "do not delete existing sessions silently"). Every
 * other Planned session, in this module or any other, is left completely
 * untouched. `approvedNewSessions` is the exact list the user approved from
 * the preview (a subset of, never more than, what api_previewExamFocusPlan
 * proposed) -- each becomes a real Study Planner session via the same
 * appendStudySessionRow_ every other v1.2.0 session-creation path already
 * uses, with Origin = "Exam Focus" and its own concrete starter tasks via
 * addStarterTasksForSession_ (spec §7: never a vague "Study Module" entry).
 */
function api_activateExamFocus(moduleName, approvedSkipIds, approvedNewSessions) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(LOCK_WAIT_MS)) return fail_(new Error("Workbook is busy — try again in a moment."));
    if (!moduleName) return fail_(new Error("No module given."));

    var skippedCount = 0;
    (approvedSkipIds || []).forEach(function (sessionId) {
      var row = findRowByStableId_(STUDY_SHEET, "Study Session ID", sessionId);
      if (row === -1) return;
      var sheet = SpreadsheetApp.getActive().getSheetByName(STUDY_SHEET);
      var map = getColMap_(sheet);
      if (sheet.getRange(row, col_(map, "Module")).getValue() !== moduleName) return; // never touch another module's row
      if (map["Status"]) sheet.getRange(row, col_(map, "Status")).setValue("Skipped");
      skippedCount++;
    });

    var createdIds = [];
    (approvedNewSessions || []).forEach(function (item) {
      if (!item || VALID_STUDY_TYPES.indexOf(item.studyType) === -1) return;
      var result = appendStudySessionRow_({
        module: moduleName, date: new Date(), studyType: item.studyType, topic: item.topic || "",
        priority: "High", notes: "Exam Focus Mode -- " + (item.reason || ""),
        planned: true, completed: false, status: "Planned", origin: "Exam Focus"
      });
      addStarterTasksForSession_(result.id, moduleName, item.topic || "", item.studyType);
      createdIds.push(result.id);
    });

    var sheet = SpreadsheetApp.getActive().getSheetByName(EXAM_FOCUS_SHEET);
    if (!sheet) return fail_(new Error('Sheet not found: "' + EXAM_FOCUS_SHEET + '" — run Academic_Intelligence_v1.3.0_Migration.gs first.'));
    var existingActiveRow = findActiveExamFocusRow_(sheet, moduleName);
    if (existingActiveRow === -1) {
      var map = getColMap_(sheet);
      var row = sheet.getLastRow() + 1;
      if (row <= HEADER_ROW) row = HEADER_ROW + 1;
      sheet.getRange(row, col_(map, "Module")).setValue(moduleName);
      sheet.getRange(row, col_(map, "Trigger reasons")).setValue((approvedNewSessions || []).map(function (i) { return i.reason; }).join(" | "));
      sheet.getRange(row, col_(map, "Activated date")).setValue(new Date());
      sheet.getRange(row, col_(map, "Status")).setValue("Active");
    }

    logAutomation_("Exam Focus Mode activated", moduleName, "Activated", skippedCount + " session(s) skipped, " + createdIds.length + " new session(s) created");
    return ok_({ module: moduleName, skippedCount: skippedCount, createdSessionIds: createdIds });
  } catch (e) { return fail_(e); } finally { lock.releaseLock(); }
}

/** Exits Exam Focus Mode for a module. Only ever writes the Deactivated date
 *  on its "22 Exam Focus Log" row -- never touches a single Study Planner
 *  session (sessions created while Exam Focus was active simply remain,
 *  exactly like any other session, editable/deletable the normal way). */
function api_exitExamFocus(moduleName) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(LOCK_WAIT_MS)) return fail_(new Error("Workbook is busy — try again in a moment."));
    var sheet = SpreadsheetApp.getActive().getSheetByName(EXAM_FOCUS_SHEET);
    if (!sheet) return fail_(new Error('Sheet not found: "' + EXAM_FOCUS_SHEET + '" — run Academic_Intelligence_v1.3.0_Migration.gs first.'));
    var row = findActiveExamFocusRow_(sheet, moduleName);
    if (row === -1) return fail_(new Error("Exam Focus Mode is not currently active for " + moduleName + "."));
    var map = getColMap_(sheet);
    sheet.getRange(row, col_(map, "Deactivated date")).setValue(new Date());
    sheet.getRange(row, col_(map, "Status")).setValue("Exited");
    logAutomation_("Exam Focus Mode exited", moduleName, "Exited", "");
    return ok_({});
  } catch (e) { return fail_(e); } finally { lock.releaseLock(); }
}

// ------------------------------------------------------------------
// 6. IMPROVEMENT TRACKER (spec §6) -- every metric requires >=2 comparable,
// dated data points already present in the workbook, or it is simply
// omitted (never padded, never inferred).
// ------------------------------------------------------------------
/** Counts Revision Tracker rows that are both flagged Weak and overdue
 *  (Next revision date in the past, relative to todayIso). Pure. */
function computeOverdueWeakCount_(revisionTracker, todayIso) {
  return (revisionTracker || []).filter(function (r) {
    if (!r.weakTopic || !r.nextRevision) return false;
    var d = daysBetween_(r.nextRevision, todayIso);
    return typeof d === "number" && d > 0;
  }).length;
}

/**
 * Pure. `bundle` carries the already-loaded v1.2.0 arrays plus `logSnapshots`
 * (Automation Log rows whose Detail starts with "OverdueWeakCount:" --
 * see maybeLogOverdueWeakSnapshot_ below for how those get created, always
 * from a real computed count, never fabricated) and `logRows` (the full
 * recent Automation Log, for the study-streak metric). Returns an array of
 * { key, category, module, message, direction, evidenceCount } -- one entry
 * per metric that actually had >=2 comparable points.
 */
function ai_computeImprovementTracker_(bundle) {
  var metrics = [];

  // 1. Mark trend per module (AF Components entries, each already dated).
  var byModuleMarks = {};
  (bundle.afComponents || []).forEach(function (a) {
    if (typeof a.pct !== "number" || !a.date) return;
    (byModuleMarks[a.module] = byModuleMarks[a.module] || []).push({ date: a.date, pct: a.pct, label: a.itemName });
  });
  Object.keys(byModuleMarks).forEach(function (mod) {
    var list = byModuleMarks[mod].sort(function (a, b) { return a.date < b.date ? -1 : (a.date > b.date ? 1 : 0); });
    if (list.length < 2) return;
    var first = list[0], last = list[list.length - 1];
    metrics.push({ key: "mark_trend_" + mod, category: "Mark trend", module: mod,
      message: "Your " + mod + " AF marks moved from " + round1_(first.pct) + "% (" + first.label + ") to " + round1_(last.pct) + "% (" + last.label + ") across " + list.length + " recorded entries.",
      direction: last.pct > first.pct ? "up" : (last.pct < first.pct ? "down" : "flat"), evidenceCount: list.length });
  });

  // 2. Past-paper/practice score trend -- only resources with a real
  // "Score updated" date (see Academic_Intelligence_v1.3.0_Migration.gs).
  var byModTopicScored = {};
  (bundle.resources || []).forEach(function (r) {
    if ((r.type !== "Past paper" && r.type !== "Practice set") || typeof r.score !== "number" || !r.scoreUpdated) return;
    var key = r.module + "||" + (r.topic || "General");
    (byModTopicScored[key] = byModTopicScored[key] || []).push({ date: r.scoreUpdated, score: r.score });
  });
  Object.keys(byModTopicScored).forEach(function (key) {
    var list = byModTopicScored[key].sort(function (a, b) { return a.date < b.date ? -1 : (a.date > b.date ? 1 : 0); });
    if (list.length < 2) return;
    var parts = key.split("||");
    var scoresStr = list.map(function (x) { return round1_(x.score) + "%"; }).join(" to ");
    metrics.push({ key: "score_trend_" + key, category: "Past-paper / practice score trend", module: parts[0],
      message: "Your last " + list.length + " " + parts[0] + (parts[1] !== "General" ? " " + parts[1] : "") + " practice scores moved from " + scoresStr + ".",
      direction: list[list.length - 1].score > list[0].score ? "up" : (list[list.length - 1].score < list[0].score ? "down" : "flat"), evidenceCount: list.length });
  });

  // 3. Topic confidence trend -- Study Sessions' own confidenceAfter, by
  // module+topic, only from Completed/Partly completed sessions.
  var byModTopicConf = {};
  (bundle.studySessions || []).forEach(function (s) {
    if (typeof s.confidenceAfter !== "number" || !s.date || !s.topic) return;
    if (s.status !== "Completed" && s.status !== "Partly completed") return;
    var key = s.module + "||" + s.topic;
    (byModTopicConf[key] = byModTopicConf[key] || []).push({ date: s.date, conf: s.confidenceAfter });
  });
  Object.keys(byModTopicConf).forEach(function (key) {
    var list = byModTopicConf[key].sort(function (a, b) { return a.date < b.date ? -1 : (a.date > b.date ? 1 : 0); });
    if (list.length < 2) return;
    var parts = key.split("||");
    metrics.push({ key: "confidence_trend_" + key, category: "Topic confidence trend", module: parts[0],
      message: "Confidence for " + parts[0] + " -- " + parts[1] + " moved from " + list[0].conf + "/100 to " + list[list.length - 1].conf + "/100 across " + list.length + " completed sessions.",
      direction: list[list.length - 1].conf > list[0].conf ? "up" : (list[list.length - 1].conf < list[0].conf ? "down" : "flat"), evidenceCount: list.length });
  });

  // 4. Planned-session completion trend, weekly, last up to 4 weeks with data.
  var byWeek = {};
  (bundle.studySessions || []).forEach(function (s) {
    if (!s.date) return;
    var wk = isoWeekKey_(s.date);
    if (!wk) return;
    byWeek[wk] = byWeek[wk] || { planned: 0, completed: 0 };
    byWeek[wk].planned++;
    if (s.status === "Completed" || s.status === "Partly completed") byWeek[wk].completed++;
  });
  var weekKeys = Object.keys(byWeek).sort();
  if (weekKeys.length >= 2) {
    var recentWeeks = weekKeys.slice(-4);
    var ratesStr = recentWeeks.map(function (wk) { var d = byWeek[wk]; return (d.planned ? Math.round((d.completed / d.planned) * 100) : 0) + "%"; }).join(" to ");
    metrics.push({ key: "session_completion_trend", category: "Planned-session completion", module: null,
      message: "Weekly planned-session completion moved " + ratesStr + " across the last " + recentWeeks.length + " week(s) with recorded sessions.",
      direction: null, evidenceCount: recentWeeks.length });
  }

  // 5. Completed Study Tasks trend, weekly, last up to 4 weeks.
  var tasksByWeek = {};
  (bundle.studyTasks || []).forEach(function (t) {
    if (!t.completed || !t.lastUpdated) return;
    var wk = isoWeekKey_(t.lastUpdated);
    if (!wk) return;
    tasksByWeek[wk] = (tasksByWeek[wk] || 0) + 1;
  });
  var taskWeeks = Object.keys(tasksByWeek).sort();
  if (taskWeeks.length >= 2) {
    var recentTaskWeeks = taskWeeks.slice(-4);
    var countsStr = recentTaskWeeks.map(function (wk) { return tasksByWeek[wk]; }).join(" to ");
    metrics.push({ key: "tasks_completed_trend", category: "Completed Study Tasks", module: null,
      message: "Completed Study Tasks per week moved " + countsStr + " across the last " + recentTaskWeeks.length + " week(s).",
      direction: null, evidenceCount: recentTaskWeeks.length });
  }

  // 6. Overdue weak-topic reduction -- from real snapshots logged in 20
  // Automation Log (see maybeLogOverdueWeakSnapshot_), never fabricated.
  var snapshots = (bundle.logSnapshots || []).slice();
  if (snapshots.length >= 2) {
    var oldest = snapshots[snapshots.length - 1], newest = snapshots[0]; // logRows are newest-first
    metrics.push({ key: "overdue_weak_topics", category: "Overdue weak topics", module: null,
      message: "Overdue weak Revision Tracker topics moved from " + oldest.count + " to " + newest.count + " across " + snapshots.length + " recorded checkpoints.",
      direction: newest.count < oldest.count ? "up" : (newest.count > oldest.count ? "down" : "flat"), evidenceCount: snapshots.length });
  }

  // 7. Consistent study streaks -- consecutive distinct days with at least
  // one relevant Automation Log entry.
  var studyDays = {};
  (bundle.logRows || []).forEach(function (l) {
    if (!l.action) return;
    if (l.action.indexOf("Study session") === 0 || l.action.indexOf("Independent study") === 0 || l.action.indexOf("Study task") === 0) {
      var day = String(l.timestamp).slice(0, 10);
      if (day) studyDays[day] = true;
    }
  });
  var dayKeys = Object.keys(studyDays).sort();
  if (dayKeys.length >= 2) {
    var streak = 1, best = 1;
    for (var i = 1; i < dayKeys.length; i++) {
      var diff = daysBetween_(dayKeys[i - 1], dayKeys[i]);
      if (diff === 1) { streak++; best = Math.max(best, streak); } else { streak = 1; }
    }
    if (best >= 2) {
      metrics.push({ key: "study_streak", category: "Study streak", module: null,
        message: "Your longest current study streak (consecutive days with at least one recorded study action) is " + best + " day(s), from " + dayKeys.length + " recorded study day(s) total.",
        direction: "up", evidenceCount: dayKeys.length });
    }
  }

  return metrics;
}

/**
 * Logs a fresh "OverdueWeakCount:N" snapshot to 20 Automation Log ONLY when
 * the count has genuinely changed since the last logged snapshot -- so the
 * Improvement Tracker's evidence trail records real transitions, not one
 * row per page load. Best-effort: reuses logAutomation_'s own try/catch, so
 * a failure here can never break api_getPlannerData(). Does not touch
 * Revision Tracker or any revision-contribution logic at all.
 */
function maybeLogOverdueWeakSnapshot_(currentCount, recentLogRows) {
  try {
    var lastSnapshot = null;
    for (var i = 0; i < recentLogRows.length; i++) { // logRows are newest-first
      var detail = recentLogRows[i].detail || recentLogRows[i]["Detail"];
      if (typeof detail === "string" && detail.indexOf("OverdueWeakCount:") === 0) {
        lastSnapshot = parseInt(detail.split(":")[1], 10);
        break;
      }
    }
    if (lastSnapshot !== currentCount) {
      logAutomation_("Overdue weak-topic checkpoint", "-", "Recorded", "OverdueWeakCount:" + currentCount);
    }
  } catch (e) { /* best-effort only -- must never block a planner data read */ }
}

// ------------------------------------------------------------------
// ORCHESTRATOR -- called once from api_getPlannerData(). Reuses every array
// api_getPlannerData already built (no extra sheet reads for anything
// except "22 Exam Focus Log", which v1.2.0 never read). This is the ONLY
// function below that assembles per-module context objects; every ai_...
// function above stays a pure, independently testable unit.
// ------------------------------------------------------------------
function computeAcademicIntelligence_(bundle) {
  var todayIso = bundle.todayIso;
  var byModuleName = {};
  bundle.modules.forEach(function (m) { byModuleName[m.name] = m; });
  var marksByModule = {};
  bundle.marksTracker.forEach(function (r) { marksByModule[r.module] = r; });

  // v1.3.0 Marks Intelligence expansion -- computed early so its per-module
  // currentMark (which correctly incorporates a verified Estimated A2 when
  // A2 hasn't been published yet -- see mi_computeMarksIntelligence_ /
  // mi_inferHiddenA2_ above) can serve as the Priority Engine's fallback
  // "current standing" for a module that has real AF/A1/(A2) component data
  // but no recorded "Final / provisional (after A3)" yet in 11 Marks
  // Tracker (i.e. before A3 has happened). 11 Marks Tracker's own
  // finalOrProvisional value, when present, ALWAYS takes priority -- this is
  // purely a fallback for the gap where real component data exists but that
  // sheet-computed column is still blank, so a hidden-A2 module doesn't
  // fall into "Insufficient data" the moment its A2 hasn't been officially
  // published, when Marks Intelligence already has a verified estimate.
  var marksIntelligenceByCode = {};
  bundle.modules.filter(function (m) { return m.active; }).forEach(function (m) {
    marksIntelligenceByCode[m.code] = mi_computeMarksIntelligence_(m, marksByModule[m.name], m.passMark, m.distinctionMark, bundle.settings);
  });

  // Next assessment per module (never negative days -- only genuinely
  // upcoming assessments count).
  var nextAssessmentByModule = {};
  bundle.assessments.forEach(function (a) {
    if (!a.date) return;
    var d = daysBetween_(todayIso, a.date);
    if (d === null || d < 0) return;
    var existing = nextAssessmentByModule[a.module];
    if (!existing || d < existing.days) {
      nextAssessmentByModule[a.module] = { days: d, label: (a.type || "Assessment") + (a.title ? " -- " + a.title : ""), date: a.date, id: a.id };
    }
  });

  // v1.4.2 -- next scheduled Tutorial session per module (05 Timetable
  // Import), same "never negative, keep the soonest" shape as
  // nextAssessmentByModule above. bundle.timetableSessions is [] on a
  // workbook where 05 Timetable Import is empty or missing -- degrades to
  // "no tutorial signal" for every module, never a fabricated date.
  var nextTutorialByModule = {};
  (bundle.timetableSessions || []).forEach(function (s) {
    if (s.type !== "Tutorial" || !s.date) return;
    var d = daysBetween_(todayIso, s.date);
    if (d === null || d < 0) return;
    var existing = nextTutorialByModule[s.module];
    if (!existing || d < existing.days) {
      nextTutorialByModule[s.module] = { days: d, label: "Tutorial", date: s.date };
    }
  });

  var incompleteTasksByModule = {}, weakTopicsByModule = {}, weakSinceByModule = {};
  bundle.studyTasks.forEach(function (t) {
    if (!t.completed) incompleteTasksByModule[t.module] = (incompleteTasksByModule[t.module] || 0) + 1;
  });
  bundle.revisionTracker.forEach(function (r) {
    if (r.weakTopic) {
      weakTopicsByModule[r.module] = (weakTopicsByModule[r.module] || 0) + 1;
      if (r.weakTopicSince) {
        var days = daysBetween_(r.weakTopicSince, todayIso);
        if (typeof days === "number" && (!weakSinceByModule[r.module] || days > weakSinceByModule[r.module])) weakSinceByModule[r.module] = days;
      }
    }
  });

  var scoresByModule = {};
  bundle.resources.forEach(function (r) {
    if ((r.type === "Past paper" || r.type === "Practice set") && typeof r.score === "number") {
      (scoresByModule[r.module] = scoresByModule[r.module] || []).push(r);
    }
  });

  var reqFieldByTarget = function (marksRow, target) {
    if (!marksRow || typeof target !== "number") return null;
    var candidates = [{ t: 50, v: marksRow.req50 }, { t: 60, v: marksRow.req60 }, { t: 75, v: marksRow.req75 }, { t: 80, v: marksRow.req80 }]
      .filter(function (c) { return typeof pctToNumber_(c.v) === "number"; });
    if (!candidates.length) return null;
    candidates.sort(function (a, b) { return Math.abs(a.t - target) - Math.abs(b.t - target); });
    return pctToNumber_(candidates[0].v);
  };

  // v1.3.1 (2nd correction pass) -- the Priority Engine and Exam Focus
  // triggers use the module's REQUIRED A3 (spec: "required A3") as one of
  // their inputs. Where the module's weighting is verified, this is now the
  // BEST (lowest) REGULATION-VALID required A3 across whichever of FM2/FM3
  // structurally has an A3 component (FM1 never does under the corrected
  // route model -- see mi_computeA3Recovery_/mi_computeRequiredA3ForRoute_,
  // the exact same calculation the A3 Recovery Calculator screen shows, so
  // nothing is duplicated or computed two different ways) rather than the
  // sheet's fixed 50/60/75/80% columns. Falls back to reqFieldByTarget()
  // above only for a module whose rules aren't yet Verified.
  var requiredA3ByTarget = function (m, marksRow, target) {
    if (typeof target !== "number") return null;
    var weights = mi_getModuleWeights_(m);
    if (!weights.complete) return reqFieldByTarget(marksRow, target);
    var recovery = mi_computeA3Recovery_(m, marksRow, [{ label: "_", value: target }], bundle.settings);
    var routes = (recovery.targets[0] || { routes: [] }).routes;
    var valid = routes.filter(function (r) { return typeof r.regulationValidRequiredA3 === "number"; });
    if (!valid.length) return reqFieldByTarget(marksRow, target);
    return valid.reduce(function (a, b) { return b.regulationValidRequiredA3 < a.regulationValidRequiredA3 ? b : a; }).regulationValidRequiredA3;
  };

  // Passive vs practice session counts in the last 7 days, per module (Exam
  // Focus trigger input only).
  var passiveByModule = {}, practiceByModule = {};
  var sevenDaysAgo = todayIso; // computed relative to each session's own date via daysBetween_ below
  bundle.studySessions.forEach(function (s) {
    if (!s.date) return;
    var age = daysBetween_(s.date, todayIso);
    if (typeof age !== "number" || age < 0 || age > 7) return;
    var passiveTypes = ["Preview", "Learn", "Visualization", "Synthesis", "Weekly Review"];
    var practiceTypes = ["Practice", "Exam", "Mistake"];
    if (passiveTypes.indexOf(s.type) !== -1) passiveByModule[s.module] = (passiveByModule[s.module] || 0) + 1;
    if (practiceTypes.indexOf(s.type) !== -1) practiceByModule[s.module] = (practiceByModule[s.module] || 0) + 1;
  });

  // v1.4.2 -- "Include in study plan" (03 Modules) is a Study Planner-
  // specific filter, separate from "active" -- a module can be Active
  // (shows in Attendance/Calendar/etc.) but excluded here, or vice versa.
  var prioritized = bundle.modules.filter(function (m) { return m.active && m.includeInStudyPlan; }).map(function (m) {
    var marksRow = marksByModule[m.name];
    var nextA = nextAssessmentByModule[m.name];
    var nextT = nextTutorialByModule[m.name];
    var scores = (scoresByModule[m.name] || []).map(function (r) { return r.score; });
    // Prefer 11 Marks Tracker's own "Final / provisional (after A3)" value
    // when it's actually recorded; otherwise fall back to Marks
    // Intelligence's own currentMark (see the marksIntelligenceByCode
    // comment above) so a hidden-A2 module with real AF/A1/estimated-A2
    // signal doesn't get treated as having no data at all.
    var sheetCurrent = pctToNumber_(marksRow ? marksRow.finalOrProvisional : null);
    var miForModule = marksIntelligenceByCode[m.code];
    var usedEstimatedFallback = false;
    var currentForPriority = sheetCurrent;
    if (currentForPriority === null && miForModule && miForModule.status === "ok" && typeof miForModule.currentMark === "number") {
      currentForPriority = miForModule.currentMark;
      usedEstimatedFallback = true;
    }
    var input = {
      code: m.code, name: m.name, ruleStatus: m.ruleStatus,
      finalOrProvisional: currentForPriority,
      currentMarkIsEstimatedFallback: usedEstimatedFallback,
      targetMarkOverride: m.targetMark, settingsTargetMark: bundle.settings.targetMark,
      daysToNextAssessment: nextA ? nextA.days : null, nextAssessmentLabel: nextA ? nextA.label : null,
      daysToNextTutorial: nextT ? nextT.days : null, nextTutorialLabel: nextT ? nextT.label : null,
      repeated: m.repeated, difficulty: m.difficulty, workload: m.workload,
      incompleteTaskCount: incompleteTasksByModule[m.name] || 0, weakTopicCount: weakTopicsByModule[m.name] || 0,
      recentPracticeScores: scores, requiredFutureMark: requiredA3ByTarget(m, marksRow, (typeof m.targetMark === "number") ? m.targetMark : bundle.settings.targetMark)
    };
    var priority = ai_computePriority_(input);
    priority.minWeeklyHours = m.minWeeklyHours;
    priority.daysToNextAssessment = input.daysToNextAssessment;
    priority.nextAssessmentLabel = input.nextAssessmentLabel;
    priority.nextAssessmentId = nextA ? nextA.id : null;
    priority.daysToNextTutorial = input.daysToNextTutorial;
    priority.nextTutorialLabel = input.nextTutorialLabel;
    return priority;
  });

  var passMarkByModule = {};
  bundle.modules.forEach(function (m) { passMarkByModule[m.name] = (typeof m.passMark === "number") ? m.passMark : 50; });
  var resourceInsights = ai_computeResourceInsights_(bundle.resources, passMarkByModule);

  var nextAction = ai_computeNextAction_({ prioritized: prioritized, resourceInsights: resourceInsights, revisionTracker: bundle.revisionTracker, studyTasks: bundle.studyTasks });

  var weeklyAllocation = ai_computeAllocation_(prioritized, bundle.settings.weeklyStudyCapacity);

  var examFocusTriggers = bundle.modules.filter(function (m) { return m.active; }).map(function (m) {
    var scores = (scoresByModule[m.name] || []).map(function (r) { return r.score; });
    var latest = scores.length ? scores[scores.length - 1] : null;
    var marksRow = marksByModule[m.name];
    var target = (typeof m.targetMark === "number") ? m.targetMark : bundle.settings.targetMark;
    return ai_computeExamFocusTriggers_({
      name: m.name, daysToNextAssessment: nextAssessmentByModule[m.name] ? nextAssessmentByModule[m.name].days : null,
      nextAssessmentLabel: nextAssessmentByModule[m.name] ? nextAssessmentByModule[m.name].label : null,
      nextAssessmentDate: nextAssessmentByModule[m.name] ? nextAssessmentByModule[m.name].date : null,
      latestRelevantScore: latest, targetMark: target,
      weakTopicCount: weakTopicsByModule[m.name] || 0, weakTopicSinceDays: weakSinceByModule[m.name] || null,
      requiredFutureMark: requiredA3ByTarget(m, marksRow, target), currentMark: pctToNumber_(marksRow ? marksRow.finalOrProvisional : null),
      passiveCount: passiveByModule[m.name] || 0, practiceCount: practiceByModule[m.name] || 0
    });
  }).filter(function (t) { return t.eligible; });

  var overdueWeakCount = computeOverdueWeakCount_(bundle.revisionTracker, todayIso);
  maybeLogOverdueWeakSnapshot_(overdueWeakCount, bundle.logRows);
  var logSnapshots = (bundle.logRows || [])
    .filter(function (l) { return typeof l.detail === "string" && l.detail.indexOf("OverdueWeakCount:") === 0; })
    .map(function (l) { return { timestamp: l.timestamp, count: parseInt(l.detail.split(":")[1], 10) }; })
    .filter(function (x) { return isFinite(x.count); });

  var improvementTracker = ai_computeImprovementTracker_({
    afComponents: bundle.afComponents, resources: bundle.resources, studySessions: bundle.studySessions,
    studyTasks: bundle.studyTasks, logRows: bundle.logRows, logSnapshots: logSnapshots
  });

  var examFocusActiveRows = readExamFocusLog_().filter(function (r) { return r["Status"] === "Active" && !r["Deactivated date"]; })
    .map(function (r) { return { module: r["Module"], activatedDate: isoDate_(r["Activated date"]), triggerReasons: r["Trigger reasons"] || "" }; });

  // Reuses the SAME marksIntelligenceByCode entries computed at the top of
  // this function (never recomputed a second time) -- just attaches the
  // matching Priority Engine result and, for a fully-computed module, its
  // A3 Recovery Calculator targets.
  var marksIntelligence = bundle.modules.filter(function (m) { return m.active; }).map(function (m) {
    var marksRow = marksByModule[m.name];
    var mi = marksIntelligenceByCode[m.code];
    var matchingPriority = prioritized.filter(function (p) { return p.code === m.code; })[0];
    mi.priority = matchingPriority ? { category: matchingPriority.category, score: matchingPriority.score } : null;
    if (mi.status === "ok") {
      var defaultTargets = [{ label: "Pass", value: m.passMark }, { label: "Distinction", value: m.distinctionMark }];
      var customTarget = (typeof m.targetMark === "number") ? m.targetMark : bundle.settings.targetMark;
      if (typeof customTarget === "number" && customTarget !== m.passMark && customTarget !== m.distinctionMark) {
        defaultTargets.push({ label: "Target", value: customTarget });
      }
      mi.a3Recovery = mi_computeA3Recovery_(m, marksRow, defaultTargets.filter(function (t) { return typeof t.value === "number"; }), bundle.settings);

      // v1.3.1 Marks Risk Analysis (spec §7): recovery difficulty using
      // remaining available marks, the required A3 for this module's own
      // effective target (the "Target" entry if one was configured and
      // distinct from Pass/Distinction, else "Pass"), recent practice
      // scores (last 7 days), and assessment proximity. No AI/probabilistic
      // model -- purely the deterministic rules in mi_computeRiskAnalysis_.
      var targetEntry = mi.a3Recovery.targets.filter(function (t) { return t.label === "Target"; })[0]
        || mi.a3Recovery.targets.filter(function (t) { return t.label === "Pass"; })[0];
      var bestRoute = targetEntry ? (targetEntry.routes.filter(function (r) { return r.isBest; })[0] || targetEntry.routes[0]) : null;
      var recentScores = (scoresByModule[m.name] || []).map(function (r) { return r.score; });
      mi.riskAnalysis = mi_computeRiskAnalysis_({
        requiredA3: bestRoute ? bestRoute.regulationValidRequiredA3 : null,
        requiredA3Status: bestRoute ? bestRoute.status : "insufficient_data",
        remainingAvailableMarks: (typeof mi.officialFinalMark === "number") ? 0 : null,
        recentPracticeScores: recentScores,
        daysToNextAssessment: nextAssessmentByModule[m.name] ? nextAssessmentByModule[m.name].days : null
      });
    } else {
      mi.riskAnalysis = { level: "Insufficient data", reasons: [(mi.reasons && mi.reasons[0]) || "Marks Intelligence is not available for this module yet."] };
    }
    return mi;
  });

  return {
    priorities: prioritized,
    weeklyAllocation: weeklyAllocation,
    nextAction: nextAction,
    resourceInsights: resourceInsights,
    examFocus: { triggers: examFocusTriggers, active: examFocusActiveRows },
    improvementTracker: improvementTracker,
    marksIntelligence: marksIntelligence
  };
}

// ============================================================================================
// v1.3.0 MARKS INTELLIGENCE EXPANSION
// (Deterministic Academic Intelligence -- additive on top of everything
// above.) Adds: a full per-module Marks Intelligence bundle (current mark,
// projected mark, Pass/Distinction status, a general required-future-marks
// figure, an assessment-contribution breakdown, remaining available marks,
// and a recovery-difficulty rating), a Hidden A2 inference engine, and an
// A3 Recovery Calculator that evaluates every faculty route the workbook has
// verified weightings for. Every function below is pure (no SpreadsheetApp/
// CalendarApp calls) and independently unit-tested -- see
// build_v3/test_harness.js and V1.3.0_STATIC_VALIDATION.md.
//
// DISCLOSED JUDGMENT CALL (same spirit as pctToNumber_ above): "04 Module
// Rules" only stores AF/A1/A2 weighting explicitly. This release treats the
// four components (AF, A1, A2, A3) as summing to 100% and derives the A3
// weighting as the remainder (100 - af - a1 - a2). See
// V1.3.0_SCORING_METHODOLOGY.md for the full disclosure.
// ============================================================================================

/** Published/sheet percentages are assumed rounded to within this many
 *  percentage points -- used only to decide whether a solved A2 that falls
 *  slightly outside 0-100% is "impossible" (a genuine data/weighting
 *  inconsistency) versus a harmless rounding overshoot that can be clamped.
 *  Deliberately a SEPARATE, wider tolerance from RECALC_MATCH_TOLERANCE_PCT
 *  below -- these check two different things (see mi_inferHiddenA2_): a
 *  narrow miss should still clamp to a usable estimate, but that clamped
 *  estimate must still round-trip closely under the tighter recalculation
 *  check to be called "verified". */
var ROUNDING_TOLERANCE_PCT = 0.5;

/** How closely a clamped A2 estimate's forward-recomputed provisional final
 *  must match the published figure to be called "verified". Intentionally
 *  tighter than ROUNDING_TOLERANCE_PCT above: a raw solved A2 that falls
 *  just outside 0-100% (and therefore gets clamped rather than rejected as
 *  impossible) can still shift the recomputed provisional final measurably
 *  away from what was actually published -- this is exactly the case
 *  "verified: false" exists to catch. */
var RECALC_MATCH_TOLERANCE_PCT = 0.15;

/**
 * Pure. v1.3.1 2ND CORRECTION PASS -- replaces the retired "derive A3 as the
 * remainder to 100%" model. The Faculty's real model is ROUTE-BASED, not
 * additive: FM1 = AF+A1+A2 (which is why every confirmed module's AF+A1+A2
 * already sums to 100 -- FM1 genuinely has no A3 component), FM2 = AF+A1+A3,
 * FM3 = AF+A2+A3, each combination renormalized to its own included weights.
 * A3's weight is therefore a GENUINELY SEPARATE, independently-verified
 * number (04 Module Rules: "A3 weight (%) (verified)") -- NEVER derived as a
 * remainder, NEVER assumed equal to A2's weight, NEVER split evenly across
 * whichever components are in play. Returns
 * { complete, af, a1, a2, a3, a3Verified, reason }, where `complete` covers
 * only the AF/A1/A2 baseline (sufficient for FM1); `a3Verified` is checked
 * separately since a module can have a complete FM1 baseline while its A3
 * weight remains unconfirmed (FM2/FM3 simply aren't computed until then).
 */
function mi_getModuleWeights_(m) {
  if (!m || m.ruleStatus !== "Verified") {
    return { complete: false, a3Verified: false, reason: "Module Rules for " + (m ? m.name : "this module") + " are not marked \"Verified\" in 04 Module Rules -- no weighting is available." };
  }
  if (typeof m.afWeight !== "number" || typeof m.a1Weight !== "number" || typeof m.a2Weight !== "number") {
    return { complete: false, a3Verified: false, reason: "AF/A1/A2 weighting is not fully configured in 04 Module Rules for " + m.name + "." };
  }
  var a3Verified = typeof m.a3Weight === "number";
  return {
    complete: true, af: m.afWeight, a1: m.a1Weight, a2: m.a2Weight,
    a3: a3Verified ? m.a3Weight : null, a3Verified: a3Verified,
    reason: a3Verified ? null : ("A3's own weight is not yet verified for " + m.name + " in 04 Module Rules (\"A3 weight (%) (verified)\") -- FM2/FM3 and any A3-based calculation are unavailable until it is entered; never assumed equal to A2's weight or split evenly.")
  };
}

/**
 * Pure. Resolves one assessment's written/not-written/excused/deferred
 * status. A BLANK MARK MUST NOT BE ASSUMED "not written" (2nd correction
 * pass) -- so status is read from its own explicit column first. Only when
 * that column is itself blank do we fall back to inferring "written" from a
 * real recorded mark (a percentage cannot exist without having been
 * written -- a safe, disclosed inference, not a guess); a module with
 * neither an explicit status nor a mark is "unknown", not silently treated
 * as "not written".
 */
function mi_statusOf_(explicitStatus, mark) {
  var s = (explicitStatus || "").toString().trim().toLowerCase();
  if (s === "written") return "written";
  if (s === "not written" || s === "not_written") return "not_written";
  if (s === "excused") return "excused";
  if (s === "deferred") return "deferred";
  if (typeof mark === "number") return "written"; // a real mark implies it was written
  return "unknown";
}

/** True only for a status that yields a genuinely usable mark for arithmetic. */
function mi_isWritten_(status) { return status === "written"; }

/**
 * Pure. MTD (Marks To Date) = (W_AF*AF + W_A1*A1) / (W_AF+W_A1), computed
 * ONLY when A1 has actually been written (per mi_statusOf_ -- never inferred
 * from a blank cell alone). Returns { value, status, reasons }, status one
 * of "ok" | "a1_not_written" | "insufficient_data".
 */
function mi_computeMtd_(weights, af, a1, a1Status) {
  if (!weights || !weights.complete) {
    return { value: null, status: "insufficient_data", reasons: ["AF/A1 weighting is not fully verified."] };
  }
  if (a1Status !== "written") {
    return { value: null, status: "a1_not_written", reasons: ["No MTD -- A1 not written" + (a1Status === "unknown" ? " (status not confirmed)." : (a1Status === "excused" ? " (excused)." : (a1Status === "deferred" ? " (deferred)." : ".")))] };
  }
  if (typeof af !== "number" || typeof a1 !== "number") {
    return { value: null, status: "insufficient_data", reasons: ["AF and/or A1 mark is missing despite A1 being marked written -- cannot compute MTD."] };
  }
  var denom = weights.af + weights.a1;
  if (denom <= 0) return { value: null, status: "insufficient_data", reasons: ["AF+A1 weighting sums to 0% -- cannot compute MTD."] };
  var value = round1_((weights.af * af + weights.a1 * a1) / denom);
  return { value: value, status: "ok", reasons: ["MTD = (" + weights.af + "%×AF + " + weights.a1 + "%×A1) / " + denom + "% = " + value + "%."] };
}

// ------------------------------------------------------------------
// HIDDEN A2 INFERENCE (spec §2)
// ------------------------------------------------------------------
/** How much a published, whole-number-rounded percentage (AF, A1, or the
 *  published pre-A3 figure) could differ from its true underlying value --
 *  used only to build the "possible range" interval on an Estimated A2, per
 *  spec §2's "support intervals when rounding creates multiple solutions".
 *  Not itself configurable (it describes ordinary rounding, not a policy
 *  choice) -- what IS configurable is RECALC_MATCH_TOLERANCE_PCT's
 *  workbook/module default, via `tolerancePct` below. */
var A2_INPUT_ROUNDING_HALFWIDTH_PCT = 0.5;

/** Pure. Resolves the effective "verify by recalculation" tolerance for one
 *  module: its own "A2 rounding tolerance override (%)" (04 Module Rules)
 *  if set, else the workbook-wide "Default A2 recalculation match tolerance
 *  (%)" (02 Settings) if THAT is set, else `undefined` -- letting
 *  mi_inferHiddenA2_ fall back to its own fixed constant. Never guesses a
 *  number that wasn't actually configured somewhere. */
function mi_effectiveA2Tolerance_(module, settings) {
  if (module && typeof module.a2ToleranceOverride === "number") return module.a2ToleranceOverride;
  if (settings && typeof settings.defaultA2Tolerance === "number") return settings.defaultA2Tolerance;
  return undefined;
}

/**
 * Pure. Infers a hidden/unpublished A2 mark from AF, A1, and the module's
 * own "Published Final (Before A3)" figure (11 Marks Tracker -- see
 * Marks_Intelligence_v1.3.1_Migration.gs) using the module's verified
 * AF/A1/A2 weighting. `tolerancePct`, if provided, is the configurable
 * "verify by recalculation" match tolerance for THIS call (module override,
 * or the workbook default, or the fixed fallback constant if neither is
 * configured yet) -- see RECALC_MATCH_TOLERANCE_PCT's comment above.
 *
 * v1.3.1 CORRECTION: the earlier v1.3.0 pass read this published figure from
 * "Provisional final (after A2)" -- that column is actually this sheet's own
 * FORMULA output, computed only once a real A2 is on record, and is
 * therefore the wrong field for a module that does NOT publish A2. This
 * function now reads ONLY marksRow.publishedFinalBeforeA3, a dedicated,
 * separately-stored manual-entry field that never overwrites A2, A3,
 * "Provisional final (after A2)", or "Final / provisional (after A3)". See
 * V1.3.1_CHANGE_SUMMARY.md.
 *
 * Returns null when A2 is not actually hidden (already published) -- there
 * is nothing to infer. Otherwise returns one of:
 *   { status: "insufficient_data", reasons }
 *   { status: "impossible", estimate, reasons }              -- out of 0-100%
 *   { status: "estimated", estimate, interval:{low,high}, verified, confidence, reasons }
 *
 * The result is NEVER labelled official anywhere in this codebase -- every
 * caller (the frontend included) must render it as "Estimated A2".
 */
function mi_inferHiddenA2_(marksRow, weights, tolerancePct) {
  if (!marksRow) return null;
  var a2Published = pctToNumber_(marksRow.a2);
  if (a2Published !== null) return null; // A2 already published -- nothing to infer.

  var af = pctToNumber_(marksRow.af), a1 = pctToNumber_(marksRow.a1);
  var published = pctToNumber_(marksRow.publishedFinalBeforeA3);
  if (af === null || a1 === null || published === null) {
    return { status: "insufficient_data", reasons: [
      "AF, A1, and a \"Published Final (Before A3)\" figure are all required to infer a hidden A2; at least one is missing for " + (marksRow.module || "this module") + "."
    ] };
  }
  if (!weights || !weights.complete) {
    return { status: "insufficient_data", reasons: ["Module weighting is not fully verified -- cannot infer a hidden A2 without verified AF/A1/A2 weighting from 04 Module Rules."] };
  }
  if (weights.a2 <= 0) {
    return { status: "insufficient_data", reasons: ["This module's A2 weighting is configured as 0% -- there is nothing to infer."] };
  }

  var tolerance = (typeof tolerancePct === "number" && tolerancePct >= 0) ? tolerancePct : RECALC_MATCH_TOLERANCE_PCT;

  var denom = round1_(weights.af + weights.a1 + weights.a2);
  var solveA2 = function (pub, afv, a1v) { return (pub * denom - weights.af * afv - weights.a1 * a1v) / weights.a2; };
  var raw = solveA2(published, af, a1);
  var reasons = [
    "Published: AF " + af + "%, A1 " + a1 + "%, Published Final (Before A3) " + published + "%.",
    "Verified weighting: AF " + weights.af + "%, A1 " + weights.a1 + "%, A2 " + weights.a2 + "% (of the " + denom + "% covered before A3).",
    "Recalculation match tolerance in use: " + tolerance + " point(s) (" + ((typeof tolerancePct === "number") ? "module/workbook-configured" : "fixed fallback -- run Marks_Intelligence_v1.3.1_Migration.gs to configure one") + ")."
  ];

  if (raw < -ROUNDING_TOLERANCE_PCT || raw > 100 + ROUNDING_TOLERANCE_PCT) {
    return { status: "impossible", estimate: round1_(raw), reasons: reasons.concat([
      "Solving for A2 with these inputs produces " + round1_(raw) + "%, which is outside the possible 0-100% range -- the published figures or weighting are likely inconsistent. Treat this module's Marks Intelligence as unavailable until checked."
    ]) };
  }

  var wasClamped = (raw < 0 || raw > 100);
  var estimate = round1_(Math.max(0, Math.min(100, raw)));
  // Verify by recalculation: forward-compute the provisional final from the
  // (possibly clamped) estimate and compare it against the published figure,
  // against the (possibly configured) tighter `tolerance` -- this is what
  // actually catches a clamped estimate that no longer honestly reproduces
  // what was published (see ROUNDING_TOLERANCE_PCT/RECALC_MATCH_TOLERANCE_PCT's
  // comments above for why this must be a separate, tighter check from the
  // impossible-value band).
  var recomputed = round1_((weights.af * af + weights.a1 * a1 + weights.a2 * estimate) / denom);
  var diff = round1_(Math.abs(recomputed - published));
  var verified = diff <= tolerance;
  reasons.push("Recalculating the provisional final from this estimate gives " + recomputed + "%, versus the published " + published + "% (difference " + diff + " point(s)).");
  if (wasClamped) reasons.push("The solved value fell slightly outside 0-100% and was clamped to " + estimate + "% before this recalculation check.");

  var confidence;
  if (!verified) confidence = "Low";
  else if (weights.a2 >= 15) confidence = "High";
  else if (weights.a2 >= 5) confidence = "Medium";
  else confidence = "Low";

  if (!verified) reasons.push("Recalculation difference exceeds the " + tolerance + "-point match tolerance -- treat this estimate with caution.");
  if (weights.a2 < 15 && verified) reasons.push("A2's weighting (" + weights.a2 + "%) is relatively small, so this estimate is more sensitive to rounding in the published figures.");

  // "Support intervals when rounding creates multiple solutions" (spec §2):
  // AF, A1, and the published figure are all themselves rounded, so any A2
  // within this range is EQUALLY consistent with the rounded inputs -- not
  // just the single point estimate above. Widened toward the direction that
  // most favours the student (higher published, lower AF/A1) for the high
  // end, and the reverse for the low end, using the standard rounding
  // half-width, then clamped to 0-100%.
  var high = round1_(Math.max(0, Math.min(100,
    solveA2(published + A2_INPUT_ROUNDING_HALFWIDTH_PCT, af - A2_INPUT_ROUNDING_HALFWIDTH_PCT, a1 - A2_INPUT_ROUNDING_HALFWIDTH_PCT))));
  var low = round1_(Math.max(0, Math.min(100,
    solveA2(published - A2_INPUT_ROUNDING_HALFWIDTH_PCT, af + A2_INPUT_ROUNDING_HALFWIDTH_PCT, a1 + A2_INPUT_ROUNDING_HALFWIDTH_PCT))));
  if (low > high) { var tmp = low; low = high; high = tmp; } // defensive -- keeps the interval well-formed regardless of sign of weights
  var interval = { low: low, high: high };
  if (high - low > 0.05) {
    reasons.push("Because AF, A1, and the published figure are themselves rounded, any A2 from " + low + "% to " + high + "% is equally consistent with those rounded inputs -- " + estimate + "% is the single best-guess midpoint, not a certainty.");
  }

  return { status: "estimated", estimate: estimate, interval: interval, verified: verified, confidence: confidence, reasons: reasons };
}

// ------------------------------------------------------------------
// FACULTY RULE ENGINE: FM1 / FM2 / FM3 / FMp (spec §3), DCA (spec §4 Rule 3)
// ------------------------------------------------------------------
/**
 * Pure. v1.3.1 2ND CORRECTION PASS -- retires the "DCA replaces the lower of
 * A2/A3 with the HIGHER of the two" guess (that invented a value that was
 * never actually entered anywhere). DCA (Deemed Continuous Assessment) is a
 * genuine, separately-awarded substitute mark: `dcaMark` is this module's
 * own entered figure (11 Marks Tracker: "DCA mark (%)"). Only applies when
 * DCA is enabled for the module, a real dcaMark has been entered, AND both
 * A2 and A3 are actually known (DCA replaces whichever of the two is lower
 * -- with only one known there is nothing to compare, so DCA does nothing).
 * Applied literally and unconditionally per the release spec ("identify the
 * lower of A2 and A3; replace that mark with DCA") -- no additional "only if
 * better" check is added. Preserves the original mark for audit. Returns
 * { values, applied, replacedComponent, originalValue, dcaMark }.
 */
function mi_applyDca_(values, dcaEnabled, dcaMark) {
  if (!dcaEnabled || typeof values.a2 !== "number" || typeof values.a3 !== "number" || typeof dcaMark !== "number") {
    return { values: values, applied: false };
  }
  var lowerKey = values.a2 <= values.a3 ? "a2" : "a3";
  var newValues = { af: values.af, a1: values.a1, a2: values.a2, a3: values.a3 };
  var originalValue = newValues[lowerKey];
  newValues[lowerKey] = dcaMark;
  return { values: newValues, applied: true, replacedComponent: lowerKey, originalValue: originalValue, dcaMark: dcaMark };
}

/** Plain-English rendering of an assessment status for use in reasons. */
function mi_describeStatus_(status) {
  return status === "not_written" ? "not written" : status === "excused" ? "excused"
    : status === "deferred" ? "deferred" : status === "unknown" ? "status not confirmed" : (status || "unknown");
}

/**
 * Pure. FM1 = (W_AF·AF + W_A1·A1 + W_A2·A2) / (W_AF+W_A1+W_A2). For every
 * confirmed module, W_AF+W_A1+W_A2 already sums to 100% by definition, so
 * this denominator is a no-op in practice -- included anyway for robustness
 * and symmetry with FM2/FM3's genuine renormalization. Requires A1 AND A2 to
 * be WRITTEN (mi_statusOf_), not merely present -- never fabricated from a
 * blank-but-assumed value.
 */
function mi_computeFm1_(weights, af, a1, a2, a1Status, a2Status) {
  if (!weights || !weights.complete) return { value: null, status: "weight_unverified", reasons: ["FM1: AF/A1/A2 weighting is not fully verified for this module."] };
  var problems = [];
  if (a1Status !== "written") problems.push("A1 is " + mi_describeStatus_(a1Status));
  if (a2Status !== "written") problems.push("A2 is " + mi_describeStatus_(a2Status));
  if (problems.length) return { value: null, status: "not_written", reasons: ["FM1 (AF+A1+A2) is not computable: " + problems.join("; ") + "."] };
  if (typeof af !== "number" || typeof a1 !== "number" || typeof a2 !== "number") {
    return { value: null, status: "insufficient_data", reasons: ["FM1: AF, A1, or A2 mark is missing despite being marked written."] };
  }
  var denom = weights.af + weights.a1 + weights.a2;
  if (denom <= 0) return { value: null, status: "insufficient_data", reasons: ["FM1: combined AF+A1+A2 weighting is 0%."] };
  var value = round1_((weights.af * af + weights.a1 * a1 + weights.a2 * a2) / denom);
  return { value: value, status: "ok", reasons: ["FM1 = (" + weights.af + "%×AF + " + weights.a1 + "%×A1 + " + weights.a2 + "%×A2) / " + round1_(denom) + "% = " + value + "%."] };
}

/**
 * Pure. FM2 = (W_AF·AF + W_A1·A1 + W_A3·A3) / (W_AF+W_A1+W_A3) -- a genuine
 * renormalization, since W_A3 is independently verified and these three
 * weights are not guaranteed to sum to 100%. Requires A3's weight to be
 * VERIFIED (never assumed equal to A2's weight, never split evenly) and A1 +
 * A3 to both be WRITTEN.
 */
function mi_computeFm2_(weights, af, a1, a3, a1Status, a3Status) {
  if (!weights || !weights.complete) return { value: null, status: "weight_unverified", reasons: ["FM2: AF/A1 weighting is not fully verified for this module."] };
  if (!weights.a3Verified) return { value: null, status: "a3_unverified", reasons: ["FM2 (AF+A1+A3) is unavailable: A3's own weight has not been verified for this module -- see \"A3 weight (%) (verified)\" in 04 Module Rules."] };
  var problems = [];
  if (a1Status !== "written") problems.push("A1 is " + mi_describeStatus_(a1Status));
  if (a3Status !== "written") problems.push("A3 is " + mi_describeStatus_(a3Status));
  if (problems.length) return { value: null, status: "not_written", reasons: ["FM2 (AF+A1+A3) is not computable: " + problems.join("; ") + "."] };
  if (typeof af !== "number" || typeof a1 !== "number" || typeof a3 !== "number") {
    return { value: null, status: "insufficient_data", reasons: ["FM2: AF, A1, or A3 mark is missing despite being marked written."] };
  }
  var denom = weights.af + weights.a1 + weights.a3;
  if (denom <= 0) return { value: null, status: "insufficient_data", reasons: ["FM2: combined AF+A1+A3 weighting is 0%."] };
  var value = round1_((weights.af * af + weights.a1 * a1 + weights.a3 * a3) / denom);
  return { value: value, status: "ok", reasons: ["FM2 = (" + weights.af + "%×AF + " + weights.a1 + "%×A1 + " + weights.a3 + "%×A3) / " + round1_(denom) + "% (renormalized) = " + value + "%."] };
}

/**
 * Pure. FM3 = (W_AF·AF + W_A2·A2 + W_A3·A3) / (W_AF+W_A2+W_A3). Mirrors
 * mi_computeFm2_ exactly, over A2+A3 instead of A1+A3.
 */
function mi_computeFm3_(weights, af, a2, a3, a2Status, a3Status) {
  if (!weights || !weights.complete) return { value: null, status: "weight_unverified", reasons: ["FM3: AF/A2 weighting is not fully verified for this module."] };
  if (!weights.a3Verified) return { value: null, status: "a3_unverified", reasons: ["FM3 (AF+A2+A3) is unavailable: A3's own weight has not been verified for this module -- see \"A3 weight (%) (verified)\" in 04 Module Rules."] };
  var problems = [];
  if (a2Status !== "written") problems.push("A2 is " + mi_describeStatus_(a2Status));
  if (a3Status !== "written") problems.push("A3 is " + mi_describeStatus_(a3Status));
  if (problems.length) return { value: null, status: "not_written", reasons: ["FM3 (AF+A2+A3) is not computable: " + problems.join("; ") + "."] };
  if (typeof af !== "number" || typeof a2 !== "number" || typeof a3 !== "number") {
    return { value: null, status: "insufficient_data", reasons: ["FM3: AF, A2, or A3 mark is missing despite being marked written."] };
  }
  var denom = weights.af + weights.a2 + weights.a3;
  if (denom <= 0) return { value: null, status: "insufficient_data", reasons: ["FM3: combined AF+A2+A3 weighting is 0%."] };
  var value = round1_((weights.af * af + weights.a2 * a2 + weights.a3 * a3) / denom);
  return { value: value, status: "ok", reasons: ["FM3 = (" + weights.af + "%×AF + " + weights.a2 + "%×A2 + " + weights.a3 + "%×A3) / " + round1_(denom) + "% (renormalized) = " + value + "%."] };
}

/**
 * Pure. Evaluates Faculty Override Rules 1 and 2 against one {a1,a2,a3}
 * state (real, post-DCA, or a hypothetical post-solve state -- the caller
 * decides). Rule 1's exact trigger condition is READ FROM THE MODULE'S OWN
 * CONFIGURED "A2/A3 subminimum mode" -- resolves the contradiction between
 * "A2<40 OR A3<40" and "max(A2,A3)<40" by never guessing between them: a
 * module with the mode left unconfigured simply has Rule 1 skipped entirely
 * (disclosed), never defaulted to either reading. Rule 2 (A1,A2,A3 all
 * WRITTEN caps FMp at 50%, unless the module is an official exception) is
 * evaluated against actual write status, not a hypothetical "about to be
 * written" inference. Returns { caps:[{rule,cap,reason}], effectiveCap,
 * reasons }, where effectiveCap is the lowest of any triggered caps (or null
 * if none triggered).
 */
function mi_evaluateOverrideCaps_(module, values, statuses) {
  var reasons = [], caps = [];
  var mode = module && module.subminimumMode;
  if (mode === "EACH_WRITTEN_ASSESSMENT_MUST_REACH_40" || mode === "AT_LEAST_ONE_OF_A2_OR_A3_MUST_REACH_40") {
    var a2w = statuses.a2Status === "written" && typeof values.a2 === "number";
    var a3w = statuses.a3Status === "written" && typeof values.a3 === "number";
    var trigger = false, why = [];
    if (mode === "EACH_WRITTEN_ASSESSMENT_MUST_REACH_40") {
      if (a2w && values.a2 < 40) { trigger = true; why.push("A2 (" + values.a2 + "%) is below 40%"); }
      if (a3w && values.a3 < 40) { trigger = true; why.push("A3 (" + values.a3 + "%) is below 40%"); }
    } else {
      if (a2w || a3w) {
        var maxWritten = Math.max(a2w ? values.a2 : -Infinity, a3w ? values.a3 : -Infinity);
        if (maxWritten < 40) { trigger = true; why.push("neither of the written A2/A3 marks reaches 40% (mode: at least one of A2/A3 must reach 40%)"); }
      }
    }
    if (trigger) {
      caps.push({ rule: "Rule 1", cap: 45, reason: "Faculty Rule 1 (\"" + mode + "\"): " + why.join(" and ") + " -- FMp capped at 45%." });
    } else if (a2w || a3w) {
      reasons.push("Faculty Rule 1 (\"" + mode + "\") evaluated: no cap triggered.");
    }
  } else {
    reasons.push("Faculty Rule 1: \"A2/A3 subminimum mode\" is not configured for this module in 04 Module Rules -- Rule 1 is not evaluated (the release spec's two possible readings, \"A2<40 OR A3<40\" versus \"max(A2,A3)<40\", produce different outcomes, so neither is assumed without an explicit, module-level configuration choice).");
  }

  var rule2Active = !(module && module.rule2Exception);
  if (rule2Active) {
    var allThreeWritten = statuses.a1Status === "written" && statuses.a2Status === "written" && statuses.a3Status === "written";
    if (allThreeWritten) {
      caps.push({ rule: "Rule 2", cap: 50, reason: "Faculty Rule 2: A1, A2 and A3 are all written -- FMp capped at 50% (no exception configured)." });
    }
  } else {
    reasons.push("Faculty Rule 2: this module is configured as an official exception in 04 Module Rules -- not evaluated.");
  }

  var effectiveCap = caps.length ? Math.min.apply(null, caps.map(function (c) { return c.cap; })) : null;
  return { caps: caps, effectiveCap: effectiveCap, reasons: reasons };
}

/**
 * Pure. The complete v1.3.1 (2nd correction pass) Marks calculation pipeline
 * for one module, in the exact order specified: (1) resolve written status
 * per assessment, (2) apply DCA substitution before anything else, (3)
 * compute FM1/FM2/FM3 from the (post-DCA) values, (4) FMp = max of whichever
 * routes are actually "ok", gated on at least two of A1/A2/A3 being WRITTEN
 * (original, pre-DCA status -- DCA substitutes a VALUE, it doesn't change
 * whether something was written), (5) apply Faculty Override Rules 1/2 to
 * get the Official Final Mark, distinct from the uncapped Raw FMp. `values`
 * = {af,a1,a2,a3} (real, or a verified Estimated A2 in place of a2 -- the
 * caller resolves that upstream exactly as before); `statuses` =
 * {a1Status,a2Status,a3Status}; `dcaMark` = this module's entered DCA mark,
 * or null. Nothing here is hidden -- see reasons/warnings for every
 * intermediate decision (spec §9: Transparency).
 */
function mi_computeFmpAndOfficial_(module, values, statuses, dcaMark) {
  var reasons = [], warnings = [];
  var weights = mi_getModuleWeights_(module);

  var dca = mi_applyDca_(values, !!(module && module.dcaEnabled), dcaMark);
  if (dca.applied) {
    reasons.push("Faculty Rule 3 (DCA) applied: " + dca.replacedComponent.toUpperCase() + " (" + dca.originalValue + "%, the lower of A2/A3) was replaced with the entered DCA mark (" + dca.dcaMark + "%) before FM1/FM2/FM3 were evaluated. Original mark preserved for audit.");
  } else if (module && module.dcaEnabled) {
    reasons.push("DCA is enabled for this module, but both A2 and A3 must be known and a DCA mark entered before it has anything to substitute -- no replacement made yet.");
  }
  var v = dca.values;
  var vA2Status = statuses.a2Status, vA3Status = statuses.a3Status; // DCA substitutes a VALUE, not written-ness

  var fm1 = mi_computeFm1_(weights, v.af, v.a1, v.a2, statuses.a1Status, vA2Status);
  var fm2 = mi_computeFm2_(weights, v.af, v.a1, v.a3, statuses.a1Status, vA3Status);
  var fm3 = mi_computeFm3_(weights, v.af, v.a2, v.a3, vA2Status, vA3Status);
  [["FM1", fm1], ["FM2", fm2], ["FM3", fm3]].forEach(function (pair) {
    reasons.push(pair[1].status === "ok" ? pair[1].reasons[0] : pair[0] + ": " + pair[1].reasons[0]);
  });

  var writtenCount = ["a1Status", "a2Status", "a3Status"].filter(function (k) { return statuses[k] === "written"; }).length;
  var fmpEligible = writtenCount >= 2;
  var candidates = [];
  if (fm1.status === "ok") candidates.push({ route: "FM1", value: fm1.value });
  if (fm2.status === "ok") candidates.push({ route: "FM2", value: fm2.value });
  if (fm3.status === "ok") candidates.push({ route: "FM3", value: fm3.value });

  var rawFmp = null, winningRoute = null;
  if (!fmpEligible) {
    reasons.push("Raw FMp not computed: fewer than two of A1/A2/A3 have been written for this module (" + writtenCount + " of 3).");
  } else if (!candidates.length) {
    reasons.push("Raw FMp not computed: no route (FM1/FM2/FM3) could be fully evaluated yet.");
  } else {
    var best = candidates.reduce(function (a, b) { return b.value > a.value ? b : a; });
    rawFmp = best.value;
    winningRoute = candidates.filter(function (c) { return c.value === rawFmp; }).map(function (c) { return c.route; }).join(" / ");
    reasons.push("Raw FMp = max(" + candidates.map(function (c) { return c.route + " " + c.value + "%"; }).join(", ") + ") = " + rawFmp + "%, via " + winningRoute + ".");
  }

  var officialFinalMark = rawFmp, capsApplied = [];
  if (rawFmp !== null) {
    var override = mi_evaluateOverrideCaps_(module, { a1: v.a1, a2: v.a2, a3: v.a3 }, statuses);
    reasons.push.apply(reasons, override.reasons);
    override.caps.forEach(function (c) {
      if (officialFinalMark > c.cap) {
        warnings.push(c.reason + " (down from the uncapped " + rawFmp + "%.)");
        capsApplied.push({ rule: c.rule, cap: c.cap, rawFmp: rawFmp });
        officialFinalMark = Math.min(officialFinalMark, c.cap);
      } else {
        reasons.push(c.reason + " (already at or below this cap -- no change.)");
      }
    });
  }

  return {
    fm1: fm1, fm2: fm2, fm3: fm3, rawFmp: rawFmp, officialFinalMark: officialFinalMark,
    winningRoute: winningRoute, dcaApplied: dca.applied, dcaInfo: dca.applied ? { replacedComponent: dca.replacedComponent, originalValue: dca.originalValue, dcaMark: dca.dcaMark } : null,
    capsApplied: capsApplied, values: v, weights: weights, reasons: reasons, warnings: warnings
  };
}

// ------------------------------------------------------------------
// REQUIRED A3 CALCULATOR (spec §3, §5) -- v1.3.1 2nd correction pass:
// mathematically-required vs regulation-valid vs impossible-due-to-cap,
// kept DELIBERATELY SEPARATE (never silently merged into one "corrected"
// number). Only ever applies to FM2/FM3 -- FM1 (AF+A1+A2) structurally has
// no A3 component under the corrected route model.
// ------------------------------------------------------------------
/**
 * Pure. Shared solve for FM2 (`otherKey` = "a1") or FM3 (`otherKey` = "a2").
 * `otherValue`/`otherStatus` is the route's other non-AF, non-A3 component
 * (A1 for FM2, A2 for FM3); `spectatorValue`/`spectatorStatus` is the THIRD
 * component not in this route's formula at all (A2 for FM2, A1 for FM3) --
 * carried along only so Faculty Override Rule 2 (which cares about all
 * three) can be evaluated correctly even though this route's formula
 * doesn't use it.
 */
function mi_computeRequiredA3ForRoute_(routeName, otherLabel, weights, af, otherValue, otherStatus, spectatorKey, spectatorValue, spectatorStatus, targetMark, module) {
  if (!weights || !weights.complete) {
    return { route: routeName, targetMark: targetMark, status: "weight_unverified",
      mathematicallyRequiredA3: null, regulationValidRequiredA3: null,
      reasons: [routeName + ": AF/" + otherLabel + " weighting is not fully verified."], warnings: [] };
  }
  if (!weights.a3Verified) {
    return { route: routeName, targetMark: targetMark, status: "a3_unverified",
      mathematicallyRequiredA3: null, regulationValidRequiredA3: null,
      reasons: [routeName + ": A3's own weight has not been verified for this module -- see \"A3 weight (%) (verified)\" in 04 Module Rules."], warnings: [] };
  }
  if (otherStatus !== "written") {
    return { route: routeName, targetMark: targetMark, status: "not_written",
      mathematicallyRequiredA3: null, regulationValidRequiredA3: null,
      reasons: [routeName + " requires " + otherLabel + " to be written -- " + otherLabel + " is " + mi_describeStatus_(otherStatus) + "."], warnings: [] };
  }
  if (typeof af !== "number" || typeof otherValue !== "number") {
    return { route: routeName, targetMark: targetMark, status: "insufficient_data",
      mathematicallyRequiredA3: null, regulationValidRequiredA3: null,
      reasons: [routeName + ": AF or " + otherLabel + " mark is missing."], warnings: [] };
  }

  var otherWeight = otherLabel === "A1" ? weights.a1 : weights.a2;
  var denom = weights.af + otherWeight + weights.a3;
  if (denom <= 0) {
    return { route: routeName, targetMark: targetMark, status: "insufficient_data",
      mathematicallyRequiredA3: null, regulationValidRequiredA3: null,
      reasons: [routeName + ": combined weighting is 0%."], warnings: [] };
  }
  var used = weights.af * af + otherWeight * otherValue;
  var raw = (targetMark * denom - used) / weights.a3;
  var reasons = [routeName + " (AF " + weights.af + "%, " + otherLabel + " " + otherWeight + "%, A3 " + weights.a3 + "%, renormalized over " + round1_(denom) + "%): AF/" + otherLabel + " currently contribute toward the " + targetMark + "% target."];

  var mathStatus, mathRequired;
  if (raw <= 0) { mathStatus = "already_secured"; mathRequired = 0; }
  else if (raw > 100) { mathStatus = "impossible"; mathRequired = round1_(raw); }
  else { mathStatus = "ok"; mathRequired = round1_(raw); }
  reasons.push(mathStatus === "already_secured" ? "Mathematically required A3: already secured (0% needed)."
    : mathStatus === "impossible" ? "Mathematically required A3: " + mathRequired + "% -- exceeds 100%, not mathematically possible."
    : "Mathematically required A3: " + mathRequired + "%.");

  if (mathStatus === "impossible") {
    return { route: routeName, targetMark: targetMark, status: "impossible",
      mathematicallyRequiredA3: mathRequired, regulationValidRequiredA3: null, reasons: reasons, warnings: [] };
  }

  // Regulation check: would achieving this A3 (with the other component now
  // written, A3 about to be written, and the third/spectator component
  // exactly as it currently stands) survive Faculty Override Rules 1/2?
  var hypotheticalValues = {}; hypotheticalValues[spectatorKey] = spectatorValue; hypotheticalValues.a3 = mathRequired;
  hypotheticalValues[otherLabel === "A1" ? "a1" : "a2"] = otherValue;
  var hypotheticalStatuses = {
    a1Status: (otherLabel === "A1") ? "written" : spectatorStatus,
    a2Status: (otherLabel === "A2") ? "written" : spectatorStatus,
    a3Status: "written"
  };
  var override = mi_evaluateOverrideCaps_(module, hypotheticalValues, hypotheticalStatuses);
  var warnings = [];
  if (override.effectiveCap !== null && override.effectiveCap < targetMark) {
    override.caps.forEach(function (c) { warnings.push(c.reason); });
    return { route: routeName, targetMark: targetMark, status: "impossible_due_to_cap",
      mathematicallyRequiredA3: mathRequired, regulationValidRequiredA3: null,
      reasons: reasons.concat(["Regulation-valid required A3: NOT achievable -- even though the raw arithmetic reaches " + targetMark + "%, a Faculty Override Rule would cap the Official Final Mark at " + override.effectiveCap + "%, below the target."]),
      warnings: warnings };
  }
  reasons.push("Regulation-valid required A3: " + mathRequired + "% (no Faculty Override Rule prevents reaching " + targetMark + "% via this route with this A3).");
  return { route: routeName, targetMark: targetMark, status: mathStatus,
    mathematicallyRequiredA3: mathRequired, regulationValidRequiredA3: mathRequired,
    reasons: reasons, warnings: warnings };
}

/**
 * Pure orchestrator for one module: evaluates the required-A3 figure on
 * EVERY route that structurally has an A3 component (FM2, FM3 -- FM1 never
 * does, under the corrected route model) for each requested target
 * `{label, value}`, classifying each as "Best route" (lowest regulation-
 * valid required A3 among reachable/comparable routes), "Alternative
 * route", or "Invalid route" (impossible / impossible_due_to_cap /
 * weight_unverified / a3_unverified / not_written / insufficient_data).
 * Also reports FM1's own current status (already fully determined once
 * written, or not yet) for context, since a target already secured via FM1
 * makes any A3 discussion moot. `settings` supplies the workbook-wide
 * default A2 recalculation tolerance for Hidden A2 inference.
 */
function mi_computeA3Recovery_(module, marksRow, targets, settings) {
  var weights = mi_getModuleWeights_(module);
  var tolerance = mi_effectiveA2Tolerance_(module, settings);
  var a2Published = marksRow ? pctToNumber_(marksRow.a2) : null;
  var inferredA2 = weights.complete ? mi_inferHiddenA2_(marksRow, weights, tolerance) : null;
  var a2Value = a2Published, a2Source = "published";
  if (a2Value === null && inferredA2 && inferredA2.status === "estimated") { a2Value = inferredA2.estimate; a2Source = "estimated"; }

  var af = marksRow ? pctToNumber_(marksRow.af) : null;
  var a1 = marksRow ? pctToNumber_(marksRow.a1) : null;
  var a1Status = mi_statusOf_(marksRow ? marksRow.a1Status : null, a1);
  var a2Status = mi_statusOf_(marksRow ? marksRow.a2Status : null, a2Value);

  var known = { af: af, a1: a1, a2: a2Value };

  var results = (targets || []).map(function (t) {
    var fm2Result = mi_computeRequiredA3ForRoute_("FM2", "A1", weights, af, a1, a1Status, "a2", a2Value, a2Status, t.value, module);
    var fm3Result = mi_computeRequiredA3ForRoute_("FM3", "A2", weights, af, a2Value, a2Status, "a1", a1, a1Status, t.value, module);
    var routeResults = [fm2Result, fm3Result];
    var comparable = routeResults.filter(function (r) { return typeof r.regulationValidRequiredA3 === "number"; });
    if (comparable.length > 1) {
      var best = comparable.reduce(function (a, b) { return b.regulationValidRequiredA3 < a.regulationValidRequiredA3 ? b : a; });
      routeResults.forEach(function (r) { r.isBest = (r === best); r.isAlternative = (comparable.indexOf(r) !== -1 && r !== best); r.isInvalid = comparable.indexOf(r) === -1; });
    } else if (comparable.length === 1) {
      routeResults.forEach(function (r) { r.isBest = (r === comparable[0]); r.isAlternative = false; r.isInvalid = !r.isBest; });
    } else {
      routeResults.forEach(function (r) { r.isBest = false; r.isAlternative = false; r.isInvalid = true; });
    }
    return { label: t.label, targetMark: t.value, routes: routeResults };
  });

  return {
    module: module.name, a2Source: a2Source, inferredA2: inferredA2,
    weights: weights, known: known, targets: results
  };
}

/**
 * Read-only endpoint: A3 Recovery Calculator for one module against a
 * user-entered custom target, computed on demand (the three default targets
 * -- Pass / Distinction / module or workbook Target -- are already included
 * in every planner load via computeAcademicIntelligence_ -> marksIntelligence,
 * so this endpoint exists only for the "custom target" case, which cannot be
 * precomputed since it's whatever number the user just typed in).
 */
function api_computeCustomA3Recovery(moduleName, customTargetMark) {
  try {
    var mod = requireModule_(moduleName);
    var target = requireFiniteNumber_(customTargetMark, "Custom target mark", 0, 100);
    var planner = api_getPlannerData();
    if (!planner.ok) return planner;
    var moduleRow = planner.data.modules.filter(function (m) { return m.name === mod; })[0];
    var marksRow = planner.data.marksTracker.filter(function (r) { return r.module === mod; })[0];
    if (!moduleRow) throw new Error("Module not found: " + mod);
    var result = mi_computeA3Recovery_(moduleRow, marksRow, [{ label: "Custom", value: target }], planner.data.settings);
    return ok_(result);
  } catch (e) { return fail_(e); }
}

// ------------------------------------------------------------------
// MARKS RISK ANALYSIS (spec §7)
// ------------------------------------------------------------------
/**
 * Pure. Deterministic (no AI/probabilistic model) recovery-difficulty
 * rating: Low / Medium / High / Impossible, using the module's required A3
 * (for its own best route/effective target), remaining available marks,
 * recent (last 7 days) practice/past-paper scores, and assessment
 * proximity. `input` = { requiredA3, requiredA3Status, remainingAvailableMarks,
 * recentPracticeScores: number[], daysToNextAssessment }.
 */
function mi_computeRiskAnalysis_(input) {
  var reasons = [];
  var status = input.requiredA3Status;

  if (!status || status === "insufficient_data") {
    return { level: "Insufficient data", reasons: ["Not enough verified data (module weighting, AF/A1/A2, or a target) to assess recovery difficulty yet."] };
  }
  if (status === "impossible") {
    return { level: "Impossible", reasons: ["The required A3 to reach this target exceeds 100% on the best available route -- not mathematically achievable."] };
  }
  if (status === "impossible_due_to_cap") {
    return { level: "Impossible", reasons: ["A Faculty Override Rule (Rule 1 or Rule 2) caps the Official Final Mark below this target on this route, regardless of A3."] };
  }
  if (status === "already_secured" || (typeof input.remainingAvailableMarks === "number" && input.remainingAvailableMarks <= 0)) {
    return { level: "Low", reasons: ["The target is already secured (or nothing remains to be assessed) -- no recovery risk."] };
  }
  if (typeof input.requiredA3 !== "number") {
    return { level: "Insufficient data", reasons: ["No usable required-A3 figure is available for this module's effective target yet."] };
  }

  var rankTier = ["Low", "Medium", "High"];
  var tierRank = { Low: 0, Medium: 1, High: 2 };
  var tier = input.requiredA3 <= 50 ? "Low" : input.requiredA3 <= 70 ? "Medium" : "High";
  reasons.push("Base difficulty from required A3 (" + input.requiredA3 + "% on the best available route): " + tier + ".");

  var scores = input.recentPracticeScores || [];
  if (scores.length) {
    var avg = round1_(scores.reduce(function (a, b) { return a + b; }, 0) / scores.length);
    reasons.push("Recent (last 7 days) practice/past-paper average: " + avg + "% across " + scores.length + " scored item(s), versus the " + input.requiredA3 + "% required.");
    if (avg >= input.requiredA3 + 10 && tierRank[tier] > 0) {
      tier = rankTier[tierRank[tier] - 1];
      reasons.push("Recent practice comfortably exceeds the requirement -- difficulty eased to " + tier + ".");
    } else if (avg < input.requiredA3 - 10 && tierRank[tier] < 2) {
      tier = rankTier[tierRank[tier] + 1];
      reasons.push("Recent practice is well below the requirement -- difficulty raised to " + tier + ".");
    }
  } else {
    reasons.push("No recent (last 7 days) practice/past-paper scores recorded for this module -- difficulty not adjusted for recent practice.");
  }

  if (typeof input.daysToNextAssessment === "number") {
    reasons.push("Next assessment in " + input.daysToNextAssessment + " day(s).");
    if (input.daysToNextAssessment <= 3 && input.requiredA3 > 60 && tierRank[tier] < 2) {
      tier = rankTier[tierRank[tier] + 1];
      reasons.push("Assessment is within 3 days and the requirement is demanding (" + input.requiredA3 + "%) -- difficulty raised to " + tier + ".");
    }
  } else {
    reasons.push("No upcoming assessment date recorded for this module.");
  }

  return { level: tier, reasons: reasons };
}

// ------------------------------------------------------------------
// ASSESSMENT SIMULATOR / "WHAT IF?" (spec §6)
// ------------------------------------------------------------------
/**
 * Pure. `hypothetical` = { af?, a1?, a2?, a3? } -- ENTIRELY hypothetical,
 * never real recorded marks; a key OMITTED entirely means that component is
 * excluded from this scenario (e.g. simulating a before-A3 question), while
 * a key present with value null means "unknown -- solve for this one".
 * With exactly one included component unknown, solves for the value that
 * component would need to reach each target (spec example: "AF 75%, A1 58%,
 * A2 Unknown -> Need 62% for distinction"). With zero unknown, reports the
 * resulting contribution. With more than one unknown, explains that a
 * single figure isn't possible until only one is left blank.
 */
function mi_simulateWhatIf_(module, hypothetical, targets) {
  var weights = mi_getModuleWeights_(module);
  if (!weights.complete) return { status: "insufficient_data", reasons: [weights.reason] };

  var allKeys = ["af", "a1", "a2", "a3"];
  var included = allKeys.filter(function (k) { return Object.prototype.hasOwnProperty.call(hypothetical, k); });
  var reasons = ["Hypothetical / What-If scenario -- none of these values are real recorded marks."];

  // v1.3.1 2nd correction pass: the Faculty's real model is route-based
  // (leave-one-out), not one blended AF+A1+A2+A3 weighting -- so a scenario
  // must include AF plus exactly two of {A1,A2,A3}, mapping onto exactly one
  // of FM1 (AF+A1+A2), FM2 (AF+A1+A3), or FM3 (AF+A2+A3). Any other
  // combination is not a route this Faculty model defines and is explicitly
  // rejected rather than guessed at.
  var routeKeySets = { fm1: ["af", "a1", "a2"], fm2: ["af", "a1", "a3"], fm3: ["af", "a2", "a3"] };
  var matchedRoute = Object.keys(routeKeySets).filter(function (r) {
    var keys = routeKeySets[r];
    return keys.length === included.length && keys.every(function (k) { return included.indexOf(k) !== -1; });
  })[0];

  if (!matchedRoute) {
    return { status: "insufficient_data", reasons: reasons.concat([
      "A What-If scenario must include AF plus exactly two of A1/A2/A3, matching one of this Faculty's three routes -- FM1 (AF+A1+A2), FM2 (AF+A1+A3), or FM3 (AF+A2+A3). Components supplied: " + (included.length ? included.map(function (k) { return k.toUpperCase(); }).join(", ") : "none") + "."
    ]) };
  }
  if (matchedRoute !== "fm1" && !weights.a3Verified) {
    return { status: "a3_unverified", reasons: reasons.concat(["This scenario needs A3's own verified weight (" + matchedRoute.toUpperCase() + " includes A3), which has not been verified for this module -- see \"A3 weight (%) (verified)\" in 04 Module Rules."]) };
  }

  var routeLabel = matchedRoute.toUpperCase();
  var wMap = matchedRoute === "fm1" ? { af: weights.af, a1: weights.a1, a2: weights.a2 }
    : matchedRoute === "fm2" ? { af: weights.af, a1: weights.a1, a3: weights.a3 }
    : { af: weights.af, a2: weights.a2, a3: weights.a3 };
  var denom = Object.keys(wMap).reduce(function (s, k) { return s + wMap[k]; }, 0);

  var known = included.filter(function (k) { return typeof hypothetical[k] === "number"; });
  var unknown = included.filter(function (k) { return typeof hypothetical[k] !== "number"; });
  var usedPoints = known.reduce(function (sum, k) { return sum + wMap[k] * hypothetical[k]; }, 0);

  reasons.push(routeLabel + " scenario (" + Object.keys(wMap).map(function (k) { return k.toUpperCase(); }).join("+") + "): " + included.map(function (k) {
    return k.toUpperCase() + (typeof hypothetical[k] === "number" ? " " + hypothetical[k] + "%" : " (unknown -- solving for this)");
  }).join(", ") + ".");

  var results = (targets || []).map(function (t) {
    if (unknown.length === 0) {
      var fm = round1_(usedPoints / denom);
      return { label: t.label, targetMark: t.value, status: "computed", contribution: fm,
        reasons: [routeLabel + " = " + fm + "% with all three components hypothetically known."] };
    }
    if (unknown.length === 1) {
      var uKey = unknown[0], uWeight = wMap[uKey];
      if (uWeight <= 0) {
        return { label: t.label, targetMark: t.value, status: "no_component", requiredValue: null, component: uKey,
          reasons: ["\"" + uKey.toUpperCase() + "\" has a 0% weighting on " + routeLabel + " -- it cannot affect the outcome."] };
      }
      var raw = (t.value * denom - usedPoints) / uWeight;
      if (raw <= 0) {
        return { label: t.label, targetMark: t.value, status: "already_secured", requiredValue: 0, component: uKey,
          reasons: ["The " + t.value + "% target is already secured from the other hypothetical inputs alone on " + routeLabel + "."] };
      }
      if (raw > 100) {
        return { label: t.label, targetMark: t.value, status: "impossible", requiredValue: round1_(raw), component: uKey,
          reasons: ["Reaching " + t.value + "% would require " + round1_(raw) + "% on " + uKey.toUpperCase() + " via " + routeLabel + " -- not mathematically possible."] };
      }
      return { label: t.label, targetMark: t.value, status: "ok", requiredValue: round1_(raw), component: uKey,
        reasons: ["Need " + round1_(raw) + "% on " + uKey.toUpperCase() + " for " + t.label.toLowerCase() + " (" + t.value + "%) via " + routeLabel + "."] };
    }
    return { label: t.label, targetMark: t.value, status: "multiple_unknown", requiredValue: null,
      reasons: ["More than one component is still unknown -- leave only one blank to get a single required figure."] };
  });

  return { status: "ok", module: module.name, route: routeLabel, weights: weights, hypothetical: hypothetical, targets: results, reasons: reasons };
}

/**
 * Read-only endpoint: Assessment Simulator ("What If?"). `hypothetical` is a
 * plain object from the client, e.g. { af: 75, a1: 58, a2: null } (a2 sent
 * as null/""/"Unknown" all mean "solve for this"; a key omitted entirely
 * means "exclude this component from the scenario"). Also returns the full
 * FM1/FM2/FM3/Raw-FMp/Official-Final-Mark pipeline evaluated against the
 * SAME hypothetical values, for full transparency (spec §9). A hypothetical
 * key with a real number is treated as "written" for this simulation ONLY --
 * never affects any real recorded status. DCA is not applied to a
 * hypothetical scenario (no real dcaMark exists for a what-if).
 */
function api_simulateAssessment(moduleName, hypothetical) {
  try {
    var mod = requireModule_(moduleName);
    var planner = api_getPlannerData();
    if (!planner.ok) return planner;
    var moduleRow = planner.data.modules.filter(function (m) { return m.name === mod; })[0];
    if (!moduleRow) throw new Error("Module not found: " + mod);

    var h = {};
    ["af", "a1", "a2", "a3"].forEach(function (k) {
      if (hypothetical && Object.prototype.hasOwnProperty.call(hypothetical, k)) {
        var v = hypothetical[k];
        if (v === "" || v === null || typeof v === "undefined" || v === "Unknown") { h[k] = null; }
        else { h[k] = requireFiniteNumber_(v, k.toUpperCase() + " (What-If)", 0, 100); }
      }
    });

    var targets = [
      { label: "Pass", value: moduleRow.passMark },
      { label: "Distinction", value: moduleRow.distinctionMark }
    ];
    if (typeof moduleRow.targetMark === "number" && moduleRow.targetMark !== moduleRow.passMark && moduleRow.targetMark !== moduleRow.distinctionMark) {
      targets.push({ label: "Target", value: moduleRow.targetMark });
    }
    targets = targets.filter(function (t) { return typeof t.value === "number"; });

    var result = mi_simulateWhatIf_(moduleRow, h, targets);
    var statuses = {
      a1Status: typeof h.a1 === "number" ? "written" : "unknown",
      a2Status: typeof h.a2 === "number" ? "written" : "unknown",
      a3Status: typeof h.a3 === "number" ? "written" : "unknown"
    };
    result.fmpAndOfficial = mi_computeFmpAndOfficial_(moduleRow, { af: h.af, a1: h.a1, a2: h.a2, a3: h.a3 }, statuses, null);
    return ok_(result);
  } catch (e) { return fail_(e); }
}

// ------------------------------------------------------------------
// MARKS INTELLIGENCE ENGINE (spec §1) -- the full per-module bundle
// ------------------------------------------------------------------
/**
 * Pure. Produces the complete per-module Marks Intelligence bundle: current
 * mark, projected mark, Pass/Distinction status, a general required-future-
 * marks figure (average % needed across whatever weighted percentage points
 * are still undetermined), an assessment-contribution breakdown, remaining
 * available marks, and a recovery-difficulty rating. Every field is
 * null/"insufficient_data" rather than fabricated when the underlying inputs
 * are missing. `passMark`/`distinctionMark` are this module's own configured
 * thresholds (04 Module Rules, defaulting 50/75 -- see api_getPlannerData).
 * `settings` (optional -- v1.3.1) supplies the workbook-wide default A2
 * recalculation tolerance; a per-module override in `module` always takes
 * priority over it (see mi_effectiveA2Tolerance_).
 */
function mi_computeMarksIntelligence_(module, marksRow, passMark, distinctionMark, settings) {
  var weights = mi_getModuleWeights_(module);
  if (!weights.complete) {
    return { code: module.code, name: module.name, status: "insufficient_data", reasons: [weights.reason] };
  }

  var af = marksRow ? pctToNumber_(marksRow.af) : null;
  var a1 = marksRow ? pctToNumber_(marksRow.a1) : null;
  var a3 = marksRow ? pctToNumber_(marksRow.a3) : null;
  var a2Published = marksRow ? pctToNumber_(marksRow.a2) : null;
  var inferredA2 = mi_inferHiddenA2_(marksRow, weights, mi_effectiveA2Tolerance_(module, settings));
  var a2Value = a2Published, a2Source = "published";
  if (a2Value === null && inferredA2 && inferredA2.status === "estimated") { a2Value = inferredA2.estimate; a2Source = "estimated"; }

  // v1.3.1 2nd correction pass: status (written/not-written/excused/deferred)
  // is tracked independently of the mark itself -- a blank mark never
  // silently means "not written". Estimated A2 counts as "written" for
  // pipeline purposes since it stands in for a real value, disclosed via
  // a2Source below.
  var a1Status = mi_statusOf_(marksRow ? marksRow.a1Status : null, a1);
  var a2Status = (a2Source === "estimated") ? "written" : mi_statusOf_(marksRow ? marksRow.a2Status : null, a2Value);
  var a3Status = mi_statusOf_(marksRow ? marksRow.a3Status : null, a3);

  var reasons = [];
  if (a2Source === "estimated") reasons.push("A2 is not yet published for this module -- Estimated A2 (" + a2Value + "%, confidence " + inferredA2.confidence + ") was used. See the A2 inference panel for the full calculation.");
  if (!weights.a3Verified) reasons.push("A3 rules incomplete -- required-A3 and FM2/FM3 calculations unavailable.");

  var mtd = mi_computeMtd_(weights, af, a1, a1Status);
  reasons.push(mtd.reasons[0]);

  var dcaMark = (marksRow && typeof marksRow.dcaMark === "number") ? marksRow.dcaMark : null;
  var pipeline = mi_computeFmpAndOfficial_(module, { af: af, a1: a1, a2: a2Value, a3: a3 }, { a1Status: a1Status, a2Status: a2Status, a3Status: a3Status }, dcaMark);
  reasons = reasons.concat(pipeline.reasons);

  // v1.4.0 -- "Published Final (After A3)": a module that publishes a new
  // combined final percentage after A3 instead of a raw A3 mark. Folded in
  // as an alternative candidate for the Official Final Mark, never a
  // replacement of the computed pipeline outright -- the higher of the two
  // wins, so a real A3 result (raw or published) can only ever raise the
  // mark, exactly like every other FMp route (FM1/FM2/FM3 already share
  // this same "best route wins" property).
  var publishedAfterA3 = pctToNumber_(marksRow ? marksRow.publishedFinalAfterA3 : null);
  if (typeof publishedAfterA3 === "number") {
    if (typeof pipeline.officialFinalMark === "number") {
      if (publishedAfterA3 > pipeline.officialFinalMark) {
        reasons.push("Published Final (After A3) of " + publishedAfterA3 + "% is higher than the computed Official Final Mark of " + pipeline.officialFinalMark + "% -- the published figure is used (a real A3 result can only raise your mark, never lower it).");
        pipeline.officialFinalMark = publishedAfterA3;
      } else {
        reasons.push("Published Final (After A3) of " + publishedAfterA3 + "% entered, but the computed Official Final Mark of " + pipeline.officialFinalMark + "% is already at least as high -- the computed value is kept.");
      }
    } else {
      reasons.push("Published Final (After A3) of " + publishedAfterA3 + "% used directly as the Official Final Mark (not enough of A1/A2/A3 recorded individually to compute FM1/FM2/FM3).");
      pipeline.officialFinalMark = publishedAfterA3;
    }
  }

  function statusFor(target, label) {
    if (typeof target !== "number") return { label: label, status: "Insufficient data", reasons: ["No " + label.toLowerCase() + " mark is configured for this module."] };
    if (typeof pipeline.officialFinalMark === "number") {
      return pipeline.officialFinalMark >= target
        ? { label: label, status: "Secured", reasons: [label + " is secured: the Official Final Mark of " + pipeline.officialFinalMark + "% already reaches " + target + "%."] }
        : { label: label, status: "At risk", reasons: [label + " target of " + target + "% is not met by the Official Final Mark of " + pipeline.officialFinalMark + "%. See the Required A3 Calculator for what's still needed."] };
    }
    if (typeof pipeline.rawFmp === "number") {
      return pipeline.rawFmp >= target
        ? { label: label, status: "On track", reasons: [label + " target of " + target + "% is on track at the current raw FMp of " + pipeline.rawFmp + "% (" + pipeline.winningRoute + ")."] }
        : { label: label, status: "At risk", reasons: [label + " target of " + target + "% is not yet met at the current raw FMp of " + pipeline.rawFmp + "% (" + pipeline.winningRoute + "). See the Required A3 Calculator."] };
    }
    if (typeof mtd.value === "number") {
      return { label: label, status: "Insufficient data", reasons: [label + ": no Faculty route (FM1/FM2/FM3) can be calculated yet -- only MTD (" + mtd.value + "%) is available so far."] };
    }
    return { label: label, status: "Insufficient data", reasons: ["Not enough written assessments yet to evaluate " + label.toLowerCase() + "."] };
  }
  var passStatus = statusFor(passMark, "Pass");
  var distinctionStatus = statusFor(distinctionMark, "Distinction");

  var recoveryDifficulty;
  if (typeof pipeline.officialFinalMark === "number") recoveryDifficulty = "Not applicable";
  else if (typeof pipeline.rawFmp !== "number") recoveryDifficulty = "Insufficient data";
  else recoveryDifficulty = "See Required A3 Calculator";

  var currentMark = (typeof pipeline.officialFinalMark === "number") ? pipeline.officialFinalMark
    : (typeof pipeline.rawFmp === "number") ? pipeline.rawFmp
    : (typeof mtd.value === "number") ? mtd.value
    : null;

  return {
    code: module.code, name: module.name, status: "ok",
    currentMark: currentMark,
    values: pipeline.values, weights: pipeline.weights,
    mtd: mtd, fm1: pipeline.fm1, fm2: pipeline.fm2, fm3: pipeline.fm3,
    rawFmp: pipeline.rawFmp, officialFinalMark: pipeline.officialFinalMark, winningRoute: pipeline.winningRoute,
    dcaApplied: pipeline.dcaApplied, dcaInfo: pipeline.dcaInfo, capsApplied: pipeline.capsApplied,
    passStatus: passStatus, distinctionStatus: distinctionStatus, recoveryDifficulty: recoveryDifficulty,
    a1Status: a1Status, a2Status: a2Status, a3Status: a3Status,
    a2Source: a2Source, inferredA2: (a2Source === "estimated") ? inferredA2 : null,
    publishedFinalBeforeA3: pctToNumber_(marksRow ? marksRow.publishedFinalBeforeA3 : null),
    publishedFinalAfterA3: publishedAfterA3,
    dcaMarkOnFile: dcaMark, subminimumMode: module.subminimumMode || "",
    ruleStatus: module.ruleStatus || "", a3Verified: weights.a3Verified,
    reasons: reasons, warnings: pipeline.warnings || []
  };
}
