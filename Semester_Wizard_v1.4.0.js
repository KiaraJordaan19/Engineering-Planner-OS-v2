/**
 * Semester_Wizard_v1.4.0.gs — Engineering Planner OS v1.4.0, Semester
 * Lifecycle & Analytics.
 *
 * New Semester Wizard (spec §8). Every step (1-6) is collected CLIENT-SIDE
 * only — nothing is written to the workbook until the Review step (step 7)
 * calls semWizardApply_(config) with everything the user confirmed, applied
 * inside a single lock so the whole wizard either fully commits or fully
 * fails (no half-applied semester). semWizardGetOptions_ is the one
 * read-only lookup the wizard needs to populate its steps' choices.
 *
 * Reset (Semester_Delete_Reset_v1.4.0.gs) deliberately never clears "03
 * Modules", "04 Module Rules", or "05 Timetable Import" — so "reuse
 * previous semester" for Modules/Timetable is very often already a no-op
 * (the previous semester's rows are simply still there). "25 Module
 * Templates" exists for the genuinely different case: a module that isn't
 * currently in "03 Modules" at all (a new module selection, or one
 * returning after a gap) but whose colour/contacts/weightings/rules were
 * saved from an earlier semester and should be restored rather than
 * retyped from scratch.
 */

// ------------------------------------------------------------------
// READ-ONLY LOOKUPS (populate the wizard's steps)
// ------------------------------------------------------------------
function semWizardGetOptions_() {
  var pd = api_getPlannerData();
  if (!pd.ok) throw new Error('Could not read workbook data: ' + pd.error);
  return {
    currentModules: pd.data.modules.map(function (m) { return { code: m.code, name: m.name, color: m.color, active: m.active }; }),
    templates: semListModuleTemplates_(),
    currentCalendarId: pd.data.settings.calendarId || '',
    currentSemesterName: pd.data.settings.semesterName || '',
    currentAcademicYear: getSetting_('Academic year') || '',
    timetableRowCount: readSheetRows_(TIMETABLE_SHEET, HEADER_ROW, 'Module').length,
    semesterStatus: getSetting_('Semester status') || 'Active'
  };
}

// ------------------------------------------------------------------
// APPLY (spec §8 step 7 — everything below runs from ONE confirmed config)
// ------------------------------------------------------------------
/**
 * config = {
 *   semesterName, academicYear, startDate, endDate,               // step 1
 *   modulesMode: 'reuse' | 'new', reuseModuleCodes: [code, ...],   // step 2
 *   coloursMode: 'preserve' | 'edit', colourOverrides: {code: hex},// step 3
 *   timetableMode: 'import' | 'skip',                              // step 4
 *   rulesMode: 'reuse' | 'new',                                    // step 5
 *   calendarMode: 'reuse' | 'new' | 'skip', newCalendarName: ''    // step 6
 * }
 * Returns a snapshot summary for step 7's confirmation display.
 */
function semWizardApply_(config) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(LOCK_WAIT_MS)) throw new Error('Workbook is busy — try again in a moment.');
  try {
    var name = requireText_(config.semesterName, 'Semester name');
    var start = parseValidDate_(config.startDate, 'Start date');
    var end = parseValidDate_(config.endDate, 'End date');
    if (!start || !end) throw new Error('Both a start date and an end date are required.');

    // ---- Step 1: identity + a brand-new Semester ID ----
    var settingsSheet = SpreadsheetApp.getActive().getSheetByName(SETTINGS_SHEET);
    var existingIds = readSheetRows_(SEM_ARCHIVE_SHEET, HEADER_ROW, 'Semester ID').map(function (r) { return r['Semester ID']; })
      .concat(readSheetRows_(SEM_HISTORY_SHEET, HEADER_ROW, 'Semester ID').map(function (r) { return r['Semester ID']; }));
    var newSemesterId = nextId_('SEM', existingIds);
    semSetSetting_('Semester name', name);
    semSetSetting_('Academic year', config.academicYear || '');
    semSetSetting_('Semester start date', start);
    semSetSetting_('Semester end date', end);
    semSetSetting_('Semester ID', newSemesterId);
    semSetSetting_('Semester status', 'Active');
    semSetSetting_('Semester manually finished (TRUE/FALSE)', false);

    // ---- Step 2: modules ----
    var modulesApplied = [];
    if (config.modulesMode === 'reuse' && config.reuseModuleCodes && config.reuseModuleCodes.length) {
      var templatesByCode = {};
      semListModuleTemplates_().forEach(function (t) { templatesByCode[t.code] = t; });
      config.reuseModuleCodes.forEach(function (code) {
        var t = templatesByCode[code];
        if (!t) return;
        semApplyModuleIdentityFromTemplate_(t);
        modulesApplied.push(code);
      });
    }

    // ---- Step 3: colours ----
    if (config.coloursMode === 'edit' && config.colourOverrides) {
      Object.keys(config.colourOverrides).forEach(function (code) {
        semSetModuleColour_(code, config.colourOverrides[code]);
      });
    }

    // ---- Step 4: timetable ----
    if (config.timetableMode === 'skip') {
      var tt = SpreadsheetApp.getActive().getSheetByName(TIMETABLE_SHEET);
      if (tt) {
        var lr = tt.getLastRow(), lc = tt.getLastColumn();
        if (lr > HEADER_ROW) tt.getRange(HEADER_ROW + 1, 1, lr - HEADER_ROW, lc).clearContent();
      }
    }
    // 'import' mode is intentionally a no-op here: Reset never clears "05
    // Timetable Import", so a previous semester's rows (or rows the user
    // pasted in immediately before running the Wizard) are simply left as
    // they are -- see file header note for why v1.4.0 does not invent a new
    // timetable file-import parser.

    // ---- Step 5: assessment frameworks / module rules ----
    var rulesApplied = [];
    if (config.rulesMode === 'reuse') {
      var currentCodes = readSheetRows_(MODULES_SHEET, HEADER_ROW, 'Module name').map(function (m) { return m['Module code']; });
      var templatesByCode2 = {};
      semListModuleTemplates_().forEach(function (t) { templatesByCode2[t.code] = t; });
      currentCodes.forEach(function (code) {
        var t = templatesByCode2[code];
        if (!t) return;
        semApplyModuleRulesFromTemplate_(t);
        rulesApplied.push(code);
      });
    }

    // ---- Step 6: calendar ----
    var calendarResult = { mode: config.calendarMode || 'reuse' };
    if (config.calendarMode === 'new') {
      var cal = CalendarApp.createCalendar(config.newCalendarName || (name + ' Planner'));
      semSetSetting_('Planner Calendar ID (read/write)', cal.getId());
      semSetSetting_('Calendar sync enabled (TRUE/FALSE)', true);
      calendarResult.calendarId = cal.getId();
      calendarResult.calendarName = cal.getName();
    } else if (config.calendarMode === 'skip') {
      semSetSetting_('Planner Calendar ID (read/write)', '');
      semSetSetting_('Calendar sync enabled (TRUE/FALSE)', false);
    }
    // 'reuse' -- no-op, existing Settings calendar ID is left exactly as is.

    logAutomation_('New Semester Wizard completed', newSemesterId, 'Success',
      name + ' — modules: ' + modulesApplied.length + ', rules: ' + rulesApplied.length + ', calendar: ' + calendarResult.mode);

    return {
      semesterId: newSemesterId, semesterName: name, academicYear: config.academicYear || '',
      startDate: isoDate_(start), endDate: isoDate_(end),
      modulesApplied: modulesApplied, rulesApplied: rulesApplied, calendar: calendarResult
    };
  } finally {
    lock.releaseLock();
  }
}

/** Upserts one module's IDENTITY fields (name/code/colour/contacts/
 *  study-preference/attendance fields) into "03 Modules" from a saved
 *  template. Never touches "04 Module Rules" -- see semApplyModuleRulesFromTemplate_. */
function semApplyModuleIdentityFromTemplate_(t) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(MODULES_SHEET);
  var map = getColMap_(sheet);
  var rows = readSheetRows_(MODULES_SHEET, HEADER_ROW, 'Module name');
  var existing = rows.filter(function (r) { return r['Module code'] === t.code; })[0];
  var row = existing ? existing.__row : Math.max(sheet.getLastRow() + 1, HEADER_ROW + 1);
  function setIf(header, value) { if (map[header] && typeof value !== 'undefined') sheet.getRange(row, map[header]).setValue(value); }
  setIf('Module name', t.name);
  setIf('Module code', t.code);
  setIf('Colour (hex)', t.color);
  setIf('Active status', 'Active');
  var c = t.contacts || {}, sp = t.studyPreferences || {}, ar = t.attendanceRules || {};
  setIf('Coordinator name', c.coordinatorName); setIf('Coordinator email', c.coordinatorEmail);
  setIf('Lecturer name', c.lecturerName); setIf('Lecturer email', c.lecturerEmail); setIf('Contact notes', c.contactNotes);
  setIf('Conceptual difficulty (1-5)', sp.difficulty); setIf('Workload intensity (1-5)', sp.workload);
  setIf('Repeated module (TRUE/FALSE)', sp.repeated); setIf('Target mark (%)', sp.targetMark); setIf('Min weekly study hours', sp.minWeeklyHours);
  setIf('Attendance-sensitive lecture (TRUE/FALSE)', ar.attendanceSensitive); setIf('Compulsory practical/lab (TRUE/FALSE)', ar.compulsoryPractical);
}

/** Upserts one module's RULE/WEIGHTING fields into "04 Module Rules" from a
 *  saved template. Never touches "03 Modules". */
function semApplyModuleRulesFromTemplate_(t) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(MODULE_RULES_SHEET);
  var map = getColMap_(sheet);
  var rows = readSheetRows_(MODULE_RULES_SHEET, HEADER_ROW, 'Module code');
  var existing = rows.filter(function (r) { return r['Module code'] === t.code; })[0];
  var row = existing ? existing.__row : Math.max(sheet.getLastRow() + 1, HEADER_ROW + 1);
  function setIf(header, value) { if (map[header] && typeof value !== 'undefined' && value !== null) sheet.getRange(row, map[header]).setValue(value); }
  var w = t.weightings || {}, f = t.assessmentFramework || {};
  setIf('Module code', t.code);
  setIf('AF weighting', w.afWeight); setIf('A1 weighting', w.a1Weight); setIf('A2 weighting', w.a2Weight);
  setIf('A3 weight (%) (verified)', w.a3Weight); setIf('Pass mark', w.passMark); setIf('Distinction mark', w.distinctionMark);
  setIf('Rule status', f.ruleStatus); setIf('AF calculation method', f.afMethod); setIf('A2/A3 subminimum mode', f.subminimumMode);
  setIf('Rule 2 exception (TRUE/FALSE)', f.rule2Exception); setIf('DCA enabled (TRUE/FALSE)', f.dcaEnabled);
  setIf('A2 rounding tolerance override (%)', f.a2ToleranceOverride); setIf('General notes', f.generalNotes);
}

function semSetModuleColour_(code, hex) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(MODULES_SHEET);
  var map = getColMap_(sheet);
  var rows = readSheetRows_(MODULES_SHEET, HEADER_ROW, 'Module name');
  var existing = rows.filter(function (r) { return r['Module code'] === code; })[0];
  if (existing && map['Colour (hex)']) sheet.getRange(existing.__row, map['Colour (hex)']).setValue(hex);
}
