/**
 * Semester_Delete_Reset_v1.4.0.gs — Engineering Planner OS v1.4.0, Semester
 * Lifecycle & Analytics.
 *
 * Reset Active Semester (spec §7) and the Delete Semester workflow (spec
 * §6): Generate Report -> Generate Backup -> Archive Semester -> Display
 * Summary -> Require typed confirmation ("DELETE SEMESTER") -> only THEN
 * clear active data. The two halves are deliberately split into two
 * functions (semStartDeleteWorkflow_ / semConfirmDelete_) so the typed
 * confirmation gate genuinely sits between "the permanent record now
 * exists" and "active data is cleared" — nothing before that gate can
 * delete anything.
 */

// ------------------------------------------------------------------
// RESET ACTIVE SEMESTER (spec §7)
// ------------------------------------------------------------------
/**
 * Clears ONLY: Assignments, Assessments, Study Sessions (+ Study Tasks),
 * Revision, Marks (+ AF Components), Attendance, Inbox, Automation history.
 * Never touches: Settings, Module Rules, Modules (incl. colours & contacts),
 * Module Templates, Guide, Automation configuration (triggers), or either
 * Archive sheet (17 Archive / 23 Semester Archive) — none of those sheets
 * appear in sheetsToClear below, so there is no code path here that can
 * reach them.
 *
 * options.deleteCalendarEvents (boolean, default false): if true, also
 * deletes every planner-created Calendar event referenced by a "Calendar
 * Event ID" cell on Assignments/Assessments/Study Planner BEFORE those
 * sheets are cleared (so the IDs are still readable). This is the ONE
 * genuinely destructive side effect outside the workbook itself, which is
 * why spec §7 calls it out as needing its own confirmation — the client is
 * responsible for getting that confirmation before setting this flag to
 * true; this function trusts the flag it's given.
 */
function semResetActiveSemester_(options) {
  options = options || {};
  var ss = SpreadsheetApp.getActive();
  var deletedEvents = 0, calendarErrors = [];

  if (options.deleteCalendarEvents) {
    var eventSources = [ASSIGNMENTS_SHEET, ASSESSMENTS_SHEET, STUDY_SHEET];
    var calendar = null;
    try { calendar = getTargetCalendar_(); } catch (e) { calendarErrors.push('Could not resolve the planner calendar: ' + e.message); }
    if (calendar) {
      eventSources.forEach(function (sheetName) {
        readSheetRows_(sheetName, HEADER_ROW, 'Calendar Event ID').forEach(function (r) {
          var eventId = r['Calendar Event ID'];
          if (!eventId) return;
          try {
            var ev = calendar.getEventById(eventId);
            if (ev) { ev.deleteEvent(); deletedEvents++; }
          } catch (e) { calendarErrors.push(sheetName + ': ' + (e && e.message)); }
        });
      });
    }
  }

  var sheetsToClear = [
    { name: ASSIGNMENTS_SHEET, header: HEADER_ROW },
    { name: ASSESSMENTS_SHEET, header: HEADER_ROW },
    { name: STUDY_SHEET, header: HEADER_ROW },
    { name: STUDY_TASKS_SHEET, header: HEADER_ROW },
    { name: REVISION_SHEET, header: HEADER_ROW },
    { name: MARKS_TRACKER_SHEET, header: MARKS_TRACKER_HEADER_ROW },
    { name: AF_COMPONENTS_SHEET, header: HEADER_ROW },
    { name: ATTENDANCE_SHEET, header: HEADER_ROW },
    { name: INBOX_SHEET, header: HEADER_ROW },
    { name: AUTOMATION_LOG_SHEET, header: HEADER_ROW }
  ];
  var cleared = [];
  sheetsToClear.forEach(function (s) {
    var sheet = ss.getSheetByName(s.name);
    if (!sheet) return;
    var lr = sheet.getLastRow(), lc = sheet.getLastColumn();
    if (lr > s.header) {
      sheet.getRange(s.header + 1, 1, lr - s.header, lc).clearContent();
      cleared.push(s.name);
    }
  });

  // Clear the "manually finished" flag now that the semester it referred to
  // has actually been reset -- the New Semester Wizard starts every new
  // semester with a clean completion-detection state.
  semSetSetting_('Semester manually finished (TRUE/FALSE)', false);

  // logAutomation_ writes AFTER the clear above, into the now-empty
  // Automation Log -- this is the first row of the new semester's history,
  // which is the correct, intentional behavior (see file header note).
  logAutomation_('Semester reset (active data cleared)', '-', 'Success',
    cleared.join(', ') + (options.deleteCalendarEvents ? (' + ' + deletedEvents + ' calendar event(s) deleted') : '') +
    (calendarErrors.length ? (' — calendar warnings: ' + calendarErrors.join('; ')) : ''));

  return { clearedSheets: cleared, deletedEvents: deletedEvents, calendarErrors: calendarErrors };
}

/** Small helper: write one "02 Settings" value by label, creating nothing
 *  (silently no-ops if the label doesn't exist — every label this file
 *  writes is guaranteed present by Semester_Lifecycle_v1.4.0_Migration.gs). */
function semSetSetting_(label, value) {
  var settings = SpreadsheetApp.getActive().getSheetByName(SETTINGS_SHEET);
  var cell = settings.createTextFinder(label).matchEntireCell(false).findNext();
  if (cell) settings.getRange(cell.getRow(), 3).setValue(value);
}

// ------------------------------------------------------------------
// MANUAL "FINISH SEMESTER" (spec §1 — always requires this explicit action
// or full auto-detection; never triggers on its own)
// ------------------------------------------------------------------
function semMarkManuallyFinished_() {
  semSetSetting_('Semester manually finished (TRUE/FALSE)', true);
  logAutomation_('Semester marked finished', getSetting_('Semester ID') || '-', 'Success', 'Manually finished via "Finish Semester".');
  return { ok: true };
}

// ------------------------------------------------------------------
// DELETE SEMESTER WORKFLOW (spec §6)
// ------------------------------------------------------------------
/**
 * Step 1 of 2. Generates the report, backup, and archive (in that order,
 * per spec §6's diagram), refreshes Module Templates, and marks the
 * semester "Complete" in Settings. Does NOT clear any active data — that
 * only happens in semConfirmDelete_ below, after the typed confirmation.
 * Safe to call more than once accidentally EXCEPT that semCreateArchive_
 * will throw on the second call (archives are never overwritten) — the
 * client should treat that specific error as "already prepared, go straight
 * to the confirmation step" rather than a genuine failure.
 */
var SEM_WORKFLOWS_SHEET = '27 Semester Workflows';
var SEM_WORKFLOW_TTL_MS = 24 * 60 * 60 * 1000;
var SEM_LIFECYCLE_LOCK_MS = 30000;

function semWithLifecycleLock_(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(SEM_LIFECYCLE_LOCK_MS)) throw new Error('Another semester lifecycle operation is already running. Try again shortly.');
  try { return fn(); } finally { lock.releaseLock(); }
}

function semWorkflowRows_() {
  return readSheetRows_(SEM_WORKFLOWS_SHEET, HEADER_ROW, 'Workflow Token');
}

function semFindWorkflow_(token) {
  return semWorkflowRows_().filter(function (r) { return r['Workflow Token'] === token; })[0] || null;
}

function semFindPreparedWorkflowForSemester_(semesterId) {
  var now = Date.now();
  return semWorkflowRows_().filter(function (r) {
    var exp = r['Expires At'] ? new Date(r['Expires At']).getTime() : 0;
    return r['Semester ID'] === semesterId && r['Status'] === 'PREPARED' && exp > now;
  }).sort(function (a,b) { return new Date(b['Prepared At']).getTime() - new Date(a['Prepared At']).getTime(); })[0] || null;
}

function semArchiveExists_(semesterId) {
  return readSheetRows_(SEM_ARCHIVE_SHEET, HEADER_ROW, 'Semester ID').some(function (r) { return r['Semester ID'] === semesterId; });
}

function semWorkflowIsValid_(row, semesterId) {
  if (!row || row['Status'] !== 'PREPARED') return { ok:false, reason:'Workflow is not prepared.' };
  if (row['Semester ID'] !== semesterId) return { ok:false, reason:'Workflow belongs to a different semester.' };
  if (!row['Expires At'] || new Date(row['Expires At']).getTime() <= Date.now()) return { ok:false, reason:'Workflow token has expired.' };
  var verification = semVerifyBackup_(row['Backup ID']);
  if (!verification.ok) return { ok:false, reason:'Associated backup failed verification: ' + (verification.reason || verification.missingSheets.join(', ')) };
  if (!semArchiveExists_(row['Archive Semester ID'])) return { ok:false, reason:'Associated archive is missing.' };
  return { ok:true };
}

function semWriteWorkflow_(record) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(SEM_WORKFLOWS_SHEET);
  if (!sheet) throw new Error('"27 Semester Workflows" not found — run the v1.4.0 migration again.');
  var row = Math.max(sheet.getLastRow()+1, HEADER_ROW+1);
  sheet.getRange(row,1,1,11).setValues([[
    record.token, record.semesterId, record.backupId, record.archiveSemesterId,
    record.preparedAt, record.expiresAt, record.status, '', record.plannerVersion,
    record.lastStage || 'PREPARED', record.lastError || ''
  ]]);
  return row;
}

function semUpdateWorkflow_(rowNumber, fields) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(SEM_WORKFLOWS_SHEET);
  var map = getColMap_(sheet);
  Object.keys(fields).forEach(function (header) {
    if (map[header]) sheet.getRange(rowNumber, col_(map, header)).setValue(fields[header]);
  });
}

function semStartDeleteWorkflow_(reflection) {
  return semWithLifecycleLock_(function () {
    var report = semReportBuild_(reflection);
    var semesterId = report.meta.semesterId;
    var existing = semFindPreparedWorkflowForSemester_(semesterId);
    if (existing) {
      var valid = semWorkflowIsValid_(existing, semesterId);
      if (valid.ok) {
        return {
          semesterId: semesterId, workflowToken: existing['Workflow Token'],
          summary: semSummariseReport_(report),
          backup: { backupId: existing['Backup ID'] },
          archive: { semesterId: existing['Archive Semester ID'] },
          templatesSaved: 0, reused: true
        };
      }
    }

    var charts = semChartsBuild_(report);
    var backup = semCreateBackup_(semesterId);
    var verified = semVerifyBackup_(backup.backupId);
    if (!verified.ok) throw new Error('Backup verification failed; active data was not changed.');

    var archive;
    if (semArchiveExists_(semesterId)) archive = { semesterId: semesterId, reused: true };
    else archive = semCreateArchive_(report, charts);
    if (!semArchiveExists_(semesterId)) throw new Error('Archive verification failed; active data was not changed.');

    var templates = semSaveModuleTemplates_(semesterId);
    var now = new Date();
    var token = Utilities.getUuid();
    semWriteWorkflow_({
      token: token, semesterId: semesterId, backupId: backup.backupId,
      archiveSemesterId: semesterId, preparedAt: now,
      expiresAt: new Date(now.getTime()+SEM_WORKFLOW_TTL_MS), status:'PREPARED',
      plannerVersion: getSetting_('Planner version') || '1.4.0', lastStage:'PREPARED'
    });
    semSetSetting_('Semester status', 'Complete');
    logAutomation_('Delete Semester workflow prepared', semesterId, 'Success', 'Backup verified, archive verified, single-use authorization created.');
    return {
      semesterId: semesterId, workflowToken: token,
      summary: semSummariseReport_(report),
      backup: { backupId: backup.backupId, fileUrl: backup.fileUrl },
      archive: { semesterId: archive.semesterId }, templatesSaved: templates.saved, reused:false
    };
  });
}

function semConfirmDelete_(workflowToken, typedPhrase, options) {
  return semWithLifecycleLock_(function () {
    if ((typedPhrase || '').trim() !== 'DELETE SEMESTER') throw new Error('Confirmation phrase did not match — type exactly: DELETE SEMESTER');
    if (!workflowToken) throw new Error('A prepared workflow token is required. Generate the report, backup, and archive first.');
    var currentSemesterId = getSetting_('Semester ID') || '';
    var row = semFindWorkflow_(workflowToken);
    if (!row) throw new Error('Workflow token was not found.');
    var valid = semWorkflowIsValid_(row, currentSemesterId);
    if (!valid.ok) throw new Error(valid.reason);
    var result = semResetActiveSemester_(options || {});
    semUpdateWorkflow_(row.__row, { 'Status':'CONSUMED', 'Consumed At':new Date(), 'Last Stage':'RESET_COMPLETE', 'Last Error':'' });
    logAutomation_('Semester deleted (confirmed)', currentSemesterId, 'Success', 'Prepared workflow consumed; active data reset.');
    result.workflowTokenConsumed = true;
    return result;
  });
}

function semEmergencyReset_(typedPhrase, options) {
  return semWithLifecycleLock_(function () {
    if ((typedPhrase || '').trim() !== 'EMERGENCY RESET') throw new Error('Confirmation phrase did not match — type exactly: EMERGENCY RESET');
    var semesterId = getSetting_('Semester ID') || '';
    var backup = semCreateBackup_(semesterId);
    var verified = semVerifyBackup_(backup.backupId);
    if (!verified.ok) throw new Error('Emergency backup verification failed; no active data was cleared.');
    var result = semResetActiveSemester_(options || {});
    result.emergencyBackupId = backup.backupId;
    logAutomation_('Emergency reset completed', semesterId, 'Success', 'Verified backup ' + backup.backupId + ' created before reset.');
    return result;
  });
}

/** A small, display-only subset of the full report -- what spec §6's
 *  "Display Summary" step shows before asking for the typed confirmation.
 *  The full report remains available afterward via the Archive. */
function semSummariseReport_(report) {
  return {
    semester: report.meta.semesterName, academicYear: report.meta.academicYear,
    moduleCount: report.overview.modules.length,
    studyHours: report.overview.studyHours, revisionHours: report.overview.revisionHours,
    overallAverage: report.marks.overallAverage, highestMark: report.marks.highestMark, lowestMark: report.marks.lowestMark,
    distinctionCount: report.marks.distinctionCount, passCount: report.marks.passCount,
    attendanceRate: report.attendance.attendanceRate,
    assignmentsCompleted: report.overview.assignmentsCompleted, assessmentsWritten: report.overview.assessmentsWritten
  };
}
