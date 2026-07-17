/**
 * Semester_API_v1.4.0.gs — Engineering Planner OS v1.4.0, Semester
 * Lifecycle & Analytics.
 *
 * The web-app-facing surface for everything in this release: every function
 * below is an api_sem* endpoint the Semester screen (Index_v1.4.0.html)
 * calls via google.script.run, following the exact same contract every
 * other api_* function in API_v1.3.1.gs already uses — { ok: true, data }
 * or { ok: false, error } — using the SAME ok_()/fail_() helpers (not
 * reimplemented here). Deliberately kept in its own file, separate from
 * API_v1.3.1.gs, so v1.4.0 ships as a pure addition: API_v1.3.1.gs is not
 * modified by this release at all.
 *
 * Deliberately NOT folded into api_getPlannerData()'s single batched read:
 * the Semester screen is opened far less often than every other screen, and
 * a full Semester Report is a genuinely heavier computation (it calls
 * computeAcademicIntelligence_ a second time) than anything else on that
 * payload — so it is fetched lazily, on demand, the same pattern this app
 * already uses for Study Suggestions and the Assessment Simulator.
 */

// ------------------------------------------------------------------
// DASHBOARD (spec §11)
// ------------------------------------------------------------------
function api_semDashboard() {
  try {
    var pd = api_getPlannerData();
    if (!pd.ok) throw new Error(pd.error);
    var completion = semDetectCompletion_(pd.data);
    var archives = semListArchives_();
    var backups = semListBackups_();
    return ok_({
      current: {
        semesterId: getSetting_('Semester ID') || '',
        semesterName: pd.data.settings.semesterName || '',
        academicYear: getSetting_('Academic year') || '',
        startDate: pd.data.settings.semesterStart || '',
        endDate: pd.data.settings.semesterEnd || '',
        status: getSetting_('Semester status') || 'Active'
      },
      completion: completion,
      archivedSemesters: archives,
      backups: backups.slice(0, 10),
      reportAvailable: true, // the current semester's report can always be generated on demand
      backupAvailable: backups.length > 0,
      archiveStatus: archives.length ? (archives.length + ' semester(s) archived') : 'No semesters archived yet',
      plannerVersion: getSetting_('Planner version') || '1.4.0'
    });
  } catch (e) { return fail_(e); }
}

// ------------------------------------------------------------------
// SEMESTER COMPLETION (spec §1)
// ------------------------------------------------------------------
function api_semDetectCompletion() {
  try {
    var pd = api_getPlannerData();
    if (!pd.ok) throw new Error(pd.error);
    return ok_(semDetectCompletion_(pd.data));
  } catch (e) { return fail_(e); }
}

function api_semMarkFinished() {
  try { return ok_(semMarkManuallyFinished_()); } catch (e) { return fail_(e); }
}

// ------------------------------------------------------------------
// REPORT (spec §2) + CHARTS (spec §3)
// ------------------------------------------------------------------
/** Generates a live report for the CURRENT (not-yet-archived) semester —
 *  a preview, never itself a write to any archive. Safe to call as often
 *  as the user likes; only semStartDeleteWorkflow_ / semWizardApply_ /
 *  semSaveModuleTemplates_ actually write anything durable. */
function api_semGenerateReport(reflection) {
  try {
    var report = semReportBuild_(reflection || {});
    var charts = semChartsBuild_(report);
    return ok_({ report: report, charts: charts });
  } catch (e) { return fail_(e); }
}

function api_semGetArchivedReport(semesterId) {
  try {
    var found = semGetArchivedReport_(requireText_(semesterId, 'Semester ID'));
    if (!found) return fail_(new Error('No archived semester found for ' + semesterId));
    return ok_(found);
  } catch (e) { return fail_(e); }
}

function api_semListArchives() {
  try { return ok_(semListArchives_()); } catch (e) { return fail_(e); }
}

// ------------------------------------------------------------------
// BACKUP (spec §5)
// ------------------------------------------------------------------
/** Standalone "Generate Backup" button, independent of the Delete Semester
 *  workflow (which also creates one automatically) -- so the user can take
 *  a safety snapshot at any point, not only when deleting. */
function api_semCreateBackup() {
  try { return ok_(semCreateBackup_(getSetting_('Semester ID') || '')); } catch (e) { return fail_(e); }
}
function api_semListBackups() {
  try { return ok_(semListBackups_()); } catch (e) { return fail_(e); }
}
function api_semVerifyBackup(backupId) {
  try { return ok_(semVerifyBackup_(requireText_(backupId, 'Backup ID'))); } catch (e) { return fail_(e); }
}

// ------------------------------------------------------------------
// ARCHIVE (spec §4) + MODULE TEMPLATES (spec §9)
// ------------------------------------------------------------------
/** Manual "refresh templates now" action, independent of the Delete
 *  Semester workflow (which also refreshes templates automatically). Lets
 *  the user save colour/contact/rule changes into Module Templates at any
 *  time without going through the full delete workflow. */
function api_semSaveModuleTemplates() {
  try { return ok_(semSaveModuleTemplates_(getSetting_('Semester ID') || '')); } catch (e) { return fail_(e); }
}
function api_semListModuleTemplates() {
  try { return ok_(semListModuleTemplates_()); } catch (e) { return fail_(e); }
}

// ------------------------------------------------------------------
// HISTORICAL ANALYTICS (spec §10)
// ------------------------------------------------------------------
function api_semHistoricalCompare() {
  try { return ok_(semHistoricalCompare_()); } catch (e) { return fail_(e); }
}

// ------------------------------------------------------------------
// DELETE SEMESTER WORKFLOW (spec §6)
// ------------------------------------------------------------------
function api_semStartDeleteWorkflow(reflection) {
  try { return ok_(semStartDeleteWorkflow_(reflection || {})); } catch (e) { return fail_(e); }
}

/** typedPhrase must be exactly "DELETE SEMESTER" -- enforced inside
 *  semConfirmDelete_, not re-checked here, so there is exactly one place in
 *  the whole codebase that gate lives. */
function api_semConfirmDelete(workflowToken, typedPhrase, options) {
  try { return ok_(semConfirmDelete_(workflowToken, typedPhrase, options || {})); } catch (e) { return fail_(e); }
}

// ------------------------------------------------------------------
// RESET ACTIVE SEMESTER (spec §7) — exposed directly ONLY for the case
// where a workbook is already mid-semester-transition outside the Delete
// Semester workflow (e.g. re-running the Wizard after a manual archive).
// Still requires the same typed phrase, via the same semConfirmDelete_
// gate — there is no api_ endpoint anywhere in this file that reaches
// semResetActiveSemester_ without it.
// ------------------------------------------------------------------
function api_semResetActiveSemester(typedPhrase, options) {
  try { return ok_(semEmergencyReset_(typedPhrase, options || {})); } catch (e) { return fail_(e); }
}

// ------------------------------------------------------------------
// NEW SEMESTER WIZARD (spec §8)
// ------------------------------------------------------------------
function api_semWizardGetOptions() {
  try { return ok_(semWizardGetOptions_()); } catch (e) { return fail_(e); }
}
function api_semWizardApply(config) {
  try { return ok_(semWizardApply_(config || {})); } catch (e) { return fail_(e); }
}
