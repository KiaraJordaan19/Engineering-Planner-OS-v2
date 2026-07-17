/**
 * Semester_Archive_v1.4.0.gs — Engineering Planner OS v1.4.0, Semester
 * Lifecycle & Analytics.
 *
 * Semester Archive (spec §4), Module Templates (spec §9), and Historical
 * Analytics (spec §10). Every write in this file targets ONLY the three new
 * sheets this release adds ("23 Semester Archive", "25 Module Templates",
 * "26 Semester History") plus reading (never writing) "03 Modules" / "04
 * Module Rules" for templates. Nothing here touches "17 Archive" or any
 * other pre-v1.4.0 sheet.
 *
 * Depends on Semester_Lifecycle_v1.4.0_Migration.gs having been run once
 * (the three sheets above must exist), and on readSheetRows_ / getColMap_ /
 * col_ / num_ / round1_ / HEADER_ROW / nextId_ already being present.
 */

var SEM_ARCHIVE_SHEET = '23 Semester Archive';
var SEM_BACKUPS_SHEET = '24 Semester Backups';
var SEM_TEMPLATES_SHEET = '25 Module Templates';
var SEM_HISTORY_SHEET = '26 Semester History';

// ------------------------------------------------------------------
// ARCHIVE (spec §4, §12: unique ID + timestamp + planner version on every
// archive; never overwritten; read-only once written)
// ------------------------------------------------------------------
/**
 * Writes ONE new row to "23 Semester Archive" for report.meta.semesterId.
 * Refuses outright (throws) if a row for that exact Semester ID already
 * exists — archives are never overwritten, per spec §4/§12. Also appends the
 * matching rollup row to "26 Semester History" (semUpdateHistory_) in the
 * same call, since a semester only ever gets archived once and the two
 * sheets must never drift apart.
 *
 * After writing, applies Range-level sheet protection to the new row so it
 * cannot be edited from within Sheets itself (best-effort — Apps Script
 * protection cannot stop the workbook owner from removing the protection
 * manually, but it does stop accidental edits, which is the realistic
 * threat this guards against). Protection failures are caught and reported,
 * never allowed to fail the archive itself (the row is safely written
 * either way).
 */
function semCreateArchive_(report, charts) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(SEM_ARCHIVE_SHEET);
  if (!sheet) throw new Error('"23 Semester Archive" not found — run Semester_Lifecycle_v1.4.0_Migration.gs first.');

  var existing = readSheetRows_(SEM_ARCHIVE_SHEET, HEADER_ROW, 'Semester ID');
  if (existing.some(function (r) { return r['Semester ID'] === report.meta.semesterId; })) {
    throw new Error('Semester ' + report.meta.semesterId + ' is already archived — archives are never overwritten. Start a new semester (New Semester Wizard) before archiving again.');
  }

  var row = sheet.getLastRow() + 1;
  if (row <= HEADER_ROW) row = HEADER_ROW + 1;
  var now = new Date();
  var values = [
    report.meta.semesterId, report.meta.semesterName, report.meta.academicYear,
    report.overview.semesterStart || '', report.overview.semesterEnd || '',
    Utilities.formatDate(now, TIMEZONE, 'yyyy-MM-dd'), now, report.meta.plannerVersion,
    'Archived', true,
    JSON.stringify(report), JSON.stringify(charts || {}), JSON.stringify(report.reflection || {})
  ];
  sheet.getRange(row, 1, 1, values.length).setValues([values]);

  var protectionNote = 'Protected.';
  try {
    var protection = sheet.getRange(row, 1, 1, values.length).protect();
    protection.setDescription('Semester Archive ' + report.meta.semesterId + ' — immutable once written (v1.4.0)');
    protection.setWarningOnly(false);
    var me = Session.getEffectiveUser();
    try { protection.removeEditors(protection.getEditors().filter(function (e) { return e.getEmail() !== me.getEmail(); })); } catch (e2) { /* editor list not always mutable in every context -- warning-only protection above still applies */ }
  } catch (e) {
    protectionNote = 'Row written, but automatic sheet protection could not be applied in this context (' + e.message + ') — treat this row as read-only by convention.';
  }

  semUpdateHistory_(report);
  logAutomation_('Semester archived', report.meta.semesterId, 'Success', report.meta.semesterName + ' — ' + protectionNote);
  return { semesterId: report.meta.semesterId, row: row, protectionNote: protectionNote };
}

/** Lightweight list for the Semester Dashboard — metadata only, never the
 *  full Report/Charts/Reflection JSON blobs (kept out of the payload so the
 *  dashboard stays fast even with many archived semesters). */
function semListArchives_() {
  return readSheetRows_(SEM_ARCHIVE_SHEET, HEADER_ROW, 'Semester ID').map(function (r) {
    return {
      semesterId: r['Semester ID'], semesterName: r['Semester Name'], academicYear: r['Academic Year'],
      startDate: isoDate_(r['Start Date']), endDate: isoDate_(r['End Date']), completionDate: isoDate_(r['Completion Date']),
      createdTimestamp: r['Created Timestamp'] ? new Date(r['Created Timestamp']).toISOString() : '',
      plannerVersion: r['Planner Version'], status: r['Status'], immutable: r['Immutable'] === true
    };
  }).sort(function (a, b) { return (b.completionDate || '').localeCompare(a.completionDate || ''); });
}

/** Full detail for one archived semester (Report + Charts + Reflection),
 *  read-only, used when the user opens an archived semester's report from
 *  the Dashboard. Returns null if not found. */
function semGetArchivedReport_(semesterId) {
  var rows = readSheetRows_(SEM_ARCHIVE_SHEET, HEADER_ROW, 'Semester ID');
  var row = rows.filter(function (r) { return r['Semester ID'] === semesterId; })[0];
  if (!row) return null;
  var report, charts, reflection;
  try { report = JSON.parse(row['Report JSON'] || '{}'); } catch (e) { report = {}; }
  try { charts = JSON.parse(row['Charts JSON'] || '{}'); } catch (e) { charts = {}; }
  try { reflection = JSON.parse(row['Reflection JSON'] || '{}'); } catch (e) { reflection = {}; }
  return { report: report, charts: charts, reflection: reflection };
}

// ------------------------------------------------------------------
// MODULE TEMPLATES (spec §9) -- survive every semester. Refreshed (upserted
// by Module Code, never silently dropped) every time a semester is
// archived, so the New Semester Wizard's "reuse previous" options always
// reflect the most recently confirmed configuration.
// ------------------------------------------------------------------
function semSaveModuleTemplates_(semesterId) {
  var pd = api_getPlannerData();
  if (!pd.ok) throw new Error('Could not read workbook data: ' + pd.error);
  var modules = pd.data.modules;
  var sheet = SpreadsheetApp.getActive().getSheetByName(SEM_TEMPLATES_SHEET);
  if (!sheet) throw new Error('"25 Module Templates" not found — run Semester_Lifecycle_v1.4.0_Migration.gs first.');

  var existingRows = readSheetRows_(SEM_TEMPLATES_SHEET, HEADER_ROW, 'Module Code');
  var rowByCode = {};
  existingRows.forEach(function (r) { rowByCode[r['Module Code']] = r.__row; });

  var now = new Date();
  var saved = 0;
  modules.forEach(function (m) {
    var values = [
      m.code, m.name, m.color,
      JSON.stringify({ coordinatorName: m.coordinatorName, coordinatorEmail: m.coordinatorEmail, lecturerName: m.lecturerName, lecturerEmail: m.lecturerEmail, contactNotes: m.contactNotes }),
      JSON.stringify({ afWeight: m.afWeight, a1Weight: m.a1Weight, a2Weight: m.a2Weight, a3Weight: m.a3Weight, passMark: m.passMark, distinctionMark: m.distinctionMark }),
      JSON.stringify({ attendanceSensitive: m.attendanceSensitive, compulsoryPractical: m.compulsoryPractical }),
      JSON.stringify({ difficulty: m.difficulty, workload: m.workload, minWeeklyHours: m.minWeeklyHours, targetMark: m.targetMark, repeated: m.repeated }),
      JSON.stringify({ ruleStatus: m.ruleStatus, afMethod: m.afMethod, subminimumMode: m.subminimumMode, rule2Exception: m.rule2Exception, dcaEnabled: m.dcaEnabled, a2ToleranceOverride: m.a2ToleranceOverride, generalNotes: m.generalNotes }),
      now, semesterId || ''
    ];
    var existingRow = rowByCode[m.code];
    if (existingRow) {
      sheet.getRange(existingRow, 1, 1, values.length).setValues([values]);
    } else {
      var newRow = sheet.getLastRow() + 1;
      if (newRow <= HEADER_ROW) newRow = HEADER_ROW + 1;
      sheet.getRange(newRow, 1, 1, values.length).setValues([values]);
    }
    saved++;
  });
  logAutomation_('Module templates saved', semesterId || '-', 'Success', saved + ' module template(s) refreshed');
  return { saved: saved };
}

/** Read-only, for the New Semester Wizard's "reuse previous" step. */
function semListModuleTemplates_() {
  return readSheetRows_(SEM_TEMPLATES_SHEET, HEADER_ROW, 'Module Code').map(function (r) {
    function parseJson_(v) { try { return JSON.parse(v || '{}'); } catch (e) { return {}; } }
    return {
      code: r['Module Code'], name: r['Module Name'], color: r['Colour (hex)'],
      contacts: parseJson_(r['Contacts JSON']), weightings: parseJson_(r['Weightings JSON']),
      attendanceRules: parseJson_(r['Attendance Rules JSON']), studyPreferences: parseJson_(r['Study Preferences JSON']),
      assessmentFramework: parseJson_(r['Assessment Framework JSON']),
      lastUpdated: r['Last Updated'] ? new Date(r['Last Updated']).toISOString() : '',
      sourceSemesterId: r['Source Semester ID'] || ''
    };
  });
}

// ------------------------------------------------------------------
// HISTORICAL ANALYTICS (spec §10) -- deterministic comparisons only. No
// predictive analytics anywhere in this file.
// ------------------------------------------------------------------
function semUpdateHistory_(report) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(SEM_HISTORY_SHEET);
  if (!sheet) throw new Error('"26 Semester History" not found — run Semester_Lifecycle_v1.4.0_Migration.gs first.');
  var rev = report.revisionStatistics, res = report.resources, att = report.attendance, m = report.marks;
  var revisionCompletionRate = rev.totalTopics ? round1_((rev.topicsCompleted / rev.totalTopics) * 100) : null;
  var resourceCompletionRate = res.resourcesAdded ? round1_((res.resourcesCompleted / res.resourcesAdded) * 100) : null;
  var row = sheet.getLastRow() + 1;
  if (row <= HEADER_ROW) row = HEADER_ROW + 1;
  var values = [
    report.meta.semesterId, report.meta.semesterName, report.meta.academicYear,
    Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd'),
    report.overview.studyHours, m.overallAverage, m.highestMark, m.lowestMark, m.distinctionCount, m.passCount,
    att.attendanceRate, revisionCompletionRate, resourceCompletionRate
  ];
  sheet.getRange(row, 1, 1, values.length).setValues([values]);
}

/** Deterministic side-by-side comparison across every archived semester,
 *  oldest first, plus the simple (non-predictive) semester-over-semester
 *  delta for each numeric column -- a plain subtraction between two already
 *  -recorded historical figures is a comparison, not a prediction. */
function semHistoricalCompare_() {
  var rows = readSheetRows_(SEM_HISTORY_SHEET, HEADER_ROW, 'Semester ID').map(function (r) {
    return {
      semesterId: r['Semester ID'], semesterName: r['Semester Name'], academicYear: r['Academic Year'],
      completionDate: isoDate_(r['Completion Date']), studyHours: num_(r['Study Hours']), avgMark: num_(r['Avg Mark']),
      highestMark: num_(r['Highest Mark']), lowestMark: num_(r['Lowest Mark']), distinctionCount: num_(r['Distinction Count']),
      passCount: num_(r['Pass Count']), attendanceRate: num_(r['Attendance Rate (%)']),
      revisionCompletionRate: num_(r['Revision Completion Rate (%)']), resourceCompletionRate: num_(r['Resource Completion Rate (%)'])
    };
  }).sort(function (a, b) { return (a.completionDate || '').localeCompare(b.completionDate || ''); });

  var NUMERIC_FIELDS = ['studyHours', 'avgMark', 'highestMark', 'lowestMark', 'distinctionCount', 'passCount', 'attendanceRate', 'revisionCompletionRate', 'resourceCompletionRate'];
  for (var i = 0; i < rows.length; i++) {
    if (i === 0) { rows[i].deltaFromPrevious = null; continue; }
    var delta = {};
    NUMERIC_FIELDS.forEach(function (f) {
      var a = rows[i - 1][f], b = rows[i][f];
      delta[f] = (typeof a === 'number' && typeof b === 'number') ? round1_(b - a) : null;
    });
    rows[i].deltaFromPrevious = delta;
  }
  return rows;
}
