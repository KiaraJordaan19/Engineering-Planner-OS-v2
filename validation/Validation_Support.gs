/**
 * VALIDATION STUB — NOT PRODUCTION CODE
 *
 * Minimal shared infrastructure and disposable-fixture preparation for the
 * Milestone 0 backup/archive subsystem. This file must never be copied into a
 * production Apps Script project. Calendar access is deliberately disabled.
 */
var TIMEZONE = 'Africa/Johannesburg';
var HEADER_ROW = 4;
var SETTINGS_SHEET = '02 Settings';
var ASSIGNMENTS_SHEET = '08 Assignments';
var ASSESSMENTS_SHEET = '09 Assessments';
var INBOX_SHEET = '07 Academic Inbox';
var STUDY_SHEET = '12 Study Planner';
var STUDY_TASKS_SHEET = '21 Study Tasks';
var TIMETABLE_SHEET = '05 Timetable';
var M0_DISPOSABLE_SPREADSHEET_ID = '14sK9SkrKBLwKJNF755AU8FU7Y91rq1BQx8Z_GP6LBco';
var M0_DISPOSABLE_DRIVE_FOLDER_ID = '1AnvGrZ5OJj2rvgEreMPqj29ygCW3s-mn';

function getColMap_(sheet, headerRow) {
  headerRow = headerRow || HEADER_ROW;
  var values = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  var map = {};
  values.forEach(function (value, index) {
    var key = String(value || '').trim();
    if (key) map[key] = index + 1;
  });
  return map;
}

function col_(map, header) {
  if (!map[header]) throw new Error('Required column not found: ' + header);
  return map[header];
}

function getSetting_(name) {
  var rows = readSheetRows_(SETTINGS_SHEET, HEADER_ROW, 'Setting');
  for (var i = 0; i < rows.length; i++) {
    if (rows[i]['Setting'] === name) return rows[i]['Value'];
  }
  return '';
}

function nextId_(prefix) {
  return String(prefix || 'ID') + '-' + Utilities.getUuid();
}

function getTargetCalendar_() {
  throw new Error('Calendar access is disabled in the Milestone 0 disposable validation project.');
}

/** VALIDATION OVERRIDE: confines every backup artefact to disposable Drive. */
function semGetOrCreateBackupFolder_() {
  m0AssertDisposable_();
  var parent = DriveApp.getFolderById(M0_DISPOSABLE_DRIVE_FOLDER_ID);
  var name = 'Engineering Planner OS Backups — M0 Disposable';
  var folders = parent.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : parent.createFolder(name);
}

function semWizardGetOptions_() {
  throw new Error('Semester Wizard is outside the approved Milestone 0 validation boundary.');
}

function semWizardApply_() {
  throw new Error('Semester Wizard is outside the approved Milestone 0 validation boundary.');
}

function m0AssertDisposable_() {
  var ss = SpreadsheetApp.getActive();
  var name = ss.getName();
  if (ss.getId() !== M0_DISPOSABLE_SPREADSHEET_ID ||
      name.indexOf('DISPOSABLE') === -1 ||
      name.indexOf('VALIDATION') === -1 ||
      name.indexOf('DO NOT USE') === -1) {
    throw new Error('Safety stop: this is not the authorized disposable validation workbook.');
  }
  return ss;
}

function m0SheetDefinitions_() {
  return [
    ['02 Settings', ['Setting', 'Description', 'Value'], [
      ['Semester ID', 'Disposable fixture lineage', 'TEST-SEMESTER-2026-M0'],
      ['Semester name', 'Artificial validation semester', 'M0 Disposable Semester'],
      ['Academic year', 'Artificial validation year', 2026],
      ['Semester start date', 'Artificial date', new Date(2026, 0, 12)],
      ['Semester end date', 'Artificial date', new Date(2026, 5, 30)],
      ['Semester status', 'Lifecycle state', 'Active'],
      ['Semester manually finished (TRUE/FALSE)', 'Lifecycle flag', false],
      ['Planner version', 'Validation build marker', '1.5.0-M0-VALIDATION'],
      ['Pass mark (%)', 'Artificial threshold', 50],
      ['Distinction mark (%)', 'Artificial threshold', 75],
      ['Target final mark (%)', 'Artificial target', 65],
      ['Calendar sync enabled (TRUE/FALSE)', 'Must remain disabled', false],
      ['Planner Calendar ID (read/write)', 'Intentionally blank', '']
    ]],
    ['03 Modules', ['Module code', 'Module name', 'Short name', 'Active status', 'Colour (hex)', 'Conceptual difficulty (1-5)', 'Workload intensity (1-5)', 'Repeated module (TRUE/FALSE)', 'Target mark (%)', 'Min weekly study hours', 'Coordinator email', 'Fixture Formula Result'], [
      ['TEST-MOD-01', 'Artificial Engineering Module', 'TEST 101', 'Active', '#336699', 3, 4, false, 65.5, 0, 'disposable@example.invalid', '']
    ]],
    ['04 Module Rules', ['Module code', 'Rule status', 'AF weighting', 'A1 weighting', 'A2 weighting', 'A3 weight (%) (verified)', 'Pass mark', 'Distinction mark', 'AF calculation method', 'General notes', 'Rule 2 exception (TRUE/FALSE)', 'DCA enabled (TRUE/FALSE)'], [
      ['TEST-MOD-01', 'Verified', 20, 30, 50, 50, 50, 75, 'WEIGHTED', 'Artificial rule only — punctuation: !@#; unicode: Ω', false, false]
    ]],
    ['07 Academic Inbox', ['Inbox ID', 'Module', 'Item type', 'Title', 'Due date', 'Due time', 'Venue or link', 'Notes', 'Priority', 'Processed', 'Send to Calendar', 'Calendar Event ID'], [
      ['INBOX-TEST-01', 'Artificial Engineering Module', 'Other', 'Artificial inbox item', new Date(2026, 1, 1), new Date(1899, 11, 30, 9, 30), 'https://example.invalid', 'Blank calendar ID by design', 'Low', false, false, '']
    ]],
    ['08 Assignments', ['Assignment ID', 'Module', 'Title', 'Description', 'Due date', 'Due time', 'Days remaining', 'Priority', 'Estimated hours', 'Hours completed', 'Status', 'Completion %', 'Notes', 'Calendar Event ID', 'Sync to Calendar'], [
      ['ASSIGN-TEST-01', 'Artificial Engineering Module', 'Artificial assignment', 'Disposable fixture only', new Date(2026, 1, 10), new Date(1899, 11, 30, 14, 0), 24, 'Medium', 2.5, 0, 'Not started', 0, '', '', false]
    ]],
    ['09 Assessments', ['Assessment ID', 'Module', 'Assessment type', 'Title', 'Date', 'Start time', 'Venue', 'Status', 'Days remaining', 'Notes', 'Calendar Event ID', 'Sync to Calendar'], [
      ['ASSESS-TEST-01', 'Artificial Engineering Module', 'Test', 'Artificial assessment', new Date(2026, 2, 3), new Date(1899, 11, 30, 8, 0), 'Test Room', 'Planned', 45, 'No real learner data', '', false]
    ]],
    ['10 AF Components', ['AF Item ID', 'Module', 'Item name', 'Mark', 'Maximum mark', 'Percentage', 'Date', 'Written', 'Excused or excluded', 'Included in AF', 'Notes'], [
      ['AF-TEST-01', 'Artificial Engineering Module', 'Artificial quiz', 0, 10, 0, new Date(2026, 1, 5), true, false, true, 'Zero is intentional, not missing']
    ]],
    ['11 Marks Tracker', ['Module code', 'Module', 'Rule status', 'AF %', 'A1 %', 'A2 %', 'A3 %', 'MTD', 'Provisional final (after A2)', 'Final / provisional (after A3)', 'Pass status', 'Notes'], [
      ['TEST-MOD-01', 'Artificial Engineering Module', 'Verified', 0, 61.25, '', '', 42.125, '', '', 'In progress', 'Blank and zero are distinct']
    ], 6],
    ['12 Study Planner', ['Study Session ID', 'Date', 'Start time', 'End time', 'Duration (min)', 'Module', 'Study type', 'Topic', 'Priority', 'Planned', 'Completed', 'Notes', 'Calendar Event ID', 'Status', 'Objective', 'Actual duration (min)', 'Confidence after', 'Origin'], [
      ['STUDY-TEST-01', new Date(2026, 0, 20), new Date(1899, 11, 30, 18, 0), new Date(1899, 11, 30, 19, 30), 90, 'Artificial Engineering Module', 'Practice', 'Fixture topic', 'High', true, false, 'Artificial session', '', 'Planned', 'Test recovery fidelity', '', '', 'Manual']
    ]],
    ['13 Revision Tracker', ['Module', 'Topic', 'Confidence', 'Priority', 'Next revision date', 'Completed', 'Notes', 'Total actual study minutes', 'Session count', 'Last revised', 'Completed task count', 'Score', 'Weak topic (TRUE/FALSE)', 'Planned-study count', 'Independent-study count', 'Source', 'Weak topic since'], [
      ['Artificial Engineering Module', 'Fixture topic', 2, 'High', new Date(2026, 0, 25), false, 'Needs review', 0, 0, '', 0, -1.5, true, 1, 0, 'Validation fixture', new Date(2026, 0, 1)]
    ]],
    ['15 Resources', ['Resource ID', 'Module', 'Resource type', 'Title', 'Link or file reference', 'Topic', 'Description', 'Priority', 'Reviewed', 'Notes', 'Status', 'Progress Flag A', 'Progress Flag B', 'Progress Flag C', 'Progress Flag D', 'Flashcards created (TRUE/FALSE)', 'Score', 'Score updated'], [
      ['RESOURCE-TEST-01', 'Artificial Engineering Module', 'Past paper', 'Artificial resource', 'https://example.invalid/resource', 'Fixture topic', 'Disposable only', 'Low', false, '', 'Not started', false, false, false, false, false, 0, '']
    ]],
    ['16 Notes', ['Timestamp', 'Module', 'Title', 'Note'], [
      [new Date(2026, 0, 15, 12, 34, 56), 'Artificial Engineering Module', 'Fixture note', 'Long-form artificial text\nwith a preserved newline and unicode ✓.']
    ]],
    ['17 Archive', ['Semester', 'Module', 'Final mark', 'Result'], [
      ['Previous artificial semester', 'Artificial Engineering Module', 76.25, 'Pass']
    ], 12],
    ['19 Attendance', ['Attendance ID', 'Date', 'Module', 'Session type', 'Attended', 'Notes', 'Status', 'Session key'], [
      ['ATTEND-TEST-01', new Date(2026, 0, 16), 'Artificial Engineering Module', 'Lecture', true, '', 'Attended', 'TEST-MOD-01|2026-01-16|LECTURE']
    ]],
    ['20 Automation Log', ['Timestamp', 'Action', 'Item', 'Result', 'Detail'], []],
    ['21 Study Tasks', ['Task ID', 'Study Session ID', 'Module', 'Topic', 'Task type', 'Task text', 'Resource ID', 'Completed', 'Score', 'Notes', 'Sort order', 'Last updated'], [
      ['TASK-TEST-01', 'STUDY-TEST-01', 'Artificial Engineering Module', 'Fixture topic', 'Practice questions', 'Artificial task', '', false, 0, '', 1, new Date(2026, 0, 15)]
    ]],
    ['23 Semester Archive', ['Semester ID', 'Semester Name', 'Academic Year', 'Start Date', 'End Date', 'Completion Date', 'Created Timestamp', 'Planner Version', 'Status', 'Immutable', 'Report JSON', 'Charts JSON', 'Reflection JSON'], []],
    ['24 Semester Backups', ['Backup ID', 'Semester ID', 'Timestamp', 'Planner Version', 'Drive File ID', 'Drive File URL', 'Sheets Included', 'Notes'], []],
    ['25 Module Templates', ['Module Code', 'Module Name', 'Colour (hex)', 'Contacts JSON', 'Weightings JSON', 'Attendance Rules JSON', 'Study Preferences JSON', 'Assessment Framework JSON', 'Last Updated', 'Source Semester ID'], []],
    ['26 Semester History', ['Semester ID', 'Semester Name', 'Academic Year', 'Completion Date', 'Study Hours', 'Avg Mark', 'Highest Mark', 'Lowest Mark', 'Distinction Count', 'Pass Count', 'Attendance Rate (%)', 'Revision Completion Rate (%)', 'Resource Completion Rate (%)'], []],
    ['27 Semester Workflows', ['Workflow Token', 'Semester ID', 'Backup ID', 'Archive Semester ID', 'Prepared At', 'Expires At', 'Status', 'Consumed At', 'Planner Version', 'Last Stage', 'Last Error'], []]
  ];
}

function m0PrepareDisposableFixture() {
  var ss = m0AssertDisposable_();
  var definitions = m0SheetDefinitions_();
  definitions.forEach(function (definition) {
    var name = definition[0];
    var headers = definition[1];
    var rows = definition[2];
    var headerRow = definition[3] || HEADER_ROW;
    var sheet = ss.getSheetByName(name) || ss.insertSheet(name);
    sheet.clear();
    sheet.getRange(headerRow, 1, 1, headers.length).setValues([headers]);
    if (rows.length) sheet.getRange(headerRow + 1, 1, rows.length, headers.length).setValues(rows);
  });

  var modules = ss.getSheetByName('03 Modules');
  modules.getRange(HEADER_ROW + 1, 12).setFormula('=1+1');
  SpreadsheetApp.flush();

  var defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) ss.deleteSheet(defaultSheet);

  var report = m0PreparationReport_();
  console.log(JSON.stringify(report));
  return report;
}

function m0CanonicalValue_(value) {
  if (value instanceof Date) return { type: 'date', value: value.toISOString() };
  if (value === '') return { type: 'blank', value: '' };
  if (typeof value === 'number') return { type: 'number', value: String(value) };
  if (typeof value === 'boolean') return { type: 'boolean', value: value ? 'true' : 'false' };
  return { type: 'string', value: String(value) };
}

function m0Hex_(bytes) {
  return bytes.map(function (value) {
    var unsigned = value < 0 ? value + 256 : value;
    return ('0' + unsigned.toString(16)).slice(-2);
  }).join('');
}

function m0PreparationReport_() {
  var ss = m0AssertDisposable_();
  var definitions = m0SheetDefinitions_();
  var inventory = definitions.map(function (definition) {
    var sheet = ss.getSheetByName(definition[0]);
    return {
      name: definition[0],
      present: !!sheet,
      rows: sheet ? sheet.getLastRow() : 0,
      columns: sheet ? sheet.getLastColumn() : 0
    };
  });
  var recoverySheet = ss.getSheetByName('03 Modules');
  var values = recoverySheet.getDataRange().getValues().map(function (row) {
    return row.map(m0CanonicalValue_);
  });
  var digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    JSON.stringify(values),
    Utilities.Charset.UTF_8
  );
  var report = {
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName(),
    sheetCount: ss.getSheets().length,
    inventory: inventory,
    recoveryComparison: {
      sheet: '03 Modules',
      algorithm: 'SHA-256',
      canonicalization: 'typed JSON cells in row and column order; Date.toISOString()',
      checksum: m0Hex_(digest)
    },
    calendarAccess: 'DISABLED',
    source: 'VALIDATION STUB — NOT PRODUCTION CODE'
  };
  console.log(JSON.stringify(report));
  return report;
}

// ---------------------------------------------------------------------------
// GATE A VALIDATION RUNNERS — DISPOSABLE PROJECT ONLY; NOT PRODUCTION CODE.
// Each public runner executes one authorized blocking test and throws if the
// observed result differs from the Validation Record.
// ---------------------------------------------------------------------------
function m0GateASafetyCheck() {
  var folder = semGetOrCreateBackupFolder_();
  var parents = folder.getParents(), parentIds = [];
  while (parents.hasNext()) parentIds.push(parents.next().getId());
  if (parentIds.indexOf(M0_DISPOSABLE_DRIVE_FOLDER_ID) === -1) {
    throw new Error('Safety stop: backup folder is not inside the authorized disposable Drive folder.');
  }
  var result = {
    status: 'PASS', spreadsheetId: m0AssertDisposable_().getId(),
    backupFolderId: folder.getId(), parentIds: parentIds,
    calendarAccess: 'DISABLED', productionAccess: 'NONE'
  };
  console.log(JSON.stringify(result));
  return result;
}

function m0Assert_(condition, message) {
  if (!condition) throw new Error('GATE A ASSERTION FAILED: ' + message);
}

function m0ControlBackupId_() {
  var id = getSetting_('M0 Gate A Control Backup ID');
  if (!id) throw new Error('Gate A control backup ID has not been recorded.');
  return String(id);
}

function m0SetSetting_(name, description, value) {
  var sheet = m0AssertDisposable_().getSheetByName(SETTINGS_SHEET);
  var rows = readSheetRows_(SETTINGS_SHEET, HEADER_ROW, 'Setting');
  var existing = rows.filter(function (row) { return row['Setting'] === name; })[0];
  var targetRow = existing ? existing.__row : Math.max(sheet.getLastRow() + 1, HEADER_ROW + 1);
  sheet.getRange(targetRow, 1, 1, 3).setValues([[name, description, value]]);
}

function m0BackupRow_(backupId) {
  return readSheetRows_(SEM_BACKUPS_SHEET, HEADER_ROW, 'Backup ID').filter(function (row) {
    return String(row['Backup ID']) === String(backupId);
  })[0];
}

function m0AppendBackupRow_(backupId, file, semesterId, plannerVersion, sheetsIncluded, notes) {
  var sheet = m0AssertDisposable_().getSheetByName(SEM_BACKUPS_SHEET);
  var row = Math.max(sheet.getLastRow() + 1, HEADER_ROW + 1);
  sheet.getRange(row, 1, 1, 8).setValues([[
    backupId, semesterId, new Date(), plannerVersion, file.getId(), file.getUrl(), sheetsIncluded, notes
  ]]);
  return row;
}

function m0DeleteBackupRow_(backupId) {
  var row = m0BackupRow_(backupId);
  if (row) m0AssertDisposable_().getSheetByName(SEM_BACKUPS_SHEET).deleteRow(row.__row);
}

function m0ControlPayload_() {
  var row = m0BackupRow_(m0ControlBackupId_());
  if (!row) throw new Error('Gate A control backup metadata row is missing.');
  var file = DriveApp.getFileById(row['Drive File ID']);
  return { row: row, file: file, payload: JSON.parse(file.getBlob().getDataAsString()) };
}

function m0GateAId_(testNumber) {
  return 'M0-GA-' + testNumber + '-' + Utilities.getUuid();
}

function m0CreateTestFile_(name, content) {
  return semGetOrCreateBackupFolder_().createFile(name, content, MimeType.PLAIN_TEXT);
}

function m0FolderFileCount_() {
  var files = semGetOrCreateBackupFolder_().getFiles();
  var count = 0;
  while (files.hasNext()) { files.next(); count++; }
  return count;
}

function m0RunGateATest01() {
  var started = new Date();
  var beforeRows = readSheetRows_(SEM_BACKUPS_SHEET, HEADER_ROW, 'Backup ID').length;
  var result = semCreateBackup_('TEST-SEMESTER-2026-M0');
  var row = m0BackupRow_(result.backupId);
  var file = DriveApp.getFileById(result.fileId);
  var payload = JSON.parse(file.getBlob().getDataAsString());
  m0Assert_(!!row, 'matching metadata row was not created');
  m0Assert_(String(row['Drive File ID']) === result.fileId, 'metadata file ID differs from response');
  m0Assert_(payload.backupId === result.backupId, 'payload backup ID differs from response');
  m0Assert_(payload.semesterId === 'TEST-SEMESTER-2026-M0', 'payload semester ID differs');
  m0Assert_(semBackupSheetSetMatches_(Object.keys(payload.sheets)), 'payload sheet set differs');
  m0Assert_(!!semBackupChecksumFromNotes_(row['Notes']), 'metadata fingerprint is missing');
  m0Assert_(readSheetRows_(SEM_BACKUPS_SHEET, HEADER_ROW, 'Backup ID').length === beforeRows + 1, 'metadata row count did not increase by one');
  m0SetSetting_('M0 Gate A Control Backup ID', 'VALIDATION STATE — control backup for Gate A', result.backupId);
  var evidence = {
    testId: 'M0-LV-01', status: 'PASS', durationMs: new Date() - started,
    response: result, metadataRow: row.__row, formatVersion: payload.formatVersion,
    semesterId: payload.semesterId, plannerVersion: payload.plannerVersion,
    spreadsheetId: payload.spreadsheetId, indexedSheetCount: String(row['Sheets Included']).split(',').length,
    payloadSheetCount: Object.keys(payload.sheets).length, fingerprint: semBackupChecksumFromNotes_(row['Notes'])
  };
  console.log(JSON.stringify(evidence));
  return evidence;
}

function m0RunGateATest02() {
  var started = new Date();
  var backupId = m0ControlBackupId_();
  var verification = semVerifyBackup_(backupId);
  m0Assert_(verification.ok === true, 'valid control backup did not verify');
  m0Assert_(verification.fileFound === true && verification.jsonValid === true, 'valid control file/JSON flags differ');
  m0Assert_(verification.checksumValid === true, 'valid control checksum flag differs');
  m0Assert_(Array.isArray(verification.missingSheets) && verification.missingSheets.length === 0, 'valid control reports missing sheets');
  var evidence = { testId: 'M0-LV-02', status: 'PASS', durationMs: new Date() - started, backupId: backupId, verification: verification };
  console.log(JSON.stringify(evidence));
  return evidence;
}

function m0RunGateATest03() {
  var started = new Date(), control = m0ControlPayload_(), id = m0GateAId_('03');
  var file = m0CreateTestFile_('M0-LV-03-malformed-' + id + '.json', '{not valid JSON');
  var rowNumber = m0AppendBackupRow_(id, file, control.payload.semesterId, control.payload.plannerVersion, SEM_BACKUP_SHEETS.join(', '), control.row['Notes']);
  var verification = semVerifyBackup_(id);
  var cleanup = semCleanupIncompleteBackupFile_(file);
  m0DeleteBackupRow_(id);
  m0Assert_(verification.ok === false && verification.fileFound === true && verification.jsonValid === false, 'malformed JSON flags differ');
  m0Assert_(String(verification.reason).indexOf('not valid JSON') !== -1, 'malformed JSON reason differs');
  m0Assert_(cleanup.status === 'SUCCEEDED', 'malformed artefact cleanup failed');
  var evidence = { testId: 'M0-LV-03', status: 'PASS', durationMs: new Date() - started, artefactId: id, fileId: file.getId(), metadataRow: rowNumber, mutation: 'content replaced with {not valid JSON', verification: verification, cleanup: cleanup.status };
  console.log(JSON.stringify(evidence));
  return evidence;
}

function m0RunGateATest04() {
  var started = new Date(), control = m0ControlPayload_(), id = m0GateAId_('04');
  var payload = JSON.parse(JSON.stringify(control.payload));
  var before = payload.spreadsheetName;
  payload.backupId = id;
  payload.spreadsheetName = before + ' — ALTERED';
  var file = m0CreateTestFile_('M0-LV-04-altered-' + id + '.json', JSON.stringify(payload));
  var rowNumber = m0AppendBackupRow_(id, file, payload.semesterId, payload.plannerVersion, SEM_BACKUP_SHEETS.join(', '), control.row['Notes']);
  var verification = semVerifyBackup_(id);
  var cleanup = semCleanupIncompleteBackupFile_(file);
  m0DeleteBackupRow_(id);
  m0Assert_(verification.ok === false && String(verification.reason).indexOf('fingerprint') !== -1, 'altered JSON did not fail fingerprint verification');
  m0Assert_(cleanup.status === 'SUCCEEDED', 'altered artefact cleanup failed');
  var evidence = { testId: 'M0-LV-04', status: 'PASS', durationMs: new Date() - started, artefactId: id, fileId: file.getId(), metadataRow: rowNumber, changedField: 'spreadsheetName', before: before, after: payload.spreadsheetName, expectedFingerprint: semBackupChecksumFromNotes_(control.row['Notes']), verification: verification, cleanup: cleanup.status };
  console.log(JSON.stringify(evidence));
  return evidence;
}

function m0RunGateATest05() {
  var started = new Date(), control = m0ControlPayload_(), id = m0GateAId_('05');
  var payload = JSON.parse(JSON.stringify(control.payload));
  payload.backupId = id;
  payload.contentChecksum = semBackupChecksum_(payload);
  var notes = semBackupNotes_(payload.contentChecksum);
  var file = m0CreateTestFile_('M0-LV-05-deleted-' + id + '.json', JSON.stringify(payload));
  var rowNumber = m0AppendBackupRow_(id, file, payload.semesterId, payload.plannerVersion, SEM_BACKUP_SHEETS.join(', '), notes);
  var beforeTrash = semVerifyBackup_(id);
  file.setTrashed(true);
  var verification = semVerifyBackup_(id);
  m0DeleteBackupRow_(id);
  m0Assert_(beforeTrash.ok === true, 'deleted-file test artefact was invalid before trashing');
  m0Assert_(verification.ok === false && verification.fileFound === false, 'trashed Drive file was not reported as missing');
  m0Assert_(verification.reason === 'Backup Drive file is trashed and unavailable.', 'trashed Drive file reason differs from the frozen contract');
  var evidence = { testId: 'M0-LV-05', status: 'PASS', durationMs: new Date() - started, artefactId: id, fileId: file.getId(), metadataRow: rowNumber, trashed: true, verificationBeforeTrash: beforeTrash, verification: verification, cleanup: 'File remains trashed; metadata row removed' };
  console.log(JSON.stringify(evidence));
  return evidence;
}

function m0RunGateATest06() {
  var started = new Date(), control = m0ControlPayload_(), id = m0GateAId_('06');
  var payload = JSON.parse(JSON.stringify(control.payload));
  var payloadId = payload.backupId;
  var file = m0CreateTestFile_('M0-LV-06-wrong-id-' + id + '.json', JSON.stringify(payload));
  var rowNumber = m0AppendBackupRow_(id, file, payload.semesterId, payload.plannerVersion, SEM_BACKUP_SHEETS.join(', '), control.row['Notes']);
  var verification = semVerifyBackup_(id);
  var cleanup = semCleanupIncompleteBackupFile_(file);
  m0DeleteBackupRow_(id);
  m0Assert_(verification.ok === false && String(verification.reason).indexOf('Backup ID does not match') !== -1, 'wrong backup ID reason differs');
  m0Assert_(cleanup.status === 'SUCCEEDED', 'wrong-ID artefact cleanup failed');
  var evidence = { testId: 'M0-LV-06', status: 'PASS', durationMs: new Date() - started, requestedId: id, metadataId: id, payloadId: payloadId, fileId: file.getId(), metadataRow: rowNumber, verification: verification, cleanup: cleanup.status };
  console.log(JSON.stringify(evidence));
  return evidence;
}

function m0RunGateATest07() {
  var started = new Date(), control = m0ControlPayload_(), id = m0GateAId_('07');
  var payload = JSON.parse(JSON.stringify(control.payload));
  var expectedId = payload.spreadsheetId;
  payload.backupId = id;
  payload.spreadsheetId = 'DISPOSABLE-WRONG-SPREADSHEET-ID';
  payload.contentChecksum = semBackupChecksum_(payload);
  var file = m0CreateTestFile_('M0-LV-07-wrong-sheet-' + id + '.json', JSON.stringify(payload));
  var rowNumber = m0AppendBackupRow_(id, file, payload.semesterId, payload.plannerVersion, SEM_BACKUP_SHEETS.join(', '), semBackupNotes_(payload.contentChecksum));
  var verification = semVerifyBackup_(id);
  var cleanup = semCleanupIncompleteBackupFile_(file);
  m0DeleteBackupRow_(id);
  m0Assert_(verification.ok === false && verification.reason === 'Backup belongs to a different spreadsheet.', 'wrong spreadsheet ID reason differs');
  m0Assert_(cleanup.status === 'SUCCEEDED', 'wrong-spreadsheet artefact cleanup failed');
  var evidence = { testId: 'M0-LV-07', status: 'PASS', durationMs: new Date() - started, expectedSpreadsheetId: expectedId, alteredSpreadsheetId: payload.spreadsheetId, fileId: file.getId(), metadataRow: rowNumber, verification: verification, cleanup: cleanup.status };
  console.log(JSON.stringify(evidence));
  return evidence;
}

function m0RunGateATest08() {
  var started = new Date(), ss = m0AssertDisposable_(), requiredName = '03 Modules', temporaryName = '03 Modules — M0-LV-08 TEMP';
  var sheet = ss.getSheetByName(requiredName), rowsBefore = readSheetRows_(SEM_BACKUPS_SHEET, HEADER_ROW, 'Backup ID').length, filesBefore = m0FolderFileCount_(), creationError = '';
  sheet.setName(temporaryName);
  try { semCreateBackup_('TEST-SEMESTER-2026-M0'); } catch (e) { creationError = e.message; }
  sheet.setName(requiredName);
  SpreadsheetApp.flush();
  var rowsAfter = readSheetRows_(SEM_BACKUPS_SHEET, HEADER_ROW, 'Backup ID').length, filesAfter = m0FolderFileCount_();
  m0Assert_(creationError.indexOf(requiredName) !== -1, 'missing live sheet did not block creation');
  m0Assert_(rowsBefore === rowsAfter && filesBefore === filesAfter, 'missing-sheet creation produced an artefact');

  var control = m0ControlPayload_(), id = m0GateAId_('08'), payload = JSON.parse(JSON.stringify(control.payload));
  payload.backupId = id;
  delete payload.sheets[requiredName];
  payload.contentChecksum = semBackupChecksum_(payload);
  var file = m0CreateTestFile_('M0-LV-08-missing-sheet-' + id + '.json', JSON.stringify(payload));
  var rowNumber = m0AppendBackupRow_(id, file, payload.semesterId, payload.plannerVersion, SEM_BACKUP_SHEETS.join(', '), semBackupNotes_(payload.contentChecksum));
  var verification = semVerifyBackup_(id), cleanup = semCleanupIncompleteBackupFile_(file);
  m0DeleteBackupRow_(id);
  m0Assert_(verification.ok === false && verification.missingSheets.indexOf(requiredName) !== -1, 'crafted missing sheet was not identified');
  m0Assert_(cleanup.status === 'SUCCEEDED', 'missing-sheet artefact cleanup failed');
  var evidence = { testId: 'M0-LV-08', status: 'PASS', durationMs: new Date() - started, omittedSheet: requiredName, creationError: creationError, driveFileCountBefore: filesBefore, driveFileCountAfter: filesAfter, metadataCountBefore: rowsBefore, metadataCountAfter: rowsAfter, craftedFileId: file.getId(), craftedMetadataRow: rowNumber, verification: verification, cleanup: cleanup.status, fixtureRestored: !!ss.getSheetByName(requiredName) && !ss.getSheetByName(temporaryName) };
  console.log(JSON.stringify(evidence));
  return evidence;
}

function m0RunGateATest09() {
  var started = new Date(), ss = m0AssertDisposable_(), requiredName = '04 Module Rules', temporaryName = '04 Module Rules — M0-LV-09 TEMP';
  var original = ss.getSheetByName(requiredName), rowsBefore = readSheetRows_(SEM_BACKUPS_SHEET, HEADER_ROW, 'Backup ID').length, filesBefore = m0FolderFileCount_(), creationError = '';
  original.setName(temporaryName);
  ss.insertSheet(requiredName);
  try { semCreateBackup_('TEST-SEMESTER-2026-M0'); } catch (e) { creationError = e.message; }
  ss.deleteSheet(ss.getSheetByName(requiredName));
  original.setName(requiredName);
  SpreadsheetApp.flush();
  var rowsAfter = readSheetRows_(SEM_BACKUPS_SHEET, HEADER_ROW, 'Backup ID').length, filesAfter = m0FolderFileCount_();
  m0Assert_(creationError.indexOf(requiredName) !== -1, 'empty live sheet did not block creation');
  m0Assert_(rowsBefore === rowsAfter && filesBefore === filesAfter, 'empty-sheet creation produced an artefact');

  var control = m0ControlPayload_(), id = m0GateAId_('09'), payload = JSON.parse(JSON.stringify(control.payload));
  payload.backupId = id;
  payload.sheets[requiredName].values = [];
  payload.contentChecksum = semBackupChecksum_(payload);
  var file = m0CreateTestFile_('M0-LV-09-empty-sheet-' + id + '.json', JSON.stringify(payload));
  var rowNumber = m0AppendBackupRow_(id, file, payload.semesterId, payload.plannerVersion, SEM_BACKUP_SHEETS.join(', '), semBackupNotes_(payload.contentChecksum));
  var verification = semVerifyBackup_(id), cleanup = semCleanupIncompleteBackupFile_(file);
  m0DeleteBackupRow_(id);
  m0Assert_(verification.ok === false && verification.missingSheets.indexOf(requiredName) !== -1, 'crafted empty sheet was not identified');
  m0Assert_(cleanup.status === 'SUCCEEDED', 'empty-sheet artefact cleanup failed');
  var evidence = { testId: 'M0-LV-09', status: 'PASS', durationMs: new Date() - started, emptiedSheet: requiredName, creationError: creationError, driveFileCountBefore: filesBefore, driveFileCountAfter: filesAfter, metadataCountBefore: rowsBefore, metadataCountAfter: rowsAfter, craftedFileId: file.getId(), craftedMetadataRow: rowNumber, verification: verification, cleanup: cleanup.status, fixtureRestored: !!ss.getSheetByName(requiredName) && !ss.getSheetByName(temporaryName) };
  console.log(JSON.stringify(evidence));
  return evidence;
}

function m0RunGateATest10() {
  var started = new Date(), control = m0ControlPayload_(), id = m0GateAId_('10');
  var payload = JSON.parse(JSON.stringify(control.payload));
  payload.backupId = id;
  delete payload.formatVersion;
  delete payload.contentChecksum;
  var file = m0CreateTestFile_('M0-LV-10-legacy-' + id + '.json', JSON.stringify(payload));
  var legacyNotes = 'Legacy backup — no checksum';
  var rowNumber = m0AppendBackupRow_(id, file, payload.semesterId, payload.plannerVersion, SEM_BACKUP_SHEETS.join(', '), legacyNotes);
  var listable = semListBackups_().some(function (backup) { return backup.backupId === id; });
  var verification = semVerifyBackup_(id), cleanup = semCleanupIncompleteBackupFile_(file);
  m0DeleteBackupRow_(id);
  m0Assert_(listable, 'legacy backup was not listable');
  m0Assert_(verification.ok === false && String(verification.reason).indexOf('no verifiable SHA-256 fingerprint') !== -1, 'legacy rejection reason differs');
  m0Assert_(cleanup.status === 'SUCCEEDED', 'legacy artefact cleanup failed');
  var evidence = { testId: 'M0-LV-10', status: 'PASS', durationMs: new Date() - started, backupId: id, fileId: file.getId(), metadataRow: rowNumber, legacyNotes: legacyNotes, listable: listable, verification: verification, cleanup: cleanup.status };
  console.log(JSON.stringify(evidence));
  return evidence;
}


// ---------------------------------------------------------------------------
// GATE B VALIDATION RUNNERS — DISPOSABLE PROJECT ONLY; NOT PRODUCTION CODE.
// Owner-approved design: docs/validation/Gate_B_Validation_Runner_Specification.md
// These functions orchestrate public production APIs. They contain no semester
// lifecycle, backup-verification, archive, or reset business rules.
// ---------------------------------------------------------------------------
var M0_GATE_B_TARGETS = [
  { name: '08 Assignments', header: 4 },
  { name: '09 Assessments', header: 4 },
  { name: '12 Study Planner', header: 4 },
  { name: '21 Study Tasks', header: 4 },
  { name: '13 Revision Tracker', header: 4 },
  { name: '11 Marks Tracker', header: 6 },
  { name: '10 AF Components', header: 4 },
  { name: '19 Attendance', header: 4 },
  { name: '07 Academic Inbox', header: 4 },
  { name: '20 Automation Log', header: 4 }
];
var M0_GATE_B_INFRA = [
  '02 Settings', '23 Semester Archive', '24 Semester Backups',
  '25 Module Templates', '26 Semester History', '27 Semester Workflows'
];
var M0_GATE_B_WORKFLOW_HEADERS = [
  'Workflow Token', 'Semester ID', 'Backup ID', 'Archive Semester ID',
  'Prepared At', 'Expires At', 'Status', 'Consumed At', 'Planner Version',
  'Last Stage', 'Last Error'
];
var M0_GATE_B_ARCHIVE_HEADERS = [
  'Semester ID', 'Semester Name', 'Academic Year', 'Start Date', 'End Date',
  'Completion Date', 'Created Timestamp', 'Planner Version', 'Status',
  'Immutable', 'Report JSON', 'Charts JSON', 'Reflection JSON'
];
var M0_GATE_B_TEMPLATE_HEADERS = [
  'Module Code', 'Module Name', 'Colour (hex)', 'Contacts JSON', 'Weightings JSON',
  'Attendance Rules JSON', 'Study Preferences JSON', 'Assessment Framework JSON',
  'Last Updated', 'Source Semester ID'
];

function m0BError_(classification, message) {
  var error = new Error(message);
  error.m0Classification = classification;
  return error;
}

function m0BAssertHarness_(condition, message) {
  if (!condition) throw m0BError_('VALIDATION HARNESS FAILURE', message);
}

function m0BAssertProduction_(condition, message) {
  if (!condition) throw m0BError_('IMPLEMENTATION FAILURE', message);
}

function m0BAssertEnvironment_(condition, message) {
  if (!condition) throw m0BError_('ENVIRONMENT BLOCKED', message);
}

function m0BAssertCleanup_(condition, message) {
  if (!condition) throw m0BError_('CLEANUP INCOMPLETE', message);
}

function m0BCloneValue_(value) {
  if (value instanceof Date) return new Date(value.getTime());
  if (Array.isArray(value)) return value.map(m0BCloneValue_);
  return value;
}

function m0BSnapshotSheet_(name) {
  var sheet = m0AssertDisposable_().getSheetByName(name);
  m0BAssertEnvironment_(!!sheet, 'required validation sheet missing: ' + name);
  var rows = sheet.getLastRow(), columns = sheet.getLastColumn();
  var values = rows && columns ? sheet.getRange(1, 1, rows, columns).getValues() : [];
  return {
    name: name, sheetId: sheet.getSheetId(), rows: rows, columns: columns,
    values: values.map(function (row) { return row.map(m0BCloneValue_); }),
    checksum: m0BChecksum_(values)
  };
}

function m0BChecksum_(values) {
  var canonical = values.map(function (row) {
    return row.map(m0CanonicalValue_);
  });
  return m0Hex_(Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    JSON.stringify(canonical),
    Utilities.Charset.UTF_8
  ));
}

function m0BSnapshotSet_(names) {
  var result = {};
  names.forEach(function (name) { result[name] = m0BSnapshotSheet_(name); });
  return result;
}

function m0BTargetNames_() {
  return M0_GATE_B_TARGETS.map(function (target) { return target.name; });
}

function m0BCurrentState_(names) {
  var result = {};
  names.forEach(function (name) {
    var snapshot = m0BSnapshotSheet_(name);
    result[name] = {
      sheetId: snapshot.sheetId, rows: snapshot.rows,
      columns: snapshot.columns, checksum: snapshot.checksum
    };
  });
  return result;
}

function m0BNonTargetBaseline_(permitted) {
  var excluded = m0BTargetNames_().concat(permitted || []);
  var names = m0AssertDisposable_().getSheets().map(function (sheet) {
    return sheet.getName();
  }).filter(function (name) { return excluded.indexOf(name) === -1; });
  return m0BSnapshotSet_(names);
}

function m0BAssertSame_(baseline, names, classification) {
  var assertSame = classification === 'CLEANUP INCOMPLETE' ?
    m0BAssertCleanup_ : m0BAssertProduction_;
  names.forEach(function (name) {
    var current = m0BSnapshotSheet_(name), expected = baseline[name];
    assertSame(current.sheetId === expected.sheetId, name + ' sheet identity changed');
    assertSame(current.rows === expected.rows && current.columns === expected.columns,
      name + ' dimensions changed');
    assertSame(current.checksum === expected.checksum, name + ' values changed');
  });
}

function m0BAssertCleared_(name, header) {
  var sheet = m0AssertDisposable_().getSheetByName(name);
  m0BAssertProduction_(!!sheet, 'destructive target missing: ' + name);
  m0BAssertProduction_(sheet.getLastRow() <= header, name + ' retained data below header ' + header);
}

function m0BRestoreBusinessSheet_(snapshot, expectedCurrentChecksum) {
  m0BAssertCleanup_(M0_GATE_B_INFRA.indexOf(snapshot.name) === -1,
    'whole-sheet infrastructure restoration is prohibited: ' + snapshot.name);
  var sheet = m0AssertDisposable_().getSheetByName(snapshot.name);
  m0BAssertCleanup_(!!sheet && sheet.getSheetId() === snapshot.sheetId,
    'cleanup identity mismatch for ' + snapshot.name);
  var current = m0BSnapshotSheet_(snapshot.name);
  if (current.checksum === snapshot.checksum &&
      current.rows === snapshot.rows && current.columns === snapshot.columns) return;
  m0BAssertCleanup_(current.checksum === expectedCurrentChecksum,
    'cleanup refused unexpected current state for ' + snapshot.name);
  var maxRows = Math.max(sheet.getLastRow(), snapshot.rows);
  var maxColumns = Math.max(sheet.getLastColumn(), snapshot.columns);
  if (maxRows && maxColumns) sheet.getRange(1, 1, maxRows, maxColumns).clearContent();
  if (snapshot.rows && snapshot.columns) {
    sheet.getRange(1, 1, snapshot.rows, snapshot.columns).setValues(snapshot.values);
  }
  var restored = m0BSnapshotSheet_(snapshot.name);
  m0BAssertCleanup_(restored.checksum === snapshot.checksum &&
    restored.rows === snapshot.rows && restored.columns === snapshot.columns,
    'cleanup did not restore ' + snapshot.name);
}

function m0BRestoreBusinessSet_(baseline, postState, names) {
  names.forEach(function (name) {
    m0BRestoreBusinessSheet_(baseline[name], postState[name].checksum);
  });
}

function m0BPublicData_(response, apiName) {
  m0BAssertProduction_(response && response.ok === true && response.data,
    apiName + ' did not return the required public success envelope');
  return response.data;
}

function m0BValidateWorkflowSchema_() {
  var sheet = m0AssertDisposable_().getSheetByName('27 Semester Workflows');
  m0BAssertEnvironment_(!!sheet, 'workflow sheet is missing');
  var actual = sheet.getRange(HEADER_ROW, 1, 1, M0_GATE_B_WORKFLOW_HEADERS.length)
    .getValues()[0].map(function (value) { return String(value || '').trim(); });
  m0BAssertEnvironment_(JSON.stringify(actual) === JSON.stringify(M0_GATE_B_WORKFLOW_HEADERS),
    'workflow sheet schema does not match approved production persistence order');
  return actual;
}

function m0BValidateHeaders_(sheetName, expected) {
  var sheet = m0AssertDisposable_().getSheetByName(sheetName);
  m0BAssertEnvironment_(!!sheet, sheetName + ' is missing');
  var actual = sheet.getRange(HEADER_ROW, 1, 1, expected.length).getValues()[0]
    .map(function (value) { return String(value || '').trim(); });
  m0BAssertEnvironment_(JSON.stringify(actual) === JSON.stringify(expected),
    sheetName + ' schema does not match approved production persistence order');
  return actual;
}

function m0BWorkflowRowByToken_(token) {
  return readSheetRows_('27 Semester Workflows', HEADER_ROW, 'Workflow Token').filter(function (row) {
    return String(row['Workflow Token']) === String(token);
  })[0] || null;
}

function m0BBackupRowById_(backupId) {
  return readSheetRows_('24 Semester Backups', HEADER_ROW, 'Backup ID').filter(function (row) {
    return String(row['Backup ID']) === String(backupId);
  })[0] || null;
}

function m0BRequireNoPreparedState_() {
  m0BValidateWorkflowSchema_();
  m0BValidateHeaders_('23 Semester Archive', M0_GATE_B_ARCHIVE_HEADERS);
  m0BValidateHeaders_('25 Module Templates', M0_GATE_B_TEMPLATE_HEADERS);
  var semesterId = String(getSetting_('Semester ID') || '');
  var prepared = readSheetRows_('27 Semester Workflows', HEADER_ROW, 'Workflow Token').filter(function (row) {
    return String(row['Semester ID']) === semesterId && String(row['Status']) === 'PREPARED';
  });
  var archives = readSheetRows_('23 Semester Archive', HEADER_ROW, 'Semester ID').filter(function (row) {
    return String(row['Semester ID']) === semesterId;
  });
  var history = readSheetRows_('26 Semester History', HEADER_ROW, 'Semester ID').filter(function (row) {
    return String(row['Semester ID']) === semesterId;
  });
  if (prepared.length || archives.length || history.length) {
    throw m0BError_('ENVIRONMENT BLOCKED',
      'pre-existing prepared workflow, archive, or history for ' + semesterId);
  }
}

function m0BRequireNoTemplateOverlap_(baselineCodes) {
  var activeCodes = readSheetRows_('03 Modules', HEADER_ROW, 'Module code').map(function (row) {
    return String(row['Module code'] || '');
  }).filter(String);
  m0BAssertEnvironment_(!activeCodes.some(function (code) {
    return baselineCodes.indexOf(code) !== -1;
  }), 'pre-existing Module Template would be overwritten by validation preparation');
}

function m0BRequireOwnedPreparation_(data) {
  m0BAssertProduction_(data.reused === false, 'preparation reused an existing artefact');
  m0BAssertProduction_(!!data.workflowToken && data.backup && data.backup.backupId &&
    data.archive && data.archive.semesterId, 'preparation identifiers incomplete');
  var workflow = m0BWorkflowRowByToken_(data.workflowToken);
  var backup = m0BBackupRowById_(data.backup.backupId);
  m0BAssertProduction_(!!workflow, 'production workflow row not found by returned token');
  m0BAssertProduction_(!!backup, 'production backup row not found by returned backup ID');
  m0BAssertProduction_(String(workflow['Backup ID']) === String(data.backup.backupId),
    'workflow-to-backup linkage mismatch');
  m0BAssertProduction_(String(workflow['Archive Semester ID']) === String(data.archive.semesterId),
    'workflow-to-archive linkage mismatch');
  m0BAssertProduction_(String(workflow['Status']) === 'PREPARED',
    'new workflow was not persisted as PREPARED');
  m0BAssertProduction_(workflow['Prepared At'] instanceof Date &&
    !isNaN(workflow['Prepared At'].getTime()), 'prepared timestamp is invalid');
  m0BAssertProduction_(workflow['Expires At'] instanceof Date &&
    workflow['Expires At'].getTime() > Date.now(), 'prepared workflow is already expired');
  return { workflow: workflow, backup: backup };
}

function m0BVerifyPublic_(backupId) {
  var envelope = api_semVerifyBackup(backupId);
  var data = m0BPublicData_(envelope, 'api_semVerifyBackup');
  m0BAssertProduction_(data.ok === true, 'public backup verification did not pass');
  return envelope;
}

function m0BTrashOwnedBackup_(backupId) {
  var row = m0BBackupRowById_(backupId);
  m0BAssertProduction_(!!row && !!row['Drive File ID'], 'owned backup Drive identity unavailable');
  var file = DriveApp.getFileById(String(row['Drive File ID']));
  m0BAssertEnvironment_(!file.isTrashed(), 'owned backup was already trashed');
  file.setTrashed(true);
  m0BAssertHarness_(file.isTrashed() === true, 'owned backup trash operation was not confirmed');
  return { backupId: backupId, fileId: file.getId(), trashed: true };
}

function m0BDisposeOwnedBackup_(backupId) {
  var row = m0BBackupRowById_(backupId);
  if (!row) return { backupId: backupId, status: 'SUCCEEDED', alreadyAbsent: true };
  m0BAssertCleanup_(!!row['Drive File ID'], 'backup cleanup lacks Drive identity');
  var file = null, fileId = String(row['Drive File ID']);
  try { file = DriveApp.getFileById(fileId); } catch (alreadyAbsent) { file = null; }
  if (file) {
    if (!file.isTrashed()) file.setTrashed(true);
    m0BAssertCleanup_(file.isTrashed() === true, 'backup cleanup trash state unknown');
  }
  m0AssertDisposable_().getSheetByName('24 Semester Backups').deleteRow(row.__row);
  m0BAssertCleanup_(!m0BBackupRowById_(backupId), 'backup metadata cleanup failed');
  return { backupId: backupId, fileId: fileId, status: 'SUCCEEDED', alreadyAbsent: !file };
}

function m0BDeleteOwnedRow_(sheetName, keyHeader, keyValue, ownership) {
  var rows = readSheetRows_(sheetName, HEADER_ROW, keyHeader).filter(function (row) {
    return String(row[keyHeader]) === String(keyValue);
  });
  if (!rows.length) return { status: 'SUCCEEDED', alreadyAbsent: true };
  m0BAssertCleanup_(rows.length === 1, 'cleanup ownership is ambiguous in ' + sheetName);
  m0BAssertCleanup_(ownership(rows[0]) === true,
    'cleanup ownership could not be proven in ' + sheetName);
  var sheet = m0AssertDisposable_().getSheetByName(sheetName), rowNumber = rows[0].__row;
  if (sheetName === '23 Semester Archive') {
    sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function (protection) {
      var range = protection.getRange();
      if (range.getRow() <= rowNumber && range.getLastRow() >= rowNumber &&
          String(protection.getDescription() || '').indexOf(String(keyValue)) !== -1) {
        protection.remove();
      }
    });
  }
  sheet.deleteRow(rowNumber);
  m0BAssertCleanup_(!readSheetRows_(sheetName, HEADER_ROW, keyHeader).some(function (row) {
    return String(row[keyHeader]) === String(keyValue);
  }), 'owned row cleanup failed in ' + sheetName);
  return { status: 'SUCCEEDED', key: String(keyValue) };
}

function m0BRestoreSetting_(label, baselineValue, expectedCurrentValue) {
  return m0BRestoreOwnedSetting_(label, {
    originalValue: baselineValue,
    testWrittenValue: expectedCurrentValue
  });
}

function m0BReadUniqueSettingRow_(label) {
  var rows = readSheetRows_('02 Settings', HEADER_ROW, 'Setting').filter(function (row) {
    return String(row['Setting']) === String(label);
  });
  m0BAssertCleanup_(rows.length === 1, 'setting identity is ambiguous: ' + label);
  return rows[0];
}

function m0BReadUniqueSetting_(label) {
  return m0BReadUniqueSettingRow_(label)['Value'];
}

function m0BTypedValueEquals_(left, right) {
  return JSON.stringify(m0CanonicalValue_(left)) === JSON.stringify(m0CanonicalValue_(right));
}

function m0BRestoreOwnedSetting_(label, ownership) {
  m0BAssertCleanup_(ownership &&
    Object.prototype.hasOwnProperty.call(ownership, 'originalValue') &&
    Object.prototype.hasOwnProperty.call(ownership, 'testWrittenValue'),
    'setting ownership evidence is incomplete: ' + label);
  var sheet = m0AssertDisposable_().getSheetByName('02 Settings');
  var row = m0BReadUniqueSettingRow_(label);
  var current = row['Value'];
  if (m0BTypedValueEquals_(current, ownership.originalValue)) return;
  m0BAssertCleanup_(m0BTypedValueEquals_(current, ownership.testWrittenValue),
    'setting changed outside test ownership: ' + label);
  sheet.getRange(row.__row, col_(getColMap_(sheet), 'Value')).setValue(ownership.originalValue);
  m0BAssertCleanup_(m0BTypedValueEquals_(m0BReadUniqueSetting_(label), ownership.originalValue),
    'setting restoration did not persist: ' + label);
}

function m0BAuditRows_() {
  return readSheetRows_('20 Automation Log', HEADER_ROW, 'Action');
}

function m0BAssertAuditActions_(rows, allowed, forbidden) {
  rows.forEach(function (row) {
    m0BAssertProduction_(allowed.indexOf(String(row['Action'])) !== -1,
      'unexpected Automation Log action: ' + row['Action']);
  });
  forbidden.forEach(function (action) {
    m0BAssertProduction_(!rows.some(function (row) {
      return String(row['Action']) === action;
    }), 'forbidden Automation Log action present: ' + action);
  });
}

function m0BRestoreAuditBaseline_(baseline, expectedRows) {
  var sheet = m0AssertDisposable_().getSheetByName('20 Automation Log');
  var current = m0BAuditRows_();
  var currentValues = current.map(function (row) {
    return [row['Timestamp'], row['Action'], row['Item'], row['Result'], row['Detail']];
  });
  m0BAssertCleanup_(m0BChecksum_(currentValues) === m0BChecksum_(expectedRows.map(function (row) {
    return [row['Timestamp'], row['Action'], row['Item'], row['Result'], row['Detail']];
  })), 'Automation Log changed after evidence capture');
  var baselineData = baseline.values.slice(HEADER_ROW);
  var prefixMatches = currentValues.length >= baselineData.length &&
    m0BChecksum_(currentValues.slice(0, baselineData.length)) === m0BChecksum_(baselineData);
  if (prefixMatches) {
    for (var suffixRow = sheet.getLastRow(); suffixRow > HEADER_ROW + baselineData.length; suffixRow--) {
      sheet.deleteRow(suffixRow);
    }
  } else {
    for (var ownedRow = sheet.getLastRow(); ownedRow > HEADER_ROW; ownedRow--) sheet.deleteRow(ownedRow);
    if (baselineData.length) {
      sheet.getRange(HEADER_ROW + 1, 1, baselineData.length, baselineData[0].length)
        .setValues(baselineData);
    }
  }
  m0BAssertCleanup_(m0BSnapshotSheet_('20 Automation Log').checksum === baseline.checksum,
    'Automation Log baseline restoration failed');
}

function m0BCleanupTemplates_(semesterId, baselineCodes) {
  var rows = readSheetRows_('25 Module Templates', HEADER_ROW, 'Module Code');
  var owned = rows.filter(function (row) {
    return baselineCodes.indexOf(String(row['Module Code'])) === -1 &&
      String(row['Source Semester ID']) === String(semesterId);
  }).sort(function (a, b) { return b.__row - a.__row; });
  var sheet = m0AssertDisposable_().getSheetByName('25 Module Templates');
  owned.forEach(function (row) { sheet.deleteRow(row.__row); });
  m0BAssertCleanup_(!readSheetRows_('25 Module Templates', HEADER_ROW, 'Module Code').some(function (row) {
    return baselineCodes.indexOf(String(row['Module Code'])) === -1 &&
      String(row['Source Semester ID']) === String(semesterId);
  }), 'test-owned Module Templates cleanup failed');
}

function m0BCleanupPreparation_(owned, baselineTemplateCodes) {
  var token = owned.workflowToken, backupId = owned.backupId;
  var archiveId = owned.archiveSemesterId, semesterId = owned.semesterId;
  m0BDeleteOwnedRow_('27 Semester Workflows', 'Workflow Token', token, function (row) {
    return String(row['Backup ID']) === String(backupId) &&
      String(row['Archive Semester ID']) === String(archiveId) &&
      String(row['Semester ID']) === String(semesterId);
  });
  m0BDisposeOwnedBackup_(backupId);
  m0BDeleteOwnedRow_('23 Semester Archive', 'Semester ID', archiveId, function (row) {
    return String(row['Semester ID']) === String(semesterId) && row['Immutable'] === true;
  });
  m0BDeleteOwnedRow_('26 Semester History', 'Semester ID', semesterId, function (row) {
    return String(row['Semester ID']) === String(semesterId);
  });
  m0BCleanupTemplates_(semesterId, baselineTemplateCodes);
}

function m0BAssertBoundaryCleared_(returnedCleared, auditAllowed, auditRequired, nonTargetBaseline) {
  var expected = M0_GATE_B_TARGETS.map(function (target) { return target.name; }).sort();
  var actual = (returnedCleared || []).slice().sort();
  m0BAssertProduction_(JSON.stringify(actual) === JSON.stringify(expected),
    'returned cleared-sheet set differs from frozen ten-sheet boundary');
  M0_GATE_B_TARGETS.slice(0, 9).forEach(function (target) {
    m0BAssertCleared_(target.name, target.header);
  });
  var logRows = readSheetRows_('20 Automation Log', HEADER_ROW, 'Action');
  m0BAssertAuditActions_(logRows, auditAllowed, []);
  auditRequired.forEach(function (action) {
    m0BAssertProduction_(logRows.some(function (row) {
      return String(row['Action']) === action;
    }), 'required Automation Log action missing: ' + action);
  });
  m0BAssertSame_(nonTargetBaseline, Object.keys(nonTargetBaseline));
  return logRows;
}

function m0BClassify_(error) {
  if (error && error.m0Classification) return error.m0Classification;
  return 'VALIDATION HARNESS FAILURE';
}

function m0BResult_(testId, started, status, evidence, error, cleanup) {
  var result = {
    testId: testId, status: status, durationMs: new Date() - started,
    evidence: evidence || {}, cleanup: cleanup || { status: 'UNKNOWN' }
  };
  if (error) result.error = String(error && error.message || error);
  console.log(JSON.stringify(result));
  return result;
}

function m0BGateBSafety_() {
  var safety = m0GateASafetyCheck();
  m0BAssertEnvironment_(safety.calendarAccess === 'DISABLED' && safety.productionAccess === 'NONE',
    'disposable safety controls failed');
  return safety;
}

function m0RunGateBTest11() {
  var started = new Date(), testId = 'M0-LV-11', evidence = {}, cleanup = { status: 'UNKNOWN' };
  var baseline, preparation, owned, auditAfter, statusOwnership;
  try {
    evidence.safety = m0BGateBSafety_();
    m0BRequireNoPreparedState_();
    baseline = m0BSnapshotSet_(m0BTargetNames_().concat(['20 Automation Log']));
    var baselineStatus = m0BReadUniqueSetting_('Semester status');
    var baselineTemplateCodes = readSheetRows_('25 Module Templates', HEADER_ROW, 'Module Code')
      .map(function (row) { return String(row['Module Code']); });
    m0BRequireNoTemplateOverlap_(baselineTemplateCodes);
    preparation = api_semStartDeleteWorkflow({ validationTest: testId, note: 'Deterministic disposable validation reflection' });
    var prepared = m0BPublicData_(preparation, 'api_semStartDeleteWorkflow');
    var ownership = m0BRequireOwnedPreparation_(prepared);
    statusOwnership = {
      originalValue: baselineStatus,
      testWrittenValue: m0BReadUniqueSetting_('Semester status')
    };
    owned = {
      workflowToken: prepared.workflowToken, backupId: prepared.backup.backupId,
      archiveSemesterId: prepared.archive.semesterId, semesterId: prepared.semesterId,
      driveFileId: ownership.backup['Drive File ID']
    };
    evidence.preparation = preparation;
    evidence.identifiers = owned;
    evidence.preCorruptionVerification = m0BVerifyPublic_(owned.backupId);
    evidence.corruption = m0BTrashOwnedBackup_(owned.backupId);
    evidence.confirmation = api_semConfirmDelete(
      prepared.workflowToken, 'DELETE SEMESTER', { deleteCalendarEvents: false });
    m0BAssertProduction_(evidence.confirmation && evidence.confirmation.ok === false,
      'production accepted a workflow whose backup was trashed');
    m0BAssertProduction_(evidence.confirmation.error ===
      'Associated backup failed verification: Backup Drive file is trashed and unavailable.',
      'production rejection reason differed from frozen contract');
    m0BAssertSame_(baseline, m0BTargetNames_().slice(0, 9));
    var workflowAfter = m0BWorkflowRowByToken_(prepared.workflowToken);
    m0BAssertProduction_(workflowAfter && String(workflowAfter['Status']) === 'PREPARED',
      'rejected workflow did not remain exactly PREPARED');
    m0BAssertProduction_(!workflowAfter['Consumed At'] &&
      String(workflowAfter['Last Stage']) === 'PREPARED' && !workflowAfter['Last Error'],
      'rejected workflow persisted an invalid consumed/stage/error state');
    auditAfter = m0BAuditRows_();
    var auditDelta = auditAfter.slice(Math.max(0, baseline['20 Automation Log'].rows - HEADER_ROW));
    m0BAssertAuditActions_(auditDelta, [
      'Semester backup created', 'Semester archived', 'Module templates saved',
      'Delete Semester workflow prepared', 'API error'
    ], ['Semester reset (active data cleared)', 'Semester deleted (confirmed)', 'Emergency reset completed']);
    m0BAssertProduction_(auditDelta.some(function (row) { return row['Action'] === 'API error'; }),
      'expected rejection API-error audit entry missing');
    evidence.auditDelta = auditDelta;
    evidence.after = m0BCurrentState_(m0BTargetNames_());
    m0BCleanupPreparation_(owned, baselineTemplateCodes);
    m0BRestoreOwnedSetting_('Semester status', statusOwnership);
    m0BRestoreAuditBaseline_(baseline['20 Automation Log'], auditAfter);
    m0BAssertSame_(baseline, m0BTargetNames_(), 'CLEANUP INCOMPLETE');
    cleanup = { status: 'SUCCEEDED', corruptedFileRetainedInTrash: true };
    return m0BResult_(testId, started, 'PASS', evidence, null, cleanup);
  } catch (error) {
    try {
      if (baseline && owned) {
        m0BCleanupPreparation_(owned, baselineTemplateCodes || []);
        m0BRestoreOwnedSetting_('Semester status', statusOwnership);
        if (auditAfter) m0BRestoreAuditBaseline_(baseline['20 Automation Log'], auditAfter);
        else m0BAssertCleanup_(m0BSnapshotSheet_('20 Automation Log').checksum ===
          baseline['20 Automation Log'].checksum, 'unsafe Automation Log cleanup state');
        cleanup = { status: 'SUCCEEDED' };
      }
    } catch (cleanupError) {
      cleanup = { status: 'FAILED', error: String(cleanupError.message || cleanupError) };
      error = m0BError_('CLEANUP INCOMPLETE', cleanup.error);
    }
    return m0BResult_(testId, started, m0BClassify_(error), evidence, error, cleanup);
  }
}

function m0RunGateBTest12() {
  var started = new Date(), testId = 'M0-LV-12', evidence = {}, cleanup = { status: 'UNKNOWN' };
  var baseline, prepared, owned, postState, auditAfter, statusOwnership, manualOwnership;
  try {
    evidence.safety = m0BGateBSafety_();
    m0BRequireNoPreparedState_();
    baseline = m0BSnapshotSet_(m0BTargetNames_());
    var auditBaseline = m0BSnapshotSheet_('20 Automation Log');
    var nonTargetBaseline = m0BNonTargetBaseline_(M0_GATE_B_INFRA.concat(['20 Automation Log']));
    var baselineStatus = m0BReadUniqueSetting_('Semester status');
    var baselineManual = m0BReadUniqueSetting_('Semester manually finished (TRUE/FALSE)');
    var baselineTemplateCodes = readSheetRows_('25 Module Templates', HEADER_ROW, 'Module Code')
      .map(function (row) { return String(row['Module Code']); });
    m0BRequireNoTemplateOverlap_(baselineTemplateCodes);
    var preparation = api_semStartDeleteWorkflow({ validationTest: testId, note: 'Deterministic disposable validation reflection' });
    prepared = m0BPublicData_(preparation, 'api_semStartDeleteWorkflow');
    var ownership = m0BRequireOwnedPreparation_(prepared);
    statusOwnership = {
      originalValue: baselineStatus,
      testWrittenValue: m0BReadUniqueSetting_('Semester status')
    };
    manualOwnership = {
      originalValue: baselineManual,
      testWrittenValue: m0BReadUniqueSetting_('Semester manually finished (TRUE/FALSE)')
    };
    owned = {
      workflowToken: prepared.workflowToken, backupId: prepared.backup.backupId,
      archiveSemesterId: prepared.archive.semesterId, semesterId: prepared.semesterId,
      driveFileId: ownership.backup['Drive File ID']
    };
    evidence.preparation = preparation;
    evidence.verification = m0BVerifyPublic_(owned.backupId);
    evidence.identifiers = owned;
    evidence.confirmation = api_semConfirmDelete(
      prepared.workflowToken, 'DELETE SEMESTER', { deleteCalendarEvents: false });
    var confirmed = m0BPublicData_(evidence.confirmation, 'api_semConfirmDelete');
    m0BAssertProduction_(confirmed.workflowTokenConsumed === true, 'workflow token was not consumed');
    m0BAssertProduction_(confirmed.deletedEvents === 0 &&
      (!confirmed.calendarErrors || confirmed.calendarErrors.length === 0),
      'Calendar deletion was not disabled');
    auditAfter = m0BAssertBoundaryCleared_(confirmed.clearedSheets, [
      'Semester reset (active data cleared)', 'Semester deleted (confirmed)'
    ], ['Semester reset (active data cleared)', 'Semester deleted (confirmed)'], nonTargetBaseline);
    var persisted = m0BWorkflowRowByToken_(owned.workflowToken);
    m0BAssertProduction_(persisted && String(persisted['Status']) === 'CONSUMED',
      'persisted workflow status is not CONSUMED');
    m0BAssertProduction_(persisted['Consumed At'] instanceof Date &&
      !isNaN(persisted['Consumed At'].getTime()), 'persisted consumed timestamp is invalid');
    m0BAssertProduction_(String(persisted['Last Stage']) === 'RESET_COMPLETE' &&
      String(persisted['Backup ID']) === String(owned.backupId) &&
      String(persisted['Archive Semester ID']) === String(owned.archiveSemesterId),
      'persisted workflow stage or identity linkage is invalid');
    evidence.after = m0BCurrentState_(m0BTargetNames_());
    postState = m0BSnapshotSet_(m0BTargetNames_());
    m0BRestoreBusinessSet_(baseline, postState, m0BTargetNames_().slice(0, 9));
    m0BCleanupPreparation_(owned, baselineTemplateCodes);
    m0BRestoreOwnedSetting_('Semester status', statusOwnership);
    m0BRestoreOwnedSetting_('Semester manually finished (TRUE/FALSE)', manualOwnership);
    m0BRestoreAuditBaseline_(auditBaseline, auditAfter);
    m0BAssertSame_(baseline, m0BTargetNames_(), 'CLEANUP INCOMPLETE');
    cleanup = { status: 'SUCCEEDED', backupDisposed: true };
    return m0BResult_(testId, started, 'PASS', evidence, null, cleanup);
  } catch (error) {
    try {
      if (baseline) {
        if (postState) m0BRestoreBusinessSet_(baseline, postState, m0BTargetNames_().slice(0, 9));
        else m0BAssertSame_(baseline, m0BTargetNames_().slice(0, 9), 'CLEANUP INCOMPLETE');
        if (owned) {
          m0BCleanupPreparation_(owned, baselineTemplateCodes || []);
        }
        m0BRestoreOwnedSetting_('Semester status', statusOwnership);
        m0BRestoreOwnedSetting_('Semester manually finished (TRUE/FALSE)', manualOwnership);
        if (auditAfter) m0BRestoreAuditBaseline_(auditBaseline, auditAfter);
        else m0BAssertCleanup_(m0BSnapshotSheet_('20 Automation Log').checksum ===
          auditBaseline.checksum, 'unsafe Automation Log cleanup state');
        cleanup = { status: 'SUCCEEDED' };
      }
    } catch (cleanupError) {
      cleanup = { status: 'FAILED', error: String(cleanupError.message || cleanupError) };
      error = m0BError_('CLEANUP INCOMPLETE', cleanup.error);
    }
    return m0BResult_(testId, started, m0BClassify_(error), evidence, error, cleanup);
  }
}

function m0RunGateBTest13() {
  var started = new Date(), testId = 'M0-LV-13', evidence = {}, cleanup = { status: 'UNKNOWN' };
  var baseline, modules, originalId, renamed = false, auditAfter;
  try {
    evidence.safety = m0BGateBSafety_();
    baseline = m0BSnapshotSet_(m0BTargetNames_().concat(['03 Modules', '24 Semester Backups']));
    modules = m0AssertDisposable_().getSheetByName('03 Modules');
    originalId = modules.getSheetId();
    m0BAssertEnvironment_(!m0AssertDisposable_().getSheetByName('03 Modules — M0-LV-13 TEMP'),
      'temporary M0-LV-13 sheet already exists');
    modules.setName('03 Modules — M0-LV-13 TEMP');
    renamed = true;
    evidence.fault = { sheetId: originalId, temporaryName: '03 Modules — M0-LV-13 TEMP' };
    var driveBefore = m0FolderFileCount_();
    evidence.response = api_semResetActiveSemester(
      'EMERGENCY RESET', { deleteCalendarEvents: false });
    m0BAssertProduction_(evidence.response && evidence.response.ok === false,
      'emergency reset accepted missing required sheet');
    m0BAssertProduction_(evidence.response.error ===
      'Backup was not created because required sheets are missing or empty: 03 Modules',
      'emergency rejection reason differed from frozen contract');
    m0BAssertSame_(baseline, m0BTargetNames_().slice(0, 9));
    m0BAssertProduction_(m0FolderFileCount_() === driveBefore, 'failed backup left a Drive artefact');
    m0BAssertProduction_(m0BSnapshotSheet_('24 Semester Backups').checksum ===
      baseline['24 Semester Backups'].checksum, 'failed backup left metadata');
    auditAfter = m0BAuditRows_();
    var auditDelta = auditAfter.slice(Math.max(0, baseline['20 Automation Log'].rows - HEADER_ROW));
    m0BAssertAuditActions_(auditDelta, ['API error'], [
      'Semester backup created', 'Semester reset (active data cleared)',
      'Semester deleted (confirmed)', 'Emergency reset completed'
    ]);
    m0BAssertProduction_(auditDelta.length >= 1, 'expected API-error audit delta missing');
    evidence.auditDelta = auditDelta;
    modules.setName('03 Modules');
    renamed = false;
    m0BAssertSame_(baseline, ['03 Modules'].concat(m0BTargetNames_().slice(0, 9)));
    m0BRestoreAuditBaseline_(baseline['20 Automation Log'], auditAfter);
    m0BAssertSame_(baseline, m0BTargetNames_().concat(['03 Modules', '24 Semester Backups']), 'CLEANUP INCOMPLETE');
    cleanup = { status: 'SUCCEEDED', sheetRestored: true };
    return m0BResult_(testId, started, 'PASS', evidence, null, cleanup);
  } catch (error) {
    try {
      var ss = m0AssertDisposable_();
      var temp = ss.getSheetByName('03 Modules — M0-LV-13 TEMP');
      var original = ss.getSheetByName('03 Modules');
      if (renamed) {
        m0BAssertCleanup_(!!temp && !original && temp.getSheetId() === originalId,
          'cleanup identity mismatch for 03 Modules');
        temp.setName('03 Modules');
      }
      if (baseline) {
        if (auditAfter) m0BRestoreAuditBaseline_(baseline['20 Automation Log'], auditAfter);
        m0BAssertSame_(baseline, ['03 Modules'].concat(m0BTargetNames_().slice(0, 9)), 'CLEANUP INCOMPLETE');
      }
      cleanup = { status: 'SUCCEEDED', sheetRestored: true };
    } catch (cleanupError) {
      cleanup = { status: 'FAILED', error: String(cleanupError.message || cleanupError) };
      error = m0BError_('CLEANUP INCOMPLETE', cleanup.error);
    }
    return m0BResult_(testId, started, m0BClassify_(error), evidence, error, cleanup);
  }
}

function m0RunGateBTest14() {
  var started = new Date(), testId = 'M0-LV-14', evidence = {}, cleanup = { status: 'UNKNOWN' };
  var baseline, emergencyBackupId, postState, auditAfter;
  try {
    evidence.safety = m0BGateBSafety_();
    baseline = m0BSnapshotSet_(m0BTargetNames_());
    var auditBaseline = m0BSnapshotSheet_('20 Automation Log');
    var nonTargetBaseline = m0BNonTargetBaseline_([
      '02 Settings', '20 Automation Log', '24 Semester Backups'
    ]);
    var baselineManual = getSetting_('Semester manually finished (TRUE/FALSE)');
    evidence.response = api_semResetActiveSemester(
      'EMERGENCY RESET', { deleteCalendarEvents: false });
    var reset = m0BPublicData_(evidence.response, 'api_semResetActiveSemester');
    emergencyBackupId = reset.emergencyBackupId;
    m0BAssertProduction_(!!emergencyBackupId, 'emergency reset returned no backup ID');
    m0BAssertProduction_(reset.deletedEvents === 0 &&
      (!reset.calendarErrors || reset.calendarErrors.length === 0),
      'Calendar deletion was not disabled');
    auditAfter = m0BAssertBoundaryCleared_(reset.clearedSheets, [
      'Semester reset (active data cleared)', 'Emergency reset completed'
    ], ['Semester reset (active data cleared)', 'Emergency reset completed'], nonTargetBaseline);
    evidence.verification = m0BVerifyPublic_(emergencyBackupId);
    var backup = m0BBackupRowById_(emergencyBackupId);
    m0BAssertProduction_(!!backup && !!backup['Drive File ID'],
      'emergency backup metadata linkage unavailable');
    evidence.identifiers = {
      emergencyBackupId: emergencyBackupId, driveFileId: backup['Drive File ID']
    };
    evidence.after = m0BCurrentState_(m0BTargetNames_());
    postState = m0BSnapshotSet_(m0BTargetNames_());
    m0BRestoreBusinessSet_(baseline, postState, m0BTargetNames_().slice(0, 9));
    m0BRestoreSetting_('Semester manually finished (TRUE/FALSE)', baselineManual, false);
    m0BRestoreAuditBaseline_(auditBaseline, auditAfter);
    m0BAssertSame_(baseline, m0BTargetNames_(), 'CLEANUP INCOMPLETE');
    m0BAssertProduction_(m0BPublicData_(api_semVerifyBackup(emergencyBackupId),
      'api_semVerifyBackup').ok === true, 'retained emergency backup became invalid');
    cleanup = { status: 'SUCCEEDED', emergencyBackupRetained: true };
    return m0BResult_(testId, started, 'PASS', evidence, null, cleanup);
  } catch (error) {
    try {
      if (baseline) {
        if (postState) m0BRestoreBusinessSet_(baseline, postState, m0BTargetNames_().slice(0, 9));
        else m0BAssertSame_(baseline, m0BTargetNames_().slice(0, 9), 'CLEANUP INCOMPLETE');
        m0BRestoreSetting_('Semester manually finished (TRUE/FALSE)', baselineManual, false);
        if (auditAfter) m0BRestoreAuditBaseline_(auditBaseline, auditAfter);
        else m0BAssertCleanup_(m0BSnapshotSheet_('20 Automation Log').checksum ===
          auditBaseline.checksum, 'unsafe Automation Log cleanup state');
        m0BAssertSame_(baseline, m0BTargetNames_(), 'CLEANUP INCOMPLETE');
        cleanup = { status: 'SUCCEEDED', emergencyBackupRetained: !!emergencyBackupId };
      }
    } catch (cleanupError) {
      cleanup = { status: 'FAILED', error: String(cleanupError.message || cleanupError) };
      error = m0BError_('CLEANUP INCOMPLETE', cleanup.error);
    }
    return m0BResult_(testId, started, m0BClassify_(error), evidence, error, cleanup);
  }
}

// ---------------------------------------------------------------------------
// M0-LV-15 RECOVERY DRILL — VALIDATION ONLY; NOT A PRODUCTION RESTORE API.
// Recovers only 03 Modules from the uniquely newest retained verified backup.
// ---------------------------------------------------------------------------
function m0R15Error_(classification, message) {
  var error = new Error(message);
  error.m0R15Classification = classification;
  return error;
}

function m0R15Assert_(condition, classification, message) {
  if (!condition) throw m0R15Error_(classification, message);
}

function m0R15Classify_(error) {
  return error && error.m0R15Classification ? error.m0R15Classification : 'FAIL';
}

function m0R15NewestBackup_() {
  var rows = readSheetRows_('24 Semester Backups', HEADER_ROW, 'Backup ID');
  m0R15Assert_(rows.length > 0, 'BLOCKED', 'no indexed backup is available');
  var candidates = rows.map(function (row) {
    var timestamp = row['Timestamp'] instanceof Date ? row['Timestamp'] : new Date(row['Timestamp']);
    m0R15Assert_(!isNaN(timestamp.getTime()), 'BLOCKED',
      'indexed backup timestamp is invalid: ' + row['Backup ID']);
    return { row: row, epoch: timestamp.getTime() };
  }).sort(function (left, right) { return right.epoch - left.epoch; });
  m0R15Assert_(candidates.length === 1 || candidates[0].epoch > candidates[1].epoch,
    'BLOCKED', 'newest retained backup identity is ambiguous');
  var selected = candidates[0].row;
  m0R15Assert_(!!selected['Backup ID'] && !!selected['Drive File ID'], 'BLOCKED',
    'newest retained backup has incomplete identity');
  m0R15Assert_(!rows.some(function (row) {
    return row.__row !== selected.__row &&
      (String(row['Backup ID']) === String(selected['Backup ID']) ||
       String(row['Drive File ID']) === String(selected['Drive File ID']));
  }), 'BLOCKED', 'newest retained backup identity is duplicated');
  return selected;
}

function m0R15BindPayload_(row) {
  var backupId = String(row['Backup ID']);
  var verificationEnvelope = api_semVerifyBackup(backupId);
  m0R15Assert_(verificationEnvelope && verificationEnvelope.ok === true &&
    verificationEnvelope.data, 'BLOCKED',
    'api_semVerifyBackup did not return the required public success envelope');
  var verification = verificationEnvelope.data;
  m0R15Assert_(verification.ok === true && verification.fileFound === true &&
    verification.jsonValid === true && verification.checksumValid === true &&
    Array.isArray(verification.missingSheets) && verification.missingSheets.length === 0,
    'BLOCKED', 'retained emergency backup did not pass public verification');

  var file = DriveApp.getFileById(String(row['Drive File ID']));
  m0R15Assert_(file.getId() === String(row['Drive File ID']) && file.isTrashed() === false,
    'BLOCKED', 'bound backup Drive identity is unavailable');
  var payload = JSON.parse(file.getBlob().getDataAsString());
  var expectedChecksum = semBackupChecksumFromNotes_(row['Notes']);
  var actualChecksum = semBackupChecksum_(payload);
  m0R15Assert_(!!expectedChecksum && payload.contentChecksum === expectedChecksum &&
    actualChecksum === expectedChecksum, 'BLOCKED',
    'bound payload fingerprint did not match independent metadata');
  m0R15Assert_(payload.backupId === backupId &&
    String(payload.semesterId || '') === String(row['Semester ID'] || '') &&
    String(payload.plannerVersion || '') === String(row['Planner Version'] || '') &&
    String(payload.spreadsheetId || '') === String(m0AssertDisposable_().getId()) &&
    payload.formatVersion === SEM_BACKUP_FORMAT_VERSION,
    'BLOCKED', 'bound payload identity did not match verified metadata');
  m0R15Assert_(semBackupSheetSetMatches_(Object.keys(payload.sheets || {})), 'BLOCKED',
    'bound payload sheet set is invalid');
  m0R15Assert_(String(row['Sheets Included'] || '').split(',').map(function (name) {
    return name.trim();
  }).filter(Boolean).length === SEM_BACKUP_SHEETS.length &&
    semBackupSheetSetMatches_(String(row['Sheets Included'] || '').split(',').map(function (name) {
      return name.trim();
    }).filter(Boolean)), 'BLOCKED', 'indexed sheet set is invalid');
  return {
    rowNumber: row.__row,
    backupId: backupId,
    driveFileId: file.getId(),
    timestamp: row['Timestamp'] instanceof Date ? row['Timestamp'].toISOString() : String(row['Timestamp']),
    checksum: expectedChecksum,
    verification: verificationEnvelope,
    file: file,
    payload: payload
  };
}

function m0R15Reconstruct_(payloadValues, baselineValues) {
  m0R15Assert_(Array.isArray(payloadValues) && payloadValues.length === baselineValues.length,
    'BLOCKED', '03 Modules backup row count differs from approved baseline');
  return payloadValues.map(function (row, rowIndex) {
    m0R15Assert_(Array.isArray(row) && row.length === baselineValues[rowIndex].length,
      'BLOCKED', '03 Modules backup column count differs from approved baseline');
    return row.map(function (value, columnIndex) {
      var baseline = baselineValues[rowIndex][columnIndex];
      if (baseline instanceof Date) {
        m0R15Assert_(typeof value === 'string' && value === baseline.toISOString(),
          'FAIL', 'serialized Date differs at R' + (rowIndex + 1) + 'C' + (columnIndex + 1));
        return new Date(value);
      }
      m0R15Assert_(value === '' || typeof value === 'string' || typeof value === 'boolean' ||
        (typeof value === 'number' && isFinite(value)), 'FAIL',
        'unsupported backup value at R' + (rowIndex + 1) + 'C' + (columnIndex + 1));
      return value;
    });
  });
}

function m0R15MismatchList_(expected, actual) {
  var mismatches = [];
  for (var row = 0; row < expected.length; row++) {
    for (var column = 0; column < expected[row].length; column++) {
      var expectedValue = m0CanonicalValue_(expected[row][column]);
      var actualValue = m0CanonicalValue_(actual[row][column]);
      if (JSON.stringify(expectedValue) !== JSON.stringify(actualValue)) {
        mismatches.push({
          row: row + 1, column: column + 1,
          expected: expectedValue, actual: actualValue
        });
      }
    }
  }
  return mismatches;
}

function m0R15DriveInventory_() {
  var parent = DriveApp.getFolderById(M0_DISPOSABLE_DRIVE_FOLDER_ID);
  var folders = parent.getFoldersByName('Engineering Planner OS Backups — M0 Disposable');
  m0R15Assert_(folders.hasNext(), 'BLOCKED', 'disposable backup folder is missing');
  var folder = folders.next();
  m0R15Assert_(!folders.hasNext(), 'BLOCKED', 'disposable backup folder identity is ambiguous');
  var files = folder.getFiles(), inventory = [];
  while (files.hasNext()) {
    var file = files.next();
    inventory.push({ id: file.getId(), trashed: file.isTrashed() });
  }
  return inventory.sort(function (left, right) { return left.id.localeCompare(right.id); });
}

function m0RunRecoveryDrill15() {
  var started = new Date(), evidence = {}, cleanup = { status: 'UNKNOWN' };
  var targetBaseline = null, target = null, corruptedChecksum = '', bound = null;
  try {
    evidence.safety = m0BGateBSafety_();
    var ss = m0AssertDisposable_();
    target = ss.getSheetByName('03 Modules');
    m0R15Assert_(!!target && ss.getSheets().filter(function (sheet) {
      return sheet.getName() === '03 Modules';
    }).length === 1, 'BLOCKED', '03 Modules target identity is not unique');

    targetBaseline = m0BSnapshotSheet_('03 Modules');
    m0R15Assert_(targetBaseline.rows === 5 && targetBaseline.columns === 12,
      'BLOCKED', '03 Modules approved dimensions are not 5x12');
    m0R15Assert_(targetBaseline.checksum ===
      '73180c7dfd6a6863ffc44b4e0bf3b422db9eee7a4d3bdbcf92c84d7f360b404a',
      'BLOCKED', '03 Modules does not match the approved fixture checksum');

    var nonTargetNames = ss.getSheets().map(function (sheet) { return sheet.getName(); })
      .filter(function (name) { return name !== '03 Modules'; });
    var nonTargetBaseline = m0BSnapshotSet_(nonTargetNames);
    var driveBaseline = m0R15DriveInventory_();
    var backupRow = m0R15NewestBackup_();
    bound = m0R15BindPayload_(backupRow);
    var dump = bound.payload.sheets['03 Modules'];
    m0R15Assert_(dump && dump.sheetName === '03 Modules', 'BLOCKED',
      'bound payload does not contain the sole approved target');
    var recoveredValues = m0R15Reconstruct_(dump.values, targetBaseline.values);
    m0R15Assert_(m0BChecksum_(recoveredValues) === targetBaseline.checksum, 'FAIL',
      'backup-derived 03 Modules values do not match approved baseline');

    evidence.backup = {
      backupId: bound.backupId, driveFileId: bound.driveFileId,
      timestamp: bound.timestamp, checksum: bound.checksum
    };
    evidence.verification = bound.verification;
    evidence.preCorruption = {
      rows: targetBaseline.rows, columns: targetBaseline.columns,
      checksum: targetBaseline.checksum
    };

    target.getRange(HEADER_ROW + 1, 1, 1, targetBaseline.columns).clearContent();
    SpreadsheetApp.flush();
    m0R15Assert_(target.getSheetId() === targetBaseline.sheetId, 'FAIL',
      '03 Modules sheet identity changed during corruption');
    corruptedChecksum = m0BSnapshotSheet_('03 Modules').checksum;
    m0R15Assert_(corruptedChecksum !== targetBaseline.checksum, 'FAIL',
      'controlled corruption did not change the target checksum');
    evidence.postCorruption = { checksum: corruptedChecksum };
    m0BAssertSame_(nonTargetBaseline, nonTargetNames);

    var reboundRow = m0BBackupRowById_(bound.backupId);
    m0R15Assert_(reboundRow && reboundRow.__row === bound.rowNumber &&
      String(reboundRow['Drive File ID']) === bound.driveFileId, 'FAIL',
      'verified backup metadata changed before recovery write');
    var rebound = m0R15BindPayload_(reboundRow);
    m0R15Assert_(rebound.driveFileId === bound.driveFileId &&
      rebound.checksum === bound.checksum, 'FAIL',
      'verified backup identity or fingerprint was swapped before write');
    recoveredValues = m0R15Reconstruct_(rebound.payload.sheets['03 Modules'].values,
      targetBaseline.values);
    m0R15Assert_(m0BSnapshotSheet_('03 Modules').checksum === corruptedChecksum, 'FAIL',
      '03 Modules changed outside test ownership before recovery write');

    var maxRows = Math.max(target.getLastRow(), targetBaseline.rows);
    var maxColumns = Math.max(target.getLastColumn(), targetBaseline.columns);
    target.getRange(1, 1, maxRows, maxColumns).clearContent();
    target.getRange(1, 1, recoveredValues.length, recoveredValues[0].length)
      .setValues(recoveredValues);
    SpreadsheetApp.flush();

    var restored = m0BSnapshotSheet_('03 Modules');
    var mismatches = m0R15MismatchList_(targetBaseline.values, restored.values);
    evidence.postRecovery = {
      rows: restored.rows, columns: restored.columns, checksum: restored.checksum,
      mismatchCount: mismatches.length, mismatches: mismatches
    };
    m0R15Assert_(restored.sheetId === targetBaseline.sheetId &&
      restored.rows === targetBaseline.rows && restored.columns === targetBaseline.columns &&
      restored.checksum === targetBaseline.checksum && mismatches.length === 0,
      'FAIL', 'backup-derived recovery did not exactly reproduce 03 Modules');
    m0BAssertSame_(nonTargetBaseline, nonTargetNames);
    m0R15Assert_(JSON.stringify(m0R15DriveInventory_()) === JSON.stringify(driveBaseline),
      'FAIL', 'Drive backup inventory changed during recovery drill');
    var finalVerification = m0R15BindPayload_(m0BBackupRowById_(bound.backupId));
    m0R15Assert_(finalVerification.checksum === bound.checksum, 'FAIL',
      'retained backup changed after recovery');
    cleanup = { status: 'SUCCEEDED', environmentReusable: true, backupRetained: true };
    evidence.totalRuntimeMs = new Date() - started;
    return m0BResult_('M0-LV-15', started, 'PASS', evidence, null, cleanup);
  } catch (error) {
    try {
      if (targetBaseline && target && corruptedChecksum &&
          m0BSnapshotSheet_('03 Modules').checksum === corruptedChecksum && bound) {
        var cleanupBound = m0R15BindPayload_(m0BBackupRowById_(bound.backupId));
        var cleanupValues = m0R15Reconstruct_(cleanupBound.payload.sheets['03 Modules'].values,
          targetBaseline.values);
        target.getRange(1, 1, Math.max(target.getLastRow(), targetBaseline.rows),
          Math.max(target.getLastColumn(), targetBaseline.columns)).clearContent();
        target.getRange(1, 1, cleanupValues.length, cleanupValues[0].length).setValues(cleanupValues);
        SpreadsheetApp.flush();
      }
      if (targetBaseline) {
        var cleanupSnapshot = m0BSnapshotSheet_('03 Modules');
        m0R15Assert_(cleanupSnapshot.sheetId === targetBaseline.sheetId &&
          cleanupSnapshot.rows === targetBaseline.rows &&
          cleanupSnapshot.columns === targetBaseline.columns &&
          cleanupSnapshot.checksum === targetBaseline.checksum,
          'CLEANUP INCOMPLETE', '03 Modules reusable baseline could not be proven');
      }
      cleanup = { status: 'SUCCEEDED', environmentReusable: true, backupRetained: !!bound };
    } catch (cleanupError) {
      cleanup = { status: 'FAILED', environmentReusable: false,
        error: String(cleanupError.message || cleanupError) };
      error = m0R15Error_('CLEANUP INCOMPLETE', cleanup.error);
    }
    var classification = m0R15Classify_(error);
    m0R15Assert_(['FAIL', 'CLEANUP INCOMPLETE', 'BLOCKED'].indexOf(classification) !== -1,
      'FAIL', 'invalid M0-LV-15 classification');
    return m0BResult_('M0-LV-15', started, classification, evidence, error, cleanup);
  }
}

// Residual cleanup for the single failed M0-LV-11 execution at 2026-07-18
// 09:16 SAST. Validation-only, stable-identity scoped, and safe to rerun.
function m0RecoverGateBTest11Residuals() {
  var semesterId = 'TEST-SEMESTER-2026-M0';
  var backupId = 'BKP-4d7f3a2c-70a6-4e4f-8090-f1fbd043dc0e';
  var moduleCode = 'TEST-MOD-01';
  var ss = m0AssertDisposable_();
  var legacyArchiveHeaders = [
    'Semester ID', 'Semester Name', 'Academic Year', 'Start Date', 'End Date',
    'Status', 'Module Count', 'Final Average', 'Pass Count', 'Fail Count',
    'Distinction Count', 'Archived Date', 'Notes'
  ];
  var legacyTemplateHeaders = [
    'Module Code', 'Module Name', 'Short Name', 'Colour', 'AF Weighting',
    'A1 Weighting', 'A2 Weighting', 'Pass Mark', 'Distinction Mark',
    'AF Calculation Method', 'Rule Status', 'Saved Date'
  ];

  function repairExactLegacyHeaders_(sheetName, expected, legacy) {
    var sheet = ss.getSheetByName(sheetName);
    m0BAssertCleanup_(!!sheet, sheetName + ' is missing');
    var width = Math.max(expected.length, legacy.length);
    var actual = sheet.getRange(HEADER_ROW, 1, 1, width).getValues()[0]
      .map(function (value) { return String(value || '').trim(); });
    var expectedPadded = expected.concat(Array(width - expected.length).fill(''));
    var legacyPadded = legacy.concat(Array(width - legacy.length).fill(''));
    if (JSON.stringify(actual) === JSON.stringify(expectedPadded)) return 'ALREADY_CURRENT';
    m0BAssertCleanup_(JSON.stringify(actual) === JSON.stringify(legacyPadded),
      sheetName + ' header state is neither approved nor the captured legacy schema');
    sheet.getRange(HEADER_ROW, 1, 1, width).clearContent();
    sheet.getRange(HEADER_ROW, 1, 1, expected.length).setValues([expected]);
    m0BValidateHeaders_(sheetName, expected);
    return 'REPAIRED_EXACT_LEGACY_SCHEMA';
  }

  var archiveSchema = repairExactLegacyHeaders_(
    '23 Semester Archive', M0_GATE_B_ARCHIVE_HEADERS, legacyArchiveHeaders);
  var templateSchema = repairExactLegacyHeaders_(
    '25 Module Templates', M0_GATE_B_TEMPLATE_HEADERS, legacyTemplateHeaders);

  var archiveRows = readSheetRows_('23 Semester Archive', HEADER_ROW, 'Semester ID')
    .filter(function (row) { return String(row['Semester ID']) === semesterId; });
  m0BAssertCleanup_(archiveRows.length <= 1, 'residual archive identity is ambiguous');
  if (archiveRows.length) {
    var archive = archiveRows[0], report;
    try { report = JSON.parse(archive['Report JSON'] || '{}'); }
    catch (error) { throw m0BError_('CLEANUP INCOMPLETE', 'residual archive Report JSON is invalid'); }
    m0BAssertCleanup_(report.meta && report.meta.semesterId === semesterId &&
      String(archive['Planner Version']) === '1.5.0-M0-VALIDATION' &&
      String(archive['Status']) === 'Archived' && archive['Immutable'] === true,
      'residual archive composite ownership could not be proven');
    var archiveSheet = ss.getSheetByName('23 Semester Archive');
    var matchingProtections = archiveSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE)
      .filter(function (protection) {
        var range = protection.getRange();
        return range.getRow() <= archive.__row && range.getLastRow() >= archive.__row &&
          String(protection.getDescription() || '') ===
            'Semester Archive ' + semesterId + ' — immutable once written (v1.4.0)';
      });
    m0BAssertCleanup_(matchingProtections.length <= 1,
      'residual archive protection identity is ambiguous');
  }

  var historyRows = readSheetRows_('26 Semester History', HEADER_ROW, 'Semester ID')
    .filter(function (row) { return String(row['Semester ID']) === semesterId; });
  m0BAssertCleanup_(historyRows.length <= 1, 'residual history identity is ambiguous');
  if (historyRows.length && archiveRows.length) {
    m0BAssertCleanup_(String(historyRows[0]['Semester Name']) ===
      String(archiveRows[0]['Semester Name']) &&
      String(historyRows[0]['Academic Year']) === String(archiveRows[0]['Academic Year']),
      'residual history does not belong to the archive composite');
  }

  var templateRows = readSheetRows_('25 Module Templates', HEADER_ROW, 'Module Code')
    .filter(function (row) { return String(row['Module Code']) === moduleCode; });
  m0BAssertCleanup_(templateRows.length <= 1, 'residual template identity is ambiguous');
  if (templateRows.length) {
    m0BAssertCleanup_(String(templateRows[0]['Source Semester ID']) === semesterId,
      'residual template ownership could not be proven');
  }

  var auditRows = m0BAuditRows_();
  var expectedSuffix = [
    ['Semester backup created', backupId],
    ['Semester archived', semesterId],
    ['Module templates saved', semesterId],
    ['Delete Semester workflow prepared', semesterId],
    ['API error', '-']
  ];
  var suffixPresent = auditRows.length >= expectedSuffix.length &&
    expectedSuffix.every(function (expected, index) {
      var row = auditRows[auditRows.length - expectedSuffix.length + index];
      return String(row['Action']) === expected[0] && String(row['Item']) === expected[1];
    });
  var anyResidualAudit = auditRows.some(function (row) {
    return (String(row['Action']) === 'Semester backup created' && String(row['Item']) === backupId) ||
      (expectedSuffix.slice(1, 4).some(function (expected) {
        return String(row['Action']) === expected[0] && String(row['Item']) === semesterId;
      })) || (String(row['Action']) === 'API error' &&
        String(row['Detail']).indexOf('Backup Drive file is trashed and unavailable') !== -1);
  });
  m0BAssertCleanup_(suffixPresent || !anyResidualAudit,
    'residual Automation Log ownership is incomplete or non-contiguous');

  if (archiveRows.length) {
    m0BDeleteOwnedRow_('23 Semester Archive', 'Semester ID', semesterId, function (row) {
      var parsed;
      try { parsed = JSON.parse(row['Report JSON'] || '{}'); } catch (ignored) { return false; }
      return parsed.meta && parsed.meta.semesterId === semesterId &&
        String(row['Planner Version']) === '1.5.0-M0-VALIDATION' &&
        row['Immutable'] === true;
    });
  }
  if (historyRows.length) {
    m0BDeleteOwnedRow_('26 Semester History', 'Semester ID', semesterId, function (row) {
      return String(row['Semester Name']) === 'M0 Disposable Semester' &&
        String(row['Academic Year']) === '2026';
    });
  }
  if (templateRows.length) {
    m0BDeleteOwnedRow_('25 Module Templates', 'Module Code', moduleCode, function (row) {
      return String(row['Source Semester ID']) === semesterId;
    });
  }

  var statusRows = readSheetRows_('02 Settings', HEADER_ROW, 'Setting').filter(function (row) {
    return String(row['Setting']) === 'Semester status';
  });
  m0BAssertCleanup_(statusRows.length === 1, 'Semester status setting identity is ambiguous');
  var currentStatus = String(statusRows[0]['Value']);
  m0BAssertCleanup_(currentStatus === 'Complete' || currentStatus === 'Active',
    'Semester status is outside the captured residual state');
  if (currentStatus === 'Complete') {
    ss.getSheetByName('02 Settings').getRange(statusRows[0].__row,
      col_(getColMap_(ss.getSheetByName('02 Settings')), 'Value')).setValue('Active');
  }

  if (suffixPresent) {
    var logSheet = ss.getSheetByName('20 Automation Log');
    for (var i = 0; i < expectedSuffix.length; i++) logSheet.deleteRow(logSheet.getLastRow());
  }

  var result = {
    status: 'PASS', testId: 'M0-LV-11-RESIDUAL-RECOVERY',
    archiveSchema: archiveSchema, templateSchema: templateSchema,
    archiveRemoved: archiveRows.length === 1, historyRemoved: historyRows.length === 1,
    templateRemoved: templateRows.length === 1, settingRestored: currentStatus === 'Complete',
    auditSuffixRemoved: suffixPresent, backupFileDisposition: 'REMAINS_IN_DISPOSABLE_TRASH'
  };
  console.log(JSON.stringify(result));
  return result;
}

// Residual cleanup for the M0-LV-11 retest at 2026-07-18 09:38 SAST.
// Validation-only compare-and-swap recovery for the preserved Setting and
// exact Automation Log suffix. It creates no workflow or business mutation.
function m0RecoverGateBTest11SettingsResiduals() {
  var ss = m0AssertDisposable_();
  var semesterId = 'TEST-SEMESTER-2026-M0';
  var workflowToken = 'a593f2b4-02c5-41fb-9505-63b82d09bdbb';
  var backupId = 'BKP-3ac91141-42fe-4977-8e56-d733c55e0bf6';
  var moduleCode = 'TEST-MOD-01';
  var settingOwnership = { originalValue: 'Active', testWrittenValue: 'Complete' };
  var settingsBefore = readSheetRows_('02 Settings', HEADER_ROW, 'Setting');
  var unrelatedSettingsBefore = settingsBefore.filter(function (row) {
    return String(row['Setting']) !== 'Semester status';
  }).map(function (row) {
    return [row['Setting'], m0CanonicalValue_(row['Value'])];
  });

  m0BAssertCleanup_(!m0BWorkflowRowByToken_(workflowToken),
    'test-owned workflow metadata remains');
  m0BAssertCleanup_(!m0BBackupRowById_(backupId),
    'test-owned backup metadata remains');
  m0BAssertCleanup_(!readSheetRows_('23 Semester Archive', HEADER_ROW, 'Semester ID')
    .some(function (row) { return String(row['Semester ID']) === semesterId; }),
    'test-owned archive row remains');
  m0BAssertCleanup_(!readSheetRows_('26 Semester History', HEADER_ROW, 'Semester ID')
    .some(function (row) { return String(row['Semester ID']) === semesterId; }),
    'test-owned history row remains');
  m0BAssertCleanup_(!readSheetRows_('25 Module Templates', HEADER_ROW, 'Module Code')
    .some(function (row) {
      return String(row['Module Code']) === moduleCode &&
        String(row['Source Semester ID']) === semesterId;
    }), 'test-owned template row remains');

  var auditRows = m0BAuditRows_();
  var expectedSuffix = [
    ['Semester backup created', backupId],
    ['Semester archived', semesterId],
    ['Module templates saved', semesterId],
    ['Delete Semester workflow prepared', semesterId],
    ['API error', '-']
  ];
  m0BAssertCleanup_(auditRows.length >= expectedSuffix.length,
    'captured Automation Log suffix is missing');
  var prefixLength = auditRows.length - expectedSuffix.length;
  var prefixBefore = auditRows.slice(0, prefixLength).map(function (row) {
    return [row['Timestamp'], row['Action'], row['Item'], row['Detail']];
  });
  expectedSuffix.forEach(function (expected, index) {
    var row = auditRows[prefixLength + index];
    m0BAssertCleanup_(String(row['Action']) === expected[0] &&
      String(row['Item']) === expected[1],
      'captured Automation Log suffix does not match exact ownership');
  });
  m0BAssertCleanup_(String(auditRows[auditRows.length - 1]['Detail'])
    .indexOf('Backup Drive file is trashed and unavailable') !== -1,
    'captured Automation Log rejection detail does not match ownership');

  var currentBefore = m0BReadUniqueSetting_('Semester status');
  m0BRestoreOwnedSetting_('Semester status', settingOwnership);
  var logSheet = ss.getSheetByName('20 Automation Log');
  for (var i = 0; i < expectedSuffix.length; i++) logSheet.deleteRow(logSheet.getLastRow());

  var auditAfter = m0BAuditRows_().map(function (row) {
    return [row['Timestamp'], row['Action'], row['Item'], row['Detail']];
  });
  m0BAssertCleanup_(JSON.stringify(auditAfter) === JSON.stringify(prefixBefore),
    'earlier Automation Log history changed during residual cleanup');
  var unrelatedSettingsAfter = readSheetRows_('02 Settings', HEADER_ROW, 'Setting')
    .filter(function (row) { return String(row['Setting']) !== 'Semester status'; })
    .map(function (row) { return [row['Setting'], m0CanonicalValue_(row['Value'])]; });
  m0BAssertCleanup_(JSON.stringify(unrelatedSettingsAfter) ===
    JSON.stringify(unrelatedSettingsBefore),
    'unrelated Settings changed during residual cleanup');

  var result = {
    status: 'PASS', testId: 'M0-LV-11-SETTINGS-RESIDUAL-RECOVERY',
    originalValue: settingOwnership.originalValue,
    testWrittenValue: settingOwnership.testWrittenValue,
    currentValueBeforeRemediation: currentBefore,
    currentValueAfterRemediation: m0BReadUniqueSetting_('Semester status'),
    automationLogSuffixRemoved: expectedSuffix.length,
    earlierAutomationLogRowsPreserved: prefixBefore.length,
    unrelatedSettingsPreserved: unrelatedSettingsBefore.length,
    corruptedBackupDisposition: 'REMAINS_IN_DISPOSABLE_TRASH'
  };
  console.log(JSON.stringify(result));
  return result;
}

// Residual cleanup for the M0-LV-12 execution at 2026-07-18 10:08 SAST.
// Uses the immutable, test-owned backup only as typed baseline evidence and
// restores infrastructure through identity-scoped rows and compare-and-swap.
function m0RecoverGateBTest12Residuals() {
  var ss = m0AssertDisposable_();
  var semesterId = 'TEST-SEMESTER-2026-M0';
  var workflowToken = 'dfc5f88d-7454-434e-8d4a-a3850d5dd499';
  var backupId = 'BKP-7388ffdb-2bd5-4b78-8669-2e89768f2c2c';
  var driveFileId = '12Oyu-RJOIexQXDdLdzeo1xfdhDbVjtEq';
  var moduleCode = 'TEST-MOD-01';
  var nonTargetBaseline = m0BNonTargetBaseline_(['02 Settings', '20 Automation Log']);

  var file = DriveApp.getFileById(driveFileId);
  m0BAssertCleanup_(file.isTrashed() === true,
    'preserved M0-LV-12 backup does not have the approved trash disposition');
  var payload;
  try { payload = JSON.parse(file.getBlob().getDataAsString()); }
  catch (error) { throw m0BError_('CLEANUP INCOMPLETE', 'preserved backup JSON is invalid'); }
  m0BAssertCleanup_(payload.backupId === backupId &&
    payload.spreadsheetId === ss.getId() && payload.semesterId === semesterId,
    'preserved backup identity does not match M0-LV-12');

  function backupValues_(sheetName) {
    var dump = payload.sheets && payload.sheets[sheetName];
    m0BAssertCleanup_(dump && dump.sheetName === sheetName && Array.isArray(dump.values),
      'preserved backup lacks typed baseline for ' + sheetName);
    return dump.values;
  }

  function comparableValues_(values) {
    return values.map(function (row) {
      return row.map(function (value) {
        return value instanceof Date ? value.toISOString() : value;
      });
    });
  }

  m0BTargetNames_().slice(0, 9).forEach(function (sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    var current = sheet.getDataRange().getValues();
    m0BAssertCleanup_(JSON.stringify(comparableValues_(current)) ===
      JSON.stringify(backupValues_(sheetName)),
      'business fixture was not restored before residual recovery: ' + sheetName);
  });

  m0BAssertCleanup_(!m0BWorkflowRowByToken_(workflowToken),
    'test-owned workflow metadata remains');
  m0BAssertCleanup_(!m0BBackupRowById_(backupId),
    'test-owned backup metadata remains');
  m0BAssertCleanup_(!readSheetRows_('23 Semester Archive', HEADER_ROW, 'Semester ID')
    .some(function (row) { return String(row['Semester ID']) === semesterId; }),
    'test-owned archive row remains');
  m0BAssertCleanup_(!readSheetRows_('26 Semester History', HEADER_ROW, 'Semester ID')
    .some(function (row) { return String(row['Semester ID']) === semesterId; }),
    'test-owned history row remains');
  m0BAssertCleanup_(!readSheetRows_('25 Module Templates', HEADER_ROW, 'Module Code')
    .some(function (row) {
      return String(row['Module Code']) === moduleCode &&
        String(row['Source Semester ID']) === semesterId;
    }), 'test-owned template row remains');

  var settingsValues = backupValues_('02 Settings');
  var settingHeaders = settingsValues[HEADER_ROW - 1];
  var settingLabelColumn = settingHeaders.indexOf('Setting');
  var settingValueColumn = settingHeaders.indexOf('Value');
  m0BAssertCleanup_(settingLabelColumn !== -1 && settingValueColumn !== -1,
    'backup Settings schema is malformed');
  function originalSetting_(label) {
    var matches = settingsValues.slice(HEADER_ROW).filter(function (row) {
      return String(row[settingLabelColumn]) === label;
    });
    m0BAssertCleanup_(matches.length === 1,
      'backup Settings identity is ambiguous: ' + label);
    return matches[0][settingValueColumn];
  }

  var obsoleteManualRows = settingsValues.slice(HEADER_ROW).filter(function (row) {
    return String(row[settingLabelColumn]) === 'Semester manually finished';
  });
  var canonicalManualRows = settingsValues.slice(HEADER_ROW).filter(function (row) {
    return String(row[settingLabelColumn]) === 'Semester manually finished (TRUE/FALSE)';
  });
  m0BAssertCleanup_(obsoleteManualRows.length === 1 &&
    obsoleteManualRows[0][settingValueColumn] === false &&
    canonicalManualRows.length === 0,
    'backup is not the exact obsolete pre-alignment validation fixture');

  var statusBefore = m0BReadUniqueSetting_('Semester status');
  var manualBefore = m0BReadUniqueSetting_('Semester manually finished (TRUE/FALSE)');
  var statusOwnership = {
    originalValue: originalSetting_('Semester status'),
    testWrittenValue: 'Complete'
  };
  var manualOwnership = {
    // This one identified disposable backup is obsolete and discarded. The
    // canonical fixture baseline is authoritative; this is not legacy-read
    // compatibility for production backups.
    originalValue: false,
    testWrittenValue: false
  };
  m0BRestoreOwnedSetting_('Semester status', statusOwnership);
  m0BRestoreOwnedSetting_('Semester manually finished (TRUE/FALSE)', manualOwnership);

  var auditRows = m0BAuditRows_();
  m0BAssertCleanup_(auditRows.length === 2 &&
    String(auditRows[0]['Action']) === 'Semester reset (active data cleared)' &&
    String(auditRows[1]['Action']) === 'Semester deleted (confirmed)' &&
    String(auditRows[1]['Item']) === semesterId,
    'current Automation Log is not the exact M0-LV-12 destructive suffix');
  var auditSheet = ss.getSheetByName('20 Automation Log');
  auditSheet.deleteRow(auditRows[1].__row);
  auditSheet.deleteRow(auditRows[0].__row);
  var auditBaselineValues = backupValues_('20 Automation Log').slice(HEADER_ROW);
  if (auditBaselineValues.length) {
    var restoredAuditValues = auditBaselineValues.map(function (row) {
      var copy = row.slice();
      if (typeof copy[0] === 'string' && /^\d{4}-\d{2}-\d{2}T/.test(copy[0])) {
        copy[0] = new Date(copy[0]);
      }
      return copy;
    });
    auditSheet.getRange(HEADER_ROW + 1, 1, restoredAuditValues.length,
      restoredAuditValues[0].length).setValues(restoredAuditValues);
  }
  m0BAssertCleanup_(JSON.stringify(comparableValues_(auditSheet.getDataRange().getValues())) ===
    JSON.stringify(backupValues_('20 Automation Log')),
    'Automation Log historical baseline restoration failed');
  m0BAssertSame_(nonTargetBaseline, Object.keys(nonTargetBaseline), 'CLEANUP INCOMPLETE');

  var result = {
    status: 'PASS', testId: 'M0-LV-12-RESIDUAL-RECOVERY',
    statusBefore: statusBefore,
    statusAfter: m0BReadUniqueSetting_('Semester status'),
    manualBefore: manualBefore,
    manualAfter: m0BReadUniqueSetting_('Semester manually finished (TRUE/FALSE)'),
    businessSheetsVerified: 9,
    destructiveAuditRowsRemoved: 2,
    historicalAuditRowsRestored: auditBaselineValues.length,
    unrelatedInfrastructurePreserved: Object.keys(nonTargetBaseline).length,
    backupDisposition: 'REMAINS_IN_DISPOSABLE_TRASH'
  };
  console.log(JSON.stringify(result));
  return result;
}
