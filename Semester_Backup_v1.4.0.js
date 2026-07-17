/**
 * Semester_Backup_v1.4.0.gs — Engineering Planner OS v1.4.0, Semester
 * Lifecycle & Analytics.
 *
 * Creates a complete JSON snapshot of the semester-lifecycle sheets and
 * records an independently stored SHA-256 fingerprint in "24 Semester
 * Backups". Google Drive stores the backup file; the sheet row is the
 * independent index and integrity record used by verification.
 */

var SEM_BACKUP_FORMAT_VERSION = 'EPOS_BACKUP_FORMAT_2';
var SEM_BACKUP_CHECKSUM_ALGORITHM = 'SHA-256';

/** Raw, complete dump of one sheet's used range -- headers included. */
function semDumpSheet_(sheetName) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sheet) return null;
  var lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return { sheetName: sheetName, values: [] };
  var values = sheet.getRange(1, 1, lastRow, lastCol).getValues().map(function (row) {
    return row.map(function (cell) { return (cell instanceof Date) ? cell.toISOString() : cell; });
  });
  return { sheetName: sheetName, values: values };
}

function semGetOrCreateBackupFolder_() {
  var name = 'Engineering Planner OS Backups';
  var it = DriveApp.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(name);
}

/** The fixed, required v1.4 backup boundary. */
var SEM_BACKUP_SHEETS = [
  '02 Settings', '04 Module Rules', '03 Modules', '17 Archive', '23 Semester Archive',
  '25 Module Templates', '27 Semester Workflows', '07 Academic Inbox', '08 Assignments', '09 Assessments',
  '16 Notes', '15 Resources', '12 Study Planner', '21 Study Tasks',
  '13 Revision Tracker', '11 Marks Tracker', '10 AF Components', '19 Attendance', '20 Automation Log'
];

function semBackupCanonicalContent_(payload) {
  var orderedSheets = {};
  SEM_BACKUP_SHEETS.forEach(function (name) { orderedSheets[name] = payload.sheets[name]; });
  return {
    formatVersion: payload.formatVersion,
    semesterId: payload.semesterId,
    plannerVersion: payload.plannerVersion,
    spreadsheetId: payload.spreadsheetId,
    spreadsheetName: payload.spreadsheetName,
    sheets: orderedSheets
  };
}

function semBackupSha256_(value) {
  var bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    value,
    Utilities.Charset.UTF_8
  );
  return bytes.map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
}

function semBackupChecksum_(payload) {
  return semBackupSha256_(JSON.stringify(semBackupCanonicalContent_(payload)));
}

function semBackupNotes_(checksum) {
  // Milestone 0 compatibility decision: the existing Notes column temporarily
  // stores machine-readable integrity metadata so no v1.4 schema migration is
  // required. Technical debt: move this value to a dedicated checksum column
  // in a future versioned migration, while retaining legacy-read support here.
  return SEM_BACKUP_FORMAT_VERSION + '; ' + SEM_BACKUP_CHECKSUM_ALGORITHM + ': ' + checksum;
}

function semBackupChecksumFromNotes_(notes) {
  var match = String(notes || '').match(/^EPOS_BACKUP_FORMAT_2; SHA-256: ([0-9a-f]{64})$/i);
  return match ? match[1].toLowerCase() : '';
}

function semBackupSheetSetMatches_(names) {
  if (!Array.isArray(names) || names.length !== SEM_BACKUP_SHEETS.length) return false;
  var seen = {};
  names.forEach(function (name) { seen[name] = (seen[name] || 0) + 1; });
  return SEM_BACKUP_SHEETS.every(function (name) { return seen[name] === 1; });
}

function semBackupDumpIsValid_(name, dump) {
  if (!dump || dump.sheetName !== name || !Array.isArray(dump.values) || dump.values.length === 0) return false;
  if (!dump.values.every(function (row) { return Array.isArray(row) && row.length > 0; })) return false;
  var width = dump.values[0].length;
  return dump.values.every(function (row) { return row.length === width; });
}

function semCleanupIncompleteBackupFile_(file) {
  if (!file || typeof file.setTrashed !== 'function') {
    return { status: 'UNKNOWN', detail: 'The incomplete Drive file cleanup status is unknown.' };
  }
  try {
    file.setTrashed(true);
    return { status: 'SUCCEEDED', detail: 'The incomplete Drive file was moved to trash.' };
  } catch (e) {
    return { status: 'FAILED', detail: 'The incomplete Drive file could not be moved to trash: ' + e.message };
  }
}

function semCreateBackup_(semesterId) {
  var ss = SpreadsheetApp.getActive();
  var indexSheet = ss.getSheetByName(SEM_BACKUPS_SHEET);
  if (!indexSheet) throw new Error('"24 Semester Backups" not found — run Semester_Lifecycle_v1.4.0_Migration.gs first.');

  var payload = {
    formatVersion: SEM_BACKUP_FORMAT_VERSION,
    createdAt: new Date().toISOString(),
    semesterId: semesterId || getSetting_('Semester ID') || '',
    plannerVersion: getSetting_('Planner version') || '1.4.0',
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName(),
    sheets: {}
  };
  var missingOrInvalid = [];
  SEM_BACKUP_SHEETS.forEach(function (name) {
    var dump = semDumpSheet_(name);
    if (!semBackupDumpIsValid_(name, dump)) missingOrInvalid.push(name);
    else payload.sheets[name] = dump;
  });
  if (missingOrInvalid.length) {
    throw new Error('Backup was not created because required sheets are missing or empty: ' + missingOrInvalid.join(', '));
  }

  var existingIds = readSheetRows_(SEM_BACKUPS_SHEET, HEADER_ROW, 'Backup ID').map(function (r) { return r['Backup ID']; });
  var backupId = nextId_('BKP', existingIds);
  payload.backupId = backupId;
  payload.contentChecksum = semBackupChecksum_(payload);

  var folder = semGetOrCreateBackupFolder_();
  var fileName = 'EngineeringPlannerOS_Backup_' + backupId + '_' + Utilities.formatDate(new Date(), TIMEZONE, 'yyyyMMdd_HHmmss') + '.json';
  var file = folder.createFile(fileName, JSON.stringify(payload), MimeType.PLAIN_TEXT);

  try {
    var row = Math.max(indexSheet.getLastRow() + 1, HEADER_ROW + 1);
    indexSheet.getRange(row, 1, 1, 8).setValues([[
      backupId, payload.semesterId, new Date(), payload.plannerVersion, file.getId(), file.getUrl(),
      SEM_BACKUP_SHEETS.join(', '), semBackupNotes_(payload.contentChecksum)
    ]]);
  } catch (e) {
    var cleanup = semCleanupIncompleteBackupFile_(file);
    throw new Error('Backup index write failed. Cleanup ' + cleanup.status + ': ' + cleanup.detail + ' Index error: ' + e.message);
  }

  try { logAutomation_('Semester backup created', backupId, 'Success', file.getUrl()); } catch (ignored) {}
  return {
    backupId: backupId, fileId: file.getId(), fileUrl: file.getUrl(),
    sheetsIncluded: SEM_BACKUP_SHEETS.slice(), timestamp: payload.createdAt
  };
}

function semListBackups_() {
  return readSheetRows_(SEM_BACKUPS_SHEET, HEADER_ROW, 'Backup ID').map(function (r) {
    return {
      backupId: r['Backup ID'], semesterId: r['Semester ID'],
      timestamp: r['Timestamp'] ? new Date(r['Timestamp']).toISOString() : '',
      plannerVersion: r['Planner Version'], fileId: r['Drive File ID'], fileUrl: r['Drive File URL'],
      sheetsIncluded: r['Sheets Included'], notes: r['Notes']
    };
  }).sort(function (a, b) { return (b.timestamp || '').localeCompare(a.timestamp || ''); });
}

/**
 * Verifies the independent index, file identity, fixed sheet boundary,
 * structural completeness, and SHA-256 fingerprint. Legacy rows without an
 * integrity fingerprint remain listable but cannot authorize destructive work.
 */
function semVerifyBackup_(backupId) {
  var failure = function (reason, fileFound, jsonValid, missingSheets) {
    return { ok: false, fileFound: fileFound, jsonValid: jsonValid, missingSheets: missingSheets || [], reason: reason };
  };
  var row = readSheetRows_(SEM_BACKUPS_SHEET, HEADER_ROW, 'Backup ID').filter(function (r) { return r['Backup ID'] === backupId; })[0];
  if (!row) return failure('No backup row found for ' + backupId, false, false);

  var file;
  try { file = DriveApp.getFileById(row['Drive File ID']); } catch (e) {
    return failure('Drive file not found: ' + e.message, false, false);
  }
  try {
    if (file.isTrashed()) return failure('Backup Drive file is trashed and unavailable.', false, false);
  } catch (e) {
    return failure('Backup Drive file availability could not be confirmed: ' + e.message, false, false);
  }
  var parsed;
  try { parsed = JSON.parse(file.getBlob().getDataAsString()); } catch (e) {
    return failure('Backup file is not valid JSON: ' + e.message, true, false);
  }

  var expectedChecksum = semBackupChecksumFromNotes_(row['Notes']);
  if (!expectedChecksum) return failure('Legacy or malformed backup metadata has no verifiable SHA-256 fingerprint.', true, true);

  if (parsed.formatVersion !== SEM_BACKUP_FORMAT_VERSION) return failure('Unsupported backup format version.', true, true);
  if (parsed.backupId !== row['Backup ID'] || parsed.backupId !== backupId) return failure('Backup ID does not match the independent index.', true, true);
  if (String(parsed.semesterId || '') !== String(row['Semester ID'] || '')) return failure('Semester ID does not match the independent index.', true, true);
  if (String(parsed.plannerVersion || '') !== String(row['Planner Version'] || '')) return failure('Planner version does not match the independent index.', true, true);
  if (String(parsed.spreadsheetId || '') !== String(SpreadsheetApp.getActive().getId())) return failure('Backup belongs to a different spreadsheet.', true, true);

  var indexedSheets = String(row['Sheets Included'] || '').split(',').map(function (s) { return s.trim(); }).filter(Boolean);
  if (!semBackupSheetSetMatches_(indexedSheets)) return failure('Indexed sheet set does not match the required backup boundary.', true, true);
  var payloadSheets = parsed.sheets && typeof parsed.sheets === 'object' ? Object.keys(parsed.sheets) : [];
  if (!semBackupSheetSetMatches_(payloadSheets)) {
    var absent = SEM_BACKUP_SHEETS.filter(function (name) { return payloadSheets.indexOf(name) === -1; });
    return failure('Payload sheet set does not match the required backup boundary.', true, true, absent);
  }
  var invalid = SEM_BACKUP_SHEETS.filter(function (name) { return !semBackupDumpIsValid_(name, parsed.sheets[name]); });
  if (invalid.length) return failure('One or more required sheet dumps are empty or structurally invalid.', true, true, invalid);

  var actualChecksum = semBackupChecksum_(parsed);
  if (parsed.contentChecksum !== expectedChecksum || actualChecksum !== expectedChecksum) {
    return failure('Backup content fingerprint does not match the independent index.', true, true);
  }
  return { ok: true, fileFound: true, jsonValid: true, missingSheets: [], checksumValid: true };
}
