/**
 * Local characterization/safety tests for Semester_Backup_v1.4.0.js.
 * Run with: node Semester_Backup_v1.4.0_Test.cjs
 */
'use strict';

const assert = require('assert');
const crypto = require('crypto');
const fs = require('fs');
const vm = require('vm');

const source = fs.readFileSync(__dirname + '/Semester_Backup_v1.4.0.js', 'utf8');

function makeContext() {
  const state = { rows: [], files: {}, createFileCalls: 0, trashed: false };
  const active = {
    getId: () => 'sheet-123',
    getName: () => 'Planner',
    getSheetByName: () => null
  };
  const context = {
    console,
    Date,
    JSON,
    Array,
    String,
    Object,
    SEM_BACKUPS_SHEET: '24 Semester Backups',
    HEADER_ROW: 11,
    TIMEZONE: 'Africa/Johannesburg',
    MimeType: { PLAIN_TEXT: 'text/plain' },
    SpreadsheetApp: { getActive: () => active },
    DriveApp: {
      getFileById(id) {
        if (!state.files[id]) throw new Error('missing');
        return { getBlob: () => ({ getDataAsString: () => state.files[id] }) };
      },
      getFoldersByName() {
        return { hasNext: () => true, next: () => context.__folder };
      }
    },
    Utilities: {
      DigestAlgorithm: { SHA_256: 'SHA_256' },
      Charset: { UTF_8: 'UTF_8' },
      computeDigest(_algorithm, value) {
        return Array.from(crypto.createHash('sha256').update(value, 'utf8').digest()).map(b => b > 127 ? b - 256 : b);
      },
      formatDate: () => '20260717_120000'
    },
    readSheetRows_: () => state.rows,
    nextId_: () => 'BKP-001',
    getSetting_: label => label === 'Planner version' ? '1.4.0' : 'SEM-001',
    logAutomation_: () => {}
  };
  context.__folder = {
    createFile(_name, body) {
      state.createFileCalls++;
      state.files['file-1'] = body;
      return {
        getId: () => 'file-1',
        getUrl: () => 'https://drive.test/file-1',
        setTrashed: value => { state.trashed = value; }
      };
    }
  };
  vm.createContext(context);
  vm.runInContext(source, context, { filename: 'Semester_Backup_v1.4.0.js' });
  return { context, state, active };
}

function validFixture(env) {
  const c = env.context;
  const sheets = {};
  c.SEM_BACKUP_SHEETS.forEach(name => {
    sheets[name] = { sheetName: name, values: [['Header'], [0]] };
  });
  const payload = {
    formatVersion: c.SEM_BACKUP_FORMAT_VERSION,
    createdAt: '2026-07-17T10:00:00.000Z',
    backupId: 'BKP-001',
    semesterId: 'SEM-001',
    plannerVersion: '1.4.0',
    spreadsheetId: 'sheet-123',
    spreadsheetName: 'Planner',
    sheets
  };
  payload.contentChecksum = c.semBackupChecksum_(payload);
  env.state.rows = [{
    'Backup ID': 'BKP-001',
    'Semester ID': 'SEM-001',
    'Planner Version': '1.4.0',
    'Drive File ID': 'file-1',
    'Sheets Included': c.SEM_BACKUP_SHEETS.join(', '),
    'Notes': c.semBackupNotes_(payload.contentChecksum)
  }];
  env.state.files['file-1'] = JSON.stringify(payload);
  return payload;
}

const tests = [];
function test(name, fn) { tests.push({ name, fn }); }

test('backup boundary includes every sheet cleared by semester reset', () => {
  const env = makeContext();
  const clearedByReset = [
    '08 Assignments', '09 Assessments', '12 Study Planner', '21 Study Tasks',
    '13 Revision Tracker', '11 Marks Tracker', '10 AF Components',
    '19 Attendance', '07 Academic Inbox', '20 Automation Log'
  ];
  clearedByReset.forEach(name => assert.ok(env.context.SEM_BACKUP_SHEETS.includes(name), name));
});

test('accepts a complete matching backup and preserves zero values', () => {
  const env = makeContext();
  validFixture(env);
  assert.strictEqual(env.context.semVerifyBackup_('BKP-001').ok, true);
});

test('rejects content changed after fingerprinting', () => {
  const env = makeContext();
  const payload = validFixture(env);
  payload.sheets['03 Modules'].values[1][0] = 99;
  env.state.files['file-1'] = JSON.stringify(payload);
  assert.match(env.context.semVerifyBackup_('BKP-001').reason, /fingerprint/);
});

test('rejects empty sheet values', () => {
  const env = makeContext();
  const payload = validFixture(env);
  payload.sheets['03 Modules'].values = [];
  payload.contentChecksum = env.context.semBackupChecksum_(payload);
  env.state.rows[0].Notes = env.context.semBackupNotes_(payload.contentChecksum);
  env.state.files['file-1'] = JSON.stringify(payload);
  const result = env.context.semVerifyBackup_('BKP-001');
  assert.strictEqual(result.ok, false);
  assert.deepStrictEqual(Array.from(result.missingSheets), ['03 Modules']);
});

test('rejects an omitted required sheet even when metadata also omits it', () => {
  const env = makeContext();
  const payload = validFixture(env);
  delete payload.sheets['03 Modules'];
  env.state.rows[0]['Sheets Included'] = env.context.SEM_BACKUP_SHEETS.filter(n => n !== '03 Modules').join(', ');
  env.state.files['file-1'] = JSON.stringify(payload);
  assert.strictEqual(env.context.semVerifyBackup_('BKP-001').ok, false);
});

test('rejects mismatched backup, semester, planner, and spreadsheet identities', () => {
  ['backupId', 'semesterId', 'plannerVersion', 'spreadsheetId'].forEach(field => {
    const env = makeContext();
    const payload = validFixture(env);
    payload[field] = 'wrong';
    payload.contentChecksum = env.context.semBackupChecksum_(payload);
    env.state.files['file-1'] = JSON.stringify(payload);
    assert.strictEqual(env.context.semVerifyBackup_('BKP-001').ok, false, field);
  });
});

test('rejects legacy metadata without a fingerprint', () => {
  const env = makeContext();
  validFixture(env);
  env.state.rows[0].Notes = 'Automatic backup';
  assert.match(env.context.semVerifyBackup_('BKP-001').reason, /no verifiable/);
});

test('reports missing files and invalid JSON', () => {
  const missing = makeContext();
  validFixture(missing);
  delete missing.state.files['file-1'];
  assert.strictEqual(missing.context.semVerifyBackup_('BKP-001').fileFound, false);

  const invalid = makeContext();
  validFixture(invalid);
  invalid.state.files['file-1'] = '{';
  const result = invalid.context.semVerifyBackup_('BKP-001');
  assert.strictEqual(result.fileFound, true);
  assert.strictEqual(result.jsonValid, false);
});

test('does not create a Drive file when the backup index is absent', () => {
  const env = makeContext();
  assert.throws(() => env.context.semCreateBackup_('SEM-001'), /not found/);
  assert.strictEqual(env.state.createFileCalls, 0);
});

test('does not create a Drive file when a required sheet is missing', () => {
  const env = makeContext();
  env.active.getSheetByName = name => name === '24 Semester Backups' ? {
    getLastRow: () => 11,
    getRange: () => ({ setValues: () => {} })
  } : null;
  env.context.semDumpSheet_ = name => name === '03 Modules' ? null : { sheetName: name, values: [['Header']] };
  assert.throws(() => env.context.semCreateBackup_('SEM-001'), /03 Modules/);
  assert.strictEqual(env.state.createFileCalls, 0);
});

test('reports successful cleanup when the independent index write fails', () => {
  const env = makeContext();
  env.active.getSheetByName = name => name === '24 Semester Backups' ? {
    getLastRow: () => 11,
    getRange: () => ({ setValues: () => { throw new Error('write denied'); } })
  } : null;
  env.context.semDumpSheet_ = name => ({ sheetName: name, values: [['Header']] });
  assert.throws(() => env.context.semCreateBackup_('SEM-001'), /Cleanup SUCCEEDED: The incomplete Drive file was moved to trash/);
  assert.strictEqual(env.state.createFileCalls, 1);
  assert.strictEqual(env.state.trashed, true);
});

test('reports cleanup failure when trashing the incomplete Drive file fails', () => {
  const env = makeContext();
  env.active.getSheetByName = name => name === '24 Semester Backups' ? {
    getLastRow: () => 11,
    getRange: () => ({ setValues: () => { throw new Error('write denied'); } })
  } : null;
  env.context.semDumpSheet_ = name => ({ sheetName: name, values: [['Header']] });
  env.context.__folder.createFile = (_name, body) => {
    env.state.createFileCalls++;
    env.state.files['file-1'] = body;
    return {
      getId: () => 'file-1',
      getUrl: () => 'https://drive.test/file-1',
      setTrashed: () => { throw new Error('trash denied'); }
    };
  };
  assert.throws(
    () => env.context.semCreateBackup_('SEM-001'),
    /Cleanup FAILED: The incomplete Drive file could not be moved to trash: trash denied/
  );
  assert.strictEqual(env.state.createFileCalls, 1);
});

let failed = 0;
tests.forEach(({ name, fn }) => {
  try {
    fn();
    console.log('PASS', name);
  } catch (error) {
    failed++;
    console.error('FAIL', name, '\n ', error.stack);
  }
});
console.log(`\n${tests.length - failed} passed, ${failed} failed`);
if (failed) process.exitCode = 1;
