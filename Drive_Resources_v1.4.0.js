/**
 * Drive_Resources_v1.4.0.gs — OPTIONAL, ONE-TIME, MANUAL migration, plus the
 * feature it unlocks.
 *
 * Do not run the migration automatically. Run it yourself from the Apps
 * Script editor (runDriveResourcesV140Migration) once, before uploading a
 * file from the Resources screen.
 *
 * "15 Resources" was deliberately reference-only before this ("Link or file
 * reference" is free text — never a file upload; see api_addResource's own
 * doc comment). This release adds an actual upload path alongside it,
 * without changing that existing text-reference flow at all:
 *
 *  1. Two new, purely additive columns on "15 Resources": "Drive category"
 *     (one of VALID_DRIVE_CATEGORIES below, blank for every existing
 *     text-reference row) and "Drive file ID" (the uploaded file's Drive ID,
 *     for future reference/cleanup — blank likewise).
 *  2. api_uploadResourceFile(form): decodes a base64 file the frontend read
 *     client-side, uploads it into
 *     "<Semester name>/<Module>/<category>" (created on first use, reused
 *     after that — see ensureModuleCategoryFolder_), and writes a normal
 *     Resources row pointing at the uploaded file's URL — same row shape
 *     api_addResource already produces, so every existing Resources screen
 *     feature (Status, Progress flags, Score, Reviewed) works on an
 *     uploaded file exactly like it does on a hand-typed link.
 *
 * The semester folder name always tracks 02 Settings > "Semester name" —
 * next semester's upload just works once that setting is updated (e.g. via
 * Start New Semester / Update Semester Dates), never hardcoded.
 *
 * Idempotent: safe to run more than once. Folder creation
 * (ensureSemesterDriveFolder_ / ensureModuleCategoryFolder_) is equally
 * idempotent — re-uploading never creates a duplicate folder tree, it just
 * reuses the one already there.
 */
var VALID_DRIVE_CATEGORIES = ["Completed tutorial tests", "Assessment or exam", "Questions for professors"];

// Hard backstop against a runaway upload; the frontend also warns well
// before this (see saveDriveUpload_ in Index.html) so this should rarely
// actually fire.
var MAX_UPLOAD_BYTES = 20 * 1024 * 1024;

function runDriveResourcesV140Migration() {
  var report = [];
  var sheet = SpreadsheetApp.getActive().getSheetByName(RESOURCES_SHEET);
  if (!sheet) {
    migrationReport_('Drive Resources v1.4.0 migration', ['"' + RESOURCES_SHEET + '" sheet not found — nothing to do.']);
    return;
  }
  ['Drive category', 'Drive file ID'].forEach(function (name) {
    var result = ensureColumn_(sheet, HEADER_ROW, name);
    report.push((result.added ? 'Added column: "' : 'Already present: "') + name + '"');
  });
  report.push('');
  report.push('Uploads go to "<Semester name from 02 Settings>/<Module>/<category>" in your Google Drive, created automatically on first upload.');
  report.push('No existing Resources row or the "Link or file reference" text-entry flow was touched.');
  report.push('Safe to run again at any time.');
  migrationReport_('Drive Resources v1.4.0 migration', report);
}

/** Gets-or-creates a single named subfolder directly under `parent`. Never
 *  creates a duplicate if one already exists (first match wins, same
 *  reuse-not-recreate principle as ensureColumn_ elsewhere in this project). */
function getOrCreateSubfolder_(parent, name) {
  var existing = parent.getFoldersByName(name);
  if (existing.hasNext()) return existing.next();
  return parent.createFolder(name);
}

/** Top-level Drive folder for the current semester, named after 02 Settings
 *  > "Semester name" (falls back to "Semester" if that setting is blank —
 *  never fails outright just because Settings hasn't been filled in). */
function ensureSemesterDriveFolder_() {
  var semesterName = (getSetting_("Semester name") || "Semester").toString().trim() || "Semester";
  return getOrCreateSubfolder_(DriveApp.getRootFolder(), semesterName);
}

/** "<Semester name>/<Module>/<category>", each level created only if it
 *  doesn't already exist. */
function ensureModuleCategoryFolder_(moduleName, category) {
  var semesterFolder = ensureSemesterDriveFolder_();
  var moduleFolder = getOrCreateSubfolder_(semesterFolder, moduleName);
  return getOrCreateSubfolder_(moduleFolder, category);
}

/**
 * `form` = {module, category, title, topic, priority, notes, fileName,
 * mimeType, base64Data}. base64Data is the file's contents, already
 * base64-encoded client-side (FileReader.readAsDataURL, prefix stripped —
 * see saveDriveUpload_ in Index.html). Uploads the file, then writes a
 * normal "15 Resources" row (same fields api_addResource writes) pointing
 * at it — Drive category / Drive file ID are the only extra columns set.
 */
function api_uploadResourceFile(form) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(LOCK_WAIT_MS)) return fail_(new Error("Workbook is busy — try again in a moment."));
    if (!form) return fail_(new Error("No form data received."));
    var module = requireModule_(form.module);
    var category = VALID_DRIVE_CATEGORIES.indexOf(form.category) !== -1 ? form.category : null;
    if (!category) return fail_(new Error("Choose a category: " + VALID_DRIVE_CATEGORIES.join(", ") + "."));
    var title = requireText_(form.title, "Resource name", 200);
    var fileName = requireText_(form.fileName, "File name", 200);
    if (!form.base64Data) return fail_(new Error("No file data received — choose a file first."));

    var decoded;
    try { decoded = Utilities.base64Decode(form.base64Data); }
    catch (e) { return fail_(new Error("Could not read the file data — try choosing the file again.")); }
    if (decoded.length > MAX_UPLOAD_BYTES) {
      return fail_(new Error("File is too large (" + Math.round(decoded.length / 1024 / 1024) + "MB). Keep uploads under " + Math.round(MAX_UPLOAD_BYTES / 1024 / 1024) + "MB."));
    }

    var blob = Utilities.newBlob(decoded, form.mimeType || "application/octet-stream", fileName);
    var folder = ensureModuleCategoryFolder_(module, category);
    var file = folder.createFile(blob);

    var topic = cleanText_(form.topic, 200);
    var notes = cleanMultiline_(form.notes, 1000);
    var priority = VALID_PRIORITIES.indexOf(form.priority) !== -1 ? form.priority : "Medium";
    // A reasonable Resource type default per category, purely so this
    // upload still shows up sensibly anywhere the Resources screen groups
    // or filters by type — never shown to the user as a choice, since the
    // category itself (a separate, additive field) is the real answer to
    // "what kind of upload is this."
    var resType = category === "Completed tutorial tests" ? "Tutorial"
      : category === "Assessment or exam" ? "Past paper" : "Other";

    var sheet = SpreadsheetApp.getActive().getSheetByName(RESOURCES_SHEET);
    if (!sheet) throw new Error("Sheet not found: " + RESOURCES_SHEET);
    var map = getColMap_(sheet);
    if (!map["Resource ID"]) throw new Error('"Resource ID" column not found — run Resources_v1.2.0_Migration.gs first.');
    if (!map["Drive category"]) throw new Error('"Drive category" column not found — run Drive_Resources_v1.4.0_Migration.gs first.');

    var row = firstBlankRow_(sheet, col_(map, "Title"));
    var newId = nextId_("RES", getColumnValues_(sheet, col_(map, "Resource ID")));
    sheet.getRange(row, col_(map, "Resource ID")).setValue(newId);
    sheet.getRange(row, col_(map, "Module")).setValue(module);
    sheet.getRange(row, col_(map, "Resource type")).setValue(resType);
    sheet.getRange(row, col_(map, "Title")).setValue(title);
    if (topic) sheet.getRange(row, col_(map, "Topic")).setValue(topic);
    sheet.getRange(row, col_(map, "Link or file reference")).setValue(file.getUrl());
    if (notes) sheet.getRange(row, col_(map, "Notes")).setValue(notes);
    sheet.getRange(row, col_(map, "Priority")).setValue(priority);
    sheet.getRange(row, col_(map, "Status")).setValue("Not started");
    sheet.getRange(row, col_(map, "Reviewed")).setValue(false);
    sheet.getRange(row, col_(map, "Drive category")).setValue(category);
    sheet.getRange(row, col_(map, "Drive file ID")).setValue(file.getId());

    logAutomation_("Resource file uploaded", module + " — " + title, "Created", newId + "; " + category + "; " + file.getUrl());
    return ok_({ id: newId, url: file.getUrl() });
  } catch (e) { return fail_(e); } finally { lock.releaseLock(); }
}
