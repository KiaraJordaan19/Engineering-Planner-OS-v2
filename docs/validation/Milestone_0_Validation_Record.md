# Milestone 0 Validation Record

## Record metadata

| Field | Value |
|---|---|
| Project | Engineering Planner OS v1.5.0 |
| Milestone | Milestone 0 — Semester Backup Recovery and Hardening |
| Validation status | GATE A PASSED — GATE B NOT STARTED |
| Tested commit | `66abd493bbec9517411091dc2d6d0be8cb336948` |
| Branch | `feature/v1.5.0-assessment-engine` |
| Tester | Codex — disposable-environment Gate A execution |
| Validation date | 2026-07-17 to 2026-07-18 |
| Disposable workbook identifier | `14sK…LBco` — `EPOS v1.5.0 — M0 DISPOSABLE VALIDATION — DO NOT USE` |
| Apps Script project identifier | `1rGb…F0e9` — newly bound disposable project |
| Apps Script version or deployment identifier | Disposable bound source with M0-LV-05 remediation; `Semester_Backup_v1.4.0.js` SHA-256 `edf4b9ed06c2cc1594055479f100cd013380b32cba073d4b2e1eeb8257ab17b2`; no deployment created |
| Google account used | `kiara…@gmail.com` — full address intentionally withheld from the repository |
| Test environment statement | Disposable Drive folder `EPOS Disposable Validation — Milestone 0` (`1Anv…s-mn`), workbook, and bound Apps Script project were newly created at 2026-07-17T22:18:03+02:00. The project is named `EPOS v1.5.0 M0 Disposable Validation`. The artificial fixture contains the 19 backed-up planner sheets plus infrastructure sheets `24 Semester Backups` and `26 Semester History`. Calendar access is hard-disabled, the project has zero triggers, and no deployment was created. M0-LV-01 through M0-LV-10 passed. M0-LV-05 passed after the recorded implementation and validation-harness remediations. Gate B was not started. Production resources were not opened or modified. |
| Production untouched | Confirmed for Gate A; final confirmation Pending |
| Cleanup completed | Pending |
| Final Milestone 0 decision | PENDING |

## Principal validation table

| Test ID | Test | Classification | Expected result | Actual result | Status | Evidence | Notes |
|---|---|---|---|---|---|---|---|
| M0-LV-01 | Backup creation | Blocking | A backup is created only when all 19 required sheets are present and structurally valid. One JSON Drive file and one matching `24 Semester Backups` metadata row are created with the same backup ID, file ID, semester ID, planner version, sheet set, and fingerprint metadata. | Created control backup `BKP-03a3…31a7a`, Drive file `1Nvm…a_SV`, and metadata row 5. Payload and index both contained 19 sheets; format, semester, planner version, and fingerprint metadata matched. | Pass | Apps Script execution 2026-07-17 22:48:07–22:48:19 SAST; sanitized artefact details are in section 4. | Control backup retained for later authorized gates; final cleanup remains Pending. |
| M0-LV-02 | Valid backup verification | Blocking | Verification of the untouched M0-LV-01 backup returns `ok: true`, `fileFound: true`, `jsonValid: true`, an empty `missingSheets` array, and `checksumValid: true`. | Returned `{"ok":true,"fileFound":true,"jsonValid":true,"missingSheets":[],"checksumValid":true}` for `BKP-03a3…31a7a`. | Pass | Apps Script execution 2026-07-17 22:48:31–22:48:33 SAST. | No additional artefact created; control backup retained. |
| M0-LV-03 | Malformed JSON detection | Blocking | Verification of a controlled test copy whose content is not valid JSON returns `ok: false`, `fileFound: true`, `jsonValid: false`, and a reason identifying invalid JSON. No destructive workflow is authorized. | Content was replaced with `{not valid JSON`. Verification returned `ok:false`, `fileFound:true`, `jsonValid:false`, and identified invalid JSON at position 1. | Pass | Artefact `M0-GA-03-855f…76d3`, file `1zPk…1OV_`, temporary row 6; execution 2026-07-17 22:48:43–22:48:48 SAST. | Cleanup succeeded: file trashed and temporary metadata removed; control unchanged. |
| M0-LV-04 | Altered JSON detection | Blocking | Verification of structurally valid JSON altered after fingerprinting returns `ok: false` with a fingerprint-mismatch reason. No destructive workflow is authorized. | `spreadsheetName` was changed by appending `— ALTERED`; independent metadata retained fingerprint `bfa363a5634f7bcca0c1f0cb15e6a66e88cd7c5d851d56692323041965bcf216`. Verification returned the expected fingerprint-mismatch reason. | Pass | Artefact `M0-GA-04-e144…888f`, file `1qTr…8QXe`, temporary row 6; execution 2026-07-17 22:49:00–22:49:06 SAST. | Cleanup succeeded: file trashed and temporary metadata removed. |
| M0-LV-05 | Deleted Drive file detection | Blocking | Verification of a retrievable but trashed backup returns `ok:false`, `fileFound:false`, and exact reason `Backup Drive file is trashed and unavailable.` No destructive workflow is authorized. | PASS on final retest. Original run failed because trash state was not inspected (`VAL-01`). After implementation remediation, the first retest returned the substantive approved fields but the runner rejected the new reason using an obsolete substring (`VAL-02`). The corrected runner then received `{"ok":false,"fileFound":false,"jsonValid":false,"missingSheets":[],"reason":"Backup Drive file is trashed and unavailable."}`. | Pass | Original failure: 2026-07-17 22:49:34–22:49:43 SAST. First substantive-pass/harness-failure retest: 23:39:41–23:39:48. Final retest: 23:46:51–23:46:59, duration 7,579 ms. Production source checksum `edf4b9ed…17b2`. | Final artefact `M0-GA-05-2ce6…47a4`, file `1Bx_…0LSf`; file remains trashed and metadata row 6 was removed. |
| M0-LV-06 | Wrong backup ID detection | Blocking | Verification fails with `ok: false` when the requested backup ID, payload backup ID, and independent metadata backup ID do not agree. The result identifies a missing row or backup-ID mismatch as appropriate. | A copied payload retained control ID `BKP-03a3…31a7a` while its request and metadata used `M0-GA-06-9dca…8a2a`. Verification returned `ok:false`, `fileFound:true`, `jsonValid:true`, and reason `Backup ID does not match the independent index.` | Pass | Artefact `M0-GA-06-9dca…8a2a`, file `19Vc…Xpqy`, temporary metadata row 6; execution 2026-07-18 00:03:12–00:03:17 SAST, duration 4,586 ms. | Cleanup succeeded: copied file trashed and temporary metadata removed. |
| M0-LV-07 | Wrong spreadsheet ID detection | Blocking | Verification of a backup payload identifying a spreadsheet other than the active disposable workbook returns `ok: false` with the reason `Backup belongs to a different spreadsheet.` | The copied payload used `DISPOSABLE-WRONG-SPREADSHEET-ID`. Verification returned `ok:false`, `fileFound:true`, `jsonValid:true`, and exact reason `Backup belongs to a different spreadsheet.` | Pass | File `1-fi…4ATF`, temporary metadata row 6; execution 2026-07-18 00:03:36–00:03:43 SAST, duration 6,565 ms. | Cleanup succeeded: copied file trashed and temporary metadata removed. |
| M0-LV-08 | Missing required sheet detection | Blocking | Backup creation fails before creating a Drive file when any mandatory sheet is absent. Verification of a crafted payload or index omitting a mandatory sheet returns `ok: false` and identifies a required-boundary mismatch; no destructive workflow is authorized. | With `03 Modules` temporarily absent, creation failed with `Backup was not created because required sheets are missing or empty: 03 Modules`; Drive and metadata counts remained 1. A crafted payload omitting that sheet returned `ok:false`, listed `03 Modules` in `missingSheets`, and identified the required-boundary mismatch. | Pass | Crafted file `19c1…bnoI`, temporary metadata row 6; execution 2026-07-18 00:04:04–00:04:29 SAST, duration 23,570 ms. | Crafted artefacts cleaned successfully; original `03 Modules` fixture restored. |
| M0-LV-09 | Empty required sheet detection | Blocking | A required sheet with no used cells causes backup creation to fail before Drive-file creation. A crafted backup containing an empty `values` array returns `ok: false`, identifies the affected sheet, and cannot authorize destructive work. | With an empty replacement `04 Module Rules`, creation failed with `Backup was not created because required sheets are missing or empty: 04 Module Rules`; Drive and metadata counts remained 1. A crafted payload with an empty `values` array returned `ok:false`, listed `04 Module Rules` in `missingSheets`, and identified an empty or structurally invalid sheet dump. | Pass | Crafted file `1_Ax…tFdb`, temporary metadata row 6; execution 2026-07-18 00:04:45–00:05:02 SAST, duration 15,565 ms. | Crafted artefacts cleaned successfully; original `04 Module Rules` fixture restored. |
| M0-LV-10 | Legacy backup rejection | Blocking | A valid legacy-format backup without `EPOS_BACKUP_FORMAT_2` SHA-256 metadata remains listable but verification returns `ok: false` with the legacy-or-malformed-fingerprint reason. It cannot authorize destructive work. | Legacy artefact `M0-GA-10-4a1b…eb6b` was listable. Verification returned `ok:false`, `fileFound:true`, `jsonValid:true`, and reason `Legacy or malformed backup metadata has no verifiable SHA-256 fingerprint.` | Pass | File `1hvN…NzuZ`, temporary metadata row 6; execution 2026-07-18 00:05:20–00:05:27 SAST, duration 5,777 ms. | Cleanup succeeded: legacy file trashed and temporary metadata removed. |
| M0-LV-11 | Normal delete gate rejects invalid backup | Blocking | The normal delete workflow refuses confirmation when its associated backup fails verification. No active semester data or Calendar event is cleared, and the workflow is not consumed as successfully completed. |  | Pending | Record the workflow token or sanitized reference, invalid backup ID, verification failure, workflow response, and before/after fixture comparison proving no destructive change. | Cleanup: invalidate or remove the disposable workflow record and retain the unchanged fixture for subsequent controlled tests. |
| M0-LV-12 | Normal delete gate accepts valid backup | Blocking | In the disposable environment, the normal delete workflow accepts an unexpired prepared workflow with a valid verified backup and archive only after the exact confirmation phrase. It clears exactly the documented active-sheet targets and consumes the workflow token once. |  | Pending | Record the workflow reference, valid backup ID, archive evidence, confirmation result, consumed status, and before/after inventory. Do not enable external Calendar deletion for this validation. | Cleanup: restore required fixture values for the recovery drill or recreate the disposable fixture; remove all workflow, archive, backup, and Drive test artefacts at final cleanup. |
| M0-LV-13 | Emergency-reset gate rejects invalid backup | Blocking | Emergency reset does not clear active data when fresh backup creation or verification fails. The response identifies the failure and before/after fixture comparison shows no destructive change. |  | Pending | Record the controlled failure, emergency-reset response, relevant backup evidence, and complete before/after fixture comparison. | Cleanup: remove failed test artefacts and restore the disposable fixture to its recorded baseline. |
| M0-LV-14 | Emergency-reset gate accepts valid backup | Blocking | In the disposable environment, emergency reset proceeds only after the exact emergency confirmation phrase and successful creation and verification of a fresh backup. It clears exactly the documented active-sheet targets and returns the emergency backup ID. |  | Pending | Record the response, emergency backup ID, verification output, before/after sheet inventory, and proof that external Calendar deletion was disabled. | Cleanup: retain the valid emergency backup until the recovery drill is complete, then remove all disposable artefacts. |
| M0-LV-15 | Recovery drill | Blocking | Values from at least one selected sheet are manually restored from the verified backup into the disposable workbook and match the recorded original fixture exactly for dimensions, row order, column order, types, blanks, zeroes, booleans, text, numbers, and serialized dates within the documented value-only scope. |  | Pending | Record the source backup and sheet, pre-reset fixture comparison record, restored range, automated or manual comparison output, mismatch count, and screenshots or sanitized logs. | Cleanup: dispose of the restored workbook and all related test files only after evidence and owner review requirements are satisfied. |

## 1. Environment isolation declaration

- [x] The workbook is a disposable copy and is not the production workbook.
- [x] The Apps Script project is dedicated to the disposable workbook or otherwise proven isolated from production.
- [x] The Google account used is recorded above and authorized for this validation.
- [x] Production spreadsheet, Drive files, Apps Script project, deployments, triggers, and Calendar were not accessed or changed.
- [x] External Calendar deletion is disabled for all validation workflows.
- [x] All identifiers and screenshots included as evidence are sanitized where required.

Isolation declaration:

> 

Tester signature or name:  
Date:  

## 2. Required planner and infrastructure fixture inventory

Record the disposable fixture state before any validation mutation.

| # | Required sheet | Present | Last row | Last column | Fixture/reference checksum | Notes |
|---:|---|---|---:|---:|---|---|
| 1 | `02 Settings` | Yes | 17 | 3 | See preparation execution log | Planner dataset |
| 2 | `04 Module Rules` | Yes | 5 | 12 | See preparation execution log | Planner dataset |
| 3 | `03 Modules` | Yes | 5 | 12 | `73180c7dfd6a6863ffc44b4e0bf3b422db9eee7a4d3bdbcf92c84d7f360b404a` | Planner dataset; selected recovery sheet |
| 4 | `17 Archive` | Yes | 13 | 4 | See preparation execution log | Planner dataset |
| 5 | `23 Semester Archive` | Yes | 4 | 13 | See preparation execution log | Planner dataset |
| 6 | `25 Module Templates` | Yes | 4 | 12 | See preparation execution log | Planner dataset |
| 7 | `27 Semester Workflows` | Yes | 4 | 11 | See preparation execution log | Planner dataset |
| 8 | `07 Academic Inbox` | Yes | 5 | 12 | See preparation execution log | Planner dataset |
| 9 | `08 Assignments` | Yes | 5 | 15 | See preparation execution log | Planner dataset |
| 10 | `09 Assessments` | Yes | 5 | 12 | See preparation execution log | Planner dataset |
| 11 | `16 Notes` | Yes | 5 | 4 | See preparation execution log | Planner dataset |
| 12 | `15 Resources` | Yes | 5 | 18 | See preparation execution log | Planner dataset |
| 13 | `12 Study Planner` | Yes | 5 | 18 | See preparation execution log | Planner dataset |
| 14 | `21 Study Tasks` | Yes | 5 | 12 | See preparation execution log | Planner dataset |
| 15 | `13 Revision Tracker` | Yes | 5 | 17 | See preparation execution log | Planner dataset |
| 16 | `11 Marks Tracker` | Yes | 7 | 12 | See preparation execution log | Planner dataset; header row 6 |
| 17 | `10 AF Components` | Yes | 5 | 11 | See preparation execution log | Planner dataset |
| 18 | `19 Attendance` | Yes | 5 | 8 | See preparation execution log | Planner dataset |
| 19 | `20 Automation Log` | Yes | 4 | 5 | See preparation execution log | Planner dataset; header only before tests |
| 20 | `24 Semester Backups` | Yes | 4 | 8 | See preparation execution log | Infrastructure; not part of backed-up semester dataset |
| 21 | `26 Semester History` | Yes | 4 | 13 | See preparation execution log | Infrastructure; not part of backed-up semester dataset |

Fixture inventory completed by:  
Date:  

## 3. Original fixture checksum or comparison record

Describe the comparison method used to establish the original value-level fixture baseline.

| Item | Value |
|---|---|
| Comparison method | SHA-256 of the selected recovery sheet's typed value matrix |
| Canonicalization or export method | JSON cells in row and column order; values discriminated as date, blank, number, boolean, or string; dates serialized with `Date.toISOString()` |
| Whole-fixture checksum, if used |  |
| Selected recovery-drill sheet | `03 Modules` |
| Selected-sheet dimensions | 5 rows × 12 columns |
| Selected-sheet checksum or reference | `73180c7dfd6a6863ffc44b4e0bf3b422db9eee7a4d3bdbcf92c84d7f360b404a` |
| Baseline evidence location | Apps Script preparation execution log for `m0PrepareDisposableFixture` |
| Recorded by | Codex — environment preparation only |
| Recorded at | 2026-07-17T22:39:08+02:00 |

## 4. Backup artefact register

| Artefact ID | Test ID | Backup ID | Drive file ID | Metadata row/reference | Format | Fingerprint | Created at | Final disposition |
|---|---|---|---|---|---|---|---|---|
| `M0-CONTROL-01` | M0-LV-01 / M0-LV-02 | `BKP-03a3…31a7a` | `1Nvm…a_SV` | Row 5 | `EPOS_BACKUP_FORMAT_2` | `bfa363a5634f7bcca0c1f0cb15e6a66e88cd7c5d851d56692323041965bcf216` | 2026-07-17 22:48:08 SAST | Retained for later authorized gates; Pending |

## 5. Corruption-test artefact register

Never corrupt the valid control backup. Create an isolated test copy for every corruption scenario.

| Artefact ID | Test ID | Source backup ID | Test file ID | Mutation | Expected failure | Observed failure | Final disposition |
|---|---|---|---|---|---|---|---|
| `M0-GA-03-855f…76d3` | M0-LV-03 | `BKP-03a3…31a7a` | `1zPk…1OV_` | Entire content replaced with malformed JSON | Invalid-JSON rejection | Invalid-JSON rejection | File trashed; metadata removed |
| `M0-GA-04-e144…888f` | M0-LV-04 | `BKP-03a3…31a7a` | `1qTr…8QXe` | Appended `— ALTERED` to `spreadsheetName` | Fingerprint mismatch | Fingerprint mismatch | File trashed; metadata removed |
| Identifier not emitted before assertion | M0-LV-05 | `BKP-03a3…31a7a` | Identifier not emitted before assertion | Valid copy verified, then `setTrashed(true)` | Drive-file-not-found rejection | Trashed file remained resolvable | File remains trashed; metadata removed |
| Identifier not emitted before assertion | M0-LV-05 first retest | `BKP-03a3…31a7a` | Identifier not emitted before assertion | Valid copy verified, trashed, then checked by remediated implementation | Exact trashed-and-unavailable rejection | Substantive fields passed; obsolete harness reason assertion failed | File remains trashed; metadata removed |
| `M0-GA-05-2ce6…47a4` | M0-LV-05 final retest | `BKP-03a3…31a7a` | `1Bx_…0LSf` | Valid copy verified, then trashed | Exact trashed-and-unavailable rejection | Exact contract matched | File remains trashed; metadata removed |
| `M0-GA-06-9dca…8a2a` | M0-LV-06 | `BKP-03a3…31a7a` | `19Vc…Xpqy` | Request/index ID differed from payload ID | Backup-ID mismatch | Backup-ID mismatch | File trashed; metadata removed |
| M0-LV-07 crafted copy | M0-LV-07 | `BKP-03a3…31a7a` | `1-fi…4ATF` | Payload spreadsheet ID replaced and fingerprint recomputed | Wrong-spreadsheet rejection | Wrong-spreadsheet rejection | File trashed; metadata removed |
| M0-LV-08 crafted copy | M0-LV-08 | `BKP-03a3…31a7a` | `19c1…bnoI` | `03 Modules` omitted from payload | Required-boundary mismatch | Required-boundary mismatch | File trashed; metadata removed; fixture restored |
| M0-LV-09 crafted copy | M0-LV-09 | `BKP-03a3…31a7a` | `1_Ax…tFdb` | `04 Module Rules.values` replaced with an empty array | Empty-sheet rejection | Empty-sheet rejection | File trashed; metadata removed; fixture restored |
| `M0-GA-10-4a1b…eb6b` | M0-LV-10 | Not applicable — legacy artefact | `1hvN…NzuZ` | Format/fingerprint metadata omitted; legacy Notes used | Legacy-fingerprint rejection | Legacy-fingerprint rejection | File trashed; metadata removed |

## 6. Destructive-workflow safety record

| Test ID | Workflow type | Backup ID | Verification status | Confirmation used | Calendar deletion enabled | Expected affected sheets | Actual affected sheets | Token/status outcome | Safety result |
|---|---|---|---|---|---|---|---|---|---|
| M0-LV-11 | Normal delete — rejection |  |  |  | No | None |  |  | Pending |
| M0-LV-12 | Normal delete — acceptance |  |  |  | No | `08 Assignments`; `09 Assessments`; `12 Study Planner`; `21 Study Tasks`; `13 Revision Tracker`; `11 Marks Tracker`; `10 AF Components`; `19 Attendance`; `07 Academic Inbox`; `20 Automation Log` |  |  | Pending |
| M0-LV-13 | Emergency reset — rejection |  |  |  | No | None |  |  | Pending |
| M0-LV-14 | Emergency reset — acceptance |  |  |  | No | `08 Assignments`; `09 Assessments`; `12 Study Planner`; `21 Study Tasks`; `13 Revision Tracker`; `11 Marks Tracker`; `10 AF Components`; `19 Attendance`; `07 Academic Inbox`; `20 Automation Log` |  |  | Pending |

## 7. Recovery comparison

| Field | Value |
|---|---|
| Test ID | M0-LV-15 |
| Backup ID |  |
| Drive file ID |  |
| Restored sheet |  |
| Original dimensions |  |
| Restored dimensions |  |
| Comparison method |  |
| Cells compared |  |
| Mismatch count |  |
| Type mismatches |  |
| Blank/zero/boolean checks |  |
| Date serialization checks |  |
| Comparison evidence |  |
| Result | Pending |

## 8. Runtime observations

| Observation ID | Test ID | Timestamp | Duration | Severity | Observation | Follow-up required |
|---|---|---|---:|---|---|---|
| `RUN-01` | M0-LV-01 | 2026-07-17 22:48:07 SAST | 11216 | Info | Backup creation completed successfully within Apps Script runtime limits. | No |
| `RUN-02` | M0-LV-02 | 2026-07-17 22:48:31 SAST | 988 | Info | Intact control verification completed successfully. | No |
| `RUN-03` | M0-LV-03 | 2026-07-17 22:48:43 SAST | 4515 | Info | Malformed JSON was rejected and cleanup succeeded. | No |
| `RUN-04` | M0-LV-04 | 2026-07-17 22:49:00 SAST | 4952 | Info | Altered JSON was rejected by fingerprint verification and cleanup succeeded. | No |
| `RUN-05` | M0-LV-05 | 2026-07-17 22:49:34 SAST | Approximately 9000 | Blocking | `DriveApp.getFileById()` continued to resolve the file after `setTrashed(true)`; verification did not classify it as missing. | Yes — implementation decision/correction required before rerun |
| `RUN-05-R1` | M0-LV-05 | 2026-07-17 23:39:41 SAST | Approximately 7000 | Harness defect | Remediated implementation returned `ok:false` and `fileFound:false`; runner rejected the precise new reason because it still required legacy text `Drive file not found`. | Resolved by exact frozen-contract assertion; incident `VAL-02` |
| `RUN-05-R2` | M0-LV-05 | 2026-07-17 23:46:51 SAST | 7579 | Info | Final retest returned `ok:false`, `fileFound:false`, and exact reason `Backup Drive file is trashed and unavailable.` | No; M0-LV-05 passed |
| `RUN-06` | M0-LV-06 | 2026-07-18 00:03:12 SAST | 4586 | Info | Payload-to-index backup-ID mismatch was rejected and cleanup succeeded. | No |
| `RUN-07` | M0-LV-07 | 2026-07-18 00:03:36 SAST | 6565 | Info | Wrong spreadsheet identity was rejected with the exact expected reason and cleanup succeeded. | No |
| `RUN-08` | M0-LV-08 | 2026-07-18 00:04:04 SAST | 23570 | Info | Missing-sheet creation failed without new artefacts; crafted payload was rejected; fixture restoration succeeded. | No |
| `RUN-09` | M0-LV-09 | 2026-07-18 00:04:45 SAST | 15565 | Info | Empty-sheet creation failed without new artefacts; crafted payload was rejected; fixture restoration succeeded. | No |
| `RUN-10` | M0-LV-10 | 2026-07-18 00:05:20 SAST | 5777 | Info | Legacy backup remained listable but could not pass integrity verification; cleanup succeeded. | No |

Record Apps Script execution time, Drive latency, authorization prompts, quotas, warnings, unexpected logs, and any runtime behavior that differs from local mocks.

## 9. Known limitations

The current backup preserves cell values within each required sheet's used range. It does **not** constitute full automated recovery of:

- formulas as formulas;
- formatting;
- data validation;
- comments or notes;
- protections;
- merged cells;
- drawings;
- named ranges;
- external Calendar events.

Additional limitations or observations:

- Recovery is a manual drill in Milestone 0; no automated restore workflow is authorized.
- A fingerprint detects content inconsistency against the independent metadata record but is not a cryptographic authenticity guarantee against an actor able to replace both records.
- Legacy backups without the required format and fingerprint metadata cannot authorize destructive workflows.
- Checksum metadata is temporarily stored in the existing `Notes` column pending a future versioned schema migration to a dedicated field.

## 10. Cleanup register

| Cleanup ID | Related test/artefact | Required action | Responsible person | Status | Completed at | Evidence/notes |
|---|---|---|---|---|---|---|
| `CLEAN-03` | M0-LV-03 malformed artefact | Trash file and remove temporary metadata | Codex | Complete | 2026-07-17 22:48:48 SAST | Runner reported `SUCCEEDED`; row removed |
| `CLEAN-04` | M0-LV-04 altered artefact | Trash file and remove temporary metadata | Codex | Complete | 2026-07-17 22:49:05 SAST | Runner reported `SUCCEEDED`; row removed |
| `CLEAN-05` | M0-LV-05 trashed-file artefact | Retain in trash pending failure review; remove temporary metadata | Codex | Partial | 2026-07-17 22:49:43 SAST | File remains trashed; row removed; identifier was not logged before assertion |
| `CLEAN-05-R2` | M0-LV-05 final retest artefact `M0-GA-05-2ce6…47a4` | Retain in trash pending final cleanup; remove temporary metadata | Codex | Partial | 2026-07-17 23:46:59 SAST | File `1Bx_…0LSf` remains trashed; metadata row 6 removed |
| `CLEAN-06` | M0-LV-06 mismatch artefact `M0-GA-06-9dca…8a2a` | Trash copied file and remove temporary metadata | Codex | Complete | 2026-07-18 00:03:17 SAST | Runner reported `SUCCEEDED`; metadata row 6 removed |
| `CLEAN-07` | M0-LV-07 wrong-spreadsheet artefact | Trash copied file and remove temporary metadata | Codex | Complete | 2026-07-18 00:03:42 SAST | Runner reported `SUCCEEDED`; metadata row 6 removed |
| `CLEAN-08` | M0-LV-08 missing-sheet artefact and fixture mutation | Trash crafted file, remove temporary metadata, and restore `03 Modules` | Codex | Complete | 2026-07-18 00:04:29 SAST | Runner reported `SUCCEEDED`; fixture restoration confirmed |
| `CLEAN-09` | M0-LV-09 empty-sheet artefact and fixture mutation | Trash crafted file, remove temporary metadata, and restore `04 Module Rules` | Codex | Complete | 2026-07-18 00:05:01 SAST | Runner reported `SUCCEEDED`; fixture restoration confirmed |
| `CLEAN-10` | M0-LV-10 legacy artefact `M0-GA-10-4a1b…eb6b` | Trash legacy file and remove temporary metadata | Codex | Complete | 2026-07-18 00:05:26 SAST | Runner reported `SUCCEEDED`; metadata row 6 removed |
| `CLEAN-CONTROL` | M0-LV-01 control backup | Retain through later authorized gates and recovery drill | Owner/Codex | Pending |  | Control file and metadata retained intentionally |

Cleanup completion checklist:

- [ ] Corrupted and malformed Drive test files removed or trashed.
- [ ] Valid backup test files disposed of after the recovery drill and owner review.
- [ ] Disposable metadata and workflow rows removed or the entire disposable workbook disposed of.
- [ ] Disposable Apps Script deployments and triggers removed where applicable.
- [ ] Disposable workbook removed or archived according to owner instruction.
- [ ] No test Calendar events were created or deleted.
- [ ] Cleanup evidence recorded.

Cleanup completed by:  
Cleanup completed at:  
Overall cleanup status: Pending

## 11. Production non-impact declaration

Confirm after validation:

- [x] Production workbook was not opened, read, or modified for Gate A validation.
- [x] Production Apps Script source, versions, deployments, properties, and triggers were not modified.
- [x] Production Drive backup files and folders were not modified; Gate A first verified that its backup folder was a child of the disposable folder.
- [x] Production Calendar and events were not accessed or modified.
- [x] All Gate A mutations occurred only in the identified disposable environment.

Declaration:

> Gate A execution was confined to disposable workbook `14sK…LBco`, disposable Apps Script project `1rGb…F0e9`, and a backup folder verified as a child of disposable folder `1Anv…s-mn`. No production resource was selected or changed.

Confirmed by: Codex — Gate A execution

Date: 2026-07-18

Status: Confirmed for Gate A; final Milestone 0 confirmation remains pending

## 12. Deviations and incidents

| ID | Test ID | Type | Description | Impact | Containment | Resolution | Owner decision |
|---|---|---|---|---|---|---|---|
| ENV-01 | Environment preparation | Resolved specification/dependency conflict | The original fixture listed 19 planner sheets but `semCreateBackup_` also requires the `24 Semester Backups` index sheet through `SEM_BACKUPS_SHEET`. | The original 19-sheet fixture could not execute backup creation. | Preparation stopped before source transfer, fixture creation, or function execution. | Owner authorized `24 Semester Backups` as a twentieth infrastructure sheet. | Approved and resolved |
| ENV-02 | Environment preparation | Resolved expanded-boundary dependency conflict | The owner-approved `Semester_Archive_v1.4.0.js` declares `SEM_HISTORY_SHEET = '26 Semester History'`, and `semCreateArchive_` always calls `semUpdateHistory_`, which throws when that sheet is absent. | The normal delete acceptance test could not complete safely in the previously authorized 20-sheet fixture. | Preparation stopped before Apps Script source transfer, fixture creation, or function execution. | Owner authorized `26 Semester History` as a twenty-first infrastructure sheet and replaced fixed-count doctrine with intent-based dependency inclusion. | Approved and resolved |
| ENV-03 | Environment preparation | Resolved application/business dependency | `semReportBuild_`, which is mandatory for normal delete workflow preparation, directly calls `api_getPlannerData()` and `computeAcademicIntelligence_()`. Both are implemented in repository `API.js`, not in the original five-file boundary, and neither may be replaced by a validation stub. | Loading only the original five files would leave the normal delete workflow non-executable or substitute unauthorized business logic. | Completed the full static dependency pass and stopped until owner authorization was received. | Owner authorized `API.js` from the same approved feature-branch HEAD. Its uploaded SHA-256 matches the repository source. Calendar access remains hard-disabled in the separate validation support file. | Approved and resolved |
| VAL-01 | M0-LV-05 | Historical implementation defect | A file moved to trash with `file.setTrashed(true)` remained retrievable through `DriveApp.getFileById()`. The original `semVerifyBackup_()` did not inspect trash state. | Initial deleted/trashed-file detection failed and Gate A stopped. | The artefact remained in disposable trash, metadata was removed, and production remained untouched. | Owner-approved minimal remediation added an immediate fail-closed `file.isTrashed()` inspection. Local suite passed 14/14; disposable remote source checksum is `edf4b9ed06c2cc1594055479f100cd013380b32cba073d4b2e1eeb8257ab17b2`; final live retest passed. | Approved and resolved; history retained |
| VAL-02 | M0-LV-05 first retest | Historical validation-harness defect | The first remediated live retest satisfied `ok:false` and `fileFound:false`, but the runner still required obsolete reason substring `Drive file not found` rather than the frozen precise trashed-file reason. | The runner reported failure despite substantive implementation compliance. No later tests ran. | Execution stopped; no production change followed. | Owner classified this as a harness defect. The validation-only assertion was changed to exact-match `Backup Drive file is trashed and unavailable.` The final retest passed. | Approved and resolved; history retained |

If there were no deviations or incidents, record `None` explicitly before final assessment.

## 13. Final pass/fail assessment

**Decision rule:** Milestone 0 may be marked PASS only when all Blocking tests pass, required evidence is recorded, cleanup is complete, and production non-impact is confirmed.

| Assessment field | Value |
|---|---|
| Blocking tests passed | Gate A: M0-LV-01 through M0-LV-10 passed; M0-LV-11 through M0-LV-15 Pending |
| Required evidence complete | Gate A evidence recorded; full milestone evidence incomplete |
| Cleanup complete | Pending |
| Production non-impact confirmed | Confirmed for Gate A; final confirmation Pending |
| Unresolved deviations/incidents | None for completed tests; historical `VAL-01` and `VAL-02` resolved and retained |
| Final Milestone 0 decision | PENDING |
| Decision rationale | All Gate A backup-integrity tests passed. Gate B was not started; final Milestone 0 decision remains pending. |
| Assessed by | Codex — Gate A execution |
| Assessment date | 2026-07-18 |

## 14. Validation sign-off summary

Complete this summary only after the detailed record and all required evidence have been reviewed.

| Sign-off field | Value |
|---|---|
| Blocking tests passed | 10/15 passed; 5 not completed |
| Evidence package complete | Complete through Gate A; Gate B and recovery evidence Pending |
| Deviations and incidents resolved or accepted | `VAL-01` and `VAL-02` resolved; historical entries retained |
| Cleanup verified complete | Pending |
| Production non-impact verified | Confirmed for Gate A; final confirmation Pending |
| Tester sign-off |  |
| Independent reviewer sign-off |  |
| Owner sign-off |  |
| Final Milestone 0 decision | PENDING |
| Sign-off date |  |

This summary is not a substitute for the evidence in the principal validation table and supporting registers. A `PASS` decision is valid only when the decision rule in section 13 is satisfied and owner approval is recorded in section 15.

## 15. Owner approval

| Field | Value |
|---|---|
| Owner name |  |
| Decision | Pending |
| Approved Milestone 0 result |  |
| Conditions or follow-up actions |  |
| Approval date |  |
| Signature or recorded approval reference |  |

Owner approval of this record does not authorize Milestone 1 unless that authorization is granted separately and explicitly.
