# Milestone 0 Validation Record

## Record metadata

| Field | Value |
|---|---|
| Project | Engineering Planner OS v1.5.0 |
| Milestone | Milestone 0 — Semester Backup Recovery and Hardening |
| Validation status | NOT STARTED |
| Tested commit | `66abd493bbec9517411091dc2d6d0be8cb336948` |
| Branch | `feature/v1.5.0-assessment-engine` |
| Tester |  |
| Validation date |  |
| Disposable workbook identifier |  |
| Apps Script project identifier |  |
| Apps Script version or deployment identifier |  |
| Google account used |  |
| Test environment statement | Validation must be performed only against an isolated disposable workbook and its dedicated Apps Script project. Production resources must not be used. |
| Production untouched | Pending confirmation |
| Cleanup completed | Pending |
| Final Milestone 0 decision | PENDING |

## Principal validation table

| Test ID | Test | Classification | Expected result | Actual result | Status | Evidence | Notes |
|---|---|---|---|---|---|---|---|
| M0-LV-01 | Backup creation | Blocking | A backup is created only when all 19 required sheets are present and structurally valid. One JSON Drive file and one matching `24 Semester Backups` metadata row are created with the same backup ID, file ID, semester ID, planner version, sheet set, and fingerprint metadata. |  | Pending | Record the backup ID, Drive file ID, sanitized file URL, metadata-row reference, creation response, and screenshots or logs showing the matching artefacts. | Cleanup: delete or trash the test Drive file and remove or retire the disposable metadata row during final environment cleanup. |
| M0-LV-02 | Valid backup verification | Blocking | Verification of the untouched M0-LV-01 backup returns `ok: true`, `fileFound: true`, `jsonValid: true`, an empty `missingSheets` array, and `checksumValid: true`. |  | Pending | Record the complete sanitized verification output and identify the backup ID verified. | Cleanup: no additional artefact expected; retain evidence until owner review, then clean the disposable environment. |
| M0-LV-03 | Malformed JSON detection | Blocking | Verification of a controlled test copy whose content is not valid JSON returns `ok: false`, `fileFound: true`, `jsonValid: false`, and a reason identifying invalid JSON. No destructive workflow is authorized. |  | Pending | Record the corruption artefact ID, exact mutation method, complete sanitized verification output, and error message. | Cleanup: trash the malformed test file and remove or retire its disposable metadata row. Never alter the valid control backup. |
| M0-LV-04 | Altered JSON detection | Blocking | Verification of structurally valid JSON altered after fingerprinting returns `ok: false` with a fingerprint-mismatch reason. No destructive workflow is authorized. |  | Pending | Record the corruption artefact ID, changed field and before/after values, unchanged expected fingerprint, and complete sanitized verification output. | Cleanup: trash the altered test file and remove or retire its disposable metadata row. |
| M0-LV-05 | Deleted Drive file detection | Blocking | Verification of a disposable metadata row whose test Drive file has been trashed or deleted returns `ok: false`, `fileFound: false`, and a Drive-file-not-found reason. No destructive workflow is authorized. |  | Pending | Record the backup ID, file ID, deletion or trash evidence, complete sanitized verification output, and error message. | Cleanup: permanently remove or retain the already-trashed test artefact according to the disposable-environment cleanup plan; remove or retire its metadata row. |
| M0-LV-06 | Wrong backup ID detection | Blocking | Verification fails with `ok: false` when the requested backup ID, payload backup ID, and independent metadata backup ID do not agree. The result identifies a missing row or backup-ID mismatch as appropriate. |  | Pending | Record the requested ID, metadata ID, payload ID, mutation method, and complete sanitized verification output. | Cleanup: remove or retire the mismatched disposable metadata row and trash its test file. |
| M0-LV-07 | Wrong spreadsheet ID detection | Blocking | Verification of a backup payload identifying a spreadsheet other than the active disposable workbook returns `ok: false` with the reason `Backup belongs to a different spreadsheet.` |  | Pending | Record sanitized expected and altered spreadsheet identifiers, mutation method, and complete verification output. | Cleanup: trash the altered test file and remove or retire its disposable metadata row. |
| M0-LV-08 | Missing required sheet detection | Blocking | Backup creation fails before creating a Drive file when any mandatory sheet is absent. Verification of a crafted payload or index omitting a mandatory sheet returns `ok: false` and identifies a required-boundary mismatch; no destructive workflow is authorized. |  | Pending | Record the omitted sheet, creation or verification output, Drive file count before/after, metadata-row count before/after, and proof that no unintended artefact was created. | Cleanup: restore the missing sheet in the disposable fixture and remove any deliberately crafted test artefacts. |
| M0-LV-09 | Empty required sheet detection | Blocking | A required sheet with no used cells causes backup creation to fail before Drive-file creation. A crafted backup containing an empty `values` array returns `ok: false`, identifies the affected sheet, and cannot authorize destructive work. |  | Pending | Record the emptied sheet, controlled fixture state, creation or verification output, and Drive/metadata counts before and after. | Cleanup: restore the sheet from the original fixture and remove any crafted test artefacts. |
| M0-LV-10 | Legacy backup rejection | Blocking | A valid legacy-format backup without `EPOS_BACKUP_FORMAT_2` SHA-256 metadata remains listable but verification returns `ok: false` with the legacy-or-malformed-fingerprint reason. It cannot authorize destructive work. |  | Pending | Record the legacy backup ID, legacy Notes value, listing evidence, complete verification output, and rejection reason. | Cleanup: remove or retire the disposable legacy metadata row and trash its test file. |
| M0-LV-11 | Normal delete gate rejects invalid backup | Blocking | The normal delete workflow refuses confirmation when its associated backup fails verification. No active semester data or Calendar event is cleared, and the workflow is not consumed as successfully completed. |  | Pending | Record the workflow token or sanitized reference, invalid backup ID, verification failure, workflow response, and before/after fixture comparison proving no destructive change. | Cleanup: invalidate or remove the disposable workflow record and retain the unchanged fixture for subsequent controlled tests. |
| M0-LV-12 | Normal delete gate accepts valid backup | Blocking | In the disposable environment, the normal delete workflow accepts an unexpired prepared workflow with a valid verified backup and archive only after the exact confirmation phrase. It clears exactly the documented active-sheet targets and consumes the workflow token once. |  | Pending | Record the workflow reference, valid backup ID, archive evidence, confirmation result, consumed status, and before/after inventory. Do not enable external Calendar deletion for this validation. | Cleanup: restore required fixture values for the recovery drill or recreate the disposable fixture; remove all workflow, archive, backup, and Drive test artefacts at final cleanup. |
| M0-LV-13 | Emergency-reset gate rejects invalid backup | Blocking | Emergency reset does not clear active data when fresh backup creation or verification fails. The response identifies the failure and before/after fixture comparison shows no destructive change. |  | Pending | Record the controlled failure, emergency-reset response, relevant backup evidence, and complete before/after fixture comparison. | Cleanup: remove failed test artefacts and restore the disposable fixture to its recorded baseline. |
| M0-LV-14 | Emergency-reset gate accepts valid backup | Blocking | In the disposable environment, emergency reset proceeds only after the exact emergency confirmation phrase and successful creation and verification of a fresh backup. It clears exactly the documented active-sheet targets and returns the emergency backup ID. |  | Pending | Record the response, emergency backup ID, verification output, before/after sheet inventory, and proof that external Calendar deletion was disabled. | Cleanup: retain the valid emergency backup until the recovery drill is complete, then remove all disposable artefacts. |
| M0-LV-15 | Recovery drill | Blocking | Values from at least one selected sheet are manually restored from the verified backup into the disposable workbook and match the recorded original fixture exactly for dimensions, row order, column order, types, blanks, zeroes, booleans, text, numbers, and serialized dates within the documented value-only scope. |  | Pending | Record the source backup and sheet, pre-reset fixture comparison record, restored range, automated or manual comparison output, mismatch count, and screenshots or sanitized logs. | Cleanup: dispose of the restored workbook and all related test files only after evidence and owner review requirements are satisfied. |

## 1. Environment isolation declaration

- [ ] The workbook is a disposable copy and is not the production workbook.
- [ ] The Apps Script project is dedicated to the disposable workbook or otherwise proven isolated from production.
- [ ] The Google account used is recorded above and authorized for this validation.
- [ ] Production spreadsheet, Drive files, Apps Script project, deployments, triggers, and Calendar were not accessed or changed.
- [ ] External Calendar deletion is disabled for all validation workflows.
- [ ] All identifiers and screenshots included as evidence are sanitized where required.

Isolation declaration:

> 

Tester signature or name:  
Date:  

## 2. Required 19-sheet fixture inventory

Record the disposable fixture state before any validation mutation.

| # | Required sheet | Present | Last row | Last column | Fixture/reference checksum | Notes |
|---:|---|---|---:|---:|---|---|
| 1 | `02 Settings` | Pending |  |  |  |  |
| 2 | `04 Module Rules` | Pending |  |  |  |  |
| 3 | `03 Modules` | Pending |  |  |  |  |
| 4 | `17 Archive` | Pending |  |  |  |  |
| 5 | `23 Semester Archive` | Pending |  |  |  |  |
| 6 | `25 Module Templates` | Pending |  |  |  |  |
| 7 | `27 Semester Workflows` | Pending |  |  |  |  |
| 8 | `07 Academic Inbox` | Pending |  |  |  |  |
| 9 | `08 Assignments` | Pending |  |  |  |  |
| 10 | `09 Assessments` | Pending |  |  |  |  |
| 11 | `16 Notes` | Pending |  |  |  |  |
| 12 | `15 Resources` | Pending |  |  |  |  |
| 13 | `12 Study Planner` | Pending |  |  |  |  |
| 14 | `21 Study Tasks` | Pending |  |  |  |  |
| 15 | `13 Revision Tracker` | Pending |  |  |  |  |
| 16 | `11 Marks Tracker` | Pending |  |  |  |  |
| 17 | `10 AF Components` | Pending |  |  |  |  |
| 18 | `19 Attendance` | Pending |  |  |  |  |
| 19 | `20 Automation Log` | Pending |  |  |  |  |

Fixture inventory completed by:  
Date:  

## 3. Original fixture checksum or comparison record

Describe the comparison method used to establish the original value-level fixture baseline.

| Item | Value |
|---|---|
| Comparison method |  |
| Canonicalization or export method |  |
| Whole-fixture checksum, if used |  |
| Selected recovery-drill sheet |  |
| Selected-sheet dimensions |  |
| Selected-sheet checksum or reference |  |
| Baseline evidence location |  |
| Recorded by |  |
| Recorded at |  |

## 4. Backup artefact register

| Artefact ID | Test ID | Backup ID | Drive file ID | Metadata row/reference | Format | Fingerprint | Created at | Final disposition |
|---|---|---|---|---|---|---|---|---|
|  |  |  |  |  |  |  |  | Pending |

## 5. Corruption-test artefact register

Never corrupt the valid control backup. Create an isolated test copy for every corruption scenario.

| Artefact ID | Test ID | Source backup ID | Test file ID | Mutation | Expected failure | Observed failure | Final disposition |
|---|---|---|---|---|---|---|---|
|  |  |  |  |  |  |  | Pending |

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
|  |  |  |  |  |  |  |

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
|  |  |  |  | Pending |  |  |

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

- [ ] Production workbook was not opened, read, or modified for validation.
- [ ] Production Apps Script source, versions, deployments, properties, and triggers were not modified.
- [ ] Production Drive backup files and folders were not modified.
- [ ] Production Calendar and events were not accessed or modified.
- [ ] All validation mutations occurred only in the identified disposable environment.

Declaration:

> 

Confirmed by:  
Date:  
Status: Pending confirmation

## 12. Deviations and incidents

| ID | Test ID | Type | Description | Impact | Containment | Resolution | Owner decision |
|---|---|---|---|---|---|---|---|
|  |  |  |  |  |  |  |  |

If there were no deviations or incidents, record `None` explicitly before final assessment.

## 13. Final pass/fail assessment

**Decision rule:** Milestone 0 may be marked PASS only when all Blocking tests pass, required evidence is recorded, cleanup is complete, and production non-impact is confirmed.

| Assessment field | Value |
|---|---|
| Blocking tests passed | Pending |
| Required evidence complete | Pending |
| Cleanup complete | Pending |
| Production non-impact confirmed | Pending |
| Unresolved deviations/incidents | Pending |
| Final Milestone 0 decision | PENDING |
| Decision rationale |  |
| Assessed by |  |
| Assessment date |  |

## 14. Validation sign-off summary

Complete this summary only after the detailed record and all required evidence have been reviewed.

| Sign-off field | Value |
|---|---|
| Blocking tests passed | Pending (0/15 completed) |
| Evidence package complete | Pending |
| Deviations and incidents resolved or accepted | Pending |
| Cleanup verified complete | Pending |
| Production non-impact verified | Pending confirmation |
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
