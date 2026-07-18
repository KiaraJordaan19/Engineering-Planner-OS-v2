# Engineering Planner OS v1.5.0

## M0-LV-15 Recovery Drill Runner Specification

| Field | Value |
|---|---|
| Project | Engineering Planner OS v1.5.0 |
| Milestone | Milestone 0 — Semester Backup Recovery and Hardening |
| Validation test | M0-LV-15 — Recovery drill |
| Document type | Normative validation-runner design specification |
| Status | EXECUTED — PASS; FROZEN FOR RELEASE CHECKPOINT REVIEW |
| Production scope change | None |

## 1. Purpose

This specification defines the validation-only runner for M0-LV-15. The runner SHALL prove, in the approved disposable environment, that the existing backup format can recover the value matrix of the approved `03 Modules` fixture from the retained M0-LV-14 emergency backup.

The runner SHALL be orchestration and evidence-collection code only. It SHALL NOT implement, expose, or imply a production restore, rollback, or disaster-recovery capability.

Successful execution SHALL establish only that the selected sheet's backed-up cell values can be reconstructed within the value-only limits in this specification. It SHALL NOT establish full workbook recovery or production disaster recovery.

## 2. Scope

### 2.1 Included scope

The runner SHALL:

1. execute only against the approved disposable workbook and disposable Apps Script project;
2. use only the retained verified emergency backup created by the successful M0-LV-14 execution;
3. recover only the sheet named exactly `03 Modules`;
4. read only the `03 Modules` dump from the bound backup payload;
5. perform value-only restoration;
6. preserve dimensions, row order, column order, and supported value types;
7. compare the restored matrix with the approved fixture baseline;
8. prove that non-target sheets and unrelated infrastructure remain invariant;
9. capture the evidence required by this specification; and
10. return one permitted final classification.

### 2.2 Excluded scope

The runner SHALL NOT restore or claim to restore:

- any sheet other than `03 Modules`;
- formulas as formulas;
- formatting;
- conditional formatting;
- merged cells;
- notes;
- comments;
- protections;
- drawings;
- images;
- named ranges;
- filters;
- hidden rows;
- hidden columns;
- data-validation rules;
- Calendar events;
- Drive artefacts;
- workflow history;
- module templates; or
- any production workbook content.

The runner SHALL NOT modify Gate A or Gate B behavior, evidence, status, or artefacts except where the retained M0-LV-14 backup is read as the authorized recovery source.

## 3. Normative terminology

The terms `MUST`, `MUST NOT`, `SHALL`, `SHALL NOT`, `SHOULD`, and `MAY` are normative.

For this specification:

- **Bound backup** means the exact M0-LV-14 emergency backup whose backup ID, Drive file ID, metadata row, payload identity, and checksum have been proven to agree.
- **Approved baseline** means the recorded typed-value fixture for `03 Modules`, including its sheet identity, dimensions, ordered matrix, and checksum.
- **Corrupted state** means the controlled, test-owned mutation of `03 Modules` used to prove that recovery changes the selected sheet from a non-baseline state back to the approved baseline.
- **Non-target state** means every workbook sheet, Setting, infrastructure record, Automation Log row, and Drive artefact outside the authorized selected-sheet mutation and documented test-owned audit suffix.
- **Typed value** means a canonical value discriminated as blank, Boolean, number, string, or ISO date under section 9.

## 4. Runner boundary

### 4.1 Callable entry point

The implementation SHALL expose exactly one validation-only callable entry point named `m0RunRecoveryDrill15`. It SHALL accept no parameters.

The entry point SHALL:

- require no UI prompts, editor selection, toast interaction, or manual cell entry;
- derive the retained backup through the deterministic identity rule in section 6.3;
- reject arbitrary sheet names;
- return a structured validation result containing the fields in section 12; and
- remain absent from production APIs and user-facing application workflows.

### 4.2 Permitted responsibilities

The runner MAY perform only:

1. safety-gate evaluation;
2. approved fixture and ownership discovery;
3. public production backup verification;
4. read-only retrieval of the exact bound backup file for validation;
5. independent binding and checksum checks;
6. controlled corruption of `03 Modules`;
7. value-only restoration of `03 Modules` from the bound payload;
8. evidence capture and assertions;
9. identity-scoped cleanup; and
10. report generation.

### 4.3 Prohibited responsibilities

The runner SHALL NOT:

- add or call a production restore API;
- call a private production verification helper as a substitute for `api_semVerifyBackup()`;
- duplicate or weaken production backup-verification rules;
- expose the backup payload through a production or public API;
- restore a sheet selected by caller input;
- write backup contents to any sheet other than the identity-proven disposable `03 Modules` sheet;
- replace infrastructure sheets wholesale;
- manufacture or alter lifecycle workflow records;
- alter the retained backup payload, its metadata row, or its checksum metadata;
- perform Calendar access or Calendar mutation;
- create triggers or deployments; or
- continue after a failed required assertion.

## 5. Required inputs

Before execution, the runner SHALL possess or derive the following:

| Input | Requirement |
|---|---|
| Disposable workbook identity | MUST equal the approved disposable workbook identity |
| Disposable Apps Script identity | MUST equal the approved disposable project identity |
| M0-LV-14 backup ID | MUST be the retained emergency backup ID recorded by M0-LV-14 |
| Drive file ID | MUST equal the file ID in the unique metadata row for that backup ID |
| Semester ID | MUST agree across current fixture, metadata, and payload |
| Spreadsheet ID | MUST agree with the disposable workbook and payload |
| Planner version | MUST agree across Settings, metadata, and payload |
| Backup-format version | MUST be the current verifiable format accepted by production verification |
| Independent checksum | MUST be read from the backup metadata and match the bound payload |
| Selected sheet | MUST be the literal `03 Modules` |
| Approved baseline | MUST include sheet ID, dimensions, typed ordered matrix, and checksum |
| Non-target baseline | MUST cover all non-target sheets and relevant infrastructure state |
| Automation Log baseline | MUST preserve exact ordered pre-test rows |

Missing, duplicate, malformed, stale, or contradictory inputs SHALL produce `BLOCKED` before any mutation.

## 6. Safety gates

All gates in this section SHALL pass before controlled corruption begins.

### 6.1 Environment isolation

The runner SHALL prove that:

- the active workbook is the approved disposable workbook;
- the active Apps Script project is the approved disposable project;
- production identifiers are absent;
- Calendar deletion is disabled;
- no validation-unsafe trigger or deployment is active; and
- the `03 Modules` sheet exists exactly once and has the expected stable sheet ID.

### 6.2 Backup availability and public verification

The runner SHALL call `api_semVerifyBackup(backupId)` before reading or mutating `03 Modules`.

The runner SHALL require:

- the public envelope to return `ok:true`;
- its nested verification result to return `ok:true`;
- `fileFound:true`;
- `jsonValid:true`;
- an empty missing-sheet list; and
- no failure reason.

Any unsuccessful, inaccessible, trashed, malformed, legacy, mismatched, or indeterminately available backup SHALL produce `BLOCKED` before mutation.

### 6.3 Exact identity binding

The runner SHALL bind verification to recovery using all of the following:

1. enumerate `24 Semester Backups` and require at least one indexed backup;
2. require exactly one row to have the greatest valid `Timestamp`; that uniquely newest row is the retained M0-LV-14 emergency-backup candidate;
3. require the candidate to be newer than every other indexed backup and to have a distinct backup ID and Drive file ID;
4. call `api_semVerifyBackup()` for that candidate and require the complete public success contract in section 6.2;
5. read the Drive file ID, semester ID, planner version, included-sheet declaration, and checksum metadata from that row;
6. retrieve exactly that Drive file without altering it;
7. require the file not to be trashed;
8. parse its JSON payload once for the recovery operation;
9. require the payload backup ID, semester ID, spreadsheet ID, planner version, and format version to agree with the verified metadata and disposable environment;
10. require the payload to contain exactly the approved backup sheet set;
11. calculate the canonical payload checksum with the loaded production checksum function used by the verified backup format;
12. require that checksum to equal both the payload checksum and independent metadata checksum; and
13. retain the bound backup ID, Drive file ID, timestamp, and checksum in the evidence record.

The runner SHALL fail closed if the file identity, metadata, payload identity, or checksum changes between public verification and the final pre-write binding check.

The runner SHALL NOT treat a successful public verification of one file as authorization to read or restore from another file.

### 6.4 Selected-sheet enforcement

The recovery target SHALL be hard-bound to the literal `03 Modules`.

The runner SHALL require:

- exactly one payload member named `03 Modules`;
- that member's embedded `sheetName` to equal `03 Modules`;
- its `values` member to be a non-empty rectangular array; and
- its dimensions to equal the approved recovery baseline dimensions.

Any caller-supplied alternative, alias, partial match, case variation, or second target SHALL produce `BLOCKED` before mutation.

### 6.5 Ownership and non-target baseline

Before mutation, the runner SHALL record:

- the stable sheet ID of `03 Modules`;
- its complete typed-value matrix, dimensions, and checksum;
- checksums and dimensions for all non-target planner sheets;
- stable identities and relevant rows for infrastructure sheets;
- exact Settings values;
- exact Automation Log rows;
- the retained backup metadata and Drive disposition; and
- the absence of unrelated concurrent mutation.

The runner SHALL stop if the approved baseline cannot be proven or if current state differs from the approved pre-drill state.

## 7. Execution sequence

The runner SHALL execute the following sequence without reordering:

1. Capture the start time and sanitized environment identifiers.
2. Evaluate every safety gate in section 6.
3. Capture the approved `03 Modules` pre-corruption matrix, dimensions, and checksum.
4. Capture complete non-target and infrastructure baselines.
5. Apply the single approved controlled corruption to `03 Modules`.
6. Capture the corrupted matrix, dimensions, and checksum.
7. Prove that the corrupted checksum differs from the approved pre-corruption checksum.
8. Reconfirm selected-sheet identity and non-target invariance.
9. Reconfirm the bound backup file remains available, untrashed, identity-matched, and checksum-valid.
10. Extract only the `03 Modules.values` matrix from the bound payload.
11. Reconstruct supported typed values under section 9.
12. Replace only the value contents of the identity-proven `03 Modules` sheet with the reconstructed matrix.
13. Capture the recovered matrix, dimensions, and checksum.
14. Perform cell-by-cell typed comparison against the approved baseline.
15. Assert non-target, Settings, infrastructure, retained-backup, and Drive invariance.
16. Capture the permitted Automation Log delta, if any.
17. Perform cleanup under section 13.
18. Verify the reusable disposable baseline.
19. Return exactly one classification under section 15.

## 8. Controlled corruption

The controlled corruption SHALL clear only the value contents of row `HEADER_ROW + 1` across the captured used columns of the identity-proven `03 Modules` sheet. The sheet object, sheet name, stable sheet ID, dimensions, and all non-value features SHALL remain in place.

The mechanism MUST:

- affect only the identity-proven `03 Modules` sheet;
- change at least one value so the typed checksum differs from baseline;
- preserve the sheet object and stable sheet ID;
- avoid formulas, formatting, protections, validation, metadata, and structural features outside value scope;
- be completely captured as test-owned state; and
- be reversible through restoration from the bound backup.

The runner SHALL NOT use deletion or renaming of the sheet as the corruption mechanism. The runner SHALL NOT use the original in-memory baseline as the recovery source.

## 9. Value reconstruction rules

### 9.1 Supported value domain

The runner SHALL support only:

- blank;
- Boolean;
- finite number;
- string; and
- an ISO date string proven to have originated from a Date value in the approved typed baseline.

Unsupported objects, arrays as cells, non-finite numbers, malformed dates, or error values SHALL produce `FAIL` before the target write.

### 9.2 Blank values

JSON `null` SHALL NOT be silently converted to an empty cell unless the approved backup format and approved baseline jointly establish that representation.

An empty string SHALL remain an empty string. Zero SHALL remain numeric zero. Boolean `false` SHALL remain Boolean false. None of these values SHALL be treated as missing.

### 9.3 ISO date handling

Because the backup payload does not discriminate ordinary strings from serialized dates, the runner SHALL NOT infer dates from string shape alone.

A payload string MAY be reconstructed as a Date only when the corresponding approved baseline cell is typed as a Date and the payload value exactly equals that baseline Date's ISO serialization. All other payload strings SHALL remain strings.

This baseline-assisted rule is validation-only and SHALL NOT be represented as a general production deserialization contract.

### 9.4 Dimensions and order

The restored matrix SHALL preserve:

- exact row count;
- exact column count;
- exact row order;
- exact column order; and
- cell position.

The runner SHALL NOT sort, normalize, trim, coerce, deduplicate, or otherwise reinterpret values.

## 10. Restoration write boundary

Immediately before writing, the runner SHALL require:

- the target sheet ID still equals the captured sheet ID;
- the current target checksum still equals the captured corrupted checksum;
- every non-target checksum still equals its baseline; and
- the bound backup identity and checksum remain valid.

The runner SHALL clear and write only the used value range necessary to reproduce the approved `03 Modules` matrix. It SHALL NOT clear or write another sheet.

The runner SHALL NOT rely on the pre-corruption in-memory matrix for restored values. Every restored cell value SHALL derive from the bound backup's `03 Modules.values` matrix, subject only to the typed reconstruction rule in section 9.

If a partial write or runtime exception occurs, the runner SHALL classify the execution as `FAIL`, SHALL attempt only the identity-proven cleanup allowed by section 13, and SHALL report whether cleanup completed.

## 11. Assertions

### 11.1 Recovery assertions

The runner SHALL prove:

1. the pre-corruption checksum equals the approved baseline checksum;
2. the corrupted checksum differs from the approved baseline checksum;
3. the recovered dimensions equal the approved dimensions;
4. the recovered checksum equals the approved baseline checksum;
5. every recovered cell has the expected canonical type and value;
6. mismatch count equals zero;
7. row order and column order are unchanged; and
8. the selected sheet's stable identity is unchanged.

### 11.2 Non-target assertions

The runner SHALL prove that:

- every non-target sheet retains its baseline dimensions and typed-value checksum;
- Settings retain their baseline labels, types, and values;
- workflow, backup, archive, history, and template records are unchanged;
- the retained M0-LV-14 backup remains available, untrashed, and verifiable;
- no unrelated Drive artefact is created, altered, or trashed;
- no Calendar access or mutation occurs; and
- Automation Log changes are limited to the exact test-owned suffix permitted by the reviewed implementation.

## 12. Evidence schema and output

The runner SHALL return and preserve the following evidence:

| Evidence group | Required fields |
|---|---|
| Test identity | Test ID `M0-LV-15`, runner version or checksum, start and finish timestamps |
| Environment | Sanitized workbook ID, sanitized Apps Script project ID, disposable-safety result |
| Backup identity | Sanitized backup ID, sanitized Drive file ID, semester ID, planner version, format version |
| Verification | Complete public `api_semVerifyBackup()` response and assertion result |
| Binding | Metadata/payload identity result, independent checksum algorithm and checksum result |
| Target identity | Literal sheet name, stable sheet ID, baseline dimensions |
| Before corruption | Typed checksum and dimensions |
| After corruption | Controlled mutation description, typed checksum and dimensions |
| After recovery | Typed checksum, dimensions, typed comparison result, mismatch count |
| Mismatches | Cell coordinates, expected canonical value, actual canonical value; empty when PASS |
| Non-target state | Invariance result for sheets, Settings, infrastructure, Drive, and Calendar |
| Automation Log | Baseline count, exact test-owned delta, cleanup result |
| Timing | Verification duration, corruption duration, recovery duration, comparison duration, cleanup duration, total runtime |
| Cleanup | Per-artefact status, retained-backup status, reusable-environment result |
| Outcome | Exactly one classification and a precise reason when not PASS |

Repository evidence SHALL sanitize external identifiers. Full identifiers MAY exist only in the controlled live execution evidence needed to prove identity.

The runner SHALL record all evidence needed to diagnose a failure without rerunning the drill.

## 13. Cleanup specification

### 13.1 Ownership model

Cleanup SHALL operate only on state proven to be owned by the current M0-LV-15 execution:

- the controlled `03 Modules` corruption;
- an exact test-owned Automation Log suffix, if one is created; and
- validation-only transient evidence that does not belong to production or earlier tests.

Stable sheet ID, exact before/after checksums, exact row identities, and captured test-written values SHALL be used to prove ownership. Row position alone SHALL NOT prove ownership.

### 13.2 Retained backup policy

The M0-LV-14 emergency backup SHALL remain retained, unmodified, untrashed, and verifiable through the end of M0-LV-15 evidence review.

The runner SHALL NOT delete its metadata row or Drive file. Final disposition requires separate owner authorization after review.

### 13.3 Selected-sheet cleanup

If recovery succeeds and `03 Modules` equals the approved baseline, selected-sheet cleanup SHALL be idempotently complete without another write.

If execution stops after controlled corruption but before successful recovery, cleanup MAY restore `03 Modules` only from the already bound and checksum-valid backup payload. It SHALL NOT restore from an unrelated snapshot or guess at intended content.

If current sheet identity or content no longer matches either the exact test-owned corrupted state or approved recovered baseline, cleanup SHALL preserve current content and return `CLEANUP INCOMPLETE`.

### 13.4 Settings and infrastructure cleanup

The runner SHOULD produce no Settings or lifecycle-infrastructure changes.

If any such change occurs unexpectedly, the runner SHALL NOT overwrite or delete it speculatively. It SHALL record the delta and return `CLEANUP INCOMPLETE` unless exact test ownership and an approved cleanup rule exist.

Whole-sheet restoration of infrastructure sheets is prohibited.

### 13.5 Automation Log cleanup

The runner MAY remove only an exact contiguous Automation Log suffix proven to have been created by the current execution. Earlier rows SHALL remain byte-for-byte and type-for-type equivalent to baseline.

If prefix identity, suffix ownership, or concurrent-log safety cannot be proven, the runner SHALL preserve the log and return `CLEANUP INCOMPLETE`.

### 13.6 Idempotency and reusable environment

Cleanup SHALL be idempotent. Repeating cleanup SHALL either:

- confirm that the approved baseline is already restored; or
- perform only an outstanding identity-proven cleanup step.

Cleanup SHALL NOT fail merely because an owned transient artefact is already absent when baseline evidence proves the intended final state.

The disposable environment is reusable only when:

- `03 Modules` equals its approved typed baseline;
- all non-target state equals baseline;
- the retained emergency backup remains valid;
- no test-owned transient state remains except explicitly retained evidence; and
- cleanup status is complete.

## 14. Failure containment

The runner SHALL stop on the first failed required assertion.

| Condition | Required treatment |
|---|---|
| Safety gate or prerequisite unavailable before mutation | `BLOCKED` |
| Production verification rejects the retained backup before mutation | `BLOCKED` |
| Bound payload is malformed, mismatched, or checksum-invalid before mutation | `BLOCKED` |
| Controlled corruption cannot be proven | `FAIL` |
| Restore write, typed comparison, invariance assertion, or expected recovery behavior fails | `FAIL` |
| Unexpected exception after mutation | `FAIL`, followed by identity-proven cleanup attempt |
| Cleanup cannot safely restore or prove the reusable baseline | `CLEANUP INCOMPLETE` |

Evidence SHALL be preserved for every stopped execution. The runner SHALL NOT perform a second speculative mutation or correction.

## 15. Final classifications

The runner SHALL return exactly one of these classifications:

- `PASS`
- `FAIL`
- `CLEANUP INCOMPLETE`
- `BLOCKED`

No other classification is permitted.

### 15.1 PASS

`PASS` requires all of the following:

1. every safety gate passed;
2. the exact retained M0-LV-14 backup was publicly verified and identity-bound;
3. checksum validation passed immediately before restoration;
4. only `03 Modules` was corrupted and restored;
5. restored dimensions, order, types, and values exactly matched baseline;
6. recovered checksum equalled the approved pre-corruption checksum;
7. mismatch count was zero;
8. all non-target invariance assertions passed;
9. the retained backup remained valid and unchanged;
10. all required evidence was recorded; and
11. cleanup completed and the disposable environment was reusable.

### 15.2 FAIL

`FAIL` SHALL mean execution reached an owned validation mutation but recovery behavior, reconstruction, comparison, or a substantive assertion did not satisfy this specification. If cleanup also fails, the final classification SHALL be `CLEANUP INCOMPLETE` and the substantive failure SHALL remain recorded as the initiating incident.

### 15.3 CLEANUP INCOMPLETE

`CLEANUP INCOMPLETE` SHALL mean the runner cannot prove or safely restore the reusable disposable baseline after any test-owned mutation. It SHALL override `FAIL` as the final classification while preserving the initiating failure evidence.

### 15.4 BLOCKED

`BLOCKED` SHALL mean a prerequisite, environment identity, approved baseline, retained backup, or safe identity binding could not be established before test mutation. A blocked run SHALL perform no controlled corruption or recovery write.

## 16. Repository and product impact

Future implementation conforming to this specification:

- SHALL modify validation-only code exclusively;
- SHALL NOT change production scope;
- SHALL NOT create or expose a restore API;
- SHALL NOT create a public endpoint;
- SHALL NOT change backup format, verification behavior, persistence schema, migrations, DTOs, domain logic, or architecture;
- SHALL NOT be loaded into or used against production resources; and
- SHALL exist solely to generate M0-LV-15 validation evidence.

The implementation SHALL require separate owner authorization, focused code review, source-parity verification, and explicit execution authorization.

## 17. Implementation acceptance gates

Runner implementation SHALL NOT begin until:

1. this specification completes design review;
2. all ambiguities and hidden assumptions recorded below are resolved;
3. the controlled corruption mechanism is approved;
4. the exact validation-only payload-access mechanism is approved;
5. the callable entry-point name and parameter contract are approved;
6. the M0-LV-14 retained-backup evidence is confirmed complete; and
7. the owner marks this specification `OWNER APPROVED`.

Execution SHALL NOT begin until:

1. implementation is shown to conform to every normative requirement;
2. only approved validation code changed;
3. local syntax and static checks pass;
4. repository/disposable source parity is confirmed;
5. the retained M0-LV-14 backup verifies successfully;
6. the disposable baseline matches the Validation Record; and
7. the owner separately authorizes M0-LV-15 execution.

## 18. Design-review decision register

The design review MUST resolve the following without expanding production scope:

| Decision | Required resolution |
|---|---|
| Validation-only callable name | `m0RunRecoveryDrill15` |
| Runner input contract | No parameters; select the uniquely newest indexed backup and bind its full identity before mutation |
| Payload access | Read the exact metadata-bound Drive file directly in validation-only code; do not expose any API |
| Controlled corruption | Clear only row `HEADER_ROW + 1` across the captured used columns of identity-proven `03 Modules` |
| Date reconstruction | Use baseline-assisted Date reconstruction; never infer a Date from string shape alone |
| Automation logging | Write no Automation Log entry; require exact log invariance |
| Evidence retention | Full identifiers remain only in the controlled Apps Script execution result; repository reporting is sanitized |
| Backup final disposition | Retain the file and metadata unchanged until separate post-review cleanup authorization |

No decision in this register authorizes implementation or execution.
