# Gate B Validation Runner Specification

## Document control

| Field | Value |
|---|---|
| Project | Engineering Planner OS v1.5.0 |
| Milestone | Milestone 0 — Semester Backup Recovery and Hardening |
| Validation gate | Gate B — Destructive Workflow Safeguards |
| Status | OWNER APPROVED — FROZEN DESIGN SPECIFICATION — IMPLEMENTATION NOT AUTHORIZED |
| Applies to | M0-LV-11 through M0-LV-14 |
| Production checkpoint | `Milestone-0-GateA-Passed` / `a1e8f8675691aa68051d2b4af0355bd75d8a0a9d` |
| Normative validation record | `docs/validation/Milestone_0_Validation_Record.md` |

## 1. Purpose

Gate B validates that Engineering Planner OS permits destructive semester operations only through the approved production safeguards. It must prove that invalid authorization cannot clear active semester data and that valid authorization permits only the intended mutations.

Validation-only runners exist to make these workflows repeatable and auditable in the approved disposable environment. They prepare controlled fixtures, invoke public production APIs, observe returned and persisted outcomes, assert those outcomes, collect evidence, and restore the fixture.

Runners are orchestration code, not business logic. They must not decide whether a workflow is valid, reproduce backup verification, clear semester data, create authorization records independently, or replace any production dependency.

## 2. Scope

This specification covers only:

- M0-LV-11 — Normal delete gate rejects invalid backup.
- M0-LV-12 — Normal delete gate accepts valid backup.
- M0-LV-13 — Emergency-reset gate rejects a backup-creation failure.
- M0-LV-14 — Emergency-reset gate accepts a valid fresh backup.

The following are explicitly excluded:

- M0-LV-01 through M0-LV-10 and all Gate A execution;
- M0-LV-15 and the recovery drill;
- production implementation changes;
- automated or production recovery workflows;
- architecture, domain, persistence, DTO, migration, or API changes;
- Calendar deletion or Calendar validation;
- production workbooks, production Drive artefacts, and production Apps Script projects.

## 3. Design principles

1. Public production APIs are authoritative for workflow preparation, confirmation, backup creation, backup verification, and destructive mutation. Where independent observational verification evidence is required, runners may call only `api_semVerifyBackup(backupId)` and must assert its public `{ok,data}` envelope; they must never call `semVerifyBackup_` directly.
2. Runners never duplicate, reinterpret, or replace business rules.
3. Production code performs the authorization and backup verification required by each workflow. Runner-side observations never substitute for those checks.
4. Runners observe behavior through API results and persisted state.
5. Every test begins from a declared, reproducible fixture state and has deterministic success and failure criteria.
6. Execution is confined to the approved disposable workbook, Drive folder, and Apps Script project.
7. Calendar deletion remains disabled for every invocation.
8. Missing, ambiguous, or unverifiable preconditions cause the runner to fail closed before invoking a destructive acceptance path.
9. Evidence collection must be complete enough to reconstruct the invocation, authorization chain, mutations, and cleanup without exposing unsanitized identifiers in repository documentation.
10. Fixture restoration is part of test completion, not optional post-test housekeeping.
11. A failed assertion cannot be converted into a pass by cleanup or by a runner-side interpretation of production intent.
12. Historical failures and incidents are appended and preserved; they are never overwritten.

## 4. Runner architecture

The validation runner architecture is a one-way orchestration sequence:

```text
Runner
  ↓
Fixture preparation
  ↓
Public production API invocation
  ↓
Evidence capture
  ↓
Assertions
  ↓
Cleanup
  ↓
Report generation
```

Fixture preparation establishes environment identity, Calendar-disabled options, lifecycle-table baselines, typed-value workbook baselines, and artefact counts. Public API invocation calls only the approved production endpoints. Evidence capture records raw responses and relevant persisted state before assertions interpret them. Assertions compare the captured evidence with this specification and the Validation Record. Cleanup restores fixture state and disposes of test artefacts in the specified order. Report generation emits the sanitized evidence package and status only after cleanup status is known.

The runner must distinguish test outcome from cleanup outcome. A passing production assertion with failed or unknown cleanup is not a completed test and must be escalated.

### 4.1 Frozen destructive boundary

The exact active-data mutation boundary for M0-LV-12 and M0-LV-14 is:

| Sheet | Header row | Expected accepted-workflow behavior |
|---|---:|---|
| `08 Assignments` | 4 | Clear data below header |
| `09 Assessments` | 4 | Clear data below header |
| `12 Study Planner` | 4 | Clear data below header |
| `21 Study Tasks` | 4 | Clear data below header |
| `13 Revision Tracker` | 4 | Clear data below header |
| `11 Marks Tracker` | 6 | Clear data below header |
| `10 AF Components` | 4 | Clear data below header |
| `19 Attendance` | 4 | Clear data below header |
| `07 Academic Inbox` | 4 | Clear data below header |
| `20 Automation Log` | 4 | Clear prior data, then contain only the documented post-reset audit rows |

No other sheet may be cleared. For rejection tests, the first nine sheets above must remain identical to their pre-invocation baselines. `20 Automation Log` is assessed separately because the public API failure envelope legitimately appends an `API error` audit row and normal workflow preparation appends its own documented audit rows.

## 5. Per-test specifications

### 5.1 M0-LV-11 — Normal delete gate rejects invalid backup

#### Purpose

Prove that a production-prepared normal-delete workflow cannot authorize active-data clearing after its associated backup becomes invalid.

#### Public API

- `api_semStartDeleteWorkflow(reflection)`
- `api_semConfirmDelete(workflowToken, typedPhrase, options)`
- `api_semVerifyBackup(backupId)` — observational evidence only; never a substitute for confirmation-time production verification.

#### Parameters

- `reflection`: a deterministic validation reflection containing only non-sensitive test text.
- `workflowToken`: exactly the token returned by `api_semStartDeleteWorkflow`.
- `typedPhrase`: exactly `DELETE SEMESTER`.
- `options`: Calendar deletion explicitly false.
- `backupId`: exactly the backup ID returned by production preparation when calling the observational verification API.

#### Preconditions

- Disposable spreadsheet, Apps Script project, and Drive-parent identities match the approved environment.
- Calendar access is disabled, with zero triggers and zero deployments.
- The semester ID is known and nonblank.
- All required fixture sheets exist and match the recorded pre-Gate-B baseline.
- All destructive target sheets contain the expected fixture data.
- No prepared workflow or archive exists for the active test semester. If either exists, classify the test `ENVIRONMENT BLOCKED`; do not reuse or delete it.
- Lifecycle, template, archive, history, settings, and automation-log baselines have been captured.

#### Fixture preparation

Capture typed-value baselines and dimensions for every destructive target. Capture backup-folder file inventory and the relevant rows in `23 Semester Archive`, `24 Semester Backups`, `25 Module Templates`, `26 Semester History`, and `27 Semester Workflows`. Invoke the production preparation API and retain its returned workflow token, backup ID, archive semester ID, and reuse status. Require `reused:false`; otherwise stop as `ENVIRONMENT BLOCKED` without cleanup of the reused artefacts. Confirm that these identifiers agree with persisted workflow metadata and mark the resulting rows and Drive file as owned by the current test run. Confirm the associated backup through `api_semVerifyBackup(backupId)` and its nested `data.ok:true` result before applying the single authorized corruption: move only the test-owned workflow backup file to trash.

#### Execution sequence

Invoke the production confirmation API with the production-generated token, exact confirmation phrase, and Calendar deletion disabled. Capture the complete response before inspecting or cleaning any state.

#### Expected production response

The API returns exactly `{ok:false,error:"Associated backup failed verification: Backup Drive file is trashed and unavailable."}`. No success data is accepted.

#### Assertions

- The preparation response succeeded and supplied a genuine workflow token.
- The backup was valid before controlled corruption.
- Preparation returned `reused:false`.
- The confirmation response is `ok:false`.
- The error exactly matches `Associated backup failed verification: Backup Drive file is trashed and unavailable.`
- The response does not claim workflow consumption or active-data clearing.

#### Persisted-state assertions

- The first nine sheets in §4.1 match their pre-invocation typed-value baselines.
- `20 Automation Log` contains only the documented preparation entries and the expected `API error` rejection entry; it contains no semester-reset or delete-confirmed entry.
- The workflow is not `CONSUMED`.
- No reset-complete stage or consumed timestamp was recorded.
- The corrupted backup cannot verify.
- The workflow-to-backup and workflow-to-archive links remain traceable.
- No Calendar mutation occurred.

#### Evidence captured

Capture the full sanitized preparation and confirmation responses; workflow token reference; backup and archive references; pre-corruption verification result; corruption action and timestamp; post-corruption failure result; before-and-after workbook checksums and dimensions; workflow row status; Drive disposition; automation-log delta; runtime; and cleanup result.

#### Cleanup order

1. Preserve failure evidence and raw API responses.
2. Confirm active data remained unchanged.
3. Verify every cleanup target is owned by this test through the captured test-run, workflow, backup, archive, semester, Drive, and baseline identities. A mismatch stops cleanup and creates `CLEANUP INCOMPLETE`.
4. Remove the test workflow row by its exact production-generated token.
5. Remove the test backup metadata row by its exact backup ID while leaving the already-trashed test file in disposable trash pending final environment disposal.
6. Remove protection only from the exact test-created archive row, then remove that row by its captured semester ID and content identity.
7. Remove the exact matching test-created history row.
8. Restore template rows, Settings values, and automation-log baseline from captured typed values only when their current state matches the expected test-produced state.
9. Verify all lifecycle tables and active-data targets match their recorded baselines.

Each cleanup step is idempotent: an already-removed test-owned row or already-restored value is success only when baseline comparison proves the intended final state. A missing target with ambiguous ownership, an unexpected current value, or a shifted row that cannot be resolved by stable identity is `CLEANUP INCOMPLETE`; cleanup must not guess by row position.

#### Failure containment

If preparation, identity linkage, pre-corruption verification, confirmation, evidence capture, or cleanup differs from the specification, stop Gate B. Do not execute M0-LV-12 or later tests. Preserve the safest available failure state and register an incident.

#### Exit criteria

The test passes only when production rejects the corrupted backup before destructive mutation, the unchanged fixture is proven, required evidence is complete, and cleanup is verified.

### 5.2 M0-LV-12 — Normal delete gate accepts valid backup

#### Purpose

Prove that a production-prepared normal-delete workflow with a valid backup and archive authorizes the intended active-data reset exactly once after the exact confirmation phrase.

#### Public API

- `api_semStartDeleteWorkflow(reflection)`
- `api_semConfirmDelete(workflowToken, typedPhrase, options)`
- `api_semVerifyBackup(backupId)` — observational evidence only; never a substitute for confirmation-time production verification.

#### Parameters

- `reflection`: deterministic validation text.
- `workflowToken`: exactly the production-generated token returned by preparation.
- `typedPhrase`: exactly `DELETE SEMESTER`.
- `options`: Calendar deletion explicitly false.
- `backupId`: exactly the backup ID returned by production preparation when calling the observational verification API.

#### Preconditions

- All environment and fixture preconditions from M0-LV-11 hold.
- M0-LV-11 cleanup is verified.
- No prepared workflow or archive exists for the active test semester; any pre-existing artefact blocks the test.
- Every destructive target contains known non-header fixture data.

#### Fixture preparation

Capture complete active-data, lifecycle, template, Settings, automation-log, Drive, and archive baselines. Invoke the production preparation API. Record the returned token, backup ID, archive ID, and reuse status. Require `reused:false`, mark all resulting artefacts as owned by the current test run, establish that persisted workflow metadata links those identifiers, and confirm the associated backup through `api_semVerifyBackup(backupId)` with outer `ok:true` and nested `data.ok:true`.

#### Execution sequence

Invoke the production confirmation API with the returned token, exact phrase, and Calendar deletion disabled. Capture the response and persisted state before any restoration.

#### Expected production response

The API returns `ok:true`. Its data reports the cleared sheets, zero deleted Calendar events, no Calendar errors, and `workflowTokenConsumed:true`.

#### Assertions

- Preparation and confirmation both return `ok:true`.
- Preparation returns `reused:false`.
- The confirmation token is exactly the production-generated token.
- The associated backup was valid when production authorized the reset.
- `workflowTokenConsumed` is true.
- `deletedEvents` is zero and Calendar errors are empty.
- The returned cleared-sheet set agrees with the populated fixture targets.

#### Persisted-state assertions

- The exact §4.1 boundary is observed: the first nine sheets are cleared below their stated headers and `20 Automation Log` contains only documented post-reset audit rows.
- Non-target planner, infrastructure, archive, and configuration sheets are not cleared.
- The workflow status is `CONSUMED` with a consumed timestamp and reset-complete stage.
- The backup and archive remain traceable.
- The manual-finished setting is false.
- `20 Automation Log` contains the intentional post-reset audit entries rather than being required to remain empty.
- No Calendar mutation occurred.

#### Evidence captured

Capture complete preparation and confirmation responses; sanitized workflow, backup, archive, and Drive identifiers; successful backup-verification evidence; before-and-after dimensions and typed-value checksums; exact cleared-sheet set; workflow status transition; Settings delta; automation-log delta; non-target comparison; runtime; and cleanup status.

#### Cleanup order

1. Preserve acceptance evidence and post-reset comparisons.
2. Restore active-data target values, dimensions, and types from the recorded fixture baseline.
3. Restore Settings and automation-log baselines.
4. Verify ownership of every cleanup target using the captured workflow token, backup ID, archive/semester ID, Drive ID, content identity, and baseline. Any mismatch is `CLEANUP INCOMPLETE`.
5. Remove the consumed workflow row by its exact token.
6. Remove protection only from the exact test-created archive row, then remove that row and its exact matching history row.
7. Restore template rows only from their typed-value baseline and only when their current values match the expected test-produced state.
8. Remove the M0-LV-12 backup metadata row by its exact test-owned backup ID and move its exact Drive file to disposable trash. This backup is not retained for M0-LV-15.
9. Verify the complete fixture and lifecycle baseline before proceeding.

Cleanup is idempotent under the rules in §5.1: already-complete test-owned steps are accepted only after baseline verification; ambiguous identity or unexpected state stops cleanup without positional guessing.

#### Failure containment

Any incorrect rejection, unexpected mutation, missing target clear, non-target mutation, Calendar activity, workflow-state mismatch, or incomplete cleanup stops Gate B immediately. Preserve the post-operation workbook when restoration could conceal the defect, unless a safe restoration step is necessary to prevent further harm.

#### Exit criteria

The test passes only when the valid production authorization permits exactly the intended reset, all evidence is recorded, and the disposable fixture is verifiably restored.

### 5.3 M0-LV-13 — Emergency-reset gate rejects backup-creation failure

#### Purpose

Prove that emergency reset cannot clear active semester data when its mandatory fresh backup cannot be created.

#### Public API

- `api_semResetActiveSemester(typedPhrase, options)`

#### Parameters

- `typedPhrase`: exactly `EMERGENCY RESET`.
- `options`: Calendar deletion explicitly false.

#### Preconditions

- Approved disposable environment and safety controls are confirmed.
- M0-LV-12 cleanup and fixture restoration are verified.
- All required backup sheets and active-data targets match the baseline before controlled preparation.
- Backup-folder and backup-index baselines are captured.

#### Fixture preparation

Capture active-data and backup artefact baselines. Preserve the complete identity, name, contents, dimensions, and typed-value checksum of `03 Modules`, then temporarily rename it to the fixed validation-only name `03 Modules — M0-LV-13 TEMP`. No empty replacement sheet is created. This is the only authorized M0-LV-13 fault injection and must be reversed regardless of test outcome.

#### Execution sequence

Invoke the public emergency-reset API with the exact phrase and Calendar deletion disabled. Capture the complete response and all before-and-after artefact counts. Restore the temporarily altered sheet in the cleanup phase, including after an unexpected exception.

#### Expected production response

The API returns exactly `{ok:false,error:"Backup was not created because required sheets are missing or empty: 03 Modules"}`.

#### Assertions

- The response is `ok:false`.
- The error exactly matches `Backup was not created because required sheets are missing or empty: 03 Modules`.
- No success data or emergency backup ID is returned.
- The test does not require or introduce a synthetic post-creation verification hook.

#### Persisted-state assertions

- The first nine sheets in §4.1 match their pre-invocation typed-value baselines.
- `20 Automation Log` differs only by the expected `API error` row for backup-creation failure and contains no reset-complete or emergency-reset-completed entry.
- No new backup-index row remains.
- No new Drive backup file remains.
- No emergency-reset completion audit entry exists.
- No Calendar mutation occurred.
- The temporarily changed required sheet is restored exactly.

#### Evidence captured

Capture selected sheet identity; reversible preparation method; precondition checksum; complete API response; Drive and backup-index counts before and after; destructive-target comparisons; automation-log delta; restoration proof; runtime; and cleanup status.

#### Cleanup order

1. Preserve the failure response and artefact-count evidence.
2. Restore the exact captured `03 Modules` sheet from `03 Modules — M0-LV-13 TEMP` to its original name, preserving the same sheet identity, values, dimensions, and types.
3. Remove any unexpected partial backup metadata only after its evidence is captured.
4. Trash any unexpected partial Drive file using the approved cleanup mechanism.
5. Verify all active-data, backup-index, Drive, and fixture baselines.

Cleanup is idempotent only when the captured sheet identity is found under either the temporary or original name and its typed-value checksum matches the baseline. If both names exist, neither exists, identity differs, or content differs, stop with `CLEANUP INCOMPLETE`; do not rename or delete by name alone. Unexpected partial artefacts may be removed only when their test ownership is proven by captured identifiers.

#### Failure containment

If any active data is cleared, a partial artefact cannot be accounted for, or the required sheet cannot be restored, stop Gate B immediately and preserve evidence. A cleanup failure is a blocking incident even when production correctly rejected the reset.

#### Exit criteria

The test passes only when backup creation fails through the public workflow, no destructive mutation occurs, no unaccounted artefact remains, and the selected sheet and full fixture are restored.

### 5.4 M0-LV-14 — Emergency-reset gate accepts valid fresh backup

#### Purpose

Prove that emergency reset creates and verifies a fresh backup before clearing exactly the intended active semester data.

#### Public API

- `api_semResetActiveSemester(typedPhrase, options)`
- `api_semVerifyBackup(backupId)` — post-operation observational evidence only.

#### Parameters

- `typedPhrase`: exactly `EMERGENCY RESET`.
- `options`: Calendar deletion explicitly false.
- `backupId`: exactly the returned `emergencyBackupId` when performing post-operation observational verification.

#### Preconditions

- Approved disposable environment and safety controls are confirmed.
- M0-LV-13 cleanup is verified.
- Every mandatory backup sheet is present and structurally valid.
- Every destructive target contains known non-header fixture data.
- Complete workbook, Settings, automation-log, backup-index, and Drive baselines are captured.

#### Fixture preparation

Record all destructive-target and non-target baselines. Record backup-folder and index inventories. Confirm no unrelated execution is in progress and prepare the exact Calendar-disabled options.

#### Execution sequence

Invoke the public emergency-reset API with the exact phrase. Capture the complete response and persisted post-operation state before restoration. After capturing the success response, call `api_semVerifyBackup(emergencyBackupId)` only to confirm the retained artefact's current validity. The production success path itself, not this later observation, is the evidence that internal verification preceded reset.

#### Expected production response

The API returns `ok:true`. Its data reports the cleared sheets, zero deleted Calendar events, no Calendar errors, and a nonblank `emergencyBackupId`.

#### Assertions

- The response is `ok:true`.
- `emergencyBackupId` is nonblank and identifies the new backup.
- The corresponding metadata and Drive artefact exist and link to the disposable workbook and active semester.
- `api_semVerifyBackup(emergencyBackupId)` returns outer `ok:true` and nested `data.ok:true`; this confirms retained-backup validity without substituting for or independently proving the internal temporal ordering.
- `deletedEvents` is zero and Calendar errors are empty.
- The returned cleared-sheet set agrees with the populated fixture targets.

#### Persisted-state assertions

- The exact §4.1 boundary is observed: the first nine sheets are cleared below their stated headers and `20 Automation Log` contains only the documented post-reset audit rows.
- Non-target sheets are unchanged except documented lifecycle, Settings, and audit effects.
- The manual-finished setting is false.
- The emergency backup remains available and valid.
- The automation log contains the expected reset and emergency completion evidence.
- No Calendar mutation occurred.

#### Evidence captured

Capture the complete API response; sanitized emergency backup and Drive identifiers; backup metadata linkage; verification result; before-and-after dimensions and typed-value checksums; cleared-sheet set; non-target comparison; Settings delta; automation-log delta; runtime; and cleanup disposition.

#### Cleanup order

1. Preserve the success response, emergency-backup evidence, and post-reset comparisons.
2. Retain the verified emergency backup for the separately authorized M0-LV-15 recovery drill.
3. Restore active-data target values, dimensions, and types from the baseline; restoration is mandatory before Gate B completion.
4. Restore Settings and automation-log baselines only after their evidence is secured; restoration is mandatory before Gate B completion.
5. Verify the restored fixture and retained backup independently.
6. Do not dispose of the emergency backup until the recovery drill and owner review are complete.

Restoration is idempotent only when each sheet is the captured disposable sheet and its current state is either the expected post-reset state or the exact baseline. Unexpected content, identity mismatch, or ambiguous state produces `CLEANUP INCOMPLETE`; the runner must not overwrite it. The retained emergency backup is never removed by Gate B cleanup.

#### Failure containment

An invalid or unavailable emergency backup, incomplete clear, non-target mutation, Calendar activity, missing evidence, or failed fixture restoration stops Gate B. Preserve the emergency backup and relevant post-reset state for investigation.

#### Exit criteria

The test passes only when production creates and validates the fresh backup, performs exactly the intended reset, captures complete evidence, retains the recovery artefact, and leaves the disposable environment in the approved post-Gate-B state.

## 6. Evidence schema

The following evidence categories and fields are normative. Repository records must sanitize external identifiers while retaining full identifiers only in the controlled live evidence context.

### 6.1 API responses

- Test ID.
- Public API name.
- Complete request inputs, including the exact confirmation phrase, exact `deleteCalendarEvents:false` option, deterministic reflection content where applicable, and sanitized production-generated identifiers.
- Complete returned envelope.
- Complete returned data or error.
- Invocation start, completion, and duration.

### 6.2 Workbook state

- Disposable workbook identifier.
- Sheet name.
- Header row.
- Last row and last column before and after.
- Typed-value checksum before and after.
- Expected mutation classification: unchanged, cleared, restored, or documented lifecycle write.
- Mismatch count and mismatch description.

### 6.3 Drive artefacts

- Sanitized file identifier.
- Associated backup ID.
- Parent-folder identity confirmation.
- Creation timestamp.
- Availability and trash state.
- Final disposition.

### 6.4 Workflow state

- Sanitized workflow token.
- Semester ID.
- Backup ID.
- Archive semester ID.
- Prepared timestamp and expiry.
- Status before and after.
- Consumed timestamp.
- Last stage and last error.

### 6.5 Backup state

- Backup ID.
- Metadata-row reference.
- Drive-file reference.
- Semester, planner-version, and spreadsheet linkage.
- Required sheet count.
- Verification result before authorization when applicable.
- Verification failure after controlled corruption for M0-LV-11.
- Public verification API envelope, recorded separately from the workflow API envelope.
- Fingerprint evidence in sanitized form.

### 6.6 Archive state

- Semester ID and archive-row reference.
- Created timestamp.
- Existence before preparation and after preparation.
- Workflow linkage.
- Protection state relevant to cleanup.
- Final disposition.

### 6.7 Automation log

- Row count before and after.
- Expected action labels.
- Sanitized references.
- Unexpected actions or errors.
- Restored or retained disposition.

### 6.8 Timing

- Environment-preflight timestamp.
- API start and completion timestamps.
- Duration in milliseconds.
- Cleanup start and completion timestamps.
- Lock, quota, latency, or timeout observations.

### 6.9 Identifiers

- Test-run identifier.
- Test ID.
- Workbook and Apps Script project identifiers in sanitized form.
- Workflow, backup, archive, Drive, and incident identifiers.
- Explicit mapping among related identifiers.

### 6.10 Screenshots

Screenshots are optional corroborating evidence, not a substitute for structured results. When captured, they are limited to the disposable project, sanitized where necessary, labelled with test ID and timestamp, and must not expose account details or full sensitive identifiers in repository files.

## 7. Cleanup specification

Cleanup is deterministic, test-specific, and verified against pre-test baselines.

| Test | Restore order | Artefact removal or retention | Post-cleanup verification |
|---|---|---|---|
| M0-LV-11 | Active-data unchanged proof; ownership checks; workflow; backup index; archive protection/row; history; templates; Settings; automation log | Exact corrupted test file remains in disposable trash; exact test-owned lifecycle rows removed | First nine boundary sheets unchanged; Automation Log restored after expected deltas are captured; lifecycle and configuration baselines match |
| M0-LV-12 | Preserve evidence; active data; Settings; automation log; ownership checks; workflow; archive protection/row; history; templates; backup | Exact M0-LV-12 backup metadata removed and exact Drive file trashed; it is not retained for recovery | Full fixture and lifecycle baseline match; no test-owned preparation artefact remains outside disposable trash |
| M0-LV-13 | Captured `03 Modules` identity restored first; unexpected metadata; unexpected Drive file; remaining fixture | No expected backup artefact; proven test-owned unexpected partial file is trashed only after evidence capture | `03 Modules`, first nine boundary sheets, Drive inventory, backup index, and restored Automation Log match baseline |
| M0-LV-14 | Preserve evidence; retain emergency backup; restore active data; restore Settings and automation log | Emergency backup retained for M0-LV-15 and never removed by Gate B cleanup | Fixture baseline matches; retained emergency backup remains available and returns a successful public verification envelope |

Cleanup must record one of `SUCCEEDED`, `FAILED`, or `UNKNOWN`. `FAILED` and `UNKNOWN` are blocking and classify the test `CLEANUP INCOMPLETE`. Every cleanup mutation is guarded by stable test ownership and expected-current-state checks; row position alone is never sufficient. Repeating cleanup is permitted only when an already-complete state can be proven against the baseline. Cleanup must not delete evidence required to explain an unexpected result. If restoration could conceal a production defect, the runner performs safety-only cleanup, preserves the remaining state, and stops unless the owner separately authorizes further restoration.

## 8. Failure containment

### Runner failure

A runner orchestration error stops the current test and all later Gate B tests. Production outcome remains unclassified unless independently captured. Perform only identity-proven safety cleanup that cannot conceal evidence. Register a validation-harness incident and classify the result `VALIDATION HARNESS FAILURE`.

### Production failure

An expected production rejection is `PASS` only when it exactly satisfies the test contract and cleanup succeeds. Any unexpected production response or error is captured verbatim, the environment is preserved where safe, later tests do not run, and the result is `IMPLEMENTATION FAILURE` only when the runner and environment are proven valid and the production boundary was actually exercised. Otherwise use the applicable harness or environment classification.

### Environment failure

Identity mismatch, Calendar enablement, trigger or deployment presence, missing fixture dependency, pre-existing workflow/archive state, lock contention that cannot be safely retried, or unauthorized artefact state prevents invocation. Do not call the production workflow. Preserve the environment, record an environment incident, classify `ENVIRONMENT BLOCKED`, and do not classify implementation behavior.

### Assertion failure

Capture the raw result and persisted state before cleanup and stop later tests immediately. Perform only identity-proven safety cleanup that cannot conceal the mismatch. Do not weaken or rewrite the assertion during the execution run. Classify `IMPLEMENTATION FAILURE` only when the captured response/state proves production behavior violated the frozen contract; otherwise classify `VALIDATION HARNESS FAILURE`.

### Unexpected exception

Record the exception, current stage, identifiers already created, workbook state, and cleanup feasibility. Stop the gate. Perform only necessary identity-proven safety cleanup and preserve all other evidence. Classify according to the proven source: `IMPLEMENTATION FAILURE`, `VALIDATION HARNESS FAILURE`, or `ENVIRONMENT BLOCKED`; if the source cannot be proven, use `VALIDATION HARNESS FAILURE` with the production result unclassified.

### Cleanup failure

Record the attempted action, target, error, and resulting state. Mark cleanup `FAILED` or `UNKNOWN`, classify `CLEANUP INCOMPLETE`, and do not continue. Do not retry against an ambiguous target or continue restoration where doing so could destroy evidence.

### Incident escalation

Every new incident receives a stable identifier, classification, timestamp, affected test, evidence references, containment, and owner decision. Historical incidents, including `ENV-04`, `VAL-01`, and `VAL-02`, remain preserved. Resolution appends information; it never erases the original event.

The only allowed per-test final classifications are `PASS`, `IMPLEMENTATION FAILURE`, `VALIDATION HARNESS FAILURE`, `ENVIRONMENT BLOCKED`, and `CLEANUP INCOMPLETE`. A blocked or unexecuted test is never classified as an implementation failure. Any classification other than `PASS` stops Gate B before the next test.

## 9. Prohibited behavior

Validation runners must not:

- call private lifecycle helpers directly;
- call `semResetActiveSemester_` directly;
- manufacture workflow tokens;
- create or edit workflow-table rows directly to simulate authorization;
- edit backups except for the single controlled M0-LV-11 corruption authorized by this specification;
- directly perform any destructive business mutation;
- duplicate or substitute backup-verification logic;
- call `semVerifyBackup_` directly; observational verification may use only `api_semVerifyBackup` and may not substitute for production workflow verification;
- bypass production lifecycle locking;
- alter production behavior or source;
- replace production dependencies with stubs;
- call Calendar services or set Calendar deletion true;
- create deployments or triggers;
- execute against production resources;
- transform a production failure into success through runner interpretation;
- continue after a blocking failure or incomplete cleanup.

## 10. Traceability matrix

| Validation test | Production API | Callable-boundary assessment | Validation Record | Expected evidence | Cleanup section |
|---|---|---|---|---|---|
| M0-LV-11 | `api_semStartDeleteWorkflow`; `api_semConfirmDelete`; observational `api_semVerifyBackup` | Existing callable boundary sufficient; production confirmation re-verifies linked backup | Principal table M0-LV-11; destructive-workflow safety record; incident register | Preparation and exact rejection envelopes; test-owned workflow/backup/archive; unchanged first nine boundary sheets; expected audit delta; unconsumed workflow | §5.1 Cleanup order; §7 M0-LV-11 |
| M0-LV-12 | `api_semStartDeleteWorkflow`; `api_semConfirmDelete`; observational `api_semVerifyBackup` | Existing callable boundary sufficient; production token and exact phrase authorize reset | Principal table M0-LV-12; destructive-workflow safety record | Preparation and success envelopes; valid test-owned backup; exact §4.1 boundary; consumed workflow; zero Calendar deletion | §5.2 Cleanup order; §7 M0-LV-12 |
| M0-LV-13 | `api_semResetActiveSemester` | Existing callable boundary sufficient through deterministic backup-creation failure | Principal table M0-LV-13; destructive-workflow safety record | Failure envelope; missing/empty required sheet; unchanged active data; no backup artefact; restoration proof | §5.3 Cleanup order; §7 M0-LV-13 |
| M0-LV-14 | `api_semResetActiveSemester`; observational `api_semVerifyBackup` | Existing callable boundary sufficient; production creates and verifies backup inside locked workflow | Principal table M0-LV-14; destructive-workflow safety record | Success envelope; retained emergency backup linkage and validity; exact §4.1 boundary; zero Calendar deletion | §5.4 Cleanup order; §7 M0-LV-14 |

## 11. Acceptance criteria

Runner implementation may begin only when all of the following are true:

1. The owner approves this specification as the frozen normative design.
2. The three workflow APIs—`api_semStartDeleteWorkflow`, `api_semConfirmDelete`, and `api_semResetActiveSemester`—and the observational verification API `api_semVerifyBackup` are confirmed unchanged at the approved implementation commit.
3. No production source change is included in the runner implementation scope.
4. Runner code is housed only in the disposable validation harness and is clearly marked validation-only.
5. Every runner has a documented mapping to its per-test specification, evidence schema, cleanup order, and failure-containment rule.
6. Environment identity and Calendar-disabled checks are mandatory preconditions.
7. M0-LV-11 obtains its workflow token only from production preparation and relies on production confirmation to reject the corrupted backup.
8. M0-LV-12 obtains its workflow token only from production preparation and proves single-use consumption.
9. M0-LV-13 uses deterministic backup-creation failure and introduces no production interception hook.
10. M0-LV-14 retains its emergency backup for the separately authorized recovery drill.
11. Baseline, evidence, cleanup, sanitization, and incident schemas are defined for all four runners.
12. Local review confirms runners contain orchestration and assertions only, with no duplicated business rules.
13. The complete runner diff is reviewed and separately approved before upload or execution.

## 12. Future implementation constraints

Future runner implementation must conform exactly to this specification and remain validation-only. Any deviation affecting public APIs, assertion contracts, evidence, cleanup, production behavior, or test scope requires a reviewed specification amendment before implementation.

Implementation requires separate owner authorization. Uploading runner code, changing `Validation_Support.gs`, modifying the disposable environment, and executing Gate B each require authorization within their applicable gates. This document authorizes none of those actions.

Production changes are not permitted under the runner implementation scope. If implementation discovers that a required boundary is unavailable or behaves differently from this frozen specification, work must stop and the discrepancy must receive separate architectural and owner review.
