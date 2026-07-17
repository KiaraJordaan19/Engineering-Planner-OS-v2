# Milestone 0 — M0-LV-05 Remediation Summary

| Field | Value |
|---|---|
| Project | Engineering Planner OS v1.5.0 |
| Milestone and gate | Milestone 0 — Gate A |
| Test | M0-LV-05 — Deleted Drive file detection |
| Incidents | `VAL-01` implementation defect; `VAL-02` validation-harness defect |
| Date | 2026-07-17 |
| Status | Remediated and verified; Gate A paused before M0-LV-06 |

## Root cause

`DriveApp.getFileById()` continued to return a retrievable file after `file.setTrashed(true)`. The original `semVerifyBackup_()` treated successful retrieval as availability and did not inspect `file.isTrashed()`. A trashed backup could therefore pass verification and potentially authorize a destructive workflow (`VAL-01`).

The minimal implementation remediation correctly returned the frozen unavailable-file result, but the first live retest runner still required the obsolete reason substring `Drive file not found`. It rejected the new precise reason even though `ok:false` and `fileFound:false` were correct (`VAL-02`).

## Risk assessment

The implementation change is low risk and fail-closed. It adds one availability decision immediately after the existing Drive lookup and before reading the blob. Valid non-trashed files continue through the unchanged verification path. No backup format, fingerprint, DTO, metadata layout, sheet schema, persistence rule, migration, or external integration changed.

The change strengthens M0-LV-11 through M0-LV-15 because a trashed or indeterminately available backup cannot authorize deletion, reset, or recovery. Earlier integrity behavior remains unchanged.

## Files changed

- `Semester_Backup_v1.4.0.js` — production availability check for trashed and indeterminate files.
- `Semester_Backup_v1.4.0_Test.cjs` — Drive mock support plus trashed-file and inspection-error characterization tests.
- Disposable Apps Script `Validation_Support.gs` — exact frozen M0-LV-05 reason assertion; validation-only and not repository production code.
- `docs/validation/Milestone_0_Validation_Record.md` — preserved both incidents, remediation evidence, checksums, runtimes, and final PASS.
- This summary — permanent audit trail for the remediation sequence.

## Local verification

- Complete backup characterization suite: **14 passed, 0 failed**.
- `node --check Semester_Backup_v1.4.0.js`: passed.
- `node --check Semester_Backup_v1.4.0_Test.cjs`: passed.
- Syntax checks for all repository `.js` files: passed.
- `git diff --check`: passed.

Local coverage confirms fail-closed behavior when a file is missing, inaccessible, trashed, or its trash state cannot be inspected, while a valid non-trashed file follows the existing successful path.

## Live verification

The corrected source was loaded only into the disposable Apps Script project. A fresh pull matched the local production source byte-for-byte.

- Original M0-LV-05 run: implementation failure; `VAL-01` recorded.
- First remediation retest: substantive result passed, but obsolete harness reason matching failed; `VAL-02` recorded.
- Final retest: **PASS**, 7,579 ms.
- Final returned result: `{"ok":false,"fileFound":false,"jsonValid":false,"missingSheets":[],"reason":"Backup Drive file is trashed and unavailable."}`
- Production, Calendar, deployments, and later Gate A tests remained untouched.

## Runner correction

The runner still requires `ok === false` and `fileFound === false`. Its reason assertion was narrowed to exact equality with the frozen reason:

> `Backup Drive file is trashed and unavailable.`

No arbitrary failure reason is accepted. Both `VAL-01` and `VAL-02` remain in the Validation Record as resolved historical incidents.

## Final checksum

Verified SHA-256 for the remediated `Semester_Backup_v1.4.0.js` loaded in the disposable project:

`edf4b9ed06c2cc1594055479f100cd013380b32cba073d4b2e1eeb8257ab17b2`

## Final approved behaviour

> A backup file is available for verification only when it can be retrieved and is not trashed. Missing, inaccessible, trashed, or indeterminately available files fail closed with `ok:false` and `fileFound:false`. A trashed file returns the exact reason `Backup Drive file is trashed and unavailable.`

No authorization to execute M0-LV-06 or later tests, create a tag, commit, push, deploy, or begin Gate B is recorded by this document.
