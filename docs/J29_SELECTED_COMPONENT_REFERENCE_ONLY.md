# J29 — Selected Components Reference-Only

## Purpose

J29 is a guarded atomic-batch writer for preparing one or more component
occurrences as Reference-Only. It does not discover targets by part number and
does not walk the assembly. The operator preselects direct component rows in
Assembly Navigator.

J28 V2 provided the runtime contract used by J29. In the accepted NX X 2506
result, sequence 177 (`264MN028171A01/A / 028060/A`) was classified
`REFERENCE_ONLY` because it owned this single direct occurrence attribute:

```text
Title:          REFERENCE_COMPONENT
Type:           STRING
Value:          <blank>
Inherited:      false
OwnedBySystem:  false
PdmBased:       false
```

`PLIST_IGNORE_MEMBER` and `PLIST_IGNORE_SUBASSEMBLY` were both absent. J29
reproduces only that exact state.

## Batch safety contract

J29:

- defaults to `APPLY` by operator decision;
- accepts 1–100 preselected NX assembly components by default;
- permits `NX_J29_MAX_SELECTION` to override the positive selection limit;
- accepts only direct children of the active assembly root, so every modified
  occurrence is owned by the same active parent assembly;
- preflights every selection before writing any target;
- blocks the complete batch if any target is invalid, suppressed, nested,
  duplicated, unreadable, conflicting, or nonstandard;
- treats an exact existing Reference-Only marker as a successful no-op while
  applying the remaining eligible targets;
- requires the parent assembly to be both Work Part and Displayed Part;
- never walks the assembly tree or loads a component/prototype;
- never checks out, checks in, or saves;
- requires an already-writable native parent assembly, or a managed parent
  assembly already checked out by the current Teamcenter user;
- recognizes the current Teamcenter user by an exact identity token, including
  the runtime form `display name (32-character-user-id)` versus the bare ID;
- creates one visible NX undo mark for the complete batch;
- writes `Component.SetInstanceUserAttribute("REFERENCE_COMPONENT", -1, "",
  NXOpen.Update.Option.Now)` to each eligible occurrence;
- rereads and verifies every write against the exact J28 V2 contract;
- rolls back the complete batch if any write, read-back, verification, or
  paired CSV/JSON evidence write fails; and
- leaves a successful verified batch unsaved and undoable for inspection.

The deferred `AtTermination` unload option follows the memory-safe J28 V2
lifecycle and avoids invalidating NX bridge objects during journal teardown.

## Run sequence

1. Open the parent assembly and make it both Work Part and Displayed Part.
2. In Teamcenter, manually check out the parent assembly. J29 will not perform
   checkout.
3. In Assembly Navigator, preselect 1–100 direct child component rows. Do not
   select their geometry in the graphics window.
4. Play J29. The file default is:

   ```python
   USER_MODE = "APPLY"
   ```

5. Require `APPLIED_VERIFIED` or `ALREADY_REFERENCE_ONLY`. Inspect the CSV and
   JSON beneath `%USERPROFILE%\Desktop\NX_REFERENCE_ONLY`, or beneath
   `NX_JOURNALS_IO_DIR\NX_REFERENCE_ONLY` when configured.
6. Inspect the Assembly Navigator and BoM result. The parent assembly remains
   modified but unsaved.
7. Save manually only if the result is correct. Otherwise use one NX Undo:
   `J29 Set selected components Reference-Only`.

Set `NX_J29_MODE=DRY_RUN` for a no-write batch preflight. Set
`NX_J29_MAX_SELECTION=<positive integer>` only when a reviewed batch must exceed
the default limit.

## Evidence format

Each run creates one paired CSV/JSON report. It contains one batch summary and
one target record per selected occurrence. Every target records its selection
index, identity, before/after controls, action, status, verification result, and
error. A blocked target remains visible even though the atomic batch writes
nothing.

## Verdicts

- `DRY_RUN_READY`: the complete selection passed; NX was not modified.
- `APPLIED_VERIFIED`: every required write was reread and verified; exact
  existing markers were retained as no-ops.
- `ALREADY_REFERENCE_ONLY`: every selected occurrence already matched; no
  write or undo mark was needed.
- `BLOCKED_BATCH`: one or more per-target preflight statuses blocked the whole
  batch; nothing was changed.
- `BLOCKED_SELECTION`: no usable selection was available.
- `BLOCKED_SELECTION_LIMIT`: selection count exceeded the configured cap.
- `BLOCKED_WRITE_ACCESS`: parent-assembly write access was not proven.
- `BLOCKED_UNDO_MARK`: NX could not create the required batch undo mark.
- `ROLLED_BACK`: a failure occurred and every selected baseline was restored
  and verified.
- `ROLLBACK_FAILED`: restoration could not be proven for every target; use NX
  Undo and inspect immediately.

Per-target statuses include `ELIGIBLE`, `DRY_RUN_READY`, `APPLIED_VERIFIED`,
`ALREADY_REFERENCE_ONLY`, `BLOCKED_SELECTION`, `BLOCKED_CONTROL_CONFLICT`,
`BLOCKED_NONSTANDARD_REFERENCE`, `WRITE_OR_VERIFY_FAILED`, and `ROLLED_BACK`.
An eligible target after the failure point is logged as `NOT_ATTEMPTED`.

## Runtime evidence and verification boundary

The NX X 2506 V1 evidence committed on 2026-09-01 showed a false access block:
checkout owner `aqil ameran (99946e1828964542b86c86d6c2cf3cbe)` and current
user `99946e1828964542b86c86d6c2cf3cbe` are the same Teamcenter identity. V2
uses the shared strong identity token instead of literal full-string equality.
The evidence also showed that a genuinely checked-in/read-only parent was
correctly blocked.

Local tests verify batch preflight, selection limits, exact API arguments,
same-user identity matching, idempotence, all-or-nothing conflicts, multi-write
read-back, complete rollback, report structure, and the absence of
load/save/checkout calls. Siemens NX is not installed on this host. Only Aqil's
next NX X 2506 CSV/JSON and Listing Window result can prove V2 runtime behavior.
