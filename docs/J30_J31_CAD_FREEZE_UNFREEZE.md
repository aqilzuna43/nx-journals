# J30/J31 — WAE CAD Freeze and Unfreeze

J30 and J31 are separate NX X 2506 UI-button journals. They use the native
Teamcenter workflows proven by J32 runtime evidence:

```text
J30: Part_Freeze_Process
J31: Part_Unfreeze_Process
```

This separates formal configuration revision from working CAD iteration:

```text
TCX Rev A / WAE 1 / editable
  -> J30 Freeze
TCX Rev A / WAE 1 / frozen baseline
  -> J31 Unfreeze
TCX Rev A / WAE 2 / editable
```

Neither journal creates or writes `DB_PART_REV`. Formal changes still use the
normal NX/Teamcenter revise UI.

## Targeting rule

- One or more preselected Assembly Navigator component rows: process only
  their unique, loaded prototypes.
- Selected geometry is normalized through `OwningComponent`; a selected
  managed part may also be used directly.
- If NX reports no selection, or none of the reported selections resolves to
  a component/managed-part target, process only the active NX work part. This
  handles stale face/body/root selections that otherwise suppress fallback.
- A selection always excludes the active parent assembly unless it is itself
  explicitly opened and targeted with no selection.
- Selecting a subassembly targets only that subassembly CAD file; descendants
  are never traversed.
- Repeated occurrences of one prototype are collapsed. J31 increments that
  prototype once, not once per occurrence.
- A suppressed, unloaded, or unmanaged component blocks the complete batch.
  An unrelated selection mixed with valid CAD targets also blocks the batch;
  unrelated selections are ignored only when active-work-part fallback is
  used.
- Audit JSON records the runtime type and resolution of every NX-selected
  object in `selected_objects`.

## Complete-batch preflight

Before any mutation, both journals resolve every target, query the configured
workflows, read checkout ownership, read display/internal release status, and
validate `DB_PART_REV` plus the exact positive-integer `WAEItem/WAE_VERSION`.

If any unique target fails preflight, the result is `BLOCKED_BATCH` and nothing
is saved, checked in/out, frozen/unfrozen, or written. A non-freeze release
status also blocks the batch so these journals cannot bypass formal release.

## J30 Freeze

For every preflighted target:

1. An already frozen target must be checked in and read-only.
2. A checked-in but unfrozen target still requires the freeze workflow;
   check-in alone is not considered frozen.
3. A checked-out target must be writable and owned by the current Teamcenter
   user. J30 saves it without changing revision or WAE, then batch-checks in
   only those checked-out targets.
4. J30 calls `PdmSession.AssignFreezeStatus(parts,
   "Part_Freeze_Process")` once for every non-frozen target in the batch.
5. Every final target must positively show a freeze status, `CHECKED_IN`,
   read-only, `IsModifiable()=False`, unchanged `DB_PART_REV`, and unchanged
   `WAE_VERSION`. `HasWriteAccess` remains diagnostic only: NX X 2506 runtime
   evidence shows it may remain true for a genuinely frozen/read-only part.

Only then does the batch report `ALL_TARGETS_FROZEN`.

## J31 Unfreeze

Every target must start positively frozen and `CHECKED_IN`. A checked-out
target blocks the complete batch, preventing a second WAE increment.

1. J31 calls `PdmSession.AssignUnfreezeStatus(parts,
   "Part_Unfreeze_Process")` once for the complete batch.
2. It verifies that the freeze status is removed without changing revision,
   WAE, or checkout state.
3. It checks out all unique targets in one batch and verifies that every target
   is writable, has write/modification access, and is owned by the current
   Teamcenter user.
4. For each unique prototype, J31 computes the next WAE value internally,
   writes exactly `current + 1`, rereads it, saves, and verifies it.
5. Every target remains checked out and unfrozen for CAD editing.

Only then does the batch report `ALL_TARGETS_UNFROZEN`.

## Runtime failure and recovery

Teamcenter workflow, check-in/out, and saved attribute changes do not form one
rollback transaction. If a mutation fails after execution begins, the journal
stops immediately and reports `RECOVERY_REQUIRED`, the exact `failed_stage`,
per-target operation flags, and fresh after-state snapshots. Operators must not
rerun blindly; inspect the JSON and recover the affected targets first.

## Output and modes

Both journals default to `APPLY` for their NX toolbar buttons. Set
`NX_J30_MODE=DRY_RUN` or `NX_J31_MODE=DRY_RUN` before starting NX for a
non-mutating complete-batch preflight.

J30 and J31 V4 require the shared helper build
`WAE-CHANGE-CONTROL-V4`. A partial deployment stops with a helper-version
mismatch instead of running older behavior under a newer journal build label.

Audit JSON is written beneath:

```text
<NX_JOURNALS_IO_DIR or Desktop>/NX_WAE_CHANGE_CONTROL
```

## NX X 2506 acceptance sequence

NX is not installed on this repository host. Local tests prove code and state
contracts only; the following disposable runtime test is required:

1. Deploy the complete `from_git` directory.
2. Open one disposable `Rev A / WAE 6` CAD part checked out by the operator.
3. Run J30 in `DRY_RUN`, then `APPLY`.
4. Require `ALL_TARGETS_FROZEN`, a positive freeze status, checked-in/read-only
   state, and unchanged `Rev A / WAE 6`.
5. Attempt normal NX checkout manually. It must be denied while frozen.
6. Run J31 in `DRY_RUN`, then `APPLY`.
7. Require `ALL_TARGETS_UNFROZEN`, no protected release status, checked-out and
   writable state, unchanged Rev A, and WAE 7.
8. Make and save a harmless CAD edit.
9. Repeat with two selected unique components, then with duplicate occurrences
   of one prototype; retain every JSON report as runtime evidence.

Do not use production CAD until the one-part freeze/checkout-denial/unfreeze
cycle is proven on the actual Teamcenter X tenant.
