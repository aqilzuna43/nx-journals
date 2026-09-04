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
- Repeated occurrences and distinct loaded NX proxies for the same
  case-insensitive `DB_PART_NO + DB_PART_REV` are collapsed. J31 increments
  that Teamcenter identity once, not once per occurrence or proxy.
- For J30, an unrelated selection mixed with valid CAD targets is skipped and
  reported while valid targets continue. J31 keeps complete-selection
  preflight and blocks before mutation if any selected target is invalid.
- Audit JSON records the runtime type and resolution of every NX-selected
  object in `selected_objects`.

## WAE lifecycle and preflight

Before any mutation, both journals resolve every target, query the configured
workflow per exact identity, read checkout ownership, read display/internal
release status, and validate `DB_PART_REV` plus `WAEItem/WAE_VERSION`.

- J30 accepts a positive whole-number working WAE or an alphabetic final WAE
  matching `DB_PART_REV`, such as `Rev E / WAE E`. Missing, malformed, and
  mismatched values are blocked for that target. Other safe targets continue.
- J31 accepts only positive whole-number working WAE values. A matching
  alphabetic value produces `BLOCKED_FINAL_RELEASE_BASELINE`; that baseline
  remains Frozen and the next engineering change must use normal Teamcenter
  revision control.
- J31 retains complete-selection preflight: any invalid target produces
  `BLOCKED_BATCH` before any selected target is changed.

A non-freeze controlled status remains a safety block, so neither journal can
bypass formal release.

## J30 Freeze

For every preflighted target:

1. An already frozen target must be checked in and read-only.
2. A checked-in but unfrozen target still requires the freeze workflow;
   check-in alone is not considered frozen.
3. A checked-out target must be writable and owned by the current Teamcenter
   user. J30 saves and checks in that one target without changing revision or
   WAE.
4. J30 calls `PdmSession.AssignFreezeStatus([part],
   "Part_Freeze_Process")` independently for each non-frozen exact identity.
5. Every final target must positively show a freeze status, `CHECKED_IN`,
   read-only, `IsModifiable()=False`, unchanged `DB_PART_REV`, and unchanged
   `WAE_VERSION`. `HasWriteAccess` remains diagnostic only: NX X 2506 runtime
   evidence shows it may remain true for a genuinely frozen/read-only part.

One target's Teamcenter failure does not stop later safe targets. A positively
verified final state reports `FROZEN`; if the API raised an error but every
postcondition passed, it reports `FROZEN_WITH_WARNING`. A target that did not
reach Frozen reports `FAILED_FREEZE_WORKFLOW`. The overall result is
`ALL_TARGETS_FROZEN`, `PARTIAL_COMPLETION`, or `NO_TARGETS_COMPLETED`.

## J31 Unfreeze

Every target must start positively frozen and `CHECKED_IN`. A checked-out
target blocks the complete batch, preventing a second WAE increment.

1. J31 calls `PdmSession.AssignUnfreezeStatus([part],
   "Part_Unfreeze_Process")` for one exact identity.
2. It verifies that the freeze status is removed without changing revision,
   WAE, or checkout state.
3. It checks out that identity and verifies it is writable, has
   write/modification access, and is owned by the current Teamcenter user.
4. J31 computes the next WAE value internally,
   writes exactly `current + 1`, rereads it, saves, and verifies it.
5. Only after the target completes does J31 proceed to the next identity.
   Completed targets remain checked out and unfrozen for CAD editing.

Only then does the batch report `ALL_TARGETS_UNFROZEN`.

## Runtime failure and recovery

Teamcenter workflow, check-in/out, and saved attribute changes do not form one
rollback transaction. J30 isolates and reports a failed identity, then
continues. J31 continues after a clean failure only when the target is
positively verified to remain unchanged and Frozen. If a J31 failure leaves an
unfrozen, checked-out, incremented, or otherwise incomplete/unknown state, it
reports `RECOVERY_REQUIRED` and marks later targets
`NOT_ATTEMPTED_AFTER_RECOVERY_REQUIRED`. Operators must not rerun blindly;
inspect the JSON and recover the affected target first.

For either journal, an API exception followed by fully successful final-state
verification is recorded as `FROZEN_WITH_WARNING` or
`UNFROZEN_WITH_WARNING`; no repeat operation is required solely because the
API raised.

## Output and modes

Both journals default to `APPLY` for their NX toolbar buttons. Set
`NX_J30_MODE=DRY_RUN` or `NX_J31_MODE=DRY_RUN` before starting NX for a
non-mutating complete-batch preflight.

J30 and J31 V5 require the shared helper build
`WAE-CHANGE-CONTROL-V5`. A partial deployment stops with a helper-version
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
