# J30/J31 — Selected-component CAD Freeze and Unfreeze

These are two separate NX X 2506 UI-button entry journals backed by the shared
`from_git/utils/wae_change_control.py` helper. They act on exactly one
preselected component prototype and never traverse or modify the BoM.

## Controlled lifecycle

```text
Rev A / WAE 1 / CHECKED_IN
  -> press J31 CAD Unfreeze
  -> checkout selected prototype
  -> WAE_VERSION 1 -> 2
  -> reread and save
  -> Rev A / WAE 2 / CHECKED_OUT (ready for CAD editing)

After editing:
  -> press J30 CAD Freeze
  -> save selected prototype
  -> check in selected prototype
  -> Rev A / WAE 2 / CHECKED_IN (frozen baseline)
```

Only the selected component prototype is targeted. `DB_PART_REV` is read and
verified but never written. Formal Teamcenter revision changes continue through
the normal NX/TCX UI.

## J30 CAD Freeze

File: `from_git/journals/30_cad_freeze.py`

- Requires exactly one preselected, loaded, Teamcenter-managed component.
- If already checked in, reports `FROZEN_VERIFIED` without mutation.
- Otherwise requires checkout ownership by the current Teamcenter user.
- Saves and checks in only the selected prototype.
- Verifies `CHECKED_IN`, unchanged `DB_PART_REV`, and unchanged `WAE_VERSION`.

## J31 CAD Unfreeze

File: `from_git/journals/31_cad_unfreeze.py`

- Requires exactly one preselected, loaded, Teamcenter-managed component.
- Requires the component to start positively `CHECKED_IN`.
- Computes the next value internally; operators cannot enter an arbitrary value.
- Checks out only the selected prototype with secondary inclusion disabled.
- Writes only `WAEItem/WAE_VERSION`, rereads it, and saves the prototype.
- Leaves the component checked out for CAD modification.
- Blocks reruns while the component remains checked out, preventing a second increment.

Both journals default to `APPLY` for direct NX button use. For a non-mutating
preflight, set `NX_J30_MODE=DRY_RUN` or `NX_J31_MODE=DRY_RUN` before launching NX.

Audit JSON files are written beneath:

```text
<NX_JOURNALS_IO_DIR or Desktop>/NX_WAE_CHANGE_CONTROL
```

## Initial NX acceptance

NX is not installed on this repository host. Before normal production use:

1. Deploy the complete `from_git` folder to the NX X 2506 machine.
2. Use a disposable Rev A / WAE 1 component.
3. Run J30 and J31 first with their mode temporarily set to `DRY_RUN`.
4. Run J31 in `APPLY`; require `UNFROZEN_READY_FOR_EDIT`, WAE 2, and checked out.
5. Make a harmless CAD edit, then run J30; require `FROZEN_CHECKED_IN`, WAE 2,
   and checked in.
6. Retain both JSON files as runtime proof before production rollout.
