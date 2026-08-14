# J24 — Guarded HLA Isolate-View Visibility Repair

Use `from_git/journals/24_repair_hla_isolate_visibility.py` after J23 proves
that mapped occurrence geometry is absent from an active view named `Isolate`.

NXOpen does not provide an `ExitIsolation` method. The supported assembly APIs
are `CreateIsolateViewWithComponents`, `ShowComponentsInIsolateView`, and
`HideComponentsInIsolateView`. J24 uses the least disruptive option: it adds
only the selected missing occurrence and its complete subtree to the current
isolate view.

## Safety contract

- Requires exactly one target occurrence.
- Refuses to run unless the active work-view name contains `Isolate`.
- Refuses to run when mapped target geometry is unavailable or already visible.
- Creates one visible NX undo mark before changing the view.
- Does not blank, unblank, suppress, unsuppress, load, save, or modify
  Teamcenter data.
- If the API throws or produces no mapped-body visibility change, it attempts
  an automatic rollback.
- On success, the display change remains so the operator can inspect it. Use
  one **Undo** to restore the previous view state.

## Run in NX X 2506

1. Pull the latest `master`.
2. Open the affected top-level HLA as both displayed and work part.
3. Keep the failing main window and its `Isolate` view active.
4. In Assembly Navigator, preselect exactly the missing
   `264MN031978A01/A` occurrence.
5. Play `from_git/journals/24_repair_hla_isolate_visibility.py`.
6. Check whether the missing subtree appears.
7. Push the generated `J24_ISOLATE_REPAIR_*.json` from
   `Desktop\NX_HLA_VISIBILITY_DIAGNOSTIC\`.

Do not use **Isolate in New Window** for this run; J24 must execute against the
failing HLA work view.

## Verdict interpretation

- `CONFIRMED / ISOLATE_VIEW_MEMBERSHIP_EXCLUDED_TARGET`: the supported isolate
  API changed exact mapped-target visibility from zero to a positive count.
  Isolation membership is causally confirmed.
- `API_ERROR / SHOW_COMPONENTS_IN_ISOLATE_VIEW_FAILED`: NX rejected the
  operation. The JSON contains the exception and rollback status.
- `INCONCLUSIVE / ISOLATE_SHOW_DID_NOT_RESTORE_MAPPED_GEOMETRY`: the API call
  completed, but the exact target bodies remained absent. J24 attempts rollback;
  the next branch is a controlled layout/work-view replacement diagnostic.
- `NOT_APPLIED`: a safety precondition failed; read the reported root-cause code.

## Current evidence baseline

The NX X 2506 J23 V2 artifact for component tag `69623` established:

- selected occurrence: `264MN031978A01/A`;
- mapped target-subtree bodies: 36;
- mapped target bodies visible in `Isolate`: 0;
- unsuppressed mapped-but-absent rows: 12;
- blanked subtree rows: 0;
- hidden HLA-layer rows: 0;
- visible dynamic sections: 0.

J24 turns the remaining isolation hypothesis into a direct before/after causal
test instead of inferring it from the view name.
