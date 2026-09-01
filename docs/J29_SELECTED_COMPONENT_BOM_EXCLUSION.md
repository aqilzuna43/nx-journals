# J29 — Selected Component BoM Exclusion

## Corrected purpose

J29 marks one or more selected component occurrences so the Celestica custom
BoM workflows exclude each selected occurrence and its complete descendant
subtree. The marker is deliberately separate from NX's native Reference-Only
control:

```text
CELESTICA_BOM_EXCLUDE_SUBTREE = "YES"
```

It is a direct string instance attribute on the selected occurrence. It does
not mark the prototype part, so another placement of the same part remains in
the BoM unless that occurrence is separately marked.

J29 does not create, delete, or modify `REFERENCE_COMPONENT`. When JT geometry
must remain visible, the operator unticks Reference-Only through the standard
NX UI. J29 records the native control before and after only to prove that it
did not change.

## Consumers

The exact custom marker prunes the selected occurrence and its subtree in:

- `NXOpenBoMExtended.py`;
- J04 assembly attribute audit; and
- J21 mass/surface roll-up updater.

Those workflows continue to support genuine native `REFERENCE_COMPONENT` and
`PLIST_IGNORE_MEMBER` controls. The custom marker is active only when its exact
direct string value is `YES`; blank, `NO`, inherited, or otherwise nonstandard
values are not accepted by J29.

## Atomic batch contract

J29:

- defaults to `APPLY`;
- accepts 1–100 selected direct child component rows by default;
- supports a positive `NX_J29_MAX_SELECTION` override;
- preflights every target before writing any target;
- blocks the complete batch for invalid, nested, suppressed, duplicated,
  unreadable, PLIST-conflicting, or nonstandard custom-marker selections;
- treats an exact existing custom marker as `ALREADY_BOM_EXCLUDED` and applies
  only the remaining eligible targets;
- requires the parent assembly to be both Work Part and Displayed Part;
- requires an already-writable native parent, or a managed parent already
  checked out by the current Teamcenter user;
- creates one visible undo mark for the complete batch;
- writes and rereads the exact custom occurrence marker;
- verifies that `REFERENCE_COMPONENT` and the PLIST controls did not change;
- rolls back the complete batch if any write, verification, or paired evidence
  write fails; and
- never loads the tree, checks out, checks in, or saves.

There is no include/reversal mode. Remove the custom marker through the normal
NX attribute UI when that occurrence must return to the custom BoM.

## Run sequence

1. Open the parent assembly and make it both Work Part and Displayed Part.
2. Manually check out the parent assembly in Teamcenter.
3. For JT visibility, use the standard NX UI to untick Reference-Only on the
   affected component occurrences.
4. In Assembly Navigator, select 1–100 direct child component rows. Do not
   select geometry in the graphics window.
5. Play `29_set_selected_component_bom_exclusion.py`.
6. Require `APPLIED_VERIFIED` or `ALREADY_BOM_EXCLUDED`.
7. Inspect the CSV and JSON beneath
   `%USERPROFILE%\Desktop\NX_BOM_EXCLUSION`, or beneath
   `NX_JOURNALS_IO_DIR\NX_BOM_EXCLUSION` when configured.
8. Inspect the NX state and custom BoM. The parent remains modified but unsaved.
9. Save manually only if correct. Otherwise use one NX Undo:
   `J29 Set selected component BoM exclusions`.

Set `NX_J29_MODE=DRY_RUN` for a no-write batch preflight.

## Main verdicts

- `DRY_RUN_READY`: the complete batch passed; NX was not changed.
- `APPLIED_VERIFIED`: every new custom marker was written and verified.
- `ALREADY_BOM_EXCLUDED`: every target already had the exact marker.
- `BLOCKED_BATCH`: at least one per-target preflight failed; nothing changed.
- `BLOCKED_SELECTION_LIMIT`: the configured cap was exceeded.
- `BLOCKED_WRITE_ACCESS`: parent write access was not proven.
- `ROLLED_BACK`: a failure occurred and all baselines were restored.
- `ROLLBACK_FAILED`: complete restoration could not be proven; inspect and use
  NX Undo immediately.

## Evidence history and verification boundary

The 2026-09-02 V2 evidence proved that multi-selection, same-Teamcenter-user
matching, occurrence writes, read-back, and undo-mark creation work in NX X
2506. It also proved that writing native `REFERENCE_COMPONENT` produced the
wrong operational outcome for JT visibility. V3 preserves the proven batch
mechanics but writes only the custom BoM marker.

Local tests verify marker scope, exact value, native-control immutability,
all-or-nothing preflight, rollback, per-target reports, and matching subtree
pruning across the exporter, J04, and J21. Siemens NX is unavailable on this
host; Aqil's next NX X 2506 CSV/JSON and Listing Window output remain the
runtime proof.
