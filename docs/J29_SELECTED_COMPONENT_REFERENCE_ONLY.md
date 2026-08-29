# J29 — Selected Component Reference-Only

## Purpose

J29 is a guarded, one-occurrence writer for preparing a component as
Reference-Only. It does not discover targets by part number and it does not
walk the assembly. The operator must preselect exactly one component row in
Assembly Navigator.

J28 V2 provided the runtime contract used by J29. In the accepted NX X 2506
result, sequence 177 (`264MN028171A01/A / 028060/A`) was a direct child of the
active assembly and was classified `REFERENCE_ONLY` because it owned this
single direct occurrence attribute:

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

## Safety contract

J29:

- defaults to `DRY_RUN`;
- accepts exactly one preselected NX assembly component;
- accepts only a direct child of the active assembly root, so the modified
  occurrence attribute is owned by the active parent assembly;
- requires the parent assembly to be both Work Part and Displayed Part;
- refuses suppressed components, root/nested selections, unreadable control
  attributes, and selections that are not component rows;
- refuses to combine `REFERENCE_COMPONENT` with either PLIST control;
- refuses to overwrite a nonstandard existing `REFERENCE_COMPONENT`;
- never walks the assembly tree or loads a component/prototype;
- never checks out, checks in, or saves;
- requires an already-writable native parent assembly, or a managed parent
  assembly already checked out by the current Teamcenter user;
- creates one visible NX undo mark before the write;
- writes `Component.SetInstanceUserAttribute("REFERENCE_COMPONENT", -1, "",
  NXOpen.Update.Option.Now)`;
- rereads the selected occurrence and verifies the exact J28 V2 contract;
- rolls back to the undo mark if the write, read-back, contract verification,
  or paired CSV/JSON evidence write fails;
- leaves a successful verified change unsaved and undoable for inspection.

The deferred `AtTermination` unload option is intentional. It follows the
memory-safe J28 V2 lifecycle and avoids invalidating NX bridge objects during
journal teardown.

## Run sequence

1. Open the parent assembly in NX and make it both Work Part and Displayed
   Part.
2. In Assembly Navigator, select exactly one direct child component row. Do
   not select its geometry in the graphics window.
3. Leave this setting in
   `from_git/journals/29_set_selected_component_reference_only.py`:

   ```python
   USER_MODE = "DRY_RUN"
   ```

4. Play J29 and require `DRY_RUN_READY`. Review the CSV and JSON under
   `%USERPROFILE%\Desktop\NX_REFERENCE_ONLY`, or under
   `NX_JOURNALS_IO_DIR\NX_REFERENCE_ONLY` when that environment variable is
   configured.
5. Ensure the parent assembly is writable. In Teamcenter, manually check out
   the parent assembly first; J29 will not do this.
6. Change `USER_MODE` to `"APPLY"`, keep the same component selected, and play
   J29 again.
7. Require `APPLIED_VERIFIED`. Inspect the Assembly Navigator/BOM behavior.
   The parent assembly is modified but unsaved.
8. Save manually only if the result is correct. Otherwise use one NX Undo:
   `J29 Set selected component Reference-Only`.

`NX_J29_MODE=DRY_RUN` or `NX_J29_MODE=APPLY` may override the file setting for
controlled automation.

## Verdicts

- `DRY_RUN_READY`: eligible; NX was not modified.
- `APPLIED_VERIFIED`: exact blank string occurrence attribute was reread and
  verified; the parent assembly remains unsaved under the visible undo mark.
- `ALREADY_REFERENCE_ONLY`: exact contract was already present; no write.
- `BLOCKED_SELECTION`: selection, depth, suppression, or occurrence reads did
  not pass the fail-closed gate.
- `BLOCKED_CONTROL_CONFLICT`: one or both native PLIST controls are present.
- `BLOCKED_NONSTANDARD_REFERENCE`: the existing reference attribute differs
  from the J28-proven contract.
- `BLOCKED_WRITE_ACCESS`: parent assembly write access was not proven.
- `BLOCKED_UNDO_MARK`: NX could not create the required undo mark.
- `ROLLED_BACK`: a failure occurred and the absent baseline was restored and
  verified.
- `ROLLBACK_FAILED`: restoration could not be proven; use NX Undo and inspect
  immediately.

## Verification boundary

Local tests verify mode parsing, exact API arguments, selection and access
gates, idempotence, conflicts, read-back verification, rollback, report
structure, and the absence of load/save/checkout calls. Siemens NX is not
installed on this repository host. Only Aqil's NX X 2506 CSV/JSON and Listing
Window result can prove the runtime behavior.
