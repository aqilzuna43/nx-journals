# NX 2312 / NX X 2506 Runtime Folder

This folder is the deployable NX journal payload.

Keep these folders together:

```text
from_git/
  config/
  journals/
  templates/
  utils/
```

Run journals from `from_git/journals` in NX 2312 or NX X 2506.

Available production journals:

```text
01_hla_step_export.py          Active work part STEP export
02_hla_multilevel_bom.py       Multilevel assembly BOM CSV
03_batch_drawing_pdf.py        Assembly drawing PDF batch export
04_assembly_attribute_audit.py Read-only 3D model business-attribute pull
05_bulk_attribute_updater.py   Approved CSV update + checkout workflow
06_auto_pdf_step_export.py     Active work part STEP + drawing PDF export
07_datapack_pdf_step_export.py DataPack-controlled assembly PDF + STEP export
08_list_loaded_drawings.py     Loaded-drawing Teamcenter identity probe
09_test_teamcenter_specification_open.py Closed-drawing specification-open test
10_test_step_export.py         STEP export and body-validation diagnostic
11_test_teamcenter_attribute_checkout.py Guarded checkout/save acceptance
12_diagnose_pdf_watermark_symbols.py PDF watermark/catalog-symbol matrix
13_test_teamcenter_part_name.py Disposable Teamcenter Item Name rename test
14_bulk_part_name_updater.py   Approved bulk Teamcenter Item Name update
15_tc_offline_drawing_workflow.py Native offline drawing export workflow
16_tc_offline_drawing_import.py Verified existing-specification drawing import
17_tc_master_drawing_import.py Separate master-drawing import workflow
18_work_part_surface_area.py Active-work-part solid surface-area CSV
19_test_teamcenter_drawing_import_contract.py Read-only J16 runtime contract probe
20_diagnose_assembly_full_load.py Component-by-component Full Load failure diagnostic
21_mass_surface_attribute_updater.py Bottom-up native roll-up mass/area updater
22_diagnose_mass_attribute_write.py One-part write-mechanism diagnostic
25_tc_single_drawing_cleanup.py Guarded reduction to one drawing specification
26_move_solid_bodies_to_layer_1.py Guarded active-part solid-body layer migration
27_move_assembly_components_to_layer_1.py Guarded direct-component layer migration
28_probe_bom_structure.py      Memory-safe targeted BoM-control checkpoint
29_set_selected_component_bom_exclusion.py Atomic custom BoM-subtree exclusion + native Reference-Only untick
30_cad_freeze.py              Freeze one selected component at its current WAE version
31_cad_unfreeze.py            Unfreeze one selected component and increment WAE_VERSION
32_probe_wae_freeze_capability.py Read-only NX/Teamcenter WAE lock API inventory
```

J04, J05, and J11 are intentionally self-contained to avoid NX2312
package/import path problems. They read
`config/attribute_reconciliation.json`; J05 production saving remains
J21 is also self-contained and follows the same BoM visibility filter as
`NXOpenBoMExtended.py` (suppressed, reference-only, and keyword-named
occurrences are excluded). APPLY fully discovers that subtree in memory, then
scans every unique prototype and runs NX's native Mass Properties update only
when the exact `NX_MassPropRollupMass` title is absent. Existing values,
including `0.0`, are not made work parts, inspected for checkout, or saved.
Load failures block missing dependent assemblies while safe sibling branches
continue. J21 performs no checkout and no direct reserved-attribute write;
selected non-writable targets are skipped and reported. Use `REFRESH_ALL` for
V5's full bottom-up rebuild with an all-or-nothing load gate. NX stores area in
mm^2; the report also presents converted m^2.
The Journal 05 save gate is enabled as `SAVE_CHANGED_PARTS`.

Other journals still use shared helpers from `utils`, so keep the full folder
together.

J04 reads unique 3D master prototypes from the active assembly state only.
Suppressed occurrences and their complete subtrees are excluded, and only
BoM-visible models are pulled (same filter as `NXOpenBoMExtended.py`):
reference-only members and keyword-named occurrences (CSYS, COORDINATE,
DATUM, REFERENCE, SKELETON) are excluded too. It produces one editable
`NX_ATTRIBUTE_UPDATE_*.csv` and a required `.baseline.json` sidecar. It does
not inspect drawings, require drawing scope, certify a BOM, or modify NX.

For normal NX use, open J05 and edit only the two user settings near its top:
paste the edited J04 CSV path into `USER_UPDATE_CSV` and leave
`USER_MODE = "DRY_RUN"` for the first run. PowerShell is not required.
`NX_ATTRIBUTE_UPDATE_FILE` and `NX_J05_MODE` remain optional overrides for
automation. An approved row authorizes every changed business field on that
row; identity, material, mass, dimensions, lifecycle, and quantity cannot be
changed. J05 rejects blank replacements and stale baselines.
`NO_CHANGE` is informational. Reruns also treat a live value that already
matches the approved replacement as `ALREADY_AT_EXPECTED_VALUE`, without a
checkout or write. Only a live third value remains stale.

J05 explicitly checks out all affected Teamcenter prototypes before writing,
using one session-wide status snapshot and one batch checkout for the unique
approved targets that are not already checked out. A target checked out by
another user is reported as `CHECKOUT_FAILED` and skipped without writes;
independently verified writable targets continue. J05 rereads each write, saves
successful targets, leaves successful checkouts checked out for review, and
never performs check-in. For large assemblies, manually pre-check out only the
approved target parts, not the complete assembly; J05 reuses those checkouts.
NX X `@DB/...` identities are treated as managed even when
`Session.IsManagedMode` incorrectly reports false. J11 remains available in
`PROBE` and guarded `FULL_REVERSIBLE` modes for runtime diagnosis.

J06 also uses the shared helpers from `utils`. It writes the active work part
STEP file and active-part drawing PDF files to `NX_JOURNALS_IO_DIR` when set,
or to the user's Desktop by default. It does not create Teamcenter datasets.

J07 is self-contained. Prepare `NX_EXPORT_SCOPE.csv` from
`templates/NX_EXPORT_SCOPE_TEMPLATE.csv`, place it in `NX_JOURNALS_IO_DIR` or
on the user's Desktop, fully load the required parts under the active NX
assembly, and play `journals/07_datapack_pdf_step_export.py`. It matches exact
part-number/revision pairs and writes a timestamped `NX_BULK_EXPORT` run with
one combined multipage PDF per drawing, AP214 STEP files, a UTF-8-BOM result
CSV, and a text log.

Every J07 PDF page receives the native NX watermark
`DRAFT_<revision>.<WAE_VERSION>`. The value is read from the already-loaded
model first and then the drawing; STEP reads it from the master part. Output
filenames embed the version after the revision —
`<number>_REV<revision>.<WAE_VERSION>` for both PDF (multi-drawing PDFs
append `_DWG<n>`) and STEP (`.stp`). Missing `WAE_VERSION` produces a
revision-only watermark such as `DRAFT_A`, keeps the revision-only filenames,
and records a report warning. Drawing words, numbers, and the run-level
`EXPORTED: YYYY-MM-DD HH:MM MYT` footer remain searchable/selectable PDF text.
The large draft value uses the normal NX PDF watermark and may also be
searchable. J07 creates the footer as a temporary drafting note, exports one
combined PDF, then undoes the note. Timing for drawing resolution, timestamp
preparation, PDF commit, and cleanup is written to the log.

J18 measures every face of each direct traditional solid body in the active
work part, including hidden bodies. It ignores sheet and convergent bodies and
writes per-body rows plus a fail-closed total in square metres under
`NX_SURFACE_AREA`. J18 does not traverse assemblies, save NX data, or calculate
paint weight.

J20 is for assemblies that fail only when **Full Load** is requested. Run it
while the failing top-level assembly is still partially or minimally loaded.
It snapshots occurrence paths first, calls `LoadThisPartFully()` once per
unique prototype, then runs a final assembly-wide `LoadFully()` verification.
It writes a UTF-8-BOM CSV plus
`NX_Assembly_Load_Diagnostic_Report.txt` under a timestamped
`NX_ASSEMBLY_FULL_LOAD_DIAGNOSTIC` folder. J20 changes the in-memory load state
but never saves, checks out, checks in, replaces, suppresses, or closes parts.

J07 accepts the documented DataPack header aliases and PDF/STEP values, merges
duplicate part/revision requests, reports invalid input, missing parts, and
revision mismatches, and verifies that each expected export file exists. It
reuses loaded drawings or opens the canonical Teamcenter specification
`@DB/<part>/<revision>/specification/<part>-<revision>-dwg<n>`. It does not
search for a different revision, modify or save NX parts, or require a JSON
configuration file.

For the NX X 2506 closed-drawing test, deploy this complete folder, fully close
`264MN028607A01/A/dwg1`, and run
`journals/09_test_teamcenter_specification_open.py`. Require three returned
sheets, the canonical `/specification/` identifier, and
`FINAL STATUS: SUCCESS`. Journal 07 must identify the deployed implementation
with:

```text
Journal build: J07-NX2506-SEARCHABLE-TEXT-NATIVE-WATERMARK-V8
Drawing resolver: canonical Teamcenter specification identifier
```

J12 is the read-only PDF rendering diagnostic. Display the affected drawing
and run it once to store the canonical target and create five `PRELOADED`
comparison PDFs. Close the drawing completely and rerun J12 to create the
matching `CLOSED_AUTO` matrix through `Parts.OpenDisplay`. Results are written
under `NX_PDF_DIAGNOSTIC`. J12 does not change layer states, update or save the
drawing, and closes only the drawing it opened.

J30 and J31 are separate NX X 2506 UI-button journals for one preselected
Assembly Navigator component. J30 saves and checks in only its loaded
prototype, preserving `WAE_VERSION`. J31 requires that prototype to start
checked in, checks out only that prototype, increments `WAE_VERSION` by exactly
one, verifies and saves it, then leaves it checked out for CAD editing. Both
journals read but never write `DB_PART_REV`, and neither scans or modifies the
BoM. See `docs/J30_J31_CAD_FREEZE_UNFREEZE.md`.

J25 handles the migration case where one 3D Item/Revision has `dwg1`, `dwg2`,
and other non-master drawing specifications but only one final DWG may remain.
It defaults to `DRY_RUN` and requires an explicit keep index plus an exact live
extra-index list. `APPLY_APPROVED` backs up and hashes every extra drawing's
associated files, then uses `DeleteExistingAttachedFiles(..., False)` to remove
the files and empty drawing dataset. This is destructive dataset removal, not a
relation-only detach. See `docs/J25_SINGLE_DRAWING_CLEANUP.md` and start from
`templates/NX_TC_SINGLE_DRAWING_SCOPE_TEMPLATE.csv`.

J26 prevents layer-filtered Teamcenter STEP and JT output from omitting imported
solid geometry. It scans only direct bodies owned by the active NX work part,
including blanked bodies and bodies on hidden layers. Traditional solid bodies
are eligible for layer 1; sheet, convergent, and other body types are reported
but never moved. A parent assembly may remain displayed while a component is
the work part. J26 never traverses or modifies other component prototypes.

J26 defaults to `USER_MODE = "DRY_RUN"`. Play it once and review the paired CSV
and JSON under `NX_LAYER_1_MIGRATION` on `NX_JOURNALS_IO_DIR` or the Desktop.
For a Teamcenter part, manually check out the active work part before changing
the setting near the top of the journal to `USER_MODE = "APPLY"`; J26 requires
a known checkout owned by the current Teamcenter user and never checks out,
checks in, or saves. APPLY uses one visible NX undo mark, batch-moves the
off-layer solids, and verifies every body, blanked state, work layer, and all
256 layer states. Any move, verification, or paired-evidence failure triggers
rollback. After `APPLIED_VERIFIED`, inspect the part, use one Undo if needed,
and save manually only when the result is correct. A successful local test run
is not NX proof: retain the office-machine CSV/JSON and confirm the downstream
Teamcenter STEP and JT contain every expected solid before declaring J26
runtime-verified.

J27 complements J26 at assembly level. It requires the target assembly to be
both Work Part and Displayed Part, then scans only the component occurrences
placed directly under that assembly root. It includes blanked, suppressed,
reference-only, non-geometric, lightweight, and unloaded occurrences without
forcing any prototype to load. For a direct subassembly, J27 changes only that
subassembly occurrence's parent-level layer option; it does not recurse into or
modify the subassembly's internal child placements or any prototype body.

J27 defaults to `USER_MODE = "DRY_RUN"` and writes paired CSV/JSON evidence
under `NX_ASSEMBLY_LAYER_1_MIGRATION`. Review DRY_RUN first. For a Teamcenter
assembly, manually check out only the parent assembly, set
`USER_MODE = "APPLY"`, and rerun. APPLY uses
`Component.SetLayerOption(1)` under one visible undo mark and verifies direct
child membership, occurrence identity, prototype identity, suppression,
blanking, reference set, non-geometric state, position/orientation, work layer,
and all 256 layer states. Any mutation, verification, or paired-evidence
failure triggers full rollback. J27 never loads, checks out, checks in, or
saves NX data. After `APPLIED_VERIFIED`, inspect the assembly, use one Undo if
needed, and save manually only when correct. Run J27 separately with a nested
subassembly as Work and Displayed Part only when its own direct children also
need normalization. Retain the office-machine evidence and verify the customer
Teamcenter STEP/JT before declaring runtime success.

No third-party Python packages are required.
