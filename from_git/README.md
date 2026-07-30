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
```

J04, J05, and J11 are intentionally self-contained to avoid NX2312
package/import path problems. They read
`config/attribute_reconciliation.json`; J05 production saving remains
The Journal 05 save gate is enabled as `SAVE_CHANGED_PARTS`.

Other journals still use shared helpers from `utils`, so keep the full folder
together.

J04 reads unique 3D master prototypes only. It produces one editable
`NX_ATTRIBUTE_UPDATE_*.csv` and a required `.baseline.json` sidecar. It does
not inspect drawings, require drawing scope, certify a BOM, or modify NX.

For normal NX use, open J05 and edit only the two user settings near its top:
paste the edited J04 CSV path into `USER_UPDATE_CSV` and leave
`USER_MODE = "DRY_RUN"` for the first run. PowerShell is not required.
`NX_ATTRIBUTE_UPDATE_FILE` and `NX_J05_MODE` remain optional overrides for
automation. An approved row authorizes every changed business field on that
row; identity, material, mass, dimensions, lifecycle, and quantity cannot be
changed. J05 rejects blank replacements and stale baselines.

J05 explicitly checks out all affected Teamcenter prototypes before writing,
aborts without attribute changes if any checkout fails, rereads each write,
saves successful targets, leaves successful checkouts checked out for review,
and never performs check-in. NX X `@DB/...` identities are treated as managed
even when `Session.IsManagedMode` incorrectly reports false. J11 remains
available in `PROBE` and guarded `FULL_REVERSIBLE` modes for runtime diagnosis.

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
model first and then the drawing. Missing `WAE_VERSION` produces a
revision-only watermark such as `DRAFT_A` plus a report warning; filenames do
not change. PDF text is converted to polylines so NX catalog symbols remain
visible; the resulting PDF text is not searchable or selectable. Every page
also receives the same run-level `EXPORTED: YYYY-MM-DD HH:MM MYT` footer.
J07 creates the footer as a temporary drafting note, exports one combined PDF,
then undoes the notes. Timing for drawing resolution, note preparation, PDF
commit, and cleanup is written to the log.

J18 measures every face of each direct traditional solid body in the active
work part, including hidden bodies. It ignores sheet and convergent bodies and
writes per-body rows plus a fail-closed total in square metres under
`NX_SURFACE_AREA`. J18 does not traverse assemblies, save NX data, or calculate
paint weight.

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
Journal build: J07-NX2506-PDF-POLYLINES-TIMESTAMP-V5
Drawing resolver: canonical Teamcenter specification identifier
```

J12 is the read-only PDF rendering diagnostic. Display the affected drawing
and run it once to store the canonical target and create five `PRELOADED`
comparison PDFs. Close the drawing completely and rerun J12 to create the
matching `CLOSED_AUTO` matrix through `Parts.OpenDisplay`. Results are written
under `NX_PDF_DIAGNOSTIC`. J12 does not change layer states, update or save the
drawing, and closes only the drawing it opened.

No third-party Python packages are required.
