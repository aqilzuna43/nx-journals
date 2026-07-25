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
```

J04, J05, and J11 are intentionally self-contained to avoid NX2312
package/import path problems. They read
`config/attribute_reconciliation.json`; J05 production saving remains
`NO_SAVE` until the disposable-item Journal 11 runtime gate is approved.

Other journals still use shared helpers from `utils`, so keep the full folder
together.

J04 reads unique 3D master prototypes only. It produces one editable
`NX_ATTRIBUTE_UPDATE_*.csv` and a required `.baseline.json` sidecar. It does
not inspect drawings, require drawing scope, certify a BOM, or modify NX.

Set `NX_ATTRIBUTE_UPDATE_FILE` to the edited J04 CSV before running J05. Use
`NX_J05_MODE=DRY_RUN` first. An approved row authorizes every changed business
field on that row; identity, material, mass, dimensions, lifecycle, and
quantity cannot be changed. J05 rejects blank replacements and stale
baselines.

Before enabling `SAVE_CHANGED_PARTS`, run J11 in its default read-only `PROBE`
mode and then `FULL_REVERSIBLE` on an explicitly identified disposable item.
J05 explicitly checks out all affected prototypes before writing, aborts
without attribute changes if any checkout fails, leaves successful checkouts
checked out for review, and never performs check-in.

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
Journal build: J07-NX2506-CANONICAL-SPEC-OPEN-V1
Drawing resolver: canonical Teamcenter specification identifier
```

No third-party Python packages are required.
