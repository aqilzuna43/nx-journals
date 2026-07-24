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
04_assembly_attribute_audit.py Assembly attribute audit reports
05_bulk_attribute_updater.py   Pull/push attribute CSV workflow
06_auto_pdf_step_export.py     Active work part STEP + drawing PDF export
07_datapack_pdf_step_export.py DataPack-controlled assembly PDF + STEP export
08_list_loaded_drawings.py     Loaded-drawing Teamcenter identity probe
09_test_teamcenter_specification_open.py Closed-drawing specification-open test
10_test_step_export.py         STEP export and body-validation diagnostic
```

`05_bulk_attribute_updater.py` is intentionally self-contained to avoid NX2312
package/import path problems. The J02/J04/J05 reconciliation workflow reads
`config/attribute_reconciliation.json`; its production save policy remains
`NO_SAVE` until the disposable-item runtime gate is approved.

The other journals still use shared helpers from `utils`, so keep the full
folder together if you run J01-J04.

J04 also requires `NX_DRAWING_SCOPE.csv` in `NX_JOURNALS_IO_DIR` (or Desktop).
If NX stages the selected journal outside this folder, set `NX_JOURNALS_ROOT`
to the complete `from_git` folder or its repository parent before starting NX.

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
