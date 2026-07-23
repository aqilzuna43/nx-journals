# NX2312 Runtime Folder

This is the complete deployable journal payload. Keep the folders together:

```text
from_git/
  config/
  journals/
  templates/
  utils/
```

The reconciliation pipeline treats NX/Teamcenter as authoritative:

```text
02_hla_multilevel_bom.py        FZ-compatible draft BOM
04_assembly_attribute_audit.py Read-only, fail-closed certification
05_bulk_attribute_updater.py   PULL / DRY_RUN / APPLY_APPROVED corrections
```

J02 and J04 use `utils/attribute_reconciliation.py`. J05 is self-contained for
NX2312 deployment but consumes the same
`config/attribute_reconciliation.json` rule model.

J04 requires `NX_DRAWING_SCOPE.csv` in `NX_JOURNALS_IO_DIR` (or Desktop) and
accepts an optional `NX_ATTRIBUTE_MASTER_REFERENCE.csv` only for downstream
drift detection. J05 consumes `NX_ATTRIBUTE_CORRECTIONS.csv`; templates are in
this folder's `templates` directory.

J05 defaults to `DRY_RUN`, and the versioned configuration remains `NO_SAVE`
until the disposable Teamcenter-item save/reopen proof is explicitly approved.
MASTER and BOM values are never authoritative correction sources.

Other journals:

```text
01_hla_step_export.py            Active work-part STEP export
03_batch_drawing_pdf.py          Legacy assembly PDF batch
06_auto_pdf_step_export.py       Legacy active-part STEP + PDF
07_datapack_pdf_step_export.py   Protected DataPack PDF + STEP export
08_nx_x_attribute_diagnostic.py  Diagnostic-only probe
09_managed_drawing_open_test.py  Canonical drawing-open regression
```

J07 remains self-contained and unchanged by the reconciliation work. It reads
`NX_EXPORT_SCOPE.csv`, matches exact part/revision pairs already present in the
loaded assembly, writes a timestamped `NX_BULK_EXPORT` tree, and never modifies
or saves NX parts.

No third-party Python packages are required.
