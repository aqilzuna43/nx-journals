# NX2312 Runtime Folder

This folder is the deployable NX journal payload.

Keep these folders together:

```text
from_git/
  config/
  journals/
  templates/
  utils/
```

Run journals from `from_git/journals` in NX 2312.

Available production journals:

```text
01_hla_step_export.py          Active work part STEP export
02_hla_multilevel_bom.py       Multilevel assembly BOM CSV
03_batch_drawing_pdf.py        Assembly drawing PDF batch export
04_assembly_attribute_audit.py Assembly attribute audit reports
05_bulk_attribute_updater.py   Pull/push attribute CSV workflow
06_auto_pdf_step_export.py     Active work part STEP + drawing PDF export
```

`05_bulk_attribute_updater.py` is intentionally self-contained to avoid NX2312
package/import path problems. It still reads `config/attribute_mapping.json`,
so keep the `config` folder beside `journals`.

The other journals still use shared helpers from `utils`, so keep the full
folder together if you run J01-J04.

J06 also uses the shared helpers from `utils`. It writes the active work part
STEP file and active-part drawing PDF files to `NX_JOURNALS_IO_DIR` when set,
or to the user's Desktop by default. It does not create Teamcenter datasets.

No third-party Python packages are required.
