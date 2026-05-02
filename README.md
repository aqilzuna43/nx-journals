# NX Open Python Journals

NX Open Python journals for NX 2312 + Teamcenter (TC) productivity.  
Run via **NX > Tools > Journal > Play** (`Alt+F8`). Each journal prompts for input paths at runtime.

## Journals

| # | File | Description |
|---|------|-------------|
| 01 | `01_hla_step_export.py` | Exports the active HLA part to STEP (AP214/AP242) using settings from `config/step_export.yaml` |
| 02 | `02_hla_multilevel_bom.py` | Traverses the full assembly tree and exports a multilevel BOM to Excel from NX part attributes |
| 03 | `03_batch_drawing_pdf.py` | Traverses the assembly, finds all associated drawing sheets, and batch-exports them to PDF |
| 04 | `04_assembly_attribute_audit.py` | Audits required part attributes across the assembly and flags missing/invalid values to Excel |
| 05 | `05_bulk_attribute_updater.py` | **Pull/Push** — dumps existing NX attributes to Excel (PULL), or writes TC attribute values back to NX (PUSH) |

---

## Single Source of Truth

The attribute schema is driven by the TC export file:

```
documents/input/Att-264MN021145A01_20Apr26.xlsx
```

This Excel file (generated from Teamcenter) defines:
- Which attributes exist on CAD Design, WAEItem, and WAEItemRevision objects
- The **display alias** shown in TC (row 2 of the export)
- The **internal NX attribute name** — shown in parentheses in the alias, e.g. `ID (DB_PART_NO)` → internal name is `DB_PART_NO`

All attribute mappings in `config/attribute_mapping.yaml` are derived from this file.  
Run **J05 PULL** to verify that the internal names in the yaml match what NX actually stores on real parts.

---

## How to Run

1. Open NX 2312 with the target assembly loaded as the work part.
2. Go to **Tools > Journal > Play** (or press `Alt+F8`).
3. Browse to the desired `.py` file in `journals/`.
4. The journal prompts for any required folder paths via NX dialogs at runtime.
5. Outputs are written to the folder you select.

---

## Recommended Workflow

### First time — verify attribute names

Before running any push, confirm that the internal NX attribute names in `attribute_mapping.yaml` match your actual parts:

```
Step 1  Open a representative part (one with TC attributes populated) in NX.
Step 2  Tools > Journal > Play → utils/discover_attributes.py
        → generates ATTR_DISCOVERY_*.txt listing every attribute on the part.
Step 3  Run J05 PULL on the open assembly
        → generates PULL_<assembly>_<timestamp>.xlsx listing current NX values.
Step 4  Compare PULL output and ATTR_DISCOVERY report against attribute_mapping.yaml.
        Update yaml values for any names that differ.
```

### Ongoing — populate TC attributes from TC export

```
Step 1  Export the TC attribute Excel from Teamcenter for your assembly
        (same format as documents/input/Att-264MN021145A01_20Apr26.xlsx).

Step 2  Run J05 PUSH, select the folder containing the TC Excel.
        → Reads TC Excel, matches parts by DB_PART_NO.
        → Reads existing NX attributes first — only writes to empty fields.
        → Never overwrites a non-empty attribute.
        → Generates PUSH_REPORT_<timestamp>.xlsx showing every decision:
             UPDATED   — TC value written to NX (was empty)
             KEPT      — NX already had a value, TC value ignored
             BOTH_EMPTY — both NX and TC are empty, no action
             NO_MATCH  — part number from TC not found in open assembly

Step 3  Review PUSH_REPORT.xlsx, then verify spot-check values in the NX attribute editor.
```

### BOM export and audit

```
J02  →  BOM_<part>_<timestamp>.xlsx     Multilevel BOM, one row per component
J04  →  AUDIT_<part>_<timestamp>.xlsx   Attribute audit, green = PASS, amber = FAIL
```

J02 and J04 read the same attribute names resolved through `config/attribute_mapping.yaml → columns`.  
Run J05 PUSH first to populate TC attributes, then J04 to verify completeness.

---

## Attribute Mapping Configuration

`config/attribute_mapping.yaml` has two sections:

### `columns` — used by J02 and J04

Maps the BOM/audit column header names to NX internal attribute names:

```yaml
columns:
  PART_NUMBER:    DB_PART_NO        # TC alias: "ID (DB_PART_NO)"
  DESCRIPTION:    DB_PART_NAME      # TC alias: "Name (DB_PART_NAME)"
  MATERIAL:       MATERIAL
  FINISH:         SURFACE_FINISH    # verify via discover_attributes
  REVISION:       DB_PART_REV       # TC alias: "Revision (DB_PART_REV)"
  DRAWING_NUMBER: DRAWING_NUMBER    # stored attribute on NX part — verify via discover_attributes
```

`DRAWING_NUMBER` is a stored user attribute on the NX part. It may differ from the part number (e.g. a shared drawing covers multiple PNs), so it is read directly via `GetUserAttribute` and is included in the J04 audit.

### `tc_columns` — used by J05

Maps the TC display alias (row 2 of the TC Excel) to the NX internal attribute name.  
Covers all 26 writable attributes across CAD Design, WAEItem, and WAEItemRevision groups.  
WAEItem/WAEItemRevision internal names are best-guess from the TC export — **verify with PULL mode before first push**.

### `skip_columns` — used by J05

TC columns that exist in the export but must never be written to NX (user IDs, read-only flags, computed values).

---

## Output File Naming

| Journal | Output pattern |
|---------|---------------|
| J01 | `<DB_PART_NO>_REV<DB_PART_REV>.stp` |
| J02 | `BOM_<DB_PART_NO>_<timestamp>.xlsx` |
| J04 | `AUDIT_<DB_PART_NO>_<timestamp>.xlsx` |
| J05 PULL | `PULL_<assembly_name>_<timestamp>.xlsx` |
| J05 PUSH | `PUSH_REPORT_<timestamp>.xlsx` (alongside TC Excel) |

---

## Dependencies

Install into NX's Python environment (`<NX_ROOT>\NXBIN\python.exe`):

```
python.exe -m pip install xlsxwriter pyyaml openpyxl
```

- **NXOpen / NXOpen.UF** — bundled with NX 2312
- **xlsxwriter** — Excel output (J02, J04, J05)
- **PyYAML** — attribute mapping config
- **openpyxl** — reading TC Excel in J05

---

## Notes

- All journals operate entirely within NX using `GetUserAttribute` / `SetUserAttribute` on NX part files. No connection to Teamcenter is made at runtime.
- The TC Excel used by J05 PUSH is read as a plain spreadsheet file — it supplies the desired attribute values but J05 writes them into NX, not back into TC.
- `DB_PART_NO` / `DB_PART_REV` are TC-propagated attributes stored on the NX part file. Legacy parts may have `PART_NUMBER` / `REVISION` instead — all journals fall back to these if the TC names are missing.
- All output files are excluded from version control via `.gitignore`.
- Edit `config/step_export.yaml` to control STEP version (AP214/AP242) and output naming for J01.
