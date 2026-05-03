# NX Open Python Journals

NX Open Python journals for **Siemens NX 2312** + Teamcenter productivity.
Run via **NX > Tools > Journal > Play** (`Alt+F8`). The journals target NX 2312 embedded Python 3.10 and avoid third-party Python packages.

## Journals

| # | File | Description |
|---|------|-------------|
| 01 | `01_hla_step_export.py` | Exports the active HLA part to STEP using `config/step_export.json` |
| 02 | `02_hla_multilevel_bom.py` | Traverses the assembly and exports a multilevel BOM CSV from NX part attributes |
| 03 | `03_batch_drawing_pdf.py` | Traverses unique prototype parts and exports drawing sheets to PDF |
| 04 | `04_assembly_attribute_audit.py` | Audits required attributes and writes audit/summary CSV reports |
| 05 | `05_bulk_attribute_updater.py` | **Pull/Push** - dumps NX attributes to CSV or writes Teamcenter CSV values back to empty NX attributes |

## Key Runtime Notes

- Deployment target: NX 2312 embedded Python 3.10.
- Required external Python packages: none.
- Config format: JSON.
- Report format: CSV with UTF-8 BOM so Excel opens it cleanly.
- Errors and summaries are written to the NX Listing Window because NX may run journals through `ugraf.exe`.

## Recommended Workflow

### First time - verify attribute names

```
Step 1  Open a representative part with Teamcenter attributes populated.
Step 2  Tools > Journal > Play -> utils/discover_attributes.py
        -> generates ATTR_DISCOVERY_<part>_<timestamp>.txt.
Step 3  Run J05 PULL on a representative part or assembly.
        -> generates PULL_<part>_<timestamp>.csv.
Step 4  Compare discovery/PULL output against config/attribute_mapping.json.
        Update JSON values for any names that differ.
```

### Ongoing - populate NX attributes from Teamcenter CSV

```
Step 1  Export the Teamcenter attribute sheet and save it as Att-*.csv.
        Keep the same two header rows.

Step 2  Run J05 PUSH and select the folder containing Att-*.csv.
        -> Matches parts by DB_PART_NO.
        -> Writes only to empty NX attributes.
        -> Never overwrites a non-empty NX attribute.
        -> Generates PUSH_REPORT_<timestamp>.csv.

Step 3  Review PUSH_REPORT_<timestamp>.csv, then spot-check values in NX.
```

## Attribute Mapping

`config/attribute_mapping.json` is the source of truth for attribute names.

```json
{
  "columns": {
    "PART_NUMBER": "DB_PART_NO",
    "DESCRIPTION": "DB_PART_NAME",
    "MATERIAL": "MATERIAL",
    "FINISH": "SURFACE_FINISH",
    "REVISION": "DB_PART_REV",
    "DRAWING_NUMBER": "DRAWING_NUMBER"
  }
}
```

- `columns` drives J02 BOM output and J04 audit output.
- `tc_columns` maps Teamcenter CSV row-2 aliases to NX internal attribute names for J05 PUSH.
- `skip_columns` documents Teamcenter columns that must not be written to NX.

## Output File Naming

| Journal | Output pattern |
|---------|---------------|
| J01 | `<DB_PART_NO>_REV<DB_PART_REV>.stp` |
| J02 | `BOM_<DB_PART_NO>_<timestamp>.csv` |
| J03 | `<drawing_number>_REV<revision>.pdf` |
| J04 | `AUDIT_<DB_PART_NO>_<timestamp>.csv` and `AUDIT_SUMMARY_<DB_PART_NO>_<timestamp>.csv` |
| J05 PULL | `PULL_<part_name>_<timestamp>.csv` |
| J05 PUSH | `PUSH_REPORT_<timestamp>.csv` |

## Notes

- All journals operate directly on NX part files through `GetUserAttribute` and `SetUserAttribute`.
- No Teamcenter connection is made at journal runtime.
- Legacy parts may have `PART_NUMBER` / `REVISION`; journals fall back to those when TC names are missing.
- Edit `config/step_export.json` to control STEP version and output naming for J01.
