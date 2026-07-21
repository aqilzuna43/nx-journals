# NX Open Python Journals

NX Open Python journals for **Siemens NX 2312** + Teamcenter productivity.
Run via **NX > Tools > Journal > Play** (`Alt+F8`). The deployable runtime is the `from_git/` folder, which targets NX 2312 embedded Python 3.10 and avoids third-party Python packages.

## Deployment Layout

Copy or pull the whole `from_git/` folder to the office PC. The folder must keep this shape:

```text
from_git/
  config/
  journals/
  templates/
  utils/
```

In NX, browse to journals inside that folder, for example:

```text
...\from_git\journals\05_bulk_attribute_updater.py
```

For J05, the production updater is self-contained to avoid NX2312 import-path failures, but it still reads `from_git\config\attribute_mapping.json`. Keep `config` beside `journals`.

For J01-J04, keep the full `from_git` folder together because those journals still use shared helpers from `from_git\utils`.

## Journals

| # | File | Description |
|---|------|-------------|
| 01 | `from_git/journals/01_hla_step_export.py` | Exports the active work part to STEP |
| 02 | `from_git/journals/02_hla_multilevel_bom.py` | Traverses the assembly and exports a multilevel BOM CSV from NX part attributes |
| 03 | `from_git/journals/03_batch_drawing_pdf.py` | Traverses unique prototype parts and exports drawing sheets to PDF |
| 04 | `from_git/journals/04_assembly_attribute_audit.py` | Audits required attributes and writes audit/summary CSV reports |
| 05 | `from_git/journals/05_bulk_attribute_updater.py` | **Pull/Push** - dumps NX attributes to CSV or writes Teamcenter CSV values back to empty NX attributes |
| 06 | `from_git/journals/06_auto_pdf_step_export.py` | Exports the active work part to STEP and its drawing sheets to PDF in one run |
| 07 | `from_git/journals/07_datapack_pdf_step_export.py` | Exports DataPack-controlled drawing PDFs and AP214 STEP files from the loaded assembly |

## Key Runtime Notes

- Deployment target: NX 2312 embedded Python 3.10.
- Required external Python packages: none.
- Config format: JSON.
- Report format: CSV with UTF-8 BOM so Excel opens it cleanly.
- Errors and summaries are written to the NX Listing Window because NX may run journals through `ugraf.exe`.
- NX 2312 does not expose the folder picker API used by older journals, so these scripts follow the known-good `Export_BOM.py` pattern and use the Desktop by default:
  - Input CSV files: `%USERPROFILE%\Desktop`
  - Generated reports/STEP/PDF files: `%USERPROFILE%\Desktop`
- To use a shared or custom location, set `NX_JOURNALS_IO_DIR` before launching NX.

## Journal 07 - DataPack PDF + STEP Export

Journal 07 reads a manually prepared DataPack scope and exports only the PDF
and STEP outputs explicitly enabled in that CSV. It matches each request by the
normalized combination of `DB_PART_NO` and `DB_PART_REV`; it does not decide
which parts are BTP, determine drawing readiness, search Teamcenter, or select a
different revision.

### Prepare the input

1. Refresh and filter the FZ-PowerSystem DataPack tracker to the required BTP
   scope.
2. Export or copy the selected rows to CSV and confirm the `PDF` and `STEP`
   controls.
3. Use `from_git/templates/NX_EXPORT_SCOPE_TEMPLATE.csv` as the starting
   format, then save the working file with the exact name
   `NX_EXPORT_SCOPE.csv`.
4. Close Excel after saving the CSV.
5. Put the file in `NX_JOURNALS_IO_DIR` when configured, or on the current
   user's Desktop otherwise. Journal 07 never searches for the "latest" CSV.

Required logical columns and accepted aliases:

| Logical value | Accepted headers |
|---|---|
| Part number | `DB_PART_NO`, `Item Number`, `PART_NUMBER`, `Part Number` |
| Revision | `DB_PART_REV`, `Item Rev`, `REVISION`, `Revision` |
| PDF control | `PDF`, `Export_PDF`, `EXPORT_PDF` |
| STEP control | `STEP`, `Export_STEP`, `EXPORT_STEP` |

Optional traceability columns are `DATA_PACK_STATUS`/`Status`,
`PRIMARY_MODULE`/`Primary Module`, `PART_DESCRIPTION`/`Part Description`, and
`OWNER`/`Owner`. Enabled controls are `YES`, `Y`, `TRUE`, `1`, or `X`;
disabled controls are blank, `NO`, `N`, `FALSE`, or `0`. An unknown nonblank
control is reported as a warning and treated as disabled. Rows with both
controls explicitly disabled are ignored. Duplicate part/revision rows are
merged, with PDF and STEP enabled when any contributing row requests them.

### Prepare NX and run

Before playing the journal, open the correct top-level HLA assembly in managed
NX 2312, apply the intended Teamcenter revision rule, fully load the required
components, and confirm the expected revisions in Assembly Navigator. Then run:

```text
NX > Tools > Journal > Play
from_git\journals\07_datapack_pdf_step_export.py
```

The first version can use only prototype parts already available through the
loaded assembly. It does not query Teamcenter, open missing revisions, save or
modify NX parts, create datasets, or upload generated files.

### Journal 07 outputs

Each run creates an audit-preserving folder:

```text
<I/O root>\NX_BULK_EXPORT\YYYYMMDD_HHMMSS\
  PDF\
  STEP\
  REPORTS\EXPORT_RESULT_YYYYMMDD_HHMMSS.csv
  LOGS\EXPORT_LOG_YYYYMMDD_HHMMSS.txt
```

STEP files use `<DB_PART_NO>_REV<DB_PART_REV>.stp` and AP214. PDF files use
`DRAWING_NUMBER` when available, otherwise the requested part number, and add a
deterministic `_SHEET01`, `_SHEET02`, and so on for multi-sheet drawings.
Journal 07 creates one PDF per drawing sheet; it does not create a combined PDF.

The UTF-8-BOM result CSV contains one row per valid unique request plus each
invalid input row. Principal results are `SUCCESS`, `PARTIAL_SUCCESS`,
`NOT_REQUESTED`, `SKIPPED_NO_DRAWING`, `NOT_FOUND`, `REVISION_MISMATCH`,
`INVALID_INPUT`, `FAILED`, and `FAILED_NO_OUTPUT_FILE`. PDF and STEP outcomes
are independent, and the expected file must exist before an export is recorded
as successful. A valid-header CSV containing only invalid or ignored rows still
produces a report but performs no conversion.

The NX Listing Window shows progress, traversal diagnostics, collisions, and a
final file-count summary. Journal 07 restores the original display and work
parts even when an individual export fails.

## Recommended Workflow

### First time - verify attribute names

```
Step 1  Open a representative part with Teamcenter attributes populated.
Step 2  Tools > Journal > Play -> from_git/utils/discover_attributes.py
        -> generates ATTR_DISCOVERY_<part>_<timestamp>.txt on Desktop.
Step 3  Run J05 PULL on a representative part or assembly.
        -> generates PULL_<part>_<timestamp>.csv on Desktop.
Step 4  Compare discovery/PULL output against from_git/config/attribute_mapping.json.
        Update JSON values for any names that differ.
```

### Ongoing - populate NX attributes from Teamcenter CSV

```
Step 1  Export the Teamcenter attribute sheet and save it as Att-*.csv.
        Keep the same two header rows.
        Put the CSV on Desktop.

Step 2  Run J05 PUSH.
        -> Matches parts by DB_PART_NO.
        -> Writes only to empty NX attributes.
        -> Never overwrites a non-empty NX attribute.
        -> Generates PUSH_REPORT_<timestamp>.csv on Desktop.

Step 3  Review PUSH_REPORT_<timestamp>.csv, then spot-check values in NX.
```

## Attribute Mapping

`from_git/config/attribute_mapping.json` is the source of truth for attribute names.

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
| J06 | STEP: `<DB_PART_NO>_REV<DB_PART_REV>.stp`; PDF: `<DRAWING_NUMBER>_REV<revision>.pdf` |
| J07 | `NX_BULK_EXPORT\<timestamp>\PDF`, `STEP`, `REPORTS`, and `LOGS` |

## Notes

- All journals operate directly on NX part files through `GetUserAttribute` and `SetUserAttribute`.
- No Teamcenter connection is made at journal runtime.
- Legacy parts may have `PART_NUMBER` / `REVISION`; journals fall back to those when TC names are missing.
- J01 exports the currently open work part as AP214 STEP and names the file from `DB_PART_NO` / `DB_PART_REV` when available.
- J06 combines the J01 STEP path and active-part drawing PDF export into one no-prompt journal. It writes files to the configured output folder and does not create Teamcenter datasets.
- J07 is self-contained and needs no shared utility or JSON configuration file. It only processes exact part-number/revision matches already loaded under the active assembly.
