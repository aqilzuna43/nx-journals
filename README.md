# NX Open Python Journals

NX Open Python journals for **Siemens NX 2312 and NX X 2506** + Teamcenter productivity.
Run via **NX > Tools > Journal > Play** (`Alt+F8`). The deployable runtime is
the `from_git/` folder, supports the embedded Python runtimes in both NX
versions, and avoids third-party Python packages.

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

J04, J05, and J11 are self-contained to avoid NX journal import-path
failures, but they read `from_git\config\attribute_reconciliation.json`. Keep
`config` beside `journals`.

Keep the full `from_git` folder together because other journals still use
shared helpers from `from_git\utils`.

## Journals

| # | File | Description |
|---|------|-------------|
| 01 | `from_git/journals/01_hla_step_export.py` | Exports the active work part to STEP |
| 02 | `from_git/journals/02_hla_multilevel_bom.py` | Exports an NX-authoritative draft multilevel BOM |
| 03 | `from_git/journals/03_batch_drawing_pdf.py` | Traverses unique prototype parts and exports drawing sheets to PDF |
| 04 | `from_git/journals/04_assembly_attribute_audit.py` | Read-only 3D master business-attribute pull |
| 05 | `from_git/journals/05_bulk_attribute_updater.py` | Approved CSV update with stale-value and checkout gates |
| 06 | `from_git/journals/06_auto_pdf_step_export.py` | Exports the active work part to STEP and its drawing sheets to PDF in one run |
| 07 | `from_git/journals/07_datapack_pdf_step_export.py` | Exports DataPack-controlled drawing PDFs and AP214 STEP files from the loaded assembly |
| 08 | `from_git/journals/08_list_loaded_drawings.py` | Reports exact Teamcenter identities for drawings already loaded in NX |
| 09 | `from_git/journals/09_test_teamcenter_specification_open.py` | Tests automatic opening of one canonical Teamcenter drawing specification |
| 10 | `from_git/journals/10_test_step_export.py` | Diagnoses STEP export and validates body geometry |
| 11 | `from_git/journals/11_test_teamcenter_attribute_checkout.py` | Guarded Teamcenter checkout/save/reopen/restoration acceptance |

## Key Runtime Notes

- Deployment target: NX 2312 or NX X 2506 embedded Python.
- Required external Python packages: none.
- Config format: JSON.
- Report format: CSV with UTF-8 BOM so Excel opens it cleanly.
- Errors and summaries are written to the NX Listing Window because NX may run journals through `ugraf.exe`.
- NX 2312 does not expose the folder picker API used by older journals, so these scripts follow the known-good `Export_BOM.py` pattern and use the Desktop by default:
  - Input CSV files: `%USERPROFILE%\Desktop`
  - Generated reports/STEP/PDF files: `%USERPROFILE%\Desktop`
- To use a shared or custom location, set `NX_JOURNALS_IO_DIR` before launching NX.

## 3D Business-Attribute Round Trip

J04, J05, Extended BOM, and J11 use the exact XML-backed NX attribute titles.
The normal workflow is:

1. Open and fully load the intended 3D assembly.
2. Run J04. It reads prototype models only and creates
   `NX_ATTRIBUTE_UPDATE_<root>_<timestamp>.csv` with a matching
   `.baseline.json`.
3. Edit only business columns. Set `APPROVED=YES` and populate `ENGINEER` on
   each row to apply.
4. Set `NX_ATTRIBUTE_UPDATE_FILE` to that CSV and run J05 in `DRY_RUN`.
5. Resolve every stale, identity, controlled-value, permission, or checkout
   error before considering apply mode.

Part number, part name, revision, quantity, lifecycle, material, dimensions,
mass, and roll-up mass are read-only NX/CAD values. J05 can change only the
business allowlist in `attribute_reconciliation.json`, including
`WAE_VERSION`, `NX_FINISH`, commodity, traceability, service, manufacturer,
stocking, UOM, and export metadata. It rejects blank replacements.

J05 production saving remains disabled while `save_policy` is `NO_SAVE`.
Before changing that gate, run J11 `PROBE`, then its explicitly guarded
`FULL_REVERSIBLE` test on a disposable Teamcenter item. Apply mode explicitly
checks out every affected master part before changing anything, aborts the
batch if any checkout fails, and never checks a part in automatically.

See `docs/J04_J05_ATTRIBUTE_RECONCILIATION_PLAN.md` for CSV columns, checkout
guards, recovery behavior, and the J11 acceptance procedure.

## Journal 07 - DataPack PDF + STEP Export

Journal 07 reads a manually prepared DataPack scope and exports only the PDF
and STEP outputs explicitly enabled in that CSV. It matches each request by the
normalized combination of `DB_PART_NO` and `DB_PART_REV`; it does not decide
which parts are BTP, determine drawing readiness, or select a different
revision. It reuses a loaded drawing when possible and otherwise attempts to
open its canonical Teamcenter specification:

```text
@DB/<part>/<revision>/specification/<part>-<revision>-dwg<n>
```

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
NX 2312 or NX X 2506, apply the intended Teamcenter revision rule, fully load
the required components, and confirm the expected revisions in Assembly
Navigator. Then run:

```text
NX > Tools > Journal > Play
from_git\journals\07_datapack_pdf_step_export.py
```

The journal uses only prototype revisions already available through the loaded
assembly. It may open a drawing specification for that exact revision, but it
does not search for another revision, save or modify NX parts, create datasets,
or upload generated files.

The listing window must identify the current deployment before export:

```text
Journal build: J07-NX2506-CANONICAL-SPEC-OPEN-V1
Drawing resolver: canonical Teamcenter specification identifier
```

### Journal 07 outputs

Each run creates an audit-preserving folder:

```text
<I/O root>\NX_BULK_EXPORT\YYYYMMDD_HHMMSS\
  PDF\
  STEP\
  REPORTS\EXPORT_RESULT_YYYYMMDD_HHMMSS.csv
  LOGS\EXPORT_LOG_YYYYMMDD_HHMMSS.txt
```

STEP files use `<DB_PART_NO>_REV<DB_PART_REV>.stp` and AP214. Journal 07
creates one combined multipage PDF per resolved drawing. PDF files use
`DRAWING_NUMBER` when available, otherwise the requested part number. When
multiple drawing items resolve, the drawing index is appended as `_DWG<n>` to
avoid collisions.

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

### NX X 2506 closed-drawing acceptance

1. Deploy the complete current `from_git` directory; do not copy only Journal 07.
2. Close `264MN028607A01/A/dwg1` completely so it is absent from the NX session.
3. Run Journal 09 with its defaults.
4. Require the canonical `/specification/` identifier, `Drawing sheets returned: 3`, and `FINAL STATUS: SUCCESS`.
5. Run Journal 07 with PDF and STEP enabled.
6. Require the Journal 07 build and resolver banners, one multipage PDF, successful STEP body validation, and restored display/work parts.
7. Repeat Journal 07 with the drawing preloaded and compare the resulting PDF.

Journal 09 can be redirected without editing the file by setting
`NX_TEST_PART_NO`, `NX_TEST_PART_REV`, `NX_TEST_DWG_INDEX`, or
`NX_TEST_EXPECTED_SHEET_COUNT` in the NX environment.

## Output File Naming

| Journal | Output pattern |
|---------|---------------|
| J01 | `<DB_PART_NO>_REV<DB_PART_REV>.stp` |
| J02 | `BOM_<DB_PART_NO>_<timestamp>.csv` |
| J03 | `<drawing_number>_REV<revision>.pdf` |
| J04 | `NX_ATTRIBUTE_UPDATE_<root>_<timestamp>.csv` and matching `.baseline.json` |
| J05 | `J05_<DRY_RUN-or-APPLY_APPROVED>_<timestamp>.csv` |
| J06 | STEP: `<DB_PART_NO>_REV<DB_PART_REV>.stp`; PDF: `<DRAWING_NUMBER>_REV<revision>.pdf` |
| J07 | `NX_BULK_EXPORT\<timestamp>\PDF`, `STEP`, `REPORTS`, and `LOGS` |
| J11 | `J11_CHECKOUT_ACCEPTANCE_<timestamp>.json` |

## Notes

- Journals use the active NX Teamcenter connection and do not create a separate Teamcenter login.
- J04 and Extended BOM read prototype attributes, so their values correspond to the 3D master object J05 updates.
- J05 requires exact `DB_PART_NO + DB_PART_REV`; it has no legacy identity fallback.
- J05 never relies on implicit Teamcenter autolock and never performs automatic check-in.
- J01 exports the currently open work part as AP214 STEP and names the file from `DB_PART_NO` / `DB_PART_REV` when available.
- J06 combines the J01 STEP path and active-part drawing PDF export into one no-prompt journal. It writes files to the configured output folder and does not create Teamcenter datasets.
- J07 is self-contained and needs no shared utility or JSON configuration file. It processes exact part-number/revision matches already loaded under the active assembly and can open their canonical drawing specifications.
