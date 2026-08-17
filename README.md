# NX Open Python Journals

NX Open Python journals for **Siemens NX 2312 and NX X 2506** + Teamcenter productivity.
Run via **NX > Tools > Journal > Play** (`Alt+F8`). The deployable runtime is
the `from_git/` folder, supports the embedded Python runtimes in both NX
versions, and avoids third-party Python packages.

The repository also contains a standalone NXOpen VB.NET assembly-load
troubleshooter at
`Assembly/Diagnostic/NX_Assembly_Load_Diagnostic.vb`. Use it to identify the
exact occurrence behind missing-file, unavailable-prototype, unloaded-part,
and invalid OM-object STEP failures. See `Assembly/Diagnostic/README.md` for
operation and report guidance.

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
| 12 | `from_git/journals/12_diagnose_pdf_watermark_symbols.py` | Compares PDF watermark, text, and NX catalog-symbol rendering settings |
| 13 | `from_git/journals/13_test_teamcenter_part_name.py` | Disposable Teamcenter Item Name rename test |
| 14 | `from_git/journals/14_bulk_part_name_updater.py` | Approved bulk Teamcenter Item Name update |
| 15 | `from_git/journals/15_tc_offline_drawing_workflow.py` | Exports native drawing packages for controlled offline editing |
| 16 | `from_git/journals/16_tc_offline_drawing_import.py` | Checkout-gated drawing import with exact-target re-export verification |
| 17 | `from_git/journals/17_tc_master_drawing_import.py` | Verified creation of a missing drawing specification beneath an existing 3D revision |
| 18 | `from_git/journals/18_work_part_surface_area.py` | Read-only active-work-part solid surface-area CSV |
| 19 | `from_git/journals/19_test_teamcenter_drawing_import_contract.py` | Read-only J16 checkout/export/runtime contract probe |
| 20 | `from_git/journals/20_diagnose_assembly_full_load.py` | Isolates component/prototype failures that occur only during assembly Full Load |
| 21 | `from_git/journals/21_mass_surface_attribute_updater.py` | Requires the active BoM-visible subtree to be fully loaded, then triggers NX's native mass-properties update separately on every unique leaf part and subassembly in bottom-up order. It never checks out or directly writes reserved attributes; non-writable Teamcenter targets are skipped and reported while writable targets are saved and verified by read-back. DRY_RUN / SMOKE / PROBE modes. |
| 22 | `from_git/journals/22_diagnose_mass_attribute_write.py` | One-part diagnostic: tests classic compute, the native mass-properties builder, and per-category attribute writes, with full before/after attribute dumps |
| 23 | `from_git/journals/23_diagnose_hla_visibility.py` | Read-only, exact-target HLA visibility proof: tri-state NX probes, subtree/reference-set occurrence mapping, cross-view and same-prototype controls, hypothesis verdicts, and fact-cited root-cause conclusions |
| 24 | `from_git/journals/24_repair_hla_isolate_visibility.py` | Guarded display-only causal test and repair: uses the NX Python one-input isolate API, tests the selected parent then mapped unsuppressed descendants, compares active/returned/layout views, records exact before/after JSON evidence, and provides an NX undo mark without saving |

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

## Journal 17 - Create a New Drawing Specification

Use J17 when a local NX drawing must become a new `/specification/` dataset
beneath an existing Teamcenter 3D Item/Revision. Use J16 instead when replacing
an existing drawing specification.

Start a fresh NX X 2506 managed session with no parts loaded. Copy
`from_git/templates/NX_TC_NEW_DRAWING_SPECIFICATION_TEMPLATE.csv` to the
configured I/O folder and complete one row per missing specification:

```csv
PART_NUMBER,REVISION,DWG_INDEX,SOURCE_DRAWING_FILE,APPROVED,ENGINEER
264MN021262A01,A,3,C:\drawings\local_drawing.prt,YES,AQIL
```

The example creates exactly
`@DB/264MN021262A01/A/specification/264MN021262A01-A-dwg3`. The parent
`@DB/264MN021262A01/A` must already exist and be checked in; the exact `dwg3`
specification must not already exist.

Run `from_git/journals/17_tc_master_drawing_import.py` once. Its production
default is `APPLY_APPROVED`; the managed checks and UF Clone dry run happen
inside that same run. J17 stages the drawing with the proven Teamcenter
AutoTranslate specification encoding, quarantines existing or ambiguous
destinations, checked-out 3D masters, stale local files, and other prewrite
failures with zero writes. Every discovered reference defaults to
`UseExisting`; only the exact staged specification receives `Overwrite` for
creation of its new managed dataset.

Require `SPECIFICATION_CREATED_VERIFIED` or
`SPECIFICATION_CREATED_VERIFIED_MANAGED_TRANSFORM` in the result CSV. The
report must also show `PRESERVE_3D_UNCHANGED=YES`, the exact new specification
identifier, at least one drawing sheet, the expected managed native `.prt`, and
a post-import `CHECKED_IN` state. If the result is `MANUAL_CHECKIN_REQUIRED`,
do not rerun J17: verify the new specification and manually check it in only
when the checkout belongs to you. J17 never checks in automatically.

## Journal 04 + 05 - Business Attribute Update

1. Open and fully load the required 3D assembly.
2. Run **J04**. It pulls only the BoM-visible 3D master models (the same
   filter as `NXOpenBoMExtended.py`): suppressed occurrences, reference-only
   members, and CSYS/datum/skeleton/keyword-named occurrences are all
   excluded together with their subtrees. It then creates
   `NX_ATTRIBUTE_UPDATE_*.csv` and a matching `.baseline.json`.
3. Edit the CSV business fields only. For rows to update, set `APPROVED=YES` and fill `ENGINEER`.
4. In **J05**, set:

```python
USER_UPDATE_CSV = r"C:\path\to\NX_ATTRIBUTE_UPDATE_....csv"
USER_MODE = "DRY_RUN"
```

5. Run J05 and resolve all errors in the report.
6. When clean, change `USER_MODE` to `"APPLY_APPROVED"` and run again.

J05 checks out affected Teamcenter parts, writes and verifies the approved attributes, saves them, and leaves them checked out. It never checks parts in automatically. Part Number, Part Name, and Revision are not changed by J05.
`NO_CHANGE` and `ALREADY_AT_EXPECTED_VALUE` are successful informational
results. A true `STALE_BASELINE_VALUE` still blocks an overwrite when the live
value matches neither the J04 baseline nor the approved replacement.

For large assemblies, J05 takes one session-wide checkout snapshot and uses
one batch checkout for the unique approved targets that are not already
checked out. You may manually check out the approved target parts first; do
not check out every part in the assembly. The Listing Window reports phase and
per-target timings during apply. A target checked out by another user is
reported as `CHECKOUT_FAILED` and skipped; independently verified writable
targets continue through update, verification, and save.

## Journal 14 - Teamcenter Part Name Update

Use J14 only for **Item Name / Part Name** changes. It does not change Part Number, Revision, description, or geometry.

Prepare a CSV with these columns:

```csv
PART_NUMBER,CURRENT_PART_NAME,NEW_PART_NAME,APPROVED,ENGINEER,APPROVAL_NOTE
264MN000000A01,OLD NAME,NEW NAME,YES,AQIL,Name correction
```

Then set J14:

```python
USER_PART_NAME_CSV = r"C:\path\to\part_name_update.csv"
USER_MODE = "DRY_RUN"
```

1. Run J14 in `DRY_RUN` and review the report.
2. If clean, change `USER_MODE` to `"APPLY_APPROVED"` and run again.
3. Confirm `UPDATED_VERIFIED` in the result CSV.

`CURRENT_PART_NAME` is the stale-value safety check. J14 uses `UF_UGMGR.SetPartNameDesc()` and verifies the name and description after each update. A managed NX/Teamcenter session is required; the target part does not need to be open.

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
does not search for another revision, save NX parts, create datasets, or upload
generated files. Temporary PDF timestamp notes are removed through an NX undo
mark before J07 continues.

The listing window must identify the current deployment before export:

```text
Journal build: J07-NX2506-SEARCHABLE-TEXT-NATIVE-WATERMARK-V8
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

Every Journal 07 PDF page receives the native NX draft watermark:

```text
DRAFT_<DB_PART_REV>.<WAE_VERSION>
```

Each page also receives one small bottom-right export timestamp:

```text
EXPORTED: YYYY-MM-DD HH:MM MYT
```

J07 creates the footer as a temporary native drafting note on every sheet,
exports the existing combined multipage PDF, and undoes the notes immediately
afterwards. It logs drawing-resolution, timestamp preparation, PDF commit, and
cleanup timings. A timestamp-cleanup failure stops later PDF work but does not
prevent independently requested STEP exports.

Journal 07 reads `WAE_VERSION` from the already-loaded model first and then
from the drawing (STEP reads it from the master part). Output filenames embed
the version after the revision: `<number>_REV<revision>.<WAE_VERSION>.pdf`
(for example `264MN021888A01_REVA.2.pdf`), with multi-drawing PDFs appending
`_DWG<n>` after the version; STEP files use the same pattern with a `.stp`
extension. If `WAE_VERSION` is unavailable, the PDF is still exported with a
revision-only watermark such as `DRAFT_A`, both outputs keep their
revision-only filenames, and the result report records a warning. Journal 07
keeps drawing text as PDF text, so drawing words, numbers, and the export
footer are searchable and selectable without OCR. The large `DRAFT_<revision>.<WAE_VERSION>` value uses
the normal NX PDF watermark feature and may also be searchable.
`CustomSymbolsInForeground` is enabled so drawing symbols remain visible.

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
6. Require the Journal 07 build and resolver banners, one multipage PDF with
   the draft watermark and catalog symbols visible on every page, successful
   STEP body validation, and restored display/work parts.
7. Repeat Journal 07 with the drawing preloaded and compare the resulting PDF.

Journal 09 can be redirected without editing the file by setting
`NX_TEST_PART_NO`, `NX_TEST_PART_REV`, `NX_TEST_DWG_INDEX`, or
`NX_TEST_EXPECTED_SHEET_COUNT` in the NX environment.

### Journal 12 PDF symbol diagnostic

Use Journal 12 when symbols such as omega or pi are visible in NX but absent
from the exported PDF:

1. Display the affected drawing and run Journal 12.
2. Inspect the five PDFs in the timestamped `PRELOADED` folder.
3. Close the drawing completely and run Journal 12 again.
4. Inspect the five PDFs in the `CLOSED_AUTO` folder.
5. Return both logs and identify the first variant containing both the
   watermark and all catalog symbols.

The first run stores the canonical drawing identity in
`NX_PDF_DIAGNOSTIC\LAST_TARGET.json`; the second run reuses it through
`session.Parts.OpenDisplay`. The journal inventories sheets, views, and
non-empty object layers but never changes layer state, updates, or saves the
drawing. It restores the original display/work parts and closes only a drawing
it opened.

## Output File Naming

| Journal | Output pattern |
|---------|---------------|
| J01 | `<DB_PART_NO>_REV<DB_PART_REV>.stp` |
| J02 | `BOM_<DB_PART_NO>_<timestamp>.csv` |
| J03 | `<drawing_number>_REV<revision>.pdf` |
| J04 | `NX_ATTRIBUTE_UPDATE_<root>_<timestamp>.csv` and matching `.baseline.json` |
| J05 | `J05_<DRY_RUN-or-APPLY_APPROVED>_<timestamp>.csv` |
| J06 | STEP: `<DB_PART_NO>_REV<DB_PART_REV>.stp`; PDF: `<DRAWING_NUMBER>_REV<revision>.pdf` |
| J07 | `NX_BULK_EXPORT\<timestamp>\PDF\<number>_REV<rev>.<WAE_VERSION>.pdf`, `STEP\<number>_REV<rev>.<WAE_VERSION>.stp`, plus `REPORTS` and `LOGS` |
| J11 | `J11_CHECKOUT_ACCEPTANCE_<timestamp>.json` |
| J12 | `NX_PDF_DIAGNOSTIC\<timestamp>_<PRELOADED-or-CLOSED_AUTO>\*.pdf` |
| J14 | `J14_PART_NAME_<DRY_RUN-or-APPLY_APPROVED>_<timestamp>.csv` |
| J18 | `NX_SURFACE_AREA\J18_SURFACE_AREA_<part>_<timestamp>.csv` |
| J21 | `NX_MASS_SURFACE_UPDATE\J21_MASS_SURFACE_<root>_<timestamp>.csv` |
| J22 | `NX_MASS_SURFACE_UPDATE\J22_DIAGNOSTIC_<root>_<timestamp>.csv` and `.json` |
| J23 | `NX_HLA_VISIBILITY_DIAGNOSTIC\J23_EVIDENCE_<target>_<timestamp>.csv` and `.json` |
| J25 | `NX_TC_SINGLE_DRAWING_CLEANUP\<timestamp>\` with CSV, JSON, log, and `BACKUP\` |

## Notes

- Journals use the active NX Teamcenter connection and do not create a separate Teamcenter login.
- J04 reads prototype attributes from BoM-visible occurrences only (same
  filter as `NXOpenBoMExtended.py`: suppressed, reference-only, and
  CSYS/datum/skeleton/keyword-named occurrences are excluded), so its values
  correspond to the 3D master objects J05 updates.
- J05 requires exact `DB_PART_NO + DB_PART_REV`; it never changes Part Name or Revision and never performs automatic check-in.
- J14 changes Teamcenter Item Name only and verifies the result through UF_UGMGR read-back.
- J01 exports the currently open work part as AP214 STEP and names the file from `DB_PART_NO` / `DB_PART_REV` when available.
- J06 combines the J01 STEP path and active-part drawing PDF export into one no-prompt journal. It writes files to the configured output folder and does not create Teamcenter datasets.
- J07 is self-contained and needs no shared utility or JSON configuration file. It processes exact part-number/revision matches already loaded under the active assembly and can open their canonical drawing specifications.
- J18 measures every face of direct traditional solid bodies in the active work part, including hidden bodies. It reports square metres only and intentionally contains no paint-weight calculation.
- J21 requires every unique BoM-visible prototype below the active work assembly to be fully loaded before APPLY begins. It processes deepest leaves first, then subassemblies, and the active assembly last. For each writable target it temporarily makes that part the work part and drives NX's native `CreateMassPropertiesBuilder` path (`UpdateOnSave` + `UpdateNow` + `Commit`), saves, and reads back `NX_MassPropRollupMass` (kg) / `NX_MassPropRollupArea` (mm^2). J21 never calls a checkout API and never writes the reserved titles directly. Checked-in, read-only, and other-user checkout targets are skipped and identified in the CSV while processing continues; the original display/work context is restored. `NX_J21_MODE=DRY_RUN` only reports current values and access state, while `NX_J21_MODE=SMOKE` updates the active work part only.
- J22 is a fast one-part diagnostic (run on a disposable part): it tests the classic compute APIs, the native mass-properties builder (CreateMassPropertiesBuilder + UpdateOnSave + UpdateNow + Commit), and per-category AttributePropertiesBuilder writes, and dumps all attributes before/after so the working mechanism and category on a given NX build are visible.
- J25 reduces an exact 3D Item/Revision to one explicitly retained `dwg<n>` specification. It defaults to `DRY_RUN`; apply requires approval, an exact live extra-index list, and backup of every removed payload. NXOpen has no relation-only detach in this API boundary, so J25 removes each extra empty drawing dataset with `DeleteExistingAttachedFiles(..., False)` and then verifies that only the retained drawing remains.
