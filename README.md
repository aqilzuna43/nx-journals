# NX Open Python Journals

NX Open Python journals for **Siemens NX 2312** and managed TeamcenterX.
Run them through **NX > Tools > Journal > Play** (`Alt+F8`). The deployable
runtime is the complete `from_git/` folder and uses only Python 3.10 standard
library modules plus NXOpen.

NX/Teamcenter is authoritative for the engineering BOM and configured
attributes. The FZ MASTER workbook is downstream reference evidence only and
is never used to overwrite NX.

## Deployment layout

Copy or pull the entire `from_git/` folder and preserve this shape:

```text
from_git/
  config/
  journals/
  templates/
  utils/
```

J02 and J04 consume the shared reconciliation engine. J05 is intentionally
self-contained for NX2312 journal loading, but reads
`config/attribute_reconciliation.json`. Keep the full folder together.

## Journals

| # | File | Role |
|---|---|---|
| 01 | `from_git/journals/01_hla_step_export.py` | Active work-part STEP export |
| 02 | `from_git/journals/02_hla_multilevel_bom.py` | NX-authoritative FZ-compatible **draft** BOM with immediate-parent quantities |
| 03 | `from_git/journals/03_batch_drawing_pdf.py` | Legacy drawing PDF batch export |
| 04 | `from_git/journals/04_assembly_attribute_audit.py` | Fail-closed, read-only NX certification and evidence reports |
| 05 | `from_git/journals/05_bulk_attribute_updater.py` | Approved typed correction workflow (`PULL`, `DRY_RUN`, `APPLY_APPROVED`) |
| 06 | `from_git/journals/06_auto_pdf_step_export.py` | Legacy active-part STEP and drawing PDF export |
| 07 | `from_git/journals/07_datapack_pdf_step_export.py` | Protected production DataPack PDF and STEP export |
| 08 | `from_git/journals/08_nx_x_attribute_diagnostic.py` | Diagnostic-only NX attribute probe |
| 09 | `from_git/journals/09_managed_drawing_open_test.py` | Canonical managed-drawing regression proof |

J07 is protected and is not part of the J02/J04/J05 reconciliation refactor.

## Runtime I/O

The default I/O directory is `%USERPROFILE%\Desktop`. Set
`NX_JOURNALS_IO_DIR` before launching NX to use a controlled directory.
Reports are UTF-8-BOM CSVs for Excel and progress is written to the NX Listing
Window.

Default reconciliation inputs:

| File | Use | Override |
|---|---|---|
| `NX_DRAWING_SCOPE.csv` | Required by J04 certification | `NX_DRAWING_SCOPE_FILE` |
| `NX_ATTRIBUTE_MASTER_REFERENCE.csv` | Optional downstream drift check | `NX_ATTRIBUTE_MASTER_FILE` |
| `NX_ATTRIBUTE_CORRECTIONS.csv` | J05 dry-run/apply input | `NX_ATTRIBUTE_CORRECTIONS_FILE` |

Templates for drawing scope and corrections are in `from_git/templates`.

## NX-authoritative BOM workflow

### J02 draft

Open and fully load the intended assembly, then play J02. It includes the
active root at level 0, excludes suppressed occurrences, aggregates duplicate
siblings, and preserves the same part under different parents.

Output:

```text
BOM_DRAFT_<root>_<timestamp>.csv
```

The exact FZ import headers are:

```text
BOM Level, DB_PART_NO, Indented Part Name, Component Name, Quantity,
DB_PART_DESC, DB_PART_NAME, DB_PART_REV, MFG, MPN, Stocking_Type
```

J02 output is always draft and performs no drawing or release certification.

### J04 certification

Prepare `NX_DRAWING_SCOPE.csv` from the template. `Drawing Required` must be
`YES` or `NO`; `REVIEW`, missing identities, invalid values, and conflicting
duplicates block certification.

J04 audits exact `category + attribute title`, model identity and hierarchy,
controlled values, placeholders, material completeness,
mass/density/volume consistency, exact transformed model-coordinate
dimensions, and applicable canonical drawings. It is strictly read-only.

Every run creates detailed and summary evidence. A certified BOM is created
only when the blocking count is zero:

```text
NX_CERTIFIED_BOM_<root>_<run_id>.csv
```

When an optional MASTER reference is supplied, any structural, revision,
quantity, or configured-field difference is `DOWNSTREAM_BOM_DRIFT` and means
MASTER must be regenerated. It is never an NX correction source.

### J05 controlled correction

Select mode through `NX_J05_MODE`:

```text
PULL
DRY_RUN          (default)
APPLY_APPROVED
```

Use `NX_ATTRIBUTE_CORRECTIONS.csv` from the template. Each row must identify
the exact part and revision, MODEL or DRAWING target, category, attribute,
type, audited raw value, expected value, audit run, engineer, and approval
evidence. The valid authorities are `ENGINEERING_APPROVAL` or a configured
`MODEL` mirror. BOM and MASTER authority are rejected.

J05 never writes identity, material, physical, dimension, system-owned,
locked, PDM-based, read-only, or unconfigured fields. Writes are grouped per
target under an undo mark and reread immediately. A failure rolls back that
target and prevents its save.

The versioned save policy is currently `NO_SAVE`. Do not switch it to
`SAVE_CHANGED_PARTS` until the disposable Teamcenter-item write/save/reopen
gate is explicitly confirmed in NX 2312.

## Journal 07 DataPack PDF and STEP export

J07 reads a manually prepared `NX_EXPORT_SCOPE.csv` and exports only the PDF
and STEP outputs enabled by the DataPack scope. It matches the normalized
combination of `DB_PART_NO` and `DB_PART_REV`; it does not select revisions or
decide drawing readiness.

Use `from_git/templates/NX_EXPORT_SCOPE_TEMPLATE.csv`, place the working file
in `NX_JOURNALS_IO_DIR` (or Desktop), fully load the intended assembly and
revision rule, then play J07. It writes:

```text
<I/O root>\NX_BULK_EXPORT\YYYYMMDD_HHMMSS\
  PDF\
  STEP\
  REPORTS\EXPORT_RESULT_YYYYMMDD_HHMMSS.csv
  LOGS\EXPORT_LOG_YYYYMMDD_HHMMSS.txt
```

J07 is self-contained, does not save or modify NX parts, and restores the
original display/work parts.

## Output naming

| Journal | Output |
|---|---|
| J01 | `<DB_PART_NO>_REV<DB_PART_REV>.stp` |
| J02 | `BOM_DRAFT_<root>_<timestamp>.csv` |
| J04 | detail and summary reports; conditional `NX_CERTIFIED_BOM_<root>_<run_id>.csv` |
| J05 | `PULL_<timestamp>.csv` or `J05_<mode>_<timestamp>.csv` |
| J06 | active-part STEP and per-sheet PDF |
| J07 | timestamped `NX_BULK_EXPORT` tree |

## Acceptance boundary

Repository tests provide static and fake-adapter evidence only. NX2312 gates
must separately prove canonical drawing behavior, placeholder suppression,
fully compliant certification, full-HLA hierarchy/dimensions, J05 approval and
stale-audit cases, disposable-item save/reopen behavior, and J07 regression.
Static validation must not be represented as NX runtime proof.

See `docs/J04_J05_ATTRIBUTE_RECONCILIATION_PLAN.md` for the complete rule,
failure, recovery, and runtime-gate specification.
