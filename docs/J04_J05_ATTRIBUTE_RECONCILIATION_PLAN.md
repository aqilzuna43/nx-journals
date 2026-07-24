# NX-Authoritative Attribute Reconciliation and Certified BOM Pipeline

## Document status

- **Repository:** `aqilzuna43/nx-journals`
- **Implementation branch:** `codex/j04-j05-attribute-reconciliation`
- **Baseline:** remote `main` at `a87258f`
- **Runtime:** Siemens NX 2312 embedded Python 3.10
- **Data management:** TeamcenterX / managed NX
- **External Python packages:** none
- **Protected J07 commit:** `6fe58f765b8a207229b2f6990e3b0224caa03771`
- **Authority decision:** NX/Teamcenter is authoritative for the engineering BOM and configured attributes. MASTER is downstream reference evidence only.
- **Current J05 save gate:** `NO_SAVE` until the approved disposable-item proof is completed.

This document supersedes the earlier proposal that treated an imported BOM or
MASTER workbook as authority over NX. The production BOM is built from exact
NX/Teamcenter model identities and assembly occurrences. No MASTER or BOM value
may be used to overwrite NX.

## 1. Definition of win

The pipeline succeeds only when all applicable blocking rules pass:

- J02 exports a clearly labelled draft FZ-compatible BOM from the active NX assembly.
- J04 reads NX and Teamcenter-managed drawing specifications without modifying or saving them.
- Part identity is the exact composite `DB_PART_NO + DB_PART_REV`.
- Attribute identity is the exact composite `category + title`.
- Assembly hierarchy and quantity are evaluated by immediate parent, not globally.
- Suppressed occurrences are excluded and unresolved prototypes are reported.
- Exact model-coordinate X/Y/Z dimensions are derived from transformed solid-body bounding boxes.
- Required model, material, classification, commodity, export, traceability, service and drawing rules pass.
- Placeholder values such as blank, `TBC`, `TBD`, and `00-Jan-0` block certification.
- J04 always emits detail and summary evidence, and emits a certified BOM only with zero blocking findings.
- J05 changes only explicitly approved, configured-writable metadata and verifies every change immediately.
- Identity, material, mass, density, volume, dimensions, system-owned, locked and read-only values are never writable.
- J07 remains unchanged from the protected commit and passes its NX regression.

A present value is not automatically valid. For example, `TBC` is populated
but is a blocking placeholder for commodity, country and export-control data.

## 2. Authority and data flow

```text
NX/Teamcenter assembly and attributes (authority)
        |
        +--> J02 --> BOM_DRAFT_<root>_<timestamp>.csv
        |
        +--> J04 read-only certification
                  |       |
                  |       +--> optional MASTER reference -> downstream drift only
                  |
                  +--> detail + summary reports (every run)
                  +--> certified BOM (zero blockers only)
                  |
                  +--> engineer-reviewed correction evidence
                              |
                              +--> J05 DRY_RUN
                              +--> J05 APPLY_APPROVED
                                      |
                                      +--> immediate reread/verification
                                      +--> save only after sandbox gate
```

Valid J05 authorities are `ENGINEERING_APPROVAL` and the configured
model-to-drawing mirror. `MASTER`, `BOM`, and similar downstream exports are
explicitly rejected as correction sources.

## 3. Versioned rule model

The sanitized runtime contract is
`from_git/config/attribute_reconciliation.json`. It contains no inspected
snapshot values, account IDs, workstation names, paths, or other source
metadata. It defines:

- NX/Teamcenter authority and the `NO_SAVE` production gate;
- category-aware, typed identities and attribute rules;
- exact model identity attributes;
- required-on applicability and certification severity;
- normalization, controlled values and placeholder rejection;
- writable targets proven as editable in the supplied NX attribute evidence;
- the canonical drawing resolver
  `@DB/{part_number}/{revision}/specification/{part_number}-{revision}-dwg{index}`
  for indexes 1 through 9;
- FZ-compatible export columns;
- mass/volume/density tolerance and absolute-model-coordinate dimensions.

No unconfirmed material aliases are configured. Material comparison uses
normalized text and remains non-writable.

## 4. Shared assembly and attribute engine

`from_git/utils/attribute_reconciliation.py` is a standard-library/NXOpen
shared engine used by J02 and J04. It provides:

- schema validation, canonical JSON hashing and deterministic run IDs;
- category-aware typed attribute reads, with set/unset/default status retained;
- deterministic recursive traversal with suppressed-occurrence exclusion;
- exact prototype identity and occurrence paths;
- root BOM row at level 0;
- sibling aggregation by exact part and revision;
- immediate-parent quantity, while preserving repeated parts under different parents;
- FZ BOM projection;
- drawing-scope validation;
- optional MASTER subtree drift comparison;
- controlled-value and placeholder checks;
- mass/density/volume consistency checks;
- transformed solid-body bounding-box union in the owning model coordinate system.

Failure to derive an exact dimension is blocking. The implementation does not
fall back to approximate or untransformed extents.

## 5. J02 draft BOM

Run `from_git/journals/02_hla_multilevel_bom.py` with the intended active root
assembly fully loaded. It writes:

```text
BOM_DRAFT_<root>_<timestamp>.csv
```

The exact columns are:

```text
BOM Level, DB_PART_NO, Indented Part Name, Component Name, Quantity,
DB_PART_DESC, DB_PART_NAME, DB_PART_REV, MFG, MPN, Stocking_Type
```

The active root is included at BOM level 0. `Indented Part Name` is derived
from `DB_PART_NAME`, matching the downstream FZ convention. J02 performs no
drawing or release certification; every J02 output is draft.

## 6. J04 fail-closed certification

Run `from_git/journals/04_assembly_attribute_audit.py`. J04 is read-only and
requires `NX_DRAWING_SCOPE.csv` in `NX_JOURNALS_IO_DIR`, unless
`NX_DRAWING_SCOPE_FILE` specifies an explicit path.

Required drawing-scope headers:

```text
Item Number,Item Rev,Drawing Required
```

Drawing decisions are:

- `YES`: a canonical drawing must open, match part/revision, contain a sheet and pass configured comparisons;
- `NO`: drawing checks are `NOT_APPLICABLE`;
- `REVIEW`, a missing identity, an invalid value, or conflicting duplicates: certification is blocked.

An optional `NX_ATTRIBUTE_MASTER_REFERENCE.csv` can be supplied through the
default I/O directory or `NX_ATTRIBUTE_MASTER_FILE`. Structural, quantity,
revision or configured-field differences are reported as
`DOWNSTREAM_BOM_DRIFT` with an instruction to regenerate MASTER. They never
produce an NX correction.

Every run writes detailed and summary CSV evidence. Only a run with zero
blocking findings also writes:

```text
NX_CERTIFIED_BOM_<root>_<run_id>.csv
```

The failure vocabulary includes `PLACEHOLDER_VALUE`,
`INVALID_CONTROLLED_VALUE`, `DRAWING_SCOPE_REVIEW`,
`DIMENSION_DERIVATION_FAILED`, and `DOWNSTREAM_BOM_DRIFT`. Raw NX exception
text and `ErrorCode`, when available, are retained in detail rows.

J04 restores the original work/display state in a `finally` path and closes
only drawings it opened.

## 7. J05 controlled correction

`from_git/journals/05_bulk_attribute_updater.py` remains self-contained for
NX2312 journal deployment. It reads the shared JSON rule model but imports no
repository-local Python module.

Modes are selected with `NX_J05_MODE`:

- `PULL`: create typed, category-aware evidence for the loaded model set and applicable drawings;
- `DRY_RUN`: validate correction rows and report proposed/no-change/rejected actions;
- `APPLY_APPROVED`: apply only rows that pass every approval and stale-audit gate.

The default mode is `DRY_RUN`. The default correction input is
`NX_ATTRIBUTE_CORRECTIONS.csv`; override it with
`NX_ATTRIBUTE_CORRECTIONS_FILE`.

Each row carries approval, exact part/revision, MODEL or DRAWING target,
drawing index, category, title, type, audited raw value, expected value,
authoritative source, audit run ID, engineer, evidence reference and note.

J05 rejects:

- unapproved or incomplete approval evidence;
- part-only or wrong-revision matches;
- wrong target/category/title/type;
- blank expected values;
- stale audited raw values;
- MASTER/BOM authority;
- non-writable, locked, system-owned, PDM-based or read-only fields;
- physical, identity, material or dimension changes.

Approved writes use the NX attribute builder with category and type. Changes
are grouped per target under an undo mark. A failed write or reread verification
rolls back that target and prevents its save. When saving is eventually
enabled, each verified changed target is saved once with component saving
disabled. A save failure stops later saves and leaves the affected target open
and visibly modified with recovery details.

Preloaded parts remain open. Journal-opened parts close only after unchanged
processing or a confirmed successful save. Any target with uncertain rollback
or unsaved changes remains open.

## 8. Inputs, templates and overrides

Default files under `NX_JOURNALS_IO_DIR` (Desktop when unset):

| Input | Required | Override |
|---|---:|---|
| `NX_DRAWING_SCOPE.csv` | J04 certification | `NX_DRAWING_SCOPE_FILE` |
| `NX_ATTRIBUTE_MASTER_REFERENCE.csv` | No | `NX_ATTRIBUTE_MASTER_FILE` |
| `NX_ATTRIBUTE_CORRECTIONS.csv` | J05 DRY/APPLY | `NX_ATTRIBUTE_CORRECTIONS_FILE` |

Templates are kept in `from_git/templates`. Close Excel before running a
journal so the CSV is not locked.

## 9. Evidence case

`264MN025450A01/A` is the fail-closed evidence case. Its current identity and
material evidence is expected to pass, while `TBC` commodity, country and
export-control values (and any other configured placeholder) must prevent
certified BOM output. The inspected source snapshot is local evidence only and
is not versioned.

## 10. Acceptance gates

Static acceptance in this repository covers:

- JSON schema and sensitive-evidence separation;
- typed normalization, placeholders and controlled values;
- hierarchy, per-parent quantities, repeated subtrees and suppression;
- transformed bounding-box behavior;
- drawing scope and MASTER drift;
- J04 mutation/save-call absence;
- J05 approval, stale-value, prohibited-field, rollback, verification and save behavior with fakes;
- J07 byte equality with protected commit `6fe58f7`.

NX 2312 runtime acceptance remains user-confirmed:

1. Run J09 and confirm canonical drawing open and state restoration.
2. Run single-part J04 on `264MN025450A01/A`; placeholders must block certification.
3. Run J04 on a fully compliant item; certified FZ-compatible output must be created.
4. Run full-HLA J04 and validate hierarchy, per-parent quantities, drawings and dimensions.
5. Exercise J05 `PULL`, `DRY_RUN`, approved, unapproved and stale cases.
6. On a user-supplied disposable Teamcenter item, prove reversible write, reread, save, reopen and restoration.
7. Only after explicit approval, change `save_policy` to `SAVE_CHANGED_PARTS`.
8. Rerun J04, J09 and J07 PDF/STEP regressions. Certified output must contain no mismatches.

Static success is not NX runtime proof. Until gate 6 is confirmed, the
versioned configuration must remain `NO_SAVE`.
