# Journal 15 — Teamcenter X Offline Drawing Workflow

## Supported runtime

**NX X 2506 only.**

J15 is no longer designed or maintained for NX 2312. The NX X 2506 Python binding used by this workflow is `NXOpen.UF.Clone`.

## Goal

Export a Teamcenter-managed 3D model plus its drawing to native NX files, work on the drawing locally, then update **only the existing drawing specification** in Teamcenter X.

The Teamcenter 3D master remains authoritative and must never be overwritten by the local reference copy.

Files:

```text
from_git/journals/15_tc_offline_drawing_workflow.py
from_git/templates/NX_TC_OFFLINE_SCOPE_TEMPLATE.csv
```

## Keep the Teamcenter AutoTranslate filename

Do **not** rename exported Teamcenter native files.

Example:

```text
264MN028282A01_A_s_264MN028282A01-A-dwg1.prt
```

The long filename is intentional. J15 uses it as another identity check before allowing the drawing to be returned to Teamcenter.

## Safety rules built into J15

- `EXPORT` never writes to Teamcenter.
- Export keeps Teamcenter AutoTranslate filenames unchanged.
- Local 3D/reference `.prt` files are marked Windows read-only after export.
- The target drawing remains writable.
- The drawing gets a SHA-256 baseline in the manifest.
- `IMPORT_DRY_RUN` must be used before a real import.
- `IMPORT_APPLY` requires `APPROVED=YES` and a nonblank `ENGINEER`.
- Import default action is `UseExisting` for every discovered object.
- Only the exact target drawing gets `Overwrite`.
- A renamed drawing or wrong `/specification/` identity is rejected.
- Unchanged drawings are skipped by SHA-256.
- Any failed export/import row makes the overall J15 run report `FINAL STATUS: FAILED`.

## 1. Deploy

Pull the latest `master` branch and keep the complete `from_git` folder together.

Run journals from:

```text
NX > Tools > Journal > Play
```

J15 uses `NX_JOURNALS_IO_DIR` if configured; otherwise it uses the current user's Desktop.

Optional example, set before starting NX:

```bat
set NX_JOURNALS_IO_DIR=D:\NX_OFFLINE_WORK
```

## 2. Prepare the scope CSV

Copy:

```text
from_git/templates/NX_TC_OFFLINE_SCOPE_TEMPLATE.csv
```

into the J15 I/O root and save the working copy with the exact name:

```text
NX_TC_OFFLINE_SCOPE.csv
```

Example content:

```csv
PART_NUMBER,REVISION,DWG_INDEX
264MN028282A01,A,1
```

This resolves:

```text
Model   @DB/264MN028282A01/A
Drawing @DB/264MN028282A01/A/specification/264MN028282A01-A-dwg1
```

## 3. Export

Default J15 settings are:

```python
USER_MODE = "EXPORT"
USER_SCOPE_CSV = r""
USER_MANIFEST_CSV = r""
```

When `USER_SCOPE_CSV` is blank, J15 expects:

```text
<I/O root>\NX_TC_OFFLINE_SCOPE.csv
```

Environment-variable override remains available:

```bat
set NX_TC_OFFLINE_MODE=EXPORT
set NX_TC_OFFLINE_SCOPE_FILE=D:\NX_OFFLINE_WORK\NX_TC_OFFLINE_SCOPE.csv
```

Environment variables must exist before the NX process starts if you want NX to inherit them.

Run:

```text
from_git\journals\15_tc_offline_drawing_workflow.py
```

The NX Listing Window should begin with something similar to:

```text
J15 TEAMCENTER X OFFLINE DRAWING WORKFLOW
Build: J15-TCX-OFFLINE-DRAWING-NX2506-V3
Runtime target: NX X 2506 only
UF Clone binding: NXOpen.UF.Clone
```

Typical output:

```text
NX_TC_OFFLINE_DRAWINGS\<timestamp>\
  264MN028282A01_A_DWG1\
    <3D/reference native .prt files - READ ONLY>
    264MN028282A01_A_s_264MN028282A01-A-dwg1.prt
    EXPORT_264MN028282A01_A_DWG1.clone
  TCX_OFFLINE_MANIFEST_<timestamp>.csv
```

## 4. Work locally

Open and edit only the target `_s_...dwg<n>.prt` drawing.

Do not:

- rename exported `.prt` files;
- move individual `.prt` files out of their package;
- remove read-only protection from 3D/reference files;
- Save As a local 3D reference under the same Teamcenter identity.

If a required reference is missing, fix the export coverage instead of recreating the geometry locally.

## 5. Approve a drawing for return

Open the generated manifest and, for drawings that should be returned, set:

```text
APPROVED=YES
ENGINEER=<your engineering identifier>
```

Do not manually change:

```text
DRAWING_IDENTIFIER
DRAWING_FILE
EXPORT_SHA256
```

## 6. Dry-run import

Set in J15:

```python
USER_MODE = "IMPORT_DRY_RUN"
USER_MANIFEST_CSV = r"D:\...\TCX_OFFLINE_MANIFEST_<timestamp>.csv"
```

or set before launching NX:

```bat
set NX_TC_OFFLINE_MODE=IMPORT_DRY_RUN
set NX_TC_OFFLINE_MANIFEST_FILE=D:\...\TCX_OFFLINE_MANIFEST_<timestamp>.csv
```

For a changed drawing, require:

```text
CHANGED=YES
DEFAULT_IMPORT_ACTION=UseExisting
DRAWING_IMPORT_ACTION=Overwrite
DRY_RUN=YES
RESULT=DRY_RUN_OK
```

Do not proceed to apply if any row is `FAILED` or the identity is unexpected.

## 7. Apply

After a clean dry run:

```python
USER_MODE = "IMPORT_APPLY"
```

The core rule is:

```text
All 3D/reference/native related parts -> UseExisting
Exact target drawing                  -> Overwrite
```

`IMPORT_APPLY` will not process a changed drawing unless the manifest contains both:

```text
APPROVED=YES
ENGINEER=<nonblank>
```

## 8. First acceptance test

Use only one drawing until the workflow is proven on the actual Teamcenter X tenant.

Example:

```csv
PART_NUMBER,REVISION,DWG_INDEX
264MN024819A01,A,1
```

Verify:

```text
[ ] Listing Window shows NX X 2506 only
[ ] Listing Window shows UF Clone binding: NXOpen.UF.Clone
[ ] Export succeeds
[ ] Target native drawing exists
[ ] 3D/reference .prt files are read-only
[ ] Drawing opens locally with references resolved
[ ] Drawing can be saved locally
[ ] IMPORT_DRY_RUN returns DRY_RUN_OK
[ ] IMPORT_APPLY updates the existing Teamcenter drawing only
[ ] Managed 3D master remains unchanged
[ ] No unexpected Item/Revision/UGMASTER/drawing dataset is created
[ ] Drawing associativity remains correct after reopening from Teamcenter
```

Do not run a mass batch until this passes.

## Troubleshooting

### `module 'NXOpen.UF' has no attribute 'UFClone'`

You are running an old J15 revision. Pull the latest `master`. J15 for NX X 2506 uses:

```python
NXOpen.UF.Clone
```

not `NXOpen.UF.UFClone`.

### Template filename contains `DRAWING`

The obsolete template name was removed. Use:

```text
NX_TC_OFFLINE_SCOPE_TEMPLATE.csv
```

and copy it to the I/O root as:

```text
NX_TC_OFFLINE_SCOPE.csv
```

### `SKIPPED_UNCHANGED`

The local drawing still has the same SHA-256 as the export snapshot, so there is nothing to import.

### Apply blocked

Confirm `APPROVED=YES` and `ENGINEER` are populated in the manifest.

## Core invariant

```text
3D master/reference = UseExisting
Drawing specification = Overwrite only when exact manifest identity + exact native filename match
```
