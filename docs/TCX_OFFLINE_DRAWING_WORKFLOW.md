# Journal 15 — Teamcenter X Offline Drawing Workflow

## Goal

Export a Teamcenter-managed model plus one drawing to native NX, edit the drawing locally, then update **only the existing drawing specification** in Teamcenter X. The managed 3D master remains authoritative and must never be overwritten by the local copy.

Files:

```text
from_git/journals/15_tc_offline_drawing_workflow.py
from_git/templates/NX_TC_OFFLINE_DRAWING_SCOPE_TEMPLATE.csv
```

Target: NX 2312 / NX X 2506 embedded Python in a managed Teamcenter session. No third-party packages.

## Keep the long filename

Do **not** rename exported Teamcenter native files. Example:

```text
264MN028282A01_A_s_264MN028282A01-A-dwg1.prt
```

The workflow relies on Teamcenter AutoTranslate naming when mapping the native drawing back to the existing managed object. Identity safety is more important than a shorter local filename.

## Safety rules built into J15

- `EXPORT` never writes to Teamcenter.
- Export keeps AutoTranslate filenames unchanged.
- After export, every `.prt` except the target drawing is marked Windows read-only.
- The drawing gets a SHA-256 baseline in the manifest.
- `IMPORT_DRY_RUN` validates before a real import.
- `IMPORT_APPLY` requires `APPROVED=YES` and nonblank `ENGINEER`.
- Import default action is `UseExisting` for every discovered object.
- Only the exact target drawing file gets `Overwrite`.
- A renamed file or a target that is not `/specification/...` is rejected.
- Unchanged drawings are skipped by SHA-256.
- Apply mode stops after the first runtime failure.

## 1. Deploy

Pull `master` and keep the complete `from_git` folder together. Run from:

```text
NX > Tools > Journal > Play
```

I/O follows the repo convention: `NX_JOURNALS_IO_DIR` if set, otherwise Desktop.

Optional before launching NX:

```bat
set NX_JOURNALS_IO_DIR=D:\NX_OFFLINE_WORK
```

## 2. Prepare scope

Copy the template to the I/O root and name it:

```text
NX_TC_OFFLINE_SCOPE.csv
```

Example:

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

In J15:

```python
USER_MODE = "EXPORT"
USER_SCOPE_CSV = r""
USER_MANIFEST_CSV = r""
```

Or use environment variables:

```bat
set NX_TC_OFFLINE_MODE=EXPORT
set NX_TC_OFFLINE_SCOPE_FILE=D:\NX_OFFLINE_WORK\NX_TC_OFFLINE_SCOPE.csv
```

Run `15_tc_offline_drawing_workflow.py` in managed NX.

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
- move individual `.prt` files out of the package;
- remove read-only protection from 3D/reference files;
- Save As a local 3D reference under the same Teamcenter identity.

If a required 3D reference is missing, fix export coverage instead of recreating geometry locally.

## 5. Approve the drawing for return

Open the generated manifest. For drawings to return to Teamcenter, set:

```text
APPROVED=YES
ENGINEER=<your engineering identifier>
```

Do not edit:

```text
DRAWING_IDENTIFIER
DRAWING_FILE
EXPORT_SHA256
```

## 6. Dry-run import

Set:

```python
USER_MODE = "IMPORT_DRY_RUN"
USER_MANIFEST_CSV = r"D:\...\TCX_OFFLINE_MANIFEST_<timestamp>.csv"
```

or:

```bat
set NX_TC_OFFLINE_MODE=IMPORT_DRY_RUN
set NX_TC_OFFLINE_MANIFEST_FILE=D:\...\TCX_OFFLINE_MANIFEST_<timestamp>.csv
```

Run J15 again. For a changed drawing, require:

```text
CHANGED=YES
DEFAULT_IMPORT_ACTION=UseExisting
DRAWING_IMPORT_ACTION=Overwrite
DRY_RUN=YES
RESULT=DRY_RUN_OK
```

Do not continue if the report shows a wrong identity, unexpected filename, missing file, or `FAILED`.

## 7. Apply

After a clean dry run:

```python
USER_MODE = "IMPORT_APPLY"
```

The import is intentionally configured as:

```text
All 3D/reference/native related parts -> UseExisting
Exact target drawing                  -> Overwrite
```

`IMPORT_APPLY` will not process a changed drawing unless the manifest has `APPROVED=YES` and `ENGINEER` filled in.

## 8. First acceptance test — mandatory before mass use

Start with one drawing only, for example:

```csv
PART_NUMBER,REVISION,DWG_INDEX
264MN028282A01,A,1
```

Verify:

```text
[ ] Export succeeds
[ ] Target drawing native file exists
[ ] 3D/reference native files are read-only
[ ] Drawing opens locally with references resolved
[ ] Drawing can be saved locally
[ ] IMPORT_DRY_RUN returns DRY_RUN_OK
[ ] IMPORT_APPLY updates the existing Teamcenter drawing
[ ] Managed 3D master is unchanged
[ ] No unexpected Item, Revision, UGMASTER, or drawing dataset is created
[ ] Drawing associativity to managed 3D remains correct after reopen
```

Do not run hundreds of drawings until this passes on the actual Teamcenter X tenant. Site preferences, ownership, permissions, dataset rules, and revision rules can affect UFClone behavior.

## Troubleshooting

**Drawing not found after export:** verify the Teamcenter specification dataset follows `<PART>-<REV>-dwg<n>`. Journal 09 can independently test the canonical `/specification/` identifier.

**Drawing renamed locally:** restore the original exported filename. J15 intentionally refuses to guess the Teamcenter target.

**`SKIPPED_UNCHANGED`:** the current SHA-256 still matches the export baseline, so no import is needed.

**Apply blocked:** confirm `APPROVED=YES` and `ENGINEER` are populated.

**Unexpected tenant behavior:** stop the batch and retain the J15 log, package `.clone` log, manifest, and import report for diagnosis.

## Core invariant

```text
3D master/reference = UseExisting
Drawing specification = Overwrite only when exact manifest identity + exact native filename match
```
