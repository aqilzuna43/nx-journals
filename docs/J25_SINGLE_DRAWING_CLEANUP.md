# J25 Single Drawing Cleanup

J25 reduces one Teamcenter-managed 3D Item/Revision from multiple drawing
specifications (`dwg1`, `dwg2`, ...) to one explicitly selected final drawing.
It is intended for a customer migration rule that permits only one DWG.

## Exact mutation semantics

NXOpen does not expose a supported relation-only detach call in the available
NX 2506 API. J25 therefore cannot leave an extra UGPART drawing dataset in
Teamcenter as an unassociated orphan.

In `APPLY_APPROVED`, J25 first downloads and hashes every associated file of
each approved extra drawing. It then calls:

```text
FileManagement.DeleteExistingAttachedFiles(files, keepEmptyDataset=False)
```

The call removes the drawing's files and the now-empty drawing dataset. Its
`IMAN_specification` relationship disappears because the dataset is removed.
This is destructive Teamcenter cleanup, not merely hiding or unlinking a DWG.

### Drawings that carry no files

Some extra specifications are empty: NX reports `DrawingSheets = 0` for them
and Teamcenter associates no file with the dataset. There is nothing to back
up, so J25 records the payload count and then attempts the dataset removal
anyway. Because the delete API is file-driven, a zero-file call may remove
nothing at all — so the outcome is decided by evidence, not by the call:

- Dataset gone after the delete, and the final inventory shows only the keep
drawing → `SINGLE_DWG_VERIFIED` as usual.
- Dataset still openable → the row stops being automatic and reports
  `EMPTY_DATASET_REMAINS` (nothing else was removed) or
  `PARTIAL_EMPTY_DATASET_REMAINS` (the other extras were removed). The
  remaining empty dataset must be cut off the revision in the **rich
  Teamcenter client**; the free TCX client has no Cut command.

A blank file *name* is deliberately still a hard failure (`Could not prove all
associated file names`): files exist that cannot be identified, so neither a
provable backup nor a provable delete is possible.

Two report columns make the payload visible per row: `EXTRA_ASSOCIATED_FILE_COUNTS`
(for example `DWG1:0 | DWG2:0 | DWG3:1`) and `EMPTY_DATASET_DWG_INDICES`.
`WRITE_ATTEMPTED` now becomes `YES` only once a delete API call is actually made,
so a row that fails before the first removal reports `NO`.

## Prepare the input

### Option A: generate the scope from a J07 export report (recommended)

A J07 `EXPORT_RESULT_<timestamp>.csv` already records every live drawing per
part (`PDF_FILES` names end in `_DWG<n>.pdf` when more than one drawing was
exported). `from_git/utils/single_drawing_scope.py` converts any such report
into a J25 scope CSV in one step (pure Python; runs on any machine, no NX):

```text
python from_git/utils/single_drawing_scope.py \
    --input  EXPORT_RESULT_20260815_093237.csv \
    --output NX_TC_SINGLE_DRAWING_SCOPE.csv
```

Generated policy (operator-overridable): `KEEP_DWG_INDEX=1` when `dwg1` is
live (otherwise the lowest live index), `EXPECTED_REMOVE_DWG_INDICES` = every
other live index (`2` or `2|3`), `APPROVED=NO`, blank `ENGINEER` and
`CONFIRMATION`. The tool fails closed (writes nothing) when a multi-drawing
row is not `SUCCESS`, its file list does not parse to exactly
`PDF_FILE_COUNT` `_DWG<n>` names, or a part appears twice. J25's apply gates
are unchanged: DRY_RUN review, then per-row sign-off.

Example: the J07 run inside `from_git/templates/LOGS/REPORTS.zip`
(2026-08-15) found 16 masters with extra specifications. The generated scope
is committed at `from_git/templates/NX_TC_SINGLE_DRAWING_SCOPE.csv` (15 rows
with `dwg1`+`dwg2`, one row with `dwg1`+`dwg2`+`dwg3`) and was cross-checked
against the run's text log. Copy that file to the NX machine's I/O root as
`NX_TC_SINGLE_DRAWING_SCOPE.csv`, or point `NX_TC_SINGLE_DRAWING_FILE` at it.

### Option B: copy the template manually

Copy `from_git/templates/NX_TC_SINGLE_DRAWING_SCOPE_TEMPLATE.csv` to the I/O
root as `NX_TC_SINGLE_DRAWING_SCOPE.csv`.

| Column | Meaning |
|---|---|
| `PART_NUMBER` | Exact 3D master Item ID |
| `REVISION` | Exact Item Revision |
| `KEEP_DWG_INDEX` | The one final drawing to retain |
| `EXPECTED_REMOVE_DWG_INDICES` | Exact live extras, for example `2|3` |
| `APPROVED` | Must be `YES` in apply mode |
| `ENGINEER` | Must identify the approving engineer in apply mode |
| `CONFIRMATION` | Must be `REMOVE_EXTRA_DRAWINGS` in apply mode |

J25 scans canonical specifications `dwg1` through `dwg9`. It blocks the row if
the discovered extras differ from `EXPECTED_REMOVE_DWG_INDICES`, if the keep
drawing is missing or has no sheets, if any target is not proven checked in,
or if an extra drawing is already loaded in NX.

### TCX note

The full Teamcenter client can "Cut" an extra drawing dataset off an item
revision; the free TCX web client does not offer that command, so J25 is the
journal-side workaround. J25 cannot perform a relation-only detach (no such
NXOpen API): it downloads every file of each approved extra drawing into the
run's `BACKUP` folder, records SHA-256 evidence, then removes the dataset and
its files with `DeleteExistingAttachedFiles(..., keepEmptyDataset=False)`.
For a cut-forever workflow (no re-attach planned) the visible result is the
same as Cut, and no orphan dataset is left behind in Teamcenter.

## Run safely

1. Leave `USER_MODE = "DRY_RUN"` and play J25 in managed NX X 2506.
2. Review `DISCOVERED_DWG_INDICES`, `LIVE_REMOVE_DWG_INDICES`, the keep drawing,
   and every checkout state in the CSV/JSON result.
3. Close every drawing part in NX.
4. Complete the approval fields, set `USER_MODE = "APPLY_APPROVED"`, and run
   again.
5. Preserve the complete timestamped output folder. Its `BACKUP` directory is
   the recovery evidence for removed drawing payloads.

After every deletion, J25 proves that the removed exact specification no
longer opens. Final success is `SINGLE_DWG_VERIFIED`, which additionally
requires that only the selected DWG remains and still contains drawing sheets.
`EMPTY_DATASET_REMAINS` and `PARTIAL_EMPTY_DATASET_REMAINS` are terminal
results that mean the automatic work is done and one or more empty datasets
still need a Cut in the rich Teamcenter client; update the row's
`EXPECTED_REMOVE_DWG_INDICES` to the remaining live extras before re-running,
because J25 blocks a plan whose expected list does not match the live one.

Because NX is not installed on this repository host, local tests prove only
parsing, safety gates, API call shape, and report logic. Aqil must run the
`DRY_RUN` in NX X 2506 first and return the CSV, JSON, and log before the apply
path is treated as runtime-proven.
