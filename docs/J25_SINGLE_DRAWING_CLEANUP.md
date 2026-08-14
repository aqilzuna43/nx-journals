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

## Prepare the input

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

Because NX is not installed on this repository host, local tests prove only
parsing, safety gates, API call shape, and report logic. Aqil must run the
`DRY_RUN` in NX X 2506 first and return the CSV, JSON, and log before the apply
path is treated as runtime-proven.
