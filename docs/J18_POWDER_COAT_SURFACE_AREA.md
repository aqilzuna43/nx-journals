# Journal 18 — Powder-Coat Surface Area and Material Estimate

`from_git/journals/18_powder_coat_surface_area.py` avoids selecting or fully measuring a huge assembly. It processes an explicit CSV scope, opens one unique Teamcenter master part at a time, measures all solid bodies, caches the result, closes only journal-opened parts, and multiplies the area by the required quantity.

## Safety and scope

- Read-only: no checkout, save, check-in, dataset creation, or geometry modification.
- Exact identity: `DB_PART_NO + DB_PART_REV`.
- Managed opening attempts:
  - `@DB/<part>/<revision>`
  - `@DB/<part>/<revision>/master`
- Drawing non-master parts are rejected.
- Sheet bodies and assembly-only parts are not included.
- The measured value is full geometric solid-body area. NX does not know which faces are masked, internal, mating, inaccessible, or intentionally uncoated.

## Prepare the CSV

Copy `from_git/templates/NX_POWDER_COAT_SCOPE_TEMPLATE.csv` to the Desktop or `NX_JOURNALS_IO_DIR` and save it with the exact name:

```text
NX_POWDER_COAT_SCOPE.csv
```

Required logical columns:

| Value | Accepted headings |
|---|---|
| Part number | `DB_PART_NO`, `PART_NUMBER`, `Part Number`, `Item Number` |
| Revision | `DB_PART_REV`, `REVISION`, `Item Rev` |
| Quantity | `QUANTITY`, `QTY` |

Recommended columns:

| Column | Meaning | Default |
|---|---|---:|
| `INCLUDE` | `YES` for powder-coated items; blank/`NO` is excluded when the column exists | All rows included when column is absent |
| `POWDER_CODE` | Colour or powder purchasing code used for summary grouping | `UNSPECIFIED` |
| `COATED_AREA_FACTOR` | Manufacturing correction for masked/uncoated faces | `1.00` |
| `COATS` | Number of identical coats | `1` |
| `DFT_UM` | Cured dry-film thickness in micrometres | `70` |
| `SPECIFIC_GRAVITY` | Powder-specific value from supplier TDS | `1.50` |
| `UTILISATION` | Overall transfer/reclaim efficiency | `0.85` |
| `CONTINGENCY` | Reserve for startup, colour change and rejects | `0.10` |
| `PACK_SIZE_KG` | Supplier bag or box size | `20` |

Fractions may be entered as `0.90`, `90`, or `90%`.

The input should contain the **final total quantity per unique part/revision**. Duplicate rows with identical coating assumptions are summed automatically. A raw multilevel BOM must first be converted to rolled-up quantities; parent-level quantities are not inferred by J18.

## Run

1. Start managed NX connected to Teamcenter X.
2. The top-level assembly does not need to be fully loaded. An existing open part may remain displayed.
3. Close Excel so the CSV is not locked.
4. Run:

```text
NX > Tools > Journal > Play
from_git\journals\18_powder_coat_surface_area.py
```

Optional environment overrides:

```text
NX_JOURNALS_IO_DIR=C:\NX_JOURNALS_IO
NX_POWDER_COAT_INPUT=C:\path\custom_scope.csv
NX_POWDER_COAT_ACCURACY=0.99
```

## Calculation model

For each CSV row:

```text
Coated area per part = NX solid-body area × coated-area factor
Total coated area = coated area per part × quantity × coats
Cured film volume (L) = total coated area (m²) × DFT (µm) / 1000
Theoretical powder (kg) = cured film volume (L) × specific gravity
Required powder (kg) = theoretical powder / utilisation × (1 + contingency)
```

The summary rounds purchasing quantity upward to full packs:

```text
Bags required = CEILING(required powder / pack size)
Purchase quantity = bags required × pack size
Estimated spare = purchase quantity − required powder
```

## Outputs

Each run creates:

```text
<I/O root>\NX_POWDER_COAT\YYYYMMDD_HHMMSS\
  REPORTS\POWDER_COAT_DETAIL_<timestamp>.csv
  REPORTS\POWDER_COAT_SUMMARY_<timestamp>.csv
  LOGS\POWDER_COAT_LOG_<timestamp>.txt
```

The detail report contains one aggregated input row per part/revision/coating setup. The summary groups demand by powder code and matching DFT, specific gravity, utilisation, contingency and pack size.

## Acceptance test

Use two simple test parts before the production batch:

1. A known rectangular solid with manually calculable surface area.
2. A multi-solid part containing two identical solids.

Require:

- `RAW_AREA_M2_PER_PART` agrees with NX Measure within the selected accuracy.
- The two-solid test returns approximately twice the one-solid area.
- A duplicated CSV row is measured once and its quantities are summed.
- The journal restores the original display/work part.
- No part is saved, checked out or checked in.
- `POWDER_COAT_SUMMARY` matches a manual calculation for one sample row.

## Manufacturing validation

Before purchasing production material, confirm the approved powder TDS values and compare the estimate against the supplier's historical powder issued versus accepted coated parts. Replace the default area factor, DFT, specific gravity, utilisation and contingency with line-specific data.
