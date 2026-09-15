# Journal 38 - Attribute Variant Probe

## Problem

After pushing a native NX file to the customer's Teamcenter, the customer
sees duplicate business attributes with identical values, always as a
display-name / internal-name pair:

| Display-name variant | Internal-name variant |
|---|---|
| Commodity Code | `Commodity_Code` |
| Commodity Type | `COMMODITYTYPE` |
| Country of Origin | `Country_of_Origin` |
| Export Control Number | `Export_Control_Number` |
| Mfr. Name | `MFG` |
| Mfr. Part Number | `MPN` |
| Stocking Type | `Stocking_Type` |
| UOM | `Unit_Of_Measure` |
| WAE Version | `WAE_VERSION` |
| WAE Hazardous | `WAE_Hazardous` |

These 10 pairs match, one for one, the `alias` / `title` pairs deployed in
the NX part attribute template `tests/NXPartAttribute_FZ.xml`. The alias is
display-only inside NX; the title is the name physically stored in the part
file.

## What the journals do (verified from code)

- Journal 05 writes each business field exactly once via
  `AttributePropertiesBuilder` with `Category = "WAEItem"` and the canonical
  title (e.g. `Commodity_Code`). It never writes the alias spelling.
- Journal 04 reads by exact category/title and would not see an alias-named
  attribute even if one existed on the part.

So the NX journals cannot be creating a second variant at write time. The
duplicate is born either (A) inside the part file before export, or (B) on
the customer side during their TC conversion/import. Journal 38 exists to
prove which one.

## Running the probe

1. Open the exact part/assembly you export and push to the customer.
2. Tools > Journal > Play > `from_git/journals/38_attribute_variant_probe.py`.
3. It is strictly read-only (no writes, no saves, no checkout).
4. Collect the output from the desktop (or `NX_JOURNALS_IO_DIR`):
   - `NX_ATTRIBUTE_VARIANT_PROBE\J38_VARIANT_PROBE_<root>_<timestamp>.json`
   - `NX_ATTRIBUTE_VARIANT_PROBE\J38_VARIANT_PROBE_<root>_<timestamp>.csv`
5. Commit the JSON back to the repo (or paste the Listing Window verdict).

If the customer returns native files to you, probe those as well and compare
the two reports.

## Interpreting the verdict

### `NX_FILE_CLEAN` - customer TC conversion problem

The pushed file carries only canonical titles. Every duplicate the customer
sees was created by their conversion/import step. Likely mechanisms on their
side:

- their importer maps one NX attribute into two TC storages (a BMIDE-mapped
  property with a localized display name, plus a second property or form
  field keyed by the raw title);
- their mapping matches by display name (alias) while the CAD attribute is
  keyed by title, causing a second write path;
- the customer TC item type genuinely has two properties per field (one
  localized, one named like the raw title).

Hand them the `parts[].fields[]` block from the JSON: it is the exact
mapping contract (`expected_category = WAEItem`, canonical titles, alias
values) and the per-field evidence that the source file contains one
attribute per field. The six fields absent from their duplicate list
(Temperature_Sensitive, LIFED, Serviceable_item_flag,
SERIAL_NUMBERED_PART, COMPONENT_CLASS, NX_FINISH) are a further clue: those
are exactly the template fields the customer's TC item type does not map.

### `NX_FILE_CARRIES_DUPLICATES` - fix the file before export

The part file itself contains both spellings. Read
`parts[].attributes[]` flags:

- `pdm_based: true` or `owned_by_system: true` on the alias-named attribute
  means it was mirrored into the part by a Teamcenter round trip (e.g. a
  file returned by the customer and re-pushed). Stop re-pushing
  round-tripped files or strip the mirrored attributes first.
- Plain title+alias pairs with both flags false mean the variant was written
  locally (older template, manual entry, or a previous import). Clean those
  parts (a follow-up cleanup journal can delete the alias-named variants by
  exact category/title) and re-export.

`CONFLICTING_DUPLICATE` relations mean the two variants hold different
values - that is worse than cosmetic duplication and needs cleanup before
any further push.

## Trial matrix

| Trial | Where | Result |
|---|---|---|
| 1. Probe the exact file you push | your NX | Proves whether the duplicate is born in the file |
| 2. Probe a file returned by the customer (if any) | your NX | Shows whether their TC round trip injects variants |
| 3. Customer checks if duplicates are two ItemRevision properties or a CAD-attributes tab | customer TC | Locates the second storage on their side |
| 4. Deliver a one-field marker test (unique value in `Commodity_Code`) | both | Shows exactly which TC storages receive the value |
