# Journal 28 — Raw BoM Structure Checkpoint

## Purpose

Journal 28 captures the complete occurrence structure of the active NX
assembly before any Reference-Only cleanup design is finalized. It records the
root, suppressed occurrences, Reference-Only occurrences, repeated placements,
and every descendant returned by NX without grouping or pruning.

This is an evidence journal, not a production BoM exporter. It does not prove
that a replacement marker is safe for JT or native NX parts lists.

Build: `J28-NX2506-BOM-STRUCTURE-PROBE-V1`

Primary runtime: NX X 2506 embedded Python in managed or native mode.

## Read-only contract

J28 does not:

- load or fully load components;
- change suppression, visibility, reference sets, layers, or arrangements;
- create, edit, or delete attributes;
- update, save, check out, or check in NX/Teamcenter data; or
- group repeated components or filter any returned occurrence.

The JSON records `IsModified` before and after the probe. This is supporting
evidence, not a substitute for checking the NX session after the run.

## Run instructions

1. Open the representative top-level assembly in NX X 2506.
2. Keep the normal customer load and revision-rule settings. Do not clean up
   Reference-Only occurrences before this checkpoint.
3. In NX, select **Tools > Journal > Play**.
4. Run `from_git/journals/28_probe_bom_structure.py`.
5. Wait for `Run status`, `CSV`, and `JSON` to appear in the Listing Window.
   Progress is reported every 500 occurrences.
6. Do not save the assembly merely because the probe was run. Confirm that NX
   has not marked the work part modified.

By default, output is written beneath:

```text
%USERPROFILE%\Desktop\NX_BOM_STRUCTURE_PROBE\
```

Set `NX_JOURNALS_IO_DIR` before launching NX to use another local output root.
Each run receives a separate timestamped folder and run ID.

The occurrence safety cap defaults to 100,000.  Set
`NX_J28_MAX_OCCURRENCES` before launching NX to tighten it for a
per-subassembly checkpoint run (for example `10000`); a hit cap makes the
run `INCOMPLETE` with `safety_limit_reached` true.

## Artifacts

### Occurrence CSV

The UTF-8-BOM CSV contains one row per occurrence in NX preorder. The root is
level 0. Structural paths contain sibling indexes so repeated component names
remain distinguishable.

Important evidence groups include:

- component, prototype, part-number, revision, and stocking metadata;
- suppression, reference set, representation, non-geometric, and layer probes;
- separate presence, raw value, value-state, and read-status fields for
  `REFERENCE_COMPONENT`, `PLIST_IGNORE_MEMBER`, and
  `PLIST_IGNORE_SUBASSEMBLY`;
- direct control classification and nearest controlling ancestor; and
- a prediction of what the current `NXOpenBoMExtended.py` logic would do.

The prediction is diagnostic only. J28 never uses it to prune traversal.

### JSON summary

The JSON contains:

- build, schema, run, NX runtime, and root metadata;
- `COMPLETE`, `INCOMPLETE`, or `FAILED` run status;
- occurrence and classification totals;
- descendant counts beneath every directly controlled occurrence;
- traversal and read failures;
- complete typed instance-attribute inventories for flagged, inconsistent, or
  unreadable occurrences only;
- CSV/schema SHA-256 values; and
- work-part modified state before and after.

An inaccessible branch, unreadable required BoM evidence, changed modified
state, or the 100,000-occurrence safety cap makes the run `INCOMPLETE`.
Optional metadata failures, such as an unavailable stable instance ID, remain
visible in `read_errors` but do not invalidate an otherwise complete structural
checkpoint. A root-access failure produces a failed JSON and no CSV.
Unexpected failures retain `.partial` files for troubleshooting.

## Artifact return checklist

Before sharing the output, verify:

- the assembly contains at least one current Reference-Only subassembly with
  descendants;
- the CSV contains that occurrence and its descendants;
- the JSON reports no unexpected traversal gaps;
- the CSV and JSON have the same run ID;
- `work_part_modified.changed` is `false`; and
- the assembly did not become checked out or modified because of the run.

Return both files together. They can contain proprietary part numbers,
structure, paths, and custom attributes. Do not commit raw customer artifacts
without deliberate review and sanitization.

## Acceptance boundary

Local pytest verifies traversal, classification, report construction, and the
absence of known mutation calls. Siemens NX is not installed on this machine,
so only Aqil's NX X 2506 run can establish runtime behavior.

The checkpoint does not authorize changes to `NXOpenBoMExtended.py`. After the
artifacts are reviewed, resume the BoM design and decide whether native
`PLIST_IGNORE_MEMBER` / `PLIST_IGNORE_SUBASSEMBLY`, a custom occurrence
attribute, or another mechanism is appropriate. Siemens documents that newer
managed-mode NX releases retain Reference-Only intent in Teamcenter relations,
but the customer's JT and parts-list recipes still require direct validation:

- [Siemens managed-mode Reference-Only update](https://blogs.sw.siemens.com/designcenter/whats-new-in-nx-june-2024-managed-mode/)
- [Historical NX parts-list attribute discussion](https://www.eng-tips.com/threads/nx5-assemblies-parts-list.215349/)
