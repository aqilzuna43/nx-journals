# J23 V2 — Exact-target HLA Visibility Evidence

Use `from_git/journals/23_diagnose_hla_visibility.py` when an occurrence or
subassembly is missing in the main HLA window but displays correctly through
**Isolate in New Window**.

J23 V2 is read-only. It does not show, hide, blank, suppress, load, switch a
view, change a reference set, update, or save anything.

## Why V2 replaced the first report

The first NX X 2506 run proved useful facts, but also exposed unsupported API
calls. V1 treated some unavailable properties as false and therefore produced
unsupported findings such as `PROTOTYPE_NOT_FULLY_LOADED`. V2 uses this rule:

> A failed or unavailable probe is `ERROR` or `UNAVAILABLE`; it is never
> converted into `NO`, `False`, zero, or a root cause.

Every V2 conclusion cites fact IDs from `evidence_ledger`. Each hypothesis is
`CONFIRMED`, `STRONGLY_SUPPORTED`, `RULED_OUT`, or `INCONCLUSIVE`.

## Current target

The checked-in fallback is:

```python
USER_TARGET = "264MN031978A01"
```

Target resolution order is:

1. Exactly one component preselected in Assembly Navigator.
2. Exact part number from `NX_J23_TARGET`.
3. `USER_TARGET` in the journal.

Preselection is best when the same part number occurs more than once because
it identifies the exact component tag and assembly path.

## Run on the office NX machine

1. Pull the current `master`.
2. Reproduce the missing component in the **main HLA window**.
3. Make the HLA both displayed part and work part.
4. In Assembly Navigator, select exactly the missing
   `264MN031978A01/A` occurrence. Do not select its prototype in a separate
   window.
5. Keep the failing main view active.
6. Play `from_git/journals/23_diagnose_hla_visibility.py`.
7. Push the generated `J23_EVIDENCE_*.json` under `docs/`.

The report is written under:

```text
<NX_JOURNALS_IO_DIR or Desktop>\NX_HLA_VISIBILITY_DIAGNOSTIC\
  J23_EVIDENCE_<target>_<timestamp>.csv
  J23_EVIDENCE_<target>_<timestamp>.json
```

## What V2 proves

For the selected occurrence and its complete subtree, V2 records:

- exact component/prototype tags and assembly paths;
- direct suppression and blanking as independent NX probes;
- HLA component-layer state from the displayed HLA only;
- non-geometric and representation probes without boolean fallbacks;
- runtime prototype type and load-property availability;
- actual reference-set body **and component** members;
- `Component.FindOccurrence` mappings for those members;
- exact mapped body tags present in the active work view;
- the same body tags present in every readable saved modeling view;
- visible same-part/revision controls outside the target subtree;
- dynamic-section visibility in the active view;
- a hypothesis table and fact-cited conclusion.

`ISOLATE_VIEW_MECHANISM` is intentionally not marked `CONFIRMED` merely because
the work view is named `Isolate`. NXOpen exposes commands that create and edit
isolate membership, but the available read-only interface does not expose a
direct membership query. It becomes `STRONGLY_SUPPORTED` only when independent
geometry/view comparisons support it.

## Facts already established for 264MN031978A01

The first NX artifact under
`docs/J23_HLA_VISIBILITY_264MN024625A01_20260814_115209.json` proves:

- target subtree occurrence rows: **28**;
- rows with successfully mapped geometry absent from the work view: **27**;
- mapped-but-absent rows whose components report unsuppressed: **21**;
- blanked target-subtree rows: **0**;
- target-subtree rows with mapped geometry visible in the main view: **0**;
- active work-view name: **Isolate**.

Therefore, suppression and blanking do not explain the entire missing
subassembly. V2's alternate-view and same-prototype control probes are designed
to close the remaining gap between “current-view exclusion observed” and a
confirmed view-specific root cause.

## Verification boundary

Local tests verify traversal, exact targeting, tri-state probes, subassembly
reference-set mapping, evidence citations, hypothesis rules, report structure,
and absence of mutation calls. Only the next NX X 2506 JSON can verify the new
runtime probes and establish the final root cause.

## If NX X 2506 has no visible “Exit Isolation” command

Run `from_git/journals/24_repair_hla_isolate_visibility.py`. NXOpen exposes no
`ExitIsolation` method, so J24 uses the supported
`ComponentAssembly.ShowComponentsInIsolateView` operation against the exact
selected target subtree. It records mapped-body visibility before and after,
does not save, and creates a visible undo mark. See
`docs/NX_HLA_ISOLATE_VISIBILITY_REPAIR.md` for the guarded run procedure.
