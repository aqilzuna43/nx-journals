# J23 — HLA Assembly Visibility Diagnostic

Use `from_git/journals/23_diagnose_hla_visibility.py` when a part displays
normally in its own NX window but its occurrence geometry is missing from the
top-level HLA assembly.

J23 is read-only. It does not change hide/show state, layers, reference sets,
arrangements, representations, load state, or Teamcenter data, and it never
saves a part. It inventories the current failing display state so that the
first correction can be based on evidence.

## Before running

1. Reproduce the problem in the affected NX session.
2. Make the top-level HLA both the **displayed part** and **work part**.
3. Keep the failing modeling work view and active arrangement selected.
4. Do not open the missing prototype in a new window immediately before the
   run if that action changes the HLA display state.
5. Play `from_git/journals/23_diagnose_hla_visibility.py` from
   **Tools > Journal > Play**.

For a large HLA, optionally define `NX_J23_TARGET` before starting NX. Its
value can be any case-insensitive component name, part number, or assembly-path
substring. J23 still captures the entire structure but prints matching rows
first. Example:

```text
NX_J23_TARGET=264MN012345A01
```

If setting an environment variable is inconvenient, run without it and use
the missing component name to filter the CSV.

## Output

J23 writes two files under:

```text
<NX_JOURNALS_IO_DIR or Desktop>\NX_HLA_VISIBILITY_DIAGNOSTIC\
  J23_HLA_VISIBILITY_<HLA>_<timestamp>.csv
  J23_HLA_VISIBILITY_<HLA>_<timestamp>.json
```

The CSV is for quick filtering. The JSON preserves the active arrangement,
work-view inventory, dynamic-section context, ranked occurrences, and probe
errors. Return the **JSON**, the exact missing component name/path, and a
screenshot of the HLA Assembly Navigator and graphics window.

## Root-cause ranking

The first issue code on a row is J23's highest-ranked explanation. The main
high-confidence results are:

| Issue code | Meaning |
|---|---|
| `SUPPRESSED_CURRENT_ARRANGEMENT` / `ANCESTOR_SUPPRESSED` | Active-arrangement suppression hides the occurrence or its subtree. |
| `COMPONENT_BLANKED` / `ANCESTOR_BLANKED` | Occurrence-level blanking hides the component or a parent subtree. |
| `COMPONENT_LAYER_HIDDEN` / `ANCESTOR_LAYER_HIDDEN` | The component or a parent is placed on a hidden **HLA** layer. |
| `NON_GEOMETRIC_OCCURRENCE` | The HLA occurrence is deliberately marked non-geometric. |
| `EMPTY_REFERENCE_SET` | The occurrence uses NX's Empty reference set. |
| `REFERENCE_SET_NOT_FOUND` | The occurrence names a reference set absent from the resolved revision. |
| `REFERENCE_SET_HAS_NO_GEOMETRY` | The selected reference set exists but contains no body/component geometry. |
| `NO_OCCURRENCE_GEOMETRY` | Prototype members exist, but NX cannot map them to HLA occurrences; stale/corrupt occurrence or representation data is suspected. |
| `ALL_OCCURRENCE_GEOMETRY_BLANKED` | Every mapped geometry member is blanked in the HLA context. |

Medium-confidence results include work-view/isolate visibility, hidden
prototype body layers, body-level blanking, incomplete loading, and
lightweight/partial representation. `ACTIVE_DYNAMIC_SECTION` is deliberately
low confidence because clipping can be valid and unrelated.

`NO_DIRECT_CAUSE_FOUND` is not proof that the occurrence is healthy. It means
the supported read-only NXOpen properties did not expose a direct cause; the
returned JSON is then used to design a narrower second probe.

## Verification boundary

Local tests only verify imports, traversal, ranking, report structure, and the
absence of mutation calls. Siemens NX is not installed on this repository
host. Only an NX 2312/NX X 2506 run on the office machine can verify the API
behavior and identify the real root cause.
