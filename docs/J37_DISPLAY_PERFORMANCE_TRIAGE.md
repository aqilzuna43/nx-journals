# J37 Display Performance Triage — Findings and Open Decisions

Status: **OPEN** — root cause identified, one decision outstanding (reference-set
change), no Teamcenter write attempted. Pick this up at
[Next actions](#next-actions).

Symptom under investigation: *a simple view change on the PDU assembly takes
about 10 minutes.*

## Evidence base

| Item | Value |
| --- | --- |
| Report | `from_git/templates/LOGS/REPORTS/J37_EVIDENCE_20260912_113048.json` |
| Build | `J37-NX2506-DISPLAY-PERF-TRIAGE-V2` |
| Commit | `d86cca4 lATEST LOG` (preceded by `03c20dd` v2 fixes, `ae76293` journal) |
| Mode / scope | `PROBE` / `ALL` |
| Assembly root | `264MN021888A01/A;1-ASY-PDU-WORKING` |
| Journal | `from_git/journals/37_diagnose_display_performance.py` |
| Prior analysis | Pi session `01a09348-4fb1-73a2-ba7a-f214a73e40e4` (crashed 2026-09-12 11:34 local; its final message is the complete V2 review) |

Measured inventory: 2,947 occurrences, 719 prototypes, 255,622 faces, 11,306
bodies, 30,432 visible objects (11.26 s scan). Evidence ledger: 15,486 `OK`,
10,552 `UNAVAILABLE`. Suspects: 459 HIGH, 169 MEDIUM, 6 INFO. Load state: 672
FullyLoaded, 1 PartiallyLoaded (`264MN019808A01/A`), 46 unreadable. The view
rendering style is `0` = `ShadedWithEdges` (`NXOpen.View.RenderingStyleType`,
verified in `documents/input/20230921Intellisense/NX12/NXOpen/__init__.py:513`
and `Release2023/NXOpen/__init__.py:768`).

## What the view actually draws

| Visible object type | Count | Share |
| --- | --- | --- |
| Body | 18,719 | 61.5% |
| **DisplayedConstraint** | **3,674** | **12.1%** |
| Component | 2,725 | 9.0% |
| Line + Spline + Arc + ControlPoint + SplineSegment | 3,734 | 12.3% |
| Routing ports (Fitting/Stock/Extract/Fixture/Multi) | 604 | 2.0% |
| DatumAxis + DatumPlane + CSYS | 593 | 2.0% |
| Point / Group / Face / Edge / misc | 383 | 1.3% |
| **Total** | **30,432** | (scan 11.3 s) |

## Root cause, ranked

1. **Reference set `Entire Part` on 1,414 occurrences / 404 parts — the
   multiplier.** Those parts hold 6,522 bodies / 84,814 faces (57.7% of bodies,
   33.2% of faces as *unique-part* shares).
2. **Facets are regenerated and edges are drawn.** `ShadedWithEdges` plus a
   stored-facets setting that cannot be confirmed (see
   [Evidence gaps](#evidence-gaps-and-corrections)), so every view change
   re-tessellates and then strokes edges over a very large face count.
3. **Constraint display is ON — 3,674 `DisplayedConstraint` objects (12.1% of
   the view)** of pure overhead. These are the mate/align glyphs; NX's own
   large-assembly guidance is to hide them.
4. **Wireframe/port clutter.** `264MN024578A01/A` alone holds 8,174
   datum/curve/line/point objects; 604 routing ports + ~4,200 curves/splines/arcs
   are cable/harness content; 100 dynamic sections across 69 parts.
5. **256 visible layers on the root** (`264MN021888A01/A`) with only 3 populated.
6. **Hot geometry:** `264MN025484A01/A` 51,495 faces / 1,655 blanked bodies;
   `032609/A` 31,637 faces / 2,203 bodies.
7. **46 unreadable prototypes**, all direct children of the root — hardware and
   harness/routing library parts whose prototype exposes no Name, no Bodies and
   no Preferences. They contribute the visible ports/splines/arcs.

Ruled out as the view-change cause: 223 zero-density bodies across 77 parts and
5 parts with impossible density ranges (e.g. `min=1470` vs `max=78,306,400,000`)
— a real data defect, but not the cause of slow view updates.

## Corrected arithmetic — the prize is bigger than the first review said

The first review quoted "84,814 faces = 33% of the assembly". That is the
**unweighted, per-unique-part** share. Weighted by occurrence count — what the
view actually draws — the picture is different:

| Measure | Value | Share |
| --- | --- | --- |
| `Entire Part` occurrences / parts | 1,372–1,414 of 2,947 · 404 of 719 | ~48% of occurrences |
| Entire-Part unique-part bodies | 6,522 of 11,306 | 57.7% |
| Entire-Part unique-part faces | 84,814 of 255,622 | 33.2% |
| **Entire-Part occurrence-weighted bodies** | **13,647** of 18,719 drawn | **up to 72.9%** |
| Entire-Part occurrence-weighted faces | 174,087 face instances | — |
| Bodies per occurrence | Entire Part ≈ **9.95** vs everything else ≈ **3.22** | ~3× inflation |

The 72.9% is an **upper bound**: blanked, suppressed and off-layer content is
not drawn. The arithmetic does close, which is why it is worth taking seriously
— the non-Entire-Part residual of 5,072 drawn bodies is consistent with the
4,784 unique bodies those 315 prototypes hold.

Reproduce with:

```bash
cd from_git/templates/LOGS/REPORTS
node -e '
const j=JSON.parse(require("fs").readFileSync("J37_EVIDENCE_20260912_113048.json","utf8"));
const n=f=>Number(f.value)||0; const B={},F={};
for(const f of j.facts){const o=f.owner;
  if(o&&!["session","Top","occurrences=2947"].includes(o)){
    if(f.probe==="Bodies")B[o]=n(f); if(f.probe==="GetFaces")F[o]=n(f);}}
const ep=j.suspects.filter(s=>s.CODE==="ENTIRE_PART_REFSET");
let b=0,f=0,ob=0,of=0;
for(const s of ep){const o=+s.OCCURRENCE_COUNT||0;
  b+=B[s.IDENTITY]||0; f+=F[s.IDENTITY]||0;
  ob+=o*(B[s.IDENTITY]||0); of+=o*(F[s.IDENTITY]||0);}
console.log({uniqueBodies:b,uniqueFaces:f,occBodies:ob,occFaces:of});'
```

## The open decision: risk of `Entire Part` → `MODEL`

**Tier 1 — functional**

1. A `MODEL` reference set shows bodies only. Everything non-solid leaves the
   display: ~3,700 visible curves/splines/arcs + 604 routing ports + 593
   datums/CSYS (12% of 30,432 visible objects). For the harness/cable prototypes
   (`W240 EXTEND CABLE/A`, `036460/A` `028508/A` and most of the 46 unreadable
   root children) that routing geometry **is** the product. J23/J24
   visibility tooling would then report refset-hidden harness parts as
   "missing", pushing false positives into the visibility-repair pipeline.
2. Assembly constraints can break: a constraint bound to a datum plane, sketch
   edge or curve that only `Entire Part` exposes loses its reference on update.
   3,674 displayed constraints is a large exposure surface.
3. Blanked bodies stay blanked — a `MODEL` reference set does not unblank
   `264MN025484A01/A` (1,655 of 1,702 bodies blanked). Expect components that
   render empty and read as "deleted".
4. The reference set must already exist. J37 only tests
   `ReferenceSet == EntirePartRefsetName` (aggregate facts `F26032`/`F26033`);
   it never enumerates available reference sets per part. Parts without a usable
   `MODEL` set either refuse the change or go invisible. The 46 unnamed
   prototypes cannot be switched programmatically at all.

**Tier 2 — structural / governance**

5. A reference set is a component property, so it lives in the **assembly**
   files (root plus every sub-assembly on the path). Fixing 404 parts means
   check-out/check-in across the assembly tree and a new revision of the top
   assembly for a display setting — inside the J30/J31 and J34/J35 admin-freeze
   gates.
6. **Creating** missing `MODEL` reference sets writes to the 404 component part
   files themselves: 404 revised parts, each possibly used by assemblies outside
   this analysis. Much larger blast radius; requires the admin-freeze path.
7. Per-component the change is a cheap toggle back, but not after check-in. A
   mistake costs revisions, not seconds.

**Tier 3 — measurement**

8. No baseline exists: this run is `PROBE`, `J37_TIMING_*.csv` was never
   produced, and `RepresentationMode` returned `<unavailable>=2947` for every
   occurrence. It is not known whether the expensive path is Exact or
   Lightweight, so the refset share of the 10 minutes cannot be predicted.
9. Report inconsistency: `totals.entire_part_refset_occurrences` = 1,414 but the
   404 `ENTIRE_PART_REFSET` suspects sum to 1,372 (Δ42, ~3%). One Entire-Part
   part has no `Bodies`/`GetFaces` facts. Settle this before using the number as
   an acceptance criterion.
10. It is one of ten levers. `ShadedWithEdges`, unconfirmed stored facets, 3,674
    constraints, 100 dynamic sections and the 51,495-face part may dominate.
11. Everyone else opening `264MN021888A01/A` in Teamcenter sees different content
    afterwards — a standards decision that needs announcing.

## Problem part list

### A. `Entire Part` reference set — cost concentrated in 12 parts

| Part | Occ | Bodies | Faces | Occ × faces |
| --- | --- | --- | --- | --- |
| `264MN034794A01/A` | 43 | 1 | 477 | **20,511** |
| `264LN035633A01/A` | 6 | 1 | 1,915 | 11,490 |
| `993UN00002A01/E` | 1 | 1 | 11,474 | 11,474 |
| `264MN025204A01/A` | 1 | 101 | 8,790 | 8,790 |
| `264MN021171A01/A` | 2 | 5,685 | 3,010 | 6,020 |
| `264MN020438A01/A` | 2 | 11 | 2,290 | 4,580 |
| `264MN025139A01/A` | 2 | 1 | 2,243 | 4,486 |
| `1698UN207797A01/C` | 1 | 32 | 4,170 | 4,170 |
| `264MN032846A01/A` | 17 | 1 | 222 | 3,774 |
| `264LN034947A01/A` | 2 | 1 | 1,857 | 3,714 |
| `264LN034957A01/A` | 2 | 1 | 1,778 | 3,556 |
| `264LN035646A01/A` | 7 | 1 | 502 | 3,514 |

**Top 12 = 86 of 1,372 occurrences (6%) but 86,079 of 174,087 occurrence-weighted
faces (49%).** This is the pilot set.

Breadth over depth: `264MN021420A01/A` (82), `264MN028285A01/A` (50),
`264MN028286A01/A` (50), `264MN027026A01/A` (35), `264MN028301A01/A` (25),
`032625/A` (23), `264MN035957A01/A`, `031926/A`, `264MN030221A01/A` (19 each).

### B. Hot geometry (not Entire Part)

`264MN025484A01/A` 51,495 faces / 1,655 blanked (drawn ≈ 47),
`032609/A` 31,637 / 2,203 bodies, `264MN025228A01/A` 5,193 × 4 = 20,772,
`993UN00908A01/A` 5,625 × 2 = 11,250, `264MN020474A01/A` 8,607,
`1698UN177469A01/B` 7,534. Only the first two multiply.

> Note: `HIGH_FACE_GEOMETRY` counts **part** faces, not drawn faces.
> `264MN025484A01/A` occurs once and 1,655 of its 1,702 bodies are blanked, so
> ~47 bodies are drawn. It is top of the file-size/update-time list and near the
> bottom of the view-change list.

### C. Non-solid clutter

`264MN024578A01/A` **8,174** (worst by 4×), `264MN025997A01/A` 1,939,
`264MN021549A01/A` 1,586, `264MN024101A01/A` 1,586, `264MN035282A01/A` 942,
`264MN019766A01/A` 744, `264MN025370A01/A` 683, `264MN021858A01/A` 654,
`264MN025344A01/A` 585, `264MN035149A01/A` 580, `264MN035096A01/A` 566,
`264MN024580A01/A` 538, `264MN024848A01/A` 535.

### D. Dynamic sections — 100 across 69 parts

`264MN024625A01/A` **14**, root `264MN021888A01/A` **10**, then 2 each:
`264MN021541A01/A`, `264MN024578A01/A`, `264MN030208A01/A`,
`264MN032797A01/A`, `264MN032800A01/A`, `264MN034340A01/A`,
`264MN035269A99/B`, `264MN035270A99/A`. Dynamic sections are session view
objects, not model data.

### E. Convergent/faceted + junk density — one cluster

`264MN035702A01/A`, `264MN035868A01/A`, `264MN035870A01/A`,
`264MN036688A01/A` (occ 4) all report density `min=1,470` /
`max=78,306,400,000` — physically impossible; plus `264MN035931A01/A`,
`264MN036282A01/A`, `264MN036288A01/A` (facets 644–1,380). Data-quality defect
that also poisons J18/J21 mass rollups.

### F. Zero-density bodies — 78 bodies across 77 parts

Including the **root `264MN021888A01/A` (2 bodies)**, `027024/A` (2),
`264MN025241A01/A` (2 occurrences), `264MN019761A01/A`, `264MN032259A01/A`,
`264MN032260A01/A`, `264MN035143A01/A`, `264MN035618A01/A` … `264MN035622A01/A`,
`264MN036113A01/A`, `264MN036650A01/A`. Root zero-density makes every rollup
above it wrong.

### G. Load truth

`264MN019808A01/A` PartiallyLoaded. The 46 unreadable prototypes are all root
children, harness/hardware, up to 7 levels deep, e.g. `036460/A`, `028508/A`,
`W240 EXTEND CABLE/A`, `264MN036457A01/A`, `032645/A`, and the
`264MN035688A99/A → 264MN024625A01/A → 264LN035968A99/A → 264LN035866A01/A →
264LN034946A01/A → 264LN035015A01/A` chains. Names are recoverable from the
`owner` field of the `PartLoadState` facts even though the suspects show
`part/-`.

### H. Layers

Root `264MN021888A01/A`: **256 visible layers, 3 populated**.

## Quick wins, ranked

All of these are session-level — no check-in, no revision, instantly reversible.

| # | Action | Kills | Risk |
| --- | --- | --- | --- |
| 1 | View rendering style → `Shaded` (from `ShadedWithEdges`) | edge stroking over the whole face count | none |
| 2 | Hide constraints (Assembly Navigator → Constraints) | 3,674 `DisplayedConstraint` = 12.1% of the view | none |
| 3 | Delete the 100 dynamic sections (root 10, `264MN024625A01/A` 14) | per-view section evaluation | none |
| 4 | Hide the clutter parts — `264MN024578A01/A` first, then `264MN025997A01/A` | ~8,174 + 1,939 objects | hide **without saving** to stay session-only |
| 5 | Hide the 253 unused layers on the root | per-view layer walk | none |
| 6 | Fully load `264MN019808A01/A` | load-on-demand during view changes | load only |

### Evidence gaps and corrections

- **"Turn ON Render solids using stored facets" is not a free win, and its
  evidence is thin.** `PREF_RenderSolidsUsingStoredFacets` is `UNAVAILABLE` for
  **1,347 of 1,348** probes (`property not exposed: NXOpen.Preferences.
  SessionPreferences has no attribute…`); the single `OK` reading is
  `session.PerformanceVisualization = False` from a fallback path.
- `PREF_SaveAdvancedDisplayFacets = False` on **all 673 readable parts**, so no
  facets are stored in the parts at all. Enabling the rendering preference alone
  is a no-op until facets are generated **and the parts are saved** — a governed
  batch change that also grows file size. Plan a targeted facet save only for the
  hot parts (`032609/A`, `264MN025228A01/A`, `993UN00002A01/E`,
  `264MN025204A01/A`).
- `PREF_ShowFacetEdges`, both shading tolerances and all five load-on-demand
  preferences are also `UNAVAILABLE`, so half the display-preference story is
  unmeasured. Re-run with `NX_J37_DISCOVER=YES` before acting on any of them.
- `RepresentationMode` is unmeasured for all 2,947 occurrences.
- `J37_OCCURRENCES_*.csv` is referenced by the report but not committed; it is
  where the 1,414 vs 1,372 discrepancy gets resolved.
- No per-prototype **drawn**-body metric exists in J37 — only part-level faces.
  This is why `264MN025484A01/A` ranks as "hot geometry" while contributing
  ~47 drawn bodies.

## Next actions

1. **Cheap, no approval:** quick wins 1–6 above.
2. **Then measure:** `NX_J37_MODE=TIMED NX_J37_ROTATIONS=1` **and**
   `NX_J37_DISCOVER=YES` on the fully loaded assembly → baseline seconds plus
   corrected probe paths.
3. **Commit `J37_OCCURRENCES_*.csv`** from the same run; resolve 1,414 vs 1,372.
4. **Read-only refset inventory for the 404 parts** — available ref sets and
   their member body counts, reusing the machinery already in
   `from_git/journals/23_diagnose_hla_visibility.py`
   (`REFSET_FOUND`, `REFSET_BODY_MEMBERS`, `REFSET_COMPONENT_MEMBERS`). Output is
   the exclusion list: no `MODEL` set, `MODEL` set with 0 bodies, harness parts
   whose routable geometry sits outside `MODEL`.
   The machinery is already proven on real data: `docs/J23_HLA_VISIBILITY_264MN024625A01_20260814_115209.json`
   carries 1,653 `ranked_occurrences` rows each with `REFERENCE_SET`,
   `REFERENCE_SET_FOUND` and `REFERENCE_SET_MEMBER_COUNT` — for example
   `264MN025454A01/A` resolves to `REFERENCE_SET: MODEL`,
   `REFERENCE_SET_FOUND: YES`, `REFERENCE_SET_MEMBER_COUNT: 1566`. J37 agrees on
   the same part (`FullyLoaded`, 1 body, 446 faces, 27 features, 1 dynamic
   section) and does **not** list it under `ENTIRE_PART_REFSET` — so the refset
   probe and the triage report are consistent, and this part is not in the
   problem set.
5. **Pilot the 12 parts in list A** in a scratch copy, `DRY_RUN` first per house
   style. Acceptance = before/after J37 diff on targets/occurrences/visible
   census, no constraint or update errors, drawings unchanged.
6. **Then the full 404** through the admin-freeze path with the reversal
   documented.
7. **Separately (correctness, not performance):** zero-density bodies, the
   convergent/density cluster, the 46 unreadable prototypes, and the
   `264MN025484A01/A` blanking review — the last one needs a Teamcenter
   where-used first, because 1,655 blanked bodies of 1,702 is the signature of a
   multi-body master part whose other bodies are used by other assemblies.
   **Do not delete unused/blanked features in that part without where-used
   evidence and a backup.**

## J37 v3 candidates

- Report unreadable load state as its own `LOAD_STATE_UNREADABLE` (MEDIUM)
  instead of merging it into `NOT_FULLY_LOADED` (HIGH) — 46 of the 47 findings
  are "unreadable", not "not loaded".
- Map `RenderingStyle` `0` → `ShadedWithEdges` in the report instead of a bare
  `0`.
- Add a share-threshold suspect so a single visible type above ~5% self-reports
  — this run would then have flagged the 3,674 constraints automatically.
- Add occurrence-weighted faces/bodies per prototype and a drawn-vs-blanked
  split, so `HIGH_FACE_GEOMETRY` cannot rank fully blanked parts as hot.
- Include the unnamed prototypes' component paths in the suspect rows (the data
  exists in the `owner` field but not in the suspect list).

## Re-running

```text
NX_J37_MODE=PROBE            # default: inventory + ranked suspects + evidence ledger
NX_J37_MODE=TIMED            # also times Regenerate / UpdateDisplay / Rotate / Fit, restores view
NX_J37_SCOPE=ALL|BOM
NX_J37_DISCOVER=YES          # dumps real preference/component member names for probe paths
NX_J37_ROTATIONS=<n>
```

Run on the assembly already open and fully loaded. Outputs land in
`NX_DISPLAY_PERF\<timestamp>\REPORTS\` plus `LOGS\J37_LOG_<timestamp>.txt`; the
JSON evidence file is the artifact to commit.
