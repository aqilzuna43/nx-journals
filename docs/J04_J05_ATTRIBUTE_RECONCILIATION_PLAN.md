# J04/J05 BOM, Model, and Drawing Attribute Reconciliation Plan

## Document status

- **Repository:** `aqilzuna43/nx-journals`
- **Target branch:** `main`
- **Runtime:** Siemens NX 2312 embedded Python 3.10
- **Data management:** TeamcenterX / managed NX
- **External Python packages:** None
- **Production-proven journal:** `from_git/journals/07_datapack_pdf_step_export.py`
- **Protected J07 commit:** `6fe58f765b8a207229b2f6990e3b0224caa03771`
- **Primary development targets:** Journal 04 and Journal 05

---

## 1. Objective

Redesign Journal 04 and Journal 05 so that NX attributes are not merely populated, but are demonstrably aligned with:

1. the authoritative BOM,
2. the NX master model,
3. the applicable Teamcenter drawing specification, and
4. the intended drawing information used for manufacturing and release.

The work must create a controlled, auditable reconciliation process rather than a blind attribute-copying tool.

---

## 2. Definition of win

The solution is successful when a part is reported as **PASS** only when all applicable conditions are true:

- The part number and revision in the authoritative BOM match the NX master model.
- The correct model revision is loaded or resolved.
- The expected drawing specification exists when a drawing is required.
- The drawing identity and revision match the BOM and model intent.
- Required attributes exist on the correct NX object.
- Attribute values match the authoritative source after approved normalization.
- BOM parent-child structure and quantity match the NX assembly structure.
- Every mismatch is reported explicitly.
- Journal 05 changes only approved values, verifies the result, and preserves unapproved values.
- A second Journal 04 run after Journal 05 shows the approved items as PASS.
- Journal 07 remains operational and passes its regression test.

A non-empty attribute is **not** sufficient for PASS.

Example:

```text
BOM material:        SUS304
NX model material:   ALUMINIUM
Drawing material:    SUS304

Current J04 result:  PASS because MATERIAL is non-empty
Required J04 result: ATTRIBUTE_MISMATCH
```

---

## 3. Journal portfolio decision

| Journal | Current role | Decision |
|---|---|---|
| J01 | Active-part STEP export | Retain as a small utility |
| J02 | NX assembly BOM snapshot | Retain, but do not use as the independent BOM source of truth |
| J03 | Legacy batch PDF export | Mark as superseded or rebuild using J07 drawing resolution |
| **J04** | Current attribute completeness audit | Redesign as the read-only reconciliation engine |
| **J05** | Current bulk attribute pull/push | Redesign as a controlled correction engine |
| J06 | Legacy active-part PDF + STEP | Mark as superseded by J07 |
| **J07** | Proven production PDF + STEP export | Freeze and protect |
| J08 | NX X identity diagnostic probe | Move or label as diagnostic-only |
| J09 | Canonical drawing-open proof | Retain as a permanent regression test |

J07 must not be refactored as part of J04/J05 work unless a separately approved change is required.

---

## 4. Existing problems

### 4.1 Journal 04 limitations

The current J04:

- checks only a small set of attributes,
- checks whether values are present rather than correct,
- returns only the first failing attribute,
- does not consume an independent authoritative BOM,
- does not open or inspect Teamcenter drawing specifications,
- does not compare model values with drawing values,
- calculates quantity globally rather than by immediate BOM parent,
- cannot distinguish missing, unreadable, mismatched, and not-applicable values.

Current logic is effectively:

```text
Attribute is non-empty = PASS
```

Required logic is:

```text
Applicable attribute exists
AND value matches the configured authoritative source
AND model and drawing are mutually consistent where required
= PASS
```

### 4.2 Journal 05 limitations

The current J05:

- selects the newest `Att-*.csv` automatically,
- matches parts by part number without requiring revision,
- writes only when an NX attribute is empty,
- keeps an existing but incorrect NX value,
- does not require engineering approval for each change,
- does not target model and drawing objects independently,
- does not perform a structured post-write verification,
- does not use J04 findings as its correction basis,
- has no formal dry-run-to-apply workflow.

Current incorrect-value behavior:

```text
BOM value:        SUS304
NX current value: ALUMINIUM
J05 action:       KEPT
```

Required behavior:

```text
J04 action:       ATTRIBUTE_MISMATCH
J05 dry run:      PROPOSED_UPDATE, ALUMINIUM -> SUS304
J05 apply:        update only when explicitly approved
J05 verification: reread value and confirm SUS304
```

### 4.3 Circular comparison risk

J02 reads the NX assembly and exports an NX-derived BOM snapshot. It cannot be the only source used to validate NX attributes because that would create a circular comparison:

```text
NX data -> J02 CSV -> compare against the same NX data
```

J04 requires an independent authoritative BOM input, such as an approved Teamcenter BOM export, ERP BOM export, or controlled project BOM file.

---

## 5. Design principles

1. **Audit before update.** J04 must be trusted before J05 is allowed to write.
2. **Read-only by default.** J04 must never save or modify NX parts.
3. **Dry run by default.** J05 must not write unless explicitly placed in apply mode.
4. **Match by part number and revision.** Part number alone is insufficient.
5. **Identify the target object.** MODEL and DRAWING must be treated as separate NX objects.
6. **Use proven Teamcenter drawing resolution.** Drawing specifications use:

   ```text
   @DB/<PART>/<REV>/specification/<PART>-<REV>-dwg<n>
   ```

7. **Use `session.Parts.OpenDisplay()` for drawing specifications.** This method was proven by J09.
8. **Record every comparison.** Do not stop at the first failure.
9. **Normalize explicitly.** Case, whitespace, units, and aliases must be governed by configuration.
10. **No silent fallback.** Missing APIs, missing columns, or unreadable attributes must be reported as explicit statuses.
11. **Minimal production diffs.** Preserve working code and keep each change reviewable.
12. **Restore NX state.** Restore the original display/work parts and close only parts opened by the journal.
13. **No direct Teamcenter API in the first implementation.** Use NXOpen and managed identifiers unless evidence proves otherwise.
14. **No third-party packages.** Use standard-library CSV/JSON and NXOpen only.

---

## 6. Proposed end-to-end workflow

```text
Authoritative BOM export
        |
        v
J04 read-only reconciliation
        |
        +--> PASS rows
        |
        +--> mismatch rows
                 |
                 v
         Engineer review and approval
                 |
                 v
           J05 DRY_RUN
                 |
                 v
       Review proposed before/after values
                 |
                 v
        J05 APPLY_APPROVED
                 |
                 v
       Reread and verify every written value
                 |
                 v
           Rerun J04
                 |
                 v
       Approved scope returns PASS
```

---

## 7. Source-of-truth model

Source ownership must be configurable per attribute. Do not assume that every BOM column is authoritative for every model or drawing property.

Possible source types:

- `BOM`
- `MODEL`
- `DRAWING`
- `TEAMCENTER_EXPORT`
- `DERIVED`
- `MANUAL_REVIEW`

Example ownership rules:

| Logical field | Expected owner | Typical comparison |
|---|---|---|
| Part number | BOM / Teamcenter identity | Exact after trimming |
| Revision | BOM / Teamcenter identity | Exact, case-insensitive |
| Description | BOM or released item revision | Normalized text |
| Material | Drawing intent or approved BOM field | Controlled alias comparison |
| Finish | Drawing intent | Controlled alias comparison |
| Drawing number | Drawing specification | Exact or configured relationship |
| Quantity | Authoritative BOM | Integer comparison per parent |
| UOM | BOM / ERP | Controlled vocabulary |
| Commodity classification | Teamcenter/ERP | Exact or controlled alias |

Any attribute without a confirmed source owner must be marked `MANUAL_REVIEW`, not automatically corrected.

---

## 8. Proposed shared attribute-rule configuration

Replace the disconnected `columns` and `tc_columns` behavior with one shared rule model used by both J04 and J05.

Recommended configuration file:

```text
from_git/config/attribute_reconciliation.json
```

Illustrative schema:

```json
{
  "schema_version": 1,
  "identity": {
    "bom_part_number_columns": ["DB_PART_NO", "Item Number", "PART_NUMBER"],
    "bom_revision_columns": ["DB_PART_REV", "Item Rev", "REVISION"],
    "model_part_number_attributes": ["DB_PART_NO", "PART_NUMBER"],
    "model_revision_attributes": ["DB_PART_REV", "REVISION"],
    "drawing_relation": "specification",
    "drawing_name_template": "{part_number}-{revision}-dwg{index}",
    "max_drawing_index": 9
  },
  "attributes": [
    {
      "logical_name": "PART_NUMBER",
      "bom_columns": ["DB_PART_NO", "Item Number"],
      "model_attributes": ["DB_PART_NO", "PART_NUMBER"],
      "drawing_attributes": ["DB_PART_NO", "PART_NUMBER"],
      "authoritative_source": "BOM",
      "required_on": ["MODEL", "DRAWING"],
      "comparison": "TRIMMED_CASE_INSENSITIVE",
      "write_target": "NONE"
    },
    {
      "logical_name": "MATERIAL",
      "bom_columns": ["Material", "MATERIAL"],
      "model_attributes": ["MATERIAL"],
      "drawing_attributes": ["MATERIAL"],
      "authoritative_source": "DRAWING",
      "required_on": ["MODEL", "DRAWING"],
      "comparison": "NORMALIZED_TEXT",
      "write_target": "MODEL_AND_DRAWING",
      "allowed_aliases": {
        "SUS304": ["SUS 304", "SS304", "STAINLESS STEEL 304"]
      }
    }
  ]
}
```

Required rule properties:

- logical name,
- authoritative source,
- BOM column aliases,
- model attribute aliases,
- drawing attribute aliases,
- applicability,
- comparison mode,
- normalization rules,
- write permission,
- target object,
- whether a missing value is an error,
- whether conflict requires manual review.

---

## 9. Journal 04 v2 requirements

### 9.1 Purpose

J04 v2 is the trusted, read-only reconciliation engine.

It compares:

- authoritative BOM records,
- NX assembly occurrences,
- NX master model attributes,
- Teamcenter drawing specification attributes,
- configured drawing intent rules.

### 9.2 Inputs

Required:

- Authoritative BOM CSV with a fixed, explicit filename or command-line/environment override.
- Active HLA assembly in NX 2312.
- `attribute_reconciliation.json`.

Recommended default BOM filename:

```text
NX_ATTRIBUTE_BOM_SCOPE.csv
```

Do not automatically select the newest arbitrary CSV.

### 9.3 BOM identity

Each BOM line must resolve at minimum:

- parent part number,
- parent revision when available,
- child part number,
- child revision,
- quantity,
- level or parent-child relationship.

Matching key:

```text
PARENT_PART_NO + PARENT_REV + CHILD_PART_NO + CHILD_REV
```

For flat BOMs without parent information, the report must state that structural reconciliation is unavailable.

### 9.4 NX assembly traversal

J04 must:

- traverse the active assembly recursively,
- skip suppressed occurrences unless the BOM explicitly includes them,
- preserve immediate parent-child relationships,
- calculate quantity per immediate parent,
- distinguish occurrence data from prototype data,
- detect unloaded or invalid prototypes,
- collect unique model prototypes by part number and revision.

### 9.5 Model resolution

For each BOM part/revision:

1. Reuse an exact loaded model when present.
2. Verify model identity from `DB_PART_NO + DB_PART_REV` or configured aliases.
3. Report revision mismatch rather than silently selecting a different revision.
4. Do not modify or save the model.

Opening unloaded master parts can be introduced only after a separate minimal test proves the exact supported identifier and NX API.

### 9.6 Drawing resolution

For each applicable part/revision:

1. Reuse an exact drawing specification already loaded.
2. Otherwise generate sequential canonical identifiers:

   ```text
   @DB/<PN>/<REV>/specification/<PN>-<REV>-dwg1
   @DB/<PN>/<REV>/specification/<PN>-<REV>-dwg2
   ...
   ```

3. Open using `session.Parts.OpenDisplay()`.
4. Confirm the returned `JournalIdentifier`.
5. Confirm at least one drawing sheet when a drawing is required.
6. Read configured drawing attributes.
7. Restore the original display/work state.
8. Close only drawing parts opened by J04.

### 9.7 Comparison output

J04 must report every applicable logical attribute, not only the first failure.

Recommended detailed report columns:

```text
RUN_TIMESTAMP
BOM_SOURCE_FILE
PARENT_PART_NO
PARENT_REV
PART_NUMBER
REVISION
BOM_LEVEL
BOM_QUANTITY
NX_QUANTITY
MODEL_FOUND
MODEL_IDENTIFIER
DRAWING_REQUIRED
DRAWING_FOUND
DRAWING_IDENTIFIER
LOGICAL_ATTRIBUTE
AUTHORITATIVE_SOURCE
EXPECTED_VALUE
MODEL_VALUE
DRAWING_VALUE
NORMALIZED_EXPECTED
NORMALIZED_MODEL
NORMALIZED_DRAWING
COMPARISON_RESULT
FAILURE_CODE
MESSAGE
```

Recommended summary report:

```text
TOTAL_BOM_LINES
TOTAL_NX_OCCURRENCES
PASS_PARTS
FAIL_PARTS
MISSING_MODEL
MISSING_DRAWING
REVISION_MISMATCH
BOM_STRUCTURE_MISMATCH
QUANTITY_MISMATCH
ATTRIBUTE_MISSING
ATTRIBUTE_MISMATCH
DRAWING_INTENT_MISMATCH
MANUAL_REVIEW
UNREADABLE_ATTRIBUTE
```

### 9.8 Failure codes

Use deterministic status codes:

- `PASS`
- `MISSING_MODEL`
- `MISSING_DRAWING`
- `REVISION_MISMATCH`
- `BOM_STRUCTURE_MISMATCH`
- `QUANTITY_MISMATCH`
- `ATTRIBUTE_MISSING_MODEL`
- `ATTRIBUTE_MISSING_DRAWING`
- `ATTRIBUTE_MISMATCH`
- `MODEL_DRAWING_MISMATCH`
- `DRAWING_INTENT_MISMATCH`
- `UNREADABLE_ATTRIBUTE`
- `UNSUPPORTED_ATTRIBUTE_TYPE`
- `AMBIGUOUS_MATCH`
- `MANUAL_REVIEW`
- `NOT_APPLICABLE`

### 9.9 J04 safety requirements

J04 must not:

- call `SetUserAttribute`,
- save a part,
- change Teamcenter data,
- create datasets,
- select a different revision silently,
- close parts that were already loaded before the journal started.

---

## 10. Journal 05 v2 requirements

### 10.1 Purpose

J05 v2 is a controlled correction engine that acts only on engineer-approved J04 findings.

It must not independently decide that a mismatch should be overwritten.

### 10.2 Modes

Required modes:

```text
PULL
DRY_RUN
APPLY_APPROVED
```

Default mode:

```text
DRY_RUN
```

#### PULL

- Export current model and drawing attribute values.
- Include part number and revision.
- Include target object and managed identifier.
- Do not modify NX.

#### DRY_RUN

- Read an approved correction CSV.
- Resolve model/drawing target objects.
- Compare current and expected values.
- Produce proposed actions.
- Make no changes.

#### APPLY_APPROVED

- Process only rows explicitly marked approved.
- Match by part number and revision.
- Target the exact configured object.
- Write only attributes permitted by configuration.
- Reread each value immediately.
- Record verification result.
- Save only according to an explicitly approved save policy.

### 10.3 Approved correction input

Recommended columns:

```text
APPROVED
PART_NUMBER
REVISION
TARGET_OBJECT
DRAWING_INDEX
LOGICAL_ATTRIBUTE
NX_ATTRIBUTE_NAME
CURRENT_VALUE_FROM_AUDIT
EXPECTED_VALUE
AUTHORITATIVE_SOURCE
AUDIT_RUN_ID
ENGINEER
APPROVAL_NOTE
```

Valid target objects:

- `MODEL`
- `DRAWING`

Valid approval value:

```text
YES
```

All other values must be treated as not approved.

### 10.4 Matching rules

J05 must require:

```text
DB_PART_NO + DB_PART_REV
```

Part-number-only matching is prohibited.

Before writing, it must confirm:

- target part/revision is exact,
- target object type matches,
- current value still equals the value seen by J04, unless a configured override is approved,
- expected value is not blank unless blanking is explicitly supported,
- the attribute is writable under configuration.

If the current value changed after J04, return:

```text
STALE_AUDIT_VALUE
```

and do not write.

### 10.5 Write behavior

J05 must support these actions:

- `NO_CHANGE_ALREADY_MATCHES`
- `PROPOSED_UPDATE`
- `UPDATED_VERIFIED`
- `UPDATED_VERIFICATION_FAILED`
- `SKIPPED_NOT_APPROVED`
- `SKIPPED_NOT_WRITABLE`
- `SKIPPED_TARGET_NOT_FOUND`
- `SKIPPED_REVISION_MISMATCH`
- `STALE_AUDIT_VALUE`
- `ERROR`

Existing non-empty values may be updated only when the row is approved and all safety checks pass.

### 10.6 Save policy

Saving managed parts may create lifecycle and checkout implications. The first implementation must separate attribute mutation from save behavior.

Recommended configuration:

```json
{
  "save_policy": "NO_SAVE"
}
```

Supported future policies after proof:

- `NO_SAVE`
- `SAVE_CHANGED_PARTS`

`SAVE_CHANGED_PARTS` must not be enabled until a minimal controlled test proves checkout, save, failure recovery, and Teamcenter behavior.

### 10.7 J05 report

Recommended columns:

```text
RUN_TIMESTAMP
MODE
AUDIT_RUN_ID
PART_NUMBER
REVISION
TARGET_OBJECT
TARGET_IDENTIFIER
LOGICAL_ATTRIBUTE
NX_ATTRIBUTE_NAME
AUDITED_CURRENT_VALUE
ACTUAL_CURRENT_VALUE
EXPECTED_VALUE
APPROVED
ACTION
WRITE_ATTEMPTED
REREAD_VALUE
VERIFICATION_RESULT
SAVE_RESULT
MESSAGE
```

---

## 11. Implementation phases and gates

### Phase 0 - Repository classification

Tasks:

- Freeze and document J07.
- Label J08 and J09 as diagnostics/regression journals.
- Mark J03 and J06 as superseded.
- Document J04/J05 as redevelopment targets.

Gate:

- No runtime behavior changes.

### Phase 1 - Evidence capture

Use one representative known-good part and one known mismatch.

Capture for both model and drawing:

- `Name`
- `Leaf`
- `FullPath`
- `JournalIdentifier`
- all relevant user attributes,
- attribute types,
- BOM expected values,
- drawing sheet count,
- exact drawing title-block-linked attributes where accessible.

Gate:

- Internal attribute names and source ownership are confirmed, not guessed.

### Phase 2 - Shared configuration

Tasks:

- Create `attribute_reconciliation.json`.
- Define identity, applicability, comparison, normalization, and write rules.
- Add schema validation using standard Python only.

Gate:

- Configuration loads successfully in NX 2312.
- Unknown or malformed rules produce explicit errors.

### Phase 3 - J04 single-part proof

Tasks:

- Hard-code or isolate one BOM line.
- Resolve one loaded model.
- Open one drawing using the proven `/specification/` identifier.
- Compare one attribute.
- Restore NX state.
- Produce one detailed report row.

Gate:

- Known-good part returns PASS.
- Known mismatch returns the correct failure code.

### Phase 4 - J04 full reconciliation

Tasks:

- Parse full BOM CSV.
- Traverse assembly by parent-child relationship.
- Compare quantity per parent.
- Resolve all applicable drawings.
- Compare every configured attribute.
- Produce detailed and summary reports.

Gate:

- Test matrix passes.
- J04 remains read-only.

### Phase 5 - J05 DRY_RUN

Tasks:

- Consume approved correction CSV.
- Resolve exact target objects.
- Validate stale-audit protection.
- Generate proposed before/after report.
- Make no changes.

Gate:

- Proposed actions match engineering expectation exactly.

### Phase 6 - J05 APPLY_APPROVED without save

Tasks:

- Write only approved changes.
- Reread and verify.
- Restore NX state.
- Produce complete report.

Gate:

- Approved changes verify in memory.
- Unapproved and stale rows remain unchanged.

### Phase 7 - Controlled save proof

Tasks:

- Use one test part.
- Prove managed checkout/save behavior.
- Prove error recovery.
- Prove no unintended parts are saved.

Gate:

- Save policy is approved before production use.

### Phase 8 - End-to-end acceptance

Sequence:

```text
Run J04
Review failures
Approve corrections
Run J05 DRY_RUN
Review proposal
Run J05 APPLY_APPROVED
Rerun J04
Run J09 drawing-open regression
Run J07 production regression
```

Gate:

- Approved scope passes J04.
- J07 remains operational.

---

## 12. Test matrix

### 12.1 J04 tests

| Case | Expected result |
|---|---|
| Exact BOM/model/drawing match | PASS |
| Attribute blank on model | ATTRIBUTE_MISSING_MODEL |
| Attribute blank on drawing | ATTRIBUTE_MISSING_DRAWING |
| Non-empty wrong model value | ATTRIBUTE_MISMATCH |
| Model and drawing disagree | MODEL_DRAWING_MISMATCH |
| Drawing required but absent | MISSING_DRAWING |
| Drawing not required | NOT_APPLICABLE for drawing checks |
| Correct part, wrong revision loaded | REVISION_MISMATCH |
| Duplicate occurrence under same parent | Correct parent-level quantity |
| Same part used under two parents | Separate quantity result per parent |
| BOM item missing from NX | MISSING_MODEL or BOM_STRUCTURE_MISMATCH |
| NX item missing from BOM | BOM_STRUCTURE_MISMATCH |
| Suppressed occurrence excluded | Matches configured suppression rule |
| Attribute cannot be read | UNREADABLE_ATTRIBUTE |
| Two candidate models for same identity | AMBIGUOUS_MATCH |
| Alias values such as SS304/SUS304 | PASS only when configured alias exists |
| Whitespace/case-only difference | Result follows configured normalization |
| J04 failure midway | Original NX display/work state restored |

### 12.2 J05 tests

| Case | Expected result |
|---|---|
| Dry run approved mismatch | PROPOSED_UPDATE, no NX change |
| Apply approved mismatch | UPDATED_VERIFIED |
| Already matching value | NO_CHANGE_ALREADY_MATCHES |
| Row not approved | SKIPPED_NOT_APPROVED |
| Wrong revision | SKIPPED_REVISION_MISMATCH |
| Target not found | SKIPPED_TARGET_NOT_FOUND |
| Attribute not writable | SKIPPED_NOT_WRITABLE |
| Current value changed after J04 | STALE_AUDIT_VALUE |
| Verification reread differs | UPDATED_VERIFICATION_FAILED |
| Model target requested | Only model object is touched |
| Drawing target requested | Only drawing specification is touched |
| Failure on one row | Remaining rows continue safely |
| Journal-opened drawing | Closed after processing |
| Preloaded drawing | Remains loaded |
| No-save policy | No part is saved |

### 12.3 Regression tests

- J09 must continue to return `FINAL STATUS: SUCCESS`.
- J07 must continue to open `/specification/` drawings and export requested PDF/STEP outputs.
- J04/J05 must not change J07 output naming or CSV behavior.

---

## 13. Acceptance criteria

### J04 acceptance

- Reads the authoritative BOM input deterministically.
- Matches BOM, model, and drawing by part number and revision.
- Uses parent-level quantity comparisons.
- Reads model and drawing attributes separately.
- Opens drawing specifications using the J09-proven path.
- Reports all discrepancies.
- Produces deterministic detailed and summary CSVs.
- Makes no NX changes.
- Restores NX state after success or failure.

### J05 acceptance

- Defaults to dry run.
- Requires explicit row-level approval.
- Matches by part number and revision.
- Enforces target object and write permissions.
- Prevents stale-audit overwrites.
- Rereads and verifies every attempted write.
- Reports before, expected, actual, and final values.
- Does not save unless an approved save policy is enabled.
- Leaves unapproved data untouched.

### End-to-end acceptance

```text
J04 before J05: mismatches correctly identified
J05 dry run: proposed changes correct
J05 apply: approved changes written and verified
J04 after J05: approved scope PASS
Unapproved differences: unchanged
J09: SUCCESS
J07: production regression PASS
```

---

## 14. Required evidence before implementation

The following must be provided or captured before finalizing the rule configuration:

1. One representative authoritative BOM CSV.
2. One known-good part/revision.
3. One known mismatch part/revision.
4. Expected drawing specification name for each test item.
5. Confirmed internal model attribute names.
6. Confirmed internal drawing attribute names.
7. Identification of which fields are true drawing intent versus copied metadata.
8. Approved source owner for each logical attribute.
9. Expected normalization and alias rules.
10. Confirmation of whether J05 should eventually save managed parts.

---

## 15. Development instructions for Codex or another implementation agent

1. Read this document and the current J04/J05/J07/J09 code before editing.
2. Do not modify J07 during initial J04/J05 development.
3. Do not begin with full batch behavior.
4. Build a minimal single-part, single-attribute, single-drawing proof first.
5. Present the proposed configuration and report schema before writing production code.
6. Keep J04 read-only.
7. Keep J05 dry-run by default.
8. Do not infer proprietary attribute names; capture them from NX evidence.
9. Treat unavailable APIs as `NOT_AVAILABLE`, not `False`.
10. Include exact NX exception messages and `ErrorCode` values in logs.
11. Use the canonical drawing identifier proven by J09.
12. Restore display/work state in `finally` blocks.
13. Close only parts opened by the journal.
14. Perform local Python syntax validation before committing.
15. NX runtime success must be confirmed by the user; do not claim runtime proof from static review.
16. Push implementation commits to `main` only after the user approves the plan or phase.

---

## 16. Recommended first development task

Do not repair the current J05 write behavior first.

Start with:

> Build J04 v2 as a read-only single-part reconciliation proof that reads one authoritative BOM row, resolves one exact NX model and one exact Teamcenter drawing specification, compares identity plus one confirmed attribute, reports all three values, restores NX state, and produces a deterministic CSV result.

Once J04 reliably identifies what is wrong, J05 can safely become the controlled executor.

---

## 17. Completion checklist

### Planning

- [x] Define production journal status.
- [x] Define J04/J05 objective.
- [x] Define definition of win.
- [x] Define safety principles.
- [x] Define implementation phases.
- [x] Define acceptance criteria.

### Evidence

- [ ] Obtain authoritative BOM sample.
- [ ] Select known-good test item.
- [ ] Select known-mismatch test item.
- [ ] Capture model attributes.
- [ ] Capture drawing attributes.
- [ ] Confirm source ownership.
- [ ] Confirm normalization rules.

### J04 v2

- [ ] Create shared reconciliation configuration.
- [ ] Create single-part proof.
- [ ] Validate canonical drawing open.
- [ ] Validate one known PASS.
- [ ] Validate one known mismatch.
- [ ] Implement full BOM structure comparison.
- [ ] Implement full attribute comparison.
- [ ] Produce detailed report.
- [ ] Produce summary report.
- [ ] Confirm read-only behavior.

### J05 v2

- [ ] Define approved correction CSV.
- [ ] Implement PULL.
- [ ] Implement DRY_RUN.
- [ ] Implement stale-audit protection.
- [ ] Implement APPLY_APPROVED.
- [ ] Implement reread verification.
- [ ] Prove no-save behavior.
- [ ] Prove controlled save behavior separately if required.

### Regression

- [ ] J09 success confirmed.
- [ ] J07 PDF regression confirmed.
- [ ] J07 STEP regression confirmed.
- [ ] Original NX state restored after all test cases.
