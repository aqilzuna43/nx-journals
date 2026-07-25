# 3D Business-Attribute Round Trip and Checkout Gate

## Purpose

Journal 04 shows the current business attributes on every unique 3D master
prototype in the active assembly. The engineer edits that same wide CSV and
Journal 05 validates and applies the approved differences. Extended BOM reads
the same prototype attributes and produces the PLM-facing BOM.

NX/Teamcenter remains authoritative for part number, part name, revision,
quantity, lifecycle, material, dimensions, mass, and roll-up mass. Those
values are context only and are never writable through Journal 05.

## Journal 04 pull

Journal 04 is self-contained, read-only, and requires no drawing-scope file or
repository-local Python import. It traverses the active assembly, ignores
suppressed occurrences, deduplicates prototypes, and writes:

```text
NX_ATTRIBUTE_UPDATE_<root>_<timestamp>.csv
NX_ATTRIBUTE_UPDATE_<root>_<timestamp>.baseline.json
```

The CSV control fields are:

```text
AUDIT_RUN_ID, APPROVED, ENGINEER, APPROVAL_NOTE, PULL_STATUS, PULL_MESSAGE
```

Read-only identity fields are:

```text
Item Number, Part Description, Item Rev
```

Editable business fields are:

```text
UOM, Mfr. Name, Mfr. Part Number, Reference Notes, WAE_VERSION, NX_FINISH,
COMPONENT_CLASS, LIFED, SERIAL_NUMBERED_PART, Temperature_Sensitive,
Hazardous, COMMODITYTYPE, Commodity_Code, Serviceable_item_flag,
Export_Control_Number, Country_of_Origin
```

`APPROVED=YES` authorizes every detected business-field change on that row.
Populate `ENGINEER` on every approved row. Missing business values are blank
and may be filled; unreadable identity or ambiguous prototypes produce
`PULL_STATUS=REVIEW` and cannot be applied.

The hard-coded CAD identities `DB_PART_NO`, `DB_PART_NAME`, and `DB_PART_REV`
are read by exact NX title without relying on a Teamcenter category. Editable
business fields remain exact category/title reads under `WAEItem`.

The baseline sidecar stores exact typed original values. It must remain beside
the edited CSV and must not be edited.

## Journal 05 validation and checkout

Set `NX_ATTRIBUTE_UPDATE_FILE` to the edited Journal 04 CSV. Modes are selected
with `NX_J05_MODE`:

- `DRY_RUN` is the default and performs no checkout, attribute write, or save.
- `APPLY_APPROVED` applies only approved changed rows after all gates pass.

Journal 05 rejects:

- missing or mismatched baseline sidecars and audit IDs;
- duplicate, missing, stale, or edited part/revision identities;
- edits to the read-only part description;
- rows marked `REVIEW`;
- approved rows without an engineer;
- populated-to-blank changes;
- invalid controlled values;
- locked, system-owned, or PDM-based attributes;
- ambiguous loaded prototypes or current NX values different from the J04
  baseline.

`TBC` and `N/A` are accepted for unrestricted text fields. They remain invalid
where an attribute has a controlled domain such as Y/N, UOM, component class,
traceability, stocking type, or commodity type.

In managed mode, apply performs a complete preflight and then explicitly
checks out every changed prototype. No attribute is changed unless all
required checkouts succeed. Journal 05 never steals another user's checkout,
bypasses release/access rules, creates another revision, trusts implicit
autolock, or checks a part in.

After checkout, each prototype is updated under one undo mark. Every value is
reread immediately. A failure rolls back that prototype and prevents its save.
Each verified prototype is saved once. A save failure stops later saves and
leaves the affected part open, modified, and checked out with recovery details.

Reports are written as:

```text
J05_DRY_RUN_<timestamp>.csv
J05_APPLY_APPROVED_<timestamp>.csv
```

They include baseline/current/expected values, checkout state/action/result,
read-only state before/after, write/rollback/verification/save results, and NX
exception/error-code evidence.

## Production save gate

The versioned configuration remains:

```json
"save_policy": "NO_SAVE"
```

With this gate, `APPLY_APPROVED` performs no checkout and no write; it reports
`SAVE_GATE_DISABLED`. Change it to `SAVE_CHANGED_PARTS` only after Journal 11
passes in the deployed NX/Teamcenter runtime.

## Journal 11 acceptance

`11_test_teamcenter_attribute_checkout.py` defaults to the read-only `PROBE`
mode. It records managed-mode, PDM/checkout APIs, autolock, checkout status,
and part read-only state.

The mutating test requires a disposable Teamcenter item and all of:

```text
NX_J11_MODE=FULL_REVERSIBLE
NX_J11_ALLOW_MUTATION=YES
NX_J11_EXPECTED_PART_NUMBER=<exact disposable item>
NX_J11_EXPECTED_REVISION=<exact revision>
NX_J11_ATTRIBUTE=<one title from the business allowlist>
NX_J11_TEST_VALUE=<temporary non-blank value>
```

It verifies this sequence:

1. Match the exact active master part and capture the original value.
2. Explicitly check out the part and prove it is writable.
3. Write, reread, and save the temporary value.
4. Reopen and verify persistence.
5. Restore, reread, and save the original value.
6. Reopen and verify restoration.
7. Leave the part checked out for review.

Evidence is written to `J11_CHECKOUT_ACCEPTANCE_<timestamp>.json`. If
restoration cannot be proven, the result is `RESTORATION_REQUIRED`; leave the
item checked out and restore it manually before further work.

If the deployed Python binding does not expose a proven explicit checkout
path, Journal 05 must require manual checkout. It must never fall back to
implicit autolock.

## Acceptance

Repository tests cover mapping parity, prototype-only reads, unique-prototype
pulls, whole-row approval, protected identities, blanks and controlled values,
stale baselines, checkout failure, no-save behavior, rollback, verification,
save handling, and Journal 11 mutation guards.

Runtime acceptance requires:

1. Run J04 on a representative assembly and review its CSV/sidecar.
2. Edit one disposable business field and run J05 `DRY_RUN`.
3. Run Journal 11 `PROBE`.
4. Run Journal 11 `FULL_REVERSIBLE` on the explicitly guarded disposable item.
5. Confirm temporary persistence, original-value restoration, and that the
   item remains checked out.
6. Only then enable `SAVE_CHANGED_PARTS` and rerun J05 on a controlled case.
7. Re-export Extended BOM and confirm the saved business value is present.
