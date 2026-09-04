# J34/J35 — CSV Administrative CAD Freeze

J34 and J35 are two NX X 2506 toolbar journals for a controlled, repeatable
Teamcenter freeze campaign. They are freeze-only: neither journal checks parts
in or out, writes attributes, creates revisions, unfreezes CAD, or saves CAD.

Keep the complete package together:

```text
from_git/admin_freeze/
  34_validate_freeze_csv.py
  35_apply_freeze_csv.py
  admin_freeze_common.py
  NX_ADMIN_FREEZE_SCOPE.csv
  reports/                         # created at runtime
```

Both journals resolve the input and manifest relative to their own folder. No
script path needs to be edited.

## Input CSV

Minimum input:

```csv
FREEZE,DB_PART_NO,DB_PART_REV
YES,264MN033038A01,A
NO,264MN000000A01,A
```

`WAE_VERSION` is optional for compatibility with J4's wide CSV:

```csv
FREEZE,DB_PART_NO,DB_PART_REV,WAE_VERSION
YES,264MN033038A01,A,9
YES,993UN00002A01,E,E
```

When a CSV WAE value is present, J34 requires it to equal the live Teamcenter
value. When absent, J34 discovers the live value and records it in the
validation manifest. J35 always requires the live value to remain identical
to the value observed by J34.

Only `FREEZE=YES` rows are eligible. Blank/NO rows are reported once as
`SKIPPED_DISABLED`. Duplicate occurrences of the same case-insensitive
`DB_PART_NO + DB_PART_REV` are collapsed to one result. Conflicting FREEZE or
WAE instructions block that identity without blocking unrelated identities.

## WAE lifecycle rules

- Positive whole numbers (`1`, `2`, ...) are working-iteration baselines.
- An alphabetic WAE is valid only when it matches `DB_PART_REV`, such as
  `A/A`, `B/B`, or `E/E`. It is an immutable final-release baseline.
- Blank/missing WAE is blocked. Initialize an approved numeric value through
  the controlled J4/J5 process before freezing.
- Alphabetic mismatch such as `Rev A / WAE B` is blocked for lifecycle/data
  repair. The freeze journal never guesses or corrects it.
- J30 uses the same WAE classification: positive numeric working values and
  matching alphabetic final baselines may be frozen without changing WAE.
- J31 remains numeric-only and explicitly reports
  `BLOCKED_FINAL_RELEASE_BASELINE` instead of unfreezing an alphabetic final
  baseline. A later engineering change uses the normal Teamcenter revision
  process.

## J34 validation

J34 reads every row, opens only the exact Teamcenter master identity using the
J7-proven `@DB/<part>/<revision>` resolution, and produces timestamped CSV and
JSON reports plus `NX_ADMIN_FREEZE_VALIDATION.json` beside the input.

Validation checks include exact identity, live/CSV WAE, WAE lifecycle class,
checkout, release status, read-only/modifiable state for an existing Frozen
status, and availability of `Part_Freeze_Process`.

Checked-out CAD is skipped and reported with its owner. Another controlled
status such as Released is skipped. J34 never mutates CAD or Teamcenter.

## J35 apply

J35 refuses to start if the input CSV hash differs from the J34 manifest. It
also rechecks live Teamcenter state immediately before every row.

Eligible identities run independently, one part per
`PdmSession.AssignFreezeStatus([part], "Part_Freeze_Process")` call. A failure
does not stop later identities. This isolates Teamcenter errors such as
`3520110` to the exact part/revision. Successful earlier rows remain Frozen;
the journal never attempts an unfreeze rollback.

Final state is authoritative:

- `FROZEN`: workflow returned and Frozen state was verified.
- `FROZEN_WITH_WARNING`: workflow raised an error, but Frozen, checked-in,
  read-only, non-modifiable, identity, and WAE postconditions all passed.
- `ALREADY_FROZEN`: the exact baseline was already positively Frozen.
- `FAILED_FREEZE_WORKFLOW`: Frozen postconditions were not achieved.
- `BLOCKED_*` / `NOT_FOUND`: the row was unsafe or unavailable and was skipped.

Every result includes `MESSAGE` and `RESOLUTION_HINT` for Aqil. Both journals
show a concise NX dialog and Listing Window summary; detailed evidence is
written under `admin_freeze/reports/`.

J34/J35 close only parts they opened. A part that unexpectedly becomes
modified or cannot close is left open and reported.

## Acceptance sequence

NX is not installed on this repository host. Local tests verify parsing,
policy, manifest, and reporting logic only.

1. Deploy the complete `admin_freeze` folder.
2. Leave both template rows as NO and run J34 to prove path/report handling.
3. Add one disposable checked-in numeric target with `FREEZE=YES`; run J34.
4. Review `READY`, the actual WAE, and the manifest/report paths.
5. Run J35 and require `FROZEN` plus native NX Status `Frozen`.
6. Repeat with one matching alphabetic final target.
7. Test one missing WAE, checked-out target, Released target, duplicate row,
   nonexistent identity, and a known `3520110` target. Safe rows must continue
   while each unsafe row receives its specific result and resolution hint.
