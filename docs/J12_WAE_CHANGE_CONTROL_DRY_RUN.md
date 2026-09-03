# J12 WAE Change Control — Option B DRY RUN

`from_git/journals/12_wae_change_control_dry_run.py` is a read-only preflight for the proposed Teamcenter-controlled freeze/unfreeze process.

## Scope

- Keeps `DB_PART_REV` as the formal TCX revision.
- Reads `WAEItem/WAE_VERSION` as the working iteration within that revision.
- Simulates `FREEZE` or `UNFREEZE` without changing NX or Teamcenter data.
- Does not modify Journal 04 or Journal 05.

## Run

By default the journal simulates `UNFREEZE`.

Optional environment variable:

```text
NX_J12_ACTION=FREEZE
```

or

```text
NX_J12_ACTION=UNFREEZE
```

Run the journal in NX X / NX 2506 with the managed master part as the active work part.

## DRY_RUN behavior

The journal reads:

- `DB_PART_NO`
- `DB_PART_REV`
- `WAEItem/WAE_VERSION`
- `CHECKED_OUT_USER`
- `CHECKED_OUT`
- managed-mode evidence (`@DB/` work-part identity or NX managed mode)

For `UNFREEZE`, it simulates:

```text
explicit TCX checkout
-> WAE_VERSION N -> N+1
-> reread/verify
-> save
```

For `FREEZE`, it simulates:

```text
validate/save
-> keep WAE_VERSION unchanged
-> TCX check-in / frozen baseline
```

No production operation is executed.

## Safety guarantees

The current J12 version performs none of the following:

- Teamcenter checkout
- Teamcenter check-in
- NX attribute write
- NX save
- TCX revision creation

It writes only a local JSON audit file, normally under `%TEMP%`, named:

```text
J12_WAE_CHANGE_CONTROL_DRY_RUN_<timestamp>.json
```

## Expected example

For an active part at:

```text
DB_PART_REV=A
WAE_VERSION=3
```

an `UNFREEZE` dry run should report:

```text
SIMULATION: WAE_VERSION 3 -> 4
Production intent: explicit TCX checkout -> increment -> verify -> save
```

while confirming that no mutation was performed.
