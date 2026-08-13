# AGENTS.md — nx-journals

Guidance for AI agents working in this repository.

## Runtime constraint: NX is NOT on this machine

- Siemens NX (2312 / 2506) runs on a **different office machine**. It is
  **not installed here and cannot be launched or tested from this repo's
  host**.
- Journals in `from_git/journals/` are **NX-embedded Python** scripts: they
  only execute inside NX (Tools > Journal > Play). No unit test here can
  prove NX runtime behavior — tests only verify the code loads, the modes
  parse, and the report structure is built.
- The NX user (Aqil) is the **only execution path**: he runs a journal on
  the office PC and brings back the output (CSV / JSON / Listing Window log).

## Feedback loop (how work gets verified)

1. Agent writes/edits a journal + its unit tests here.
2. `python3 -m pytest tests/ -q` must pass locally (import/parse/mode/logic
   coverage only).
3. Agent commits and pushes to `master`.
4. Aqil runs the journal in NX on the office machine, then **pushes the
   resulting log/CSV/JSON back to the repo** (or pastes it) for the agent to
   troubleshoot.
5. The agent analyzes the real output and iterates.

Do **not** claim a journal "works" based on unit tests alone — the repo's
proof artifacts (e.g. `docs/NX2506_J21_SMOKE_LOG.md`,
`J22_DIAGNOSTIC_*.json`) are the only evidence of real NX behavior.

## Provenance

- NX output files (`*.prt`, `*.log`, `*.syslog`, `*.csv`, `*.txt`, `*.pdf`,
  `*.xlsx`, `*.stp`, `*.step`) are gitignored by default; log/JSON artifacts
  worth keeping are committed explicitly under `docs/` or the repo root.
- `documents/` holds a local NXOpen Intellisense reference (15 MB,
  untracked) — do not commit it.
