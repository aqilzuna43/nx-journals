# NX Assembly Load Diagnostic

`NX_Assembly_Load_Diagnostic.vb` is a read-only NXOpen VB.NET journal for
finding the assembly occurrence that can cause STEP export to fail with
`IM0541: An operation was attempted on an invalid or unsuitable OM object`.

## Supported environment

- Siemens NX X 2506
- Teamcenter X managed assemblies
- Local NX assemblies

The journal does not load, save, check out, replace, suppress, or otherwise
modify parts. It inspects the active work-part assembly and continues after an
individual occurrence raises an NXOpen exception.

## Run

1. Open the same top-level assembly that fails during STEP export.
2. Apply the intended Teamcenter revision rule and normal assembly load
   options. Do not manually repair or remove the failing occurrence first.
3. In NX, select **Tools > Journal > Play** (`Alt+F8`).
4. Select `NX_Assembly_Load_Diagnostic.vb`.
5. Follow progress in the NX Listing Window.

The report is written as:

```text
NX_Assembly_Load_Diagnostic_Report.txt
```

By default it is placed on the current user's Desktop. Set
`NX_JOURNALS_IO_DIR` before starting NX to use another output folder. An
existing report with the same name is replaced by the latest run.

## Result interpretation

| Status | Meaning | First corrective action |
|---|---|---|
| `OK` | The inspected occurrence exposed a usable, fully loaded prototype | None |
| `MISSING_FILE` | A local prototype path was recorded but the `.prt` file is absent | Restore the file or correct assembly search paths |
| `PROTOTYPE_UNAVAILABLE` | The occurrence exists but NX cannot return a part prototype | Check revision rule, access, dataset, and load options |
| `UNLOADED` | The part is not fully loaded, or is minimally/partially loaded | Fully load exact geometry and rerun |
| `INVALID_OBJECT` | An inspected NX call raised the IM0541/invalid OM-object signature | Repair or replace the occurrence shown by `Assembly path` |
| `ERROR` | Another NXOpen error interrupted one inspection operation | Review `Failed operation` and `Exception` |

The report includes every scanned occurrence so that its parent, hierarchy
level, full assembly path, prototype/file identity, Teamcenter Item/Revision
when available, reference set, load state, and inspection result can be
compared. Start with `INVALID_OBJECT`, then `MISSING_FILE`,
`PROTOTYPE_UNAVAILABLE`, and `UNLOADED` entries.

## Important limitation

The journal diagnoses assembly-load and occurrence health; it does not replay
the STEP translator itself. If every occurrence is `OK` but STEP still fails,
use the existing STEP source/geometry diagnostic next because the problem may
be a body, interpart reference, or translator setting rather than component
loading.
