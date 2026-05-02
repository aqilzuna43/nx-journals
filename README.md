# NX Open Python Journals

NX Open Python journals for NX 2312 + TC X productivity. Run via **NX > Tools > Journal > Play**. Each journal prompts for input paths at runtime.

## Journals

| # | File | Description |
|---|------|-------------|
| 01 | `01_hla_step_export.py` | Exports the active HLA part to STEP (AP214/AP242) using settings from `config/step_export.yaml` |
| 02 | `02_hla_multilevel_bom.py` | Traverses the full assembly tree and exports a multilevel BOM to Excel from NX part attributes |
| 03 | `03_batch_drawing_pdf.py` | Traverses the assembly, finds all associated drawing sheets, and batch-exports them to PDF |
| 04 | `04_assembly_attribute_audit.py` | Audits required part attributes across the assembly and flags missing/invalid values to Excel |
| 05 | `05_bulk_attribute_updater.py` | Reads a J02/J04 BOM xlsx and writes attribute values back into the open NX assembly |

## How to Run

1. Open NX 2312 with the HLA assembly loaded as the work part.
2. Go to **Tools > Journal > Play** (or press `Alt+F8`).
3. Browse to the desired journal `.py` file in this repository.
4. The journal will prompt for any required folder paths via NX dialogs at runtime.
5. Outputs are written to the folder you select.

## Master Template

All journals share a single column schema defined in `utils/template_generator.py`:

```
LEVEL | PART_NUMBER | DESCRIPTION | REVISION | DRAWING_NUMBER | MATERIAL | FINISH | QUANTITY | STATUS | NOTES
```

The mapping between these column names and NX's internal attribute names is stored in
`config/attribute_mapping.yaml`. NX often exposes attributes under internal titles
(e.g. `DB_TITLE`) that differ from what the attribute editor displays.

### First-time setup

1. Open a representative part in NX — one with all standard attributes populated.
2. **Tools > Journal > Play** → select `utils/discover_attributes.py`.
3. Open the generated `ATTR_DISCOVERY_*.txt` file and match titles against the NX
   attribute editor to find the real internal names.
4. Update the `columns:` section in `config/attribute_mapping.yaml` with the
   discovered names.

### End-to-end workflow

```
J02  →  BOM_*.xlsx        (one row per component, MASTER_COLUMNS schema)
J04  →  AUDIT_*.xlsx      (same schema, STATUS = PASS / FAIL: <attr>)
J05  ←  either xlsx       (pushes non-skip columns back to NX, keyed by PART_NUMBER)
```

J02 and J04 output can be fed directly into J05 with no reformatting required.

## Dependencies

- **NX Open Python** — bundled with NX 2312 (`NXOpen`, `NXOpen.UF`, etc.)
- **xlsxwriter** — for Excel output. Install into NX's Python environment:
  ```
  "<NX_ROOT>\NXBIN\python.exe" -m pip install xlsxwriter
  ```
- **PyYAML** — for reading yaml config files:
  ```
  "<NX_ROOT>\NXBIN\python.exe" -m pip install pyyaml
  ```
- **openpyxl** — for reading xlsx in journal 05:
  ```
  "<NX_ROOT>\NXBIN\python.exe" -m pip install openpyxl
  ```

## Configuration

Edit `config/step_export.yaml` to control STEP export behaviour before running journal 01.

## Notes

- Journals read NX part attributes directly from the part file, **not** from Teamcenter, for reliability in both connected and offline modes.
- All output files are excluded from version control via `.gitignore`.
