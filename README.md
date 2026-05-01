# NX Open Python Journals

NX Open Python journals for NX 2312 + TC X productivity. Run via **NX > Tools > Journal > Play**. Each journal prompts for input paths at runtime.

## Journals

| # | File | Description |
|---|------|-------------|
| 01 | `01_hla_step_export.py` | Exports the active HLA part to STEP (AP214/AP242) using settings from `config/step_export.yaml` |
| 02 | `02_hla_multilevel_bom.py` | Traverses the full assembly tree and exports a multilevel BOM to Excel from NX part attributes |
| 03 | `03_batch_drawing_pdf.py` | Traverses the assembly, finds all associated drawing sheets, and batch-exports them to PDF |
| 04 | `04_assembly_attribute_audit.py` | Audits required part attributes across the assembly and flags missing/invalid values to Excel |

## How to Run

1. Open NX 2312 with the HLA assembly loaded as the work part.
2. Go to **Tools > Journal > Play** (or press `Alt+F8`).
3. Browse to the desired journal `.py` file in this repository.
4. The journal will prompt for any required folder paths via NX dialogs at runtime.
5. Outputs are written to the folder you select.

## Dependencies

- **NX Open Python** — bundled with NX 2312 (`NXOpen`, `NXOpen.UF`, etc.)
- **xlsxwriter** — for Excel output. Install into NX's Python environment:
  ```
  "<NX_ROOT>\NXBIN\python.exe" -m pip install xlsxwriter
  ```
- **PyYAML** — for reading `config/step_export.yaml`:
  ```
  "<NX_ROOT>\NXBIN\python.exe" -m pip install pyyaml
  ```

## Configuration

Edit `config/step_export.yaml` to control STEP export behaviour before running journal 01.

## Notes

- Journals read NX part attributes directly from the part file, **not** from Teamcenter, for reliability in both connected and offline modes.
- All output files are excluded from version control via `.gitignore`.
