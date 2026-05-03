# Templates

The journals now write CSV reports directly and no longer generate an Excel
template at runtime. `utils/template_generator.py` only stores the shared
`MASTER_COLUMNS` definition used by BOM and audit CSV outputs.
