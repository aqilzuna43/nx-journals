"""Generates the master BOM Excel template used by all journals."""

import os
import xlsxwriter

MASTER_COLUMNS = [
    "LEVEL", "PART_NUMBER", "DESCRIPTION", "REVISION",
    "DRAWING_NUMBER", "MATERIAL", "FINISH", "QUANTITY", "STATUS", "NOTES",
]

_HEADER_BG = "#1a1a2e"
_HEADER_FG = "#ffffff"

_TEMPLATES_DIR = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "templates"
)
_TEMPLATE_NAME = "master_bom_attributes.xlsx"


def get_template_path():
    """Returns absolute path to templates/master_bom_attributes.xlsx"""
    return os.path.join(_TEMPLATES_DIR, _TEMPLATE_NAME)


def ensure_template_exists():
    """Creates the template file if it doesn't exist. Returns path."""
    path = get_template_path()
    if not os.path.exists(path):
        os.makedirs(_TEMPLATES_DIR, exist_ok=True)
        wb, _ = create_workbook_with_headers(path)
        wb.close()
    return path


def create_workbook_with_headers(output_path, sheet_name="BOM"):
    """
    Creates a new xlsxwriter workbook at output_path pre-loaded with
    MASTER_COLUMNS header row styled: bold, white text, #1a1a2e background.
    Returns (workbook, worksheet) tuple — caller must call workbook.close().
    """
    workbook = xlsxwriter.Workbook(output_path)
    worksheet = workbook.add_worksheet(sheet_name)
    header_fmt = workbook.add_format({
        "bold": True,
        "font_color": _HEADER_FG,
        "bg_color": _HEADER_BG,
        "border": 1,
    })
    for col, name in enumerate(MASTER_COLUMNS):
        worksheet.write(0, col, name, header_fmt)
    return workbook, worksheet


if __name__ == "__main__":
    path = ensure_template_exists()
    print(f"Template ready: {path}")
