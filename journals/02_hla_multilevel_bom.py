# Attributes pulled from NX part file directly (not TC) for reliability.
"""
Journal 02 — HLA Multilevel BOM to Excel
Traverses the full assembly tree and exports a multilevel BOM from NX part attributes.
Run via: NX > Tools > Journal > Play
"""

import os
import sys
from collections import Counter
from datetime import datetime

import NXOpen
import NXOpen.UF
import yaml

_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _REPO_ROOT not in sys.path:
    sys.path.insert(0, _REPO_ROOT)

from utils.nx_helpers import get_session, get_work_part, prompt_folder  # noqa: E402
from utils.template_generator import MASTER_COLUMNS, create_workbook_with_headers  # noqa: E402

_MAPPING_FILE = os.path.join(_REPO_ROOT, "config", "attribute_mapping.yaml")


def _load_mapping():
    with open(_MAPPING_FILE, encoding="utf-8") as f:
        cfg = yaml.safe_load(f)
    return cfg.get("columns", {})


def _get_attr(nxobj, attr_name):
    """Read a string user attribute; returns empty string if not found."""
    try:
        attr = nxobj.GetUserAttribute(attr_name, NXOpen.NXObject.AttributeType.String, -1)
        return attr.StringValue.strip()
    except Exception:
        return ""


def _collect_components(root_component):
    """
    Walk the assembly, collecting (component, depth) tuples.
    Also builds a part-number → occurrence count map for QUANTITY.
    """
    rows = []
    part_counter = Counter()

    def _walk(component, depth):
        for child in component.GetChildren():
            proto = child.Prototype
            pn = _get_attr(proto, "PART_NUMBER") if proto else ""
            part_counter[pn or child.Name] += 1
            rows.append((child, depth))
            _walk(child, depth + 1)

    _walk(root_component, 1)
    return rows, part_counter


def _build_row(component, depth, part_counter, col_map):
    """Build a MASTER_COLUMNS-ordered list of values for one component."""
    proto = component.Prototype
    pn = _get_attr(proto, "PART_NUMBER") if proto else ""
    qty = part_counter.get(pn or component.Name, 1)

    values = {}
    for col in MASTER_COLUMNS:
        if col == "LEVEL":
            values[col] = depth
        elif col == "PART_NUMBER":
            values[col] = pn
        elif col == "QUANTITY":
            values[col] = qty
        elif col in ("STATUS", "NOTES"):
            values[col] = ""
        elif proto is not None:
            nx_attr = col_map.get(col, col)
            values[col] = _get_attr(proto, nx_attr)
        else:
            values[col] = ""

    return [values[col] for col in MASTER_COLUMNS]


def main():
    session = get_session()
    part = get_work_part()

    if part is None:
        print("ERROR: No work part loaded.")
        return

    output_folder = prompt_folder("Select BOM Output Folder")
    if output_folder is None:
        print("BOM export cancelled by user.")
        return

    col_map = _load_mapping()

    hla_pn = _get_attr(part, "PART_NUMBER") or os.path.splitext(os.path.basename(part.FullPath))[0]
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"BOM_{hla_pn}_{timestamp}.xlsx"
    output_path = os.path.join(output_folder, filename)

    root = part.ComponentAssembly.RootComponent
    component_rows, part_counter = _collect_components(root)

    workbook, worksheet = create_workbook_with_headers(output_path)
    cell_fmt = workbook.add_format({"border": 1})

    for row_idx, (component, depth) in enumerate(component_rows, start=1):
        for col_idx, value in enumerate(_build_row(component, depth, part_counter, col_map)):
            worksheet.write(row_idx, col_idx, value, cell_fmt)

    workbook.close()
    print(f"BOM exported ({len(component_rows)} components): {output_path}")


if __name__ == "__main__":
    main()
