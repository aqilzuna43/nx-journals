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

_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _REPO_ROOT not in sys.path:
    sys.path.insert(0, _REPO_ROOT)

from utils.nx_helpers import get_session, get_work_part, prompt_folder, traverse_assembly  # noqa: E402
from utils.excel_writer import ExcelWriter  # noqa: E402


_REQUIRED_ATTRS = [
    "PART_NUMBER",
    "DESCRIPTION",
    "MATERIAL",
    "FINISH",
    "REVISION",
    "DRAWING_NUMBER",
]

_BOM_HEADERS = [
    "Level",
    "Part Number",
    "Description",
    "Material",
    "Finish",
    "Revision",
    "Drawing Number",
    "Quantity",
]


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

    hla_pn = _get_attr(part, "PART_NUMBER") or os.path.splitext(os.path.basename(part.FullPath))[0]
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"BOM_{hla_pn}_{timestamp}.xlsx"
    output_path = os.path.join(output_folder, filename)

    root = part.ComponentAssembly.RootComponent
    component_rows, part_counter = _collect_components(root)

    writer = ExcelWriter(output_path)
    writer.add_header_row("BOM", _BOM_HEADERS)

    for component, depth in component_rows:
        proto = component.Prototype
        values = [
            depth,
            _get_attr(proto, "PART_NUMBER"),
            _get_attr(proto, "DESCRIPTION"),
            _get_attr(proto, "MATERIAL"),
            _get_attr(proto, "FINISH"),
            _get_attr(proto, "REVISION"),
            _get_attr(proto, "DRAWING_NUMBER"),
            part_counter.get(_get_attr(proto, "PART_NUMBER") or component.Name, 1),
        ]
        writer.add_row("BOM", values)

    writer.save()
    print(f"BOM exported ({len(component_rows)} components): {output_path}")


if __name__ == "__main__":
    main()
