"""
Journal 04 — Assembly Attribute Audit
Checks required NX part attributes across the full assembly and reports issues to Excel.
Run via: NX > Tools > Journal > Play
"""

import os
import re
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
from utils.excel_writer import ExcelWriter  # noqa: E402
from utils.template_generator import MASTER_COLUMNS  # noqa: E402

_MAPPING_FILE = os.path.join(_REPO_ROOT, "config", "attribute_mapping.yaml")

_REVISION_PATTERN = re.compile(r"^([A-Z]|\d+)$")

# Columns whose NX attr values are audited for presence and validity.
# These are MASTER_COLUMNS names; the yaml mapping resolves the internal NX name.
_AUDITED_COLS = ["PART_NUMBER", "DESCRIPTION", "MATERIAL", "FINISH", "REVISION", "DRAWING_NUMBER"]

_AMBER = "#FFF3CD"
_GREEN = "#D4EDDA"

_SUMMARY_HEADERS = ["Metric", "Value"]


def _load_mapping():
    with open(_MAPPING_FILE, encoding="utf-8") as f:
        cfg = yaml.safe_load(f)
    return cfg.get("columns", {})


def _get_attr(nxobj, attr_name):
    try:
        attr = nxobj.GetUserAttribute(attr_name, NXOpen.NXObject.AttributeType.String, -1)
        return attr.StringValue.strip()
    except Exception:
        return ""


def _collect_components(root_component):
    """Return (component, depth) tuples and a part-number occurrence Counter."""
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


def _audit_proto(proto, col_map):
    """
    Check required attributes on a prototype.
    Returns the name of the first failing MASTER_COLUMNS column, or None if all pass.
    """
    for col in _AUDITED_COLS:
        nx_attr = col_map.get(col, col)
        value = _get_attr(proto, nx_attr)

        if not value:
            return col

        if col == "REVISION" and not _REVISION_PATTERN.match(value):
            return col

    return None


def _build_row(component, depth, part_counter, col_map):
    """
    Build a MASTER_COLUMNS-ordered list of values for one component,
    filling STATUS with 'PASS' or 'FAIL: {col}'.
    Returns (row_values, first_failing_col_or_None).
    """
    proto = component.Prototype
    pn = _get_attr(proto, "PART_NUMBER") if proto else ""
    qty = part_counter.get(pn or component.Name, 1)

    first_fail = _audit_proto(proto, col_map) if proto is not None else "PART_NUMBER"
    status = "PASS" if first_fail is None else f"FAIL: {first_fail}"

    values = {}
    for col in MASTER_COLUMNS:
        if col == "LEVEL":
            values[col] = depth
        elif col == "PART_NUMBER":
            values[col] = pn
        elif col == "QUANTITY":
            values[col] = qty
        elif col == "STATUS":
            values[col] = status
        elif col == "NOTES":
            values[col] = ""
        elif proto is not None:
            nx_attr = col_map.get(col, col)
            values[col] = _get_attr(proto, nx_attr)
        else:
            values[col] = ""

    return [values[col] for col in MASTER_COLUMNS], first_fail


def main():
    session = get_session()
    part = get_work_part()

    if part is None:
        print("ERROR: No work part loaded.")
        return

    output_folder = prompt_folder("Select Audit Output Folder")
    if output_folder is None:
        print("Audit cancelled by user.")
        return

    col_map = _load_mapping()

    hla_pn = _get_attr(part, "PART_NUMBER") or os.path.splitext(os.path.basename(part.FullPath))[0]
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"AUDIT_{hla_pn}_{timestamp}.xlsx"
    output_path = os.path.join(output_folder, filename)

    root = part.ComponentAssembly.RootComponent
    component_rows, part_counter = _collect_components(root)

    writer = ExcelWriter(output_path)
    writer.write_master_header("Audit")

    pass_count = 0
    fail_count = 0
    fail_by_col = {col: 0 for col in _AUDITED_COLS}

    for component, depth in component_rows:
        row_values, first_fail = _build_row(component, depth, part_counter, col_map)
        if first_fail is None:
            writer.add_row_with_color("Audit", row_values, _GREEN)
            pass_count += 1
        else:
            writer.add_row_with_color("Audit", row_values, _AMBER)
            fail_count += 1
            fail_by_col[first_fail] = fail_by_col.get(first_fail, 0) + 1

    # Summary sheet
    writer.add_header_row("Summary", _SUMMARY_HEADERS)
    writer.add_row("Summary", ["Total components checked", len(component_rows)])
    writer.add_row("Summary", ["PASS", pass_count])
    writer.add_row("Summary", ["FAIL", fail_count])
    writer.add_row("Summary", ["", ""])
    writer.add_row("Summary", ["First-failing attribute", "FAIL count"])
    for col in _AUDITED_COLS:
        writer.add_row("Summary", [col, fail_by_col.get(col, 0)])

    writer.save()

    print(f"Audit complete.")
    print(f"  Components checked : {len(component_rows)}")
    print(f"  PASS               : {pass_count}")
    print(f"  FAIL               : {fail_count}")
    print(f"  Report written to  : {output_path}")


if __name__ == "__main__":
    main()
