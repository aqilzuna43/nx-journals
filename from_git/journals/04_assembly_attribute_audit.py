"""
Journal 04 - Assembly Attribute Audit
Checks required NX part attributes across the assembly and reports issues to CSV.
Run via: NX > Tools > Journal > Play
"""

import os
import re
import sys
from collections import Counter
from datetime import datetime

import NXOpen
import NXOpen.UF

_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _REPO_ROOT not in sys.path:
    sys.path.insert(0, _REPO_ROOT)

from utils.config_loader import load_json_config  # noqa: E402
from utils.csv_reports import write_csv  # noqa: E402
from utils.nx_helpers import (  # noqa: E402
    get_output_folder,
    get_string_attr,
    iter_occurrences,
    log_info,
    require_work_part,
    run_journal,
    safe_part_name,
    get_root_component,
)
from utils.template_generator import MASTER_COLUMNS  # noqa: E402

_MAPPING_FILE = os.path.join("config", "attribute_mapping.json")
_REVISION_PATTERN = re.compile(r"^([A-Z]|\d+)$")
_AUDITED_COLS = ["PART_NUMBER", "DESCRIPTION", "MATERIAL", "FINISH", "REVISION", "DRAWING_NUMBER"]


def _load_mapping():
    cfg = load_json_config(_REPO_ROOT, _MAPPING_FILE)
    return cfg.get("columns", {})


def _part_number(part, col_map):
    pn_attr = col_map.get("PART_NUMBER", "DB_PART_NO")
    return get_string_attr(part, pn_attr) or get_string_attr(part, "PART_NUMBER")


def _collect_rows(work_part, col_map):
    rows = []
    part_counter = Counter()

    occurrences = list(iter_occurrences(work_part))
    for component in occurrences:
        proto = component.Prototype
        pn = _part_number(proto, col_map) if proto is not None else ""
        part_counter[pn or component.Name] += 1

    def walk(component, depth):
        for child in component.GetChildren():
            rows.append((child, depth))
            walk(child, depth + 1)

    root = get_root_component(work_part)
    if root is not None:
        walk(root, 1)

    return rows, part_counter


def _audit_proto(proto, col_map):
    for col in _AUDITED_COLS:
        value = ""
        if proto is not None:
            value = get_string_attr(proto, col_map.get(col, col))
        if not value:
            return col
        if col == "REVISION" and not _REVISION_PATTERN.match(value):
            return col
    return None


def _build_row(component, depth, part_counter, col_map):
    proto = component.Prototype
    pn = _part_number(proto, col_map) if proto is not None else ""
    qty = part_counter.get(pn or component.Name, 1)
    first_fail = _audit_proto(proto, col_map)
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
            values[col] = get_string_attr(proto, col_map.get(col, col))
        else:
            values[col] = ""

    return [values[col] for col in MASTER_COLUMNS], first_fail


def main(session):
    part = require_work_part(session)
    if part is None:
        return

    output_folder = get_output_folder()
    log_info(session, f"Audit output folder: {output_folder}")

    col_map = _load_mapping()
    hla_pn = _part_number(part, col_map) or safe_part_name(part)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    audit_path = os.path.join(output_folder, f"AUDIT_{hla_pn}_{timestamp}.csv")
    summary_path = os.path.join(output_folder, f"AUDIT_SUMMARY_{hla_pn}_{timestamp}.csv")

    component_rows, part_counter = _collect_rows(part, col_map)
    audit_rows = []
    pass_count = 0
    fail_count = 0
    fail_by_col = {col: 0 for col in _AUDITED_COLS}

    for component, depth in component_rows:
        row_values, first_fail = _build_row(component, depth, part_counter, col_map)
        audit_rows.append(row_values)
        if first_fail is None:
            pass_count += 1
        else:
            fail_count += 1
            fail_by_col[first_fail] = fail_by_col.get(first_fail, 0) + 1

    summary_rows = [
        ["Total components checked", len(component_rows)],
        ["PASS", pass_count],
        ["FAIL", fail_count],
        ["", ""],
        ["First-failing attribute", "FAIL count"],
    ]
    summary_rows.extend([[col, fail_by_col.get(col, 0)] for col in _AUDITED_COLS])

    write_csv(audit_path, MASTER_COLUMNS, audit_rows)
    write_csv(summary_path, ["Metric", "Value"], summary_rows)

    log_info(
        session,
        "\n".join(
            [
                "Audit complete.",
                f"  Components checked : {len(component_rows)}",
                f"  PASS               : {pass_count}",
                f"  FAIL               : {fail_count}",
                f"  Report written to  : {audit_path}",
                f"  Summary written to : {summary_path}",
            ]
        ),
    )


if __name__ == "__main__":
    run_journal(main)
