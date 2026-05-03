"""
Journal 05 - Bulk Attribute Updater (Pull / Push)
Run via: NX > Tools > Journal > Play

PULL reads configured attributes from the work part and all unique assembly
prototypes, then writes PULL_<part>_<timestamp>.csv.

PUSH reads an Att-*.csv Teamcenter export, matches rows to NX parts by
DB_PART_NO, and writes values only when the NX attribute is currently empty.
It never overwrites non-empty NX attributes.
"""

import os
import re
import sys
from datetime import datetime

import NXOpen
import NXOpen.UF

_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _REPO_ROOT not in sys.path:
    sys.path.insert(0, _REPO_ROOT)

from utils.config_loader import load_json_config  # noqa: E402
from utils.csv_reports import find_newest_csv, read_csv_rows, write_csv  # noqa: E402
from utils.nx_helpers import (  # noqa: E402
    get_string_attr,
    log_error,
    log_info,
    prompt_folder,
    require_work_part,
    run_journal,
    safe_part_name,
    set_string_attr,
    unique_prototype_parts,
)

_MAPPING_FILE = os.path.join("config", "attribute_mapping.json")
_IDENTITY_INTERNALS = {"DB_PART_NO", "DB_PART_REV"}

_PULL_IDENTITY_COLS = ["DB_PART_NO", "DB_PART_NAME", "DB_PART_REV"]
_PULL_ATTR_COLS = [
    "UOM",
    "Mfr. Name",
    "Mfr. Part Number",
    "PREFERRED",
    "NX_MassPropRollupMass",
    "MATERIAL",
    "Part Classification",
    "Commodity Code",
    "Commodity Type",
    "Country of Origin",
    "Export Control Number",
    "Traceability",
    "Shelf Life Limited",
    "Temperature Sensitive",
    "Serviceable item flag",
    "HANDEDNESS",
    "MANUFACTURINGCODE",
    "PARTCATEGORY",
    "PROGRAMIDENTIFIER",
    "PROJECT_IDS",
    "WAEItemID",
    "WAEItemItemID",
    "COMPONENT_CLASS",
    "LIFED",
    "SERIAL_NUMBERED_PART",
    "Stocking Type",
]
_PULL_ALL_COLS = _PULL_IDENTITY_COLS + _PULL_ATTR_COLS


def _load_mapping():
    cfg = load_json_config(_REPO_ROOT, _MAPPING_FILE)
    return cfg.get("tc_columns", {})


def _decode_alias(alias):
    """'ID (DB_PART_NO)' -> 'DB_PART_NO'; 'HANDEDNESS' -> 'HANDEDNESS'."""
    match = re.search(r"\(([^)]+)\)", str(alias))
    return match.group(1) if match else str(alias).strip()


def _prompt_mode(session):
    ui = NXOpen.UI.GetUI()
    response = ui.NXMessageBox.Show(
        "J05 - Bulk Attribute Updater",
        NXOpen.NXMessageBox.DialogType.Question,
        "Select mode:\n\nYes = PULL (read NX attributes -> CSV)\nNo = PUSH (write TC CSV -> NX)",
    )
    try:
        is_yes = int(response) == 6
    except (TypeError, ValueError):
        is_yes = str(response).strip().lower() in ("yes", "6", "1")
    return "pull" if is_yes else "push"


def _part_number(part):
    return get_string_attr(part, "DB_PART_NO") or get_string_attr(part, "PART_NUMBER")


def _pull_rows(part):
    rows = []
    for proto in unique_prototype_parts(part, include_work_part=True):
        row = {}
        for col in _PULL_ALL_COLS:
            row[col] = get_string_attr(proto, col)
        if not row["DB_PART_NO"]:
            row["DB_PART_NO"] = get_string_attr(proto, "PART_NUMBER")
        rows.append([row.get(col, "") for col in _PULL_ALL_COLS])
    return rows


def _run_pull(session, part):
    output_folder = prompt_folder("Select Output Folder for PULL Report")
    if output_folder is None:
        log_info(session, "PULL cancelled by user.")
        return

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = os.path.join(output_folder, f"PULL_{safe_part_name(part)}_{timestamp}.csv")
    rows = _pull_rows(part)
    write_csv(output_path, _PULL_ALL_COLS, rows)

    log_info(
        session,
        "\n".join(
            [
                "PULL complete.",
                f"  Parts found  : {len(rows)}",
                f"  Report saved : {output_path}",
            ]
        ),
    )


def _build_part_map(part):
    part_map = {}
    for proto in unique_prototype_parts(part, include_work_part=True):
        pn = _part_number(proto)
        if pn:
            part_map.setdefault(pn, []).append(proto)
    return part_map


def _build_alias_index(alias_row):
    alias_to_idx = {}
    for idx, alias in enumerate(alias_row):
        clean_alias = str(alias).strip() if alias is not None else ""
        alias_to_idx[clean_alias] = idx
        decoded = _decode_alias(clean_alias)
        if decoded != clean_alias:
            alias_to_idx[decoded] = idx
    return alias_to_idx


def _get_cell(row, idx):
    if idx is None or idx >= len(row):
        return ""
    value = row[idx]
    return str(value).strip() if value is not None else ""


def _run_push(session, part):
    input_folder = prompt_folder("Select Folder Containing TC Attribute CSV (Att-*.csv)")
    if input_folder is None:
        log_info(session, "PUSH cancelled by user.")
        return

    tc_path = find_newest_csv(input_folder, prefix="Att-")
    if tc_path is None:
        log_error(session, f"No TC CSV found in: {input_folder}")
        return

    try:
        all_rows = read_csv_rows(tc_path)
    except RuntimeError as exc:
        log_error(session, exc)
        return
    if len(all_rows) < 3:
        log_error(session, "TC CSV has fewer than 3 rows (expected 2 header rows + data).")
        return

    tc_columns = _load_mapping()
    writable_map = {
        alias: internal
        for alias, internal in tc_columns.items()
        if internal not in _IDENTITY_INTERNALS
    }

    alias_to_idx = _build_alias_index(all_rows[1])
    pn_col_idx = alias_to_idx.get("DB_PART_NO")
    if pn_col_idx is None:
        pn_col_idx = alias_to_idx.get("ID (DB_PART_NO)")
    if pn_col_idx is None:
        log_error(session, "DB_PART_NO column not found in TC CSV header row 2.")
        return

    part_map = _build_part_map(part)
    report_rows = []
    updated_count = 0
    kept_count = 0
    empty_count = 0
    no_match_count = 0
    error_count = 0

    for alias, internal_name in writable_map.items():
        if alias not in alias_to_idx and internal_name not in alias_to_idx:
            log_info(session, f"WARN: alias '{alias}' (internal: '{internal_name}') not found in TC CSV headers.")

    for data_row in all_rows[2:]:
        pn = _get_cell(data_row, pn_col_idx)
        if not pn:
            continue

        protos = part_map.get(pn)
        if not protos:
            report_rows.append([pn, "-", "", "", "NO_MATCH"])
            no_match_count += 1
            continue

        for alias, internal_name in writable_map.items():
            tc_idx = alias_to_idx.get(alias, alias_to_idx.get(internal_name))
            if tc_idx is None:
                continue

            tc_value = _get_cell(data_row, tc_idx)
            for proto in protos:
                nx_existing = get_string_attr(proto, internal_name)
                if nx_existing:
                    action = "KEPT"
                    kept_count += 1
                elif tc_value:
                    try:
                        set_string_attr(proto, internal_name, tc_value)
                        action = "UPDATED"
                        updated_count += 1
                    except Exception as exc:
                        action = f"ERROR: {exc}"
                        error_count += 1
                else:
                    action = "BOTH_EMPTY"
                    empty_count += 1

                report_rows.append([pn, internal_name, nx_existing, tc_value, action])

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_path = os.path.join(input_folder, f"PUSH_REPORT_{timestamp}.csv")
    write_csv(report_path, ["PART_NUMBER", "ATTR_NAME", "NX_EXISTING", "TC_VALUE", "ACTION"], report_rows)

    log_info(
        session,
        "\n".join(
            [
                "PUSH complete.",
                f"  TC CSV     : {tc_path}",
                f"  UPDATED    : {updated_count}",
                f"  KEPT       : {kept_count}",
                f"  BOTH_EMPTY : {empty_count}",
                f"  NO_MATCH   : {no_match_count}",
                f"  ERRORS     : {error_count}",
                f"  Report     : {report_path}",
            ]
        ),
    )


def main(session):
    part = require_work_part(session)
    if part is None:
        return

    mode = _prompt_mode(session)
    if mode == "pull":
        _run_pull(session, part)
    else:
        _run_push(session, part)


if __name__ == "__main__":
    run_journal(main)
