"""
Journal 05 - Bulk Attribute Updater (Pull / Push)
NX 2312 rescue version: single-file journal, no repo-local imports.

Run via: NX > Tools > Journal > Play

PULL reads configured attributes from the work part and all unique assembly
prototypes, then writes PULL_<part>_<timestamp>.csv.

PUSH reads an Att-*.csv Teamcenter export, matches rows to NX parts by
DB_PART_NO, and writes values only when the NX attribute is currently empty.
It never overwrites non-empty NX attributes.
"""

import csv
import json
import os
import re
import sys
import traceback
from datetime import datetime

import NXOpen
import NXOpen.UF

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


# ---------------------------------------------------------------------------
# NX runtime helpers
# ---------------------------------------------------------------------------

def _get_session():
    return NXOpen.Session.GetSession()


def _listing_window(session):
    lw = session.ListingWindow
    lw.Open()
    return lw


def _log_lines(session, lines):
    lw = _listing_window(session)
    for line in lines:
        for text in str(line).splitlines() or [""]:
            print(text)
            lw.WriteFullline(text)


def _log_info(session, message):
    _log_lines(session, [message])


def _log_error(session, message):
    _log_lines(session, [f"ERROR: {message}"])


def _prompt_folder(title):
    ui = NXOpen.UI.GetUI()
    dialog = ui.CreateFolderSelectionDialog()
    dialog.SetTitle(title)
    try:
        response = dialog.Show()
        if response == NXOpen.SelectionDialog.DialogResponse.Ok:
            return dialog.Path
        return None
    finally:
        dialog.Destroy()


def _require_work_part(session):
    part = session.Parts.Work
    if part is None:
        _log_error(session, "No work part loaded.")
        return None
    return part


def _journal_path():
    try:
        return os.path.abspath(__file__)
    except Exception:
        return ""


def _runtime_root():
    path = _journal_path()
    if path:
        # Expected shape: <runtime_root>/journals/05_bulk_attribute_updater.py
        return os.path.dirname(os.path.dirname(path))
    return os.getcwd()


def _config_path():
    return os.path.join(_runtime_root(), "config", "attribute_mapping.json")


def _startup_diagnostics(session):
    root = _runtime_root()
    cfg = _config_path()
    lines = [
        "J05 startup diagnostics",
        f"  Journal path : {_journal_path() or '(unavailable)'}",
        f"  Runtime root : {root}",
        f"  Config path  : {cfg}",
        f"  Python       : {sys.version}",
        "  sys.path     :",
    ]
    for item in sys.path[:8]:
        lines.append(f"    {item}")
    _log_lines(session, lines)


def _run_journal(main_func):
    session = None
    try:
        session = _get_session()
        _startup_diagnostics(session)
        main_func(session)
    except Exception:
        lines = ["ERROR: Unhandled journal exception.", traceback.format_exc()]
        if session is not None:
            _log_lines(session, lines)
        else:
            for line in lines:
                print(line)


# ---------------------------------------------------------------------------
# File and CSV helpers
# ---------------------------------------------------------------------------

def _load_mapping():
    path = _config_path()
    if not os.path.exists(path):
        raise RuntimeError(
            "Missing attribute mapping config. Expected file at:\n"
            f"{path}\n"
            "Place the config folder beside the journals folder inside from_git."
        )
    with open(path, "r", encoding="utf-8") as fh:
        cfg = json.load(fh)
    tc_columns = cfg.get("tc_columns", {})
    if not isinstance(tc_columns, dict):
        raise RuntimeError("attribute_mapping.json must contain a 'tc_columns' object.")
    return tc_columns


def _write_csv(path, headers, rows):
    with open(path, "w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.writer(fh)
        writer.writerow(headers)
        for row in rows:
            writer.writerow(["" if value is None else value for value in row])
    return path


def _read_csv_rows(path):
    encodings = ("utf-8-sig", "utf-8", "cp1252")
    last_error = None
    for encoding in encodings:
        try:
            with open(path, "r", encoding=encoding, newline="") as fh:
                return [row for row in csv.reader(fh)]
        except UnicodeDecodeError as exc:
            last_error = exc
    raise RuntimeError(
        f"Unable to read CSV with supported encodings (utf-8-sig, utf-8, cp1252): {path}"
    ) from last_error


def _find_newest_csv(folder, prefix=None):
    csvs = [name for name in os.listdir(folder) if name.lower().endswith(".csv")]
    if prefix:
        preferred = [name for name in csvs if name.lower().startswith(prefix.lower())]
        csvs = preferred or csvs
    if not csvs:
        return None
    return max((os.path.join(folder, name) for name in csvs), key=os.path.getmtime)


# ---------------------------------------------------------------------------
# NX part and attribute helpers
# ---------------------------------------------------------------------------

def _get_string_attr(nxobj, attr_name, fallback=""):
    if nxobj is None or not attr_name:
        return fallback
    try:
        attr = nxobj.GetUserAttribute(attr_name, NXOpen.NXObject.AttributeType.String, -1)
        return attr.StringValue.strip()
    except Exception:
        return fallback


def _set_string_attr(nxobj, attr_name, value):
    nxobj.SetUserAttribute(attr_name, -1, value, NXOpen.Update.Option.Now)


def _safe_part_name(part, fallback="part"):
    try:
        full_path = part.FullPath
        if full_path:
            return os.path.splitext(os.path.basename(full_path))[0] or fallback
    except Exception:
        pass
    try:
        return part.Leaf or fallback
    except Exception:
        return fallback


def _get_root_component(part):
    try:
        return part.ComponentAssembly.RootComponent
    except Exception:
        return None


def _traverse_assembly(component):
    if component is None:
        return
    for child in component.GetChildren():
        yield child
        yield from _traverse_assembly(child)


def _unique_prototype_parts(part, include_work_part=True):
    unique = []
    seen = set()

    def add_part(candidate):
        if candidate is None:
            return
        key = getattr(candidate, "Tag", None) or getattr(candidate, "FullPath", None) or id(candidate)
        if key in seen:
            return
        seen.add(key)
        unique.append(candidate)

    if include_work_part:
        add_part(part)

    root = _get_root_component(part)
    for component in _traverse_assembly(root):
        add_part(component.Prototype)

    return unique


def _part_number(part):
    return _get_string_attr(part, "DB_PART_NO") or _get_string_attr(part, "PART_NUMBER")


# ---------------------------------------------------------------------------
# J05 behavior
# ---------------------------------------------------------------------------

def _decode_alias(alias):
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


def _pull_rows(part):
    rows = []
    for proto in _unique_prototype_parts(part, include_work_part=True):
        row = {}
        for col in _PULL_ALL_COLS:
            row[col] = _get_string_attr(proto, col)
        if not row["DB_PART_NO"]:
            row["DB_PART_NO"] = _get_string_attr(proto, "PART_NUMBER")
        rows.append([row.get(col, "") for col in _PULL_ALL_COLS])
    return rows


def _run_pull(session, part):
    output_folder = _prompt_folder("Select Output Folder for PULL Report")
    if output_folder is None:
        _log_info(session, "PULL cancelled by user.")
        return

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = os.path.join(output_folder, f"PULL_{_safe_part_name(part)}_{timestamp}.csv")
    rows = _pull_rows(part)
    _write_csv(output_path, _PULL_ALL_COLS, rows)

    _log_info(
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
    for proto in _unique_prototype_parts(part, include_work_part=True):
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
    input_folder = _prompt_folder("Select Folder Containing TC Attribute CSV (Att-*.csv)")
    if input_folder is None:
        _log_info(session, "PUSH cancelled by user.")
        return

    tc_path = _find_newest_csv(input_folder, prefix="Att-")
    if tc_path is None:
        _log_error(session, f"No TC CSV found in: {input_folder}")
        return

    try:
        all_rows = _read_csv_rows(tc_path)
        tc_columns = _load_mapping()
    except RuntimeError as exc:
        _log_error(session, exc)
        return

    if len(all_rows) < 3:
        _log_error(session, "TC CSV has fewer than 3 rows (expected 2 header rows + data).")
        return

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
        _log_error(session, "DB_PART_NO column not found in TC CSV header row 2.")
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
            _log_info(session, f"WARN: alias '{alias}' (internal: '{internal_name}') not found in TC CSV headers.")

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
                nx_existing = _get_string_attr(proto, internal_name)
                if nx_existing:
                    action = "KEPT"
                    kept_count += 1
                elif tc_value:
                    try:
                        _set_string_attr(proto, internal_name, tc_value)
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
    _write_csv(report_path, ["PART_NUMBER", "ATTR_NAME", "NX_EXISTING", "TC_VALUE", "ACTION"], report_rows)

    _log_info(
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
    part = _require_work_part(session)
    if part is None:
        return

    mode = _prompt_mode(session)
    if mode == "pull":
        _run_pull(session, part)
    else:
        _run_push(session, part)


if __name__ == "__main__":
    _run_journal(main)
