"""
Journal 05 — Bulk Attribute Updater (Pull / Push)
Run via: NX > Tools > Journal > Play

Two modes:

  Everything runs inside NX and operates directly on NX part files via
  GetUserAttribute / SetUserAttribute.  No connection to Teamcenter is made
  at runtime — the TC Excel is used only as a plain input file to supply values.

  PULL — Reads existing attribute values from every unique NX prototype in the
         open assembly and writes them to PULL_<assembly>_<timestamp>.xlsx.
         Run this first to verify that the internal NX attribute names in
         attribute_mapping.yaml match what is actually stored on the parts.

  PUSH — Reads the TC-exported attribute Excel (e.g. Att-*.xlsx) as a plain
         spreadsheet, matches rows to NX prototypes by DB_PART_NO, and calls
         SetUserAttribute on the NX part — but ONLY for attributes that are
         currently empty (never overwrites existing data).
         Writes PUSH_REPORT_<timestamp>.xlsx alongside the TC Excel showing
         every decision (UPDATED / KEPT / BOTH_EMPTY / NO_MATCH).

Workflow:
  1. Run PULL on the open assembly.  Review PULL_*.xlsx to confirm attribute names.
  2. Update config/attribute_mapping.yaml if any names differ.
  3. Place the TC attribute Excel in a folder.
  4. Run PUSH, select that folder.  Review PUSH_REPORT_*.xlsx and verify in NX.
"""

import os
import re
import sys
from datetime import datetime

import NXOpen
import NXOpen.UF
import openpyxl
import yaml

_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _REPO_ROOT not in sys.path:
    sys.path.insert(0, _REPO_ROOT)

from utils.nx_helpers import get_session, get_work_part, prompt_folder, traverse_assembly  # noqa: E402
from utils.excel_writer import ExcelWriter  # noqa: E402

_MAPPING_FILE = os.path.join(_REPO_ROOT, "config", "attribute_mapping.yaml")

# Pull output columns ordered to match FZ-PowerSystem_v1 - MASTER.csv.
# Prefixed identity columns (DB_PART_NO / DB_PART_REV) are always written first.
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
# Helpers
# ---------------------------------------------------------------------------

def _load_mapping():
    with open(_MAPPING_FILE, encoding="utf-8") as f:
        cfg = yaml.safe_load(f)
    tc_columns = cfg.get("tc_columns", {})   # alias → internal_name
    skip = set(cfg.get("skip_columns", []))
    return tc_columns, skip


def _get_attr(nxobj, attr_name):
    """Return string value of a user attribute, or '' if missing/error.

    Tries String type only — all TC-managed attributes are strings.
    Integer/Real probe is omitted to avoid false '0' returns on missing attrs.
    """
    try:
        attr = nxobj.GetUserAttribute(attr_name, NXOpen.NXObject.AttributeType.String, -1)
        return attr.StringValue.strip()
    except Exception:
        return ""


def _set_attr(nxobj, attr_name, value):
    nxobj.SetUserAttribute(attr_name, -1, value, NXOpen.Update.Option.Now)


def _decode_alias(alias):
    """'ID (DB_PART_NO)' → 'DB_PART_NO'; 'HANDEDNESS' → 'HANDEDNESS'."""
    m = re.search(r'\(([^)]+)\)', str(alias))
    return m.group(1) if m else str(alias).strip()


def _build_part_map(root_component):
    """Return {db_part_no: [prototype, ...]} for every unique prototype."""
    part_map = {}
    seen = set()
    for component in traverse_assembly(root_component):
        proto = component.Prototype
        if proto is None:
            continue
        tag = proto.Tag
        if tag in seen:
            continue
        seen.add(tag)
        pn = _get_attr(proto, "DB_PART_NO")
        if not pn:
            pn = _get_attr(proto, "PART_NUMBER")
        if pn:
            part_map.setdefault(pn, []).append(proto)
    return part_map


def _prompt_mode(session):
    """Prompt user to choose PULL or PUSH via NX dialog.

    NXMessageBox.Show returns an integer whose exact value varies by NX version.
    For DialogType.Question the Yes button is IDYES (6 on Windows).
    We also guard against versions that return a string or enum object.
    """
    ui = NXOpen.UI.GetUI()
    response = ui.NXMessageBox.Show(
        "J05 — Bulk Attribute Updater",
        NXOpen.NXMessageBox.DialogType.Question,
        "Select mode:\n\nYes = PULL (read NX attributes → Excel)\nNo = PUSH (write TC Excel → NX)",
    )
    # Normalise response to a comparable string
    try:
        is_yes = int(response) == 6  # IDYES on Windows / NX
    except (TypeError, ValueError):
        is_yes = str(response).strip().lower() in ("yes", "6", "1")
    return "pull" if is_yes else "push"


def _find_tc_excel(folder):
    """Return newest xlsx in folder that starts with 'Att-' or 'att-', or any xlsx."""
    xlsxs = [f for f in os.listdir(folder) if f.lower().endswith(".xlsx")]
    att = [f for f in xlsxs if f.lower().startswith("att-")]
    candidates = att if att else xlsxs
    if not candidates:
        return None
    return max(
        (os.path.join(folder, f) for f in candidates),
        key=os.path.getmtime,
    )


# ---------------------------------------------------------------------------
# PULL mode
# ---------------------------------------------------------------------------

def _run_pull(session, part):
    output_folder = prompt_folder("Select Output Folder for PULL Report")
    if output_folder is None:
        print("PULL cancelled.")
        return

    root = part.ComponentAssembly.RootComponent
    seen = set()
    rows = []

    for component in traverse_assembly(root):
        proto = component.Prototype
        if proto is None:
            continue
        tag = proto.Tag
        if tag in seen:
            continue
        seen.add(tag)

        # Read every column we care about
        row = {}
        for col in _PULL_ALL_COLS:
            row[col] = _get_attr(proto, col)

        # Fallback: also try PART_NUMBER for DB_PART_NO
        if not row["DB_PART_NO"]:
            row["DB_PART_NO"] = _get_attr(proto, "PART_NUMBER")

        rows.append(row)

    # Write Excel
    part_name = os.path.splitext(os.path.basename(part.FullPath))[0]
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = os.path.join(output_folder, f"PULL_{part_name}_{ts}.xlsx")

    writer = ExcelWriter(out_path)
    writer.add_header_row("PULL", _PULL_ALL_COLS)
    for row in rows:
        writer.add_row("PULL", [row.get(c, "") for c in _PULL_ALL_COLS])
    writer.save()

    summary = (
        f"PULL complete.\n"
        f"  Parts found  : {len(rows)}\n"
        f"  Report saved : {out_path}"
    )
    print(summary)
    lw = session.ListingWindow
    lw.Open()
    lw.WriteFullline(summary)


# ---------------------------------------------------------------------------
# PUSH mode
# ---------------------------------------------------------------------------

def _run_push(session, part):
    input_folder = prompt_folder("Select Folder Containing TC Attribute Excel (Att-*.xlsx)")
    if input_folder is None:
        print("PUSH cancelled.")
        return

    tc_path = _find_tc_excel(input_folder)
    if tc_path is None:
        print(f"ERROR: No xlsx found in: {input_folder}")
        return

    print(f"Reading TC Excel: {tc_path}")

    tc_columns, _ = _load_mapping()
    # Build {internal_name: alias} reverse for lookups, and writable set
    # tc_columns = {alias: internal_name}; writable = those not identity
    identity_internals = {"DB_PART_NO", "DB_PART_REV"}
    writable_map = {
        alias: internal
        for alias, internal in tc_columns.items()
        if internal not in identity_internals
    }

    wb = openpyxl.load_workbook(tc_path, read_only=True, data_only=True)
    ws = wb.active
    all_rows = list(ws.iter_rows(values_only=True))
    wb.close()

    if len(all_rows) < 3:
        print("ERROR: TC Excel has fewer than 3 rows (expected 2 header rows + data).")
        return

    # Row 1 = group headers (index 0), row 2 = display aliases (index 1)
    alias_row = [str(c).strip() if c is not None else "" for c in all_rows[1]]
    data_rows = all_rows[2:]

    # Map alias → column index; decode parenthetical names
    alias_to_idx = {}
    for idx, alias in enumerate(alias_row):
        alias_to_idx[alias] = idx
        decoded = _decode_alias(alias)
        if decoded != alias:
            alias_to_idx[decoded] = idx

    # Locate DB_PART_NO column
    pn_col_idx = alias_to_idx.get("DB_PART_NO")
    if pn_col_idx is None:
        # Try the raw alias form
        pn_col_idx = alias_to_idx.get("ID (DB_PART_NO)")
    if pn_col_idx is None:
        print("ERROR: DB_PART_NO column not found in TC Excel header row 2.")
        return

    # Build NX part map
    root = part.ComponentAssembly.RootComponent
    part_map = _build_part_map(root)

    # Warn about writable attributes whose alias wasn't found in the TC Excel header
    for alias, internal_name in writable_map.items():
        if alias not in alias_to_idx and internal_name not in alias_to_idx:
            print(f"  WARN: alias '{alias}' (internal: '{internal_name}') not found in TC Excel headers — column skipped.")

    # Collect report rows
    report_rows = []   # (part_number, attr_name, nx_existing, tc_value, action)
    updated_count = 0
    kept_count = 0
    empty_count = 0
    no_match_count = 0
    error_count = 0

    for data_row in data_rows:
        pn_cell = data_row[pn_col_idx]
        pn = str(pn_cell).strip() if pn_cell is not None else ""
        if not pn:
            continue

        protos = part_map.get(pn)
        if not protos:
            report_rows.append((pn, "—", "", "", "NO_MATCH"))
            no_match_count += 1
            continue

        for alias, internal_name in writable_map.items():
            tc_idx = alias_to_idx.get(alias, alias_to_idx.get(internal_name))
            if tc_idx is None:
                continue
            tc_cell = data_row[tc_idx]
            tc_value = str(tc_cell).strip() if tc_cell is not None else ""

            for proto in protos:
                nx_existing = _get_attr(proto, internal_name)

                if nx_existing:
                    action = "KEPT"
                    kept_count += 1
                elif tc_value:
                    try:
                        _set_attr(proto, internal_name, tc_value)
                        action = "UPDATED"
                        updated_count += 1
                    except Exception as exc:
                        action = f"ERROR: {exc}"
                        error_count += 1
                else:
                    action = "BOTH_EMPTY"
                    empty_count += 1

                report_rows.append((pn, internal_name, nx_existing, tc_value, action))

    # Write PUSH_REPORT Excel
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_path = os.path.join(input_folder, f"PUSH_REPORT_{ts}.xlsx")
    writer = ExcelWriter(report_path)
    report_headers = ["PART_NUMBER", "ATTR_NAME", "NX_EXISTING", "TC_VALUE", "ACTION"]
    writer.add_header_row("REPORT", report_headers)
    for row in report_rows:
        highlight = row[4] == "UPDATED"
        writer.add_row("REPORT", list(row), highlight_amber=highlight)
    writer.save()

    summary = (
        f"PUSH complete.\n"
        f"  UPDATED    : {updated_count}\n"
        f"  KEPT       : {kept_count}\n"
        f"  BOTH_EMPTY : {empty_count}\n"
        f"  NO_MATCH   : {no_match_count}\n"
        f"  ERRORS     : {error_count}\n"
        f"  Report     : {report_path}"
    )
    print(summary)
    lw = session.ListingWindow
    lw.Open()
    lw.WriteFullline(summary)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    session = get_session()
    part = get_work_part()

    if part is None:
        print("ERROR: No work part loaded.")
        return

    mode = _prompt_mode(session)

    if mode == "pull":
        _run_pull(session, part)
    else:
        _run_push(session, part)


if __name__ == "__main__":
    main()
