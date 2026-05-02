"""
Journal 05 — Bulk Attribute Updater
Reads a J02 or J04 BOM xlsx and writes attribute values back into the open NX assembly.
Run via: NX > Tools > Journal > Play

Workflow:
  1. Run J02 or J04 to generate a BOM xlsx.
  2. Edit STATUS/attribute columns in the xlsx as needed.
  3. Run this journal, select the folder containing the xlsx.
     The most recently modified BOM_*.xlsx or AUDIT_*.xlsx in that folder is used.
  4. Attributes are pushed back to each prototype matched by PART_NUMBER.
"""

import os
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
from utils.template_generator import MASTER_COLUMNS  # noqa: E402

_MAPPING_FILE = os.path.join(_REPO_ROOT, "config", "attribute_mapping.yaml")


def _load_mapping():
    with open(_MAPPING_FILE, encoding="utf-8") as f:
        cfg = yaml.safe_load(f)
    columns = cfg.get("columns", {})
    skip = set(cfg.get("skip_columns", []))
    return columns, skip


def _find_bom_file(folder):
    """Return the newest BOM_*.xlsx or AUDIT_*.xlsx in folder, or None."""
    candidates = [
        os.path.join(folder, f)
        for f in os.listdir(folder)
        if f.lower().endswith(".xlsx") and (f.startswith("BOM_") or f.startswith("AUDIT_"))
    ]
    if not candidates:
        return None
    return max(candidates, key=os.path.getmtime)


def _get_attr(nxobj, attr_name):
    try:
        attr = nxobj.GetUserAttribute(attr_name, NXOpen.NXObject.AttributeType.String, -1)
        return attr.StringValue.strip()
    except Exception:
        return ""


def _build_part_map(root_component):
    """Return {part_number: [prototype, ...]} for every unique prototype in the assembly."""
    part_map = {}
    seen_tags = set()
    for component in traverse_assembly(root_component):
        proto = component.Prototype
        if proto is None:
            continue
        tag = proto.Tag
        if tag in seen_tags:
            continue
        seen_tags.add(tag)
        pn = _get_attr(proto, "PART_NUMBER")
        if pn:
            part_map.setdefault(pn, []).append(proto)
    return part_map


def _set_attr(nxobj, attr_name, value):
    nxobj.SetUserAttribute(attr_name, -1, value, NXOpen.Update.Option.Now)


def main():
    session = get_session()
    part = get_work_part()

    if part is None:
        print("ERROR: No work part loaded.")
        return

    input_folder = prompt_folder("Select Folder Containing BOM / Audit xlsx")
    if input_folder is None:
        print("Update cancelled by user.")
        return

    bom_path = _find_bom_file(input_folder)
    if bom_path is None:
        print(f"ERROR: No BOM_*.xlsx or AUDIT_*.xlsx found in: {input_folder}")
        return

    print(f"Reading: {bom_path}")

    col_map, skip_cols = _load_mapping()

    wb = openpyxl.load_workbook(bom_path, read_only=True, data_only=True)
    ws = wb.active

    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        print("ERROR: xlsx is empty.")
        wb.close()
        return

    # Build column-name → index map from header row
    header = [str(cell).strip() if cell is not None else "" for cell in rows[0]]
    col_index = {name: idx for idx, name in enumerate(header)}

    if "PART_NUMBER" not in col_index:
        print("ERROR: PART_NUMBER column not found in header row.")
        wb.close()
        return

    # Columns to write back: those in col_map that are not in skip_cols
    write_cols = [
        col for col in MASTER_COLUMNS
        if col not in skip_cols and col in col_map and col in col_index
    ]

    root = part.ComponentAssembly.RootComponent
    part_map = _build_part_map(root)

    updated_parts = 0
    skipped_rows = 0
    attr_write_count = 0

    for data_row in rows[1:]:
        pn_idx = col_index["PART_NUMBER"]
        pn = str(data_row[pn_idx]).strip() if data_row[pn_idx] is not None else ""

        if not pn:
            skipped_rows += 1
            continue

        protos = part_map.get(pn)
        if not protos:
            print(f"  SKIP: PART_NUMBER '{pn}' not found in assembly.")
            skipped_rows += 1
            continue

        for proto in protos:
            for col in write_cols:
                cell_idx = col_index[col]
                cell_val = data_row[cell_idx]
                if cell_val is None:
                    continue
                value = str(cell_val).strip()
                nx_attr = col_map[col]
                try:
                    _set_attr(proto, nx_attr, value)
                    attr_write_count += 1
                except Exception as exc:
                    print(f"  WARN: could not set {nx_attr} on '{pn}': {exc}")

        updated_parts += 1

    wb.close()

    print(f"Bulk update complete.")
    print(f"  Parts updated      : {updated_parts}")
    print(f"  Attributes written : {attr_write_count}")
    print(f"  Rows skipped       : {skipped_rows}")


if __name__ == "__main__":
    main()
