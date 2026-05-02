"""
discover_attributes.py — Run once in NX to reveal internal attribute names.

NX Open shows display aliases in the UI (e.g. "@ Part Number") but the
internal title used by GetUserAttribute() may differ (e.g. "DB_PART_NO").
This journal dumps every attribute on the current work part so you can
build the correct attribute_mapping.json.

How to use:
  1. Open a part in NX that has all your standard attributes populated.
  2. Tools > Journal > Play > select this file.
  3. Open the output .txt file — match titles against what you see in the
NX attribute editor to build your attribute_mapping.json.
"""

import os
import sys
from datetime import datetime

import NXOpen
import NXOpen.UF

_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _REPO_ROOT not in sys.path:
    sys.path.insert(0, _REPO_ROOT)

from utils.nx_helpers import get_session, get_work_part, prompt_folder  # noqa: E402

# Known TC-propagated attribute names to probe explicitly in addition to the
# bulk GetUserAttributes() call, in case they are not returned by the bulk API.
_TC_PROBE_NAMES = [
    "DB_PART_NO",
    "DB_TITLE",
    "DB_PART_REV",
    "SURFACE_FINISH",
    "DRAWING_NUMBER",
    "MATERIAL",
    "PART_NUMBER",
    "DESCRIPTION",
    "FINISH",
    "REVISION",
]


def _attr_value_str(attr_info):
    """Return a readable string for an AttributeInformation value."""
    t = attr_info.Type
    try:
        if t == NXOpen.NXObject.AttributeType.String:
            return attr_info.StringValue
        if t == NXOpen.NXObject.AttributeType.Integer:
            return str(attr_info.IntegerValue)
        if t == NXOpen.NXObject.AttributeType.Real:
            return str(attr_info.RealValue)
        if t == NXOpen.NXObject.AttributeType.Boolean:
            return str(attr_info.BooleanValue)
        if t == NXOpen.NXObject.AttributeType.Time:
            return str(attr_info.TimeValue)
        return "(unsupported type)"
    except Exception:
        return "(error reading value)"


def _type_name(attr_type):
    """Return a short string label for an attribute type enum value."""
    mapping = {
        NXOpen.NXObject.AttributeType.String: "String",
        NXOpen.NXObject.AttributeType.Integer: "Integer",
        NXOpen.NXObject.AttributeType.Real: "Real",
        NXOpen.NXObject.AttributeType.Boolean: "Boolean",
        NXOpen.NXObject.AttributeType.Time: "Time",
        NXOpen.NXObject.AttributeType.Null: "Null",
        NXOpen.NXObject.AttributeType.Any: "Any",
        NXOpen.NXObject.AttributeType.Reference: "Reference",
    }
    return mapping.get(attr_type, "Unknown")


def main():
    session = get_session()
    part = get_work_part()

    if part is None:
        print("ERROR: No work part loaded. Open a representative part and try again.")
        return

    output_folder = prompt_folder("Select Output Folder for Attribute Report")
    if output_folder is None:
        print("Attribute discovery cancelled by user.")
        return

    part_name = os.path.splitext(os.path.basename(part.FullPath))[0]
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_name = f"ATTR_DISCOVERY_{part_name}_{timestamp}.txt"
    report_path = os.path.join(output_folder, report_name)

    # --- collect bulk attributes ---
    rows = {}  # title → (type_name, value)
    try:
        for attr_info in part.GetUserAttributes():
            title = attr_info.Title
            rows[title] = (_type_name(attr_info.Type), _attr_value_str(attr_info))
    except Exception as exc:
        rows["__GetUserAttributes_error__"] = ("Error", str(exc))

    # --- probe known TC attribute names explicitly ---
    for probe_name in _TC_PROBE_NAMES:
        if probe_name in rows:
            continue
        for attr_type in (
            NXOpen.NXObject.AttributeType.String,
            NXOpen.NXObject.AttributeType.Integer,
            NXOpen.NXObject.AttributeType.Real,
        ):
            try:
                attr_info = part.GetUserAttribute(probe_name, attr_type, -1)
                rows[probe_name] = (_type_name(attr_info.Type), _attr_value_str(attr_info))
                break
            except Exception:
                pass

    # --- build sorted report lines ---
    header_line = f"{'Title':<40} | {'Type':<12} | Value"
    separator = "-" * 80
    sorted_rows = sorted(rows.items())
    data_lines = [
        f"{title:<40} | {type_str:<12} | {value}"
        for title, (type_str, value) in sorted_rows
    ]

    report_lines = [
        f"NX Attribute Discovery Report",
        f"Part file : {part.FullPath}",
        f"Date      : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"Attributes: {len(rows)}",
        separator,
        header_line,
        separator,
        *data_lines,
        separator,
    ]

    # --- write txt report ---
    with open(report_path, "w", encoding="utf-8") as f:
        f.write("\n".join(report_lines))

    # --- echo to NX listing window ---
    lw = session.ListingWindow
    lw.Open()
    for line in report_lines:
        lw.WriteFullline(line)

    print(f"Report saved to: {report_path}")


if __name__ == "__main__":
    main()
