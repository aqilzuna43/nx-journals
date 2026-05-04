"""
discover_attributes.py - Run once in NX to reveal internal attribute names.

NX Open shows display aliases in the UI, but the internal title used by
GetUserAttribute() may differ. This journal dumps every attribute on the
current work part so you can build config/attribute_mapping.json.
"""

import os
import sys
from datetime import datetime

import NXOpen
import NXOpen.UF

_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _REPO_ROOT not in sys.path:
    sys.path.insert(0, _REPO_ROOT)

from utils.nx_helpers import (  # noqa: E402
    get_output_folder,
    log_info,
    require_work_part,
    run_journal,
    safe_part_name,
)

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


def main(session):
    part = require_work_part(session)
    if part is None:
        return

    output_folder = get_output_folder()
    log_info(session, f"Attribute discovery output folder: {output_folder}")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_path = os.path.join(output_folder, f"ATTR_DISCOVERY_{safe_part_name(part)}_{timestamp}.txt")

    rows = {}
    try:
        for attr_info in part.GetUserAttributes():
            rows[attr_info.Title] = (_type_name(attr_info.Type), _attr_value_str(attr_info))
    except Exception as exc:
        rows["__GetUserAttributes_error__"] = ("Error", str(exc))

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

    header_line = f"{'Title':<40} | {'Type':<12} | Value"
    separator = "-" * 80
    data_lines = [
        f"{title:<40} | {type_str:<12} | {value}"
        for title, (type_str, value) in sorted(rows.items())
    ]
    report_lines = [
        "NX Attribute Discovery Report",
        f"Part file : {part.FullPath}",
        f"Date      : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"Attributes: {len(rows)}",
        separator,
        header_line,
        separator,
        *data_lines,
        separator,
        f"Report saved to: {report_path}",
    ]

    with open(report_path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(report_lines))

    log_info(session, "\n".join(report_lines))


if __name__ == "__main__":
    run_journal(main)
