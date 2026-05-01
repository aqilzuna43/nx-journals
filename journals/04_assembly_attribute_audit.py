"""
Journal 04 — Assembly Attribute Audit
Checks required NX part attributes across the full assembly and reports issues to Excel.
Run via: NX > Tools > Journal > Play
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

_REVISION_PATTERN = re.compile(r"^([A-Z]|\d+)$")

_ISSUES_HEADERS = [
    "Part Number",
    "Attribute",
    "Issue",
    "Current Value",
    "Component Path",
]

_SUMMARY_HEADERS = [
    "Metric",
    "Value",
]


def _get_attr(nxobj, attr_name):
    try:
        attr = nxobj.GetUserAttribute(attr_name, NXOpen.NXObject.AttributeType.String, -1)
        return attr.StringValue.strip()
    except Exception:
        return ""


def _component_path(component):
    """Build a slash-separated path string for a component's location in the tree."""
    parts = []
    obj = component
    while obj is not None:
        try:
            parts.append(obj.Name)
            obj = obj.Parent
        except Exception:
            break
    return " / ".join(reversed(parts))


def _audit_component(component):
    """
    Check a single component's required attributes.
    Returns a list of (part_number, attr_name, issue, current_value, path) tuples.
    """
    issues = []
    proto = component.Prototype
    if proto is None:
        return issues

    pn = _get_attr(proto, "PART_NUMBER")
    path = _component_path(component)

    for attr in _REQUIRED_ATTRS:
        value = _get_attr(proto, attr)

        if not value:
            issues.append((pn, attr, "Missing or empty", value, path))
            continue

        if attr == "REVISION" and not _REVISION_PATTERN.match(value):
            issues.append((
                pn,
                attr,
                "Invalid format (expected single letter or integer)",
                value,
                path,
            ))

    return issues


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

    hla_pn = _get_attr(part, "PART_NUMBER") or os.path.splitext(os.path.basename(part.FullPath))[0]
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"AUDIT_{hla_pn}_{timestamp}.xlsx"
    output_path = os.path.join(output_folder, filename)

    root = part.ComponentAssembly.RootComponent
    all_components = list(traverse_assembly(root))

    all_issues = []
    pass_counts = {attr: 0 for attr in _REQUIRED_ATTRS}
    fail_counts = {attr: 0 for attr in _REQUIRED_ATTRS}

    for component in all_components:
        proto = component.Prototype
        if proto is None:
            continue

        component_issues = _audit_component(component)
        all_issues.extend(component_issues)

        failed_attrs = {issue[1] for issue in component_issues}
        for attr in _REQUIRED_ATTRS:
            if attr in failed_attrs:
                fail_counts[attr] += 1
            else:
                pass_counts[attr] += 1

    writer = ExcelWriter(output_path)

    # Issues sheet
    writer.add_header_row("Issues", _ISSUES_HEADERS)
    for issue_row in all_issues:
        writer.add_row("Issues", list(issue_row), highlight_amber=True)

    # Summary sheet
    writer.add_header_row("Summary", _SUMMARY_HEADERS)
    writer.add_row("Summary", ["Total components checked", len(all_components)])
    writer.add_row("Summary", ["Total issues found", len(all_issues)])
    writer.add_row("Summary", ["", ""])
    writer.add_row("Summary", ["Attribute", "Pass / Fail"])
    for attr in _REQUIRED_ATTRS:
        writer.add_row("Summary", [attr, f"{pass_counts[attr]} pass / {fail_counts[attr]} fail"])

    writer.save()

    print(f"Audit complete.")
    print(f"  Components checked : {len(all_components)}")
    print(f"  Issues found       : {len(all_issues)}")
    print(f"  Report written to  : {output_path}")


if __name__ == "__main__":
    main()
