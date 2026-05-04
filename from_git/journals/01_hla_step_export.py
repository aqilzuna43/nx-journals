"""
Journal 01 - HLA STEP Export
Exports the active work part to STEP.
Run via: NX > Tools > Journal > Play
"""

import os
import sys

import NXOpen
import NXOpen.UF

_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _REPO_ROOT not in sys.path:
    sys.path.insert(0, _REPO_ROOT)

from utils.nx_helpers import (  # noqa: E402
    get_output_folder,
    get_string_attr,
    log_info,
    require_work_part,
    run_journal,
    safe_part_name,
)


def _build_output_filename(part):
    part_number = get_string_attr(part, "DB_PART_NO") or get_string_attr(part, "PART_NUMBER")
    revision = get_string_attr(part, "DB_PART_REV") or get_string_attr(part, "REVISION")

    if part_number:
        if revision:
            return f"{part_number}_REV{revision}.stp"
        return f"{part_number}.stp"

    return safe_part_name(part) + ".stp"


def main(session):
    part = require_work_part(session)
    if part is None:
        return

    output_folder = get_output_folder()
    output_path = os.path.join(output_folder, _build_output_filename(part))

    exporter = session.DexManager.CreateStepCreator()
    try:
        exporter.OutputFile = output_path
        exporter.ObjectTypes.ExportSelectionBlock.SelectionScope = (
            NXOpen.ObjectTypes.SelectionScope.WorkPart
        )
        exporter.ExportAs = NXOpen.StepCreator.ExportAsOption.Ap214
        exporter.Commit()
    finally:
        exporter.Destroy()

    log_info(session, f"STEP export complete: {output_path}")


if __name__ == "__main__":
    run_journal(main)
