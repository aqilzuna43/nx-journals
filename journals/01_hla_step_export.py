"""
Journal 01 - HLA STEP Export
Exports the active work part to STEP using config/step_export.json.
Run via: NX > Tools > Journal > Play
"""

import os
import sys

import NXOpen
import NXOpen.UF

_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _REPO_ROOT not in sys.path:
    sys.path.insert(0, _REPO_ROOT)

from utils.config_loader import load_json_config  # noqa: E402
from utils.nx_helpers import (  # noqa: E402
    get_string_attr,
    log_info,
    prompt_folder,
    require_work_part,
    run_journal,
    safe_part_name,
)


def _load_config():
    return load_json_config(_REPO_ROOT, os.path.join("config", "step_export.json"))


def _build_output_filename(part, config):
    part_number = get_string_attr(part, "DB_PART_NO") or get_string_attr(part, "PART_NUMBER")
    revision = get_string_attr(part, "DB_PART_REV") or get_string_attr(part, "REVISION")

    if part_number:
        template = config.get("output_naming", "{part_number}_REV{revision}.stp")
        return template.format(part_number=part_number, revision=revision)

    return safe_part_name(part) + ".stp"


def main(session):
    part = require_work_part(session)
    if part is None:
        return

    output_folder = prompt_folder("Select STEP Output Folder")
    if output_folder is None:
        log_info(session, "STEP export cancelled by user.")
        return

    config = _load_config()
    output_path = os.path.join(output_folder, _build_output_filename(part, config))

    step_version = str(config.get("step_version", "AP214")).upper()
    export_as = (
        NXOpen.StepCreator.ExportAsOption.Ap242
        if step_version == "AP242"
        else NXOpen.StepCreator.ExportAsOption.Ap214
    )

    exporter = session.DexManager.CreateStepCreator()
    try:
        exporter.OutputFile = output_path
        exporter.ObjectTypes.ExportSelectionBlock.SelectionScope = (
            NXOpen.ObjectTypes.SelectionScope.WorkPart
        )
        exporter.ExportAs = export_as
        exporter.Commit()
    finally:
        exporter.Destroy()

    log_info(session, f"STEP export complete: {output_path}")


if __name__ == "__main__":
    run_journal(main)
