"""
Journal 01 — HLA STEP Export
Exports the active work part to STEP using settings from config/step_export.yaml.
Run via: NX > Tools > Journal > Play
"""

import os
import sys

import NXOpen
import NXOpen.UF

# Allow imports from the repo utils/ directory
_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _REPO_ROOT not in sys.path:
    sys.path.insert(0, _REPO_ROOT)

import yaml  # noqa: E402
from utils.nx_helpers import get_session, get_work_part, prompt_folder  # noqa: E402


def _load_config():
    config_path = os.path.join(_REPO_ROOT, "config", "step_export.yaml")
    with open(config_path, "r") as fh:
        return yaml.safe_load(fh)


def _get_part_attr(part, attr_name, fallback=""):
    """Read a string user attribute from a part; return fallback if absent."""
    try:
        attr_obj = part.GetUserAttribute(attr_name, NXOpen.NXObject.AttributeType.String, -1)
        return attr_obj.StringValue.strip()
    except Exception:
        return fallback


def _build_output_filename(part, config):
    part_number = _get_part_attr(part, "PART_NUMBER")
    revision = _get_part_attr(part, "REVISION")

    if part_number:
        template = config.get("output_naming", "{part_number}_REV{revision}.stp")
        return template.format(part_number=part_number, revision=revision)

    # Fallback: use the part file name without extension
    base = os.path.splitext(os.path.basename(part.FullPath))[0]
    return base + ".stp"


def main():
    session = get_session()
    part = get_work_part()

    if part is None:
        print("ERROR: No work part loaded.")
        return

    config = _load_config()

    output_folder = prompt_folder("Select STEP Output Folder")
    if output_folder is None:
        print("Export cancelled by user.")
        return

    filename = _build_output_filename(part, config)
    output_path = os.path.join(output_folder, filename)

    step_version_str = config.get("step_version", "AP214").upper()
    export_as = (
        NXOpen.StepCreator.ExportAsOption.Ap242
        if step_version_str == "AP242"
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

    print(f"STEP export complete: {output_path}")


if __name__ == "__main__":
    main()
