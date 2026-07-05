"""
Journal 06 - Automatic PDF + STEP Export
Exports the active work part to STEP and its drawing sheets to PDF.
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


_INVALID_FILENAME_CHARS = '<>:"/\\|?*'


def _clean_filename_token(value, fallback="part"):
    text = str(value or "").strip()
    if not text:
        return fallback
    cleaned = "".join("_" if char in _INVALID_FILENAME_CHARS else char for char in text)
    cleaned = cleaned.strip(" .")
    return cleaned or fallback


def _get_revision(part):
    return get_string_attr(part, "DB_PART_REV") or get_string_attr(part, "REVISION")


def _build_step_output_filename(part):
    part_number = get_string_attr(part, "DB_PART_NO") or get_string_attr(part, "PART_NUMBER")
    revision = _get_revision(part)

    if part_number:
        part_number = _clean_filename_token(part_number)
        if revision:
            revision = _clean_filename_token(revision, fallback="")
            return f"{part_number}_REV{revision}.stp"
        return f"{part_number}.stp"

    return _clean_filename_token(safe_part_name(part)) + ".stp"


def _build_pdf_output_filename(part, sheet, sheet_count):
    drawing_number = get_string_attr(part, "DRAWING_NUMBER")
    revision = _get_revision(part)
    sheet_suffix = ""

    if sheet_count > 1:
        sheet_name = _clean_filename_token(getattr(sheet, "Name", ""), fallback="sheet")
        sheet_suffix = f"_{sheet_name}"

    if drawing_number:
        base_name = _clean_filename_token(drawing_number)
        if revision:
            revision = _clean_filename_token(revision, fallback="")
            return f"{base_name}_REV{revision}{sheet_suffix}.pdf"
        return f"{base_name}{sheet_suffix}.pdf"

    return _clean_filename_token(safe_part_name(part)) + sheet_suffix + ".pdf"


def _sheet_count(drawing_sheets):
    try:
        return drawing_sheets.Count
    except Exception:
        pass
    try:
        return len(drawing_sheets.ToArray())
    except Exception:
        return 0


def _export_step(session, part, output_folder):
    output_path = os.path.join(output_folder, _build_step_output_filename(part))

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
    return output_path


def _export_current_sheet_to_pdf(session, pdf_path):
    exporter = session.DexManager.CreatePdfExporter()
    try:
        exporter.OutputFile = pdf_path
        exporter.Commit()
    finally:
        exporter.Destroy()


def _export_drawing_pdfs(session, part, output_folder):
    drawing_sheets = part.DrawingSheets
    sheet_count = _sheet_count(drawing_sheets)
    if sheet_count == 0:
        log_info(session, "No drawing sheets found on active work part; PDF export skipped.")
        return []

    pdf_paths = []
    original_display_part = session.Parts.Display
    try:
        session.Parts.SetDisplay(part, False, True, NXOpen.PartCollection.SdpsAction.None_)
        for sheet in drawing_sheets:
            pdf_name = _build_pdf_output_filename(part, sheet, sheet_count)
            pdf_path = os.path.join(output_folder, pdf_name)
            part.DrawingSheets.SetCurrentSheet(sheet)
            _export_current_sheet_to_pdf(session, pdf_path)
            pdf_paths.append(pdf_path)
            log_info(session, f"PDF export complete: {pdf_path}")
    finally:
        if original_display_part is not None:
            session.Parts.SetDisplay(
                original_display_part,
                False,
                True,
                NXOpen.PartCollection.SdpsAction.None_,
            )

    return pdf_paths


def main(session):
    part = require_work_part(session)
    if part is None:
        return

    output_folder = get_output_folder()
    log_info(session, f"Automatic PDF + STEP output folder: {output_folder}")

    step_path = _export_step(session, part, output_folder)
    pdf_paths = _export_drawing_pdfs(session, part, output_folder)

    log_info(
        session,
        "Automatic PDF + STEP export complete. "
        f"STEP files: 1; PDF files: {len(pdf_paths)}; Output folder: {output_folder}",
    )
    log_info(session, f"STEP file: {step_path}")


if __name__ == "__main__":
    run_journal(main)
