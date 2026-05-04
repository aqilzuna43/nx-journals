"""
Journal 03 - Batch Drawing PDF Export
Traverses the assembly, finds drawing sheets, and exports each to PDF.
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
    unique_prototype_parts,
)


def _export_current_sheet_to_pdf(session, pdf_path):
    exporter = session.DexManager.CreatePdfExporter()
    try:
        exporter.OutputFile = pdf_path
        exporter.Commit()
    finally:
        exporter.Destroy()


def _sheet_count(drawing_sheets):
    try:
        return drawing_sheets.Count
    except Exception:
        return 0


def main(session):
    part = require_work_part(session)
    if part is None:
        return

    output_folder = get_output_folder()
    log_info(session, f"PDF output folder: {output_folder}")

    pdf_count = 0
    original_display_part = session.Parts.Display
    try:
        for proto_part in unique_prototype_parts(part, include_work_part=True):
            drawing_sheets = proto_part.DrawingSheets
            count = _sheet_count(drawing_sheets)
            if count == 0:
                continue

            drawing_number = get_string_attr(proto_part, "DRAWING_NUMBER")
            revision = get_string_attr(proto_part, "DB_PART_REV") or get_string_attr(proto_part, "REVISION")

            for sheet in drawing_sheets:
                if drawing_number:
                    rev_suffix = f"_REV{revision}" if revision else ""
                    sheet_suffix = f"_{sheet.Name}" if count > 1 else ""
                    pdf_name = f"{drawing_number}{rev_suffix}{sheet_suffix}.pdf"
                else:
                    sheet_suffix = f"_{sheet.Name}" if count > 1 else ""
                    pdf_name = f"{safe_part_name(proto_part)}{sheet_suffix}.pdf"

                pdf_path = os.path.join(output_folder, pdf_name)
                session.Parts.SetDisplay(proto_part, False, True, NXOpen.PartCollection.SdpsAction.None_)
                proto_part.DrawingSheets.SetCurrentSheet(sheet)
                _export_current_sheet_to_pdf(session, pdf_path)
                pdf_count += 1
                log_info(session, f"Exported PDF: {pdf_name}")
    finally:
        if original_display_part is not None:
            session.Parts.SetDisplay(original_display_part, False, True, NXOpen.PartCollection.SdpsAction.None_)

    log_info(session, f"Batch PDF export complete. {pdf_count} PDF(s) written to: {output_folder}")


if __name__ == "__main__":
    run_journal(main)
