"""
Journal 03 — Batch Drawing PDF Export
Traverses the assembly, finds all associated drawing sheets, and exports each to PDF.
Run via: NX > Tools > Journal > Play
"""

import os
import sys

import NXOpen
import NXOpen.UF

_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _REPO_ROOT not in sys.path:
    sys.path.insert(0, _REPO_ROOT)

from utils.nx_helpers import get_session, get_work_part, prompt_folder, traverse_assembly  # noqa: E402


def _get_attr(nxobj, attr_name):
    try:
        attr = nxobj.GetUserAttribute(attr_name, NXOpen.NXObject.AttributeType.String, -1)
        return attr.StringValue.strip()
    except Exception:
        return ""


def _collect_unique_part_tags(root_component):
    """Return a set of unique Prototype tags found in the assembly tree."""
    seen = set()
    seen.add(root_component.Prototype.Tag if root_component.Prototype else None)
    for comp in traverse_assembly(root_component):
        proto = comp.Prototype
        if proto is not None:
            seen.add(proto.Tag)
    seen.discard(None)
    return seen


def _export_drawing_sheet(session, sheet, pdf_path):
    """Export a single NX drawing sheet to a PDF file."""
    pdfExporter = session.DexManager.CreatePdfExporter()
    try:
        pdfExporter.OutputFile = pdf_path
        pdfExporter.Commit()
    finally:
        pdfExporter.Destroy()


def main():
    session = get_session()
    part = get_work_part()

    if part is None:
        print("ERROR: No work part loaded.")
        return

    output_folder = prompt_folder("Select PDF Output Folder")
    if output_folder is None:
        print("PDF export cancelled by user.")
        return

    root = part.ComponentAssembly.RootComponent
    unique_proto_tags = _collect_unique_part_tags(root)

    pdf_count = 0

    for proto_tag in unique_proto_tags:
        try:
            proto_part = session.Parts.FindObject(str(proto_tag))
        except Exception:
            continue

        drawing_sheets = proto_part.DrawingSheets
        if drawing_sheets is None or drawing_sheets.Count == 0:
            continue

        drawing_number = _get_attr(proto_part, "DRAWING_NUMBER")
        revision = _get_attr(proto_part, "REVISION")

        for sheet in drawing_sheets:
            # Build PDF filename using DRAWING_NUMBER and REVISION attributes
            if drawing_number:
                rev_suffix = f"_REV{revision}" if revision else ""
                sheet_suffix = f"_{sheet.Name}" if drawing_sheets.Count > 1 else ""
                pdf_name = f"{drawing_number}{rev_suffix}{sheet_suffix}.pdf"
            else:
                base = os.path.splitext(os.path.basename(proto_part.FullPath))[0]
                sheet_suffix = f"_{sheet.Name}" if drawing_sheets.Count > 1 else ""
                pdf_name = f"{base}{sheet_suffix}.pdf"

            pdf_path = os.path.join(output_folder, pdf_name)

            # Make the drawing sheet the display part so the exporter targets it
            session.Parts.SetDisplay(proto_part, False, True, NXOpen.PartCollection.SdpsAction.None_)
            proto_part.DrawingSheets.SetCurrentSheet(sheet)

            _export_drawing_sheet(session, sheet, pdf_path)
            pdf_count += 1
            print(f"  Exported: {pdf_name}")

    print(f"\nBatch PDF export complete. {pdf_count} PDF(s) written to: {output_folder}")


if __name__ == "__main__":
    main()
