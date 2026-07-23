import importlib.util
import inspect
import sys
import tempfile
import types
import unittest
from pathlib import Path
from unittest import mock


ROOT = Path(__file__).resolve().parents[1]
JOURNAL = (
    ROOT / "from_git" / "journals" / "07_datapack_pdf_step_export.py"
)


def load_journal():
    sys.modules.setdefault("NXOpen", types.ModuleType("NXOpen"))
    spec = importlib.util.spec_from_file_location("journal07", JOURNAL)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class StepValidationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.journal = load_journal()

    def write_step(self, content):
        folder = tempfile.TemporaryDirectory()
        path = Path(folder.name) / "sample.stp"
        path.write_text(content, encoding="utf-8")
        self.addCleanup(folder.cleanup)
        return path

    def test_header_only_step_has_no_body_geometry(self):
        path = self.write_step(
            "ISO-10303-21;\nHEADER;\nENDSEC;\nDATA;\n"
            "#1=PRODUCT('X','','',());\nENDSEC;\nEND-ISO-10303-21;\n"
        )
        self.assertEqual(
            self.journal.step_body_signature_count(path),
            0,
        )

    def test_body_entities_in_data_section_are_detected(self):
        path = self.write_step(
            "ISO-10303-21;\nHEADER;\nENDSEC;\nDATA;\n"
            "#1=MANIFOLD_SOLID_BREP('',#2);\n"
            "#2=CLOSED_SHELL('',());\nENDSEC;\nEND-ISO-10303-21;\n"
        )
        self.assertEqual(
            self.journal.step_body_signature_count(path),
            2,
        )

    def test_body_words_outside_data_section_do_not_pass(self):
        path = self.write_step(
            "ISO-10303-21;\nHEADER;\n"
            "FILE_DESCRIPTION(('MANIFOLD_SOLID_BREP'),'2;1');\n"
            "ENDSEC;\nDATA;\n#1=PRODUCT('X','','',());\nENDSEC;\n"
        )
        self.assertEqual(
            self.journal.step_body_signature_count(path),
            0,
        )

    def test_export_uses_proven_display_scope_and_layer_mask(self):
        source = inspect.getsource(self.journal.export_step_from_part)
        self.assertIn("ExportFromOption.DisplayPart", source)
        self.assertIn("Scope.EntirePart", source)
        self.assertIn("exporter.LayerMask = STEP_LAYER_MASK", source)
        self.assertNotIn("exporter.InputFile", source)
        self.assertIn('"FAILED_ZERO_GEOMETRY"', source)


class PdfGroupingTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.journal = load_journal()

    def test_single_drawing_uses_plain_part_revision_name(self):
        self.assertEqual(
            self.journal.build_pdf_filename(
                "264MN020016A01",
                "A",
                "DWG1",
                1,
            ),
            "264MN020016A01_REVA.pdf",
        )

    def test_multiple_drawings_receive_dwg_suffixes(self):
        self.assertEqual(
            self.journal.build_pdf_filename(
                "264MN020016A01",
                "A",
                "DWG2",
                2,
            ),
            "264MN020016A01_REVA_DWG2.pdf",
        )

    def test_duplicate_or_missing_tokens_are_made_unique(self):
        candidates = [
            {"part": object(), "drawing_index": 1},
            {"part": object(), "drawing_index": 1},
            {"part": object(), "drawing_index": None},
            {"part": object(), "drawing_index": 3},
        ]
        with mock.patch.object(
            self.journal,
            "drawing_index_from_part",
            return_value=None,
        ):
            self.assertEqual(
                self.journal.unique_drawing_tokens(candidates),
                ["DWG1", "DWG2", "DWG4", "DWG3"],
            )

    def test_all_sheets_are_sent_to_one_pdf_builder_commit(self):
        class Sheet:
            def __init__(self):
                self.open_count = 0

            def Open(self):
                self.open_count += 1

        class SourceBuilder:
            def __init__(self):
                self.sheets = None

            def SetSheets(self, sheets):
                self.sheets = sheets

        class Builder:
            def __init__(self):
                self.SourceBuilder = SourceBuilder()
                self.commit_count = 0
                self.destroy_count = 0

            def Commit(self):
                self.commit_count += 1

            def Destroy(self):
                self.destroy_count += 1

        builder = Builder()
        drawing_part = types.SimpleNamespace(
            PlotManager=types.SimpleNamespace(
                CreatePrintPdfbuilder=lambda: builder
            )
        )
        sheets = [Sheet(), Sheet(), Sheet()]
        self.journal.NXOpen.PrintPDFBuilder = types.SimpleNamespace(
            ActionOption=types.SimpleNamespace(Native="Native")
        )

        self.journal.export_drawing_pdf(
            drawing_part,
            sheets,
            "combined.pdf",
        )

        self.assertIs(builder.SourceBuilder.sheets, sheets)
        self.assertEqual(builder.Filename, "combined.pdf")
        self.assertFalse(builder.Append)
        self.assertEqual(builder.commit_count, 1)
        self.assertEqual(builder.destroy_count, 1)
        self.assertEqual(sheets[0].open_count, 1)
        self.assertEqual(sheets[1].open_count, 0)
        self.assertEqual(sheets[2].open_count, 0)

    def run_grouped_export(self, candidates):
        session = types.SimpleNamespace()

        def create_pdf(_part, _sheets, output_path):
            Path(output_path).write_bytes(b"%PDF-test")

        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        with mock.patch.object(
            self.journal,
            "resolve_drawing_candidates",
            return_value=(candidates, []),
        ), mock.patch.object(
            self.journal,
            "set_display_part",
        ), mock.patch.object(
            self.journal,
            "restore_parts",
        ), mock.patch.object(
            self.journal,
            "export_drawing_pdf",
            side_effect=create_pdf,
        ) as exporter:
            result = self.journal.export_pdfs_for_instruction(
                session,
                folder.name,
                "264MN020016A01",
                "A",
                None,
                None,
                [],
            )
        return result, exporter

    def test_three_sheet_drawing_returns_one_pdf_path(self):
        sheets = [object(), object(), object()]
        candidate = {
            "part": types.SimpleNamespace(
                Name="drawing",
                DrawingSheets=sheets,
            ),
            "drawing_index": 1,
            "opened_by_journal": False,
        }

        result, exporter = self.run_grouped_export([candidate])

        self.assertEqual(result["result"], "SUCCESS")
        self.assertEqual(len(result["paths"]), 1)
        self.assertTrue(
            result["paths"][0].endswith(
                "264MN020016A01_REVA.pdf"
            )
        )
        self.assertEqual(exporter.call_count, 1)
        self.assertEqual(exporter.call_args.args[1], sheets)

    def test_two_drawings_return_two_suffixed_pdf_paths(self):
        candidates = [
            {
                "part": types.SimpleNamespace(
                    Name="drawing1",
                    DrawingSheets=[object(), object()],
                ),
                "drawing_index": 1,
                "opened_by_journal": False,
            },
            {
                "part": types.SimpleNamespace(
                    Name="drawing2",
                    DrawingSheets=[object()],
                ),
                "drawing_index": 2,
                "opened_by_journal": False,
            },
        ]

        result, exporter = self.run_grouped_export(candidates)

        self.assertEqual(result["result"], "SUCCESS")
        self.assertEqual(exporter.call_count, 2)
        self.assertEqual(
            [Path(path).name for path in result["paths"]],
            [
                "264MN020016A01_REVA_DWG1.pdf",
                "264MN020016A01_REVA_DWG2.pdf",
            ],
        )


if __name__ == "__main__":
    unittest.main()
