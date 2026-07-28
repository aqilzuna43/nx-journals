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
        validator = inspect.getsource(
            self.journal.evaluate_step_validation
        )
        self.assertIn("ExportFromOption.DisplayPart", source)
        self.assertIn("Scope.EntirePart", source)
        self.assertIn("exporter.LayerMask = STEP_LAYER_MASK", source)
        self.assertNotIn("exporter.InputFile", source)
        self.assertIn("ensure_step_source_loaded", source)
        self.assertIn("parse_step_translator_log", source)
        self.assertIn('"FAILED_ZERO_GEOMETRY"', validator)

    def test_runtime_version_uses_capability_before_environment(self):
        session = types.SimpleNamespace(
            GetEnvironmentVariableValue=lambda name: (
                "2506.5000" if name == "UGII_VERSION" else ""
            )
        )
        self.assertEqual(
            self.journal.runtime_nx_version(session),
            "2506.5000",
        )

    def test_full_load_capability_includes_children_and_disposes_status(self):
        class Status:
            NumberUnloadedParts = 0

            def __init__(self):
                self.disposed = False

            def Dispose(self):
                self.disposed = True

        status = Status()
        ensure = mock.Mock(return_value=status)
        session = types.SimpleNamespace(
            Parts=types.SimpleNamespace(
                EnsurePartsLoadedFully=ensure,
            )
        )
        part = types.SimpleNamespace(IsFullyLoaded=True)
        result = self.journal.ensure_step_source_loaded(session, part)
        self.assertEqual(result["status"], "SUCCESS")
        ensure.assert_called_once_with([part], True)
        self.assertTrue(status.disposed)

    def test_small_valid_step_is_not_rejected_by_file_size(self):
        validation = self.journal.evaluate_step_validation(
            "2506",
            {"status": "SUCCESS"},
            {
                "direct_solid_body_count": 1,
                "component_occurrence_count": 0,
                "descendant_solid_body_occurrence_count": 0,
                "component_limit_reached": False,
            },
            {
                "body_geometry_signatures": 2,
                "assembly_signatures": 0,
            },
            {
                "path": "",
                "solids_input": "",
                "solids_processed": "",
                "solids_as_sheets": "",
                "solids_not_processed": "",
            },
            64,
        )
        self.assertEqual(validation["result"], "SUCCESS")
        self.assertIn("STEP size=64 bytes", validation["message"])

    def test_translator_partial_geometry_is_rejected(self):
        validation = self.journal.evaluate_step_validation(
            "2506",
            {"status": "SUCCESS"},
            {
                "direct_solid_body_count": 2,
                "component_occurrence_count": 0,
                "descendant_solid_body_occurrence_count": 0,
                "component_limit_reached": False,
            },
            {
                "body_geometry_signatures": 4,
                "assembly_signatures": 0,
            },
            {
                "path": "sample.log",
                "solids_input": "2",
                "solids_processed": "1",
                "solids_as_sheets": "0",
                "solids_not_processed": "1",
            },
            285000,
        )
        self.assertEqual(
            validation["result"],
            "FAILED_TRANSLATOR_PARTIAL_GEOMETRY",
        )

    def test_translator_log_parser_reads_nx_counts(self):
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        step_path = Path(folder.name) / "sample.stp"
        step_path.write_text("ISO-10303-21;", encoding="utf-8")
        log_path = Path(folder.name) / "sample.log"
        log_path.write_text(
            "File_Name sample.stp\n"
            "Total number of solids input for this translation : 2\n"
            "Number of solids processed without problems : 2\n"
            "Number of solids with problems output as sheets : 0\n"
            "Number of solids not processed : 0\n",
            encoding="utf-8",
        )
        parsed = self.journal.parse_step_translator_log(step_path)
        self.assertEqual(parsed["solids_input"], "2")
        self.assertEqual(parsed["solids_processed"], "2")
        self.assertEqual(parsed["solids_not_processed"], "0")


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

    def test_draft_watermark_combines_revision_and_exact_wae_version(self):
        self.assertEqual(
            self.journal.build_pdf_watermark("A", "2"),
            "DRAFT_A.2",
        )
        self.assertEqual(
            self.journal.build_pdf_watermark("A", "2.0"),
            "DRAFT_A.2.0",
        )

    def test_watermark_prefers_loaded_model_wae_version(self):
        model = types.SimpleNamespace(
            GetStringAttribute=lambda name: (
                "2" if name == "WAE_VERSION" else ""
            )
        )
        drawing = types.SimpleNamespace(
            GetStringAttribute=lambda name: (
                "9" if name == "WAE_VERSION" else ""
            )
        )
        with mock.patch.object(
            self.journal,
            "loaded_master_candidate",
            return_value=model,
        ):
            watermark, source, warning = (
                self.journal.resolve_pdf_watermark(
                    types.SimpleNamespace(),
                    "264MN020016A01",
                    "A",
                    [{"part": drawing}],
                )
            )

        self.assertEqual(watermark, "DRAFT_A.2")
        self.assertEqual(source, "loaded model WAE_VERSION")
        self.assertEqual(warning, "")

    def test_watermark_falls_back_to_drawing_wae_version(self):
        drawing = types.SimpleNamespace(
            GetStringAttribute=lambda name: (
                "3" if name == "WAE_VERSION" else ""
            )
        )
        with mock.patch.object(
            self.journal,
            "loaded_master_candidate",
            return_value=None,
        ):
            watermark, source, warning = (
                self.journal.resolve_pdf_watermark(
                    types.SimpleNamespace(),
                    "264MN020016A01",
                    "A",
                    [{"part": drawing}],
                )
            )

        self.assertEqual(watermark, "DRAFT_A.3")
        self.assertEqual(source, "drawing WAE_VERSION")
        self.assertEqual(warning, "")

    def test_missing_wae_version_uses_revision_only_with_warning(self):
        with mock.patch.object(
            self.journal,
            "loaded_master_candidate",
            return_value=None,
        ):
            watermark, source, warning = (
                self.journal.resolve_pdf_watermark(
                    types.SimpleNamespace(),
                    "264MN020016A01",
                    "A",
                    [{"part": object()}],
                )
            )

        self.assertEqual(watermark, "DRAFT_A")
        self.assertEqual(source, "revision-only fallback")
        self.assertIn("WAE_VERSION is blank or unavailable", warning)

    def test_drawing_specs_use_canonical_specification_identifier(self):
        self.assertEqual(
            self.journal.teamcenter_drawing_specs(
                "264MN020016A01",
                "A",
                2,
            ),
            [
                (
                    "@DB/264MN020016A01/A/specification/"
                    "264MN020016A01-A-dwg2"
                )
            ],
        )

    def test_runtime_identity_marks_canonical_nx2506_build(self):
        self.assertEqual(
            self.journal.JOURNAL_BUILD_ID,
            "J07-NX2506-STEP-VALIDATION-V5",
        )
        self.assertTrue(
            self.journal.runtime_source_path().endswith(
                "07_datapack_pdf_step_export.py"
            )
        )

    def test_drawing_open_uses_open_display(self):
        class Status:
            def __init__(self):
                self.disposed = False

            def Dispose(self):
                self.disposed = True

        part = types.SimpleNamespace(Tag=42, Name="drawing")
        status = Status()
        parts = types.SimpleNamespace(
            OpenDisplay=mock.Mock(return_value=(part, status)),
            OpenBase=mock.Mock(
                side_effect=AssertionError("OpenBase must not open drawings")
            ),
        )
        session = types.SimpleNamespace(Parts=parts)

        opened = self.journal.open_display_part(
            session,
            (
                "@DB/264MN020016A01/A/specification/"
                "264MN020016A01-A-dwg1"
            ),
            set(),
            [],
            "drawing",
        )

        self.assertIs(opened["part"], part)
        self.assertTrue(opened["opened_by_journal"])
        self.assertTrue(status.disposed)
        parts.OpenDisplay.assert_called_once()
        parts.OpenBase.assert_not_called()

    def test_closed_dwg1_opens_when_later_drawings_are_missing(self):
        drawing = types.SimpleNamespace(
            Tag=42,
            Name="drawing",
            JournalIdentifier=(
                "@DB/264MN020016A01/A/specification/"
                "264MN020016A01-A-dwg1"
            ),
            DrawingSheets=[object(), object(), object()],
        )

        class Parts:
            Display = None

            def __iter__(self):
                return iter(())

            def OpenDisplay(self, specification):
                if specification.endswith("-dwg1"):
                    return drawing, None
                raise RuntimeError("drawing does not exist")

        candidates, attempts = self.journal.resolve_drawing_candidates(
            types.SimpleNamespace(Parts=Parts()),
            "264MN020016A01",
            "A",
            [],
        )

        self.assertEqual(len(candidates), 1)
        self.assertIs(candidates[0]["part"], drawing)
        self.assertEqual(candidates[0]["drawing_index"], 1)
        self.assertEqual(len(attempts), 9)

    def test_preloaded_drawing_is_reused_without_reopening_dwg1(self):
        drawing = types.SimpleNamespace(
            Tag=42,
            Name="drawing",
            JournalIdentifier=(
                "@DB/264MN020016A01/A/specification/"
                "264MN020016A01-A-dwg1"
            ),
            DrawingSheets=[object(), object(), object()],
        )
        open_display = mock.Mock(side_effect=RuntimeError("missing"))

        class Parts:
            Display = drawing

            def __iter__(self):
                return iter((drawing,))

            OpenDisplay = open_display

        candidates, attempts = self.journal.resolve_drawing_candidates(
            types.SimpleNamespace(Parts=Parts()),
            "264MN020016A01",
            "A",
            [],
        )

        self.assertEqual(len(candidates), 1)
        self.assertIs(candidates[0]["part"], drawing)
        self.assertEqual(len(attempts), 8)
        self.assertEqual(open_display.call_count, 8)
        self.assertFalse(
            any(specification.endswith("-dwg1") for specification in attempts)
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
            ActionOption=types.SimpleNamespace(Native="Native"),
            OutputTextOption=types.SimpleNamespace(Polylines="Polylines"),
        )

        self.journal.export_drawing_pdf(
            drawing_part,
            sheets,
            "combined.pdf",
            "DRAFT_A.2",
        )

        self.assertIs(builder.SourceBuilder.sheets, sheets)
        self.assertEqual(builder.Filename, "combined.pdf")
        self.assertFalse(builder.Append)
        self.assertTrue(builder.AddWatermark)
        self.assertEqual(builder.Watermark, "DRAFT_A.2")
        self.assertEqual(builder.OutputText, "Polylines")
        self.assertFalse(hasattr(builder, "CustomSymbolsInForeground"))
        self.assertEqual(builder.commit_count, 1)
        self.assertEqual(builder.destroy_count, 1)
        self.assertEqual(sheets[0].open_count, 1)
        self.assertEqual(sheets[1].open_count, 0)
        self.assertEqual(sheets[2].open_count, 0)

    def test_pdf_export_fails_when_nx_rejects_required_watermark(self):
        class Sheet:
            def Open(self):
                pass

        class Builder:
            def __init__(self):
                self.SourceBuilder = types.SimpleNamespace(
                    SetSheets=lambda sheets: None
                )
                self.destroy_count = 0

            def __setattr__(self, name, value):
                if name == "Watermark":
                    raise RuntimeError("watermark API unavailable")
                object.__setattr__(self, name, value)

            def Commit(self):
                raise AssertionError("Commit must not run")

            def Destroy(self):
                self.destroy_count += 1

        builder = Builder()
        drawing_part = types.SimpleNamespace(
            PlotManager=types.SimpleNamespace(
                CreatePrintPdfbuilder=lambda: builder
            )
        )
        self.journal.NXOpen.PrintPDFBuilder = types.SimpleNamespace(
            ActionOption=types.SimpleNamespace(Native="Native"),
            OutputTextOption=types.SimpleNamespace(Polylines="Polylines"),
        )

        with self.assertRaisesRegex(
            RuntimeError,
            "could not apply required watermark DRAFT_A.2",
        ):
            self.journal.export_drawing_pdf(
                drawing_part,
                [Sheet()],
                "combined.pdf",
                "DRAFT_A.2",
            )

        self.assertEqual(builder.destroy_count, 1)

    def test_pdf_export_fails_when_nx_rejects_watermark_enable(self):
        class Sheet:
            def Open(self):
                pass

        class Builder:
            def __init__(self):
                self.SourceBuilder = types.SimpleNamespace(
                    SetSheets=lambda sheets: None
                )
                self.destroy_count = 0

            def __setattr__(self, name, value):
                if name == "AddWatermark":
                    raise RuntimeError("watermark enable unavailable")
                object.__setattr__(self, name, value)

            def Commit(self):
                raise AssertionError("Commit must not run")

            def Destroy(self):
                self.destroy_count += 1

        builder = Builder()
        drawing_part = types.SimpleNamespace(
            PlotManager=types.SimpleNamespace(
                CreatePrintPdfbuilder=lambda: builder
            )
        )
        self.journal.NXOpen.PrintPDFBuilder = types.SimpleNamespace(
            ActionOption=types.SimpleNamespace(Native="Native"),
            OutputTextOption=types.SimpleNamespace(Polylines="Polylines"),
        )

        with self.assertRaisesRegex(
            RuntimeError,
            "could not apply required watermark DRAFT_A.2",
        ):
            self.journal.export_drawing_pdf(
                drawing_part,
                [Sheet()],
                "combined.pdf",
                "DRAFT_A.2",
            )

        self.assertEqual(builder.destroy_count, 1)

    def test_pdf_export_fails_when_nx_rejects_polyline_text_output(self):
        class Sheet:
            def Open(self):
                pass

        class Builder:
            def __init__(self):
                self.SourceBuilder = types.SimpleNamespace(
                    SetSheets=lambda sheets: None
                )
                self.destroy_count = 0

            def __setattr__(self, name, value):
                if name == "OutputText":
                    raise RuntimeError("polyline output unavailable")
                object.__setattr__(self, name, value)

            def Commit(self):
                raise AssertionError("Commit must not run")

            def Destroy(self):
                self.destroy_count += 1

        builder = Builder()
        drawing_part = types.SimpleNamespace(
            PlotManager=types.SimpleNamespace(
                CreatePrintPdfbuilder=lambda: builder
            )
        )
        self.journal.NXOpen.PrintPDFBuilder = types.SimpleNamespace(
            ActionOption=types.SimpleNamespace(Native="Native"),
            OutputTextOption=types.SimpleNamespace(Polylines="Polylines"),
        )

        with self.assertRaisesRegex(
            RuntimeError,
            "could not apply required polyline text output",
        ):
            self.journal.export_drawing_pdf(
                drawing_part,
                [Sheet()],
                "combined.pdf",
                "DRAFT_A.2",
            )

        self.assertEqual(builder.destroy_count, 1)

    def run_grouped_export(self, candidates, wae_version="2"):
        session = types.SimpleNamespace()
        original_display = object()
        original_work = object()
        logs = []
        model = (
            types.SimpleNamespace(
                GetStringAttribute=lambda name: (
                    wae_version if name == "WAE_VERSION" else ""
                )
            )
            if wae_version is not None
            else None
        )

        def create_pdf(_part, _sheets, output_path, _watermark):
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
        ) as restorer, mock.patch.object(
            self.journal,
            "loaded_master_candidate",
            return_value=model,
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
                original_display,
                original_work,
                logs,
            )
        restorer.assert_called_once()
        restored_session, restored_display, restored_work, restored_logs = (
            restorer.call_args.args
        )
        self.assertIs(restored_session, session)
        self.assertIs(restored_display, original_display)
        self.assertIs(restored_work, original_work)
        self.assertIs(restored_logs, logs)
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
        self.assertEqual(exporter.call_args.args[3], "DRAFT_A.2")

    def test_missing_wae_warning_is_reported_without_failing_pdf(self):
        candidate = {
            "part": types.SimpleNamespace(
                Name="drawing",
                DrawingSheets=[object()],
            ),
            "drawing_index": 1,
            "opened_by_journal": False,
        }

        result, exporter = self.run_grouped_export(
            [candidate],
            wae_version=None,
        )

        self.assertEqual(result["result"], "SUCCESS")
        self.assertEqual(result["watermark"], "DRAFT_A")
        self.assertIn("WAE_VERSION is blank or unavailable", result["message"])
        self.assertEqual(exporter.call_args.args[3], "DRAFT_A")

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
