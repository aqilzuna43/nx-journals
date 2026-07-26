import importlib.util
import inspect
import json
import os
import sys
import tempfile
import types
import unittest
from pathlib import Path
from unittest import mock


ROOT = Path(__file__).resolve().parents[1]
JOURNAL = (
    ROOT
    / "from_git"
    / "journals"
    / "12_diagnose_pdf_watermark_symbols.py"
)


def load_journal():
    nxopen = types.ModuleType("NXOpen")
    nxopen.PrintPDFBuilder = types.SimpleNamespace(
        ActionOption=types.SimpleNamespace(Native="Native"),
        OutputTextOption=types.SimpleNamespace(
            Text="Text",
            Polylines="Polylines",
        ),
    )
    nxopen.BasePart = types.SimpleNamespace(
        CloseWholeTree=types.SimpleNamespace(FalseValue="FalseValue"),
        CloseModified=types.SimpleNamespace(CloseModified="CloseModified"),
    )
    nxopen.Session = types.SimpleNamespace()
    sys.modules["NXOpen"] = nxopen
    spec = importlib.util.spec_from_file_location("journal12", JOURNAL)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class FakeLogger:
    def __init__(self):
        self.lines = []

    def write(self, message=""):
        self.lines.append(str(message))


class FakeStatus:
    def __init__(self):
        self.disposed = False

    def Dispose(self):
        self.disposed = True


class PdfDiagnosticTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.journal = load_journal()

    def make_drawing(self, tag=42):
        sheet = types.SimpleNamespace(
            Name="Sheet 1",
            Open=mock.Mock(),
        )
        return types.SimpleNamespace(
            Tag=tag,
            Name="affected-drawing",
            JournalIdentifier=(
                "@DB/TEST-PART/A/specification/TEST-PART-A-dwg1"
            ),
            DrawingSheets=[sheet],
            Close=mock.Mock(),
        )

    def test_variant_matrix_is_complete_and_ordered(self):
        variants = self.journal.variant_definitions()

        self.assertEqual(
            [variant["token"] for variant in variants],
            [
                "00_WATERMARK_DISABLED_BASELINE",
                "01_TEXT_WATERMARK",
                "02_TEXT_FOREGROUND_SYMBOLS",
                "03_POLYLINE_WATERMARK",
                "04_POLYLINE_FOREGROUND_SYMBOLS",
            ],
        )
        self.assertEqual(len(variants), 5)
        self.assertFalse(variants[0]["add_watermark"])
        self.assertEqual(variants[1]["output_text"], "TEXT")
        self.assertTrue(variants[2]["custom_symbols_in_foreground"])
        self.assertEqual(variants[3]["output_text"], "POLYLINES")
        self.assertTrue(variants[4]["custom_symbols_in_foreground"])

    def test_identifier_comparison_normalizes_case_and_slashes(self):
        self.assertEqual(
            self.journal.identifier_key(
                "@DB/TEST/A/specification/TEST-A-dwg1"
            ),
            self.journal.identifier_key(
                "@db\\test\\a\\SPECIFICATION\\test-a-DWG1"
            ),
        )

    def test_preloaded_run_captures_canonical_target_without_opening(self):
        drawing = self.make_drawing()
        parts = types.SimpleNamespace(
            Display=drawing,
            OpenDisplay=mock.Mock(
                side_effect=AssertionError("must not open preloaded drawing")
            ),
        )
        session = types.SimpleNamespace(Parts=parts)
        logger = FakeLogger()

        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        target_path = os.path.join(folder.name, "LAST_TARGET.json")

        resolved = self.journal.resolve_diagnostic_drawing(
            session,
            target_path,
            logger,
        )

        self.assertEqual(resolved["mode"], "PRELOADED")
        self.assertIs(resolved["part"], drawing)
        self.assertFalse(resolved["opened_by_journal"])
        parts.OpenDisplay.assert_not_called()
        with open(target_path, "r", encoding="utf-8") as handle:
            stored = json.load(handle)
        self.assertEqual(
            stored["journal_identifier"],
            drawing.JournalIdentifier,
        )

    def test_closed_run_reuses_target_and_opens_with_open_display(self):
        drawing = self.make_drawing()
        assembly = types.SimpleNamespace(Tag=10, Name="assembly")
        status = FakeStatus()

        class Parts:
            Display = assembly

            def __iter__(self):
                return iter((assembly,))

            def __init__(self):
                self.OpenDisplay = mock.Mock(
                    return_value=(drawing, status)
                )

        parts = Parts()
        session = types.SimpleNamespace(Parts=parts)
        logger = FakeLogger()
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        target_path = os.path.join(folder.name, "LAST_TARGET.json")
        self.journal.write_target(
            target_path,
            self.journal.target_record(drawing),
        )

        resolved = self.journal.resolve_diagnostic_drawing(
            session,
            target_path,
            logger,
        )

        self.assertEqual(resolved["mode"], "CLOSED_AUTO")
        self.assertIs(resolved["part"], drawing)
        self.assertTrue(resolved["opened_by_journal"])
        parts.OpenDisplay.assert_called_once_with(drawing.JournalIdentifier)
        self.assertTrue(status.disposed)

    def test_variant_applies_text_and_foreground_controls(self):
        builder = types.SimpleNamespace()
        variant = self.journal.variant_definitions()[2]

        self.journal.apply_variant(
            builder,
            variant,
            "J12_WATERMARK_TEST",
        )

        self.assertTrue(builder.AddWatermark)
        self.assertEqual(builder.Watermark, "J12_WATERMARK_TEST")
        self.assertEqual(builder.OutputText, "Text")
        self.assertTrue(builder.CustomSymbolsInForeground)

    def test_variant_applies_polyline_output(self):
        builder = types.SimpleNamespace()
        variant = self.journal.variant_definitions()[3]

        self.journal.apply_variant(
            builder,
            variant,
            "J12_WATERMARK_TEST",
        )

        self.assertTrue(builder.AddWatermark)
        self.assertEqual(builder.OutputText, "Polylines")
        self.assertFalse(builder.CustomSymbolsInForeground)

    def test_export_records_defaults_and_commits_all_sheets_once(self):
        class SourceBuilder:
            def __init__(self):
                self.sheets = None

            def SetSheets(self, sheets):
                self.sheets = sheets

        class Builder:
            def __init__(self):
                self.SourceBuilder = SourceBuilder()
                self.AddWatermark = False
                self.Watermark = ""
                self.OutputText = "DefaultText"
                self.CustomSymbolsInForeground = False
                self.Colors = "Colors"
                self.RasterImages = False
                self.ShadedGeometry = False
                self.ImageResolution = "Resolution"
                self.commit_count = 0
                self.destroy_count = 0

            def Commit(self):
                self.commit_count += 1
                Path(self.Filename).write_bytes(b"%PDF-test")

            def Destroy(self):
                self.destroy_count += 1

        builder = Builder()
        sheets = [
            types.SimpleNamespace(Open=mock.Mock()),
            types.SimpleNamespace(Open=mock.Mock()),
        ]
        drawing = types.SimpleNamespace(
            PlotManager=types.SimpleNamespace(
                CreatePrintPdfbuilder=lambda: builder
            )
        )
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        output_path = os.path.join(folder.name, "variant.pdf")

        defaults, applied = self.journal.export_pdf_variant(
            drawing,
            sheets,
            output_path,
            self.journal.variant_definitions()[4],
            "J12_WATERMARK_TEST",
        )

        self.assertEqual(defaults["AddWatermark"], "False")
        self.assertEqual(applied["AddWatermark"], "True")
        self.assertEqual(applied["OutputText"], "Polylines")
        self.assertEqual(applied["CustomSymbolsInForeground"], "True")
        self.assertIs(builder.SourceBuilder.sheets, sheets)
        self.assertEqual(builder.commit_count, 1)
        self.assertEqual(builder.destroy_count, 1)
        sheets[0].Open.assert_called_once()
        sheets[1].Open.assert_not_called()

    def test_variant_matrix_continues_after_one_variant_fails(self):
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        logger = FakeLogger()
        calls = []

        def export_variant(
            _drawing,
            _sheets,
            output_path,
            variant,
            _watermark,
        ):
            calls.append(variant["token"])
            if variant["token"] == "01_TEXT_WATERMARK":
                raise RuntimeError("controlled failure")
            Path(output_path).write_bytes(b"%PDF-test")
            return {"default": "value"}, {"applied": "value"}

        with mock.patch.object(
            self.journal,
            "export_pdf_variant",
            side_effect=export_variant,
        ):
            outputs, failures = self.journal.run_variant_matrix(
                object(),
                [object()],
                folder.name,
                logger,
            )

        self.assertEqual(len(calls), 5)
        self.assertEqual(len(outputs), 4)
        self.assertEqual(len(failures), 1)
        self.assertEqual(failures[0]["token"], "01_TEXT_WATERMARK")
        self.assertIn("controlled failure", failures[0]["message"])

    def test_cleanup_restores_state_and_closes_only_journal_opened_part(self):
        drawing = self.make_drawing()
        original_display = types.SimpleNamespace(Name="assembly")
        original_work = types.SimpleNamespace(Name="work")
        original_sheet = types.SimpleNamespace(Open=mock.Mock())
        parts = types.SimpleNamespace(
            SetDisplay=mock.Mock(return_value=(None, None)),
            SetWork=mock.Mock(),
        )
        session = types.SimpleNamespace(Parts=parts)
        logger = FakeLogger()

        self.journal.cleanup_diagnostic(
            session,
            drawing,
            True,
            original_display,
            original_work,
            original_sheet,
            logger,
        )

        parts.SetDisplay.assert_called_once_with(
            original_display,
            False,
            True,
        )
        parts.SetWork.assert_called_once_with(original_work)
        drawing.Close.assert_called_once()
        original_sheet.Open.assert_not_called()

        drawing.Close.reset_mock()
        parts.SetDisplay.reset_mock()
        parts.SetWork.reset_mock()
        self.journal.cleanup_diagnostic(
            session,
            original_display,
            False,
            original_display,
            original_work,
            original_sheet,
            logger,
        )
        drawing.Close.assert_not_called()
        original_sheet.Open.assert_called_once()

    def test_diagnostic_has_no_layer_mutation_update_or_save_calls(self):
        source = inspect.getsource(self.journal)
        self.assertNotIn(".SetState(", source)
        self.assertNotIn(".ChangeStates(", source)
        self.assertNotIn(".SetObjectsVisibilityOnLayer(", source)
        self.assertNotIn(".UpdateViews(", source)
        self.assertNotIn(".Save(", source)


if __name__ == "__main__":
    unittest.main()
