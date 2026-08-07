import datetime
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
    nxopen = sys.modules.setdefault("NXOpen", types.ModuleType("NXOpen"))
    annotations = sys.modules.setdefault(
        "NXOpen.Annotations", types.ModuleType("NXOpen.Annotations")
    )
    nxopen.Annotations = annotations
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

    def test_annotations_submodule_is_explicitly_imported(self):
        source_lines = JOURNAL.read_text(encoding="utf-8").splitlines()
        self.assertIn("import NXOpen.Annotations", source_lines)

    def test_native_watermark_has_no_custom_raster_path(self):
        source = JOURNAL.read_text(encoding="utf-8")
        for forbidden in (
            "import NXOpen.UF",
            "CreateImageFromFile",
            "create_watermark_png",
            "RasterImages = True",
            ".j07-watermark-",
        ):
            self.assertNotIn(forbidden, source)

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

    def run_pdf_export(
        self,
        drawing_part,
        sheets,
        output_path="combined.pdf",
        watermark="DRAFT_A.2",
        timestamp_text="EXPORTED: 2026-07-31 00:13 MYT",
    ):
        self.journal.NXOpen.Session = types.SimpleNamespace(
            MarkVisibility=types.SimpleNamespace(Invisible="Invisible")
        )
        session = types.SimpleNamespace(
            SetUndoMark=mock.Mock(return_value="undo-mark"),
        )
        with mock.patch.object(
            self.journal,
            "create_timestamp_notes",
        ) as note_creator, mock.patch.object(
            self.journal,
            "undo_timestamp_notes",
        ) as note_cleanup:
            result = self.journal.export_drawing_pdf(
                session,
                drawing_part,
                sheets,
                output_path,
                watermark,
                timestamp_text,
            )
        return result, session, note_creator, note_cleanup

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
            "J07-NX2506-SEARCHABLE-TEXT-NATIVE-WATERMARK-V8",
        )
        self.assertTrue(
            self.journal.runtime_source_path().endswith(
                "07_datapack_pdf_step_export.py"
            )
        )

    def test_export_timestamp_uses_one_explicit_myt_datetime(self):
        run_datetime = datetime.datetime(
            2026,
            7,
            31,
            0,
            13,
            tzinfo=datetime.timezone(datetime.timedelta(hours=8)),
        )

        self.assertEqual(
            self.journal.build_export_timestamp_text(run_datetime),
            "EXPORTED: 2026-07-31 00:13 MYT",
        )

    def test_sheet_units_accept_nx2506_numeric_values(self):
        self.journal.NXOpen.Drawings = types.SimpleNamespace()

        self.assertTrue(
            self.journal.sheet_uses_inches(
                types.SimpleNamespace(Units=1)
            )
        )
        self.assertFalse(
            self.journal.sheet_uses_inches(
                types.SimpleNamespace(Units=2)
            )
        )

    def test_sheet_units_accept_both_nx_sheet_enum_class_names(self):
        self.journal.NXOpen.Drawings = types.SimpleNamespace(
            DrawingSheet=types.SimpleNamespace(
                Unit=types.SimpleNamespace(
                    UnitInches="legacy-inches",
                    UnitMillimeters="legacy-millimeters",
                )
            ),
            DraftingDrawingSheet=types.SimpleNamespace(
                Unit=types.SimpleNamespace(
                    Inches="drafting-inches",
                    Millimeters="drafting-millimeters",
                )
            ),
        )

        self.assertTrue(
            self.journal.sheet_uses_inches(
                types.SimpleNamespace(Units="legacy-inches")
            )
        )
        self.assertFalse(
            self.journal.sheet_uses_inches(
                types.SimpleNamespace(Units="drafting-millimeters")
            )
        )

    def test_sheet_units_still_reject_unknown_numeric_value(self):
        self.journal.NXOpen.Drawings = types.SimpleNamespace()

        with self.assertRaisesRegex(
            RuntimeError,
            "Unsupported or unavailable drawing-sheet units: 3",
        ):
            self.journal.sheet_uses_inches(
                types.SimpleNamespace(Units=3)
            )

    def test_timestamp_note_placement_handles_metric_and_inch_sheets(self):
        class TextBlock:
            def __init__(self):
                self.value = None

            def SetText(self, value):
                self.value = value

        class NoteBuilder:
            def __init__(self):
                self.Origin = types.SimpleNamespace(
                    Anchor=None,
                    OriginPoint=None,
                )
                self.Style = types.SimpleNamespace(
                    LetteringStyle=types.SimpleNamespace(
                        GeneralTextSize=None,
                        GeneralTextLineWidth=None,
                        HorizontalTextJustification=None,
                    )
                )
                self.Text = types.SimpleNamespace(TextBlock=TextBlock())
                self.destroyed = False

            def Commit(self):
                return object()

            def Destroy(self):
                self.destroyed = True

        metric_builder = NoteBuilder()
        inch_builder = NoteBuilder()
        builders = iter((metric_builder, inch_builder))
        drawing = types.SimpleNamespace(
            Annotations=types.SimpleNamespace(
                CreateDraftingNoteBuilder=lambda _note: next(builders)
            )
        )
        metric_sheet = types.SimpleNamespace(
            Units="Millimeters",
            Length=1000.0,
            Height=500.0,
            Open=mock.Mock(),
        )
        inch_sheet = types.SimpleNamespace(
            Units="Inches",
            Length=10.0,
            Height=8.0,
            Open=mock.Mock(),
        )
        self.journal.NXOpen.Drawings = types.SimpleNamespace(
            DrawingSheet=types.SimpleNamespace(
                Unit=types.SimpleNamespace(
                    Inches="Inches",
                    Millimeters="Millimeters",
                )
            )
        )
        self.journal.NXOpen.Annotations = types.SimpleNamespace(
            OriginBuilder=types.SimpleNamespace(
                AlignmentPosition=types.SimpleNamespace(
                    BottomRight="BottomRight"
                )
            ),
            LineWidth=types.SimpleNamespace(Normal="Normal"),
            TextJustification=types.SimpleNamespace(Right="Right"),
        )
        self.journal.NXOpen.Point3d = lambda x, y, z: types.SimpleNamespace(
            X=x,
            Y=y,
            Z=z,
        )

        self.journal.create_pdf_timestamp_note(
            drawing,
            metric_sheet,
            "EXPORTED: 2026-07-31 00:13 MYT",
        )
        self.journal.create_pdf_timestamp_note(
            drawing,
            inch_sheet,
            "EXPORTED: 2026-07-31 00:13 MYT",
        )

        self.assertAlmostEqual(metric_builder.Origin.OriginPoint.X, 992.0)
        self.assertAlmostEqual(metric_builder.Origin.OriginPoint.Y, 5.0)
        self.assertAlmostEqual(
            metric_builder.Style.LetteringStyle.GeneralTextSize,
            2.5,
        )
        self.assertAlmostEqual(
            inch_builder.Origin.OriginPoint.X,
            10.0 - (8.0 / 25.4),
        )
        self.assertAlmostEqual(
            inch_builder.Origin.OriginPoint.Y,
            5.0 / 25.4,
        )
        self.assertAlmostEqual(
            inch_builder.Style.LetteringStyle.GeneralTextSize,
            2.5 / 25.4,
        )
        self.assertEqual(
            metric_builder.Origin.Anchor,
            "BottomRight",
        )
        self.assertEqual(
            metric_builder.Style.LetteringStyle.GeneralTextLineWidth,
            "Normal",
        )
        self.assertEqual(
            metric_builder.Style.LetteringStyle.HorizontalTextJustification,
            "Right",
        )
        self.assertTrue(metric_builder.destroyed)
        self.assertTrue(inch_builder.destroyed)

    def test_timestamp_notes_are_created_per_sheet_with_one_update(self):
        sheets = [object(), object(), object()]
        update_manager = types.SimpleNamespace(
            DoUpdate=mock.Mock(return_value=0)
        )
        session = types.SimpleNamespace(UpdateManager=update_manager)

        with mock.patch.object(
            self.journal,
            "create_pdf_timestamp_note",
            side_effect=[object(), object(), object()],
        ) as creator:
            notes = self.journal.create_timestamp_notes(
                session,
                object(),
                sheets,
                "EXPORTED: 2026-07-31 00:13 MYT",
                "undo-mark",
            )

        self.assertEqual(len(notes), 3)
        self.assertEqual(creator.call_count, 3)
        update_manager.DoUpdate.assert_called_once_with("undo-mark")

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
            OutputTextOption=types.SimpleNamespace(Text="Text"),
        )

        metrics, session, note_creator, note_cleanup = self.run_pdf_export(
            drawing_part, sheets
        )

        self.assertIs(builder.SourceBuilder.sheets, sheets)
        self.assertEqual(builder.Filename, "combined.pdf")
        self.assertFalse(builder.Append)
        self.assertTrue(builder.AddWatermark)
        self.assertEqual(builder.Watermark, "DRAFT_A.2")
        self.assertFalse(hasattr(builder, "RasterImages"))
        self.assertTrue(builder.CustomSymbolsInForeground)
        self.assertEqual(builder.OutputText, "Text")
        self.assertEqual(builder.commit_count, 1)
        self.assertEqual(builder.destroy_count, 1)
        self.assertEqual(sheets[0].open_count, 1)
        self.assertEqual(sheets[1].open_count, 0)
        self.assertEqual(sheets[2].open_count, 0)
        self.assertEqual(metrics["pdf_commit_seconds"] >= 0.0, True)
        note_creator.assert_called_once()
        note_cleanup.assert_called_once()
        session.SetUndoMark.assert_called_once()

    def test_pdf_export_fails_when_nx_rejects_native_watermark_text(self):
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
                    raise RuntimeError("watermark text unavailable")
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
            OutputTextOption=types.SimpleNamespace(Text="Text"),
        )

        with self.assertRaisesRegex(
            RuntimeError,
            "could not apply required native watermark DRAFT_A.2",
        ):
            self.run_pdf_export(
                drawing_part,
                [Sheet()],
            )

        self.assertEqual(builder.destroy_count, 1)

    def test_pdf_export_fails_when_nx_cannot_enable_native_watermark(self):
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
            OutputTextOption=types.SimpleNamespace(Text="Text"),
        )

        with self.assertRaisesRegex(
            RuntimeError,
            "could not apply required native watermark DRAFT_A.2",
        ):
            self.run_pdf_export(
                drawing_part,
                [Sheet()],
            )

        self.assertEqual(builder.destroy_count, 1)

    def test_pdf_export_fails_when_nx_rejects_searchable_text_output(self):
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
                    raise RuntimeError("text output unavailable")
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
            OutputTextOption=types.SimpleNamespace(Text="Text"),
        )

        with self.assertRaisesRegex(
            RuntimeError,
            "could not apply required searchable text output",
        ):
            self.run_pdf_export(
                drawing_part,
                [Sheet()],
            )

        self.assertEqual(builder.destroy_count, 1)

    def test_cleanup_failure_is_a_batch_halting_error(self):
        original_sheet = types.SimpleNamespace(Open=mock.Mock())
        session = types.SimpleNamespace(
            UndoToMark=mock.Mock(side_effect=RuntimeError("undo unavailable")),
            DeleteUndoMark=mock.Mock(),
        )
        drawing = types.SimpleNamespace(IsModified=False)

        with self.assertRaisesRegex(
            self.journal.TimestampCleanupError,
            "discard unsaved drawing changes",
        ):
            self.journal.undo_timestamp_notes(
                session,
                "undo-mark",
                "temporary timestamp",
                drawing,
                False,
                original_sheet,
            )

        original_sheet.Open.assert_called_once()
        session.DeleteUndoMark.assert_called_once()

    def test_pdf_export_reports_cleanup_failure_even_after_commit(self):
        class Builder:
            def __init__(self):
                self.SourceBuilder = types.SimpleNamespace(
                    SetSheets=lambda _sheets: None
                )

            def Commit(self):
                pass

            def Destroy(self):
                pass

        drawing = types.SimpleNamespace(
            PlotManager=types.SimpleNamespace(
                CreatePrintPdfbuilder=lambda: Builder()
            )
        )
        self.journal.NXOpen.PrintPDFBuilder = types.SimpleNamespace(
            ActionOption=types.SimpleNamespace(Native="Native"),
            OutputTextOption=types.SimpleNamespace(Text="Text"),
        )
        self.journal.NXOpen.Session = types.SimpleNamespace(
            MarkVisibility=types.SimpleNamespace(Invisible="Invisible")
        )
        session = types.SimpleNamespace(
            SetUndoMark=mock.Mock(return_value="undo-mark"),
        )
        sheet = types.SimpleNamespace(Open=mock.Mock())

        with mock.patch.object(
            self.journal,
            "create_timestamp_notes",
        ), mock.patch.object(
            self.journal,
            "undo_timestamp_notes",
            side_effect=self.journal.TimestampCleanupError(
                "cleanup could not be proven"
            ),
        ):
            with self.assertRaisesRegex(
                self.journal.TimestampCleanupError,
                "cleanup could not be proven",
            ):
                self.journal.export_drawing_pdf(
                    session,
                    drawing,
                    [sheet],
                    "combined.pdf",
                    "DRAFT_A.2",
                    "EXPORTED: 2026-07-31 00:13 MYT",
                )

    def test_pdf_builder_failure_still_undoes_timestamp_notes(self):
        class Builder:
            def __init__(self):
                self.SourceBuilder = types.SimpleNamespace(
                    SetSheets=lambda _sheets: None
                )

            def Commit(self):
                raise RuntimeError("PDF commit failed")

            def Destroy(self):
                pass

        drawing = types.SimpleNamespace(
            PlotManager=types.SimpleNamespace(
                CreatePrintPdfbuilder=lambda: Builder()
            )
        )
        self.journal.NXOpen.PrintPDFBuilder = types.SimpleNamespace(
            ActionOption=types.SimpleNamespace(Native="Native"),
            OutputTextOption=types.SimpleNamespace(Text="Text"),
        )
        self.journal.NXOpen.Session = types.SimpleNamespace(
            MarkVisibility=types.SimpleNamespace(Invisible="Invisible")
        )
        session = types.SimpleNamespace(
            SetUndoMark=mock.Mock(return_value="undo-mark"),
        )
        sheet = types.SimpleNamespace(Open=mock.Mock())

        with mock.patch.object(
            self.journal,
            "create_timestamp_notes",
        ), mock.patch.object(
            self.journal,
            "undo_timestamp_notes",
        ) as cleanup:
            with self.assertRaisesRegex(RuntimeError, "PDF commit failed"):
                self.journal.export_drawing_pdf(
                    session,
                    drawing,
                    [sheet],
                    "combined.pdf",
                    "DRAFT_A.2",
                    "EXPORTED: 2026-07-31 00:13 MYT",
                )

        cleanup.assert_called_once()

    def test_halted_pdf_batch_still_runs_independent_step_export(self):
        instruction = {
            "part_number": "264MN000001A01",
            "revision": "A",
            "pdf_requested": True,
            "step_requested": True,
            "warnings": [],
        }
        state = {
            "halted": True,
            "reason": "timestamp cleanup failed",
        }
        step_result = {
            "result": "SUCCESS",
            "path": "part.stp",
            "message": "",
            "size": 42,
        }

        with mock.patch.object(
            self.journal,
            "export_pdfs_for_instruction",
            side_effect=AssertionError("PDF export must remain halted"),
        ), mock.patch.object(
            self.journal,
            "export_step_for_instruction",
            return_value=step_result,
        ) as step_exporter, mock.patch.object(
            self.journal,
            "restore_parts",
        ):
            result = self.journal.process_instruction(
                types.SimpleNamespace(),
                instruction,
                {"pdf": "PDF", "step": "STEP"},
                "20260731_001300",
                "EXPORTED: 2026-07-31 00:13 MYT",
                state,
                object(),
                object(),
                [],
            )

        self.assertEqual(
            result["PDF_RESULT"],
            "FAILED_TIMESTAMP_CLEANUP",
        )
        self.assertEqual(result["STEP_RESULT"], "SUCCESS")
        step_exporter.assert_called_once()

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

        def create_pdf(
            _session,
            _part,
            _sheets,
            output_path,
            _watermark,
            _timestamp_text,
        ):
            Path(output_path).write_bytes(b"%PDF-test")
            return {
                "timestamp_prepare_seconds": 0.01,
                "pdf_commit_seconds": 1.0,
                "timestamp_cleanup_seconds": 0.01,
                "pdf_total_seconds": 1.02,
                "timestamp_overhead_percent": 2.0,
            }

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
                "EXPORTED: 2026-07-31 00:13 MYT",
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
        self.assertEqual(exporter.call_args.args[2], sheets)
        self.assertEqual(exporter.call_args.args[4], "DRAFT_A.2")
        self.assertEqual(
            exporter.call_args.args[5],
            "EXPORTED: 2026-07-31 00:13 MYT",
        )

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
        self.assertEqual(exporter.call_args.args[4], "DRAFT_A")

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
