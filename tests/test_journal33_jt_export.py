import importlib.util
import sys
import tempfile
import types
import unittest
from pathlib import Path
from unittest import mock


ROOT = Path(__file__).resolve().parents[1]
JOURNAL = ROOT / "from_git" / "journals" / "33_datapack_jt_export.py"


def load_journal():
    nxopen = types.ModuleType("NXOpen")
    spec = importlib.util.spec_from_file_location("journal33", JOURNAL)
    module = importlib.util.module_from_spec(spec)
    with mock.patch.dict(sys.modules, {"NXOpen": nxopen}):
        spec.loader.exec_module(module)
    return module


class Journal33JtExportTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.journal = load_journal()

    def write_scope(self, content):
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        path = Path(folder.name) / "NX_EXPORT_SCOPE.csv"
        path.write_text(content, encoding="utf-8")
        return path

    def install_jt_enums(self):
        self.journal.NXOpen.JtCreator = types.SimpleNamespace(
            FileStructure=types.SimpleNamespace(Monolithic="MONOLITHIC"),
            FileWrite=types.SimpleNamespace(All="ALL"),
            TessellationOption=types.SimpleNamespace(Nx="NX"),
            RefsetOption=types.SimpleNamespace(Default="DEFAULT"),
            PmiOption=types.SimpleNamespace(PartAndAsm="PART_AND_ASM"),
        )

    def test_scope_requires_jt_but_accepts_j07_columns(self):
        headers, warnings = self.journal.resolve_headers(
            ["DB_PART_NO", "DB_PART_REV", "PDF", "STEP", "JT"]
        )
        self.assertEqual("DB_PART_NO", headers["part_number"])
        self.assertEqual("DB_PART_REV", headers["revision"])
        self.assertEqual("JT", headers["jt"])
        self.assertEqual([], warnings)

        with self.assertRaisesRegex(ValueError, "jt"):
            self.journal.resolve_headers(
                ["DB_PART_NO", "DB_PART_REV", "PDF", "STEP"]
            )

    def test_scope_merges_duplicate_jt_requests_and_ignores_disabled_rows(self):
        path = self.write_scope(
            "DB_PART_NO,DB_PART_REV,PDF,STEP,JT,PART_DESCRIPTION\n"
            "P2,B,NO,YES,NO,ignored\n"
            "P1,A,YES,NO,YES,\n"
            "p1,a,NO,NO,X,Bracket\n"
        )

        parsed = self.journal.read_export_scope(path)

        self.assertEqual(3, parsed["input_row_count"])
        self.assertEqual(1, parsed["ignored_row_count"])
        self.assertEqual(1, len(parsed["instructions"]))
        instruction = parsed["instructions"][0]
        self.assertEqual("P1", instruction["part_number"])
        self.assertEqual(2, instruction["source_row_count"])
        self.assertEqual(1, instruction["merged_row_count"])
        self.assertEqual("Bracket", instruction["part_description"])

    def test_repository_template_drives_j33_without_breaking_j07_columns(self):
        template = (
            ROOT / "from_git" / "templates" / "NX_EXPORT_SCOPE_TEMPLATE.csv"
        )

        parsed = self.journal.read_export_scope(template)

        self.assertEqual(1, len(parsed["instructions"]))
        self.assertEqual(1, parsed["ignored_row_count"])
        self.assertTrue(parsed["instructions"][0]["jt_requested"])
        header = template.read_text(encoding="utf-8-sig").splitlines()[0]
        self.assertIn("PDF", header.split(","))
        self.assertIn("STEP", header.split(","))
        self.assertIn("JT", header.split(","))

    def test_unknown_jt_control_is_reported_as_invalid(self):
        path = self.write_scope(
            "DB_PART_NO,DB_PART_REV,JT\nP1,A,MAYBE\n"
        )

        parsed = self.journal.read_export_scope(path)

        self.assertEqual([], parsed["instructions"])
        self.assertEqual(1, len(parsed["invalid_rows"]))
        result = self.journal.invalid_result(
            "20260904_120000",
            parsed["invalid_rows"][0],
        )
        self.assertEqual("INVALID_INPUT", result["OVERALL_RESULT"])
        self.assertIn("unknown JT control value", result["MESSAGE"])

    def test_versioned_jt_name_matches_j07_convention(self):
        self.assertEqual(
            "264MN020016A01_REVA.2",
            self.journal.build_versioned_base("264MN020016A01", "A", "2"),
        )
        self.assertEqual(
            "264MN020016A01_REVA",
            self.journal.build_versioned_base("264MN020016A01", "A", ""),
        )

    def test_jt_builder_contract_is_explicit(self):
        self.install_jt_enums()
        builder = types.SimpleNamespace()

        self.journal.configure_jt_builder(builder, "part.jt")

        self.assertEqual("part.jt", builder.OutputJtFile)
        self.assertEqual("MONOLITHIC", builder.JtfileStructure)
        self.assertEqual("ALL", builder.JtWrite)
        self.assertTrue(builder.JtParts)
        self.assertTrue(builder.AsmStructure)
        self.assertTrue(builder.PreciseGeom)
        self.assertEqual("NX", builder.TessOption)
        self.assertEqual("DEFAULT", builder.UseRefset)
        self.assertEqual("PART_AND_ASM", builder.IncludePmi)
        self.assertTrue(builder.ApplyPmi)
        self.assertFalse(builder.MergeSolids)
        self.assertFalse(builder.WireFrame)

    def test_export_creates_one_versioned_jt_and_destroys_builder(self):
        self.install_jt_enums()
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)

        class Builder:
            def __init__(self):
                self.destroyed = False

            def Commit(self):
                Path(self.OutputJtFile).write_bytes(b"JT-test")

            def Destroy(self):
                self.destroyed = True

        builder = Builder()
        session = types.SimpleNamespace(
            Parts=types.SimpleNamespace(
                SetDisplay=mock.Mock(return_value=None),
                SetWork=mock.Mock(),
            ),
            PvtransManager=types.SimpleNamespace(
                CreateJtCreator=mock.Mock(return_value=builder)
            ),
        )
        part = object()

        result = self.journal.export_jt_from_part(
            session,
            part,
            folder.name,
            "264MN020016A01",
            "A",
            "2",
        )

        self.assertEqual("SUCCESS", result["result"])
        self.assertEqual("264MN020016A01_REVA.2.jt", Path(result["path"]).name)
        self.assertEqual(7, result["size"])
        self.assertTrue(builder.destroyed)
        session.Parts.SetDisplay.assert_called_once_with(part, False, True)
        session.Parts.SetWork.assert_called_once_with(part)

    def test_missing_output_after_commit_is_rejected(self):
        self.install_jt_enums()
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        builder = types.SimpleNamespace(
            Commit=mock.Mock(),
            Destroy=mock.Mock(),
        )
        session = types.SimpleNamespace(
            Parts=types.SimpleNamespace(
                SetDisplay=mock.Mock(return_value=None),
                SetWork=mock.Mock(),
            ),
            PvtransManager=types.SimpleNamespace(
                CreateJtCreator=mock.Mock(return_value=builder)
            ),
        )

        result = self.journal.export_jt_from_part(
            session,
            object(),
            folder.name,
            "P1",
            "A",
            "",
        )

        self.assertEqual("FAILED_NO_OUTPUT_FILE", result["result"])
        builder.Destroy.assert_called_once()

    def test_exact_identity_mismatch_blocks_builder_commit(self):
        part = types.SimpleNamespace(
            GetStringAttribute=lambda name: {
                "DB_PART_NO": "P1",
                "DB_PART_REV": "B",
            }.get(name, "")
        )
        candidate = {
            "part": part,
            "opened_by_journal": False,
            "source": "loaded session",
        }

        with mock.patch.object(
            self.journal,
            "resolve_master_candidate",
            return_value=(candidate, []),
        ), mock.patch.object(
            self.journal,
            "export_jt_from_part",
        ) as exporter:
            result = self.journal.export_jt_for_instruction(
                types.SimpleNamespace(),
                "JT",
                "P1",
                "A",
                object(),
                object(),
                [],
            )

        self.assertEqual("REVISION_MISMATCH", result["result"])
        exporter.assert_not_called()

    def test_source_contains_no_checkout_save_or_dataset_creation(self):
        source = JOURNAL.read_text(encoding="utf-8")
        for forbidden in (
            "CheckoutParts",
            "CheckinParts",
            ".Save(",
            "CreateSpecification",
            "CreateDataset",
        ):
            self.assertNotIn(forbidden, source)


if __name__ == "__main__":
    unittest.main()
