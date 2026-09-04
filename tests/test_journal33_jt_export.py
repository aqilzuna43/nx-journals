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

    def setUp(self):
        config_folder = tempfile.TemporaryDirectory()
        self.addCleanup(config_folder.cleanup)
        self.config_path = Path(config_folder.name) / "tessUG.config"
        self.config_path.write_text("# test JT config\n", encoding="utf-8")
        resolver = mock.patch.object(
            self.journal,
            "resolve_jt_config_file",
            return_value=str(self.config_path),
        )
        wait_time = mock.patch.object(
            self.journal,
            "jt_output_wait_seconds",
            return_value=0.0,
        )
        resolver.start()
        wait_time.start()
        self.addCleanup(resolver.stop)
        self.addCleanup(wait_time.stop)

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
        self.journal.NXOpen.ListCreator = types.SimpleNamespace(
            TessellationOption=types.SimpleNamespace(Defined="DEFINED")
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
        appended = []
        builder = types.SimpleNamespace(
            NewLevel=lambda: types.SimpleNamespace(),
            LodList=types.SimpleNamespace(Append=appended.append),
        )

        self.journal.configure_jt_builder(
            builder,
            "part.jt",
            "tessUG.config",
        )

        self.assertEqual("tessUG.config", builder.ConfigFile)
        self.assertEqual("part.jt", builder.OutputJtFile)
        self.assertEqual("MONOLITHIC", builder.JtfileStructure)
        self.assertEqual("ALL", builder.JtWrite)
        self.assertTrue(builder.JtParts)
        self.assertTrue(builder.AsmStructure)
        self.assertTrue(builder.PreciseGeom)
        self.assertTrue(builder.AutolowLod)
        self.assertEqual("DEFAULT", builder.UseRefset)
        self.assertEqual("PART_AND_ASM", builder.IncludePmi)
        self.assertTrue(builder.ApplyPmi)
        self.assertEqual(1, len(appended))
        level = appended[0]
        self.assertEqual("DEFINED", level.TessOption)
        self.assertEqual(0.001, level.Chordal)
        self.assertEqual(20.0, level.Angular)

    def test_export_creates_one_versioned_jt_and_destroys_builder(self):
        self.install_jt_enums()
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)

        class Builder:
            def __init__(self):
                self.destroyed = False
                self.LodList = types.SimpleNamespace(Append=lambda level: None)

            def NewLevel(self):
                return types.SimpleNamespace()

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
            NewLevel=lambda: types.SimpleNamespace(),
            LodList=types.SimpleNamespace(Append=lambda level: None),
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
        self.assertIn(str(self.config_path), result["message"])
        builder.Destroy.assert_called_once()

    def test_builder_stays_alive_until_delayed_output_is_observed(self):
        self.install_jt_enums()
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)

        class Builder:
            def __init__(self):
                self.destroyed = False
                self.LodList = types.SimpleNamespace(Append=lambda level: None)

            def NewLevel(self):
                return types.SimpleNamespace()

            def Commit(self):
                return None

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

        def complete_output(path, _timeout):
            self.assertFalse(builder.destroyed)
            Path(path).write_bytes(b"JT")
            return 2, 0.25

        with mock.patch.object(
            self.journal,
            "wait_for_nonzero_file",
            side_effect=complete_output,
        ):
            result = self.journal.export_jt_from_part(
                session,
                object(),
                folder.name,
                "P1",
                "A",
                "",
            )

        self.assertEqual("SUCCESS", result["result"])
        self.assertTrue(builder.destroyed)

    def test_failed_builder_validation_blocks_commit(self):
        self.install_jt_enums()
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        builder = types.SimpleNamespace(
            NewLevel=lambda: types.SimpleNamespace(),
            LodList=types.SimpleNamespace(Append=lambda level: None),
            Validate=mock.Mock(return_value=False),
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

        self.assertEqual("FAILED_BUILDER_VALIDATION", result["result"])
        builder.Commit.assert_not_called()
        builder.Destroy.assert_called_once()

    def test_unimplemented_validate_does_not_block_commit(self):
        # NX 2506 JtCreator.Validate is declared but not implemented and
        # raises NXException ("Not yet implemented") when called.
        self.install_jt_enums()
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)

        class Builder:
            def __init__(self):
                self.destroyed = False
                self.committed = False
                self.LodList = types.SimpleNamespace(Append=lambda level: None)

            def NewLevel(self):
                return types.SimpleNamespace()

            def Validate(self):
                raise RuntimeError("Not yet implemented")

            def Commit(self):
                self.committed = True
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

        result = self.journal.export_jt_from_part(
            session,
            object(),
            folder.name,
            "264MN020016A01",
            "A",
            "2",
        )

        self.assertEqual("SUCCESS", result["result"])
        self.assertTrue(builder.committed)
        self.assertTrue(builder.destroyed)

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


class Journal33JtConfigResolutionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.journal = load_journal()

    def test_explicit_config_override_wins(self):
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        config_path = Path(folder.name) / "custom.config"
        config_path.write_text("# custom\n", encoding="utf-8")

        with mock.patch.dict(
            self.journal.os.environ,
            {"NX_JT_CONFIG_FILE": str(config_path)},
            clear=True,
        ):
            resolved = self.journal.resolve_jt_config_file()

        self.assertEqual(str(config_path.resolve()), resolved)

    def test_ugii_root_nxbin_finds_sibling_pvtrans_config(self):
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        install = Path(folder.name) / "NX2506"
        nxbin = install / "NXBIN"
        pvtrans = install / "PVTRANS"
        nxbin.mkdir(parents=True)
        pvtrans.mkdir()
        config_path = pvtrans / "tessUG.config"
        config_path.write_text("# installed\n", encoding="utf-8")

        with mock.patch.dict(
            self.journal.os.environ,
            {"UGII_ROOT_DIR": str(nxbin)},
            clear=True,
        ):
            resolved = self.journal.resolve_jt_config_file()

        self.assertEqual(str(config_path.resolve()), resolved)

    def test_missing_explicit_config_is_reported(self):
        missing = str(ROOT / "does-not-exist" / "tessUG.config")
        with mock.patch.dict(
            self.journal.os.environ,
            {"NX_JT_CONFIG_FILE": missing},
            clear=True,
        ):
            with self.assertRaisesRegex(RuntimeError, "points to a missing"):
                self.journal.resolve_jt_config_file()


if __name__ == "__main__":
    unittest.main()
