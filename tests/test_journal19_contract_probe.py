import ast
import importlib.util
import json
import os
import sys
import tempfile
import types
import unittest
import zipfile
from pathlib import Path
from unittest import mock


ROOT = Path(__file__).resolve().parents[1]
J19_PATH = (
    ROOT
    / "from_git"
    / "journals"
    / "19_test_teamcenter_drawing_import_contract.py"
)


def load_j19():
    nxopen = types.ModuleType("NXOpen")
    nxuf = types.ModuleType("NXOpen.UF")
    nxopen.UF = nxuf
    nxopen.BasePart = types.SimpleNamespace(
        CloseWholeTree=types.SimpleNamespace(FalseValue="FalseValue"),
        CloseModified=types.SimpleNamespace(CloseModified="CloseModified"),
    )
    sys.modules["NXOpen"] = nxopen
    sys.modules["NXOpen.UF"] = nxuf
    spec = importlib.util.spec_from_file_location("journal19_contract", J19_PATH)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class FakeLog:
    def __init__(self, _session=None):
        self.lines = []

    def write(self, message=""):
        self.lines.append(str(message))


class FakePdmFile:
    def __init__(self, path):
        self.path = path
        self.released = False

    def GetFileName(self):
        return self.path

    def GetFileSize(self):
        return str(os.path.getsize(self.path)) if os.path.isfile(self.path) else "0"

    def GetFileLastModifiedDate(self):
        return "2026-07-31T00:00:00"

    def FreeResource(self):
        self.released = True


class Journal19ContractTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.journal = load_j19()

    def test_default_target_is_controlled_drawing(self):
        with mock.patch.dict(os.environ, {}, clear=True):
            self.assertEqual(
                ("264MN021218A01", "A", 1),
                self.journal.configured_target(),
            )

    def test_invoke_records_direct_return_shape(self):
        def method(value):
            return 0, [value]

        record, raw = self.journal.invoke(method, ("exported.prt",))

        self.assertEqual("RETURNED", record["status"])
        self.assertEqual((0, ["exported.prt"]), raw)
        self.assertEqual([0], record["pdi_codes"])
        self.assertEqual(["exported.prt"], record["returned_strings"])
        self.assertEqual(1, record["attempts"][0]["argument_count"])

    def test_invoke_retries_only_for_required_output_parameter(self):
        def method(*args):
            if len(args) == 1:
                raise TypeError("output parameter required")
            args[-1].append("download-folder")
            return 0

        output = []
        record, raw = self.journal.invoke(method, ("input",), output)

        self.assertEqual(0, raw)
        self.assertEqual("RETURNED", record["status"])
        self.assertEqual("TYPE_ERROR", record["attempts"][0]["status"])
        self.assertEqual("RETURNED", record["attempts"][1]["status"])
        self.assertEqual(["download-folder"], output)
        self.assertIn("download-folder", record["returned_strings"])

    def test_directory_snapshot_records_every_file_and_sha(self):
        with tempfile.TemporaryDirectory() as folder:
            nested = os.path.join(folder, "nested")
            os.makedirs(nested)
            path = os.path.join(nested, "drawing.prt")
            with open(path, "wb") as handle:
                handle.write(b"managed drawing")

            snapshot = self.journal.directory_snapshot(folder)

        self.assertEqual([os.path.join("nested", "drawing.prt")], [
            item["relative_path"] for item in snapshot["files"]
        ])
        self.assertEqual(64, len(snapshot["files"][0]["sha256"]))

    def test_associated_file_probe_downloads_and_hashes_native_file(self):
        with tempfile.TemporaryDirectory() as folder:
            native = os.path.join(folder, "downloaded.prt")
            pdm_file = FakePdmFile(native)

            class FileManagement:
                def GetAssociatedFiles(self, parts, excluded):
                    self.get_args = (parts, excluded)
                    return [pdm_file]

                def DownloadAssociatedFiles(self, parts, files):
                    self.download_args = (parts, files)
                    with open(native, "wb") as handle:
                        handle.write(b"downloaded payload")

            result = self.journal.probe_associated_files(
                FileManagement(), object(), os.path.join(folder, "evidence")
            )

        self.assertEqual("PASS_NATIVE_FILE_FOUND", result["status"])
        self.assertEqual(native, result["native_files"][0]["path"])
        self.assertEqual(64, len(result["native_files"][0]["sha256"]))
        self.assertTrue(pdm_file.released)

    def test_named_reference_probe_captures_physical_payload(self):
        with tempfile.TemporaryDirectory() as folder:
            class FileManagement:
                def ExportNamedReferences(self, *args):
                    path = os.path.join(args[6], "named-reference.prt")
                    with open(path, "wb") as handle:
                        handle.write(b"named reference")
                    return 0, [path]

            result = self.journal.probe_export_named_reference(
                FileManagement(), ("ITEM", "A", 1), folder
            )

        self.assertEqual("PASS_NATIVE_FILE_FOUND", result["status"])
        self.assertEqual([0], result["call"]["pdi_codes"])
        self.assertEqual(1, len(result["native_files"]))

    def test_legacy_export_zero_without_file_is_not_success(self):
        class FileManagement:
            def ExportFiles(self, *args):
                return 0, []

        with tempfile.TemporaryDirectory() as folder:
            result = self.journal.probe_legacy_export_files(
                FileManagement(), ("ITEM", "A", 1), folder
            )

        self.assertEqual("COMPLETE_NO_NATIVE_FILE", result["status"])
        self.assertEqual([0], result["call"]["pdi_codes"])
        self.assertEqual([], result["native_files"])

    def test_main_continues_all_three_probes_and_packages_evidence(self):
        identifier = self.journal.J16.drawing_id("264MN021218A01", "A", 1)
        part = types.SimpleNamespace(
            JournalIdentifier=identifier,
            FullPath=identifier,
            Leaf="264MN021218A01-A-dwg1",
            UniqueIdentifier="uid",
            IsReadOnly=True,
            HasWriteAccess=False,
            IsModified=False,
            DrawingSheets=[object()],
        )

        class Parts:
            def __iter__(self):
                return iter([])

            def OpenBase(self, _identifier):
                return part, None

        session = types.SimpleNamespace(Parts=Parts(), PdmSession=object())
        self.journal.NXOpen.Session = types.SimpleNamespace(
            GetSession=mock.Mock(return_value=session)
        )

        with tempfile.TemporaryDirectory() as folder:
            root = os.path.join(folder, "contract")
            paths = (
                root,
                os.path.join(root, "report.json"),
                os.path.join(root, "report.log"),
            )
            statuses = [
                {"status": "ERROR"},
                {"status": "PASS_NATIVE_FILE_FOUND"},
                {"status": "COMPLETE_NO_NATIVE_FILE"},
            ]
            with mock.patch.object(
                self.journal, "output_paths", return_value=paths
            ), mock.patch.object(
                self.journal.J16, "Log", FakeLog
            ), mock.patch.object(
                self.journal.J16,
                "query_pdm_checkout",
                return_value={"state": "CHECKED_IN", "owner": "", "raw": "(False, '')"},
            ), mock.patch.object(
                self.journal.J16,
                "new_file_management",
                return_value=(object(), object()),
            ), mock.patch.object(
                self.journal.J16, "close_opened_part"
            ), mock.patch.object(
                self.journal,
                "probe_associated_files",
                return_value=statuses[0],
            ) as associated, mock.patch.object(
                self.journal,
                "probe_export_named_reference",
                return_value=statuses[1],
            ) as named, mock.patch.object(
                self.journal,
                "probe_legacy_export_files",
                return_value=statuses[2],
            ) as legacy:
                zip_path = self.journal.main()

            report = json.loads(Path(paths[1]).read_text(encoding="utf-8"))
            self.assertEqual("PROBE_COMPLETE_RETRIEVAL_FOUND", report["result"])
            self.assertFalse(report["teamcenter_write_attempted"])
            associated.assert_called_once()
            named.assert_called_once()
            legacy.assert_called_once()
            self.assertTrue(os.path.isfile(zip_path))
            with zipfile.ZipFile(zip_path) as archive:
                names = archive.namelist()
            self.assertTrue(any(name.endswith("report.json") for name in names))
            self.assertTrue(any(name.endswith("report.log") for name in names))

    def test_source_contains_no_teamcenter_mutation_or_clone_calls(self):
        source = J19_PATH.read_text(encoding="utf-8")
        tree = ast.parse(source)
        forbidden = {
            "Checkout",
            "CheckoutParts",
            "CheckinParts",
            "Save",
            "SaveAs",
            "SaveAll",
            "ImportFiles",
            "ImportFilesAndCreateDatasets",
            "PerformClone",
            "SetDryrun",
        }
        called = set()
        for node in ast.walk(tree):
            if isinstance(node, ast.Call) and isinstance(node.func, ast.Attribute):
                called.add(node.func.attr)
        self.assertEqual(set(), forbidden.intersection(called))
        self.assertNotIn("J16.import_one", source)
        self.assertNotIn("NXOpen.UF", source)
        self.assertIn("DownloadAssociatedFiles", source)
        self.assertIn("ExportNamedReferences", source)
        self.assertIn("ExportFiles", source)


if __name__ == "__main__":
    unittest.main()
