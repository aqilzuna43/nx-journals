import importlib.util
import os
import sys
import tempfile
import types
import unittest
from pathlib import Path
from unittest import mock


ROOT = Path(__file__).resolve().parents[1]
JOURNAL = ROOT / "from_git" / "journals" / "25_tc_single_drawing_cleanup.py"


def load_journal():
    nxopen = types.ModuleType("NXOpen")
    nxuf = types.ModuleType("NXOpen.UF")
    nxopen.UF = nxuf
    nxopen.Session = types.SimpleNamespace(
        LibraryUnloadOption=types.SimpleNamespace(Immediately="Immediately")
    )
    nxopen.BasePart = types.SimpleNamespace(
        CloseWholeTree=types.SimpleNamespace(FalseValue="FalseValue"),
        CloseModified=types.SimpleNamespace(CloseModified="CloseModified"),
    )
    prior = sys.modules.get("NXOpen")
    prior_uf = sys.modules.get("NXOpen.UF")
    sys.modules["NXOpen"] = nxopen
    sys.modules["NXOpen.UF"] = nxuf
    try:
        spec = importlib.util.spec_from_file_location("journal25", JOURNAL)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        return module
    finally:
        if prior is None:
            sys.modules.pop("NXOpen", None)
        else:
            sys.modules["NXOpen"] = prior
        if prior_uf is None:
            sys.modules.pop("NXOpen.UF", None)
        else:
            sys.modules["NXOpen.UF"] = prior_uf


class FakeLog:
    def __init__(self):
        self.lines = []

    def write(self, value=""):
        self.lines.append(str(value))


class SingleDrawingCleanupTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.journal = load_journal()

    def row(self, **changes):
        result = {
            "_CSV_ROW": 2,
            "PART_NUMBER": "MODEL100",
            "REVISION": "A",
            "KEEP_DWG_INDEX": "1",
            "EXPECTED_REMOVE_DWG_INDICES": "2|3",
            "APPROVED": "YES",
            "ENGINEER": "AQIL",
            "CONFIRMATION": "REMOVE_EXTRA_DRAWINGS",
        }
        result.update(changes)
        return result

    def inspection(self, sheets=1, loaded=False, checkout="CHECKED_IN"):
        return {
            "state": "EXISTS",
            "checkout_state": checkout,
            "checkout_owner": "",
            "drawing_sheet_count": sheets,
            "loaded_at_start": loaded,
            "detail": "exact",
        }

    def plan(self, extras=(2, 3), loaded=()):
        drawings = {1: self.inspection(sheets=3)}
        for index in extras:
            drawings[index] = self.inspection(loaded=index in loaded)
        return {
            "part_number": "MODEL100",
            "revision": "A",
            "keep": 1,
            "expected_remove": list(extras),
            "live_remove": list(extras),
            "master": self.inspection(sheets=-1),
            "drawings": drawings,
            "discovered": [1] + list(extras),
        }

    def test_parse_index_list_accepts_auditable_separators(self):
        self.assertEqual([2, 3, 4], self.journal.parse_index_list("4|2,3"))
        with self.assertRaisesRegex(RuntimeError, "duplicate"):
            self.journal.parse_index_list("2|2")
        with self.assertRaisesRegex(RuntimeError, "1 to 9"):
            self.journal.parse_index_list("10")

    def test_validate_plan_requires_exact_live_extra_list(self):
        drawings = {
            1: self.inspection(sheets=3),
            2: self.inspection(),
            3: self.inspection(),
        }
        with mock.patch.object(
            self.journal, "inspect_exact", return_value=self.inspection(sheets=-1)
        ), mock.patch.object(
            self.journal, "discover_drawings", return_value=drawings
        ):
            plan = self.journal.validate_plan(self.row(), object(), FakeLog())
            self.assertEqual([2, 3], plan["live_remove"])
            with self.assertRaisesRegex(RuntimeError, "Live extras"):
                self.journal.validate_plan(
                    self.row(EXPECTED_REMOVE_DWG_INDICES="2"),
                    object(), FakeLog(),
                )

    def test_validate_plan_never_allows_keep_index_in_remove_list(self):
        with self.assertRaisesRegex(RuntimeError, "also listed"):
            self.journal.validate_plan(
                self.row(EXPECTED_REMOVE_DWG_INDICES="1|2"),
                object(), FakeLog(),
            )

    def test_apply_authorization_blocks_loaded_extra(self):
        with self.assertRaisesRegex(RuntimeError, "Close these extra"):
            self.journal.require_apply_authorization(
                self.row(), self.plan(loaded=(2,))
            )

    def test_dry_run_builds_plan_without_delete(self):
        with tempfile.TemporaryDirectory() as folder, mock.patch.object(
            self.journal, "validate_plan", return_value=self.plan()
        ), mock.patch.object(
            self.journal, "backup_and_delete_target"
        ) as delete:
            reports = self.journal.execute(
                [self.row()], object(), object(), "DRY_RUN", folder,
                "20260814_120000", FakeLog()
            )
        self.assertEqual("DRY_RUN_READY", reports[0]["RESULT"])
        self.assertEqual("NO", reports[0]["WRITE_ATTEMPTED"])
        delete.assert_not_called()

    def test_apply_batch_preflights_every_row_before_any_delete(self):
        second = self.row(PART_NUMBER="MODEL200", _CSV_ROW=3)
        with tempfile.TemporaryDirectory() as folder, mock.patch.object(
            self.journal, "validate_plan",
            side_effect=[self.plan(), RuntimeError("bad second row")],
        ), mock.patch.object(
            self.journal, "backup_and_delete_target"
        ) as delete:
            reports = self.journal.execute(
                [self.row(), second], object(), object(), "APPLY_APPROVED",
                folder, "20260814_120000", FakeLog(),
            )
        self.assertEqual("BLOCKED_BY_BATCH_PREFLIGHT", reports[0]["RESULT"])
        self.assertEqual("BLOCKED", reports[1]["RESULT"])
        delete.assert_not_called()

    def test_apply_deletes_only_extras_and_verifies_keep(self):
        initial = self.plan()
        final = {1: self.inspection(sheets=3)}
        backup = lambda index: {
            "identifier": self.journal.drawing_id("MODEL100", "A", index),
            "backup": [{"file": "dwg{0}.prt".format(index), "sha256": str(index) * 64}],
            "delete_result": "[0]", "delete_statuses": [0],
            "keep_empty_dataset": False,
        }
        post_absent = {
            "state": "NOT_OPENABLE", "loaded_at_start": False, "detail": "not found",
        }
        with tempfile.TemporaryDirectory() as folder, mock.patch.object(
            self.journal, "validate_plan", side_effect=[initial, initial]
        ), mock.patch.object(
            self.journal, "backup_and_delete_target",
            side_effect=lambda *args: backup(args[4]),
        ) as delete, mock.patch.object(
            self.journal, "inspect_exact", return_value=post_absent
        ), mock.patch.object(
            self.journal, "discover_drawings", return_value=final
        ):
            reports = self.journal.execute(
                [self.row()], object(), object(), "APPLY_APPROVED", folder,
                "20260814_120000", FakeLog()
            )
        report = reports[0]
        self.assertEqual("SINGLE_DWG_VERIFIED", report["RESULT"])
        self.assertEqual("2|3", report["REMOVED_DWG_INDICES"])
        self.assertEqual("1", report["POSTCHECK_DWG_INDICES"])
        self.assertEqual([2, 3], [call.args[4] for call in delete.call_args_list])

    def test_delete_api_uses_false_keep_empty_dataset(self):
        class FakePart:
            JournalIdentifier = "@DB/MODEL100/A/specification/MODEL100-A-dwg2"

        class FakePdmFile:
            def __init__(self, name):
                self.name = name

            def GetFileName(self):
                return self.name

            def FreeResource(self):
                pass

        class Parts:
            def OpenBase(self, identifier):
                return FakePart(), None

        class FileManagement:
            def __init__(self):
                self.delete_args = None

            def GetAssociatedFiles(self, parts, excluded):
                return [FakePdmFile("MODEL100_A_dwg2.prt")]

            def DownloadAssociatedFiles(self, parts, files):
                return []

            def DeleteExistingAttachedFiles(self, files, keep_empty):
                self.delete_args = (list(files), keep_empty)
                return [0]

        with tempfile.TemporaryDirectory() as folder:
            native = os.path.join(folder, "MODEL100_A_dwg2.prt")
            with open(native, "wb") as handle:
                handle.write(b"drawing")
            manager = FileManagement()
            session = types.SimpleNamespace(Parts=Parts())
            with mock.patch.object(
                self.journal.J16, "find_loaded_by_identifier", return_value=None
            ), mock.patch.object(
                self.journal.J16, "locate_downloaded_files",
                return_value={os.path.normcase(native): native},
            ), mock.patch.object(self.journal.J16, "close_opened_part"):
                outcome = self.journal.backup_and_delete_target(
                    session, manager, "MODEL100", "A", 2, folder, FakeLog()
                )
        self.assertIs(manager.delete_args[1], False)
        self.assertEqual([0], outcome["delete_statuses"])
        self.assertEqual(1, len(outcome["backup"]))

    def test_source_declares_destructive_semantics_and_postcheck(self):
        source = JOURNAL.read_text(encoding="utf-8")
        self.assertIn("DeleteExistingAttachedFiles", source)
        self.assertIn("keepEmptyDataset=False", source)
        self.assertIn("This is not a relation-only detach", source)
        self.assertIn("POSTCHECK_DWG_INDICES", source)


if __name__ == "__main__":
    unittest.main()
