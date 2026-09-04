import importlib.util
import json
import pathlib
import sys
import tempfile
import types
import unittest
from unittest import mock


ROOT = pathlib.Path(__file__).resolve().parents[1]
PACKAGE = ROOT / "from_git" / "admin_freeze"
COMMON_PATH = PACKAGE / "admin_freeze_common.py"
J34_PATH = PACKAGE / "34_validate_freeze_csv.py"
J35_PATH = PACKAGE / "35_apply_freeze_csv.py"
TEMPLATE_PATH = PACKAGE / "NX_ADMIN_FREEZE_SCOPE.csv"


def load_common():
    nxopen = types.ModuleType("NXOpen")
    nxopen_pdm = types.ModuleType("NXOpen.PDM")
    nxopen.PDM = nxopen_pdm
    spec = importlib.util.spec_from_file_location("admin_freeze_test", COMMON_PATH)
    module = importlib.util.module_from_spec(spec)
    with mock.patch.dict(sys.modules, {"NXOpen": nxopen, "NXOpen.PDM": nxopen_pdm}):
        spec.loader.exec_module(module)
    return module


def planned(common, number="P1", revision="A", csv_wae=""):
    return common.base_result([2], "YES", number, revision, csv_wae)


def snap(number="P1", revision="A", wae="1", state="CHECKED_IN",
         owner="", statuses=None, read_only=True, modifiable=False):
    values = list(statuses or [])
    return {
        "part_number": number,
        "revision": revision,
        "wae_version": wae,
        "checkout": {"state": state, "owner": owner, "raw": ""},
        "release_status": {
            "display": values[0] if values else "",
            "internal": values[1:] if len(values) > 1 else [],
            "errors": [],
        },
        "read_only": read_only,
        "pdm_modifiable": modifiable,
        "pdm_modifiable_error": "",
        "modified": False,
    }


class TestAdminFreeze(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.common = load_common()

    def test_package_and_two_button_contract(self):
        self.assertTrue(J34_PATH.exists())
        self.assertTrue(J35_PATH.exists())
        self.assertTrue(TEMPLATE_PATH.exists())
        self.assertEqual("NX-ADMIN-FREEZE-V1", self.common.COMMON_BUILD)
        self.assertIn('run_ui("VALIDATE"', J34_PATH.read_text(encoding="utf-8"))
        self.assertIn('run_ui("APPLY"', J35_PATH.read_text(encoding="utf-8"))
        self.assertIn(
            'EXPECTED_COMMON_BUILD = "NX-ADMIN-FREEZE-V1"',
            J34_PATH.read_text(encoding="utf-8"),
        )

    def test_wae_classification(self):
        self.assertEqual(("NUMERIC_WORKING", ""), self.common.classify_wae("12", "A"))
        self.assertEqual(("ALPHABETIC_FINAL", ""), self.common.classify_wae("a", "A"))
        for value, revision in (("", "A"), ("0", "A"), ("B", "A"), ("1.2", "A")):
            with self.subTest(value=value):
                kind, error = self.common.classify_wae(value, revision)
                self.assertEqual("", kind)
                self.assertTrue(error)

    def test_duplicate_identity_collapses_once(self):
        rows = [
            {"csv_row": 2, "freeze": "YES", "part_number": "P1",
             "revision": "A", "csv_wae_version": "1"},
            {"csv_row": 8, "freeze": "Y", "part_number": "p1",
             "revision": "a", "csv_wae_version": "1"},
        ]
        results = self.common.plan_scope(rows)
        self.assertEqual(1, len(results))
        self.assertEqual([2, 8], results[0]["csv_rows"])
        self.assertEqual("PENDING", results[0]["result"])

    def test_conflicting_duplicate_wae_blocks_identity(self):
        rows = [
            {"csv_row": 2, "freeze": "YES", "part_number": "P1",
             "revision": "A", "csv_wae_version": "1"},
            {"csv_row": 3, "freeze": "YES", "part_number": "P1",
             "revision": "A", "csv_wae_version": "2"},
        ]
        result = self.common.plan_scope(rows)[0]
        self.assertEqual("BLOCKED_CONFLICTING_CSV_WAE", result["result"])

    def test_disabled_duplicates_report_once(self):
        rows = [
            {"csv_row": 2, "freeze": "NO", "part_number": "P1",
             "revision": "A", "csv_wae_version": ""},
            {"csv_row": 3, "freeze": "", "part_number": "P1",
             "revision": "A", "csv_wae_version": ""},
        ]
        results = self.common.plan_scope(rows)
        self.assertEqual(1, len(results))
        self.assertEqual("SKIPPED_DISABLED", results[0]["result"])

    def test_read_scope_accepts_minimal_csv_without_wae_column(self):
        with tempfile.TemporaryDirectory() as folder:
            path = pathlib.Path(folder) / "scope.csv"
            path.write_text("FREEZE,DB_PART_NO,DB_PART_REV\nYES,P1,A\n", encoding="utf-8")
            rows = self.common.read_scope(str(path))
        self.assertEqual("", rows[0]["csv_wae_version"])

    def test_validation_accepts_matching_alphabetic_final(self):
        result = planned(self.common, revision="E", csv_wae="E")
        opened = {"part": object(), "opened_by_journal": False, "source": "@DB/P1/E"}
        with mock.patch.object(self.common, "open_exact_part", return_value=opened), \
             mock.patch.object(self.common, "part_snapshot", return_value=snap(revision="E", wae="E")), \
             mock.patch.object(self.common, "get_workflows", return_value=["Part_Freeze_Process"]):
            self.common.validate_one(object(), result)
        self.assertEqual("READY", result["result"])
        self.assertEqual("ALPHABETIC_FINAL", result["wae_class"])

    def test_validation_blocks_stale_csv_wae(self):
        result = planned(self.common, csv_wae="2")
        opened = {"part": object(), "opened_by_journal": False, "source": "@DB/P1/A"}
        with mock.patch.object(self.common, "open_exact_part", return_value=opened), \
             mock.patch.object(self.common, "part_snapshot", return_value=snap(wae="3")):
            self.common.validate_one(object(), result)
        self.assertEqual("BLOCKED_STALE_CSV_WAE", result["result"])

    def test_validation_skips_checked_out_target_and_reports_owner(self):
        result = planned(self.common)
        opened = {"part": object(), "opened_by_journal": False, "source": "@DB/P1/A"}
        with mock.patch.object(self.common, "open_exact_part", return_value=opened), \
             mock.patch.object(
                 self.common, "part_snapshot",
                 return_value=snap(state="CHECKED_OUT", owner="Aqil", read_only=False, modifiable=True),
             ):
            self.common.validate_one(object(), result)
        self.assertEqual("BLOCKED_CHECKED_OUT", result["result"])
        self.assertIn("Aqil", result["message"])

    def test_validation_skips_other_release_status(self):
        result = planned(self.common)
        opened = {"part": object(), "opened_by_journal": False, "source": "@DB/P1/A"}
        with mock.patch.object(self.common, "open_exact_part", return_value=opened), \
             mock.patch.object(self.common, "part_snapshot", return_value=snap(statuses=["Released"])):
            self.common.validate_one(object(), result)
        self.assertEqual("BLOCKED_OTHER_RELEASE_STATUS", result["result"])

    def test_apply_error_with_verified_frozen_state_is_warning_success(self):
        validated = planned(self.common)
        validated.update({"result": "READY", "actual_wae_version": "1"})
        result = self.common.clone_manifest_result(validated)
        opened = {"part": object(), "opened_by_journal": False, "source": "@DB/P1/A"}
        after = snap(statuses=["Frozen", "Cad0Frozen"])
        with mock.patch.object(self.common, "open_exact_part", return_value=opened), \
             mock.patch.object(self.common, "part_snapshot", side_effect=[snap(), after]), \
             mock.patch.object(self.common, "get_workflows", return_value=["Part_Freeze_Process"]), \
             mock.patch.object(self.common, "assign_freeze", side_effect=RuntimeError("TC [3520110]")):
            self.common.apply_one(object(), result, validated)
        self.assertEqual("FROZEN_WITH_WARNING", result["result"])
        self.assertIn("3520110", result["workflow_error"])

    def test_apply_error_without_frozen_state_fails_but_does_not_raise(self):
        validated = planned(self.common)
        validated.update({"result": "READY", "actual_wae_version": "1"})
        result = self.common.clone_manifest_result(validated)
        opened = {"part": object(), "opened_by_journal": False, "source": "@DB/P1/A"}
        with mock.patch.object(self.common, "open_exact_part", return_value=opened), \
             mock.patch.object(self.common, "part_snapshot", side_effect=[snap(), snap()]), \
             mock.patch.object(self.common, "get_workflows", return_value=["Part_Freeze_Process"]), \
             mock.patch.object(self.common, "assign_freeze", side_effect=RuntimeError("TC failed")):
            self.common.apply_one(object(), result, validated)
        self.assertEqual("FAILED_FREEZE_WORKFLOW", result["result"])

    def test_apply_requires_exact_validated_csv_hash(self):
        with tempfile.TemporaryDirectory() as folder:
            root = pathlib.Path(folder)
            csv_path = root / self.common.INPUT_FILENAME
            csv_path.write_text("FREEZE,DB_PART_NO,DB_PART_REV\nYES,P1,A\n", encoding="utf-8")
            manifest = {
                "build": self.common.COMMON_BUILD,
                "mode": "VALIDATE",
                "input_sha256": "not-the-current-hash",
                "results": [],
            }
            (root / self.common.MANIFEST_FILENAME).write_text(
                json.dumps(manifest), encoding="utf-8"
            )
            with mock.patch.object(self.common, "package_root", return_value=str(root)):
                with self.assertRaisesRegex(RuntimeError, "changed after J34"):
                    self.common.run_apply(object())

    def test_apply_continues_across_independent_results(self):
        with tempfile.TemporaryDirectory() as folder:
            root = pathlib.Path(folder)
            csv_path = root / self.common.INPUT_FILENAME
            csv_path.write_text("FREEZE,DB_PART_NO,DB_PART_REV\nYES,P1,A\nYES,P2,A\n", encoding="utf-8")
            manifest_results = []
            for number in ("P1", "P2"):
                item = planned(self.common, number=number)
                item.update({"result": "READY", "actual_wae_version": "1"})
                manifest_results.append(item)
            manifest = {
                "build": self.common.COMMON_BUILD,
                "mode": "VALIDATE",
                "input_sha256": self.common.file_sha256(str(csv_path)),
                "results": manifest_results,
            }
            (root / self.common.MANIFEST_FILENAME).write_text(json.dumps(manifest), encoding="utf-8")

            def finish(_session, result, _validated):
                result["result"] = "FROZEN" if result["part_number"] == "P1" else "FAILED_FREEZE_WORKFLOW"
                return result

            with mock.patch.object(self.common, "package_root", return_value=str(root)), \
                 mock.patch.object(self.common, "apply_one", side_effect=finish) as apply:
                payload = self.common.run_apply(object())
            self.assertEqual(2, apply.call_count)
            self.assertEqual(1, payload["counts"]["FROZEN"])
            self.assertEqual(1, payload["counts"]["FAILED_FREEZE_WORKFLOW"])

    def test_freeze_only_source_contract(self):
        source = COMMON_PATH.read_text(encoding="utf-8")
        self.assertIn("AssignFreezeStatus", source)
        self.assertNotIn("AssignUnfreezeStatus", source)
        self.assertNotIn("CheckoutParts", source)
        self.assertNotIn("CheckinParts", source)
        self.assertNotIn("AttributePropertiesBuilder", source)
        self.assertNotIn("CreateNewRevision", source)


if __name__ == "__main__":
    unittest.main()
