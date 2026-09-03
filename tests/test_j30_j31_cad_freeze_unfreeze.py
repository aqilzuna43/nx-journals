import importlib.util
import pathlib
import sys
import types
import unittest
from unittest import mock


ROOT = pathlib.Path(__file__).resolve().parents[1]
COMMON_PATH = ROOT / "from_git" / "utils" / "wae_change_control.py"
J30_PATH = ROOT / "from_git" / "journals" / "30_cad_freeze.py"
J31_PATH = ROOT / "from_git" / "journals" / "31_cad_unfreeze.py"


def load_common():
    nxopen = types.ModuleType("NXOpen")
    nxopen_pdm = types.ModuleType("NXOpen.PDM")
    nxopen.PDM = nxopen_pdm
    spec = importlib.util.spec_from_file_location("wae_change_control_test", COMMON_PATH)
    module = importlib.util.module_from_spec(spec)
    with mock.patch.dict(sys.modules, {"NXOpen": nxopen, "NXOpen.PDM": nxopen_pdm}):
        spec.loader.exec_module(module)
    return module


def snapshot(state, version=1, owner="", current_user="aqil", read_only=True):
    return {
        "component_name": "COMPONENT",
        "component_tag": "10",
        "part_identifier": "@DB/P1/A",
        "part_number": "P1",
        "db_part_rev": "A",
        "wae_version": version,
        "wae_version_raw": str(version),
        "wae_attribute": {
            "value": str(version), "type": "STRING", "unset": False,
            "locked": False, "owned_by_system": False, "pdm_based": False,
            "not_saved": False,
        },
        "checkout": {
            "state": state,
            "owner": owner,
            "current_user": current_user,
            "owner_is_current_user": state == "CHECKED_OUT" and owner == current_user,
            "raw": "",
        },
        "read_only": read_only,
        "part_modified": False,
    }


class FakeSelectionManager:
    def __init__(self, selected):
        self.selected = list(selected)

    def GetNumSelectedObjects(self):
        return len(self.selected)

    def GetSelectedTaggedObject(self, index):
        return self.selected[index]


class TestJ30J31Contract(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.common = load_common()

    def test_two_distinct_button_journals_exist(self):
        self.assertTrue(J30_PATH.exists())
        self.assertTrue(J31_PATH.exists())
        self.assertIn('run_ui("FREEZE"', J30_PATH.read_text(encoding="utf-8"))
        self.assertIn('run_ui("UNFREEZE"', J31_PATH.read_text(encoding="utf-8"))

    def test_journals_default_to_apply_for_direct_button_use(self):
        self.assertIn('USER_MODE = "APPLY"', J30_PATH.read_text(encoding="utf-8"))
        self.assertIn('USER_MODE = "APPLY"', J31_PATH.read_text(encoding="utf-8"))

    def test_wae_version_is_strict_positive_integer(self):
        self.assertEqual(1, self.common.parse_wae_version("1"))
        self.assertEqual(42, self.common.parse_wae_version(" 42 "))
        for invalid in ("", "0", "-1", "1.0", "A"):
            with self.subTest(invalid=invalid), self.assertRaises(RuntimeError):
                self.common.parse_wae_version(invalid)

    def test_exactly_one_loaded_component_must_be_preselected(self):
        component = types.SimpleNamespace(
            Prototype=types.SimpleNamespace(PDMPart=object()), IsSuppressed=False
        )
        selected, prototype = self.common.selected_component_target(
            FakeSelectionManager([component])
        )
        self.assertIs(component, selected)
        self.assertIs(component.Prototype, prototype)
        for values in ([], [component, component]):
            with self.subTest(count=len(values)), self.assertRaises(RuntimeError):
                self.common.selected_component_target(FakeSelectionManager(values))

    def test_freeze_already_checked_in_is_verification_only(self):
        before = snapshot("CHECKED_IN")
        report = self.common.base_report("FREEZE", "J30", "APPLY")
        with mock.patch.object(
            self.common, "target_snapshot", side_effect=[before, before]
        ), mock.patch.object(self.common, "save_part") as save, mock.patch.object(
            self.common, "checkin_part"
        ) as checkin:
            result = self.common.freeze(object(), object(), object(), report)
        self.assertEqual("FROZEN_VERIFIED", result["result"])
        save.assert_not_called()
        checkin.assert_not_called()

    def test_freeze_saves_checks_in_and_preserves_revision_and_wae(self):
        before = snapshot("CHECKED_OUT", owner="aqil", read_only=False)
        saved = snapshot("CHECKED_OUT", owner="aqil", read_only=False)
        after = snapshot("CHECKED_IN")
        report = self.common.base_report("FREEZE", "J30", "APPLY")
        with mock.patch.object(
            self.common, "target_snapshot", side_effect=[before, saved, after]
        ), mock.patch.object(self.common, "save_part") as save, mock.patch.object(
            self.common, "checkin_part", return_value="no errors"
        ) as checkin:
            result = self.common.freeze(object(), object(), object(), report)
        self.assertEqual("FROZEN_CHECKED_IN", result["result"])
        save.assert_called_once()
        checkin.assert_called_once()
        self.assertEqual(1, result["after"]["wae_version"])
        self.assertEqual("A", result["after"]["db_part_rev"])

    def test_freeze_blocks_checkout_owned_by_someone_else(self):
        before = snapshot("CHECKED_OUT", owner="other", read_only=False)
        report = self.common.base_report("FREEZE", "J30", "APPLY")
        with mock.patch.object(
            self.common, "target_snapshot", return_value=before
        ), mock.patch.object(self.common, "save_part") as save:
            result = self.common.freeze(object(), object(), object(), report)
        self.assertEqual("BLOCKED", result["result"])
        self.assertIn("another user", result["message"])
        save.assert_not_called()

    def test_unfreeze_blocks_rerun_while_checked_out(self):
        before = snapshot("CHECKED_OUT", owner="aqil", read_only=False)
        report = self.common.base_report("UNFREEZE", "J31", "APPLY")
        with mock.patch.object(
            self.common, "target_snapshot", return_value=before
        ), mock.patch.object(self.common, "checkout_part") as checkout:
            result = self.common.unfreeze(object(), object(), object(), report)
        self.assertEqual("BLOCKED", result["result"])
        checkout.assert_not_called()

    def test_unfreeze_checkouts_increments_once_saves_and_leaves_checked_out(self):
        before = snapshot("CHECKED_IN")
        checked_out = snapshot("CHECKED_OUT", owner="aqil", read_only=False)
        after = snapshot("CHECKED_OUT", version=2, owner="aqil", read_only=False)
        session = types.SimpleNamespace(
            SetUndoMark=mock.Mock(return_value="mark"),
            DeleteUndoMark=mock.Mock(),
            UndoToMark=mock.Mock(),
        )
        report = self.common.base_report("UNFREEZE", "J31", "APPLY")
        with mock.patch.object(
            self.common, "target_snapshot", side_effect=[before, checked_out, after]
        ), mock.patch.object(
            self.common, "checkout_part", return_value="no errors"
        ) as checkout, mock.patch.object(
            self.common, "write_wae_version"
        ) as write, mock.patch.object(
            self.common, "read_wae_attribute", return_value={"value": "2"}
        ), mock.patch.object(self.common, "save_part") as save, mock.patch.object(
            self.common.NXOpen, "Session", create=True
        ) as nx_session:
            nx_session.MarkVisibility.Invisible = "invisible"
            result = self.common.unfreeze(session, object(), object(), report)
        self.assertEqual("UNFROZEN_READY_FOR_EDIT", result["result"])
        checkout.assert_called_once()
        write.assert_called_once_with(session, mock.ANY, 2)
        save.assert_called_once()
        self.assertEqual("CHECKED_OUT", result["after"]["checkout"]["state"])
        self.assertEqual("A", result["after"]["db_part_rev"])

    def test_dry_runs_do_not_mutate(self):
        frozen = snapshot("CHECKED_IN")
        working = snapshot("CHECKED_OUT", owner="aqil", read_only=False)
        with mock.patch.object(self.common, "target_snapshot", return_value=working), \
                mock.patch.object(self.common, "save_part") as save:
            freeze_result = self.common.freeze(
                object(), object(), object(),
                self.common.base_report("FREEZE", "J30", "DRY_RUN"),
            )
        self.assertEqual("DRY_RUN_READY_TO_FREEZE", freeze_result["result"])
        save.assert_not_called()
        with mock.patch.object(self.common, "target_snapshot", return_value=frozen), \
                mock.patch.object(self.common, "checkout_part") as checkout:
            unfreeze_result = self.common.unfreeze(
                object(), object(), object(),
                self.common.base_report("UNFREEZE", "J31", "DRY_RUN"),
            )
        self.assertEqual("DRY_RUN_READY_TO_UNFREEZE", unfreeze_result["result"])
        checkout.assert_not_called()

    def test_scope_is_single_component_and_revision_is_never_written(self):
        source = COMMON_PATH.read_text(encoding="utf-8")
        self.assertIn('"ONE_PRESELECTED_COMPONENT_PROTOTYPE"', source)
        self.assertIn("GetSelectedTaggedObject(0)", source)
        self.assertNotIn("GetChildren(", source)
        self.assertNotIn("DB_PART_REV_TITLE,", source)
        self.assertNotIn("CreateNewRevision", source)


if __name__ == "__main__":
    unittest.main()
