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


def snapshot(state, version=1, owner="", current_user="aqil", read_only=True,
             release_status="", internal_status=None):
    return {
        "component_name": "COMPONENT", "component_tag": "10",
        "part_identifier": "@DB/P1/A", "part_number": "P1",
        "db_part_rev": "A", "wae_version": version,
        "wae_version_raw": str(version),
        "wae_attribute": {
            "value": str(version), "type": "STRING", "unset": False,
            "locked": False, "owned_by_system": False, "pdm_based": False,
            "not_saved": False,
        },
        "checkout": {
            "state": state, "owner": owner, "current_user": current_user,
            "owner_is_current_user": state == "CHECKED_OUT" and owner == current_user,
            "raw": "",
        },
        "release_status": {
            "display": release_status, "internal": list(internal_status or []),
            "display_raw": repr(release_status),
            "internal_raw": repr(internal_status or []), "errors": [],
        },
        "modifiability": {
            "has_write_access": not read_only,
            "pdm_modifiable": not read_only,
            "errors": [],
        },
        "read_only": read_only, "part_modified": False,
    }


def target(part, component=None, indexes=None, occurrences=1):
    return {
        "component": component, "part": part,
        "source": "ASSEMBLY_NAVIGATOR_SELECTION" if component else "ACTIVE_WORK_PART",
        "selected_indexes": list(indexes or []), "occurrence_count": occurrences,
    }


def ready_report(common, action, before, planned, index=1):
    report = common.base_report(action, "BUILD", "APPLY")
    report.update({
        "target_index": index, "before": before, "planned_action": planned,
        "result": "PREFLIGHT_READY", "message": "Target preflight passed.",
    })
    if action == "UNFREEZE":
        report["next_wae_version"] = before["wae_version"] + 1
    return report


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

    def test_two_distinct_apply_button_journals_exist(self):
        self.assertTrue(J30_PATH.exists())
        self.assertTrue(J31_PATH.exists())
        self.assertIn('run_ui("FREEZE"', J30_PATH.read_text(encoding="utf-8"))
        self.assertIn('run_ui("UNFREEZE"', J31_PATH.read_text(encoding="utf-8"))
        self.assertIn('USER_MODE = "APPLY"', J30_PATH.read_text(encoding="utf-8"))
        self.assertIn('USER_MODE = "APPLY"', J31_PATH.read_text(encoding="utf-8"))

    def test_wae_version_is_strict_positive_integer(self):
        self.assertEqual(1, self.common.parse_wae_version("1"))
        self.assertEqual(42, self.common.parse_wae_version(" 42 "))
        for invalid in ("", "0", "-1", "1.0", "A"):
            with self.subTest(invalid=invalid), self.assertRaises(RuntimeError):
                self.common.parse_wae_version(invalid)

    def test_no_preselection_targets_only_active_work_part(self):
        work = types.SimpleNamespace(PDMPart=object(), Tag=100)
        session = types.SimpleNamespace(Parts=types.SimpleNamespace(Work=work))
        targets, count = self.common.selected_or_work_targets(
            session, FakeSelectionManager([])
        )
        self.assertEqual(0, count)
        self.assertEqual(1, len(targets))
        self.assertIs(work, targets[0]["part"])
        self.assertIsNone(targets[0]["component"])

    def test_selections_override_active_part_and_duplicate_prototypes_collapse(self):
        work = types.SimpleNamespace(PDMPart=object(), Tag=100)
        shared = types.SimpleNamespace(PDMPart=object(), Tag=200)
        other = types.SimpleNamespace(PDMPart=object(), Tag=201)
        components = [
            types.SimpleNamespace(Prototype=shared, IsSuppressed=False),
            types.SimpleNamespace(Prototype=other, IsSuppressed=False),
            types.SimpleNamespace(Prototype=shared, IsSuppressed=False),
        ]
        session = types.SimpleNamespace(Parts=types.SimpleNamespace(Work=work))
        targets, count = self.common.selected_or_work_targets(
            session, FakeSelectionManager(components)
        )
        self.assertEqual(3, count)
        self.assertEqual(2, len(targets))
        self.assertNotIn(work, [row["part"] for row in targets])
        self.assertEqual([0, 2], targets[0]["selected_indexes"])
        self.assertEqual(2, targets[0]["occurrence_count"])

    def test_selected_subassembly_does_not_recurse(self):
        prototype = types.SimpleNamespace(PDMPart=object(), Tag=300)
        component = types.SimpleNamespace(
            Prototype=prototype, IsSuppressed=False,
            GetChildren=mock.Mock(side_effect=AssertionError("must not recurse")),
        )
        targets, _ = self.common.selected_or_work_targets(
            object(), FakeSelectionManager([component])
        )
        self.assertEqual([prototype], [row["part"] for row in targets])
        component.GetChildren.assert_not_called()

    def test_selected_non_component_is_rejected(self):
        with self.assertRaisesRegex(RuntimeError, "not a loaded Assembly Navigator component"):
            self.common.selected_or_work_targets(
                object(), FakeSelectionManager([types.SimpleNamespace(Tag=400)])
            )

    def test_freeze_and_unfreeze_status_classification(self):
        frozen = snapshot("CHECKED_IN", release_status="Frozen",
                          internal_status=["WAE_FROZEN"])
        released = snapshot("CHECKED_IN", release_status="Released")
        mixed = snapshot("CHECKED_IN", release_status="Frozen; Released")
        unfrozen = snapshot("CHECKED_IN", release_status="Unfrozen")
        self.assertTrue(self.common.is_frozen_status(frozen))
        self.assertFalse(self.common.has_other_release_status(frozen))
        self.assertTrue(self.common.has_other_release_status(released))
        self.assertTrue(self.common.has_other_release_status(mixed))
        self.assertFalse(self.common.is_frozen_status(unfrozen))
        self.assertFalse(self.common.has_other_release_status(unfrozen))

    def test_freeze_preflight_distinguishes_checkin_from_freeze(self):
        checked_in = snapshot("CHECKED_IN", read_only=True)
        with mock.patch.object(self.common, "target_snapshot", return_value=checked_in):
            report = self.common.make_target_report(
                "FREEZE", object(), target(object()), "J30", "APPLY", 0
            )
        self.assertEqual("PREFLIGHT_READY", report["result"])
        self.assertEqual("ASSIGN_FREEZE_STATUS", report["planned_action"])

    def test_freeze_preflight_accepts_positive_frozen_read_only_target(self):
        frozen = snapshot("CHECKED_IN", release_status="Frozen", read_only=True)
        with mock.patch.object(self.common, "target_snapshot", return_value=frozen):
            report = self.common.make_target_report(
                "FREEZE", object(), target(object()), "J30", "APPLY", 0
            )
        self.assertEqual("PREFLIGHT_READY", report["result"])
        self.assertEqual("ALREADY_FROZEN", report["planned_action"])

    def test_unfreeze_preflight_requires_positive_freeze_status(self):
        with mock.patch.object(
            self.common, "target_snapshot", return_value=snapshot("CHECKED_IN")
        ):
            report = self.common.make_target_report(
                "UNFREEZE", object(), target(object()), "J31", "APPLY", 0
            )
        self.assertEqual("PREFLIGHT_BLOCKED", report["result"])
        self.assertIn("positive freeze status", report["message"])

    def test_assign_status_uses_exact_proven_workflows(self):
        errors = types.SimpleNamespace(FreeResource=mock.Mock())
        freeze_method = mock.Mock(return_value=errors)
        unfreeze_method = mock.Mock(return_value=errors)
        session = types.SimpleNamespace(PdmSession=types.SimpleNamespace(
            AssignFreezeStatus=freeze_method, AssignUnfreezeStatus=unfreeze_method,
        ))
        parts = [object(), object()]
        self.common.assign_status_workflow(session, parts, "FREEZE")
        freeze_method.assert_called_once_with(parts, "Part_Freeze_Process")
        self.common.assign_status_workflow(session, parts, "UNFREEZE")
        unfreeze_method.assert_called_once_with(parts, "Part_Unfreeze_Process")
        self.assertEqual(2, errors.FreeResource.call_count)

    def test_complete_batch_preflight_blocks_every_mutation(self):
        targets = [target(object()), target(object())]
        good = ready_report(
            self.common, "FREEZE",
            snapshot("CHECKED_OUT", owner="aqil", read_only=False),
            "SAVE_CHECKIN_AND_ASSIGN_FREEZE_STATUS", 1,
        )
        blocked = self.common.base_report("FREEZE", "J30", "APPLY")
        blocked.update({"target_index": 2, "result": "PREFLIGHT_BLOCKED",
                        "message": "bad target"})
        with mock.patch.object(
            self.common, "selected_or_work_targets", return_value=(targets, 2)
        ), mock.patch.object(
            self.common, "get_available_workflows",
            return_value={"names": ["Part_Freeze_Process"]}
        ), mock.patch.object(
            self.common, "make_target_report", side_effect=[good, blocked]
        ), mock.patch.object(self.common, "execute_freeze_batch") as mutate:
            report = self.common.execute("FREEZE", object(), object(), "J30", "APPLY")
        self.assertEqual("BLOCKED_BATCH", report["result"])
        self.assertFalse(report["preflight"]["passed"])
        mutate.assert_not_called()

    def test_freeze_batch_saves_checkins_assigns_status_and_verifies(self):
        session = object()
        first_part, second_part = object(), object()
        targets = [target(first_part), target(second_part)]
        first_before = snapshot("CHECKED_OUT", owner="aqil", read_only=False)
        second_before = snapshot("CHECKED_IN", read_only=True)
        reports = [
            ready_report(self.common, "FREEZE", first_before,
                         "SAVE_CHECKIN_AND_ASSIGN_FREEZE_STATUS", 1),
            ready_report(self.common, "FREEZE", second_before,
                         "ASSIGN_FREEZE_STATUS", 2),
        ]
        frozen_first = snapshot("CHECKED_IN", release_status="Frozen", read_only=True)
        frozen_second = snapshot("CHECKED_IN", release_status="Frozen", read_only=True)
        batch = self.common.batch_report("FREEZE", "J30", "APPLY")
        with mock.patch.object(
            self.common, "target_snapshot",
            side_effect=[first_before, snapshot("CHECKED_IN"), frozen_first, frozen_second]
        ), mock.patch.object(self.common, "save_part") as save, mock.patch.object(
            self.common, "checkin_parts", return_value="checkin"
        ) as checkin, mock.patch.object(
            self.common, "assign_status_workflow", return_value="freeze"
        ) as assign:
            self.common.execute_freeze_batch(session, targets, reports, batch)
        save.assert_called_once_with(first_part)
        checkin.assert_called_once_with([first_part])
        assign.assert_called_once_with(session, [first_part, second_part], "FREEZE")
        self.assertEqual(["FROZEN", "FROZEN"], [row["result"] for row in reports])

    def test_unfreeze_batch_status_checkout_increment_save_sequence(self):
        part = object()
        targets = [target(part)]
        frozen = snapshot("CHECKED_IN", version=6, release_status="Frozen")
        report = ready_report(self.common, "UNFREEZE", frozen,
                              "UNFREEZE_CHECKOUT_INCREMENT_AND_SAVE")
        unfrozen = snapshot("CHECKED_IN", version=6, read_only=True)
        checked_out = snapshot("CHECKED_OUT", version=6, owner="aqil", read_only=False)
        saved = snapshot("CHECKED_OUT", version=7, owner="aqil", read_only=False)
        batch = self.common.batch_report("UNFREEZE", "J31", "APPLY")
        session = types.SimpleNamespace(
            SetUndoMark=mock.Mock(return_value="mark"), DeleteUndoMark=mock.Mock()
        )
        with mock.patch.object(
            self.common, "target_snapshot", side_effect=[unfrozen, checked_out, saved]
        ), mock.patch.object(
            self.common, "assign_status_workflow", return_value="unfreeze"
        ) as unfreeze, mock.patch.object(
            self.common, "checkout_parts", return_value="checkout"
        ) as checkout, mock.patch.object(
            self.common, "write_wae_version"
        ) as write, mock.patch.object(
            self.common, "read_wae_attribute", return_value={"value": "7"}
        ), mock.patch.object(self.common, "save_part") as save, mock.patch.object(
            self.common.NXOpen, "Session", create=True
        ) as nx_session:
            nx_session.MarkVisibility.Invisible = "invisible"
            self.common.execute_unfreeze_batch(session, targets, [report], batch)
        unfreeze.assert_called_once_with(session, [part], "UNFREEZE")
        checkout.assert_called_once_with([part])
        write.assert_called_once_with(session, part, 7)
        save.assert_called_once_with(part)
        self.assertEqual("UNFROZEN_READY_FOR_EDIT", report["result"])

    def test_runtime_failure_stops_and_requires_recovery(self):
        targets = [target(object())]
        ready = ready_report(
            self.common, "FREEZE",
            snapshot("CHECKED_OUT", owner="aqil", read_only=False),
            "SAVE_CHECKIN_AND_ASSIGN_FREEZE_STATUS",
        )
        with mock.patch.object(
            self.common, "selected_or_work_targets", return_value=(targets, 0)
        ), mock.patch.object(
            self.common, "get_available_workflows",
            return_value={"names": ["Part_Freeze_Process"]}
        ), mock.patch.object(
            self.common, "make_target_report", return_value=ready
        ), mock.patch.object(
            self.common, "execute_freeze_batch",
            side_effect=RuntimeError("Teamcenter failed")
        ), mock.patch.object(self.common, "capture_after_states"):
            report = self.common.execute("FREEZE", object(), object(), "J30", "APPLY")
        self.assertEqual("RECOVERY_REQUIRED", report["result"])
        self.assertIn("Teamcenter failed", report["message"])

    def test_dry_run_preflights_but_does_not_mutate(self):
        targets = [target(object())]
        ready = ready_report(
            self.common, "UNFREEZE",
            snapshot("CHECKED_IN", release_status="Frozen"),
            "UNFREEZE_CHECKOUT_INCREMENT_AND_SAVE",
        )
        with mock.patch.object(
            self.common, "selected_or_work_targets", return_value=(targets, 0)
        ), mock.patch.object(
            self.common, "get_available_workflows",
            return_value={"names": ["Part_Unfreeze_Process"]}
        ), mock.patch.object(
            self.common, "make_target_report", return_value=ready
        ), mock.patch.object(self.common, "execute_unfreeze_batch") as mutate:
            report = self.common.execute("UNFREEZE", object(), object(), "J31", "DRY_RUN")
        self.assertEqual("DRY_RUN_READY", report["result"])
        mutate.assert_not_called()

    def test_scope_is_surgical_and_formal_revision_is_never_written(self):
        source = COMMON_PATH.read_text(encoding="utf-8")
        self.assertIn('"SELECTED_COMPONENT_PROTOTYPES_OR_ACTIVE_WORK_PART"', source)
        self.assertIn("GetSelectedTaggedObject(index)", source)
        self.assertIn('safe_property(parts, "Work")', source)
        self.assertIn('FREEZE_WORKFLOW = "Part_Freeze_Process"', source)
        self.assertIn('UNFREEZE_WORKFLOW = "Part_Unfreeze_Process"', source)
        self.assertNotIn("GetChildren(", source)
        self.assertNotIn("DB_PART_REV_TITLE,", source)
        self.assertNotIn("CreateNewRevision", source)


if __name__ == "__main__":
    unittest.main()
