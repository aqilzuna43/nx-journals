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
             release_status="", internal_status=None, has_write_access=None,
             pdm_modifiable=None, part_number="P1", revision="A"):
    if has_write_access is None:
        has_write_access = not read_only
    if pdm_modifiable is None:
        pdm_modifiable = not read_only
    return {
        "component_name": "COMPONENT", "component_tag": "10",
        "part_identifier": "@DB/{0}/{1}".format(part_number, revision),
        "part_number": part_number,
        "db_part_rev": revision, "wae_version": version,
        "wae_version_raw": str(version),
        "wae_class": (
            "NUMERIC_WORKING" if isinstance(version, int) else "ALPHABETIC_FINAL"
        ),
        "wae_validation_error": "",
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
            "has_write_access": has_write_access,
            "pdm_modifiable": pdm_modifiable,
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
        self.assertEqual("WAE-CHANGE-CONTROL-V5", self.common.COMMON_BUILD)
        self.assertIn(
            'EXPECTED_COMMON_BUILD = "WAE-CHANGE-CONTROL-V5"',
            J30_PATH.read_text(encoding="utf-8"),
        )
        self.assertIn(
            'EXPECTED_COMMON_BUILD = "WAE-CHANGE-CONTROL-V5"',
            J31_PATH.read_text(encoding="utf-8"),
        )

    def test_wae_version_is_strict_positive_integer(self):
        self.assertEqual(1, self.common.parse_wae_version("1"))
        self.assertEqual(42, self.common.parse_wae_version(" 42 "))
        for invalid in ("", "0", "-1", "1.0", "A"):
            with self.subTest(invalid=invalid), self.assertRaises(RuntimeError):
                self.common.parse_wae_version(invalid)

    def test_action_neutral_wae_classification_matches_j34_j35(self):
        self.assertEqual(
            ("NUMERIC_WORKING", ""), self.common.classify_wae_version("7", "A")
        )
        self.assertEqual(
            ("ALPHABETIC_FINAL", ""), self.common.classify_wae_version("e", "E")
        )
        self.assertTrue(self.common.classify_wae_version("B", "A")[1])
        self.assertTrue(self.common.classify_wae_version("", "A")[1])

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

    def test_active_assembly_selection_is_excluded_when_children_are_selected(self):
        work = types.SimpleNamespace(PDMPart=object(), Tag=100)
        child_part = types.SimpleNamespace(PDMPart=object(), Tag=200)
        root = types.SimpleNamespace(Prototype=work, IsSuppressed=False, Tag=10)
        child = types.SimpleNamespace(
            Prototype=child_part, IsSuppressed=False, Tag=11
        )
        session = types.SimpleNamespace(Parts=types.SimpleNamespace(Work=work))
        report = self.common.batch_report("UNFREEZE", "J31", "APPLY")
        targets, count = self.common.selected_or_work_targets(
            session, FakeSelectionManager([root, child]), report
        )
        self.assertEqual(2, count)
        self.assertEqual([child_part], [row["part"] for row in targets])
        self.assertEqual(
            "ACTIVE_WORK_PART_FALLBACK_CANDIDATE",
            report["selected_objects"][0]["resolution"],
        )

    def test_selected_geometry_resolves_through_owning_component(self):
        work = types.SimpleNamespace(PDMPart=object(), Tag=100)
        child_part = types.SimpleNamespace(PDMPart=object(), Tag=200)
        child = types.SimpleNamespace(
            Prototype=child_part, IsSuppressed=False, Tag=11
        )
        face = types.SimpleNamespace(OwningComponent=child, Tag=501)
        session = types.SimpleNamespace(Parts=types.SimpleNamespace(Work=work))
        report = self.common.batch_report("UNFREEZE", "J31", "APPLY")
        targets, _ = self.common.selected_or_work_targets(
            session, FakeSelectionManager([face]), report
        )
        self.assertEqual([child_part], [row["part"] for row in targets])
        self.assertEqual("OWNING_COMPONENT_SELECTION", targets[0]["source"])
        self.assertEqual(
            "OWNING_COMPONENT_PROTOTYPE",
            report["selected_objects"][0]["resolution"],
        )

    def test_selected_managed_part_is_a_direct_target(self):
        work = types.SimpleNamespace(PDMPart=object(), Tag=100)
        selected_part = types.SimpleNamespace(PDMPart=object(), Tag=200)
        session = types.SimpleNamespace(Parts=types.SimpleNamespace(Work=work))
        targets, count = self.common.selected_or_work_targets(
            session, FakeSelectionManager([selected_part])
        )
        self.assertEqual(1, count)
        self.assertEqual([selected_part], [row["part"] for row in targets])
        self.assertEqual("MANAGED_PART_SELECTION", targets[0]["source"])

    def test_only_unresolved_selection_falls_back_to_active_work_part(self):
        work = types.SimpleNamespace(PDMPart=object(), Tag=100)
        unrelated = types.SimpleNamespace(Tag=400)
        session = types.SimpleNamespace(Parts=types.SimpleNamespace(Work=work))
        report = self.common.batch_report("UNFREEZE", "J31", "APPLY")
        targets, count = self.common.selected_or_work_targets(
            session, FakeSelectionManager([unrelated]), report
        )
        self.assertEqual(1, count)
        self.assertEqual([work], [row["part"] for row in targets])
        self.assertEqual("ACTIVE_WORK_PART", targets[0]["source"])
        self.assertEqual([0], targets[0]["selected_indexes"])
        self.assertEqual(
            "IGNORED_FOR_ACTIVE_WORK_PART_FALLBACK",
            report["selected_objects"][0]["resolution"],
        )

    def test_mixed_valid_and_unresolved_selection_blocks_complete_batch(self):
        work = types.SimpleNamespace(PDMPart=object(), Tag=100)
        child_part = types.SimpleNamespace(PDMPart=object(), Tag=200)
        child = types.SimpleNamespace(
            Prototype=child_part, IsSuppressed=False, Tag=11
        )
        unrelated = types.SimpleNamespace(Tag=400)
        session = types.SimpleNamespace(Parts=types.SimpleNamespace(Work=work))
        report = self.common.batch_report("UNFREEZE", "J31", "APPLY")
        with self.assertRaisesRegex(RuntimeError, "complete batch was blocked"):
            self.common.selected_or_work_targets(
                session, FakeSelectionManager([child, unrelated]), report
            )
        self.assertEqual(2, report["selected_object_count"])
        self.assertEqual(2, len(report["selected_objects"]))

    def test_j30_mixed_selection_skips_unresolved_and_keeps_valid_target(self):
        work = types.SimpleNamespace(PDMPart=object(), Tag=100)
        child_part = types.SimpleNamespace(PDMPart=object(), Tag=200)
        child = types.SimpleNamespace(
            Prototype=child_part, IsSuppressed=False, Tag=11
        )
        unrelated = types.SimpleNamespace(Tag=400)
        session = types.SimpleNamespace(Parts=types.SimpleNamespace(Work=work))
        report = self.common.batch_report("FREEZE", "J30", "APPLY")
        targets, _ = self.common.selected_or_work_targets(
            session, FakeSelectionManager([child, unrelated]), report,
            allow_partial_selection=True,
        )
        self.assertEqual([child_part], [row["part"] for row in targets])
        self.assertEqual(1, len(report["selection_warnings"]))

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
        frozen = snapshot(
            "CHECKED_IN", release_status="Frozen", read_only=True,
            has_write_access=True, pdm_modifiable=False,
        )
        with mock.patch.object(self.common, "target_snapshot", return_value=frozen):
            report = self.common.make_target_report(
                "FREEZE", object(), target(object()), "J30", "APPLY", 0
            )
        self.assertEqual("PREFLIGHT_READY", report["result"])
        self.assertEqual("ALREADY_FROZEN", report["planned_action"])

    def test_freeze_accepts_matching_alphabetic_final_baseline(self):
        final = snapshot("CHECKED_IN", version="E", revision="E", read_only=True)
        with mock.patch.object(self.common, "target_snapshot", return_value=final):
            report = self.common.make_target_report(
                "FREEZE", object(), target(object()), "J30", "APPLY", 0
            )
        self.assertEqual("PREFLIGHT_READY", report["result"])
        self.assertEqual("ASSIGN_FREEZE_STATUS", report["planned_action"])

    def test_unfreeze_blocks_matching_alphabetic_final_baseline(self):
        final = snapshot(
            "CHECKED_IN", version="E", revision="E", release_status="Frozen"
        )
        with mock.patch.object(self.common, "target_snapshot", return_value=final):
            report = self.common.make_target_report(
                "UNFREEZE", object(), target(object()), "J31", "APPLY", 0
            )
        self.assertEqual("BLOCKED_FINAL_RELEASE_BASELINE", report["result"])
        self.assertIn("immutable", report["message"])

    def test_missing_wae_has_explicit_block_result(self):
        missing = snapshot("CHECKED_IN")
        missing.update({
            "wae_version": "", "wae_version_raw": "", "wae_class": "",
            "wae_validation_error": "WAE_VERSION is blank.",
        })
        with mock.patch.object(self.common, "target_snapshot", return_value=missing):
            report = self.common.make_target_report(
                "FREEZE", object(), target(object()), "J30", "APPLY", 0
            )
        self.assertEqual("BLOCKED_MISSING_WAE_VERSION", report["result"])

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

    def test_freeze_preflight_skips_bad_target_and_runs_safe_target(self):
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
        ), mock.patch.object(
            self.common, "execute_freeze_batch",
            side_effect=lambda session, targets, reports, batch: reports[0].update(
                {"result": "FROZEN"}
            ),
        ) as mutate:
            report = self.common.execute("FREEZE", object(), object(), "J30", "APPLY")
        self.assertEqual("PARTIAL_COMPLETION", report["result"])
        self.assertFalse(report["preflight"]["passed"])
        mutate.assert_called_once()
        self.assertEqual(1, report["counts"]["succeeded"])
        self.assertEqual(1, report["counts"]["blocked"])

    def test_unfreeze_keeps_complete_selection_preflight(self):
        targets = [target(object()), target(object())]
        good = ready_report(
            self.common, "UNFREEZE",
            snapshot("CHECKED_IN", version=2, release_status="Frozen"),
            "UNFREEZE_CHECKOUT_INCREMENT_AND_SAVE", 1,
        )
        final = self.common.base_report("UNFREEZE", "J31", "APPLY")
        final.update({
            "target_index": 2,
            "before": snapshot(
                "CHECKED_IN", version="B", revision="B", part_number="P2",
                release_status="Frozen",
            ),
            "result": "BLOCKED_FINAL_RELEASE_BASELINE",
            "message": "immutable final baseline",
        })
        with mock.patch.object(
            self.common, "selected_or_work_targets", return_value=(targets, 2)
        ), mock.patch.object(
            self.common, "get_available_workflows",
            return_value={"names": ["Part_Unfreeze_Process"]},
        ), mock.patch.object(
            self.common, "make_target_report", side_effect=[good, final]
        ), mock.patch.object(self.common, "execute_unfreeze_batch") as mutate:
            report = self.common.execute(
                "UNFREEZE", object(), object(), "J31", "APPLY"
            )
        self.assertEqual("BLOCKED_BATCH", report["result"])
        mutate.assert_not_called()
        self.assertEqual("NOT_ATTEMPTED_BATCH_BLOCKED", report["targets"][0]["result"])

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
        frozen_first = snapshot(
            "CHECKED_IN", release_status="Frozen", read_only=True,
            has_write_access=True, pdm_modifiable=False,
        )
        frozen_second = snapshot(
            "CHECKED_IN", release_status="Frozen", read_only=True,
            has_write_access=True, pdm_modifiable=False,
        )
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
        self.assertEqual(2, assign.call_count)
        assign.assert_has_calls([
            mock.call(session, [first_part], "FREEZE"),
            mock.call(session, [second_part], "FREEZE"),
        ])
        self.assertEqual(["FROZEN", "FROZEN"], [row["result"] for row in reports])

    def test_freeze_workflow_error_is_warning_when_final_state_is_frozen(self):
        part = object()
        before = snapshot("CHECKED_IN", read_only=True)
        frozen = snapshot(
            "CHECKED_IN", release_status="Frozen", read_only=True,
            has_write_access=True, pdm_modifiable=False,
        )
        report = ready_report(
            self.common, "FREEZE", before, "ASSIGN_FREEZE_STATUS"
        )
        batch = self.common.batch_report("FREEZE", "J30", "APPLY")
        with mock.patch.object(
            self.common, "assign_status_workflow", side_effect=RuntimeError("3520110")
        ), mock.patch.object(self.common, "target_snapshot", return_value=frozen):
            self.common.execute_freeze_batch(
                object(), [target(part)], [report], batch
            )
        self.assertEqual("FROZEN_WITH_WARNING", report["result"])
        self.assertIn("3520110", report["message"])

    def test_freeze_failure_isolated_and_later_target_continues(self):
        first, second = object(), object()
        first_before = snapshot("CHECKED_IN", read_only=True, part_number="P1")
        second_before = snapshot("CHECKED_IN", read_only=True, part_number="P2")
        reports = [
            ready_report(self.common, "FREEZE", first_before, "ASSIGN_FREEZE_STATUS", 1),
            ready_report(self.common, "FREEZE", second_before, "ASSIGN_FREEZE_STATUS", 2),
        ]
        unchanged = snapshot("CHECKED_IN", read_only=True, part_number="P1")
        frozen = snapshot(
            "CHECKED_IN", read_only=True, release_status="Frozen",
            pdm_modifiable=False, part_number="P2",
        )
        batch = self.common.batch_report("FREEZE", "J30", "APPLY")
        with mock.patch.object(
            self.common, "assign_status_workflow",
            side_effect=[RuntimeError("first failed"), "second ok"],
        ) as assign, mock.patch.object(
            self.common, "target_snapshot", side_effect=[unchanged, frozen]
        ):
            self.common.execute_freeze_batch(
                object(), [target(first), target(second)], reports, batch
            )
        self.assertEqual(2, assign.call_count)
        self.assertEqual(
            ["FAILED_FREEZE_WORKFLOW", "FROZEN"],
            [row["result"] for row in reports],
        )

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

    def test_unfreeze_incomplete_target_stops_later_targets(self):
        first, second = object(), object()
        frozen_first = snapshot("CHECKED_IN", version=6, release_status="Frozen")
        frozen_second = snapshot(
            "CHECKED_IN", version=3, release_status="Frozen", part_number="P2"
        )
        reports = [
            ready_report(
                self.common, "UNFREEZE", frozen_first,
                "UNFREEZE_CHECKOUT_INCREMENT_AND_SAVE", 1,
            ),
            ready_report(
                self.common, "UNFREEZE", frozen_second,
                "UNFREEZE_CHECKOUT_INCREMENT_AND_SAVE", 2,
            ),
        ]
        unfrozen = snapshot("CHECKED_IN", version=6, read_only=True)
        batch = self.common.batch_report("UNFREEZE", "J31", "APPLY")
        with mock.patch.object(
            self.common, "assign_status_workflow", return_value="unfreeze"
        ) as assign, mock.patch.object(
            self.common, "target_snapshot", side_effect=[unfrozen, unfrozen]
        ), mock.patch.object(
            self.common, "checkout_part", side_effect=RuntimeError("checkout failed")
        ):
            stopped = self.common.execute_unfreeze_batch(
                object(), [target(first), target(second)], reports, batch
            )
        self.assertTrue(stopped)
        self.assertEqual(1, assign.call_count)
        self.assertEqual("RECOVERY_REQUIRED", reports[0]["result"])
        self.assertEqual("NOT_ATTEMPTED_AFTER_RECOVERY_REQUIRED", reports[1]["result"])

    def test_exact_teamcenter_identity_collapses_distinct_loaded_proxies(self):
        targets = [target(object(), indexes=[0]), target(object(), indexes=[1])]
        reports = [
            ready_report(self.common, "FREEZE", snapshot("CHECKED_IN"),
                         "ASSIGN_FREEZE_STATUS", 1),
            ready_report(self.common, "FREEZE", snapshot("CHECKED_IN"),
                         "ASSIGN_FREEZE_STATUS", 2),
        ]
        collapsed_targets, collapsed_reports = self.common.collapse_exact_identity_targets(
            targets, reports
        )
        self.assertEqual(1, len(collapsed_targets))
        self.assertEqual(1, len(collapsed_reports))
        self.assertEqual([0, 1], collapsed_targets[0]["selected_indexes"])
        self.assertEqual(2, collapsed_reports[0]["selected_occurrence_count"])

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

    def test_recovery_recognizes_frozen_part_when_has_write_access_is_true(self):
        before = snapshot("CHECKED_OUT", owner="aqil", read_only=False)
        after = snapshot(
            "CHECKED_IN", release_status="Frozen", read_only=True,
            has_write_access=True, pdm_modifiable=False,
        )
        report = self.common.base_report("FREEZE", "J30", "APPLY")
        report.update({"before": before, "after": after})
        self.common.mark_recovery_results("FREEZE", [report])
        self.assertEqual("FROZEN_WITH_WARNING", report["result"])

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
