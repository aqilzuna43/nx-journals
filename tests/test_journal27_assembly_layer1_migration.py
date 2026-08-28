import csv
import datetime
import importlib.util
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
    ROOT / "from_git" / "journals"
    / "27_move_assembly_components_to_layer_1.py"
)


def load_journal():
    nxopen = types.ModuleType("NXOpen")
    nxopen.NXObject = types.SimpleNamespace(
        AttributeType=types.SimpleNamespace(String="String")
    )
    nxopen.Session = types.SimpleNamespace(
        MarkVisibility=types.SimpleNamespace(Visible="Visible"),
        LibraryUnloadOption=types.SimpleNamespace(Immediately="Immediately"),
    )
    prior = sys.modules.get("NXOpen")
    sys.modules["NXOpen"] = nxopen
    try:
        spec = importlib.util.spec_from_file_location("journal27", JOURNAL)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        return module
    finally:
        if prior is None:
            sys.modules.pop("NXOpen", None)
        else:
            sys.modules["NXOpen"] = prior


class Point:
    def __init__(self, x=0.0, y=0.0, z=0.0):
        self.X, self.Y, self.Z = x, y, z


class Matrix:
    def __init__(self):
        self.Xx, self.Xy, self.Xz = 1.0, 0.0, 0.0
        self.Yx, self.Yy, self.Yz = 0.0, 1.0, 0.0
        self.Zx, self.Zy, self.Zz = 0.0, 0.0, 1.0


class FakePrototype:
    def __init__(self, name, tag):
        self.Name = name
        self.Tag = tag


class FakeComponent:
    def __init__(
        self,
        name,
        tag,
        layer_option,
        reported_layer=None,
        suppressed=False,
        blanked=False,
        reference_set="MODEL",
        non_geometric=False,
        prototype=True,
        position=None,
    ):
        self.Name = name
        self.Tag = tag
        self.Layer = layer_option if reported_layer is None else reported_layer
        self._layer_option = layer_option
        self.IsSuppressed = suppressed
        self.IsBlanked = blanked
        self.ReferenceSet = reference_set
        self._non_geometric = non_geometric
        self.Prototype = (
            FakePrototype(name + "_PROTO", tag + 1000) if prototype else None
        )
        self.Parent = None
        self.position = position or Point(float(tag), 0.0, 0.0)
        self.orientation = Matrix()
        self.set_calls = []
        self.set_error = None
        self.ignore_set = False
        self.mutate_blanking = False
        self.mutate_position = False
        self.nested_get_children_calls = 0

    def GetLayerOption(self):
        return self._layer_option

    def SetLayerOption(self, layer):
        self.set_calls.append(layer)
        if self.set_error:
            raise RuntimeError(self.set_error)
        if not self.ignore_set:
            self._layer_option = layer
            self.Layer = layer
        if self.mutate_blanking:
            self.IsBlanked = not self.IsBlanked
        if self.mutate_position:
            self.position.X += 10.0

    def GetNonGeometricState(self):
        return self._non_geometric

    def GetPosition(self):
        return self.position, self.orientation

    def GetChildren(self):
        self.nested_get_children_calls += 1
        raise AssertionError("J27 must not recurse into component children")


class FakeRoot:
    def __init__(self, children, tag=5000):
        self.Tag = tag
        self.children = list(children)
        self.get_children_calls = 0
        self.scan_error = None
        for child in self.children:
            child.Parent = self

    def GetChildren(self):
        self.get_children_calls += 1
        if self.scan_error:
            raise RuntimeError(self.scan_error)
        return list(self.children)


class FakeLayers:
    def __init__(self, work_layer=7):
        self.WorkLayer = work_layer
        self.states = {
            index: "STATE_{0}".format(index) for index in range(1, 257)
        }

    def GetState(self, layer):
        return self.states[layer]


class FakePdmPart:
    def __init__(self, checked=True, owner="aqil"):
        self.checked = checked
        self.owner = owner

    def GetCheckedoutStatusAndUser(self):
        if self.checked == "ERROR":
            raise RuntimeError("checkout lookup failed")
        if self.checked == "UNKNOWN":
            return "unknown-status", self.owner
        return self.checked, self.owner


class FakePdmSession:
    def __init__(self, user="aqil"):
        self.user = user

    def GetUserName(self):
        return self.user


class FakePart:
    _next_tag = 9000

    def __init__(
        self,
        children,
        managed=False,
        read_only=False,
        checked=True,
        owner="aqil",
        assembly=True,
    ):
        FakePart._next_tag += 1
        self.Tag = FakePart._next_tag
        self.Name = "TOP_ASSEMBLY"
        self.Leaf = "TOP_ASSEMBLY"
        self.FullPath = (
            "@DB/TOP_ASSEMBLY/A"
            if managed else r"C:\temp\TOP_ASSEMBLY.prt"
        )
        self.JournalIdentifier = self.FullPath
        self.IsReadOnly = read_only
        self.Layers = FakeLayers()
        self.root = FakeRoot(children) if assembly else None
        self.ComponentAssembly = (
            types.SimpleNamespace(RootComponent=self.root)
            if assembly else None
        )
        self.PDMPart = FakePdmPart(checked, owner) if managed else None
        self.attributes = {
            "DB_PART_NO": "TOP_ASSEMBLY",
            "DB_PART_REV": "A",
        }

    def GetStringAttribute(self, name):
        return self.attributes.get(name, "")


class FakeListingWindow:
    def __init__(self):
        self.lines = []
        self.opened = False

    def Open(self):
        self.opened = True

    def WriteFullline(self, value):
        self.lines.append(str(value))


class FakeSession:
    def __init__(
        self,
        part,
        display_part="SAME",
        managed=False,
        user="aqil",
        mark_error=False,
        undo_error=False,
    ):
        display = part if display_part == "SAME" else display_part
        self.Parts = types.SimpleNamespace(Work=part, Display=display)
        self.IsManagedMode = managed
        self.PdmSession = FakePdmSession(user)
        self.ListingWindow = FakeListingWindow()
        self.mark_error = mark_error
        self.undo_error = undo_error
        self.set_mark_calls = []
        self.undo_calls = []
        self.delete_calls = []
        self._baseline = None

    def SetUndoMark(self, visibility, name):
        self.set_mark_calls.append((visibility, name))
        if self.mark_error:
            raise RuntimeError("undo mark unavailable")
        part = self.Parts.Work
        self._baseline = {
            "children": list(part.root.children),
            "components": {
                item.Tag: {
                    "layer_option": item._layer_option,
                    "layer": item.Layer,
                    "suppressed": item.IsSuppressed,
                    "blanked": item.IsBlanked,
                    "reference_set": item.ReferenceSet,
                    "non_geometric": item._non_geometric,
                    "position": (item.position.X, item.position.Y, item.position.Z),
                }
                for item in part.root.children
            },
            "work_layer": part.Layers.WorkLayer,
            "states": dict(part.Layers.states),
        }
        return "MARK-1"

    def UndoToMark(self, mark, name):
        self.undo_calls.append((mark, name))
        if self.undo_error:
            raise RuntimeError("undo failed")
        part = self.Parts.Work
        part.root.children = list(self._baseline["children"])
        for item in part.root.children:
            state = self._baseline["components"][item.Tag]
            item._layer_option = state["layer_option"]
            item.Layer = state["layer"]
            item.IsSuppressed = state["suppressed"]
            item.IsBlanked = state["blanked"]
            item.ReferenceSet = state["reference_set"]
            item._non_geometric = state["non_geometric"]
            item.position.X, item.position.Y, item.position.Z = state["position"]
        part.Layers.WorkLayer = self._baseline["work_layer"]
        part.Layers.states = dict(self._baseline["states"])

    def DeleteUndoMark(self, mark, name):
        self.delete_calls.append((mark, name))


class AssemblyLayerOneMigrationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.journal = load_journal()

    def run_in_temp(self, session, mode="DRY_RUN"):
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        now = datetime.datetime(
            2026,
            8,
            28,
            15,
            0,
            tzinfo=datetime.timezone(datetime.timedelta(hours=8)),
        )
        with mock.patch.object(self.journal, "io_root", return_value=folder.name):
            return self.journal.run(session, run_datetime=now, mode=mode)

    def test_mode_defaults_override_and_validation(self):
        with mock.patch.dict(os.environ, {}, clear=True), mock.patch.object(
            self.journal, "USER_MODE", "DRY_RUN"
        ):
            self.assertEqual(self.journal.configured_mode(), "DRY_RUN")
        with mock.patch.dict(os.environ, {"NX_J27_MODE": "apply"}, clear=True):
            self.assertEqual(self.journal.configured_mode(), "APPLY")
        with mock.patch.dict(os.environ, {"NX_J27_MODE": "MOVE"}, clear=True):
            with self.assertRaisesRegex(RuntimeError, "DRY_RUN or APPLY"):
                self.journal.configured_mode()

    def test_dry_run_includes_all_direct_occurrence_states_without_recursing(self):
        components = [
            FakeComponent("SUPPRESSED", 1, 20, suppressed=True),
            FakeComponent("BLANKED", 2, -1, reported_layer=25, blanked=True),
            FakeComponent("REFERENCE", 3, 30, reference_set="EMPTY"),
            FakeComponent("NON_GEOMETRIC", 4, 40, non_geometric=True),
            FakeComponent("UNLOADED", 5, 50, prototype=False),
        ]
        part = FakePart(components)
        session = FakeSession(part)

        csv_path, json_path, report = self.run_in_temp(session)

        self.assertEqual(report["verdict"]["status"], "DRY_RUN_READY")
        self.assertEqual(report["counts"]["direct_component_count"], 5)
        self.assertEqual(report["counts"]["move_candidate_count"], 5)
        self.assertEqual(report["counts"]["suppressed_count"], 1)
        self.assertEqual(report["counts"]["blanked_count"], 1)
        self.assertEqual(report["counts"]["non_geometric_count"], 1)
        self.assertEqual(report["counts"]["prototype_unavailable_count"], 1)
        self.assertEqual(part.root.get_children_calls, 1)
        self.assertTrue(all(item.nested_get_children_calls == 0 for item in components))
        self.assertTrue(all(item.set_calls == [] for item in components))
        self.assertEqual(session.set_mark_calls, [])
        self.assertEqual(Path(csv_path).parent, Path(json_path).parent)
        self.assertEqual(
            Path(csv_path).parent.parent.name,
            "NX_ASSEMBLY_LAYER_1_MIGRATION",
        )
        self.assertTrue(Path(csv_path).read_bytes().startswith(b"\xef\xbb\xbf"))
        with open(csv_path, "r", encoding="utf-8-sig", newline="") as handle:
            rows = list(csv.DictReader(handle))
        self.assertEqual(rows[0]["ROW_TYPE"], "SUMMARY")
        self.assertEqual(len(rows), 6)
        payload = json.loads(Path(json_path).read_text(encoding="utf-8"))
        self.assertFalse(payload["configuration"]["recursive"])
        self.assertFalse(payload["configuration"]["force_load"])

    def test_apply_calls_set_layer_option_only_for_noncompliant_children(self):
        change_a = FakeComponent("A", 1, -1, reported_layer=12)
        change_b = FakeComponent("B", 2, 7)
        compliant = FakeComponent("C", 3, 1)
        components = [change_a, change_b, compliant]
        part = FakePart(components)
        session = FakeSession(part)
        positions = [item.position.X for item in components]
        states = dict(part.Layers.states)

        _, _, report = self.run_in_temp(session, mode="APPLY")

        self.assertEqual(report["verdict"]["status"], "APPLIED_VERIFIED")
        self.assertEqual(change_a.set_calls, [1])
        self.assertEqual(change_b.set_calls, [1])
        self.assertEqual(compliant.set_calls, [])
        self.assertEqual([item.Layer for item in components], [1, 1, 1])
        self.assertEqual([item.position.X for item in components], positions)
        self.assertEqual(part.Layers.states, states)
        self.assertEqual(part.Layers.WorkLayer, 7)
        self.assertEqual(session.undo_calls, [])
        self.assertEqual(session.delete_calls, [])
        self.assertTrue(report["action"]["successful_change_left_undoable"])

    def test_already_compliant_is_noop(self):
        component = FakeComponent("A", 1, 1)
        session = FakeSession(FakePart([component]))

        _, _, report = self.run_in_temp(session, mode="APPLY")

        self.assertEqual(report["verdict"]["status"], "ALREADY_COMPLIANT")
        self.assertEqual(component.set_calls, [])
        self.assertEqual(session.set_mark_calls, [])

    def test_empty_assembly_is_noop(self):
        session = FakeSession(FakePart([]))

        csv_path, _, report = self.run_in_temp(session, mode="APPLY")

        self.assertEqual(
            report["verdict"]["status"], "NO_COMPONENT_OCCURRENCES"
        )
        self.assertTrue(Path(csv_path).is_file())
        self.assertEqual(session.set_mark_calls, [])

    def test_invalid_work_display_or_nonassembly_context_writes_evidence(self):
        cases = (
            (FakeSession(None, display_part=None), "FAILED_NO_WORK_PART"),
            (
                FakeSession(FakePart([]), display_part=FakePart([])),
                "FAILED_CONTEXT",
            ),
            (FakeSession(FakePart([], assembly=False)), "FAILED_NOT_ASSEMBLY"),
        )
        for session, expected in cases:
            with self.subTest(expected=expected):
                csv_path, json_path, report = self.run_in_temp(session)
                self.assertEqual(report["verdict"]["status"], expected)
                self.assertTrue(Path(csv_path).is_file())
                self.assertTrue(Path(json_path).is_file())

    def test_root_scan_failure_fails_closed_without_mutation(self):
        component = FakeComponent("A", 1, 2)
        part = FakePart([component])
        part.root.scan_error = "root unavailable"
        session = FakeSession(part)

        _, _, report = self.run_in_temp(session, mode="APPLY")

        self.assertEqual(report["verdict"]["status"], "FAILED_SCAN")
        self.assertEqual(component.set_calls, [])
        self.assertEqual(session.set_mark_calls, [])

    def test_native_read_only_assembly_is_blocked(self):
        component = FakeComponent("A", 1, 2)
        part = FakePart([component], read_only=True)
        session = FakeSession(part)

        _, _, report = self.run_in_temp(session, mode="APPLY")

        self.assertEqual(report["verdict"]["status"], "BLOCKED_WRITE_ACCESS")
        self.assertEqual(component.set_calls, [])

    def test_native_unknown_read_only_state_is_blocked(self):
        component = FakeComponent("A", 1, 2)
        part = FakePart([component])
        del part.IsReadOnly
        session = FakeSession(part)

        _, _, report = self.run_in_temp(session, mode="APPLY")

        self.assertEqual(report["verdict"]["status"], "BLOCKED_WRITE_ACCESS")
        self.assertIn("unavailable", report["access"]["message"])
        self.assertEqual(component.set_calls, [])

    def test_managed_write_gate_requires_current_user_checkout_and_writable(self):
        cases = (
            (False, "aqil", "aqil", False),
            ("UNKNOWN", "", "aqil", False),
            (True, "other", "aqil", False),
            (True, "aqil", "", False),
            (True, "aqil", "aqil", True),
        )
        for checked, owner, user, read_only in cases:
            with self.subTest(checked=checked, owner=owner, user=user):
                component = FakeComponent("A", 1, 2)
                part = FakePart(
                    [component], managed=True, checked=checked,
                    owner=owner, read_only=read_only,
                )
                session = FakeSession(part, managed=True, user=user)

                _, _, report = self.run_in_temp(session, mode="APPLY")

                self.assertEqual(
                    report["verdict"]["status"], "BLOCKED_WRITE_ACCESS"
                )
                self.assertEqual(component.set_calls, [])

    def test_managed_current_user_checkout_can_apply(self):
        component = FakeComponent("A", 1, 2)
        part = FakePart(
            [component], managed=True, checked=True, owner="aqil"
        )
        session = FakeSession(part, managed=True, user="AQIL")

        _, _, report = self.run_in_temp(session, mode="APPLY")

        self.assertEqual(report["verdict"]["status"], "APPLIED_VERIFIED")
        self.assertEqual(component.set_calls, [1])
        self.assertEqual(report["access"]["checkout_owner"], "aqil")

    def test_managed_unknown_read_only_state_is_blocked(self):
        component = FakeComponent("A", 1, 2)
        part = FakePart(
            [component], managed=True, checked=True, owner="aqil"
        )
        del part.IsReadOnly
        session = FakeSession(part, managed=True, user="aqil")

        _, _, report = self.run_in_temp(session, mode="APPLY")

        self.assertEqual(report["verdict"]["status"], "BLOCKED_WRITE_ACCESS")
        self.assertIn("unavailable", report["access"]["message"])

    def test_partial_api_failure_rolls_back_every_component(self):
        first = FakeComponent("A", 1, 2)
        second = FakeComponent("B", 2, 3)
        second.set_error = "NX layer error"
        part = FakePart([first, second])
        session = FakeSession(part)

        _, _, report = self.run_in_temp(session, mode="APPLY")

        self.assertEqual(report["verdict"]["status"], "ROLLED_BACK")
        self.assertEqual([first.Layer, second.Layer], [2, 3])
        self.assertEqual(len(session.undo_calls), 1)
        self.assertEqual(len(session.delete_calls), 1)

    def test_layer_verification_mismatch_rolls_back(self):
        component = FakeComponent("A", 1, 2)
        component.ignore_set = True
        part = FakePart([component])
        session = FakeSession(part)

        _, _, report = self.run_in_temp(session, mode="APPLY")

        self.assertEqual(report["verdict"]["status"], "ROLLED_BACK")
        self.assertTrue(report["action"]["verification_errors"])
        self.assertEqual(component.Layer, 2)

    def test_blanking_or_position_change_rolls_back(self):
        for flag in ("mutate_blanking", "mutate_position"):
            with self.subTest(flag=flag):
                component = FakeComponent("A", 1, 2)
                setattr(component, flag, True)
                part = FakePart([component])
                session = FakeSession(part)

                _, _, report = self.run_in_temp(session, mode="APPLY")

                self.assertEqual(report["verdict"]["status"], "ROLLED_BACK")
                self.assertEqual(component.Layer, 2)
                self.assertFalse(component.IsBlanked)
                self.assertEqual(component.position.X, 1.0)

    def test_membership_or_layer_state_change_rolls_back(self):
        scenarios = ("membership", "layer_state")
        for scenario in scenarios:
            with self.subTest(scenario=scenario):
                component = FakeComponent("A", 1, 2)
                part = FakePart([component])
                original_set = component.SetLayerOption

                def changed(layer):
                    original_set(layer)
                    if scenario == "membership":
                        part.root.children = []
                    else:
                        part.Layers.states[8] = "MUTATED"

                component.SetLayerOption = changed
                session = FakeSession(part)

                _, _, report = self.run_in_temp(session, mode="APPLY")

                self.assertEqual(report["verdict"]["status"], "ROLLED_BACK")
                self.assertEqual(part.root.children, [component])
                self.assertEqual(part.Layers.states[8], "STATE_8")

    def test_undo_failure_is_prominent(self):
        first = FakeComponent("A", 1, 2)
        second = FakeComponent("B", 2, 3)
        second.set_error = "second failed"
        part = FakePart([first, second])
        session = FakeSession(part, undo_error=True)

        _, _, report = self.run_in_temp(session, mode="APPLY")

        self.assertEqual(report["verdict"]["status"], "ROLLBACK_FAILED")
        self.assertEqual(first.Layer, 1)
        self.assertIn("UndoToMark failed", report["rollback"]["error"])

    def test_evidence_failure_after_success_rolls_back_then_reports(self):
        component = FakeComponent("A", 1, 2)
        part = FakePart([component])
        session = FakeSession(part)
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        original_write = self.journal.write_outputs
        calls = []

        def flaky(report, output_folder, stem):
            calls.append(report["verdict"]["status"])
            if len(calls) == 1:
                raise OSError("disk interrupted")
            return original_write(report, output_folder, stem)

        with mock.patch.object(
            self.journal, "io_root", return_value=folder.name
        ), mock.patch.object(
            self.journal, "write_outputs", side_effect=flaky
        ):
            csv_path, json_path, report = self.journal.run(
                session,
                run_datetime=datetime.datetime(2026, 8, 28, 15, 0),
                mode="APPLY",
            )

        self.assertEqual(calls, ["APPLIED_VERIFIED", "ROLLED_BACK"])
        self.assertEqual(report["verdict"]["status"], "ROLLED_BACK")
        self.assertEqual(component.Layer, 2)
        self.assertTrue(Path(csv_path).is_file())
        self.assertTrue(Path(json_path).is_file())

    def test_undo_mark_failure_never_changes_components(self):
        component = FakeComponent("A", 1, 2)
        part = FakePart([component])
        session = FakeSession(part, mark_error=True)

        _, _, report = self.run_in_temp(session, mode="APPLY")

        self.assertEqual(report["verdict"]["status"], "ROLLED_BACK")
        self.assertEqual(component.set_calls, [])
        self.assertFalse(report["action"]["attempted"])

    def test_listing_window_output(self):
        session = FakeSession(None, display_part=None)

        self.journal.log_line(session, "first\nsecond")

        self.assertTrue(session.ListingWindow.opened)
        self.assertEqual(session.ListingWindow.lines, ["first", "second"])

    def test_source_has_no_forbidden_mutation_or_loading_calls(self):
        source = JOURNAL.read_text(encoding="utf-8")

        for token in (
            ".Save(", ".Checkout(", ".Checkin(", "LoadThisPartFully(",
            "LoadFully(", "MoveDisplayableObjects(", ".Bodies",
        ):
            self.assertNotIn(token, source)
        self.assertEqual(source.count('required_call(root, "GetChildren"'), 1)


if __name__ == "__main__":
    unittest.main()
