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
    / "29_set_selected_component_bom_exclusion.py"
)


def load_journal():
    nxopen = types.ModuleType("NXOpen")
    nxopen.NXObject = types.SimpleNamespace(
        AttributeType=types.SimpleNamespace(String="String")
    )
    nxopen.Update = types.SimpleNamespace(
        Option=types.SimpleNamespace(Now="Now")
    )
    nxopen.Session = types.SimpleNamespace(
        MarkVisibility=types.SimpleNamespace(Visible="Visible"),
        LibraryUnloadOption=types.SimpleNamespace(AtTermination="AtTermination"),
    )
    prior = sys.modules.get("NXOpen")
    sys.modules["NXOpen"] = nxopen
    try:
        spec = importlib.util.spec_from_file_location("journal29", JOURNAL)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        return module
    finally:
        if prior is None:
            sys.modules.pop("NXOpen", None)
        else:
            sys.modules["NXOpen"] = prior


class FakeAttributeInfo:
    def __init__(self, title, value="", type_value="String", **metadata):
        self.Title = title
        self.StringValue = value
        self.Type = type_value
        self.Unset = metadata.get("unset", False)
        self.Inherited = metadata.get("inherited", False)
        self.IsOverride = metadata.get("is_override", False)
        self.OwnedBySystem = metadata.get("owned_by_system", False)
        self.PdmBased = metadata.get("pdm_based", False)
        self.NotSaved = metadata.get("not_saved", False)


class FakePrototype:
    def __init__(self, name="028060/A", tag=220):
        self.Name = name
        self.Tag = tag


class FakeComponent:
    def __init__(self, name="028060/A", tag=177, attributes=None):
        self.Name = name
        self.DisplayName = name
        self.Tag = tag
        self.Parent = None
        self.Prototype = FakePrototype(name=name + ";1", tag=tag + 10000)
        self.IsSuppressed = False
        self.attributes = dict(attributes or {})
        self.attribute_events = []
        self.set_calls = []
        self.delete_attribute_calls = []
        self.set_error = None
        self.delete_error = None
        self.ignore_set = False
        self.ignore_delete = False
        self.force_wrong_value = None
        self.metadata = {}

    def HasInstanceUserAttribute(self, title, attribute_type, index):
        return title in self.attributes

    def GetInstanceUserAttribute(self, title, attribute_type, index):
        if title not in self.attributes:
            raise RuntimeError("attribute absent")
        metadata = self.metadata.get(title, {})
        return FakeAttributeInfo(title, self.attributes[title], **metadata)

    def SetInstanceUserAttribute(self, title, index, value, update_option):
        self.attribute_events.append(("SET", title))
        self.set_calls.append((title, index, value, update_option))
        if self.set_error:
            raise RuntimeError(self.set_error)
        if not self.ignore_set:
            self.attributes[title] = (
                self.force_wrong_value
                if self.force_wrong_value is not None
                else value
            )

    def DeleteInstanceUserAttribute(
        self, attribute_type, title, delete_entire_array, update_option
    ):
        self.attribute_events.append(("DELETE", title))
        self.delete_attribute_calls.append(
            (attribute_type, title, delete_entire_array, update_option)
        )
        if self.delete_error:
            raise RuntimeError(self.delete_error)
        if not self.ignore_delete:
            self.attributes.pop(title, None)


class FakeRoot:
    def __init__(self, components=None, tag=5000):
        self.Tag = tag
        for component in components or []:
            component.Parent = self


class FakePdmPart:
    def __init__(self, checked=True, owner="aqil"):
        self.checked = checked
        self.owner = owner

    def GetCheckedoutStatusAndUser(self):
        return self.checked, self.owner


class FakePdmSession:
    def __init__(self, user="aqil"):
        self.user = user

    def GetUserName(self):
        return self.user


class FakePart:
    _tag = 9000

    def __init__(self, components, managed=False, read_only=False):
        FakePart._tag += 1
        self.Tag = FakePart._tag
        self.Name = "264MN028171A01/A"
        self.Leaf = self.Name
        self.FullPath = "@DB/264MN028171A01/A" if managed else r"C:\temp\top.prt"
        self.JournalIdentifier = self.FullPath
        self.IsReadOnly = read_only
        self.IsModified = False
        self.components = list(components)
        self.root = FakeRoot(self.components)
        self.ComponentAssembly = types.SimpleNamespace(RootComponent=self.root)
        self.PDMPart = FakePdmPart() if managed else None
        self.attributes = {
            "DB_PART_NO": "264MN028171A01",
            "DB_PART_REV": "A",
        }

    def GetStringAttribute(self, title):
        return self.attributes.get(title, "")


class FakeSession:
    def __init__(
        self, part, components, managed=False, display="SAME", user="aqil",
        mark_error=False, undo_error=False,
    ):
        self.Parts = types.SimpleNamespace(
            Work=part,
            Display=part if display == "SAME" else display,
        )
        self.IsManagedMode = managed
        self.PdmSession = FakePdmSession(user)
        self.components = list(components)
        self.mark_error = mark_error
        self.undo_error = undo_error
        self.set_mark_calls = []
        self.undo_calls = []
        self.delete_calls = []
        self.baselines = None

    def SetUndoMark(self, visibility, name):
        self.set_mark_calls.append((visibility, name))
        if self.mark_error:
            raise RuntimeError("mark failed")
        self.baselines = [dict(component.attributes) for component in self.components]
        return 42

    def UndoToMark(self, mark, name):
        self.undo_calls.append((mark, name))
        if self.undo_error:
            raise RuntimeError("undo failed")
        for component, baseline in zip(self.components, self.baselines):
            component.attributes = dict(baseline)

    def DeleteUndoMark(self, mark, name):
        self.delete_calls.append((mark, name))


class FakeSelectionManager:
    def __init__(self, selected):
        self.selected = list(selected)

    def GetNumSelectedObjects(self):
        return len(self.selected)

    def GetSelectedTaggedObject(self, index):
        return self.selected[index]


class Journal29Tests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.journal = load_journal()
        cls.now = datetime.datetime(
            2026, 8, 29, 23, 0, 0,
            tzinfo=datetime.timezone(datetime.timedelta(hours=8)),
        )

    def make_context(
        self, attributes=None, managed=False, read_only=False,
        components=None, user="aqil",
    ):
        if components is None:
            components = [FakeComponent(attributes=attributes)]
        part = FakePart(components, managed=managed, read_only=read_only)
        session = FakeSession(
            part, components, managed=managed, user=user
        )
        selection = FakeSelectionManager(components)
        return components, part, session, selection

    def run_in_temp(self, session, selection, mode=None, **kwargs):
        with tempfile.TemporaryDirectory() as folder:
            with mock.patch.object(self.journal, "io_root", return_value=folder):
                csv_path, json_path, report = self.journal.run(
                    session, selection, run_datetime=self.now, mode=mode, **kwargs
                )
                self.assertTrue(Path(csv_path).is_file())
                self.assertTrue(Path(json_path).is_file())
                persisted = json.loads(Path(json_path).read_text(encoding="utf-8"))
                self.assertEqual(persisted["verdict"], report["verdict"])
                return report

    def test_apply_is_default_and_writes_exact_attribute(self):
        components, part, session, selection = self.make_context()
        component = components[0]
        report = self.run_in_temp(session, selection)
        self.assertEqual(report["mode"], "APPLY")
        self.assertEqual(report["verdict"]["status"], "APPLIED_VERIFIED")
        self.assertEqual(
            component.set_calls,
            [("CELESTICA_BOM_EXCLUDE_SUBTREE", -1, "YES", "Now")],
        )
        self.assertEqual(session.set_mark_calls, [("Visible", self.journal.UNDO_MARK_NAME)])
        self.assertEqual(session.undo_calls, [])
        self.assertTrue(report["action"]["successful_change_left_undoable"])
        control = report["targets"][0]["after"]["controls"][
            "CELESTICA_BOM_EXCLUDE_SUBTREE"
        ]
        self.assertEqual(control["type"], "STRING")
        self.assertEqual(control["raw_value"], "YES")
        self.assertFalse(control["inherited"])

    def test_explicit_dry_run_preflights_batch_without_write(self):
        components = [FakeComponent(tag=177), FakeComponent(name="028061/A", tag=178)]
        components, part, session, selection = self.make_context(components=components)
        report = self.run_in_temp(session, selection, mode="DRY_RUN")
        self.assertEqual(report["verdict"]["status"], "DRY_RUN_READY")
        self.assertEqual([target["status"] for target in report["targets"]], [
            "DRY_RUN_READY", "DRY_RUN_READY",
        ])
        self.assertTrue(all(component.set_calls == [] for component in components))
        self.assertTrue(
            all(component.delete_attribute_calls == [] for component in components)
        )
        self.assertEqual(session.set_mark_calls, [])

    def test_atomic_batch_applies_every_eligible_component(self):
        components = [FakeComponent(tag=177), FakeComponent(name="028061/A", tag=178)]
        components, part, session, selection = self.make_context(components=components)
        report = self.run_in_temp(session, selection, mode="APPLY")
        self.assertEqual(report["verdict"]["status"], "APPLIED_VERIFIED")
        self.assertEqual(report["action"]["applied_count"], 2)
        self.assertEqual(len(report["targets"]), 2)
        self.assertTrue(all(
            component.attributes.get("CELESTICA_BOM_EXCLUDE_SUBTREE") == "YES"
            for component in components
        ))
        self.assertEqual(len(session.set_mark_calls), 1)

    def test_already_bom_excluded_is_noop_while_other_target_applies(self):
        components = [
            FakeComponent(
                tag=177,
                attributes={"CELESTICA_BOM_EXCLUDE_SUBTREE": "YES"},
            ),
            FakeComponent(name="028061/A", tag=178),
        ]
        components, part, session, selection = self.make_context(components=components)
        report = self.run_in_temp(session, selection, mode="APPLY")
        self.assertEqual(report["verdict"]["status"], "APPLIED_VERIFIED")
        self.assertEqual(components[0].set_calls, [])
        self.assertEqual(len(components[1].set_calls), 1)
        self.assertEqual(report["targets"][0]["status"], "ALREADY_BOM_EXCLUDED")
        self.assertEqual(report["targets"][1]["status"], "APPLIED_VERIFIED")

    def test_all_already_bom_excluded_is_idempotent(self):
        components, part, session, selection = self.make_context(
            {"CELESTICA_BOM_EXCLUDE_SUBTREE": "YES"}
        )
        report = self.run_in_temp(session, selection, mode="APPLY")
        self.assertEqual(report["verdict"]["status"], "ALREADY_BOM_EXCLUDED")
        self.assertEqual(components[0].set_calls, [])
        self.assertEqual(session.set_mark_calls, [])

    def test_all_already_bom_excluded_does_not_require_checkout(self):
        component = FakeComponent(
            attributes={"CELESTICA_BOM_EXCLUDE_SUBTREE": "YES"}
        )
        components, part, session, selection = self.make_context(
            components=[component], managed=True, read_only=True
        )
        part.PDMPart.checked = False
        report = self.run_in_temp(session, selection, mode="APPLY")
        self.assertEqual(report["verdict"]["status"], "ALREADY_BOM_EXCLUDED")
        self.assertEqual(component.set_calls, [])

    def test_one_conflict_blocks_complete_batch(self):
        components = [
            FakeComponent(tag=177),
            FakeComponent(tag=178, attributes={"PLIST_IGNORE_MEMBER": ""}),
        ]
        components, part, session, selection = self.make_context(components=components)
        report = self.run_in_temp(session, selection, mode="APPLY")
        self.assertEqual(report["verdict"]["status"], "BLOCKED_BATCH")
        self.assertEqual(report["targets"][1]["status"], "BLOCKED_CONTROL_CONFLICT")
        self.assertTrue(all(component.set_calls == [] for component in components))
        self.assertEqual(session.set_mark_calls, [])

    def test_nonstandard_or_inherited_bom_exclusion_blocks_batch(self):
        component = FakeComponent(
            attributes={"CELESTICA_BOM_EXCLUDE_SUBTREE": "NO"}
        )
        components, part, session, selection = self.make_context(components=[component])
        report = self.run_in_temp(session, selection, mode="APPLY")
        self.assertEqual(report["verdict"]["status"], "BLOCKED_BATCH")
        self.assertEqual(
            report["targets"][0]["status"],
            "BLOCKED_NONSTANDARD_BOM_EXCLUSION",
        )

        component.attributes["CELESTICA_BOM_EXCLUDE_SUBTREE"] = "YES"
        component.metadata["CELESTICA_BOM_EXCLUDE_SUBTREE"] = {
            "inherited": True
        }
        report = self.run_in_temp(session, selection, mode="APPLY")
        self.assertEqual(report["verdict"]["status"], "BLOCKED_BATCH")

    def test_native_reference_component_is_automatically_unticked(self):
        component = FakeComponent(attributes={"REFERENCE_COMPONENT": ""})
        components, part, session, selection = self.make_context(
            components=[component]
        )
        report = self.run_in_temp(session, selection, mode="APPLY")
        self.assertEqual(report["verdict"]["status"], "APPLIED_VERIFIED")
        self.assertNotIn("REFERENCE_COMPONENT", component.attributes)
        self.assertEqual(
            component.delete_attribute_calls,
            [("String", "REFERENCE_COMPONENT", True, "Now")],
        )
        self.assertEqual(component.attribute_events, [
            ("SET", "CELESTICA_BOM_EXCLUDE_SUBTREE"),
            ("DELETE", "REFERENCE_COMPONENT"),
        ])
        self.assertEqual(
            component.attributes["CELESTICA_BOM_EXCLUDE_SUBTREE"], "YES"
        )
        self.assertTrue(
            report["targets"][0]["before"]["controls"][
                "REFERENCE_COMPONENT"
            ]["present"]
        )
        self.assertFalse(
            report["targets"][0]["after"]["controls"][
                "REFERENCE_COMPONENT"
            ]["present"]
        )
        self.assertEqual(report["action"]["reference_components_removed_count"], 1)

    def test_existing_custom_marker_still_unticks_reference_only(self):
        component = FakeComponent(attributes={
            "CELESTICA_BOM_EXCLUDE_SUBTREE": "YES",
            "REFERENCE_COMPONENT": "",
        })
        components, part, session, selection = self.make_context(
            components=[component]
        )

        report = self.run_in_temp(session, selection, mode="APPLY")

        self.assertEqual(report["verdict"]["status"], "APPLIED_VERIFIED")
        self.assertEqual(component.set_calls, [])
        self.assertEqual(
            component.delete_attribute_calls,
            [("String", "REFERENCE_COMPONENT", True, "Now")],
        )
        self.assertNotIn("REFERENCE_COMPONENT", component.attributes)
        self.assertEqual(
            report["targets"][0]["action"]["status"],
            "UNTICK_REFERENCE_ONLY",
        )

    def test_empty_duplicate_or_noncomponent_selection_blocks(self):
        components, part, session, selection = self.make_context()
        component = components[0]
        report = self.run_in_temp(session, FakeSelectionManager([]), mode="APPLY")
        self.assertEqual(report["verdict"]["status"], "BLOCKED_SELECTION")

        report = self.run_in_temp(
            session, FakeSelectionManager([component, component]), mode="APPLY"
        )
        self.assertEqual(report["verdict"]["status"], "BLOCKED_BATCH")
        self.assertTrue(all(target["status"].startswith("BLOCKED_") or target["status"] == "ELIGIBLE" for target in report["targets"]))

        report = self.run_in_temp(
            session, FakeSelectionManager([types.SimpleNamespace(Tag=123)]), mode="APPLY"
        )
        self.assertEqual(report["verdict"]["status"], "BLOCKED_BATCH")
        self.assertEqual(report["targets"][0]["status"], "BLOCKED_SELECTION")

    def test_selection_limit_blocks_without_writes(self):
        components = [FakeComponent(tag=177), FakeComponent(tag=178)]
        components, part, session, selection = self.make_context(components=components)
        with mock.patch.dict(os.environ, {"NX_J29_MAX_SELECTION": "1"}):
            report = self.run_in_temp(session, selection, mode="APPLY")
        self.assertEqual(report["verdict"]["status"], "BLOCKED_SELECTION_LIMIT")
        self.assertTrue(all(component.set_calls == [] for component in components))

    def test_nested_or_suppressed_target_blocks_complete_batch(self):
        components = [FakeComponent(tag=177), FakeComponent(tag=178)]
        components, part, session, selection = self.make_context(components=components)
        components[1].Parent = FakeRoot(tag=6000)
        report = self.run_in_temp(session, selection, mode="APPLY")
        self.assertEqual(report["verdict"]["status"], "BLOCKED_BATCH")
        self.assertTrue(all(component.set_calls == [] for component in components))

        components[1].Parent = part.root
        components[1].IsSuppressed = True
        report = self.run_in_temp(session, selection, mode="APPLY")
        self.assertEqual(report["verdict"]["status"], "BLOCKED_BATCH")

    def test_apply_requires_proven_write_access(self):
        components, part, session, selection = self.make_context(read_only=True)
        report = self.run_in_temp(session, selection, mode="APPLY")
        self.assertEqual(report["verdict"]["status"], "BLOCKED_WRITE_ACCESS")
        self.assertEqual(components[0].set_calls, [])

        components, part, session, selection = self.make_context(managed=True)
        part.PDMPart.checked = False
        report = self.run_in_temp(session, selection, mode="APPLY")
        self.assertEqual(report["verdict"]["status"], "BLOCKED_WRITE_ACCESS")
        self.assertIn("never performs checkout", report["verdict"]["message"])

    def test_runtime_owner_format_matches_same_teamcenter_identifier(self):
        user_id = "99946e1828964542b86c86d6c2cf3cbe"
        components, part, session, selection = self.make_context(
            managed=True, user=user_id
        )
        part.PDMPart.owner = "aqil ameran ({0})".format(user_id)
        report = self.run_in_temp(session, selection, mode="APPLY")
        self.assertEqual(report["verdict"]["status"], "APPLIED_VERIFIED")
        self.assertTrue(report["access"]["owner_is_current_user"])

    def test_different_teamcenter_identifier_still_blocks(self):
        components, part, session, selection = self.make_context(
            managed=True, user="99946e1828964542b86c86d6c2cf3cbe"
        )
        part.PDMPart.owner = "other user (aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa)"
        report = self.run_in_temp(session, selection, mode="APPLY")
        self.assertEqual(report["verdict"]["status"], "BLOCKED_WRITE_ACCESS")

    def test_managed_checkout_enum_shape_is_accepted_for_current_user(self):
        components, part, session, selection = self.make_context(managed=True)
        part.PDMPart.GetCheckedoutStatusAndUser = lambda: (
            types.SimpleNamespace(name="CheckedOut"), "aqil"
        )
        report = self.run_in_temp(session, selection, mode="APPLY")
        self.assertEqual(report["verdict"]["status"], "APPLIED_VERIFIED")

    def test_second_target_verification_failure_rolls_back_complete_batch(self):
        components = [
            FakeComponent(tag=177), FakeComponent(tag=178), FakeComponent(tag=179)
        ]
        components[1].force_wrong_value = "NO"
        components, part, session, selection = self.make_context(components=components)
        report = self.run_in_temp(session, selection, mode="APPLY")
        self.assertEqual(report["verdict"]["status"], "ROLLED_BACK")
        self.assertTrue(all(
            "CELESTICA_BOM_EXCLUDE_SUBTREE" not in component.attributes
            for component in components
        ))
        self.assertEqual(len(session.undo_calls), 1)
        self.assertEqual(len(session.delete_calls), 1)
        self.assertEqual(report["rollback"]["status"], "ROLLED_BACK")
        self.assertEqual(report["targets"][2]["status"], "NOT_ATTEMPTED")

    def test_write_exception_rolls_back_complete_batch(self):
        components = [FakeComponent(tag=177), FakeComponent(tag=178)]
        components[1].set_error = "NX write rejected"
        components, part, session, selection = self.make_context(components=components)
        report = self.run_in_temp(session, selection, mode="APPLY")
        self.assertEqual(report["verdict"]["status"], "ROLLED_BACK")
        self.assertTrue(all(
            "CELESTICA_BOM_EXCLUDE_SUBTREE" not in component.attributes
            for component in components
        ))

    def test_reference_untick_failure_rolls_back_custom_marker(self):
        component = FakeComponent(attributes={"REFERENCE_COMPONENT": ""})
        component.delete_error = "NX reference untick rejected"
        components, part, session, selection = self.make_context(
            components=[component]
        )

        report = self.run_in_temp(session, selection, mode="APPLY")

        self.assertEqual(report["verdict"]["status"], "ROLLED_BACK")
        self.assertEqual(component.attributes, {"REFERENCE_COMPONENT": ""})
        self.assertEqual(len(session.undo_calls), 1)

    def test_reference_untick_readback_failure_rolls_back(self):
        component = FakeComponent(attributes={"REFERENCE_COMPONENT": ""})
        component.ignore_delete = True
        components, part, session, selection = self.make_context(
            components=[component]
        )

        report = self.run_in_temp(session, selection, mode="APPLY")

        self.assertEqual(report["verdict"]["status"], "ROLLED_BACK")
        self.assertEqual(component.attributes, {"REFERENCE_COMPONENT": ""})
        self.assertIn(
            "REFERENCE_COMPONENT remains present",
            report["targets"][0]["action"]["verification_errors"][0],
        )

    def test_context_must_be_same_work_and_display_assembly(self):
        components, part, session, selection = self.make_context()
        session.Parts.Display = types.SimpleNamespace(Tag=999)
        report = self.run_in_temp(session, selection, mode="APPLY")
        self.assertEqual(report["verdict"]["status"], "FAILED_CONTEXT")
        self.assertEqual(components[0].set_calls, [])

    def test_unload_is_deferred_to_nx_termination(self):
        self.assertEqual(self.journal.get_unload_option(None), "AtTermination")

    def test_source_contains_no_load_checkout_save_or_checkin_calls(self):
        source = JOURNAL.read_text(encoding="utf-8")
        forbidden = (
            ".Save(", ".SaveAs(", ".OpenBase(", ".OpenDisplay(",
            ".LoadFully(", ".LoadThisPartFully(", ".Checkout",
            ".CheckOut", ".Checkin", ".CheckIn",
        )
        for token in forbidden:
            self.assertNotIn(token, source)

    def test_source_only_removes_native_reference_component(self):
        source = JOURNAL.read_text(encoding="utf-8")
        self.assertNotIn(
            "SetInstanceUserAttribute(\n                    REFERENCE_ATTRIBUTE",
            source,
        )
        self.assertIn("component.DeleteInstanceUserAttribute(", source)
        self.assertIn("REFERENCE_ATTRIBUTE,", source)
        self.assertIn("automatically unticks native Reference-Only", source)

    def test_configured_mode_supports_environment_override(self):
        with mock.patch.dict(os.environ, {}, clear=True):
            self.assertEqual(self.journal.configured_mode(), "APPLY")
        with mock.patch.dict(os.environ, {"NX_J29_MODE": "apply"}):
            self.assertEqual(self.journal.configured_mode(), "APPLY")
        with mock.patch.dict(os.environ, {"NX_J29_MODE": "bad"}):
            with self.assertRaises(RuntimeError):
                self.journal.configured_mode()

    def test_configured_selection_limit(self):
        with mock.patch.dict(os.environ, {}, clear=True):
            self.assertEqual(
                self.journal.configured_max_selection(),
                self.journal.DEFAULT_MAX_SELECTION,
            )
        with mock.patch.dict(os.environ, {"NX_J29_MAX_SELECTION": "25"}):
            self.assertEqual(self.journal.configured_max_selection(), 25)
        with mock.patch.dict(os.environ, {"NX_J29_MAX_SELECTION": "0"}):
            with self.assertRaises(RuntimeError):
                self.journal.configured_max_selection()


if __name__ == "__main__":
    unittest.main()
