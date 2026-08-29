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
    / "29_set_selected_component_reference_only.py"
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
        self.Prototype = FakePrototype()
        self.IsSuppressed = False
        self.attributes = dict(attributes or {})
        self.set_calls = []
        self.set_error = None
        self.ignore_set = False
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
        self.set_calls.append((title, index, value, update_option))
        if self.set_error:
            raise RuntimeError(self.set_error)
        if not self.ignore_set:
            self.attributes[title] = (
                self.force_wrong_value
                if self.force_wrong_value is not None
                else value
            )


class FakeRoot:
    def __init__(self, component=None, tag=5000):
        self.Tag = tag
        if component is not None:
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

    def __init__(self, component, managed=False, read_only=False):
        FakePart._tag += 1
        self.Tag = FakePart._tag
        self.Name = "264MN028171A01/A"
        self.Leaf = self.Name
        self.FullPath = "@DB/264MN028171A01/A" if managed else r"C:\temp\top.prt"
        self.JournalIdentifier = self.FullPath
        self.IsReadOnly = read_only
        self.IsModified = False
        self.root = FakeRoot(component)
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
        self, part, component, managed=False, display="SAME", user="aqil",
        mark_error=False, undo_error=False,
    ):
        self.Parts = types.SimpleNamespace(
            Work=part,
            Display=part if display == "SAME" else display,
        )
        self.IsManagedMode = managed
        self.PdmSession = FakePdmSession(user)
        self.component = component
        self.mark_error = mark_error
        self.undo_error = undo_error
        self.set_mark_calls = []
        self.undo_calls = []
        self.delete_calls = []
        self.baseline = None

    def SetUndoMark(self, visibility, name):
        self.set_mark_calls.append((visibility, name))
        if self.mark_error:
            raise RuntimeError("mark failed")
        self.baseline = dict(self.component.attributes)
        return 42

    def UndoToMark(self, mark, name):
        self.undo_calls.append((mark, name))
        if self.undo_error:
            raise RuntimeError("undo failed")
        self.component.attributes = dict(self.baseline)

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

    def make_context(self, attributes=None, managed=False, read_only=False):
        component = FakeComponent(attributes=attributes)
        part = FakePart(component, managed=managed, read_only=read_only)
        session = FakeSession(part, component, managed=managed)
        selection = FakeSelectionManager([component])
        return component, part, session, selection

    def run_in_temp(self, session, selection, mode="DRY_RUN", **kwargs):
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

    def test_dry_run_is_default_and_does_not_write(self):
        component, part, session, selection = self.make_context()
        report = self.run_in_temp(session, selection)
        self.assertEqual(report["verdict"]["status"], "DRY_RUN_READY")
        self.assertEqual(component.set_calls, [])
        self.assertEqual(session.set_mark_calls, [])
        self.assertNotIn("REFERENCE_COMPONENT", component.attributes)
        self.assertFalse(report["configuration"]["force_load"])
        self.assertFalse(report["configuration"]["automatic_save"])

    def test_apply_writes_exact_blank_occurrence_attribute_and_leaves_undo(self):
        component, part, session, selection = self.make_context()
        report = self.run_in_temp(session, selection, mode="APPLY")
        self.assertEqual(report["verdict"]["status"], "APPLIED_VERIFIED")
        self.assertEqual(
            component.set_calls,
            [("REFERENCE_COMPONENT", -1, "", "Now")],
        )
        self.assertEqual(session.set_mark_calls, [("Visible", self.journal.UNDO_MARK_NAME)])
        self.assertEqual(session.undo_calls, [])
        self.assertEqual(session.delete_calls, [])
        self.assertTrue(report["action"]["successful_change_left_undoable"])
        control = report["after"]["controls"]["REFERENCE_COMPONENT"]
        self.assertEqual(control["type"], "STRING")
        self.assertEqual(control["raw_value"], "")
        self.assertFalse(control["inherited"])

    def test_already_reference_only_is_idempotent(self):
        component, part, session, selection = self.make_context(
            {"REFERENCE_COMPONENT": ""}
        )
        report = self.run_in_temp(session, selection, mode="APPLY")
        self.assertEqual(report["verdict"]["status"], "ALREADY_REFERENCE_ONLY")
        self.assertEqual(component.set_calls, [])
        self.assertEqual(session.set_mark_calls, [])

    def test_conflicting_plist_control_blocks(self):
        component, part, session, selection = self.make_context(
            {"PLIST_IGNORE_MEMBER": ""}
        )
        report = self.run_in_temp(session, selection, mode="APPLY")
        self.assertEqual(report["verdict"]["status"], "BLOCKED_CONTROL_CONFLICT")
        self.assertEqual(component.set_calls, [])

    def test_nonstandard_existing_reference_blocks(self):
        component, part, session, selection = self.make_context(
            {"REFERENCE_COMPONENT": "YES"}
        )
        report = self.run_in_temp(session, selection, mode="APPLY")
        self.assertEqual(report["verdict"]["status"], "BLOCKED_NONSTANDARD_REFERENCE")
        self.assertEqual(component.set_calls, [])

    def test_inherited_reference_blocks(self):
        component, part, session, selection = self.make_context(
            {"REFERENCE_COMPONENT": ""}
        )
        component.metadata["REFERENCE_COMPONENT"] = {"inherited": True}
        report = self.run_in_temp(session, selection, mode="APPLY")
        self.assertEqual(report["verdict"]["status"], "BLOCKED_NONSTANDARD_REFERENCE")

    def test_selection_must_be_exactly_one_component(self):
        component, part, session, selection = self.make_context()
        report = self.run_in_temp(
            session, FakeSelectionManager([]), mode="APPLY"
        )
        self.assertEqual(report["verdict"]["status"], "BLOCKED_SELECTION")
        self.assertEqual(component.set_calls, [])

        report = self.run_in_temp(
            session, FakeSelectionManager([component, component]), mode="APPLY"
        )
        self.assertEqual(report["verdict"]["status"], "BLOCKED_SELECTION")

        report = self.run_in_temp(
            session, FakeSelectionManager([types.SimpleNamespace(Tag=123)]), mode="APPLY"
        )
        self.assertEqual(report["verdict"]["status"], "BLOCKED_SELECTION")

    def test_nested_or_suppressed_component_blocks(self):
        component, part, session, selection = self.make_context()
        component.Parent = FakeRoot(tag=6000)
        report = self.run_in_temp(session, selection, mode="APPLY")
        self.assertEqual(report["verdict"]["status"], "BLOCKED_SELECTION")
        self.assertEqual(component.set_calls, [])

        component.Parent = part.root
        component.IsSuppressed = True
        report = self.run_in_temp(session, selection, mode="APPLY")
        self.assertEqual(report["verdict"]["status"], "BLOCKED_SELECTION")

    def test_apply_requires_proven_write_access(self):
        component, part, session, selection = self.make_context(read_only=True)
        report = self.run_in_temp(session, selection, mode="APPLY")
        self.assertEqual(report["verdict"]["status"], "BLOCKED_WRITE_ACCESS")
        self.assertEqual(component.set_calls, [])

        component, part, session, selection = self.make_context(managed=True)
        part.PDMPart.checked = False
        report = self.run_in_temp(session, selection, mode="APPLY")
        self.assertEqual(report["verdict"]["status"], "BLOCKED_WRITE_ACCESS")
        self.assertIn("never performs checkout", report["verdict"]["message"])

    def test_managed_checkout_enum_shape_is_accepted_for_current_user(self):
        component, part, session, selection = self.make_context(managed=True)
        part.PDMPart.GetCheckedoutStatusAndUser = lambda: (
            types.SimpleNamespace(name="CheckedOut"), "aqil"
        )
        report = self.run_in_temp(session, selection, mode="APPLY")
        self.assertEqual(report["verdict"]["status"], "APPLIED_VERIFIED")

    def test_verification_failure_rolls_back_and_proves_absent_baseline(self):
        component, part, session, selection = self.make_context()
        component.force_wrong_value = "YES"
        report = self.run_in_temp(session, selection, mode="APPLY")
        self.assertEqual(report["verdict"]["status"], "ROLLED_BACK")
        self.assertNotIn("REFERENCE_COMPONENT", component.attributes)
        self.assertEqual(len(session.undo_calls), 1)
        self.assertEqual(len(session.delete_calls), 1)
        self.assertEqual(report["rollback"]["status"], "ROLLED_BACK")

    def test_write_exception_rolls_back(self):
        component, part, session, selection = self.make_context()
        component.set_error = "NX write rejected"
        report = self.run_in_temp(session, selection, mode="APPLY")
        self.assertEqual(report["verdict"]["status"], "ROLLED_BACK")
        self.assertEqual(len(session.undo_calls), 1)

    def test_context_must_be_same_work_and_display_assembly(self):
        component, part, session, selection = self.make_context()
        session.Parts.Display = types.SimpleNamespace(Tag=999)
        report = self.run_in_temp(session, selection, mode="APPLY")
        self.assertEqual(report["verdict"]["status"], "FAILED_CONTEXT")
        self.assertEqual(component.set_calls, [])

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

    def test_configured_mode_supports_environment_override(self):
        with mock.patch.dict(os.environ, {"NX_J29_MODE": "apply"}):
            self.assertEqual(self.journal.configured_mode(), "APPLY")
        with mock.patch.dict(os.environ, {"NX_J29_MODE": "bad"}):
            with self.assertRaises(RuntimeError):
                self.journal.configured_mode()


if __name__ == "__main__":
    unittest.main()
