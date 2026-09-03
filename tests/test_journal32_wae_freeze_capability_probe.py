import ast
import importlib.util
import pathlib
import sys
import types
import unittest
from unittest import mock


ROOT = pathlib.Path(__file__).resolve().parents[1]
JOURNAL_PATH = ROOT / "from_git" / "journals" / "32_probe_wae_freeze_capability.py"


def load_journal():
    nxopen = types.ModuleType("NXOpen")
    nxopen_pdm = types.ModuleType("NXOpen.PDM")
    nxopen.PDM = nxopen_pdm
    spec = importlib.util.spec_from_file_location("journal32_test", JOURNAL_PATH)
    module = importlib.util.module_from_spec(spec)
    with mock.patch.dict(sys.modules, {"NXOpen": nxopen, "NXOpen.PDM": nxopen_pdm}):
        spec.loader.exec_module(module)
    return module


class FakeSelectionManager:
    def __init__(self, values):
        self.values = list(values)

    def GetNumSelectedObjects(self):
        return len(self.values)

    def GetSelectedTaggedObject(self, index):
        return self.values[index]


class TestJ32CapabilityProbe(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.journal = load_journal()

    def test_no_selection_uses_only_active_work_part(self):
        work = types.SimpleNamespace(PDMPart=object(), JournalIdentifier="@DB/P1/A")
        session = types.SimpleNamespace(
            Parts=types.SimpleNamespace(Work=work), IsManagedMode=True
        )
        component, part, source = self.journal.resolve_target(
            session, FakeSelectionManager([])
        )
        self.assertIsNone(component)
        self.assertIs(work, part)
        self.assertEqual("ACTIVE_WORK_PART", source)

    def test_one_selection_uses_component_prototype(self):
        part = types.SimpleNamespace(PDMPart=object(), JournalIdentifier="@DB/P2/A")
        component = types.SimpleNamespace(
            Prototype=part, IsSuppressed=False, DisplayName="P2/A"
        )
        session = types.SimpleNamespace(IsManagedMode=True)
        selected, target, source = self.journal.resolve_target(
            session, FakeSelectionManager([component])
        )
        self.assertIs(component, selected)
        self.assertIs(part, target)
        self.assertEqual("ASSEMBLY_NAVIGATOR_SELECTION", source)

    def test_multiple_selections_are_blocked(self):
        part = types.SimpleNamespace(PDMPart=object(), JournalIdentifier="@DB/P2/A")
        component = types.SimpleNamespace(Prototype=part, IsSuppressed=False)
        with self.assertRaisesRegex(RuntimeError, "zero or one"):
            self.journal.resolve_target(
                types.SimpleNamespace(IsManagedMode=True),
                FakeSelectionManager([component, component]),
            )

    def test_candidate_filter_finds_relevant_names(self):
        value = types.SimpleNamespace(
            CheckoutParts=object(), ReleaseStatus=object(), Save=object(), Name="P1"
        )
        names = self.journal.candidate_member_names(value)
        self.assertIn("CheckoutParts", names)
        self.assertIn("ReleaseStatus", names)
        self.assertNotIn("Save", names)
        self.assertNotIn("Name", names)

    def test_candidate_metadata_captures_signature_without_invocation(self):
        calls = []

        def assign_freeze_status(parts, include_secondary=False):
            """Fake freeze signature for read-only metadata testing."""
            calls.append((parts, include_secondary))

        value = types.SimpleNamespace(AssignFreezeStatus=assign_freeze_status)
        rows = self.journal.candidate_member_metadata(value)
        self.assertEqual([], calls)
        self.assertEqual(1, len(rows))
        self.assertEqual("AssignFreezeStatus", rows[0]["name"])
        self.assertTrue(rows[0]["callable"])
        self.assertIn("parts", rows[0]["inspect_signature"])
        self.assertIn("Fake freeze signature", rows[0]["doc"])

    def test_report_declares_every_mutation_guard_false(self):
        report = self.journal.base_report()
        self.assertTrue(report["strictly_read_only"])
        self.assertTrue(report["operations"])
        self.assertFalse(any(report["operations"].values()))

    def test_source_never_calls_mutating_nx_methods(self):
        tree = ast.parse(JOURNAL_PATH.read_text(encoding="utf-8"))
        forbidden = {
            "CheckoutParts",
            "CheckinParts",
            "Save",
            "SetUserAttribute",
            "DeleteUserAttribute",
            "Commit",
            "CreateNewRevision",
            "AddReleaseStatus",
            "RemoveReleaseStatus",
            "Lock",
            "Unlock",
            "Promote",
            "Demote",
        }
        called = {
            node.func.attr
            for node in ast.walk(tree)
            if isinstance(node, ast.Call) and isinstance(node.func, ast.Attribute)
        }
        self.assertFalse(forbidden.intersection(called))


if __name__ == "__main__":
    unittest.main()
