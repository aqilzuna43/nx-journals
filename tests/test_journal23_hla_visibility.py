import datetime
import importlib.util
import json
import os
import sys
import tempfile
import types
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
JOURNAL = ROOT / "from_git" / "journals" / "23_diagnose_hla_visibility.py"


def load_journal():
    nxopen = types.ModuleType("NXOpen")
    nxopen.NXObject = types.SimpleNamespace(
        AttributeType=types.SimpleNamespace(String="String")
    )
    nxopen.Session = types.SimpleNamespace(
        LibraryUnloadOption=types.SimpleNamespace(Immediately="Immediately")
    )
    prior = sys.modules.get("NXOpen")
    sys.modules["NXOpen"] = nxopen
    try:
        spec = importlib.util.spec_from_file_location("journal23", JOURNAL)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        return module
    finally:
        if prior is None:
            sys.modules.pop("NXOpen", None)
        else:
            sys.modules["NXOpen"] = prior


class State:
    def __init__(self, name):
        self.name = name

    def __str__(self):
        return self.name


class Named:
    def __init__(self, name):
        self.Name = name


class FakeLayers:
    def __init__(self, states=None):
        self.states = dict(states or {})

    def GetState(self, layer):
        return State(self.states.get(layer, "Selectable"))


class FakeBody:
    def __init__(self, tag, layer=1, blanked=False, solid=True):
        self.Tag = tag
        self.Layer = layer
        self.IsBlanked = blanked
        self.IsSolidBody = solid


class FakeBodyOccurrence:
    def __init__(self, tag, blanked=False):
        self.Tag = tag
        self.IsBlanked = blanked


class FakeReferenceSet:
    def __init__(self, name, members):
        self.Name = name
        self.members = list(members)

    def AskMembersInReferenceSet(self):
        return list(self.members)


class FakePrototype:
    _next_tag = 1000

    def __init__(
        self,
        name,
        bodies=None,
        reference_sets=None,
        fully_loaded=True,
        part_number=None,
    ):
        FakePrototype._next_tag += 1
        self.Tag = FakePrototype._next_tag
        self.Name = name
        self.Leaf = name
        self.Bodies = list(bodies or [])
        self.Layers = FakeLayers()
        self.PartLoadState = State("FullyLoaded" if fully_loaded else "MinimallyLoaded")
        self.IsFullyLoaded = fully_loaded
        self.EntirePartRefsetName = "Entire Part"
        self.EmptyPartRefsetName = "Empty"
        self._reference_sets = list(reference_sets or [])
        self.part_number = part_number or name
        root = types.SimpleNamespace(GetChildren=lambda: [])
        self.ComponentAssembly = types.SimpleNamespace(RootComponent=root)

    def GetAllReferenceSets(self):
        return list(self._reference_sets)

    def GetStringAttribute(self, name):
        if name == "DB_PART_NO":
            return self.part_number
        if name == "DB_PART_REV":
            return "A"
        raise RuntimeError("attribute missing")

    def GetUserAttribute(self, *args):
        raise RuntimeError("attribute missing")


class FakeComponent:
    _next_tag = 2000

    def __init__(
        self,
        name,
        prototype,
        children=None,
        layer=1,
        blanked=False,
        suppressed=False,
        non_geometric=False,
        representation="Exact",
        reference_set="MODEL",
        occurrence_map=None,
    ):
        FakeComponent._next_tag += 1
        self.Tag = FakeComponent._next_tag
        self.DisplayName = name
        self.Name = name
        self.Prototype = prototype
        self.Layer = layer
        self.IsBlanked = blanked
        self.IsSuppressed = suppressed
        self.ReferenceSet = reference_set
        self.SuppressingArrangement = None
        self.UsedArrangement = None
        self.non_geometric = non_geometric
        self.representation = representation
        self.children = list(children or [])
        self.occurrence_map = dict(occurrence_map or {})

    def GetChildren(self):
        return list(self.children)

    def GetComponentRepresentationMode(self):
        return State(self.representation)

    def FindOccurrence(self, member):
        return self.occurrence_map.get(member)


class FakeAssembly:
    def __init__(self, children):
        self.RootComponent = types.SimpleNamespace(GetChildren=lambda: list(children))
        self.ActiveArrangement = Named("Arrangement 1")

    def GetSuppressedState(self, component, controlled):
        return State("Suppressed" if component.IsSuppressed else "Unsuppressed")

    def GetSuppressionExpression(self, component):
        raise RuntimeError("not suppression controlled")

    def GetNonGeometricState(self, component):
        return component.non_geometric


class FakeView:
    def __init__(self, visible):
        self.Name = "Trimetric"
        self.visible = list(visible)

    def AskVisibleObjects(self):
        return list(self.visible)


class FakeDynamicSections:
    def ToArray(self):
        return []


class FakeWorkPart:
    def __init__(self, children, visible=None, layer_states=None):
        self.Name = "TOP-HLA"
        self.ComponentAssembly = FakeAssembly(children)
        self.Layers = FakeLayers(layer_states)
        self.ModelingViews = types.SimpleNamespace(WorkView=FakeView(visible or []))
        self.DynamicSections = FakeDynamicSections()

    def GetStringAttribute(self, name):
        if name == "DB_PART_NO":
            return "TOP-HLA"
        raise RuntimeError("attribute missing")

    def GetUserAttribute(self, *args):
        raise RuntimeError("attribute missing")


def healthy_component(name="VISIBLE"):
    body = FakeBody(10)
    occurrence = FakeBodyOccurrence(20)
    reference_set = FakeReferenceSet("MODEL", [body])
    prototype = FakePrototype(name, [body], [reference_set])
    component = FakeComponent(name, prototype, occurrence_map={body: occurrence})
    return component, occurrence


class Journal23Tests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.journal = load_journal()

    def test_source_is_read_only_and_has_nx_unload_hook(self):
        source = JOURNAL.read_text(encoding="utf-8")
        forbidden = (
            ".Blank(",
            ".Unblank(",
            ".Suppress(",
            ".Unsuppress(",
            ".ReplaceReferenceSet(",
            ".SetState(",
            ".Save(",
            ".LoadFully(",
        )
        for token in forbidden:
            self.assertNotIn(token, source)
        self.assertIn("def get_unload_option(dummy):", source)
        self.assertIn("LibraryUnloadOption.Immediately", source)

    def test_healthy_exact_occurrence_has_no_direct_cause(self):
        component, occurrence = healthy_component()
        work = FakeWorkPart([component], visible=[occurrence])
        view = self.journal.work_view_snapshot(work)
        sections = self.journal.dynamic_section_snapshot(work)
        rows, errors = self.journal.collect_records(
            work, "2026-08-14T10:00:00+08:00", "", view, sections
        )
        self.assertEqual([], errors)
        self.assertEqual(1, len(rows))
        self.assertEqual("NO_DIRECT_CAUSE_FOUND", rows[0]["ISSUE_CODES"])
        self.assertEqual("LOW", rows[0]["CONFIDENCE"])
        self.assertEqual(1, rows[0]["OCCURRENCE_MEMBERS_VISIBLE_IN_WORK_VIEW"])
        self.assertNotIn("SUPPRESSED_CURRENT_ARRANGEMENT", rows[0]["ISSUE_CODES"])

    def test_hidden_hla_component_layer_is_ranked_high(self):
        component, occurrence = healthy_component("HIDDEN-LAYER")
        component.Layer = 77
        work = FakeWorkPart([component], visible=[], layer_states={77: "Hidden"})
        view = self.journal.work_view_snapshot(work)
        rows, _ = self.journal.collect_records(
            work,
            "2026-08-14T10:00:00+08:00",
            "",
            view,
            self.journal.dynamic_section_snapshot(work),
        )
        self.assertEqual("COMPONENT_LAYER_HIDDEN", rows[0]["ISSUE_CODES"].split(" | ")[0])
        self.assertEqual("HIGH", rows[0]["CONFIDENCE"])
        self.assertIn("hidden layer in the top-level assembly", rows[0]["ROOT_CAUSE"])

    def test_blanked_parent_is_inherited_by_child(self):
        child, child_occurrence = healthy_component("MISSING-CHILD")
        parent_body = FakeBody(30)
        parent_occurrence = FakeBodyOccurrence(31)
        parent_refset = FakeReferenceSet("MODEL", [parent_body])
        parent_proto = FakePrototype("SUBASM", [parent_body], [parent_refset])
        parent = FakeComponent(
            "BLANKED-PARENT",
            parent_proto,
            children=[child],
            blanked=True,
            occurrence_map={parent_body: parent_occurrence},
        )
        work = FakeWorkPart([parent], visible=[child_occurrence])
        view = self.journal.work_view_snapshot(work)
        rows, _ = self.journal.collect_records(
            work,
            "2026-08-14T10:00:00+08:00",
            "MISSING-CHILD",
            view,
            self.journal.dynamic_section_snapshot(work),
        )
        child_row = rows[1]
        self.assertEqual("YES", child_row["ANCESTOR_BLANKED"])
        self.assertEqual("ANCESTOR_BLANKED", child_row["ISSUE_CODES"].split(" | ")[0])
        self.assertEqual("YES", child_row["TARGET_MATCH"])

    def test_empty_and_missing_reference_sets_are_distinguished(self):
        body = FakeBody(40)
        prototype = FakePrototype("P", [body], [FakeReferenceSet("MODEL", [body])])
        empty = FakeComponent("EMPTY", prototype, reference_set="Empty")
        stale = FakeComponent("STALE", prototype, reference_set="OLD_MODEL")
        work = FakeWorkPart([empty, stale])
        view = self.journal.work_view_snapshot(work)
        rows, _ = self.journal.collect_records(
            work,
            "2026-08-14T10:00:00+08:00",
            "",
            view,
            self.journal.dynamic_section_snapshot(work),
        )
        self.assertEqual("EMPTY_REFERENCE_SET", rows[0]["ISSUE_CODES"].split(" | ")[0])
        self.assertEqual("REFERENCE_SET_NOT_FOUND", rows[1]["ISSUE_CODES"].split(" | ")[0])

    def test_target_filter_matches_prototype_part_number(self):
        body = FakeBody(50)
        occurrence = FakeBodyOccurrence(51)
        reference_set = FakeReferenceSet("MODEL", [body])
        prototype = FakePrototype(
            "GENERIC-NAME",
            [body],
            [reference_set],
            part_number="264MN099999A01",
        )
        component = FakeComponent(
            "OCCURRENCE-7", prototype, occurrence_map={body: occurrence}
        )
        work = FakeWorkPart([component], visible=[occurrence])
        view = self.journal.work_view_snapshot(work)
        rows, _ = self.journal.collect_records(
            work,
            "2026-08-14T10:00:00+08:00",
            "264MN099999A01",
            view,
            self.journal.dynamic_section_snapshot(work),
        )
        self.assertEqual("YES", rows[0]["TARGET_MATCH"])

    def test_run_writes_ranked_csv_and_json_without_changing_part(self):
        component, occurrence = healthy_component("TARGET-264")
        work = FakeWorkPart([component], visible=[occurrence])
        session = types.SimpleNamespace(
            Parts=types.SimpleNamespace(Work=work, Display=work)
        )
        with tempfile.TemporaryDirectory() as folder:
            old_io = os.environ.get("NX_JOURNALS_IO_DIR")
            old_target = os.environ.get("NX_J23_TARGET")
            os.environ["NX_JOURNALS_IO_DIR"] = folder
            os.environ["NX_J23_TARGET"] = "TARGET-264"
            try:
                csv_path, json_path, report = self.journal.run(
                    session,
                    datetime.datetime(2026, 8, 14, 10, 30, 0),
                )
            finally:
                if old_io is None:
                    os.environ.pop("NX_JOURNALS_IO_DIR", None)
                else:
                    os.environ["NX_JOURNALS_IO_DIR"] = old_io
                if old_target is None:
                    os.environ.pop("NX_J23_TARGET", None)
                else:
                    os.environ["NX_J23_TARGET"] = old_target
            self.assertTrue(Path(csv_path).is_file())
            self.assertTrue(Path(json_path).is_file())
            payload = json.loads(Path(json_path).read_text(encoding="utf-8"))
            self.assertEqual(self.journal.BUILD, payload["journal_build"])
            self.assertEqual(1, payload["target_match_count"])
            self.assertEqual("READ_ONLY_HLA_VISIBILITY_DIAGNOSTIC", payload["scope"])
            self.assertEqual(report["ranked_occurrences"], payload["ranked_occurrences"])

    def test_run_requires_same_work_and_display_hla(self):
        component, occurrence = healthy_component()
        work = FakeWorkPart([component], visible=[occurrence])
        other = FakeWorkPart([], visible=[])
        session = types.SimpleNamespace(
            Parts=types.SimpleNamespace(Work=work, Display=other)
        )
        with self.assertRaisesRegex(RuntimeError, "both the displayed part and work part"):
            self.journal.run(session)


if __name__ == "__main__":
    unittest.main()
