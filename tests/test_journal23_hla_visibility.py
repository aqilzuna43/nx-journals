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
JOURNAL24 = ROOT / "from_git" / "journals" / "24_repair_hla_isolate_visibility.py"
NX_V1_ARTIFACT = (
    ROOT
    / "docs"
    / "J23_HLA_VISIBILITY_264MN024625A01_20260814_115209.json"
)


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


def load_journal24():
    nxopen = types.ModuleType("NXOpen")
    nxopen.NXObject = types.SimpleNamespace(
        AttributeType=types.SimpleNamespace(String="String")
    )
    nxopen.Session = types.SimpleNamespace(
        LibraryUnloadOption=types.SimpleNamespace(Immediately="Immediately"),
        MarkVisibility=types.SimpleNamespace(Visible="Visible"),
    )
    prior = sys.modules.get("NXOpen")
    sys.modules["NXOpen"] = nxopen
    try:
        spec = importlib.util.spec_from_file_location("journal24", JOURNAL24)
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
        prototype_children=None,
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
        root = types.SimpleNamespace(GetChildren=lambda: list(prototype_children or []))
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

    def GetNonGeometricState(self):
        return self.non_geometric

    def FindOccurrence(self, member):
        return self.occurrence_map.get(member)


class FakeAssembly:
    def __init__(self, children):
        self.RootComponent = types.SimpleNamespace(GetChildren=lambda: list(children))
        self.ActiveArrangement = Named("Arrangement 1")
        self.show_calls = []
        self.show_exception = None
        self.show_errors = []

    def GetSuppressedState(self, component, controlled):
        return State("Suppressed" if component.IsSuppressed else "Unsuppressed")

    def GetSuppressionExpression(self, component):
        raise RuntimeError("not suppression controlled")

    def GetNonGeometricState(self, component):
        return component.non_geometric

    def ShowComponentsInIsolateView(self, components, view):
        self.show_calls.append((list(components), view))
        if self.show_exception is not None:
            raise self.show_exception
        for component in components:
            for occurrence in component.occurrence_map.values():
                if isinstance(occurrence, FakeBodyOccurrence) and occurrence not in view.visible:
                    view.visible.append(occurrence)
        return FakeErrorList(self.show_errors)


class FakeErrorInfo:
    def __init__(self, code, description):
        self.ErrorCode = code
        self.Description = description
        self.ErrorObject = None
        self.ErrorObjectDescription = ""


class FakeErrorList:
    def __init__(self, errors=None):
        self.errors = list(errors or [])
        self.Length = len(self.errors)
        self.freed = False

    def GetErrorInfo(self, index):
        return self.errors[index]

    def FreeResource(self):
        self.freed = True


class FakeView:
    _next_tag = 3000

    def __init__(self, visible, name="Trimetric", visible_sections=None):
        FakeView._next_tag += 1
        self.Tag = FakeView._next_tag
        self.Name = name
        self.visible = list(visible)
        self.visible_sections = set(visible_sections or [])

    def AskVisibleObjects(self):
        return list(self.visible)

    def IsDynamicSectionVisible(self, section):
        return section in self.visible_sections

    def Regenerate(self):
        return None


class FakeModelingViews:
    def __init__(self, work_view, others=None):
        self.WorkView = work_view
        self.views = [work_view] + list(others or [])

    def ToArray(self):
        return list(self.views)


class FakeDynamicSections:
    def ToArray(self):
        return []


class FakeWorkPart:
    def __init__(
        self,
        children,
        visible=None,
        layer_states=None,
        view_name="Trimetric",
        other_views=None,
    ):
        self.Name = "TOP-HLA"
        self.ComponentAssembly = FakeAssembly(children)
        self.Layers = FakeLayers(layer_states)
        work_view = FakeView(visible or [], name=view_name)
        self.ModelingViews = FakeModelingViews(work_view, other_views)
        self.DynamicSections = FakeDynamicSections()
        self.Views = types.SimpleNamespace(UpdateDisplay=lambda: None)

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

    def test_failed_and_missing_probes_never_become_false(self):
        class Broken:
            @property
            def IsBlanked(self):
                raise RuntimeError("NX probe failed")

        error = self.journal.property_probe(Broken(), "IsBlanked")
        missing = self.journal.property_probe(Broken(), "IsSuppressed")
        self.assertEqual(self.journal.ERROR, error["status"])
        self.assertIsNone(error["value"])
        self.assertEqual(self.journal.UNAVAILABLE, missing["status"])
        self.assertIsNone(missing["value"])

    def test_exact_target_view_exclusion_is_confirmed_by_other_view(self):
        component, occurrence = healthy_component("TARGET-264")
        alternate = FakeView([occurrence], name="Trimetric")
        work = FakeWorkPart(
            [component], visible=[], view_name="Isolate", other_views=[alternate]
        )
        nodes, errors = self.journal.collect_nodes(work)
        self.assertEqual([], errors)
        analysis = self.journal.analyze_target(nodes[0], nodes, work)
        self.assertEqual("CONFIRMED", analysis["conclusion"]["status"])
        self.assertEqual(
            "CURRENT_WORK_VIEW_EXCLUSION",
            analysis["conclusion"]["root_cause_code"],
        )
        verdicts = {item["code"]: item["verdict"] for item in analysis["hypotheses"]}
        self.assertEqual("RULED_OUT", verdicts["SUPPRESSION_AS_PRIMARY_CAUSE"])
        self.assertEqual("RULED_OUT", verdicts["BLANKING_AS_PRIMARY_CAUSE"])
        self.assertEqual("RULED_OUT", verdicts["REFERENCE_SET_AS_PRIMARY_CAUSE"])
        self.assertEqual("STRONGLY_SUPPORTED", verdicts["ISOLATE_VIEW_MECHANISM"])

    def test_entire_part_subassembly_maps_children_and_descendant_bodies(self):
        child_body = FakeBody(61)
        child_occ_body = FakeBodyOccurrence(62)
        child_refset = FakeReferenceSet("MODEL", [child_body])
        child_proto = FakePrototype("CHILD", [child_body], [child_refset])
        child_occ = FakeComponent(
            "CHILD/A", child_proto, occurrence_map={child_body: child_occ_body}
        )
        prototype_child = FakeComponent("CHILD-PROTOTYPE", child_proto)
        target_proto = FakePrototype(
            "TARGET-SUBASM",
            [],
            [],
            part_number="264MN031978A01",
            prototype_children=[prototype_child],
        )
        target = FakeComponent(
            "264MN031978A01/A",
            target_proto,
            children=[child_occ],
            reference_set="Entire Part",
            occurrence_map={prototype_child: child_occ},
        )
        alternate = FakeView([child_occ_body], name="Trimetric")
        work = FakeWorkPart(
            [target], visible=[], view_name="Isolate", other_views=[alternate]
        )
        nodes, _ = self.journal.collect_nodes(work)
        analysis = self.journal.analyze_target(nodes[0], nodes, work)
        target_row = analysis["subtree_occurrences"][0]
        self.assertEqual("ENTIRE_PART", target_row["reference_set"]["kind"])
        self.assertEqual(1, target_row["reference_set"]["component_member_count"])
        self.assertEqual(1, target_row["mapping"]["mapped_component_count"])
        self.assertEqual(1, analysis["subtree_summary"]["mapped_body_occurrences"])

    def test_hidden_hla_layer_is_observed_not_assumed(self):
        component, occurrence = healthy_component("HIDDEN-LAYER")
        component.Layer = 77
        work = FakeWorkPart([component], visible=[], layer_states={77: "Hidden"})
        nodes, _ = self.journal.collect_nodes(work)
        analysis = self.journal.analyze_target(nodes[0], nodes, work)
        verdicts = {item["code"]: item["verdict"] for item in analysis["hypotheses"]}
        self.assertEqual("CONFIRMED", verdicts["HIDDEN_HLA_COMPONENT_LAYER"])
        layer_probe = analysis["subtree_occurrences"][0]["component_state"]["layer_state"]
        self.assertEqual(self.journal.OBSERVED, layer_probe["status"])
        self.assertEqual("Hidden", layer_probe["value"])

    def test_unavailable_load_property_is_not_reported_as_no(self):
        component, occurrence = healthy_component("TARGET")
        del component.Prototype.IsFullyLoaded
        work = FakeWorkPart([component], visible=[occurrence])
        nodes, _ = self.journal.collect_nodes(work)
        row = self.journal.analyze_target(nodes[0], nodes, work)["subtree_occurrences"][0]
        self.assertEqual(self.journal.UNAVAILABLE, row["prototype"]["fully_loaded"]["status"])
        self.assertIsNone(row["prototype"]["fully_loaded"]["value"])

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
            self.assertEqual(2, payload["schema_version"])
            self.assertEqual("READ_ONLY_EXACT_TARGET_ROOT_CAUSE_PROOF", payload["scope"])
            self.assertIn("truth_policy", payload)
            analysis = payload["target_analyses"][0]
            self.assertTrue(analysis["evidence_ledger"])
            fact_ids = {item["id"] for item in analysis["evidence_ledger"]}
            self.assertTrue(set(analysis["conclusion"]["evidence_ids"]) <= fact_ids)

    def test_run_requires_same_work_and_display_hla(self):
        component, occurrence = healthy_component()
        work = FakeWorkPart([component], visible=[occurrence])
        other = FakeWorkPart([], visible=[])
        session = types.SimpleNamespace(
            Parts=types.SimpleNamespace(Work=work, Display=other)
        )
        with self.assertRaisesRegex(RuntimeError, "both work and displayed part"):
            self.journal.run(session)

    def test_real_nx_artifact_rules_out_suppression_and_blanking_for_target_subtree(self):
        payload = json.loads(NX_V1_ARTIFACT.read_text(encoding="utf-8"))
        prefix = (
            "264MN024625A01/A;1-ASSY-DCDC / 264MN031978A01/A"
        )
        rows = [
            row
            for row in payload["ranked_occurrences"]
            if row["ASSEMBLY_PATH"] == prefix
            or row["ASSEMBLY_PATH"].startswith(prefix + " / ")
        ]
        mapped_absent = [
            row
            for row in rows
            if int(row["OCCURRENCE_MEMBERS_FOUND"]) > 0
            and int(row["OCCURRENCE_MEMBERS_VISIBLE_IN_WORK_VIEW"]) == 0
        ]
        unsuppressed_mapped_absent = [
            row for row in mapped_absent if row["SUPPRESSED"] == "NO"
        ]
        self.assertEqual(28, len(rows))
        self.assertEqual(27, len(mapped_absent))
        self.assertEqual(21, len(unsuppressed_mapped_absent))
        self.assertEqual(0, sum(row["IS_BLANKED"] == "YES" for row in rows))
        self.assertEqual(
            0,
            sum(
                int(row["OCCURRENCE_MEMBERS_VISIBLE_IN_WORK_VIEW"]) > 0
                for row in rows
            ),
        )
        self.assertEqual("Isolate", payload["work_view"]["name"])
        self.assertTrue(
            all("GetSuppressedState: No overload" in row["PROBE_ERRORS"] for row in rows)
        )


class Journal24Tests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.j23 = load_journal()
        cls.j24 = load_journal24()

    def session_for(self, work):
        session = types.SimpleNamespace(
            Parts=types.SimpleNamespace(Work=work, Display=work),
            DisplayManager=types.SimpleNamespace(MakeUpToDate=lambda: None),
            undo_marks=[],
            rollbacks=[],
        )

        def set_undo_mark(visibility, name):
            mark = len(session.undo_marks) + 1
            session.undo_marks.append((visibility, name, mark))
            return mark

        def undo_to_mark(mark, name):
            session.rollbacks.append((mark, name))

        session.SetUndoMark = set_undo_mark
        session.UndoToMark = undo_to_mark
        return session

    def run_in_temp(self, session):
        with tempfile.TemporaryDirectory() as folder:
            old_io = os.environ.get("NX_JOURNALS_IO_DIR")
            old_target = os.environ.get("NX_J23_TARGET")
            os.environ["NX_JOURNALS_IO_DIR"] = folder
            os.environ["NX_J23_TARGET"] = "TARGET-264"
            try:
                path, report = self.j24.run(
                    session,
                    datetime.datetime(2026, 8, 14, 13, 0, 0),
                    dependency=self.j23,
                )
                payload = json.loads(Path(path).read_text(encoding="utf-8"))
                return report, payload
            finally:
                if old_io is None:
                    os.environ.pop("NX_JOURNALS_IO_DIR", None)
                else:
                    os.environ["NX_JOURNALS_IO_DIR"] = old_io
                if old_target is None:
                    os.environ.pop("NX_J23_TARGET", None)
                else:
                    os.environ["NX_J23_TARGET"] = old_target

    def test_source_is_display_only_and_has_undo_guard(self):
        source = JOURNAL24.read_text(encoding="utf-8")
        self.assertIn("ShowComponentsInIsolateView", source)
        self.assertIn("SetUndoMark", source)
        self.assertIn("UndoToMark", source)
        for token in (".Save(", ".Suppress(", ".Unsuppress(", ".Blank(", ".Unblank("):
            self.assertNotIn(token, source)

    def test_isolate_show_restores_target_and_confirms_cause(self):
        component, occurrence = healthy_component("TARGET-264")
        work = FakeWorkPart([component], visible=[], view_name="Isolate")
        session = self.session_for(work)
        report, payload = self.run_in_temp(session)
        self.assertEqual(0, report["before"]["mapped_target_count_visible"])
        self.assertEqual(1, report["after"]["mapped_target_count_visible"])
        self.assertEqual("CONFIRMED", report["verdict"]["status"])
        self.assertEqual(
            "ISOLATE_VIEW_MEMBERSHIP_EXCLUDED_TARGET",
            report["verdict"]["root_cause_code"],
        )
        self.assertEqual(1, len(work.ComponentAssembly.show_calls))
        self.assertEqual([], session.rollbacks)
        self.assertEqual(self.j24.BUILD, payload["journal_build"])

    def test_non_isolate_view_is_not_mutated(self):
        component, occurrence = healthy_component("TARGET-264")
        work = FakeWorkPart([component], visible=[], view_name="Trimetric")
        session = self.session_for(work)
        report, _ = self.run_in_temp(session)
        self.assertEqual("NOT_APPLIED", report["verdict"]["status"])
        self.assertEqual([], work.ComponentAssembly.show_calls)
        self.assertEqual([], session.undo_marks)

    def test_api_failure_rolls_back_and_preserves_error_evidence(self):
        component, occurrence = healthy_component("TARGET-264")
        work = FakeWorkPart([component], visible=[], view_name="Isolate#2")
        work.ComponentAssembly.show_exception = RuntimeError("NX rejected view")
        session = self.session_for(work)
        report, _ = self.run_in_temp(session)
        self.assertEqual("API_ERROR", report["verdict"]["status"])
        self.assertIn("NX rejected view", report["action"]["exception"])
        self.assertEqual("ROLLED_BACK", report["rollback"]["status"])
        self.assertEqual(1, len(session.rollbacks))

    def test_no_visibility_change_is_inconclusive_and_rolled_back(self):
        component, occurrence = healthy_component("TARGET-264")
        work = FakeWorkPart([component], visible=[], view_name="Isolate")
        work.ComponentAssembly.ShowComponentsInIsolateView = (
            lambda components, view: FakeErrorList()
        )
        session = self.session_for(work)
        report, _ = self.run_in_temp(session)
        self.assertEqual("INCONCLUSIVE", report["verdict"]["status"])
        self.assertEqual(
            "ISOLATE_SHOW_DID_NOT_RESTORE_MAPPED_GEOMETRY",
            report["verdict"]["root_cause_code"],
        )
        self.assertEqual("ROLLED_BACK", report["rollback"]["status"])
        self.assertEqual(1, len(session.rollbacks))


if __name__ == "__main__":
    unittest.main()
