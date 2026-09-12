import contextlib
import csv
import importlib.util
import os
import sys
import tempfile
import types
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
JOURNAL = ROOT / "from_git" / "journals" / "36_session_tree_pdf_step_export.py"
SCOPE_TOOL = ROOT / "from_git" / "utils" / "single_drawing_scope.py"


def build_nxopen():
    nxopen = types.ModuleType("NXOpen")
    nxopen.NXObject = types.SimpleNamespace(
        AttributeType=types.SimpleNamespace(String="String")
    )
    nxopen.BasePart = types.SimpleNamespace(
        CloseWholeTree=types.SimpleNamespace(FalseValue="FALSE"),
        CloseModified=types.SimpleNamespace(CloseModified="CLOSE_MODIFIED"),
    )
    nxopen.Session = types.SimpleNamespace(
        MarkVisibility=types.SimpleNamespace(Invisible="INVISIBLE")
    )
    nxopen.StepCreator = types.SimpleNamespace(
        ExportFromOption=types.SimpleNamespace(
            DisplayPart="DISPLAY_PART",
            ExistingPart="EXISTING_PART",
        ),
        ExportAsOption=types.SimpleNamespace(
            Ap203="AP203",
            Ap214="AP214",
            Ap242="AP242",
        ),
    )
    nxopen.ObjectSelector = types.SimpleNamespace(
        Scope=types.SimpleNamespace(
            EntirePart="ENTIRE_PART",
            SelectedObjects="SELECTED_OBJECTS",
        )
    )
    nxopen.PrintPDFBuilder = types.SimpleNamespace(
        ActionOption=types.SimpleNamespace(Native="NATIVE"),
        OutputTextOption=types.SimpleNamespace(Text="TEXT"),
    )
    nxopen.Annotations = types.SimpleNamespace(
        OriginBuilder=types.SimpleNamespace(
            AlignmentPosition=types.SimpleNamespace(BottomRight="BOTTOM_RIGHT")
        ),
        LineWidth=types.SimpleNamespace(Normal="NORMAL"),
        TextJustification=types.SimpleNamespace(Right="RIGHT"),
    )
    nxopen.Drawings = types.SimpleNamespace(
        DrawingSheet=types.SimpleNamespace(
            Unit=types.SimpleNamespace(Millimeters="MM", Inches="INCH")
        )
    )
    nxopen.Point3d = lambda x, y, z: ("POINT", x, y, z)
    return nxopen


def load_journal():
    """Return (journal_module, fake_nxopen_module)."""
    nxopen = build_nxopen()
    prior = sys.modules.get("NXOpen")
    sys.modules["NXOpen"] = nxopen
    try:
        spec = importlib.util.spec_from_file_location("journal36", JOURNAL)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        return module, nxopen
    finally:
        if prior is None:
            sys.modules.pop("NXOpen", None)
        else:
            sys.modules["NXOpen"] = prior


def load_scope_tool():
    spec = importlib.util.spec_from_file_location("scope_tool", SCOPE_TOOL)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


@contextlib.contextmanager
def env_overrides(**values):
    keys = (
        "NX_J36_MODE",
        "NX_J36_SCOPE",
        "NX_J36_LOAD_MODE",
        "NX_J36_STEP_SCOPE",
    )
    saved = {key: os.environ.pop(key, None) for key in keys}
    try:
        for key, value in values.items():
            if value is not None:
                os.environ[key] = value
        yield
    finally:
        for key, value in saved.items():
            if value is None:
                os.environ.pop(key, None)
            else:
                os.environ[key] = value


# ---------------------------------------------------------------------------
# Fakes
# ---------------------------------------------------------------------------


class FakeListingWindow:
    def __init__(self):
        self.lines = []

    def Open(self):
        return None

    def WriteFullline(self, line):
        self.lines.append(str(line))


class FakeUserAttribute:
    def __init__(self, value):
        self.StringValue = value


class FakePart:
    def __init__(
        self,
        name,
        tag=None,
        number="",
        revision="",
        attributes=None,
        sheets=None,
        fully_loaded=True,
        load_state="FULLY_LOADED",
        body_count=1,
    ):
        self.Name = name
        self.Tag = tag if tag is not None else name
        self.Leaf = name
        self.FullPath = name
        self.PartName = name
        self.JournalIdentifier = name
        self._attributes = dict(attributes or {})
        if number:
            self._attributes["DB_PART_NO"] = number
        if revision:
            self._attributes["DB_PART_REV"] = revision
        self._sheets = list(sheets or [])
        self.DrawingSheets = FakeSheetCollection(self._sheets)
        self.IsFullyLoaded = fully_loaded
        self.PartLoadState = load_state
        self.Bodies = list(range(body_count))
        self.ComponentAssembly = None
        self.IsModified = False
        self.Annotations = FakeAnnotations()
        self.PlotManager = FakePlotManager()
        self.load_calls = 0
        self.closed = False
        self.display_work_calls = []

    def GetStringAttribute(self, name):
        if name in self._attributes:
            return self._attributes[name]
        raise RuntimeError("attribute not found: {0}".format(name))

    def GetUserAttribute(self, name, attribute_type, index):
        if name in self._attributes:
            return FakeUserAttribute(self._attributes[name])
        raise RuntimeError("attribute not found: {0}".format(name))

    def LoadThisPartFully(self):
        self.load_calls += 1
        self.IsFullyLoaded = True
        return FakeLoadStatus(0)

    def Close(self, whole_tree, close_modified, status):
        self.closed = True


class FakeLoadStatus:
    def __init__(self, unloaded_count=0, entries=None):
        self.NumberUnloadedParts = unloaded_count
        self._entries = list(entries or [])

    def GetPartName(self, index):
        return self._entries[index][0]

    def GetStatus(self, index):
        return self._entries[index][1]

    def GetStatusDescription(self, index):
        return self._entries[index][2]

    def Dispose(self):
        return None


class FakeSheet:
    def __init__(self, name, length=297.0, height=210.0):
        self.Name = name
        self.Length = length
        self.Height = height
        self.Units = "MM"
        self.opened = 0

    def Open(self):
        self.opened += 1


class FakeSheetCollection:
    def __init__(self, sheets):
        self._sheets = list(sheets)
        self.CurrentDrawingSheet = self._sheets[0] if self._sheets else None

    @property
    def Count(self):
        return len(self._sheets)

    def __iter__(self):
        return iter(self._sheets)

    def SetCurrentSheet(self, sheet):
        self.CurrentDrawingSheet = sheet


class FakeTextBlock:
    def __init__(self):
        self.text = []

    def SetText(self, value):
        self.text = list(value) if isinstance(value, list) else [value]


class FakeNoteBuilder:
    def __init__(self):
        self.Origin = types.SimpleNamespace(Anchor=None, OriginPoint=None)
        self.Style = types.SimpleNamespace(
            LetteringStyle=types.SimpleNamespace(
                GeneralTextSize=None,
                GeneralTextLineWidth=None,
                HorizontalTextJustification=None,
            )
        )
        self.Text = types.SimpleNamespace(TextBlock=FakeTextBlock())
        self.committed = 0
        self.destroyed = 0

    def Commit(self):
        self.committed += 1
        return "NOTE"

    def Destroy(self):
        self.destroyed += 1


class FakeAnnotations:
    def __init__(self):
        self.builders = []

    def CreateDraftingNoteBuilder(self, obj):
        builder = FakeNoteBuilder()
        self.builders.append(builder)
        return builder


class FakeSourceBuilder:
    def __init__(self):
        self.sheets = []

    def SetSheets(self, sheets):
        self.sheets = list(sheets)


class FakePdfBuilder:
    def __init__(self, write_pdf=True):
        self.Action = None
        self.Filename = None
        self.Append = None
        self.OutputText = None
        self.AddWatermark = None
        self.Watermark = None
        self.CustomSymbolsInForeground = None
        self.SourceBuilder = FakeSourceBuilder()
        self.committed = 0
        self.destroyed = 0
        self._write_pdf = write_pdf

    def Commit(self):
        self.committed += 1
        if self._write_pdf and self.Filename:
            with open(self.Filename, "w", encoding="utf-8") as handle:
                handle.write("%PDF-1.4 fake\n")

    def Destroy(self):
        self.destroyed += 1


class FakePlotManager:
    def __init__(self, write_pdf=True):
        self.builders = []
        self._write_pdf = write_pdf

    def CreatePrintPdfbuilder(self):
        builder = FakePdfBuilder(write_pdf=self._write_pdf)
        self.builders.append(builder)
        return builder


class FakeStepCreator:
    def __init__(self, write_step=True, body_token="MANIFOLD_SOLID_BREP"):
        self.OutputFile = None
        self.ExportFrom = None
        self.ExportSelectionBlock = types.SimpleNamespace(SelectionScope=None)
        self.LayerMask = None
        self.ObjectTypes = types.SimpleNamespace(
            Solids=None, Surfaces=None, Curves=None
        )
        self.ExportAs = None
        self.ProcessHoldFlag = None
        self.committed = 0
        self.destroyed = 0
        self._write_step = write_step
        self._body_token = body_token

    def Commit(self):
        self.committed += 1
        if not self._write_step or not self.OutputFile:
            return
        with open(self.OutputFile, "w", encoding="utf-8") as handle:
            handle.write("ISO-10303-21;\nHEADER;\nENDSEC;\n")
            handle.write("DATA;\n")
            if self._body_token:
                handle.write("#1={0}('x',(#2));\n".format(self._body_token))
            handle.write("ENDSEC;\nEND-ISO-10303-21;\n")

    def Destroy(self):
        self.destroyed += 1


class FakeComponent:
    def __init__(
        self,
        name,
        prototype=None,
        children=None,
        suppressed=False,
        attributes=None,
    ):
        self.Name = name
        self.DisplayName = name
        self.Prototype = prototype
        self._children = list(children or [])
        self.IsSuppressed = suppressed
        self._attributes = dict(attributes or {})

    def GetChildren(self):
        return list(self._children)

    def GetStringAttribute(self, title):
        if title in self._attributes:
            return self._attributes[title]
        raise RuntimeError("component attribute not found: {0}".format(title))


class FakePartsCollection:
    def __init__(self, session):
        self._session = session
        self.Display = None
        self.Work = None
        self.set_display_calls = []
        self.set_work_calls = []
        self.open_display_calls = []

    def __iter__(self):
        return iter(self._session.loaded_parts)

    def SetDisplay(self, part, visibility, activate):
        self.Display = part
        self.set_display_calls.append(part)
        part.display_work_calls.append(("display", part.Name))
        return (part, None)

    def SetWork(self, part):
        self.Work = part
        self.set_work_calls.append(part)
        part.display_work_calls.append(("work", part.Name))

    def OpenDisplay(self, specification):
        self.open_display_calls.append(specification)
        opener = self._session.open_display_hook
        if opener is None:
            raise RuntimeError("no such specification: " + specification)
        return opener(specification)


class FakeDexManager:
    def __init__(self, step_creator):
        self.step_creator = step_creator
        self.created = 0

    def CreateStepCreator(self):
        self.created += 1
        return self.step_creator


class FakeSession:
    def __init__(
        self,
        work_part=None,
        loaded_parts=None,
        step_creator=None,
        open_display_hook=None,
    ):
        self.ListingWindow = FakeListingWindow()
        self.Parts = FakePartsCollection(self)
        self.DexManager = FakeDexManager(step_creator or FakeStepCreator())
        self.loaded_parts = list(loaded_parts or [])
        self.open_display_hook = open_display_hook
        self.undo_marks = []
        self.undo_to_mark_calls = []
        self.deleted_marks = []
        self.updated = 0
        self.IsManagedMode = False
        if work_part is not None:
            self.Parts.Work = work_part
            self.Parts.Display = work_part
            if work_part not in self.loaded_parts:
                self.loaded_parts.append(work_part)

    def SetUndoMark(self, visibility, name):
        mark = ("MARK", name)
        self.undo_marks.append(mark)
        return mark

    def UndoToMark(self, mark, name):
        self.undo_to_mark_calls.append((mark, name))

    def DeleteUndoMark(self, mark, name):
        self.deleted_marks.append((mark, name))

    @property
    def UpdateManager(self):
        session = self

        class UpdateManager:
            def DoUpdate(self, mark):
                session.updated += 1
                return 0

        return UpdateManager()


def make_assembly():
    """Build a 3-level fake assembly tree.

    root (L0)
      sub_assy (L1)
        leaf_a (L2)
        shared (L2)
      leaf_b (L1)  -> duplicate of shared prototype
      leaf_c (L1)
      suppressed_leaf (L1, suppressed)
      csys_leaf (L1, keyword name)
      reference_leaf (L1, REFERENCE_COMPONENT flagged)
      excluded_assy (L1, CELESTICA_BOM_EXCLUDE_SUBTREE=YES)
        hidden_leaf (L2)
    """
    root = FakePart("ROOT_ASSY", tag="root", number="264MN000001A01", revision="A")
    sub_assy = FakePart(
        "SUB_ASSY", tag="sub", number="264MN000002A01", revision="A"
    )
    leaf_a = FakePart("LEAF_A", tag="leafa", number="264MN000003A01", revision="A")
    shared = FakePart("SHARED", tag="shared", number="264MN000004A01", revision="A")
    leaf_b = FakePart("LEAF_B", tag="leafb", number="264MN000005A01", revision="A")
    leaf_c = FakePart("LEAF_C", tag="leafc", number="264MN000006A01", revision="A")
    suppressed_leaf = FakePart(
        "SUPPRESSED", tag="sup", number="264MN000007A01", revision="A"
    )
    csys_leaf = FakePart("CSYS_1", tag="csys", number="264MN000008A01", revision="A")
    reference_leaf = FakePart(
        "REF_LEAF", tag="ref", number="264MN000009A01", revision="A"
    )
    excluded_assy = FakePart(
        "EXCLUDED_ASSY", tag="excl", number="264MN00000AA01", revision="A"
    )
    hidden_leaf = FakePart(
        "HIDDEN_LEAF", tag="hidden", number="264MN00000BA01", revision="A"
    )

    root.ComponentAssembly = types.SimpleNamespace(
        RootComponent=FakeComponent(
            "ROOT_ASSY",
            prototype=root,
            children=[
                FakeComponent(
                    "SUB_ASSY",
                    prototype=sub_assy,
                    children=[
                        FakeComponent("LEAF_A", prototype=leaf_a),
                        FakeComponent("SHARED", prototype=shared),
                    ],
                ),
                FakeComponent("LEAF_B", prototype=shared),
                FakeComponent("LEAF_C", prototype=leaf_c),
                FakeComponent("SUPPRESSED", prototype=suppressed_leaf, suppressed=True),
                FakeComponent("CSYS_1", prototype=csys_leaf),
                FakeComponent(
                    "REF_LEAF",
                    prototype=reference_leaf,
                    attributes={"REFERENCE_COMPONENT": "YES"},
                ),
                FakeComponent(
                    "EXCLUDED_ASSY",
                    prototype=excluded_assy,
                    attributes={"CELESTICA_BOM_EXCLUDE_SUBTREE": "YES"},
                    children=[
                        FakeComponent("HIDDEN_LEAF", prototype=hidden_leaf),
                    ],
                ),
            ],
        )
    )
    sub_assy.ComponentAssembly = types.SimpleNamespace(
        RootComponent=FakeComponent(
            "SUB_ASSY",
            prototype=sub_assy,
            children=[
                FakeComponent("LEAF_A", prototype=leaf_a),
                FakeComponent("SHARED", prototype=shared),
            ],
        )
    )
    return {
        "root": root,
        "sub_assy": sub_assy,
        "leaf_a": leaf_a,
        "shared": shared,
        "leaf_b_occurrence_prototype": shared,
        "leaf_c": leaf_c,
        "suppressed_leaf": suppressed_leaf,
        "csys_leaf": csys_leaf,
        "reference_leaf": reference_leaf,
        "excluded_assy": excluded_assy,
        "hidden_leaf": hidden_leaf,
    }


class Journal36SessionTreeExportTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.journal, cls.nxopen = load_journal()

    # --- modes ---------------------------------------------------------

    def test_mode_defaults(self):
        with env_overrides():
            self.assertEqual("DRY_RUN", self.journal.resolve_write_mode())
            self.assertEqual("BOM", self.journal.resolve_scope_filter())
            self.assertEqual("LOAD", self.journal.resolve_load_mode())
            self.assertEqual("DRAWING_ONLY", self.journal.resolve_step_scope())

    def test_mode_env_overrides(self):
        with env_overrides(
            NX_J36_MODE="apply",
            NX_J36_SCOPE="all",
            NX_J36_LOAD_MODE="report_only",
            NX_J36_STEP_SCOPE="all",
        ):
            self.assertEqual("APPLY", self.journal.resolve_write_mode())
            self.assertEqual("ALL", self.journal.resolve_scope_filter())
            self.assertEqual("REPORT_ONLY", self.journal.resolve_load_mode())
            self.assertEqual("ALL", self.journal.resolve_step_scope())

    def test_invalid_modes_fail_closed(self):
        for key, value in (
            ("NX_J36_MODE", "EXPORT"),
            ("NX_J36_SCOPE", "SOMETHING"),
            ("NX_J36_LOAD_MODE", "MAYBE"),
            ("NX_J36_STEP_SCOPE", "LEAVES"),
        ):
            with env_overrides(**{key: value}):
                resolver = {
                    "NX_J36_MODE": self.journal.resolve_write_mode,
                    "NX_J36_SCOPE": self.journal.resolve_scope_filter,
                    "NX_J36_LOAD_MODE": self.journal.resolve_load_mode,
                    "NX_J36_STEP_SCOPE": self.journal.resolve_step_scope,
                }[key]
                with self.assertRaisesRegex(ValueError, key):
                    resolver()

    # --- scope ---------------------------------------------------------

    def test_bom_scope_excludes_suppressed_keyword_reference_and_excluded(self):
        tree = make_assembly()
        targets, diagnostics = self.journal.collect_session_scope(
            tree["root"], "BOM"
        )
        names = {target["part"].Name for target in targets}
        self.assertEqual(
            {"ROOT_ASSY", "SUB_ASSY", "LEAF_A", "SHARED", "LEAF_C"}, names
        )
        for excluded in (
            "SUPPRESSED",
            "CSYS_1",
            "REF_LEAF",
            "EXCLUDED_ASSY",
            "HIDDEN_LEAF",
        ):
            self.assertNotIn(excluded, names)
        self.assertEqual([], diagnostics)

    def test_scope_all_keeps_every_loaded_occurrence(self):
        tree = make_assembly()
        targets, _diagnostics = self.journal.collect_session_scope(
            tree["root"], "ALL"
        )
        names = {target["part"].Name for target in targets}
        self.assertIn("SUPPRESSED", names)
        self.assertIn("CSYS_1", names)
        self.assertIn("HIDDEN_LEAF", names)

    def test_shared_prototype_is_deduplicated_at_topmost_level(self):
        tree = make_assembly()
        targets, _diagnostics = self.journal.collect_session_scope(
            tree["root"], "BOM"
        )
        by_name = {target["part"].Name: target for target in targets}
        shared = by_name["SHARED"]
        self.assertEqual(1, shared["level"])
        self.assertEqual(2, shared["deepest_level"])
        self.assertEqual(2, shared["occurrence_count"])
        self.assertIn("LEAF_B", shared["component_path"])
        root_target = by_name["ROOT_ASSY"]
        self.assertTrue(root_target["is_work_part"])
        self.assertEqual(0, root_target["level"])
        self.assertEqual("ASSEMBLY", self.journal.part_kind(tree["root"]))
        self.assertEqual("PART", self.journal.part_kind(tree["leaf_a"]))
        self.assertEqual(
            "LEAF_A", by_name["LEAF_A"]["component_path"].split(" / ")[-1]
        )

    def test_scope_is_sorted_by_level(self):
        tree = make_assembly()
        targets, _diagnostics = self.journal.collect_session_scope(
            tree["root"], "BOM"
        )
        levels = [target["level"] for target in targets]
        self.assertEqual(sorted(levels), levels)

    def test_missing_prototype_is_reported_not_fatal(self):
        root = FakePart("ROOT", tag="root", number="N1", revision="A")
        root.ComponentAssembly = types.SimpleNamespace(
            RootComponent=FakeComponent(
                "ROOT",
                prototype=root,
                children=[FakeComponent("UNLOADED", prototype=None)],
            )
        )
        targets, diagnostics = self.journal.collect_session_scope(root, "BOM")
        self.assertEqual(["ROOT"], [t["part"].Name for t in targets])
        self.assertEqual(["MISSING_MODEL"], [d["code"] for d in diagnostics])

    # --- load gate -----------------------------------------------------

    def test_load_mode_report_only_never_loads(self):
        tree = make_assembly()
        tree["leaf_a"].IsFullyLoaded = False
        tree["leaf_a"].PartLoadState = "PARTIALLY_LOADED"
        ok, targets, records, diagnostics = self.journal.load_session_scope(
            tree["root"], "BOM", "REPORT_ONLY"
        )
        self.assertTrue(ok)
        self.assertEqual(0, tree["leaf_a"].load_calls)
        record = records[self.journal._object_key(tree["leaf_a"])]
        self.assertEqual("NOT_ATTEMPTED", record["load_action"])
        self.assertEqual("NOT_EVALUATED", record["load_status"])
        self.assertEqual([], diagnostics)
        self.assertIn(
            "LEAF_A", {target["part"].Name for target in targets}
        )

    def test_load_gate_loads_unloaded_and_retraverses(self):
        root = FakePart("ROOT", tag="root", number="N1", revision="A")
        sub = FakePart(
            "SUB",
            tag="sub",
            number="N2",
            revision="A",
            fully_loaded=False,
            load_state="PARTIALLY_LOADED",
        )
        sub_occurrence = FakeComponent("SUB", prototype=sub)
        leaf = FakePart("LEAF", tag="leaf", number="N3", revision="A")

        def load_hook():
            # Fully loading the sub-assembly exposes its child occurrence,
            # exactly like the Journal 21 descendant-discovery loop.
            sub.load_calls += 1
            sub.IsFullyLoaded = True
            sub.PartLoadState = "FULLY_LOADED"
            sub_occurrence._children = [
                FakeComponent("LEAF", prototype=leaf)
            ]
            return FakeLoadStatus(0)

        sub.LoadThisPartFully = load_hook
        root.ComponentAssembly = types.SimpleNamespace(
            RootComponent=FakeComponent(
                "ROOT", prototype=root, children=[sub_occurrence]
            )
        )

        ok, targets, records, diagnostics = self.journal.load_session_scope(
            root, "BOM", "LOAD"
        )
        self.assertTrue(ok)
        self.assertEqual([], diagnostics)
        self.assertEqual(
            {"ROOT", "SUB", "LEAF"},
            {target["part"].Name for target in targets},
        )
        record = records[self.journal._object_key(sub)]
        self.assertEqual("LOAD_THIS_PART_FULLY", record["load_action"])
        self.assertEqual("SUCCESS", record["load_status"])
        self.assertEqual("FULLY_LOADED", record["final_load_state"])
        self.assertEqual(1, sub.load_calls)

    def test_load_failure_is_recorded_and_not_fatal(self):
        tree = make_assembly()
        failing = tree["leaf_a"]

        def fail():
            failing.IsFullyLoaded = False
            return FakeLoadStatus(
                1,
                [("OTHER_PART", "MISSING_FILE", "not found using current search options")],
            )

        failing.IsFullyLoaded = False
        failing.PartLoadState = "PARTIALLY_LOADED"
        failing.LoadThisPartFully = fail

        ok, targets, records, diagnostics = self.journal.load_session_scope(
            tree["root"], "BOM", "LOAD"
        )
        self.assertFalse(ok)
        record = records[self.journal._object_key(failing)]
        self.assertEqual("MISSING_FILE", record["load_status"])
        codes = [diagnostic["code"] for diagnostic in diagnostics]
        self.assertIn("MISSING_FILE", codes)
        # Sibling branches survive a failed load.
        self.assertIn("LEAF_C", {target["part"].Name for target in targets})

    # --- naming --------------------------------------------------------

    def test_versioned_base_and_pdf_filename(self):
        journal = self.journal
        self.assertEqual(
            "264MN000001A01_REVA.V1.2",
            journal.build_versioned_base("264MN000001A01", "A", "V1.2"),
        )
        self.assertEqual(
            "264MN000001A01_REVA",
            journal.build_versioned_base("264MN000001A01", "A", ""),
        )
        self.assertEqual(
            "264MN000001A01_REVA.V1.2.pdf",
            journal.build_pdf_filename(
                "264MN000001A01", "A", "V1.2", "DWG1", 1
            ),
        )
        self.assertEqual(
            "264MN000001A01_REVA.V1.2_DWG2.pdf",
            journal.build_pdf_filename(
                "264MN000001A01", "A", "V1.2", "DWG2", 3
            ),
        )
        self.assertEqual(
            "PART_WITH_QUESTION_REVA.stp",
            journal.build_versioned_base("PART?WITH:QUESTION", "A", "") + ".stp",
        )
        self.assertEqual(
            "264MN000001A01_REVA.V1.2.stp",
            journal.build_versioned_base(
                "264MN000001A01", "A", "V1.2"
            )
            + ".stp",
        )

    def test_drawing_tokens_are_unique_and_ordered(self):
        candidates = [
            {
                "part": FakePart("264MN000001A01-A-DWG1", tag="dwg1"),
                "drawing_index": 1,
            },
            {
                "part": FakePart("264MN000001A01-A-DWG2", tag="dwg2"),
                "drawing_index": 2,
            },
            {
                "part": FakePart("264MN000001A01-A-DWG2B", tag="dwg2b"),
                "drawing_index": 2,
            },
            {
                "part": FakePart("UNKNOWN_DRAWING", tag="unknown"),
                "drawing_index": None,
            },
        ]
        tokens = self.journal.unique_drawing_tokens(candidates)
        self.assertEqual(len(tokens), len(candidates))
        self.assertEqual(len(set(tokens)), len(candidates))
        self.assertEqual("DWG1", tokens[0])
        self.assertEqual("DWG2", tokens[1])
        self.assertNotEqual("DWG2", tokens[2])
        self.assertTrue(tokens[3].startswith("DWG"))

    def test_drawing_index_detection(self):
        journal = self.journal
        self.assertEqual(3, journal.drawing_index_from_text("ITEM-A-DWG3"))
        self.assertEqual(12, journal.drawing_index_from_text("x_dwg12_extra"))
        self.assertIsNone(journal.drawing_index_from_text("ITEM-A-NODRAWING"))

    def test_teamcenter_drawing_specification_format(self):
        self.assertEqual(
            [
                "@DB/264MN000001A01/A/specification/"
                "264MN000001A01-A-dwg4"
            ],
            self.journal.teamcenter_drawing_specs(
                "264MN000001A01", "A", 4
            ),
        )

    def test_step_scope_switch(self):
        journal = self.journal
        self.assertTrue(journal.step_requested_for_target(True, "DRAWING_ONLY"))
        self.assertFalse(journal.step_requested_for_target(False, "DRAWING_ONLY"))
        self.assertTrue(journal.step_requested_for_target(False, "ALL"))

    # --- result aggregation -------------------------------------------

    def test_overall_result_matrix(self):
        journal = self.journal
        self.assertEqual(
            "SUCCESS", journal.overall_result("SUCCESS", "SUCCESS")
        )
        self.assertEqual(
            "PARTIAL_SUCCESS",
            journal.overall_result("SKIPPED_EXISTS", "SUCCESS"),
        )
        self.assertEqual(
            "SKIPPED_NO_DRAWING",
            journal.overall_result("SKIPPED_NO_DRAWING", "SKIPPED_NO_DRAWING"),
        )
        self.assertEqual(
            "SKIPPED_EXISTS",
            journal.overall_result("SKIPPED_EXISTS", "SKIPPED_EXISTS"),
        )
        self.assertEqual(
            "SKIPPED_PARTIAL",
            journal.overall_result("SKIPPED_NO_DRAWING", "SKIPPED_EXISTS"),
        )
        self.assertEqual("FAILED", journal.overall_result("FAILED", "FAILED"))
        self.assertEqual(
            "DRY_RUN", journal.overall_result("PLANNED", "PLANNED", "DRY_RUN")
        )
        self.assertEqual(
            "SKIPPED_NO_DRAWING",
            journal.overall_result(
                "SKIPPED_NO_DRAWING", "SKIPPED_NO_DRAWING", "DRY_RUN"
            ),
        )

    def test_report_columns_are_complete_everywhere(self):
        journal = self.journal
        tree = make_assembly()
        targets, diagnostics = journal.collect_session_scope(tree["root"], "BOM")
        rows = [
            journal.new_result("TS", "DRY_RUN", "BOM", "LOAD", targets[0]),
            journal.diagnostic_row("TS", "DRY_RUN", "BOM", "LOAD", {
                "code": "MISSING_MODEL",
                "message": "m",
                "component_path": "p",
                "level": 1,
            }),
            journal.summary_row(
                "TS", "DRY_RUN", "BOM", "LOAD", {"SUCCESS": 1}, {"targets": 1}
            ),
        ]
        for row in rows:
            for column in journal._RESULT_COLUMNS:
                self.assertIn(column, row, column)
        # Columns the J25 scope generator reads must stay present.
        for required in (
            "DB_PART_NO",
            "DB_PART_REV",
            "PDF_RESULT",
            "PDF_FILE_COUNT",
            "PDF_FILES",
            "OVERALL_RESULT",
            "MESSAGE",
        ):
            self.assertIn(required, journal._RESULT_COLUMNS)
        self.assertEqual("TARGET", rows[0]["ROW_TYPE"])
        self.assertEqual("DIAGNOSTIC", rows[1]["ROW_TYPE"])
        self.assertEqual("SUMMARY", rows[2]["ROW_TYPE"])

    def test_result_csv_round_trip(self):
        journal = self.journal
        tree = make_assembly()
        targets, _diagnostics = journal.collect_session_scope(tree["root"], "BOM")
        rows = [
            journal.new_result("TS", "DRY_RUN", "BOM", "LOAD", targets[0])
        ]
        with tempfile.TemporaryDirectory() as folder:
            path = os.path.join(folder, "EXPORT_RESULT_TS.csv")
            journal.write_result_csv(path, rows)
            with open(path, "r", encoding="utf-8-sig", newline="") as handle:
                reader = csv.DictReader(handle)
                self.assertEqual(list(journal._RESULT_COLUMNS), reader.fieldnames)
                data = list(reader)
            self.assertEqual(1, len(data))
            self.assertEqual("ROOT_ASSY", data[0]["PART_NAME"])

    def test_export_result_stays_j25_scope_generator_compatible(self):
        journal = self.journal
        scope_tool = load_scope_tool()
        tree = make_assembly()
        targets, _diagnostics = journal.collect_session_scope(tree["root"], "BOM")
        row = journal.new_result("TS", "APPLY", "BOM", "LOAD", targets[0])
        row["PDF_RESULT"] = "SUCCESS"
        row["PDF_FILE_COUNT"] = 2
        row["PDF_FILES"] = ";".join(
            [
                "/out/264MN000001A01_REVA.V1.2_DWG1.pdf",
                "/out/264MN000001A01_REVA.V1.2_DWG2.pdf",
            ]
        )
        rows, skipped, errors = scope_tool.scope_rows([row])
        self.assertEqual([], errors)
        self.assertEqual([], skipped)
        self.assertEqual(1, len(rows))
        self.assertEqual("264MN000001A01", rows[0]["PART_NUMBER"])
        self.assertEqual("1", rows[0]["KEEP_DWG_INDEX"])
        self.assertEqual("2", rows[0]["EXPECTED_REMOVE_DWG_INDICES"])

    # --- STEP ----------------------------------------------------------

    def test_step_body_signature_count(self):
        journal = self.journal
        with tempfile.TemporaryDirectory() as folder:
            path = os.path.join(folder, "sample.stp")
            with open(path, "w", encoding="utf-8") as handle:
                handle.write(
                    "ISO-10303-21;\nHEADER;\nENDSEC;\nDATA;\n"
                    "#1=MANIFOLD_SOLID_BREP('x',(#2));\n"
                    "#2=CLOSED_SHELL('y',(#3));\n"
                    "ENDSEC;\n"
                )
            self.assertEqual(2, journal.step_body_signature_count(path))

            header_only = os.path.join(folder, "empty.stp")
            with open(header_only, "w", encoding="utf-8") as handle:
                handle.write("ISO-10303-21;\nHEADER;\nENDSEC;\nDATA;\nENDSEC;\n")
            self.assertEqual(0, journal.step_body_signature_count(header_only))

    def test_export_step_skips_existing_output(self):
        journal = self.journal
        tree = make_assembly()
        with tempfile.TemporaryDirectory() as folder:
            existing = os.path.join(
                folder,
                journal.build_versioned_base("264MN000003A01", "A", "")
                + ".stp",
            )
            Path(existing).write_text("already here", encoding="utf-8")
            session = FakeSession(work_part=tree["leaf_a"])
            result = journal.export_step_from_part(
                session,
                tree["leaf_a"],
                folder,
                "264MN000003A01",
                "A",
                "",
            )
            self.assertEqual("SKIPPED_EXISTS", result["result"])
            self.assertEqual(existing, result["path"])
            self.assertEqual(0, session.DexManager.created)
            self.assertEqual(
                "already here", Path(existing).read_text(encoding="utf-8")
            )

    def test_export_step_uses_the_j07_proven_option_set(self):
        journal = self.journal
        tree = make_assembly()
        creator = FakeStepCreator()
        session = FakeSession(work_part=tree["leaf_c"], step_creator=creator)
        with tempfile.TemporaryDirectory() as folder:
            result = journal.export_step_from_part(
                session,
                tree["leaf_c"],
                folder,
                "264MN000006A01",
                "A",
                "V1.0",
            )
            self.assertEqual("SUCCESS", result["result"])
            self.assertEqual("DISPLAY_PART", creator.ExportFrom)
            self.assertEqual("ENTIRE_PART", creator.ExportSelectionBlock.SelectionScope)
            self.assertEqual("1-256", creator.LayerMask)
            self.assertEqual("AP214", creator.ExportAs)
            self.assertTrue(creator.ProcessHoldFlag)
            self.assertTrue(creator.ObjectTypes.Solids)
            self.assertTrue(creator.ObjectTypes.Surfaces)
            self.assertTrue(creator.ObjectTypes.Curves)
            self.assertEqual(1, creator.committed)
            self.assertEqual(1, creator.destroyed)
            self.assertTrue(os.path.isfile(result["path"]))
            self.assertGreater(int(result["size"]), 0)
            self.assertIn(tree["leaf_c"], session.Parts.set_display_calls)
            self.assertIn(tree["leaf_c"], session.Parts.set_work_calls)

    def test_export_step_reports_zero_geometry(self):
        journal = self.journal
        tree = make_assembly()
        creator = FakeStepCreator(body_token="")
        session = FakeSession(work_part=tree["leaf_a"], step_creator=creator)
        with tempfile.TemporaryDirectory() as folder:
            result = journal.export_step_from_part(
                session, tree["leaf_a"], folder, "264MN000003A01", "A", ""
            )
            self.assertEqual("FAILED_ZERO_GEOMETRY", result["result"])

    def test_export_step_reports_missing_output_file(self):
        journal = self.journal
        tree = make_assembly()
        creator = FakeStepCreator(write_step=False)
        session = FakeSession(work_part=tree["leaf_a"], step_creator=creator)
        with tempfile.TemporaryDirectory() as folder:
            result = journal.export_step_from_part(
                session, tree["leaf_a"], folder, "264MN000003A01", "A", ""
            )
            self.assertEqual("FAILED_NO_OUTPUT_FILE", result["result"])

    # --- PDF helpers ---------------------------------------------------

    def test_sheet_units_convert_millimeters_to_inches(self):
        journal = self.journal
        sheet = FakeSheet("S1")
        self.assertFalse(journal.sheet_uses_inches(sheet))
        self.assertEqual(25.4, journal.millimeters_to_sheet_units(25.4, sheet))
        sheet.Units = "INCH"
        self.assertTrue(journal.sheet_uses_inches(sheet))
        self.assertEqual(1.0, journal.millimeters_to_sheet_units(25.4, sheet))

    def test_watermark_falls_back_to_revision_only(self):
        journal = self.journal
        part = FakePart("LEAF", tag="leaf")
        wae_version, source, warning = journal.resolve_pdf_watermark(part, [])
        self.assertEqual("", wae_version)
        self.assertEqual("revision-only fallback", source)
        self.assertIn("WAE_VERSION", warning)
        self.assertEqual("DRAFT_B", journal.build_pdf_watermark("B", wae_version))

        part._attributes["WAE_VERSION"] = "V9.1"
        wae_version, source, warning = journal.resolve_pdf_watermark(part, [])
        self.assertEqual("V9.1", wae_version)
        self.assertEqual("", warning)
        self.assertEqual("DRAFT_B.V9.1", journal.build_pdf_watermark("B", wae_version))

    def test_watermark_prefers_drawing_when_part_is_blank(self):
        journal = self.journal
        part = FakePart("LEAF", tag="leaf")
        drawing = FakePart(
            "DRAWING", tag="dwg", attributes={"WAE_VERSION": "V2.0"}
        )
        wae_version, source, _warning = journal.resolve_pdf_watermark(
            part, [{"part": drawing}]
        )
        self.assertEqual("V2.0", wae_version)
        self.assertEqual("drawing WAE_VERSION", source)

    # --- per-target processing ----------------------------------------

    def build_drawing_target(self, tree, wae_version="V1.2"):
        root = tree["root"]
        root._attributes["WAE_VERSION"] = wae_version
        drawing = FakePart(
            "264MN000001A01-A-DWG1",
            tag="rootdrawing",
            attributes={
                "DB_PART_NO": "264MN000001A01",
                "DB_PART_REV": "A",
                "WAE_VERSION": wae_version,
            },
            sheets=[FakeSheet("Sheet1"), FakeSheet("Sheet2")],
        )
        return drawing

    def test_process_target_dry_run_plans_without_writing(self):
        journal = self.journal
        tree = make_assembly()
        drawing = self.build_drawing_target(tree)
        session = FakeSession(work_part=tree["root"], loaded_parts=[drawing])
        targets, _diagnostics = journal.collect_session_scope(tree["root"], "BOM")
        target = [t for t in targets if t["part"].Name == "ROOT_ASSY"][0]

        with tempfile.TemporaryDirectory() as folder:
            folders = {
                "run": folder,
                "pdf": os.path.join(folder, "PDF"),
                "step": os.path.join(folder, "STEP"),
                "reports": os.path.join(folder, "REPORTS"),
                "logs": os.path.join(folder, "LOGS"),
            }
            for path in folders.values():
                os.makedirs(path, exist_ok=True)
            result = journal.process_target(
                session,
                target,
                folders,
                "TS",
                journal.build_export_timestamp_text(
                    __import__("datetime").datetime.now()
                ),
                "DRY_RUN",
                "BOM",
                "LOAD",
                "DRAWING_ONLY",
                {"load_status": "SUCCESS", "initial_load_state": "FULLY_LOADED",
                 "load_action": "NOT_REQUIRED", "final_load_state": "FULLY_LOADED"},
                {"halted": False, "reason": ""},
                tree["root"],
                tree["root"],
                [],
            )
            self.assertEqual("PLANNED", result["PDF_RESULT"])
            self.assertEqual("PLANNED", result["STEP_RESULT"])
            self.assertEqual("DRY_RUN", result["OVERALL_RESULT"])
            self.assertEqual(1, result["DRAWING_COUNT"])
            self.assertEqual("YES", result["HAS_DRAWING"])
            self.assertEqual(0, result["LEVEL"])
            self.assertEqual(0, session.DexManager.created)
            self.assertEqual(
                [], os.listdir(os.path.join(folder, "PDF"))
            )
            self.assertEqual(
                [], os.listdir(os.path.join(folder, "STEP"))
            )

    def test_dry_run_closes_journal_opened_teamcenter_drawings(self):
        """DRY_RUN must not leave @DB drawings it opened in the session."""
        journal = self.journal
        tree = make_assembly()
        drawing = self.build_drawing_target(tree)
        session = FakeSession(
            work_part=tree["root"],
            loaded_parts=[],
            open_display_hook=lambda specification: drawing,
        )
        targets, _diagnostics = journal.collect_session_scope(tree["root"], "BOM")
        target = [t for t in targets if t["part"].Name == "ROOT_ASSY"][0]

        with tempfile.TemporaryDirectory() as folder:
            folders = {
                "run": folder,
                "pdf": os.path.join(folder, "PDF"),
                "step": os.path.join(folder, "STEP"),
                "reports": os.path.join(folder, "REPORTS"),
                "logs": os.path.join(folder, "LOGS"),
            }
            for path in folders.values():
                os.makedirs(path, exist_ok=True)
            result = journal.process_target(
                session,
                target,
                folders,
                "TS",
                journal.build_export_timestamp_text(
                    __import__("datetime").datetime.now()
                ),
                "DRY_RUN",
                "BOM",
                "LOAD",
                "DRAWING_ONLY",
                {"load_status": "SUCCESS", "initial_load_state": "FULLY_LOADED",
                 "load_action": "NOT_REQUIRED", "final_load_state": "FULLY_LOADED"},
                {"halted": False, "reason": ""},
                tree["root"],
                tree["root"],
                [],
            )
            self.assertEqual("PLANNED", result["PDF_RESULT"])
            self.assertEqual("YES", result["HAS_DRAWING"])
            self.assertTrue(
                any(
                    "@DB/264MN000001A01/A/specification" in call
                    for call in session.Parts.open_display_calls
                )
            )
            self.assertTrue(drawing.closed)
            self.assertEqual(tree["root"], session.Parts.Display)
            self.assertEqual(tree["root"], session.Parts.Work)
            self.assertEqual([], os.listdir(os.path.join(folder, "PDF")))

    def test_dry_run_plans_step_for_drawingless_level_in_all_scope(self):
        journal = self.journal
        tree = make_assembly()
        session = FakeSession(work_part=tree["root"], loaded_parts=[])
        targets, _diagnostics = journal.collect_session_scope(tree["root"], "BOM")
        target = [t for t in targets if t["part"].Name == "LEAF_C"][0]

        with tempfile.TemporaryDirectory() as folder:
            folders = {
                "run": folder,
                "pdf": os.path.join(folder, "PDF"),
                "step": os.path.join(folder, "STEP"),
                "reports": os.path.join(folder, "REPORTS"),
                "logs": os.path.join(folder, "LOGS"),
            }
            for path in folders.values():
                os.makedirs(path, exist_ok=True)
            result = journal.process_target(
                session,
                target,
                folders,
                "TS",
                journal.build_export_timestamp_text(
                    __import__("datetime").datetime.now()
                ),
                "DRY_RUN",
                "BOM",
                "LOAD",
                "ALL",
                None,
                {"halted": False, "reason": ""},
                tree["root"],
                tree["root"],
                [],
            )
            self.assertEqual("SKIPPED_NO_DRAWING", result["PDF_RESULT"])
            self.assertEqual("PLANNED", result["STEP_RESULT"])
            self.assertEqual("YES", result["STEP_REQUESTED"])
            self.assertEqual("DRY_RUN", result["OVERALL_RESULT"])
            self.assertTrue(
                result["STEP_FILE"].endswith("264MN000006A01_REVA.stp")
            )
            self.assertIn("would export 1 STEP", result["MESSAGE"])

    def test_process_target_apply_exports_pdf_and_step(self):
        journal = self.journal
        tree = make_assembly()
        drawing = self.build_drawing_target(tree)
        creator = FakeStepCreator()
        session = FakeSession(
            work_part=tree["root"], loaded_parts=[drawing], step_creator=creator
        )
        targets, _diagnostics = journal.collect_session_scope(tree["root"], "BOM")
        target = [t for t in targets if t["part"].Name == "ROOT_ASSY"][0]

        with tempfile.TemporaryDirectory() as folder:
            folders = {
                "run": folder,
                "pdf": os.path.join(folder, "PDF"),
                "step": os.path.join(folder, "STEP"),
                "reports": os.path.join(folder, "REPORTS"),
                "logs": os.path.join(folder, "LOGS"),
            }
            for path in folders.values():
                os.makedirs(path, exist_ok=True)
            result = journal.process_target(
                session,
                target,
                folders,
                "TS",
                "EXPORTED: 2026-09-07 10:00 MYT",
                "APPLY",
                "BOM",
                "LOAD",
                "DRAWING_ONLY",
                {"load_status": "SUCCESS", "initial_load_state": "FULLY_LOADED",
                 "load_action": "NOT_REQUIRED", "final_load_state": "FULLY_LOADED"},
                {"halted": False, "reason": ""},
                tree["root"],
                tree["root"],
                [],
            )
            self.assertEqual("SUCCESS", result["PDF_RESULT"])
            self.assertEqual(1, result["PDF_FILE_COUNT"])
            self.assertEqual("SUCCESS", result["STEP_RESULT"])
            self.assertEqual("SUCCESS", result["OVERALL_RESULT"])
            pdf_path = result["PDF_FILES"].split(";")[0]
            self.assertTrue(os.path.isfile(pdf_path))
            self.assertTrue(pdf_path.endswith("264MN000001A01_REVA.V1.2.pdf"))
            self.assertTrue(os.path.isfile(result["STEP_FILE"]))
            # Native watermark and searchable text were applied.
            builder = drawing.PlotManager.builders[0]
            self.assertTrue(builder.AddWatermark)
            self.assertEqual("DRAFT_A.V1.2", builder.Watermark)
            self.assertEqual("NATIVE", builder.Action)
            self.assertEqual("TEXT", builder.OutputText)
            # Temporary timestamp notes were created then undone.
            self.assertEqual(2, len(drawing.Annotations.builders))
            # One drawing specification -> one PDF commit -> one undo mark.
            self.assertEqual(1, len(session.undo_to_mark_calls))
            self.assertEqual(1, len(session.deleted_marks))
            self.assertEqual(2, len(builder.SourceBuilder.sheets))
            # Work and display parts were restored to the assembly root.
            self.assertEqual(tree["root"], session.Parts.Work)
            self.assertEqual(tree["root"], session.Parts.Display)

    def test_process_target_apply_reports_existing_outputs(self):
        journal = self.journal
        tree = make_assembly()
        drawing = self.build_drawing_target(tree)
        session = FakeSession(work_part=tree["root"], loaded_parts=[drawing])
        targets, _diagnostics = journal.collect_session_scope(tree["root"], "BOM")
        target = [t for t in targets if t["part"].Name == "ROOT_ASSY"][0]

        with tempfile.TemporaryDirectory() as folder:
            folders = {
                "run": folder,
                "pdf": os.path.join(folder, "PDF"),
                "step": os.path.join(folder, "STEP"),
                "reports": os.path.join(folder, "REPORTS"),
                "logs": os.path.join(folder, "LOGS"),
            }
            for path in folders.values():
                os.makedirs(path, exist_ok=True)
            base = journal.build_versioned_base("264MN000001A01", "A", "V1.2")
            Path(os.path.join(folders["pdf"], base + ".pdf")).write_text(
                "existing", encoding="utf-8"
            )
            Path(os.path.join(folders["step"], base + ".stp")).write_text(
                "existing", encoding="utf-8"
            )
            result = journal.process_target(
                session,
                target,
                folders,
                "TS",
                "EXPORTED: 2026-09-07 10:00 MYT",
                "APPLY",
                "BOM",
                "LOAD",
                "DRAWING_ONLY",
                {"load_status": "SUCCESS", "initial_load_state": "FULLY_LOADED",
                 "load_action": "NOT_REQUIRED", "final_load_state": "FULLY_LOADED"},
                {"halted": False, "reason": ""},
                tree["root"],
                tree["root"],
                [],
            )
            self.assertEqual("SKIPPED_EXISTS", result["PDF_RESULT"])
            self.assertEqual("SKIPPED_EXISTS", result["STEP_RESULT"])
            self.assertEqual("SKIPPED_EXISTS", result["OVERALL_RESULT"])
            self.assertEqual(0, result["PDF_FILE_COUNT"])
            self.assertIn(base + ".pdf", result["PDF_SKIPPED_FILES"])
            self.assertEqual(
                "existing",
                Path(os.path.join(folders["pdf"], base + ".pdf")).read_text(
                    encoding="utf-8"
                ),
            )

    def test_process_target_without_drawing_skips_step_in_drawing_only_scope(self):
        journal = self.journal
        tree = make_assembly()
        leaf = tree["leaf_c"]
        session = FakeSession(work_part=tree["root"], loaded_parts=[])
        targets, _diagnostics = journal.collect_session_scope(tree["root"], "BOM")
        target = [t for t in targets if t["part"].Name == "LEAF_C"][0]

        with tempfile.TemporaryDirectory() as folder:
            folders = {
                "run": folder,
                "pdf": os.path.join(folder, "PDF"),
                "step": os.path.join(folder, "STEP"),
                "reports": os.path.join(folder, "REPORTS"),
                "logs": os.path.join(folder, "LOGS"),
            }
            for path in folders.values():
                os.makedirs(path, exist_ok=True)
            result = journal.process_target(
                session,
                target,
                folders,
                "TS",
                "EXPORTED: 2026-09-07 10:00 MYT",
                "APPLY",
                "BOM",
                "LOAD",
                "DRAWING_ONLY",
                {"load_status": "SUCCESS", "initial_load_state": "FULLY_LOADED",
                 "load_action": "NOT_REQUIRED", "final_load_state": "FULLY_LOADED"},
                {"halted": False, "reason": ""},
                tree["root"],
                tree["root"],
                [],
            )
            self.assertEqual("NO", result["HAS_DRAWING"])
            self.assertEqual("SKIPPED_NO_DRAWING", result["PDF_RESULT"])
            self.assertEqual("SKIPPED_NO_DRAWING", result["STEP_RESULT"])
            self.assertEqual("NO", result["STEP_REQUESTED"])
            self.assertEqual(0, session.DexManager.created)
            self.assertIn(
                "264MN000006A01", session.Parts.open_display_calls[0]
            )
            self.assertEqual(
                [], os.listdir(os.path.join(folder, "PDF"))
            )

    def test_process_target_step_scope_all_exports_without_drawing(self):
        journal = self.journal
        tree = make_assembly()
        creator = FakeStepCreator()
        session = FakeSession(
            work_part=tree["root"], loaded_parts=[], step_creator=creator
        )
        targets, _diagnostics = journal.collect_session_scope(tree["root"], "BOM")
        target = [t for t in targets if t["part"].Name == "LEAF_C"][0]

        with tempfile.TemporaryDirectory() as folder:
            folders = {
                "run": folder,
                "pdf": os.path.join(folder, "PDF"),
                "step": os.path.join(folder, "STEP"),
                "reports": os.path.join(folder, "REPORTS"),
                "logs": os.path.join(folder, "LOGS"),
            }
            for path in folders.values():
                os.makedirs(path, exist_ok=True)
            result = journal.process_target(
                session,
                target,
                folders,
                "TS",
                "EXPORTED: 2026-09-07 10:00 MYT",
                "APPLY",
                "BOM",
                "LOAD",
                "ALL",
                {"load_status": "SUCCESS", "initial_load_state": "FULLY_LOADED",
                 "load_action": "NOT_REQUIRED", "final_load_state": "FULLY_LOADED"},
                {"halted": False, "reason": ""},
                tree["root"],
                tree["root"],
                [],
            )
            self.assertEqual("SKIPPED_NO_DRAWING", result["PDF_RESULT"])
            self.assertEqual("SUCCESS", result["STEP_RESULT"])
            self.assertEqual("PARTIAL_SUCCESS", result["OVERALL_RESULT"])
            self.assertEqual(1, creator.committed)

    def test_process_target_reports_missing_identity_and_skips_tc_fallback(self):
        journal = self.journal
        part = FakePart("LOCAL_ONLY_PART", tag="local", sheets=[])
        session = FakeSession(work_part=part, loaded_parts=[])
        target = {
            "key": ("TAG", "local"),
            "part": part,
            "level": 0,
            "deepest_level": 0,
            "component_path": "LOCAL_ONLY_PART",
            "occurrence_count": 1,
            "is_work_part": True,
        }
        with tempfile.TemporaryDirectory() as folder:
            folders = {
                "run": folder,
                "pdf": os.path.join(folder, "PDF"),
                "step": os.path.join(folder, "STEP"),
                "reports": os.path.join(folder, "REPORTS"),
                "logs": os.path.join(folder, "LOGS"),
            }
            for path in folders.values():
                os.makedirs(path, exist_ok=True)
            result = journal.process_target(
                session,
                target,
                folders,
                "TS",
                "EXPORTED: 2026-09-07 10:00 MYT",
                "APPLY",
                "BOM",
                "LOAD",
                "DRAWING_ONLY",
                None,
                {"halted": False, "reason": ""},
                part,
                part,
                [],
            )
            self.assertEqual([], session.Parts.open_display_calls)
            self.assertEqual("LOCAL_ONLY_PART", result["DB_PART_NO"])
            self.assertEqual("SKIPPED_NO_DRAWING", result["OVERALL_RESULT"])
            self.assertIn("DB_PART_NO is unavailable", result["MESSAGE"])

    # --- end-to-end main() smoke test ---------------------------------

    def test_main_dry_run_writes_report_and_log(self):
        journal = self.journal
        tree = make_assembly()
        session = FakeSession(work_part=tree["root"], loaded_parts=[])
        session.IsManagedMode = False

        nxopen = self.nxopen
        nxopen.Session.GetSession = staticmethod(lambda: session)
        try:
            with tempfile.TemporaryDirectory() as folder:
                with env_overrides(
                    NX_J36_MODE="DRY_RUN",
                    NX_JOURNALS_IO_DIR=folder,
                ):
                    journal.main()
                run_root = os.path.join(folder, journal.OUTPUT_ROOT_FOLDER)
                runs = os.listdir(run_root)
                self.assertEqual(1, len(runs))
                run = os.path.join(run_root, runs[0])
                self.assertTrue(os.path.isdir(os.path.join(run, "PDF")))
                self.assertTrue(os.path.isdir(os.path.join(run, "STEP")))
                self.assertTrue(os.path.isdir(os.path.join(run, "REPORTS")))
                self.assertTrue(os.path.isdir(os.path.join(run, "LOGS")))
                self.assertEqual([], os.listdir(os.path.join(run, "PDF")))
                self.assertEqual([], os.listdir(os.path.join(run, "STEP")))

                reports = [
                    name
                    for name in os.listdir(os.path.join(run, "REPORTS"))
                    if name.startswith("EXPORT_RESULT_")
                ]
                self.assertEqual(1, len(reports))
                with open(
                    os.path.join(run, "REPORTS", reports[0]),
                    "r",
                    encoding="utf-8-sig",
                    newline="",
                ) as handle:
                    rows = list(csv.DictReader(handle))
                target_rows = [
                    row for row in rows if row["ROW_TYPE"] == "TARGET"
                ]
                summary_rows = [
                    row for row in rows if row["ROW_TYPE"] == "SUMMARY"
                ]
                self.assertEqual(1, len(summary_rows))
                self.assertEqual(
                    {"ROOT_ASSY", "SUB_ASSY", "LEAF_A", "SHARED", "LEAF_C"},
                    {row["PART_NAME"] for row in target_rows},
                )
                for row in target_rows:
                    self.assertEqual("DRY_RUN", row["WRITE_MODE"])
                    self.assertEqual("SKIPPED_NO_DRAWING", row["OVERALL_RESULT"])
                levels = sorted(int(row["LEVEL"]) for row in target_rows)
                self.assertEqual([0, 1, 1, 1, 2], levels)

                logs = os.listdir(os.path.join(run, "LOGS"))
                self.assertEqual(1, len(logs))
                with open(
                    os.path.join(run, "LOGS", logs[0]),
                    "r",
                    encoding="utf-8",
                ) as handle:
                    log_text = handle.read()
                self.assertIn("J36-NX2506-SESSION-TREE-PDF-STEP-V1", log_text)
                self.assertIn("Mode: DRY_RUN; scope: BOM; load: LOAD; step scope: DRAWING_ONLY", log_text)
                self.assertIn("DRY_RUN: no STEP/PDF file was written.", log_text)
        finally:
            # The fake NXOpen module is per-test-instance, so removing the
            # injected GetSession keeps later tests independent.
            nxopen.Session.__dict__.pop("GetSession", None)


if __name__ == "__main__":
    unittest.main()
