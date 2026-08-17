import importlib.util
import os
import sys
import tempfile
import types
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
JOURNAL = (
    ROOT
    / "from_git"
    / "journals"
    / "21_mass_surface_attribute_updater.py"
)


def load_journal():
    nxopen = types.ModuleType("NXOpen")
    nxopen.NXObject = types.SimpleNamespace(
        AttributeType=types.SimpleNamespace(String="String")
    )
    nxopen.BasePart = types.SimpleNamespace(
        SaveComponents=types.SimpleNamespace(FalseValue="FALSE"),
        CloseAfterSave=types.SimpleNamespace(FalseValue="FALSE"),
    )
    nxopen.MassPropertiesBuilder = types.SimpleNamespace(
        UpdateOptions=types.SimpleNamespace(Yes="YES")
    )
    nxopen.Session = types.SimpleNamespace()
    prior = sys.modules.get("NXOpen")
    sys.modules["NXOpen"] = nxopen
    try:
        spec = importlib.util.spec_from_file_location("journal21", JOURNAL)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        return module
    finally:
        if prior is None:
            sys.modules.pop("NXOpen", None)
        else:
            sys.modules["NXOpen"] = prior


def walk_prototypes(obj):
    """Simulate NX writing attributes to every part under the measured root."""
    prototype = getattr(obj, "Prototype", None)
    if prototype is not None:
        yield prototype
    for child in list(getattr(obj, "_children", [])):
        for part in walk_prototypes(child):
            yield part


class FakeManagerFactory:
    def __init__(self, name, events=None):
        self.name = name
        self.events = events
        self.builder_calls = []
        self.update_now_calls = 0
        self.commit_calls = 0

    def CreateMassPropertiesBuilder(self, objects):
        builder = FakeBuilder(self, objects)
        self.builder_calls.append(builder)
        return builder


class FakeBuilder:
    def __init__(self, manager, objects):
        self.manager = manager
        self.objects = list(objects)
        self.Accuracy = None
        self.RollUp = None
        self.UpdateOnSave = "NO"
        self.UpdateOptions = types.SimpleNamespace(Yes="YES")
        self.UpdateNow_called = False
        self.Commit_called = False

    def UpdateNow(self):
        self.UpdateNow_called = True
        self.manager.update_now_calls += 1
        for obj in self.objects:
            prototype = getattr(obj, "Prototype", None)
            targets = [prototype] if prototype is not None else []
            if not targets and getattr(obj, "real_attributes", None) is not None:
                targets = [obj]
            for part in targets:
                if self.manager.events is not None:
                    self.manager.events.append("update:" + part.Name)
                part.real_attributes["NX_MassPropRollupMass"] = 0.25
                part.real_attributes["NX_MassPropRollupArea"] = 20000.0
                part.real_attributes["NX_Mass"] = 0.1
                part.real_attributes["NX_Area"] = 8000.0

    def Commit(self):
        self.Commit_called = True
        self.manager.commit_calls += 1


class FakePart:
    _tag = 0

    def __init__(self, name, component_children=(), save_error=False):
        FakePart._tag += 1
        self.Name = name
        self.Leaf = name
        self.FullPath = "C:\\NX\\{0}.prt".format(name)
        self.Tag = FakePart._tag
        self.PropertiesManager = FakeManagerFactory("PropertiesManager")
        self.MeasureManager = FakeManagerFactory("MeasureManager")
        self.ComponentAssembly = types.SimpleNamespace(RootComponent=None)
        self.IsFullyLoaded = True
        self.PartLoadState = "FullyLoaded"
        self.IsReadOnly = False
        self.PDMPart = None
        self.real_attributes = {}
        self.string_attributes = {}
        self.save_error = save_error
        self.saved = False
        self.events = None
        if component_children:
            self.ComponentAssembly.RootComponent = FakeComponent(
                "ROOT-" + name,
                self,
                component_children,
            )

    def GetStringAttribute(self, title):
        if title not in self.string_attributes:
            raise AttributeError("No such attribute: " + title)
        return self.string_attributes[title]

    def GetUserAttribute(self, *args):
        raise AttributeError("unavailable")

    def GetRealAttribute(self, title):
        if title not in self.real_attributes:
            raise KeyError("No such attribute: " + title)
        return self.real_attributes[title]

    def Save(self, *args):
        if self.save_error:
            raise RuntimeError("NX part is not writable")
        self.saved = True
        if self.events is not None:
            self.events.append("save:" + self.Name)
        return types.SimpleNamespace(
            NumberUnsavedParts=0, NumberUnsavedObjects=0
        )


class FakeComponent:
    _tag = 0

    def __init__(
        self,
        name,
        prototype=None,
        children=(),
        suppressed=False,
        string_attributes=None,
    ):
        FakeComponent._tag += 1
        self.Name = name
        self.DisplayName = name
        self.Prototype = prototype
        self._children = list(children)
        self.IsSuppressed = suppressed
        self.string_attributes = dict(string_attributes or {})
        self.Tag = FakeComponent._tag

    def GetChildren(self):
        return list(self._children)

    def GetStringAttribute(self, title):
        if title not in self.string_attributes:
            raise AttributeError("No such attribute: " + title)
        return self.string_attributes[title]


class FakeListingWindow:
    def __init__(self):
        self.lines = []

    def Open(self):
        pass

    def WriteFullline(self, line):
        self.lines.append(line)


class FakeSession:
    def __init__(self, work_part, managed=False, user="aqil"):
        self.events = []
        self.Parts = FakePartCollection(work_part)
        self.ListingWindow = FakeListingWindow()
        self.IsManagedMode = managed
        self.PdmSession = FakePdmSession(user)
        if work_part is not None:
            root = work_part.ComponentAssembly.RootComponent
            targets = list(walk_prototypes(root)) if root is not None else [work_part]
            for part in targets:
                part.events = self.events
                part.PropertiesManager = FakeManagerFactory(
                    "PropertiesManager", self.events
                )
                part.MeasureManager = FakeManagerFactory(
                    "MeasureManager", self.events
                )


class FakePartCollection:
    def __init__(self, work_part):
        self.Work = work_part
        self.Display = work_part
        self.work_history = []

    def SetWork(self, part):
        self.Work = part
        self.work_history.append(part.Name)

    def SetDisplay(self, part, *_args):
        self.Display = part
        return part, None


class FakePdmSession:
    def __init__(self, user):
        self.user = user

    def GetUserName(self):
        return self.user


class FakePdmPart:
    def __init__(self, checked_out, owner=""):
        self.checked_out = checked_out
        self.owner = owner

    def GetCheckedoutStatusAndUser(self, *_args):
        return self.checked_out, self.owner


def rows_by_part_number(rows):
    return {
        row["DB_PART_NO"] or row["PART_NAME"]: row
        for row in rows
    }


class J21Tests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.j21 = load_journal()

    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp_dir.cleanup)
        os.environ["NX_JOURNALS_IO_DIR"] = self.temp_dir.name
        # Dry-run is the test default; APPLY tests remove it to exercise the
        # journal's built-in WRITE_MODE default.
        os.environ["NX_J21_MODE"] = "DRY_RUN"

    def make_assembly(self):
        child_a = FakePart("264MN000002A01")
        child_a.string_attributes = {
            "DB_PART_NO": "264MN000002A01",
            "DB_PART_REV": "A",
        }
        child_b = FakePart("264MN000003A01")
        child_b.string_attributes = {
            "DB_PART_NO": "264MN000003A01",
            "DB_PART_REV": "A",
        }
        root = FakePart(
            "264MN000001A01",
            component_children=[
                FakeComponent("CHILD-A-1", child_a),
                FakeComponent("CHILD-B-1", child_b),
            ],
        )
        root.string_attributes = {
            "DB_PART_NO": "264MN000001A01",
            "DB_PART_REV": "A",
        }
        return root, [child_a, child_b]

    def test_apply_triggers_native_update_reads_back_and_saves(self):
        root, children = self.make_assembly()
        session = FakeSession(root)

        os.environ.pop("NX_J21_MODE", None)  # exercise the APPLY default
        _path, rows, diagnostics = self.j21.run(session)

        for part in (children[0], children[1], root):
            manager = part.PropertiesManager
            self.assertEqual(1, len(manager.builder_calls))
            builder = manager.builder_calls[0]
            self.assertEqual(0.99, builder.Accuracy)
            self.assertIsNone(builder.RollUp)
            self.assertEqual("YES", builder.UpdateOnSave)
            self.assertTrue(builder.UpdateNow_called)
            self.assertTrue(builder.Commit_called)
            self.assertEqual(1, manager.update_now_calls)
            self.assertEqual(1, manager.commit_calls)

        self.assertEqual(
            [
                "update:264MN000002A01",
                "save:264MN000002A01",
                "update:264MN000003A01",
                "save:264MN000003A01",
                "update:264MN000001A01",
                "save:264MN000001A01",
            ],
            session.events,
        )
        self.assertIs(root, session.Parts.Work)
        self.assertIs(root, session.Parts.Display)

        self.assertFalse(diagnostics)
        by_number = rows_by_part_number(rows)
        self.assertEqual(3, len(rows))
        for part, number in (
            (root, "264MN000001A01"),
            (children[0], "264MN000002A01"),
            (children[1], "264MN000003A01"),
        ):
            row = by_number[number]
            self.assertEqual("0.250000", row["ROLLUP_MASS_KG"])
            self.assertEqual("20000.00", row["ROLLUP_AREA_MM2"])
            self.assertEqual("0.0200", row["ROLLUP_AREA_M2"])
            self.assertEqual("POPULATED", row["ROLLUP_MASS_ATTRIBUTE"])
            self.assertEqual("POPULATED", row["ROLLUP_AREA_ATTRIBUTE"])
            self.assertEqual("SAVED", row["SAVED"])
            self.assertEqual("SUCCESS", row["STATUS"])
            self.assertEqual("UPDATED", row["UPDATE"])
            self.assertTrue(part.saved)
        # NX itself wrote the standard attributes - no journal-created ones.
        self.assertEqual(
            "NX_MassPropRollupMass",
            self.j21.ROLLUP_MASS_ATTRIBUTE,
        )
        self.assertEqual(
            "NX_MassPropRollupArea",
            self.j21.ROLLUP_AREA_ATTRIBUTE,
        )

    def test_nested_assembly_is_processed_leaf_to_subassembly_to_root(self):
        leaf = FakePart("LEAF")
        leaf.string_attributes = {"DB_PART_NO": "LEAF", "DB_PART_REV": "A"}
        leaf_occurrence = FakeComponent("LEAF-1", leaf)
        subassembly = FakePart("SUB", component_children=[leaf_occurrence])
        subassembly.string_attributes = {
            "DB_PART_NO": "SUB",
            "DB_PART_REV": "A",
        }
        sub_occurrence = FakeComponent(
            "SUB-1", subassembly, children=[leaf_occurrence]
        )
        root = FakePart("TOP", component_children=[sub_occurrence])
        root.string_attributes = {"DB_PART_NO": "TOP", "DB_PART_REV": "A"}
        session = FakeSession(root)

        os.environ.pop("NX_J21_MODE", None)
        _path, rows, _diagnostics = self.j21.run(session)

        self.assertEqual(
            ["LEAF", "SUB", "TOP"],
            [row["DB_PART_NO"] for row in rows],
        )
        self.assertEqual([2, 1, 0], [row["LEVEL"] for row in rows])
        self.assertEqual(
            ["LEAF", "ASSEMBLY", "ASSEMBLY"],
            [row["PART_KIND"] for row in rows],
        )
        self.assertEqual(
            [
                "update:LEAF",
                "save:LEAF",
                "update:SUB",
                "save:SUB",
                "update:TOP",
                "save:TOP",
            ],
            session.events,
        )

    def test_shared_prototype_is_updated_once_at_its_deepest_level(self):
        shared = FakePart("SHARED")
        shared.string_attributes = {
            "DB_PART_NO": "SHARED",
            "DB_PART_REV": "A",
        }
        nested_occurrence = FakeComponent("SHARED-NESTED", shared)
        subassembly = FakePart(
            "SUB", component_children=[nested_occurrence]
        )
        subassembly.string_attributes = {
            "DB_PART_NO": "SUB",
            "DB_PART_REV": "A",
        }
        root = FakePart(
            "TOP",
            component_children=[
                FakeComponent("SHARED-DIRECT", shared),
                FakeComponent(
                    "SUB-1", subassembly, children=[nested_occurrence]
                ),
            ],
        )
        root.string_attributes = {"DB_PART_NO": "TOP", "DB_PART_REV": "A"}
        session = FakeSession(root)

        os.environ.pop("NX_J21_MODE", None)
        _path, rows, _diagnostics = self.j21.run(session)

        shared_rows = [row for row in rows if row["DB_PART_NO"] == "SHARED"]
        self.assertEqual(1, len(shared_rows))
        self.assertEqual(2, shared_rows[0]["LEVEL"])
        self.assertEqual(1, len(shared.PropertiesManager.builder_calls))

    def test_apply_aborts_before_any_change_when_target_is_not_fully_loaded(self):
        root, children = self.make_assembly()
        children[1].IsFullyLoaded = False
        children[1].PartLoadState = "MinimallyLoaded"
        session = FakeSession(root)

        os.environ.pop("NX_J21_MODE", None)
        with self.assertRaisesRegex(RuntimeError, "FULL_LOAD_REQUIRED"):
            self.j21.run(session)

        self.assertEqual([], session.events)
        self.assertEqual([], session.Parts.work_history)
        for part in (root, children[0], children[1]):
            self.assertFalse(part.saved)
            self.assertEqual(0, len(part.PropertiesManager.builder_calls))

    def test_part_checked_out_by_other_user_is_skipped_and_run_continues(self):
        root, children = self.make_assembly()
        root.PDMPart = FakePdmPart(True, "aqil")
        children[0].PDMPart = FakePdmPart(True, "aqil")
        children[1].PDMPart = FakePdmPart(True, "other.user")
        children[1].IsReadOnly = True
        session = FakeSession(root, managed=True, user="aqil")

        os.environ.pop("NX_J21_MODE", None)
        _path, rows, _diagnostics = self.j21.run(session)

        blocked = rows_by_part_number(rows)["264MN000003A01"]
        self.assertEqual("CHECKED_OUT", blocked["CHECKOUT_STATE"])
        self.assertEqual("other.user", blocked["CHECKOUT_OWNER"])
        self.assertEqual("SKIPPED_NOT_WRITABLE", blocked["UPDATE"])
        self.assertEqual("NOT_SAVED", blocked["SAVED"])
        self.assertEqual("SKIPPED", blocked["STATUS"])
        self.assertIn("another user", blocked["MESSAGE"])
        self.assertFalse(children[1].saved)
        self.assertTrue(children[0].saved)
        self.assertTrue(root.saved)

    def test_original_work_part_is_restored_when_root_is_not_writable(self):
        root, children = self.make_assembly()
        root.PDMPart = FakePdmPart(False, "")
        root.IsReadOnly = True
        for child in children:
            child.PDMPart = FakePdmPart(True, "aqil")
        session = FakeSession(root, managed=True, user="aqil")

        os.environ.pop("NX_J21_MODE", None)
        _path, rows, diagnostics = self.j21.run(session)

        root_row = rows_by_part_number(rows)["264MN000001A01"]
        self.assertEqual("SKIPPED", root_row["STATUS"])
        self.assertIs(root, session.Parts.Work)
        self.assertIs(root, session.Parts.Display)
        self.assertFalse(diagnostics)

    def test_apply_falls_back_to_measure_manager(self):
        root, children = self.make_assembly()
        session = FakeSession(root)
        root.PropertiesManager = None

        os.environ.pop("NX_J21_MODE", None)
        _path, rows, _diagnostics = self.j21.run(session)

        manager = root.MeasureManager
        self.assertEqual(1, len(manager.builder_calls))
        self.assertTrue(manager.builder_calls[0].UpdateNow_called)
        row = rows_by_part_number(rows)["264MN000001A01"]
        self.assertEqual("0.250000", row["ROLLUP_MASS_KG"])
        self.assertEqual("0.0200", row["ROLLUP_AREA_M2"])
        self.assertEqual("SAVED", row["SAVED"])

    def test_smoke_runs_update_on_work_part_only(self):
        root, children = self.make_assembly()
        session = FakeSession(root)

        os.environ["NX_J21_MODE"] = "SMOKE"
        _path, rows, _diagnostics = self.j21.run(session)

        manager = root.PropertiesManager
        self.assertEqual(1, len(manager.builder_calls))
        # SMOKE measures only the work assembly root, not each child target.
        self.assertEqual(
            [root.ComponentAssembly.RootComponent],
            manager.builder_calls[0].objects,
        )
        self.assertTrue(manager.builder_calls[0].UpdateNow_called)
        self.assertTrue(manager.builder_calls[0].Commit_called)
        self.assertEqual(1, len(rows))
        self.assertEqual("264MN000001A01", rows[0]["DB_PART_NO"])
        self.assertEqual("POPULATED", rows[0]["ROLLUP_MASS_ATTRIBUTE"])
        self.assertEqual("SAVED", rows[0]["SAVED"])
        self.assertTrue(root.saved)
        self.assertFalse(children[0].saved)
        self.assertFalse(children[1].saved)

    def test_dry_run_reports_current_values_without_update_or_save(self):
        root, children = self.make_assembly()
        for part in (root, children[0], children[1]):
            part.real_attributes["NX_MassPropRollupMass"] = 0.5
            part.real_attributes["NX_MassPropRollupArea"] = 10000.0
        session = FakeSession(root)

        _path, rows, _diagnostics = self.j21.run(session)

        manager = root.PropertiesManager
        self.assertEqual(0, manager.update_now_calls)
        self.assertEqual(0, len(manager.builder_calls))
        for part in (root, children[0], children[1]):
            self.assertFalse(part.saved)
        row = rows_by_part_number(rows)["264MN000002A01"]
        self.assertEqual("0.500000", row["ROLLUP_MASS_KG"])
        self.assertEqual("10000.00", row["ROLLUP_AREA_MM2"])
        self.assertEqual("0.0100", row["ROLLUP_AREA_M2"])
        self.assertEqual("STORED", row["ROLLUP_MASS_ATTRIBUTE"])
        self.assertEqual("DRY_RUN", row["SAVED"])
        self.assertEqual("SUCCESS", row["STATUS"])

    def test_boM_visibility_filters_noise_from_scope(self):
        child_a = FakePart("264MN000002A01")
        child_a.string_attributes = {
            "DB_PART_NO": "264MN000002A01",
            "DB_PART_REV": "A",
        }
        suppressed = FakePart("264MN000003A01")
        suppressed.string_attributes = {
            "DB_PART_NO": "264MN000003A01",
            "DB_PART_REV": "A",
        }
        reference = FakePart("264MN000004A01")
        reference.string_attributes = {
            "DB_PART_NO": "264MN000004A01",
            "DB_PART_REV": "A",
        }
        csys = FakePart("264MN000005A01")
        csys.string_attributes = {
            "DB_PART_NO": "264MN000005A01",
            "DB_PART_REV": "A",
        }
        root = FakePart(
            "264MN000001A01",
            component_children=[
                FakeComponent("CHILD-A-1", child_a),
                FakeComponent("SUPPRESSED-1", suppressed, suppressed=True),
                FakeComponent(
                    "REFERENCE-1",
                    reference,
                    string_attributes={"REFERENCE_COMPONENT": ""},
                ),
                FakeComponent("CSYS_ORIGIN", csys),
            ],
        )
        root.string_attributes = {
            "DB_PART_NO": "264MN000001A01",
            "DB_PART_REV": "A",
        }
        session = FakeSession(root)

        os.environ.pop("NX_J21_MODE", None)
        _path, rows, _diagnostics = self.j21.run(session)

        numbers = [row["DB_PART_NO"] for row in rows]
        self.assertEqual(["264MN000002A01", "264MN000001A01"], numbers)
        self.assertTrue(child_a.saved)
        self.assertFalse(suppressed.saved)
        self.assertFalse(reference.saved)
        self.assertFalse(csys.saved)

    def test_blank_attributes_report_partial(self):
        root, children = self.make_assembly()
        session = FakeSession(root)

        # Simulate NX not writing attributes for one part (e.g. empty refset).
        class LimitedManager(FakeManagerFactory):
            def __init__(self):
                super().__init__("PropertiesManager")

            def CreateMassPropertiesBuilder(self, objects):
                builder = FakeBuilder(self, objects)
                self.builder_calls.append(builder)

                def limited_update():
                    builder.UpdateNow_called = True
                    self.update_now_calls += 1
                    # Builder runs, but NX leaves the measured child's
                    # reserved attributes blank.

                builder.UpdateNow = limited_update
                return builder

        children[1].PropertiesManager = LimitedManager()

        os.environ.pop("NX_J21_MODE", None)
        _path, rows, _diagnostics = self.j21.run(session)

        row = rows_by_part_number(rows)["264MN000003A01"]
        self.assertEqual("", row["ROLLUP_MASS_KG"])
        self.assertEqual("BLANK", row["ROLLUP_MASS_ATTRIBUTE"])
        self.assertEqual("BLANK", row["ROLLUP_AREA_ATTRIBUTE"])
        self.assertEqual("PARTIAL", row["STATUS"])
        self.assertIn("MASS ATTRIBUTE:", row["MESSAGE"])
        self.assertIn("AREA ATTRIBUTE:", row["MESSAGE"])

    def test_apply_save_failure_is_reported_and_run_continues(self):
        root, children = self.make_assembly()
        children[1].save_error = True
        session = FakeSession(root)

        os.environ.pop("NX_J21_MODE", None)
        _path, rows, _diagnostics = self.j21.run(session)

        by_number = rows_by_part_number(rows)
        self.assertEqual("SAVED", by_number["264MN000002A01"]["SAVED"])
        self.assertEqual("SAVE_FAILED", by_number["264MN000003A01"]["SAVED"])
        self.assertEqual("SAVE_FAILED", by_number["264MN000003A01"]["STATUS"])
        self.assertIn("SAVE:", by_number["264MN000003A01"]["MESSAGE"])

    def test_probe_returns_builder_api_surface(self):
        root, _children = self.make_assembly()
        session = FakeSession(root)

        os.environ["NX_J21_MODE"] = "PROBE"
        path, rows, diagnostics = self.j21.run(session)

        self.assertIsNone(path)
        self.assertTrue(rows)
        self.assertTrue(
            any("PropertiesManager members" in line for line in rows)
        )
        self.assertTrue(
            any("MassPropertiesBuilder via" in line for line in rows)
        )
        self.assertFalse(diagnostics)

    def test_no_work_part_raises(self):
        session = FakeSession(None)
        with self.assertRaises(RuntimeError):
            self.j21.run(session)

    def test_native_only_no_manual_compute_or_attribute_write(self):
        source = JOURNAL.read_text(encoding="utf-8")
        self.assertNotIn("NewFaceProperties", source)
        self.assertNotIn("NewMassProperties", source)
        # The journal must not CALL the attribute writer (the docstring may
        # mention it when explaining why the native update is required).
        self.assertNotIn("CreateAttributePropertiesBuilder", source)
        self.assertNotIn("SetRealAttribute", source)
        self.assertIn("CreateMassPropertiesBuilder", source)
        self.assertIn("PropertiesManager", source)
        self.assertIn("UpdateNow", source)
        self.assertIn("NXOpenBoMExtended", source)
        self.assertIn("GetCheckedoutStatusAndUser", source)
        self.assertNotIn(".Checkout(", source)
        self.assertNotIn("CheckoutParts", source)
        self.assertNotIn("LoadThisPartFully", source)
        self.assertNotIn(".LoadFully(", source)


if __name__ == "__main__":
    unittest.main()
