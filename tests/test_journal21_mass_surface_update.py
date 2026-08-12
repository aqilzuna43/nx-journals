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


class FakeBuilder:
    def __init__(self, manager, objects):
        self.manager = manager
        self.objects = list(objects)
        self.Accuracy = None
        self.RollUp = False
        self.UpdateOnSave = "NO"
        self.UpdateNow_called = False

    def UpdateNow(self):
        self.UpdateNow_called = True
        self.manager.update_now_calls += 1
        for obj in self.objects:
            for part in walk_prototypes(obj):
                part.real_attributes["NX_MassPropRollupMass"] = 0.25
                part.real_attributes["NX_MassPropRollupArea"] = 20000.0
                part.real_attributes["NX_Mass"] = 0.1
                part.real_attributes["NX_Area"] = 8000.0


class FakeMeasureManager:
    def __init__(self):
        self.builder_calls = []
        self.update_now_calls = 0

    def CreateMassPropertiesBuilder(self, objects):
        builder = FakeBuilder(self, objects)
        self.builder_calls.append(builder)
        return builder


class FakePart:
    _tag = 0

    def __init__(self, name, component_children=(), save_error=False):
        FakePart._tag += 1
        self.Name = name
        self.Leaf = name
        self.Tag = FakePart._tag
        self.MeasureManager = None
        self.ComponentAssembly = types.SimpleNamespace(RootComponent=None)
        self.real_attributes = {}
        self.string_attributes = {}
        self.save_error = save_error
        self.saved = False
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
    def __init__(self, work_part):
        self.Parts = types.SimpleNamespace(Work=work_part)
        self.ListingWindow = FakeListingWindow()
        if work_part is not None:
            work_part.MeasureManager = FakeMeasureManager()


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

        manager = root.MeasureManager
        self.assertEqual(1, len(manager.builder_calls))
        builder = manager.builder_calls[0]
        self.assertEqual(0.99, builder.Accuracy)
        self.assertIs(True, builder.RollUp)
        self.assertEqual("YES", builder.UpdateOnSave)
        self.assertTrue(builder.UpdateNow_called)
        self.assertEqual(1, manager.update_now_calls)

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
            self.assertEqual("POPULATED", row["ROLLUP_MASS_ATTRIBUTE"])
            self.assertEqual("POPULATED", row["ROLLUP_AREA_ATTRIBUTE"])
            self.assertEqual("SAVED", row["SAVED"])
            self.assertEqual("SUCCESS", row["STATUS"])
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

    def test_dry_run_reports_current_values_without_update_or_save(self):
        root, children = self.make_assembly()
        for part in (root, children[0], children[1]):
            part.real_attributes["NX_MassPropRollupMass"] = 0.5
            part.real_attributes["NX_MassPropRollupArea"] = 10000.0
        session = FakeSession(root)

        _path, rows, _diagnostics = self.j21.run(session)

        manager = root.MeasureManager
        self.assertEqual(0, manager.update_now_calls)
        self.assertEqual(0, len(manager.builder_calls))
        for part in (root, children[0], children[1]):
            self.assertFalse(part.saved)
        row = rows_by_part_number(rows)["264MN000002A01"]
        self.assertEqual("0.500000", row["ROLLUP_MASS_KG"])
        self.assertEqual("10000.00", row["ROLLUP_AREA_MM2"])
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
        self.assertEqual(["264MN000001A01", "264MN000002A01"], numbers)
        self.assertTrue(child_a.saved)
        self.assertFalse(suppressed.saved)
        self.assertFalse(reference.saved)
        self.assertFalse(csys.saved)

    def test_blank_attributes_report_partial(self):
        root, children = self.make_assembly()
        session = FakeSession(root)

        # Simulate NX not writing attributes for one part (e.g. empty refset).
        class LimitedManager(FakeMeasureManager):
            def CreateMassPropertiesBuilder(self, objects):
                builder = FakeBuilder(self, objects)
                self.builder_calls.append(builder)

                def limited_update():
                    builder.UpdateNow_called = True
                    self.update_now_calls += 1
                    for obj in objects:
                        for part in walk_prototypes(obj):
                            if part.Name == "264MN000003A01":
                                continue
                            part.real_attributes["NX_MassPropRollupMass"] = 0.25
                            part.real_attributes["NX_MassPropRollupArea"] = 20000.0

                builder.UpdateNow = limited_update
                return builder

        root.MeasureManager = LimitedManager()

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
            any("MassPropertiesBuilder members" in line for line in rows)
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
        self.assertNotIn("AttributePropertiesBuilder", source)
        self.assertNotIn("SetRealAttribute", source)
        self.assertIn("CreateMassPropertiesBuilder", source)
        self.assertIn("RollUp", source)
        self.assertIn("UpdateNow", source)
        self.assertIn("NXOpenBoMExtended", source)
        self.assertNotIn("Checkout", source)


if __name__ == "__main__":
    unittest.main()
