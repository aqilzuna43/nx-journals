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
    nxopen.AttributePropertiesBuilder = types.SimpleNamespace(
        OperationType=types.SimpleNamespace(Save="SAVE")
    )
    nxopen.AttributePropertiesBaseBuilder = types.SimpleNamespace(
        DataTypeOptions=types.SimpleNamespace(Number="NUMBER")
    )
    nxopen.BasePart = types.SimpleNamespace(
        SaveComponents=types.SimpleNamespace(FalseValue="FALSE"),
        CloseAfterSave=types.SimpleNamespace(FalseValue="FALSE"),
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


class FakeUnit:
    def __init__(self, name, symbol, measure):
        self.Name = name
        self.Symbol = symbol
        self.Abbreviation = symbol
        self.TypeName = name
        self.Measure = measure


class FakeUnits:
    def __init__(self):
        self.area = FakeUnit("SquareMeter", "m^2", "Area")
        self.length = FakeUnit("Meter", "m", "Length")
        self.mass = FakeUnit("Kilogram", "kg", "Mass")

    def FindObject(self, name):
        if name in ("SquareMeter", "SquareMetre"):
            return self.area
        if name in ("Meter", "Metre"):
            return self.length
        if name == "Kilogram":
            return self.mass
        raise RuntimeError("not found: " + name)

    def GetMeasureTypes(self, measure):
        if measure == "Area":
            return [self.area]
        if measure == "Length":
            return [self.length]
        if measure == "Mass":
            return [self.mass]
        return []


class FakeAreaMeasurement:
    def __init__(self, area):
        self.Area = area
        self.disposed = False

    def Dispose(self):
        self.disposed = True


class FakeMassMeasurement:
    def __init__(self, mass):
        self.Mass = mass
        self.disposed = False

    def Dispose(self):
        self.disposed = True


class FakeSaveStatus:
    def __init__(self):
        self.NumberUnsavedParts = 0
        self.NumberUnsavedObjects = 0

    def Dispose(self):
        pass


class FakeAttributeBuilder:
    def __init__(self, manager, target, objects, operation):
        self.manager = manager
        self.target = target
        self.objects = objects
        self.operation = operation
        self.Category = ""
        self.Title = ""
        self.DataType = ""
        self.NumberValue = None
        self.committed = False

    def Commit(self):
        if self.manager.fail_commit_categories and (
            self.Category in self.manager.fail_commit_categories
        ):
            raise RuntimeError("NX attribute commit failed")
        self.committed = True
        self.manager.writes.append(
            (
                self.target,
                self.Category,
                self.Title,
                self.DataType,
                self.NumberValue,
            )
        )
        # Simulate NX storing the attribute in the part immediately.
        self.target.real_attributes[self.Title] = self.NumberValue


class FakeAttributeManager:
    def __init__(self, fail_commit_categories=None):
        self.writes = []
        self.fail_commit_categories = tuple(fail_commit_categories or ())

    def CreateAttributePropertiesBuilder(
        self, target, objects, operation
    ):
        return FakeAttributeBuilder(self, target, objects, operation)


class FakeMeasureManager:
    def __init__(self, areas, masses, failing_face=None):
        self.areas = areas
        self.masses = masses
        self.failing_face = failing_face
        self.face_calls = []
        self.mass_calls = []

    def NewFaceProperties(self, area_unit, length_unit, accuracy, faces):
        self.face_calls.append((area_unit, length_unit, accuracy, list(faces)))
        key = faces[0]
        if key == self.failing_face:
            raise RuntimeError("NX face measurement failed")
        return FakeAreaMeasurement(self.areas[key])

    def NewMassProperties(self, mass_units, accuracy, bodies):
        self.mass_calls.append((list(mass_units), accuracy, list(bodies)))
        key = tuple(sorted(body.Tag for body in bodies))
        return FakeMassMeasurement(self.masses[key])


class FakeBody:
    _tag = 0

    def __init__(self, name, faces, solid=True, sheet=False, convergent=False):
        FakeBody._tag += 1
        self.Name = name
        self.Tag = FakeBody._tag
        self.IsSolidBody = solid
        self.IsSheetBody = sheet
        self.IsConvergentBody = convergent
        self._faces = list(faces)

    def GetFaces(self):
        return list(self._faces)


class FakeFace:
    _tag = 1000

    def __init__(self):
        FakeFace._tag += 1
        self.Tag = FakeFace._tag


def body_with_area(name, area):
    face = FakeFace()
    body = FakeBody(name, [face])
    return body, face


class FakePart:
    _tag = 0

    def __init__(
        self,
        name,
        bodies,
        component_children=(),
        attributes=None,
        save_error=False,
    ):
        FakePart._tag += 1
        self.Name = name
        self.Leaf = name
        self.Tag = FakePart._tag
        self.Bodies = list(bodies)
        self.MeasureManager = None
        self.UnitCollection = None
        self.ComponentAssembly = types.SimpleNamespace(RootComponent=None)
        self.attributes = dict(attributes or {})
        self.real_attributes = {}
        self.save_error = save_error
        self.saved = False
        if component_children:
            self.ComponentAssembly.RootComponent = FakeComponent(
                "ROOT-" + name,
                self,
                component_children,
            )

    def GetStringAttribute(self, title):
        if title not in self.attributes:
            raise AttributeError("No such attribute: " + title)
        return self.attributes[title]

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
        return FakeSaveStatus()


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
    def __init__(self, work_part, fail_commit_categories=None):
        self.Parts = types.SimpleNamespace(Work=work_part)
        self.ListingWindow = FakeListingWindow()
        self.AttributeManager = FakeAttributeManager(fail_commit_categories)
        if work_part is not None:
            work_part.MeasureManager = work_part.MeasureManager
            work_part.UnitCollection = work_part.UnitCollection


def rows_by_part_number(rows):
    return {
        row["DB_PART_NO"] or row["PART_NAME"]: row
        for row in rows
    }


def make_manager(areas, masses, failing_face=None):
    return FakeMeasureManager(areas, masses, failing_face)


class J21Tests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.j21 = load_journal()

    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp_dir.cleanup)
        os.environ["NX_JOURNALS_IO_DIR"] = self.temp_dir.name
        os.environ["NX_J21_MODE"] = "DRY_RUN"

    def make_leaf(self, area=2.5, mass=1.25):
        body, face = body_with_area("BODY1", area)
        part = FakePart("264MN000001A01", [body], attributes={
            "DB_PART_NO": "264MN000001A01",
            "DB_PART_REV": "A",
        })
        part.UnitCollection = FakeUnits()
        part.MeasureManager = make_manager(
            {face: area}, {(body.Tag,): mass}
        )
        return part, body

    def make_assembly(self):
        child_a_body, child_a_face = body_with_area("CHILD_A_BODY", 3.0)
        child_b_body, child_b_face = body_with_area("CHILD_B_BODY", 4.0)
        child_a = FakePart(
            "264MN000002A01",
            [child_a_body],
            attributes={"DB_PART_NO": "264MN000002A01", "DB_PART_REV": "A"},
        )
        child_b = FakePart(
            "264MN000003A01",
            [child_b_body],
            attributes={"DB_PART_NO": "264MN000003A01", "DB_PART_REV": "A"},
        )
        root_body, root_face = body_with_area("ROOT_BODY", 1.0)
        root = FakePart(
            "264MN000001A01",
            [root_body],
            component_children=[
                FakeComponent("CHILD-A-1", child_a),
                FakeComponent("CHILD-B-1", child_b),
            ],
            attributes={"DB_PART_NO": "264MN000001A01", "DB_PART_REV": "A"},
        )
        root.UnitCollection = FakeUnits()
        root.MeasureManager = make_manager(
            {
                root_face: 1.0,
                child_a_face: 3.0,
                child_b_face: 4.0,
            },
            {
                (root_body.Tag,): 1.0,
                (child_a_body.Tag,): 2.0,
                (child_b_body.Tag,): 3.0,
                tuple(sorted((root_body.Tag, child_a_body.Tag))): 3.0,
                tuple(sorted((root_body.Tag, child_a_body.Tag, child_b_body.Tag))): 6.0,
                tuple(sorted((root_body.Tag, child_b_body.Tag))): 4.0,
            },
        )
        return root, [child_a, child_b]

    def test_dry_run_computes_without_writing(self):
        part, _body = self.make_leaf()
        session = FakeSession(part)

        _path, rows, _diagnostics = self.j21.run(session)

        row = rows[0]
        self.assertEqual("2.5000", row["ROLLUP_AREA_M2"])
        self.assertEqual("1.250000", row["ROLLUP_MASS_KG"])
        self.assertEqual("DRY_RUN", row["ROLLUP_AREA_ATTRIBUTE"])
        self.assertEqual("DRY_RUN", row["ROLLUP_MASS_ATTRIBUTE"])
        self.assertEqual("DRY_RUN", row["SAVED"])
        self.assertEqual("SUCCESS", row["STATUS"])
        self.assertEqual([], session.AttributeManager.writes)
        self.assertFalse(part.saved)

    def test_apply_writes_standard_attributes_and_saves(self):
        part, _body = self.make_leaf()
        session = FakeSession(part)

        os.environ.pop("NX_J21_MODE", None)  # exercise APPLY default
        _path, rows, _diagnostics = self.j21.run(session)

        row = rows[0]
        self.assertEqual("WRITTEN", row["ROLLUP_AREA_ATTRIBUTE"])
        self.assertEqual("WRITTEN", row["ROLLUP_MASS_ATTRIBUTE"])
        self.assertEqual("SAVED", row["SAVED"])
        self.assertEqual("SUCCESS", row["STATUS"])
        self.assertTrue(part.saved)
        writes = session.AttributeManager.writes
        self.assertEqual(
            [
                (
                    part,
                    "Rolled-Up Mass Properties",
                    "NX_MassPropRollupArea",
                    "NUMBER",
                    2.5 * 1_000_000.0,
                ),
                (
                    part,
                    "Rolled-Up Mass Properties",
                    "NX_MassPropRollupMass",
                    "NUMBER",
                    1.25,
                ),
            ],
            writes,
        )
        # Read-back verification passed: attributes readable after write.
        self.assertNotIn("VERIFY:", row["MESSAGE"])

    def test_category_fallback_when_standard_rejected(self):
        part, _body = self.make_leaf()
        session = FakeSession(
            part,
            fail_commit_categories=("Rolled-Up Mass Properties",),
        )

        os.environ.pop("NX_J21_MODE", None)
        _path, rows, _diagnostics = self.j21.run(session)

        row = rows[0]
        self.assertEqual("WRITTEN", row["ROLLUP_AREA_ATTRIBUTE"])
        self.assertEqual("WRITTEN", row["ROLLUP_MASS_ATTRIBUTE"])
        self.assertIn("fallback category Materials", row["MESSAGE"])
        self.assertEqual("SUCCESS", row["STATUS"])

    def test_apply_write_failure_reported(self):
        part, _body = self.make_leaf()
        session = FakeSession(
            part,
            fail_commit_categories=(
                "Rolled-Up Mass Properties",
                "Materials",
            ),
        )

        os.environ.pop("NX_J21_MODE", None)
        _path, rows, _diagnostics = self.j21.run(session)

        row = rows[0]
        self.assertEqual("WRITE_FAILED", row["ROLLUP_AREA_ATTRIBUTE"])
        self.assertEqual("WRITE_FAILED", row["ROLLUP_MASS_ATTRIBUTE"])
        self.assertEqual("PARTIAL", row["STATUS"])
        self.assertIn("AREA ATTRIBUTE:", row["MESSAGE"])
        self.assertIn("MASS ATTRIBUTE:", row["MESSAGE"])

    def test_rollup_scope_and_visibility_filter(self):
        root, children = self.make_assembly()
        suppressed_body, _ = body_with_area("SUPPRESSED_BODY", 99.0)
        suppressed = FakePart(
            "264MN000004A01",
            [suppressed_body],
            attributes={"DB_PART_NO": "264MN000004A01", "DB_PART_REV": "A"},
        )
        root.ComponentAssembly.RootComponent._children.append(
            FakeComponent("SUPPRESSED-1", suppressed, suppressed=True)
        )

        os.environ.pop("NX_J21_MODE", None)
        session = FakeSession(root)
        _path, rows, _diagnostics = self.j21.run(session)

        by_number = rows_by_part_number(rows)
        self.assertEqual(
            ["264MN000001A01", "264MN000002A01", "264MN000003A01"],
            [row["DB_PART_NO"] for row in rows],
        )
        # Root roll-up includes own + both children (mass 1+2+3=6).
        self.assertEqual("6.000000", by_number["264MN000001A01"]["ROLLUP_MASS_KG"])
        self.assertEqual(3, by_number["264MN000001A01"]["ROLLUP_SOLID_BODY_COUNT"])
        # Child roll-up is its own mass only.
        self.assertEqual("2.000000", by_number["264MN000002A01"]["ROLLUP_MASS_KG"])
        self.assertNotIn("264MN000004A01", by_number)

    def test_smoke_runs_work_part_only(self):
        part, _body = self.make_leaf()
        session = FakeSession(part)

        os.environ["NX_J21_MODE"] = "SMOKE"
        _path, rows, _diagnostics = self.j21.run(session)

        self.assertEqual(1, len(rows))
        self.assertEqual("264MN000001A01", rows[0]["DB_PART_NO"])
        self.assertEqual("2.5000", rows[0]["ROLLUP_AREA_M2"])
        self.assertEqual("1.250000", rows[0]["ROLLUP_MASS_KG"])
        self.assertEqual("WRITTEN", rows[0]["ROLLUP_MASS_ATTRIBUTE"])
        self.assertEqual("SAVED", rows[0]["SAVED"])
        self.assertTrue(part.saved)

    def test_save_failure_reported_and_run_continues(self):
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

    def test_measurement_failure_is_fail_closed(self):
        good_body, good_face = body_with_area("GOOD_BODY", 1.0)
        failing_body, failing_face = body_with_area("FAILING_BODY", 2.0)
        part = FakePart(
            "264MN000001A01",
            [good_body, failing_body],
            attributes={"DB_PART_NO": "264MN000001A01", "DB_PART_REV": "A"},
        )
        part.UnitCollection = FakeUnits()
        part.MeasureManager = make_manager(
            {good_face: 1.0, failing_face: 2.0},
            {(good_body.Tag, failing_body.Tag): 3.0},
            failing_face=failing_face,
        )
        session = FakeSession(part)

        _path, rows, _diagnostics = self.j21.run(session)

        row = rows[0]
        self.assertEqual("", row["ROLLUP_AREA_M2"])
        self.assertEqual("FAILED", row["ROLLUP_AREA_ATTRIBUTE"])
        self.assertEqual("PARTIAL", row["STATUS"])
        self.assertIn("AREA:", row["MESSAGE"])
        self.assertEqual("3.000000", row["ROLLUP_MASS_KG"])

    def test_no_work_part_raises(self):
        session = FakeSession(None)
        with self.assertRaises(RuntimeError):
            self.j21.run(session)

    def test_static_proven_mechanism_no_native_builder(self):
        source = JOURNAL.read_text(encoding="utf-8")
        self.assertIn("NewFaceProperties", source)
        self.assertIn("NewMassProperties", source)
        self.assertIn("AttributePropertiesBuilder", source)
        self.assertNotIn("CreateMassPropertiesBuilder", source)
        self.assertNotIn("Checkout", source)
        self.assertIn("NXOpenBoMExtended", source)


if __name__ == "__main__":
    unittest.main()
