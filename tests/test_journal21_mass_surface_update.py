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
        self.fail_commit = getattr(manager, "fail_commit", False)

    def Commit(self):
        if self.fail_commit:
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


class FakeAttributeManager:
    def __init__(self, fail_commit=False):
        self.writes = []
        self.fail_commit = fail_commit

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

    def NewFaceProperties(
        self, area_unit, length_unit, accuracy, faces
    ):
        self.face_calls.append((area_unit, length_unit, accuracy, list(faces)))
        key = faces[0]
        if key == self.failing_face:
            raise RuntimeError("NX face measurement failed")
        return FakeAreaMeasurement(self.areas[key])

    def NewMassProperties(self, mass_units, accuracy, bodies):
        self.mass_calls.append(
            (list(mass_units), accuracy, list(bodies))
        )
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
        self.ComponentAssembly = types.SimpleNamespace(
            RootComponent=None
        )
        self.attributes = dict(attributes or {})
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
    def __init__(self, work_part, manager, units, fail_commit=False):
        self.Parts = types.SimpleNamespace(Work=work_part)
        self.ListingWindow = FakeListingWindow()
        self.AttributeManager = FakeAttributeManager(fail_commit=fail_commit)
        if work_part is not None:
            work_part.MeasureManager = manager
            work_part.UnitCollection = units


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
        os.environ["NX_J21_MODE"] = "DRY_RUN"

    def make_manager(self, areas, masses, failing_face=None):
        return FakeMeasureManager(areas, masses, failing_face)

    def test_dry_run_leaf_computes_area_and_rollup_mass_without_writing(self):
        body, face = body_with_area("BODY1", 2.5)
        part = FakePart("264MN000001A01", [body])
        part.attributes = {"DB_PART_NO": "264MN000001A01", "DB_PART_REV": "A"}
        manager = self.make_manager(
            {face: 2.5},
            {(body.Tag,): 1.25},
        )
        session = FakeSession(part, manager, FakeUnits())

        path, rows, diagnostics = self.j21.run(session)

        self.assertEqual(1, len(rows))
        row = rows[0]
        self.assertEqual("264MN000001A01", row["DB_PART_NO"])
        self.assertEqual("2.5000", row["SURFACE_AREA_M2"])
        self.assertEqual("1.250000", row["ROLLUP_MASS_KG"])
        self.assertEqual("DRY_RUN", row["SURFACE_AREA_ATTRIBUTE"])
        self.assertEqual("DRY_RUN", row["ROLLUP_MASS_ATTRIBUTE"])
        self.assertEqual("DRY_RUN", row["SAVED"])
        self.assertEqual("SUCCESS", row["STATUS"])
        self.assertFalse(diagnostics)
        self.assertEqual([], session.AttributeManager.writes)
        self.assertFalse(part.saved)
        self.assertTrue(os.path.isfile(path))

    def test_dry_run_subassembly_rolls_up_child_masses(self):
        root_body, root_face = body_with_area("ROOT_BODY", 1.0)
        child_body, child_face = body_with_area("CHILD_BODY", 3.0)
        child = FakePart("264MN000002A01", [child_body])
        child.attributes = {"DB_PART_NO": "264MN000002A01", "DB_PART_REV": "A"}
        root = FakePart(
            "264MN000001A01",
            [root_body],
            component_children=[FakeComponent("CHILD-1", child)],
        )
        root.attributes = {"DB_PART_NO": "264MN000001A01", "DB_PART_REV": "A"}
        manager = self.make_manager(
            {root_face: 1.0, child_face: 3.0},
            {
                (root_body.Tag,): 1.0,
                (child_body.Tag,): 2.0,
                (root_body.Tag, child_body.Tag): 3.0,
            },
        )
        session = FakeSession(root, manager, FakeUnits())

        _path, rows, _diagnostics = self.j21.run(session)

        by_number = rows_by_part_number(rows)
        root_row = by_number["264MN000001A01"]
        child_row = by_number["264MN000002A01"]
        self.assertEqual("3.000000", root_row["ROLLUP_MASS_KG"])
        self.assertEqual(2, root_row["ROLLUP_SOLID_BODY_COUNT"])
        self.assertEqual("2.000000", child_row["ROLLUP_MASS_KG"])
        self.assertEqual(1, child_row["ROLLUP_SOLID_BODY_COUNT"])
        self.assertEqual("3.0000", child_row["SURFACE_AREA_M2"])
        self.assertEqual(0, root_row["LEVEL"])
        self.assertEqual(1, child_row["LEVEL"])

    def test_bom_visibility_filters_noise_from_rows_and_rollup(self):
        root_body, root_face = body_with_area("ROOT_BODY", 1.0)
        normal_body, normal_face = body_with_area("NORMAL_BODY", 2.0)
        suppressed_body = FakeBody("SUPPRESSED_BODY", [FakeFace()])
        reference_body = FakeBody("REFERENCE_BODY", [FakeFace()])
        csys_body = FakeBody("CSYS_BODY", [FakeFace()])

        normal = FakePart("264MN000002A01", [normal_body])
        normal.attributes = {"DB_PART_NO": "264MN000002A01", "DB_PART_REV": "A"}
        suppressed = FakePart("264MN000003A01", [suppressed_body])
        suppressed.attributes = {"DB_PART_NO": "264MN000003A01", "DB_PART_REV": "A"}
        reference = FakePart("264MN000004A01", [reference_body])
        reference.attributes = {"DB_PART_NO": "264MN000004A01", "DB_PART_REV": "A"}
        csys = FakePart("264MN000005A01", [csys_body])
        csys.attributes = {"DB_PART_NO": "264MN000005A01", "DB_PART_REV": "A"}

        root = FakePart(
            "264MN000001A01",
            [root_body],
            component_children=[
                FakeComponent("NORMAL-1", normal),
                FakeComponent("SUPPRESSED-1", suppressed, suppressed=True),
                FakeComponent(
                    "REFERENCE-1",
                    reference,
                    string_attributes={"REFERENCE_COMPONENT": ""},
                ),
                FakeComponent("CSYS_ORIGIN", csys),
            ],
        )
        root.attributes = {"DB_PART_NO": "264MN000001A01", "DB_PART_REV": "A"}

        manager = self.make_manager(
            {root_face: 1.0, normal_face: 2.0},
            {
                (root_body.Tag,): 1.0,
                (normal_body.Tag,): 2.0,
                (suppressed_body.Tag,): 100.0,
                (reference_body.Tag,): 100.0,
                (csys_body.Tag,): 100.0,
                (root_body.Tag, normal_body.Tag): 3.0,
            },
        )
        session = FakeSession(root, manager, FakeUnits())

        _path, rows, _diagnostics = self.j21.run(session)

        numbers = [row["DB_PART_NO"] for row in rows]
        self.assertEqual(
            ["264MN000001A01", "264MN000002A01"],
            numbers,
        )
        root_row = rows_by_part_number(rows)["264MN000001A01"]
        self.assertEqual("3.000000", root_row["ROLLUP_MASS_KG"])
        self.assertEqual(2, root_row["ROLLUP_SOLID_BODY_COUNT"])
        # The noise bodies must never be handed to the mass measurement.
        measured_tags = {
            body.Tag
            for call in manager.mass_calls
            for body in call[2]
        }
        self.assertNotIn(suppressed_body.Tag, measured_tags)
        self.assertNotIn(reference_body.Tag, measured_tags)
        self.assertNotIn(csys_body.Tag, measured_tags)

    def test_apply_writes_attributes_and_saves_each_part(self):
        body, face = body_with_area("BODY1", 4.0)
        part = FakePart("264MN000001A01", [body])
        part.attributes = {"DB_PART_NO": "264MN000001A01", "DB_PART_REV": "A"}
        manager = self.make_manager({face: 4.0}, {(body.Tag,): 2.0})
        session = FakeSession(part, manager, FakeUnits())

        os.environ.pop("NX_J21_MODE", None)  # exercise the APPLY default
        _path, rows, _diagnostics = self.j21.run(session)

        row = rows[0]
        self.assertEqual("WRITTEN", row["SURFACE_AREA_ATTRIBUTE"])
        self.assertEqual("WRITTEN", row["ROLLUP_MASS_ATTRIBUTE"])
        self.assertEqual("SAVED", row["SAVED"])
        self.assertEqual("SUCCESS", row["STATUS"])
        self.assertTrue(part.saved)
        self.assertEqual(
            [
                (
                    part,
                    "Materials",
                    "NX_SURFACE_AREA",
                    "NUMBER",
                    4.0,
                ),
                (
                    part,
                    "Materials",
                    "NX_MassPropRollupMass",
                    "NUMBER",
                    2.0,
                ),
            ],
            session.AttributeManager.writes,
        )

    def test_apply_save_failure_is_reported_and_run_continues(self):
        ok_body, ok_face = body_with_area("OK_BODY", 1.0)
        bad_body, bad_face = body_with_area("BAD_BODY", 2.0)
        ok = FakePart("264MN000002A01", [ok_body])
        ok.attributes = {"DB_PART_NO": "264MN000002A01", "DB_PART_REV": "A"}
        bad = FakePart("264MN000003A01", [bad_body], save_error=True)
        bad.attributes = {"DB_PART_NO": "264MN000003A01", "DB_PART_REV": "A"}
        root = FakePart(
            "264MN000001A01",
            [],
            component_children=[
                FakeComponent("OK-1", ok),
                FakeComponent("BAD-1", bad),
            ],
        )
        root.attributes = {"DB_PART_NO": "264MN000001A01", "DB_PART_REV": "A"}
        manager = self.make_manager(
            {ok_face: 1.0, bad_face: 2.0},
            {
                (ok_body.Tag,): 1.0,
                (bad_body.Tag,): 2.0,
                (ok_body.Tag, bad_body.Tag): 3.0,
            },
        )
        session = FakeSession(root, manager, FakeUnits())

        os.environ.pop("NX_J21_MODE", None)
        _path, rows, _diagnostics = self.j21.run(session)

        by_number = rows_by_part_number(rows)
        self.assertEqual("SAVED", by_number["264MN000002A01"]["SAVED"])
        self.assertEqual(
            "SAVE_FAILED", by_number["264MN000003A01"]["SAVED"]
        )
        self.assertEqual(
            "SAVE_FAILED", by_number["264MN000003A01"]["STATUS"]
        )
        self.assertIn("SAVE:", by_number["264MN000003A01"]["MESSAGE"])

    def test_apply_attribute_commit_failure_is_reported(self):
        body, face = body_with_area("BODY1", 4.0)
        part = FakePart("264MN000001A01", [body])
        part.attributes = {"DB_PART_NO": "264MN000001A01", "DB_PART_REV": "A"}
        manager = self.make_manager({face: 4.0}, {(body.Tag,): 2.0})
        session = FakeSession(part, manager, FakeUnits(), fail_commit=True)

        os.environ.pop("NX_J21_MODE", None)
        _path, rows, _diagnostics = self.j21.run(session)

        row = rows[0]
        self.assertEqual("WRITE_FAILED", row["SURFACE_AREA_ATTRIBUTE"])
        self.assertEqual("WRITE_FAILED", row["ROLLUP_MASS_ATTRIBUTE"])
        self.assertEqual("PARTIAL", row["STATUS"])
        self.assertIn("AREA ATTRIBUTE:", row["MESSAGE"])

    def test_no_work_part_raises(self):
        session = FakeSession(None, None, None)
        with self.assertRaises(RuntimeError):
            self.j21.run(session)

    def test_measurement_failure_is_fail_closed(self):
        good_body, good_face = body_with_area("GOOD_BODY", 1.0)
        failing_body, failing_face = body_with_area("FAILING_BODY", 2.0)
        part = FakePart("264MN000001A01", [good_body, failing_body])
        part.attributes = {"DB_PART_NO": "264MN000001A01", "DB_PART_REV": "A"}
        manager = self.make_manager(
            {good_face: 1.0, failing_face: 2.0},
            {
                (good_body.Tag, failing_body.Tag): 3.0,
                (good_body.Tag,): 1.0,
                (failing_body.Tag,): 2.0,
            },
            failing_face=failing_face,
        )
        session = FakeSession(part, manager, FakeUnits())

        _path, rows, _diagnostics = self.j21.run(session)

        row = rows[0]
        self.assertEqual("", row["SURFACE_AREA_M2"])
        self.assertEqual("FAILED", row["SURFACE_AREA_ATTRIBUTE"])
        self.assertEqual("PARTIAL", row["STATUS"])
        self.assertIn("AREA:", row["MESSAGE"])
        # Mass is independent of the failed face measurement.
        self.assertEqual("3.000000", row["ROLLUP_MASS_KG"])

    def test_static_self_contained_no_checkout_dependency(self):
        source = JOURNAL.read_text(encoding="utf-8")
        self.assertNotIn("Checkout", source)
        self.assertNotIn("import utils", source)
        self.assertIn("NXOpenBoMExtended", source)


if __name__ == "__main__":
    unittest.main()
