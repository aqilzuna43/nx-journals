import importlib.util
import json
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
    / "22_diagnose_mass_attribute_write.py"
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
        spec = importlib.util.spec_from_file_location("journal22", JOURNAL)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        return module
    finally:
        if prior is None:
            sys.modules.pop("NXOpen", None)
        else:
            sys.modules["NXOpen"] = prior


class FakeUnit:
    def __init__(self, name, symbol):
        self.Name = name
        self.Symbol = symbol
        self.Abbreviation = symbol
        self.TypeName = name


class FakeUnits:
    def __init__(self):
        self.area = FakeUnit("SquareMeter", "m^2")
        self.length = FakeUnit("Meter", "m")
        self.mass = FakeUnit("Kilogram", "kg")

    def FindObject(self, name):
        if name in ("SquareMeter", "SquareMetre"):
            return self.area
        if name in ("Meter", "Metre"):
            return self.length
        if name == "Kilogram":
            return self.mass
        raise RuntimeError("not found: " + name)


class FakeAreaMeasurement:
    def __init__(self, area):
        self.Area = area

    def Dispose(self):
        pass


class FakeMassMeasurement:
    def __init__(self, mass, volume=None, area=None):
        self.Mass = mass
        self.Volume = volume
        self.Area = area

    def Dispose(self):
        pass


class FakeAttributeType:
    def __init__(self, name):
        self.name = name


class FakeAttributeInfo:
    def __init__(self, category, title, value, kind="String", unset=False):
        self.Category = category
        self.Title = title
        self.Type = FakeAttributeType(kind)
        self.StringValue = value if kind == "String" else ""
        self.RealValue = value if kind == "Real" else None
        self.IntegerValue = None
        self.BooleanValue = None
        self.Unset = unset


class FakeIterator:
    def SetIncludeOnlyCategory(self, value):
        self.category = value

    def SetIncludeOnlyTitle(self, value):
        self.title = value

    def SetIncludeAlsoUnset(self, value):
        self.include_unset = value

    def FreeResource(self):
        pass


class FakeMeasureManager:
    def __init__(self, mass=1.25, area_m2=2.5, fail_mass=False, fail_area=False):
        self.mass = mass
        self.area_m2 = area_m2
        self.fail_mass = fail_mass
        self.fail_area = fail_area
        self.mass_calls = 0
        self.face_calls = 0

    def NewMassProperties(self, mass_units, accuracy, bodies):
        self.mass_calls += 1
        if self.fail_mass:
            raise RuntimeError("mass compute failed")
        return FakeMassMeasurement(
            self.mass, volume=100.0, area=self.area_m2 * 1_000_000.0
        )

    def NewFaceProperties(self, area_unit, length_unit, accuracy, faces):
        self.face_calls += 1
        if self.fail_area:
            raise RuntimeError("face compute failed")
        return FakeAreaMeasurement(self.area_m2)


class FakeMassPropsBuilder:
    def __init__(self):
        self.Accuracy = None
        self.UpdateOnSave = "NO"
        self.UpdateOptions = types.SimpleNamespace(Yes="YES")
        self.UpdateNow_called = False
        self.Commit_called = False

    def UpdateNow(self):
        self.UpdateNow_called = True

    def Commit(self):
        self.Commit_called = True


class FakePropertiesManager:
    def __init__(self):
        self.builder = FakeMassPropsBuilder()

    def CreateMassPropertiesBuilder(self, objects):
        return self.builder


class FakeAttributeBuilder:
    def __init__(self, manager, target, objects, operation):
        self.manager = manager
        self.target = target
        self.Category = ""
        self.Title = ""
        self.DataType = ""
        self.NumberValue = None

    def Commit(self):
        if self.Category in self.manager.fail_categories:
            raise RuntimeError("commit failed")
        self.manager.writes.append(
            (self.target, self.Category, self.Title, self.NumberValue)
        )
        self.target.real_attributes[self.Title] = self.NumberValue
        # Simulate NX adding the attribute to the part's attribute list.
        self.target.attribute_infos.append(
            FakeAttributeInfo(
                self.Category, self.Title, self.NumberValue, kind="Real"
            )
        )


class FakeAttributeManager:
    def __init__(self, fail_categories=()):
        self.writes = []
        self.fail_categories = tuple(fail_categories)

    def CreateAttributePropertiesBuilder(self, target, objects, operation):
        return FakeAttributeBuilder(self, target, objects, operation)


class FakeBody:
    _tag = 0

    def __init__(self, name, faces=()):
        FakeBody._tag += 1
        self.Name = name
        self.Tag = FakeBody._tag
        self.IsSolidBody = True
        self.IsSheetBody = False
        self.IsConvergentBody = False
        self._faces = list(faces)

    def GetFaces(self):
        return list(self._faces)


class FakeFace:
    _tag = 500

    def __init__(self):
        FakeFace._tag += 1
        self.Tag = FakeFace._tag


class FakePart:
    _tag = 0

    def __init__(
        self,
        name,
        bodies=(),
        attribute_infos=(),
        with_properties_manager=True,
        save_error=False,
    ):
        FakePart._tag += 1
        self.Name = name
        self.Leaf = name
        self.Tag = FakePart._tag
        self.Bodies = list(bodies)
        self.MeasureManager = None
        self.UnitCollection = FakeUnits()
        self.PropertiesManager = (
            FakePropertiesManager() if with_properties_manager else None
        )
        self.ComponentAssembly = types.SimpleNamespace(RootComponent=None)
        self.attribute_infos = list(attribute_infos)
        self.string_attributes = {}
        self.real_attributes = {}
        self.save_error = save_error

    def GetStringAttribute(self, title):
        if title not in self.string_attributes:
            raise AttributeError("No such attribute: " + title)
        return self.string_attributes[title]

    def GetUserAttribute(self, *args):
        raise AttributeError("unavailable")

    def CreateAttributeIterator(self):
        return FakeIterator()

    def GetUserAttributes(self, iterator=None):
        return list(self.attribute_infos)

    def GetRealAttribute(self, title):
        if title not in self.real_attributes:
            raise KeyError("No such attribute: " + title)
        return self.real_attributes[title]

    def Save(self, *args):
        if self.save_error:
            raise RuntimeError("save failed")
        return types.SimpleNamespace(
            NumberUnsavedParts=0, NumberUnsavedObjects=0
        )


class FakeListingWindow:
    def __init__(self):
        self.lines = []

    def Open(self):
        pass

    def WriteFullline(self, line):
        self.lines.append(line)


class FakeSession:
    def __init__(self, work_part, fail_categories=()):
        self.Parts = types.SimpleNamespace(Work=work_part)
        self.ListingWindow = FakeListingWindow()
        self.AttributeManager = FakeAttributeManager(fail_categories)
        if work_part is not None:
            work_part.MeasureManager = FakeMeasureManager()


class J22Tests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.j22 = load_journal()

    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp_dir.cleanup)
        os.environ["NX_JOURNALS_IO_DIR"] = self.temp_dir.name

    def make_part(self, **kwargs):
        body = FakeBody("BODY1", [FakeFace()])
        part = FakePart(
            "264MN000001A01",
            [body],
            attribute_infos=[
                FakeAttributeInfo(
                    "Rolled-Up Mass Properties",
                    "NX_MassPropRollupMass",
                    0.1,
                    kind="Real",
                )
            ],
            **kwargs,
        )
        part.string_attributes = {
            "DB_PART_NO": "264MN000001A01",
            "DB_PART_REV": "A",
        }
        return part

    def test_run_writes_csv_json_and_reports_all_tests(self):
        part = self.make_part()
        session = FakeSession(part)

        csv_path, json_path, report = self.j22.run(session)

        self.assertTrue(os.path.isfile(csv_path))
        self.assertTrue(os.path.isfile(json_path))
        with open(json_path, encoding="utf-8") as handle:
            payload = json.load(handle)
        self.assertEqual("264MN000001A01", payload["work_part"]["number"])
        self.assertEqual(1, payload["solid_body_count"])
        statuses = {
            (item.get("test"), item.get("category", "")): item["status"]
            for item in payload["findings"]
        }
        self.assertEqual("OK", statuses[("A_classic_compute", "")])
        self.assertEqual("OK", statuses[("B_native_builder", "")])
        self.assertEqual(
            "OK", statuses[("C_direct_write", "Rolled-Up Mass Properties")]
        )
        self.assertEqual("OK", statuses[("C_direct_write", "Materials")])
        self.assertEqual("OK", statuses[("save", "")])
        self.assertTrue(
            any(
                item["title"] == "NX_MassPropRollupArea"
                for item in payload["after_attributes"]
            )
        )

    def test_classic_compute_failure_reported(self):
        part = self.make_part()
        session = FakeSession(part)
        part.MeasureManager.fail_mass = True

        _csv_path, _json_path, report = self.j22.run(session)

        statuses = {
            item.get("test"): item["status"] for item in report["findings"]
        }
        self.assertEqual(
            "FAILED", statuses["A_classic_compute_NewMassProperties"]
        )

    def test_native_builder_absent_reported(self):
        part = self.make_part(with_properties_manager=False)
        session = FakeSession(part)

        _csv_path, _json_path, report = self.j22.run(session)

        statuses = {
            item.get("test"): item["status"] for item in report["findings"]
        }
        self.assertEqual("NO_PROPERTIES_MANAGER", statuses["B_native_builder"])

    def test_direct_write_category_failure_reported(self):
        part = self.make_part()
        session = FakeSession(
            part, fail_categories=("Rolled-Up Mass Properties",)
        )

        _csv_path, _json_path, report = self.j22.run(session)

        statuses = {
            (item.get("test"), item.get("category", "")): item["status"]
            for item in report["findings"]
        }
        self.assertEqual(
            "FAILED", statuses[("C_direct_write", "Rolled-Up Mass Properties")]
        )
        self.assertEqual("OK", statuses[("C_direct_write", "Materials")])

    def test_no_work_part_raises(self):
        session = FakeSession(None)
        with self.assertRaises(RuntimeError):
            self.j22.run(session)


if __name__ == "__main__":
    unittest.main()
