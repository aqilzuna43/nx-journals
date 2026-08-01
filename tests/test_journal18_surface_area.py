import csv
import datetime
import importlib.util
import inspect
import sys
import tempfile
import types
import unittest
from pathlib import Path
from unittest import mock


ROOT = Path(__file__).resolve().parents[1]
JOURNAL = (
    ROOT / "from_git" / "journals" / "18_work_part_surface_area.py"
)


def load_journal():
    nxopen = types.ModuleType("NXOpen")
    nxopen.NXObject = types.SimpleNamespace(
        AttributeType=types.SimpleNamespace(String="String")
    )
    nxopen.Session = types.SimpleNamespace()
    sys.modules["NXOpen"] = nxopen
    spec = importlib.util.spec_from_file_location("journal18", JOURNAL)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class FakeUnit:
    def __init__(self, name, symbol, measure):
        self.Name = name
        self.Symbol = symbol
        self.Abbreviation = symbol
        self.TypeName = name
        self.Measure = measure


class FakeUnits:
    def __init__(self, expose_find=True):
        self.area = FakeUnit("SquareMeter", "m^2", "Area")
        self.length = FakeUnit("Meter", "m", "Length")
        self.expose_find = expose_find

    def FindObject(self, name):
        if not self.expose_find:
            raise RuntimeError("FindObject unavailable")
        if name in ("SquareMeter", "SquareMetre"):
            return self.area
        if name in ("Meter", "Metre"):
            return self.length
        raise RuntimeError("not found")

    def GetMeasureTypes(self, measure):
        if measure == "Area":
            return [self.area]
        if measure == "Length":
            return [self.length]
        return []


class FakeMeasurement:
    def __init__(self, area):
        self.Area = area
        self.disposed = False

    def Dispose(self):
        self.disposed = True


class FakeMeasureManager:
    def __init__(self, areas, failing_face=None):
        self.areas = areas
        self.failing_face = failing_face
        self.calls = []
        self.measurements = []

    def NewFaceProperties(
        self,
        area_unit,
        length_unit,
        accuracy,
        faces,
    ):
        self.calls.append(
            (area_unit, length_unit, accuracy, list(faces))
        )
        key = faces[0]
        if key == self.failing_face:
            raise RuntimeError("NX measurement failed")
        measurement = FakeMeasurement(self.areas[key])
        self.measurements.append(measurement)
        return measurement


class FakeBody:
    def __init__(
        self,
        name,
        tag,
        faces,
        solid=True,
        sheet=False,
        convergent=False,
        hidden=False,
    ):
        self.Name = name
        self.Tag = tag
        self.IsSolidBody = solid
        self.IsSheetBody = sheet
        self.IsConvergentBody = convergent
        self.IsBlanked = hidden
        self._faces = list(faces)

    def GetFaces(self):
        return list(self._faces)


class FakePart:
    def __init__(self, bodies, manager, units=None):
        self.Bodies = list(bodies)
        self.MeasureManager = manager
        self.UnitCollection = units or FakeUnits()
        self.Name = "TEST_PART"
        self.Leaf = "TEST_PART"
        self.attributes = {
            "DB_PART_NO": "264MN000001A01",
            "DB_PART_REV": "A",
        }

    def GetStringAttribute(self, name):
        return self.attributes.get(name, "")


class SurfaceAreaTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.journal = load_journal()

    def test_multibody_surface_area_includes_hidden_solids(self):
        body1 = FakeBody("BODY_A", 101, ["face-a"], hidden=True)
        body2 = FakeBody("BODY_B", 102, ["face-b", "face-c"])
        sheet = FakeBody(
            "SHEET",
            103,
            ["face-sheet"],
            solid=False,
            sheet=True,
        )
        convergent = FakeBody(
            "CONVERGENT",
            104,
            ["face-convergent"],
            solid=True,
            convergent=True,
        )
        manager = FakeMeasureManager(
            {
                "face-a": 1.23444,
                "face-b": 2.34556,
            }
        )
        part = FakePart(
            [body1, body2, sheet, convergent],
            manager,
        )

        rows = self.journal.calculate_surface_rows(
            part,
            "2026-07-31T00:13:00+08:00",
        )

        self.assertEqual(len(rows), 3)
        self.assertEqual(
            [row["ROW_TYPE"] for row in rows],
            ["BODY", "BODY", "TOTAL"],
        )
        self.assertEqual(rows[0]["SURFACE_AREA_M2"], "1.2344")
        self.assertEqual(rows[1]["SURFACE_AREA_M2"], "2.3456")
        self.assertEqual(rows[2]["SURFACE_AREA_M2"], "3.5800")
        self.assertEqual(rows[2]["INCLUDED_SOLID_BODY_COUNT"], 2)
        self.assertEqual(rows[2]["SKIPPED_SHEET_BODY_COUNT"], 1)
        self.assertEqual(
            rows[2]["SKIPPED_CONVERGENT_BODY_COUNT"],
            1,
        )
        self.assertEqual(len(manager.calls), 2)
        self.assertTrue(
            all(
                call[2] == self.journal.MEASUREMENT_ACCURACY
                for call in manager.calls
            )
        )
        self.assertTrue(
            all(item.disposed for item in manager.measurements)
        )

    def test_one_body_failure_blanks_fail_closed_total(self):
        body1 = FakeBody("GOOD", 1, ["good"])
        body2 = FakeBody("BAD", 2, ["bad"])
        manager = FakeMeasureManager(
            {"good": 1.0, "bad": 2.0},
            failing_face="bad",
        )
        part = FakePart([body1, body2], manager)

        rows = self.journal.calculate_surface_rows(part, "timestamp")

        self.assertEqual(rows[0]["STATUS"], "SUCCESS")
        self.assertEqual(rows[1]["STATUS"], "FAILED_MEASUREMENT")
        self.assertEqual(rows[2]["STATUS"], "FAILED_BODY_MEASUREMENT")
        self.assertEqual(rows[2]["SURFACE_AREA_M2"], "")
        self.assertIn("fail-closed total is blank", rows[2]["MESSAGE"])

    def test_no_traditional_solids_produces_failure_total(self):
        sheet = FakeBody(
            "SHEET",
            1,
            ["sheet"],
            solid=False,
            sheet=True,
        )
        manager = FakeMeasureManager({})
        part = FakePart([sheet], manager)

        rows = self.journal.calculate_surface_rows(part, "timestamp")

        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["ROW_TYPE"], "TOTAL")
        self.assertEqual(rows[0]["STATUS"], "FAILED_NO_SOLID_BODIES")
        self.assertEqual(rows[0]["SURFACE_AREA_M2"], "")
        self.assertEqual(rows[0]["SKIPPED_SHEET_BODY_COUNT"], 1)
        self.assertEqual(manager.calls, [])

    def test_unit_resolution_falls_back_to_measure_types(self):
        units = FakeUnits(expose_find=False)
        part = types.SimpleNamespace(UnitCollection=units)

        area, length = self.journal.resolve_square_metre_units(part)

        self.assertIs(area, units.area)
        self.assertIs(length, units.length)

    def test_missing_square_metre_unit_fails_with_body_count(self):
        body = FakeBody("BODY", 1, ["face"])
        missing_units = types.SimpleNamespace(
            FindObject=mock.Mock(side_effect=RuntimeError("not found")),
            GetMeasureTypes=mock.Mock(return_value=[]),
        )
        part = FakePart(
            [body],
            FakeMeasureManager({"face": 1.0}),
            units=missing_units,
        )

        rows = self.journal.calculate_surface_rows(part, "timestamp")

        self.assertEqual(rows[-1]["STATUS"], "FAILED_UNIT_RESOLUTION")
        self.assertEqual(rows[-1]["INCLUDED_SOLID_BODY_COUNT"], 1)
        self.assertEqual(rows[-1]["SURFACE_AREA_M2"], "")

    def test_run_writes_utf8_bom_body_and_total_csv(self):
        body = FakeBody("BODY", 42, ["face"])
        part = FakePart(
            [body],
            FakeMeasureManager({"face": 0.12345}),
        )
        session = types.SimpleNamespace(
            Parts=types.SimpleNamespace(Work=part)
        )
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        run_datetime = datetime.datetime(
            2026,
            7,
            31,
            0,
            13,
            tzinfo=datetime.timezone(datetime.timedelta(hours=8)),
        )

        with mock.patch.object(
            self.journal,
            "io_root",
            return_value=folder.name,
        ):
            path, rows = self.journal.run(session, run_datetime)

        data = Path(path).read_bytes()
        self.assertTrue(data.startswith(b"\xef\xbb\xbf"))
        self.assertEqual(Path(path).parent.name, "NX_SURFACE_AREA")
        self.assertIn("264MN000001A01", Path(path).name)
        with open(path, "r", encoding="utf-8-sig", newline="") as handle:
            written = list(csv.DictReader(handle))
        self.assertEqual(len(written), 2)
        self.assertEqual(written[-1]["ROW_TYPE"], "TOTAL")
        self.assertEqual(written[-1]["SURFACE_AREA_M2"], "0.1235")
        self.assertEqual(rows[-1]["STATUS"], "SUCCESS")

    def test_no_work_part_still_writes_failure_csv(self):
        session = types.SimpleNamespace(
            Parts=types.SimpleNamespace(Work=None)
        )
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)

        with mock.patch.object(
            self.journal,
            "io_root",
            return_value=folder.name,
        ):
            path, rows = self.journal.run(session)

        self.assertTrue(Path(path).is_file())
        self.assertEqual(rows[-1]["STATUS"], "FAILED_NO_WORK_PART")
        self.assertEqual(rows[-1]["SURFACE_AREA_M2"], "")

    def test_source_is_surface_only_and_read_only(self):
        source = JOURNAL.read_text(encoding="utf-8")

        self.assertIn("NewFaceProperties", source)
        self.assertNotIn("NewMassProperties", source)
        self.assertNotIn("PdmSession", source)
        self.assertNotIn("ComponentAssembly", source)
        self.assertNotIn(".Save(", source)
        self.assertNotIn("SetUserAttribute", source)
        self.assertNotIn("CreateFeature(", source)
        self.assertNotIn("CreateEmbedded", source)
        self.assertNotIn("GetChildren(", source)

    def test_measurement_uses_faces_not_body_mass_properties(self):
        source = inspect.getsource(self.journal.measure_body_area_m2)

        self.assertIn("body.GetFaces()", source)
        self.assertIn("NewFaceProperties", source)
        self.assertEqual(
            self.journal.BUILD,
            "J18-NX2506-WORK-PART-SURFACE-AREA-V1",
        )


if __name__ == "__main__":
    unittest.main()
