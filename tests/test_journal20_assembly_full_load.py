import importlib.util
import os
import sys
import tempfile
import types
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
JOURNAL = ROOT / "from_git" / "journals" / "20_diagnose_assembly_full_load.py"


def load_journal():
    nxopen = types.ModuleType("NXOpen")
    nxopen.NXObject = types.SimpleNamespace(
        AttributeType=types.SimpleNamespace(String="String")
    )
    nxopen.Session = types.SimpleNamespace(
        LibraryUnloadOption=types.SimpleNamespace(Immediately="Immediately")
    )
    sys.modules["NXOpen"] = nxopen
    spec = importlib.util.spec_from_file_location("journal20", JOURNAL)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class State:
    def __init__(self, name):
        self.name = name

    def __str__(self):
        return self.name


class FakeLoadStatus:
    def __init__(self, failures=None):
        self.failures = list(failures or [])
        self.disposed = False

    @property
    def NumberUnloadedParts(self):
        return len(self.failures)

    def GetPartName(self, index):
        return self.failures[index][0]

    def GetStatus(self, index):
        return self.failures[index][1]

    def GetStatusDescription(self, index):
        return self.failures[index][2]

    def Dispose(self):
        self.disposed = True


class FakePrototype:
    def __init__(self, name, tag, behavior="success", load_status=None):
        self.Name = name
        self.Leaf = name
        self.FullPath = "@DB/{0}/A".format(name)
        self.Tag = tag
        self.PartLoadState = State("MinimallyLoaded")
        self.IsFullyLoaded = False
        self.behavior = behavior
        self.load_status = load_status or FakeLoadStatus()
        self.load_calls = 0

    def GetStringAttribute(self, name):
        if name == "DB_PART_NO":
            return self.Name
        if name == "DB_PART_REV":
            return "A"
        raise RuntimeError("attribute missing")

    def GetUserAttribute(self, *args):
        raise RuntimeError("attribute missing")

    def LoadThisPartFully(self):
        self.load_calls += 1
        if self.behavior == "invalid":
            raise RuntimeError(
                "IM0541: An operation was attempted on an invalid or unsuitable OM object"
            )
        if not self.load_status.NumberUnloadedParts:
            self.IsFullyLoaded = True
            self.PartLoadState = State("FullyLoaded")
        return self.load_status


class FakeComponent:
    def __init__(self, name, prototype=None, children=None, suppressed=False):
        self.DisplayName = name
        self.Name = name
        self.ReferenceSet = "MODEL"
        self.IsSuppressed = suppressed
        self._prototype = prototype
        self._children = list(children or [])

    @property
    def Prototype(self):
        return self._prototype

    def GetChildren(self):
        return list(self._children)


class FakeWorkPart:
    def __init__(self, children, assembly_behavior="success"):
        self.Name = "TOP"
        root = FakeComponent("ROOT", children=children)
        self.ComponentAssembly = types.SimpleNamespace(RootComponent=root)
        self.assembly_behavior = assembly_behavior
        self.load_calls = 0
        self.load_status = FakeLoadStatus()

    def LoadFully(self):
        self.load_calls += 1
        if self.assembly_behavior == "invalid":
            raise RuntimeError("invalid or unsuitable OM object")
        return self.load_status


class Journal20Tests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.journal = load_journal()

    def test_traversal_captures_exact_nested_occurrence_path(self):
        leaf_proto = FakePrototype("LEAF_PART", 2)
        leaf = FakeComponent("LEAF_OCC", leaf_proto)
        sub_proto = FakePrototype("SUB_PART", 1)
        sub = FakeComponent("SUB_OCC", sub_proto, children=[leaf])
        work = FakeWorkPart([sub])

        rows, errors, is_assembly = self.journal.collect_occurrences(
            work, "2026-08-11T20:00:00+08:00"
        )

        self.assertTrue(is_assembly)
        self.assertEqual([], errors)
        self.assertEqual(2, len(rows))
        self.assertEqual("TOP / SUB_OCC", rows[0]["ASSEMBLY_PATH"])
        self.assertEqual("TOP / SUB_OCC / LEAF_OCC", rows[1]["ASSEMBLY_PATH"])
        self.assertEqual(2, rows[1]["LEVEL"])

    def test_shared_prototype_is_loaded_once_and_result_maps_to_all_occurrences(self):
        prototype = FakePrototype("SHARED", 10)
        work = FakeWorkPart(
            [
                FakeComponent("SHARED_1", prototype),
                FakeComponent("SHARED_2", prototype),
            ]
        )
        rows, _, _ = self.journal.collect_occurrences(work, "timestamp")

        unique_count = self.journal.probe_all_prototypes(rows)

        self.assertEqual(1, unique_count)
        self.assertEqual(1, prototype.load_calls)
        self.assertEqual(["SUCCESS", "SUCCESS"], [row["FULL_LOAD_PROBE"] for row in rows])
        self.assertEqual(["OK", "OK"], [row["STATUS"] for row in rows])
        self.assertTrue(all(row["FINAL_LOAD_STATE"] == "FullyLoaded" for row in rows))

    def test_invalid_om_exception_identifies_prototype_and_every_occurrence(self):
        prototype = FakePrototype("BAD_PART", 20, behavior="invalid")
        work = FakeWorkPart(
            [FakeComponent("BAD_1", prototype), FakeComponent("BAD_2", prototype)]
        )
        rows, _, _ = self.journal.collect_occurrences(work, "timestamp")

        self.journal.probe_all_prototypes(rows)

        self.assertEqual(1, prototype.load_calls)
        self.assertTrue(all(row["STATUS"] == "INVALID_OBJECT" for row in rows))
        self.assertTrue(all(row["FULL_LOAD_PROBE"] == "FAILED" for row in rows))
        self.assertTrue(
            all(row["FAILED_OPERATION"] == "BasePart.LoadThisPartFully" for row in rows)
        )
        self.assertIn("IM0541", rows[0]["EXCEPTION"])

    def test_part_load_status_missing_file_is_classified_and_disposed(self):
        status = FakeLoadStatus(
            [("MISSING.prt", 641044, "Failed to find file using current search options")]
        )
        prototype = FakePrototype("MISSING", 30, load_status=status)
        work = FakeWorkPart([FakeComponent("MISSING_OCC", prototype)])
        rows, _, _ = self.journal.collect_occurrences(work, "timestamp")

        self.journal.probe_all_prototypes(rows)

        self.assertEqual("MISSING_FILE", rows[0]["STATUS"])
        self.assertIn("MISSING.prt", rows[0]["LOAD_STATUS_DETAILS"])
        self.assertTrue(status.disposed)

    def test_unresolved_occurrences_are_not_incorrectly_grouped(self):
        work = FakeWorkPart(
            [FakeComponent("MISSING_A"), FakeComponent("MISSING_B")]
        )
        rows, _, _ = self.journal.collect_occurrences(work, "timestamp")

        unique_count = self.journal.probe_all_prototypes(rows)

        self.assertEqual(2, unique_count)
        self.assertNotEqual(rows[0]["_prototype_key"], rows[1]["_prototype_key"])

    def test_final_assembly_load_reproduces_invalid_object_failure(self):
        work = FakeWorkPart([], assembly_behavior="invalid")

        result = self.journal.probe_assembly_full_load(work)

        self.assertEqual("INVALID_OBJECT", result["status"])
        self.assertEqual(1, work.load_calls)
        self.assertIn("invalid or unsuitable OM object", result["exception"])

    def test_reports_include_component_and_assembly_results(self):
        prototype = FakePrototype("GOOD", 40)
        work = FakeWorkPart([FakeComponent("GOOD_OCC", prototype)])
        rows, errors, _ = self.journal.collect_occurrences(work, "timestamp")
        self.journal.probe_all_prototypes(rows)
        assembly_result = self.journal.probe_assembly_full_load(work)

        with tempfile.TemporaryDirectory() as folder:
            csv_path = os.path.join(folder, "report.csv")
            text_path = os.path.join(folder, "report.txt")
            output_rows = rows + [
                self.journal.assembly_summary_row("TOP", "timestamp", assembly_result)
            ]
            self.journal.write_csv(csv_path, output_rows)
            self.journal.write_text_report(
                text_path, "TOP", "timestamp", rows, errors, assembly_result
            )

            csv_text = Path(csv_path).read_text(encoding="utf-8-sig")
            report_text = Path(text_path).read_text(encoding="utf-8-sig")

        self.assertIn("ASSEMBLY_SUMMARY", csv_text)
        self.assertIn("TOP / GOOD_OCC", csv_text)
        self.assertIn("Final assembly full-load result: SUCCESS", report_text)

    def test_source_has_load_calls_but_no_persistence_or_structure_mutation(self):
        source = JOURNAL.read_text(encoding="utf-8")
        self.assertIn("LoadThisPartFully", source)
        self.assertIn("work_part.LoadFully()", source)
        for forbidden in (
            ".Save(",
            ".SaveAs(",
            ".Checkout",
            ".Checkin",
            ".Suppress(",
            ".Unsuppress(",
            ".ReplaceComponent(",
            ".Close(",
        ):
            self.assertNotIn(forbidden, source)


if __name__ == "__main__":
    unittest.main()
