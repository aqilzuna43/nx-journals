import csv
import hashlib
import importlib.util
import json
import os
import sys
import tempfile
import types
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
JOURNAL = ROOT / "from_git" / "journals" / "28_probe_bom_structure.py"


def load_journal():
    nxopen = types.ModuleType("NXOpen")
    nxopen.__path__ = []
    nxopen.NXObject = types.SimpleNamespace(
        AttributeType=types.SimpleNamespace(String="String")
    )
    nxopen.Session = types.SimpleNamespace(
        LibraryUnloadOption=types.SimpleNamespace(
            Immediately="Immediately",
            AtTermination="AtTermination",
        )
    )
    nxopen_uf = types.ModuleType("NXOpen.UF")
    nxopen_uf.UFSession = types.SimpleNamespace(GetUFSession=lambda: None)
    nxopen.UF = nxopen_uf

    prior_nx = sys.modules.get("NXOpen")
    prior_uf = sys.modules.get("NXOpen.UF")
    sys.modules["NXOpen"] = nxopen
    sys.modules["NXOpen.UF"] = nxopen_uf
    try:
        spec = importlib.util.spec_from_file_location("journal28", JOURNAL)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        return module
    finally:
        if prior_nx is None:
            sys.modules.pop("NXOpen", None)
        else:
            sys.modules["NXOpen"] = prior_nx
        if prior_uf is None:
            sys.modules.pop("NXOpen.UF", None)
        else:
            sys.modules["NXOpen.UF"] = prior_uf


class NamedValue:
    def __init__(self, name):
        self.name = name

    def __str__(self):
        return self.name


class FakeAttributeInfo:
    def __init__(self, title, value, attribute_type="String", unset=False):
        self.Title = title
        self.TitleAlias = ""
        self.Category = "DB Component Instance"
        self.Type = NamedValue(attribute_type)
        self.StringValue = value if attribute_type == "String" else ""
        self.BooleanValue = value if attribute_type == "Boolean" else False
        self.IntegerValue = value if attribute_type == "Integer" else 0
        self.RealValue = value if attribute_type == "Real" else 0.0
        self.TimeValue = value if attribute_type == "Time" else ""
        self.ReferenceValue = value if attribute_type == "Reference" else ""
        self.Unset = unset
        self.Locked = False
        self.Inherited = False
        self.IsOverride = False
        self.OwnedBySystem = False
        self.PdmBased = False
        self.NotSaved = False
        self.Array = False
        self.ArrayElementIndex = -1

    def ToString(self):
        return str(self.StringValue)


class FakePrototype:
    def __init__(self, part_number, tag, attributes=None):
        self.Name = part_number
        self.Leaf = part_number
        self.FullPath = "@DB/{0}/A".format(part_number)
        self.Tag = tag
        self.PartLoadState = NamedValue("FullyLoaded")
        values = {
            "DB_PART_NO": part_number,
            "DB_PART_REV": "A",
            "DB_PART_NAME": part_number + " NAME",
            "Stocking_Type": "MAKE",
        }
        values.update(attributes or {})
        self.attributes = values

    def GetStringAttribute(self, title):
        if title not in self.attributes:
            raise KeyError(title)
        return self.attributes[title]

    def GetUserAttribute(self, *args):
        raise KeyError(args[0])


class FakeComponent:
    def __init__(
        self,
        name,
        tag,
        prototype=None,
        children=None,
        attributes=None,
        suppressed=False,
        children_error=None,
        inventory_error=None,
    ):
        self.Name = name
        self.DisplayName = name
        self.Tag = tag
        self.JournalIdentifier = "COMPONENT {0}".format(tag)
        self.Prototype = prototype
        self.ReferenceSet = "MODEL"
        self.Layer = 1
        self.IsSuppressed = suppressed
        self._children = list(children or [])
        self._attributes = list(attributes or [])
        self._children_error = children_error
        self._inventory_error = inventory_error
        self.targeted_presence_calls = 0
        self.targeted_value_calls = 0
        self.full_inventory_calls = 0

    def GetChildren(self):
        if self._children_error:
            raise RuntimeError(self._children_error)
        return list(self._children)

    def GetInstanceUserAttributes(self, include_unset=False):
        self.full_inventory_calls += 1
        if self._inventory_error:
            raise RuntimeError(self._inventory_error)
        if include_unset:
            return list(self._attributes)
        return [item for item in self._attributes if not item.Unset]

    def HasInstanceUserAttribute(self, title, attribute_type, index):
        self.targeted_presence_calls += 1
        if self._inventory_error:
            raise RuntimeError(self._inventory_error)
        return any(item.Title == title and not item.Unset for item in self._attributes)

    def GetInstanceUserAttribute(self, title, attribute_type, index):
        self.targeted_value_calls += 1
        if self._inventory_error:
            raise RuntimeError(self._inventory_error)
        for item in self._attributes:
            if item.Title == title and not item.Unset:
                return item
        raise KeyError(title)

    def GetNonGeometricState(self):
        return False

    def GetComponentRepresentationMode(self):
        return NamedValue("Exact")


class FakeWorkPart(FakePrototype):
    def __init__(self, root, modified=False):
        super().__init__("TOP", 1000)
        self.DisplayName = "TOP"
        self.IsModified = modified
        self.ComponentAssembly = types.SimpleNamespace(RootComponent=root)


class FakeListingWindow:
    def __init__(self):
        self.lines = []

    def Open(self):
        return None

    def WriteLine(self, text):
        self.lines.append(text)


class FakeSession:
    def __init__(self, work_part):
        self.Parts = types.SimpleNamespace(Work=work_part)
        self.ReleaseName = "NX X 2506"
        self.ReleaseNumber = 2506
        self.BuildNumber = 1
        self.ApplicationName = "UG_APP_MODELING"
        self.ListingWindow = FakeListingWindow()


class FakeAssem:
    def AskStableIdOfInstance(self, tag):
        return "STABLE-{0}".format(tag)


class FakeUFSession:
    def __init__(self):
        self.Assem = FakeAssem()


class FailingAssem:
    def __init__(self):
        self.calls = 0

    def AskStableIdOfInstance(self, tag):
        self.calls += 1
        raise RuntimeError("stable ID unsupported for this occurrence")


class FailingUFSession:
    def __init__(self):
        self.Assem = FailingAssem()


def attribute(title, value, attribute_type="String", unset=False):
    return FakeAttributeInfo(title, value, attribute_type, unset)


def make_work(children, root_attributes=None):
    work_placeholder = FakePrototype("TOP", 1000)
    root = FakeComponent(
        "ROOT",
        1,
        prototype=work_placeholder,
        children=children,
        attributes=root_attributes,
    )
    work = FakeWorkPart(root)
    root.Prototype = work
    return work


class Journal28Tests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.journal = load_journal()

    def collect(self, work):
        return self.journal.collect_occurrences(
            work,
            FakeUFSession(),
            "run28",
            "2026-08-29T10:00:00+08:00",
        )

    def test_raw_preorder_keeps_reference_suppressed_and_all_descendants(self):
        reference_leaf = FakeComponent(
            "CONTEXT_LEAF", 4, FakePrototype("REF_LEAF", 104)
        )
        reference_parent = FakeComponent(
            "CONTEXT_PARENT",
            3,
            FakePrototype("REF_PARENT", 103),
            children=[reference_leaf],
            attributes=[attribute("REFERENCE_COMPONENT", "")],
        )
        suppressed_leaf = FakeComponent(
            "SUPPRESSED_LEAF", 6, FakePrototype("SUP_LEAF", 106)
        )
        suppressed_parent = FakeComponent(
            "SUPPRESSED_PARENT",
            5,
            FakePrototype("SUP_PARENT", 105),
            children=[suppressed_leaf],
            suppressed=True,
        )
        work = make_work([reference_parent, suppressed_parent])

        rows, errors, capped = self.collect(work)

        self.assertEqual([], errors)
        self.assertFalse(capped)
        self.assertEqual(
            [
                "ROOT",
                "CONTEXT_PARENT",
                "CONTEXT_LEAF",
                "SUPPRESSED_PARENT",
                "SUPPRESSED_LEAF",
            ],
            [row["COMPONENT_DISPLAY_NAME"] for row in rows],
        )
        self.assertEqual("REFERENCE_ONLY", rows[1]["DIRECT_CONTROL_CLASSIFICATION"])
        self.assertEqual("YES", rows[1]["REFERENCE_COMPONENT_PRESENT"])
        self.assertEqual("PRESENT_BLANK", rows[1]["REFERENCE_COMPONENT_VALUE_STATE"])
        self.assertEqual("EXCLUDE_REFERENCE_SUBTREE", rows[1]["CURRENT_EXTENDED_BOM_PREDICTION"])
        self.assertEqual("EXCLUDED_BY_ANCESTOR", rows[2]["CURRENT_EXTENDED_BOM_PREDICTION"])
        self.assertEqual(rows[1]["STRUCTURAL_PATH"], rows[2]["NEAREST_CONTROL_ANCESTOR_PATH"])
        self.assertEqual("YES", rows[3]["SUPPRESSED"])
        self.assertEqual("EXCLUDE_SUPPRESSED_SUBTREE", rows[3]["CURRENT_EXTENDED_BOM_PREDICTION"])
        self.assertEqual("EXCLUDED_BY_ANCESTOR", rows[4]["CURRENT_EXTENDED_BOM_PREDICTION"])

    def test_duplicate_prototypes_and_repeated_names_remain_distinct_rows(self):
        shared = FakePrototype("SHARED", 200)
        work = make_work(
            [
                FakeComponent("SAME_NAME", 10, shared),
                FakeComponent("SAME_NAME", 11, shared),
            ]
        )

        rows, _, _ = self.collect(work)

        self.assertEqual(3, len(rows))
        self.assertEqual("SHARED", rows[1]["DB_PART_NO"])
        self.assertEqual("SHARED", rows[2]["DB_PART_NO"])
        self.assertEqual(rows[1]["PROTOTYPE_TAG"], rows[2]["PROTOTYPE_TAG"])
        self.assertNotEqual(rows[1]["STRUCTURAL_PATH"], rows[2]["STRUCTURAL_PATH"])
        self.assertEqual([1, 2, 3], [row["SEQUENCE"] for row in rows])

    def test_missing_prototype_keeps_row_and_marks_required_evidence_incomplete(self):
        work = make_work(
            [
                FakeComponent("MISSING_MODEL", 15, prototype=None),
                FakeComponent("GOOD_MODEL", 16, FakePrototype("GOOD_MODEL", 216)),
            ]
        )

        rows, errors, _ = self.collect(work)

        self.assertEqual([], errors)
        self.assertEqual(
            ["ROOT", "MISSING_MODEL", "GOOD_MODEL"],
            [row["COMPONENT_DISPLAY_NAME"] for row in rows],
        )
        self.assertEqual("", rows[1]["PROTOTYPE_NAME"])
        self.assertEqual("ERROR", rows[1]["ROW_EVIDENCE_STATUS"])
        self.assertIn("NX returned no prototype object", rows[1]["PROBE_ERRORS"])
        self.assertIn(
            "NX returned no prototype object",
            " | ".join(rows[1]["_critical_read_error_items"]),
        )

    def test_legacy_keyword_prediction_is_reported_without_pruning(self):
        child = FakeComponent("REAL_CHILD", 18, FakePrototype("REAL_CHILD", 218))
        datum = FakeComponent(
            "DATUM_GUIDE",
            17,
            FakePrototype("DATUM_GUIDE", 217),
            children=[child],
        )
        work = make_work([datum])

        rows, _, _ = self.collect(work)

        self.assertEqual("DATUM", rows[1]["LEGACY_IGNORE_KEYWORD_MATCH"])
        self.assertEqual(
            "EXCLUDE_NAME_KEYWORD_SUBTREE",
            rows[1]["CURRENT_EXTENDED_BOM_PREDICTION"],
        )
        self.assertEqual("EXCLUDED_BY_ANCESTOR", rows[2]["CURRENT_EXTENDED_BOM_PREDICTION"])

    def test_presence_value_and_unreadable_states_are_not_collapsed(self):
        work = make_work(
            [
                FakeComponent(
                    "BLANK", 20, FakePrototype("BLANK", 220),
                    attributes=[attribute("REFERENCE_COMPONENT", "")],
                ),
                FakeComponent("ABSENT", 21, FakePrototype("ABSENT", 221)),
                FakeComponent(
                    "NONSTANDARD", 22, FakePrototype("NONSTANDARD", 222),
                    attributes=[attribute("PLIST_IGNORE_MEMBER", "off")],
                ),
                FakeComponent(
                    "UNREADABLE", 23, FakePrototype("UNREADABLE", 223),
                    inventory_error="cannot read instance attributes",
                ),
            ]
        )

        rows, _, _ = self.collect(work)
        blank, absent, nonstandard, unreadable = rows[1:]

        self.assertEqual("PRESENT_BLANK", blank["REFERENCE_COMPONENT_VALUE_STATE"])
        self.assertEqual("NO", absent["REFERENCE_COMPONENT_PRESENT"])
        self.assertEqual("ABSENT", absent["REFERENCE_COMPONENT_VALUE_STATE"])
        self.assertEqual("PRESENT_NONSTANDARD", nonstandard["PLIST_IGNORE_MEMBER_VALUE_STATE"])
        self.assertEqual("NATIVE_MEMBER_ONLY", nonstandard["DIRECT_CONTROL_CLASSIFICATION"])
        self.assertEqual("INCLUDE", nonstandard["CURRENT_EXTENDED_BOM_PREDICTION"])
        self.assertEqual("UNREADABLE", unreadable["DIRECT_CONTROL_CLASSIFICATION"])
        self.assertEqual("ERROR", unreadable["REFERENCE_COMPONENT_READ_STATUS"])
        self.assertIn("cannot read instance attributes", unreadable["PROBE_ERRORS"])

    def test_all_direct_control_classifications(self):
        cases = [
            ([], "NONE"),
            ([attribute("REFERENCE_COMPONENT", "YES")], "REFERENCE_ONLY"),
            ([attribute("PLIST_IGNORE_MEMBER", "YES")], "NATIVE_MEMBER_ONLY"),
            ([attribute("PLIST_IGNORE_SUBASSEMBLY", "YES")], "NATIVE_SUBASSEMBLY_ONLY"),
            (
                [
                    attribute("PLIST_IGNORE_MEMBER", "YES"),
                    attribute("PLIST_IGNORE_SUBASSEMBLY", "YES"),
                ],
                "NATIVE_EXCLUDE_PAIR",
            ),
            (
                [
                    attribute("REFERENCE_COMPONENT", "YES"),
                    attribute("PLIST_IGNORE_MEMBER", "YES"),
                ],
                "MULTIPLE_CONTROLS",
            ),
        ]
        for index, (attributes, expected) in enumerate(cases):
            with self.subTest(expected=expected):
                work = make_work(
                    [
                        FakeComponent(
                            "CASE_{0}".format(index),
                            30 + index,
                            FakePrototype("CASE_{0}".format(index), 300 + index),
                            attributes=attributes,
                        )
                    ]
                )
                rows, _, _ = self.collect(work)
                self.assertEqual(expected, rows[1]["DIRECT_CONTROL_CLASSIFICATION"])

    def test_failed_child_enumeration_preserves_unaffected_sibling(self):
        bad = FakeComponent(
            "BAD_BRANCH",
            40,
            FakePrototype("BAD", 440),
            children_error="branch unavailable",
        )
        good = FakeComponent("GOOD_BRANCH", 41, FakePrototype("GOOD", 441))
        work = make_work([bad, good])

        rows, errors, _ = self.collect(work)

        self.assertEqual(["ROOT", "BAD_BRANCH", "GOOD_BRANCH"], [row["COMPONENT_DISPLAY_NAME"] for row in rows])
        self.assertEqual(1, len(errors))
        self.assertEqual(rows[1]["STRUCTURAL_PATH"], errors[0]["path"])
        self.assertIn("branch unavailable", errors[0]["error"])

    def test_control_descendant_counts_handle_nested_controls_linearly(self):
        nested_leaf = FakeComponent("LEAF", 52, FakePrototype("LEAF", 552))
        nested_control = FakeComponent(
            "NESTED",
            51,
            FakePrototype("NESTED", 551),
            children=[nested_leaf],
            attributes=[attribute("PLIST_IGNORE_MEMBER", "YES")],
        )
        outer = FakeComponent(
            "OUTER",
            50,
            FakePrototype("OUTER", 550),
            children=[nested_control],
            attributes=[attribute("REFERENCE_COMPONENT", "")],
        )
        work = make_work([outer])
        rows, _, _ = self.collect(work)

        counts = self.journal.control_descendant_counts(rows)

        self.assertEqual(2, counts[0]["descendant_count"])
        self.assertEqual(1, counts[1]["descendant_count"])

    def test_run_writes_matching_utf8_bom_csv_and_json_summary(self):
        work = make_work(
            [
                FakeComponent(
                    "REFERENCE_PARENT",
                    60,
                    FakePrototype("REFERENCE_PARENT", 660),
                    attributes=[attribute("REFERENCE_COMPONENT", "")],
                )
            ]
        )
        session = FakeSession(work)
        now = __import__("datetime").datetime(2026, 8, 29, 10, 30, 0)

        with tempfile.TemporaryDirectory() as folder:
            csv_path, json_path, report = self.journal.run(
                session,
                uf_session=FakeUFSession(),
                run_datetime=now,
                output_root=folder,
                run_id="abc12345",
            )
            csv_bytes = Path(csv_path).read_bytes()
            parsed = json.loads(Path(json_path).read_text(encoding="utf-8"))
            with Path(csv_path).open(encoding="utf-8-sig", newline="") as handle:
                csv_rows = list(csv.DictReader(handle))
            remaining_partials = list(Path(folder).rglob("*.partial"))

        self.assertEqual("COMPLETE", report["run_status"])
        self.assertEqual(
            "J28-NX2506-BOM-STRUCTURE-PROBE-V2-MEMSAFE",
            parsed["journal_build"],
        )
        self.assertEqual(2, parsed["schema_version"])
        self.assertEqual(b"\xef\xbb\xbf", csv_bytes[:3])
        self.assertEqual(2, parsed["summary"]["occurrence_count"])
        self.assertEqual(2, len(csv_rows))
        self.assertTrue(
            all(row["INSTANCE_STABLE_ID_STATUS"] == "NOT_APPLICABLE" for row in csv_rows)
        )
        self.assertEqual(
            ["REFERENCE_COMPONENT"],
            [item["title"] for item in parsed["flagged_occurrences"][0]["control_attributes"]],
        )
        self.assertNotIn("instance_attributes", parsed["flagged_occurrences"][0])
        self.assertEqual(
            hashlib.sha256(csv_bytes).hexdigest(), parsed["csv_sha256"]
        )
        self.assertEqual("abc12345", parsed["run_id"])
        self.assertFalse(parsed["work_part_modified"]["changed"])
        self.assertEqual([], remaining_partials)

    def test_run_marks_traversal_failure_incomplete(self):
        work = make_work(
            [
                FakeComponent(
                    "BAD", 70, FakePrototype("BAD", 770),
                    children_error="cannot enumerate",
                ),
                FakeComponent("GOOD", 71, FakePrototype("GOOD", 771)),
            ]
        )
        session = FakeSession(work)

        with tempfile.TemporaryDirectory() as folder:
            csv_path, json_path, report = self.journal.run(
                session,
                uf_session=FakeUFSession(),
                output_root=folder,
                run_id="incomplete",
            )
            parsed = json.loads(Path(json_path).read_text(encoding="utf-8"))

        self.assertTrue(csv_path)
        self.assertEqual("INCOMPLETE", report["run_status"])
        self.assertEqual(1, parsed["summary"]["traversal_error_count"])
        self.assertEqual(3, parsed["summary"]["occurrence_count"])

    def test_uf_stable_id_probe_is_never_called(self):
        work = make_work(
            [FakeComponent("GOOD", 75, FakePrototype("GOOD", 775))]
        )
        session = FakeSession(work)
        uf_session = FailingUFSession()

        with tempfile.TemporaryDirectory() as folder:
            _, json_path, report = self.journal.run(
                session,
                uf_session=uf_session,
                output_root=folder,
                run_id="no_stable_id",
            )
            parsed = json.loads(Path(json_path).read_text(encoding="utf-8"))

        self.assertEqual("COMPLETE", report["run_status"])
        self.assertEqual(0, parsed["summary"]["read_error_occurrence_count"])
        self.assertEqual(0, parsed["summary"]["critical_read_error_occurrence_count"])
        self.assertEqual([], parsed["flagged_occurrences"])
        self.assertEqual(0, uf_session.Assem.calls)

    def test_numeric_nx_attribute_type_decodes_string_value(self):
        info = FakeAttributeInfo("REFERENCE_COMPONENT", "YES")
        info.Type = 5

        item = self.journal.attribute_info_dict(info)

        self.assertEqual("STRING", item["type"])
        self.assertEqual("YES", item["value"])

    def test_targeted_controls_do_not_enumerate_or_retain_components(self):
        component = FakeComponent(
            "REFERENCE",
            9001,
            FakePrototype("REFERENCE", 9002),
            attributes=[attribute("REFERENCE_COMPONENT", "")],
        )
        work = make_work([component])

        rows, _, _ = self.collect(work)

        self.assertEqual(3, component.targeted_presence_calls)
        self.assertEqual(1, component.targeted_value_calls)
        self.assertEqual(0, component.full_inventory_calls)
        self.assertNotIn("_component", rows[1])
        self.assertEqual(
            ["REFERENCE_COMPONENT"],
            [item["title"] for item in rows[1]["_control_attributes"]],
        )

    def test_unloaded_parent_skips_targeted_control_calls(self):
        child = FakeComponent(
            "CHILD",
            9011,
            FakePrototype("CHILD", 9012),
            inventory_error="targeted read must not run",
        )
        parent_prototype = FakePrototype("PARENT", 9013)
        parent_prototype.PartLoadState = NamedValue("NotLoaded")
        parent = FakeComponent(
            "PARENT",
            9014,
            parent_prototype,
            children=[child],
        )
        work = make_work([parent])

        rows, _, _ = self.collect(work)

        self.assertEqual(0, child.targeted_presence_calls)
        self.assertEqual("UNAVAILABLE", rows[2]["REFERENCE_COMPONENT_READ_STATUS"])
        self.assertIn("parent assembly prototype is not loaded", rows[2]["PROBE_ERRORS"])

    def test_unload_is_deferred_until_nx_termination(self):
        self.assertEqual("AtTermination", self.journal.get_unload_option(None))

    def test_max_occurrences_env_override_limits_traversal(self):
        work = make_work(
            [
                FakeComponent("A", 1, FakePrototype("A", 101)),
                FakeComponent("B", 2, FakePrototype("B", 102)),
                FakeComponent("C", 3, FakePrototype("C", 103)),
            ]
        )
        previous = os.environ.get("NX_J28_MAX_OCCURRENCES")
        os.environ["NX_J28_MAX_OCCURRENCES"] = "2"
        try:
            rows, errors, capped = self.journal.collect_occurrences(
                work, FakeUFSession(), "run28", "2026-08-29T10:00:00+08:00"
            )
        finally:
            if previous is None:
                os.environ.pop("NX_J28_MAX_OCCURRENCES", None)
            else:
                os.environ["NX_J28_MAX_OCCURRENCES"] = previous

        self.assertTrue(capped)
        self.assertEqual(2, len(rows))
        self.assertIn("Safety limit reached: 2", errors[0]["error"])

    def test_max_occurrences_invalid_env_falls_back_to_default(self):
        previous = os.environ.get("NX_J28_MAX_OCCURRENCES")
        os.environ["NX_J28_MAX_OCCURRENCES"] = "not-a-number"
        try:
            resolved = self.journal.max_occurrences()
        finally:
            if previous is None:
                os.environ.pop("NX_J28_MAX_OCCURRENCES", None)
            else:
                os.environ["NX_J28_MAX_OCCURRENCES"] = previous

        self.assertEqual(self.journal.MAX_OCCURRENCES, resolved)

    def test_root_failure_writes_failed_json_without_csv(self):
        work = make_work([])
        work.ComponentAssembly.RootComponent = None
        session = FakeSession(work)

        with tempfile.TemporaryDirectory() as folder:
            csv_path, json_path, report = self.journal.run(
                session,
                uf_session=FakeUFSession(),
                output_root=folder,
                run_id="failed",
            )
            parsed = json.loads(Path(json_path).read_text(encoding="utf-8"))
            csv_files = list(Path(folder).rglob("*.csv"))

        self.assertEqual("", csv_path)
        self.assertEqual("FAILED", report["run_status"])
        self.assertIn("not an assembly", parsed["fatal_error"])
        self.assertEqual([], csv_files)

    def test_unexpected_json_write_failure_retains_partial_files(self):
        work = make_work([])
        rows, _, _ = self.collect(work)
        report = self.journal.base_report("partial", "timestamp", FakeSession(work))
        report["not_json_serializable"] = object()

        with tempfile.TemporaryDirectory() as folder:
            with self.assertRaises(TypeError):
                self.journal.write_artifacts(folder, "J28_PARTIAL", rows, report)
            csv_partial = Path(folder) / "J28_PARTIAL.csv.partial"
            json_partial = Path(folder) / "J28_PARTIAL.json.partial"
            csv_final = Path(folder) / "J28_PARTIAL.csv"

            self.assertTrue(csv_partial.exists())
            self.assertTrue(json_partial.exists())
            self.assertFalse(csv_final.exists())

    def test_source_contains_no_nx_mutation_or_load_calls(self):
        source = JOURNAL.read_text(encoding="utf-8")
        self.assertIn("HasInstanceUserAttribute", source)
        self.assertIn("GetInstanceUserAttribute", source)
        self.assertNotIn("GetInstanceUserAttributes(", source)
        self.assertNotIn("AskStableIdOfInstance", source)
        self.assertNotIn('row["_component"]', source)
        self.assertNotIn('row["_attribute_inventory"]', source)
        for forbidden in (
            ".Save(",
            ".SaveAs(",
            ".Checkout",
            ".Checkin",
            "LoadThisPartFully",
            ".LoadFully(",
            "SetInstanceUserAttribute",
            "DeleteInstanceUserAttribute",
            "SuppressComponents",
            ".Suppress(",
            ".Unsuppress(",
            "ReplaceReferenceSet",
            "SetNonGeometricState",
            ".Blank(",
            ".Unblank(",
            "DoUpdate",
            "SetWorkPart",
        ):
            self.assertNotIn(forbidden, source)


if __name__ == "__main__":
    unittest.main()
