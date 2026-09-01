import csv
import importlib.util
import json
import sys
import tempfile
import types
import unittest
from pathlib import Path
from unittest import mock


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "from_git" / "utils"))

import attribute_reconciliation as core


class EnumValue:
    def __init__(self, name):
        self.name = name


class FakeIterator:
    def SetIncludeOnlyCategory(self, value):
        self.category = value

    def SetIncludeOnlyTitle(self, value):
        self.title = value

    def SetIncludeAlsoUnset(self, value):
        self.include_unset = value

    def FreeResource(self):
        self.freed = True


class FakeAttribute:
    def __init__(self, category, title, value, unset=False, **flags):
        self.Category = category
        self.Title = title
        self.Type = EnumValue("String")
        self.StringValue = value
        self.Unset = unset
        self.Locked = flags.get("locked", False)
        self.OwnedBySystem = flags.get("owned_by_system", False)
        self.PdmBased = flags.get("pdm_based", False)
        self.Required = flags.get("required", False)
        self.NotSaved = flags.get("not_saved", False)


class FakePart:
    _tag = 1

    def __init__(self, name, attributes=(), bodies=()):
        self.Name = name
        self.JournalIdentifier = name
        self.attributes = list(attributes)
        self.Bodies = list(bodies)
        self.Tag = FakePart._tag
        FakePart._tag += 1
        self.ComponentAssembly = types.SimpleNamespace(RootComponent=None)

    def CreateAttributeIterator(self):
        return FakeIterator()

    def GetUserAttributes(self, iterator=None):
        if iterator is None:
            return list(self.attributes)
        return [
            item
            for item in self.attributes
            if item.Category == iterator.category and item.Title == iterator.title
        ]


class FakeComponent:
    _tag = 100

    def __init__(
        self,
        name,
        prototype=None,
        children=(),
        suppressed=False,
        position=(0, 0, 0),
        orientation=((1, 0, 0), (0, 1, 0), (0, 0, 1)),
        string_attributes=None,
    ):
        self.Name = name
        self.DisplayName = name
        self.Prototype = prototype
        self._children = list(children)
        self.IsSuppressed = suppressed
        self.position = position
        self.orientation = orientation
        self.string_attributes = dict(string_attributes or {})
        self.Tag = FakeComponent._tag
        FakeComponent._tag += 1

    def GetChildren(self):
        return list(self._children)

    def GetStringAttribute(self, title):
        if title not in self.string_attributes:
            raise AttributeError("No such attribute: " + title)
        return self.string_attributes[title]

    def GetPosition(self):
        return self.position, self.orientation


class FakeBody:
    _tag = 500

    def __init__(self):
        self.Tag = FakeBody._tag
        FakeBody._tag += 1
        self.IsSolidBody = True
        self.JournalIdentifier = "BODY_{0}".format(self.Tag)


def attrs(part_number, revision="A", name=None):
    return [
        FakeAttribute("Cad0Design", "DB_PART_NO", part_number),
        FakeAttribute("Cad0DesignRevision", "DB_PART_REV", revision),
        FakeAttribute("Cad0Design", "DB_PART_NAME", name or part_number),
    ]


MINI_CONFIG = {
    "attributes": [
        {"logical_name": "part_number", "category": "Cad0Design", "attribute": "DB_PART_NO", "type": "String"},
        {"logical_name": "revision", "category": "Cad0DesignRevision", "attribute": "DB_PART_REV", "type": "String"},
        {"logical_name": "part_name", "category": "Cad0Design", "attribute": "DB_PART_NAME", "type": "String"},
    ]
}


def load_j05():
    nxopen = types.ModuleType("NXOpen")
    nxopen_uf = types.ModuleType("NXOpen.UF")
    nxopen_pdm = types.ModuleType("NXOpen.PDM")

    class CheckoutInput:
        def __init__(
            self,
            input_comment,
            input_change_id,
            allow_remote,
            explicit_checkout,
            include_secondary,
        ):
            self.InputComment = input_comment
            self.InputChangeId = input_change_id
            self.AllowRemote = allow_remote
            self.ExplicitCheckOut = explicit_checkout
            self.IncludeSecondary = include_secondary

    nxopen_pdm.PdmPart = types.SimpleNamespace(CheckoutInput=CheckoutInput)
    nxopen.UF = nxopen_uf
    nxopen.PDM = nxopen_pdm
    nxopen.Session = types.SimpleNamespace(
        MarkVisibility=types.SimpleNamespace(Invisible="Invisible")
    )
    prior_nx = sys.modules.get("NXOpen")
    prior_uf = sys.modules.get("NXOpen.UF")
    prior_pdm = sys.modules.get("NXOpen.PDM")
    sys.modules["NXOpen"] = nxopen
    sys.modules["NXOpen.UF"] = nxopen_uf
    sys.modules["NXOpen.PDM"] = nxopen_pdm
    try:
        spec = importlib.util.spec_from_file_location(
            "j05_under_test", ROOT / "from_git" / "journals" / "05_bulk_attribute_updater.py"
        )
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
        if prior_pdm is None:
            sys.modules.pop("NXOpen.PDM", None)
        else:
            sys.modules["NXOpen.PDM"] = prior_pdm


def load_j04():
    nxopen = types.ModuleType("NXOpen")
    prior_nx = sys.modules.get("NXOpen")
    sys.modules["NXOpen"] = nxopen
    try:
        spec = importlib.util.spec_from_file_location(
            "j04_under_test",
            ROOT
            / "from_git"
            / "journals"
            / "04_assembly_attribute_audit.py",
        )
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        return module
    finally:
        if prior_nx is None:
            sys.modules.pop("NXOpen", None)
        else:
            sys.modules["NXOpen"] = prior_nx


class ConfigAndValuesTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.config_path = ROOT / "from_git" / "config" / "attribute_reconciliation.json"
        cls.config = json.loads(cls.config_path.read_text(encoding="utf-8"))

    def test_config_is_valid_nx_authoritative_and_save_enabled(self):
        self.assertIs(core.validate_config(self.config), self.config)
        self.assertEqual("NX_TEAMCENTER", self.config["authority"])
        self.assertEqual(
            "SAVE_CHANGED_PARTS", self.config["save_policy"]
        )

    def test_runtime_config_contains_no_snapshot_or_evidence_identity(self):
        payload = self.config_path.read_text(encoding="utf-8")
        for forbidden in ("source_snapshot", '"snapshot"', "264MN025450A01", "workstation", "user_id"):
            self.assertNotIn(forbidden, payload)

    def test_schema_rejects_duplicate_composite_attribute(self):
        config = json.loads(json.dumps(self.config))
        config["attributes"].append(dict(config["attributes"][0], logical_name="duplicate_identity"))
        with self.assertRaises(core.ReconciliationError):
            core.validate_config(config)

    def test_placeholders_typed_normalization_and_controlled_values(self):
        self.assertTrue(core.is_placeholder("  tBc ", self.config))
        self.assertTrue(core.is_placeholder("00-Jan-0", self.config))
        boolean_rule = {"comparison": "BOOLEAN_ALIAS"}
        self.assertEqual("Y", core.normalize_value(" yes ", boolean_rule, self.config))
        number_rule = {"comparison": "NUMBER"}
        self.assertEqual("1.25", core.normalize_value("1.2500", number_rule, self.config))
        rule = {
            "comparison": "TRIMMED_CASE_INSENSITIVE",
            "required_for_certification": True,
            "allowed_values": ["MAKE", "BUY"],
        }
        result = {"status": "POPULATED", "raw_value": "invalid"}
        self.assertEqual("INVALID_CONTROLLED_VALUE", core.validate_attribute_value(result, rule, self.config)[1])

    def test_set_unset_missing_and_category_title_identity(self):
        part = FakePart(
            "P",
            [
                FakeAttribute("Right", "VALUE", "set"),
                FakeAttribute("Wrong", "VALUE", "shadow"),
                FakeAttribute("Right", "UNSET", "default", unset=True),
            ],
        )
        set_result = core.read_attribute(part, {"category": "Right", "attribute": "VALUE", "type": "String"})
        unset_result = core.read_attribute(part, {"category": "Right", "attribute": "UNSET", "type": "String"})
        missing = core.read_attribute(part, {"category": "Right", "attribute": "MISSING", "type": "String"})
        self.assertEqual(("POPULATED", "set"), (set_result["status"], set_result["raw_value"]))
        self.assertEqual(("UNSET", "default"), (unset_result["status"], unset_result["raw_value"]))
        self.assertEqual("MISSING", missing["status"])

    def test_nx2506_numeric_attribute_type_values_are_decoded(self):
        string_info = FakeAttribute("Cad0Design", "DB_PART_NO", "PART-001")
        string_info.Type = 5
        real_info = FakeAttribute("Materials", "NX_Mass", "")
        real_info.Type = 4
        real_info.RealValue = 12.5

        self.assertEqual(("PART-001", "String"), core._attribute_value(string_info))
        self.assertEqual((12.5, "Real"), core._attribute_value(real_info))


class AssemblyTests(unittest.TestCase):
    def test_immediate_parent_quantity_repeated_subtree_and_suppression(self):
        root = FakePart("ROOT", attrs("ROOT"))
        part_a = FakePart("A", attrs("A"))
        part_b = FakePart("B", attrs("B"))
        part_c = FakePart("C", attrs("C"))
        suppressed_part = FakePart("S", attrs("S"))

        a1 = FakeComponent("A1", part_a, [FakeComponent("C-under-A", part_c)])
        a2 = FakeComponent("A2", part_a, [FakeComponent("C-under-A2", part_c)])
        b = FakeComponent("B1", part_b, [FakeComponent("C-under-B", part_c)])
        suppressed = FakeComponent("suppressed", suppressed_part, suppressed=True)
        root.ComponentAssembly.RootComponent = FakeComponent("ROOT-COMP", root, [a1, a2, b, suppressed])

        nodes, findings = core.collect_bom_nodes(root, MINI_CONFIG)
        identities = [(node["parent_part_number"], node["part_number"], node["quantity"]) for node in nodes]
        self.assertEqual(
            [("", "ROOT", 1), ("ROOT", "A", 2), ("A", "C", 1), ("ROOT", "B", 1), ("B", "C", 1)],
            identities,
        )
        self.assertFalse(findings)
        self.assertNotIn("S", [node["part_number"] for node in nodes])

    def test_unloaded_prototype_is_reported(self):
        root = FakePart("ROOT", attrs("ROOT"))
        root.ComponentAssembly.RootComponent = FakeComponent(
            "ROOT-COMP", root, [FakeComponent("unloaded", None)]
        )
        nodes, findings = core.collect_bom_nodes(root, MINI_CONFIG)
        self.assertEqual(1, len(nodes))
        self.assertEqual("MISSING_MODEL", findings[0]["code"])

    def test_duplicate_identity_on_distinct_prototypes_is_blocking(self):
        root = FakePart("ROOT", attrs("ROOT"))
        first = FakePart("DUP-1", attrs("DUP"))
        second = FakePart("DUP-2", attrs("DUP"))
        root.ComponentAssembly.RootComponent = FakeComponent(
            "ROOT-COMP",
            root,
            [
                FakeComponent("first", first),
                FakeComponent("other-parent", FakePart("P", attrs("P")), [FakeComponent("second", second)]),
            ],
        )
        _, findings = core.collect_bom_nodes(root, MINI_CONFIG)
        self.assertIn("AMBIGUOUS_MATCH", [finding["code"] for finding in findings])

    def test_transformed_bounding_box_union(self):
        root_body = FakeBody()
        child_body = FakeBody()
        root = FakePart("ROOT", attrs("ROOT"), [root_body])
        child = FakePart("CHILD", attrs("CHILD"), [child_body])
        child_component = FakeComponent("child", child, position=(10, 0, 0))
        root.ComponentAssembly.RootComponent = FakeComponent("ROOT-COMP", root, [child_component])

        boxes = {
            root_body.Tag: ([0, 0, 0], [1, 0, 0, 0, 1, 0, 0, 0, 1], [1, 1, 1]),
            child_body.Tag: ([0, 0, 0], [1, 0, 0, 0, 1, 0, 0, 0, 1], [2, 2, 2]),
        }

        class Modl:
            def AskBoundingBoxExact(self, tag, csys):
                return boxes[tag]

        dimensions = core.exact_model_dimensions(root, types.SimpleNamespace(Modl=Modl()))
        self.assertEqual((12.0, 2.0, 2.0), dimensions)

    def test_rotated_occurrence_bounding_box(self):
        child_body = FakeBody()
        root = FakePart("ROOT", attrs("ROOT"))
        child = FakePart("CHILD", attrs("CHILD"), [child_body])
        quarter_turn = ((0, -1, 0), (1, 0, 0), (0, 0, 1))
        root.ComponentAssembly.RootComponent = FakeComponent(
            "ROOT-COMP", root, [FakeComponent("child", child, orientation=quarter_turn)]
        )

        class Modl:
            def AskBoundingBoxExact(self, tag, csys):
                return [0, 0, 0], [1, 0, 0, 0, 1, 0, 0, 0, 1], [2, 3, 1]

        dimensions = core.exact_model_dimensions(root, types.SimpleNamespace(Modl=Modl()))
        self.assertEqual((3.0, 2.0, 1.0), dimensions)

    def test_nx2506_aligned_bounding_box_fallback(self):
        body = FakeBody()
        root = FakePart("ROOT", attrs("ROOT"), [body])

        class Modl:
            def AskBoundingBoxAligned(self, tag, csys, expand):
                self.call = (tag, csys, expand)
                return [0, 0, 0], [1, 0, 0, 0, 1, 0, 0, 0, 1], [4, 5, 6]

        modl = Modl()
        dimensions = core.exact_model_dimensions(root, types.SimpleNamespace(Modl=modl))

        self.assertEqual((4.0, 5.0, 6.0), dimensions)
        self.assertEqual((body.Tag, 0, False), modl.call)

    def test_dimension_derivation_fails_without_solids(self):
        with self.assertRaises(core.ReconciliationError):
            core.exact_model_dimensions(FakePart("EMPTY", attrs("EMPTY")), types.SimpleNamespace())


class ScopeAndDriftTests(unittest.TestCase):
    def test_drawing_scope_conflict_and_missing_are_review(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "scope.csv"
            with path.open("w", encoding="utf-8", newline="") as handle:
                writer = csv.writer(handle)
                writer.writerow(["Item Number", "Item Rev", "Drawing Required"])
                writer.writerow(["P1", "A", "YES"])
                writer.writerow(["P1", "A", "NO"])
            scope, findings = core.load_drawing_scope(path)
        self.assertEqual("REVIEW", core.drawing_decision(scope, "P1", "A"))
        self.assertEqual("REVIEW", core.drawing_decision(scope, "MISSING", "A"))
        self.assertEqual("DRAWING_SCOPE_REVIEW", findings[0]["code"])

    def test_master_difference_is_downstream_drift(self):
        nx_row = {
            "BOM Level": 0,
            "DB_PART_NO": "ROOT",
            "DB_PART_NAME": "NX NAME",
            "DB_PART_REV": "A",
            "Quantity": 1,
            "MFG": "M",
            "MPN": "N",
            "Stocking_Type": "MAKE",
        }
        headers = list(core.MASTER_TO_NX_COLUMNS)
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "master.csv"
            with path.open("w", encoding="utf-8", newline="") as handle:
                writer = csv.DictWriter(handle, fieldnames=headers)
                writer.writeheader()
                writer.writerow(
                    {
                        "Level": 5,
                        "Item Number": "ROOT",
                        "Part Description": "STALE NAME",
                        "Item Rev": "A",
                        "Qty": 1,
                        "Mfr. Name": "M",
                        "Mfr. Part Number": "N",
                        "Reference Notes": "MAKE",
                    }
                )
            findings = core.compare_master_reference(path, "ROOT", [nx_row])
        self.assertTrue(findings)
        self.assertTrue(all(item["code"] == "DOWNSTREAM_BOM_DRIFT" for item in findings))


class J04Tests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.j04 = load_j04()
        cls.config = json.loads(
            (
                ROOT
                / "from_git"
                / "config"
                / "attribute_reconciliation.json"
            ).read_text(encoding="utf-8")
        )

    def test_unique_prototype_pull_and_missing_business_values_are_ready(self):
        root = FakePart(
            "ROOT",
            attrs("ROOT", "A", "ROOT MODEL")
            + [FakeAttribute("WAEItem", "Commodity_Code", "ROOT-CODE")],
        )
        child = FakePart(
            "CHILD",
            attrs("CHILD", "A", "CHILD MODEL")
            + [FakeAttribute("WAEItem", "WAE_VERSION", "1.0")],
        )
        root.ComponentAssembly.RootComponent = FakeComponent(
            "ROOT-COMP",
            root,
            [
                FakeComponent("CHILD-1", child),
                FakeComponent("CHILD-2", child),
            ],
        )

        records, diagnostics = self.j04.build_pull_records(
            root, self.config, "RUN1"
        )

        self.assertFalse(diagnostics)
        self.assertEqual(["ROOT", "CHILD"], [
            record["row"]["Item Number"] for record in records
        ])
        self.assertTrue(all(
            record["row"]["PULL_STATUS"] == "READY"
            for record in records
        ))
        self.assertEqual("1.0", records[1]["row"]["WAE_VERSION"])

    def test_identity_reads_do_not_depend_on_teamcenter_category(self):
        part = FakePart(
            "DIRECT",
            [FakeAttribute("WAEItem", "Commodity_Code", "CODE")],
        )
        direct_values = {
            "DB_PART_NO": "DIRECT-PN",
            "DB_PART_NAME": "DIRECT NAME",
            "DB_PART_REV": "B",
        }
        part.GetStringAttribute = direct_values.__getitem__

        records, _ = self.j04.build_pull_records(
            part, self.config, "RUN2"
        )

        self.assertEqual("DIRECT-PN", records[0]["row"]["Item Number"])
        self.assertEqual("DIRECT NAME", records[0]["row"]["Part Description"])
        self.assertEqual("B", records[0]["row"]["Item Rev"])
        self.assertEqual("READY", records[0]["row"]["PULL_STATUS"])

    def test_pull_includes_only_active_unsuppressed_occurrences(self):
        root = FakePart("ROOT", attrs("ROOT"))
        active = FakePart("ACTIVE", attrs("ACTIVE"))
        suppressed_only = FakePart("SUPPRESSED", attrs("SUPPRESSED"))
        suppressed_descendant = FakePart("DESCENDANT", attrs("DESCENDANT"))
        root.ComponentAssembly.RootComponent = FakeComponent(
            "ROOT-COMP",
            root,
            [
                FakeComponent("ACTIVE", active),
                FakeComponent("ACTIVE-SUPPRESSED-DUPLICATE", active, suppressed=True),
                FakeComponent("SUPPRESSED", suppressed_only, suppressed=True),
                FakeComponent(
                    "SUPPRESSED-SUBASSEMBLY",
                    FakePart("SUB", attrs("SUB")),
                    [FakeComponent("DESCENDANT", suppressed_descendant)],
                    suppressed=True,
                ),
            ],
        )

        records, diagnostics = self.j04.build_pull_records(
            root, self.config, "RUN3"
        )

        self.assertFalse(diagnostics)
        self.assertEqual(
            ["ROOT", "ACTIVE"],
            [record["row"]["Item Number"] for record in records],
        )

    def test_pull_excludes_occurrence_when_suppression_state_is_unreadable(self):
        class UnreadableSuppressionComponent(FakeComponent):
            @property
            def IsSuppressed(self):
                raise RuntimeError("suppression unavailable")

            @IsSuppressed.setter
            def IsSuppressed(self, value):
                pass

        root = FakePart("ROOT", attrs("ROOT"))
        uncertain = FakePart("UNCERTAIN", attrs("UNCERTAIN"))
        root.ComponentAssembly.RootComponent = FakeComponent(
            "ROOT-COMP",
            root,
            [UnreadableSuppressionComponent("UNCERTAIN", uncertain)],
        )

        records, diagnostics = self.j04.build_pull_records(
            root, self.config, "RUN4"
        )

        self.assertEqual(
            ["ROOT"],
            [record["row"]["Item Number"] for record in records],
        )
        self.assertEqual("SUPPRESSION_STATE_UNAVAILABLE", diagnostics[0]["code"])

    def test_pull_excludes_keyword_named_occurrences(self):
        root = FakePart("ROOT", attrs("ROOT"))
        csys = FakePart("CSYS", attrs("CSYS"))
        datum = FakePart("DATUM", attrs("DATUM"))
        skeleton = FakePart("SKELETON", attrs("SKELETON"))
        real = FakePart("REAL", attrs("REAL"))
        root.ComponentAssembly.RootComponent = FakeComponent(
            "ROOT-COMP",
            root,
            [
                FakeComponent("CSYS_ORIGIN", csys),
                FakeComponent("DATUM_PLANE_A", datum),
                FakeComponent("SKELETON_MASTER", skeleton),
                FakeComponent("REAL-1", real),
            ],
        )

        records, diagnostics = self.j04.build_pull_records(
            root, self.config, "RUN5"
        )

        self.assertFalse(diagnostics)
        self.assertEqual(
            ["ROOT", "REAL"],
            [record["row"]["Item Number"] for record in records],
        )

    def test_pull_excludes_reference_flagged_occurrences(self):
        root = FakePart("ROOT", attrs("ROOT"))
        empty_ref = FakePart("EMPTY_REF", attrs("EMPTY_REF"))
        yes_ref = FakePart("YES_REF", attrs("YES_REF"))
        plist = FakePart("PLIST", attrs("PLIST"))
        custom = FakePart("CUSTOM", attrs("CUSTOM"))
        real = FakePart("REAL", attrs("REAL"))
        root.ComponentAssembly.RootComponent = FakeComponent(
            "ROOT-COMP",
            root,
            [
                FakeComponent(
                    "EMPTY_REF",
                    empty_ref,
                    string_attributes={"REFERENCE_COMPONENT": ""},
                ),
                FakeComponent(
                    "YES_REF",
                    yes_ref,
                    string_attributes={"REFERENCE_COMPONENT": "YES"},
                ),
                FakeComponent(
                    "PLIST_IGNORED",
                    plist,
                    string_attributes={"PLIST_IGNORE_MEMBER": "YES"},
                ),
                FakeComponent(
                    "CUSTOM_BOM_EXCLUDED",
                    custom,
                    string_attributes={
                        "CELESTICA_BOM_EXCLUDE_SUBTREE": "YES"
                    },
                ),
                FakeComponent("REAL-1", real),
            ],
        )

        records, diagnostics = self.j04.build_pull_records(
            root, self.config, "RUN6"
        )

        self.assertFalse(diagnostics)
        self.assertEqual(
            ["ROOT", "REAL"],
            [record["row"]["Item Number"] for record in records],
        )

    def test_part_with_any_visible_occurrence_is_pulled_once(self):
        root = FakePart("ROOT", attrs("ROOT"))
        shared = FakePart("SHARED", attrs("SHARED"))
        root.ComponentAssembly.RootComponent = FakeComponent(
            "ROOT-COMP",
            root,
            [
                FakeComponent(
                    "SHARED-REF",
                    shared,
                    string_attributes={"PLIST_IGNORE_MEMBER": "YES"},
                ),
                FakeComponent("SHARED-NORMAL", shared),
            ],
        )

        records, _ = self.j04.build_pull_records(root, self.config, "RUN7")

        self.assertEqual(
            ["ROOT", "SHARED"],
            [record["row"]["Item Number"] for record in records],
        )


class J05Tests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.j05 = load_j05()
        cls.config = json.loads(
            (ROOT / "from_git" / "config" / "attribute_reconciliation.json").read_text(encoding="utf-8")
        )

    def target(self, value="OLD", **flags):
        part = FakePart(
            "P1",
            attrs("P1", "A", "PART ONE")
            + [
                FakeAttribute(
                    "WAEItem",
                    "Commodity_Code",
                    value,
                    **flags,
                )
            ],
        )
        part.IsReadOnly = False
        return part

    def baseline(self, value="OLD"):
        values = {}
        for column, _rule in self.j05._business_specs(self.config):
            values[column] = {
                "status": (
                    "POPULATED" if column == "Commodity_Code" else "MISSING"
                ),
                "raw_value": value if column == "Commodity_Code" else "",
            }
        return {
            "schema_version": 1,
            "audit_run_id": "RUN1",
            "identity_columns": self.j05._baseline_contract(
                self.config, "identity_columns"
            ),
            "business_columns": self.j05._baseline_contract(
                self.config, "business_columns"
            ),
            "parts": [
                {
                    "part_number": "P1",
                    "part_name": "PART ONE",
                    "revision": "A",
                    "business_values": values,
                }
            ],
        }

    def row(self, **updates):
        row = {
            column: ""
            for column in self.j05.update_columns(self.config)
        }
        row.update(
            {
                "AUDIT_RUN_ID": "RUN1",
                "APPROVED": "YES",
                "ENGINEER": "Engineer",
                "PULL_STATUS": "READY",
                "Item Number": "P1",
                "Part Description": "PART ONE",
                "Item Rev": "A",
                "Commodity_Code": "NEW",
                "_CSV_ROW": 2,
            }
        )
        row.update(updates)
        return row

    def prepare(self, row=None, target=None, baseline=None):
        target = target or self.target()
        reports, proposals = self.j05.prepare_updates(
            types.SimpleNamespace(IsManagedMode=True),
            target,
            self.config,
            [row or self.row()],
            baseline or self.baseline(),
            "timestamp",
            "DRY_RUN",
        )
        return reports, proposals

    def test_manual_settings_are_safe_and_environment_can_override(self):
        self.assertEqual("", self.j05.USER_UPDATE_CSV)
        self.assertEqual("DRY_RUN", self.j05.USER_MODE)
        with mock.patch.dict(
            self.j05.os.environ,
            {
                "NX_ATTRIBUTE_UPDATE_FILE": r"C:\temp\from-environment.csv",
                "NX_J05_MODE": "apply_approved",
            },
            clear=False,
        ), mock.patch.object(
            self.j05, "USER_UPDATE_CSV", r"C:\temp\from-code.csv"
        ), mock.patch.object(self.j05, "USER_MODE", "DRY_RUN"):
            self.assertEqual(
                r"C:\temp\from-environment.csv",
                self.j05.configured_input_path(),
            )
            self.assertEqual(
                "APPLY_APPROVED", self.j05.configured_mode()
            )

    def test_manual_settings_are_used_without_environment(self):
        with mock.patch.dict(
            self.j05.os.environ,
            {},
            clear=True,
        ), mock.patch.object(
            self.j05, "USER_UPDATE_CSV", r"C:\temp\from-code.csv"
        ), mock.patch.object(self.j05, "USER_MODE", "apply_approved"):
            self.assertEqual(
                r"C:\temp\from-code.csv",
                self.j05.configured_input_path(),
            )
            self.assertEqual(
                "APPLY_APPROVED", self.j05.configured_mode()
            )

    def test_business_allowlist_and_model_write_targets(self):
        columns = [
            column
            for column, _rule in self.j05._business_specs(self.config)
        ]
        self.assertIn("WAE_VERSION", columns)
        self.assertIn("NX_FINISH", columns)
        self.assertNotIn("NX_MATERIAL", columns)
        self.assertNotIn("NX_MASS", columns)
        finish = next(
            rule
            for rule in self.config["attributes"]
            if rule["logical_name"] == "finish"
        )
        self.assertEqual(["MODEL"], finish["write_targets"])

    def test_approved_change_unapproved_change_and_stale_value(self):
        reports, proposals = self.prepare()
        self.assertEqual(["PROPOSED_UPDATE"], [r["ACTION"] for r in reports])
        self.assertEqual(1, len(proposals))

        reports, proposals = self.prepare(self.row(APPROVED="NO"))
        self.assertEqual("SKIPPED_NOT_APPROVED", reports[0]["ACTION"])
        self.assertFalse(proposals)

        reports, proposals = self.prepare(target=self.target("CURRENT"))
        self.assertEqual("STALE_BASELINE_VALUE", reports[0]["ACTION"])
        self.assertFalse(proposals)

    def test_current_value_already_expected_is_idempotent_success(self):
        reports, proposals = self.prepare(target=self.target("NEW"))

        self.assertEqual(
            "ALREADY_AT_EXPECTED_VALUE", reports[0]["ACTION"]
        )
        self.assertEqual("ALREADY_MATCHED", reports[0]["VERIFICATION_RESULT"])
        self.assertEqual("NOT_REQUIRED", reports[0]["SAVE_RESULT"])
        self.assertFalse(proposals)
        self.assertFalse(self.j05._hard_preflight_error(reports[0]))

    def test_normalized_current_baseline_difference_is_not_stale(self):
        reports, proposals = self.prepare(target=self.target(" old "))

        self.assertEqual("PROPOSED_UPDATE", reports[0]["ACTION"])
        self.assertEqual(1, len(proposals))

    def test_buy_reference_compatibility_survives_older_config_copy(self):
        rule = {
            "logical_name": "stocking_type",
            "type": "String",
            "comparison": "TRIMMED_CASE_INSENSITIVE",
            "allowed_values": ["MAKE", "BUY", "REF"],
        }

        self.assertEqual(
            "", self.j05._validate_expected("BUY/REF", rule, self.config)
        )

    def test_commodity_type_accepts_value_outside_shared_controlled_vocabulary(self):
        baseline = self.baseline()
        baseline["parts"][0]["business_values"]["COMMODITYTYPE"].update(
            status="POPULATED", raw_value="Assembly"
        )
        row = self.row(
            Commodity_Code="OLD", COMMODITYTYPE="Future Commodity Group"
        )
        target = self.target()
        target.attributes.append(
            FakeAttribute("WAEItem", "COMMODITYTYPE", "Assembly")
        )

        reports, proposals = self.prepare(row, target, baseline)

        self.assertEqual(["PROPOSED_UPDATE"], [r["ACTION"] for r in reports])
        self.assertEqual(1, len(proposals))
        self.assertEqual("Future Commodity Group", proposals[0]["expected"])

    def test_commodity_type_keeps_tbc_but_rejects_blank_updates(self):
        rule = next(
            rule
            for rule in self.config["attributes"]
            if rule["logical_name"] == "commodity_type"
        )

        self.assertEqual(
            "", self.j05._validate_expected("TBC", rule, self.config)
        )
        self.assertEqual(
            "Populated-to-blank updates are not supported.",
            self.j05._validate_expected("   ", rule, self.config),
        )

        uom_rule = next(
            rule
            for rule in self.config["attributes"]
            if rule["logical_name"] == "uom"
        )
        self.assertEqual(
            "Expected value is outside the controlled value set.",
            self.j05._validate_expected("box", uom_rule, self.config),
        )

    def test_baseline_mapping_must_match_deployed_contract(self):
        baseline = self.baseline()
        baseline["business_columns"][0]["attribute"] = "WRONG"

        with self.assertRaisesRegex(RuntimeError, "business mapping"):
            self.prepare(baseline=baseline)

    def test_legacy_traceability_csv_and_baseline_remain_compatible(self):
        baseline = self.baseline()
        traceability = next(
            item
            for item in baseline["business_columns"]
            if item["logical_name"] == "traceability"
        )
        traceability["csv_column"] = "SERIAL_NUMBERED_PART"
        values = baseline["parts"][0]["business_values"]
        values["SERIAL_NUMBERED_PART"] = values.pop("Traceability")

        row = self.row()
        row["SERIAL_NUMBERED_PART"] = row.pop("Traceability")
        row["Traceability"] = row["SERIAL_NUMBERED_PART"]

        reports, proposals = self.prepare(row=row, baseline=baseline)

        self.assertEqual(["PROPOSED_UPDATE"], [r["ACTION"] for r in reports])
        self.assertEqual(1, len(proposals))

    def test_legacy_traceability_header_is_canonicalized_on_read(self):
        headers = self.j05.update_columns(self.config)
        headers[headers.index("Traceability")] = "SERIAL_NUMBERED_PART"
        values = {column: "" for column in headers}
        values.update(
            {
                "AUDIT_RUN_ID": "RUN1",
                "Item Number": "P1",
                "Item Rev": "A",
                "SERIAL_NUMBERED_PART": "BATCH",
            }
        )

        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "legacy.csv"
            with path.open("w", encoding="utf-8", newline="") as handle:
                writer = csv.DictWriter(handle, fieldnames=headers)
                writer.writeheader()
                writer.writerow(values)

            rows = self.j05._read_csv(str(path), self.config)

        self.assertEqual("BATCH", rows[0]["Traceability"])

    def test_duplicate_traceability_headers_are_rejected(self):
        headers = self.j05.update_columns(self.config)
        headers.append("SERIAL_NUMBERED_PART")

        with self.assertRaisesRegex(RuntimeError, "ambiguous columns"):
            self.j05._input_column_sources(headers, headers[:-1])

    def test_identity_blank_controlled_and_runtime_flags_fail_closed(self):
        reports, _ = self.prepare(
            self.row(**{"Part Description": "EDITED"})
        )
        self.assertEqual(
            "ERROR_PROTECTED_IDENTITY_EDIT", reports[0]["ACTION"]
        )

        reports, _ = self.prepare(self.row(Commodity_Code=""))
        self.assertEqual("ERROR_VALUE", reports[0]["ACTION"])

        base = self.baseline()
        base["parts"][0]["business_values"]["UOM"]["raw_value"] = "ea"
        row = self.row(Commodity_Code="OLD", UOM="TBC")
        target = self.target()
        target.attributes.append(
            FakeAttribute("WAEItem", "Unit_Of_Measure", "ea")
        )
        reports, _ = self.prepare(row, target, base)
        self.assertEqual("ERROR_VALUE", reports[0]["ACTION"])

        reports, _ = self.prepare(
            target=self.target(pdm_based=True)
        )
        self.assertEqual(
            "ERROR_ATTRIBUTE_NOT_WRITABLE", reports[0]["ACTION"]
        )

    def test_free_text_tbc_is_allowed(self):
        base = self.baseline()
        base["parts"][0]["business_values"]["Mfr. Name"][
            "raw_value"
        ] = "OLD MFG"
        row = self.row(Commodity_Code="OLD", **{"Mfr. Name": "TBC"})
        target = self.target()
        target.attributes.append(
            FakeAttribute("WAEItem", "MFG", "OLD MFG")
        )
        reports, proposals = self.prepare(row, target, base)
        self.assertEqual("PROPOSED_UPDATE", reports[0]["ACTION"])
        self.assertEqual(1, len(proposals))

    def test_no_save_gate_prevents_checkout_and_write(self):
        target = self.target()
        config = dict(self.config, save_policy="NO_SAVE")
        with mock.patch.object(
            self.j05, "checkout_targets"
        ) as checkout, mock.patch.object(
            self.j05, "apply_groups"
        ) as apply:
            reports, unsaved = self.j05.execute(
                types.SimpleNamespace(IsManagedMode=True),
                target,
                config,
                [self.row()],
                self.baseline(),
                "timestamp",
                "APPLY_APPROVED",
            )
        checkout.assert_not_called()
        apply.assert_not_called()
        self.assertFalse(unsaved)
        self.assertEqual("SAVE_GATE_DISABLED", reports[0]["ACTION"])

    def test_checkout_failure_aborts_before_attribute_write(self):
        target = self.target()
        config = dict(self.config, save_policy="SAVE_CHANGED_PARTS")
        failed = {
            "success": False,
            "before": "NOT_CHECKED_OUT",
            "action": "EXPLICIT_CHECKOUT",
            "result": "FAILED",
            "read_only_before": True,
            "read_only_after": True,
            "message": "checked out by another user",
            "exception_type": "NXException",
            "error_code": "123",
        }
        with mock.patch.object(
            self.j05,
            "checkout_targets",
            return_value={self.j05._object_key(target): failed},
        ), mock.patch.object(self.j05, "apply_groups") as apply:
            reports, unsaved = self.j05.execute(
                types.SimpleNamespace(IsManagedMode=True),
                target,
                config,
                [self.row()],
                self.baseline(),
                "timestamp",
                "APPLY_APPROVED",
            )
        apply.assert_not_called()
        self.assertFalse(unsaved)
        self.assertEqual("CHECKOUT_FAILED", reports[0]["ACTION"])

    def test_blocked_target_does_not_abort_independent_writable_target(self):
        writable = self.target()
        blocked = FakePart("P2", attrs("P2"))
        writable_proposal = self.proposal(writable)
        blocked_proposal = self.proposal(blocked)
        reports = [
            writable_proposal["report"],
            blocked_proposal["report"],
        ]
        checkout_results = {
            self.j05._object_key(writable): {
                "success": True,
                "message": "Already checked out and writable.",
            },
            self.j05._object_key(blocked): {
                "success": False,
                "message": "Checkout user: another.user",
            },
        }
        progress = []

        with mock.patch.object(
            self.j05,
            "prepare_updates",
            return_value=(
                reports,
                [writable_proposal, blocked_proposal],
            ),
        ), mock.patch.object(
            self.j05,
            "checkout_targets",
            return_value=checkout_results,
        ), mock.patch.object(
            self.j05, "apply_groups", return_value=set()
        ) as apply:
            returned_reports, unsaved = self.j05.execute(
                types.SimpleNamespace(IsManagedMode=True),
                writable,
                dict(self.config, save_policy="SAVE_CHANGED_PARTS"),
                [self.row()],
                self.baseline(),
                "timestamp",
                "APPLY_APPROVED",
                progress=progress.append,
            )

        apply.assert_called_once()
        self.assertEqual(
            [writable_proposal], apply.call_args.args[1]
        )
        self.assertFalse(unsaved)
        self.assertIs(reports, returned_reports)
        self.assertEqual(
            "PROPOSED_UPDATE", writable_proposal["report"]["ACTION"]
        )
        self.assertEqual(
            "CHECKOUT_FAILED", blocked_proposal["report"]["ACTION"]
        )
        self.assertIn(
            "another.user", blocked_proposal["report"]["MESSAGE"]
        )
        self.assertTrue(any(
            "1 writable target(s), 1 blocked target(s)" in item
            for item in progress
        ))

    def test_prechecked_targets_use_one_snapshot_and_no_checkout_call(self):
        first = self.target()
        second = FakePart("P2", attrs("P2"))
        second.IsReadOnly = False
        proposals = [
            self.proposal(first),
            self.proposal(first),
            self.proposal(second),
        ]
        checked = {
            self.j05._object_key(first),
            self.j05._object_key(second),
        }
        session = types.SimpleNamespace(IsManagedMode=True)

        with mock.patch.object(
            self.j05,
            "_session_checkout_snapshot",
            return_value=(checked, set(), ""),
        ) as snapshot, mock.patch.object(
            self.j05, "_batch_checkout"
        ) as batch:
            results = self.j05.checkout_targets(
                session, first, proposals
            )

        snapshot.assert_called_once_with(session)
        batch.assert_not_called()
        self.assertEqual(2, len(results))
        self.assertTrue(all(result["success"] for result in results.values()))
        self.assertTrue(all(
            result["result"] == "ALREADY_CHECKED_OUT"
            for result in results.values()
        ))

    def test_mixed_targets_use_one_batch_checkout_and_post_snapshot(self):
        prechecked = self.target()
        pending = FakePart("P2", attrs("P2"))
        pending.IsReadOnly = False
        proposals = [
            self.proposal(prechecked),
            self.proposal(pending),
            self.proposal(pending),
        ]
        prechecked_key = self.j05._object_key(prechecked)
        pending_key = self.j05._object_key(pending)
        session = types.SimpleNamespace(IsManagedMode=True)

        with mock.patch.object(
            self.j05,
            "_session_checkout_snapshot",
            side_effect=[
                ({prechecked_key}, {pending_key}, ""),
                ({prechecked_key, pending_key}, set(), ""),
            ],
        ) as snapshot, mock.patch.object(
            self.j05, "_batch_checkout"
        ) as batch:
            results = self.j05.checkout_targets(
                session, prechecked, proposals
            )

        self.assertEqual(2, snapshot.call_count)
        batch.assert_called_once_with([pending])
        self.assertEqual(
            "ALREADY_CHECKED_OUT", results[prechecked_key]["result"]
        )
        self.assertEqual("BATCH_CHECKOUT", results[pending_key]["result"])
        self.assertEqual(
            "BATCH_CHECKOUT", proposals[1]["report"]["CHECKOUT_ACTION"]
        )

    def test_batch_checkout_failure_and_unavailable_snapshot_fail_closed(self):
        target = self.target()
        target.IsReadOnly = True
        proposal = self.proposal(target)
        key = self.j05._object_key(target)
        session = types.SimpleNamespace(IsManagedMode=True)

        with mock.patch.object(
            self.j05,
            "_session_checkout_snapshot",
            side_effect=[(set(), {key}, ""), (set(), {key}, "")],
        ), mock.patch.object(
            self.j05,
            "_batch_checkout",
            side_effect=RuntimeError("batch checkout failed"),
        ):
            results = self.j05.checkout_targets(
                session, target, [proposal]
            )

        self.assertFalse(results[key]["success"])
        self.assertEqual("FAILED", results[key]["result"])
        self.assertIn("batch checkout failed", results[key]["message"])

        proposal = self.proposal(target)
        with mock.patch.object(
            self.j05,
            "_session_checkout_snapshot",
            return_value=(None, None, "API_UNAVAILABLE"),
        ), mock.patch.object(self.j05, "_batch_checkout") as batch:
            results = self.j05.checkout_targets(
                session, target, [proposal]
            )

        batch.assert_not_called()
        self.assertFalse(results[key]["success"])
        self.assertIn("API_UNAVAILABLE", results[key]["message"])

    def test_checked_out_but_read_only_target_is_blocked(self):
        target = self.target()
        target.IsReadOnly = True
        target.PDMPart = types.SimpleNamespace(
            GetCheckedoutStatusAndUser=lambda: (True, "another.user")
        )
        proposal = self.proposal(target)
        key = self.j05._object_key(target)

        with mock.patch.object(
            self.j05,
            "_session_checkout_snapshot",
            return_value=({key}, set(), ""),
        ), mock.patch.object(self.j05, "_batch_checkout") as batch:
            results = self.j05.checkout_targets(
                types.SimpleNamespace(IsManagedMode=True),
                target,
                [proposal],
            )

        batch.assert_not_called()
        self.assertFalse(results[key]["success"])
        self.assertIn("another.user", results[key]["message"])

    def proposal(self, target):
        rule = next(
            rule
            for rule in self.config["attributes"]
            if rule["logical_name"] == "commodity_code"
        )
        report = {
            column: "" for column in self.j05.REPORT_COLUMNS
        }
        report["ACTION"] = "PROPOSED_UPDATE"
        return {
            "source_row": self.row(),
            "report": report,
            "rule": rule,
            "target": target,
            "expected": "NEW",
        }

    def test_verification_failure_rolls_back_without_save(self):
        target = self.target()
        session = mock.Mock()
        session.SetUndoMark.return_value = 7
        proposal = self.proposal(target)
        with mock.patch.object(
            self.j05, "_write_attribute"
        ), mock.patch.object(
            self.j05,
            "_read_attribute",
            return_value={"raw": "WRONG"},
        ), mock.patch.object(self.j05, "_save_target") as save:
            unsaved = self.j05.apply_groups(
                session,
                [proposal],
                dict(self.config, save_policy="SAVE_CHANGED_PARTS"),
            )
        session.UndoToMark.assert_called_once()
        save.assert_not_called()
        self.assertNotIn(self.j05._object_key(target), unsaved)
        self.assertEqual(
            "UPDATED_VERIFICATION_FAILED",
            proposal["report"]["ACTION"],
        )

    def test_successful_apply_reports_per_target_progress(self):
        target = self.target()
        session = mock.Mock()
        session.SetUndoMark.return_value = 7
        proposal = self.proposal(target)
        progress = []
        with mock.patch.object(
            self.j05, "_write_attribute"
        ), mock.patch.object(
            self.j05,
            "_read_attribute",
            return_value={"raw": "NEW"},
        ), mock.patch.object(self.j05, "_save_target") as save:
            unsaved = self.j05.apply_groups(
                session,
                [proposal],
                dict(self.config, save_policy="SAVE_CHANGED_PARTS"),
                progress=progress.append,
            )

        save.assert_called_once_with(target)
        self.assertFalse(unsaved)
        self.assertTrue(any("Updating target 1/1" in item for item in progress))
        self.assertTrue(any("Verified and saved" in item for item in progress))

    def test_batch_checkout_uses_explicit_input_and_disposes_result(self):
        target = self.target()

        class OperationErrors:
            def __init__(inner_self):
                inner_self.disposed = False

            def Dispose(inner_self):
                inner_self.disposed = True

        operation_errors = OperationErrors()
        calls = []

        def checkout(parts, checkout_input):
            calls.append((parts, checkout_input))
            return operation_errors

        target.PDMPart = types.SimpleNamespace(CheckoutParts=checkout)

        self.j05._batch_checkout([target])

        self.assertEqual([target], calls[0][0])
        self.assertTrue(calls[0][1].ExplicitCheckOut)
        self.assertFalse(calls[0][1].IncludeSecondary)
        self.assertTrue(operation_errors.disposed)

    def test_db_identity_is_managed_when_session_flag_is_false(self):
        target = self.target()
        target.JournalIdentifier = "@DB/P1/A"

        self.assertTrue(
            self.j05._is_teamcenter_target(
                types.SimpleNamespace(IsManagedMode=False), target
            )
        )

    def test_sample_changes_are_business_writable_and_valid(self):
        expected = {
            "Unit_Of_Measure": "ea",
            "MFG": "FABRICATOR",
            "MPN": "264MN033036A01",
        }
        rules = {
            rule["attribute"]: rule
            for _column, rule in self.j05._business_specs(self.config)
        }
        for title, value in expected.items():
            self.assertIn(title, rules)
            self.assertTrue(rules[title]["writable"])
            self.assertEqual(
                "", self.j05._validate_expected(value, rules[title], self.config)
            )

    def test_j05_uses_session_snapshot_and_batch_checkout_only(self):
        source = (
            ROOT
            / "from_git"
            / "journals"
            / "05_bulk_attribute_updater.py"
        ).read_text(encoding="utf-8")

        self.assertIn("GetCheckedoutStatusOfAllObjectsInSession", source)
        self.assertIn("CheckoutParts", source)
        self.assertNotIn("GetCheckedoutStatusOfObjects", source)
        self.assertNotIn(".SaveAll(", source)
        self.assertNotIn(".Checkin", source)


class StaticSafetyTests(unittest.TestCase):
    def test_j04_has_no_mutation_or_save_calls(self):
        source = (ROOT / "from_git" / "journals" / "04_assembly_attribute_audit.py").read_text(
            encoding="utf-8"
        )
        for forbidden in ("SetUserAttribute", "CreateAttributePropertiesBuilder", ".Save("):
            self.assertNotIn(forbidden, source)
        self.assertNotIn("drawing_scope", source.lower())
        self.assertIn(".baseline.json", source)
        self.assertIn("collect_unique_prototypes", source)

    def test_j11_is_guarded_and_never_checks_in(self):
        source = (
            ROOT
            / "from_git"
            / "journals"
            / "11_test_teamcenter_attribute_checkout.py"
        ).read_text(encoding="utf-8")
        self.assertIn("NX_J11_ALLOW_MUTATION", source)
        self.assertIn("FULL_REVERSIBLE", source)
        self.assertIn("RESTORATION_REQUIRED", source)
        self.assertNotIn(".Checkin", source)

    def test_j07_step_fix_is_fail_closed(self):
        source = (
            ROOT / "from_git" / "journals"
            / "07_datapack_pdf_step_export.py"
        ).read_text(encoding="utf-8")
        self.assertIn('STEP_LAYER_MASK = "1-256"', source)
        self.assertIn("ExportFromOption.DisplayPart", source)
        self.assertIn("Scope.EntirePart", source)
        self.assertIn("FAILED_ZERO_GEOMETRY", source)


if __name__ == "__main__":
    unittest.main()
