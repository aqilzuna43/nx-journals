import csv
import importlib.util
import json
import subprocess
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
    ):
        self.Name = name
        self.DisplayName = name
        self.Prototype = prototype
        self._children = list(children)
        self.IsSuppressed = suppressed
        self.position = position
        self.orientation = orientation
        self.Tag = FakeComponent._tag
        FakeComponent._tag += 1

    def GetChildren(self):
        return list(self._children)

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
    nxopen.UF = nxopen_uf
    nxopen.Session = types.SimpleNamespace(
        MarkVisibility=types.SimpleNamespace(Invisible="Invisible")
    )
    prior_nx = sys.modules.get("NXOpen")
    prior_uf = sys.modules.get("NXOpen.UF")
    sys.modules["NXOpen"] = nxopen
    sys.modules["NXOpen.UF"] = nxopen_uf
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


class ConfigAndValuesTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.config_path = ROOT / "from_git" / "config" / "attribute_reconciliation.json"
        cls.config = json.loads(cls.config_path.read_text(encoding="utf-8"))

    def test_config_is_valid_nx_authoritative_and_no_save(self):
        self.assertIs(core.validate_config(self.config), self.config)
        self.assertEqual("NX_TEAMCENTER", self.config["authority"])
        self.assertEqual("NO_SAVE", self.config["save_policy"])

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


class J05Tests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.j05 = load_j05()
        cls.config = json.loads(
            (ROOT / "from_git" / "config" / "attribute_reconciliation.json").read_text(encoding="utf-8")
        )

    def row(self, **updates):
        row = {column: "" for column in self.j05.CORRECTION_COLUMNS}
        row.update(
            APPROVED="YES",
            PART_NUMBER="P1",
            REVISION="A",
            TARGET_OBJECT="MODEL",
            CATEGORY="WAEItem",
            NX_ATTRIBUTE_NAME="Commodity_Code",
            NX_ATTRIBUTE_TYPE="String",
            CURRENT_VALUE_FROM_AUDIT="OLD",
            EXPECTED_VALUE="NEW",
            AUTHORITATIVE_SOURCE="ENGINEERING_APPROVAL",
            AUDIT_RUN_ID="RUN1",
            ENGINEER="Engineer",
            EVIDENCE_REFERENCE="EVIDENCE-1",
        )
        row.update(updates)
        return row

    def validate(self, row, target=None):
        target = target or FakePart("P1", [FakeAttribute("WAEItem", "Commodity_Code", "OLD")])
        return self.j05._validate_row(
            None,
            "timestamp",
            "DRY_RUN",
            row,
            self.config,
            self.j05._rule_map(self.config),
            {("P1", "A"): [target]},
            {"P1": {"A"}},
            [],
        )

    def test_unapproved_master_prohibited_and_physical_fields_are_rejected(self):
        report, proposal = self.validate(self.row(APPROVED="NO"))
        self.assertEqual("SKIPPED_NOT_APPROVED", report["ACTION"])
        self.assertIsNone(proposal)
        report, _ = self.validate(self.row(AUTHORITATIVE_SOURCE="MASTER"))
        self.assertEqual("ERROR", report["ACTION"])
        report, _ = self.validate(
            self.row(CATEGORY="Materials", NX_ATTRIBUTE_NAME="NX_Mass", NX_ATTRIBUTE_TYPE="Number")
        )
        self.assertEqual("SKIPPED_NOT_WRITABLE", report["ACTION"])

    def test_stale_pdm_no_change_and_valid_proposal(self):
        report, _ = self.validate(self.row(CURRENT_VALUE_FROM_AUDIT="STALE"))
        self.assertEqual("STALE_AUDIT_VALUE", report["ACTION"])
        pdm_target = FakePart(
            "P1", [FakeAttribute("WAEItem", "Commodity_Code", "OLD", pdm_based=True)]
        )
        report, _ = self.validate(self.row(), pdm_target)
        self.assertEqual("SKIPPED_NOT_WRITABLE", report["ACTION"])
        report, proposal = self.validate(self.row(EXPECTED_VALUE="OLD"))
        self.assertEqual("NO_CHANGE_ALREADY_MATCHES", report["ACTION"])
        self.assertIsNone(proposal)
        report, proposal = self.validate(self.row())
        self.assertEqual("PROPOSED_UPDATE", report["ACTION"])
        self.assertIsNotNone(proposal)

    def test_exact_revision_and_placeholder_expected_value_are_rejected(self):
        report, _ = self.validate(self.row(REVISION="B"))
        self.assertEqual("SKIPPED_REVISION_MISMATCH", report["ACTION"])
        report, _ = self.validate(self.row(EXPECTED_VALUE="TBC"))
        self.assertEqual("ERROR", report["ACTION"])
        self.assertIn("placeholder", report["MESSAGE"])

    def proposal(self, target):
        rule = next(rule for rule in self.config["attributes"] if rule["logical_name"] == "commodity_code")
        report = self.j05._base_report("timestamp", "APPLY_APPROVED", self.row())
        return {
            "source_row": self.row(),
            "report": report,
            "rule": rule,
            "target": target,
            "newly_opened": True,
        }

    def test_verification_failure_rolls_back_without_save(self):
        target = FakePart("P1")
        session = mock.Mock()
        session.SetUndoMark.return_value = 7
        proposal = self.proposal(target)
        with mock.patch.object(self.j05, "_write_attribute"), mock.patch.object(
            self.j05, "_read_attribute", return_value={"raw": "WRONG"}
        ), mock.patch.object(self.j05, "_save_target") as save:
            unsaved = self.j05._apply_groups(session, [proposal], dict(self.config, save_policy="SAVE_CHANGED_PARTS"))
        session.UndoToMark.assert_called_once()
        save.assert_not_called()
        self.assertNotIn(self.j05._object_key(target), unsaved)
        self.assertEqual("UPDATED_VERIFICATION_FAILED", proposal["report"]["ACTION"])

    def test_save_failure_stops_later_saves_and_preserves_unsaved_keys(self):
        first = FakePart("P1")
        second = FakePart("P2")
        first_proposal = self.proposal(first)
        second_proposal = self.proposal(second)
        second_proposal["source_row"] = self.row(PART_NUMBER="P2")
        session = mock.Mock()
        session.SetUndoMark.side_effect = [1, 2]
        with mock.patch.object(self.j05, "_write_attribute"), mock.patch.object(
            self.j05, "_read_attribute", side_effect=lambda target, rule: {"raw": "NEW"}
        ), mock.patch.object(self.j05, "_save_target", side_effect=RuntimeError("save failed")) as save:
            unsaved = self.j05._apply_groups(
                session, [first_proposal, second_proposal], dict(self.config, save_policy="SAVE_CHANGED_PARTS")
            )
        self.assertEqual(1, save.call_count)
        self.assertEqual(
            {self.j05._object_key(first), self.j05._object_key(second)}, unsaved
        )
        self.assertEqual("SAVE_FAILED_PART_LEFT_MODIFIED", first_proposal["report"]["SAVE_RESULT"])
        self.assertEqual("NOT_ATTEMPTED", second_proposal["report"]["SAVE_RESULT"])


class StaticSafetyTests(unittest.TestCase):
    def test_j04_has_no_mutation_or_save_calls(self):
        source = (ROOT / "from_git" / "journals" / "04_assembly_attribute_audit.py").read_text(
            encoding="utf-8"
        )
        for forbidden in ("SetUserAttribute", "CreateAttributePropertiesBuilder", ".Save("):
            self.assertNotIn(forbidden, source)
        self.assertIn("if blocker_count == 0:", source)
        self.assertIn("NX_CERTIFIED_BOM_", source)

    def test_j07_matches_protected_commit(self):
        relative = "from_git/journals/07_datapack_pdf_step_export.py"
        expected = subprocess.check_output(
            ["git", "show", "6fe58f765b8a207229b2f6990e3b0224caa03771:" + relative],
            cwd=ROOT,
        )
        actual = (ROOT / relative).read_bytes()
        # The Windows checkout uses CRLF while git-show returns canonical LF.
        self.assertEqual(expected.replace(b"\r\n", b"\n"), actual.replace(b"\r\n", b"\n"))


if __name__ == "__main__":
    unittest.main()
