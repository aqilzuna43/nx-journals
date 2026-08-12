import importlib.util
import csv
import io
import sys
import tempfile
import types
import unittest
import xml.etree.ElementTree as ElementTree
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
EXTENDED_BOM_PATH = ROOT / "from_git" / "journals" / "NXOpenBoMExtended.py"
JOURNAL_04_PATH = ROOT / "from_git" / "journals" / "04_assembly_attribute_audit.py"
FZ_TEMPLATE_PATH = ROOT / "docs" / "FZ-PowerSystem_v1_22Jun.csv"
ATTRIBUTE_XML_PATH = ROOT / "tests" / "NXPartAttribute_FZ.xml"
EXPECTED_EXTENDED_COLUMNS = [
    "Level",
    "Item Number",
    "Part Description",
    "Item Rev",
    "Lifecycle",
    "Qty",
    "UOM",
    "Mfr. Name",
    "Mfr. Part Number",
    "Reference Notes",
    "WAE_VERSION",
    "NX_MATERIAL",
    "NX_FINISH",
    "NX_MASS",
    "NX_MassPropRollupMass",
    "NX_MassPropRollupArea_m2",
    "COMPONENT_CLASS",
    "LIFED",
    "SERIAL_NUMBERED_PART",
    "Temperature_Sensitive",
    "Hazardous",
    "Dimensions",
    "COMMODITYTYPE",
    "Commodity_Code",
    "Serviceable_item_flag",
    "Export_Control_Number",
    "Country_of_Origin",
]


def load_with_fake_nxopen(module_name, path):
    nxopen = types.ModuleType("NXOpen")
    nxopen.__path__ = []
    nxopen_uf = types.ModuleType("NXOpen.UF")
    nxopen.UF = nxopen_uf

    prior_nx = sys.modules.get("NXOpen")
    prior_uf = sys.modules.get("NXOpen.UF")
    sys.modules["NXOpen"] = nxopen
    sys.modules["NXOpen.UF"] = nxopen_uf
    try:
        spec = importlib.util.spec_from_file_location(module_name, path)
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


class FakeNxObject:
    def __init__(
        self,
        strings=None,
        numbers=None,
        prototype=None,
        name="COMPONENT",
        display_name="DISPLAY",
        children=None,
    ):
        self.strings = dict(strings or {})
        self.numbers = dict(numbers or {})
        self.Prototype = prototype
        self.Name = name
        self.DisplayName = display_name
        self.IsSuppressed = False
        self.children = list(children or [])

    def GetStringAttribute(self, title):
        if title not in self.strings:
            raise KeyError(title)
        return self.strings[title]

    def GetRealAttribute(self, title):
        if title not in self.numbers:
            raise KeyError(title)
        return self.numbers[title]

    def GetChildren(self):
        return list(self.children)


class ExtendedBomAttributeTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.module = load_with_fake_nxopen("extended_bom_under_test", EXTENDED_BOM_PATH)

    def test_extended_header_matches_feedback_and_retains_fz_prefix(self):
        self.assertEqual(EXPECTED_EXTENDED_COLUMNS, self.module.FZ_COLUMNS)

        if FZ_TEMPLATE_PATH.exists():
            with FZ_TEMPLATE_PATH.open(encoding="utf-8-sig", newline="") as handle:
                template_columns = next(csv.reader(handle))
            self.assertEqual(template_columns, self.module.FZ_COLUMNS[:len(template_columns)])
        self.assertEqual(
            [
                ("Part Description", "DB_PART_NAME", "String"),
                ("Item Rev", "DB_PART_REV", "String"),
                ("Lifecycle", "ItemRev_REL_STATUS", "String"),
                ("UOM", "Unit_Of_Measure", "String"),
                ("Mfr. Name", "MFG", "String"),
                ("Mfr. Part Number", "MPN", "String"),
                ("Reference Notes", "Stocking_Type", "String"),
                ("WAE_VERSION", "WAE_VERSION", "String"),
                ("NX_MATERIAL", "NX_MATERIAL", "String"),
                ("NX_FINISH", "NX_FINISH", "String"),
                ("NX_MASS", "NX_Mass", "Number"),
                ("NX_MassPropRollupMass", "NX_MassPropRollupMass", "Number"),
                ("NX_MassPropRollupArea", "NX_MassPropRollupArea", "Number"),
                ("COMPONENT_CLASS", "COMPONENT_CLASS", "String"),
                ("LIFED", "LIFED", "String"),
                ("SERIAL_NUMBERED_PART", "SERIAL_NUMBERED_PART", "String"),
                ("Temperature_Sensitive", "Temperature_Sensitive", "String"),
                ("Hazardous", "WAE_Hazardous", "String"),
                ("Dimensions", "Dimensions", "String"),
                ("COMMODITYTYPE", "COMMODITYTYPE", "String"),
                ("Commodity_Code", "Commodity_Code", "String"),
                ("Serviceable_item_flag", "Serviceable_item_flag", "String"),
                ("Export_Control_Number", "Export_Control_Number", "String"),
                ("Country_of_Origin", "Country_of_Origin", "String"),
            ],
            self.module.FZ_ATTRIBUTE_SPECS,
        )

    def test_xml_backed_columns_use_exact_internal_titles_and_types(self):
        templates = ElementTree.parse(ATTRIBUTE_XML_PATH).getroot().findall("Template")
        xml_types = {
            template.attrib["title"]: template.attrib["type"]
            for template in templates
        }
        configured_types = {
            attribute_title: attribute_type
            for _column, attribute_title, attribute_type in self.module.FZ_ATTRIBUTE_SPECS
            if attribute_title in xml_types
        }

        for attribute_title, attribute_type in configured_types.items():
            self.assertEqual(xml_types[attribute_title], attribute_type)
        self.assertEqual("String", configured_types["WAE_VERSION"])
        self.assertIn(("Hazardous", "WAE_Hazardous", "String"), self.module.FZ_ATTRIBUTE_SPECS)

    def test_j04_j05_business_mapping_matches_extended_bom(self):
        import json

        config = json.loads(
            (
                ROOT
                / "from_git"
                / "config"
                / "attribute_reconciliation.json"
            ).read_text(encoding="utf-8")
        )
        rules = {
            rule["logical_name"]: rule
            for rule in config["attributes"]
        }
        extended = {
            column: (title, attribute_type)
            for column, title, attribute_type
            in self.module.FZ_ATTRIBUTE_SPECS
        }
        xml_templates = {
            template.attrib["title"]: template.attrib
            for template in ElementTree.parse(
                ATTRIBUTE_XML_PATH
            ).getroot().findall("Template")
        }

        for mapping in config["update_workflow"]["business_columns"]:
            rule = rules[mapping["logical_name"]]
            self.assertEqual(
                (rule["attribute"], rule["type"]),
                extended[mapping["csv_column"]],
            )
            template = xml_templates[rule["attribute"]]
            self.assertEqual("WAEItem", template["category"])
            self.assertEqual("false", template["ownedBySystem"])
            self.assertEqual("false", template["pdmBasedPartAttribute"])

    def test_typed_reads_preserve_numeric_zero_and_missing_values(self):
        nx_object = FakeNxObject(
            strings={"NX_MATERIAL": "Copper"},
            numbers={"NX_Mass": 0.0},
        )

        self.assertEqual(
            "Copper",
            self.module.get_safe_attribute(nx_object, "NX_MATERIAL", "String"),
        )
        self.assertEqual(
            0.0,
            self.module.get_safe_attribute(nx_object, "NX_Mass", "Number"),
        )
        self.assertIsNone(
            self.module.get_safe_attribute(nx_object, "WAE_VERSION", "String")
        )
        with self.assertRaises(ValueError):
            self.module.get_safe_attribute(nx_object, "NX_Mass", "Unsupported")

    def test_component_values_always_come_from_prototype(self):
        prototype = FakeNxObject(
            strings={"NX_MATERIAL": "Prototype material"},
            numbers={"NX_Mass": 4.25},
        )
        component = FakeNxObject(
            strings={"NX_MATERIAL": "Occurrence material"},
            prototype=prototype,
        )

        self.assertEqual(
            "Prototype material",
            self.module.get_component_attribute(component, "NX_MATERIAL", "String"),
        )
        self.assertEqual(
            4.25,
            self.module.get_component_attribute(component, "NX_Mass", "Number"),
        )
        self.assertIsNone(
            self.module.get_component_attribute(component, "MISSING", "String")
        )

    def test_exported_row_matches_fz_template_projection(self):
        component = FakeNxObject(
            strings={
                "DB_PART_NO": "264MN180801A01",
                "DB_PART_NAME": "ASSY-GENX",
                "DB_PART_REV": "A",
                "ItemRev_REL_STATUS": "RELEASED",
                "Unit_Of_Measure": "ea",
                "MFG": "CELESTICA",
                "MPN": "264MN180801A01",
                "Stocking_Type": "MAKE",
                "WAE_VERSION": "22.1",
                "NX_MATERIAL": "Copper",
                "NX_FINISH": "TIN",
                "COMPONENT_CLASS": "A",
                "LIFED": "N",
                "SERIAL_NUMBERED_PART": "SERIAL",
                "Temperature_Sensitive": "N",
                "WAE_Hazardous": "Y",
                "Dimensions": "10 x 20 x 30",
                "COMMODITYTYPE": "Assembly",
                "Commodity_Code": "123",
                "Serviceable_item_flag": "Y",
                "Export_Control_Number": "EAR99",
                "Country_of_Origin": "MY",
            },
            numbers={
                "NX_Mass": 1.25,
                "NX_MassPropRollupMass": 2.5,
                "NX_MassPropRollupArea": 12500000.0,
            },
            name="ROOT",
            display_name="ROOT DISPLAY",
        )
        output = io.StringIO()
        writer = csv.writer(output, lineterminator="\n")

        self.module.walk_assembly_tree(component, 0, writer, quantity=1)

        self.assertEqual(
            [[
                "0",
                "264MN180801A01",
                "ASSY-GENX",
                "A",
                "RELEASED",
                "1",
                "ea",
                "CELESTICA",
                "264MN180801A01",
                "MAKE",
                "22.1",
                "Copper",
                "TIN",
                "1.25",
                "2.5",
                "12.5",
                "A",
                "N",
                "SERIAL",
                "N",
                "Y",
                "10 x 20 x 30",
                "Assembly",
                "123",
                "Y",
                "EAR99",
                "MY",
            ]],
            list(csv.reader(io.StringIO(output.getvalue()))),
        )

    def test_unset_lifecycle_defaults_to_draft(self):
        component = FakeNxObject(
            strings={
                "DB_PART_NO": "ITEM",
                "DB_PART_NAME": "ITEM NAME",
                "DB_PART_REV": "A",
            }
        )
        output = io.StringIO()

        self.module.walk_assembly_tree(component, 0, csv.writer(output), quantity=1)

        row = next(csv.reader(io.StringIO(output.getvalue())))
        self.assertEqual("DRAFT", row[4])


class Journal04ModelPullTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.journal = load_with_fake_nxopen("journal_04_under_test", JOURNAL_04_PATH)

    def test_update_columns_are_business_only_and_bom_aligned(self):
        import json

        config = json.loads(
            (
                ROOT
                / "from_git"
                / "config"
                / "attribute_reconciliation.json"
            ).read_text(encoding="utf-8")
        )
        columns = self.journal.update_columns(config)

        self.assertIn("Item Number", columns)
        self.assertIn("Part Description", columns)
        self.assertIn("Item Rev", columns)
        self.assertIn("WAE_VERSION", columns)
        self.assertIn("Hazardous", columns)
        self.assertNotIn("Qty", columns)
        self.assertNotIn("NX_MASS", columns)
        self.assertNotIn("NX_MATERIAL", columns)

    def test_journal04_has_no_shared_utils_dependency(self):
        source = JOURNAL_04_PATH.read_text(encoding="utf-8")

        self.assertNotIn("from utils", source)
        self.assertNotIn("drawing_scope", source.lower())
        self.assertIn(".baseline.json", source)


if __name__ == "__main__":
    unittest.main()
