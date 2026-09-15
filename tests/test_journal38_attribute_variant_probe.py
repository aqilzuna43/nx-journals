import csv
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
    / "38_attribute_variant_probe.py"
)


def load_journal():
    nxopen = types.ModuleType("NXOpen")
    nxopen.NXObject = types.SimpleNamespace(
        AttributeType=types.SimpleNamespace(String="String")
    )
    nxopen.Session = types.SimpleNamespace()
    prior = sys.modules.get("NXOpen")
    sys.modules["NXOpen"] = nxopen
    try:
        spec = importlib.util.spec_from_file_location("journal38", JOURNAL)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        return module
    finally:
        if prior is None:
            sys.modules.pop("NXOpen", None)
        else:
            sys.modules["NXOpen"] = prior


class FakeAttributeType:
    def __init__(self, name):
        self.name = name


class FakeAttributeInfo:
    def __init__(
        self,
        category,
        title,
        value,
        kind="String",
        unset=False,
        locked=False,
        owned_by_system=False,
        pdm_based=False,
        required=False,
        not_saved=False,
    ):
        self.Category = category
        self.Title = title
        self.Type = FakeAttributeType(kind)
        self.StringValue = value if kind == "String" else ""
        self.RealValue = value if kind in ("Real", "Number") else None
        self.IntegerValue = None
        self.BooleanValue = None
        self.Unset = unset
        self.Locked = locked
        self.OwnedBySystem = owned_by_system
        self.PdmBased = pdm_based
        self.Required = required
        self.NotSaved = not_saved


class FakeIterator:
    def __init__(self):
        self.include_unset = None

    def SetIncludeAlsoUnset(self, value):
        self.include_unset = value

    def FreeResource(self):
        self.freed = True


class FakeComponent:
    def __init__(self, prototype, children=None, suppressed=False):
        self._prototype = prototype
        self._children = list(children or [])
        self.IsSuppressed = suppressed

    def GetChildren(self):
        return list(self._children)

    @property
    def Prototype(self):
        return self._prototype


class FakeComponentAssembly:
    def __init__(self, root):
        self.RootComponent = root


class FakePart:
    counter = 100

    def __init__(self, name="part", attributes=None, string_attrs=None, path=""):
        FakePart.counter += 1
        self.Tag = FakePart.counter
        self.Name = name
        self.Leaf = name + ".prt"
        self.FullPath = path or name + ".prt"
        self.attribute_infos = list(attributes or [])
        self.string_attrs = dict(string_attrs or {})
        self.iterator = FakeIterator()

    def GetStringAttribute(self, name):
        if name in self.string_attrs:
            return self.string_attrs[name]
        raise RuntimeError("missing attribute " + name)

    def CreateAttributeIterator(self):
        return self.iterator

    def GetUserAttributes(self, iterator=None):
        return list(self.attribute_infos)


def make_part(name, number, infos, path="", string_attrs=None):
    string_attrs = dict(string_attrs or {})
    string_attrs.setdefault("DB_PART_NO", number)
    return FakePart(name, infos, string_attrs, path=path)


def clean_part(number="PN1"):
    """Part carrying only canonical titles under WAEItem."""
    infos = [
        FakeAttributeInfo("WAEItem", "Commodity_Code", "C123"),
        FakeAttributeInfo("WAEItem", "Unit_Of_Measure", "ea"),
        FakeAttributeInfo("Cad0Design", "DB_PART_NO", number),
    ]
    return make_part("clean", number, infos)


def fake_session(work_part):
    listing = types.SimpleNamespace(
        Open=lambda: None,
        WriteFullline=lambda line: None,
    )
    session = types.SimpleNamespace(
        Parts=types.SimpleNamespace(Work=work_part),
        ListingWindow=listing,
        IsManagedMode=False,
    )
    return session


class AttributeVariantTests(unittest.TestCase):
    def setUp(self):
        self.journal = load_journal()

    def test_normalize_title_key_collapses_alias_and_title(self):
        self.assertEqual(
            self.journal.normalize_title_key("Commodity_Code"),
            self.journal.normalize_title_key("Commodity Code"),
        )
        self.assertNotEqual(
            self.journal.normalize_title_key("MFG"),
            self.journal.normalize_title_key("Mfr. Name"),
        )

    def test_values_equal(self):
        self.assertTrue(self.journal.values_equal("EA", "ea"))
        self.assertTrue(self.journal.values_equal("1.0", "1"))
        self.assertFalse(self.journal.values_equal("X", "Y"))

    def test_dump_attributes_captures_flags_via_iterator(self):
        part = FakePart(
            "flagged",
            [
                FakeAttributeInfo(
                    "WAEItem",
                    "Commodity_Code",
                    "C123",
                    pdm_based=True,
                    owned_by_system=True,
                    locked=True,
                )
            ],
        )
        infos = self.journal.dump_attributes(part)
        self.assertEqual(len(infos), 1)
        info = infos[0]
        self.assertTrue(info["pdm_based"])
        self.assertTrue(info["owned_by_system"])
        self.assertTrue(info["locked"])
        self.assertEqual(info["category"], "WAEItem")

    def test_dump_attributes_plain_fallback(self):
        class PlainPart(FakePart):
            CreateAttributeIterator = None

        plain = PlainPart("plain", [FakeAttributeInfo("WAEItem", "MFG", "ACME")])
        infos = self.journal.dump_attributes(plain)
        self.assertEqual(len(infos), 1)
        self.assertEqual(infos[0]["title"], "MFG")

    def test_clean_part_is_clean(self):
        report = self.journal.analyze_part(clean_part())
        self.assertEqual(report["verdict"], "NX_FILE_CLEAN")
        commodity = next(
            field
            for field in report["fields"]
            if field["canonical_title"] == "Commodity_Code"
        )
        self.assertEqual(commodity["status"], "TITLE_ONLY")

    def test_both_variants_identical(self):
        part = clean_part()
        part.attribute_infos.append(
            FakeAttributeInfo("WAEItem", "Commodity Code", "C123")
        )
        report = self.journal.analyze_part(part)
        self.assertEqual(report["verdict"], "NX_FILE_CARRIES_DUPLICATES")
        self.assertEqual(report["identical_duplicate_fields"], ["Commodity_Code"])
        commodity = next(
            field
            for field in report["fields"]
            if field["canonical_title"] == "Commodity_Code"
        )
        self.assertEqual(commodity["status"], "BOTH_VARIANTS")
        self.assertEqual(commodity["relation"], "IDENTICAL_DUPLICATE")

    def test_alias_only_is_flagged(self):
        part = clean_part()
        part.attribute_infos.append(
            FakeAttributeInfo("WAEItem", "Mfr. Name", "ACME")
        )
        report = self.journal.analyze_part(part)
        self.assertEqual(report["verdict"], "NX_FILE_CARRIES_DUPLICATES")
        mfg = next(
            field for field in report["fields"] if field["canonical_title"] == "MFG"
        )
        self.assertEqual(mfg["status"], "ALIAS_ONLY")
        self.assertEqual(report["conflicting_duplicate_fields"], [])
        self.assertEqual(report["identical_duplicate_fields"], [])

    def test_conflicting_variants(self):
        part = clean_part()
        part.attribute_infos.append(
            FakeAttributeInfo("WAEItem", "Commodity Code", "DIFFERENT")
        )
        report = self.journal.analyze_part(part)
        self.assertEqual(report["conflicting_duplicate_fields"], ["Commodity_Code"])

    def test_generic_variant_group_discovery(self):
        part = clean_part()
        part.attribute_infos.append(
            FakeAttributeInfo("WAEItem", "Vendor", "ACME")
        )
        part.attribute_infos.append(
            FakeAttributeInfo("WAEItem", "VENDOR", "ACME")
        )
        report = self.journal.analyze_part(part)
        keys = [group["normalized_key"] for group in report["variant_groups"]]
        self.assertIn("vendor", keys)
        group = report["variant_groups"][keys.index("vendor")]
        self.assertEqual(group["relation"], "IDENTICAL_DUPLICATE")
        # A lowercase spelling of a canonical title is treated as the title
        # variant, not a separate unknown group.
        part2 = clean_part()
        part2.attribute_infos.append(
            FakeAttributeInfo("WAEItem", "commodity_code", "C123")
        )
        report2 = self.journal.analyze_part(part2)
        self.assertEqual(report2["variant_groups"], [])
        commodity = next(
            field
            for field in report2["fields"]
            if field["canonical_title"] == "Commodity_Code"
        )
        self.assertEqual(commodity["status"], "TITLE_ONLY")

    def test_unique_prototypes_dedupe_and_skip_suppressed(self):
        work = clean_part("ROOT")
        sub = clean_part("SUB")
        leaf = clean_part("LEAF")
        sub_occ_a = FakeComponent(sub)
        sub_occ_b = FakeComponent(sub)
        suppressed = FakeComponent(leaf, suppressed=True)
        sub_root = FakeComponent(sub, children=[sub_occ_a, sub_occ_b])
        leaf_occ = FakeComponent(leaf)
        root = FakeComponent(work, children=[sub_root, suppressed, leaf_occ])
        work.ComponentAssembly = FakeComponentAssembly(root)
        parts = self.journal.unique_prototypes(work)
        self.assertEqual(len(parts), 3)
        self.assertIs(parts[0], work)
        self.assertIn(sub, parts)
        self.assertIn(leaf, parts)

    def test_run_writes_json_and_csv(self):
        part = clean_part()
        part.attribute_infos.append(
            FakeAttributeInfo("WAEItem", "Commodity Code", "C123")
        )
        session = fake_session(part)
        with tempfile.TemporaryDirectory() as tmp:
            csv_path, json_path, report = self.journal.run(
                session, io_dir=tmp, run_datetime=self.journal.datetime.datetime(2026, 1, 2, 3, 4, 5)
            )
            self.assertTrue(os.path.isfile(json_path))
            self.assertTrue(os.path.isfile(csv_path))
            with open(json_path, "r", encoding="utf-8") as handle:
                loaded = json.load(handle)
            self.assertEqual(loaded["build"], self.journal.BUILD)
            self.assertEqual(loaded["parts_carrying_duplicates"], 1)
            self.assertIn("interpretation", loaded)
            with open(csv_path, "r", encoding="utf-8-sig", newline="") as handle:
                rows = list(csv.reader(handle))
            self.assertEqual(
                rows[0],
                [
                    "PART_NUMBER",
                    "PART_REV",
                    "PART_NAME",
                    "CATEGORY",
                    "TITLE",
                    "TYPE",
                    "VALUE",
                    "UNSET",
                    "LOCKED",
                    "OWNED_BY_SYSTEM",
                    "PDM_BASED",
                    "REQUIRED",
                    "NOT_SAVED",
                    "KNOWN_FIELD",
                    "VARIANT_ROLE",
                    "PART_VERDICT",
                ],
            )
            titles = {row[4] for row in rows[1:]}
            self.assertIn("Commodity_Code", titles)
            self.assertIn("Commodity Code", titles)
            alias_row = next(row for row in rows[1:] if row[4] == "Commodity Code")
            self.assertEqual(alias_row[13], "Commodity_Code")
            self.assertEqual(alias_row[14], "ALIAS")
            title_row = next(row for row in rows[1:] if row[4] == "Commodity_Code")
            self.assertEqual(title_row[14], "TITLE")

    def test_run_requires_work_part(self):
        session = fake_session(None)
        with tempfile.TemporaryDirectory() as tmp:
            with self.assertRaises(RuntimeError):
                self.journal.run(session, io_dir=tmp)

    def test_run_clean_part_reports_clean(self):
        session = fake_session(clean_part())
        with tempfile.TemporaryDirectory() as tmp:
            _, _, report = self.journal.run(session, io_dir=tmp)
            self.assertEqual(report["parts_clean"], 1)
            self.assertEqual(report["parts_carrying_duplicates"], 0)

    def test_part_identity_reads_db_attributes(self):
        part = clean_part("PN42")
        identity = self.journal.part_identity(part)
        self.assertEqual(identity["number"], "PN42")

    def test_pdm_path_detected(self):
        part = clean_part("PN1")
        part.FullPath = "@DB/PN1/A;1/part.prt"
        report = self.journal.analyze_part(part)
        self.assertTrue(report["part"]["pdm_part"])


if __name__ == "__main__":
    unittest.main()
