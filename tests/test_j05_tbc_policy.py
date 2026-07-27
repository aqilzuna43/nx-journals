import json
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
CONFIG_PATH = ROOT / "from_git" / "config" / "attribute_reconciliation.json"


def _rules():
    config = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    return {
        rule["logical_name"]: rule
        for rule in config.get("attributes", [])
    }


def test_tbc_is_allowed_for_intentionally_deferred_business_fields():
    rules = _rules()
    for logical_name in (
        "temperature_sensitive",
        "serviceable_item",
        "commodity_type",
    ):
        assert "TBC" in rules[logical_name]["allowed_values"]


def test_tbc_does_not_weaken_other_controlled_fields():
    rules = _rules()
    for logical_name in (
        "uom",
        "component_class",
        "traceability",
        "hazardous",
        "shelf_life_limited",
        "stocking_type",
    ):
        assert "TBC" not in rules[logical_name].get("allowed_values", [])
