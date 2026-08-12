"""Journal 04 - pull editable business attributes from 3D master models.

The journal is deliberately read-only.  It traverses the active assembly and
writes one wide CSV row per 3D master model that would actually show in the
BoM export (same visibility filter as NXOpenBoMExtended.py): suppressed,
reference-only, and keyword-named occurrences (CSYS, datum, skeleton, ...)
are excluded together with their subtrees.  The CSV can be edited and passed
directly to Journal 05.  A JSON sidecar keeps the exact typed baseline
required for stale-value protection.
"""

import csv
import json
import os
import re
import traceback
from collections import OrderedDict
from datetime import datetime

import NXOpen


CONTROL_COLUMNS = [
    "AUDIT_RUN_ID",
    "APPROVED",
    "ENGINEER",
    "APPROVAL_NOTE",
    "PULL_STATUS",
    "PULL_MESSAGE",
]

# --- BOM VISIBILITY (mirrors NXOpenBoMExtended.py) ---
# Only components that would show in the BoM export are pulled.  Everything
# else (CSYS, datums, skeletons, reference-only members, ...) is noise for the
# update CSV and must not create rows here.
IGNORE_KEYWORDS = ["CSYS", "COORDINATE", "DATUM", "REFERENCE", "SKELETON"]
BOM_REFERENCE_ATTRIBUTES = ("REFERENCE_COMPONENT", "PLIST_IGNORE_MEMBER")
# NX marks native reference components with an empty string; manual overrides
# use YES/1/True/true/yes.
BOM_REFERENCE_FLAG_VALUES = ("", "YES", "1", "True", "true", "yes")


def _text(value):
    return "" if value is None else str(value)


def _clean(value):
    return _text(value).strip()


def _normalized(value):
    return " ".join(_clean(value).split()).upper()


def _enum_name(value):
    if value is None:
        return ""
    name = getattr(value, "name", None)
    return _text(name if name is not None else value).split(".")[-1]


def _dispose(value):
    if value is None:
        return
    for name in ("Dispose", "FreeResource"):
        method = getattr(value, name, None)
        if callable(method):
            try:
                method()
            except Exception:
                pass
            return


def _exception_details(error):
    code = getattr(error, "ErrorCode", "")
    return "{0}{1}".format(
        type(error).__name__,
        ": {0}".format(code) if code not in ("", None) else "",
    ) + " - " + _text(error)


def _journal_path():
    return os.path.abspath(__file__)


def _runtime_candidates():
    script = _journal_path()
    script_parent = os.path.dirname(script)
    configured = _clean(os.environ.get("NX_JOURNALS_ROOT"))
    candidates = [
        configured,
        os.path.join(configured, "from_git") if configured else "",
        os.path.dirname(script_parent),
        os.getcwd(),
        os.path.join(os.getcwd(), "from_git"),
    ]
    result = []
    for candidate in candidates:
        if not candidate:
            continue
        absolute = os.path.abspath(candidate)
        if absolute not in result:
            result.append(absolute)
    return result


def _runtime_root():
    attempted = []
    for candidate in _runtime_candidates():
        attempted.append(candidate)
        if os.path.isfile(
            os.path.join(candidate, "config", "attribute_reconciliation.json")
        ):
            return candidate
    raise RuntimeError(
        "Journal 04 configuration was not found. Deploy the config folder beside "
        "journals or set NX_JOURNALS_ROOT. Attempted: {0}".format(
            " | ".join(attempted)
        )
    )


def _io_root():
    configured = _clean(os.environ.get("NX_JOURNALS_IO_DIR"))
    if configured:
        return os.path.abspath(configured)
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    return desktop if os.path.isdir(desktop) else os.getcwd()


def _load_config():
    path = os.path.join(
        _runtime_root(), "config", "attribute_reconciliation.json"
    )
    with open(path, "r", encoding="utf-8-sig") as handle:
        config = json.load(handle)
    workflow = config.get("update_workflow", {})
    if workflow.get("schema_version") != 1:
        raise RuntimeError("Unsupported or missing update_workflow schema.")
    if not workflow.get("identity_columns") or not workflow.get(
        "business_columns"
    ):
        raise RuntimeError("Update workflow has no identity/business columns.")
    return config


def _rule_map(config):
    rules = {}
    for rule in config.get("attributes", []):
        logical_name = _clean(rule.get("logical_name"))
        if logical_name:
            rules[logical_name] = rule
    return rules


def _column_specs(config, key):
    rules = _rule_map(config)
    specs = []
    seen_columns = set()
    for item in config["update_workflow"][key]:
        column = _clean(item.get("csv_column"))
        logical_name = _clean(item.get("logical_name"))
        rule = rules.get(logical_name)
        if not column or rule is None:
            raise RuntimeError(
                "Invalid {0} mapping for {1}.".format(key, logical_name)
            )
        if column in seen_columns:
            raise RuntimeError("Duplicate update CSV column: " + column)
        seen_columns.add(column)
        specs.append((column, rule))
    return specs


def update_columns(config):
    return (
        list(CONTROL_COLUMNS)
        + [column for column, _rule in _column_specs(config, "identity_columns")]
        + [column for column, _rule in _column_specs(config, "business_columns")]
    )


def _attribute_value(info):
    kind = _enum_name(getattr(info, "Type", ""))
    numeric_kind = getattr(info, "Type", None)
    if kind in ("String", "5") or numeric_kind == 5:
        return getattr(info, "StringValue", ""), "String"
    if kind in ("Real", "Number", "4") or numeric_kind == 4:
        return getattr(info, "RealValue", None), "Number"
    if kind in ("Integer", "3") or numeric_kind == 3:
        return getattr(info, "IntegerValue", None), "Integer"
    if kind in ("Boolean", "1") or numeric_kind == 1:
        return getattr(info, "BooleanValue", None), "Boolean"
    return getattr(info, "StringValue", ""), kind or "String"


def _read_attribute(nx_object, rule):
    iterator = None
    try:
        iterator = nx_object.CreateAttributeIterator()
        iterator.SetIncludeOnlyCategory(rule["category"])
        iterator.SetIncludeOnlyTitle(rule["attribute"])
        iterator.SetIncludeAlsoUnset(True)
        matches = [
            info
            for info in nx_object.GetUserAttributes(iterator)
            if _clean(getattr(info, "Category", "")) == rule["category"]
            and _clean(getattr(info, "Title", "")) == rule["attribute"]
        ]
        if not matches:
            return {
                "status": "MISSING",
                "raw_value": "",
                "type": rule["type"],
                "flags": {},
            }
        if len(matches) > 1:
            return {
                "status": "AMBIGUOUS",
                "raw_value": "",
                "type": rule["type"],
                "flags": {},
                "message": "Multiple category/title matches.",
            }
        info = matches[0]
        raw_value, actual_type = _attribute_value(info)
        unset = bool(getattr(info, "Unset", False))
        status = "UNSET" if unset else (
            "BLANK" if _clean(raw_value) == "" else "POPULATED"
        )
        return {
            "status": status,
            "raw_value": raw_value,
            "type": actual_type,
            "flags": {
                "locked": bool(getattr(info, "Locked", False)),
                "owned_by_system": bool(
                    getattr(info, "OwnedBySystem", False)
                ),
                "pdm_based": bool(getattr(info, "PdmBased", False)),
                "not_saved": bool(getattr(info, "NotSaved", False)),
            },
        }
    except Exception as exc:
        return {
            "status": "UNREADABLE",
            "raw_value": "",
            "type": rule.get("type", ""),
            "flags": {},
            "message": _exception_details(exc),
        }
    finally:
        _dispose(iterator)


def _read_identity_attribute(nx_object, rule):
    """Read hard-coded NX/CAD identity by title, independent of category."""
    try:
        raw_value = nx_object.GetStringAttribute(rule["attribute"])
        return {
            "status": (
                "BLANK" if _clean(raw_value) == "" else "POPULATED"
            ),
            "raw_value": raw_value,
            "type": "String",
            "flags": {},
        }
    except AttributeError:
        return _read_attribute(nx_object, rule)
    except Exception as exc:
        fallback = _read_attribute(nx_object, rule)
        if fallback["status"] != "MISSING":
            return fallback
        fallback["message"] = _exception_details(exc)
        return fallback


def _object_key(nx_object):
    tag = getattr(nx_object, "Tag", None)
    return ("TAG", _text(tag)) if tag is not None else ("PY", id(nx_object))


def _object_identifier(nx_object):
    for name in ("JournalIdentifier", "FullPath", "Name", "Leaf"):
        value = getattr(nx_object, name, "")
        if callable(value):
            try:
                value = value()
            except Exception:
                value = ""
        if _clean(value):
            return _clean(value)
    return _text(_object_key(nx_object))


def _children(component):
    try:
        return list(component.GetChildren())
    except Exception:
        return []


def _suppression_state(component):
    """Return suppression state and an error; unreadable is not active."""
    try:
        return bool(component.IsSuppressed), ""
    except Exception as exc:
        return None, _exception_details(exc)


def _component_string_attribute(component, title):
    """Safe read of a component-level string attribute; None when absent."""
    try:
        return component.GetStringAttribute(title)
    except Exception:
        return None


def _is_bom_visible(component):
    """Mirror NXOpenBoMExtended.py: only BoM-visible components are pulled.

    Suppression is handled separately by the caller.  Keyword-named and
    reference-flagged occurrences are excluded together with their subtrees.
    """
    name = _clean(getattr(component, "Name", ""))
    display_name = _clean(getattr(component, "DisplayName", ""))
    combined = " ".join((name, display_name)).upper()
    for keyword in IGNORE_KEYWORDS:
        if keyword in combined:
            return False
    for title in BOM_REFERENCE_ATTRIBUTES:
        raw = _component_string_attribute(component, title)
        if raw is not None and _clean(raw) in BOM_REFERENCE_FLAG_VALUES:
            return False
    return True


def collect_unique_prototypes(work_part):
    """Return ordered unique prototype parts and traversal diagnostics."""
    unique = OrderedDict()
    diagnostics = []

    def add_part(part):
        if part is not None:
            unique.setdefault(_object_key(part), part)

    add_part(work_part)
    root_component = getattr(
        getattr(work_part, "ComponentAssembly", None), "RootComponent", None
    )
    if root_component is None:
        return list(unique.values()), diagnostics

    stack = list(reversed(_children(root_component)))
    while stack:
        component = stack.pop()
        suppressed, suppression_error = _suppression_state(component)
        if suppression_error:
            diagnostics.append(
                {
                    "code": "SUPPRESSION_STATE_UNAVAILABLE",
                    "message": (
                        "Component excluded because its active suppression "
                        "state could not be read: {0}".format(
                            suppression_error
                        )
                    ),
                }
            )
            continue
        if suppressed:
            continue
        if not _is_bom_visible(component):
            continue
        prototype = getattr(component, "Prototype", None)
        if prototype is None:
            diagnostics.append(
                {
                    "code": "MISSING_MODEL",
                    "message": "Component has no loaded prototype: {0}".format(
                        _clean(getattr(component, "DisplayName", ""))
                        or _clean(getattr(component, "Name", ""))
                        or "<unknown>"
                    ),
                }
            )
        else:
            add_part(prototype)
        stack.extend(reversed(_children(component)))
    return list(unique.values()), diagnostics


def _sidecar_value(result):
    raw = result.get("raw_value", "")
    return {
        "status": result.get("status", ""),
        "raw_value": raw,
        "normalized_value": _normalized(raw),
        "type": result.get("type", ""),
        "flags": result.get("flags", {}),
        "message": result.get("message", ""),
    }


def build_pull_records(work_part, config, run_id):
    identity_specs = _column_specs(config, "identity_columns")
    business_specs = _column_specs(config, "business_columns")
    parts, traversal_diagnostics = collect_unique_prototypes(work_part)
    records = []

    for part in parts:
        messages = []
        identity_values = OrderedDict()
        identity_results = {}
        for column, rule in identity_specs:
            result = _read_identity_attribute(part, rule)
            identity_results[column] = result
            identity_values[column] = _text(result.get("raw_value", ""))
            if result["status"] in (
                "MISSING",
                "UNSET",
                "BLANK",
                "UNREADABLE",
                "AMBIGUOUS",
            ):
                messages.append(
                    "{0}: {1}".format(column, result["status"])
                )

        business_values = OrderedDict()
        business_results = {}
        for column, rule in business_specs:
            result = _read_attribute(part, rule)
            business_results[column] = result
            business_values[column] = _text(result.get("raw_value", ""))
            if result["status"] in ("UNREADABLE", "AMBIGUOUS"):
                messages.append(
                    "{0}: {1}{2}".format(
                        column,
                        result["status"],
                        " - " + result.get("message", "")
                        if result.get("message")
                        else "",
                    )
                )

        row = OrderedDict()
        row.update(
            {
                "AUDIT_RUN_ID": run_id,
                "APPROVED": "NO",
                "ENGINEER": "",
                "APPROVAL_NOTE": "",
                "PULL_STATUS": "REVIEW" if messages else "READY",
                "PULL_MESSAGE": " | ".join(messages),
            }
        )
        row.update(identity_values)
        row.update(business_values)
        records.append(
            {
                "part": part,
                "row": row,
                "identity_results": identity_results,
                "business_results": business_results,
                "messages": messages,
            }
        )

    identities = {}
    for record in records:
        key = (
            _normalized(record["row"].get("Item Number")),
            _normalized(record["row"].get("Item Rev")),
        )
        identities.setdefault(key, []).append(record)
    for key, matches in identities.items():
        if not all(key) or len(matches) < 2:
            continue
        message = "Duplicate prototype identity: {0}/{1}".format(*key)
        for record in matches:
            record["messages"].append(message)
            record["row"]["PULL_STATUS"] = "REVIEW"
            record["row"]["PULL_MESSAGE"] = " | ".join(record["messages"])

    return records, traversal_diagnostics


def _safe_filename(value):
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", _clean(value))
    return cleaned.strip("._") or "UNKNOWN"


def _write_csv(path, columns, rows):
    with open(path, "w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=columns)
        writer.writeheader()
        for row in rows:
            writer.writerow({column: row.get(column, "") for column in columns})


def _write_json(path, payload):
    with open(path, "w", encoding="utf-8") as handle:
        json.dump(payload, handle, indent=2, ensure_ascii=False)


def _listing_window(session):
    listing = getattr(session, "ListingWindow", None)
    if listing is not None:
        try:
            listing.Open()
        except Exception:
            pass
    return listing


def _log(listing, message):
    if listing is not None:
        try:
            listing.WriteLine(_text(message))
            return
        except Exception:
            pass
    print(_text(message))


def main(session):
    work_part = getattr(getattr(session, "Parts", None), "Work", None)
    if work_part is None:
        raise RuntimeError("Open an NX 3D master part or assembly first.")

    config = _load_config()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    run_id = "{0}_{1}".format(
        _safe_filename(_object_identifier(work_part)), timestamp
    )
    records, traversal_diagnostics = build_pull_records(
        work_part, config, run_id
    )

    output_root = _io_root()
    os.makedirs(output_root, exist_ok=True)
    stem = "NX_ATTRIBUTE_UPDATE_{0}".format(run_id)
    csv_path = os.path.join(output_root, stem + ".csv")
    sidecar_path = os.path.join(output_root, stem + ".baseline.json")
    columns = update_columns(config)
    _write_csv(csv_path, columns, [record["row"] for record in records])

    identity_specs = _column_specs(config, "identity_columns")
    business_specs = _column_specs(config, "business_columns")
    sidecar = {
        "schema_version": 1,
        "audit_run_id": run_id,
        "generated_at": datetime.now().isoformat(),
        "source_journal": "04_assembly_attribute_audit.py",
        "root_identifier": _object_identifier(work_part),
        "csv_filename": os.path.basename(csv_path),
        "identity_columns": [
            {
                "csv_column": column,
                "logical_name": rule["logical_name"],
                "category": rule["category"],
                "attribute": rule["attribute"],
                "type": rule["type"],
            }
            for column, rule in identity_specs
        ],
        "business_columns": [
            {
                "csv_column": column,
                "logical_name": rule["logical_name"],
                "category": rule["category"],
                "attribute": rule["attribute"],
                "type": rule["type"],
            }
            for column, rule in business_specs
        ],
        "traversal_diagnostics": traversal_diagnostics,
        "parts": [
            {
                "part_number": record["row"].get("Item Number", ""),
                "part_name": record["row"].get("Part Description", ""),
                "revision": record["row"].get("Item Rev", ""),
                "model_identifier": _object_identifier(record["part"]),
                "pull_status": record["row"]["PULL_STATUS"],
                "pull_message": record["row"]["PULL_MESSAGE"],
                "identity": {
                    column: _sidecar_value(record["identity_results"][column])
                    for column, _rule in identity_specs
                },
                "business_values": {
                    column: _sidecar_value(record["business_results"][column])
                    for column, _rule in business_specs
                },
            }
            for record in records
        ],
    }
    _write_json(sidecar_path, sidecar)

    listing = _listing_window(session)
    ready_count = sum(
        1 for record in records if record["row"]["PULL_STATUS"] == "READY"
    )
    _log(listing, "Journal 04 model attribute pull complete.")
    _log(listing, "  Unique 3D master models: {0}".format(len(records)))
    _log(listing, "  Ready rows: {0}".format(ready_count))
    _log(listing, "  Review rows: {0}".format(len(records) - ready_count))
    _log(
        listing,
        "  Traversal diagnostics: {0}".format(len(traversal_diagnostics)),
    )
    _log(listing, "  Editable CSV: " + csv_path)
    _log(listing, "  Baseline: " + sidecar_path)
    _log(
        listing,
        "Edit business fields, set APPROVED=YES, populate ENGINEER, "
        "then run Journal 05.",
    )
    return csv_path, sidecar_path


def _run_journal():
    session = NXOpen.Session.GetSession()
    listing = _listing_window(session)
    try:
        main(session)
    except Exception as exc:
        _log(listing, "JOURNAL 04 FAILED: " + _exception_details(exc))
        _log(listing, traceback.format_exc())
        raise


if __name__ == "__main__":
    _run_journal()
