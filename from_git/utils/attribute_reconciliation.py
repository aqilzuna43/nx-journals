"""Shared NX-authoritative reconciliation helpers.

This module deliberately imports only the Python standard library.  NX journals
pass NXOpen objects into the duck-typed helpers, while unit tests use fakes.
"""

import csv
import hashlib
import json
import os
import re
from collections import OrderedDict
from datetime import datetime


CONFIG_RELATIVE_PATH = os.path.join("config", "attribute_reconciliation.json")
FZ_BOM_COLUMNS = [
    "BOM Level",
    "DB_PART_NO",
    "Indented Part Name",
    "Component Name",
    "Quantity",
    "DB_PART_DESC",
    "DB_PART_NAME",
    "DB_PART_REV",
    "MFG",
    "MPN",
    "Stocking_Type",
]
VALID_ATTRIBUTE_TYPES = {"STRING", "NUMBER", "REAL", "INTEGER", "BOOLEAN", "DATE", "TIME"}
PROTECTED_LOGICAL_WRITES = {
    "part_number",
    "revision",
    "part_name",
    "part_description",
    "part_type",
    "reference_set",
    "mass_kg",
    "volume_mm3",
    "density_kg_per_mm3",
    "material",
    "material_missing_assignments",
    "material_multiple_assigned",
    "length_x_mm",
    "width_y_mm",
    "height_z_mm",
}
MASTER_TO_NX_COLUMNS = OrderedDict(
    [
        ("Level", "BOM Level"),
        ("Item Number", "DB_PART_NO"),
        ("Part Description", "DB_PART_NAME"),
        ("Item Rev", "DB_PART_REV"),
        ("Qty", "Quantity"),
        ("Mfr. Name", "MFG"),
        ("Mfr. Part Number", "MPN"),
        ("Reference Notes", "Stocking_Type"),
    ]
)


class ReconciliationError(RuntimeError):
    """Raised for deterministic configuration or input failures."""


def text(value):
    return "" if value is None else str(value)


def enum_name(value):
    name = getattr(value, "name", None)
    if name:
        return str(name)
    raw = text(value)
    return raw.rsplit(".", 1)[-1]


def clean_text(value):
    return text(value).strip()


def normalized_text(value):
    return " ".join(clean_text(value).split()).upper()


def config_path(repo_root):
    return os.path.join(repo_root, CONFIG_RELATIVE_PATH)


def load_config(repo_root):
    path = config_path(repo_root)
    try:
        with open(path, "r", encoding="utf-8") as handle:
            config = json.load(handle)
    except (OSError, ValueError) as exc:
        raise ReconciliationError("Unable to load reconciliation config {0}: {1}".format(path, exc))
    validate_config(config)
    return config


def validate_config(config):
    if not isinstance(config, dict):
        raise ReconciliationError("Reconciliation config must be a JSON object.")
    if config.get("schema_version") != 1:
        raise ReconciliationError("Unsupported reconciliation schema_version; expected 1.")
    if config.get("authority") != "NX_TEAMCENTER":
        raise ReconciliationError("authority must be NX_TEAMCENTER.")
    if config.get("save_policy") not in ("NO_SAVE", "SAVE_CHANGED_PARTS"):
        raise ReconciliationError("save_policy must be NO_SAVE or SAVE_CHANGED_PARTS.")
    inputs = config.get("inputs", {})
    for key in ("drawing_scope_filename", "master_reference_filename", "corrections_filename"):
        if not clean_text(inputs.get(key)):
            raise ReconciliationError("inputs.{0} is required.".format(key))
    release_policy = config.get("release_policy", {})
    placeholders = release_policy.get("placeholder_values")
    if not isinstance(placeholders, list) or not placeholders:
        raise ReconciliationError("release_policy.placeholder_values must be a non-empty array.")
    aliases = release_policy.get("boolean_aliases", {})
    if not all(isinstance(aliases.get(key), list) and aliases[key] for key in ("Y", "N")):
        raise ReconciliationError("release_policy.boolean_aliases must define non-empty Y and N arrays.")
    try:
        if float(release_policy.get("mass_relative_tolerance")) <= 0:
            raise ValueError()
    except (TypeError, ValueError):
        raise ReconciliationError("release_policy.mass_relative_tolerance must be positive.")
    attributes = config.get("attributes")
    if not isinstance(attributes, list) or not attributes:
        raise ReconciliationError("attributes must be a non-empty array.")
    logical_names = set()
    composite_keys = set()
    for index, rule in enumerate(attributes):
        if not isinstance(rule, dict):
            raise ReconciliationError("attributes[{0}] must be an object.".format(index))
        required = ("logical_name", "category", "attribute", "type", "source_owner")
        missing = [name for name in required if not clean_text(rule.get(name))]
        if missing:
            raise ReconciliationError(
                "attributes[{0}] missing: {1}".format(index, ", ".join(missing))
            )
        logical = rule["logical_name"]
        composite = (rule["category"], rule["attribute"])
        if logical in logical_names:
            raise ReconciliationError("Duplicate logical_name: {0}".format(logical))
        if composite in composite_keys:
            raise ReconciliationError(
                "Duplicate category/attribute mapping: {0}/{1}".format(*composite)
            )
        logical_names.add(logical)
        composite_keys.add(composite)
        if normalized_text(rule["type"]) not in VALID_ATTRIBUTE_TYPES:
            raise ReconciliationError("Unsupported attribute type on {0}.".format(logical))
        required_on = rule.get("required_on", [])
        if not isinstance(required_on, list) or any(
            target not in ("MODEL", "DRAWING") for target in required_on
        ):
            raise ReconciliationError("Invalid required_on target on {0}.".format(logical))
        targets = rule.get("write_targets", [])
        if rule.get("writable") and not targets:
            raise ReconciliationError("Writable rule {0} has no write_targets.".format(logical))
        if any(target not in ("MODEL", "DRAWING") for target in targets):
            raise ReconciliationError("Invalid write target on {0}.".format(logical))
        if rule.get("writable") and logical in PROTECTED_LOGICAL_WRITES:
            raise ReconciliationError("Protected attribute cannot be writable: {0}.".format(logical))
    identity = config.get("identity", {})
    for logical in ("part_number", "revision"):
        if logical not in identity or logical not in logical_names:
            raise ReconciliationError("Missing identity mapping: {0}".format(logical))
    drawing = config.get("drawing", {})
    template = clean_text(drawing.get("identifier_template"))
    if not template:
        raise ReconciliationError("drawing.identifier_template is required.")
    if not all(token in template for token in ("{part_number}", "{revision}", "{index}")):
        raise ReconciliationError("drawing.identifier_template is missing required placeholders.")
    if int(drawing.get("max_index", 0)) < 1:
        raise ReconciliationError("drawing.max_index must be positive.")
    columns = config.get("bom_export", {}).get("columns")
    if columns != FZ_BOM_COLUMNS:
        raise ReconciliationError("bom_export.columns must match the FZ NX import contract.")
    return config


def rules_by_logical(config):
    return {rule["logical_name"]: rule for rule in config["attributes"]}


def file_sha256(path):
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def config_sha256(config):
    payload = json.dumps(config, sort_keys=True, separators=(",", ":"), ensure_ascii=True)
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def new_run_id(root_part_number, now=None):
    stamp = (now or datetime.now()).strftime("%Y%m%d_%H%M%S")
    safe_root = re.sub(r"[^A-Za-z0-9_.-]+", "_", clean_text(root_part_number)) or "UNKNOWN"
    return "{0}_{1}".format(safe_root, stamp)


def read_csv_rows(path):
    last_error = None
    for encoding in ("utf-8-sig", "utf-8", "cp1252"):
        try:
            with open(path, "r", encoding=encoding, newline="") as handle:
                return list(csv.reader(handle))
        except UnicodeDecodeError as exc:
            last_error = exc
    raise ReconciliationError("Unable to decode CSV {0}: {1}".format(path, last_error))


def read_csv_dicts(path):
    rows = read_csv_rows(path)
    if not rows:
        raise ReconciliationError("CSV is empty: {0}".format(path))
    headers = [clean_text(value) for value in rows[0]]
    if len(headers) != len(set(headers)):
        raise ReconciliationError("CSV contains duplicate headers: {0}".format(path))
    result = []
    for row_number, row in enumerate(rows[1:], 2):
        if not any(clean_text(value) for value in row):
            continue
        values = list(row) + [""] * max(0, len(headers) - len(row))
        record = {header: clean_text(values[index]) for index, header in enumerate(headers)}
        record["__ROW_NUMBER__"] = row_number
        result.append(record)
    return headers, result


def write_csv(path, headers, rows):
    with open(path, "w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(headers)
        for row in rows:
            if isinstance(row, dict):
                writer.writerow([row.get(header, "") for header in headers])
            else:
                writer.writerow(["" if value is None else value for value in row])
    return path


def normalize_value(value, rule, config):
    mode = rule.get("comparison", "NORMALIZED_TEXT")
    raw = clean_text(value)
    if mode == "EXACT":
        return raw
    if mode == "NUMBER":
        if raw == "":
            return ""
        try:
            return format(float(raw), ".15g")
        except (TypeError, ValueError):
            return raw
    if mode == "BOOLEAN_ALIAS":
        upper = raw.upper()
        aliases = config.get("release_policy", {}).get("boolean_aliases", {})
        for canonical, values in aliases.items():
            if upper in {clean_text(item).upper() for item in values}:
                return canonical
        return upper
    if mode == "TRIMMED_CASE_INSENSITIVE":
        return raw.upper()
    return normalized_text(raw)


def is_placeholder(value, config):
    candidate = normalized_text(value)
    placeholders = config.get("release_policy", {}).get("placeholder_values", [])
    return candidate in {normalized_text(item) for item in placeholders}


def validate_attribute_value(result, rule, config):
    """Return (status, code, message) for one configured attribute result."""
    required = bool(rule.get("required_for_certification"))
    status = result.get("status")
    raw = result.get("raw_value", "")
    if status == "AMBIGUOUS":
        return "BLOCK", "AMBIGUOUS_MATCH", "Multiple attributes matched category and title."
    if status == "UNREADABLE":
        return "BLOCK", "UNREADABLE_ATTRIBUTE", result.get("message", "Attribute could not be read.")
    if status in ("MISSING", "UNSET"):
        if required:
            return "BLOCK", "ATTRIBUTE_MISSING_MODEL", "Required attribute is not set."
        return "INFO", "NOT_APPLICABLE", "Optional attribute is not set."
    if is_placeholder(raw, config):
        if required:
            return "BLOCK", "PLACEHOLDER_VALUE", "Required attribute is blank or a placeholder."
        return "WARN", "PLACEHOLDER_VALUE", "Optional attribute is blank or a placeholder."
    normalized = normalize_value(raw, rule, config)
    allowed = rule.get("allowed_values") or []
    if allowed:
        allowed_normalized = {normalize_value(item, rule, config) for item in allowed}
        if normalized not in allowed_normalized:
            return "BLOCK" if required else "WARN", "INVALID_CONTROLLED_VALUE", (
                "Value is outside the configured controlled vocabulary."
            )
    if rule.get("validation") == "GREATER_THAN_ZERO":
        try:
            if float(raw) <= 0:
                raise ValueError()
        except (TypeError, ValueError):
            return "BLOCK" if required else "WARN", "INVALID_CONTROLLED_VALUE", (
                "Value must be numeric and greater than zero."
            )
    return "PASS", "PASS", "Value satisfies the configured rule."


def _attribute_value(info):
    kind = enum_name(getattr(info, "Type", ""))
    field_by_type = {
        "String": "StringValue",
        "Real": "RealValue",
        "Number": "RealValue",
        "Integer": "IntegerValue",
        "Boolean": "BooleanValue",
        "Time": "TimeValue",
        "Reference": "ReferenceValue",
    }
    field = field_by_type.get(kind)
    if field is None:
        return "", kind
    return getattr(info, field, ""), kind


def read_attribute(nx_object, rule, include_unset=True):
    category = clean_text(rule.get("category"))
    title = clean_text(rule.get("attribute"))
    iterator = None
    try:
        try:
            iterator = nx_object.CreateAttributeIterator()
            iterator.SetIncludeOnlyCategory(category)
            iterator.SetIncludeOnlyTitle(title)
            iterator.SetIncludeAlsoUnset(bool(include_unset))
            infos = list(nx_object.GetUserAttributes(iterator))
        except Exception:
            infos = list(nx_object.GetUserAttributes())
        matches = [
            info
            for info in infos
            if clean_text(getattr(info, "Category", "")) == category
            and clean_text(getattr(info, "Title", "")) == title
        ]
        if not matches:
            return {
                "status": "MISSING",
                "raw_value": "",
                "category": category,
                "title": title,
                "type": clean_text(rule.get("type")),
                "flags": {},
            }
        if len(matches) > 1:
            return {
                "status": "AMBIGUOUS",
                "raw_value": "",
                "category": category,
                "title": title,
                "type": clean_text(rule.get("type")),
                "flags": {},
            }
        info = matches[0]
        raw, kind = _attribute_value(info)
        unset = bool(getattr(info, "Unset", False))
        return {
            "status": "UNSET" if unset else ("BLANK" if clean_text(raw) == "" else "POPULATED"),
            "raw_value": raw,
            "category": category,
            "title": title,
            "type": kind,
            "flags": {
                "locked": bool(getattr(info, "Locked", False)),
                "owned_by_system": bool(getattr(info, "OwnedBySystem", False)),
                "pdm_based": bool(getattr(info, "PdmBased", False)),
                "required": bool(getattr(info, "Required", False)),
                "not_saved": bool(getattr(info, "NotSaved", False)),
            },
        }
    except Exception as exc:
        return {
            "status": "UNREADABLE",
            "raw_value": "",
            "category": category,
            "title": title,
            "type": clean_text(rule.get("type")),
            "flags": {},
            "message": "{0}: {1}".format(type(exc).__name__, exc),
        }
    finally:
        if iterator is not None:
            try:
                iterator.FreeResource()
            except Exception:
                pass


def read_attributes(nx_object, config):
    """Read all configured attributes with one NX enumeration when possible."""
    try:
        infos = list(nx_object.GetUserAttributes())
    except Exception:
        return {rule["logical_name"]: read_attribute(nx_object, rule) for rule in config["attributes"]}
    indexed = {}
    for info in infos:
        key = (clean_text(getattr(info, "Category", "")), clean_text(getattr(info, "Title", "")))
        indexed.setdefault(key, []).append(info)
    results = {}
    for rule in config["attributes"]:
        category = clean_text(rule.get("category"))
        title = clean_text(rule.get("attribute"))
        matches = indexed.get((category, title), [])
        if not matches:
            results[rule["logical_name"]] = {
                "status": "MISSING",
                "raw_value": "",
                "category": category,
                "title": title,
                "type": clean_text(rule.get("type")),
                "flags": {},
            }
            continue
        if len(matches) > 1:
            results[rule["logical_name"]] = {
                "status": "AMBIGUOUS",
                "raw_value": "",
                "category": category,
                "title": title,
                "type": clean_text(rule.get("type")),
                "flags": {},
            }
            continue
        info = matches[0]
        raw, kind = _attribute_value(info)
        unset = bool(getattr(info, "Unset", False))
        results[rule["logical_name"]] = {
            "status": "UNSET" if unset else ("BLANK" if clean_text(raw) == "" else "POPULATED"),
            "raw_value": raw,
            "category": category,
            "title": title,
            "type": kind,
            "flags": {
                "locked": bool(getattr(info, "Locked", False)),
                "owned_by_system": bool(getattr(info, "OwnedBySystem", False)),
                "pdm_based": bool(getattr(info, "PdmBased", False)),
                "required": bool(getattr(info, "Required", False)),
                "not_saved": bool(getattr(info, "NotSaved", False)),
            },
        }
    return results


def identity_from_attributes(values):
    pn = clean_text(values.get("part_number", {}).get("raw_value"))
    rev = clean_text(values.get("revision", {}).get("raw_value"))
    return pn, rev


def object_identifier(nx_object):
    for name in ("JournalIdentifier", "FullPath", "Leaf", "Name"):
        try:
            value = clean_text(getattr(nx_object, name))
            if value:
                return value
        except Exception:
            pass
    return ""


def object_key(nx_object):
    try:
        tag = getattr(nx_object, "Tag")
        if tag:
            return ("TAG", str(tag))
    except Exception:
        pass
    return ("ID", object_identifier(nx_object) or str(id(nx_object)))


def _component_name(component):
    for name in ("DisplayName", "Name"):
        try:
            value = clean_text(getattr(component, name))
            if value:
                return value
        except Exception:
            pass
    return "COMPONENT"


def _is_suppressed(component):
    try:
        return bool(component.IsSuppressed), ""
    except Exception as exc:
        return False, "{0}: {1}".format(type(exc).__name__, exc)


def root_component(part):
    try:
        return part.ComponentAssembly.RootComponent
    except Exception:
        return None


def collect_bom_nodes(work_part, config):
    """Return deterministic grouped BOM nodes and traversal findings."""
    cache = {}
    findings = []

    def values_for(part):
        key = object_key(part)
        if key not in cache:
            cache[key] = read_attributes(part, config)
        return cache[key]

    root_values = values_for(work_part)
    root_pn, root_rev = identity_from_attributes(root_values)
    nodes = [
        {
            "level": 0,
            "quantity": 1,
            "part": work_part,
            "component": None,
            "parent_part_number": "",
            "parent_revision": "",
            "part_number": root_pn,
            "revision": root_rev,
            "occurrence_path": root_pn or object_identifier(work_part),
            "attributes": root_values,
        }
    ]

    def walk(parent_component, parent_node):
        try:
            children = list(parent_component.GetChildren()) if parent_component is not None else []
        except Exception as exc:
            findings.append(
                {
                    "code": "BOM_STRUCTURE_MISMATCH",
                    "message": "Unable to enumerate children: {0}: {1}".format(type(exc).__name__, exc),
                    "occurrence_path": parent_node["occurrence_path"],
                }
            )
            return
        grouped = OrderedDict()
        for child in children:
            suppressed, suppression_error = _is_suppressed(child)
            if suppression_error:
                findings.append(
                    {
                        "code": "UNREADABLE_ATTRIBUTE",
                        "message": "Unable to read suppression state: " + suppression_error,
                        "occurrence_path": parent_node["occurrence_path"],
                    }
                )
            if suppressed:
                continue
            prototype_error_reported = False
            try:
                prototype = child.Prototype
            except Exception as exc:
                prototype = None
                prototype_error_reported = True
                findings.append(
                    {
                        "code": "MISSING_MODEL",
                        "message": "Component prototype is unavailable: {0}: {1}".format(type(exc).__name__, exc),
                        "occurrence_path": parent_node["occurrence_path"] + "/" + _component_name(child),
                    }
                )
            if prototype is None:
                if not prototype_error_reported:
                    findings.append(
                        {
                            "code": "MISSING_MODEL",
                            "message": "Component prototype is unavailable or not loaded.",
                            "occurrence_path": parent_node["occurrence_path"]
                            + "/"
                            + _component_name(child),
                        }
                    )
                continue
            values = values_for(prototype)
            pn, rev = identity_from_attributes(values)
            group_key = (normalized_text(pn), normalized_text(rev))
            if not pn or not rev:
                group_key = ("UNRESOLVED", _component_name(child), object_key(prototype))
            grouped.setdefault(group_key, []).append((child, prototype, values, pn, rev))
        for members in grouped.values():
            child, prototype, values, pn, rev = members[0]
            prototype_keys = {object_key(member[1]) for member in members}
            path = parent_node["occurrence_path"] + "/" + (pn or _component_name(child))
            if len(prototype_keys) > 1:
                findings.append(
                    {
                        "code": "AMBIGUOUS_MATCH",
                        "message": "Multiple prototypes share the same part-number/revision under one parent.",
                        "occurrence_path": path,
                    }
                )
            node = {
                "level": parent_node["level"] + 1,
                "quantity": len(members),
                "part": prototype,
                "component": child,
                "parent_part_number": parent_node["part_number"],
                "parent_revision": parent_node["revision"],
                "part_number": pn,
                "revision": rev,
                "occurrence_path": path,
                "attributes": values,
            }
            nodes.append(node)
            walk(child, node)

    root = root_component(work_part)
    if root is not None:
        walk(root, nodes[0])
    prototypes_by_identity = {}
    for node in nodes:
        identity = (normalized_text(node["part_number"]), normalized_text(node["revision"]))
        if not all(identity):
            continue
        prototypes_by_identity.setdefault(identity, set()).add(object_key(node["part"]))
    for (part_number, revision), prototype_keys in prototypes_by_identity.items():
        if len(prototype_keys) > 1:
            findings.append(
                {
                    "code": "AMBIGUOUS_MATCH",
                    "message": (
                        "Multiple loaded prototype objects share exact identity {0}/{1}."
                    ).format(part_number, revision),
                    "part_number": part_number,
                    "revision": revision,
                }
            )
    return nodes, findings


def _raw(node, logical_name):
    return clean_text(node.get("attributes", {}).get(logical_name, {}).get("raw_value"))


def bom_export_row(node):
    level = int(node["level"])
    name = _raw(node, "part_name")
    return {
        "BOM Level": level,
        "DB_PART_NO": node.get("part_number", ""),
        "Indented Part Name": ("    " * level) + name,
        "Component Name": name,
        "Quantity": node.get("quantity", 1),
        "DB_PART_DESC": _raw(node, "part_description"),
        "DB_PART_NAME": name,
        "DB_PART_REV": node.get("revision", ""),
        "MFG": _raw(node, "manufacturer"),
        "MPN": _raw(node, "manufacturer_part_number"),
        "Stocking_Type": _raw(node, "stocking_type"),
    }


def bom_export_rows(nodes):
    return [bom_export_row(node) for node in nodes]


def load_drawing_scope(path):
    headers, rows = read_csv_dicts(path)
    required = ["Item Number", "Item Rev", "Drawing Required"]
    missing = [name for name in required if name not in headers]
    if missing:
        raise ReconciliationError("Drawing scope missing columns: {0}".format(", ".join(missing)))
    scope = {}
    findings = []
    for row in rows:
        pn = normalized_text(row["Item Number"])
        rev = normalized_text(row["Item Rev"])
        decision = normalized_text(row["Drawing Required"])
        if decision in ("Y", "TRUE", "1"):
            decision = "YES"
        elif decision in ("N", "FALSE", "0"):
            decision = "NO"
        if decision not in ("YES", "NO", "REVIEW"):
            findings.append(
                {
                    "code": "INVALID_CONTROLLED_VALUE",
                    "message": "Invalid Drawing Required value on row {0}.".format(row["__ROW_NUMBER__"]),
                    "part_number": row["Item Number"],
                    "revision": row["Item Rev"],
                }
            )
            decision = "REVIEW"
        key = (pn, rev)
        if key in scope and scope[key] != decision:
            findings.append(
                {
                    "code": "DRAWING_SCOPE_REVIEW",
                    "message": "Conflicting drawing-scope rows for part/revision.",
                    "part_number": row["Item Number"],
                    "revision": row["Item Rev"],
                }
            )
            scope[key] = "REVIEW"
        else:
            scope[key] = decision
    return scope, findings


def drawing_decision(scope, part_number, revision):
    return scope.get((normalized_text(part_number), normalized_text(revision)), "REVIEW")


def compare_master_reference(path, root_part_number, nx_rows):
    headers, rows = read_csv_dicts(path)
    missing = [name for name in MASTER_TO_NX_COLUMNS if name not in headers]
    if missing:
        raise ReconciliationError("MASTER reference missing columns: {0}".format(", ".join(missing)))
    matches = [
        index
        for index, row in enumerate(rows)
        if normalized_text(row.get("Item Number")) == normalized_text(root_part_number)
    ]
    if len(matches) != 1:
        return [
            {
                "code": "AMBIGUOUS_MATCH" if matches else "DOWNSTREAM_BOM_DRIFT",
                "message": "MASTER root match count is {0}; expected exactly one.".format(len(matches)),
            }
        ]
    start = matches[0]
    try:
        root_level = int(float(rows[start]["Level"]))
    except (TypeError, ValueError):
        return [{"code": "DOWNSTREAM_BOM_DRIFT", "message": "MASTER root Level is invalid."}]
    end = start + 1
    while end < len(rows):
        try:
            level = int(float(rows[end]["Level"]))
        except (TypeError, ValueError):
            level = root_level
        if level <= root_level:
            break
        end += 1
    master_rows = rows[start:end]
    findings = []
    if len(master_rows) != len(nx_rows):
        findings.append(
            {
                "code": "DOWNSTREAM_BOM_DRIFT",
                "message": "MASTER subtree has {0} rows; NX has {1}.".format(len(master_rows), len(nx_rows)),
            }
        )
    for index, (master, nx_row) in enumerate(zip(master_rows, nx_rows), 1):
        for master_column, nx_column in MASTER_TO_NX_COLUMNS.items():
            left = master.get(master_column, "")
            right = nx_row.get(nx_column, "")
            if master_column == "Level":
                try:
                    left = int(float(left)) - root_level
                    right = int(float(right))
                except (TypeError, ValueError):
                    pass
            elif master_column == "Qty":
                try:
                    left = float(left)
                    right = float(right)
                except (TypeError, ValueError):
                    pass
            else:
                left = normalized_text(left)
                right = normalized_text(right)
            if left != right:
                findings.append(
                    {
                        "code": "DOWNSTREAM_BOM_DRIFT",
                        "message": "Row {0} {1} differs: MASTER={2!r}, NX={3!r}.".format(
                            index, master_column, master.get(master_column, ""), nx_row.get(nx_column, "")
                        ),
                        "part_number": nx_row.get("DB_PART_NO", ""),
                        "revision": nx_row.get("DB_PART_REV", ""),
                    }
                )
    return findings


def approximately_equal(left, right, relative_tolerance):
    try:
        left_value = float(left)
        right_value = float(right)
    except (TypeError, ValueError):
        return False
    scale = max(abs(left_value), abs(right_value), 1e-30)
    return abs(left_value - right_value) <= float(relative_tolerance) * scale


def mass_density_volume_consistent(mass, density, volume, relative_tolerance):
    try:
        calculated = float(density) * float(volume)
    except (TypeError, ValueError):
        return False, ""
    return approximately_equal(mass, calculated, relative_tolerance), calculated


def _point_tuple(point):
    if isinstance(point, (list, tuple)):
        return tuple(float(value) for value in point[:3])
    return (float(point.X), float(point.Y), float(point.Z))


def _matrix_tuple(matrix):
    if isinstance(matrix, (list, tuple)):
        if len(matrix) == 9:
            return tuple(tuple(float(matrix[row * 3 + col]) for col in range(3)) for row in range(3))
        return tuple(tuple(float(value) for value in row[:3]) for row in matrix[:3])
    return (
        (float(matrix.Xx), float(matrix.Xy), float(matrix.Xz)),
        (float(matrix.Yx), float(matrix.Yy), float(matrix.Yz)),
        (float(matrix.Zx), float(matrix.Zy), float(matrix.Zz)),
    )


IDENTITY_MATRIX = ((1.0, 0.0, 0.0), (0.0, 1.0, 0.0), (0.0, 0.0, 1.0))


def transform_point(point, rotation=IDENTITY_MATRIX, translation=(0.0, 0.0, 0.0)):
    return tuple(
        translation[row] + sum(rotation[row][col] * point[col] for col in range(3))
        for row in range(3)
    )


def compose_transform(parent_rotation, parent_translation, child_rotation, child_translation):
    rotation = tuple(
        tuple(
            sum(parent_rotation[row][k] * child_rotation[k][col] for k in range(3))
            for col in range(3)
        )
        for row in range(3)
    )
    translation = transform_point(child_translation, parent_rotation, parent_translation)
    return rotation, translation


def corners_from_exact_box(min_corner, directions, distances):
    origin = _point_tuple(min_corner)
    matrix = _matrix_tuple(directions)
    lengths = tuple(float(value) for value in distances[:3])
    axes = tuple(tuple(matrix[row][axis] for row in range(3)) for axis in range(3))
    corners = []
    for x_bit in (0, 1):
        for y_bit in (0, 1):
            for z_bit in (0, 1):
                factors = (x_bit, y_bit, z_bit)
                corners.append(
                    tuple(
                        origin[row]
                        + sum(axes[axis][row] * lengths[axis] * factors[axis] for axis in range(3))
                        for row in range(3)
                    )
                )
    return corners


def _ask_exact_box(uf_session, body):
    method = uf_session.Modl.AskBoundingBoxExact
    try:
        result = method(body.Tag, 0)
        if isinstance(result, tuple) and len(result) >= 3:
            return result[0], result[1], result[2]
    except TypeError:
        pass
    min_corner = [0.0, 0.0, 0.0]
    directions = [0.0] * 9
    distances = [0.0, 0.0, 0.0]
    result = method(body.Tag, 0, min_corner, directions, distances)
    if isinstance(result, tuple) and len(result) >= 3:
        return result[0], result[1], result[2]
    return min_corner, directions, distances


def _component_transform(component):
    try:
        result = component.GetPosition()
        if isinstance(result, tuple) and len(result) >= 2:
            return _matrix_tuple(result[1]), _point_tuple(result[0])
    except TypeError:
        pass
    except Exception:
        raise
    raise ReconciliationError("Component.GetPosition did not return position and orientation.")


def exact_model_dimensions(part, uf_session):
    """Return X/Y/Z exact bounding-box extents for a part and its occurrences."""
    world_points = []

    def add_part_bodies(candidate, rotation, translation):
        try:
            bodies = list(candidate.Bodies)
        except Exception:
            bodies = []
        for body in bodies:
            try:
                if hasattr(body, "IsSolidBody") and not body.IsSolidBody:
                    continue
                min_corner, directions, distances = _ask_exact_box(uf_session, body)
                for point in corners_from_exact_box(min_corner, directions, distances):
                    world_points.append(transform_point(point, rotation, translation))
            except Exception as exc:
                raise ReconciliationError(
                    "Exact bounding box failed for body {0}: {1}: {2}".format(
                        object_identifier(body), type(exc).__name__, exc
                    )
                )

    def walk(component, rotation, translation):
        try:
            children = list(component.GetChildren())
        except Exception as exc:
            raise ReconciliationError("Bounding-box traversal failed: {0}: {1}".format(type(exc).__name__, exc))
        for child in children:
            suppressed, suppression_error = _is_suppressed(child)
            if suppression_error:
                raise ReconciliationError("Suppression state unavailable: " + suppression_error)
            if suppressed:
                continue
            child_rotation, child_translation = _component_transform(child)
            next_rotation, next_translation = compose_transform(
                rotation, translation, child_rotation, child_translation
            )
            prototype = child.Prototype
            add_part_bodies(prototype, next_rotation, next_translation)
            walk(child, next_rotation, next_translation)

    add_part_bodies(part, IDENTITY_MATRIX, (0.0, 0.0, 0.0))
    root = root_component(part)
    if root is not None:
        walk(root, IDENTITY_MATRIX, (0.0, 0.0, 0.0))
    if not world_points:
        raise ReconciliationError("No solid bodies were available for exact bounding-box derivation.")
    minimum = tuple(min(point[axis] for point in world_points) for axis in range(3))
    maximum = tuple(max(point[axis] for point in world_points) for axis in range(3))
    return tuple(maximum[axis] - minimum[axis] for axis in range(3))
