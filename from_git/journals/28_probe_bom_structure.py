"""J28 - read-only raw BoM structure checkpoint.

Capture every occurrence returned by the active NX assembly, including the
root, suppressed occurrences, Reference-Only occurrences, and all descendants.
The journal does not group, filter, load, update, or persist NX/Teamcenter data.

The CSV is the occurrence ledger.  The JSON is the run summary and includes
full typed instance-attribute inventories only for occurrences whose BoM
controls are present, inconsistent, or unreadable.

Target: NX X 2506 embedded Python (NX 2312-compatible APIs where practical).
Run via: NX > Tools > Journal > Play
"""

import csv
import datetime
import hashlib
import json
import os
import traceback
import uuid

import NXOpen


BUILD = "J28-NX2506-BOM-STRUCTURE-PROBE-V1"
SCHEMA_VERSION = 1
OUTPUT_FOLDER = "NX_BOM_STRUCTURE_PROBE"
MAX_OCCURRENCES = 100000
PROGRESS_INTERVAL = 500

OBSERVED = "OBSERVED"
ERROR = "ERROR"
UNAVAILABLE = "UNAVAILABLE"
NOT_APPLICABLE = "NOT_APPLICABLE"

CONTROL_ATTRIBUTES = (
    "REFERENCE_COMPONENT",
    "PLIST_IGNORE_MEMBER",
    "PLIST_IGNORE_SUBASSEMBLY",
)
LEGACY_IGNORE_KEYWORDS = (
    "CSYS",
    "COORDINATE",
    "DATUM",
    "REFERENCE",
    "SKELETON",
)
LEGACY_ACTIVE_VALUES = ("", "YES", "1", "True", "true", "yes")

CSV_COLUMNS = (
    "SCHEMA_VERSION",
    "RUN_ID",
    "RUN_TIMESTAMP",
    "JOURNAL_BUILD",
    "SEQUENCE",
    "LEVEL",
    "STRUCTURAL_PATH",
    "DISPLAY_PATH",
    "PARENT_PATH",
    "SIBLING_INDEX",
    "COMPONENT_NAME",
    "COMPONENT_DISPLAY_NAME",
    "COMPONENT_TAG",
    "COMPONENT_JOURNAL_IDENTIFIER",
    "INSTANCE_STABLE_ID_STATUS",
    "INSTANCE_STABLE_ID",
    "INSTANCE_STABLE_ID_ERROR",
    "PROTOTYPE_STATUS",
    "PROTOTYPE_TYPE",
    "PROTOTYPE_TAG",
    "PROTOTYPE_NAME",
    "PROTOTYPE_PATH",
    "PROTOTYPE_LOAD_STATE_STATUS",
    "PROTOTYPE_LOAD_STATE",
    "DB_PART_NO_STATUS",
    "DB_PART_NO",
    "DB_PART_REV_STATUS",
    "DB_PART_REV",
    "DB_PART_NAME_STATUS",
    "DB_PART_NAME",
    "STOCKING_TYPE_STATUS",
    "STOCKING_TYPE",
    "SUPPRESSED_STATUS",
    "SUPPRESSED",
    "REFERENCE_SET_STATUS",
    "REFERENCE_SET",
    "NON_GEOMETRIC_STATUS",
    "NON_GEOMETRIC",
    "REPRESENTATION_MODE_STATUS",
    "REPRESENTATION_MODE",
    "COMPONENT_LAYER_STATUS",
    "COMPONENT_LAYER",
    "REFERENCE_COMPONENT_PRESENT",
    "REFERENCE_COMPONENT_RAW_VALUE",
    "REFERENCE_COMPONENT_VALUE_STATE",
    "REFERENCE_COMPONENT_READ_STATUS",
    "REFERENCE_COMPONENT_READ_ERROR",
    "PLIST_IGNORE_MEMBER_PRESENT",
    "PLIST_IGNORE_MEMBER_RAW_VALUE",
    "PLIST_IGNORE_MEMBER_VALUE_STATE",
    "PLIST_IGNORE_MEMBER_READ_STATUS",
    "PLIST_IGNORE_MEMBER_READ_ERROR",
    "PLIST_IGNORE_SUBASSEMBLY_PRESENT",
    "PLIST_IGNORE_SUBASSEMBLY_RAW_VALUE",
    "PLIST_IGNORE_SUBASSEMBLY_VALUE_STATE",
    "PLIST_IGNORE_SUBASSEMBLY_READ_STATUS",
    "PLIST_IGNORE_SUBASSEMBLY_READ_ERROR",
    "DIRECT_CONTROL_CLASSIFICATION",
    "NEAREST_CONTROL_ANCESTOR_PATH",
    "LEGACY_IGNORE_KEYWORD_MATCH",
    "CURRENT_EXTENDED_BOM_PREDICTION",
    "CURRENT_EXTENDED_BOM_CONTROLLING_PATH",
    "ROW_EVIDENCE_STATUS",
    "PROBE_ERRORS",
)

JSON_CONTRACT_KEYS = (
    "schema_version",
    "journal_build",
    "run_id",
    "run_timestamp",
    "run_status",
    "scope",
    "nx_runtime",
    "root_assembly",
    "work_part_modified",
    "summary",
    "classification_counts",
    "control_descendant_counts",
    "traversal_errors",
    "read_errors",
    "flagged_occurrences",
    "schema_hashes",
    "csv_sha256",
)


def clean(value):
    if value is None:
        return ""
    try:
        return str(value).strip()
    except Exception:
        return ""


def enum_text(value):
    if value is None:
        return ""
    try:
        return clean(value.name)
    except Exception:
        return clean(value)


def error_text(error):
    message = clean(error) or type(error).__name__
    code = clean(getattr(error, "ErrorCode", ""))
    if code:
        return "{0} [NX error {1}]".format(message, code)
    return message


def probe(status, value=None, source="", error=""):
    return {
        "status": status,
        "value": value,
        "source": source,
        "error": clean(error),
    }


def observed(value, source):
    return probe(OBSERVED, value=value, source=source)


def failed(source, error):
    return probe(ERROR, value=None, source=source, error=error_text(error))


def unavailable(source, reason):
    return probe(UNAVAILABLE, value=None, source=source, error=reason)


def property_probe(value, property_name):
    source = "{0}.{1}".format(object_kind(value) or "object", property_name)
    if value is None:
        return unavailable(source, "Object is unavailable.")
    try:
        result = getattr(value, property_name)
    except AttributeError:
        return unavailable(source, "Property is not exposed by this runtime object type.")
    except Exception as error:
        return failed(source, error)
    try:
        return observed(result() if callable(result) else result, source)
    except Exception as error:
        return failed(source, error)


def method_probe(value, method_name, *args):
    source = "{0}.{1}".format(object_kind(value) or "object", method_name)
    if value is None:
        return unavailable(source, "Object is unavailable.")
    try:
        method = getattr(value, method_name)
    except AttributeError:
        return unavailable(source, "Method is not exposed by this runtime object type.")
    except Exception as error:
        return failed(source, error)
    try:
        return observed(method(*args), source)
    except Exception as error:
        return failed(source, error)


def safe_value(value, property_name, fallback=""):
    if value is None:
        return fallback
    try:
        result = getattr(value, property_name)
        return result() if callable(result) else result
    except Exception:
        return fallback


def safe_name(value, fallback="<unavailable>"):
    for property_name in ("DisplayName", "Name", "Leaf", "JournalIdentifier", "FullPath"):
        result = clean(safe_value(value, property_name))
        if result:
            return result
    return fallback


def object_tag(value):
    try:
        tag = clean(value.Tag)
        return tag if tag and tag != "0" else ""
    except Exception:
        return ""


def object_kind(value):
    if value is None:
        return ""
    try:
        name = clean(getattr(value.GetType(), "Name", ""))
        if name:
            return name
    except Exception:
        pass
    return clean(type(value).__name__)


def log_line(session, message):
    text = str(message)
    try:
        window = session.ListingWindow
        window.Open()
        writer = getattr(window, "WriteFullline", None)
        if not callable(writer):
            writer = getattr(window, "WriteLine", None)
        if callable(writer):
            for line in text.splitlines() or [""]:
                writer(line)
    except Exception:
        pass
    try:
        print(text)
    except Exception:
        pass


def io_root():
    configured = clean(os.environ.get("NX_JOURNALS_IO_DIR"))
    if configured:
        return os.path.abspath(os.path.expanduser(configured))
    profile = clean(os.environ.get("USERPROFILE"))
    if profile:
        return os.path.join(profile, "Desktop")
    fallback = os.path.expanduser("~")
    if fallback and fallback != "~":
        return os.path.join(fallback, "Desktop")
    return os.getcwd()


def filename_token(value):
    text = clean(value) or "UNKNOWN"
    return "".join(
        character if character.isalnum() or character in "-_" else "_"
        for character in text
    )[:80]


def canonical_sha256(value):
    payload = json.dumps(
        value,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    ).encode("utf-8")
    return hashlib.sha256(payload).hexdigest()


def file_sha256(path):
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        while True:
            block = handle.read(1024 * 1024)
            if not block:
                break
            digest.update(block)
    return digest.hexdigest()


def attribute_type_name(info):
    return enum_text(safe_value(info, "Type")) or "UNKNOWN"


def attribute_value(info):
    type_name = attribute_type_name(info).upper()
    property_by_type = {
        "BOOLEAN": "BooleanValue",
        "INTEGER": "IntegerValue",
        "REAL": "RealValue",
        "STRING": "StringValue",
        "TIME": "TimeValue",
        "REFERENCE": "ReferenceValue",
    }
    property_name = property_by_type.get(type_name)
    if property_name:
        try:
            return getattr(info, property_name)
        except Exception:
            return ""
    try:
        return clean(info.ToString())
    except Exception:
        return ""


def attribute_info_dict(info):
    array_index = safe_value(info, "ArrayElementIndex", -1)
    try:
        array_index = int(array_index)
    except Exception:
        array_index = -1
    return {
        "title": clean(safe_value(info, "Title")),
        "title_alias": clean(safe_value(info, "TitleAlias")),
        "category": clean(safe_value(info, "Category")),
        "type": attribute_type_name(info),
        "value": attribute_value(info),
        "unset": bool(safe_value(info, "Unset", False)),
        "locked": bool(safe_value(info, "Locked", False)),
        "inherited": bool(safe_value(info, "Inherited", False)),
        "is_override": bool(safe_value(info, "IsOverride", False)),
        "owned_by_system": bool(safe_value(info, "OwnedBySystem", False)),
        "pdm_based": bool(safe_value(info, "PdmBased", False)),
        "not_saved": bool(safe_value(info, "NotSaved", False)),
        "array": bool(safe_value(info, "Array", False)),
        "array_element_index": array_index,
    }


def instance_attribute_inventory(component):
    source = "{0}.GetInstanceUserAttributes(False)".format(
        object_kind(component) or "Component"
    )
    try:
        method = getattr(component, "GetInstanceUserAttributes")
    except AttributeError:
        return unavailable(source, "Instance-attribute enumeration is unavailable.")
    except Exception as error:
        return failed(source, error)
    try:
        try:
            values = method(False)
        except TypeError:
            values = method()
        inventory = [attribute_info_dict(info) for info in list(values)]
        inventory.sort(
            key=lambda item: (
                item["title"].upper(),
                item["array_element_index"],
                item["type"],
            )
        )
        return observed(inventory, source)
    except Exception as error:
        return failed(source, error)


def value_state(present, raw_value):
    if not present:
        return "ABSENT"
    if raw_value == "":
        return "PRESENT_BLANK"
    if clean(raw_value) in LEGACY_ACTIVE_VALUES:
        return "PRESENT_STANDARD"
    return "PRESENT_NONSTANDARD"


def control_probes(inventory_probe):
    controls = {}
    if inventory_probe["status"] != OBSERVED:
        for title in CONTROL_ATTRIBUTES:
            controls[title] = {
                "present": None,
                "raw_value": "",
                "value_state": "UNREADABLE",
                "status": inventory_probe["status"],
                "source": inventory_probe["source"],
                "error": inventory_probe["error"],
            }
        return controls

    by_title = {}
    for item in inventory_probe["value"]:
        title = clean(item.get("title")).upper()
        if title and not item.get("unset"):
            by_title.setdefault(title, []).append(item)

    for title in CONTROL_ATTRIBUTES:
        matches = by_title.get(title.upper(), [])
        present = bool(matches)
        raw = ""
        if matches:
            raw = clean(matches[0].get("value"))
            if len(matches) > 1:
                raw = " | ".join(clean(item.get("value")) for item in matches)
        controls[title] = {
            "present": present,
            "raw_value": raw,
            "value_state": value_state(present, raw),
            "status": OBSERVED,
            "source": inventory_probe["source"],
            "error": "",
        }
    return controls


def direct_control_classification(controls):
    if any(controls[title]["status"] != OBSERVED for title in CONTROL_ATTRIBUTES):
        return "UNREADABLE"
    reference = controls["REFERENCE_COMPONENT"]["present"]
    member = controls["PLIST_IGNORE_MEMBER"]["present"]
    subassembly = controls["PLIST_IGNORE_SUBASSEMBLY"]["present"]
    if reference and (member or subassembly):
        return "MULTIPLE_CONTROLS"
    if reference:
        return "REFERENCE_ONLY"
    if member and subassembly:
        return "NATIVE_EXCLUDE_PAIR"
    if member:
        return "NATIVE_MEMBER_ONLY"
    if subassembly:
        return "NATIVE_SUBASSEMBLY_ONLY"
    return "NONE"


def component_string(component, property_name):
    result = property_probe(component, property_name)
    if result["status"] == OBSERVED:
        result["value"] = clean(result["value"])
    return result


def prototype_string_attribute(prototype, title):
    source = "{0}.GetStringAttribute({1})".format(
        object_kind(prototype) or "prototype", title
    )
    if prototype is None:
        return unavailable(source, "Prototype is unavailable.")
    try:
        return observed(clean(prototype.GetStringAttribute(title)), source)
    except Exception as first_error:
        try:
            info = prototype.GetUserAttribute(
                title,
                NXOpen.NXObject.AttributeType.String,
                -1,
            )
            return observed(clean(getattr(info, "StringValue", "")), source)
        except Exception:
            return unavailable(source, error_text(first_error))


def stable_instance_id_probe(component, uf_session):
    source = "UFSession.Assem.AskStableIdOfInstance"
    if uf_session is None:
        return unavailable(source, "UF session is unavailable.")
    tag = safe_value(component, "Tag", None)
    if tag is None:
        return unavailable(source, "Component tag is unavailable.")
    try:
        method = uf_session.Assem.AskStableIdOfInstance
    except AttributeError:
        return unavailable(source, "Method is not exposed by this runtime object type.")
    except Exception as error:
        return failed(source, error)

    # The UF wrapper rejects the raw Tag on some NX builds (NX error 650004
    # "Incorrect object for this operation"), so marshal the tag to the
    # integer form the UF API expects before calling.
    candidates = []
    for raw in (tag,):
        if raw not in candidates:
            candidates.append(raw)
        converted = None
        try:
            converted = int(raw)
        except Exception:
            converted = None
        if converted is not None and converted not in candidates:
            candidates.append(converted)
        if converted is None:
            for attribute_name in ("Value", "Tag", "Handle"):
                try:
                    nested = getattr(raw, attribute_name)
                except Exception:
                    continue
                try:
                    nested_int = int(nested)
                except Exception:
                    continue
                if nested_int not in candidates:
                    candidates.append(nested_int)

    last_error = None
    for candidate in candidates:
        try:
            result = method(candidate)
        except TypeError:
            try:
                result = method(candidate, "")
            except Exception as error:
                last_error = error
                continue
        except Exception as error:
            last_error = error
            continue
        if isinstance(result, (tuple, list)):
            result = next((item for item in result if clean(item)), "")
        text = clean(result)
        if text:
            return observed(text, source)
        last_error = ValueError("NX returned no stable instance ID.")
    if last_error is not None:
        return failed(source, last_error)
    return unavailable(source, "NX returned no stable instance ID.")


def keyword_match(component):
    combined = "{0} {1}".format(
        clean(safe_value(component, "Name")),
        clean(safe_value(component, "DisplayName")),
    ).upper()
    return next((word for word in LEGACY_IGNORE_KEYWORDS if word in combined), "")


def boolean_csv(value):
    if value is None:
        return ""
    return "YES" if bool(value) else "NO"


def legacy_flag_is_active(control):
    return (
        control["status"] == OBSERVED
        and bool(control["present"])
        and control["raw_value"] in LEGACY_ACTIVE_VALUES
    )


def direct_legacy_prediction(level, suppression, keyword, controls):
    if level == 0:
        return "INCLUDE_ROOT", ""
    if suppression["status"] != OBSERVED:
        return "UNDETERMINED_SUPPRESSION_READ", "SUPPRESSION"
    if bool(suppression["value"]):
        return "EXCLUDE_SUPPRESSED_SUBTREE", "SUPPRESSION"
    if keyword:
        return "EXCLUDE_NAME_KEYWORD_SUBTREE", "NAME_KEYWORD:{0}".format(keyword)
    if legacy_flag_is_active(controls["REFERENCE_COMPONENT"]):
        return "EXCLUDE_REFERENCE_SUBTREE", "REFERENCE_COMPONENT"
    if legacy_flag_is_active(controls["PLIST_IGNORE_MEMBER"]):
        return "EXCLUDE_PLIST_MEMBER_SUBTREE", "PLIST_IGNORE_MEMBER"
    if any(controls[title]["status"] != OBSERVED for title in CONTROL_ATTRIBUTES):
        return "UNDETERMINED_CONTROL_READ", "CONTROL_ATTRIBUTE"
    return "INCLUDE", ""


def structural_segment(sibling_index, component):
    name = safe_name(component, "<component>")
    escaped = name.replace("\\", "\\\\").replace("/", "\\/")
    return "{0:06d}[{1}]".format(sibling_index, escaped)


def row_probe_errors(probes, controls):
    errors = []
    for label, item in probes:
        if item["status"] == ERROR:
            errors.append("{0}: {1}".format(label, item["error"]))
    for title in CONTROL_ATTRIBUTES:
        item = controls[title]
        if item["status"] != OBSERVED:
            errors.append("{0}: {1}".format(title, item["error"] or item["status"]))
    return list(dict.fromkeys(errors))


def critical_row_errors(probes, controls):
    """Return failures that make BoM classification evidence incomplete."""
    errors = []
    for label, item in probes:
        if item["status"] != OBSERVED:
            errors.append(
                "{0}: {1}".format(label, item["error"] or item["status"])
            )
    for title in CONTROL_ATTRIBUTES:
        item = controls[title]
        if item["status"] != OBSERVED:
            errors.append("{0}: {1}".format(title, item["error"] or item["status"]))
    return list(dict.fromkeys(errors))


def create_occurrence_row(
    component,
    work_part,
    uf_session,
    sequence,
    level,
    sibling_index,
    parent_path,
    parent_display_path,
    nearest_control_ancestor,
    legacy_controlling_path,
    run_id,
    run_timestamp,
):
    segment = structural_segment(sibling_index, component)
    structural_path = segment if not parent_path else parent_path + "/" + segment
    component_name = clean(safe_value(component, "Name"))
    display_name = safe_name(component, "<component>")
    display_path = display_name if not parent_display_path else parent_display_path + " / " + display_name

    prototype_probe = property_probe(component, "Prototype")
    prototype = prototype_probe["value"] if prototype_probe["status"] == OBSERVED else None
    if level == 0 and prototype is None:
        prototype = work_part
        prototype_probe = observed(work_part, "Root component fallback to work part")

    inventory_probe = instance_attribute_inventory(component)
    controls = control_probes(inventory_probe)
    classification = direct_control_classification(controls)
    is_direct_control = any(
        controls[title]["status"] == OBSERVED and controls[title]["present"]
        for title in CONTROL_ATTRIBUTES
    )

    suppression = property_probe(component, "IsSuppressed")
    reference_set = component_string(component, "ReferenceSet")
    non_geometric = method_probe(component, "GetNonGeometricState")
    representation = method_probe(component, "GetComponentRepresentationMode")
    if representation["status"] == OBSERVED:
        representation["value"] = enum_text(representation["value"])
    layer = property_probe(component, "Layer")
    stable_id = stable_instance_id_probe(component, uf_session)
    load_state = property_probe(prototype, "PartLoadState")
    if load_state["status"] == OBSERVED:
        load_state["value"] = enum_text(load_state["value"])

    part_number = prototype_string_attribute(prototype, "DB_PART_NO")
    revision = prototype_string_attribute(prototype, "DB_PART_REV")
    part_name = prototype_string_attribute(prototype, "DB_PART_NAME")
    stocking_type = prototype_string_attribute(prototype, "Stocking_Type")
    keyword = keyword_match(component)

    if legacy_controlling_path:
        prediction = "EXCLUDED_BY_ANCESTOR"
        prediction_control = legacy_controlling_path
        next_legacy_control = legacy_controlling_path
    else:
        prediction, direct_reason = direct_legacy_prediction(
            level, suppression, keyword, controls
        )
        prediction_control = structural_path if prediction.startswith("EXCLUDE_") else direct_reason
        next_legacy_control = (
            structural_path if prediction.startswith("EXCLUDE_") else ""
        )

    all_probes = (
        ("Prototype", prototype_probe),
        ("Suppression", suppression),
        ("ReferenceSet", reference_set),
        ("NonGeometric", non_geometric),
        ("RepresentationMode", representation),
        ("ComponentLayer", layer),
        ("StableInstanceId", stable_id),
        ("PrototypeLoadState", load_state),
        ("DB_PART_NO", part_number),
        ("DB_PART_REV", revision),
        ("DB_PART_NAME", part_name),
        ("Stocking_Type", stocking_type),
        ("InstanceAttributeInventory", inventory_probe),
    )
    probe_errors = row_probe_errors(
        all_probes,
        controls,
    )
    critical_errors = critical_row_errors(
        (
            ("Prototype", prototype_probe),
            ("Suppression", suppression),
            ("InstanceAttributeInventory", inventory_probe),
        ),
        controls,
    )
    if level > 0 and prototype is None:
        critical_errors.append("Prototype: NX returned no prototype object.")
    probe_errors = list(dict.fromkeys(probe_errors + critical_errors))
    evidence_status = "ERROR" if critical_errors else (
        "PARTIAL" if any(
            item["status"] == UNAVAILABLE
            for item in (
                prototype_probe,
                suppression,
                reference_set,
                non_geometric,
                representation,
                layer,
                stable_id,
                load_state,
                part_number,
                revision,
                part_name,
                stocking_type,
            )
        ) or probe_errors else "COMPLETE"
    )

    row = {column: "" for column in CSV_COLUMNS}
    row.update(
        {
            "SCHEMA_VERSION": SCHEMA_VERSION,
            "RUN_ID": run_id,
            "RUN_TIMESTAMP": run_timestamp,
            "JOURNAL_BUILD": BUILD,
            "SEQUENCE": sequence,
            "LEVEL": level,
            "STRUCTURAL_PATH": structural_path,
            "DISPLAY_PATH": display_path,
            "PARENT_PATH": parent_path,
            "SIBLING_INDEX": sibling_index,
            "COMPONENT_NAME": component_name,
            "COMPONENT_DISPLAY_NAME": display_name,
            "COMPONENT_TAG": object_tag(component),
            "COMPONENT_JOURNAL_IDENTIFIER": clean(
                safe_value(component, "JournalIdentifier")
            ),
            "INSTANCE_STABLE_ID_STATUS": stable_id["status"],
            "INSTANCE_STABLE_ID": clean(stable_id["value"]),
            "INSTANCE_STABLE_ID_ERROR": stable_id["error"],
            "PROTOTYPE_STATUS": prototype_probe["status"],
            "PROTOTYPE_TYPE": object_kind(prototype),
            "PROTOTYPE_TAG": object_tag(prototype),
            "PROTOTYPE_NAME": safe_name(prototype, ""),
            "PROTOTYPE_PATH": clean(safe_value(prototype, "FullPath")),
            "PROTOTYPE_LOAD_STATE_STATUS": load_state["status"],
            "PROTOTYPE_LOAD_STATE": clean(load_state["value"]),
            "DB_PART_NO_STATUS": part_number["status"],
            "DB_PART_NO": clean(part_number["value"]),
            "DB_PART_REV_STATUS": revision["status"],
            "DB_PART_REV": clean(revision["value"]),
            "DB_PART_NAME_STATUS": part_name["status"],
            "DB_PART_NAME": clean(part_name["value"]),
            "STOCKING_TYPE_STATUS": stocking_type["status"],
            "STOCKING_TYPE": clean(stocking_type["value"]),
            "SUPPRESSED_STATUS": suppression["status"],
            "SUPPRESSED": boolean_csv(suppression["value"]),
            "REFERENCE_SET_STATUS": reference_set["status"],
            "REFERENCE_SET": clean(reference_set["value"]),
            "NON_GEOMETRIC_STATUS": non_geometric["status"],
            "NON_GEOMETRIC": boolean_csv(non_geometric["value"]),
            "REPRESENTATION_MODE_STATUS": representation["status"],
            "REPRESENTATION_MODE": clean(representation["value"]),
            "COMPONENT_LAYER_STATUS": layer["status"],
            "COMPONENT_LAYER": clean(layer["value"]),
            "DIRECT_CONTROL_CLASSIFICATION": classification,
            "NEAREST_CONTROL_ANCESTOR_PATH": nearest_control_ancestor,
            "LEGACY_IGNORE_KEYWORD_MATCH": keyword,
            "CURRENT_EXTENDED_BOM_PREDICTION": prediction,
            "CURRENT_EXTENDED_BOM_CONTROLLING_PATH": prediction_control,
            "ROW_EVIDENCE_STATUS": evidence_status,
            "PROBE_ERRORS": " | ".join(probe_errors),
        }
    )
    for title in CONTROL_ATTRIBUTES:
        prefix = title
        control = controls[title]
        row[prefix + "_PRESENT"] = boolean_csv(control["present"])
        row[prefix + "_RAW_VALUE"] = control["raw_value"]
        row[prefix + "_VALUE_STATE"] = control["value_state"]
        row[prefix + "_READ_STATUS"] = control["status"]
        row[prefix + "_READ_ERROR"] = control["error"]

    row["_component"] = component
    row["_attribute_inventory"] = (
        inventory_probe["value"] if inventory_probe["status"] == OBSERVED else []
    )
    row["_controls"] = controls
    row["_is_direct_control"] = is_direct_control
    row["_next_control_ancestor"] = (
        structural_path if is_direct_control else nearest_control_ancestor
    )
    row["_next_legacy_control"] = next_legacy_control
    row["_read_error_items"] = probe_errors
    row["_critical_read_error_items"] = critical_errors
    return row


def collect_occurrences(work_part, uf_session, run_id, run_timestamp, progress=None):
    traversal_errors = []
    try:
        assembly = work_part.ComponentAssembly
        root = assembly.RootComponent
    except Exception as error:
        raise RuntimeError("RootComponent: {0}".format(error_text(error)))
    if root is None:
        raise RuntimeError("The active work part is not an assembly.")

    rows = []
    stack = [(root, 0, 0, "", "", "", "")]
    safety_limit_reached = False
    while stack:
        if len(rows) >= MAX_OCCURRENCES:
            safety_limit_reached = True
            traversal_errors.append(
                {
                    "path": rows[-1]["STRUCTURAL_PATH"] if rows else "",
                    "operation": "Traversal safety limit",
                    "status": ERROR,
                    "error": "Safety limit reached: {0}".format(MAX_OCCURRENCES),
                }
            )
            break

        (
            component,
            level,
            sibling_index,
            parent_path,
            parent_display_path,
            nearest_control_ancestor,
            legacy_controlling_path,
        ) = stack.pop()
        row = create_occurrence_row(
            component,
            work_part,
            uf_session,
            len(rows) + 1,
            level,
            sibling_index,
            parent_path,
            parent_display_path,
            nearest_control_ancestor,
            legacy_controlling_path,
            run_id,
            run_timestamp,
        )
        rows.append(row)
        if progress and len(rows) % PROGRESS_INTERVAL == 0:
            progress(len(rows), row["DISPLAY_PATH"])

        children_probe = method_probe(component, "GetChildren")
        if children_probe["status"] != OBSERVED:
            traversal_errors.append(
                {
                    "path": row["STRUCTURAL_PATH"],
                    "operation": children_probe["source"],
                    "status": children_probe["status"],
                    "error": children_probe["error"],
                }
            )
            continue
        try:
            children = list(children_probe["value"])
        except Exception as error:
            traversal_errors.append(
                {
                    "path": row["STRUCTURAL_PATH"],
                    "operation": "GetChildren result enumeration",
                    "status": ERROR,
                    "error": error_text(error),
                }
            )
            continue
        for child_index in range(len(children) - 1, -1, -1):
            stack.append(
                (
                    children[child_index],
                    level + 1,
                    child_index,
                    row["STRUCTURAL_PATH"],
                    row["DISPLAY_PATH"],
                    row["_next_control_ancestor"],
                    row["_next_legacy_control"],
                )
            )

    return rows, traversal_errors, safety_limit_reached


def control_descendant_counts(rows):
    result = []
    open_controls = []
    for index, row in enumerate(rows):
        while open_controls and int(row["LEVEL"]) <= open_controls[-1][0]:
            _level, control_index = open_controls.pop()
            control = rows[control_index]
            result.append(
                {
                    "sequence": control["SEQUENCE"],
                    "structural_path": control["STRUCTURAL_PATH"],
                    "display_path": control["DISPLAY_PATH"],
                    "classification": control["DIRECT_CONTROL_CLASSIFICATION"],
                    "descendant_count": index - control_index - 1,
                }
            )
        if row.get("_is_direct_control"):
            open_controls.append((int(row["LEVEL"]), index))
    while open_controls:
        _level, control_index = open_controls.pop()
        control = rows[control_index]
        result.append(
            {
                "sequence": control["SEQUENCE"],
                "structural_path": control["STRUCTURAL_PATH"],
                "display_path": control["DISPLAY_PATH"],
                "classification": control["DIRECT_CONTROL_CLASSIFICATION"],
                "descendant_count": len(rows) - control_index - 1,
            }
        )
    result.sort(key=lambda item: item["sequence"])
    return result


def classification_counts(rows):
    counts = {}
    for row in rows:
        key = row["DIRECT_CONTROL_CLASSIFICATION"]
        counts[key] = counts.get(key, 0) + 1
    return dict(sorted(counts.items()))


def read_error_records(rows):
    records = []
    for row in rows:
        if not row.get("_read_error_items"):
            continue
        records.append(
            {
                "sequence": row["SEQUENCE"],
                "structural_path": row["STRUCTURAL_PATH"],
                "display_path": row["DISPLAY_PATH"],
                "errors": list(row["_read_error_items"]),
            }
        )
    return records


def critical_read_error_records(rows):
    records = []
    for row in rows:
        if not row.get("_critical_read_error_items"):
            continue
        records.append(
            {
                "sequence": row["SEQUENCE"],
                "structural_path": row["STRUCTURAL_PATH"],
                "display_path": row["DISPLAY_PATH"],
                "errors": list(row["_critical_read_error_items"]),
            }
        )
    return records


def flagged_occurrence_records(rows):
    records = []
    flagged_classes = {
        "REFERENCE_ONLY",
        "NATIVE_MEMBER_ONLY",
        "NATIVE_SUBASSEMBLY_ONLY",
        "NATIVE_EXCLUDE_PAIR",
        "MULTIPLE_CONTROLS",
        "UNREADABLE",
    }
    for row in rows:
        if (
            row["DIRECT_CONTROL_CLASSIFICATION"] not in flagged_classes
            and not row.get("_critical_read_error_items")
        ):
            continue
        records.append(
            {
                "sequence": row["SEQUENCE"],
                "structural_path": row["STRUCTURAL_PATH"],
                "display_path": row["DISPLAY_PATH"],
                "direct_control_classification": row[
                    "DIRECT_CONTROL_CLASSIFICATION"
                ],
                "controls": {
                    title: dict(row["_controls"][title])
                    for title in CONTROL_ATTRIBUTES
                },
                "probe_errors": list(row.get("_read_error_items", [])),
                "instance_attributes": list(row.get("_attribute_inventory", [])),
            }
        )
    return records


def csv_public_row(row):
    return {column: row.get(column, "") for column in CSV_COLUMNS}


def write_csv(path, rows):
    with open(path, "w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(handle, fieldnames=CSV_COLUMNS, extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            writer.writerow(csv_public_row(row))


def write_json(path, report):
    with open(path, "w", encoding="utf-8") as handle:
        json.dump(
            report,
            handle,
            ensure_ascii=False,
            indent=2,
            sort_keys=True,
        )
        handle.write("\n")


def write_artifacts(folder, stem, rows, report):
    os.makedirs(folder, exist_ok=True)
    csv_final = os.path.join(folder, stem + ".csv")
    json_final = os.path.join(folder, stem + ".json")
    csv_partial = csv_final + ".partial"
    json_partial = json_final + ".partial"

    write_csv(csv_partial, rows)
    report["csv_sha256"] = file_sha256(csv_partial)
    write_json(json_partial, report)
    os.replace(csv_partial, csv_final)
    os.replace(json_partial, json_final)
    return csv_final, json_final


def write_failed_json(folder, stem, report, partial=False):
    os.makedirs(folder, exist_ok=True)
    suffix = ".json.partial" if partial else ".json"
    path = os.path.join(folder, stem + suffix)
    write_json(path, report)
    return path


def nx_runtime_metadata(session):
    release_name = property_probe(session, "ReleaseName")
    release_number = property_probe(session, "ReleaseNumber")
    build_number = property_probe(session, "BuildNumber")
    application_name = property_probe(session, "ApplicationName")
    return {
        "release_name": clean(release_name["value"]),
        "release_name_status": release_name["status"],
        "release_number": clean(release_number["value"]),
        "release_number_status": release_number["status"],
        "build_number": clean(build_number["value"]),
        "build_number_status": build_number["status"],
        "application_name": clean(application_name["value"]),
        "application_name_status": application_name["status"],
    }


def modified_state(part):
    result = property_probe(part, "IsModified")
    return {
        "status": result["status"],
        "value": result["value"] if result["status"] == OBSERVED else None,
        "error": result["error"],
    }


def base_report(run_id, run_timestamp, session):
    return {
        "schema_version": SCHEMA_VERSION,
        "journal_build": BUILD,
        "run_id": run_id,
        "run_timestamp": run_timestamp,
        "run_status": "FAILED",
        "scope": "READ_ONLY_RAW_ACTIVE_ASSEMBLY_OCCURRENCES",
        "nx_runtime": nx_runtime_metadata(session),
        "root_assembly": {},
        "work_part_modified": {"before": {}, "after": {}, "changed": None},
        "summary": {},
        "classification_counts": {},
        "control_descendant_counts": [],
        "traversal_errors": [],
        "read_errors": [],
        "flagged_occurrences": [],
        "schema_hashes": {
            "csv_columns_sha256": canonical_sha256(list(CSV_COLUMNS)),
            "json_contract_sha256": canonical_sha256(list(JSON_CONTRACT_KEYS)),
        },
        "csv_sha256": "",
    }


def run(session, uf_session=None, run_datetime=None, output_root=None, run_id=None):
    now = run_datetime or datetime.datetime.now().astimezone()
    run_timestamp = now.isoformat(timespec="seconds")
    run_id = clean(run_id) or uuid.uuid4().hex[:8]
    report = base_report(run_id, run_timestamp, session)

    try:
        work_part = session.Parts.Work
    except Exception as error:
        work_part = None
        work_error = error_text(error)
    else:
        work_error = ""
    root_name = safe_name(work_part, "UNKNOWN")
    root_token = filename_token(root_name)
    stamp = now.strftime("%Y%m%d_%H%M%S")
    folder = os.path.join(
        os.path.abspath(output_root or io_root()),
        OUTPUT_FOLDER,
        "{0}_{1}_{2}".format(root_token, stamp, run_id),
    )
    stem = "J28_BOM_STRUCTURE_{0}_{1}".format(root_token, run_id)

    if work_part is None:
        report["fatal_error"] = work_error or "No active work part."
        json_path = write_failed_json(folder, stem, report)
        return "", json_path, report

    report["work_part_modified"]["before"] = modified_state(work_part)
    if uf_session is None:
        try:
            uf_module = __import__("NXOpen.UF", fromlist=["UFSession"])
            uf_session = uf_module.UFSession.GetUFSession()
        except Exception:
            uf_session = None

    try:
        def progress(count, path):
            log_line(session, "J28 progress: {0} occurrences; {1}".format(count, path))

        rows, traversal_errors, safety_limit_reached = collect_occurrences(
            work_part,
            uf_session,
            run_id,
            run_timestamp,
            progress=progress,
        )
        report["root_assembly"] = {
            "name": root_name,
            "part_number": rows[0]["DB_PART_NO"] if rows else "",
            "revision": rows[0]["DB_PART_REV"] if rows else "",
            "prototype_path": rows[0]["PROTOTYPE_PATH"] if rows else "",
        }
        report["traversal_errors"] = traversal_errors
        report["read_errors"] = read_error_records(rows)
        critical_read_errors = critical_read_error_records(rows)
        report["classification_counts"] = classification_counts(rows)
        report["control_descendant_counts"] = control_descendant_counts(rows)
        report["flagged_occurrences"] = flagged_occurrence_records(rows)
        report["summary"] = {
            "occurrence_count": len(rows),
            "suppressed_count": sum(
                1
                for row in rows
                if row["SUPPRESSED_STATUS"] == OBSERVED
                and row["SUPPRESSED"] == "YES"
            ),
            "flagged_occurrence_count": len(report["flagged_occurrences"]),
            "read_error_occurrence_count": len(report["read_errors"]),
            "critical_read_error_occurrence_count": len(critical_read_errors),
            "traversal_error_count": len(traversal_errors),
            "safety_limit_reached": safety_limit_reached,
        }
        report["work_part_modified"]["after"] = modified_state(work_part)
        before = report["work_part_modified"]["before"]
        after = report["work_part_modified"]["after"]
        changed = None
        if before.get("status") == OBSERVED and after.get("status") == OBSERVED:
            changed = before.get("value") != after.get("value")
        report["work_part_modified"]["changed"] = changed

        incomplete = bool(traversal_errors or critical_read_errors or changed)
        report["run_status"] = "INCOMPLETE" if incomplete else "COMPLETE"
        csv_path, json_path = write_artifacts(folder, stem, rows, report)
        return csv_path, json_path, report
    except RuntimeError as error:
        report["fatal_error"] = error_text(error)
        report["work_part_modified"]["after"] = modified_state(work_part)
        json_path = write_failed_json(folder, stem, report)
        return "", json_path, report
    except Exception as error:
        report["fatal_error"] = error_text(error)
        report["fatal_traceback"] = traceback.format_exc()
        report["work_part_modified"]["after"] = modified_state(work_part)
        write_failed_json(folder, stem, report, partial=True)
        raise


def main():
    session = NXOpen.Session.GetSession()
    log_line(session, "=" * 72)
    log_line(session, "J28 RAW BOM STRUCTURE CHECKPOINT")
    log_line(session, "Build: " + BUILD)
    log_line(
        session,
        "Read-only: no load, update, visibility, assembly, attribute, or persistence changes.",
    )
    log_line(session, "=" * 72)
    try:
        csv_path, json_path, report = run(session)
        log_line(session, "Run status: " + report["run_status"])
        log_line(
            session,
            "Occurrences: {0}".format(report.get("summary", {}).get("occurrence_count", 0)),
        )
        if csv_path:
            log_line(session, "CSV: " + csv_path)
        log_line(session, "JSON: " + json_path)
        if report["run_status"] != "COMPLETE":
            log_line(session, "Return both artifacts; review all incomplete evidence first.")
        else:
            log_line(session, "Return both artifacts for BoM checkpoint analysis.")
    except Exception as error:
        log_line(session, "J28 FAILED: " + error_text(error))
        log_line(session, traceback.format_exc())
        raise


def get_unload_option(dummy):
    return NXOpen.Session.LibraryUnloadOption.Immediately


if __name__ == "__main__":
    main()
