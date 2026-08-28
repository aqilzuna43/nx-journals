"""
Journal 27 - Move Active-Assembly Components to Layer 1

Normalizes the layer option of every component occurrence placed directly
under the active assembly.  It does not recurse into subassemblies and does
not modify component prototype bodies.  Blanked, suppressed, reference-only,
non-geometric, lightweight, and unloaded direct occurrences are included.

DRY_RUN is the default.  APPLY requires the assembly to be both work and
displayed part plus a writable native part or a Teamcenter part checked out
by the current user.  J27 never loads, checks out, checks in, or saves data.
A successful APPLY remains under one visible NX undo mark for inspection.

Target: NX 2312 and NX X 2506 embedded Python
Run via: NX > Tools > Journal > Play
"""

import csv
import datetime
import json
import os
import traceback

import NXOpen


# User setting. Change to "APPLY" only after reviewing a DRY_RUN report.
USER_MODE = "DRY_RUN"

BUILD = "J27-NX2506-DIRECT-COMPONENTS-TO-LAYER-1-V1"
SCHEMA_VERSION = 1
TARGET_LAYER = 1
FIRST_LAYER = 1
LAST_LAYER = 256
OUTPUT_FOLDER = "NX_ASSEMBLY_LAYER_1_MIGRATION"
UNDO_MARK_NAME = "J27 Move direct assembly components to layer 1"
VALID_MODES = ("DRY_RUN", "APPLY")

CSV_COLUMNS = (
    "RUN_TIMESTAMP",
    "JOURNAL_BUILD",
    "SCHEMA_VERSION",
    "MODE",
    "VERDICT",
    "ROW_TYPE",
    "DB_PART_NO",
    "DB_PART_REV",
    "ASSEMBLY_NAME",
    "MANAGED_MODE",
    "CHECKOUT_STATE",
    "CHECKOUT_OWNER",
    "CURRENT_USER",
    "READ_ONLY",
    "COMPONENT_INDEX",
    "COMPONENT_PATH",
    "COMPONENT_NAME",
    "COMPONENT_TAG",
    "PROTOTYPE_NAME",
    "PROTOTYPE_TAG",
    "SUPPRESSED",
    "BLANKED",
    "REFERENCE_SET",
    "NON_GEOMETRIC",
    "ORIGINAL_LAYER_OPTION",
    "ORIGINAL_REPORTED_LAYER",
    "FINAL_LAYER_OPTION",
    "FINAL_REPORTED_LAYER",
    "ACTION",
    "STATUS",
    "MESSAGE",
)

_INVALID_FILENAME_CHARS = '<>:"/\\|?*'
_INTERNAL_RECORD_KEYS = ("_object", "_tag_key")


class VerificationError(RuntimeError):
    def __init__(self, messages, snapshot=None):
        RuntimeError.__init__(self, " | ".join(messages))
        self.messages = list(messages)
        self.snapshot = snapshot


def clean(value):
    return "" if value is None else str(value).strip()


def error_text(error):
    code = clean(getattr(error, "ErrorCode", ""))
    suffix = " [{0}]".format(code) if code else ""
    return "{0}{1}".format(clean(error) or type(error).__name__, suffix)


def log_line(session, message):
    text = str(message)
    try:
        window = session.ListingWindow
        window.Open()
        for line in text.splitlines() or [""]:
            window.WriteFullline(line)
    except Exception:
        pass
    try:
        print(text)
    except Exception:
        pass


def desktop_folder():
    profile = clean(os.environ.get("USERPROFILE"))
    if profile:
        return os.path.join(profile, "Desktop")
    fallback = os.path.expanduser("~")
    if fallback and fallback != "~":
        return os.path.join(fallback, "Desktop")
    return os.getcwd()


def io_root():
    configured = clean(os.environ.get("NX_JOURNALS_IO_DIR"))
    return os.path.abspath(os.path.expanduser(configured or desktop_folder()))


def filename_token(value, fallback="UNKNOWN"):
    text = clean(value)
    if not text:
        return fallback
    result = "".join(
        "_" if char in _INVALID_FILENAME_CHARS or ord(char) < 32 else char
        for char in text
    ).strip(" .")
    return result or fallback


def configured_mode():
    mode = clean(os.environ.get("NX_J27_MODE") or USER_MODE).upper()
    if mode not in VALID_MODES:
        raise RuntimeError(
            "Invalid J27 mode {0!r}; expected DRY_RUN or APPLY.".format(mode)
        )
    return mode


def safe_property(value, name, default=None):
    try:
        result = getattr(value, name)
        return result() if callable(result) else result
    except Exception:
        return default


def required_property(value, name, label):
    try:
        result = getattr(value, name)
        return result() if callable(result) else result
    except Exception as error:
        raise RuntimeError(
            "Could not inspect {0}.{1}: {2}".format(
                label, name, error_text(error)
            )
        )


def required_call(value, name, label, *args):
    method = getattr(value, name, None)
    if not callable(method):
        raise RuntimeError("{0}.{1} is unavailable.".format(label, name))
    try:
        return method(*args)
    except Exception as error:
        raise RuntimeError(
            "Could not call {0}.{1}: {2}".format(label, name, error_text(error))
        )


def get_string_attribute(nx_object, name):
    try:
        return clean(nx_object.GetStringAttribute(name))
    except Exception:
        pass
    try:
        info = nx_object.GetUserAttribute(
            name, NXOpen.NXObject.AttributeType.String, -1
        )
        return clean(info.StringValue)
    except Exception:
        return ""


def safe_name(value, fallback="UNKNOWN"):
    for property_name in ("Name", "Leaf", "FullPath", "JournalIdentifier"):
        result = clean(safe_property(value, property_name))
        if result:
            return result
    return fallback


def part_identity(part):
    if part is None:
        return {"number": "", "revision": "", "name": "UNKNOWN"}
    return {
        "number": (
            get_string_attribute(part, "DB_PART_NO")
            or get_string_attribute(part, "PART_NUMBER")
            or get_string_attribute(part, "ITEM_ID")
        ),
        "revision": (
            get_string_attribute(part, "DB_PART_REV")
            or get_string_attribute(part, "REVISION")
            or get_string_attribute(part, "ITEM_REVISION")
        ),
        "name": safe_name(part),
    }


def object_tag(value, required=False, label="object"):
    try:
        tag = getattr(value, "Tag")
        tag = tag() if callable(tag) else tag
        text = clean(tag)
    except Exception as error:
        if required:
            raise RuntimeError(
                "Could not inspect {0}.Tag: {1}".format(label, error_text(error))
            )
        return ""
    if required and (not text or text == "0"):
        raise RuntimeError("{0} has no usable NX tag.".format(label))
    return text


def same_nx_object(left, right):
    if left is right:
        return True
    left_tag = object_tag(left)
    right_tag = object_tag(right)
    return bool(left_tag and right_tag and left_tag == right_tag)


def state_token(value):
    name = clean(getattr(value, "name", ""))
    return name or clean(value)


def read_layer_snapshot(work_part):
    layers = required_property(work_part, "Layers", "work part")
    work_layer = int(required_property(layers, "WorkLayer", "work_part.Layers"))
    if not FIRST_LAYER <= work_layer <= LAST_LAYER:
        raise RuntimeError("NX reported invalid work layer {0}.".format(work_layer))
    states = {}
    for layer in range(FIRST_LAYER, LAST_LAYER + 1):
        try:
            states[str(layer)] = state_token(layers.GetState(layer))
        except Exception as error:
            raise RuntimeError(
                "Could not inspect state of layer {0}: {1}".format(
                    layer, error_text(error)
                )
            )
    return {"work_layer": work_layer, "states": states}


def point_tuple(point):
    try:
        return [float(point.X), float(point.Y), float(point.Z)]
    except Exception as error:
        raise RuntimeError("Invalid component position: " + error_text(error))


def matrix_tuple(matrix):
    names = ("Xx", "Xy", "Xz", "Yx", "Yy", "Yz", "Zx", "Zy", "Zz")
    try:
        return [float(getattr(matrix, name)) for name in names]
    except Exception as error:
        raise RuntimeError("Invalid component orientation: " + error_text(error))


def component_transform(component, label):
    result = required_call(component, "GetPosition", label)
    if not isinstance(result, tuple) or len(result) < 2:
        raise RuntimeError(
            "{0}.GetPosition did not return position and orientation.".format(label)
        )
    return {
        "position": point_tuple(result[0]),
        "orientation": matrix_tuple(result[1]),
    }


def prototype_record(component):
    prototype = safe_property(component, "Prototype")
    if prototype is None:
        return {"available": False, "tag": "", "name": ""}
    return {
        "available": True,
        "tag": object_tag(prototype),
        "name": safe_name(prototype, "<unavailable>"),
    }


def inspect_component(component, index, root):
    label = "component {0}".format(index)
    tag = object_tag(component, required=True, label=label)
    parent = required_property(component, "Parent", label)
    if parent is None or not same_nx_object(parent, root):
        raise RuntimeError(
            "{0} (tag {1}) is not a direct child of the active assembly root.".format(
                label, tag
            )
        )
    layer_option = int(required_call(component, "GetLayerOption", label))
    reported_layer = int(required_property(component, "Layer", label))
    if layer_option < -1 or layer_option > LAST_LAYER:
        raise RuntimeError(
            "{0} has invalid layer option {1}.".format(label, layer_option)
        )
    if not FIRST_LAYER <= reported_layer <= LAST_LAYER:
        raise RuntimeError(
            "{0} has invalid reported layer {1}.".format(label, reported_layer)
        )
    name = clean(required_property(component, "Name", label)) or "COMPONENT_{0}".format(index)
    return {
        "index": index,
        "path": name,
        "name": name,
        "tag": tag,
        "prototype": prototype_record(component),
        "parent_tag": object_tag(parent, required=True, label=label + " parent"),
        "suppressed": bool(required_property(component, "IsSuppressed", label)),
        "blanked": bool(required_property(component, "IsBlanked", label)),
        "reference_set": clean(required_property(component, "ReferenceSet", label)),
        "non_geometric": bool(required_call(component, "GetNonGeometricState", label)),
        "layer_option": layer_option,
        "reported_layer": reported_layer,
        "transform": component_transform(component, label),
        "_tag_key": tag,
        "_object": component,
    }


def assembly_context(session):
    parts = safe_property(session, "Parts")
    work_part = safe_property(parts, "Work")
    display_part = safe_property(parts, "Display")
    if work_part is None:
        raise RuntimeError("No active NX work part is available.")
    if display_part is None or not same_nx_object(work_part, display_part):
        raise RuntimeError(
            "Make the target assembly both Work Part and Displayed Part."
        )
    component_assembly = required_property(
        work_part, "ComponentAssembly", "work part"
    )
    root = required_property(
        component_assembly, "RootComponent", "work_part.ComponentAssembly"
    )
    if root is None:
        raise RuntimeError("The active work/display part is not an assembly.")
    return work_part, component_assembly, root


def capture_snapshot(work_part, root):
    children = list(required_call(root, "GetChildren", "assembly root"))
    records = []
    tags = set()
    for index, component in enumerate(children, 1):
        record = inspect_component(component, index, root)
        if record["tag"] in tags:
            raise RuntimeError(
                "Duplicate direct-component tag {0} was returned.".format(
                    record["tag"]
                )
            )
        tags.add(record["tag"])
        records.append(record)
    layer_snapshot = read_layer_snapshot(work_part)
    return {
        "components": records,
        "direct_child_tags": [item["tag"] for item in records],
        "work_layer": layer_snapshot["work_layer"],
        "layer_states": layer_snapshot["states"],
    }


def public_component_record(record):
    return {
        key: value for key, value in record.items()
        if key not in _INTERNAL_RECORD_KEYS
    }


def public_snapshot(snapshot):
    if snapshot is None:
        return None
    return {
        "components": [
            public_component_record(item) for item in snapshot["components"]
        ],
        "direct_child_tags": list(snapshot["direct_child_tags"]),
        "work_layer": snapshot["work_layer"],
        "layer_states": dict(snapshot["layer_states"]),
    }


def records_by_tag(snapshot):
    return {item["tag"]: item for item in snapshot["components"]}


def invariant_errors(before, after, require_layer_one):
    errors = []
    if before["direct_child_tags"] != after["direct_child_tags"]:
        errors.append("Direct-child membership or order changed.")
    before_by_tag = records_by_tag(before)
    after_by_tag = records_by_tag(after)
    stable_fields = (
        "name", "path", "prototype", "parent_tag", "suppressed", "blanked",
        "reference_set", "non_geometric", "transform",
    )
    for tag in sorted(set(before_by_tag) & set(after_by_tag)):
        original = before_by_tag[tag]
        final = after_by_tag[tag]
        for field in stable_fields:
            if original[field] != final[field]:
                errors.append(
                    "Component {0} field {1} changed unexpectedly.".format(
                        tag, field
                    )
                )
        if require_layer_one:
            if final["layer_option"] != TARGET_LAYER:
                errors.append(
                    "Component {0} layer option expected 1, observed {1}.".format(
                        tag, final["layer_option"]
                    )
                )
            if final["reported_layer"] != TARGET_LAYER:
                errors.append(
                    "Component {0} reported layer expected 1, observed {1}.".format(
                        tag, final["reported_layer"]
                    )
                )
        else:
            for field in ("layer_option", "reported_layer"):
                if original[field] != final[field]:
                    errors.append(
                        "Component {0} field {1} was not restored.".format(
                            tag, field
                        )
                    )
    if before["work_layer"] != after["work_layer"]:
        errors.append("The assembly work layer changed.")
    if before["layer_states"] != after["layer_states"]:
        errors.append("One or more assembly layer states changed.")
    return errors


def snapshot_counts(snapshot):
    records = snapshot["components"] if snapshot else []
    return {
        "direct_component_count": len(records),
        "move_candidate_count": sum(
            1 for item in records
            if item["layer_option"] != TARGET_LAYER
            or item["reported_layer"] != TARGET_LAYER
        ),
        "already_on_layer_1_count": sum(
            1 for item in records
            if item["layer_option"] == TARGET_LAYER
            and item["reported_layer"] == TARGET_LAYER
        ),
        "suppressed_count": sum(1 for item in records if item["suppressed"]),
        "blanked_count": sum(1 for item in records if item["blanked"]),
        "reference_set_count": sum(
            1 for item in records
            if item["reference_set"].upper() not in ("", "MODEL", "ENTIRE PART")
        ),
        "non_geometric_count": sum(
            1 for item in records if item["non_geometric"]
        ),
        "prototype_unavailable_count": sum(
            1 for item in records if not item["prototype"]["available"]
        ),
    }


def read_only_value(part):
    value = safe_property(part, "IsReadOnly")
    return None if value is None else bool(value)


def managed_mode(session, part):
    if bool(safe_property(session, "IsManagedMode", False)):
        return True
    full_path = clean(safe_property(part, "FullPath"))
    identifier = clean(safe_property(part, "JournalIdentifier"))
    return full_path.upper().startswith("@DB/") or identifier.upper().startswith("@DB/")


def checkout_result(raw):
    checked = None
    owner = ""
    if isinstance(raw, dict):
        for key in ("isCheckedOut", "is_checked_out", "checkedOut", "checked_out"):
            if isinstance(raw.get(key), bool):
                checked = raw[key]
                break
        for key in ("checkedOutBy", "checked_out_by", "owner", "user"):
            if key in raw:
                owner = clean(raw[key])
                break
    elif isinstance(raw, (tuple, list)):
        for value in raw:
            if checked is None and isinstance(value, bool):
                checked = value
            elif isinstance(value, str) and not owner:
                normalized = clean(value).upper().replace("_", "")
                if normalized.startswith("NOT") and "CHECKEDOUT" in normalized:
                    checked = False
                elif "CHECKEDOUT" in normalized:
                    checked = True
                else:
                    owner = clean(value)
            elif checked is None:
                normalized = clean(getattr(value, "name", value)).upper().replace("_", "")
                if normalized.startswith("NOT") and "CHECKEDOUT" in normalized:
                    checked = False
                elif "CHECKEDOUT" in normalized:
                    checked = True
    elif isinstance(raw, bool):
        checked = raw
    else:
        for name in ("IsCheckedOut", "isCheckedOut", "CheckedOut"):
            value = getattr(raw, name, None)
            if isinstance(value, bool):
                checked = value
                break
        for name in ("CheckedOutBy", "checkedOutBy", "Owner", "User"):
            value = getattr(raw, name, None)
            if value is not None:
                owner = clean(value)
                break
    state = "CHECKED_OUT" if checked is True else "CHECKED_IN" if checked is False else "UNKNOWN"
    return state, owner


def checkout_status(part):
    pdm_part = safe_property(part, "PDMPart")
    method = getattr(pdm_part, "GetCheckedoutStatusAndUser", None)
    if not callable(method):
        return "UNKNOWN", "", "GetCheckedoutStatusAndUser unavailable"
    try:
        raw = method()
    except TypeError:
        try:
            raw = method(False, "")
        except Exception as error:
            return "UNKNOWN", "", error_text(error)
    except Exception as error:
        return "UNKNOWN", "", error_text(error)
    state, owner = checkout_result(raw)
    return state, owner, clean(repr(raw))[:2000]


def current_teamcenter_user(session):
    pdm_session = safe_property(session, "PdmSession")
    method = getattr(pdm_session, "GetUserName", None)
    if not callable(method):
        return ""
    try:
        return clean(method())
    except Exception:
        return ""


def inspect_write_access(session, part):
    read_only = read_only_value(part)
    managed = managed_mode(session, part)
    if not managed:
        allowed = read_only is False
        if read_only is True:
            message = "Native assembly is read-only."
        elif read_only is None:
            message = "Native assembly read-only state is unavailable."
        else:
            message = ""
        return {
            "allowed": allowed, "managed": False, "checkout_state": "NATIVE",
            "checkout_owner": "", "current_user": "", "read_only": read_only,
            "raw_checkout": "", "message": message,
        }
    state, owner, raw = checkout_status(part)
    current_user = current_teamcenter_user(session)
    owner_matches = bool(
        state == "CHECKED_OUT" and owner and current_user
        and owner.casefold() == current_user.casefold()
    )
    allowed = bool(owner_matches and read_only is False)
    messages = []
    if state == "CHECKED_IN":
        messages.append("Assembly is checked in; J27 never performs checkout.")
    elif state == "UNKNOWN":
        messages.append("Checkout state is unknown: {0}.".format(raw or "<none>"))
    elif not owner:
        messages.append("Checkout owner is unavailable.")
    elif not current_user:
        messages.append("Current Teamcenter user is unavailable.")
    elif owner.casefold() != current_user.casefold():
        messages.append("Assembly is checked out by another user: {0}.".format(owner))
    if read_only is True:
        messages.append("Assembly is read-only in this NX session.")
    elif read_only is None:
        messages.append("Assembly read-only state is unavailable.")
    return {
        "allowed": allowed, "managed": True, "checkout_state": state,
        "checkout_owner": owner, "current_user": current_user,
        "read_only": read_only, "raw_checkout": raw,
        "message": " ".join(messages),
    }


def empty_access(part=None, session=None):
    return {
        "allowed": None,
        "managed": bool(part is not None and managed_mode(session, part)),
        "checkout_state": "NOT_INSPECTED", "checkout_owner": "",
        "current_user": "", "read_only": read_only_value(part) if part is not None else None,
        "raw_checkout": "",
        "message": "Write access is inspected only when APPLY has components to change.",
    }


def read_only_text(value):
    return "UNKNOWN" if value is None else "YES" if value else "NO"


def unique_run_folder(identity, now):
    root = os.path.join(io_root(), OUTPUT_FOLDER)
    os.makedirs(root, exist_ok=True)
    token = filename_token(identity.get("number") or identity.get("name") or "UNKNOWN")
    base_stem = "J27_ASSEMBLY_LAYER_1_{0}_{1}".format(
        token, now.strftime("%Y%m%d_%H%M%S")
    )
    folder = os.path.join(root, base_stem)
    suffix = 1
    while os.path.exists(folder):
        folder = os.path.join(root, "{0}_{1}".format(base_stem, suffix))
        suffix += 1
    os.makedirs(folder)
    return folder, os.path.basename(folder)


def preflight_artifact_folder(folder):
    paths = [
        os.path.join(folder, ".j27_csv_probe.tmp"),
        os.path.join(folder, ".j27_json_probe.tmp"),
    ]
    try:
        for path in paths:
            with open(path, "w", encoding="utf-8") as handle:
                handle.write("J27 evidence preflight\n")
    finally:
        for path in paths:
            try:
                if os.path.exists(path):
                    os.remove(path)
            except Exception:
                pass


def base_report(mode, now, identity, access):
    return {
        "schema_version": SCHEMA_VERSION,
        "journal_build": BUILD,
        "run_timestamp": now.isoformat(timespec="seconds"),
        "mode": mode,
        "configuration": {
            "scope": "ACTIVE_ASSEMBLY_DIRECT_COMPONENT_OCCURRENCES",
            "target_layer": TARGET_LAYER,
            "recursive": False,
            "force_load": False,
            "include_all_direct_occurrence_states": True,
            "automatic_checkout": False,
            "automatic_save": False,
            "automatic_checkin": False,
        },
        "assembly": identity,
        "access": access,
        "counts": {}, "before": None, "after_attempt": None, "after": None,
        "action": {
            "api": "Component.SetLayerOption(1)", "attempted": False,
            "candidate_tags": [], "completed_tags": [],
            "undo_mark_name": UNDO_MARK_NAME, "undo_mark_created": False,
            "successful_change_left_undoable": False,
            "verification_errors": [], "error": "",
        },
        "rollback": {
            "attempted": False, "status": "NOT_REQUIRED",
            "verification_errors": [], "error": "",
        },
        "component_results": [], "errors": [],
        "artifacts": {"folder": "", "csv": "", "json": ""},
        "verdict": {"status": "INITIALIZING", "message": "J27 has not completed."},
    }


def set_verdict(report, status, message):
    report["verdict"] = {"status": status, "message": message}


def build_component_results(report):
    before = report.get("before")
    if not before:
        return []
    final = report.get("after") or before
    final_by_tag = {item["tag"]: item for item in final.get("components", [])}
    verdict = report["verdict"]["status"]
    results = []
    for original in before["components"]:
        current = final_by_tag.get(original["tag"])
        compliant = (
            original["layer_option"] == TARGET_LAYER
            and original["reported_layer"] == TARGET_LAYER
        )
        if compliant:
            action, status, message = (
                "UNCHANGED", "ALREADY_ON_LAYER_1", "Component was already on specified layer 1."
            )
        elif report["mode"] == "DRY_RUN":
            action, status, message = (
                "WOULD_CHANGE", "DRY_RUN", "APPLY would set this component layer option to 1."
            )
        elif verdict == "APPLIED_VERIFIED":
            action, status, message = (
                "CHANGED", "APPLIED_VERIFIED", "Component layer option was set to 1 and verified."
            )
        elif verdict in ("ROLLED_BACK", "ROLLBACK_FAILED"):
            action, status, message = (
                "ROLLBACK_ATTEMPTED", verdict, report["verdict"]["message"]
            )
        else:
            action, status, message = (
                "NOT_CHANGED", verdict, report["verdict"]["message"]
            )
        results.append({
            "index": original["index"], "path": original["path"],
            "name": original["name"], "tag": original["tag"],
            "prototype": original["prototype"],
            "suppressed": original["suppressed"], "blanked": original["blanked"],
            "reference_set": original["reference_set"],
            "non_geometric": original["non_geometric"],
            "original_layer_option": original["layer_option"],
            "original_reported_layer": original["reported_layer"],
            "final_layer_option": current["layer_option"] if current else None,
            "final_reported_layer": current["reported_layer"] if current else None,
            "action": action, "status": status, "message": message,
        })
    return results


def csv_rows(report):
    identity = report["assembly"]
    access = report["access"]
    common = {
        "RUN_TIMESTAMP": report["run_timestamp"],
        "JOURNAL_BUILD": report["journal_build"],
        "SCHEMA_VERSION": report["schema_version"], "MODE": report["mode"],
        "VERDICT": report["verdict"]["status"],
        "DB_PART_NO": identity.get("number", ""),
        "DB_PART_REV": identity.get("revision", ""),
        "ASSEMBLY_NAME": identity.get("name", ""),
        "MANAGED_MODE": "YES" if access.get("managed") else "NO",
        "CHECKOUT_STATE": access.get("checkout_state", ""),
        "CHECKOUT_OWNER": access.get("checkout_owner", ""),
        "CURRENT_USER": access.get("current_user", ""),
        "READ_ONLY": read_only_text(access.get("read_only")),
    }
    summary = dict(common)
    summary.update({
        "ROW_TYPE": "SUMMARY", "ACTION": "RUN_SUMMARY",
        "STATUS": report["verdict"]["status"],
        "MESSAGE": report["verdict"]["message"],
    })
    rows = [summary]
    for component in report["component_results"]:
        prototype = component["prototype"]
        row = dict(common)
        row.update({
            "ROW_TYPE": "COMPONENT", "COMPONENT_INDEX": component["index"],
            "COMPONENT_PATH": component["path"], "COMPONENT_NAME": component["name"],
            "COMPONENT_TAG": component["tag"],
            "PROTOTYPE_NAME": prototype["name"], "PROTOTYPE_TAG": prototype["tag"],
            "SUPPRESSED": "YES" if component["suppressed"] else "NO",
            "BLANKED": "YES" if component["blanked"] else "NO",
            "REFERENCE_SET": component["reference_set"],
            "NON_GEOMETRIC": "YES" if component["non_geometric"] else "NO",
            "ORIGINAL_LAYER_OPTION": component["original_layer_option"],
            "ORIGINAL_REPORTED_LAYER": component["original_reported_layer"],
            "FINAL_LAYER_OPTION": "" if component["final_layer_option"] is None else component["final_layer_option"],
            "FINAL_REPORTED_LAYER": "" if component["final_reported_layer"] is None else component["final_reported_layer"],
            "ACTION": component["action"], "STATUS": component["status"],
            "MESSAGE": component["message"],
        })
        rows.append(row)
    return rows


def write_outputs(report, folder, stem):
    csv_path = os.path.join(folder, stem + ".csv")
    json_path = os.path.join(folder, stem + ".json")
    csv_temp, json_temp = csv_path + ".tmp", json_path + ".tmp"
    report["artifacts"] = {"folder": folder, "csv": csv_path, "json": json_path}
    try:
        with open(csv_temp, "w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=CSV_COLUMNS, extrasaction="ignore")
            writer.writeheader()
            writer.writerows(csv_rows(report))
        with open(json_temp, "w", encoding="utf-8") as handle:
            json.dump(report, handle, indent=2, ensure_ascii=False)
        os.replace(csv_temp, csv_path)
        os.replace(json_temp, json_path)
    except Exception:
        for path in (csv_temp, json_temp, csv_path, json_path):
            try:
                if os.path.exists(path):
                    os.remove(path)
            except Exception:
                pass
        report["artifacts"] = {"folder": folder, "csv": "", "json": ""}
        raise
    return csv_path, json_path


def delete_undo_mark(session, mark):
    try:
        session.DeleteUndoMark(mark, UNDO_MARK_NAME)
        return ""
    except Exception as error:
        return error_text(error)


def rollback_to_before(session, mark, work_part, root, before):
    result = {
        "attempted": True, "status": "ROLLBACK_FAILED",
        "verification_errors": [], "error": "",
    }
    try:
        session.UndoToMark(mark, UNDO_MARK_NAME)
    except Exception as error:
        result["error"] = "UndoToMark failed: " + error_text(error)
        try:
            after = capture_snapshot(work_part, root)
        except Exception as capture_error:
            after = None
            result["error"] += " | Rollback snapshot failed: " + error_text(capture_error)
        return result, after
    try:
        after = capture_snapshot(work_part, root)
        errors = invariant_errors(before, after, require_layer_one=False)
        result["verification_errors"] = errors
        if errors:
            result["error"] = "Rollback verification did not restore the baseline."
        else:
            result["status"] = "ROLLED_BACK"
    except Exception as error:
        after = None
        result["error"] = "Rollback snapshot failed: " + error_text(error)
    delete_error = delete_undo_mark(session, mark)
    if delete_error:
        result["error"] = " | ".join(
            item for item in (
                result["error"], "DeleteUndoMark failed: " + delete_error
            ) if item
        )
    return result, after


def perform_apply(session, work_part, root, before, report):
    candidates = [
        item for item in before["components"]
        if item["layer_option"] != TARGET_LAYER
        or item["reported_layer"] != TARGET_LAYER
    ]
    report["action"]["candidate_tags"] = [item["tag"] for item in candidates]
    try:
        mark = session.SetUndoMark(
            NXOpen.Session.MarkVisibility.Visible, UNDO_MARK_NAME
        )
        report["action"]["undo_mark_created"] = True
    except Exception as error:
        report["action"]["error"] = error_text(error)
        set_verdict(
            report, "ROLLED_BACK",
            "J27 refused to change components because it could not create an NX undo mark.",
        )
        return None
    report["action"]["attempted"] = True
    try:
        for item in candidates:
            item["_object"].SetLayerOption(TARGET_LAYER)
            report["action"]["completed_tags"].append(item["tag"])
        after = capture_snapshot(work_part, root)
        report["after_attempt"] = public_snapshot(after)
        errors = invariant_errors(before, after, require_layer_one=True)
        report["action"]["verification_errors"] = errors
        if errors:
            raise VerificationError(errors, snapshot=after)
        report["after"] = public_snapshot(after)
        report["action"]["successful_change_left_undoable"] = True
        set_verdict(
            report, "APPLIED_VERIFIED",
            "Every direct component occurrence is on specified layer 1; the assembly remains unsaved under one visible NX undo mark.",
        )
        return mark
    except Exception as error:
        report["action"]["error"] = error_text(error)
        if isinstance(error, VerificationError) and error.snapshot is not None:
            report["after_attempt"] = public_snapshot(error.snapshot)
        rollback, final_snapshot = rollback_to_before(
            session, mark, work_part, root, before
        )
        report["rollback"] = rollback
        report["after"] = public_snapshot(final_snapshot)
        if rollback["status"] == "ROLLED_BACK":
            set_verdict(
                report, "ROLLED_BACK",
                "The component-layer change failed verification or raised an NX error; the original assembly state was restored.",
            )
        else:
            set_verdict(
                report, "ROLLBACK_FAILED",
                "The component-layer change failed and J27 could not prove restoration. Use NX Undo and inspect the assembly immediately.",
            )
        return None


def write_with_apply_rollback(report, folder, stem, session, work_part, root, before, mark):
    report["component_results"] = build_component_results(report)
    try:
        return write_outputs(report, folder, stem)
    except Exception as error:
        report["errors"].append("Evidence write failed: " + error_text(error))
        if report["verdict"]["status"] != "APPLIED_VERIFIED" or mark is None:
            raise
    report["action"]["successful_change_left_undoable"] = False
    rollback, final_snapshot = rollback_to_before(
        session, mark, work_part, root, before
    )
    report["rollback"] = rollback
    report["after"] = public_snapshot(final_snapshot)
    if rollback["status"] == "ROLLED_BACK":
        set_verdict(
            report, "ROLLED_BACK",
            "The verified component-layer change was rolled back because paired CSV/JSON evidence could not be completed.",
        )
    else:
        set_verdict(
            report, "ROLLBACK_FAILED",
            "Evidence writing failed and J27 could not prove restoration. Use NX Undo and inspect the assembly immediately.",
        )
    report["component_results"] = build_component_results(report)
    try:
        return write_outputs(report, folder, stem)
    except Exception as second_error:
        report["errors"].append(
            "Rollback evidence write also failed: " + error_text(second_error)
        )
        report["artifacts"] = {"folder": folder, "csv": "", "json": ""}
        return "", ""


def context_failure(session):
    parts = safe_property(session, "Parts")
    work_part = safe_property(parts, "Work")
    display_part = safe_property(parts, "Display")
    if work_part is None:
        return "FAILED_NO_WORK_PART", "No active NX work part is available."
    if display_part is None or not same_nx_object(work_part, display_part):
        return "FAILED_CONTEXT", "Make the target assembly both Work Part and Displayed Part."
    assembly = safe_property(work_part, "ComponentAssembly")
    if assembly is None or safe_property(assembly, "RootComponent") is None:
        return "FAILED_NOT_ASSEMBLY", "The active work/display part is not an assembly."
    return "", ""


def run(session, run_datetime=None, mode=None):
    selected_mode = clean(mode).upper() if mode is not None else configured_mode()
    if selected_mode not in VALID_MODES:
        raise RuntimeError(
            "Invalid J27 mode {0!r}; expected DRY_RUN or APPLY.".format(selected_mode)
        )
    now = run_datetime or datetime.datetime.now().astimezone()
    parts = safe_property(session, "Parts")
    work_part = safe_property(parts, "Work")
    identity = part_identity(work_part)
    report = base_report(selected_mode, now, identity, empty_access(work_part, session))
    folder, stem = unique_run_folder(identity, now)
    report["artifacts"]["folder"] = folder
    preflight_artifact_folder(folder)

    failure_status, failure_message = context_failure(session)
    if failure_status:
        set_verdict(report, failure_status, failure_message)
        csv_path, json_path = write_outputs(report, folder, stem)
        return csv_path, json_path, report

    work_part, component_assembly, root = assembly_context(session)
    try:
        before = capture_snapshot(work_part, root)
    except Exception as error:
        report["errors"].append("Assembly scan failed: " + error_text(error))
        set_verdict(
            report, "FAILED_SCAN",
            "J27 could not establish a complete fail-closed direct-component baseline.",
        )
        csv_path, json_path = write_outputs(report, folder, stem)
        return csv_path, json_path, report

    report["before"] = public_snapshot(before)
    report["counts"] = snapshot_counts(before)
    counts = report["counts"]
    mark = None
    if counts["direct_component_count"] == 0:
        report["after"] = report["before"]
        set_verdict(
            report, "NO_COMPONENT_OCCURRENCES",
            "The active assembly contains no direct component occurrences.",
        )
    elif counts["move_candidate_count"] == 0:
        report["after"] = report["before"]
        set_verdict(
            report, "ALREADY_COMPLIANT",
            "Every direct component occurrence is already on specified layer 1.",
        )
    elif selected_mode == "DRY_RUN":
        report["after"] = report["before"]
        set_verdict(
            report, "DRY_RUN_READY",
            "DRY_RUN found {0} direct component occurrence(s) to normalize; NX was not modified.".format(
                counts["move_candidate_count"]
            ),
        )
    else:
        access = inspect_write_access(session, work_part)
        report["access"] = access
        if not access["allowed"]:
            report["after"] = report["before"]
            set_verdict(
                report, "BLOCKED_WRITE_ACCESS",
                access["message"] or "J27 could not prove assembly write access.",
            )
        else:
            mark = perform_apply(session, work_part, root, before, report)

    report["component_results"] = build_component_results(report)
    csv_path, json_path = write_with_apply_rollback(
        report, folder, stem, session, work_part, root, before, mark
    )
    return csv_path, json_path, report


def main():
    session = NXOpen.Session.GetSession()
    mode = configured_mode()
    log_line(session, "=" * 72)
    log_line(session, "J27 NORMALIZE DIRECT ASSEMBLY COMPONENTS TO LAYER 1")
    log_line(session, "Build: " + BUILD)
    log_line(session, "Mode: " + mode)
    log_line(
        session,
        "Scope: direct component occurrences only; all occurrence states included; no recursion or forced loading.",
    )
    log_line(
        session,
        "J27 never saves, checks out, checks in, changes prototype bodies, or exports STEP/JT.",
    )
    log_line(session, "=" * 72)
    try:
        csv_path, json_path, report = run(session, mode=mode)
        log_line(session, "Verdict: " + report["verdict"]["status"])
        log_line(session, report["verdict"]["message"])
        counts = report.get("counts", {})
        if counts:
            log_line(
                session,
                "Components: direct={0}; to change={1}; already layer 1={2}; suppressed={3}; blanked={4}; non-geometric={5}; prototype unavailable={6}".format(
                    counts["direct_component_count"], counts["move_candidate_count"],
                    counts["already_on_layer_1_count"], counts["suppressed_count"],
                    counts["blanked_count"], counts["non_geometric_count"],
                    counts["prototype_unavailable_count"],
                ),
            )
        if csv_path:
            log_line(session, "CSV: " + csv_path)
        if json_path:
            log_line(session, "JSON: " + json_path)
        if report["verdict"]["status"] == "APPLIED_VERIFIED":
            log_line(session, "The parent assembly is modified but UNSAVED. Inspect and save manually if correct.")
            log_line(session, "Undo once to revert: " + UNDO_MARK_NAME)
        elif mode == "DRY_RUN" and report["verdict"]["status"] == "DRY_RUN_READY":
            log_line(
                session,
                "Review both artifacts, check out the parent assembly if applicable, then set USER_MODE = \"APPLY\" and rerun.",
            )
    except Exception as error:
        log_line(session, "J27 FAILED: " + error_text(error))
        log_line(session, traceback.format_exc())
        raise


def get_unload_option(dummy):
    return NXOpen.Session.LibraryUnloadOption.Immediately


if __name__ == "__main__":
    main()
