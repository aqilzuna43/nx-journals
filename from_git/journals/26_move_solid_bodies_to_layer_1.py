"""
Journal 26 - Move Active-Part Solid Bodies to Layer 1

Scans every direct body owned by the active NX work part and moves only
traditional solid bodies to layer 1.  Blanked bodies and bodies on hidden
layers are included.  Sheet, convergent, and other body types are reported
but never moved.

DRY_RUN is the default.  APPLY requires a writable native part or a
Teamcenter part that is checked out by the current user.  J26 never performs
checkout, check-in, save, assembly traversal, STEP export, or JT generation.
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


# User setting.  Change to "APPLY" only after reviewing a DRY_RUN report.
USER_MODE = "DRY_RUN"

BUILD = "J26-NX2506-ACTIVE-PART-SOLIDS-TO-LAYER-1-V1"
SCHEMA_VERSION = 1
TARGET_LAYER = 1
FIRST_LAYER = 1
LAST_LAYER = 256
OUTPUT_FOLDER = "NX_LAYER_1_MIGRATION"
UNDO_MARK_NAME = "J26 Move solid bodies to layer 1"
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
    "PART_NAME",
    "MANAGED_MODE",
    "CHECKOUT_STATE",
    "CHECKOUT_OWNER",
    "CURRENT_USER",
    "READ_ONLY",
    "BODY_INDEX",
    "BODY_NAME",
    "BODY_TAG",
    "BODY_TYPE",
    "BLANKED",
    "ORIGINAL_LAYER",
    "FINAL_LAYER",
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
    mode = clean(os.environ.get("NX_J26_MODE") or USER_MODE).upper()
    if mode not in VALID_MODES:
        raise RuntimeError(
            "Invalid J26 mode {0!r}; expected DRY_RUN or APPLY.".format(mode)
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


def get_string_attribute(nx_object, name):
    try:
        return clean(nx_object.GetStringAttribute(name))
    except Exception:
        pass
    try:
        info = nx_object.GetUserAttribute(
            name,
            NXOpen.NXObject.AttributeType.String,
            -1,
        )
        return clean(info.StringValue)
    except Exception:
        return ""


def safe_part_name(part):
    for property_name in ("Name", "Leaf", "FullPath", "JournalIdentifier"):
        value = clean(safe_property(part, property_name))
        if value:
            return value
    return "UNKNOWN"


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
        "name": safe_part_name(part),
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
    if name:
        return name
    return clean(value)


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


def inspect_body(body, index, work_part):
    label = "body {0}".format(index)
    tag = object_tag(body, required=True, label=label)
    owner = required_property(body, "OwningPart", label)
    if owner is None or not same_nx_object(owner, work_part):
        raise RuntimeError(
            "{0} (tag {1}) is not owned by the active work part.".format(
                label, tag
            )
        )

    solid = bool(required_property(body, "IsSolidBody", label))
    sheet = bool(required_property(body, "IsSheetBody", label))
    convergent = bool(required_property(body, "IsConvergentBody", label))
    blanked = bool(required_property(body, "IsBlanked", label))
    layer = int(required_property(body, "Layer", label))
    if not FIRST_LAYER <= layer <= LAST_LAYER:
        raise RuntimeError(
            "{0} (tag {1}) has invalid layer {2}.".format(label, tag, layer)
        )

    if convergent:
        body_type = "CONVERGENT"
    elif solid:
        body_type = "TRADITIONAL_SOLID"
    elif sheet:
        body_type = "SHEET"
    else:
        body_type = "OTHER"

    name = clean(safe_property(body, "Name")) or "BODY_{0}".format(index)
    return {
        "index": index,
        "name": name,
        "tag": tag,
        "body_type": body_type,
        "is_solid": solid,
        "is_sheet": sheet,
        "is_convergent": convergent,
        "blanked": blanked,
        "layer": layer,
        "owner_tag": object_tag(owner),
        "_tag_key": tag,
        "_object": body,
    }


def capture_snapshot(work_part):
    try:
        bodies = list(work_part.Bodies)
    except Exception as error:
        raise RuntimeError(
            "Could not enumerate active work-part bodies: {0}".format(
                error_text(error)
            )
        )

    records = []
    tags = set()
    for index, body in enumerate(bodies, 1):
        record = inspect_body(body, index, work_part)
        if record["_tag_key"] in tags:
            raise RuntimeError(
                "Duplicate body tag {0} was returned by work_part.Bodies.".format(
                    record["tag"]
                )
            )
        tags.add(record["_tag_key"])
        records.append(record)

    layer_snapshot = read_layer_snapshot(work_part)
    return {
        "bodies": records,
        "work_layer": layer_snapshot["work_layer"],
        "layer_states": layer_snapshot["states"],
    }


def public_body_record(record):
    return {
        key: value
        for key, value in record.items()
        if key not in _INTERNAL_RECORD_KEYS
    }


def public_snapshot(snapshot):
    if snapshot is None:
        return None
    return {
        "bodies": [public_body_record(item) for item in snapshot["bodies"]],
        "work_layer": snapshot["work_layer"],
        "layer_states": dict(snapshot["layer_states"]),
    }


def records_by_tag(snapshot):
    return {record["_tag_key"]: record for record in snapshot["bodies"]}


def exact_snapshot_errors(expected, actual):
    errors = []
    expected_by_tag = records_by_tag(expected)
    actual_by_tag = records_by_tag(actual)
    if set(expected_by_tag) != set(actual_by_tag):
        errors.append("Direct body membership changed.")
    for tag in sorted(set(expected_by_tag) & set(actual_by_tag)):
        before = expected_by_tag[tag]
        after = actual_by_tag[tag]
        for field in (
            "body_type",
            "is_solid",
            "is_sheet",
            "is_convergent",
            "blanked",
            "layer",
            "owner_tag",
        ):
            if before[field] != after[field]:
                errors.append(
                    "Body {0} field {1} changed from {2!r} to {3!r}.".format(
                        tag, field, before[field], after[field]
                    )
                )
    if expected["work_layer"] != actual["work_layer"]:
        errors.append(
            "Work layer changed from {0} to {1}.".format(
                expected["work_layer"], actual["work_layer"]
            )
        )
    if expected["layer_states"] != actual["layer_states"]:
        changed = [
            layer for layer in expected["layer_states"]
            if expected["layer_states"].get(layer)
            != actual["layer_states"].get(layer)
        ]
        errors.append(
            "Layer states changed for: {0}.".format(
                ", ".join(changed) if changed else "unknown layers"
            )
        )
    return errors


def apply_verification_errors(before, after):
    errors = []
    before_by_tag = records_by_tag(before)
    after_by_tag = records_by_tag(after)
    if set(before_by_tag) != set(after_by_tag):
        errors.append("Direct body membership changed during the layer move.")
    for tag in sorted(set(before_by_tag) & set(after_by_tag)):
        original = before_by_tag[tag]
        final = after_by_tag[tag]
        for field in (
            "body_type",
            "is_solid",
            "is_sheet",
            "is_convergent",
            "blanked",
            "owner_tag",
        ):
            if original[field] != final[field]:
                errors.append(
                    "Body {0} field {1} changed unexpectedly.".format(tag, field)
                )
        expected_layer = (
            TARGET_LAYER
            if original["body_type"] == "TRADITIONAL_SOLID"
            else original["layer"]
        )
        if final["layer"] != expected_layer:
            errors.append(
                "Body {0} expected layer {1}, observed {2}.".format(
                    tag, expected_layer, final["layer"]
                )
            )
    if before["work_layer"] != after["work_layer"]:
        errors.append("The NX work layer changed during APPLY.")
    if before["layer_states"] != after["layer_states"]:
        errors.append("One or more NX layer visibility/selectability states changed.")
    return errors


def snapshot_counts(snapshot):
    records = snapshot["bodies"] if snapshot else []
    eligible = [
        item for item in records
        if item["body_type"] == "TRADITIONAL_SOLID"
    ]
    return {
        "direct_body_count": len(records),
        "eligible_solid_count": len(eligible),
        "move_candidate_count": sum(
            1 for item in eligible if item["layer"] != TARGET_LAYER
        ),
        "already_on_layer_1_count": sum(
            1 for item in eligible if item["layer"] == TARGET_LAYER
        ),
        "skipped_sheet_count": sum(
            1 for item in records if item["body_type"] == "SHEET"
        ),
        "skipped_convergent_count": sum(
            1 for item in records if item["body_type"] == "CONVERGENT"
        ),
        "skipped_other_count": sum(
            1 for item in records if item["body_type"] == "OTHER"
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
    return full_path.upper().startswith("@DB/") or identifier.upper().startswith(
        "@DB/"
    )


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
                normalized = clean(getattr(value, "name", value)).upper().replace(
                    "_", ""
                )
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
    state = (
        "CHECKED_OUT" if checked is True
        else "CHECKED_IN" if checked is False
        else "UNKNOWN"
    )
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
        allowed = read_only is not True
        return {
            "allowed": allowed,
            "managed": False,
            "checkout_state": "NATIVE",
            "checkout_owner": "",
            "current_user": "",
            "read_only": read_only,
            "raw_checkout": "",
            "message": "" if allowed else "Native part is read-only.",
        }

    state, owner, raw = checkout_status(part)
    current_user = current_teamcenter_user(session)
    owner_matches = bool(
        state == "CHECKED_OUT"
        and owner
        and current_user
        and owner.casefold() == current_user.casefold()
    )
    allowed = bool(owner_matches and read_only is False)
    messages = []
    if state == "CHECKED_IN":
        messages.append("Part is checked in; J26 never performs checkout.")
    elif state == "UNKNOWN":
        messages.append("Checkout state is unknown: {0}.".format(raw or "<none>"))
    elif not owner:
        messages.append("Checkout owner is unavailable.")
    elif not current_user:
        messages.append("Current Teamcenter user is unavailable.")
    elif owner.casefold() != current_user.casefold():
        messages.append("Part is checked out by another user: {0}.".format(owner))
    if read_only is True:
        messages.append("Part is read-only in this NX session.")
    elif read_only is None:
        messages.append("Part read-only state is unavailable.")
    return {
        "allowed": allowed,
        "managed": True,
        "checkout_state": state,
        "checkout_owner": owner,
        "current_user": current_user,
        "read_only": read_only,
        "raw_checkout": raw,
        "message": " ".join(messages),
    }


def empty_access(part=None, session=None):
    managed = bool(part is not None and managed_mode(session, part))
    return {
        "allowed": None,
        "managed": managed,
        "checkout_state": "NOT_INSPECTED",
        "checkout_owner": "",
        "current_user": "",
        "read_only": read_only_value(part) if part is not None else None,
        "raw_checkout": "",
        "message": "Write access is inspected only when APPLY has bodies to move.",
    }


def read_only_text(value):
    return "UNKNOWN" if value is None else "YES" if value else "NO"


def unique_run_folder(identity, now):
    root = os.path.join(io_root(), OUTPUT_FOLDER)
    os.makedirs(root, exist_ok=True)
    identity_token = filename_token(
        identity.get("number") or identity.get("name") or "UNKNOWN"
    )
    base_stem = "J26_LAYER_1_{0}_{1}".format(
        identity_token, now.strftime("%Y%m%d_%H%M%S")
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
        os.path.join(folder, ".j26_csv_probe.tmp"),
        os.path.join(folder, ".j26_json_probe.tmp"),
    ]
    try:
        for path in paths:
            with open(path, "w", encoding="utf-8") as handle:
                handle.write("J26 evidence preflight\n")
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
            "scope": "ACTIVE_WORK_PART_DIRECT_BODIES",
            "target_layer": TARGET_LAYER,
            "include_blanked_and_hidden": True,
            "traditional_solids_only": True,
            "automatic_checkout": False,
            "automatic_save": False,
            "automatic_checkin": False,
        },
        "part": identity,
        "access": access,
        "counts": {},
        "before": None,
        "after_attempt": None,
        "after": None,
        "action": {
            "api": "work_part.Layers.MoveDisplayableObjects",
            "attempted": False,
            "candidate_tags": [],
            "undo_mark_name": UNDO_MARK_NAME,
            "undo_mark_created": False,
            "successful_change_left_undoable": False,
            "verification_errors": [],
            "error": "",
        },
        "rollback": {
            "attempted": False,
            "status": "NOT_REQUIRED",
            "verification_errors": [],
            "error": "",
        },
        "body_results": [],
        "errors": [],
        "artifacts": {"folder": "", "csv": "", "json": ""},
        "verdict": {
            "status": "INITIALIZING",
            "message": "J26 has not completed.",
        },
    }


def set_verdict(report, status, message):
    report["verdict"] = {"status": status, "message": message}


def final_snapshot_for_results(report):
    return report.get("after") or report.get("before")


def build_body_results(report):
    before = report.get("before")
    if not before:
        return []
    before_records = {item["tag"]: item for item in before["bodies"]}
    final = final_snapshot_for_results(report)
    final_records = {
        item["tag"]: item for item in (final or {}).get("bodies", [])
    }
    verdict = report["verdict"]["status"]
    mode = report["mode"]
    results = []
    for tag, original in before_records.items():
        final_record = final_records.get(tag)
        final_layer = final_record["layer"] if final_record else None
        body_type = original["body_type"]
        if body_type != "TRADITIONAL_SOLID":
            action = "SKIPPED"
            status = "SKIPPED_BODY_TYPE"
            message = "Only traditional solid bodies are eligible."
        elif original["layer"] == TARGET_LAYER:
            action = "UNCHANGED"
            status = "ALREADY_ON_LAYER_1"
            message = "Body was already on layer 1."
        elif mode == "DRY_RUN":
            action = "WOULD_MOVE"
            status = "DRY_RUN"
            message = "APPLY would move this body to layer 1."
        elif verdict == "APPLIED_VERIFIED":
            action = "MOVED"
            status = "APPLIED_VERIFIED"
            message = "Body moved to layer 1 and passed post-verification."
        elif verdict in ("ROLLED_BACK", "ROLLBACK_FAILED"):
            action = "ROLLBACK_ATTEMPTED"
            status = verdict
            message = report["verdict"]["message"]
        else:
            action = "NOT_MOVED"
            status = verdict
            message = report["verdict"]["message"]
        results.append({
            "index": original["index"],
            "name": original["name"],
            "tag": tag,
            "body_type": body_type,
            "blanked": original["blanked"],
            "original_layer": original["layer"],
            "final_layer": final_layer,
            "action": action,
            "status": status,
            "message": message,
        })
    return results


def csv_rows(report):
    identity = report["part"]
    access = report["access"]
    common = {
        "RUN_TIMESTAMP": report["run_timestamp"],
        "JOURNAL_BUILD": report["journal_build"],
        "SCHEMA_VERSION": report["schema_version"],
        "MODE": report["mode"],
        "VERDICT": report["verdict"]["status"],
        "DB_PART_NO": identity.get("number", ""),
        "DB_PART_REV": identity.get("revision", ""),
        "PART_NAME": identity.get("name", ""),
        "MANAGED_MODE": "YES" if access.get("managed") else "NO",
        "CHECKOUT_STATE": access.get("checkout_state", ""),
        "CHECKOUT_OWNER": access.get("checkout_owner", ""),
        "CURRENT_USER": access.get("current_user", ""),
        "READ_ONLY": read_only_text(access.get("read_only")),
    }
    summary = dict(common)
    summary.update({
        "ROW_TYPE": "SUMMARY",
        "ACTION": "RUN_SUMMARY",
        "STATUS": report["verdict"]["status"],
        "MESSAGE": report["verdict"]["message"],
    })
    rows = [summary]
    for body in report["body_results"]:
        row = dict(common)
        row.update({
            "ROW_TYPE": "BODY",
            "BODY_INDEX": body["index"],
            "BODY_NAME": body["name"],
            "BODY_TAG": body["tag"],
            "BODY_TYPE": body["body_type"],
            "BLANKED": "YES" if body["blanked"] else "NO",
            "ORIGINAL_LAYER": body["original_layer"],
            "FINAL_LAYER": (
                "" if body["final_layer"] is None else body["final_layer"]
            ),
            "ACTION": body["action"],
            "STATUS": body["status"],
            "MESSAGE": body["message"],
        })
        rows.append(row)
    return rows


def write_outputs(report, folder, stem):
    csv_path = os.path.join(folder, stem + ".csv")
    json_path = os.path.join(folder, stem + ".json")
    csv_temp = csv_path + ".tmp"
    json_temp = json_path + ".tmp"
    report["artifacts"] = {
        "folder": folder,
        "csv": csv_path,
        "json": json_path,
    }
    try:
        with open(csv_temp, "w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.DictWriter(
                handle, fieldnames=CSV_COLUMNS, extrasaction="ignore"
            )
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


def rollback_to_before(session, mark, work_part, before):
    result = {
        "attempted": True,
        "status": "ROLLBACK_FAILED",
        "verification_errors": [],
        "error": "",
    }
    try:
        session.UndoToMark(mark, UNDO_MARK_NAME)
    except Exception as error:
        result["error"] = "UndoToMark failed: " + error_text(error)
        try:
            after = capture_snapshot(work_part)
        except Exception as capture_error:
            after = None
            result["error"] += " | Rollback snapshot failed: " + error_text(
                capture_error
            )
        return result, after

    try:
        after = capture_snapshot(work_part)
        errors = exact_snapshot_errors(before, after)
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
                result["error"],
                "DeleteUndoMark failed: " + delete_error,
            ) if item
        )
    return result, after


def perform_apply(session, work_part, before, report):
    candidates = [
        record for record in before["bodies"]
        if record["body_type"] == "TRADITIONAL_SOLID"
        and record["layer"] != TARGET_LAYER
    ]
    report["action"]["candidate_tags"] = [item["tag"] for item in candidates]
    try:
        mark = session.SetUndoMark(
            NXOpen.Session.MarkVisibility.Visible,
            UNDO_MARK_NAME,
        )
        report["action"]["undo_mark_created"] = True
    except Exception as error:
        report["action"]["error"] = error_text(error)
        set_verdict(
            report,
            "ROLLED_BACK",
            "J26 refused to move bodies because it could not create an NX undo mark.",
        )
        return None

    report["action"]["attempted"] = True
    try:
        work_part.Layers.MoveDisplayableObjects(
            TARGET_LAYER,
            [item["_object"] for item in candidates],
        )
        after = capture_snapshot(work_part)
        report["after_attempt"] = public_snapshot(after)
        verification_errors = apply_verification_errors(before, after)
        report["action"]["verification_errors"] = verification_errors
        if verification_errors:
            raise VerificationError(verification_errors, snapshot=after)
        report["after"] = public_snapshot(after)
        report["action"]["successful_change_left_undoable"] = True
        set_verdict(
            report,
            "APPLIED_VERIFIED",
            "Every eligible traditional solid body is on layer 1; the part remains unsaved under one visible NX undo mark.",
        )
        return mark
    except Exception as error:
        report["action"]["error"] = error_text(error)
        if isinstance(error, VerificationError) and error.snapshot is not None:
            report["after_attempt"] = public_snapshot(error.snapshot)
        rollback, final_snapshot = rollback_to_before(
            session, mark, work_part, before
        )
        report["rollback"] = rollback
        report["after"] = public_snapshot(final_snapshot)
        if rollback["status"] == "ROLLED_BACK":
            set_verdict(
                report,
                "ROLLED_BACK",
                "The layer move failed verification or raised an NX error; the original state was restored.",
            )
        else:
            set_verdict(
                report,
                "ROLLBACK_FAILED",
                "The layer move failed and J26 could not prove restoration of the original state. Use NX Undo and inspect the part immediately.",
            )
        return None


def write_with_apply_rollback(
    report, folder, stem, session, work_part, before, success_mark
):
    report["body_results"] = build_body_results(report)
    try:
        return write_outputs(report, folder, stem)
    except Exception as error:
        report["errors"].append("Evidence write failed: " + error_text(error))
        if report["verdict"]["status"] != "APPLIED_VERIFIED" or success_mark is None:
            raise

    report["action"]["successful_change_left_undoable"] = False
    rollback, final_snapshot = rollback_to_before(
        session, success_mark, work_part, before
    )
    report["rollback"] = rollback
    report["after"] = public_snapshot(final_snapshot)
    if rollback["status"] == "ROLLED_BACK":
        set_verdict(
            report,
            "ROLLED_BACK",
            "The verified layer move was rolled back because paired CSV/JSON evidence could not be completed.",
        )
    else:
        set_verdict(
            report,
            "ROLLBACK_FAILED",
            "Evidence writing failed and J26 could not prove restoration of the original state. Use NX Undo and inspect the part immediately.",
        )
    report["body_results"] = build_body_results(report)
    try:
        return write_outputs(report, folder, stem)
    except Exception as second_error:
        report["errors"].append(
            "Rollback evidence write also failed: " + error_text(second_error)
        )
        report["artifacts"] = {"folder": folder, "csv": "", "json": ""}
        return "", ""


def run(session, run_datetime=None, mode=None):
    selected_mode = clean(mode).upper() if mode is not None else configured_mode()
    if selected_mode not in VALID_MODES:
        raise RuntimeError(
            "Invalid J26 mode {0!r}; expected DRY_RUN or APPLY.".format(
                selected_mode
            )
        )
    now = run_datetime or datetime.datetime.now().astimezone()
    work_part = safe_property(safe_property(session, "Parts"), "Work")
    identity = part_identity(work_part)
    access = empty_access(work_part, session)
    report = base_report(selected_mode, now, identity, access)

    folder, stem = unique_run_folder(identity, now)
    report["artifacts"]["folder"] = folder
    preflight_artifact_folder(folder)

    if work_part is None:
        set_verdict(
            report,
            "FAILED_NO_WORK_PART",
            "No active NX work part is available.",
        )
        report["body_results"] = []
        csv_path, json_path = write_outputs(report, folder, stem)
        return csv_path, json_path, report

    try:
        before = capture_snapshot(work_part)
    except Exception as error:
        report["errors"].append("Body/layer scan failed: " + error_text(error))
        set_verdict(
            report,
            "FAILED_SCAN",
            "J26 could not establish a complete fail-closed body and layer baseline.",
        )
        report["body_results"] = []
        csv_path, json_path = write_outputs(report, folder, stem)
        return csv_path, json_path, report

    report["before"] = public_snapshot(before)
    report["counts"] = snapshot_counts(before)
    counts = report["counts"]
    success_mark = None

    if counts["eligible_solid_count"] == 0:
        report["after"] = report["before"]
        set_verdict(
            report,
            "NO_ELIGIBLE_SOLIDS",
            "The work part contains no direct traditional solid bodies.",
        )
    elif counts["move_candidate_count"] == 0:
        report["after"] = report["before"]
        set_verdict(
            report,
            "ALREADY_COMPLIANT",
            "Every eligible traditional solid body is already on layer 1.",
        )
    elif selected_mode == "DRY_RUN":
        report["after"] = report["before"]
        set_verdict(
            report,
            "DRY_RUN_READY",
            "DRY_RUN found {0} traditional solid body/bodies to move to layer 1; NX was not modified.".format(
                counts["move_candidate_count"]
            ),
        )
    else:
        access = inspect_write_access(session, work_part)
        report["access"] = access
        if not access["allowed"]:
            report["after"] = report["before"]
            set_verdict(
                report,
                "BLOCKED_WRITE_ACCESS",
                access["message"] or "J26 could not prove write access.",
            )
        else:
            success_mark = perform_apply(session, work_part, before, report)

    report["body_results"] = build_body_results(report)
    csv_path, json_path = write_with_apply_rollback(
        report,
        folder,
        stem,
        session,
        work_part,
        before,
        success_mark,
    )
    return csv_path, json_path, report


def main():
    session = NXOpen.Session.GetSession()
    mode = configured_mode()
    log_line(session, "=" * 72)
    log_line(session, "J26 MOVE ACTIVE-PART SOLID BODIES TO LAYER 1")
    log_line(session, "Build: " + BUILD)
    log_line(session, "Mode: " + mode)
    log_line(
        session,
        "Scope: direct traditional solids in the active work part; blanked/hidden included.",
    )
    log_line(
        session,
        "J26 never saves, checks out, checks in, traverses assemblies, or exports STEP/JT.",
    )
    log_line(session, "=" * 72)
    try:
        csv_path, json_path, report = run(session, mode=mode)
        counts = report.get("counts", {})
        log_line(
            session,
            "Verdict: {0}".format(report["verdict"]["status"]),
        )
        log_line(session, report["verdict"]["message"])
        if counts:
            log_line(
                session,
                "Bodies: direct={0}; eligible solids={1}; to move={2}; already layer 1={3}; sheets={4}; convergent={5}; other={6}".format(
                    counts["direct_body_count"],
                    counts["eligible_solid_count"],
                    counts["move_candidate_count"],
                    counts["already_on_layer_1_count"],
                    counts["skipped_sheet_count"],
                    counts["skipped_convergent_count"],
                    counts["skipped_other_count"],
                ),
            )
        if csv_path:
            log_line(session, "CSV: " + csv_path)
        if json_path:
            log_line(session, "JSON: " + json_path)
        if not csv_path or not json_path:
            log_line(
                session,
                "WARNING: paired evidence could not be written; inspect rollback status in the messages above.",
            )
        if report["verdict"]["status"] == "APPLIED_VERIFIED":
            log_line(
                session,
                "The part is modified but UNSAVED. Inspect it, then save manually if correct.",
            )
            log_line(session, "Undo once to revert: " + UNDO_MARK_NAME)
        elif mode == "DRY_RUN" and report["verdict"]["status"] == "DRY_RUN_READY":
            log_line(
                session,
                "Review both artifacts, check out the Teamcenter part if applicable, then set USER_MODE = \"APPLY\" and rerun.",
            )
    except Exception as error:
        log_line(session, "J26 FAILED: " + error_text(error))
        log_line(session, traceback.format_exc())
        raise


def get_unload_option(dummy):
    return NXOpen.Session.LibraryUnloadOption.Immediately


if __name__ == "__main__":
    main()
