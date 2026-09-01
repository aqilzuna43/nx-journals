"""
Journal 29 - Set Selected Components to Reference-Only

Adds the exact occurrence attribute proven by J28 V2 to one or more
preselected direct components of the active assembly:

    REFERENCE_COMPONENT = ""  (string, index -1, component-instance scope)

APPLY is the default. It requires the active assembly to be both Work and
Display part and already writable. J29 never scans the assembly tree, loads
components, checks out, checks in, or saves. The complete selection is
preflighted before one atomic batch write. A verified APPLY remains unsaved
under one visible NX undo mark. Any write/verification/evidence failure
triggers UndoToMark and a baseline reread for every selected occurrence.

Target: NX 2312 and NX X 2506 embedded Python
Run via: NX > Tools > Journal > Play
"""

import csv
import datetime
import json
import os
import re
import traceback

import NXOpen


# User setting. APPLY is intentionally the operator-selected default; all
# targets still pass a fail-closed batch preflight before any write occurs.
USER_MODE = "APPLY"

BUILD = "J29-NX2506-BATCH-REFERENCE-ONLY-V2"
SCHEMA_VERSION = 2
OUTPUT_FOLDER = "NX_REFERENCE_ONLY"
UNDO_MARK_NAME = "J29 Set selected components Reference-Only"
VALID_MODES = ("DRY_RUN", "APPLY")
DEFAULT_MAX_SELECTION = 100
REFERENCE_ATTRIBUTE = "REFERENCE_COMPONENT"
CONFLICT_ATTRIBUTES = (
    "PLIST_IGNORE_MEMBER",
    "PLIST_IGNORE_SUBASSEMBLY",
)
CONTROL_ATTRIBUTES = (REFERENCE_ATTRIBUTE,) + CONFLICT_ATTRIBUTES

CSV_COLUMNS = (
    "RUN_TIMESTAMP", "JOURNAL_BUILD", "SCHEMA_VERSION", "MODE", "VERDICT",
    "ROW_TYPE", "DB_PART_NO", "DB_PART_REV", "ASSEMBLY_NAME",
    "MANAGED_MODE", "CHECKOUT_STATE", "CHECKOUT_OWNER", "CURRENT_USER",
    "READ_ONLY", "SELECTION_COUNT", "SELECTION_INDEX", "TARGET_STATUS",
    "TARGET_NAME", "TARGET_DISPLAY_NAME",
    "TARGET_TAG", "PARENT_TAG", "PROTOTYPE_NAME", "PROTOTYPE_TAG",
    "SUPPRESSED", "REFERENCE_COMPONENT_BEFORE",
    "REFERENCE_COMPONENT_AFTER", "PLIST_IGNORE_MEMBER_PRESENT",
    "PLIST_IGNORE_SUBASSEMBLY_PRESENT", "ACTION", "STATUS", "MESSAGE",
)

_INVALID_FILENAME_CHARS = '<>:"/\\|?*'


class VerificationError(RuntimeError):
    def __init__(self, messages, snapshot=None):
        RuntimeError.__init__(self, " | ".join(messages))
        self.messages = list(messages)
        self.snapshot = snapshot


class SelectionLimitError(RuntimeError):
    pass


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
            "Could not inspect {0}.{1}: {2}".format(label, name, error_text(error))
        )


def safe_name(value, fallback="UNKNOWN"):
    for name in ("Name", "DisplayName", "Leaf", "FullPath", "JournalIdentifier"):
        result = clean(safe_property(value, name))
        if result:
            return result
    return fallback


def object_tag(value, required=False, label="object"):
    try:
        tag = getattr(value, "Tag")
        tag = tag() if callable(tag) else tag
        result = clean(tag)
    except Exception as error:
        if required:
            raise RuntimeError(
                "Could not inspect {0}.Tag: {1}".format(label, error_text(error))
            )
        return ""
    if required and (not result or result == "0"):
        raise RuntimeError("{0} has no usable NX tag.".format(label))
    return result


def same_nx_object(left, right):
    if left is right:
        return True
    left_tag, right_tag = object_tag(left), object_tag(right)
    return bool(left_tag and right_tag and left_tag == right_tag)


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
    mode = clean(os.environ.get("NX_J29_MODE") or USER_MODE).upper()
    if mode not in VALID_MODES:
        raise RuntimeError(
            "Invalid J29 mode {0!r}; expected DRY_RUN or APPLY.".format(mode)
        )
    return mode


def configured_max_selection():
    raw = clean(os.environ.get("NX_J29_MAX_SELECTION"))
    if not raw:
        return DEFAULT_MAX_SELECTION
    try:
        value = int(raw)
    except (TypeError, ValueError):
        raise RuntimeError(
            "Invalid NX_J29_MAX_SELECTION {0!r}; expected a positive integer.".format(raw)
        )
    if value <= 0:
        raise RuntimeError("NX_J29_MAX_SELECTION must be greater than zero.")
    return value


def get_string_attribute(nx_object, title):
    try:
        return clean(nx_object.GetStringAttribute(title))
    except Exception:
        pass
    try:
        info = nx_object.GetUserAttribute(
            title, NXOpen.NXObject.AttributeType.String, -1
        )
        return clean(info.StringValue)
    except Exception:
        return ""


def part_identity(part):
    if part is None:
        return {"number": "", "revision": "", "name": "UNKNOWN", "tag": ""}
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
        "tag": object_tag(part),
    }


def assembly_context(session):
    parts = safe_property(session, "Parts")
    work_part = safe_property(parts, "Work")
    display_part = safe_property(parts, "Display")
    if work_part is None:
        raise RuntimeError("No active NX work part is available.")
    if display_part is None or not same_nx_object(work_part, display_part):
        raise RuntimeError("Make the parent assembly both Work Part and Displayed Part.")
    assembly = required_property(work_part, "ComponentAssembly", "work part")
    root = required_property(assembly, "RootComponent", "work_part.ComponentAssembly")
    if root is None:
        raise RuntimeError("The active work/display part is not an assembly.")
    return work_part, root


def selected_components(selection_manager, max_selection):
    try:
        count = int(selection_manager.GetNumSelectedObjects())
    except Exception as error:
        raise RuntimeError("Could not inspect NX preselection: " + error_text(error))
    if count < 1:
        raise RuntimeError(
            "Select at least one component row in Assembly Navigator."
        )
    if count > max_selection:
        raise SelectionLimitError(
            "Selection count {0} exceeds the configured J29 limit of {1}.".format(
                count, max_selection
            )
        )
    selected = []
    for index in range(count):
        try:
            component = selection_manager.GetSelectedTaggedObject(index)
        except Exception as error:
            raise RuntimeError(
                "Could not read selected NX object {0}: {1}".format(
                    index + 1, error_text(error)
                )
            )
        selected.append(component)
    return selected, count


def attribute_type_name(info):
    raw = clean(safe_property(info, "Type"))
    normalized = raw.split(".")[-1].upper()
    names = {
        "0": "INVALID", "1": "NULL", "2": "BOOLEAN", "3": "INTEGER",
        "4": "REAL", "5": "STRING", "6": "TIME", "7": "REFERENCE",
        "100": "ANY",
    }
    return names.get(normalized, normalized or "UNKNOWN")


def read_control(component, title):
    attribute_type = NXOpen.NXObject.AttributeType.String
    try:
        present = bool(component.HasInstanceUserAttribute(title, attribute_type, -1))
    except Exception as error:
        raise RuntimeError(
            "Could not test selected occurrence attribute {0}: {1}".format(
                title, error_text(error)
            )
        )
    if not present:
        return {
            "title": title, "present": False, "type": "", "raw_value": "",
            "value_state": "ABSENT", "unset": None, "inherited": None,
            "is_override": None, "owned_by_system": None, "pdm_based": None,
            "not_saved": None,
        }
    try:
        info = component.GetInstanceUserAttribute(title, attribute_type, -1)
    except Exception as error:
        raise RuntimeError(
            "Could not read selected occurrence attribute {0}: {1}".format(
                title, error_text(error)
            )
        )
    raw_value = safe_property(info, "StringValue", "")
    raw_value = "" if raw_value is None else str(raw_value)
    unset = bool(safe_property(info, "Unset", False))
    return {
        "title": clean(safe_property(info, "Title")) or title,
        "present": not unset,
        "type": attribute_type_name(info),
        "raw_value": raw_value,
        "value_state": "PRESENT_BLANK" if raw_value == "" else "PRESENT_VALUE",
        "unset": unset,
        "inherited": bool(safe_property(info, "Inherited", False)),
        "is_override": bool(safe_property(info, "IsOverride", False)),
        "owned_by_system": bool(safe_property(info, "OwnedBySystem", False)),
        "pdm_based": bool(safe_property(info, "PdmBased", False)),
        "not_saved": bool(safe_property(info, "NotSaved", False)),
    }


def component_snapshot(component, root, work_part):
    tag = object_tag(component, required=True, label="selected component")
    parent = required_property(component, "Parent", "selected component")
    if parent is None or not same_nx_object(parent, root):
        raise RuntimeError(
            "The selected component must be a direct child of the active assembly root."
        )
    suppressed = bool(required_property(component, "IsSuppressed", "selected component"))
    if suppressed:
        raise RuntimeError("The selected component is suppressed; J29 will not modify it.")
    prototype = safe_property(component, "Prototype")
    controls = {title: read_control(component, title) for title in CONTROL_ATTRIBUTES}
    return {
        "tag": tag,
        "name": safe_name(component),
        "display_name": clean(safe_property(component, "DisplayName")),
        "parent_tag": object_tag(parent, required=True, label="selected component parent"),
        "prototype": {
            "available": prototype is not None,
            "name": safe_name(prototype, "<unavailable>"),
            "tag": object_tag(prototype),
        },
        "suppressed": suppressed,
        "controls": controls,
        "work_part_modified": bool(safe_property(work_part, "IsModified", False)),
    }


def reference_contract_errors(snapshot, require_present):
    errors = []
    reference = snapshot["controls"][REFERENCE_ATTRIBUTE]
    if require_present:
        if not reference["present"]:
            errors.append("REFERENCE_COMPONENT is absent after the write.")
        if reference["type"] != "STRING":
            errors.append("REFERENCE_COMPONENT is not a string attribute.")
        if reference["raw_value"] != "":
            errors.append("REFERENCE_COMPONENT is not blank.")
        if reference["inherited"]:
            errors.append("REFERENCE_COMPONENT is inherited instead of direct.")
        if reference["owned_by_system"]:
            errors.append("REFERENCE_COMPONENT is unexpectedly system-owned.")
        if reference["pdm_based"]:
            errors.append("REFERENCE_COMPONENT is unexpectedly PDM-based.")
    else:
        if reference["present"]:
            errors.append("REFERENCE_COMPONENT was not restored to absent.")
    return errors


def stable_snapshot_errors(before, after, expect_reference):
    errors = []
    for field in ("tag", "name", "display_name", "parent_tag", "prototype", "suppressed"):
        if before[field] != after[field]:
            errors.append("Selected component field {0} changed unexpectedly.".format(field))
    for title in CONFLICT_ATTRIBUTES:
        if before["controls"][title] != after["controls"][title]:
            errors.append("{0} changed unexpectedly.".format(title))
    errors.extend(reference_contract_errors(after, expect_reference))
    return errors


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
    checked, owner = None, ""
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
    method = getattr(safe_property(session, "PdmSession"), "GetUserName", None)
    if not callable(method):
        return ""
    try:
        return clean(method())
    except Exception:
        return ""


def teamcenter_identity_tokens(value):
    """Return conservative identity tokens without fuzzy display-name matching."""
    text = clean(value).casefold()
    if not text:
        return set()
    tokens = {text}
    for match in re.finditer(r"\(([^()]+)\)|\[([^\[\]]+)\]", text):
        token = clean(match.group(1) or match.group(2)).casefold()
        if token:
            tokens.add(token)
    for match in re.finditer(
        r"(?<![0-9a-f])(?:[0-9a-f]{8}-?[0-9a-f]{4}-?[0-9a-f]{4}-?[0-9a-f]{4}-?[0-9a-f]{12})(?![0-9a-f])",
        text,
    ):
        tokens.add(match.group(0).replace("-", ""))
    for match in re.finditer(
        r"[a-z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-z0-9.-]+\.[a-z]{2,}",
        text,
    ):
        tokens.add(match.group(0))
    return tokens


def same_teamcenter_user(owner, current_user):
    owner_tokens = teamcenter_identity_tokens(owner)
    user_tokens = teamcenter_identity_tokens(current_user)
    return bool(owner_tokens and user_tokens and owner_tokens.intersection(user_tokens))


def inspect_write_access(session, part):
    read_only = read_only_value(part)
    managed = managed_mode(session, part)
    if not managed:
        allowed = read_only is False
        message = "" if allowed else (
            "Native parent assembly is read-only." if read_only is True
            else "Native parent assembly read-only state is unavailable."
        )
        return {
            "allowed": allowed, "managed": False, "checkout_state": "NATIVE",
            "checkout_owner": "", "current_user": "", "read_only": read_only,
            "owner_is_current_user": None,
            "raw_checkout": "", "message": message,
        }
    state, owner, raw = checkout_status(part)
    user = current_teamcenter_user(session)
    owner_is_current_user = same_teamcenter_user(owner, user)
    allowed = bool(
        state == "CHECKED_OUT" and owner and user
        and owner_is_current_user and read_only is False
    )
    messages = []
    if state == "CHECKED_IN":
        messages.append("Parent assembly is checked in; J29 never performs checkout.")
    elif state == "UNKNOWN":
        messages.append("Parent assembly checkout state is unknown: {0}.".format(raw or "<none>"))
    elif not owner:
        messages.append("Checkout owner is unavailable.")
    elif not user:
        messages.append("Current Teamcenter user is unavailable.")
    elif not owner_is_current_user:
        messages.append("Parent assembly is checked out by another user: {0}.".format(owner))
    if read_only is True:
        messages.append("Parent assembly is read-only in this NX session.")
    elif read_only is None:
        messages.append("Parent assembly read-only state is unavailable.")
    return {
        "allowed": allowed, "managed": True, "checkout_state": state,
        "checkout_owner": owner, "current_user": user, "read_only": read_only,
        "owner_is_current_user": owner_is_current_user,
        "raw_checkout": raw, "message": " ".join(messages),
    }


def empty_access(part=None, session=None):
    return {
        "allowed": None,
        "managed": bool(part is not None and managed_mode(session, part)),
        "checkout_state": "NOT_INSPECTED", "checkout_owner": "",
        "current_user": "", "read_only": read_only_value(part) if part is not None else None,
        "owner_is_current_user": None,
        "raw_checkout": "",
        "message": "Write access is inspected only for an APPLY candidate.",
    }


def unique_run_folder(identity, now):
    root = os.path.join(io_root(), OUTPUT_FOLDER)
    os.makedirs(root, exist_ok=True)
    token = filename_token(identity.get("number") or identity.get("name") or "UNKNOWN")
    base = "J29_REFERENCE_ONLY_{0}_{1}".format(token, now.strftime("%Y%m%d_%H%M%S"))
    folder = os.path.join(root, base)
    suffix = 1
    while os.path.exists(folder):
        folder = os.path.join(root, "{0}_{1}".format(base, suffix))
        suffix += 1
    os.makedirs(folder)
    return folder, os.path.basename(folder)


def preflight_artifact_folder(folder):
    paths = [
        os.path.join(folder, ".j29_csv_probe.tmp"),
        os.path.join(folder, ".j29_json_probe.tmp"),
    ]
    try:
        for path in paths:
            with open(path, "w", encoding="utf-8") as handle:
                handle.write("J29 evidence preflight\n")
    finally:
        for path in paths:
            try:
                if os.path.exists(path):
                    os.remove(path)
            except Exception:
                pass


def set_verdict(report, status, message):
    report["verdict"] = {"status": status, "message": message}


def base_report(mode, now, identity, access):
    return {
        "schema_version": SCHEMA_VERSION,
        "journal_build": BUILD,
        "run_timestamp": now.isoformat(timespec="seconds"),
        "mode": mode,
        "configuration": {
            "scope": "PRESELECTED_DIRECT_COMPONENT_OCCURRENCES_ATOMIC_BATCH",
            "attribute_title": REFERENCE_ATTRIBUTE,
            "attribute_type": "STRING", "attribute_index": -1,
            "attribute_value": "", "force_load": False,
            "max_selection": None,
            "automatic_checkout": False, "automatic_save": False,
            "automatic_checkin": False,
        },
        "assembly": identity, "selection": {"count": 0, "source": "NX_PRESELECTION"},
        "access": access, "targets": [],
        "action": {
            "api": "Component.SetInstanceUserAttribute(REFERENCE_COMPONENT, -1, blank, Update.Option.Now)",
            "attempted": False, "attempted_count": 0, "applied_count": 0,
            "undo_mark_name": UNDO_MARK_NAME,
            "undo_mark_created": False, "successful_change_left_undoable": False,
            "verification_errors": [], "error": "",
        },
        "rollback": {
            "attempted": False, "status": "NOT_REQUIRED",
            "verification_errors": [], "error": "",
        },
        "errors": [], "artifacts": {"folder": "", "csv": "", "json": ""},
        "verdict": {"status": "INITIALIZING", "message": "J29 has not completed."},
    }


def read_only_text(value):
    return "UNKNOWN" if value is None else "YES" if value else "NO"


def control_summary(snapshot, title):
    if not snapshot:
        return ""
    item = snapshot["controls"][title]
    if not item["present"]:
        return "ABSENT"
    return "{0}:{1}".format(item["type"], item["value_state"])


def csv_rows(report):
    identity, access = report["assembly"], report["access"]
    common = {
        "RUN_TIMESTAMP": report["run_timestamp"], "JOURNAL_BUILD": report["journal_build"],
        "SCHEMA_VERSION": report["schema_version"], "MODE": report["mode"],
        "VERDICT": report["verdict"]["status"], "DB_PART_NO": identity.get("number", ""),
        "DB_PART_REV": identity.get("revision", ""), "ASSEMBLY_NAME": identity.get("name", ""),
        "MANAGED_MODE": "YES" if access.get("managed") else "NO",
        "CHECKOUT_STATE": access.get("checkout_state", ""),
        "CHECKOUT_OWNER": access.get("checkout_owner", ""),
        "CURRENT_USER": access.get("current_user", ""),
        "READ_ONLY": read_only_text(access.get("read_only")),
        "SELECTION_COUNT": report["selection"]["count"],
    }
    summary = dict(common)
    summary.update({
        "ROW_TYPE": "SUMMARY", "ACTION": "RUN_SUMMARY",
        "STATUS": report["verdict"]["status"], "MESSAGE": report["verdict"]["message"],
    })
    rows = [summary]
    for target in report.get("targets", []):
        before = target.get("before")
        after = target.get("after")
        snapshot = after or target.get("after_attempt") or before or {}
        controls = snapshot.get("controls", {})
        row = dict(common)
        row.update({
            "ROW_TYPE": "TARGET", "SELECTION_INDEX": target.get("selection_index", ""),
            "TARGET_STATUS": target.get("status", ""),
            "TARGET_NAME": snapshot.get("name", target.get("selected_name", "")),
            "TARGET_DISPLAY_NAME": snapshot.get("display_name", ""),
            "TARGET_TAG": snapshot.get("tag", target.get("selected_tag", "")),
            "PARENT_TAG": snapshot.get("parent_tag", ""),
            "PROTOTYPE_NAME": snapshot.get("prototype", {}).get("name", ""),
            "PROTOTYPE_TAG": snapshot.get("prototype", {}).get("tag", ""),
            "SUPPRESSED": "YES" if snapshot.get("suppressed") else "NO",
            "REFERENCE_COMPONENT_BEFORE": control_summary(before, REFERENCE_ATTRIBUTE),
            "REFERENCE_COMPONENT_AFTER": control_summary(after, REFERENCE_ATTRIBUTE),
            "PLIST_IGNORE_MEMBER_PRESENT": (
                "YES" if controls.get("PLIST_IGNORE_MEMBER", {}).get("present") else "NO"
            ),
            "PLIST_IGNORE_SUBASSEMBLY_PRESENT": (
                "YES" if controls.get("PLIST_IGNORE_SUBASSEMBLY", {}).get("present") else "NO"
            ),
            "ACTION": target.get("action", {}).get("status", "NO_WRITE"),
            "STATUS": target.get("status", ""), "MESSAGE": target.get("message", ""),
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


def rollback_to_before_batch(session, mark, entries, root, work_part):
    result = {
        "attempted": True, "status": "ROLLBACK_FAILED",
        "verification_errors": [], "error": "",
    }
    try:
        session.UndoToMark(mark, UNDO_MARK_NAME)
    except Exception as error:
        result["error"] = "UndoToMark failed: " + error_text(error)
        return result
    for entry in entries:
        target = entry["report"]
        before = target.get("before")
        if before is None:
            continue
        try:
            after = component_snapshot(entry["component"], root, work_part)
            target["after"] = after
            expect_reference = before["controls"][REFERENCE_ATTRIBUTE]["present"]
            errors = stable_snapshot_errors(
                before, after, expect_reference=expect_reference
            )
            if errors:
                result["verification_errors"].append({
                    "selection_index": target["selection_index"],
                    "target_tag": before["tag"], "errors": errors,
                })
            elif target["action"]["attempted"]:
                target["status"] = "ROLLED_BACK"
                target["message"] = "The batch write was rolled back to the verified baseline."
                target["action"]["status"] = "ROLLED_BACK"
            elif target["status"] == "ELIGIBLE":
                target["status"] = "NOT_ATTEMPTED"
                target["message"] = "The batch stopped before this eligible occurrence was written."
        except Exception as error:
            result["verification_errors"].append({
                "selection_index": target["selection_index"],
                "target_tag": before.get("tag", ""),
                "errors": ["Rollback snapshot failed: " + error_text(error)],
            })
    if result["verification_errors"]:
        result["error"] = "Rollback verification did not restore every selected baseline."
    else:
        result["status"] = "ROLLED_BACK"
    delete_error = delete_undo_mark(session, mark)
    if delete_error:
        result["error"] = " | ".join(
            item for item in (result["error"], "DeleteUndoMark failed: " + delete_error) if item
        )
    return result


def perform_apply_batch(session, work_part, root, entries, report):
    eligible = [
        entry for entry in entries if entry["report"]["status"] == "ELIGIBLE"
    ]
    if not eligible:
        set_verdict(
            report, "ALREADY_REFERENCE_ONLY",
            "Every selected occurrence already has the exact J28-proven Reference-Only attribute; nothing was changed.",
        )
        return None
    try:
        mark = session.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, UNDO_MARK_NAME)
        report["action"]["undo_mark_created"] = True
    except Exception as error:
        report["action"]["error"] = error_text(error)
        set_verdict(report, "BLOCKED_UNDO_MARK", "J29 could not create the required NX undo mark; nothing was changed.")
        return None
    report["action"]["attempted"] = True
    try:
        for entry in eligible:
            component = entry["component"]
            target = entry["report"]
            before = target["before"]
            target["action"]["attempted"] = True
            target["action"]["status"] = "SET_REFERENCE_ONLY"
            report["action"]["attempted_count"] += 1
            try:
                component.SetInstanceUserAttribute(
                    REFERENCE_ATTRIBUTE, -1, "", NXOpen.Update.Option.Now,
                )
                after = component_snapshot(component, root, work_part)
                target["after_attempt"] = after
                errors = stable_snapshot_errors(
                    before, after, expect_reference=True
                )
                target["action"]["verification_errors"] = errors
                if errors:
                    raise VerificationError(errors, snapshot=after)
                target["after"] = after
                target["status"] = "APPLIED_VERIFIED"
                target["message"] = "The exact blank REFERENCE_COMPONENT occurrence attribute was written and verified."
                report["action"]["applied_count"] += 1
            except Exception as error:
                target["action"]["error"] = error_text(error)
                target["status"] = "WRITE_OR_VERIFY_FAILED"
                target["message"] = target["action"]["error"]
                raise
        report["action"]["successful_change_left_undoable"] = True
        set_verdict(
            report, "APPLIED_VERIFIED",
            "Applied and verified {0} occurrence(s); {1} were already Reference-Only. The parent assembly remains unsaved under one visible NX undo mark.".format(
                report["action"]["applied_count"],
                len(entries) - len(eligible),
            ),
        )
        return mark
    except Exception as error:
        report["action"]["error"] = error_text(error)
        rollback = rollback_to_before_batch(session, mark, entries, root, work_part)
        report["rollback"] = rollback
        if rollback["status"] == "ROLLED_BACK":
            set_verdict(report, "ROLLED_BACK", "A batch write or verification failed; every selected occurrence was restored to its verified baseline.")
        else:
            set_verdict(report, "ROLLBACK_FAILED", "The batch failed and J29 could not prove restoration for every occurrence. Use NX Undo and inspect the parent assembly immediately.")
        return None


def write_with_apply_rollback(report, folder, stem, session, work_part, root, entries, mark):
    try:
        return write_outputs(report, folder, stem)
    except Exception as error:
        report["errors"].append("Evidence write failed: " + error_text(error))
        if report["verdict"]["status"] != "APPLIED_VERIFIED" or mark is None:
            raise
    report["action"]["successful_change_left_undoable"] = False
    rollback = rollback_to_before_batch(session, mark, entries, root, work_part)
    report["rollback"] = rollback
    if rollback["status"] == "ROLLED_BACK":
        set_verdict(report, "ROLLED_BACK", "The verified batch was rolled back because paired CSV/JSON evidence could not be completed.")
    else:
        set_verdict(report, "ROLLBACK_FAILED", "Evidence writing failed and J29 could not prove restoration. Use NX Undo and inspect the parent assembly immediately.")
    try:
        return write_outputs(report, folder, stem)
    except Exception as second_error:
        report["errors"].append("Rollback evidence write also failed: " + error_text(second_error))
        report["artifacts"] = {"folder": folder, "csv": "", "json": ""}
        return "", ""


def run(session, selection_manager, run_datetime=None, mode=None):
    selected_mode = clean(mode).upper() if mode is not None else configured_mode()
    if selected_mode not in VALID_MODES:
        raise RuntimeError(
            "Invalid J29 mode {0!r}; expected DRY_RUN or APPLY.".format(selected_mode)
        )
    now = run_datetime or datetime.datetime.now().astimezone()
    parts = safe_property(session, "Parts")
    initial_work_part = safe_property(parts, "Work")
    identity = part_identity(initial_work_part)
    report = base_report(selected_mode, now, identity, empty_access(initial_work_part, session))
    folder, stem = unique_run_folder(identity, now)
    report["artifacts"]["folder"] = folder
    preflight_artifact_folder(folder)

    try:
        max_selection = configured_max_selection()
        report["configuration"]["max_selection"] = max_selection
    except Exception as error:
        report["errors"].append(error_text(error))
        set_verdict(report, "FAILED_CONFIGURATION", error_text(error))
        csv_path, json_path = write_outputs(report, folder, stem)
        return csv_path, json_path, report

    try:
        work_part, root = assembly_context(session)
    except Exception as error:
        report["errors"].append(error_text(error))
        set_verdict(report, "FAILED_CONTEXT", error_text(error))
        csv_path, json_path = write_outputs(report, folder, stem)
        return csv_path, json_path, report

    try:
        components, count = selected_components(selection_manager, max_selection)
        report["selection"]["count"] = count
    except SelectionLimitError as error:
        report["errors"].append(error_text(error))
        set_verdict(report, "BLOCKED_SELECTION_LIMIT", error_text(error))
        csv_path, json_path = write_outputs(report, folder, stem)
        return csv_path, json_path, report
    except Exception as error:
        report["errors"].append(error_text(error))
        set_verdict(report, "BLOCKED_SELECTION", error_text(error))
        csv_path, json_path = write_outputs(report, folder, stem)
        return csv_path, json_path, report

    entries = []
    seen_tags = set()
    required_methods = (
        "HasInstanceUserAttribute", "GetInstanceUserAttribute",
        "SetInstanceUserAttribute",
    )
    for index, component in enumerate(components, 1):
        target = {
            "selection_index": index,
            "selected_name": safe_name(component),
            "selected_tag": object_tag(component),
            "status": "INITIALIZING", "message": "",
            "before": None, "after_attempt": None, "after": None,
            "action": {
                "attempted": False, "status": "NO_WRITE",
                "verification_errors": [], "error": "",
            },
        }
        report["targets"].append(target)
        entry = {"component": component, "report": target}
        entries.append(entry)
        try:
            if not all(callable(getattr(component, name, None)) for name in required_methods):
                raise RuntimeError(
                    "The selected object is not an assembly component. Select component rows in Assembly Navigator."
                )
            before = component_snapshot(component, root, work_part)
            target["before"] = before
            tag = before["tag"]
            if tag in seen_tags:
                raise RuntimeError("This component occurrence is duplicated in the selection.")
            seen_tags.add(tag)
            reference = before["controls"][REFERENCE_ATTRIBUTE]
            conflicts = [
                title for title in CONFLICT_ATTRIBUTES
                if before["controls"][title]["present"]
            ]
            if conflicts:
                target["status"] = "BLOCKED_CONTROL_CONFLICT"
                target["message"] = "The occurrence already has {0}; J29 will not combine or replace native BoM controls.".format(
                    ", ".join(conflicts)
                )
            elif reference["present"]:
                errors = reference_contract_errors(before, require_present=True)
                target["action"]["verification_errors"] = errors
                if errors:
                    target["status"] = "BLOCKED_NONSTANDARD_REFERENCE"
                    target["message"] = "REFERENCE_COMPONENT exists but does not match the J28 V2 contract; J29 will not overwrite it."
                else:
                    target["status"] = "ALREADY_REFERENCE_ONLY"
                    target["message"] = "The exact Reference-Only occurrence attribute is already present."
                    target["after"] = before
            else:
                target["status"] = "ELIGIBLE"
                target["message"] = "The occurrence is eligible for the exact blank REFERENCE_COMPONENT attribute."
        except Exception as error:
            target["status"] = "BLOCKED_SELECTION"
            target["message"] = error_text(error)

    blockers = [
        target for target in report["targets"]
        if target["status"].startswith("BLOCKED_")
    ]
    mark = None
    if blockers:
        set_verdict(
            report, "BLOCKED_BATCH",
            "The atomic batch contains {0} blocked target(s); nothing was changed. Review the per-target CSV/JSON statuses.".format(len(blockers)),
        )
    elif selected_mode == "DRY_RUN":
        eligible_count = 0
        for target in report["targets"]:
            if target["status"] == "ELIGIBLE":
                target["status"] = "DRY_RUN_READY"
                target["after"] = target["before"]
                eligible_count += 1
        if eligible_count:
            set_verdict(report, "DRY_RUN_READY", "The complete selection passed preflight; APPLY would update {0} occurrence(s).".format(eligible_count))
        else:
            set_verdict(report, "ALREADY_REFERENCE_ONLY", "Every selected occurrence is already Reference-Only; nothing would change.")
    elif not any(target["status"] == "ELIGIBLE" for target in report["targets"]):
        set_verdict(
            report, "ALREADY_REFERENCE_ONLY",
            "Every selected occurrence is already Reference-Only; no write access or undo mark was required.",
        )
    else:
        access = inspect_write_access(session, work_part)
        report["access"] = access
        if not access["allowed"]:
            set_verdict(report, "BLOCKED_WRITE_ACCESS", access["message"] or "J29 could not prove parent-assembly write access.")
        else:
            mark = perform_apply_batch(session, work_part, root, entries, report)

    csv_path, json_path = write_with_apply_rollback(
        report, folder, stem, session, work_part, root, entries, mark
    )
    return csv_path, json_path, report


def main():
    session = NXOpen.Session.GetSession()
    mode = configured_mode()
    log_line(session, "=" * 72)
    log_line(session, "J29 SET SELECTED COMPONENTS TO REFERENCE-ONLY")
    log_line(session, "Build: " + BUILD)
    log_line(session, "Mode: " + mode)
    log_line(
        session,
        "Preselect 1-{0} direct component rows in Assembly Navigator.".format(
            configured_max_selection()
        ),
    )
    log_line(session, "J29 never scans/loads the tree, checks out, checks in, or saves.")
    log_line(session, "=" * 72)
    try:
        selection_manager = NXOpen.UI.GetUI().SelectionManager
        csv_path, json_path, report = run(session, selection_manager, mode=mode)
        log_line(session, "Verdict: " + report["verdict"]["status"])
        log_line(session, report["verdict"]["message"])
        if csv_path:
            log_line(session, "CSV: " + csv_path)
        if json_path:
            log_line(session, "JSON: " + json_path)
        if report["verdict"]["status"] == "APPLIED_VERIFIED":
            log_line(session, "The parent assembly is modified but UNSAVED. Inspect and save manually if correct.")
            log_line(session, "Undo once to revert: " + UNDO_MARK_NAME)
        elif mode == "DRY_RUN" and report["verdict"]["status"] == "DRY_RUN_READY":
            log_line(session, "Review both artifacts, ensure the parent assembly is writable, then rerun in APPLY with the same components selected.")
    except Exception as error:
        log_line(session, "J29 FAILED: " + error_text(error))
        log_line(session, traceback.format_exc())
        raise


def get_unload_option(dummy):
    # AtTermination avoids invalidating NX bridge objects while the journal
    # teardown/listing-window work is still completing on large assemblies.
    return NXOpen.Session.LibraryUnloadOption.AtTermination


if __name__ == "__main__":
    main()
