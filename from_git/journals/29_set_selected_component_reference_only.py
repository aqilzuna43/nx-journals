"""
Journal 29 - Set One Selected Component to Reference-Only

Adds the exact occurrence attribute proven by J28 V2 to one preselected
direct component of the active assembly:

    REFERENCE_COMPONENT = ""  (string, index -1, component-instance scope)

DRY_RUN is the default. APPLY requires the active assembly to be both Work
and Display part and already writable. J29 never scans the assembly tree,
loads components, checks out, checks in, or saves. A verified APPLY remains
unsaved under one visible NX undo mark. Any write/verification/evidence
failure triggers UndoToMark and a baseline reread.

Target: NX 2312 and NX X 2506 embedded Python
Run via: NX > Tools > Journal > Play
"""

import csv
import datetime
import json
import os
import traceback

import NXOpen


# User setting. Run DRY_RUN first; change to APPLY only after reviewing it.
USER_MODE = "DRY_RUN"

BUILD = "J29-NX2506-SELECTED-REFERENCE-ONLY-V1"
SCHEMA_VERSION = 1
OUTPUT_FOLDER = "NX_REFERENCE_ONLY"
UNDO_MARK_NAME = "J29 Set selected component Reference-Only"
VALID_MODES = ("DRY_RUN", "APPLY")
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
    "READ_ONLY", "SELECTION_COUNT", "TARGET_NAME", "TARGET_DISPLAY_NAME",
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


def selected_component(selection_manager):
    try:
        count = int(selection_manager.GetNumSelectedObjects())
    except Exception as error:
        raise RuntimeError("Could not inspect NX preselection: " + error_text(error))
    if count != 1:
        raise RuntimeError(
            "Select exactly one component in Assembly Navigator; observed {0} selected object(s).".format(count)
        )
    try:
        selected = selection_manager.GetSelectedTaggedObject(0)
    except Exception as error:
        raise RuntimeError("Could not read the selected NX object: " + error_text(error))
    required_methods = (
        "HasInstanceUserAttribute", "GetInstanceUserAttribute",
        "SetInstanceUserAttribute",
    )
    if not all(callable(getattr(selected, name, None)) for name in required_methods):
        raise RuntimeError(
            "The selected object is not an assembly component. Select its component row in Assembly Navigator."
        )
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
            "raw_checkout": "", "message": message,
        }
    state, owner, raw = checkout_status(part)
    user = current_teamcenter_user(session)
    allowed = bool(
        state == "CHECKED_OUT" and owner and user
        and owner.casefold() == user.casefold() and read_only is False
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
    elif owner.casefold() != user.casefold():
        messages.append("Parent assembly is checked out by another user: {0}.".format(owner))
    if read_only is True:
        messages.append("Parent assembly is read-only in this NX session.")
    elif read_only is None:
        messages.append("Parent assembly read-only state is unavailable.")
    return {
        "allowed": allowed, "managed": True, "checkout_state": state,
        "checkout_owner": owner, "current_user": user, "read_only": read_only,
        "raw_checkout": raw, "message": " ".join(messages),
    }


def empty_access(part=None, session=None):
    return {
        "allowed": None,
        "managed": bool(part is not None and managed_mode(session, part)),
        "checkout_state": "NOT_INSPECTED", "checkout_owner": "",
        "current_user": "", "read_only": read_only_value(part) if part is not None else None,
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
            "scope": "ONE_PRESELECTED_DIRECT_COMPONENT_OCCURRENCE",
            "attribute_title": REFERENCE_ATTRIBUTE,
            "attribute_type": "STRING", "attribute_index": -1,
            "attribute_value": "", "force_load": False,
            "automatic_checkout": False, "automatic_save": False,
            "automatic_checkin": False,
        },
        "assembly": identity, "selection": {"count": 0, "source": "NX_PRESELECTION"},
        "access": access, "before": None, "after_attempt": None, "after": None,
        "action": {
            "api": "Component.SetInstanceUserAttribute(REFERENCE_COMPONENT, -1, blank, Update.Option.Now)",
            "attempted": False, "undo_mark_name": UNDO_MARK_NAME,
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
    before, after = report.get("before"), report.get("after")
    target = after or before or {}
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
    if target:
        row = dict(common)
        row.update({
            "ROW_TYPE": "TARGET", "TARGET_NAME": target.get("name", ""),
            "TARGET_DISPLAY_NAME": target.get("display_name", ""), "TARGET_TAG": target.get("tag", ""),
            "PARENT_TAG": target.get("parent_tag", ""),
            "PROTOTYPE_NAME": target.get("prototype", {}).get("name", ""),
            "PROTOTYPE_TAG": target.get("prototype", {}).get("tag", ""),
            "SUPPRESSED": "YES" if target.get("suppressed") else "NO",
            "REFERENCE_COMPONENT_BEFORE": control_summary(before, REFERENCE_ATTRIBUTE),
            "REFERENCE_COMPONENT_AFTER": control_summary(after, REFERENCE_ATTRIBUTE),
            "PLIST_IGNORE_MEMBER_PRESENT": "YES" if target["controls"]["PLIST_IGNORE_MEMBER"]["present"] else "NO",
            "PLIST_IGNORE_SUBASSEMBLY_PRESENT": "YES" if target["controls"]["PLIST_IGNORE_SUBASSEMBLY"]["present"] else "NO",
            "ACTION": "SET_REFERENCE_ONLY" if report["action"]["attempted"] else "NO_WRITE",
            "STATUS": report["verdict"]["status"], "MESSAGE": report["verdict"]["message"],
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


def rollback_to_before(session, mark, component, root, work_part, before):
    result = {
        "attempted": True, "status": "ROLLBACK_FAILED",
        "verification_errors": [], "error": "",
    }
    try:
        session.UndoToMark(mark, UNDO_MARK_NAME)
    except Exception as error:
        result["error"] = "UndoToMark failed: " + error_text(error)
        try:
            after = component_snapshot(component, root, work_part)
        except Exception as capture_error:
            after = None
            result["error"] += " | Rollback snapshot failed: " + error_text(capture_error)
        return result, after
    try:
        after = component_snapshot(component, root, work_part)
        errors = stable_snapshot_errors(before, after, expect_reference=False)
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
            item for item in (result["error"], "DeleteUndoMark failed: " + delete_error) if item
        )
    return result, after


def perform_apply(session, work_part, root, component, before, report):
    try:
        mark = session.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, UNDO_MARK_NAME)
        report["action"]["undo_mark_created"] = True
    except Exception as error:
        report["action"]["error"] = error_text(error)
        set_verdict(report, "BLOCKED_UNDO_MARK", "J29 could not create the required NX undo mark; nothing was changed.")
        return None
    report["action"]["attempted"] = True
    try:
        component.SetInstanceUserAttribute(
            REFERENCE_ATTRIBUTE,
            -1,
            "",
            NXOpen.Update.Option.Now,
        )
        after = component_snapshot(component, root, work_part)
        report["after_attempt"] = after
        errors = stable_snapshot_errors(before, after, expect_reference=True)
        report["action"]["verification_errors"] = errors
        if errors:
            raise VerificationError(errors, snapshot=after)
        report["after"] = after
        report["action"]["successful_change_left_undoable"] = True
        set_verdict(
            report, "APPLIED_VERIFIED",
            "The selected component occurrence now has the exact J28-proven blank string REFERENCE_COMPONENT attribute; the parent assembly remains unsaved under one visible NX undo mark.",
        )
        return mark
    except Exception as error:
        report["action"]["error"] = error_text(error)
        if isinstance(error, VerificationError) and error.snapshot is not None:
            report["after_attempt"] = error.snapshot
        rollback, final_snapshot = rollback_to_before(
            session, mark, component, root, work_part, before
        )
        report["rollback"] = rollback
        report["after"] = final_snapshot
        if rollback["status"] == "ROLLED_BACK":
            set_verdict(report, "ROLLED_BACK", "The occurrence write failed or did not match the J28 contract; the original state was restored.")
        else:
            set_verdict(report, "ROLLBACK_FAILED", "The occurrence write failed and J29 could not prove restoration. Use NX Undo and inspect the parent assembly immediately.")
        return None


def write_with_apply_rollback(report, folder, stem, session, work_part, root, component, before, mark):
    try:
        return write_outputs(report, folder, stem)
    except Exception as error:
        report["errors"].append("Evidence write failed: " + error_text(error))
        if report["verdict"]["status"] != "APPLIED_VERIFIED" or mark is None:
            raise
    report["action"]["successful_change_left_undoable"] = False
    rollback, final_snapshot = rollback_to_before(
        session, mark, component, root, work_part, before
    )
    report["rollback"] = rollback
    report["after"] = final_snapshot
    if rollback["status"] == "ROLLED_BACK":
        set_verdict(report, "ROLLED_BACK", "The verified occurrence write was rolled back because paired CSV/JSON evidence could not be completed.")
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
        work_part, root = assembly_context(session)
    except Exception as error:
        report["errors"].append(error_text(error))
        set_verdict(report, "FAILED_CONTEXT", error_text(error))
        csv_path, json_path = write_outputs(report, folder, stem)
        return csv_path, json_path, report

    try:
        component, count = selected_component(selection_manager)
        report["selection"]["count"] = count
        before = component_snapshot(component, root, work_part)
    except Exception as error:
        report["errors"].append(error_text(error))
        set_verdict(report, "BLOCKED_SELECTION", error_text(error))
        csv_path, json_path = write_outputs(report, folder, stem)
        return csv_path, json_path, report

    report["before"] = before
    reference = before["controls"][REFERENCE_ATTRIBUTE]
    conflicts = [
        title for title in CONFLICT_ATTRIBUTES
        if before["controls"][title]["present"]
    ]
    mark = None
    if conflicts:
        report["after"] = before
        set_verdict(
            report, "BLOCKED_CONTROL_CONFLICT",
            "The selected occurrence already has {0}; J29 will not combine or replace native BoM controls.".format(", ".join(conflicts)),
        )
    elif reference["present"]:
        contract_errors = reference_contract_errors(before, require_present=True)
        report["after"] = before
        if contract_errors:
            report["action"]["verification_errors"] = contract_errors
            set_verdict(report, "BLOCKED_NONSTANDARD_REFERENCE", "REFERENCE_COMPONENT exists but does not match the J28 V2 contract; J29 will not overwrite it.")
        else:
            set_verdict(report, "ALREADY_REFERENCE_ONLY", "The selected occurrence already has the exact J28-proven Reference-Only attribute; nothing was changed.")
    elif selected_mode == "DRY_RUN":
        report["after"] = before
        set_verdict(report, "DRY_RUN_READY", "The selected direct component is eligible; APPLY would add one blank string REFERENCE_COMPONENT occurrence attribute.")
    else:
        access = inspect_write_access(session, work_part)
        report["access"] = access
        if not access["allowed"]:
            report["after"] = before
            set_verdict(report, "BLOCKED_WRITE_ACCESS", access["message"] or "J29 could not prove parent-assembly write access.")
        else:
            mark = perform_apply(session, work_part, root, component, before, report)

    csv_path, json_path = write_with_apply_rollback(
        report, folder, stem, session, work_part, root, component, before, mark
    )
    return csv_path, json_path, report


def main():
    session = NXOpen.Session.GetSession()
    mode = configured_mode()
    log_line(session, "=" * 72)
    log_line(session, "J29 SET ONE SELECTED COMPONENT TO REFERENCE-ONLY")
    log_line(session, "Build: " + BUILD)
    log_line(session, "Mode: " + mode)
    log_line(session, "Preselect exactly one direct component row in Assembly Navigator.")
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
            log_line(session, "Review both artifacts, ensure the parent assembly is writable, then set USER_MODE = \"APPLY\" and rerun with the same component selected.")
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
