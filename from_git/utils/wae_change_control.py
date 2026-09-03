"""Shared NX X 2506 implementation for J30 Freeze and J31 Unfreeze."""

import datetime
import json
import os
import re
import traceback

import NXOpen
import NXOpen.PDM


VALID_ACTIONS = ("FREEZE", "UNFREEZE")
VALID_MODES = ("DRY_RUN", "APPLY")
ATTRIBUTE_CATEGORY = "WAEItem"
WAE_VERSION_TITLE = "WAE_VERSION"
DB_PART_NO_TITLE = "DB_PART_NO"
DB_PART_REV_TITLE = "DB_PART_REV"
OUTPUT_FOLDER = "NX_WAE_CHANGE_CONTROL"


def clean(value):
    return "" if value is None else str(value).strip()


def error_text(error):
    code = clean(getattr(error, "ErrorCode", ""))
    suffix = " [{0}]".format(code) if code else ""
    return "{0}{1}".format(clean(error) or type(error).__name__, suffix)


def dispose(value):
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


def safe_property(value, name, default=None):
    try:
        result = getattr(value, name)
        return result() if callable(result) else result
    except Exception:
        return default


def object_key(value):
    tag = safe_property(value, "Tag")
    return ("TAG", clean(tag)) if tag is not None else ("PY", id(value))


def same_nx_object(first, second):
    return first is second or (
        first is not None and second is not None and object_key(first) == object_key(second)
    )


def part_identifier(part):
    for name in ("JournalIdentifier", "FullPath", "Name", "Leaf"):
        value = safe_property(part, name, "")
        if clean(value):
            return clean(value)
    return "<unknown>"


def managed_mode(session, part):
    if bool(safe_property(session, "IsManagedMode", False)):
        return True
    return part_identifier(part).upper().startswith("@DB/")


def selected_component_target(selection_manager):
    try:
        count = int(selection_manager.GetNumSelectedObjects())
    except Exception as error:
        raise RuntimeError("Could not inspect NX preselection: " + error_text(error))
    if count != 1:
        raise RuntimeError(
            "Preselect exactly one component row in Assembly Navigator; found {0}.".format(
                count
            )
        )
    try:
        component = selection_manager.GetSelectedTaggedObject(0)
    except Exception as error:
        raise RuntimeError("Could not read the selected NX object: " + error_text(error))
    prototype = safe_property(component, "Prototype")
    if prototype is None:
        raise RuntimeError(
            "The selected NX object is not a loaded assembly component prototype."
        )
    suppressed = safe_property(component, "IsSuppressed")
    if suppressed is True:
        raise RuntimeError("The selected component is suppressed.")
    if safe_property(prototype, "PDMPart") is None:
        raise RuntimeError(
            "The selected component prototype has no PDMPart; fully load a managed CAD component."
        )
    return component, prototype


def read_identity(part, title):
    method = getattr(part, "GetStringAttribute", None)
    if not callable(method):
        return ""
    try:
        return clean(method(title))
    except Exception:
        return ""


def attribute_type_name(info):
    raw = clean(safe_property(info, "Type"))
    normalized = raw.split(".")[-1].upper()
    names = {
        "0": "INVALID", "1": "BOOLEAN", "3": "INTEGER", "4": "REAL",
        "5": "STRING", "6": "TIME", "7": "REFERENCE", "100": "ANY",
    }
    return names.get(normalized, normalized or "UNKNOWN")


def read_wae_attribute(part):
    iterator = None
    try:
        iterator = part.CreateAttributeIterator()
        iterator.SetIncludeOnlyCategory(ATTRIBUTE_CATEGORY)
        iterator.SetIncludeOnlyTitle(WAE_VERSION_TITLE)
        iterator.SetIncludeAlsoUnset(True)
        matches = []
        for info in part.GetUserAttributes(iterator):
            if (
                clean(safe_property(info, "Category")) == ATTRIBUTE_CATEGORY
                and clean(safe_property(info, "Title")) == WAE_VERSION_TITLE
            ):
                matches.append(info)
        if len(matches) != 1:
            raise RuntimeError(
                "Expected exactly one WAEItem/WAE_VERSION attribute; found {0}.".format(
                    len(matches)
                )
            )
        info = matches[0]
        return {
            "value": clean(safe_property(info, "StringValue", "")),
            "type": attribute_type_name(info),
            "unset": bool(safe_property(info, "Unset", False)),
            "locked": bool(safe_property(info, "Locked", False)),
            "owned_by_system": bool(safe_property(info, "OwnedBySystem", False)),
            "pdm_based": bool(safe_property(info, "PdmBased", False)),
            "not_saved": bool(safe_property(info, "NotSaved", False)),
        }
    finally:
        dispose(iterator)


def parse_wae_version(value):
    raw = clean(value)
    if not raw:
        raise RuntimeError("WAE_VERSION is blank.")
    if not re.fullmatch(r"[1-9][0-9]*", raw):
        raise RuntimeError(
            "WAE_VERSION must be a positive whole number; found {0!r}.".format(raw)
        )
    return int(raw)


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


def identity_tokens(value):
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
    return tokens


def same_teamcenter_user(owner, current_user):
    owner_tokens = identity_tokens(owner)
    user_tokens = identity_tokens(current_user)
    return bool(owner_tokens and user_tokens and owner_tokens.intersection(user_tokens))


def checkout_snapshot(session, part):
    pdm_part = safe_property(part, "PDMPart")
    method = getattr(pdm_part, "GetCheckedoutStatusAndUser", None)
    if not callable(method):
        return {
            "state": "UNKNOWN", "owner": "", "current_user": "",
            "owner_is_current_user": None,
            "raw": "PDMPart.GetCheckedoutStatusAndUser unavailable",
        }
    try:
        raw = method()
    except TypeError:
        try:
            raw = method(False, "")
        except Exception as error:
            return {
                "state": "UNKNOWN", "owner": "", "current_user": "",
                "owner_is_current_user": None, "raw": error_text(error),
            }
    except Exception as error:
        return {
            "state": "UNKNOWN", "owner": "", "current_user": "",
            "owner_is_current_user": None, "raw": error_text(error),
        }
    state, owner = checkout_result(raw)
    get_user = getattr(safe_property(session, "PdmSession"), "GetUserName", None)
    try:
        current_user = clean(get_user()) if callable(get_user) else ""
    except Exception:
        current_user = ""
    return {
        "state": state,
        "owner": owner,
        "current_user": current_user,
        "owner_is_current_user": (
            same_teamcenter_user(owner, current_user) if state == "CHECKED_OUT" else None
        ),
        "raw": repr(raw)[:2000],
    }


def read_only_value(part):
    value = safe_property(part, "IsReadOnly")
    return None if value is None else bool(value)


def target_snapshot(session, component, part):
    wae = read_wae_attribute(part)
    version = parse_wae_version(wae["value"])
    part_number = read_identity(part, DB_PART_NO_TITLE)
    revision = read_identity(part, DB_PART_REV_TITLE)
    if not managed_mode(session, part):
        raise RuntimeError("The selected component is not positively Teamcenter-managed.")
    if not part_number:
        raise RuntimeError("DB_PART_NO is blank or unavailable.")
    if not revision:
        raise RuntimeError("DB_PART_REV is blank or unavailable.")
    return {
        "component_name": clean(safe_property(component, "DisplayName"))
        or clean(safe_property(component, "Name")),
        "component_tag": clean(safe_property(component, "Tag")),
        "part_identifier": part_identifier(part),
        "part_number": part_number,
        "db_part_rev": revision,
        "wae_version": version,
        "wae_version_raw": wae["value"],
        "wae_attribute": wae,
        "checkout": checkout_snapshot(session, part),
        "read_only": read_only_value(part),
        "part_modified": bool(safe_property(part, "IsModified", False)),
    }


def validate_owned_checkout(snapshot):
    checkout = snapshot["checkout"]
    if checkout["state"] != "CHECKED_OUT":
        return "Selected component is not checked out."
    if not checkout["owner"]:
        return "Checkout owner is unavailable."
    if not checkout["current_user"]:
        return "Current Teamcenter user is unavailable."
    if checkout["owner_is_current_user"] is not True:
        return "Selected component is checked out by another user: {0}.".format(
            checkout["owner"]
        )
    if snapshot["read_only"] is not False:
        return "Selected component is not positively writable after checkout."
    return ""


def save_part(part):
    status = None
    try:
        status = part.Save(
            NXOpen.BasePart.SaveComponents.FalseValue,
            NXOpen.BasePart.CloseAfterSave.FalseValue,
        )
        unsaved_parts = int(safe_property(status, "NumberUnsavedParts", 0) or 0)
        unsaved_objects = int(safe_property(status, "NumberUnsavedObjects", 0) or 0)
        if unsaved_parts or unsaved_objects:
            raise RuntimeError(
                "NX reported {0} unsaved part(s) and {1} unsaved object(s).".format(
                    unsaved_parts, unsaved_objects
                )
            )
    finally:
        dispose(status)


def checkout_part(part):
    pdm_part = safe_property(part, "PDMPart")
    method = getattr(pdm_part, "CheckoutParts", None)
    if not callable(method):
        raise RuntimeError("PDMPart.CheckoutParts is unavailable.")
    checkout_input = NXOpen.PDM.PdmPart.CheckoutInput(
        "J31 WAE unfreeze", "", True, True, False
    )
    errors = None
    try:
        errors = method([part], checkout_input)
        return repr(errors)[:2000]
    finally:
        dispose(errors)


def checkin_part(part):
    pdm_part = safe_property(part, "PDMPart")
    method = getattr(pdm_part, "CheckinParts", None)
    if not callable(method):
        raise RuntimeError("PDMPart.CheckinParts is unavailable.")
    constructor = NXOpen.PDM.PdmPart.CheckinInput
    try:
        checkin_input = constructor(True, True, False)
    except TypeError:
        checkin_input = constructor()
        checkin_input.AllowRemote = True
        checkin_input.ExplicitCheckIn = True
        checkin_input.IncludeSecondary = False
    errors = None
    try:
        errors = method([part], checkin_input)
        return repr(errors)[:2000]
    finally:
        dispose(errors)


def write_wae_version(session, part, expected):
    before = read_wae_attribute(part)
    if before["type"] != "STRING":
        raise RuntimeError("WAE_VERSION is not a string attribute.")
    if before["locked"] or before["owned_by_system"] or before["pdm_based"]:
        raise RuntimeError("Runtime WAE_VERSION flags prohibit this write.")
    builder = None
    try:
        builder = session.AttributeManager.CreateAttributePropertiesBuilder(
            part, [part], NXOpen.AttributePropertiesBuilder.OperationType.Save
        )
        builder.Category = ATTRIBUTE_CATEGORY
        builder.Title = WAE_VERSION_TITLE
        builder.DataType = NXOpen.AttributePropertiesBaseBuilder.DataTypeOptions.String
        builder.StringValue = str(expected)
        builder.Commit()
    finally:
        dispose(builder)


def base_report(action, build, mode):
    return {
        "build": build,
        "timestamp": datetime.datetime.now().astimezone().isoformat(timespec="seconds"),
        "action": action,
        "mode": mode,
        "scope": "ONE_PRESELECTED_COMPONENT_PROTOTYPE",
        "result": "BLOCKED",
        "message": "",
        "before": {},
        "after": {},
        "operations": {
            "checkout_attempted": False,
            "wae_write_attempted": False,
            "save_attempted": False,
            "checkin_attempted": False,
            "formal_revision_created": False,
        },
        "operation_raw": "",
    }


def freeze(session, component, part, report):
    before = target_snapshot(session, component, part)
    report["before"] = before
    state = before["checkout"]["state"]
    if state == "CHECKED_IN":
        report["after"] = target_snapshot(session, component, part)
        report["result"] = "FROZEN_VERIFIED"
        report["message"] = "Selected component is checked in at its current WAE baseline."
        return report
    owned_error = validate_owned_checkout(before)
    if owned_error:
        report["message"] = owned_error
        return report
    if report["mode"] == "DRY_RUN":
        report["result"] = "DRY_RUN_READY_TO_FREEZE"
        report["message"] = "Would save and check in only the selected component prototype."
        return report
    try:
        report["operations"]["save_attempted"] = True
        save_part(part)
        saved = target_snapshot(session, component, part)
        if saved["db_part_rev"] != before["db_part_rev"]:
            raise RuntimeError("DB_PART_REV changed during the freeze save.")
        if saved["wae_version"] != before["wae_version"]:
            raise RuntimeError("WAE_VERSION changed during the freeze save.")
        report["operations"]["checkin_attempted"] = True
        report["operation_raw"] = checkin_part(part)
        after = target_snapshot(session, component, part)
        report["after"] = after
        if after["checkout"]["state"] != "CHECKED_IN":
            raise RuntimeError(
                "Check-in postcondition failed; observed {0}.".format(
                    after["checkout"]["state"]
                )
            )
        if after["db_part_rev"] != before["db_part_rev"]:
            raise RuntimeError("DB_PART_REV changed during freeze.")
        if after["wae_version"] != before["wae_version"]:
            raise RuntimeError("WAE_VERSION changed during freeze.")
        report["result"] = "FROZEN_CHECKED_IN"
        report["message"] = "Selected component was saved, checked in, and verified."
    except Exception as error:
        report["result"] = "FREEZE_FAILED"
        report["message"] = error_text(error)
        try:
            report["after"] = target_snapshot(session, component, part)
        except Exception:
            pass
    return report


def unfreeze(session, component, part, report):
    before = target_snapshot(session, component, part)
    report["before"] = before
    if before["checkout"]["state"] != "CHECKED_IN":
        report["message"] = (
            "UNFREEZE requires CHECKED_IN. Reruns against a checked-out component are blocked."
        )
        return report
    next_version = before["wae_version"] + 1
    report["next_wae_version"] = next_version
    if report["mode"] == "DRY_RUN":
        report["result"] = "DRY_RUN_READY_TO_UNFREEZE"
        report["message"] = "Would checkout this component and advance WAE_VERSION to {0}.".format(
            next_version
        )
        return report
    checkout_may_have_changed_state = False
    mark = None
    mark_name = "J31 WAE_VERSION increment"
    try:
        report["operations"]["checkout_attempted"] = True
        report["operation_raw"] = checkout_part(part)
        checkout_may_have_changed_state = True
        checked_out = target_snapshot(session, component, part)
        owned_error = validate_owned_checkout(checked_out)
        if owned_error:
            raise RuntimeError("Checkout postcondition failed: " + owned_error)
        if checked_out["db_part_rev"] != before["db_part_rev"]:
            raise RuntimeError("DB_PART_REV changed during checkout.")
        if checked_out["wae_version"] != before["wae_version"]:
            raise RuntimeError("WAE_VERSION changed before the controlled increment.")

        mark = session.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, mark_name)
        report["operations"]["wae_write_attempted"] = True
        write_wae_version(session, part, next_version)
        immediate = read_wae_attribute(part)
        if parse_wae_version(immediate["value"]) != next_version:
            raise RuntimeError("Immediate WAE_VERSION reread did not match the increment.")
        report["operations"]["save_attempted"] = True
        save_part(part)
        after = target_snapshot(session, component, part)
        report["after"] = after
        owned_error = validate_owned_checkout(after)
        if owned_error:
            raise RuntimeError("Post-save checkout state failed: " + owned_error)
        if after["db_part_rev"] != before["db_part_rev"]:
            raise RuntimeError("DB_PART_REV changed during unfreeze.")
        if after["wae_version"] != next_version:
            raise RuntimeError("Saved WAE_VERSION does not match the controlled increment.")
        report["result"] = "UNFROZEN_READY_FOR_EDIT"
        report["message"] = (
            "Selected component remains checked out at WAE_VERSION {0}; CAD editing may begin."
        ).format(next_version)
    except Exception as error:
        if checkout_may_have_changed_state and mark is not None and not report["operations"]["save_attempted"]:
            try:
                session.UndoToMark(mark, mark_name)
            except Exception:
                pass
        report["result"] = (
            "RECOVERY_REQUIRED_CHECKOUT_ATTEMPTED"
            if report["operations"]["checkout_attempted"] else "UNFREEZE_FAILED"
        )
        report["message"] = error_text(error)
        try:
            report["after"] = target_snapshot(session, component, part)
        except Exception:
            pass
    finally:
        if mark is not None:
            try:
                session.DeleteUndoMark(mark, mark_name)
            except Exception:
                pass
    return report


def execute(action, session, selection_manager, build, mode):
    action = clean(action).upper()
    mode = clean(mode).upper()
    if action not in VALID_ACTIONS:
        raise RuntimeError("Action must be FREEZE or UNFREEZE.")
    if mode not in VALID_MODES:
        raise RuntimeError("Mode must be DRY_RUN or APPLY.")
    report = base_report(action, build, mode)
    try:
        component, part = selected_component_target(selection_manager)
        return (
            freeze(session, component, part, report)
            if action == "FREEZE"
            else unfreeze(session, component, part, report)
        )
    except Exception as error:
        report["message"] = error_text(error)
        return report


def io_root():
    configured = clean(os.environ.get("NX_JOURNALS_IO_DIR"))
    if configured:
        return os.path.abspath(configured)
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    return desktop if os.path.isdir(desktop) else os.getcwd()


def write_report(report):
    folder = os.path.join(io_root(), OUTPUT_FOLDER)
    os.makedirs(folder, exist_ok=True)
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    path = os.path.join(folder, "{0}_{1}.json".format(report["action"], stamp))
    with open(path, "w", encoding="utf-8") as handle:
        json.dump(report, handle, indent=2, sort_keys=True)
    return path


def log_line(session, message):
    text = str(message)
    try:
        window = session.ListingWindow
        window.Open()
        window.WriteLine(text)
    except Exception:
        try:
            print(text)
        except Exception:
            pass


def run_ui(action, build, user_mode, environment_name):
    session = NXOpen.Session.GetSession()
    mode = clean(os.environ.get(environment_name) or user_mode).upper()
    log_line(session, "=" * 72)
    log_line(session, "{0} | {1}".format(build, mode))
    log_line(session, "Preselect exactly one component row in Assembly Navigator.")
    try:
        selection_manager = NXOpen.UI.GetUI().SelectionManager
        report = execute(action, session, selection_manager, build, mode)
    except Exception as error:
        report = base_report(action, build, mode)
        report["result"] = "FAILED"
        report["message"] = error_text(error)
        report["traceback"] = traceback.format_exc()
    try:
        path = write_report(report)
    except Exception as error:
        path = ""
        report["message"] += " | Could not write audit JSON: " + error_text(error)
    log_line(session, "Result: " + report["result"])
    log_line(session, report["message"])
    before = report.get("before") or {}
    after = report.get("after") or {}
    if before:
        log_line(
            session,
            "Before: {0}/{1} WAE {2} {3}".format(
                before.get("part_number", ""), before.get("db_part_rev", ""),
                before.get("wae_version_raw", ""),
                (before.get("checkout") or {}).get("state", ""),
            ),
        )
    if after:
        log_line(
            session,
            "After:  {0}/{1} WAE {2} {3}".format(
                after.get("part_number", ""), after.get("db_part_rev", ""),
                after.get("wae_version_raw", ""),
                (after.get("checkout") or {}).get("state", ""),
            ),
        )
    if path:
        log_line(session, "Audit JSON: " + path)
    log_line(session, "=" * 72)
    return report
