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
FREEZE_WORKFLOW = "Part_Freeze_Process"
UNFREEZE_WORKFLOW = "Part_Unfreeze_Process"


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


def component_target(component, index):
    """Validate one selected Assembly Navigator component."""
    prototype = safe_property(component, "Prototype")
    if prototype is None:
        raise RuntimeError(
            "Selected object {0} is not a loaded Assembly Navigator component.".format(
                index + 1
            )
        )
    suppressed = safe_property(component, "IsSuppressed")
    if suppressed is True:
        raise RuntimeError("Selected component {0} is suppressed.".format(index + 1))
    if safe_property(prototype, "PDMPart") is None:
        raise RuntimeError(
            "Selected component {0} has no PDMPart; fully load a managed CAD component.".format(
                index + 1
            )
        )
    return prototype


def selected_or_work_targets(session, selection_manager):
    """Resolve selected unique prototypes, or the active work part if none."""
    try:
        count = int(selection_manager.GetNumSelectedObjects())
    except Exception as error:
        raise RuntimeError("Could not inspect NX preselection: " + error_text(error))

    if count == 0:
        parts = safe_property(session, "Parts")
        work_part = safe_property(parts, "Work")
        if work_part is None:
            raise RuntimeError(
                "No Assembly Navigator components are selected and there is no active work part."
            )
        if safe_property(work_part, "PDMPart") is None:
            raise RuntimeError("The active work part has no PDMPart; open a managed CAD part.")
        return [{
            "component": None,
            "part": work_part,
            "source": "ACTIVE_WORK_PART",
            "selected_indexes": [],
            "occurrence_count": 1,
        }], count

    targets = []
    targets_by_key = {}
    for index in range(count):
        try:
            component = selection_manager.GetSelectedTaggedObject(index)
        except Exception as error:
            raise RuntimeError(
                "Could not read selected NX object {0}: {1}".format(
                    index + 1, error_text(error)
                )
            )
        prototype = component_target(component, index)
        key = object_key(prototype)
        existing = targets_by_key.get(key)
        if existing is not None:
            existing["selected_indexes"].append(index)
            existing["occurrence_count"] += 1
            continue
        target = {
            "component": component,
            "part": prototype,
            "source": "ASSEMBLY_NAVIGATOR_SELECTION",
            "selected_indexes": [index],
            "occurrence_count": 1,
        }
        targets_by_key[key] = target
        targets.append(target)
    if not targets:
        raise RuntimeError(
            "No unique loaded component prototypes were resolved from the selection."
        )
    return targets, count


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


def release_status_snapshot(part):
    pdm_part = safe_property(part, "PDMPart")
    result = {
        "display": "",
        "internal": [],
        "display_raw": "",
        "internal_raw": "",
        "errors": [],
    }
    display_method = getattr(pdm_part, "GetReleaseStatus", None)
    if not callable(display_method):
        result["errors"].append("PDMPart.GetReleaseStatus unavailable")
    else:
        try:
            raw = display_method()
            result["display_raw"] = repr(raw)[:4000]
            result["display"] = clean(raw)
        except Exception as error:
            result["errors"].append("GetReleaseStatus: " + error_text(error))

    internal_method = getattr(pdm_part, "GetInternalReleaseStatus", None)
    if not callable(internal_method):
        result["errors"].append("PDMPart.GetInternalReleaseStatus unavailable")
    else:
        try:
            raw = internal_method([part])
            result["internal_raw"] = repr(raw)[:4000]
            values = [raw] if isinstance(raw, str) else list(raw)
            result["internal"] = [clean(value) for value in values if clean(value)]
        except Exception as error:
            result["errors"].append("GetInternalReleaseStatus: " + error_text(error))
    return result


def release_status_values(snapshot):
    status = snapshot.get("release_status") or {}
    values = [clean(status.get("display"))] + list(status.get("internal") or [])
    return [clean(value) for value in values if clean(value)]


def is_frozen_status(snapshot):
    for value in release_status_values(snapshot):
        normalized = value.upper().replace(" ", "_").replace("-", "_")
        if (
            "UNFREEZ" not in normalized
            and "UNFROZ" not in normalized
            and ("FREEZ" in normalized or "FROZ" in normalized)
        ):
            return True
    return False


def has_other_release_status(snapshot):
    for value in release_status_values(snapshot):
        normalized = value.upper().replace(" ", "_").replace("-", "_")
        if "RELEAS" in normalized:
            return True
        if "FREEZ" not in normalized and "FROZ" not in normalized:
            return True
    return False


def read_only_value(part):
    value = safe_property(part, "IsReadOnly")
    return None if value is None else bool(value)


def modifiability_snapshot(part):
    result = {
        "has_write_access": None,
        "pdm_modifiable": None,
        "errors": [],
    }
    access = safe_property(part, "HasWriteAccess")
    if access is None:
        result["errors"].append("Part.HasWriteAccess unavailable")
    else:
        result["has_write_access"] = bool(access)
    pdm_part = safe_property(part, "PDMPart")
    method = getattr(pdm_part, "IsModifiable", None)
    if not callable(method):
        result["errors"].append("PDMPart.IsModifiable unavailable")
    else:
        try:
            result["pdm_modifiable"] = bool(method())
        except Exception as error:
            result["errors"].append("PDMPart.IsModifiable: " + error_text(error))
    return result


def target_snapshot(session, component, part):
    wae = read_wae_attribute(part)
    version = parse_wae_version(wae["value"])
    part_number = read_identity(part, DB_PART_NO_TITLE)
    revision = read_identity(part, DB_PART_REV_TITLE)
    if not managed_mode(session, part):
        raise RuntimeError("The target CAD part is not positively Teamcenter-managed.")
    if not part_number:
        raise RuntimeError("DB_PART_NO is blank or unavailable.")
    if not revision:
        raise RuntimeError("DB_PART_REV is blank or unavailable.")
    return {
        "component_name": clean(safe_property(component, "DisplayName"))
        or clean(safe_property(component, "Name"))
        or ("<active work part>" if component is None else ""),
        "component_tag": clean(safe_property(component, "Tag")),
        "part_identifier": part_identifier(part),
        "part_number": part_number,
        "db_part_rev": revision,
        "wae_version": version,
        "wae_version_raw": wae["value"],
        "wae_attribute": wae,
        "checkout": checkout_snapshot(session, part),
        "release_status": release_status_snapshot(part),
        "modifiability": modifiability_snapshot(part),
        "read_only": read_only_value(part),
        "part_modified": bool(safe_property(part, "IsModified", False)),
    }


def validate_owned_checkout(snapshot):
    checkout = snapshot["checkout"]
    if checkout["state"] != "CHECKED_OUT":
        return "Target CAD part is not checked out."
    if not checkout["owner"]:
        return "Checkout owner is unavailable."
    if not checkout["current_user"]:
        return "Current Teamcenter user is unavailable."
    if checkout["owner_is_current_user"] is not True:
        return "Target CAD part is checked out by another user: {0}.".format(
            checkout["owner"]
        )
    if snapshot["read_only"] is not False:
        return "Target CAD part is not positively writable after checkout."
    modifiability = snapshot.get("modifiability") or {}
    if modifiability.get("errors"):
        return "Target modifiability query failed: {0}.".format(
            " | ".join(modifiability["errors"])
        )
    if modifiability.get("has_write_access") is not True:
        return "Target does not positively have write access after checkout."
    if modifiability.get("pdm_modifiable") is not True:
        return "PDMPart does not report modifiable after checkout."
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


def checkout_parts(parts):
    if not parts:
        return ""
    pdm_part = safe_property(parts[0], "PDMPart")
    method = getattr(pdm_part, "CheckoutParts", None)
    if not callable(method):
        raise RuntimeError("PDMPart.CheckoutParts is unavailable.")
    checkout_input = NXOpen.PDM.PdmPart.CheckoutInput(
        "J31 WAE unfreeze", "", True, True, False
    )
    errors = None
    try:
        errors = method(list(parts), checkout_input)
        return repr(errors)[:2000]
    finally:
        dispose(errors)


def checkout_part(part):
    return checkout_parts([part])


def checkin_parts(parts):
    if not parts:
        return ""
    pdm_part = safe_property(parts[0], "PDMPart")
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
        errors = method(list(parts), checkin_input)
        return repr(errors)[:2000]
    finally:
        dispose(errors)


def checkin_part(part):
    return checkin_parts([part])


def get_available_workflows(session, parts):
    pdm_session = safe_property(session, "PdmSession")
    method = getattr(pdm_session, "GetNXWorkflows", None)
    if not callable(method):
        raise RuntimeError("PdmSession.GetNXWorkflows is unavailable.")
    errors = None
    try:
        raw = method(list(parts))
        if not isinstance(raw, (tuple, list)) or len(raw) < 2:
            raise RuntimeError(
                "GetNXWorkflows returned an unexpected value: {0}".format(
                    repr(raw)[:2000]
                )
            )
        errors = raw[0]
        names = [raw[1]] if isinstance(raw[1], str) else list(raw[1])
        return {
            "names": [clean(name) for name in names if clean(name)],
            "operation_errors": repr(errors)[:2000],
            "raw": repr(raw)[:4000],
        }
    finally:
        dispose(errors)


def assign_status_workflow(session, parts, action):
    if not parts:
        return ""
    pdm_session = safe_property(session, "PdmSession")
    if action == "FREEZE":
        method_name = "AssignFreezeStatus"
        workflow = FREEZE_WORKFLOW
    elif action == "UNFREEZE":
        method_name = "AssignUnfreezeStatus"
        workflow = UNFREEZE_WORKFLOW
    else:
        raise RuntimeError("Status action must be FREEZE or UNFREEZE.")
    method = getattr(pdm_session, method_name, None)
    if not callable(method):
        raise RuntimeError("PdmSession.{0} is unavailable.".format(method_name))
    errors = None
    try:
        errors = method(list(parts), workflow)
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
        "action": action,
        "mode": mode,
        "scope": "ONE_RESOLVED_CAD_PART",
        "result": "PREFLIGHT_PENDING",
        "message": "",
        "target_index": 0,
        "source": "",
        "selected_indexes": [],
        "selected_occurrence_count": 1,
        "planned_action": "",
        "before": {},
        "after": {},
        "operations": {
            "checkout_attempted": False,
            "wae_write_attempted": False,
            "save_attempted": False,
            "checkin_attempted": False,
            "freeze_status_attempted": False,
            "unfreeze_status_attempted": False,
            "formal_revision_created": False,
        },
        "operation_raw": {},
    }


def batch_report(action, build, mode):
    return {
        "build": build,
        "timestamp": datetime.datetime.now().astimezone().isoformat(timespec="seconds"),
        "action": action,
        "mode": mode,
        "scope": "SELECTED_COMPONENT_PROTOTYPES_OR_ACTIVE_WORK_PART",
        "targeting_rule": (
            "Use all preselected Assembly Navigator component prototypes; "
            "otherwise use only the active work part."
        ),
        "target_source": "",
        "selected_object_count": 0,
        "unique_target_count": 0,
        "duplicate_occurrences_collapsed": 0,
        "result": "BLOCKED",
        "message": "",
        "required_workflow": "",
        "workflow_query": {},
        "preflight": {
            "passed": False,
            "errors": [],
        },
        "failed_stage": "",
        "operations": {
            "save_attempted": False,
            "checkin_attempted": False,
            "freeze_status_attempted": False,
            "unfreeze_status_attempted": False,
            "checkout_attempted": False,
            "wae_write_attempted": False,
            "formal_revision_created": False,
        },
        "operation_raw": {},
        "counts": {
            "succeeded": 0,
            "blocked": 0,
            "failed": 0,
        },
        "targets": [],
    }


def make_target_report(action, session, target, build, mode, index):
    report = base_report(action, build, mode)
    report["target_index"] = index + 1
    report["source"] = target["source"]
    report["selected_indexes"] = [value + 1 for value in target["selected_indexes"]]
    report["selected_occurrence_count"] = target["occurrence_count"]
    try:
        before = target_snapshot(session, target["component"], target["part"])
        report["before"] = before
        status_errors = (before.get("release_status") or {}).get("errors") or []
        if status_errors:
            raise RuntimeError("Release-status query failed: " + " | ".join(status_errors))
        if action == "FREEZE":
            if has_other_release_status(before):
                raise RuntimeError(
                    "Target has a non-freeze release status: {0}.".format(
                        ", ".join(release_status_values(before))
                    )
                )
            if is_frozen_status(before):
                modifiability = before.get("modifiability") or {}
                if (
                    before["checkout"]["state"] != "CHECKED_IN"
                    or before["read_only"] is not True
                    or modifiability.get("errors")
                    or modifiability.get("has_write_access") is not False
                    or modifiability.get("pdm_modifiable") is not False
                ):
                    raise RuntimeError(
                        "Frozen status is inconsistent with checkout/read-only/access state."
                    )
                report["planned_action"] = "ALREADY_FROZEN"
            elif before["checkout"]["state"] == "CHECKED_IN":
                report["planned_action"] = "ASSIGN_FREEZE_STATUS"
            else:
                owned_error = validate_owned_checkout(before)
                if owned_error:
                    raise RuntimeError(owned_error)
                report["planned_action"] = "SAVE_CHECKIN_AND_ASSIGN_FREEZE_STATUS"
        else:
            if before["checkout"]["state"] != "CHECKED_IN":
                raise RuntimeError(
                    "UNFREEZE requires a checked-in frozen target; reruns while checked out are blocked."
                )
            if has_other_release_status(before):
                raise RuntimeError(
                    "Target has a non-freeze release status: {0}.".format(
                        ", ".join(release_status_values(before))
                    )
                )
            if not is_frozen_status(before):
                raise RuntimeError("Target does not have a positive freeze status.")
            wae = before["wae_attribute"]
            if wae["type"] != "STRING" or wae["owned_by_system"] or wae["pdm_based"]:
                raise RuntimeError("WAE_VERSION is not an operator-writable string attribute.")
            report["planned_action"] = "UNFREEZE_CHECKOUT_INCREMENT_AND_SAVE"
            report["next_wae_version"] = before["wae_version"] + 1
        report["result"] = "PREFLIGHT_READY"
        report["message"] = "Target preflight passed."
    except Exception as error:
        report["result"] = "PREFLIGHT_BLOCKED"
        report["message"] = error_text(error)
    return report


def verify_identity_unchanged(before, after, stage):
    if after["db_part_rev"] != before["db_part_rev"]:
        raise RuntimeError("DB_PART_REV changed during {0}.".format(stage))
    if after["wae_version"] != before["wae_version"]:
        raise RuntimeError("WAE_VERSION changed during {0}.".format(stage))


def capture_after_states(session, targets, reports):
    for target, target_report in zip(targets, reports):
        try:
            target_report["after"] = target_snapshot(
                session, target["component"], target["part"]
            )
        except Exception as error:
            target_report["after_snapshot_error"] = error_text(error)


def mark_recovery_results(action, reports):
    completed_results = ("FROZEN", "UNFROZEN_READY_FOR_EDIT")
    for report in reports:
        if report["result"] in completed_results:
            continue
        before = report.get("before") or {}
        after = report.get("after") or {}
        completed = False
        if before and after:
            if action == "FREEZE":
                completed = (
                    after.get("db_part_rev") == before.get("db_part_rev")
                    and after.get("wae_version") == before.get("wae_version")
                    and (after.get("checkout") or {}).get("state") == "CHECKED_IN"
                    and after.get("read_only") is True
                    and (after.get("modifiability") or {}).get("has_write_access") is False
                    and (after.get("modifiability") or {}).get("pdm_modifiable") is False
                    and is_frozen_status(after)
                )
                if completed:
                    report["result"] = "FROZEN"
                    report["message"] = "Completed before the later batch failure."
            else:
                completed = (
                    after.get("db_part_rev") == before.get("db_part_rev")
                    and after.get("wae_version") == report.get("next_wae_version")
                    and (after.get("checkout") or {}).get("state") == "CHECKED_OUT"
                    and (after.get("checkout") or {}).get("owner_is_current_user") is True
                    and after.get("read_only") is False
                    and (after.get("modifiability") or {}).get("has_write_access") is True
                    and (after.get("modifiability") or {}).get("pdm_modifiable") is True
                    and not is_frozen_status(after)
                    and not has_other_release_status(after)
                )
                if completed:
                    report["result"] = "UNFROZEN_READY_FOR_EDIT"
                    report["message"] = "Completed before the later batch failure."
        if not completed:
            report["result"] = "RECOVERY_REQUIRED"
            report["message"] = "Inspect before/after state and recover manually."


def execute_freeze_batch(session, targets, reports, batch):
    to_checkin = [
        (target, report)
        for target, report in zip(targets, reports)
        if report["planned_action"] == "SAVE_CHECKIN_AND_ASSIGN_FREEZE_STATUS"
    ]
    to_freeze = [
        (target, report)
        for target, report in zip(targets, reports)
        if report["planned_action"] != "ALREADY_FROZEN"
    ]

    for target, report in to_checkin:
        batch["failed_stage"] = "SAVE_BEFORE_CHECKIN"
        batch["operations"]["save_attempted"] = True
        report["operations"]["save_attempted"] = True
        save_part(target["part"])
        saved = target_snapshot(session, target["component"], target["part"])
        verify_identity_unchanged(report["before"], saved, "freeze save")

    if to_checkin:
        batch["failed_stage"] = "BATCH_CHECKIN"
        batch["operations"]["checkin_attempted"] = True
        for _, report in to_checkin:
            report["operations"]["checkin_attempted"] = True
        batch["operation_raw"]["checkin"] = checkin_parts(
            [target["part"] for target, _ in to_checkin]
        )
        for target, report in to_checkin:
            checked_in = target_snapshot(session, target["component"], target["part"])
            verify_identity_unchanged(report["before"], checked_in, "check-in")
            if checked_in["checkout"]["state"] != "CHECKED_IN":
                raise RuntimeError(
                    "Check-in postcondition failed for {0}.".format(
                        checked_in["part_number"]
                    )
                )

    if to_freeze:
        batch["failed_stage"] = "ASSIGN_FREEZE_STATUS"
        batch["operations"]["freeze_status_attempted"] = True
        for _, report in to_freeze:
            report["operations"]["freeze_status_attempted"] = True
        batch["operation_raw"]["freeze_status"] = assign_status_workflow(
            session, [target["part"] for target, _ in to_freeze], "FREEZE"
        )

    batch["failed_stage"] = "VERIFY_FROZEN_POSTCONDITIONS"
    for target, report in zip(targets, reports):
        after = target_snapshot(session, target["component"], target["part"])
        report["after"] = after
        verify_identity_unchanged(report["before"], after, "freeze")
        if after["checkout"]["state"] != "CHECKED_IN":
            raise RuntimeError("Frozen target is not checked in: " + after["part_number"])
        if after["read_only"] is not True:
            raise RuntimeError("Frozen target is not read-only: " + after["part_number"])
        modifiability = after.get("modifiability") or {}
        if modifiability.get("errors"):
            raise RuntimeError(
                "Frozen modifiability query failed for {0}: {1}.".format(
                    after["part_number"], " | ".join(modifiability["errors"])
                )
            )
        if modifiability.get("has_write_access") is not False:
            raise RuntimeError("Frozen target still has write access: " + after["part_number"])
        if modifiability.get("pdm_modifiable") is not False:
            raise RuntimeError("Frozen target remains PDM-modifiable: " + after["part_number"])
        if not is_frozen_status(after):
            raise RuntimeError(
                "Freeze status was not observed for {0}; status={1}.".format(
                    after["part_number"], release_status_values(after)
                )
            )
        report["result"] = "FROZEN"
        report["message"] = "Freeze status, check-in, read-only state, revision, and WAE verified."


def execute_unfreeze_batch(session, targets, reports, batch):
    parts = [target["part"] for target in targets]
    batch["failed_stage"] = "ASSIGN_UNFREEZE_STATUS"
    batch["operations"]["unfreeze_status_attempted"] = True
    for report in reports:
        report["operations"]["unfreeze_status_attempted"] = True
    batch["operation_raw"]["unfreeze_status"] = assign_status_workflow(
        session, parts, "UNFREEZE"
    )

    batch["failed_stage"] = "VERIFY_UNFREEZE_STATUS"
    for target, report in zip(targets, reports):
        unfrozen = target_snapshot(session, target["component"], target["part"])
        verify_identity_unchanged(report["before"], unfrozen, "unfreeze status assignment")
        if is_frozen_status(unfrozen) or has_other_release_status(unfrozen):
            raise RuntimeError(
                "Unfreeze status postcondition failed for {0}; status={1}.".format(
                    unfrozen["part_number"], release_status_values(unfrozen)
                )
            )
        if unfrozen["checkout"]["state"] != "CHECKED_IN":
            raise RuntimeError("Unfreeze unexpectedly changed checkout state.")

    batch["failed_stage"] = "BATCH_CHECKOUT"
    batch["operations"]["checkout_attempted"] = True
    for report in reports:
        report["operations"]["checkout_attempted"] = True
    batch["operation_raw"]["checkout"] = checkout_parts(parts)

    batch["failed_stage"] = "VERIFY_CHECKOUT"
    for target, report in zip(targets, reports):
        checked_out = target_snapshot(session, target["component"], target["part"])
        verify_identity_unchanged(report["before"], checked_out, "checkout")
        owned_error = validate_owned_checkout(checked_out)
        if owned_error:
            raise RuntimeError(
                "Checkout postcondition failed for {0}: {1}".format(
                    checked_out["part_number"], owned_error
                )
            )

    for target, report in zip(targets, reports):
        mark = None
        mark_name = "J31 WAE_VERSION increment"
        try:
            batch["failed_stage"] = "WAE_INCREMENT_TARGET_{0}".format(
                report["target_index"]
            )
            mark = session.SetUndoMark(
                NXOpen.Session.MarkVisibility.Invisible, mark_name
            )
            batch["operations"]["wae_write_attempted"] = True
            report["operations"]["wae_write_attempted"] = True
            next_version = report["next_wae_version"]
            write_wae_version(session, target["part"], next_version)
            immediate = read_wae_attribute(target["part"])
            if parse_wae_version(immediate["value"]) != next_version:
                raise RuntimeError("Immediate WAE_VERSION reread did not match increment.")
            batch["operations"]["save_attempted"] = True
            report["operations"]["save_attempted"] = True
            save_part(target["part"])
            after = target_snapshot(session, target["component"], target["part"])
            report["after"] = after
            owned_error = validate_owned_checkout(after)
            if owned_error:
                raise RuntimeError("Post-save checkout state failed: " + owned_error)
            if after["db_part_rev"] != report["before"]["db_part_rev"]:
                raise RuntimeError("DB_PART_REV changed during unfreeze.")
            if after["wae_version"] != next_version:
                raise RuntimeError("Saved WAE_VERSION does not match controlled increment.")
            if is_frozen_status(after) or has_other_release_status(after):
                raise RuntimeError("Release status returned after controlled increment.")
            report["result"] = "UNFROZEN_READY_FOR_EDIT"
            report["message"] = "Unfrozen, checked out, and advanced to WAE_VERSION {0}.".format(
                next_version
            )
        finally:
            if mark is not None:
                try:
                    session.DeleteUndoMark(mark, mark_name)
                except Exception:
                    pass


def execute(action, session, selection_manager, build, mode):
    action = clean(action).upper()
    mode = clean(mode).upper()
    if action not in VALID_ACTIONS:
        raise RuntimeError("Action must be FREEZE or UNFREEZE.")
    if mode not in VALID_MODES:
        raise RuntimeError("Mode must be DRY_RUN or APPLY.")
    report = batch_report(action, build, mode)
    try:
        targets, selected_count = selected_or_work_targets(session, selection_manager)
    except Exception as error:
        report["message"] = error_text(error)
        return report

    report["selected_object_count"] = selected_count
    report["unique_target_count"] = len(targets)
    report["duplicate_occurrences_collapsed"] = max(0, selected_count - len(targets))
    report["target_source"] = targets[0]["source"]
    required_workflow = FREEZE_WORKFLOW if action == "FREEZE" else UNFREEZE_WORKFLOW
    report["required_workflow"] = required_workflow
    try:
        report["workflow_query"] = get_available_workflows(
            session, [target["part"] for target in targets]
        )
        if required_workflow not in report["workflow_query"]["names"]:
            raise RuntimeError(
                "Required workflow {0!r} is unavailable; found {1}.".format(
                    required_workflow, report["workflow_query"]["names"]
                )
            )
    except Exception as error:
        report["preflight"]["errors"].append(error_text(error))

    for index, target in enumerate(targets):
        target_report = make_target_report(action, session, target, build, mode, index)
        report["targets"].append(target_report)
        if target_report["result"] == "PREFLIGHT_BLOCKED":
            report["preflight"]["errors"].append(
                "Target {0}: {1}".format(index + 1, target_report["message"])
            )

    if report["preflight"]["errors"]:
        report["result"] = "BLOCKED_BATCH"
        report["message"] = "Complete batch preflight failed; nothing changed."
        report["counts"]["blocked"] = len(targets)
        return report
    report["preflight"]["passed"] = True

    if mode == "DRY_RUN":
        report["result"] = "DRY_RUN_READY"
        report["message"] = "Complete batch preflight passed; no mutations attempted."
        for target_report in report["targets"]:
            target_report["result"] = "DRY_RUN_READY"
        report["counts"]["succeeded"] = len(targets)
        return report

    try:
        if action == "FREEZE":
            execute_freeze_batch(session, targets, report["targets"], report)
            report["result"] = "ALL_TARGETS_FROZEN"
        else:
            execute_unfreeze_batch(session, targets, report["targets"], report)
            report["result"] = "ALL_TARGETS_UNFROZEN"
        report["failed_stage"] = ""
        report["counts"]["succeeded"] = len(targets)
        report["message"] = "All {0} unique target(s) completed and verified.".format(
            len(targets)
        )
    except Exception as error:
        capture_after_states(session, targets, report["targets"])
        mark_recovery_results(action, report["targets"])
        report["result"] = "RECOVERY_REQUIRED"
        report["message"] = error_text(error)
        report["counts"]["succeeded"] = sum(
            1
            for target_report in report["targets"]
            if target_report["result"] in ("FROZEN", "UNFROZEN_READY_FOR_EDIT")
        )
        report["counts"]["failed"] = len(targets) - report["counts"]["succeeded"]
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
    log_line(
        session,
        "Target: all preselected Assembly Navigator components, or active work part when none are selected.",
    )
    try:
        selection_manager = NXOpen.UI.GetUI().SelectionManager
        report = execute(action, session, selection_manager, build, mode)
    except Exception as error:
        report = batch_report(action, build, mode)
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
    if report.get("failed_stage"):
        log_line(session, "Failed stage: " + report["failed_stage"])
    for target in report.get("targets") or []:
        before = target.get("before") or {}
        after = target.get("after") or {}
        identity = before or after
        log_line(
            session,
            "[{0}] {1}/{2} | {3} | {4}".format(
                target.get("target_index", ""),
                identity.get("part_number", ""),
                identity.get("db_part_rev", ""),
                target.get("result", ""),
                target.get("message", ""),
            ),
        )
        if before:
            log_line(
                session,
                "  Before: WAE {0} {1} status={2}".format(
                    before.get("wae_version_raw", ""),
                    (before.get("checkout") or {}).get("state", ""),
                    release_status_values(before),
                ),
            )
        if after:
            log_line(
                session,
                "  After:  WAE {0} {1} status={2}".format(
                    after.get("wae_version_raw", ""),
                    (after.get("checkout") or {}).get("state", ""),
                    release_status_values(after),
                ),
            )
    if path:
        log_line(session, "Audit JSON: " + path)
    log_line(session, "=" * 72)
    return report
