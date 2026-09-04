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
COMMON_BUILD = "WAE-CHANGE-CONTROL-V5"
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


def runtime_type_name(value):
    value_type = type(value)
    module_name = clean(getattr(value_type, "__module__", ""))
    type_name = clean(getattr(value_type, "__name__", "")) or clean(value_type)
    return "{0}.{1}".format(module_name, type_name) if module_name else type_name


def selection_diagnostic(selected, index):
    owner = safe_property(selected, "OwningComponent")
    prototype = safe_property(selected, "Prototype")
    return {
        "index": index + 1,
        "runtime_type": runtime_type_name(selected),
        "tag": clean(safe_property(selected, "Tag")),
        "name": clean(safe_property(selected, "DisplayName"))
        or clean(safe_property(selected, "Name")),
        "has_prototype": prototype is not None,
        "has_pdm_part": safe_property(selected, "PDMPart") is not None,
        "owning_component_tag": clean(safe_property(owner, "Tag")),
        "resolution": "UNRESOLVED",
    }


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


def active_work_target(session, selected_indexes=None):
    parts = safe_property(session, "Parts")
    work_part = safe_property(parts, "Work")
    if work_part is None:
        raise RuntimeError(
            "No component target was resolved and there is no active work part."
        )
    if safe_property(work_part, "PDMPart") is None:
        raise RuntimeError("The active work part has no PDMPart; open a managed CAD part.")
    return {
        "component": None,
        "part": work_part,
        "source": "ACTIVE_WORK_PART",
        "selected_indexes": list(selected_indexes or []),
        "occurrence_count": 1,
    }


def selected_or_work_targets(
    session, selection_manager, report=None, allow_partial_selection=False
):
    """Resolve unique selected CAD parts, or fall back to the active work part."""
    try:
        count = int(selection_manager.GetNumSelectedObjects())
    except Exception as error:
        raise RuntimeError("Could not inspect NX preselection: " + error_text(error))

    diagnostics = []
    if report is not None:
        report["selected_object_count"] = count
        report["selected_objects"] = diagnostics

    if count == 0:
        return [active_work_target(session)], count

    parts = safe_property(session, "Parts")
    work_part = safe_property(parts, "Work")
    targets = []
    targets_by_key = {}
    unresolved_indexes = []
    active_part_indexes = []
    for index in range(count):
        try:
            selected = selection_manager.GetSelectedTaggedObject(index)
        except Exception as error:
            raise RuntimeError(
                "Could not read selected NX object {0}: {1}".format(
                    index + 1, error_text(error)
                )
            )
        diagnostic = selection_diagnostic(selected, index)
        diagnostics.append(diagnostic)

        component = None
        part = None
        source = ""
        prototype_marker = object()
        direct_prototype = safe_property(selected, "Prototype", prototype_marker)
        if direct_prototype is not prototype_marker:
            component = selected
            part = component_target(component, index)
            source = "ASSEMBLY_NAVIGATOR_SELECTION"
            diagnostic["resolution"] = "COMPONENT_PROTOTYPE"
        else:
            owner = safe_property(selected, "OwningComponent")
            if owner is not None:
                component = owner
                part = component_target(component, index)
                source = "OWNING_COMPONENT_SELECTION"
                diagnostic["resolution"] = "OWNING_COMPONENT_PROTOTYPE"
            elif safe_property(selected, "PDMPart") is not None:
                part = selected
                source = "MANAGED_PART_SELECTION"
                diagnostic["resolution"] = "MANAGED_PART"
            else:
                unresolved_indexes.append(index)
                continue

        if work_part is not None and same_nx_object(part, work_part):
            active_part_indexes.append(index)
            diagnostic["resolution"] = "ACTIVE_WORK_PART_FALLBACK_CANDIDATE"
            continue

        key = object_key(part)
        existing = targets_by_key.get(key)
        if existing is not None:
            existing["selected_indexes"].append(index)
            existing["occurrence_count"] += 1
            if existing["component"] is None and component is not None:
                existing["component"] = component
            continue
        target = {
            "component": component,
            "part": part,
            "source": source,
            "selected_indexes": [index],
            "occurrence_count": 1,
        }
        targets_by_key[key] = target
        targets.append(target)

    if unresolved_indexes and targets and not allow_partial_selection:
        unresolved = ", ".join(str(index + 1) for index in unresolved_indexes)
        raise RuntimeError(
            "Selected object(s) {0} did not resolve to managed CAD targets; "
            "the complete batch was blocked.".format(unresolved)
        )
    if unresolved_indexes and targets and report is not None:
        unresolved = ", ".join(str(index + 1) for index in unresolved_indexes)
        report.setdefault("selection_warnings", []).append(
            "Selected object(s) {0} did not resolve to managed CAD targets and "
            "were skipped by J30.".format(unresolved)
        )
    if not targets:
        fallback_indexes = active_part_indexes + unresolved_indexes
        for diagnostic in diagnostics:
            if diagnostic["resolution"] == "UNRESOLVED":
                diagnostic["resolution"] = "IGNORED_FOR_ACTIVE_WORK_PART_FALLBACK"
        return [active_work_target(session, fallback_indexes)], count
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


def classify_wae_version(value, revision):
    """Classify the shared J30/J31 lifecycle value without changing it."""
    raw = clean(value)
    db_revision = clean(revision)
    if not raw:
        return "", "WAE_VERSION is blank."
    if re.fullmatch(r"[1-9][0-9]*", raw):
        return "NUMERIC_WORKING", ""
    if re.fullmatch(r"[A-Za-z]+", raw):
        if raw.casefold() == db_revision.casefold():
            return "ALPHABETIC_FINAL", ""
        return "", (
            "Alphabetic WAE_VERSION {0!r} does not match DB_PART_REV {1!r}.".format(
                raw, db_revision
            )
        )
    return "", (
        "WAE_VERSION is neither a positive whole number nor a matching "
        "alphabetic revision."
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
    part_number = read_identity(part, DB_PART_NO_TITLE)
    revision = read_identity(part, DB_PART_REV_TITLE)
    if not managed_mode(session, part):
        raise RuntimeError("The target CAD part is not positively Teamcenter-managed.")
    if not part_number:
        raise RuntimeError("DB_PART_NO is blank or unavailable.")
    if not revision:
        raise RuntimeError("DB_PART_REV is blank or unavailable.")
    wae_class, wae_error = classify_wae_version(wae["value"], revision)
    version = (
        int(clean(wae["value"]))
        if wae_class == "NUMERIC_WORKING"
        else clean(wae["value"])
    )
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
        "wae_class": wae_class,
        "wae_validation_error": wae_error,
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
        "block_reason": "",
        "message": "",
        "target_index": 0,
        "source": "",
        "selected_indexes": [],
        "selected_occurrence_count": 1,
        "planned_action": "",
        "failed_stage": "",
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
        "helper_build": COMMON_BUILD,
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
        "selected_objects": [],
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
        wae_error = before.get("wae_validation_error", "")
        if wae_error:
            report["result"] = (
                "BLOCKED_MISSING_WAE_VERSION"
                if not clean(before.get("wae_version_raw"))
                else "BLOCKED_INVALID_WAE_VERSION"
            )
            report["block_reason"] = report["result"]
            report["message"] = wae_error
            return report
        if action == "UNFREEZE" and before.get("wae_class") == "ALPHABETIC_FINAL":
            report["result"] = "BLOCKED_FINAL_RELEASE_BASELINE"
            report["block_reason"] = report["result"]
            report["message"] = (
                "WAE_VERSION {0!r} matches TCX revision {1!r} and is an immutable "
                "final-release baseline. Create the next formal TCX revision instead."
            ).format(before["wae_version_raw"], before["db_part_rev"])
            return report
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
        message = error_text(error)
        if "WAE_VERSION" in message and ("blank" in message or "found 0" in message):
            report["result"] = "BLOCKED_MISSING_WAE_VERSION"
        else:
            report["result"] = "PREFLIGHT_BLOCKED"
        report["block_reason"] = report["result"]
        report["message"] = message
    return report


def verify_identity_unchanged(before, after, stage):
    if after["db_part_rev"] != before["db_part_rev"]:
        raise RuntimeError("DB_PART_REV changed during {0}.".format(stage))
    if clean(after.get("wae_version_raw")).casefold() != clean(
        before.get("wae_version_raw")
    ).casefold():
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
    completed_results = (
        "ALREADY_FROZEN", "FROZEN", "FROZEN_WITH_WARNING",
        "UNFROZEN_READY_FOR_EDIT", "UNFROZEN_WITH_WARNING",
    )
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
                    and clean(after.get("wae_version_raw")).casefold()
                    == clean(before.get("wae_version_raw")).casefold()
                    and (after.get("checkout") or {}).get("state") == "CHECKED_IN"
                    and after.get("read_only") is True
                    and (after.get("modifiability") or {}).get("pdm_modifiable") is False
                    and is_frozen_status(after)
                )
                if completed:
                    report["result"] = "FROZEN_WITH_WARNING"
                    report["message"] = "Verified Frozen despite a reported operation failure."
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
                    report["result"] = "UNFROZEN_WITH_WARNING"
                    report["message"] = "Verified complete despite a reported operation failure."
        if not completed:
            report["result"] = "RECOVERY_REQUIRED"
            report["message"] = "Inspect before/after state and recover manually."


def frozen_postconditions(before, after):
    return (
        clean(after.get("part_number")).casefold()
        == clean(before.get("part_number")).casefold()
        and clean(after.get("db_part_rev")).casefold()
        == clean(before.get("db_part_rev")).casefold()
        and clean(after.get("wae_version_raw")).casefold()
        == clean(before.get("wae_version_raw")).casefold()
        and (after.get("checkout") or {}).get("state") == "CHECKED_IN"
        and after.get("read_only") is True
        and (after.get("modifiability") or {}).get("pdm_modifiable") is False
        and is_frozen_status(after)
        and not has_other_release_status(after)
    )


def completed_unfreeze_postconditions(report, after):
    before = report.get("before") or {}
    return (
        clean(after.get("part_number")).casefold()
        == clean(before.get("part_number")).casefold()
        and clean(after.get("db_part_rev")).casefold()
        == clean(before.get("db_part_rev")).casefold()
        and after.get("wae_version") == report.get("next_wae_version")
        and (after.get("checkout") or {}).get("state") == "CHECKED_OUT"
        and (after.get("checkout") or {}).get("owner_is_current_user") is True
        and after.get("read_only") is False
        and (after.get("modifiability") or {}).get("has_write_access") is True
        and (after.get("modifiability") or {}).get("pdm_modifiable") is True
        and not is_frozen_status(after)
        and not has_other_release_status(after)
    )


def snapshot_after(session, target, report):
    try:
        after = target_snapshot(session, target["component"], target["part"])
        report["after"] = after
        return after, ""
    except Exception as error:
        message = error_text(error)
        report["after_snapshot_error"] = message
        return {}, message


def execute_freeze_batch(session, targets, reports, batch):
    """Freeze exact identities independently; one failure does not stop others."""
    for target, report in zip(targets, reports):
        if report["planned_action"] == "ALREADY_FROZEN":
            report["after"] = report["before"]
            report["result"] = "ALREADY_FROZEN"
            report["message"] = "Target is already a verified Frozen baseline."
            continue

        operation_error = ""
        try:
            if report["planned_action"] == "SAVE_CHECKIN_AND_ASSIGN_FREEZE_STATUS":
                report["failed_stage"] = "SAVE_BEFORE_CHECKIN"
                batch["operations"]["save_attempted"] = True
                report["operations"]["save_attempted"] = True
                save_part(target["part"])
                saved = target_snapshot(session, target["component"], target["part"])
                verify_identity_unchanged(report["before"], saved, "freeze save")

                report["failed_stage"] = "CHECKIN"
                batch["operations"]["checkin_attempted"] = True
                report["operations"]["checkin_attempted"] = True
                report["operation_raw"]["checkin"] = checkin_part(target["part"])
                checked_in = target_snapshot(
                    session, target["component"], target["part"]
                )
                verify_identity_unchanged(report["before"], checked_in, "check-in")
                if checked_in["checkout"]["state"] != "CHECKED_IN":
                    raise RuntimeError("Check-in postcondition failed.")

            report["failed_stage"] = "ASSIGN_FREEZE_STATUS"
            batch["operations"]["freeze_status_attempted"] = True
            report["operations"]["freeze_status_attempted"] = True
            report["operation_raw"]["freeze_status"] = assign_status_workflow(
                session, [target["part"]], "FREEZE"
            )
        except Exception as error:
            operation_error = error_text(error)
            report["operation_raw"]["warning"] = operation_error

        report["failed_stage"] = "VERIFY_FROZEN_POSTCONDITIONS"
        after, verification_error = snapshot_after(session, target, report)
        if after and frozen_postconditions(report["before"], after):
            report["failed_stage"] = ""
            report["result"] = "FROZEN_WITH_WARNING" if operation_error else "FROZEN"
            report["message"] = (
                "Verified Frozen despite operation warning: " + operation_error
                if operation_error else
                "Freeze status, check-in, read-only state, revision, and WAE verified."
            )
        else:
            report["result"] = "FAILED_FREEZE_WORKFLOW"
            report["message"] = operation_error or verification_error or (
                "Freeze workflow returned without a valid Frozen final state."
            )


def execute_unfreeze_batch(session, targets, reports, batch):
    """Run each target end-to-end and stop only on an incomplete mutation."""
    stop_remaining = False
    for target, report in zip(targets, reports):
        if stop_remaining:
            report["result"] = "NOT_ATTEMPTED_AFTER_RECOVERY_REQUIRED"
            report["message"] = "A prior target requires recovery; this target was not changed."
            continue

        mark = None
        mark_name = "J31 WAE_VERSION increment"
        operation_warning = ""
        try:
            report["failed_stage"] = "ASSIGN_UNFREEZE_STATUS"
            batch["operations"]["unfreeze_status_attempted"] = True
            report["operations"]["unfreeze_status_attempted"] = True
            try:
                report["operation_raw"]["unfreeze_status"] = assign_status_workflow(
                    session, [target["part"]], "UNFREEZE"
                )
            except Exception as error:
                operation_warning = error_text(error)
                report["operation_raw"]["unfreeze_warning"] = operation_warning

            report["failed_stage"] = "VERIFY_UNFREEZE_STATUS"
            unfrozen = target_snapshot(session, target["component"], target["part"])
            report["after"] = unfrozen
            verify_identity_unchanged(
                report["before"], unfrozen, "unfreeze status assignment"
            )
            if is_frozen_status(unfrozen) or has_other_release_status(unfrozen):
                if frozen_postconditions(report["before"], unfrozen):
                    report["result"] = "FAILED_UNFREEZE_WORKFLOW"
                    report["message"] = operation_warning or (
                        "Unfreeze workflow returned but the target remained safely Frozen."
                    )
                    continue
                raise RuntimeError("Unfreeze left an inconsistent controlled status.")
            if unfrozen["checkout"]["state"] != "CHECKED_IN":
                raise RuntimeError("Unfreeze unexpectedly changed checkout state.")

            report["failed_stage"] = "CHECKOUT"
            batch["operations"]["checkout_attempted"] = True
            report["operations"]["checkout_attempted"] = True
            report["operation_raw"]["checkout"] = checkout_part(target["part"])

            report["failed_stage"] = "VERIFY_CHECKOUT"
            checked_out = target_snapshot(session, target["component"], target["part"])
            report["after"] = checked_out
            verify_identity_unchanged(report["before"], checked_out, "checkout")
            owned_error = validate_owned_checkout(checked_out)
            if owned_error:
                raise RuntimeError("Checkout postcondition failed: " + owned_error)

            report["failed_stage"] = "WAE_INCREMENT"
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
            if not completed_unfreeze_postconditions(report, after):
                raise RuntimeError("Final unfreeze/WAE postconditions were not satisfied.")
            report["failed_stage"] = ""
            report["result"] = (
                "UNFROZEN_WITH_WARNING" if operation_warning
                else "UNFROZEN_READY_FOR_EDIT"
            )
            report["message"] = (
                "Unfrozen and advanced to WAE_VERSION {0} despite workflow warning: {1}"
                .format(next_version, operation_warning)
                if operation_warning else
                "Unfrozen, checked out, and advanced to WAE_VERSION {0}.".format(
                    next_version
                )
            )
        except Exception as error:
            failure = error_text(error)
            after, verification_error = snapshot_after(session, target, report)
            if after and completed_unfreeze_postconditions(report, after):
                report["failed_stage"] = ""
                report["result"] = "UNFROZEN_WITH_WARNING"
                report["message"] = "Verified complete despite operation warning: " + failure
            elif after and frozen_postconditions(report["before"], after):
                report["result"] = "FAILED_UNFREEZE_WORKFLOW"
                report["message"] = failure
            else:
                report["result"] = "RECOVERY_REQUIRED"
                report["message"] = failure or verification_error
                batch["failed_stage"] = report["failed_stage"]
                stop_remaining = True
        finally:
            if mark is not None:
                try:
                    session.DeleteUndoMark(mark, mark_name)
                except Exception:
                    pass
    return stop_remaining


def collapse_exact_identity_targets(targets, reports):
    """Collapse loaded NX proxies by authoritative Teamcenter part/revision identity."""
    collapsed_targets = []
    collapsed_reports = []
    by_identity = {}
    for target, report in zip(targets, reports):
        before = report.get("before") or {}
        number = clean(before.get("part_number"))
        revision = clean(before.get("db_part_rev"))
        key = (number.casefold(), revision.casefold()) if number and revision else None
        if key is None or key not in by_identity:
            if key is not None:
                by_identity[key] = len(collapsed_targets)
            collapsed_targets.append(target)
            collapsed_reports.append(report)
            continue

        existing_index = by_identity[key]
        existing_target = collapsed_targets[existing_index]
        existing_report = collapsed_reports[existing_index]
        existing_target["selected_indexes"].extend(target["selected_indexes"])
        existing_target["occurrence_count"] += target["occurrence_count"]
        existing_report["selected_indexes"] = sorted(set(
            existing_report["selected_indexes"] + report["selected_indexes"]
        ))
        existing_report["selected_occurrence_count"] += report[
            "selected_occurrence_count"
        ]
        existing_before = existing_report.get("before") or {}
        if (
            clean(existing_before.get("wae_version_raw")).casefold()
            != clean(before.get("wae_version_raw")).casefold()
        ):
            existing_report["result"] = "PREFLIGHT_BLOCKED"
            existing_report["block_reason"] = "BLOCKED_CONFLICTING_LOADED_STATE"
            existing_report["message"] = (
                "Loaded NX objects for the same DB_PART_NO + DB_PART_REV expose "
                "different WAE_VERSION values. Reload the assembly before retrying."
            )

    for index, report in enumerate(collapsed_reports):
        report["target_index"] = index + 1
    return collapsed_targets, collapsed_reports


def preflight_workflow(session, target, target_report, required_workflow):
    if target_report["result"] != "PREFLIGHT_READY":
        return
    if target_report.get("planned_action") == "ALREADY_FROZEN":
        return
    try:
        query = get_available_workflows(session, [target["part"]])
        target_report["workflow_query"] = query
        if required_workflow not in query["names"]:
            raise RuntimeError(
                "Required workflow {0!r} is unavailable; found {1}.".format(
                    required_workflow, query["names"]
                )
            )
    except Exception as error:
        target_report["result"] = "PREFLIGHT_BLOCKED"
        target_report["block_reason"] = "BLOCKED_WORKFLOW_UNAVAILABLE"
        target_report["message"] = error_text(error)


def result_counts(target_reports):
    success_results = {
        "ALREADY_FROZEN", "FROZEN", "FROZEN_WITH_WARNING",
        "UNFROZEN_READY_FOR_EDIT", "UNFROZEN_WITH_WARNING",
        "DRY_RUN_READY",
    }
    blocked_results = {
        "PREFLIGHT_BLOCKED", "BLOCKED_MISSING_WAE_VERSION",
        "BLOCKED_INVALID_WAE_VERSION", "BLOCKED_FINAL_RELEASE_BASELINE",
        "NOT_ATTEMPTED_BATCH_BLOCKED", "NOT_ATTEMPTED_AFTER_RECOVERY_REQUIRED",
    }
    return {
        "succeeded": sum(1 for row in target_reports if row["result"] in success_results),
        "blocked": sum(1 for row in target_reports if row["result"] in blocked_results),
        "failed": sum(
            1 for row in target_reports
            if row["result"] not in success_results and row["result"] not in blocked_results
        ),
    }


def execute(action, session, selection_manager, build, mode):
    action = clean(action).upper()
    mode = clean(mode).upper()
    if action not in VALID_ACTIONS:
        raise RuntimeError("Action must be FREEZE or UNFREEZE.")
    if mode not in VALID_MODES:
        raise RuntimeError("Mode must be DRY_RUN or APPLY.")
    report = batch_report(action, build, mode)
    try:
        targets, selected_count = selected_or_work_targets(
            session, selection_manager, report,
            allow_partial_selection=(action == "FREEZE"),
        )
    except Exception as error:
        report["message"] = error_text(error)
        return report

    report["selected_object_count"] = selected_count
    report["target_source"] = targets[0]["source"]
    required_workflow = FREEZE_WORKFLOW if action == "FREEZE" else UNFREEZE_WORKFLOW
    report["required_workflow"] = required_workflow
    report["preflight"]["errors"].extend(report.get("selection_warnings") or [])
    for index, target in enumerate(targets):
        target_report = make_target_report(action, session, target, build, mode, index)
        report["targets"].append(target_report)

    targets, report["targets"] = collapse_exact_identity_targets(
        targets, report["targets"]
    )
    report["unique_target_count"] = len(targets)
    report["duplicate_occurrences_collapsed"] = sum(
        max(0, target["occurrence_count"] - 1) for target in targets
    )

    for target, target_report in zip(targets, report["targets"]):
        preflight_workflow(session, target, target_report, required_workflow)
        if target_report["result"] != "PREFLIGHT_READY":
            report["preflight"]["errors"].append(
                "Target {0}: {1}".format(
                    target_report["target_index"], target_report["message"]
                )
            )

    if action == "UNFREEZE" and report["preflight"]["errors"]:
        for target_report in report["targets"]:
            if target_report["result"] == "PREFLIGHT_READY":
                target_report["result"] = "NOT_ATTEMPTED_BATCH_BLOCKED"
                target_report["message"] = (
                    "Another selected target failed J31 preflight; nothing was changed."
                )
        report["result"] = "BLOCKED_BATCH"
        report["message"] = "Complete batch preflight failed; nothing changed."
        report["counts"] = result_counts(report["targets"])
        return report

    ready_pairs = [
        (target, target_report)
        for target, target_report in zip(targets, report["targets"])
        if target_report["result"] == "PREFLIGHT_READY"
    ]
    report["preflight"]["passed"] = not report["preflight"]["errors"]
    if not ready_pairs:
        report["result"] = "BLOCKED_ALL_TARGETS"
        report["message"] = "No target passed preflight; nothing changed."
        report["counts"] = result_counts(report["targets"])
        return report

    if mode == "DRY_RUN":
        for _, target_report in ready_pairs:
            target_report["result"] = "DRY_RUN_READY"
        report["counts"] = result_counts(report["targets"])
        report["result"] = (
            "DRY_RUN_READY" if not report["preflight"]["errors"]
            else "DRY_RUN_PARTIAL"
        )
        report["message"] = (
            "Preflight completed; no mutations attempted. "
            "Review each target result before APPLY."
        )
        return report

    ready_targets = [pair[0] for pair in ready_pairs]
    ready_reports = [pair[1] for pair in ready_pairs]
    unexpected_runtime_error = ""
    try:
        if action == "FREEZE":
            execute_freeze_batch(session, ready_targets, ready_reports, report)
        else:
            execute_unfreeze_batch(session, ready_targets, ready_reports, report)
    except Exception as error:
        unexpected_runtime_error = error_text(error)
        capture_after_states(session, ready_targets, ready_reports)
        mark_recovery_results(action, ready_reports)
        report["failed_stage"] = report["failed_stage"] or "UNEXPECTED_RUNTIME_ERROR"
        report["message"] = unexpected_runtime_error

    report["counts"] = result_counts(report["targets"])
    if report["counts"]["succeeded"] == len(report["targets"]):
        report["result"] = (
            "ALL_TARGETS_FROZEN" if action == "FREEZE" else "ALL_TARGETS_UNFROZEN"
        )
        report["failed_stage"] = ""
        report["message"] = "All {0} unique target(s) completed and verified.".format(
            len(targets)
        )
    elif any(row["result"] == "RECOVERY_REQUIRED" for row in report["targets"]):
        report["result"] = "RECOVERY_REQUIRED"
        report["message"] = unexpected_runtime_error or (
            "A J31 target is in an incomplete state. Inspect it before rerunning."
        )
    elif report["counts"]["succeeded"]:
        report["result"] = "PARTIAL_COMPLETION"
        report["message"] = (
            "Safe targets completed; blocked or failed targets were isolated and reported."
        )
    else:
        report["result"] = "NO_TARGETS_COMPLETED"
        report["message"] = "No selected target completed; review target results."
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
        "Target: unique selected component/owning-component CAD parts, or the active work part when none resolve.",
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
