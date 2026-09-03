"""J32 - read-only NX X 2506 WAE freeze capability probe.

Run with exactly one preselected Assembly Navigator component, or with no
selection to inspect the active work part.  The journal inventories runtime
members whose names may expose locking, lifecycle, workflow, release-status,
or access-control capabilities.  V3 invokes only documented read queries and
never checks out, checks in, saves, releases, or writes an NX attribute.
"""

import datetime
import inspect
import json
import os
import re
import traceback

import NXOpen
import NXOpen.PDM


BUILD = "J32-NX2506-WAE-FREEZE-CAPABILITY-PROBE-V3"
OUTPUT_FOLDER = "NX_WAE_CHANGE_CONTROL"
CANDIDATE_PATTERN = re.compile(
    r"lock|unlock|release|status|workflow|access|checkout|checkin|reserve|"
    r"revision|permission|protect|freeze|maturity|lifecycle|promote|demote|state",
    re.IGNORECASE,
)
IDENTITY_TITLES = ("DB_PART_NO", "DB_PART_REV", "WAE_VERSION")


def clean(value):
    return "" if value is None else str(value).strip()


def error_text(error):
    code = clean(getattr(error, "ErrorCode", ""))
    suffix = " [{0}]".format(code) if code else ""
    return "{0}{1}".format(clean(error) or type(error).__name__, suffix)


def safe_property(value, name, default=None):
    try:
        result = getattr(value, name)
        return result() if callable(result) else result
    except Exception:
        return default


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


def object_type(value):
    python_type = type(value)
    return "{0}.{1}".format(
        clean(getattr(python_type, "__module__", "")),
        clean(getattr(python_type, "__name__", "")),
    ).strip(".")


def part_identifier(part):
    for name in ("JournalIdentifier", "FullPath", "Name", "Leaf"):
        value = safe_property(part, name, "")
        if clean(value):
            return clean(value)
    return "<unknown>"


def resolve_target(session, selection_manager):
    try:
        count = int(selection_manager.GetNumSelectedObjects())
    except Exception as error:
        raise RuntimeError("Could not inspect NX preselection: " + error_text(error))
    if count > 1:
        raise RuntimeError(
            "J32 accepts zero or one selected Assembly Navigator component; found {0}.".format(
                count
            )
        )
    if count == 0:
        parts = safe_property(session, "Parts")
        part = safe_property(parts, "Work")
        if part is None:
            raise RuntimeError("There is no active work part.")
        source = "ACTIVE_WORK_PART"
        component = None
    else:
        try:
            component = selection_manager.GetSelectedTaggedObject(0)
        except Exception as error:
            raise RuntimeError("Could not read the selected object: " + error_text(error))
        part = safe_property(component, "Prototype")
        if part is None:
            raise RuntimeError(
                "The selected object is not a loaded Assembly Navigator component."
            )
        if safe_property(component, "IsSuppressed") is True:
            raise RuntimeError("The selected component is suppressed.")
        source = "ASSEMBLY_NAVIGATOR_SELECTION"
    if safe_property(part, "PDMPart") is None:
        raise RuntimeError("The target has no PDMPart; use a loaded Teamcenter-managed CAD part.")
    managed = bool(safe_property(session, "IsManagedMode", False))
    if not managed and not part_identifier(part).upper().startswith("@DB/"):
        raise RuntimeError("The target is not positively Teamcenter-managed.")
    return component, part, source


def read_identity(part, title):
    method = getattr(part, "GetStringAttribute", None)
    if not callable(method):
        return ""
    try:
        return clean(method(title))
    except Exception:
        return ""


def checkout_snapshot(session, part):
    pdm_part = safe_property(part, "PDMPart")
    method = getattr(pdm_part, "GetCheckedoutStatusAndUser", None)
    if not callable(method):
        return {"state": "UNKNOWN", "owner": "", "raw": "API unavailable"}
    try:
        raw = method()
    except TypeError:
        try:
            raw = method(False, "")
        except Exception as error:
            return {"state": "UNKNOWN", "owner": "", "raw": error_text(error)}
    except Exception as error:
        return {"state": "UNKNOWN", "owner": "", "raw": error_text(error)}
    checked = None
    owner = ""
    if isinstance(raw, (tuple, list)):
        for value in raw:
            if checked is None and isinstance(value, bool):
                checked = value
            elif isinstance(value, str) and not owner:
                owner = clean(value)
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
    return {"state": state, "owner": owner, "raw": repr(raw)[:2000]}


def public_member_names(value):
    try:
        return sorted({name for name in dir(value) if not name.startswith("_")})
    except Exception:
        return []


def candidate_member_names(value):
    return [name for name in public_member_names(value) if CANDIDATE_PATTERN.search(name)]


def annotation_strings(value):
    annotations = getattr(value, "__annotations__", None)
    if not isinstance(annotations, dict):
        return {}
    return {
        clean(key): clean(annotation)
        for key, annotation in annotations.items()
    }


def candidate_member_metadata(value, names=None):
    """Inspect callable metadata without invoking any discovered member."""
    rows = []
    for name in names or candidate_member_names(value):
        row = {
            "name": name,
            "lookup_error": "",
            "python_type": "",
            "callable": False,
            "repr": "",
            "doc": "",
            "text_signature": "",
            "inspect_signature": "",
            "inspect_signature_error": "",
            "annotations": {},
            "overloads_repr": "",
        }
        try:
            member = getattr(value, name)
            row["python_type"] = object_type(member)
            row["callable"] = callable(member)
            row["repr"] = repr(member)[:4000]
            row["doc"] = clean(getattr(member, "__doc__", ""))[:12000]
            row["text_signature"] = clean(
                getattr(member, "__text_signature__", "")
            )[:4000]
            row["annotations"] = annotation_strings(member)
            overloads = getattr(member, "__overloads__", None)
            if overloads is None:
                overloads = getattr(member, "Overloads", None)
            if overloads is not None:
                row["overloads_repr"] = repr(overloads)[:12000]
            if row["callable"]:
                try:
                    row["inspect_signature"] = str(inspect.signature(member))
                except Exception as error:
                    row["inspect_signature_error"] = error_text(error)
        except Exception as error:
            row["lookup_error"] = error_text(error)
        rows.append(row)
    return rows


def reflection_signatures(value):
    """Read .NET metadata only; do not invoke any discovered candidate member."""
    result = {"runtime_type": "", "methods": [], "properties": [], "error": ""}
    get_type = getattr(value, "GetType", None)
    if not callable(get_type):
        result["error"] = "GetType unavailable"
        return result
    try:
        runtime_type = get_type()
        result["runtime_type"] = clean(runtime_type)
        for method in runtime_type.GetMethods():
            name = clean(method.Name)
            if not CANDIDATE_PATTERN.search(name):
                continue
            parameters = []
            for parameter in method.GetParameters():
                parameters.append({
                    "name": clean(parameter.Name),
                    "type": clean(parameter.ParameterType),
                })
            result["methods"].append({
                "name": name,
                "return_type": clean(method.ReturnType),
                "parameters": parameters,
            })
        for prop in runtime_type.GetProperties():
            name = clean(prop.Name)
            if CANDIDATE_PATTERN.search(name):
                result["properties"].append({
                    "name": name,
                    "type": clean(prop.PropertyType),
                })
        result["methods"].sort(key=lambda row: (row["name"], repr(row["parameters"])))
        result["properties"].sort(key=lambda row: row["name"])
    except Exception as error:
        result["error"] = error_text(error)
    return result


def inspect_object(label, value):
    members = public_member_names(value)
    candidates = [name for name in members if CANDIDATE_PATTERN.search(name)]
    return {
        "label": label,
        "python_type": object_type(value),
        "all_public_members": members,
        "candidate_members": candidates,
        "candidate_details": candidate_member_metadata(value, candidates),
        "dotnet_reflection": reflection_signatures(value),
    }


def relevant_attributes(part):
    iterator = None
    rows = []
    try:
        iterator = part.CreateAttributeIterator()
        iterator.SetIncludeAlsoUnset(True)
        for info in part.GetUserAttributes(iterator):
            title = clean(safe_property(info, "Title", ""))
            category = clean(safe_property(info, "Category", ""))
            if title not in IDENTITY_TITLES and not CANDIDATE_PATTERN.search(
                "{0} {1}".format(category, title)
            ):
                continue
            rows.append({
                "category": category,
                "title": title,
                "type": clean(safe_property(info, "Type", "")),
                "string_value": clean(safe_property(info, "StringValue", "")),
                "locked": bool(safe_property(info, "Locked", False)),
                "owned_by_system": bool(safe_property(info, "OwnedBySystem", False)),
                "pdm_based": bool(safe_property(info, "PdmBased", False)),
                "unset": bool(safe_property(info, "Unset", False)),
            })
    except Exception as error:
        rows.append({"probe_error": error_text(error)})
    finally:
        dispose(iterator)
    return rows


def target_snapshot(session, component, part, source):
    return {
        "source": source,
        "component_name": clean(safe_property(component, "DisplayName", ""))
        or clean(safe_property(component, "Name", "")),
        "component_tag": clean(safe_property(component, "Tag", "")),
        "part_identifier": part_identifier(part),
        "part_number": read_identity(part, "DB_PART_NO"),
        "db_part_rev": read_identity(part, "DB_PART_REV"),
        "wae_version": read_identity(part, "WAE_VERSION"),
        "read_only": safe_property(part, "IsReadOnly"),
        "modified": bool(safe_property(part, "IsModified", False)),
        "checkout": checkout_snapshot(session, part),
        "relevant_attributes": relevant_attributes(part),
    }


def read_workflows(pdm_session, part):
    result = {
        "api": "PdmSession.GetNXWorkflows",
        "attempted": True,
        "workflow_names": [],
        "operation_errors_repr": "",
        "raw_repr": "",
        "error": "",
    }
    operation_errors = None
    try:
        method = getattr(pdm_session, "GetNXWorkflows", None)
        if not callable(method):
            raise RuntimeError("PdmSession.GetNXWorkflows is unavailable.")
        raw = method([part])
        result["raw_repr"] = repr(raw)[:12000]
        if not isinstance(raw, (tuple, list)) or len(raw) < 2:
            raise RuntimeError(
                "GetNXWorkflows returned an unexpected value: {0}".format(
                    result["raw_repr"]
                )
            )
        operation_errors = raw[0]
        names = raw[1]
        result["operation_errors_repr"] = repr(operation_errors)[:4000]
        if isinstance(names, str):
            names = [names]
        result["workflow_names"] = [clean(name) for name in names if clean(name)]
    except Exception as error:
        result["error"] = error_text(error)
    finally:
        dispose(operation_errors)
    return result


def read_release_status(pdm_part):
    result = {
        "api": "PdmPart.GetReleaseStatus",
        "attempted": True,
        "value": "",
        "raw_repr": "",
        "error": "",
    }
    try:
        method = getattr(pdm_part, "GetReleaseStatus", None)
        if not callable(method):
            raise RuntimeError("PdmPart.GetReleaseStatus is unavailable.")
        raw = method()
        result["raw_repr"] = repr(raw)[:12000]
        result["value"] = clean(raw)
    except Exception as error:
        result["error"] = error_text(error)
    return result


def read_internal_release_status(pdm_part, part):
    result = {
        "api": "PdmPart.GetInternalReleaseStatus",
        "attempted": True,
        "values": [],
        "raw_repr": "",
        "error": "",
    }
    try:
        method = getattr(pdm_part, "GetInternalReleaseStatus", None)
        if not callable(method):
            raise RuntimeError("PdmPart.GetInternalReleaseStatus is unavailable.")
        raw = method([part])
        result["raw_repr"] = repr(raw)[:12000]
        values = [raw] if isinstance(raw, str) else raw
        result["values"] = [clean(value) for value in values]
    except Exception as error:
        result["error"] = error_text(error)
    return result


def snapshot_differences(before, after):
    keys = (
        "part_identifier",
        "part_number",
        "db_part_rev",
        "wae_version",
        "read_only",
        "modified",
        "checkout",
        "relevant_attributes",
    )
    return [key for key in keys if before.get(key) != after.get(key)]


def base_report():
    return {
        "build": BUILD,
        "timestamp": datetime.datetime.now().astimezone().isoformat(timespec="seconds"),
        "action": "WAE_FREEZE_CAPABILITY_PROBE",
        "scope": "ONE_SELECTED_COMPONENT_OR_ACTIVE_WORK_PART",
        "result": "BLOCKED",
        "message": "",
        "strictly_read_only": True,
        "operations": {
            "checkout_attempted": False,
            "checkin_attempted": False,
            "save_attempted": False,
            "attribute_write_attempted": False,
            "release_attempted": False,
            "workflow_attempted": False,
            "mutating_candidate_api_invoked": False,
            "read_only_api_calls_attempted": False,
        },
        "target": {},
        "target_after_read_queries": {},
        "read_only_postcondition": {
            "unchanged": None,
            "different_fields": [],
        },
        "read_api_results": {},
        "runtime_objects": [],
    }


def execute(session, selection_manager):
    report = base_report()
    try:
        component, part, source = resolve_target(session, selection_manager)
        pdm_part = safe_property(part, "PDMPart")
        pdm_session = safe_property(session, "PdmSession")
        report["target"] = target_snapshot(session, component, part, source)
        objects = [
            ("target_part", part),
            ("target_pdm_part", pdm_part),
            ("session_pdm_session", pdm_session),
            ("NXOpen.PDM.PdmPart_class", getattr(NXOpen.PDM, "PdmPart", None)),
            ("NXOpen.PDM_module", NXOpen.PDM),
        ]
        for label, value in objects:
            if value is None:
                report["runtime_objects"].append({
                    "label": label,
                    "python_type": "",
                    "all_public_members": [],
                    "candidate_members": [],
                    "candidate_details": [],
                    "dotnet_reflection": {"error": "Object unavailable"},
                })
            else:
                report["runtime_objects"].append(inspect_object(label, value))
        report["operations"]["read_only_api_calls_attempted"] = True
        report["read_api_results"] = {
            "get_nx_workflows": read_workflows(pdm_session, part),
            "get_release_status": read_release_status(pdm_part),
            "get_internal_release_status": read_internal_release_status(pdm_part, part),
        }
        report["target_after_read_queries"] = target_snapshot(
            session, component, part, source
        )
        differences = snapshot_differences(
            report["target"], report["target_after_read_queries"]
        )
        report["read_only_postcondition"] = {
            "unchanged": not differences,
            "different_fields": differences,
        }
        query_errors = [
            value["error"]
            for value in report["read_api_results"].values()
            if value["error"]
        ]
        if differences:
            report["result"] = "READ_ONLY_POSTCONDITION_FAILED"
            report["message"] = "Read-query state changed: " + ", ".join(differences)
        elif query_errors:
            report["result"] = "PROBE_COMPLETE_WITH_QUERY_ERRORS"
            report["message"] = "Read-only query errors: " + " | ".join(query_errors)
        else:
            report["result"] = "PROBE_COMPLETE"
            report["message"] = (
                "Workflow and release-status queries completed with unchanged target state."
            )
    except Exception as error:
        report["message"] = error_text(error)
        report["traceback"] = traceback.format_exc()
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
    path = os.path.join(folder, "WAE_FREEZE_CAPABILITY_{0}.json".format(stamp))
    with open(path, "w", encoding="utf-8") as handle:
        json.dump(report, handle, indent=2, sort_keys=True)
    return path


def log_line(session, message):
    try:
        window = session.ListingWindow
        window.Open()
        window.WriteLine(str(message))
    except Exception:
        try:
            print(message)
        except Exception:
            pass


def main():
    session = NXOpen.Session.GetSession()
    log_line(session, "=" * 72)
    log_line(session, BUILD)
    log_line(session, "READ ONLY: no mutating freeze/release/lock API will be invoked.")
    try:
        selection_manager = NXOpen.UI.GetUI().SelectionManager
        report = execute(session, selection_manager)
    except Exception as error:
        report = base_report()
        report["result"] = "FAILED"
        report["message"] = error_text(error)
        report["traceback"] = traceback.format_exc()
    try:
        path = write_report(report)
    except Exception as error:
        path = ""
        report["message"] += " | Could not write JSON: " + error_text(error)
    log_line(session, "Result: " + report["result"])
    log_line(session, report["message"])
    if report.get("target"):
        target = report["target"]
        log_line(
            session,
            "Target: {0}/{1} WAE {2} {3}".format(
                target.get("part_number", ""),
                target.get("db_part_rev", ""),
                target.get("wae_version", ""),
                (target.get("checkout") or {}).get("state", ""),
            ),
        )
    if path:
        log_line(session, "Capability JSON: " + path)
    log_line(session, "=" * 72)
    return report


if __name__ == "__main__":
    main()
