"""J20 - Assembly full-load failure diagnostic.

Use this journal when a partially/minimally loaded assembly is usable, but NX
fails when the user requests Full Load (including IM0541 invalid OM-object
errors). The journal first records the unloaded assembly structure, then fully
loads each unique component prototype individually. A final assembly-wide
LoadFully call verifies the same transition that fails interactively.

The journal changes only the in-memory load state. It never saves, checks out,
checks in, suppresses, replaces, or edits any NX or Teamcenter object.

Target: NX X 2506, local and Teamcenter-managed assemblies.
Run via: NX > Tools > Journal > Play
"""

import csv
import datetime
import os
import traceback

import NXOpen


BUILD = "J20-NX2506-ASSEMBLY-FULL-LOAD-DIAGNOSTIC-V1"
OUTPUT_FOLDER = "NX_ASSEMBLY_FULL_LOAD_DIAGNOSTIC"
REPORT_NAME = "NX_Assembly_Load_Diagnostic_Report.txt"
CSV_NAME = "NX_Assembly_Full_Load_Diagnostic.csv"
MAX_OCCURRENCES = 100000

CSV_COLUMNS = (
    "ROW_TYPE",
    "RUN_TIMESTAMP",
    "JOURNAL_BUILD",
    "ASSEMBLY_NAME",
    "COMPONENT_NAME",
    "PARENT_ASSEMBLY",
    "ASSEMBLY_PATH",
    "LEVEL",
    "PART_NUMBER",
    "REVISION",
    "PROTOTYPE_NAME",
    "PROTOTYPE_PATH",
    "REFERENCE_SET",
    "SUPPRESSED",
    "INITIAL_LOAD_STATE",
    "FINAL_LOAD_STATE",
    "FULL_LOAD_PROBE",
    "STATUS",
    "FAILED_OPERATION",
    "REASON",
    "EXCEPTION",
    "LOAD_STATUS_DETAILS",
    "RECOMMENDATION",
)

INVALID_OBJECT_TOKENS = (
    "im0541",
    "invalid or unsuitable om object",
    "invalid om object",
)

MISSING_FILE_TOKENS = (
    "failed to find file",
    "file not found",
    "cannot find the file",
    "could not find file",
    "not found using current search options",
    "no such file",
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
    return "{0}{1}".format(message, " [NX error {0}]".format(code) if code else "")


def contains_token(value, tokens):
    lowered = clean(value).lower()
    return any(token in lowered for token in tokens)


def classify_failure(details, default="ERROR"):
    if contains_token(details, INVALID_OBJECT_TOKENS):
        return "INVALID_OBJECT"
    if contains_token(details, MISSING_FILE_TOKENS):
        return "MISSING_FILE"
    return default


def recommendation(status):
    return {
        "OK": "No corrective action required.",
        "MISSING_FILE": (
            "Restore the referenced file or correct the Teamcenter revision "
            "rule and NX assembly search/load options."
        ),
        "PROTOTYPE_UNAVAILABLE": (
            "Check Teamcenter access, revision rule, dataset availability, "
            "and local assembly search paths, then reopen the assembly."
        ),
        "UNLOADED": (
            "Review the load-status details; NX returned without making the "
            "prototype fully loaded."
        ),
        "INVALID_OBJECT": (
            "Repair or replace the prototype/occurrence at the reported "
            "assembly path, then reopen NX and rerun J20."
        ),
        "ERROR": (
            "Review the failed operation and exception, then inspect the "
            "reported occurrence in Assembly Navigator."
        ),
    }.get(status, "Review the diagnostic details.")


def dispose(value):
    if value is None:
        return
    for method_name in ("Dispose", "FreeResource"):
        method = getattr(value, method_name, None)
        if callable(method):
            try:
                method()
            except Exception:
                pass
            return


def log_line(session, message):
    text = str(message)
    try:
        window = session.ListingWindow
        window.Open()
        writer = getattr(window, "WriteFullline", None)
        if not callable(writer):
            writer = getattr(window, "WriteLine", None)
        for line in text.splitlines() or [""]:
            if callable(writer):
                writer(line)
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


def timestamp_text(now=None):
    value = now or datetime.datetime.now().astimezone()
    return value.isoformat(timespec="seconds")


def run_folder(now=None):
    value = now or datetime.datetime.now()
    folder = os.path.join(io_root(), OUTPUT_FOLDER, value.strftime("%Y%m%d_%H%M%S"))
    os.makedirs(folder, exist_ok=True)
    return folder


def safe_value(value, property_name, fallback=""):
    if value is None:
        return fallback
    try:
        result = getattr(value, property_name)
        if callable(result):
            result = result()
        return result
    except Exception:
        return fallback


def safe_name(value, fallback="<unavailable>"):
    for property_name in ("DisplayName", "Name", "Leaf", "JournalIdentifier", "FullPath"):
        result = clean(safe_value(value, property_name))
        if result:
            return result
    return fallback


def get_string_attribute(nx_object, names):
    for name in names:
        try:
            result = clean(nx_object.GetStringAttribute(name))
            if result:
                return result
        except Exception:
            pass
        try:
            info = nx_object.GetUserAttribute(
                name,
                NXOpen.NXObject.AttributeType.String,
                -1,
            )
            result = clean(getattr(info, "StringValue", ""))
            if result:
                return result
        except Exception:
            pass
    return ""


def object_key(value):
    if value is None:
        return ("NONE", "")
    try:
        return ("TAG", clean(value.Tag))
    except Exception:
        pass
    for property_name in ("JournalIdentifier", "FullPath", "Name", "Leaf"):
        result = clean(safe_value(value, property_name))
        if result:
            return (property_name.upper(), result.upper())
    return ("PYTHON_ID", str(id(value)))


def component_children(component):
    try:
        return list(component.GetChildren())
    except Exception:
        return []


def initial_record(assembly_name, component, parent_name, parent_path, level, run_timestamp):
    component_name = safe_name(component, "<component unavailable>")
    path = "{0} / {1}".format(parent_path, component_name)
    row = {column: "" for column in CSV_COLUMNS}
    row.update(
        {
            "ROW_TYPE": "COMPONENT",
            "RUN_TIMESTAMP": run_timestamp,
            "JOURNAL_BUILD": BUILD,
            "ASSEMBLY_NAME": assembly_name,
            "COMPONENT_NAME": component_name,
            "PARENT_ASSEMBLY": parent_name,
            "ASSEMBLY_PATH": path,
            "LEVEL": level,
            "REFERENCE_SET": clean(safe_value(component, "ReferenceSet")),
            "SUPPRESSED": "YES" if bool(safe_value(component, "IsSuppressed", False)) else "NO",
            "FULL_LOAD_PROBE": "NOT_RUN",
            "STATUS": "OK",
            "REASON": "Awaiting controlled full-load probe.",
            "RECOMMENDATION": recommendation("OK"),
            "_component": component,
            "_prototype": None,
            "_prototype_key": ("NONE", ""),
        }
    )

    try:
        prototype = component.Prototype
        row["_prototype"] = prototype
        row["_prototype_key"] = object_key(prototype)
    except Exception as error:
        details = error_text(error)
        status = classify_failure(details)
        row["_prototype_key"] = ("UNRESOLVED_OCCURRENCE", path)
        row.update(
            {
                "STATUS": status,
                "FAILED_OPERATION": "Component.Prototype",
                "REASON": "NX failed while resolving the occurrence prototype.",
                "EXCEPTION": details,
                "RECOMMENDATION": recommendation(status),
            }
        )
        return row

    prototype = row["_prototype"]
    if prototype is None:
        row["_prototype_key"] = ("UNRESOLVED_OCCURRENCE", path)
        row.update(
            {
                "STATUS": "PROTOTYPE_UNAVAILABLE",
                "REASON": "The occurrence exists, but NX returned no prototype object.",
                "RECOMMENDATION": recommendation("PROTOTYPE_UNAVAILABLE"),
            }
        )
        return row

    row["PROTOTYPE_NAME"] = safe_name(prototype)
    row["PROTOTYPE_PATH"] = clean(safe_value(prototype, "FullPath"))
    row["PART_NUMBER"] = get_string_attribute(
        prototype,
        ("DB_PART_NO", "ITEM_ID", "PART_NUMBER"),
    )
    row["REVISION"] = get_string_attribute(
        prototype,
        ("DB_PART_REV", "ITEM_REVISION", "REVISION"),
    )
    state = enum_text(safe_value(prototype, "PartLoadState"))
    row["INITIAL_LOAD_STATE"] = state
    row["FINAL_LOAD_STATE"] = state
    return row


def collect_occurrences(work_part, run_timestamp):
    assembly_name = safe_name(work_part, "<work part unavailable>")
    errors = []
    try:
        root = work_part.ComponentAssembly.RootComponent
    except Exception as error:
        return [], ["RootComponent: {0}".format(error_text(error))], False

    if root is None:
        return [], [], False

    try:
        root_children = list(root.GetChildren())
    except Exception as error:
        return [], ["RootComponent.GetChildren: {0}".format(error_text(error))], True

    if not root_children:
        return [], [], False

    records = []
    stack = [
        (component, 1, assembly_name, assembly_name)
        for component in reversed(root_children)
    ]
    while stack:
        if len(records) >= MAX_OCCURRENCES:
            errors.append(
                "Traversal stopped at the safety limit of {0} occurrences.".format(
                    MAX_OCCURRENCES
                )
            )
            break

        component, level, parent_name, parent_path = stack.pop()
        record = initial_record(
            assembly_name,
            component,
            parent_name,
            parent_path,
            level,
            run_timestamp,
        )
        records.append(record)

        try:
            children = list(component.GetChildren())
        except Exception as error:
            details = error_text(error)
            status = classify_failure(details)
            record.update(
                {
                    "STATUS": status,
                    "FAILED_OPERATION": "Component.GetChildren",
                    "REASON": "NX failed while reading child occurrences.",
                    "EXCEPTION": details,
                    "RECOMMENDATION": recommendation(status),
                }
            )
            children = []

        for child in reversed(children):
            stack.append((child, level + 1, record["COMPONENT_NAME"], record["ASSEMBLY_PATH"]))

    return records, errors, True


def part_load_status_details(load_status):
    if load_status is None:
        return [], 0

    details = []
    try:
        count = int(load_status.NumberUnloadedParts)
    except Exception:
        count = 0
    details.append("NumberUnloadedParts={0}".format(count))

    for index in range(count):
        try:
            name = clean(load_status.GetPartName(index))
        except Exception:
            name = "<not available>"
        try:
            code = clean(load_status.GetStatus(index))
        except Exception:
            code = "<not available>"
        try:
            description = clean(load_status.GetStatusDescription(index))
        except Exception:
            description = "<not available>"
        details.append(
            "part={0}; status={1}; description={2}".format(name, code, description)
        )
    return details, count


def unwrap_load_status(value):
    if isinstance(value, (tuple, list)):
        return value[0] if value else None
    return value


def set_group_result(group, probe, status, operation, reason, exception, details, final_state):
    for row in group:
        row.update(
            {
                "FULL_LOAD_PROBE": probe,
                "STATUS": status,
                "FAILED_OPERATION": operation,
                "REASON": reason,
                "EXCEPTION": exception,
                "LOAD_STATUS_DETAILS": " | ".join(details),
                "FINAL_LOAD_STATE": final_state or row["FINAL_LOAD_STATE"],
                "RECOMMENDATION": recommendation(status),
            }
        )


def probe_prototype_group(group, logger=None):
    prototype = group[0].get("_prototype")
    paths = [row["ASSEMBLY_PATH"] for row in group]

    if all(row["SUPPRESSED"] == "YES" for row in group):
        set_group_result(
            group,
            "SKIPPED_SUPPRESSED",
            "OK",
            "",
            "All occurrences of this prototype are suppressed.",
            "",
            [],
            group[0]["FINAL_LOAD_STATE"],
        )
        return

    if prototype is None:
        existing = group[0]["STATUS"]
        status = existing if existing in ("INVALID_OBJECT", "MISSING_FILE") else "PROTOTYPE_UNAVAILABLE"
        set_group_result(
            group,
            "FAILED",
            status,
            group[0]["FAILED_OPERATION"] or "Component.Prototype",
            group[0]["REASON"],
            group[0]["EXCEPTION"],
            ["No prototype object was available for full loading."],
            group[0]["FINAL_LOAD_STATE"],
        )
        return

    try:
        if bool(prototype.IsFullyLoaded):
            state = enum_text(safe_value(prototype, "PartLoadState"))
            set_group_result(
                group,
                "NOT_REQUIRED_ALREADY_FULLY_LOADED",
                "OK",
                "",
                "The prototype was fully loaded before J20 reached it.",
                "",
                [],
                state,
            )
            return
    except Exception as error:
        details = error_text(error)
        status = classify_failure(details)
        set_group_result(
            group,
            "FAILED",
            status,
            "Prototype.IsFullyLoaded",
            "NX failed while checking the prototype load state.",
            details,
            [],
            group[0]["FINAL_LOAD_STATE"],
        )
        return

    if logger:
        logger(
            "Full-loading prototype: {0}\n  Occurrences: {1}\n  First path: {2}".format(
                group[0]["PROTOTYPE_NAME"], len(paths), paths[0]
            )
        )

    load_status = None
    try:
        load_status = unwrap_load_status(prototype.LoadThisPartFully())
        details, unloaded_count = part_load_status_details(load_status)
        final_state = enum_text(safe_value(prototype, "PartLoadState"))

        if unloaded_count:
            detail_text = " | ".join(details)
            status = classify_failure(detail_text, "PROTOTYPE_UNAVAILABLE")
            set_group_result(
                group,
                "FAILED",
                status,
                "BasePart.LoadThisPartFully",
                "NX reported one or more parts that could not be fully loaded.",
                "",
                details,
                final_state,
            )
            return

        if not bool(safe_value(prototype, "IsFullyLoaded", False)):
            set_group_result(
                group,
                "FAILED",
                "UNLOADED",
                "BasePart.LoadThisPartFully verification",
                "NX returned without an exception but the prototype is still not fully loaded.",
                "",
                details,
                final_state,
            )
            return

        set_group_result(
            group,
            "SUCCESS",
            "OK",
            "",
            "The component prototype fully loaded successfully.",
            "",
            details,
            final_state,
        )
    except Exception as error:
        details = error_text(error)
        status = classify_failure(details)
        set_group_result(
            group,
            "FAILED",
            status,
            "BasePart.LoadThisPartFully",
            "The controlled prototype full-load call raised an exception.",
            details,
            [],
            enum_text(safe_value(prototype, "PartLoadState")),
        )
    finally:
        dispose(load_status)


def probe_all_prototypes(records, logger=None):
    groups = {}
    order = []
    for row in records:
        key = row["_prototype_key"]
        if key not in groups:
            groups[key] = []
            order.append(key)
        groups[key].append(row)

    for key in order:
        probe_prototype_group(groups[key], logger=logger)
    return len(order)


def probe_assembly_full_load(work_part):
    result = {
        "status": "NOT_RUN",
        "operation": "BasePart.LoadFully",
        "details": [],
        "exception": "",
    }
    load_status = None
    try:
        load_status = unwrap_load_status(work_part.LoadFully())
        details, unloaded_count = part_load_status_details(load_status)
        result["details"] = details
        if unloaded_count:
            result["status"] = classify_failure(
                " | ".join(details),
                "PROTOTYPE_UNAVAILABLE",
            )
        else:
            result["status"] = "SUCCESS"
    except Exception as error:
        result["exception"] = error_text(error)
        result["status"] = classify_failure(result["exception"])
    finally:
        dispose(load_status)
    return result


def public_row(row):
    return {column: row.get(column, "") for column in CSV_COLUMNS}


def assembly_summary_row(assembly_name, run_timestamp, result):
    row = {column: "" for column in CSV_COLUMNS}
    status = result["status"]
    row.update(
        {
            "ROW_TYPE": "ASSEMBLY_SUMMARY",
            "RUN_TIMESTAMP": run_timestamp,
            "JOURNAL_BUILD": BUILD,
            "ASSEMBLY_NAME": assembly_name,
            "FULL_LOAD_PROBE": status,
            "STATUS": "OK" if status == "SUCCESS" else status,
            "FAILED_OPERATION": "" if status == "SUCCESS" else result["operation"],
            "REASON": (
                "The final assembly-wide full-load verification succeeded."
                if status == "SUCCESS"
                else "The final assembly-wide full-load verification failed."
            ),
            "EXCEPTION": result["exception"],
            "LOAD_STATUS_DETAILS": " | ".join(result["details"]),
            "RECOMMENDATION": recommendation("OK" if status == "SUCCESS" else status),
        }
    )
    return row


def write_csv(path, rows):
    with open(path, "w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(handle, fieldnames=CSV_COLUMNS, extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            writer.writerow(public_row(row))


def write_text_report(path, assembly_name, run_timestamp, records, traversal_errors, assembly_probe):
    failed = [row for row in records if row["STATUS"] != "OK"]
    with open(path, "w", encoding="utf-8-sig") as handle:
        handle.write("================================================\n")
        handle.write("NX Assembly Full Load Diagnostic Report\n")
        handle.write("================================================\n\n")
        handle.write("Build: {0}\n".format(BUILD))
        handle.write("Generated: {0}\n".format(run_timestamp))
        handle.write("Assembly: {0}\n".format(assembly_name))
        handle.write("Component occurrences scanned: {0}\n".format(len(records)))
        handle.write("Failed component occurrences: {0}\n".format(len(failed)))
        handle.write("Final assembly full-load result: {0}\n\n".format(assembly_probe["status"]))

        if assembly_probe["exception"]:
            handle.write("Assembly full-load exception:\n{0}\n\n".format(assembly_probe["exception"]))
        if assembly_probe["details"]:
            handle.write("Assembly load-status details:\n{0}\n\n".format(
                "\n".join("- " + item for item in assembly_probe["details"])
            ))
        if traversal_errors:
            handle.write("Traversal errors:\n{0}\n\n".format(
                "\n".join("- " + item for item in traversal_errors)
            ))

        handle.write("================================================\n")
        handle.write("FAILED COMPONENTS\n")
        handle.write("================================================\n")
        if not failed:
            handle.write("No component-level full-load failure was detected.\n")
        for row in failed:
            handle.write("\n------------------------------------------------\n")
            for label, key in (
                ("Component", "COMPONENT_NAME"),
                ("Assembly path", "ASSEMBLY_PATH"),
                ("Level", "LEVEL"),
                ("Part number", "PART_NUMBER"),
                ("Revision", "REVISION"),
                ("Prototype", "PROTOTYPE_NAME"),
                ("Initial load state", "INITIAL_LOAD_STATE"),
                ("Final load state", "FINAL_LOAD_STATE"),
                ("Full-load probe", "FULL_LOAD_PROBE"),
                ("Status", "STATUS"),
                ("Failed operation", "FAILED_OPERATION"),
                ("Reason", "REASON"),
                ("Exception", "EXCEPTION"),
                ("Load-status details", "LOAD_STATUS_DETAILS"),
                ("Recommended action", "RECOMMENDATION"),
            ):
                handle.write("{0}:\n{1}\n\n".format(label, clean(row.get(key)) or "<not available>"))


def main():
    session = NXOpen.Session.GetSession()
    run_timestamp = timestamp_text()
    work_part = getattr(session.Parts, "Work", None)

    log_line(session, "NX Assembly Full Load Diagnostic Started...")
    log_line(session, "Journal build: {0}".format(BUILD))
    log_line(session, "This run will fully load components in memory but will not save any part.")

    if work_part is None:
        log_line(session, "FAILED: No NX work part is open.")
        return

    assembly_name = safe_name(work_part, "<work part unavailable>")
    records, traversal_errors, is_assembly = collect_occurrences(work_part, run_timestamp)
    if not is_assembly:
        log_line(session, "FAILED: The current work part is not an assembly.")
        return

    log_line(session, "Assembly: {0}".format(assembly_name))
    log_line(session, "Occurrences captured before full load: {0}".format(len(records)))

    unique_count = probe_all_prototypes(
        records,
        logger=lambda message: log_line(session, message),
    )
    log_line(session, "Unique prototypes probed: {0}".format(unique_count))
    log_line(session, "Running final assembly-wide LoadFully verification...")
    assembly_probe = probe_assembly_full_load(work_part)

    folder = run_folder()
    csv_path = os.path.join(folder, CSV_NAME)
    report_path = os.path.join(folder, REPORT_NAME)
    rows = list(records)
    rows.append(assembly_summary_row(assembly_name, run_timestamp, assembly_probe))
    write_csv(csv_path, rows)
    write_text_report(
        report_path,
        assembly_name,
        run_timestamp,
        records,
        traversal_errors,
        assembly_probe,
    )

    failed_count = sum(1 for row in records if row["STATUS"] != "OK")
    log_line(session, "Failed component occurrences: {0}".format(failed_count))
    log_line(session, "Final assembly full-load result: {0}".format(assembly_probe["status"]))
    log_line(session, "CSV: {0}".format(csv_path))
    log_line(session, "Report: {0}".format(report_path))


def get_unload_option(dummy):
    return NXOpen.Session.LibraryUnloadOption.Immediately


if __name__ == "__main__":
    try:
        main()
    except Exception:
        try:
            session = NXOpen.Session.GetSession()
            log_line(session, "J20 FAILED:\n{0}".format(traceback.format_exc()))
        except Exception:
            pass
        raise
