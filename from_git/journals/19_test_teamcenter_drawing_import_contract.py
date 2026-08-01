"""J19 - one-shot, read-only Teamcenter drawing payload contract probe.

NX X 2506 only.

This journal investigates how the exact managed UGPART drawing payload can be
retrieved for J16 verification.  It opens one canonical /specification/ part
and probes three local-download paths independently:

1. FileManagement.GetAssociatedFiles + DownloadAssociatedFiles
2. FileManagement.ExportNamedReferences for the UGPART named reference
3. FileManagement.ExportFiles for the known "has specification" relation

It never checks out, checks in, saves, imports, or invokes UF Clone.  The only
files it creates are local evidence files and downloads requested from
Teamcenter.  Run it from a fresh managed NX X 2506 session with no parts open.
"""

import hashlib
import importlib.util
import inspect
import json
import os
import traceback
import zipfile

import NXOpen


USER_PART_NUMBER = "264MN021218A01"
USER_REVISION = "A"
USER_DWG_INDEX = 1
BUILD = "J19-J16-TEAMCENTER-PAYLOAD-CONTRACT-NX2506-V2"
RELATION_TYPE = "has specification"
NAMED_REFERENCE = "UGPART"


def load_j16():
    path = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "16_tc_offline_drawing_import.py",
    )
    spec = importlib.util.spec_from_file_location("nx_journal_16_probe", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


J16 = load_j16()


def configured_target():
    part_number = J16.env("NX_J19_PART_NUMBER") or USER_PART_NUMBER
    revision = J16.env("NX_J19_REVISION") or USER_REVISION
    raw_index = J16.env("NX_J19_DWG_INDEX") or str(USER_DWG_INDEX)
    try:
        drawing_index = int(raw_index)
    except Exception:
        raise RuntimeError("NX_J19_DWG_INDEX must be an integer.")
    if drawing_index < 1:
        raise RuntimeError("NX_J19_DWG_INDEX must be >= 1.")
    return part_number, revision, drawing_index


def current_teamcenter_user(session):
    pdm = getattr(session, "PdmSession", None)
    method = getattr(pdm, "GetUserName", None)
    if not callable(method):
        return ""
    try:
        return J16.clean(method())
    except Exception:
        return ""


def sha256(path):
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        while True:
            block = handle.read(1024 * 1024)
            if not block:
                break
            digest.update(block)
    return digest.hexdigest()


def type_name(value):
    if value is None:
        return "NoneType"
    value_type = type(value)
    module = getattr(value_type, "__module__", "")
    name = getattr(value_type, "__name__", str(value_type))
    return "{0}.{1}".format(module, name) if module else name


def compact_repr(value, limit=12000):
    try:
        rendered = repr(value)
    except Exception as exc:
        rendered = "<repr failed: {0}>".format(J16.error_text(exc))
    return rendered if len(rendered) <= limit else rendered[: limit - 3] + "..."


def describe_value(value, depth=0):
    result = {
        "type": type_name(value),
        "repr": compact_repr(value),
    }
    if depth >= 2:
        return result
    if isinstance(value, (tuple, list)):
        result["items"] = [describe_value(item, depth + 1) for item in value]
    elif isinstance(value, dict):
        result["items"] = {
            str(key): describe_value(item, depth + 1)
            for key, item in value.items()
        }
    return result


def exception_record(error):
    return {
        "type": type_name(error),
        "message": str(error),
        "nx_error_code": J16.clean(getattr(error, "ErrorCode", "")),
        "repr": compact_repr(error),
    }


def callable_contract(method):
    result = {
        "available": callable(method),
        "type": type_name(method),
        "repr": compact_repr(method, 4000),
        "signature": "",
        "signature_error": "",
        "doc": J16.clean(getattr(method, "__doc__", ""))[:8000],
    }
    if callable(method):
        try:
            result["signature"] = str(inspect.signature(method))
        except Exception as exc:
            result["signature_error"] = J16.error_text(exc)
    return result


def integer_values(value):
    values = []
    if isinstance(value, bool) or value is None:
        return values
    if isinstance(value, int):
        return [int(value)]
    if isinstance(value, dict):
        for item in value.values():
            values.extend(integer_values(item))
    elif isinstance(value, (tuple, list)):
        for item in value:
            values.extend(integer_values(item))
    return values


def string_values(value):
    values = []
    if isinstance(value, str):
        return [value] if value else []
    if isinstance(value, dict):
        for item in value.values():
            values.extend(string_values(item))
    elif isinstance(value, (tuple, list)):
        for item in value:
            values.extend(string_values(item))
    return values


def invoke(method, args, optional_output=None):
    """Call one runtime shape and retry only when an omitted out-param is required."""
    record = {
        "contract": callable_contract(method),
        "input_argument_count": len(args),
        "input_arguments": [describe_value(value) for value in args],
        "attempts": [],
        "status": "NOT_RUN",
        "raw_return": describe_value(None),
        "output_parameter": describe_value(optional_output),
        "pdi_codes": [],
        "returned_strings": [],
    }
    if not callable(method):
        record["status"] = "UNAVAILABLE"
        return record, None

    raw = None
    try:
        raw = method(*args)
        record["attempts"].append(
            {"argument_count": len(args), "status": "RETURNED"}
        )
    except TypeError as first_error:
        record["attempts"].append(
            {
                "argument_count": len(args),
                "status": "TYPE_ERROR",
                "error": exception_record(first_error),
            }
        )
        if optional_output is None:
            record["status"] = "ERROR"
            record["error"] = exception_record(first_error)
            return record, None
        fallback_args = args + (optional_output,)
        try:
            raw = method(*fallback_args)
            record["attempts"].append(
                {"argument_count": len(fallback_args), "status": "RETURNED"}
            )
        except Exception as fallback_error:
            record["attempts"].append(
                {
                    "argument_count": len(fallback_args),
                    "status": "ERROR",
                    "error": exception_record(fallback_error),
                }
            )
            record["status"] = "ERROR"
            record["error"] = exception_record(fallback_error)
            record["output_parameter"] = describe_value(optional_output)
            return record, None
    except Exception as error:
        record["attempts"].append(
            {
                "argument_count": len(args),
                "status": "ERROR",
                "error": exception_record(error),
            }
        )
        record["status"] = "ERROR"
        record["error"] = exception_record(error)
        return record, None

    record["status"] = "RETURNED"
    record["raw_return"] = describe_value(raw)
    record["output_parameter"] = describe_value(optional_output)
    record["pdi_codes"] = integer_values(raw)
    record["returned_strings"] = string_values(raw) + string_values(optional_output)
    return record, raw


def file_record(path, root=None):
    absolute = os.path.abspath(path)
    record = {
        "path": absolute,
        "relative_path": "",
        "exists": os.path.isfile(absolute),
        "size": "",
        "sha256": "",
        "error": "",
    }
    if root:
        try:
            record["relative_path"] = os.path.relpath(absolute, root)
        except Exception:
            pass
    if record["exists"]:
        try:
            record["size"] = os.path.getsize(absolute)
            record["sha256"] = sha256(absolute)
        except Exception as exc:
            record["error"] = J16.error_text(exc)
    return record


def directory_snapshot(root):
    result = {"root": os.path.abspath(root), "directories": [], "files": []}
    if not os.path.isdir(root):
        return result
    for folder, directories, files in os.walk(root):
        directories.sort(key=lambda value: value.lower())
        files.sort(key=lambda value: value.lower())
        for name in directories:
            result["directories"].append(
                os.path.relpath(os.path.join(folder, name), root)
            )
        for name in files:
            result["files"].append(file_record(os.path.join(folder, name), root))
    return result


def unique_file_records(groups):
    by_path = {}
    for group in groups:
        for record in group:
            path = record.get("path", "")
            if path:
                by_path[os.path.normcase(os.path.abspath(path))] = record
    return sorted(by_path.values(), key=lambda value: value["path"].lower())


def physical_files_from_strings(values, search_roots):
    records = []
    for value in values:
        if not isinstance(value, str) or not value:
            continue
        candidates = [value]
        if not os.path.isabs(value):
            candidates.extend(os.path.join(root, value) for root in search_roots)
        for candidate in candidates:
            if os.path.isfile(candidate):
                records.append(file_record(candidate))
            elif os.path.isdir(candidate):
                records.extend(directory_snapshot(candidate)["files"])
    return unique_file_records([records])


def property_record(value, name):
    try:
        return {"status": "PASS", "value": getattr(value, name)}
    except Exception as exc:
        return {"status": "ERROR", "error": exception_record(exc)}


def method_value_record(value, name):
    method = getattr(value, name, None)
    if not callable(method):
        return {"status": "UNAVAILABLE", "value": ""}
    try:
        return {"status": "PASS", "value": method()}
    except Exception as exc:
        return {"status": "ERROR", "error": exception_record(exc), "value": ""}


def describe_pdm_file(value):
    return {
        "type": type_name(value),
        "repr": compact_repr(value),
        "file_name": method_value_record(value, "GetFileName"),
        "file_size": method_value_record(value, "GetFileSize"),
        "last_modified": method_value_record(value, "GetFileLastModifiedDate"),
    }


def collect_pdm_files(value):
    files = []
    seen = set()

    def visit(candidate):
        if isinstance(candidate, (tuple, list)):
            for item in candidate:
                visit(item)
            return
        if callable(getattr(candidate, "GetFileName", None)):
            marker = id(candidate)
            if marker not in seen:
                seen.add(marker)
                files.append(candidate)

    visit(value)
    return files


def physical_pdm_files(records, search_roots):
    paths = []
    for record in records:
        value = record.get("file_name", {}).get("value", "")
        if not isinstance(value, str) or not value:
            continue
        candidates = [value]
        if not os.path.isabs(value):
            candidates.extend(os.path.join(root, value) for root in search_roots)
        for candidate in candidates:
            if os.path.isfile(candidate):
                paths.append(os.path.abspath(candidate))
    unique = sorted(set(paths), key=lambda value: value.lower())
    return [file_record(path) for path in unique]


def release_pdm_files(files):
    for value in files:
        method = getattr(value, "FreeResource", None)
        if not callable(method):
            method = getattr(value, "Dispose", None)
        if callable(method):
            try:
                method()
            except Exception:
                pass


def probe_associated_files(file_management, part, output_root):
    probe_root = os.path.join(output_root, "01_ASSOCIATED_FILES")
    os.makedirs(probe_root, exist_ok=True)
    result = {
        "status": "NOT_RUN",
        "probe_directory": probe_root,
        "before_tree": directory_snapshot(probe_root),
        "get_associated_files": {},
        "download_associated_files": {},
        "files_before_download": [],
        "files_after_download": [],
        "physical_files": [],
        "native_files": [],
    }
    pdm_files = []
    try:
        get_method = getattr(file_management, "GetAssociatedFiles", None)
        get_call, raw_files = invoke(get_method, ([part], []))
        result["get_associated_files"] = get_call
        pdm_files = collect_pdm_files(raw_files)
        result["files_before_download"] = [
            describe_pdm_file(value) for value in pdm_files
        ]

        if get_call["status"] != "RETURNED":
            result["status"] = "ERROR"
            return result
        if not pdm_files:
            result["status"] = "COMPLETE_NO_ASSOCIATED_FILES"
            return result

        download_method = getattr(file_management, "DownloadAssociatedFiles", None)
        download_call, download_raw = invoke(download_method, ([part], pdm_files))
        result["download_associated_files"] = download_call
        for value in collect_pdm_files(download_raw):
            if all(value is not existing for existing in pdm_files):
                pdm_files.append(value)
        result["files_after_download"] = [
            describe_pdm_file(value) for value in pdm_files
        ]
        result["after_tree"] = directory_snapshot(probe_root)

        search_roots = [probe_root, output_root, os.getcwd()]
        for name in ("UGII_TMP_DIR", "TEMP", "TMP"):
            value = J16.env(name)
            if value and os.path.isdir(value):
                search_roots.append(value)
        physical = physical_pdm_files(result["files_after_download"], search_roots)
        result["physical_files"] = physical
        result["native_files"] = [
            item for item in physical if item["path"].lower().endswith(".prt")
        ]
        result["status"] = (
            "PASS_NATIVE_FILE_FOUND"
            if result["native_files"]
            else "COMPLETE_NO_NATIVE_FILE"
            if download_call["status"] == "RETURNED"
            else "ERROR"
        )
        return result
    except Exception as exc:
        result["status"] = "ERROR"
        result["error"] = exception_record(exc)
        result["traceback"] = traceback.format_exc()
        return result
    finally:
        result.setdefault("after_tree", directory_snapshot(probe_root))
        release_pdm_files(pdm_files)


def probe_export_named_reference(file_management, target, output_root):
    part_number, revision, drawing_index = target
    probe_root = os.path.join(output_root, "02_EXPORT_NAMED_REFERENCE")
    os.makedirs(probe_root, exist_ok=True)
    method = getattr(file_management, "ExportNamedReferences", None)
    output_references = []
    args = (
        part_number,
        revision,
        J16.dataset_name(part_number, revision, drawing_index),
        J16.configured_dataset_type(),
        RELATION_TYPE,
        NAMED_REFERENCE,
        probe_root,
    )
    result = {
        "status": "NOT_RUN",
        "probe_directory": probe_root,
        "relation_type": RELATION_TYPE,
        "named_reference": NAMED_REFERENCE,
        "before_tree": directory_snapshot(probe_root),
    }
    call, _ = invoke(method, args, output_references)
    result["call"] = call
    result["after_tree"] = directory_snapshot(probe_root)
    result["returned_physical_files"] = physical_files_from_strings(
        call["returned_strings"], [probe_root, output_root, os.getcwd()]
    )
    physical = unique_file_records(
        [result["after_tree"]["files"], result["returned_physical_files"]]
    )
    result["native_files"] = [
        item for item in physical if item["path"].lower().endswith(".prt")
    ]
    result["status"] = (
        "PASS_NATIVE_FILE_FOUND"
        if result["native_files"]
        else "COMPLETE_NO_NATIVE_FILE"
        if call["status"] == "RETURNED"
        else call["status"]
    )
    return result


def probe_legacy_export_files(file_management, target, output_root):
    part_number, revision, drawing_index = target
    probe_root = os.path.join(output_root, "03_LEGACY_EXPORT_FILES")
    os.makedirs(probe_root, exist_ok=True)
    method = getattr(file_management, "ExportFiles", None)
    output_directories = []
    args = (
        [part_number],
        [revision],
        [J16.dataset_name(part_number, revision, drawing_index)],
        [J16.configured_dataset_type()],
        [RELATION_TYPE],
        [probe_root],
        [J16.configured_export_tool()],
    )
    result = {
        "status": "NOT_RUN",
        "probe_directory": probe_root,
        "relation_type": RELATION_TYPE,
        "before_tree": directory_snapshot(probe_root),
    }
    call, _ = invoke(method, args, output_directories)
    result["call"] = call
    result["after_tree"] = directory_snapshot(probe_root)
    result["returned_physical_files"] = physical_files_from_strings(
        call["returned_strings"], [probe_root, output_root, os.getcwd()]
    )
    physical = unique_file_records(
        [result["after_tree"]["files"], result["returned_physical_files"]]
    )
    result["native_files"] = [
        item for item in physical if item["path"].lower().endswith(".prt")
    ]
    result["status"] = (
        "PASS_NATIVE_FILE_FOUND"
        if result["native_files"]
        else "COMPLETE_NO_NATIVE_FILE"
        if call["status"] == "RETURNED"
        else call["status"]
    )
    return result


def loaded_part_records(session):
    try:
        parts = list(session.Parts)
    except Exception as exc:
        return [], exception_record(exc)
    return [
        {
            "journal_identifier": J16.journal_identifier(part),
            "leaf": J16.clean(getattr(part, "Leaf", "")),
            "type": type_name(part),
        }
        for part in parts
    ], None


def drawing_sheet_count(part):
    try:
        return len(list(part.DrawingSheets))
    except Exception:
        try:
            return int(part.DrawingSheets.Count)
        except Exception:
            return None


def part_snapshot(part):
    return {
        "journal_identifier": J16.journal_identifier(part),
        "full_path": property_record(part, "FullPath"),
        "leaf": property_record(part, "Leaf"),
        "unique_identifier": property_record(part, "UniqueIdentifier"),
        "is_read_only": property_record(part, "IsReadOnly"),
        "has_write_access": property_record(part, "HasWriteAccess"),
        "is_modified": property_record(part, "IsModified"),
        "drawing_sheet_count": drawing_sheet_count(part),
    }


def write_json(path, value):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as handle:
        json.dump(value, handle, indent=2, sort_keys=True, default=str)
        handle.write("\n")


def zip_evidence(output_root):
    zip_path = output_root + ".zip"
    base = os.path.dirname(output_root)
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as archive:
        for folder, directories, files in os.walk(output_root):
            directories.sort(key=lambda value: value.lower())
            files.sort(key=lambda value: value.lower())
            for name in files:
                path = os.path.join(folder, name)
                archive.write(path, os.path.relpath(path, base))
    return zip_path


def output_paths(timestamp):
    configured = J16.env("NX_J19_OUTPUT_DIR")
    base = os.path.abspath(os.path.expanduser(configured)) if configured else J16.io_root()
    root = os.path.join(base, "J19_CONTRACT_{0}".format(timestamp))
    return (
        root,
        os.path.join(root, "J19_CONTRACT_{0}.json".format(timestamp)),
        os.path.join(root, "J19_CONTRACT_{0}.log".format(timestamp)),
    )


def main():
    session = NXOpen.Session.GetSession()
    timestamp = J16.stamp()
    output_root, json_path, log_path = output_paths(timestamp)
    os.makedirs(output_root, exist_ok=True)
    log = J16.Log(session)
    target = configured_target()
    part_number, revision, drawing_index = target
    identifier = J16.drawing_id(part_number, revision, drawing_index)

    report = {
        "journal": BUILD,
        "timestamp": timestamp,
        "target": {
            "part_number": part_number,
            "revision": revision,
            "drawing_index": drawing_index,
            "identifier": identifier,
            "dataset_name": J16.dataset_name(part_number, revision, drawing_index),
            "dataset_type": J16.configured_dataset_type(),
        },
        "current_teamcenter_user": current_teamcenter_user(session),
        "local_environment": {
            "current_working_directory": os.getcwd(),
            "output_root": output_root,
            "UGII_TMP_DIR": J16.env("UGII_TMP_DIR"),
            "TEMP": J16.env("TEMP"),
            "TMP": J16.env("TMP"),
        },
        "fresh_session": {},
        "opened_part": {},
        "checkout": {},
        "probes": {
            "associated_files": {"status": "NOT_RUN"},
            "named_reference": {"status": "NOT_RUN"},
            "legacy_export_files": {"status": "NOT_RUN"},
        },
        "teamcenter_write_attempted": False,
        "mutation_apis_called": [],
        "result": "PROBE_INCOMPLETE",
    }

    part = None
    load_status = None
    file_management = None
    zip_path = ""
    log.write("=" * 72)
    log.write("J19 ONE-SHOT READ-ONLY TEAMCENTER PAYLOAD CONTRACT PROBE")
    log.write("Build: {0}".format(BUILD))
    log.write("Target: {0}".format(identifier))
    log.write("Output: {0}".format(output_root))
    log.write("Teamcenter writes and UF Clone: FORBIDDEN")
    log.write("=" * 72)

    try:
        loaded, loaded_error = loaded_part_records(session)
        report["fresh_session"] = {
            "loaded_part_count": len(loaded),
            "loaded_parts": loaded,
            "enumeration_error": loaded_error,
        }
        if loaded_error is not None:
            raise RuntimeError("Could not prove that the NX session is fresh.")
        if loaded:
            raise RuntimeError(
                "J19 V2 requires a fresh NX session with no parts loaded; found {0}."
                .format(len(loaded))
            )

        part, load_status = J16.unwrap_open_result(session.Parts.OpenBase(identifier))
        if part is None:
            raise RuntimeError("OpenBase returned no part for the exact specification.")
        report["opened_part"] = part_snapshot(part)
        actual = report["opened_part"]["journal_identifier"]
        if J16.upper(actual).replace("\\", "/") != J16.upper(identifier).replace("\\", "/"):
            raise RuntimeError(
                "Opened JournalIdentifier does not match the exact target: {0}"
                .format(actual or "<blank>")
            )

        checkout = J16.query_pdm_checkout(part)
        report["checkout"] = checkout
        log.write(
            "Checkout: state={0}; owner={1}; raw={2}".format(
                checkout.get("state", "UNKNOWN"),
                checkout.get("owner", "") or "<blank>",
                checkout.get("raw", "") or "<blank>",
            )
        )
        if checkout.get("state") != "CHECKED_IN":
            raise RuntimeError(
                "The one-shot target must be proven CHECKED_IN before payload probing."
            )

        _, file_management = J16.new_file_management(session)
        report["probes"]["associated_files"] = probe_associated_files(
            file_management, part, output_root
        )
        log.write(
            "Get/DownloadAssociatedFiles: {0}".format(
                report["probes"]["associated_files"]["status"]
            )
        )

        report["probes"]["named_reference"] = probe_export_named_reference(
            file_management, target, output_root
        )
        log.write(
            "ExportNamedReferences: {0}".format(
                report["probes"]["named_reference"]["status"]
            )
        )

        report["probes"]["legacy_export_files"] = probe_legacy_export_files(
            file_management, target, output_root
        )
        log.write(
            "Legacy ExportFiles: {0}".format(
                report["probes"]["legacy_export_files"]["status"]
            )
        )

        statuses = [
            value.get("status", "") for value in report["probes"].values()
        ]
        if any(status == "PASS_NATIVE_FILE_FOUND" for status in statuses):
            report["result"] = "PROBE_COMPLETE_RETRIEVAL_FOUND"
        else:
            report["result"] = "PROBE_COMPLETE_NO_RETRIEVAL"
        log.write("FINAL STATUS: {0}".format(report["result"]))
    except Exception as exc:
        report["error"] = exception_record(exc)
        report["traceback"] = traceback.format_exc()
        log.write("FINAL STATUS: PROBE_INCOMPLETE")
        log.write(J16.error_text(exc))
        log.write(report["traceback"])
    finally:
        J16.dispose(load_status)
        J16.dispose(file_management)
        if part is not None:
            try:
                J16.close_opened_part(part, log)
                report["cleanup"] = {"opened_part_closed": True}
            except Exception as exc:
                report["cleanup"] = {
                    "opened_part_closed": False,
                    "error": exception_record(exc),
                }
        write_json(json_path, report)
        zip_path = output_root + ".zip"
        log.write("Evidence ZIP: {0}".format(zip_path))
        J16.write_log(log_path, log.lines)
        zip_path = zip_evidence(output_root)

    return zip_path


if __name__ == "__main__":
    main()
