"""
Journal 33 - CSV-Driven Teamcenter JT Export

JT companion to Journal 07. Processes every enabled DB_PART_NO + DB_PART_REV
row in NX_EXPORT_SCOPE.csv whose JT control is enabled.

For JT:
- Reuse an exact matching loaded 3D master part when available.
- Otherwise open the exact Teamcenter master revision from the CSV identity.
- Export one monolithic JT file from the active display/work part.
- Include assembly structure, precise geometry, and part/assembly PMI.
- Restore the original display/work parts and close only journal-opened parts.

Output names use <number>_REV<revision>.<WAE_VERSION>.jt. A missing
WAE_VERSION keeps the revision-only filename and records a warning.

Target: NX 2312 and NX X 2506 embedded Python
Run via: NX > Tools > Journal > Play
"""

import csv
import datetime
import os
import time
import traceback

import NXOpen


INPUT_FILENAME = "NX_EXPORT_SCOPE.csv"
OUTPUT_ROOT_FOLDER = "NX_BULK_EXPORT"
JOURNAL_BUILD_ID = "J33-NX2506-DATAPACK-JT-V2"
WAE_VERSION_ATTRIBUTE = "WAE_VERSION"
VERIFY_OUTPUT_FILES = True
CLOSE_PARTS_OPENED_BY_JOURNAL = True
JT_CONFIG_ENVIRONMENT_VARIABLE = "NX_JT_CONFIG_FILE"
JT_OUTPUT_WAIT_ENVIRONMENT_VARIABLE = "NX_JT_OUTPUT_WAIT_SECONDS"
JT_OUTPUT_WAIT_SECONDS = 120.0
JT_OUTPUT_POLL_SECONDS = 0.5
MYT_TIMEZONE = datetime.timezone(datetime.timedelta(hours=8), name="MYT")

TRUE_VALUES = {"YES", "Y", "TRUE", "1", "X"}
FALSE_VALUES = {"", "NO", "N", "FALSE", "0"}
_INVALID_FILENAME_CHARS = '<>:"/\\|?*'

_HEADER_ALIASES = {
    "part_number": ("DB_PART_NO", "Item Number", "PART_NUMBER", "Part Number"),
    "revision": ("DB_PART_REV", "Item Rev", "REVISION", "Revision"),
    "jt": ("JT", "Export_JT", "EXPORT_JT"),
    "data_pack_status": ("DATA_PACK_STATUS", "Status"),
    "primary_module": ("PRIMARY_MODULE", "Primary Module"),
    "part_description": ("PART_DESCRIPTION", "Part Description"),
    "owner": ("OWNER", "Owner"),
}

_REQUIRED_HEADERS = ("part_number", "revision", "jt")

_RESULT_COLUMNS = (
    "RUN_TIMESTAMP",
    "SOURCE_ROW_COUNT",
    "MERGED_ROW_COUNT",
    "DB_PART_NO",
    "DB_PART_REV",
    "PART_DESCRIPTION",
    "PRIMARY_MODULE",
    "DATA_PACK_STATUS",
    "OWNER",
    "JT_REQUESTED",
    "JT_RESULT",
    "JT_FILE",
    "JT_FILE_SIZE_BYTES",
    "LOADED_REVISION",
    "OVERALL_RESULT",
    "MESSAGE",
    "DURATION_SECONDS",
)


def normalize_text(value):
    return "" if value is None else str(value).strip()


def normalize_header(value):
    return " ".join(normalize_text(value).lstrip("\ufeff").split()).upper()


def clean_filename_token(value, fallback="part"):
    text = normalize_text(value)
    if not text:
        return fallback

    cleaned = "".join(
        "_" if char in _INVALID_FILENAME_CHARS or ord(char) < 32 else char
        for char in text
    ).strip(" .")
    return cleaned or fallback


def append_unique(messages, message):
    text = normalize_text(message)
    if text and text not in messages:
        messages.append(text)


def runtime_source_path():
    try:
        return os.path.abspath(__file__)
    except Exception:
        return "<unknown>"


def log_line(session, message, log_buffer=None):
    text = str(message)
    if log_buffer is not None:
        log_buffer.append(text)

    try:
        session.ListingWindow.Open()
        for line in text.splitlines() or [""]:
            session.ListingWindow.WriteFullline(line)
    except Exception:
        pass

    try:
        print(text)
    except Exception:
        pass


def dispose(value):
    if value is None:
        return
    try:
        value.Dispose()
    except Exception:
        pass


def resolve_io_root():
    configured = normalize_text(os.environ.get("NX_JOURNALS_IO_DIR"))
    if configured:
        return os.path.abspath(os.path.expandvars(os.path.expanduser(configured)))

    profile = normalize_text(os.environ.get("USERPROFILE"))
    if profile:
        return os.path.join(profile, "Desktop")

    return os.path.expanduser("~/Desktop")


def create_run_folders(io_root, timestamp):
    run_root = os.path.join(io_root, OUTPUT_ROOT_FOLDER, timestamp)
    folders = {
        "root": run_root,
        "jt": os.path.join(run_root, "JT"),
        "reports": os.path.join(run_root, "REPORTS"),
        "logs": os.path.join(run_root, "LOGS"),
    }
    for path in folders.values():
        os.makedirs(path, exist_ok=True)
    return folders


def write_text_log(path, lines):
    with open(path, "w", encoding="utf-8-sig", newline="") as handle:
        for line in lines:
            handle.write(str(line))
            handle.write("\n")


def safe_part_name(part, fallback="part"):
    for property_name in ("Name", "Leaf", "FullPath"):
        try:
            value = normalize_text(getattr(part, property_name))
            if value:
                return value
        except Exception:
            pass
    return fallback


def object_identity(nx_object):
    if nx_object is None:
        return ("NONE", "")
    try:
        return ("TAG", str(nx_object.Tag))
    except Exception:
        pass
    try:
        value = normalize_text(nx_object.FullPath)
        if value:
            return ("PATH", value.upper())
    except Exception:
        pass
    return ("OBJECT", id(nx_object))


def session_parts(session):
    try:
        return list(session.Parts)
    except Exception:
        return []


def session_part_identities(session):
    return {object_identity(part) for part in session_parts(session)}


def session_is_managed(session):
    try:
        value = session.IsManagedMode
        return bool(value() if callable(value) else value)
    except Exception:
        return False


def get_string_attribute(nx_object, attribute_name, fallback=""):
    if nx_object is None:
        return fallback
    try:
        return normalize_text(nx_object.GetStringAttribute(attribute_name))
    except Exception:
        pass
    try:
        attribute = nx_object.GetUserAttribute(
            attribute_name,
            NXOpen.NXObject.AttributeType.String,
            -1,
        )
        return normalize_text(attribute.StringValue)
    except Exception:
        return fallback


def get_part_identity(part):
    number = (
        get_string_attribute(part, "DB_PART_NO")
        or get_string_attribute(part, "PART_NUMBER")
        or get_string_attribute(part, "ITEM_ID")
    )
    revision = (
        get_string_attribute(part, "DB_PART_REV")
        or get_string_attribute(part, "REVISION")
        or get_string_attribute(part, "ITEM_REVISION")
    )
    return normalize_text(number), normalize_text(revision)


def part_identifiers(part):
    values = []
    for property_name in (
        "Name",
        "Leaf",
        "FullPath",
        "PartName",
        "JournalIdentifier",
    ):
        try:
            value = normalize_text(getattr(part, property_name))
            if value and value not in values:
                values.append(value)
        except Exception:
            pass
    return values


def part_body_count(part):
    try:
        return int(part.Bodies.Count)
    except Exception:
        pass
    try:
        return len(list(part.Bodies))
    except Exception:
        return 0


def drawing_sheet_count(part):
    try:
        return int(part.DrawingSheets.Count)
    except Exception:
        pass
    try:
        return len(list(part.DrawingSheets))
    except Exception:
        return 0


def part_has_drawing_name(part, number, revision):
    expected = "{0}-{1}-DWG".format(number.upper(), revision.upper())
    identifiers = " | ".join(part_identifiers(part)).upper()
    return expected in identifiers


def unwrap_open_result(value):
    if isinstance(value, tuple):
        part = value[0] if value else None
        status = value[1] if len(value) > 1 else None
        return part, status
    return value, None


def set_display_part(session, part):
    result = session.Parts.SetDisplay(part, False, True)
    if isinstance(result, tuple) and len(result) > 1:
        dispose(result[1])


def restore_parts(session, display_part, work_part, log_buffer):
    if display_part is not None:
        try:
            set_display_part(session, display_part)
        except Exception as error:
            log_line(
                session,
                "ERROR restoring display part: {0}".format(error),
                log_buffer,
            )
    if work_part is not None:
        try:
            session.Parts.SetWork(work_part)
        except Exception as error:
            log_line(
                session,
                "ERROR restoring work part: {0}".format(error),
                log_buffer,
            )


def close_part_best_effort(part, session, log_buffer):
    if part is None or not CLOSE_PARTS_OPENED_BY_JOURNAL:
        return
    try:
        part.Close(
            NXOpen.BasePart.CloseWholeTree.FalseValue,
            NXOpen.BasePart.CloseModified.CloseModified,
            None,
        )
    except Exception as error:
        log_line(
            session,
            "  WARNING: Could not close journal-opened part '{0}': {1}".format(
                safe_part_name(part),
                error,
            ),
            log_buffer,
        )


def open_base_part(session, specification, preloaded_identities, log_buffer):
    log_line(session, "  Attempt master open: " + specification, log_buffer)
    part = None
    status = None
    try:
        part, status = unwrap_open_result(session.Parts.OpenBase(specification))
    except Exception as error:
        log_line(session, "    Not opened: {0}".format(error), log_buffer)
        return None
    finally:
        dispose(status)

    if part is None:
        log_line(session, "    Open returned no part.", log_buffer)
        return None

    opened_by_journal = object_identity(part) not in preloaded_identities
    log_line(
        session,
        "    Opened: {0}{1}".format(
            safe_part_name(part),
            " [journal-opened]" if opened_by_journal else " [already loaded]",
        ),
        log_buffer,
    )
    return {
        "part": part,
        "opened_by_journal": opened_by_journal,
        "source": specification,
    }


def loaded_master_candidate(session, number, revision):
    matches = []
    for part in session_parts(session):
        loaded_number, loaded_revision = get_part_identity(part)
        if (
            loaded_number.upper() != number.upper()
            or loaded_revision.upper() != revision.upper()
        ):
            continue
        if part_has_drawing_name(part, number, revision):
            continue
        matches.append(part)

    if not matches:
        return None

    matches.sort(
        key=lambda part: (
            part_body_count(part) <= 0,
            drawing_sheet_count(part) > 0,
            safe_part_name(part).upper(),
        )
    )
    return matches[0]


def teamcenter_master_specs(number, revision):
    return [
        "@DB/{0}/{1}".format(number, revision),
        "@DB/{0}/{1}/master".format(number, revision),
    ]


def resolve_master_candidate(session, number, revision, log_buffer):
    loaded = loaded_master_candidate(session, number, revision)
    if loaded is not None:
        log_line(
            session,
            "  Master already loaded: " + safe_part_name(loaded),
            log_buffer,
        )
        return {
            "part": loaded,
            "opened_by_journal": False,
            "source": "loaded session",
        }, []

    preloaded_identities = session_part_identities(session)
    attempts = []
    for specification in teamcenter_master_specs(number, revision):
        attempts.append(specification)
        opened = open_base_part(
            session,
            specification,
            preloaded_identities,
            log_buffer,
        )
        if opened is not None:
            return opened, attempts
    return None, attempts


def resolve_headers(fieldnames):
    if not fieldnames:
        raise ValueError("The input CSV does not contain a header row.")

    normalized_fields = [
        (fieldname, normalize_header(fieldname)) for fieldname in fieldnames
    ]
    resolved = {}
    warnings = []
    for logical_name, aliases in _HEADER_ALIASES.items():
        matches = []
        for alias in aliases:
            wanted = normalize_header(alias)
            for original, normalized in normalized_fields:
                if normalized == wanted and original not in matches:
                    matches.append(original)
        if matches:
            resolved[logical_name] = matches[0]
            if len(matches) > 1:
                warnings.append(
                    "Multiple columns match {0}; using '{1}' and ignoring: {2}".format(
                        logical_name,
                        matches[0],
                        ", ".join(matches[1:]),
                    )
                )

    missing = [name for name in _REQUIRED_HEADERS if name not in resolved]
    if missing:
        raise ValueError(
            "Missing required logical CSV column(s): {0}".format(
                ", ".join(missing)
            )
        )
    return resolved, warnings


def row_value(row, headers, logical_name):
    fieldname = headers.get(logical_name)
    return normalize_text(row.get(fieldname, "")) if fieldname else ""


def parse_control(value, label, row_number):
    normalized = normalize_text(value).upper()
    if normalized in TRUE_VALUES:
        return True, ""
    if normalized in FALSE_VALUES:
        return False, ""
    return (
        False,
        "Source row {0}: unknown {1} control value '{2}'; treated as disabled".format(
            row_number,
            label,
            normalize_text(value),
        ),
    )


def row_is_blank(row):
    for value in row.values():
        if isinstance(value, list):
            if any(normalize_text(item) for item in value):
                return False
        elif normalize_text(value):
            return False
    return True


def read_export_scope(csv_path):
    merged = {}
    invalid_rows = []
    ignored_count = 0
    input_count = 0

    with open(csv_path, "r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        headers, header_warnings = resolve_headers(reader.fieldnames)

        for row_number, row in enumerate(reader, start=2):
            if row_is_blank(row):
                continue
            input_count += 1

            jt_requested, jt_warning = parse_control(
                row_value(row, headers, "jt"),
                "JT",
                row_number,
            )
            warnings = [warning for warning in (jt_warning,) if warning]
            if not jt_requested and not warnings:
                ignored_count += 1
                continue

            number = row_value(row, headers, "part_number")
            revision = row_value(row, headers, "revision")
            optional = {
                "part_description": row_value(row, headers, "part_description"),
                "primary_module": row_value(row, headers, "primary_module"),
                "data_pack_status": row_value(row, headers, "data_pack_status"),
                "owner": row_value(row, headers, "owner"),
            }
            errors = []
            if not number:
                errors.append("Part number is blank")
            if not revision:
                errors.append("Revision is blank")
            if not jt_requested:
                errors.append("No valid JT request remains")

            instruction = {
                "source_rows": [row_number],
                "source_row_count": 1,
                "merged_row_count": 0,
                "part_number": number,
                "revision": revision,
                "normalized_key": (number.upper(), revision.upper()),
                "jt_requested": jt_requested,
                "warnings": warnings,
                **optional,
            }

            if errors:
                instruction["warnings"].append(
                    "Source row {0}: {1}".format(row_number, "; ".join(errors))
                )
                invalid_rows.append(instruction)
                continue

            key = instruction["normalized_key"]
            existing = merged.get(key)
            if existing is None:
                merged[key] = instruction
                continue

            existing["source_rows"].append(row_number)
            existing["source_row_count"] += 1
            existing["merged_row_count"] = existing["source_row_count"] - 1
            existing["jt_requested"] = existing["jt_requested"] or jt_requested
            for name, value in optional.items():
                if value and not existing[name]:
                    existing[name] = value
            for warning in warnings:
                append_unique(existing["warnings"], warning)

    return {
        "instructions": sorted(
            merged.values(),
            key=lambda item: item["normalized_key"],
        ),
        "invalid_rows": invalid_rows,
        "ignored_row_count": ignored_count,
        "input_row_count": input_count,
        "header_warnings": header_warnings,
    }


def new_result(timestamp, instruction):
    requested = bool(instruction.get("jt_requested"))
    return {
        "RUN_TIMESTAMP": timestamp,
        "SOURCE_ROW_COUNT": instruction.get("source_row_count", 1),
        "MERGED_ROW_COUNT": instruction.get("merged_row_count", 0),
        "DB_PART_NO": instruction.get("part_number", ""),
        "DB_PART_REV": instruction.get("revision", ""),
        "PART_DESCRIPTION": instruction.get("part_description", ""),
        "PRIMARY_MODULE": instruction.get("primary_module", ""),
        "DATA_PACK_STATUS": instruction.get("data_pack_status", ""),
        "OWNER": instruction.get("owner", ""),
        "JT_REQUESTED": "YES" if requested else "NO",
        "JT_RESULT": "PENDING" if requested else "NOT_REQUESTED",
        "JT_FILE": "",
        "JT_FILE_SIZE_BYTES": "",
        "LOADED_REVISION": "",
        "OVERALL_RESULT": "PENDING",
        "MESSAGE": "",
        "DURATION_SECONDS": "",
    }


def invalid_result(timestamp, instruction):
    result = new_result(timestamp, instruction)
    if result["JT_REQUESTED"] == "YES":
        result["JT_RESULT"] = "INVALID_INPUT"
    result["OVERALL_RESULT"] = "INVALID_INPUT"
    result["MESSAGE"] = " | ".join(instruction.get("warnings", []))
    result["DURATION_SECONDS"] = "0.000"
    return result


def write_result_csv(path, results):
    with open(path, "w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=_RESULT_COLUMNS)
        writer.writeheader()
        for result in results:
            writer.writerow(
                {column: result.get(column, "") for column in _RESULT_COLUMNS}
            )


def build_versioned_base(number, revision, wae_version):
    base = "{0}_REV{1}".format(
        clean_filename_token(number),
        clean_filename_token(revision, fallback=""),
    )
    if wae_version:
        cleaned_version = clean_filename_token(wae_version, fallback="")
        if cleaned_version:
            base += "." + cleaned_version
    return base


def _add_unique_path(paths, path):
    text = normalize_text(path)
    if not text:
        return
    expanded = os.path.abspath(os.path.expandvars(os.path.expanduser(text)))
    key = os.path.normcase(os.path.normpath(expanded))
    if all(os.path.normcase(os.path.normpath(item)) != key for item in paths):
        paths.append(expanded)


def jt_config_candidates():
    """Return likely NX tessellation-config paths in priority order."""
    candidates = []
    override = normalize_text(os.environ.get(JT_CONFIG_ENVIRONMENT_VARIABLE))
    if override:
        _add_unique_path(candidates, override)
        return candidates

    for variable in ("UGII_BASE_DIR", "UGII_ROOT_DIR"):
        base = normalize_text(os.environ.get(variable))
        if not base:
            continue
        _add_unique_path(candidates, os.path.join(base, "PVTRANS", "tessUG.config"))
        _add_unique_path(candidates, os.path.join(base, "tessUG.config"))
        _add_unique_path(
            candidates,
            os.path.join(os.path.dirname(base), "PVTRANS", "tessUG.config"),
        )
        _add_unique_path(
            candidates,
            os.path.join(os.path.dirname(base), "tessUG.config"),
        )

    source_directory = os.path.dirname(runtime_source_path())
    for _unused in range(8):
        _add_unique_path(
            candidates,
            os.path.join(source_directory, "PVTRANS", "tessUG.config"),
        )
        _add_unique_path(
            candidates,
            os.path.join(source_directory, "tessUG.config"),
        )
        parent = os.path.dirname(source_directory)
        if parent == source_directory:
            break
        source_directory = parent
    return candidates


def resolve_jt_config_file():
    candidates = jt_config_candidates()
    for candidate in candidates:
        if os.path.isfile(candidate):
            return candidate

    override = normalize_text(os.environ.get(JT_CONFIG_ENVIRONMENT_VARIABLE))
    if override:
        raise RuntimeError(
            "{0} points to a missing JT config file: {1}".format(
                JT_CONFIG_ENVIRONMENT_VARIABLE,
                candidates[0] if candidates else override,
            )
        )
    raise RuntimeError(
        "NX JT config tessUG.config was not found. Checked: {0}. "
        "Set {1} to the full config-file path.".format(
            "; ".join(candidates) if candidates else "no candidate paths",
            JT_CONFIG_ENVIRONMENT_VARIABLE,
        )
    )


def jt_output_wait_seconds():
    raw = normalize_text(os.environ.get(JT_OUTPUT_WAIT_ENVIRONMENT_VARIABLE))
    if not raw:
        return JT_OUTPUT_WAIT_SECONDS
    try:
        return max(0.0, min(600.0, float(raw)))
    except ValueError:
        return JT_OUTPUT_WAIT_SECONDS


def wait_for_nonzero_file(path, timeout_seconds, poll_seconds=JT_OUTPUT_POLL_SECONDS):
    """Wait while the live JT builder finishes writing its output file."""
    started = time.monotonic()
    deadline = started + max(0.0, timeout_seconds)
    while True:
        try:
            if os.path.isfile(path):
                size = os.path.getsize(path)
                if size > 0:
                    return size, time.monotonic() - started
        except Exception:
            pass

        remaining = deadline - time.monotonic()
        if remaining <= 0:
            return None, time.monotonic() - started
        time.sleep(min(max(0.01, poll_seconds), remaining))


def configure_jt_builder(builder, output_path, config_path):
    """Apply the explicit J33 JT translation contract."""
    builder.ConfigFile = config_path
    load_config = getattr(builder, "LoadConfigSettings", None)
    if callable(load_config):
        load_config()
    builder.OutputJtFile = output_path
    builder.JtfileStructure = NXOpen.JtCreator.FileStructure.Monolithic
    builder.JtWrite = NXOpen.JtCreator.FileWrite.All
    builder.JtParts = True
    builder.AsmStructure = True
    builder.PreciseGeom = True
    builder.TessOption = NXOpen.JtCreator.TessellationOption.Nx
    builder.UseRefset = NXOpen.JtCreator.RefsetOption.Default
    builder.IncludePmi = NXOpen.JtCreator.PmiOption.PartAndAsm
    builder.ApplyPmi = True
    builder.AppendRefset = False
    builder.MergeSolids = False
    builder.MergeSheets = False
    builder.WireFrame = False


def jt_settings_summary():
    return (
        "monolithic; write=all; assembly structure=yes; precise geometry=yes; "
        "tessellation=NX; reference set=default; PMI=part+assembly"
    )


def export_jt_from_part(
    session,
    part,
    output_folder,
    number,
    revision,
    wae_version,
    log_buffer=None,
):
    set_display_part(session, part)
    session.Parts.SetWork(part)

    output_path = os.path.join(
        output_folder,
        build_versioned_base(number, revision, wae_version) + ".jt",
    )
    if os.path.exists(output_path):
        raise RuntimeError("JT output already exists: {0}".format(output_path))

    try:
        config_path = resolve_jt_config_file()
    except Exception as exc:
        return {
            "result": "FAILED_CONFIGURATION",
            "path": "",
            "size": "",
            "message": str(exc),
        }

    wait_seconds = jt_output_wait_seconds()
    log_line(session, "    JT config: {0}".format(config_path), log_buffer)
    log_line(
        session,
        "    JT output wait: up to {0:.1f} seconds".format(wait_seconds),
        log_buffer,
    )

    builder = session.PvtransManager.CreateJtCreator()
    file_size = None
    waited_seconds = 0.0
    try:
        configure_jt_builder(builder, output_path, config_path)
        validate = getattr(builder, "Validate", None)
        if callable(validate):
            valid = bool(validate())
            log_line(
                session,
                "    JT builder validation: {0}".format(valid),
                log_buffer,
            )
            if not valid:
                return {
                    "result": "FAILED_BUILDER_VALIDATION",
                    "path": "",
                    "size": "",
                    "message": (
                        "JT builder validation failed with config {0}"
                    ).format(config_path),
                }
        builder.Commit()
        file_size, waited_seconds = wait_for_nonzero_file(
            output_path,
            wait_seconds,
        )
    finally:
        builder.Destroy()

    if VERIFY_OUTPUT_FILES and file_size is None:
        return {
            "result": "FAILED_NO_OUTPUT_FILE",
            "path": "",
            "size": "",
            "message": (
                "JT builder committed with config {0}, but no nonzero output "
                "file was created within {1:.1f} seconds"
            ).format(config_path, waited_seconds),
        }

    if file_size is None:
        try:
            file_size = os.path.getsize(output_path)
        except Exception:
            file_size = ""

    if VERIFY_OUTPUT_FILES and file_size == 0:
        return {
            "result": "FAILED_ZERO_BYTE_FILE",
            "path": output_path,
            "size": file_size,
            "message": "JT output exists but is zero bytes; retained for diagnosis",
        }

    return {
        "result": "SUCCESS",
        "path": output_path,
        "size": file_size,
        "message": "",
    }


def export_jt_for_instruction(
    session,
    output_folder,
    number,
    revision,
    original_display,
    original_work,
    log_buffer,
):
    candidate, attempts = resolve_master_candidate(
        session,
        number,
        revision,
        log_buffer,
    )
    if candidate is None:
        return {
            "result": "NOT_FOUND",
            "path": "",
            "size": "",
            "message": (
                "Master part could not be loaded. See the text log for attempted "
                "@DB names."
            ),
            "attempts": attempts,
        }

    part = candidate["part"]
    try:
        loaded_number, loaded_revision = get_part_identity(part)
        if (
            loaded_number.upper() != number.upper()
            or loaded_revision.upper() != revision.upper()
        ):
            return {
                "result": "REVISION_MISMATCH",
                "path": "",
                "size": "",
                "message": (
                    "Resolved master identity {0}/{1} does not match requested "
                    "{2}/{3}"
                ).format(loaded_number, loaded_revision, number, revision),
            }

        wae_version = get_string_attribute(part, WAE_VERSION_ATTRIBUTE)
        exported = export_jt_from_part(
            session,
            part,
            output_folder,
            number,
            revision,
            wae_version,
            log_buffer,
        )
        if not wae_version:
            warning = (
                "{0} is blank or unavailable on the master; JT exported with a "
                "revision-only filename."
            ).format(WAE_VERSION_ATTRIBUTE)
            message = exported.get("message") or ""
            exported["message"] = warning if not message else message + " | " + warning
        if exported.get("result") == "SUCCESS":
            log_line(
                session,
                "    JT created: {0}".format(exported.get("path", "")),
                log_buffer,
            )
        else:
            log_line(
                session,
                "    JT rejected: {0} - {1}".format(
                    exported.get("result", "FAILED"),
                    exported.get("message", ""),
                ),
                log_buffer,
            )
        return exported
    finally:
        restore_parts(session, original_display, original_work, log_buffer)
        if candidate.get("opened_by_journal"):
            close_part_best_effort(part, session, log_buffer)


def process_instruction(
    session,
    instruction,
    folders,
    timestamp,
    original_display,
    original_work,
    log_buffer,
):
    started = datetime.datetime.now()
    result = new_result(timestamp, instruction)
    messages = list(instruction.get("warnings", []))
    number = instruction["part_number"]
    revision = instruction["revision"]

    if instruction["jt_requested"]:
        try:
            exported = export_jt_for_instruction(
                session,
                folders["jt"],
                number,
                revision,
                original_display,
                original_work,
                log_buffer,
            )
            result["JT_RESULT"] = exported["result"]
            result["JT_FILE"] = exported["path"]
            result["JT_FILE_SIZE_BYTES"] = exported.get("size", "")
            append_unique(messages, exported.get("message", ""))
        except Exception as error:
            result["JT_RESULT"] = "FAILED"
            append_unique(messages, "JT export failed: {0}".format(error))
            log_line(session, traceback.format_exc(), log_buffer)

    if result["JT_RESULT"] == "PENDING":
        result["JT_RESULT"] = "FAILED"
    result["LOADED_REVISION"] = revision
    result["OVERALL_RESULT"] = (
        "SUCCESS" if result["JT_RESULT"] == "SUCCESS" else result["JT_RESULT"]
    )
    result["MESSAGE"] = " | ".join(messages)
    result["DURATION_SECONDS"] = "{0:.3f}".format(
        (datetime.datetime.now() - started).total_seconds()
    )
    restore_parts(session, original_display, original_work, log_buffer)
    return result


def main():
    session = NXOpen.Session.GetSession()
    log_buffer = []
    folders = None
    results = []
    report_path = ""
    log_path = ""
    report_written = False
    original_display = None
    original_work = None

    try:
        try:
            original_display = session.Parts.Display
        except Exception:
            pass
        try:
            original_work = session.Parts.Work
        except Exception:
            pass

        io_root = resolve_io_root()
        input_csv = os.path.join(io_root, INPUT_FILENAME)
        log_line(session, "Journal 33 - CSV-driven JT export", log_buffer)
        log_line(session, "Journal build: " + JOURNAL_BUILD_ID, log_buffer)
        log_line(session, "Journal source: " + runtime_source_path(), log_buffer)
        log_line(session, "Input CSV: " + input_csv, log_buffer)
        log_line(session, "JT settings: " + jt_settings_summary(), log_buffer)
        log_line(
            session,
            "Master resolver: exact @DB/<part>/<rev> with /master fallback",
            log_buffer,
        )
        log_line(
            session,
            "Managed-mode flag: {0} (informational only; @DB opens are always attempted)".format(
                session_is_managed(session)
            ),
            log_buffer,
        )

        if not os.path.isfile(input_csv):
            raise FileNotFoundError("Input CSV not found: {0}".format(input_csv))

        parsed = read_export_scope(input_csv)
        run_datetime = datetime.datetime.now(MYT_TIMEZONE)
        timestamp = run_datetime.strftime("%Y%m%d_%H%M%S")
        folders = create_run_folders(io_root, timestamp)
        report_path = os.path.join(
            folders["reports"],
            "JT_EXPORT_RESULT_{0}.csv".format(timestamp),
        )
        log_path = os.path.join(
            folders["logs"],
            "JT_EXPORT_LOG_{0}.txt".format(timestamp),
        )

        for warning in parsed["header_warnings"]:
            log_line(session, "WARNING: " + warning, log_buffer)
        for invalid_instruction in parsed["invalid_rows"]:
            results.append(invalid_result(timestamp, invalid_instruction))

        instructions = parsed["instructions"]
        log_line(
            session,
            "Input rows: {0}; unique JT requests: {1}; ignored: {2}; invalid: {3}".format(
                parsed["input_row_count"],
                len(instructions),
                parsed["ignored_row_count"],
                len(parsed["invalid_rows"]),
            ),
            log_buffer,
        )

        for index, instruction in enumerate(instructions, start=1):
            log_line(
                session,
                "[{0}/{1}] {2} / {3}".format(
                    index,
                    len(instructions),
                    instruction["part_number"],
                    instruction["revision"],
                ),
                log_buffer,
            )
            result = process_instruction(
                session,
                instruction,
                folders,
                timestamp,
                original_display,
                original_work,
                log_buffer,
            )
            results.append(result)
            log_line(session, "  JT: " + result["JT_RESULT"], log_buffer)

        restore_parts(session, original_display, original_work, log_buffer)
        write_result_csv(report_path, results)
        report_written = True

        counts = {}
        for result in results:
            status = result["OVERALL_RESULT"]
            counts[status] = counts.get(status, 0) + 1

        log_line(session, "Export complete", log_buffer)
        log_line(session, "Success: {0}".format(counts.get("SUCCESS", 0)), log_buffer)
        log_line(session, "Not found: {0}".format(counts.get("NOT_FOUND", 0)), log_buffer)
        failure_count = sum(
            count
            for status, count in counts.items()
            if status not in ("SUCCESS", "NOT_FOUND")
        )
        log_line(session, "Failed: {0}".format(failure_count), log_buffer)
        log_line(
            session,
            "JT files: {0}".format(sum(1 for item in results if item["JT_FILE"])),
            log_buffer,
        )
        log_line(session, "Result report: " + report_path, log_buffer)
        write_text_log(log_path, log_buffer)

    except Exception:
        log_line(session, "ERROR: Unhandled journal exception.", log_buffer)
        log_line(session, traceback.format_exc(), log_buffer)
    finally:
        restore_parts(session, original_display, original_work, log_buffer)
        if folders is not None:
            if report_path and not report_written:
                try:
                    write_result_csv(report_path, results)
                except Exception:
                    log_line(session, traceback.format_exc(), log_buffer)
            if log_path:
                try:
                    write_text_log(log_path, log_buffer)
                except Exception:
                    pass


if __name__ == "__main__":
    main()
