"""
Journal 07 - CSV-Driven Teamcenter PDF + STEP Export

Processes every enabled DB_PART_NO + DB_PART_REV row in NX_EXPORT_SCOPE.csv.
The requested item does not need to be present in the active assembly.

For PDF:
- Reuse matching drawing specifications already loaded in the NX session.
- Otherwise try Teamcenter non-master drawing specifications dwg1..dwg9.
- Make each drawing the active display part before exporting every sheet.
- Apply DRAFT_<revision>.<WAE_VERSION> to every page using the PDF builder.
- Convert PDF text to polylines so NX catalog symbols remain visible.

For STEP:
- Reuse a matching loaded master part.
- Otherwise open the Teamcenter master directly from the CSV identity.
- Make the master the active display/work part before AP214 export.

Target: NX 2312 and NX X 2506 embedded Python
Run via: NX > Tools > Journal > Play
"""

import csv
import datetime
import os
import re
import traceback

import NXOpen


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

INPUT_FILENAME = "NX_EXPORT_SCOPE.csv"
OUTPUT_ROOT_FOLDER = "NX_BULK_EXPORT"
JOURNAL_BUILD_ID = "J07-NX2506-STEP-VALIDATION-V5"
STEP_FORMAT = "AP214"
VERIFY_OUTPUT_FILES = True
STEP_LAYER_MASK = "1-256"
PDF_DRAFT_PREFIX = "DRAFT"
WAE_VERSION_ATTRIBUTE = "WAE_VERSION"
STEP_BODY_TOKENS = (
    "MANIFOLD_SOLID_BREP",
    "BREP_WITH_VOIDS",
    "FACETED_BREP",
    "SHELL_BASED_SURFACE_MODEL",
    "CLOSED_SHELL",
    "OPEN_SHELL",
    "ADVANCED_FACE",
    "TESSELLATED_SHAPE_REPRESENTATION",
)
STEP_ASSEMBLY_TOKENS = (
    "NEXT_ASSEMBLY_USAGE_OCCURRENCE",
    "CONTEXT_DEPENDENT_SHAPE_REPRESENTATION",
    "REPRESENTATION_RELATIONSHIP_WITH_TRANSFORMATION",
    "MAPPED_ITEM",
)

MAX_DRAWING_DATASET_INDEX = 9
MAX_STEP_COMPONENT_OCCURRENCES = 20000
CLOSE_PARTS_OPENED_BY_JOURNAL = True

TRUE_VALUES = {"YES", "Y", "TRUE", "1", "X"}
FALSE_VALUES = {"", "NO", "N", "FALSE", "0"}
_INVALID_FILENAME_CHARS = '<>:"/\\|?*'

_HEADER_ALIASES = {
    "part_number": ("DB_PART_NO", "Item Number", "PART_NUMBER", "Part Number"),
    "revision": ("DB_PART_REV", "Item Rev", "REVISION", "Revision"),
    "pdf": ("PDF", "Export_PDF", "EXPORT_PDF"),
    "step": ("STEP", "Export_STEP", "EXPORT_STEP"),
    "data_pack_status": ("DATA_PACK_STATUS", "Status"),
    "primary_module": ("PRIMARY_MODULE", "Primary Module"),
    "part_description": ("PART_DESCRIPTION", "Part Description"),
    "owner": ("OWNER", "Owner"),
}

_REQUIRED_HEADERS = ("part_number", "revision", "pdf", "step")

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
    "PDF_REQUESTED",
    "PDF_RESULT",
    "PDF_FILE_COUNT",
    "PDF_FILES",
    "STEP_REQUESTED",
    "STEP_RESULT",
    "STEP_FILE",
    "LOADED_REVISION",
    "OVERALL_RESULT",
    "MESSAGE",
    "DURATION_SECONDS",
)

_DRAWING_SUFFIX_RE = re.compile(
    r"(?:^|[-_])DWG(\d+)(?:$|[^A-Z0-9])",
    re.IGNORECASE,
)


# ---------------------------------------------------------------------------
# Generic helpers
# ---------------------------------------------------------------------------

def normalize_text(value):
    return "" if value is None else str(value).strip()


def runtime_source_path():
    try:
        return os.path.abspath(__file__)
    except Exception:
        return "<unknown>"


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


def log_line(session, message, log_buffer=None):
    text = str(message)
    if log_buffer is not None:
        log_buffer.append(text)

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


def write_text_log(path, lines):
    with open(path, "w", encoding="utf-8", newline="") as handle:
        for line in lines:
            handle.write(str(line) + "\n")


def desktop_folder():
    profile = normalize_text(os.environ.get("USERPROFILE"))
    if profile:
        return os.path.join(profile, "Desktop")

    home = os.path.expanduser("~")
    if home and home != "~":
        return os.path.join(home, "Desktop")

    return os.getcwd()


def resolve_io_root():
    configured = normalize_text(os.environ.get("NX_JOURNALS_IO_DIR"))
    return os.path.abspath(os.path.expanduser(configured or desktop_folder()))


def create_run_folders(io_root, timestamp):
    run = os.path.join(io_root, OUTPUT_ROOT_FOLDER, timestamp)
    os.makedirs(run, exist_ok=False)

    folders = {"run": run}
    for name in ("PDF", "STEP", "REPORTS", "LOGS"):
        path = os.path.join(run, name)
        os.makedirs(path, exist_ok=False)
        folders[name.lower()] = path
    return folders


def dispose(value):
    if value is None:
        return
    try:
        value.Dispose()
    except Exception:
        pass


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
    """
    Informational only.

    Some NX X / TeamcenterX sessions have returned False even though @DB
    managed-mode part names are valid. The journal therefore never uses this
    value to decide whether an @DB open should be attempted.
    """
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

    for attribute_name in (
        "DB_PART_NO",
        "DB_PART_REV",
        "DB_PART_NAME",
        "DB_DATASET_NAME",
        "DB_FULL_NAME",
        "DB_MODEL_NAME",
        "PART_NUMBER",
        "REVISION",
    ):
        value = get_string_attribute(part, attribute_name)
        if value and value not in values:
            values.append(value)

    return values


def drawing_sheet_count(part):
    try:
        return int(part.DrawingSheets.Count)
    except Exception:
        pass

    try:
        return len(list(part.DrawingSheets))
    except Exception:
        return 0


def part_body_count(part):
    try:
        return int(part.Bodies.Count)
    except Exception:
        pass

    try:
        return len(list(part.Bodies))
    except Exception:
        return 0


def step_collection_values(collection):
    if collection is None:
        return []
    try:
        return list(collection)
    except Exception:
        try:
            return list(collection.ToArray())
        except Exception:
            return []


def step_safe_property(value, name, fallback=None):
    if value is None:
        return fallback
    try:
        result = getattr(value, name)
        if callable(result):
            result = result()
        return result
    except Exception:
        return fallback


def runtime_nx_version(session):
    names = (
        "UGII_VERSION",
        "UGII_FULL_VERSION",
        "UGII_PRODUCT_VERSION",
        "UGII_MAJOR_VERSION",
    )
    try:
        getter = getattr(session, "GetEnvironmentVariableValue")
    except Exception:
        getter = None
    if callable(getter):
        for name in names:
            try:
                value = normalize_text(getter(name))
            except Exception:
                value = ""
            if value:
                return value

    for name in names:
        value = normalize_text(os.environ.get(name))
        if value:
            return value
    return "UNKNOWN"


def part_fully_loaded_state(part):
    for name in ("IsFullyLoaded", "FullyLoaded"):
        value = step_safe_property(part, name, None)
        if value is not None:
            try:
                return bool(value)
            except Exception:
                pass
    return ""


def step_body_type_counts(part):
    result = {
        "body_count": 0,
        "solid_body_count": 0,
        "sheet_body_count": 0,
        "unknown_body_type_count": 0,
    }
    bodies = step_collection_values(
        step_safe_property(part, "Bodies", None)
    )
    result["body_count"] = len(bodies)
    for body in bodies:
        is_solid = step_safe_property(body, "IsSolidBody", None)
        if is_solid is None:
            result["unknown_body_type_count"] += 1
        elif bool(is_solid):
            result["solid_body_count"] += 1
        else:
            result["sheet_body_count"] += 1
    return result


def step_component_children(component):
    try:
        return list(component.GetChildren())
    except Exception:
        return []


def step_source_geometry_snapshot(part):
    direct = step_body_type_counts(part)
    result = {
        "fully_loaded": part_fully_loaded_state(part),
        "direct_body_count": direct["body_count"],
        "direct_solid_body_count": direct["solid_body_count"],
        "direct_sheet_body_count": direct["sheet_body_count"],
        "component_occurrence_count": 0,
        "suppressed_occurrence_count": 0,
        "unique_prototype_count": 0,
        "descendant_body_occurrence_count": 0,
        "descendant_solid_body_occurrence_count": 0,
        "descendant_sheet_body_occurrence_count": 0,
        "component_limit_reached": False,
    }

    assembly = step_safe_property(part, "ComponentAssembly", None)
    root = step_safe_property(assembly, "RootComponent", None)
    if root is None:
        return result

    stack = list(reversed(step_component_children(root)))
    seen_occurrences = set()
    seen_prototypes = set()
    while stack:
        if (
            result["component_occurrence_count"]
            >= MAX_STEP_COMPONENT_OCCURRENCES
        ):
            result["component_limit_reached"] = True
            break

        component = stack.pop()
        key = object_identity(component)
        if key in seen_occurrences:
            continue
        seen_occurrences.add(key)

        if bool(step_safe_property(component, "IsSuppressed", False)):
            result["suppressed_occurrence_count"] += 1
            continue

        result["component_occurrence_count"] += 1
        prototype = step_safe_property(component, "Prototype", None)
        if prototype is not None:
            seen_prototypes.add(object_identity(prototype))
            counts = step_body_type_counts(prototype)
            result["descendant_body_occurrence_count"] += counts[
                "body_count"
            ]
            result["descendant_solid_body_occurrence_count"] += counts[
                "solid_body_count"
            ]
            result["descendant_sheet_body_occurrence_count"] += counts[
                "sheet_body_count"
            ]

        children = step_component_children(component)
        if children:
            stack.extend(reversed(children))

    result["unique_prototype_count"] = len(seen_prototypes)
    return result


def step_load_status_messages(status):
    messages = []
    if status is None:
        return messages
    try:
        count = int(status.NumberUnloadedParts)
    except Exception:
        count = 0
    for index in range(count):
        try:
            messages.append(
                "{0}: {1}".format(
                    normalize_text(status.GetPartName(index)),
                    normalize_text(status.GetStatusDescription(index)),
                )
            )
        except Exception:
            pass
    return messages


def ensure_step_source_loaded(session, part):
    result = {
        "status": "NOT_AVAILABLE",
        "method": "none",
        "messages": [],
    }
    parts = step_safe_property(session, "Parts", None)
    try:
        ensure = getattr(parts, "EnsurePartsLoadedFully")
    except Exception:
        ensure = None
    if callable(ensure):
        status = None
        result["method"] = "EnsurePartsLoadedFully(includeChildren=True)"
        try:
            status = ensure([part], True)
            result["status"] = "SUCCESS"
            result["messages"] = step_load_status_messages(status)
            return result
        except Exception as error:
            result["status"] = "FAILED"
            result["messages"].append(normalize_text(error))
        finally:
            dispose(status)

    try:
        load_fully = getattr(part, "LoadFully")
    except Exception:
        load_fully = None
    if callable(load_fully):
        status = None
        result["method"] = (
            result["method"] + " -> part.LoadFully()"
            if result["method"] != "none"
            else "part.LoadFully()"
        )
        try:
            status = load_fully()
            result["status"] = "SUCCESS_FALLBACK"
            result["messages"].extend(
                step_load_status_messages(status)
            )
        except Exception as error:
            result["status"] = "FAILED"
            result["messages"].append(normalize_text(error))
        finally:
            dispose(status)
    elif part_fully_loaded_state(part) is True:
        result["status"] = "ALREADY_FULLY_LOADED"
    return result


def unwrap_open_result(value):
    """
    NXOpen Python may return either a part or (part, PartLoadStatus).
    """
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


def open_base_part(session, specification, preloaded_identities, log_buffer, label):
    """
    Open one NX/Teamcenter part name. The @DB attempt is made regardless of the
    IsManagedMode value.
    """
    log_line(
        session,
        "  Attempt {0} open: {1}".format(label, specification),
        log_buffer,
    )

    part = None
    status = None
    try:
        part, status = unwrap_open_result(session.Parts.OpenBase(specification))
    except Exception as error:
        log_line(
            session,
            "    Not opened: {0}".format(error),
            log_buffer,
        )
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


def open_display_part(
    session,
    specification,
    preloaded_identities,
    log_buffer,
    label,
):
    """
    Open a Teamcenter specification as the display part.

    Teamcenter drawing datasets use their canonical /specification/
    JournalIdentifier. OpenDisplay is required for NX to resolve that managed
    drawing identity and make its drawing sheets available.
    """
    log_line(
        session,
        "  Attempt {0} open: {1}".format(label, specification),
        log_buffer,
    )

    part = None
    status = None
    try:
        part, status = unwrap_open_result(
            session.Parts.OpenDisplay(specification)
        )
    except Exception as error:
        log_line(
            session,
            "    Not opened: {0}".format(error),
            log_buffer,
        )
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
    log_line(
        session,
        "    JournalIdentifier: {0}".format(
            normalize_text(getattr(part, "JournalIdentifier", ""))
        ),
        log_buffer,
    )
    log_line(
        session,
        "    Drawing sheets: {0}".format(drawing_sheet_count(part)),
        log_buffer,
    )

    return {
        "part": part,
        "opened_by_journal": opened_by_journal,
        "source": specification,
    }


# ---------------------------------------------------------------------------
# CSV input and reports
# ---------------------------------------------------------------------------

def resolve_headers(fieldnames):
    if not fieldnames:
        raise ValueError("The input CSV does not contain a header row.")

    normalized_fields = [
        (fieldname, normalize_header(fieldname))
        for fieldname in fieldnames
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
            pdf_requested, pdf_warning = parse_control(
                row_value(row, headers, "pdf"),
                "PDF",
                row_number,
            )
            step_requested, step_warning = parse_control(
                row_value(row, headers, "step"),
                "STEP",
                row_number,
            )
            warnings = [
                warning
                for warning in (pdf_warning, step_warning)
                if warning
            ]

            if not pdf_requested and not step_requested and not warnings:
                ignored_count += 1
                continue

            number = row_value(row, headers, "part_number")
            revision = row_value(row, headers, "revision")
            optional = {
                "part_description": row_value(
                    row,
                    headers,
                    "part_description",
                ),
                "primary_module": row_value(
                    row,
                    headers,
                    "primary_module",
                ),
                "data_pack_status": row_value(
                    row,
                    headers,
                    "data_pack_status",
                ),
                "owner": row_value(row, headers, "owner"),
            }

            errors = []
            if not number:
                errors.append("Part number is blank")
            if not revision:
                errors.append("Revision is blank")
            if not pdf_requested and not step_requested:
                errors.append("No valid PDF or STEP request remains")

            instruction = {
                "source_rows": [row_number],
                "source_row_count": 1,
                "merged_row_count": 0,
                "part_number": number,
                "revision": revision,
                "normalized_key": (number.upper(), revision.upper()),
                "pdf_requested": pdf_requested,
                "step_requested": step_requested,
                "warnings": warnings,
                **optional,
            }

            if errors:
                instruction["warnings"].append(
                    "Source row {0}: {1}".format(
                        row_number,
                        "; ".join(errors),
                    )
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
            existing["pdf_requested"] = (
                existing["pdf_requested"] or pdf_requested
            )
            existing["step_requested"] = (
                existing["step_requested"] or step_requested
            )

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
    pdf_requested = bool(instruction.get("pdf_requested"))
    step_requested = bool(instruction.get("step_requested"))

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
        "PDF_REQUESTED": "YES" if pdf_requested else "NO",
        "PDF_RESULT": "PENDING" if pdf_requested else "NOT_REQUESTED",
        "PDF_FILE_COUNT": 0,
        "PDF_FILES": "",
        "STEP_REQUESTED": "YES" if step_requested else "NO",
        "STEP_RESULT": "PENDING" if step_requested else "NOT_REQUESTED",
        "STEP_FILE": "",
        "LOADED_REVISION": "",
        "OVERALL_RESULT": "PENDING",
        "MESSAGE": "",
        "DURATION_SECONDS": "",
    }


def invalid_result(timestamp, instruction):
    result = new_result(timestamp, instruction)

    if result["PDF_REQUESTED"] == "YES":
        result["PDF_RESULT"] = "INVALID_INPUT"
    if result["STEP_REQUESTED"] == "YES":
        result["STEP_RESULT"] = "INVALID_INPUT"

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


# ---------------------------------------------------------------------------
# Drawing discovery and PDF export
# ---------------------------------------------------------------------------

def drawing_index_from_text(text):
    match = _DRAWING_SUFFIX_RE.search(normalize_text(text).upper())
    if not match:
        return None

    try:
        return int(match.group(1))
    except Exception:
        return None


def drawing_index_from_part(part):
    for identifier in part_identifiers(part):
        index = drawing_index_from_text(identifier)
        if index is not None:
            return index
    return None


def loaded_drawing_candidates(session, number, revision):
    """
    Find matching drawing parts already loaded anywhere in the NX session.

    The part does not need to be active. A candidate is accepted when:
    - its part/revision attributes match and it owns drawing sheets, or
    - its identifiers contain the expected item-revision-dwg pattern.
    """
    expected_prefix = "{0}-{1}-DWG".format(
        number.upper(),
        revision.upper(),
    )
    candidates = []
    seen = set()

    try:
        display_part = session.Parts.Display
    except Exception:
        display_part = None

    ordered_parts = []
    if display_part is not None:
        ordered_parts.append(display_part)
    ordered_parts.extend(session_parts(session))

    for part in ordered_parts:
        identity = object_identity(part)
        if part is None or identity in seen:
            continue
        seen.add(identity)

        identifiers = " | ".join(part_identifiers(part)).upper()
        loaded_number, loaded_revision = get_part_identity(part)
        exact_identity = (
            loaded_number.upper() == number.upper()
            and loaded_revision.upper() == revision.upper()
        )
        name_match = expected_prefix in identifiers
        sheet_match = drawing_sheet_count(part) > 0

        if name_match or (exact_identity and sheet_match):
            candidates.append(
                {
                    "part": part,
                    "opened_by_journal": False,
                    "source": "loaded session",
                    "drawing_index": drawing_index_from_part(part),
                }
            )

    candidates.sort(
        key=lambda item: (
            item["drawing_index"] is None,
            item["drawing_index"] or 9999,
            safe_part_name(item["part"]).upper(),
        )
    )
    return candidates


def teamcenter_drawing_specs(number, revision, index):
    dataset_name = "{0}-{1}-dwg{2}".format(
        number,
        revision,
        index,
    )

    return [
        "@DB/{0}/{1}/specification/{2}".format(
            number,
            revision,
            dataset_name,
        ),
    ]


def resolve_drawing_candidates(session, number, revision, log_buffer):
    """
    Resolve loaded and not-yet-loaded dwg1..dwgN specifications.

    This function deliberately attempts @DB opens even when
    Session.IsManagedMode reports False.
    """
    preloaded_identities = session_part_identities(session)
    resolved = loaded_drawing_candidates(session, number, revision)
    seen = {object_identity(item["part"]) for item in resolved}
    known_indices = {
        item["drawing_index"]
        for item in resolved
        if item["drawing_index"] is not None
    }
    attempts = []

    for item in resolved:
        log_line(
            session,
            "  Drawing already loaded: {0}".format(
                safe_part_name(item["part"])
            ),
            log_buffer,
        )

    for index in range(1, MAX_DRAWING_DATASET_INDEX + 1):
        if index in known_indices:
            continue

        opened_for_index = False
        for specification in teamcenter_drawing_specs(
            number,
            revision,
            index,
        ):
            attempts.append(specification)
            opened = open_display_part(
                session,
                specification,
                preloaded_identities,
                log_buffer,
                "drawing",
            )
            if opened is None:
                continue

            part = opened["part"]
            identity = object_identity(part)

            if identity in seen:
                known_indices.add(index)
                opened_for_index = True
                break

            opened["drawing_index"] = index
            resolved.append(opened)
            seen.add(identity)
            known_indices.add(index)
            opened_for_index = True
            break

        if not opened_for_index:
            log_line(
                session,
                "  No drawing opened for DWG{0}.".format(index),
                log_buffer,
            )

    resolved.sort(
        key=lambda item: (
            item["drawing_index"] is None,
            item["drawing_index"] or 9999,
            safe_part_name(item["part"]).upper(),
        )
    )
    return resolved, attempts


def drawing_token(part, index):
    if index is not None:
        return "DWG{0}".format(index)

    detected = drawing_index_from_part(part)
    if detected is not None:
        return "DWG{0}".format(detected)

    return "DRAWING"


def unique_drawing_tokens(candidates):
    preferred = [
        drawing_token(
            candidate["part"],
            candidate.get("drawing_index"),
        )
        for candidate in candidates
    ]
    reserved = {
        token.upper()
        for token in preferred
        if token.upper() != "DRAWING"
    }
    used = set()
    result = []
    next_index = 1

    for token in preferred:
        key = token.upper()
        if key != "DRAWING" and key not in used:
            resolved = token
        else:
            while True:
                resolved = "DWG{0}".format(next_index)
                next_index += 1
                resolved_key = resolved.upper()
                if resolved_key not in reserved and resolved_key not in used:
                    break
        used.add(resolved.upper())
        result.append(resolved)

    return result


def build_pdf_filename(
    number,
    revision,
    token,
    drawing_count,
):
    filename = "{0}_REV{1}".format(
        clean_filename_token(number),
        clean_filename_token(revision, fallback=""),
    )

    if drawing_count > 1:
        filename += "_" + clean_filename_token(token)

    return filename + ".pdf"


def build_pdf_watermark(revision, wae_version):
    watermark = PDF_DRAFT_PREFIX
    cleaned_revision = normalize_text(revision)
    cleaned_wae_version = normalize_text(wae_version)

    if cleaned_revision:
        watermark += "_" + cleaned_revision
    if cleaned_wae_version:
        watermark += "." + cleaned_wae_version

    return watermark


def resolve_pdf_watermark(session, number, revision, candidates):
    model_part = loaded_master_candidate(session, number, revision)
    if model_part is not None:
        wae_version = get_string_attribute(
            model_part,
            WAE_VERSION_ATTRIBUTE,
        )
        if wae_version:
            return (
                build_pdf_watermark(revision, wae_version),
                "loaded model {0}".format(WAE_VERSION_ATTRIBUTE),
                "",
            )

    for candidate in candidates:
        drawing_part = candidate.get("part")
        wae_version = get_string_attribute(
            drawing_part,
            WAE_VERSION_ATTRIBUTE,
        )
        if wae_version:
            return (
                build_pdf_watermark(revision, wae_version),
                "drawing {0}".format(WAE_VERSION_ATTRIBUTE),
                "",
            )

    watermark = build_pdf_watermark(revision, "")
    warning = (
        "{0} is blank or unavailable; PDF exported with revision-only "
        "watermark {1}."
    ).format(WAE_VERSION_ATTRIBUTE, watermark)
    return watermark, "revision-only fallback", warning


def export_drawing_pdf(
    drawing_part,
    sheets,
    output_path,
    watermark,
):
    if not sheets:
        raise RuntimeError("Drawing contains no sheets to export.")
    if not normalize_text(watermark):
        raise RuntimeError("A PDF draft watermark is required.")

    sheets[0].Open()
    builder = drawing_part.PlotManager.CreatePrintPdfbuilder()
    try:
        builder.Action = NXOpen.PrintPDFBuilder.ActionOption.Native
        builder.Filename = output_path
        builder.Append = False
        try:
            builder.OutputText = (
                NXOpen.PrintPDFBuilder.OutputTextOption.Polylines
            )
        except Exception as error:
            raise RuntimeError(
                "NX PrintPDFBuilder could not apply required polyline "
                "text output: {0}".format(error)
            )
        try:
            builder.AddWatermark = True
            builder.Watermark = watermark
        except Exception as error:
            raise RuntimeError(
                "NX PrintPDFBuilder could not apply required watermark "
                "{0}: {1}".format(watermark, error)
            )
        builder.SourceBuilder.SetSheets(sheets)
        builder.Commit()
    finally:
        builder.Destroy()


def export_pdfs_for_instruction(
    session,
    output_folder,
    number,
    revision,
    original_display,
    original_work,
    log_buffer,
):
    candidates, attempts = resolve_drawing_candidates(
        session,
        number,
        revision,
        log_buffer,
    )

    if not candidates:
        return {
            "result": "SKIPPED_NO_DRAWING",
            "paths": [],
            "message": (
                "No drawing specification could be opened. "
                "See the text log for attempted @DB names."
            ),
            "failures": [],
        }

    successful_paths = []
    failures = []
    messages = []
    drawing_count = len(candidates)
    output_tokens = unique_drawing_tokens(candidates)
    watermark, watermark_source, watermark_warning = (
        resolve_pdf_watermark(
            session,
            number,
            revision,
            candidates,
        )
    )
    log_line(
        session,
        "  PDF watermark: {0} ({1})".format(
            watermark,
            watermark_source,
        ),
        log_buffer,
    )
    if watermark_warning:
        messages.append(watermark_warning)
        log_line(
            session,
            "  WARNING: " + watermark_warning,
            log_buffer,
        )

    try:
        for candidate, token in zip(candidates, output_tokens):
            drawing_part = candidate["part"]

            try:
                set_display_part(session, drawing_part)
            except Exception as error:
                failures.append(
                    {
                        "kind": "ERROR",
                        "message": "{0}: could not activate drawing: {1}".format(
                            token,
                            error,
                        ),
                        "traceback": traceback.format_exc(),
                    }
                )
                continue

            try:
                sheets = list(drawing_part.DrawingSheets)
            except Exception as error:
                failures.append(
                    {
                        "kind": "ERROR",
                        "message": "{0}: could not enumerate drawing sheets: {1}".format(
                            token,
                            error,
                        ),
                        "traceback": traceback.format_exc(),
                    }
                )
                continue

            if not sheets:
                failures.append(
                    {
                        "kind": "NOT_DRAWING",
                        "message": "{0}: opened part contains no drawing sheets".format(
                            token
                        ),
                    }
                )
                continue

            log_line(
                session,
                "  Exporting {0}: {1} sheet(s) from {2}".format(
                    token,
                    len(sheets),
                    safe_part_name(drawing_part),
                ),
                log_buffer,
            )

            filename = build_pdf_filename(
                number,
                revision,
                token,
                drawing_count,
            )
            output_path = os.path.join(output_folder, filename)

            try:
                if os.path.exists(output_path):
                    raise RuntimeError(
                        "PDF output already exists: {0}".format(
                            output_path
                        )
                    )

                export_drawing_pdf(
                    drawing_part,
                    sheets,
                    output_path,
                    watermark,
                )

                if VERIFY_OUTPUT_FILES and not os.path.isfile(output_path):
                    failures.append(
                        {
                            "kind": "NO_OUTPUT",
                            "message": (
                                "{0}: PDF builder committed but no "
                                "combined file was created"
                            ).format(token),
                        }
                    )
                else:
                    successful_paths.append(output_path)
                    log_line(
                        session,
                        "    PDF created: {0} ({1} sheets)".format(
                            output_path,
                            len(sheets),
                        ),
                        log_buffer,
                    )
            except Exception as error:
                failures.append(
                    {
                        "kind": "ERROR",
                        "message": "{0}: {1}".format(
                            token,
                            error,
                        ),
                        "traceback": traceback.format_exc(),
                    }
                )
    finally:
        restore_parts(
            session,
            original_display,
            original_work,
            log_buffer,
        )

        for candidate in candidates:
            if candidate.get("opened_by_journal"):
                close_part_best_effort(
                    candidate["part"],
                    session,
                    log_buffer,
                )

    if successful_paths and not failures:
        result = "SUCCESS"
    elif successful_paths:
        result = "PARTIAL_SUCCESS"
    elif failures and all(
        failure["kind"] == "NO_OUTPUT"
        for failure in failures
    ):
        result = "FAILED_NO_OUTPUT_FILE"
    elif failures and all(
        failure["kind"] == "NOT_DRAWING"
        for failure in failures
    ):
        result = "SKIPPED_NO_DRAWING"
    else:
        result = "FAILED"

    return {
        "result": result,
        "paths": successful_paths,
        "message": " | ".join(
            messages
            + [
                failure["message"]
                for failure in failures
            ]
        ),
        "failures": failures,
        "attempts": attempts,
        "watermark": watermark,
    }


# ---------------------------------------------------------------------------
# Master-part discovery and STEP export
# ---------------------------------------------------------------------------

def part_has_drawing_name(part, number, revision):
    expected = "{0}-{1}-DWG".format(
        number.upper(),
        revision.upper(),
    )
    identifiers = " | ".join(part_identifiers(part)).upper()
    return expected in identifiers


def loaded_master_candidate(session, number, revision):
    """
    Prefer an exact loaded master model. Drawing non-master parts are excluded.
    """
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
            "  Master already loaded: {0}".format(
                safe_part_name(loaded)
            ),
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
            "master",
        )
        if opened is not None:
            return opened, attempts

    return None, attempts


def inspect_step_file(path):
    result = {
        "data_entity_lines": 0,
        "body_geometry_signatures": 0,
        "assembly_signatures": 0,
    }
    in_data = False
    with open(path, "r", encoding="utf-8", errors="replace") as handle:
        for line in handle:
            upper = line.upper()
            stripped = upper.strip()
            if stripped == "DATA;":
                in_data = True
                continue
            if in_data and stripped == "ENDSEC;":
                in_data = False
            if not in_data:
                continue
            if stripped.startswith("#"):
                result["data_entity_lines"] += 1
            for token in STEP_BODY_TOKENS:
                result["body_geometry_signatures"] += upper.count(token)
            for token in STEP_ASSEMBLY_TOKENS:
                result["assembly_signatures"] += upper.count(token)
    return result


def step_body_signature_count(path):
    return inspect_step_file(path)["body_geometry_signatures"]


def parse_step_translator_log(output_path):
    result = {
        "path": "",
        "solids_input": "",
        "solids_processed": "",
        "solids_as_sheets": "",
        "solids_not_processed": "",
        "ug_body_count": "",
    }
    folder = os.path.dirname(output_path)
    output_name = os.path.basename(output_path).upper()
    output_stem = os.path.splitext(output_name)[0]
    candidates = []
    try:
        names = os.listdir(folder)
    except Exception:
        return result

    for name in names:
        if not name.lower().endswith(".log"):
            continue
        path = os.path.join(folder, name)
        try:
            with open(
                path,
                "r",
                encoding="utf-8",
                errors="replace",
            ) as handle:
                content = handle.read()
        except Exception:
            continue
        upper = content.upper()
        if (
            output_name not in upper
            and output_stem not in name.upper()
        ):
            continue
        try:
            modified = os.path.getmtime(path)
        except Exception:
            modified = 0
        candidates.append((modified, path, content))

    candidates.sort(reverse=True)
    patterns = {
        "solids_input": (
            r"Total\s+number\s+of\s+solids\s+input.*?:\s*(\d+)"
        ),
        "solids_processed": (
            r"Number\s+of\s+solids\s+processed\s+without\s+"
            r"problems.*?:\s*(\d+)"
        ),
        "solids_as_sheets": (
            r"Number\s+of\s+solids.*?(?:output|processed)\s+as\s+"
            r"sheets.*?:\s*(\d+)"
        ),
        "solids_not_processed": (
            r"Number\s+of\s+solids\s+not\s+processed.*?:\s*(\d+)"
        ),
    }
    for _modified, path, content in candidates:
        matched = False
        for key, pattern in patterns.items():
            match = re.search(pattern, content, re.IGNORECASE)
            if match:
                result[key] = match.group(1)
                matched = True
        body_matches = re.findall(
            r"SUMMARY-\s*Body\s*\.*\s*:\s*(\d+)",
            content,
            re.IGNORECASE,
        )
        if body_matches:
            result["ug_body_count"] = str(
                sum(int(value) for value in body_matches)
            )
            matched = True
        if matched:
            result["path"] = path
            return result
    return result


def step_metric_int(value):
    try:
        return int(value)
    except Exception:
        return None


def evaluate_step_validation(
    nx_version,
    load_result,
    source,
    inspection,
    translator,
    file_size,
):
    warnings = []
    failure = ""
    body_signatures = int(
        inspection.get("body_geometry_signatures", 0) or 0
    )
    assembly_signatures = int(
        inspection.get("assembly_signatures", 0) or 0
    )
    solids_input = step_metric_int(translator.get("solids_input"))
    solids_processed = step_metric_int(
        translator.get("solids_processed")
    )
    solids_as_sheets = step_metric_int(
        translator.get("solids_as_sheets")
    )
    solids_not_processed = step_metric_int(
        translator.get("solids_not_processed")
    )

    if body_signatures <= 0:
        failure = "FAILED_ZERO_GEOMETRY"
    elif translator.get("path") and solids_input == 0:
        failure = "FAILED_TRANSLATOR_ZERO_GEOMETRY"
    elif (
        solids_not_processed is not None
        and solids_not_processed > 0
    ):
        failure = "FAILED_TRANSLATOR_PARTIAL_GEOMETRY"
    elif (
        solids_input is not None
        and solids_processed is not None
        and (
            solids_processed + (solids_as_sheets or 0)
            < solids_input
        )
    ):
        failure = "FAILED_TRANSLATOR_PARTIAL_GEOMETRY"

    load_status = load_result.get("status", "UNKNOWN")
    if load_status == "FAILED":
        warnings.append("full-load request failed")
    if source.get("component_limit_reached"):
        warnings.append("component diagnostic limit reached")
    if not translator.get("path"):
        warnings.append("matching NX STEP translator log not found")

    expected_solids = int(
        source.get("direct_solid_body_count", 0) or 0
    ) + int(
        source.get("descendant_solid_body_occurrence_count", 0)
        or 0
    )
    if (
        solids_input is not None
        and expected_solids > solids_input
    ):
        warnings.append(
            "source solid count exceeds translator input "
            "({0}>{1})".format(expected_solids, solids_input)
        )
    if (
        int(source.get("component_occurrence_count", 0) or 0) > 0
        and assembly_signatures <= 0
    ):
        warnings.append(
            "source is an assembly but STEP has no assembly signatures"
        )
    if solids_as_sheets:
        warnings.append(
            "{0} solid(s) were output as sheets".format(
                solids_as_sheets
            )
        )

    validation = (
        "FAILED"
        if failure
        else "VERIFIED_WITH_WARNINGS"
        if warnings
        else "VERIFIED"
    )
    result = failure or "SUCCESS"
    translator_text = (
        "input={0}, processed={1}, as_sheets={2}, not_processed={3}".format(
            translator.get("solids_input", "N/A") or "N/A",
            translator.get("solids_processed", "N/A") or "N/A",
            translator.get("solids_as_sheets", "N/A") or "N/A",
            translator.get("solids_not_processed", "N/A") or "N/A",
        )
    )
    message = (
        "STEP validation: {0}; NX={1}; load={2}; "
        "source direct solids={3}, components={4}, "
        "descendant solids={5}; STEP size={6} bytes, "
        "body signatures={7}, assembly signatures={8}; "
        "translator {9}"
    ).format(
        validation,
        nx_version,
        load_status,
        source.get("direct_solid_body_count", 0),
        source.get("component_occurrence_count", 0),
        source.get("descendant_solid_body_occurrence_count", 0),
        file_size,
        body_signatures,
        assembly_signatures,
        translator_text,
    )
    if warnings:
        message += "; warnings: " + ", ".join(warnings)
    return {
        "result": result,
        "validation": validation,
        "message": message,
        "warnings": warnings,
    }


def export_step_from_part(
    session,
    part,
    output_folder,
    number,
    revision,
    log_buffer=None,
):
    set_display_part(session, part)
    session.Parts.SetWork(part)

    nx_version = runtime_nx_version(session)
    load_result = ensure_step_source_loaded(session, part)
    source = step_source_geometry_snapshot(part)
    log_line(
        session,
        (
            "    STEP source: NX={0}; load={1} via {2}; "
            "fully_loaded={3}; direct solids={4}; components={5}; "
            "descendant solids={6}"
        ).format(
            nx_version,
            load_result.get("status", "UNKNOWN"),
            load_result.get("method", "none"),
            source.get("fully_loaded", ""),
            source.get("direct_solid_body_count", 0),
            source.get("component_occurrence_count", 0),
            source.get("descendant_solid_body_occurrence_count", 0),
        ),
        log_buffer,
    )
    for message in load_result.get("messages", []):
        log_line(
            session,
            "      Load detail: {0}".format(message),
            log_buffer,
        )

    output_path = os.path.join(
        output_folder,
        "{0}_REV{1}.stp".format(
            clean_filename_token(number),
            clean_filename_token(revision, fallback=""),
        ),
    )

    if os.path.exists(output_path):
        raise RuntimeError(
            "STEP output already exists: {0}".format(output_path)
        )

    exporter = session.DexManager.CreateStepCreator()
    try:
        exporter.OutputFile = output_path
        # Journal 10 proved this exact display/scope/layer combination:
        # the gasket changes from zero input solids to one processed solid.
        exporter.ExportFrom = (
            NXOpen.StepCreator.ExportFromOption.DisplayPart
        )
        exporter.ExportSelectionBlock.SelectionScope = (
            NXOpen.ObjectSelector.Scope.EntirePart
        )
        exporter.LayerMask = STEP_LAYER_MASK
        exporter.ObjectTypes.Solids = True
        exporter.ObjectTypes.Surfaces = True
        exporter.ObjectTypes.Curves = True

        if STEP_FORMAT != "AP214":
            raise RuntimeError(
                "Unsupported STEP_FORMAT: {0}".format(STEP_FORMAT)
            )

        exporter.ExportAs = NXOpen.StepCreator.ExportAsOption.Ap214
        exporter.ProcessHoldFlag = True
        exporter.Commit()
    finally:
        exporter.Destroy()

    if VERIFY_OUTPUT_FILES and not os.path.isfile(output_path):
        return {
            "result": "FAILED_NO_OUTPUT_FILE",
            "path": "",
            "size": "",
            "message": (
                "STEP builder committed but no output file was created"
            ),
        }

    try:
        file_size = os.path.getsize(output_path)
    except Exception:
        file_size = ""

    inspection = inspect_step_file(output_path)
    translator = parse_step_translator_log(output_path)
    validation = evaluate_step_validation(
        nx_version,
        load_result,
        source,
        inspection,
        translator,
        file_size,
    )
    log_line(session, "    " + validation["message"], log_buffer)
    return {
        "result": validation["result"],
        "path": output_path,
        "size": file_size,
        "message": validation["message"],
        "validation": validation["validation"],
        "source": source,
        "inspection": inspection,
        "translator": translator,
    }


def export_step_for_instruction(
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
                "Master part could not be loaded. "
                "See the text log for attempted @DB names."
            ),
            "attempts": attempts,
        }

    try:
        exported = export_step_from_part(
            session,
            candidate["part"],
            output_folder,
            number,
            revision,
            log_buffer,
        )
        if exported.get("result") == "SUCCESS":
            log_line(
                session,
                "    STEP created: {0}".format(
                    exported.get("path", "")
                ),
                log_buffer,
            )
        else:
            log_line(
                session,
                "    STEP rejected: {0} - {1}".format(
                    exported.get("result", "FAILED"),
                    exported.get("message", ""),
                ),
                log_buffer,
            )
        return exported
    finally:
        restore_parts(
            session,
            original_display,
            original_work,
            log_buffer,
        )

        if candidate.get("opened_by_journal"):
            close_part_best_effort(
                candidate["part"],
                session,
                log_buffer,
            )


# ---------------------------------------------------------------------------
# Per-row processing
# ---------------------------------------------------------------------------

def overall_result(
    pdf_requested,
    pdf_result,
    step_requested,
    step_result,
):
    requested_results = []

    if pdf_requested:
        requested_results.append(pdf_result)
    if step_requested:
        requested_results.append(step_result)

    if requested_results and all(
        result == "SUCCESS"
        for result in requested_results
    ):
        return "SUCCESS"

    if any(
        result in ("SUCCESS", "PARTIAL_SUCCESS")
        for result in requested_results
    ):
        return "PARTIAL_SUCCESS"

    if (
        pdf_requested
        and not step_requested
        and pdf_result == "SKIPPED_NO_DRAWING"
    ):
        return "SKIPPED_NO_DRAWING"

    if (
        step_requested
        and not pdf_requested
        and step_result == "NOT_FOUND"
    ):
        return "NOT_FOUND"

    return "FAILED"


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

    if instruction["pdf_requested"]:
        try:
            pdf_export = export_pdfs_for_instruction(
                session,
                folders["pdf"],
                number,
                revision,
                original_display,
                original_work,
                log_buffer,
            )
            result["PDF_RESULT"] = pdf_export["result"]
            result["PDF_FILE_COUNT"] = len(pdf_export["paths"])
            result["PDF_FILES"] = ";".join(pdf_export["paths"])
            append_unique(messages, pdf_export.get("message", ""))

            for failure in pdf_export.get("failures", []):
                if failure.get("traceback"):
                    log_line(
                        session,
                        failure["traceback"],
                        log_buffer,
                    )
        except Exception as error:
            result["PDF_RESULT"] = "FAILED"
            append_unique(
                messages,
                "PDF export failed: {0}".format(error),
            )
            log_line(
                session,
                traceback.format_exc(),
                log_buffer,
            )

    if instruction["step_requested"]:
        try:
            step_export = export_step_for_instruction(
                session,
                folders["step"],
                number,
                revision,
                original_display,
                original_work,
                log_buffer,
            )
            result["STEP_RESULT"] = step_export["result"]
            result["STEP_FILE"] = step_export["path"]
            append_unique(messages, step_export.get("message", ""))

            if step_export.get("size") != "":
                append_unique(
                    messages,
                    "STEP file size: {0} bytes".format(
                        step_export["size"]
                    ),
                )
        except Exception as error:
            result["STEP_RESULT"] = "FAILED"
            append_unique(
                messages,
                "STEP export failed: {0}".format(error),
            )
            log_line(
                session,
                traceback.format_exc(),
                log_buffer,
            )

    if result["PDF_RESULT"] == "PENDING":
        result["PDF_RESULT"] = "FAILED"
    if result["STEP_RESULT"] == "PENDING":
        result["STEP_RESULT"] = "FAILED"

    result["LOADED_REVISION"] = revision
    result["OVERALL_RESULT"] = overall_result(
        instruction["pdf_requested"],
        result["PDF_RESULT"],
        instruction["step_requested"],
        result["STEP_RESULT"],
    )
    result["MESSAGE"] = " | ".join(messages)
    result["DURATION_SECONDS"] = "{0:.3f}".format(
        (datetime.datetime.now() - started).total_seconds()
    )

    restore_parts(
        session,
        original_display,
        original_work,
        log_buffer,
    )
    return result


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    session = NXOpen.Session.GetSession()
    log_buffer = []
    folders = None
    results = []
    report_path = ""
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

        log_line(
            session,
            "Journal 07 - CSV-driven PDF + STEP export",
            log_buffer,
        )
        log_line(session, "Journal build: " + JOURNAL_BUILD_ID, log_buffer)
        log_line(
            session,
            "NX runtime: " + runtime_nx_version(session),
            log_buffer,
        )
        log_line(
            session,
            "Journal source: " + runtime_source_path(),
            log_buffer,
        )
        log_line(session, "Input CSV: " + input_csv, log_buffer)
        log_line(
            session,
            (
                "Drawing resolver: "
                "@DB/<part>/<rev>/specification/<part>-<rev>-dwg<n> "
                "via session.Parts.OpenDisplay"
            ),
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
            raise FileNotFoundError(
                "Input CSV not found: {0}".format(input_csv)
            )

        parsed = read_export_scope(input_csv)
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        folders = create_run_folders(io_root, timestamp)

        report_path = os.path.join(
            folders["reports"],
            "EXPORT_RESULT_{0}.csv".format(timestamp),
        )
        log_path = os.path.join(
            folders["logs"],
            "EXPORT_LOG_{0}.txt".format(timestamp),
        )

        for warning in parsed["header_warnings"]:
            log_line(
                session,
                "WARNING: " + warning,
                log_buffer,
            )

        for invalid_instruction in parsed["invalid_rows"]:
            results.append(
                invalid_result(
                    timestamp,
                    invalid_instruction,
                )
            )

        instructions = parsed["instructions"]
        log_line(
            session,
            (
                "Input rows: {0}; unique requests: {1}; "
                "ignored: {2}; invalid: {3}"
            ).format(
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

            log_line(
                session,
                "  PDF: {0} ({1} file(s))".format(
                    result["PDF_RESULT"],
                    result["PDF_FILE_COUNT"],
                ),
                log_buffer,
            )
            log_line(
                session,
                "  STEP: {0}".format(result["STEP_RESULT"]),
                log_buffer,
            )

        restore_parts(
            session,
            original_display,
            original_work,
            log_buffer,
        )

        write_result_csv(report_path, results)
        report_written = True

        counts = {}
        for result in results:
            status = result["OVERALL_RESULT"]
            counts[status] = counts.get(status, 0) + 1

        log_line(session, "Export complete", log_buffer)
        log_line(
            session,
            "Success: {0}".format(counts.get("SUCCESS", 0)),
            log_buffer,
        )
        log_line(
            session,
            "Partial: {0}".format(
                counts.get("PARTIAL_SUCCESS", 0)
            ),
            log_buffer,
        )
        log_line(
            session,
            "No drawing: {0}".format(
                counts.get("SKIPPED_NO_DRAWING", 0)
            ),
            log_buffer,
        )
        log_line(
            session,
            "Not found: {0}".format(
                counts.get("NOT_FOUND", 0)
            ),
            log_buffer,
        )
        log_line(
            session,
            "Failed: {0}".format(counts.get("FAILED", 0)),
            log_buffer,
        )
        log_line(
            session,
            "PDF files: {0}".format(
                sum(
                    int(item["PDF_FILE_COUNT"])
                    for item in results
                )
            ),
            log_buffer,
        )
        log_line(
            session,
            "STEP files: {0}".format(
                sum(
                    1
                    for item in results
                    if item["STEP_FILE"]
                )
            ),
            log_buffer,
        )
        log_line(
            session,
            "Result report: " + report_path,
            log_buffer,
        )

        write_text_log(log_path, log_buffer)

    except Exception:
        log_line(
            session,
            "ERROR: Unhandled journal exception.",
            log_buffer,
        )
        log_line(
            session,
            traceback.format_exc(),
            log_buffer,
        )

    finally:
        restore_parts(
            session,
            original_display,
            original_work,
            log_buffer,
        )

        if folders is not None:
            if report_path and not report_written:
                try:
                    write_result_csv(report_path, results)
                except Exception:
                    pass

            try:
                fallback_log = os.path.join(
                    folders["logs"],
                    "EXPORT_LOG_{0}.txt".format(
                        os.path.basename(folders["run"])
                    ),
                )
                write_text_log(fallback_log, log_buffer)
            except Exception:
                pass


if __name__ == "__main__":
    main()
