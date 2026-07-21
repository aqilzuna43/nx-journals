"""
Journal 07 - DataPack-Controlled PDF + STEP Export

Reads NX_EXPORT_SCOPE.csv, matches DB_PART_NO + DB_PART_REV against the
currently loaded assembly, and exports requested drawing PDFs and AP214 STEP
files.

Drawing specifications do not need to be open. In Teamcenter managed mode the
journal searches loaded drawing parts first, then attempts numbered non-master
specifications such as:

    <DB_PART_NO>-<DB_PART_REV>-dwg1
    <DB_PART_NO>-<DB_PART_REV>-dwg2
    ...

Target: NX 2312 embedded Python 3.10
Run via: NX > Tools > Journal > Play
"""

import csv
import datetime
import os
import re
import traceback

import NXOpen
import NXOpen.UF


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

INPUT_FILENAME = "NX_EXPORT_SCOPE.csv"
OUTPUT_ROOT_FOLDER = "NX_BULK_EXPORT"
STEP_FORMAT = "AP214"
INCLUDE_ROOT_WORK_PART = True
SKIP_SUPPRESSED_COMPONENTS = True
VERIFY_OUTPUT_FILES = True

# Supports dwg1, dwg2 and further numbered specifications.
MAX_DRAWING_DATASET_INDEX = 9
TEAMCENTER_DRAWING_DATASET_TYPE = "UGPART"

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
    root = normalize_text(os.environ.get("NX_JOURNALS_IO_DIR")) or desktop_folder()
    return os.path.abspath(os.path.expanduser(root))


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
    if value is not None:
        try:
            value.Dispose()
        except Exception:
            pass


def session_is_managed(session):
    try:
        value = session.IsManagedMode
        return bool(value() if callable(value) else value)
    except Exception:
        return False


def safe_part_name(part, fallback="part"):
    for name in ("FullPath", "Leaf", "Name"):
        try:
            value = normalize_text(getattr(part, name))
            if value:
                if name == "FullPath":
                    value = os.path.splitext(os.path.basename(value))[0]
                return value
        except Exception:
            pass
    return fallback


def object_identity(nx_object):
    try:
        return ("TAG", str(nx_object.Tag))
    except Exception:
        pass
    try:
        path = normalize_text(nx_object.FullPath)
        if path:
            return ("PATH", path.upper())
    except Exception:
        pass
    return ("OBJECT", id(nx_object))


def get_string_attribute(nx_object, name, fallback=""):
    if nx_object is None:
        return fallback
    try:
        return normalize_text(nx_object.GetStringAttribute(name))
    except Exception:
        pass
    try:
        attribute = nx_object.GetUserAttribute(
            name,
            NXOpen.NXObject.AttributeType.String,
            -1,
        )
        return normalize_text(attribute.StringValue)
    except Exception:
        return fallback


def get_part_identity(part, component=None):
    number = (
        get_string_attribute(part, "DB_PART_NO")
        or get_string_attribute(component, "DB_PART_NO")
        or get_string_attribute(part, "PART_NUMBER")
        or get_string_attribute(component, "PART_NUMBER")
    )
    revision = (
        get_string_attribute(part, "DB_PART_REV")
        or get_string_attribute(component, "DB_PART_REV")
        or get_string_attribute(part, "REVISION")
        or get_string_attribute(component, "REVISION")
    )
    return normalize_text(number), normalize_text(revision)


# ---------------------------------------------------------------------------
# CSV input and reports
# ---------------------------------------------------------------------------

def resolve_headers(fieldnames):
    if not fieldnames:
        raise ValueError("The input CSV does not contain a header row.")

    normalized = [(name, normalize_header(name)) for name in fieldnames]
    resolved = {}
    warnings = []

    for logical_name, aliases in _HEADER_ALIASES.items():
        matches = []
        for alias in aliases:
            wanted = normalize_header(alias)
            for original, current in normalized:
                if current == wanted and original not in matches:
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
    field = headers.get(logical_name)
    return normalize_text(row.get(field, "")) if field else ""


def parse_control(value, label, row_number):
    normalized = normalize_text(value).upper()
    if normalized in TRUE_VALUES:
        return True, ""
    if normalized in FALSE_VALUES:
        return False, ""
    return False, (
        "Source row {0}: unknown {1} control value '{2}'; treated as disabled"
        .format(row_number, label, normalize_text(value))
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
    invalid = []
    ignored = 0
    input_rows = 0

    with open(csv_path, "r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        headers, header_warnings = resolve_headers(reader.fieldnames)

        for row_number, row in enumerate(reader, start=2):
            if row_is_blank(row):
                continue
            input_rows += 1

            pdf, pdf_warning = parse_control(
                row_value(row, headers, "pdf"), "PDF", row_number
            )
            step, step_warning = parse_control(
                row_value(row, headers, "step"), "STEP", row_number
            )
            warnings = [item for item in (pdf_warning, step_warning) if item]

            if not pdf and not step and not warnings:
                ignored += 1
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
            if not pdf and not step:
                errors.append("No valid PDF or STEP request remains")

            item = {
                "source_rows": [row_number],
                "source_row_count": 1,
                "merged_row_count": 0,
                "part_number": number,
                "revision": revision,
                "normalized_key": (number.upper(), revision.upper()),
                "pdf_requested": pdf,
                "step_requested": step,
                "warnings": warnings,
                **optional,
            }

            if errors:
                item["warnings"].append(
                    "Source row {0}: {1}".format(row_number, "; ".join(errors))
                )
                invalid.append(item)
                continue

            key = item["normalized_key"]
            current = merged.get(key)
            if current is None:
                merged[key] = item
                continue

            current["source_rows"].append(row_number)
            current["source_row_count"] += 1
            current["merged_row_count"] = current["source_row_count"] - 1
            current["pdf_requested"] = current["pdf_requested"] or pdf
            current["step_requested"] = current["step_requested"] or step
            for name, value in optional.items():
                if value and not current[name]:
                    current[name] = value
            for warning in warnings:
                append_unique(current["warnings"], warning)

    return {
        "instructions": sorted(merged.values(), key=lambda item: item["normalized_key"]),
        "invalid_rows": invalid,
        "ignored_row_count": ignored,
        "input_row_count": input_rows,
        "header_warnings": header_warnings,
    }


def new_result(timestamp, instruction):
    pdf = bool(instruction.get("pdf_requested"))
    step = bool(instruction.get("step_requested"))
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
        "PDF_REQUESTED": "YES" if pdf else "NO",
        "PDF_RESULT": "PENDING" if pdf else "NOT_REQUESTED",
        "PDF_FILE_COUNT": 0,
        "PDF_FILES": "",
        "STEP_REQUESTED": "YES" if step else "NO",
        "STEP_RESULT": "PENDING" if step else "NOT_REQUESTED",
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
            writer.writerow({name: result.get(name, "") for name in _RESULT_COLUMNS})


# ---------------------------------------------------------------------------
# Loaded assembly traversal
# ---------------------------------------------------------------------------

def build_loaded_part_map(work_part):
    exact = {}
    revisions = {}
    diagnostics = {
        "unresolved_components": [],
        "suppressed_component_count": 0,
        "prototype_error_count": 0,
        "duplicate_loaded_key_count": 0,
        "collision_keys": set(),
    }
    seen = set()

    def add_part(part, component=None):
        if part is None:
            return False
        identity = object_identity(part)
        if identity in seen:
            return True
        seen.add(identity)

        number, revision = get_part_identity(part, component)
        if not number or not revision:
            return False

        key = (number.upper(), revision.upper())
        revisions.setdefault(key[0], set()).add(key[1])
        if key not in exact:
            exact[key] = part
        elif object_identity(exact[key]) != identity:
            diagnostics["duplicate_loaded_key_count"] += 1
            diagnostics["collision_keys"].add(key)
        return True

    if INCLUDE_ROOT_WORK_PART:
        add_part(work_part)

    try:
        root = work_part.ComponentAssembly.RootComponent
    except Exception:
        root = None
    if root is None:
        return exact, revisions, diagnostics

    stack = [root]
    while stack:
        component = stack.pop()
        label = "<component>"
        for name in ("DisplayName", "Name"):
            try:
                value = normalize_text(getattr(component, name))
                if value:
                    label = value
                    break
            except Exception:
                pass

        if SKIP_SUPPRESSED_COMPONENTS:
            try:
                if component.IsSuppressed:
                    diagnostics["suppressed_component_count"] += 1
                    continue
            except Exception:
                pass

        try:
            prototype = component.Prototype
        except Exception as error:
            prototype = None
            diagnostics["prototype_error_count"] += 1
            diagnostics["unresolved_components"].append(
                "{0}: could not access prototype ({1})".format(label, error)
            )

        if prototype is None:
            diagnostics["unresolved_components"].append(
                "{0}: no usable loaded prototype".format(label)
            )
        elif not add_part(prototype, component):
            diagnostics["unresolved_components"].append(
                "{0}: prototype has no usable part number/revision".format(label)
            )

        try:
            children = list(component.GetChildren())
        except Exception as error:
            diagnostics["unresolved_components"].append(
                "{0}: could not read child branch ({1})".format(label, error)
            )
            continue
        stack.extend(reversed(children))

    return exact, revisions, diagnostics


# ---------------------------------------------------------------------------
# Drawing discovery and PDF export
# ---------------------------------------------------------------------------

def drawing_sheet_count(part):
    try:
        return int(part.DrawingSheets.Count)
    except Exception:
        pass
    try:
        return len(list(part.DrawingSheets))
    except Exception:
        return 0


def part_identifiers(part):
    values = []
    for name in ("FullPath", "Leaf", "Name"):
        try:
            value = normalize_text(getattr(part, name))
            if value:
                values.append(value)
        except Exception:
            pass
    return values


def drawing_index(part):
    for identifier in part_identifiers(part):
        match = _DRAWING_SUFFIX_RE.search(identifier.upper())
        if match:
            try:
                return int(match.group(1))
            except Exception:
                pass
    return None


def drawing_matches(part, number, revision):
    if part is None or drawing_sheet_count(part) <= 0:
        return False

    text = " | ".join(part_identifiers(part)).upper()
    expected = "{0}-{1}-DWG".format(number.upper(), revision.upper())
    if expected in text:
        return True

    loaded_number, loaded_revision = get_part_identity(part)
    return (
        loaded_number.upper() == number.upper()
        and loaded_revision.upper() == revision.upper()
        and "DWG" in text
    )


def loaded_drawing_parts(session, number, revision):
    try:
        session_parts = list(session.Parts)
    except Exception:
        session_parts = []
    try:
        display = session.Parts.Display
    except Exception:
        display = None

    ordered = ([display] if display is not None else []) + session_parts
    results = []
    seen = set()
    for part in ordered:
        if part is None or object_identity(part) in seen:
            continue
        seen.add(object_identity(part))
        if drawing_matches(part, number, revision):
            results.append({
                "part": part,
                "opened_by_journal": False,
                "drawing_index": drawing_index(part),
                "source": "loaded session",
            })

    return sorted(
        results,
        key=lambda item: (
            item["drawing_index"] is None,
            item["drawing_index"] or 9999,
            safe_part_name(item["part"]).upper(),
        ),
    )


def unwrap_open_result(value):
    if isinstance(value, tuple):
        return (
            value[0] if value else None,
            value[1] if len(value) > 1 else None,
        )
    return value, None


def teamcenter_specs(number, revision, index):
    dataset_name = "{0}-{1}-dwg{2}".format(number, revision, index)
    return [
        "@DB/{0}/{1}/{2}/{3}".format(
            number,
            revision,
            TEAMCENTER_DRAWING_DATASET_TYPE,
            dataset_name,
        ),
        "@DB/{0}/{1}/{2}/dwg{3}".format(
            number,
            revision,
            TEAMCENTER_DRAWING_DATASET_TYPE,
            index,
        ),
        "@DB/{0}/{1}/dwg{2}".format(number, revision, index),
    ]


def close_part_best_effort(part):
    if part is None:
        return
    try:
        part.Close(
            NXOpen.BasePart.CloseWholeTree.FalseValue,
            NXOpen.BasePart.CloseModified.CloseModified,
            None,
        )
    except Exception:
        pass


def resolve_drawing_parts(session, number, revision, log_buffer):
    resolved = loaded_drawing_parts(session, number, revision)
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
            "  Drawing already loaded: {0}".format(safe_part_name(item["part"])),
            log_buffer,
        )

    if not session_is_managed(session):
        return resolved, attempts

    for index in range(1, MAX_DRAWING_DATASET_INDEX + 1):
        if index in known_indices:
            continue

        for specification in teamcenter_specs(number, revision, index):
            attempts.append(specification)
            log_line(session, "  Attempt drawing open: " + specification, log_buffer)

            part = None
            status = None
            try:
                part, status = unwrap_open_result(session.Parts.OpenBase(specification))
            except Exception as error:
                log_line(session, "    Not opened: {0}".format(error), log_buffer)
                continue
            finally:
                dispose(status)

            if part is None:
                continue
            identity = object_identity(part)
            if identity in seen:
                known_indices.add(index)
                break

            count = drawing_sheet_count(part)
            if count <= 0:
                log_line(
                    session,
                    "    Opened but no drawing sheets: {0}".format(safe_part_name(part)),
                    log_buffer,
                )
                close_part_best_effort(part)
                continue

            seen.add(identity)
            known_indices.add(index)
            resolved.append({
                "part": part,
                "opened_by_journal": True,
                "drawing_index": index,
                "source": specification,
            })
            log_line(
                session,
                "    Drawing opened: {0}; sheets: {1}".format(
                    safe_part_name(part), count
                ),
                log_buffer,
            )
            break

    return sorted(
        resolved,
        key=lambda item: (
            item["drawing_index"] is None,
            item["drawing_index"] or 9999,
            safe_part_name(item["part"]).upper(),
        ),
    ), attempts


def set_display_part(session, part):
    result = session.Parts.SetDisplay(part, False, True)
    if isinstance(result, tuple) and len(result) > 1:
        dispose(result[1])


def drawing_token(part, index):
    if index is not None:
        return "DWG{0}".format(index)
    for identifier in part_identifiers(part):
        match = _DRAWING_SUFFIX_RE.search(identifier.upper())
        if match:
            return "DWG{0}".format(match.group(1))
    return "DRAWING"


def build_pdf_filename(
    part,
    number,
    revision,
    token,
    sheet,
    sheet_index,
    sheet_count,
    drawing_count,
):
    drawing_number = get_string_attribute(part, "DRAWING_NUMBER")
    base = drawing_number or number or safe_part_name(part)
    name = "{0}_REV{1}".format(
        clean_filename_token(base),
        clean_filename_token(revision, fallback=""),
    )

    # Prevent dwg1/dwg2 filename collisions.
    if drawing_count > 1 or not drawing_number:
        name += "_" + clean_filename_token(token)

    if sheet_count > 1:
        name += "_SHEET{0:02d}".format(sheet_index)
        try:
            sheet_name = normalize_text(sheet.Name)
        except Exception:
            sheet_name = ""
        if sheet_name:
            name += "_" + clean_filename_token(sheet_name, fallback="")
    return name + ".pdf"


def export_one_sheet_pdf(part, sheet, output_path):
    builder = part.PlotManager.CreatePrintPdfbuilder()
    try:
        # Critical in Teamcenter managed mode: write a native Windows PDF,
        # rather than creating or overwriting a Teamcenter PDF dataset.
        builder.Action = NXOpen.PrintPDFBuilder.ActionOption.Native
        builder.Filename = output_path
        builder.Append = False
        builder.SourceBuilder.SetSheets([sheet])
        builder.Commit()
    finally:
        builder.Destroy()


def export_pdfs(
    session,
    output_folder,
    number,
    revision,
    original_display,
    original_work,
    log_buffer,
):
    drawing_items, attempts = resolve_drawing_parts(
        session, number, revision, log_buffer
    )
    if not drawing_items:
        return {
            "result": "SKIPPED_NO_DRAWING",
            "paths": [],
            "message": "No drawing specification with sheets was found. Attempted: {0}".format(
                "; ".join(attempts)
            ),
            "failures": [],
        }

    successes = []
    failures = []
    drawing_count = len(drawing_items)

    try:
        for item in drawing_items:
            part = item["part"]
            token = drawing_token(part, item["drawing_index"])
            try:
                set_display_part(session, part)
                sheets = list(part.DrawingSheets)
            except Exception as error:
                failures.append({
                    "kind": "ERROR",
                    "message": "{0}: unable to display/enumerate sheets: {1}".format(
                        token, error
                    ),
                    "traceback": traceback.format_exc(),
                })
                continue

            log_line(
                session,
                "  Exporting {0}: {1} sheet(s)".format(token, len(sheets)),
                log_buffer,
            )

            for sheet_index, sheet in enumerate(sheets, start=1):
                filename = build_pdf_filename(
                    part,
                    number,
                    revision,
                    token,
                    sheet,
                    sheet_index,
                    len(sheets),
                    drawing_count,
                )
                output_path = os.path.join(output_folder, filename)
                try:
                    if os.path.exists(output_path):
                        raise RuntimeError("PDF output already exists: " + output_path)
                    sheet.Open()
                    export_one_sheet_pdf(part, sheet, output_path)
                    if VERIFY_OUTPUT_FILES and not os.path.isfile(output_path):
                        failures.append({
                            "kind": "NO_OUTPUT",
                            "message": "{0} sheet {1}: builder committed but no PDF was created".format(
                                token, sheet_index
                            ),
                        })
                    else:
                        successes.append(output_path)
                        log_line(session, "    PDF created: " + output_path, log_buffer)
                except Exception as error:
                    failures.append({
                        "kind": "ERROR",
                        "message": "{0} sheet {1}: {2}".format(
                            token, sheet_index, error
                        ),
                        "traceback": traceback.format_exc(),
                    })
    finally:
        if original_display is not None:
            try:
                set_display_part(session, original_display)
            except Exception:
                pass
        if original_work is not None:
            try:
                session.Parts.SetWork(original_work)
            except Exception:
                pass
        for item in drawing_items:
            if item["opened_by_journal"]:
                close_part_best_effort(item["part"])

    if successes and not failures:
        result = "SUCCESS"
    elif successes:
        result = "PARTIAL_SUCCESS"
    elif failures and all(item["kind"] == "NO_OUTPUT" for item in failures):
        result = "FAILED_NO_OUTPUT_FILE"
    else:
        result = "FAILED"

    return {
        "result": result,
        "paths": successes,
        "message": " | ".join(item["message"] for item in failures),
        "failures": failures,
    }


# ---------------------------------------------------------------------------
# STEP export
# ---------------------------------------------------------------------------

def export_step(session, part, output_folder, number, revision):
    session.Parts.SetWork(part)
    output_path = os.path.join(
        output_folder,
        "{0}_REV{1}.stp".format(
            clean_filename_token(number),
            clean_filename_token(revision, fallback=""),
        ),
    )
    if os.path.exists(output_path):
        raise RuntimeError("STEP output already exists: " + output_path)

    exporter = session.DexManager.CreateStepCreator()
    try:
        try:
            exporter.InputFile = part.FullPath
        except Exception:
            pass
        exporter.OutputFile = output_path
        exporter.ObjectTypes.Solids = True
        exporter.ObjectTypes.Surfaces = True
        exporter.ObjectTypes.Curves = True
        if STEP_FORMAT != "AP214":
            raise RuntimeError("Unsupported STEP_FORMAT: " + STEP_FORMAT)
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
            "message": "STEP builder committed but no file was created",
        }
    try:
        size = os.path.getsize(output_path)
    except Exception:
        size = ""
    return {
        "result": "SUCCESS",
        "path": output_path,
        "size": size,
        "message": "",
    }


# ---------------------------------------------------------------------------
# Processing and entry point
# ---------------------------------------------------------------------------

def overall_result(pdf_requested, pdf_result, step_requested, step_result):
    requested = []
    if pdf_requested:
        requested.append(pdf_result)
    if step_requested:
        requested.append(step_result)
    if requested and all(item == "SUCCESS" for item in requested):
        return "SUCCESS"
    if any(item in ("SUCCESS", "PARTIAL_SUCCESS") for item in requested):
        return "PARTIAL_SUCCESS"
    if requested and all(item == "SKIPPED_NO_DRAWING" for item in requested):
        return "SKIPPED_NO_DRAWING"
    return "FAILED"


def process_instruction(
    session,
    instruction,
    exact_parts,
    revisions,
    collision_keys,
    folders,
    timestamp,
    original_display,
    original_work,
    log_buffer,
):
    started = datetime.datetime.now()
    result = new_result(timestamp, instruction)
    messages = list(instruction.get("warnings", []))
    key = instruction["normalized_key"]
    part = exact_parts.get(key)

    if part is None:
        loaded_revisions = sorted(revisions.get(key[0], set()))
        status = "REVISION_MISMATCH" if loaded_revisions else "NOT_FOUND"
        message = (
            "Requested revision {0}; loaded revisions: {1}".format(
                instruction["revision"], ", ".join(loaded_revisions)
            )
            if loaded_revisions
            else "Part number was not present in the loaded assembly"
        )
        if result["PDF_REQUESTED"] == "YES":
            result["PDF_RESULT"] = status
        if result["STEP_REQUESTED"] == "YES":
            result["STEP_RESULT"] = status
        result["LOADED_REVISION"] = ";".join(loaded_revisions)
        result["OVERALL_RESULT"] = status
        result["MESSAGE"] = " | ".join(messages + [message])
        result["DURATION_SECONDS"] = "{0:.3f}".format(
            (datetime.datetime.now() - started).total_seconds()
        )
        return result

    result["LOADED_REVISION"] = key[1]
    if key in collision_keys:
        append_unique(
            messages,
            "Multiple loaded prototypes share this key; first prototype used",
        )

    if instruction["pdf_requested"]:
        try:
            exported = export_pdfs(
                session,
                folders["pdf"],
                instruction["part_number"],
                instruction["revision"],
                original_display,
                original_work,
                log_buffer,
            )
            result["PDF_RESULT"] = exported["result"]
            result["PDF_FILE_COUNT"] = len(exported["paths"])
            result["PDF_FILES"] = ";".join(exported["paths"])
            append_unique(messages, exported.get("message", ""))
            for failure in exported.get("failures", []):
                if failure.get("traceback"):
                    log_line(session, failure["traceback"], log_buffer)
        except Exception as error:
            result["PDF_RESULT"] = "FAILED"
            append_unique(messages, "PDF export failed: {0}".format(error))
            log_line(session, traceback.format_exc(), log_buffer)

    if instruction["step_requested"]:
        try:
            exported = export_step(
                session,
                part,
                folders["step"],
                instruction["part_number"],
                instruction["revision"],
            )
            result["STEP_RESULT"] = exported["result"]
            result["STEP_FILE"] = exported["path"]
            append_unique(messages, exported.get("message", ""))
            if exported.get("size") != "":
                append_unique(
                    messages,
                    "STEP file size: {0} bytes".format(exported["size"]),
                )
        except Exception as error:
            result["STEP_RESULT"] = "FAILED"
            append_unique(messages, "STEP export failed: {0}".format(error))
            log_line(session, traceback.format_exc(), log_buffer)

    if result["PDF_RESULT"] == "PENDING":
        result["PDF_RESULT"] = "FAILED"
    if result["STEP_RESULT"] == "PENDING":
        result["STEP_RESULT"] = "FAILED"

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
    return result


def restore_parts(session, display, work, log_buffer):
    if display is not None:
        try:
            set_display_part(session, display)
        except Exception as error:
            log_line(session, "ERROR restoring display part: {0}".format(error), log_buffer)
    if work is not None:
        try:
            session.Parts.SetWork(work)
        except Exception as error:
            log_line(session, "ERROR restoring work part: {0}".format(error), log_buffer)


def main():
    session = NXOpen.Session.GetSession()
    log_buffer = []
    folders = None
    results = []
    report_path = ""
    report_written = False

    try:
        work_part = session.Parts.Work
        if work_part is None:
            raise RuntimeError("No NX work part is loaded.")

        original_work = session.Parts.Work
        original_display = session.Parts.Display
        io_root = resolve_io_root()
        input_csv = os.path.join(io_root, INPUT_FILENAME)

        log_line(session, "Journal 07 - DataPack PDF + STEP export", log_buffer)
        log_line(session, "Input CSV: " + input_csv, log_buffer)
        log_line(session, "Managed mode: {0}".format(session_is_managed(session)), log_buffer)

        if not os.path.isfile(input_csv):
            raise FileNotFoundError("Input CSV not found: " + input_csv)

        parsed = read_export_scope(input_csv)
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        folders = create_run_folders(io_root, timestamp)
        report_path = os.path.join(
            folders["reports"], "EXPORT_RESULT_{0}.csv".format(timestamp)
        )
        log_path = os.path.join(
            folders["logs"], "EXPORT_LOG_{0}.txt".format(timestamp)
        )

        for warning in parsed["header_warnings"]:
            log_line(session, "WARNING: " + warning, log_buffer)
        for item in parsed["invalid_rows"]:
            results.append(invalid_result(timestamp, item))

        exact_parts, revisions, diagnostics = build_loaded_part_map(work_part)
        instructions = parsed["instructions"]
        log_line(
            session,
            "Input rows: {0}; unique requests: {1}; ignored: {2}; invalid: {3}".format(
                parsed["input_row_count"],
                len(instructions),
                parsed["ignored_row_count"],
                len(parsed["invalid_rows"]),
            ),
            log_buffer,
        )

        try:
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
                    exact_parts,
                    revisions,
                    diagnostics["collision_keys"],
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
                        result["PDF_RESULT"], result["PDF_FILE_COUNT"]
                    ),
                    log_buffer,
                )
                log_line(session, "  STEP: " + result["STEP_RESULT"], log_buffer)
        finally:
            restore_parts(session, original_display, original_work, log_buffer)

        write_result_csv(report_path, results)
        report_written = True

        counts = {}
        for result in results:
            status = result["OVERALL_RESULT"]
            counts[status] = counts.get(status, 0) + 1

        log_line(session, "Export complete", log_buffer)
        log_line(session, "Success: {0}".format(counts.get("SUCCESS", 0)), log_buffer)
        log_line(session, "Partial: {0}".format(counts.get("PARTIAL_SUCCESS", 0)), log_buffer)
        log_line(session, "Failed: {0}".format(counts.get("FAILED", 0)), log_buffer)
        log_line(
            session,
            "PDF files: {0}".format(
                sum(int(item["PDF_FILE_COUNT"]) for item in results)
            ),
            log_buffer,
        )
        log_line(
            session,
            "STEP files: {0}".format(
                sum(1 for item in results if item["STEP_FILE"])
            ),
            log_buffer,
        )
        log_line(session, "Result report: " + report_path, log_buffer)

        for diagnostic in diagnostics["unresolved_components"]:
            log_line(session, "WARNING: " + diagnostic, log_buffer)

        write_text_log(log_path, log_buffer)

    except Exception:
        log_line(session, "ERROR: Unhandled journal exception.", log_buffer)
        log_line(session, traceback.format_exc(), log_buffer)
    finally:
        if folders is not None:
            if report_path and not report_written:
                try:
                    write_result_csv(report_path, results)
                except Exception:
                    pass
            try:
                fallback = os.path.join(
                    folders["logs"],
                    "EXPORT_LOG_{0}.txt".format(os.path.basename(folders["run"])),
                )
                write_text_log(fallback, log_buffer)
            except Exception:
                pass


if __name__ == "__main__":
    main()
