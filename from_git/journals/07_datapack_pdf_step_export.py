"""
Journal 07 - DataPack-Controlled PDF + STEP Export

Reads NX_EXPORT_SCOPE.csv, matches DB_PART_NO + DB_PART_REV against the
currently loaded assembly, and exports the explicitly requested drawing PDFs
and AP214 STEP files.

Run via: NX > Tools > Journal > Play
"""

import csv
import datetime
import os
import traceback

import NXOpen
import NXOpen.UF


# ---------------------------------------------------------------------------
# Operator-adjustable configuration
# ---------------------------------------------------------------------------

INPUT_FILENAME = "NX_EXPORT_SCOPE.csv"
OUTPUT_ROOT_FOLDER = "NX_BULK_EXPORT"
STEP_FORMAT = "AP214"
INCLUDE_ROOT_WORK_PART = True
SKIP_SUPPRESSED_COMPONENTS = True
VERIFY_OUTPUT_FILES = True

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

_REQUIRED_LOGICAL_HEADERS = ("part_number", "revision", "pdf", "step")

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


# ---------------------------------------------------------------------------
# Text, logging, and file helpers
# ---------------------------------------------------------------------------

def normalize_text(value):
    if value is None:
        return ""
    return str(value).strip()


def normalize_header(value):
    text = normalize_text(value).lstrip("\ufeff")
    return " ".join(text.split()).upper()


def clean_filename_token(value, fallback="part"):
    text = normalize_text(value)
    if not text:
        return fallback

    cleaned = []
    for char in text:
        if char in _INVALID_FILENAME_CHARS or ord(char) < 32:
            cleaned.append("_")
        else:
            cleaned.append(char)

    result = "".join(cleaned).strip(" .")
    return result or fallback


def is_enabled(value):
    return normalize_text(value).upper() in TRUE_VALUES


def _parse_control(value, label, row_number):
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


def _append_unique(messages, message):
    message = normalize_text(message)
    if message and message not in messages:
        messages.append(message)


def log_line(session, message, log_buffer=None):
    text = str(message)
    if log_buffer is not None:
        log_buffer.append(text)

    try:
        listing_window = session.ListingWindow
        listing_window.Open()
        for line in text.splitlines() or [""]:
            listing_window.WriteFullline(line)
    except Exception:
        pass

    try:
        print(text)
    except Exception:
        pass


def _write_text_log(path, lines):
    with open(path, "w", encoding="utf-8", newline="") as handle:
        for line in lines:
            handle.write(str(line))
            handle.write("\n")


def _desktop_folder():
    user_profile = normalize_text(os.environ.get("USERPROFILE"))
    if user_profile:
        return os.path.join(user_profile, "Desktop")

    home = os.path.expanduser("~")
    if home and home != "~":
        return os.path.join(home, "Desktop")

    return os.getcwd()


def resolve_io_root():
    configured = normalize_text(os.environ.get("NX_JOURNALS_IO_DIR"))
    root = configured or _desktop_folder()
    return os.path.abspath(os.path.expanduser(root))


def resolve_input_csv(io_root):
    return os.path.join(io_root, INPUT_FILENAME)


def _create_run_folders(io_root, run_timestamp):
    run_folder = os.path.join(io_root, OUTPUT_ROOT_FOLDER, run_timestamp)
    os.makedirs(run_folder, exist_ok=False)

    folders = {"run": run_folder}
    for name in ("PDF", "STEP", "REPORTS", "LOGS"):
        path = os.path.join(run_folder, name)
        os.makedirs(path, exist_ok=False)
        folders[name.lower()] = path
    return folders


# ---------------------------------------------------------------------------
# CSV input and result helpers
# ---------------------------------------------------------------------------

def _resolve_headers(fieldnames):
    if not fieldnames:
        raise ValueError("The input CSV does not contain a header row.")

    normalized_fields = []
    for fieldname in fieldnames:
        normalized_fields.append((fieldname, normalize_header(fieldname)))

    resolved = {}
    warnings = []
    for logical_name, aliases in _HEADER_ALIASES.items():
        candidates = []
        for alias in aliases:
            normalized_alias = normalize_header(alias)
            for original, normalized in normalized_fields:
                if normalized == normalized_alias and original not in candidates:
                    candidates.append(original)

        if candidates:
            resolved[logical_name] = candidates[0]
            if len(candidates) > 1:
                warnings.append(
                    "Multiple columns match {0}; using '{1}' and ignoring: {2}".format(
                        logical_name,
                        candidates[0],
                        ", ".join(str(item) for item in candidates[1:]),
                    )
                )

    missing = [
        logical_name
        for logical_name in _REQUIRED_LOGICAL_HEADERS
        if logical_name not in resolved
    ]
    if missing:
        descriptions = {
            "part_number": "part-number",
            "revision": "revision",
            "pdf": "PDF control",
            "step": "STEP control",
        }
        raise ValueError(
            "Missing required logical CSV column(s): {0}".format(
                ", ".join(descriptions[item] for item in missing)
            )
        )

    return resolved, warnings


def _row_value(row, resolved_headers, logical_name):
    fieldname = resolved_headers.get(logical_name)
    if not fieldname:
        return ""
    return normalize_text(row.get(fieldname, ""))


def _optional_values(row, resolved_headers):
    return {
        "part_description": _row_value(row, resolved_headers, "part_description"),
        "primary_module": _row_value(row, resolved_headers, "primary_module"),
        "data_pack_status": _row_value(row, resolved_headers, "data_pack_status"),
        "owner": _row_value(row, resolved_headers, "owner"),
    }


def _row_is_blank(row):
    for value in row.values():
        if isinstance(value, list):
            if any(normalize_text(item) for item in value):
                return False
        elif normalize_text(value):
            return False
    return True


def read_export_scope(csv_path):
    instructions_by_key = {}
    invalid_rows = []
    ignored_row_count = 0
    input_row_count = 0

    with open(csv_path, "r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        resolved_headers, header_warnings = _resolve_headers(reader.fieldnames)

        for row_number, row in enumerate(reader, start=2):
            if _row_is_blank(row):
                continue

            input_row_count += 1
            pdf_requested, pdf_warning = _parse_control(
                _row_value(row, resolved_headers, "pdf"),
                "PDF",
                row_number,
            )
            step_requested, step_warning = _parse_control(
                _row_value(row, resolved_headers, "step"),
                "STEP",
                row_number,
            )
            control_warnings = [
                warning for warning in (pdf_warning, step_warning) if warning
            ]

            if not pdf_requested and not step_requested and not control_warnings:
                ignored_row_count += 1
                continue

            part_number = _row_value(row, resolved_headers, "part_number")
            revision = _row_value(row, resolved_headers, "revision")
            optional = _optional_values(row, resolved_headers)

            validation_errors = []
            if not part_number:
                validation_errors.append("Part number is blank")
            if not revision:
                validation_errors.append("Revision is blank")
            if not pdf_requested and not step_requested:
                validation_errors.append("No valid PDF or STEP request remains")

            if validation_errors:
                messages = list(control_warnings)
                messages.append(
                    "Source row {0}: {1}".format(
                        row_number,
                        "; ".join(validation_errors),
                    )
                )
                invalid_rows.append(
                    {
                        "source_rows": [row_number],
                        "source_row_count": 1,
                        "merged_row_count": 0,
                        "part_number": part_number,
                        "revision": revision,
                        "normalized_key": (
                            part_number.upper(),
                            revision.upper(),
                        ),
                        "pdf_requested": pdf_requested,
                        "step_requested": step_requested,
                        "part_description": optional["part_description"],
                        "primary_module": optional["primary_module"],
                        "data_pack_status": optional["data_pack_status"],
                        "owner": optional["owner"],
                        "warnings": messages,
                    }
                )
                continue

            key = (part_number.upper(), revision.upper())
            instruction = instructions_by_key.get(key)
            if instruction is None:
                instruction = {
                    "source_rows": [],
                    "source_row_count": 0,
                    "merged_row_count": 0,
                    "part_number": part_number,
                    "revision": revision,
                    "normalized_key": key,
                    "pdf_requested": False,
                    "step_requested": False,
                    "part_description": "",
                    "primary_module": "",
                    "data_pack_status": "",
                    "owner": "",
                    "warnings": [],
                }
                instructions_by_key[key] = instruction

            instruction["source_rows"].append(row_number)
            instruction["source_row_count"] += 1
            instruction["merged_row_count"] = instruction["source_row_count"] - 1
            instruction["pdf_requested"] = (
                instruction["pdf_requested"] or pdf_requested
            )
            instruction["step_requested"] = (
                instruction["step_requested"] or step_requested
            )

            for optional_name, optional_value in optional.items():
                if optional_value and not instruction[optional_name]:
                    instruction[optional_name] = optional_value

            for warning in control_warnings:
                _append_unique(instruction["warnings"], warning)

    instructions = sorted(
        instructions_by_key.values(),
        key=lambda item: item["normalized_key"],
    )
    return {
        "instructions": instructions,
        "invalid_rows": invalid_rows,
        "ignored_row_count": ignored_row_count,
        "input_row_count": input_row_count,
        "header_warnings": header_warnings,
    }


def merge_instructions(rows):
    """Merge already parsed instruction dictionaries by normalized identity."""
    merged = {}
    for row in rows:
        key = row["normalized_key"]
        current = merged.get(key)
        if current is None:
            current = dict(row)
            current["source_rows"] = list(row.get("source_rows", []))
            current["warnings"] = list(row.get("warnings", []))
            merged[key] = current
            continue

        current["source_rows"].extend(row.get("source_rows", []))
        current["source_row_count"] += row.get("source_row_count", 1)
        current["merged_row_count"] = current["source_row_count"] - 1
        current["pdf_requested"] = (
            current["pdf_requested"] or row.get("pdf_requested", False)
        )
        current["step_requested"] = (
            current["step_requested"] or row.get("step_requested", False)
        )
        for name in (
            "part_description",
            "primary_module",
            "data_pack_status",
            "owner",
        ):
            if not current.get(name) and row.get(name):
                current[name] = row[name]
        for warning in row.get("warnings", []):
            _append_unique(current["warnings"], warning)

    return sorted(merged.values(), key=lambda item: item["normalized_key"])


def _new_result(run_timestamp, instruction):
    pdf_requested = bool(instruction.get("pdf_requested"))
    step_requested = bool(instruction.get("step_requested"))
    return {
        "RUN_TIMESTAMP": run_timestamp,
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
        "DURATION_SECONDS": "0.000",
    }


def _invalid_result(run_timestamp, invalid_row):
    result = _new_result(run_timestamp, invalid_row)
    if invalid_row.get("pdf_requested"):
        result["PDF_RESULT"] = "INVALID_INPUT"
    if invalid_row.get("step_requested"):
        result["STEP_RESULT"] = "INVALID_INPUT"
    result["OVERALL_RESULT"] = "INVALID_INPUT"
    result["MESSAGE"] = " | ".join(invalid_row.get("warnings", []))
    return result


def write_result_csv(report_path, results):
    with open(report_path, "w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=_RESULT_COLUMNS,
            extrasaction="ignore",
        )
        writer.writeheader()
        writer.writerows(results)


# ---------------------------------------------------------------------------
# NX assembly and attribute helpers
# ---------------------------------------------------------------------------

def get_string_attribute(nx_object, attribute_name):
    if nx_object is None or not attribute_name:
        return ""
    try:
        return nx_object.GetStringAttribute(attribute_name).strip()
    except Exception:
        pass
    try:
        attribute = nx_object.GetUserAttribute(
            attribute_name,
            NXOpen.NXObject.AttributeType.String,
            -1,
        )
        return attribute.StringValue.strip()
    except Exception:
        return ""


def get_part_identity(part, component=None):
    part_number = (
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
    return normalize_text(part_number), normalize_text(revision)


def get_root_component(work_part):
    try:
        return work_part.ComponentAssembly.RootComponent
    except Exception:
        return None


def _object_identity(nx_object):
    try:
        tag = nx_object.Tag
        if tag is not None:
            return ("TAG", str(tag))
    except Exception:
        pass
    try:
        full_path = nx_object.FullPath
        if full_path:
            return ("PATH", str(full_path).upper())
    except Exception:
        pass
    return ("OBJECT", id(nx_object))


def _component_label(component):
    for attribute_name in ("DisplayName", "Name"):
        try:
            value = normalize_text(getattr(component, attribute_name))
            if value:
                return value
        except Exception:
            pass
    return "<unidentified component>"


def walk_components(component):
    """Iteratively yield occurrences, isolating GetChildren errors by branch."""
    if component is None:
        return
    stack = [component]
    while stack:
        current = stack.pop()
        yield current
        try:
            children = list(current.GetChildren())
        except Exception:
            continue
        for child in reversed(children):
            stack.append(child)


def build_loaded_part_map(work_part):
    exact_part_map = {}
    revisions_by_part_number = {}
    diagnostics = {
        "unresolved_components": [],
        "suppressed_component_count": 0,
        "prototype_error_count": 0,
        "duplicate_loaded_key_count": 0,
        "collision_keys": set(),
    }
    prototype_keys = {}

    def add_candidate(part, component=None):
        if part is None:
            return False

        prototype_identity = _object_identity(part)
        if prototype_identity in prototype_keys:
            return True

        part_number, revision = get_part_identity(part, component)
        if not part_number or not revision:
            return False

        key = (part_number.upper(), revision.upper())
        prototype_keys[prototype_identity] = key
        revisions_by_part_number.setdefault(key[0], set()).add(key[1])

        existing = exact_part_map.get(key)
        if existing is None:
            exact_part_map[key] = part
        elif _object_identity(existing) != prototype_identity:
            diagnostics["duplicate_loaded_key_count"] += 1
            diagnostics["collision_keys"].add(key)
        return True

    if INCLUDE_ROOT_WORK_PART:
        add_candidate(work_part)

    root = get_root_component(work_part)
    stack = [root]
    while stack:
        component = stack.pop()
        label = _component_label(component)

        if SKIP_SUPPRESSED_COMPONENTS:
            try:
                if component.IsSuppressed:
                    diagnostics["suppressed_component_count"] += 1
                    continue
            except Exception as error:
                diagnostics["unresolved_components"].append(
                    "{0}: could not read suppression state ({1})".format(label, error)
                )

        prototype = None
        try:
            prototype = component.Prototype
        except Exception as error:
            diagnostics["prototype_error_count"] += 1
            diagnostics["unresolved_components"].append(
                "{0}: could not access prototype ({1})".format(label, error)
            )

        if prototype is None:
            diagnostics["unresolved_components"].append(
                "{0}: no usable loaded prototype".format(label)
            )
        elif not add_candidate(prototype, component):
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

        for child in reversed(children):
            stack.append(child)

    return exact_part_map, revisions_by_part_number, diagnostics


def _safe_part_name(part, fallback="part"):
    try:
        full_path = part.FullPath
        if full_path:
            return os.path.splitext(os.path.basename(full_path))[0] or fallback
    except Exception:
        pass
    try:
        return normalize_text(part.Leaf) or fallback
    except Exception:
        return fallback


def count_drawing_sheets(part):
    try:
        return part.DrawingSheets.Count
    except Exception:
        pass
    try:
        return len(list(part.DrawingSheets))
    except Exception:
        return 0


# ---------------------------------------------------------------------------
# NX export helpers
# ---------------------------------------------------------------------------

def build_step_filename(part_number, revision):
    return "{0}_REV{1}.stp".format(
        clean_filename_token(part_number),
        clean_filename_token(revision, fallback=""),
    )


def build_pdf_filename(
    part,
    part_number,
    revision,
    sheet,
    sheet_index,
    sheet_count,
):
    drawing_number = get_string_attribute(part, "DRAWING_NUMBER")
    base_identifier = (
        drawing_number or normalize_text(part_number) or _safe_part_name(part)
    )
    base_name = "{0}_REV{1}".format(
        clean_filename_token(base_identifier),
        clean_filename_token(revision, fallback=""),
    )

    if sheet_count > 1:
        base_name += "_SHEET{0:02d}".format(sheet_index)
        try:
            sheet_name = normalize_text(sheet.Name)
        except Exception:
            sheet_name = ""
        if sheet_name:
            base_name += "_" + clean_filename_token(sheet_name, fallback="")

    return base_name + ".pdf"


def export_step(session, part, output_folder, part_number, revision):
    session.Parts.SetWork(part)
    output_path = os.path.join(
        output_folder,
        build_step_filename(part_number, revision),
    )
    if os.path.exists(output_path):
        raise RuntimeError("STEP output path already exists: {0}".format(output_path))

    exporter = session.DexManager.CreateStepCreator()
    try:
        # Define the source file and destination path explicitly
        exporter.InputFile = part.FullPath
        exporter.OutputFile = output_path
        
        # Configure common object filter types to export
        exporter.ObjectTypes.Solids = True
        exporter.ObjectTypes.Surfaces = True
        exporter.ObjectTypes.Curves = True
        
        if STEP_FORMAT != "AP214":
            raise RuntimeError("Unsupported STEP_FORMAT: {0}".format(STEP_FORMAT))
        exporter.ExportAs = NXOpen.StepCreator.ExportAsOption.Ap214
        
        # Force NX to wait until the STEP file finishes writing before continuing the loop
        exporter.ProcessHoldFlag = True
        exporter.Commit()
    finally:
        exporter.Destroy()

    if VERIFY_OUTPUT_FILES and not os.path.isfile(output_path):
        return {
            "result": "FAILED_NO_OUTPUT_FILE",
            "path": "",
            "size": "",
            "message": "STEP exporter committed but no output file was created",
        }

    try:
        file_size = os.path.getsize(output_path)
    except Exception:
        file_size = ""
    return {
        "result": "SUCCESS",
        "path": output_path,
        "size": file_size,
        "message": "",
    }


def _export_current_sheet(part, sheet, output_path):
    # Initialize PDF builder using the part's PlotManager
    pdf_builder = part.PlotManager.CreatePrintPdfbuilder()
    try:
        pdf_builder.Filename = output_path
        
        # TEAMCENTER FIX: Populate DatasetName so TC managed mode doesn't throw "Empty dataset name"
        base_name = os.path.splitext(os.path.basename(output_path))[0]
        try:
            pdf_builder.DatasetName = base_name
        except Exception:
            pass

        # Pass the drawing sheet to the builder
        pdf_builder.SourceBuilder.SetSheets([sheet])
        
        # Commit PDF creation
        pdf_builder.Commit()
    finally:
        pdf_builder.Destroy()


def export_pdfs(session, part, output_folder, part_number, revision):
    dwg_part = None
    expected_dwg_token = "{0}-{1}-DWG1".format(part_number, revision).upper()
    
    # -----------------------------------------------------------------------
    # 1. Search currently loaded session parts for the matching drawing dataset
    # -----------------------------------------------------------------------
    for loaded_part in session.Parts:
        try:
            identifiers = []
            if hasattr(loaded_part, "FullPath") and loaded_part.FullPath:
                identifiers.append(loaded_part.FullPath.upper())
            if hasattr(loaded_part, "Leaf") and loaded_part.Leaf:
                identifiers.append(loaded_part.Leaf.upper())
            if hasattr(loaded_part, "Name") and loaded_part.Name:
                identifiers.append(loaded_part.Name.upper())
                
            full_ident_str = " | ".join(identifiers)
            
            # Match if both the part number and "DWG" exist in the dataset identifier
            if (part_number.upper() in full_ident_str and "DWG" in full_ident_str):
                dwg_part = loaded_part
                break
        except Exception:
            continue

    # -----------------------------------------------------------------------
    # 2. If not already loaded, attempt to open the drawing dataset
    # -----------------------------------------------------------------------
    if dwg_part is None:
        try:
            if session.IsManagedMode:
                # Teamcenter Managed Mode database specs
                tc_specs = [
                    "@DB/{0}/{1}/dwg1".format(part_number, revision),
                    "@DB/{0}/{1}/UGPART/dwg1".format(part_number, revision),
                    "@DB/{0}/{1}/manifest".format(part_number, revision),
                ]
                for spec in tc_specs:
                    try:
                        dwg_part, _ = session.Parts.OpenBase(spec)
                        if dwg_part is not None:
                            break
                    except Exception:
                        continue
            else:
                # Native mode fallback paths
                three_d_dir = os.path.dirname(part.FullPath) if part.FullPath else ""
                native_paths = [
                    os.path.join(three_d_dir, "{0}-{1}-dwg1.prt".format(part_number, revision)),
                    os.path.join(three_d_dir, "{0}-{1}-dwg.prt".format(part_number, revision)),
                ]
                for npath in native_paths:
                    if os.path.exists(npath):
                        dwg_part, _ = session.Parts.OpenBase(npath)
                        if dwg_part is not None:
                            break
        except Exception as e:
            return {
                "result": "SKIPPED_NO_DRAWING",
                "paths": [],
                "message": "Error attempting to open drawing file: {0}".format(e),
            }

    # -----------------------------------------------------------------------
    # 3. If drawing dataset cannot be resolved anywhere
    # -----------------------------------------------------------------------
    if dwg_part is None:
        return {
            "result": "SKIPPED_NO_DRAWING",
            "paths": [],
            "message": "Drawing dataset '{0}' not found in loaded session or Teamcenter.".format(expected_dwg_token),
        }

    # Switch display window to the drawing dataset so sheets render correctly
    session.Parts.SetDisplay(
        dwg_part,
        False,
        True
    )

    try:
        drawing_sheets = dwg_part.DrawingSheets
        sheets = list(drawing_sheets)
    except Exception as error:
        raise RuntimeError("Unable to enumerate drawing sheets: {0}".format(error))

    if not sheets:
        return {
            "result": "SKIPPED_NO_DRAWING",
            "paths": [],
            "message": "No drawing sheets were found inside '{0}'".format(dwg_part.Leaf),
        }

    successful_paths = []
    failures = []
    sheet_count = len(sheets)
    for sheet_index, sheet in enumerate(sheets, start=1):
        pdf_name = build_pdf_filename(
            dwg_part,
            part_number,
            revision,
            sheet,
            sheet_index,
            sheet_count,
        )
        output_path = os.path.join(output_folder, pdf_name)

        try:
            if os.path.exists(output_path):
                raise RuntimeError(
                    "PDF output path already exists: {0}".format(output_path)
                )
            sheet.Open()
            _export_current_sheet(dwg_part, sheet, output_path)
            
            if VERIFY_OUTPUT_FILES and not os.path.isfile(output_path):
                failures.append(
                    {
                        "kind": "NO_OUTPUT",
                        "message": "Sheet {0}: exporter committed but no output file was created".format(
                            sheet_index
                        ),
                    }
                )
            else:
                successful_paths.append(output_path)
        except Exception as error:
            failures.append(
                {
                    "kind": "ERROR",
                    "message": "Sheet {0}: {1}".format(sheet_index, error),
                    "traceback": traceback.format_exc(),
                }
            )

    if len(successful_paths) == sheet_count:
        result = "SUCCESS"
    elif successful_paths:
        result = "PARTIAL_SUCCESS"
    elif failures and all(item["kind"] == "NO_OUTPUT" for item in failures):
        result = "FAILED_NO_OUTPUT_FILE"
    else:
        result = "FAILED"

    return {
        "result": result,
        "paths": successful_paths,
        "message": " | ".join(item["message"] for item in failures),
        "failures": failures,
    }


def determine_overall_result(
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

    if requested_results and all(result == "SUCCESS" for result in requested_results):
        return "SUCCESS"

    has_success = any(
        result in ("SUCCESS", "PARTIAL_SUCCESS") for result in requested_results
    )
    if has_success:
        return "PARTIAL_SUCCESS"
    return "FAILED"


def _set_resolution_failure(result, status, loaded_revision, message):
    if result["PDF_REQUESTED"] == "YES":
        result["PDF_RESULT"] = status
    if result["STEP_REQUESTED"] == "YES":
        result["STEP_RESULT"] = status
    result["LOADED_REVISION"] = loaded_revision
    result["OVERALL_RESULT"] = status
    result["MESSAGE"] = message


def _process_instruction(
    session,
    instruction,
    exact_part_map,
    revisions_by_part_number,
    collision_keys,
    pdf_folder,
    step_folder,
    run_timestamp,
    log_buffer,
):
    started = datetime.datetime.now()
    result = _new_result(run_timestamp, instruction)
    messages = list(instruction.get("warnings", []))
    key = instruction["normalized_key"]

    part = exact_part_map.get(key)
    if part is None:
        loaded_revisions = sorted(revisions_by_part_number.get(key[0], set()))
        if loaded_revisions:
            loaded_text = ";".join(loaded_revisions)
            message = "Requested revision {0}; loaded revisions: {1}".format(
                instruction["revision"],
                ", ".join(loaded_revisions),
            )
            _set_resolution_failure(
                result,
                "REVISION_MISMATCH",
                loaded_text,
                message,
            )
        else:
            _set_resolution_failure(
                result,
                "NOT_FOUND",
                "",
                "Part number was not present in the loaded assembly",
            )
        if messages:
            result["MESSAGE"] = " | ".join(messages + [result["MESSAGE"]])
        result["DURATION_SECONDS"] = "{0:.3f}".format(
            (datetime.datetime.now() - started).total_seconds()
        )
        return result

    result["LOADED_REVISION"] = key[1]
    if key in collision_keys:
        _append_unique(
            messages,
            "Multiple loaded prototypes share this key; the first usable prototype was used",
        )

    try:
        if instruction["pdf_requested"]:
            try:
                pdf_export = export_pdfs(
                    session,
                    part,
                    pdf_folder,
                    instruction["part_number"],
                    instruction["revision"],
                )
                result["PDF_RESULT"] = pdf_export["result"]
                result["PDF_FILE_COUNT"] = len(pdf_export["paths"])
                result["PDF_FILES"] = ";".join(pdf_export["paths"])
                _append_unique(messages, pdf_export.get("message", ""))
                for failure in pdf_export.get("failures", []):
                    if failure.get("traceback"):
                        log_line(session, failure["traceback"], log_buffer)
            except Exception as error:
                result["PDF_RESULT"] = "FAILED"
                _append_unique(messages, "PDF export failed: {0}".format(error))
                log_line(session, traceback.format_exc(), log_buffer)

        if instruction["step_requested"]:
            try:
                step_export = export_step(
                    session,
                    part,
                    step_folder,
                    instruction["part_number"],
                    instruction["revision"],
                )
                result["STEP_RESULT"] = step_export["result"]
                result["STEP_FILE"] = step_export["path"]
                _append_unique(messages, step_export.get("message", ""))
                if step_export.get("size") != "":
                    _append_unique(
                        messages,
                        "STEP file size: {0} bytes".format(step_export["size"]),
                    )
            except Exception as error:
                result["STEP_RESULT"] = "FAILED"
                _append_unique(messages, "STEP export failed: {0}".format(error))
                log_line(session, traceback.format_exc(), log_buffer)
    except Exception as error:
        if result["PDF_RESULT"] == "PENDING":
            result["PDF_RESULT"] = "FAILED"
        if result["STEP_RESULT"] == "PENDING":
            result["STEP_RESULT"] = "FAILED"
        _append_unique(messages, "Part processing failed: {0}".format(error))
        log_line(session, traceback.format_exc(), log_buffer)

    if result["PDF_RESULT"] == "PENDING":
        result["PDF_RESULT"] = "FAILED"
    if result["STEP_RESULT"] == "PENDING":
        result["STEP_RESULT"] = "FAILED"

    result["OVERALL_RESULT"] = determine_overall_result(
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


# ---------------------------------------------------------------------------
# Journal entry point
# ---------------------------------------------------------------------------

def _restore_original_parts(session, original_display_part, original_work_part, log_buffer):
    if original_display_part is not None:
        try:
            session.Parts.SetDisplay(
                original_display_part,
                False,
                True
            )
        except Exception as error:
            log_line(
                session,
                "ERROR: Could not restore original display part: {0}".format(error),
                log_buffer,
            )

    if original_work_part is not None:
        try:
            session.Parts.SetWork(original_work_part)
        except Exception as error:
            log_line(
                session,
                "ERROR: Could not restore original work part: {0}".format(error),
                log_buffer,
            )


def _log_final_summary(session, results, report_path, diagnostics, request_count, log_buffer):
    counts = {}
    for result in results:
        status = result["OVERALL_RESULT"]
        counts[status] = counts.get(status, 0) + 1

    pdf_file_count = sum(int(result["PDF_FILE_COUNT"]) for result in results)
    step_file_count = sum(1 for result in results if result["STEP_FILE"])

    lines = [
        "Export complete",
        "Requested unique parts: {0}".format(request_count),
        "Full success: {0}".format(counts.get("SUCCESS", 0)),
        "Partial success: {0}".format(counts.get("PARTIAL_SUCCESS", 0)),
        "Not found: {0}".format(counts.get("NOT_FOUND", 0)),
        "Revision mismatch: {0}".format(counts.get("REVISION_MISMATCH", 0)),
        "Invalid input: {0}".format(counts.get("INVALID_INPUT", 0)),
        "Failed: {0}".format(counts.get("FAILED", 0)),
        "PDF files: {0}".format(pdf_file_count),
        "STEP files: {0}".format(step_file_count),
        "Suppressed components skipped: {0}".format(
            diagnostics.get("suppressed_component_count", 0)
        ),
        "Unresolved component diagnostics: {0}".format(
            len(diagnostics.get("unresolved_components", []))
        ),
        "Prototype access errors: {0}".format(
            diagnostics.get("prototype_error_count", 0)
        ),
        "Loaded-key collisions: {0}".format(
            diagnostics.get("duplicate_loaded_key_count", 0)
        ),
        "Result report: {0}".format(report_path),
    ]
    for line in lines:
        log_line(session, line, log_buffer)


def main():
    session = NXOpen.Session.GetSession()
    log_buffer = []
    run_folders = None
    report_path = ""
    results = []
    report_written = False
    diagnostics = {
        "unresolved_components": [],
        "suppressed_component_count": 0,
        "prototype_error_count": 0,
        "duplicate_loaded_key_count": 0,
        "collision_keys": set(),
    }

    log_line(session, "DataPack PDF/STEP Export", log_buffer)

    try:
        io_root = resolve_io_root()
        csv_path = resolve_input_csv(io_root)
        log_line(session, "Input: {0}".format(csv_path), log_buffer)

        work_part = session.Parts.Work
        if work_part is None:
            log_line(session, "ERROR: No work part is loaded.", log_buffer)
            return

        root_component = get_root_component(work_part)
        if root_component is None:
            log_line(
                session,
                "ERROR: The active work part is not an assembly.",
                log_buffer,
            )
            return

        if not os.path.isfile(csv_path):
            log_line(
                session,
                "ERROR: Input CSV not found: {0}".format(csv_path),
                log_buffer,
            )
            return

        try:
            scope = read_export_scope(csv_path)
        except Exception as error:
            log_line(
                session,
                "ERROR: Input CSV validation failed: {0}".format(error),
                log_buffer,
            )
            return

        instructions = scope["instructions"]
        for warning in scope["header_warnings"]:
            log_line(session, "WARNING: {0}".format(warning), log_buffer)
        log_line(
            session,
            "Valid unique requests: {0}".format(len(instructions)),
            log_buffer,
        )
        log_line(
            session,
            "Ignored rows with no requested output: {0}".format(
                scope["ignored_row_count"]
            ),
            log_buffer,
        )

        exact_part_map = {}
        revisions_by_part_number = {}
        if instructions:
            exact_part_map, revisions_by_part_number, diagnostics = (
                build_loaded_part_map(work_part)
            )
        log_line(
            session,
            "Loaded assembly keys: {0}".format(len(exact_part_map)),
            log_buffer,
        )

        run_timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        try:
            run_folders = _create_run_folders(io_root, run_timestamp)
        except FileExistsError:
            log_line(
                session,
                "ERROR: Output run already exists for timestamp {0}; no files were overwritten.".format(
                    run_timestamp
                ),
                log_buffer,
            )
            return

        report_path = os.path.join(
            run_folders["reports"],
            "EXPORT_RESULT_{0}.csv".format(run_timestamp),
        )
        log_path = os.path.join(
            run_folders["logs"],
            "EXPORT_LOG_{0}.txt".format(run_timestamp),
        )
        log_line(session, "Output: {0}".format(run_folders["run"]), log_buffer)

        for invalid_row in scope["invalid_rows"]:
            results.append(_invalid_result(run_timestamp, invalid_row))

        original_work_part = session.Parts.Work
        original_display_part = session.Parts.Display
        try:
            total = len(instructions)
            for index, instruction in enumerate(instructions, start=1):
                log_line(
                    session,
                    "[{0}/{1}] {2} / {3}".format(
                        index,
                        total,
                        instruction["part_number"],
                        instruction["revision"],
                    ),
                    log_buffer,
                )
                result = _process_instruction(
                    session,
                    instruction,
                    exact_part_map,
                    revisions_by_part_number,
                    diagnostics["collision_keys"],
                    run_folders["pdf"],
                    run_folders["step"],
                    run_timestamp,
                    log_buffer,
                )
                results.append(result)
                log_line(
                    session,
                    "  PDF: {0}{1}".format(
                        result["PDF_RESULT"],
                        " ({0} file(s))".format(result["PDF_FILE_COUNT"])
                        if result["PDF_REQUESTED"] == "YES"
                        else "",
                    ),
                    log_buffer,
                )
                log_line(
                    session,
                    "  STEP: {0}".format(result["STEP_RESULT"]),
                    log_buffer,
                )
        finally:
            _restore_original_parts(
                session,
                original_display_part,
                original_work_part,
                log_buffer,
            )

        write_result_csv(report_path, results)
        report_written = True

        for collision_key in sorted(diagnostics["collision_keys"]):
            log_line(
                session,
                "WARNING: Loaded-key collision for {0} / {1}; the first usable prototype was retained.".format(
                    collision_key[0],
                    collision_key[1],
                ),
                log_buffer,
            )
        for diagnostic in diagnostics["unresolved_components"]:
            log_line(session, "WARNING: {0}".format(diagnostic), log_buffer)

        _log_final_summary(
            session,
            results,
            report_path,
            diagnostics,
            len(instructions),
            log_buffer,
        )

        _write_text_log(log_path, log_buffer)
    except Exception:
        log_line(session, "ERROR: Unhandled journal exception.", log_buffer)
        log_line(session, traceback.format_exc(), log_buffer)
    finally:
        if run_folders is not None:
            if report_path and not report_written:
                try:
                    write_result_csv(report_path, results)
                except Exception as error:
                    log_line(
                        session,
                        "ERROR: Could not write result report: {0}".format(error),
                        log_buffer,
                    )
            try:
                fallback_log_path = os.path.join(
                    run_folders["logs"],
                    "EXPORT_LOG_{0}.txt".format(
                        os.path.basename(run_folders["run"])
                    ),
                )
                _write_text_log(fallback_log_path, log_buffer)
            except Exception:
                pass


if __name__ == "__main__":
    main()
