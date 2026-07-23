"""
Journal 10 - Teamcenter STEP Zero-Geometry Diagnostic

This journal diagnoses why Journal 07 can create a very small AP214 STEP file
whose translator log reports zero input solids even though NX reports a solid
body and a manual STEP export succeeds. It runs a table-driven export matrix
that isolates source identity, layer filtering, body representation, explicit
selection, and display-state effects.

The journal does not save or modify Teamcenter data. It restores the original
display/work parts and closes only the top-level part that it opened.

Defaults:
    Part number: 264MN020016A01
    Revision:    A

Optional environment overrides:
    NX_STEP_DIAG_PART_NO
    NX_STEP_DIAG_PART_REV
    NX_STEP_DIAG_KEEP_OPEN=1
    NX_JOURNALS_IO_DIR=<output root>

Output:
    Desktop\\NX_STEP_DIAGNOSTIC\\<timestamp>\\

Target: NX 2312 embedded Python 3.10
Run via: NX > Tools > Journal > Play
"""

import csv
import datetime
import os
import re
import traceback

import NXOpen
import NXOpen.Layer


DEFAULT_PART_NUMBER = "264MN020016A01"
DEFAULT_REVISION = "A"
OUTPUT_FOLDER_NAME = "NX_STEP_DIAGNOSTIC"
KEEP_OPEN_AFTER_TEST = False
MAX_COMPONENT_OCCURRENCES = 20000
ALL_LAYER_MASK = "1-256"

OBJECT_TYPE_NAMES = (
    "Solids",
    "Surfaces",
    "Curves",
    "ConvergentBodies",
    "FacetBodies",
    "Tessellation",
    "Structures",
)

BODY_GEOMETRY_TOKENS = (
    "MANIFOLD_SOLID_BREP",
    "BREP_WITH_VOIDS",
    "FACETED_BREP",
    "ADVANCED_BREP_SHAPE_REPRESENTATION",
    "SHELL_BASED_SURFACE_MODEL",
    "CLOSED_SHELL",
    "OPEN_SHELL",
    "ADVANCED_FACE",
    "TESSELLATED_SHAPE_REPRESENTATION",
)

ASSEMBLY_TOKENS = (
    "NEXT_ASSEMBLY_USAGE_OCCURRENCE",
    "CONTEXT_DEPENDENT_SHAPE_REPRESENTATION",
    "REPRESENTATION_RELATIONSHIP_WITH_TRANSFORMATION",
    "MAPPED_ITEM",
)

CSV_COLUMNS = (
    "TRIAL",
    "LABEL",
    "PHASE",
    "COMMIT_RESULT",
    "ERROR",
    "OUTPUT_FILE",
    "FORMAT",
    "FILE_EXISTS",
    "FILE_SIZE_BYTES",
    "DATA_ENTITY_LINES",
    "BODY_GEOMETRY_SIGNATURES",
    "ASSEMBLY_SIGNATURES",
    "HAS_BODY_GEOMETRY",
    "TRANSLATOR_SOLIDS_INPUT",
    "TRANSLATOR_SOLIDS_PROCESSED",
    "TRANSLATOR_SOLIDS_AS_SHEETS",
    "TRANSLATOR_SOLIDS_NOT_PROCESSED",
    "TRANSLATOR_UG_BODY_COUNT",
    "TRANSLATOR_LOG",
    "BUILDER_VALIDATE_RESULT",
    "EXPORT_FROM",
    "INPUT_FILE",
    "LAYER_MASK",
    "FILE_SAVE_FLAG",
    "SETTINGS_FILE",
    "SELECTION_SCOPE",
    "SELECTION_SIZE",
    "SOLIDS_ENABLED",
    "SURFACES_ENABLED",
    "CURVES_ENABLED",
    "CONVERGENT_BODIES_ENABLED",
    "FACET_BODIES_ENABLED",
    "TESSELLATION_ENABLED",
    "STRUCTURES_ENABLED",
    "PART_LOAD_STATE",
    "PART_IS_FULLY_LOADED",
    "DIRECT_BODY_COUNT",
    "DIRECT_SOLID_BODY_COUNT",
    "DIRECT_SHEET_BODY_COUNT",
    "COMPONENT_OCCURRENCE_COUNT",
    "UNIQUE_PROTOTYPE_COUNT",
    "DESCENDANT_BODY_OCCURRENCE_COUNT",
    "DESCENDANT_SOLID_BODY_OCCURRENCE_COUNT",
)

BODY_CSV_COLUMNS = (
    "INDEX",
    "TAG",
    "RUNTIME_TYPE",
    "NAME",
    "JOURNAL_IDENTIFIER",
    "OWNING_PART",
    "OWNING_PART_MATCHES_TARGET",
    "LAYER",
    "LAYER_STATE",
    "IS_BLANKED",
    "IS_SOLID_BODY",
    "IS_SHEET_BODY",
    "IS_CONVERGENT_BODY",
    "FACE_COUNT",
    "EDGE_COUNT",
    "FACET_COUNT",
    "VERTEX_COUNT",
    "REFERENCE_SETS",
)


def clean(value):
    if value is None:
        return ""
    try:
        return str(value).strip()
    except Exception:
        return ""


def env_bool(name, default=False):
    value = clean(os.environ.get(name))
    if not value:
        return default
    return value.upper() in ("1", "TRUE", "YES", "Y", "ON")


def type_name(value):
    if value is None:
        return "None"
    return "{0}.{1}".format(type(value).__module__, type(value).__name__)


def enum_text(value):
    if value is None:
        return ""
    try:
        return clean(value.name)
    except Exception:
        return clean(value)


def exception_details(error):
    details = [
        "{0}: {1}".format(type_name(error), clean(error) or "<no message>")
    ]
    for name in ("ErrorCode", "error_code", "Code", "code"):
        try:
            value = getattr(error, name)
            if not callable(value) and clean(value):
                details.append("{0}={1}".format(name, clean(value)))
        except Exception:
            pass
    return "; ".join(details)


def dispose(value):
    if value is not None:
        try:
            value.Dispose()
        except Exception:
            pass


def unwrap_open_result(value):
    if isinstance(value, (tuple, list)):
        part = value[0] if value else None
        status = value[1] if len(value) > 1 else None
        return part, status
    return value, None


def safe_property(value, name, fallback=""):
    if value is None:
        return fallback
    try:
        result = getattr(value, name)
        if callable(result):
            result = result()
        return result
    except Exception:
        return fallback


def object_key(value):
    if value is None:
        return ("NONE", "")
    try:
        return ("TAG", str(value.Tag))
    except Exception:
        pass
    for name in ("JournalIdentifier", "FullPath", "Name"):
        text = clean(safe_property(value, name))
        if text:
            return (name.upper(), text.upper())
    return ("OBJECT", str(id(value)))


def loaded_parts(session):
    try:
        return list(session.Parts)
    except Exception:
        return []


def collection_values(collection):
    if collection is None:
        return []
    try:
        return list(collection)
    except Exception:
        try:
            return list(collection.ToArray())
        except Exception:
            return []


def part_full_path(part):
    return clean(safe_property(part, "FullPath"))


def part_name(part):
    for name in ("Name", "Leaf", "JournalIdentifier", "FullPath"):
        value = clean(safe_property(part, name))
        if value:
            return value
    return "part"


def part_load_state(session, part):
    path = part_full_path(part)
    if not path:
        return ""
    try:
        return enum_text(session.Parts.GetPartLoadStateOfFileName(path))
    except Exception as error:
        return "ERROR: {0}".format(exception_details(error))


def is_fully_loaded(part):
    for name in ("IsFullyLoaded", "FullyLoaded"):
        try:
            value = getattr(part, name)
            if callable(value):
                value = value()
            return bool(value)
        except Exception:
            pass
    return ""


def body_counts(part):
    record = {
        "body_count": 0,
        "solid_body_count": 0,
        "sheet_body_count": 0,
        "unknown_body_type_count": 0,
        "body_access_error": "",
    }
    try:
        bodies = collection_values(part.Bodies)
    except Exception as error:
        record["body_access_error"] = exception_details(error)
        return record

    record["body_count"] = len(bodies)
    for body in bodies:
        try:
            is_solid = body.IsSolidBody
            if callable(is_solid):
                is_solid = is_solid()
            if bool(is_solid):
                record["solid_body_count"] += 1
            else:
                record["sheet_body_count"] += 1
        except Exception:
            record["unknown_body_type_count"] += 1
    return record


def safe_method_count(value, method_name):
    try:
        result = getattr(value, method_name)()
        try:
            return len(result)
        except Exception:
            pass
        return int(result)
    except Exception:
        return ""


def reference_sets_for_body(part, body):
    names = []
    try:
        reference_sets = list(part.GetAllReferenceSets())
    except Exception:
        reference_sets = []

    body_key = object_key(body)
    for reference_set in reference_sets:
        members = []
        for method_name in (
            "AskAllDirectMembers",
            "AskMembersInReferenceSet",
        ):
            try:
                members = list(getattr(reference_set, method_name)())
                break
            except Exception:
                pass
        if any(object_key(member) == body_key for member in members):
            names.append(
                clean(safe_property(reference_set, "Name"))
                or clean(safe_property(reference_set, "JournalIdentifier"))
                or "<unnamed>"
            )
    return names


def direct_bodies(part):
    try:
        return collection_values(part.Bodies)
    except Exception:
        return []


def body_diagnostics(part):
    records = []
    target_key = object_key(part)
    for index, body in enumerate(direct_bodies(part), start=1):
        layer = safe_property(body, "Layer", "")
        try:
            layer_state = enum_text(part.Layers.GetState(int(layer)))
        except Exception:
            layer_state = ""
        owning_part = safe_property(body, "OwningPart", None)
        records.append({
            "index": index,
            "body": body,
            "tag": clean(safe_property(body, "Tag")),
            "runtime_type": type_name(body),
            "name": clean(safe_property(body, "Name")),
            "journal_identifier": clean(
                safe_property(body, "JournalIdentifier")
            ),
            "owning_part": part_name(owning_part) if owning_part else "",
            "owning_part_matches_target": (
                object_key(owning_part) == target_key
                if owning_part is not None else False
            ),
            "layer": layer,
            "layer_state": layer_state,
            "is_blanked": safe_property(body, "IsBlanked", ""),
            "is_solid_body": safe_property(body, "IsSolidBody", ""),
            "is_sheet_body": safe_property(body, "IsSheetBody", ""),
            "is_convergent_body": safe_property(
                body,
                "IsConvergentBody",
                "",
            ),
            "face_count": safe_method_count(body, "GetFaces"),
            "edge_count": safe_method_count(body, "GetEdges"),
            "facet_count": safe_method_count(
                body,
                "GetNumberOfFacets",
            ),
            "vertex_count": safe_method_count(
                body,
                "GetNumberOfVertices",
            ),
            "reference_sets": reference_sets_for_body(part, body),
        })
    return records


def log_body_diagnostics(logger, records):
    logger.section("Direct body diagnostics")
    if not records:
        logger.write("No direct bodies were available.")
        return
    for record in records:
        logger.write(
            "Body {0}: tag={1}, type={2}".format(
                record["index"],
                record["tag"],
                record["runtime_type"],
            )
        )
        logger.write(
            "  solid={0}, sheet={1}, convergent={2}".format(
                record["is_solid_body"],
                record["is_sheet_body"],
                record["is_convergent_body"],
            )
        )
        logger.write(
            "  layer={0} ({1}), blanked={2}".format(
                record["layer"],
                record["layer_state"],
                record["is_blanked"],
            )
        )
        logger.write(
            "  owning part={0}, owner matches target={1}".format(
                record["owning_part"],
                record["owning_part_matches_target"],
            )
        )
        logger.write(
            "  faces={0}, edges={1}, facets={2}, vertices={3}".format(
                record["face_count"],
                record["edge_count"],
                record["facet_count"],
                record["vertex_count"],
            )
        )
        logger.write(
            "  reference sets: {0}".format(
                ", ".join(record["reference_sets"]) or "<none>"
            )
        )


def component_children(component):
    try:
        return list(component.GetChildren())
    except Exception:
        return []


def component_prototype(component):
    try:
        return component.Prototype
    except Exception:
        return None


def assembly_counts(part):
    result = {
        "has_root_component": False,
        "component_occurrence_count": 0,
        "suppressed_occurrence_count": 0,
        "unique_prototype_count": 0,
        "loaded_prototype_count": 0,
        "descendant_body_occurrence_count": 0,
        "descendant_solid_body_occurrence_count": 0,
        "descendant_sheet_body_occurrence_count": 0,
        "component_limit_reached": False,
        "component_access_errors": 0,
    }

    try:
        root = part.ComponentAssembly.RootComponent
    except Exception:
        root = None

    if root is None:
        return result

    result["has_root_component"] = True
    stack = list(reversed(component_children(root)))
    seen_occurrences = set()
    seen_prototypes = set()

    while stack:
        if result["component_occurrence_count"] >= MAX_COMPONENT_OCCURRENCES:
            result["component_limit_reached"] = True
            break

        component = stack.pop()
        occurrence_key = object_key(component)
        if occurrence_key in seen_occurrences:
            continue
        seen_occurrences.add(occurrence_key)
        result["component_occurrence_count"] += 1

        try:
            if bool(component.IsSuppressed):
                result["suppressed_occurrence_count"] += 1
        except Exception:
            pass

        prototype = component_prototype(component)
        if prototype is None:
            result["component_access_errors"] += 1
        else:
            prototype_key = object_key(prototype)
            seen_prototypes.add(prototype_key)
            counts = body_counts(prototype)
            if not counts["body_access_error"]:
                result["loaded_prototype_count"] += 1
            result["descendant_body_occurrence_count"] += counts["body_count"]
            result["descendant_solid_body_occurrence_count"] += counts[
                "solid_body_count"
            ]
            result["descendant_sheet_body_occurrence_count"] += counts[
                "sheet_body_count"
            ]

        children = component_children(component)
        if children:
            stack.extend(reversed(children))

    result["unique_prototype_count"] = len(seen_prototypes)
    return result


def part_snapshot(session, part, label):
    direct = body_counts(part)
    assembly = assembly_counts(part)
    return {
        "label": label,
        "runtime_type": type_name(part),
        "name": part_name(part),
        "journal_identifier": clean(safe_property(part, "JournalIdentifier")),
        "full_path": part_full_path(part),
        "load_state": part_load_state(session, part),
        "is_fully_loaded": is_fully_loaded(part),
        **direct,
        **assembly,
    }


def log_snapshot(logger, snapshot):
    logger.write("Snapshot: {0}".format(snapshot["label"]))
    logger.write("  Runtime type: {0}".format(snapshot["runtime_type"]))
    logger.write("  Name: {0}".format(snapshot["name"]))
    logger.write("  JournalIdentifier: {0}".format(
        snapshot["journal_identifier"]
    ))
    logger.write("  FullPath: {0}".format(snapshot["full_path"]))
    logger.write("  Load state: {0}".format(snapshot["load_state"]))
    logger.write("  Is fully loaded: {0}".format(
        snapshot["is_fully_loaded"]
    ))
    logger.write(
        "  Direct bodies: {0} (solid={1}, sheet={2}, unknown={3})".format(
            snapshot["body_count"],
            snapshot["solid_body_count"],
            snapshot["sheet_body_count"],
            snapshot["unknown_body_type_count"],
        )
    )
    if snapshot["body_access_error"]:
        logger.write("  Body access error: {0}".format(
            snapshot["body_access_error"]
        ))
    logger.write(
        "  Assembly: root={0}, occurrences={1}, unique prototypes={2}, "
        "loaded prototypes={3}".format(
            snapshot["has_root_component"],
            snapshot["component_occurrence_count"],
            snapshot["unique_prototype_count"],
            snapshot["loaded_prototype_count"],
        )
    )
    logger.write(
        "  Descendant occurrence bodies: {0} (solid={1}, sheet={2})".format(
            snapshot["descendant_body_occurrence_count"],
            snapshot["descendant_solid_body_occurrence_count"],
            snapshot["descendant_sheet_body_occurrence_count"],
        )
    )
    if snapshot["component_limit_reached"]:
        logger.write(
            "  WARNING: component traversal stopped at {0} occurrences.".format(
                MAX_COMPONENT_OCCURRENCES
            )
        )


def load_status_lines(status):
    lines = []
    if status is None:
        return lines
    try:
        count = int(status.NumberUnloadedParts)
    except Exception:
        count = 0
    for index in range(count):
        try:
            lines.append(
                "{0}: status={1}; {2}".format(
                    clean(status.GetPartName(index)),
                    clean(status.GetStatus(index)),
                    clean(status.GetStatusDescription(index)),
                )
            )
        except Exception as error:
            lines.append(
                "Could not inspect load-status entry {0}: {1}".format(
                    index,
                    exception_details(error),
                )
            )
    return lines


def ensure_fully_loaded(session, part, logger):
    result = {
        "method": "EnsurePartsLoadedFully(includeChildren=True)",
        "status": "NOT_RUN",
        "messages": [],
    }
    status = None
    try:
        status = session.Parts.EnsurePartsLoadedFully([part], True)
        result["status"] = "SUCCESS"
        result["messages"] = load_status_lines(status)
    except Exception as error:
        result["status"] = "FAILED"
        result["messages"].append(exception_details(error))
    finally:
        dispose(status)

    logger.write("Full-load request: {0}".format(result["status"]))
    for message in result["messages"]:
        logger.write("  {0}".format(message))

    if result["status"] == "SUCCESS":
        return result

    fallback_status = None
    result["method"] += " -> part.LoadFully() fallback"
    try:
        fallback_status = part.LoadFully()
        result["status"] = "SUCCESS_FALLBACK"
        fallback_lines = load_status_lines(fallback_status)
        result["messages"].extend(fallback_lines)
        logger.write("Fallback part.LoadFully(): SUCCESS")
        for message in fallback_lines:
            logger.write("  {0}".format(message))
    except Exception as error:
        result["status"] = "FAILED"
        message = exception_details(error)
        result["messages"].append(message)
        logger.write("Fallback part.LoadFully(): FAILED")
        logger.write("  {0}".format(message))
    finally:
        dispose(fallback_status)
    return result


def set_display(session, part):
    result = session.Parts.SetDisplay(part, False, True)
    if isinstance(result, (tuple, list)) and len(result) > 1:
        dispose(result[1])


def activate_part(session, part):
    set_display(session, part)
    session.Parts.SetWork(part)


def capture_body_display_state(part, bodies):
    states = []
    for body in bodies:
        layer = safe_property(body, "Layer", "")
        try:
            layer_state = part.Layers.GetState(int(layer))
        except Exception:
            layer_state = None
        states.append({
            "body": body,
            "layer": layer,
            "layer_state": layer_state,
            "is_blanked": bool(safe_property(body, "IsBlanked", False)),
        })
    return states


def make_bodies_visible_and_selectable(part, states, logger):
    for state in states:
        body = state["body"]
        layer = state["layer"]
        if state["is_blanked"]:
            body.Unblank()
            logger.write(
                "Temporarily unblanked body tag={0}".format(
                    clean(safe_property(body, "Tag"))
                )
            )
        if (
            layer != ""
            and enum_text(state["layer_state"]) != "WorkLayer"
        ):
            part.Layers.SetState(
                int(layer),
                NXOpen.Layer.State.Selectable,
            )
            logger.write(
                "Temporarily made layer {0} selectable.".format(layer)
            )


def restore_body_display_state(part, states, logger):
    if not states:
        return
    logger.write("Restoring temporary body visibility/layer state...")
    restored_layers = set()
    for state in states:
        body = state["body"]
        try:
            if state["is_blanked"]:
                body.Blank()
            else:
                body.Unblank()
        except Exception as error:
            logger.write(
                "  WARNING: body display restore failed: {0}".format(
                    exception_details(error)
                )
            )

        layer = state["layer"]
        layer_state = state["layer_state"]
        if layer == "" or layer_state is None or layer in restored_layers:
            continue
        try:
            part.Layers.SetState(int(layer), layer_state)
            restored_layers.add(layer)
        except Exception as error:
            logger.write(
                "  WARNING: layer {0} restore failed: {1}".format(
                    layer,
                    exception_details(error),
                )
            )


def output_base_folder():
    configured = clean(os.environ.get("NX_JOURNALS_IO_DIR"))
    if configured:
        return os.path.abspath(os.path.expanduser(configured))
    profile = clean(os.environ.get("USERPROFILE"))
    if profile:
        return os.path.join(profile, "Desktop")
    home = os.path.expanduser("~")
    if home and home != "~":
        return os.path.join(home, "Desktop")
    return os.getcwd()


def create_run_folder():
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    folder = os.path.join(
        output_base_folder(),
        OUTPUT_FOLDER_NAME,
        timestamp,
    )
    os.makedirs(folder, exist_ok=False)
    return folder, timestamp


def inspect_step_file(path):
    result = {
        "exists": os.path.isfile(path),
        "size_bytes": "",
        "data_entity_lines": 0,
        "body_token_counts": {
            token: 0 for token in BODY_GEOMETRY_TOKENS
        },
        "assembly_token_counts": {
            token: 0 for token in ASSEMBLY_TOKENS
        },
        "body_geometry_signatures": 0,
        "assembly_signatures": 0,
        "has_body_geometry": False,
        "inspection_error": "",
    }
    if not result["exists"]:
        return result

    try:
        result["size_bytes"] = os.path.getsize(path)
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
                if in_data and stripped.startswith("#") and "=" in stripped:
                    result["data_entity_lines"] += 1
                for token in BODY_GEOMETRY_TOKENS:
                    result["body_token_counts"][token] += upper.count(token)
                for token in ASSEMBLY_TOKENS:
                    result["assembly_token_counts"][token] += upper.count(token)

        result["body_geometry_signatures"] = sum(
            result["body_token_counts"].values()
        )
        result["assembly_signatures"] = sum(
            result["assembly_token_counts"].values()
        )
        result["has_body_geometry"] = (
            result["body_geometry_signatures"] > 0
        )
    except Exception as error:
        result["inspection_error"] = exception_details(error)
    return result


def parse_translator_log(output_path):
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
        for name in os.listdir(folder):
            if name.lower().endswith(".log"):
                candidates.append(os.path.join(folder, name))
    except Exception:
        return result

    preferred = []
    for path in candidates:
        try:
            with open(path, "r", encoding="utf-8", errors="replace") as handle:
                content = handle.read()
            if (
                output_name in content.upper()
                or output_stem in os.path.basename(path).upper()
            ):
                preferred.append((path, content))
        except Exception:
            pass

    def modified_time(item):
        try:
            return os.path.getmtime(item[0])
        except Exception:
            return 0

    preferred.sort(key=modified_time, reverse=True)
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

    for path, content in preferred:
        matched = False
        parsed = {}
        for key, pattern in patterns.items():
            match = re.search(pattern, content, re.IGNORECASE)
            if match:
                parsed[key] = match.group(1)
                matched = True
        body_matches = re.findall(
            r"SUMMARY-\s*Body\s*\.*\s*:\s*(\d+)",
            content,
            re.IGNORECASE,
        )
        if body_matches:
            parsed["ug_body_count"] = str(sum(
                int(value) for value in body_matches
            ))
            matched = True
        if matched:
            result.update(parsed)
            result["path"] = path
            return result
    return result


def selection_objects(selector):
    try:
        return list(selector.GetArray())
    except Exception:
        return []


def builder_state(creator):
    selector = creator.ExportSelectionBlock
    selection = selector.SelectionComp
    object_types = {}
    for name in OBJECT_TYPE_NAMES:
        object_types[name] = safe_property(creator.ObjectTypes, name, "")
    return {
        "export_from": enum_text(safe_property(creator, "ExportFrom")),
        "input_file": clean(safe_property(creator, "InputFile")),
        "layer_mask": clean(safe_property(creator, "LayerMask")),
        "file_save_flag": safe_property(creator, "FileSaveFlag", ""),
        "settings_file": clean(safe_property(creator, "SettingsFile")),
        "selection_scope": enum_text(
            safe_property(selector, "SelectionScope")
        ),
        "selection_size": safe_property(selection, "Size", ""),
        "selection_tags": [
            clean(safe_property(value, "Tag"))
            for value in selection_objects(selection)
        ],
        "object_types": object_types,
    }


def log_builder_state(logger, title, state):
    logger.write(title)
    logger.write(
        "  ExportFrom={0}, InputFile={1}".format(
            state["export_from"],
            state["input_file"] or "<empty>",
        )
    )
    logger.write(
        "  LayerMask={0}, FileSaveFlag={1}, SettingsFile={2}".format(
            state["layer_mask"] or "<empty>",
            state["file_save_flag"],
            state["settings_file"] or "<empty>",
        )
    )
    logger.write(
        "  SelectionScope={0}, SelectionSize={1}, tags={2}".format(
            state["selection_scope"],
            state["selection_size"],
            ", ".join(state["selection_tags"]) or "<none>",
        )
    )
    logger.write(
        "  ObjectTypes: {0}".format(
            ", ".join(
                "{0}={1}".format(name, state["object_types"][name])
                for name in OBJECT_TYPE_NAMES
            )
        )
    )


def configure_common_step_options(creator, output_path):
    creator.OutputFile = output_path
    creator.ObjectTypes.Solids = True
    creator.ObjectTypes.Surfaces = True
    creator.ObjectTypes.Curves = True
    creator.ExportAs = NXOpen.StepCreator.ExportAsOption.Ap214
    creator.ProcessHoldFlag = True


def enable_all_body_types(creator):
    creator.ObjectTypes.ConvergentBodies = True
    creator.ObjectTypes.FacetBodies = True
    creator.ObjectTypes.Tessellation = True


def discover_step_settings_file():
    candidates = []
    for variable in ("UGII_BASE_DIR", "UGII_ROOT_DIR"):
        root = clean(os.environ.get(variable))
        if not root:
            continue
        candidates.extend((
            os.path.join(root, "STEP214UG", "step214ug.def"),
            os.path.join(root, "step214ug", "step214ug.def"),
            os.path.join(root, "translators", "step214ug.def"),
        ))
    seen = set()
    for path in candidates:
        normalized = os.path.normcase(os.path.abspath(path))
        if normalized in seen:
            continue
        seen.add(normalized)
        if os.path.isfile(path):
            return path
    return ""


def set_selected_bodies(creator, bodies):
    selector = creator.ExportSelectionBlock
    selector.SelectionScope = NXOpen.ObjectSelector.Scope.SelectedObjects
    selection = selector.SelectionComp
    try:
        selection.SetArray(bodies)
    except Exception:
        try:
            selection.Clear()
        except Exception:
            pass
        for body in bodies:
            selection.Add(body)

    selected = selection_objects(selection)
    size = safe_property(selection, "Size", len(selected))
    try:
        size_value = int(size)
    except Exception:
        size_value = len(selected)
    if not bodies:
        raise RuntimeError("No direct bodies are available for selection.")
    if size_value <= 0 and not selected:
        raise RuntimeError(
            "STEP selected-object list stayed empty after adding bodies."
        )


def run_step_trial(
    session,
    part,
    specification,
    bodies,
    run_folder,
    trial_spec,
    logger,
):
    trial_id = trial_spec["trial"]
    label = trial_spec["label"]
    phase = trial_spec["phase"]
    logger.section("{0} - {1}".format(trial_id, label))
    output_path = os.path.join(
        run_folder,
        "{0}_{1}.stp".format(trial_id, phase),
    )
    record = {
        "trial": trial_id,
        "label": label,
        "phase": phase,
        "commit_result": "NOT_RUN",
        "error": "",
        "traceback": "",
        "output_file": output_path,
        "format": "STEP",
        "builder_defaults": {},
        "builder_final": {},
        "builder_validate_result": "NOT_RUN",
        "translator": {},
    }

    if os.path.exists(output_path):
        record["commit_result"] = "FAILED"
        record["error"] = "Output already exists; refusing to overwrite."
        record["inspection"] = inspect_step_file(output_path)
        return record

    creator = None
    try:
        # Journal 07 activates the resolved part before creating the builder.
        activate_part(session, part)
        creator = session.DexManager.CreateStepCreator()
        record["builder_defaults"] = builder_state(creator)
        log_builder_state(
            logger,
            "STEP builder defaults:",
            record["builder_defaults"],
        )
        configure_common_step_options(creator, output_path)

        input_mode = trial_spec.get("input_mode")
        if input_mode == "FULL_PATH":
            creator.InputFile = part.FullPath
        elif input_mode == "CANONICAL":
            creator.InputFile = specification
        elif input_mode:
            raise RuntimeError("Unknown STEP input mode: " + input_mode)

        scope = trial_spec.get("scope")
        if scope:
            creator.ExportFrom = (
                NXOpen.StepCreator.ExportFromOption.DisplayPart
            )
        if scope == "ENTIRE":
            creator.ExportSelectionBlock.SelectionScope = (
                NXOpen.ObjectSelector.Scope.EntirePart
            )
        elif scope == "SELECTED":
            set_selected_bodies(creator, bodies)
        elif scope:
            raise RuntimeError("Unknown STEP selection scope: " + scope)

        if trial_spec.get("layer_mask"):
            creator.LayerMask = trial_spec["layer_mask"]
        if trial_spec.get("settings_file"):
            creator.SettingsFile = trial_spec["settings_file"]
        if trial_spec.get("all_body_types"):
            enable_all_body_types(creator)
        if "structures" in trial_spec:
            creator.ObjectTypes.Structures = bool(
                trial_spec["structures"]
            )

        record["builder_final"] = builder_state(creator)
        log_builder_state(
            logger,
            "STEP builder final state:",
            record["builder_final"],
        )
        record["builder_validate_result"] = clean(creator.Validate())
        logger.write(
            "Builder Validate(): {0}".format(
                record["builder_validate_result"]
            )
        )

        creator.Commit()
        record["commit_result"] = "COMMITTED"
    except Exception as error:
        record["commit_result"] = "FAILED"
        record["error"] = exception_details(error)
        record["traceback"] = traceback.format_exc()
    finally:
        if creator is not None:
            try:
                creator.Destroy()
            except Exception as error:
                if not record["error"]:
                    record["error"] = (
                        "STEP builder destroy failed: {0}".format(
                            exception_details(error)
                        )
                    )

    record["inspection"] = inspect_step_file(output_path)
    record["translator"] = parse_translator_log(output_path)
    record["snapshot"] = part_snapshot(
        session,
        part,
        "{0} post-export".format(trial_id),
    )

    inspection = record["inspection"]
    logger.write("Commit result: {0}".format(record["commit_result"]))
    if record["error"]:
        logger.write("Error: {0}".format(record["error"]))
    logger.write("Output: {0}".format(output_path))
    logger.write(
        "File: exists={0}, size={1}, DATA entities={2}".format(
            inspection["exists"],
            inspection["size_bytes"],
            inspection["data_entity_lines"],
        )
    )
    logger.write(
        "STEP signatures: body geometry={0}, assembly={1}".format(
            inspection["body_geometry_signatures"],
            inspection["assembly_signatures"],
        )
    )
    logger.write("Has body geometry: {0}".format(
        inspection["has_body_geometry"]
    ))
    translator = record["translator"]
    if translator.get("path"):
        logger.write(
            "Translator: solids input={0}, processed={1}, as sheets={2}, "
            "not processed={3}, UG body count={4}".format(
                translator["solids_input"],
                translator["solids_processed"],
                translator["solids_as_sheets"],
                translator["solids_not_processed"],
                translator["ug_body_count"],
            )
        )
        logger.write("Translator log: {0}".format(translator["path"]))
    if inspection["inspection_error"]:
        logger.write("Inspection error: {0}".format(
            inspection["inspection_error"]
        ))
    log_snapshot(logger, record["snapshot"])
    return record


def run_parasolid_control(
    session,
    part,
    bodies,
    run_folder,
    logger,
):
    trial_id = "CONTROL_P"
    label = "Selected direct bodies through Parasolid exporter"
    phase = "selected_body_parasolid"
    logger.section("{0} - {1}".format(trial_id, label))
    output_path = os.path.join(run_folder, "{0}_{1}.x_t".format(
        trial_id,
        phase,
    ))
    record = {
        "trial": trial_id,
        "label": label,
        "phase": phase,
        "commit_result": "NOT_RUN",
        "error": "",
        "traceback": "",
        "output_file": output_path,
        "format": "PARASOLID",
        "builder_defaults": {},
        "builder_final": {},
        "builder_validate_result": "NOT_RUN",
        "translator": {},
    }

    exporter = None
    try:
        activate_part(session, part)
        exporter = session.DexManager.CreateParasolidExporter()
        exporter.OutputFile = output_path
        exporter.ExportFrom = (
            NXOpen.ParasolidExporter.ExportFromOption.DisplayedPart
        )
        exporter.ExportSelectionBlock.SelectionScope = (
            NXOpen.ObjectSelector.Scope.SelectedObjects
        )
        exporter.ObjectTypes.Solids = True
        exporter.ObjectTypes.Surfaces = True
        selection = exporter.ExportSelectionBlock.SelectionComp
        try:
            selection.SetArray(bodies)
        except Exception:
            for body in bodies:
                selection.Add(body)
        if not bodies:
            raise RuntimeError(
                "No direct bodies are available for Parasolid control."
            )
        selected = selection_objects(selection)
        size = safe_property(selection, "Size", len(selected))
        if not selected and clean(size) in ("", "0"):
            raise RuntimeError(
                "Parasolid selected-object list stayed empty."
            )
        record["builder_final"] = {
            "selection_scope": enum_text(
                exporter.ExportSelectionBlock.SelectionScope
            ),
            "selection_size": size,
            "selection_tags": [
                clean(safe_property(value, "Tag")) for value in selected
            ],
            "export_from": enum_text(exporter.ExportFrom),
        }
        record["builder_validate_result"] = clean(exporter.Validate())
        exporter.Commit()
        record["commit_result"] = "COMMITTED"
    except Exception as error:
        record["commit_result"] = "FAILED"
        record["error"] = exception_details(error)
        record["traceback"] = traceback.format_exc()
    finally:
        if exporter is not None:
            try:
                exporter.Destroy()
            except Exception:
                pass

    exists = os.path.isfile(output_path)
    try:
        size_bytes = os.path.getsize(output_path) if exists else ""
    except Exception:
        size_bytes = ""
    parasolid_body_signatures = 0
    if exists:
        try:
            with open(
                output_path,
                "r",
                encoding="utf-8",
                errors="replace",
            ) as handle:
                parasolid_body_signatures = len(re.findall(
                    r"\bbody\b",
                    handle.read(),
                    re.IGNORECASE,
                ))
        except Exception:
            pass
    has_geometry = parasolid_body_signatures > 0
    record["inspection"] = {
        "exists": exists,
        "size_bytes": size_bytes,
        "data_entity_lines": "",
        "body_geometry_signatures": parasolid_body_signatures,
        "assembly_signatures": "",
        "has_body_geometry": has_geometry,
        "inspection_error": "",
    }
    record["snapshot"] = part_snapshot(
        session,
        part,
        "{0} post-export".format(trial_id),
    )
    logger.write("Commit result: {0}".format(record["commit_result"]))
    if record["error"]:
        logger.write("Error: {0}".format(record["error"]))
    logger.write(
        "Parasolid control: exists={0}, size={1}, body signatures={2}".format(
            exists,
            size_bytes,
            parasolid_body_signatures,
        )
    )
    logger.write(
        "Selection size={0}, tags={1}".format(
            record["builder_final"].get("selection_size", ""),
            ", ".join(
                record["builder_final"].get("selection_tags", [])
            ) or "<none>",
        )
    )
    return record


def trial_has_geometry(trial):
    try:
        if bool(trial["inspection"]["has_body_geometry"]):
            return True
        solids_input = trial.get("translator", {}).get(
            "solids_input",
            "",
        )
        return solids_input != "" and int(solids_input) > 0
    except Exception:
        return False


def body_flag(records, name):
    return any(bool(record.get(name)) for record in records)


def determine_conclusion(trials, body_records, after_load, load_result):
    by_id = {trial["trial"]: trial for trial in trials}
    success = {
        trial_id: trial_has_geometry(by_id.get(trial_id, {}))
        for trial_id in (
            "TRIAL_A",
            "TRIAL_A2",
            "TRIAL_B",
            "TRIAL_C",
            "TRIAL_D",
            "TRIAL_I",
            "TRIAL_E",
            "TRIAL_F",
            "TRIAL_G",
            "TRIAL_H",
            "CONTROL_P",
        )
    }

    if success["TRIAL_A"]:
        return (
            "NOT_REPRODUCED",
            "The unchanged Journal 07 configuration produced body geometry. "
            "Do not patch Journal 07 from this run.",
        )
    if success["TRIAL_A2"]:
        return (
            "LOAD_STATE_CAUSE",
            "The unchanged Journal 07 configuration first produced geometry "
            "after the explicit full-load request. Journal 07 should ensure "
            "the target part is fully loaded before creating StepCreator.",
        )
    if success["TRIAL_B"]:
        return (
            "CANONICAL_INPUT_PATH_CAUSE",
            "The first successful STEP differs only by using the canonical "
            "Teamcenter specification as StepCreator.InputFile. Journal 07 "
            "should pass the resolved @DB specification, not Part.FullPath.",
        )
    if success["TRIAL_C"]:
        return (
            "EXPORT_SOURCE_SCOPE_CAUSE",
            "The first successful STEP explicitly uses DisplayPart with "
            "EntirePart scope. Journal 07 should use that source/scope for "
            "this Teamcenter part.",
        )
    if success["TRIAL_D"]:
        return (
            "LAYER_MASK_CAUSE",
            "The first successful STEP adds LayerMask='1-256'. Journal 07 "
            "should explicitly export all layers.",
        )
    if success["TRIAL_I"]:
        return (
            "SETTINGS_FILE_CAUSE",
            "The first successful STEP differs from the all-layer baseline "
            "only by explicitly assigning the installed AP214 settings file. "
            "Journal 07 should set that SettingsFile.",
        )
    if success["TRIAL_F"]:
        return (
            "ENTIRE_PART_SELECTION_CAUSE",
            "The same standard STEP body types work when the direct body is "
            "explicitly selected. Journal 07 should populate SelectionComp "
            "and verify its size before Commit().",
        )
    if success["TRIAL_E"]:
        if body_flag(body_records, "is_convergent_body"):
            return (
                "CONVERGENT_BODY_FILTER_CAUSE",
                "The body is convergent and STEP succeeds only after "
                "ConvergentBodies/FacetBodies/Tessellation are enabled.",
            )
        return (
            "STEP_FILTER_CONFIGURATION_CAUSE",
            "STEP succeeds only after all supported body representation "
            "filters are enabled. Apply those ObjectTypes to Journal 07.",
        )
    if success["TRIAL_G"]:
        if body_flag(body_records, "is_convergent_body"):
            return (
                "SELECTION_AND_BODY_TYPE_CAUSE",
                "The convergent body requires both explicit selection and "
                "the convergent/facet/tessellation STEP filters.",
            )
        return (
            "SELECTION_AND_BODY_TYPE_CAUSE",
            "The body requires explicit selection plus the expanded STEP "
            "body representation filters.",
        )
    if success["TRIAL_H"]:
        display_filtered = body_flag(body_records, "is_blanked") or any(
            clean(record.get("layer_state")).upper() == "HIDDEN"
            for record in body_records
        )
        if display_filtered:
            return (
                "BLANKED_BODY_FILTER_CAUSE",
                "STEP succeeds only after temporarily unblanking the body "
                "and making its layer selectable. Journal 07 must restore "
                "those display states after export.",
            )
        return (
            "DISPLAY_STATE_FILTER_CAUSE",
            "STEP succeeds only after normalizing the body visibility/layer "
            "state, although the recorded initial flags were non-obvious.",
        )
    if success["CONTROL_P"]:
        return (
            "STEP_API_VS_UI_MISMATCH",
            "The selected direct body exports through the Parasolid API, "
            "while every STEP API configuration still reports zero geometry. "
            "The next evidence required is an NX UI-recorded STEP journal for "
            "the exact gasket; do not guess a Journal 07 setting.",
        )

    if not success["CONTROL_P"]:
        available_bodies = (
            int(after_load.get("body_count", 0))
            + int(after_load.get("descendant_body_occurrence_count", 0))
        )
        if available_bodies <= 0 and load_result.get("status") in (
            "SUCCESS",
            "SUCCESS_FALLBACK",
        ):
            return (
                "EMPTY_OR_WRONG_MASTER_DATASET",
                "The fully loaded Teamcenter result contains no direct or "
                "descendant bodies.",
            )
        return (
            "BODY_GEOMETRY_CONTROL_FAILURE",
            "NX reports bodies, but neither STEP nor the selected-body "
            "Parasolid control produced a non-empty geometry file. Inspect "
            "the recorded body topology and exporter errors before changing "
            "Journal 07.",
        )

    return (
        "INCONCLUSIVE",
        "The trial pattern does not map cleanly to one isolated cause. Review "
        "the per-trial load and STEP entity diagnostics.",
    )


def trial_csv_row(trial):
    inspection = trial.get("inspection", {})
    snapshot = trial.get("snapshot", {})
    translator = trial.get("translator", {})
    builder = trial.get("builder_final", {})
    object_types = builder.get("object_types", {})
    return {
        "TRIAL": trial.get("trial", ""),
        "LABEL": trial.get("label", ""),
        "PHASE": trial.get("phase", ""),
        "COMMIT_RESULT": trial.get("commit_result", ""),
        "ERROR": trial.get("error", ""),
        "OUTPUT_FILE": trial.get("output_file", ""),
        "FORMAT": trial.get("format", ""),
        "FILE_EXISTS": inspection.get("exists", ""),
        "FILE_SIZE_BYTES": inspection.get("size_bytes", ""),
        "DATA_ENTITY_LINES": inspection.get("data_entity_lines", ""),
        "BODY_GEOMETRY_SIGNATURES": inspection.get(
            "body_geometry_signatures", ""
        ),
        "ASSEMBLY_SIGNATURES": inspection.get("assembly_signatures", ""),
        "HAS_BODY_GEOMETRY": inspection.get("has_body_geometry", ""),
        "TRANSLATOR_SOLIDS_INPUT": translator.get("solids_input", ""),
        "TRANSLATOR_SOLIDS_PROCESSED": translator.get(
            "solids_processed",
            "",
        ),
        "TRANSLATOR_SOLIDS_AS_SHEETS": translator.get(
            "solids_as_sheets",
            "",
        ),
        "TRANSLATOR_SOLIDS_NOT_PROCESSED": translator.get(
            "solids_not_processed",
            "",
        ),
        "TRANSLATOR_UG_BODY_COUNT": translator.get("ug_body_count", ""),
        "TRANSLATOR_LOG": translator.get("path", ""),
        "BUILDER_VALIDATE_RESULT": trial.get(
            "builder_validate_result",
            "",
        ),
        "EXPORT_FROM": builder.get("export_from", ""),
        "INPUT_FILE": builder.get("input_file", ""),
        "LAYER_MASK": builder.get("layer_mask", ""),
        "FILE_SAVE_FLAG": builder.get("file_save_flag", ""),
        "SETTINGS_FILE": builder.get("settings_file", ""),
        "SELECTION_SCOPE": builder.get("selection_scope", ""),
        "SELECTION_SIZE": builder.get("selection_size", ""),
        "SOLIDS_ENABLED": object_types.get("Solids", ""),
        "SURFACES_ENABLED": object_types.get("Surfaces", ""),
        "CURVES_ENABLED": object_types.get("Curves", ""),
        "CONVERGENT_BODIES_ENABLED": object_types.get(
            "ConvergentBodies",
            "",
        ),
        "FACET_BODIES_ENABLED": object_types.get("FacetBodies", ""),
        "TESSELLATION_ENABLED": object_types.get("Tessellation", ""),
        "STRUCTURES_ENABLED": object_types.get("Structures", ""),
        "PART_LOAD_STATE": snapshot.get("load_state", ""),
        "PART_IS_FULLY_LOADED": snapshot.get("is_fully_loaded", ""),
        "DIRECT_BODY_COUNT": snapshot.get("body_count", ""),
        "DIRECT_SOLID_BODY_COUNT": snapshot.get("solid_body_count", ""),
        "DIRECT_SHEET_BODY_COUNT": snapshot.get("sheet_body_count", ""),
        "COMPONENT_OCCURRENCE_COUNT": snapshot.get(
            "component_occurrence_count", ""
        ),
        "UNIQUE_PROTOTYPE_COUNT": snapshot.get(
            "unique_prototype_count", ""
        ),
        "DESCENDANT_BODY_OCCURRENCE_COUNT": snapshot.get(
            "descendant_body_occurrence_count", ""
        ),
        "DESCENDANT_SOLID_BODY_OCCURRENCE_COUNT": snapshot.get(
            "descendant_solid_body_occurrence_count", ""
        ),
    }


def write_csv_report(path, trials):
    with open(path, "w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=CSV_COLUMNS)
        writer.writeheader()
        for trial in trials:
            writer.writerow(trial_csv_row(trial))


def write_body_csv(path, records):
    with open(path, "w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=BODY_CSV_COLUMNS)
        writer.writeheader()
        for record in records:
            writer.writerow({
                "INDEX": record.get("index", ""),
                "TAG": record.get("tag", ""),
                "RUNTIME_TYPE": record.get("runtime_type", ""),
                "NAME": record.get("name", ""),
                "JOURNAL_IDENTIFIER": record.get(
                    "journal_identifier",
                    "",
                ),
                "OWNING_PART": record.get("owning_part", ""),
                "OWNING_PART_MATCHES_TARGET": record.get(
                    "owning_part_matches_target",
                    "",
                ),
                "LAYER": record.get("layer", ""),
                "LAYER_STATE": record.get("layer_state", ""),
                "IS_BLANKED": record.get("is_blanked", ""),
                "IS_SOLID_BODY": record.get("is_solid_body", ""),
                "IS_SHEET_BODY": record.get("is_sheet_body", ""),
                "IS_CONVERGENT_BODY": record.get(
                    "is_convergent_body",
                    "",
                ),
                "FACE_COUNT": record.get("face_count", ""),
                "EDGE_COUNT": record.get("edge_count", ""),
                "FACET_COUNT": record.get("facet_count", ""),
                "VERTEX_COUNT": record.get("vertex_count", ""),
                "REFERENCE_SETS": "; ".join(
                    record.get("reference_sets", [])
                ),
            })


def write_text_report(path, logger):
    with open(path, "w", encoding="utf-8", newline="") as handle:
        for line in logger.lines:
            handle.write(str(line) + "\n")


def restore_original_state(
    session,
    original_display,
    original_work,
    logger,
):
    logger.write("Restoring original NX state...")
    if original_display is not None:
        try:
            set_display(session, original_display)
            logger.write("  Display restored: {0}".format(
                part_name(original_display)
            ))
        except Exception as error:
            logger.write("  WARNING: display restore failed: {0}".format(
                exception_details(error)
            ))
    if original_work is not None:
        try:
            session.Parts.SetWork(original_work)
            logger.write("  Work part restored: {0}".format(
                part_name(original_work)
            ))
        except Exception as error:
            logger.write("  WARNING: work-part restore failed: {0}".format(
                exception_details(error)
            ))


def close_test_part(
    part,
    opened_by_test,
    keep_open,
    session,
    original_display,
    original_work,
    logger,
):
    if part is None or not opened_by_test:
        return
    if keep_open:
        logger.write("Diagnostic-opened part left open by configuration.")
        return
    if original_display is None:
        logger.write(
            "Diagnostic-opened part left open because there was no original "
            "display part to restore."
        )
        return

    restore_original_state(
        session,
        original_display,
        original_work,
        logger,
    )
    try:
        part.Close(
            NXOpen.BasePart.CloseWholeTree.FalseValue,
            NXOpen.BasePart.CloseModified.CloseModified,
            None,
        )
        logger.write("Diagnostic-opened top-level part closed.")
    except Exception as error:
        logger.write("WARNING: diagnostic part close failed: {0}".format(
            exception_details(error)
        ))


class Logger:
    def __init__(self, session):
        self.lines = []
        self.window = None
        try:
            self.window = session.ListingWindow
            self.window.Open()
        except Exception:
            pass

    def write(self, message=""):
        value = str(message)
        self.lines.append(value)
        if self.window is not None:
            try:
                for line in value.splitlines() or [""]:
                    self.window.WriteFullline(line)
            except Exception:
                pass
        try:
            print(value)
        except Exception:
            pass

    def section(self, title):
        self.write("")
        self.write("=" * 78)
        self.write(title)
        self.write("=" * 78)


def main():
    session = NXOpen.Session.GetSession()
    logger = Logger(session)
    trials = []
    body_records = []
    temporary_display_states = []
    part = None
    opened_by_test = False
    run_folder = ""
    text_report = ""
    csv_report = ""
    body_report = ""
    conclusion = "INCONCLUSIVE"

    part_number = (
        clean(os.environ.get("NX_STEP_DIAG_PART_NO"))
        or DEFAULT_PART_NUMBER
    )
    revision = (
        clean(os.environ.get("NX_STEP_DIAG_PART_REV"))
        or DEFAULT_REVISION
    )
    keep_open = env_bool(
        "NX_STEP_DIAG_KEEP_OPEN",
        KEEP_OPEN_AFTER_TEST,
    )
    specification = "@DB/{0}/{1}".format(part_number, revision)

    try:
        original_display = session.Parts.Display
    except Exception:
        original_display = None
    try:
        original_work = session.Parts.Work
    except Exception:
        original_work = None

    preloaded_keys = {object_key(value) for value in loaded_parts(session)}

    try:
        run_folder, timestamp = create_run_folder()
        text_report = os.path.join(
            run_folder,
            "STEP_DIAGNOSTIC_{0}.txt".format(timestamp),
        )
        csv_report = os.path.join(
            run_folder,
            "STEP_DIAGNOSTIC_TRIALS_{0}.csv".format(timestamp),
        )
        body_report = os.path.join(
            run_folder,
            "STEP_DIAGNOSTIC_BODIES_{0}.csv".format(timestamp),
        )

        logger.section("Journal 10 - STEP Zero-Geometry Diagnostic")
        logger.write("Target: {0} / {1}".format(part_number, revision))
        logger.write("Specification: {0}".format(specification))
        logger.write("Output folder: {0}".format(run_folder))
        logger.write("Teamcenter operation: read/open only; no save")

        logger.section("OpenBase reproduction")
        load_status = None
        try:
            part, load_status = unwrap_open_result(
                session.Parts.OpenBase(specification)
            )
            if part is None:
                raise RuntimeError("OpenBase returned no part.")
            opened_by_test = object_key(part) not in preloaded_keys
            logger.write("OpenBase returned: {0}".format(part_name(part)))
            logger.write("Opened by diagnostic: {0}".format(opened_by_test))
            for message in load_status_lines(load_status):
                logger.write("  Open load status: {0}".format(message))
        finally:
            dispose(load_status)

        before_load = part_snapshot(
            session,
            part,
            "Immediately after OpenBase",
        )
        log_snapshot(logger, before_load)

        trials.append(
            run_step_trial(
                session,
                part,
                specification,
                direct_bodies(part),
                run_folder,
                {
                    "trial": "TRIAL_A",
                    "label": (
                        "Current Journal 07 settings immediately after "
                        "OpenBase"
                    ),
                    "phase": "before_full_load",
                    "input_mode": "FULL_PATH",
                },
                logger,
            )
        )

        logger.section("Explicit full-load request")
        load_result = ensure_fully_loaded(session, part, logger)
        after_load = part_snapshot(
            session,
            part,
            "After explicit full load",
        )
        log_snapshot(logger, after_load)

        body_records = body_diagnostics(part)
        bodies = [record["body"] for record in body_records]
        log_body_diagnostics(logger, body_records)
        write_body_csv(body_report, body_records)
        logger.write("Body report: {0}".format(body_report))

        settings_file = discover_step_settings_file()
        if settings_file:
            logger.write(
                "Installed AP214 settings candidate: {0}".format(
                    settings_file
                )
            )
        else:
            logger.write(
                "No installed AP214 settings file was found from "
                "UGII_BASE_DIR/UGII_ROOT_DIR; TRIAL_I will be skipped."
            )

        trial_specs = [
            {
                "trial": "TRIAL_A2",
                "label": "Current Journal 07 settings after full load",
                "phase": "after_full_load",
                "input_mode": "FULL_PATH",
            },
            {
                "trial": "TRIAL_B",
                "label": "Canonical @DB InputFile after full load",
                "phase": "canonical_input",
                "input_mode": "CANONICAL",
            },
            {
                "trial": "TRIAL_C",
                "label": "DisplayPart + EntirePart",
                "phase": "display_entire_part",
                "scope": "ENTIRE",
                "structures": False,
            },
            {
                "trial": "TRIAL_D",
                "label": "DisplayPart + EntirePart + all layers",
                "phase": "entire_part_all_layers",
                "scope": "ENTIRE",
                "layer_mask": ALL_LAYER_MASK,
                "structures": False,
            },
        ]
        if settings_file:
            trial_specs.append({
                "trial": "TRIAL_I",
                "label": (
                    "EntirePart + all layers + installed AP214 settings"
                ),
                "phase": "installed_ap214_settings",
                "scope": "ENTIRE",
                "layer_mask": ALL_LAYER_MASK,
                "settings_file": settings_file,
                "structures": False,
            })
        trial_specs.extend((
            {
                "trial": "TRIAL_E",
                "label": (
                    "EntirePart + all layers + all body representations"
                ),
                "phase": "entire_all_body_types",
                "scope": "ENTIRE",
                "layer_mask": ALL_LAYER_MASK,
                "all_body_types": True,
                "structures": False,
            },
            {
                "trial": "TRIAL_F",
                "label": (
                    "Explicit direct-body selection with standard types"
                ),
                "phase": "selected_standard_types",
                "scope": "SELECTED",
                "layer_mask": ALL_LAYER_MASK,
                "structures": False,
            },
            {
                "trial": "TRIAL_G",
                "label": (
                    "Explicit selection + all body representations"
                ),
                "phase": "selected_all_body_types",
                "scope": "SELECTED",
                "layer_mask": ALL_LAYER_MASK,
                "all_body_types": True,
                "structures": False,
            },
        ))

        for trial_spec in trial_specs:
            trials.append(
                run_step_trial(
                    session,
                    part,
                    specification,
                    bodies,
                    run_folder,
                    trial_spec,
                    logger,
                )
            )

        logger.section("Temporary display-state normalization")
        temporary_display_states = capture_body_display_state(
            part,
            bodies,
        )
        try:
            make_bodies_visible_and_selectable(
                part,
                temporary_display_states,
                logger,
            )
            trials.append(
                run_step_trial(
                    session,
                    part,
                    specification,
                    bodies,
                    run_folder,
                    {
                        "trial": "TRIAL_H",
                        "label": (
                            "Selected all body types after unblank/layer "
                            "normalization"
                        ),
                        "phase": "selected_visible_selectable",
                        "scope": "SELECTED",
                        "layer_mask": ALL_LAYER_MASK,
                        "all_body_types": True,
                        "structures": False,
                    },
                    logger,
                )
            )
        finally:
            restore_body_display_state(
                part,
                temporary_display_states,
                logger,
            )
            temporary_display_states = []

        trials.append(
            run_parasolid_control(
                session,
                part,
                bodies,
                run_folder,
                logger,
            )
        )

        conclusion, explanation = determine_conclusion(
            trials,
            body_records,
            after_load,
            load_result,
        )
        logger.section("Diagnostic conclusion")
        logger.write("ROOT_CAUSE_CLASSIFICATION: {0}".format(conclusion))
        logger.write(explanation)

        write_csv_report(csv_report, trials)
        logger.write("CSV report: {0}".format(csv_report))

    except Exception as error:
        logger.section("Unhandled diagnostic failure")
        logger.write(exception_details(error))
        logger.write(traceback.format_exc())

    finally:
        if part is not None and temporary_display_states:
            try:
                restore_body_display_state(
                    part,
                    temporary_display_states,
                    logger,
                )
            except Exception as error:
                logger.write(
                    "WARNING: emergency body-state restore failed: {0}".format(
                        exception_details(error)
                    )
                )
        restore_original_state(
            session,
            original_display,
            original_work,
            logger,
        )
        close_test_part(
            part,
            opened_by_test,
            keep_open,
            session,
            original_display,
            original_work,
            logger,
        )
        if run_folder:
            if trials and csv_report:
                try:
                    write_csv_report(csv_report, trials)
                except Exception as error:
                    logger.write("WARNING: CSV report write failed: {0}".format(
                        exception_details(error)
                    ))
            if body_records and body_report:
                try:
                    write_body_csv(body_report, body_records)
                except Exception as error:
                    logger.write(
                        "WARNING: body CSV report write failed: {0}".format(
                            exception_details(error)
                        )
                    )
            logger.write("Final classification: {0}".format(conclusion))
            logger.write("Text report: {0}".format(text_report))
            try:
                write_text_report(text_report, logger)
            except Exception:
                pass


if __name__ == "__main__":
    main()
