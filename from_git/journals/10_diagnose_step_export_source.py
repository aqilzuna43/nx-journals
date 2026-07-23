"""
Journal 10 - Teamcenter STEP Zero-Geometry Diagnostic

This journal diagnoses why Journal 07 can create a very small AP214 STEP file
whose translator log reports zero input solids. It deliberately runs four
controlled exports so that load-state, export-scope, and assembly-structure
effects can be separated.

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


DEFAULT_PART_NUMBER = "264MN020016A01"
DEFAULT_REVISION = "A"
OUTPUT_FOLDER_NAME = "NX_STEP_DIAGNOSTIC"
KEEP_OPEN_AFTER_TEST = False
MAX_COMPONENT_OCCURRENCES = 20000

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
    "FILE_EXISTS",
    "FILE_SIZE_BYTES",
    "DATA_ENTITY_LINES",
    "BODY_GEOMETRY_SIGNATURES",
    "ASSEMBLY_SIGNATURES",
    "HAS_BODY_GEOMETRY",
    "TRANSLATOR_SOLIDS_INPUT",
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


def parse_translator_solid_count(output_path):
    folder = os.path.dirname(output_path)
    output_name = os.path.basename(output_path).upper()
    candidates = []
    try:
        for name in os.listdir(folder):
            if name.lower().endswith(".log"):
                candidates.append(os.path.join(folder, name))
    except Exception:
        return "", ""

    preferred = []
    others = []
    for path in candidates:
        try:
            with open(path, "r", encoding="utf-8", errors="replace") as handle:
                content = handle.read()
            if output_name in content.upper():
                preferred.append((path, content))
            else:
                others.append((path, content))
        except Exception:
            pass

    for path, content in preferred + others:
        match = re.search(
            r"Total\s+number\s+of\s+solids\s+input.*?:\s*(\d+)",
            content,
            re.IGNORECASE,
        )
        if match:
            return match.group(1), path
    return "", ""


def configure_common_step_options(creator, output_path):
    creator.OutputFile = output_path
    creator.ObjectTypes.Solids = True
    creator.ObjectTypes.Surfaces = True
    creator.ObjectTypes.Curves = True
    creator.ExportAs = NXOpen.StepCreator.ExportAsOption.Ap214
    creator.ProcessHoldFlag = True


def run_step_trial(
    session,
    part,
    run_folder,
    trial_id,
    label,
    phase,
    mode,
    structures,
    logger,
):
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
        "mode": mode,
        "structures": structures,
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
        configure_common_step_options(creator, output_path)

        if mode == "CURRENT_J07":
            creator.InputFile = part.FullPath
        elif mode == "DISPLAY_ENTIRE_PART":
            creator.ExportFrom = (
                NXOpen.StepCreator.ExportFromOption.DisplayPart
            )
            creator.ExportSelectionBlock.SelectionScope = (
                NXOpen.ObjectSelector.Scope.EntirePart
            )
            creator.ObjectTypes.Structures = bool(structures)
        else:
            raise RuntimeError("Unknown diagnostic trial mode: " + mode)

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
    solid_count, translator_log = parse_translator_solid_count(output_path)
    record["translator_solids_input"] = solid_count
    record["translator_log"] = translator_log
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
    if solid_count:
        logger.write("Translator solids input: {0}".format(solid_count))
        logger.write("Translator log: {0}".format(translator_log))
    if inspection["inspection_error"]:
        logger.write("Inspection error: {0}".format(
            inspection["inspection_error"]
        ))
    log_snapshot(logger, record["snapshot"])
    return record


def trial_has_geometry(trial):
    try:
        return bool(trial["inspection"]["has_body_geometry"])
    except Exception:
        return False


def determine_conclusion(trials, before_load, after_load, load_result):
    by_id = {trial["trial"]: trial for trial in trials}
    a = trial_has_geometry(by_id.get("TRIAL_A", {}))
    b = trial_has_geometry(by_id.get("TRIAL_B", {}))
    c = trial_has_geometry(by_id.get("TRIAL_C", {}))
    d = trial_has_geometry(by_id.get("TRIAL_D", {}))

    if not a and b:
        return (
            "LOAD_STATE_CAUSE",
            "The Journal 07 configuration exported geometry only after a "
            "full-load request that included assembly children.",
        )
    if not b and c:
        return (
            "EXPORT_SCOPE_CAUSE",
            "The fully loaded part exported geometry only after switching "
            "from InputFile mode to explicit DisplayPart/EntirePart scope.",
        )
    if not c and d:
        return (
            "ASSEMBLY_STRUCTURE_CAUSE",
            "Geometry appeared only when STEP assembly Structures were enabled.",
        )

    if not d:
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
                "descendant bodies, and none of the STEP configurations "
                "produced body geometry.",
            )
        return (
            "INCONCLUSIVE",
            "NX reports exportable bodies after loading, but none of the "
            "controlled STEP configurations produced body geometry.",
        )

    if a:
        return (
            "INCONCLUSIVE",
            "The current Journal 07 configuration produced body geometry in "
            "this run, so the uploaded zero-solid failure was not reproduced.",
        )

    return (
        "INCONCLUSIVE",
        "The trial pattern does not map cleanly to one isolated cause. Review "
        "the per-trial load and STEP entity diagnostics.",
    )


def trial_csv_row(trial):
    inspection = trial.get("inspection", {})
    snapshot = trial.get("snapshot", {})
    return {
        "TRIAL": trial.get("trial", ""),
        "LABEL": trial.get("label", ""),
        "PHASE": trial.get("phase", ""),
        "COMMIT_RESULT": trial.get("commit_result", ""),
        "ERROR": trial.get("error", ""),
        "OUTPUT_FILE": trial.get("output_file", ""),
        "FILE_EXISTS": inspection.get("exists", ""),
        "FILE_SIZE_BYTES": inspection.get("size_bytes", ""),
        "DATA_ENTITY_LINES": inspection.get("data_entity_lines", ""),
        "BODY_GEOMETRY_SIGNATURES": inspection.get(
            "body_geometry_signatures", ""
        ),
        "ASSEMBLY_SIGNATURES": inspection.get("assembly_signatures", ""),
        "HAS_BODY_GEOMETRY": inspection.get("has_body_geometry", ""),
        "TRANSLATOR_SOLIDS_INPUT": trial.get("translator_solids_input", ""),
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
    part = None
    opened_by_test = False
    run_folder = ""
    text_report = ""
    csv_report = ""
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
                run_folder,
                "TRIAL_A",
                "Current Journal 07 settings immediately after OpenBase",
                "before_full_load",
                "CURRENT_J07",
                False,
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

        trials.append(
            run_step_trial(
                session,
                part,
                run_folder,
                "TRIAL_B",
                "Current Journal 07 settings after full load",
                "after_full_load",
                "CURRENT_J07",
                False,
                logger,
            )
        )
        trials.append(
            run_step_trial(
                session,
                part,
                run_folder,
                "TRIAL_C",
                "DisplayPart + EntirePart after full load",
                "display_entire_part",
                "DISPLAY_ENTIRE_PART",
                False,
                logger,
            )
        )
        trials.append(
            run_step_trial(
                session,
                part,
                run_folder,
                "TRIAL_D",
                "DisplayPart + EntirePart + Structures after full load",
                "display_entire_part_structures",
                "DISPLAY_ENTIRE_PART",
                True,
                logger,
            )
        )

        conclusion, explanation = determine_conclusion(
            trials,
            before_load,
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
            logger.write("Final classification: {0}".format(conclusion))
            logger.write("Text report: {0}".format(text_report))
            try:
                write_text_report(text_report, logger)
            except Exception:
                pass


if __name__ == "__main__":
    main()
