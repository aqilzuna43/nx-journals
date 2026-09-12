"""
Journal 36 - Session-Tree PDF + STEP Packager

Packages every BoM-visible item of the assembly that is open right now, at
every level, without any operator CSV.

Open ASSY 123456 (which contains 3 sub-assemblies and 4 sub-sub-assemblies),
run this journal, and it walks the whole in-session structure, then writes one
PDF (every sheet of every drawing specification it can resolve) plus one
AP214 STEP package per level:

    NX_BULK_EXPORT\\<timestamp>\\
        PDF\\     <number>_REV<rev>.<WAE_VERSION>.pdf
                  <number>_REV<rev>.<WAE_VERSION>_DWG<n>.pdf   (2+ drawings)
        STEP\\    <number>_REV<rev>.<WAE_VERSION>.stp
        REPORTS\\ EXPORT_RESULT_<timestamp>.csv
        LOGS\\    EXPORT_LOG_<timestamp>.txt

How the target set is built
---------------------------
1. Traversal starts at the NX work part and recurses through every
   occurrence with Component.GetChildren() to unlimited depth.
2. SCOPE_FILTER = "BOM" (default) keeps only BoM-visible occurrences - the
   exact Journal 04 / Journal 21 filter: suppressed occurrences, reference
   components (REFERENCE_COMPONENT / PLIST_IGNORE_MEMBER), CSYS/DATUM/
   REFERENCE/SKELETON keyword names and CELESTICA_BOM_EXCLUDE_SUBTREE=YES
   subtrees are excluded. Set NX_J36_SCOPE=ALL for the raw tree.
3. Shared prototypes are deduplicated and exported once, at their topmost
   observed level. LEVEL / DEEPEST_LEVEL / OCCURRENCE_COUNT record the
   position that was seen.
4. LOAD_MODE = "LOAD" (default) calls BasePart.LoadThisPartFully() on every
   target that is not fully loaded and re-traverses until the scope is
   stable, so components hidden behind an unloaded sub-assembly are still
   discovered (the proven Journal 21 load-gate pattern). Load failures are
   reported per target and never abort sibling branches. Parts loaded this
   way deliberately remain loaded; nothing is saved, checked out or closed.

PDF rules (Journal 07 DataPack parity)
--------------------------------------
- Reuse drawing specifications already loaded anywhere in the session, then
  fall back to Teamcenter @DB/<number>/<rev>/specification/<number>-<rev>-dwg<n>
  for n = 1..9 through session.Parts.OpenDisplay.
- One PDF per drawing specification, containing every sheet of that
  specification. _DWG<n> is appended only when more than one drawing exists.
- Native NX PDF watermark DRAFT_<rev>.<WAE_VERSION> (WAE_VERSION comes from
  the target part or its drawing; a missing value keeps the revision-only
  watermark and records a warning) and one temporary bottom-right
  "EXPORTED: <date> <time> MYT" note per sheet that is undone immediately
  after the export, exactly like Journal 07.
- Drawing parts opened by the journal are closed again; drawings that were
  already loaded are left untouched.

STEP rules
----------
- STEP_ONLY_FOR_TARGETS_WITH_DRAWING = True (default) keeps the package
  aligned: an item is only worth a STEP package when it carries a drawing.
  Set NX_J36_STEP_SCOPE=ALL to export STEP for every BoM-visible target.
- Each target is exported as its own monolithic AP214 STEP containing its
  complete subtree (the Journal 10 / Journal 07 proven combination:
  ExportFrom = DisplayPart, SelectionScope = EntirePart, LayerMask 1-256,
  Solids + Surfaces + Curves, ProcessHoldFlag). Output is verified: missing
  or zero-geometry STEP files are reported, never silently accepted.

Modes and safety
----------------
- WRITE_MODE defaults to "DRY_RUN": the whole tree is traversed, loaded and
  its drawings resolved (so the report answers "which levels would export
  what?"), but no STEP/PDF file is written. Set NX_J36_MODE=APPLY to export.
- APPLY never overwrites: an existing output file is reported as
  SKIPPED_EXISTS and left alone.
- No part is checked out or saved. Temporary timestamp notes are undone to a
  private undo mark and the drawing's IsModified state is verified unchanged;
  a failed cleanup halts the remaining PDF batch.
- The original display part and work part are restored before the journal
  finishes, including on failure.

Environment overrides
---------------------
    NX_JOURNALS_IO_DIR     output root (default: %USERPROFILE%\\Desktop)
    NX_J36_MODE            DRY_RUN (default) | APPLY
    NX_J36_SCOPE           BOM (default) | ALL
    NX_J36_LOAD_MODE       LOAD (default) | REPORT_ONLY
    NX_J36_STEP_SCOPE      DRAWING_ONLY (default) | ALL

The EXPORT_RESULT report keeps the Journal 07 column names DB_PART_NO,
DB_PART_REV, PDF_RESULT, PDF_FILE_COUNT, PDF_FILES, so
from_git/utils/single_drawing_scope.py can still generate a J25
single-drawing scope from any run.

Target: NX 2312 and NX X 2506 embedded Python
Run via: NX > Tools > Journal > Play
"""

import csv
import datetime
import os
import re
import time
import traceback

import NXOpen


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

OUTPUT_ROOT_FOLDER = "NX_BULK_EXPORT"
JOURNAL_BUILD_ID = "J36-NX2506-SESSION-TREE-PDF-STEP-V1"

WRITE_MODE = "DRY_RUN"  # DRY_RUN / APPLY
SCOPE_FILTER = "BOM"  # BOM / ALL
LOAD_MODE = "LOAD"  # LOAD / REPORT_ONLY
STEP_ONLY_FOR_TARGETS_WITH_DRAWING = True  # NX_J36_STEP_SCOPE=ALL to disable

STEP_FORMAT = "AP214"
VERIFY_OUTPUT_FILES = True
STEP_LAYER_MASK = "1-256"
WAE_VERSION_ATTRIBUTE = "WAE_VERSION"

PDF_DRAFT_PREFIX = "DRAFT"
PDF_APPLY_DRAFT_WATERMARK = True
PDF_ADD_EXPORT_TIMESTAMP_NOTE = True
PDF_TIMESTAMP_TEXT_HEIGHT_MM = 2.5
PDF_TIMESTAMP_RIGHT_MARGIN_MM = 8.0
PDF_TIMESTAMP_BOTTOM_MARGIN_MM = 5.0
PDF_TIMESTAMP_FORMAT = "%Y-%m-%d %H:%M"
PDF_TIMESTAMP_TIMEZONE_LABEL = "MYT"
PDF_TIMESTAMP_OVERHEAD_TARGET_PERCENT = 10.0
MYT_TIMEZONE = datetime.timezone(
    datetime.timedelta(hours=8),
    name=PDF_TIMESTAMP_TIMEZONE_LABEL,
)

MAX_DRAWING_DATASET_INDEX = 9
CLOSE_PARTS_OPENED_BY_JOURNAL = True
MAX_OCCURRENCES = 100000
MAX_LOAD_PASSES = 100

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

# --- BOM VISIBILITY (mirrors NXOpenBoMExtended.py, Journal 04 and 21) ---
IGNORE_KEYWORDS = ("CSYS", "COORDINATE", "DATUM", "REFERENCE", "SKELETON")
BOM_REFERENCE_ATTRIBUTES = ("REFERENCE_COMPONENT", "PLIST_IGNORE_MEMBER")
BOM_EXCLUSION_ATTRIBUTE = "CELESTICA_BOM_EXCLUDE_SUBTREE"
BOM_EXCLUSION_VALUE = "YES"
BOM_REFERENCE_FLAG_VALUES = ("", "YES", "1", "True", "true", "yes")

# --- load failure classification (mirrors Journal 21) ---
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

TRUE_VALUES = {"YES", "Y", "TRUE", "1", "X"}
FALSE_VALUES = {"", "NO", "N", "FALSE", "0"}
_INVALID_FILENAME_CHARS = '<>:"/\\|?*'

_DRAWING_SUFFIX_RE = re.compile(
    r"(?:^|[-_])DWG(\d+)(?:$|[^A-Z0-9])",
    re.IGNORECASE,
)

_RESULT_COLUMNS = (
    "ROW_TYPE",
    "RUN_TIMESTAMP",
    "JOURNAL_BUILD",
    "WRITE_MODE",
    "SCOPE_FILTER",
    "LOAD_MODE",
    "LEVEL",
    "DEEPEST_LEVEL",
    "OCCURRENCE_COUNT",
    "COMPONENT_PATH",
    "PART_NAME",
    "DB_PART_NO",
    "DB_PART_REV",
    "WAE_VERSION",
    "TARGET_KIND",
    "PART_KIND",
    "INITIAL_LOAD_STATE",
    "LOAD_ACTION",
    "FINAL_LOAD_STATE",
    "LOAD_STATUS",
    "DRAWING_SOURCE",
    "DRAWING_COUNT",
    "HAS_DRAWING",
    "PDF_REQUESTED",
    "PDF_RESULT",
    "PDF_FILE_COUNT",
    "PDF_FILES",
    "PDF_SKIPPED_FILES",
    "STEP_REQUESTED",
    "STEP_RESULT",
    "STEP_FILE",
    "STEP_FILE_SIZE_BYTES",
    "OVERALL_RESULT",
    "MESSAGE",
    "DURATION_SECONDS",
)


# ---------------------------------------------------------------------------
# Mode resolution
# ---------------------------------------------------------------------------


def resolve_write_mode():
    value = normalize_text(
        os.environ.get("NX_J36_MODE")
    ).upper() or WRITE_MODE
    if value not in ("DRY_RUN", "APPLY"):
        raise ValueError(
            "NX_J36_MODE must be DRY_RUN or APPLY, got: {0}".format(value)
        )
    return value


def resolve_scope_filter():
    value = normalize_text(
        os.environ.get("NX_J36_SCOPE")
    ).upper() or SCOPE_FILTER
    if value not in ("BOM", "ALL"):
        raise ValueError(
            "NX_J36_SCOPE must be BOM or ALL, got: {0}".format(value)
        )
    return value


def resolve_load_mode():
    value = normalize_text(
        os.environ.get("NX_J36_LOAD_MODE")
    ).upper() or LOAD_MODE
    if value not in ("LOAD", "REPORT_ONLY"):
        raise ValueError(
            "NX_J36_LOAD_MODE must be LOAD or REPORT_ONLY, got: {0}".format(
                value
            )
        )
    return value


def resolve_step_scope():
    value = normalize_text(
        os.environ.get("NX_J36_STEP_SCOPE")
    ).upper()
    if not value:
        return (
            "DRAWING_ONLY"
            if STEP_ONLY_FOR_TARGETS_WITH_DRAWING
            else "ALL"
        )
    if value not in ("DRAWING_ONLY", "ALL"):
        raise ValueError(
            "NX_J36_STEP_SCOPE must be DRAWING_ONLY or ALL, got: {0}".format(
                value
            )
        )
    return value


# ---------------------------------------------------------------------------
# Generic helpers
# ---------------------------------------------------------------------------


class TimestampCleanupError(RuntimeError):
    """Temporary PDF timestamp notes could not be proven undone."""


def normalize_text(value):
    return "" if value is None else str(value).strip()


def clean(value):
    return normalize_text(value)


def error_text(error):
    code = clean(getattr(error, "ErrorCode", ""))
    suffix = " [{0}]".format(code) if code else ""
    return "{0}{1}".format(clean(error) or type(error).__name__, suffix)


def contains_token(value, tokens):
    lowered = clean(value).lower()
    return any(token in lowered for token in tokens)


def classify_load_failure(details, default="LOAD_FAILED"):
    if contains_token(details, INVALID_OBJECT_TOKENS):
        return "INVALID_OBJECT"
    if contains_token(details, MISSING_FILE_TOKENS):
        return "MISSING_FILE"
    return default


def runtime_source_path():
    try:
        return os.path.abspath(__file__)
    except Exception:
        return "<unknown>"


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
    for method_name in ("Dispose", "FreeResource", "Destroy"):
        method = getattr(value, method_name, None)
        if callable(method):
            try:
                method()
            except Exception:
                pass
            return


def elapsed_seconds(started):
    return max(0.0, time.perf_counter() - started)


def build_export_timestamp_text(run_datetime):
    return "EXPORTED: {0} {1}".format(
        run_datetime.strftime(PDF_TIMESTAMP_FORMAT),
        PDF_TIMESTAMP_TIMEZONE_LABEL,
    )


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


def _object_key(nx_object):
    tag = getattr(nx_object, "Tag", None)
    return ("TAG", normalize_text(tag)) if tag is not None else ("PY", id(nx_object))


def _safe_property(nx_object, name, fallback=None):
    try:
        value = getattr(nx_object, name)
        return value() if callable(value) else value
    except Exception:
        return fallback


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

    return values


def part_kind(part):
    assembly = _safe_property(part, "ComponentAssembly")
    if assembly is None:
        return "PART"
    root = _safe_property(assembly, "RootComponent")
    return "ASSEMBLY" if root is not None else "PART"


def drawing_sheet_count(part):
    try:
        return int(part.DrawingSheets.Count)
    except Exception:
        pass

    try:
        return len(list(part.DrawingSheets))
    except Exception:
        return 0


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


def close_journal_opened_drawings(session, candidates, log_buffer):
    """Close only the drawing specifications this journal opened."""
    for candidate in candidates or []:
        if candidate.get("opened_by_journal"):
            close_part_best_effort(
                candidate.get("part"),
                session,
                log_buffer,
            )


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


# ---------------------------------------------------------------------------
# Traversal: BoM-visible scope with levels, paths and load records
# ---------------------------------------------------------------------------


def _component_name(component):
    return (
        clean(_safe_property(component, "DisplayName"))
        or clean(_safe_property(component, "Name"))
        or "<unknown>"
    )


def _children(component):
    """Return (children, error) so an unreadable branch cannot look empty."""
    try:
        return list(component.GetChildren()), ""
    except Exception as error:
        return [], error_text(error)


def _component_string_attribute(component, title):
    """Safe read of a component-level string attribute; None when absent."""
    try:
        return component.GetStringAttribute(title)
    except Exception:
        return None


def _is_bom_visible(component):
    """Mirror NXOpenBoMExtended.py: only BoM-visible components are packaged.

    Suppression is handled separately. Keyword-named and reference-flagged
    occurrences are excluded together with their subtrees.
    """
    name = clean(getattr(component, "Name", ""))
    display_name = clean(getattr(component, "DisplayName", ""))
    combined = " ".join((name, display_name)).upper()
    for keyword in IGNORE_KEYWORDS:
        if keyword in combined:
            return False
    for title in BOM_REFERENCE_ATTRIBUTES:
        raw = _component_string_attribute(component, title)
        if raw is not None and clean(raw) in BOM_REFERENCE_FLAG_VALUES:
            return False
    custom = _component_string_attribute(component, BOM_EXCLUSION_ATTRIBUTE)
    if custom is not None and clean(custom) == BOM_EXCLUSION_VALUE:
        return False
    return True


def _is_active_visible(component):
    """Suppression state; unreadable suppression is treated as not active."""
    try:
        return not bool(component.IsSuppressed)
    except Exception:
        return False


def collect_session_scope(work_part, scope_filter="BOM"):
    """Return unique packaging targets and traversal diagnostics.

    Each target records its topmost observed level and path, the deepest
    level it was seen at, and how many occurrences share the prototype.

    Returns (targets, diagnostics). targets is a list of dicts sorted by
    (level, part number, name).
    """
    targets = {}
    diagnostics = []
    occurrence_count = 0

    root_number, _root_revision = get_part_identity(work_part)
    root_path = root_number or safe_part_name(work_part)

    def add_target(part, level, path):
        key = _object_key(part)
        record = targets.get(key)
        if record is None:
            targets[key] = {
                "key": key,
                "part": part,
                "level": level,
                "deepest_level": level,
                "component_path": path,
                "occurrence_count": 1,
                "is_work_part": False,
            }
            return
        record["occurrence_count"] += 1
        record["deepest_level"] = max(record["deepest_level"], level)
        if level < record["level"]:
            record["level"] = level
            record["component_path"] = path

    add_target(work_part, 0, root_path)
    targets[_object_key(work_part)]["is_work_part"] = True

    try:
        root_component = getattr(
            getattr(work_part, "ComponentAssembly", None),
            "RootComponent",
            None,
        )
    except Exception as error:
        diagnostics.append(
            {
                "code": classify_load_failure(
                    error_text(error), "ROOT_COMPONENT_UNREADABLE"
                ),
                "message": "Assembly root could not be read: " + error_text(error),
                "component_path": root_path,
                "level": 0,
            }
        )
        return _sorted_targets(targets), diagnostics

    if root_component is None:
        return _sorted_targets(targets), diagnostics

    root_children, root_error = _children(root_component)
    if root_error:
        diagnostics.append(
            {
                "code": "CHILDREN_UNREADABLE",
                "message": "Assembly root children could not be read: " + root_error,
                "component_path": root_path,
                "level": 0,
            }
        )
        return _sorted_targets(targets), diagnostics

    stack = [
        (component, 1, root_path)
        for component in reversed(root_children)
    ]
    while stack:
        component, level, parent_path = stack.pop()
        occurrence_count += 1
        component_path = "{0} / {1}".format(
            parent_path, _component_name(component)
        )

        if occurrence_count > MAX_OCCURRENCES:
            diagnostics.append(
                {
                    "code": "OCCURRENCE_LIMIT",
                    "message": "Traversal exceeded {0} occurrences.".format(
                        MAX_OCCURRENCES
                    ),
                    "component_path": component_path,
                    "level": level,
                }
            )
            break

        if scope_filter == "BOM":
            if not _is_active_visible(component):
                continue
            if not _is_bom_visible(component):
                continue

        try:
            prototype = getattr(component, "Prototype", None)
        except Exception as error:
            diagnostics.append(
                {
                    "code": classify_load_failure(
                        error_text(error), "PROTOTYPE_UNAVAILABLE"
                    ),
                    "message": "Component prototype could not be read: "
                    + error_text(error),
                    "component_path": component_path,
                    "level": level,
                }
            )
            prototype = None

        if prototype is None:
            diagnostics.append(
                {
                    "code": "MISSING_MODEL",
                    "message": (
                        "Component has no loaded prototype: {0}".format(
                            _component_name(component)
                        )
                    ),
                    "component_path": component_path,
                    "level": level,
                }
            )
        else:
            add_target(prototype, level, component_path)

        children, children_error = _children(component)
        if children_error:
            diagnostics.append(
                {
                    "code": "CHILDREN_UNREADABLE",
                    "message": (
                        "Component children could not be read for {0}: {1}".format(
                            _component_name(component), children_error
                        )
                    ),
                    "component_path": component_path,
                    "level": level,
                }
            )
        stack.extend(
            (child, level + 1, component_path)
            for child in reversed(children)
        )

    return _sorted_targets(targets), diagnostics


def _sorted_targets(targets):
    return sorted(
        targets.values(),
        key=lambda item: (
            item["level"],
            get_part_identity(item["part"])[0].upper(),
            safe_part_name(item["part"]).upper(),
        ),
    )


def part_load_state(part):
    fully_loaded = _safe_property(part, "IsFullyLoaded")
    state = clean(_safe_property(part, "PartLoadState"))
    if fully_loaded is None:
        return "UNKNOWN", state
    try:
        return (
            "FULLY_LOADED" if bool(fully_loaded) else "NOT_FULLY_LOADED",
            state,
        )
    except Exception:
        return "UNKNOWN", state


def load_state_text(part):
    status, raw_state = part_load_state(part)
    return raw_state or status


def unwrap_load_status(value):
    if isinstance(value, (tuple, list)):
        return value[0] if value else None
    return value


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
            name = "<unavailable>"
        try:
            code = clean(load_status.GetStatus(index))
        except Exception:
            code = "<unavailable>"
        try:
            description = clean(load_status.GetStatusDescription(index))
        except Exception:
            description = "<unavailable>"
        details.append(
            "part={0}; status={1}; description={2}".format(
                name, code, description
            )
        )
    return details, count


def load_target(part, level, component_path, logger=None):
    identity_number, _identity_revision = get_part_identity(part)
    label = identity_number or safe_part_name(part)
    initial_status, initial_raw = part_load_state(part)
    record = {
        "part": part,
        "level": level,
        "component_path": component_path,
        "initial_load_state": initial_raw or initial_status,
        "load_action": "NOT_REQUIRED",
        "final_load_state": initial_raw or initial_status,
        "load_status": "SUCCESS",
        "load_message": "Part was already fully loaded.",
    }
    if initial_status == "FULLY_LOADED":
        return record

    method = getattr(part, "LoadThisPartFully", None)
    if not callable(method):
        record.update(
            {
                "load_action": "LOAD_THIS_PART_FULLY",
                "load_status": "API_UNAVAILABLE",
                "load_message": "BasePart.LoadThisPartFully is unavailable.",
            }
        )
        return record

    if logger:
        logger("FULL LOAD {0}: {1}".format(label, component_path))
    record["load_action"] = "LOAD_THIS_PART_FULLY"
    load_status = None
    try:
        load_status = unwrap_load_status(method())
        details, unloaded_count = part_load_status_details(load_status)
        final_status, final_raw = part_load_state(part)
        record["final_load_state"] = final_raw or final_status
        if unloaded_count:
            detail_text = " | ".join(details)
            record["load_status"] = classify_load_failure(
                detail_text, "PROTOTYPE_UNAVAILABLE"
            )
            record["load_message"] = detail_text
        elif final_status != "FULLY_LOADED":
            record["load_status"] = "UNLOADED"
            record["load_message"] = (
                "LoadThisPartFully returned, but IsFullyLoaded is not True. "
                + " | ".join(details)
            ).strip()
        else:
            record["load_status"] = "SUCCESS"
            record["load_message"] = " | ".join(details) or "Fully loaded."
    except Exception as error:
        details = error_text(error)
        record["final_load_state"] = load_state_text(part)
        record["load_status"] = classify_load_failure(details)
        record["load_message"] = details
    finally:
        dispose(load_status)
    return record


def load_session_scope(work_part, scope_filter="BOM", load_mode="LOAD", logger=None):
    """Load BoM-visible targets, re-traversing until the scope is stable.

    Returns (loaded_ok, targets, records, diagnostics). Records are keyed by
    object key. Mirrors the Journal 21 load gate; parts that could not load
    are reported and never abort sibling branches.
    """
    records = {}
    final_targets = []
    final_diagnostics = []

    if load_mode == "REPORT_ONLY":
        final_targets, final_diagnostics = collect_session_scope(
            work_part, scope_filter
        )
        for target in final_targets:
            key = target["key"]
            initial_status, initial_raw = part_load_state(target["part"])
            records[key] = {
                "part": target["part"],
                "level": target["level"],
                "component_path": target["component_path"],
                "initial_load_state": initial_raw or initial_status,
                "load_action": "NOT_ATTEMPTED",
                "final_load_state": initial_raw or initial_status,
                "load_status": "NOT_EVALUATED",
                "load_message": "NX_J36_LOAD_MODE=REPORT_ONLY",
            }
        return True, final_targets, records, final_diagnostics

    for pass_index in range(1, MAX_LOAD_PASSES + 1):
        targets, diagnostics = collect_session_scope(
            work_part, scope_filter
        )
        attempted = False
        for target in targets:
            key = target["key"]
            if key in records:
                if target["level"] < records[key]["level"]:
                    records[key]["level"] = target["level"]
                    records[key]["component_path"] = target["component_path"]
                continue
            record = load_target(
                target["part"],
                target["level"],
                target["component_path"],
                logger=logger,
            )
            records[key] = record
            if record["load_action"] == "LOAD_THIS_PART_FULLY":
                attempted = True

        final_targets, final_diagnostics = collect_session_scope(
            work_part, scope_filter
        )
        final_keys = {target["key"] for target in final_targets}
        known_keys = set(records)
        all_recorded = final_keys.issubset(known_keys)
        all_loaded = all(
            records[key]["load_status"] == "SUCCESS"
            and part_load_state(records[key]["part"])[0] == "FULLY_LOADED"
            for key in final_keys
            if key in records
        )
        if all_recorded and all_loaded and not final_diagnostics:
            return True, final_targets, records, []

        new_targets_exist = not all_recorded
        if not attempted and not new_targets_exist:
            break

        if logger:
            logger(
                "FULL LOAD PASS {0}: targets={1}; new_targets={2}".format(
                    pass_index,
                    len(final_targets),
                    "YES" if new_targets_exist else "NO",
                )
            )
    else:
        final_diagnostics.append(
            {
                "code": "LOAD_PASS_LIMIT",
                "message": "Full-load discovery exceeded {0} passes.".format(
                    MAX_LOAD_PASSES
                ),
                "component_path": "",
                "level": "",
            }
        )

    failures = []
    for key, record in records.items():
        if record["load_status"] != "SUCCESS":
            number, _revision = get_part_identity(record["part"])
            label = number or safe_part_name(record["part"])
            failures.append(
                {
                    "code": record["load_status"],
                    "message": "{0}: {1}".format(
                        label, record["load_message"]
                    ),
                    "component_path": record["component_path"],
                    "level": record["level"],
                }
            )

    final_keys = {target["key"] for target in final_targets}
    for key, record in records.items():
        if key not in final_keys:
            final_targets.append(
                {
                    "key": key,
                    "part": record["part"],
                    "level": record["level"],
                    "deepest_level": record["level"],
                    "component_path": record["component_path"],
                    "occurrence_count": 1,
                    "is_work_part": False,
                }
            )

    return (
        False,
        _sorted_targets({target["key"]: target for target in final_targets}),
        records,
        final_diagnostics + failures,
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


def build_versioned_base(number, revision, wae_version):
    base = "{0}_REV{1}".format(
        clean_filename_token(number),
        clean_filename_token(revision, fallback=""),
    )

    if wae_version:
        cleaned_version = clean_filename_token(
            wae_version,
            fallback="",
        )
        if cleaned_version:
            base += "." + cleaned_version

    return base


def build_pdf_filename(
    number,
    revision,
    wae_version,
    token,
    drawing_count,
):
    filename = build_versioned_base(
        number,
        revision,
        wae_version,
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


def resolve_pdf_watermark(target_part, candidates):
    """Resolve the DRAFT watermark and WAE_VERSION for one target."""
    wae_version = get_string_attribute(
        target_part,
        WAE_VERSION_ATTRIBUTE,
    )
    if wae_version:
        return (
            wae_version,
            "target {0}".format(WAE_VERSION_ATTRIBUTE),
            "",
        )

    for candidate in candidates:
        wae_version = get_string_attribute(
            candidate.get("part"),
            WAE_VERSION_ATTRIBUTE,
        )
        if wae_version:
            return (
                wae_version,
                "drawing {0}".format(WAE_VERSION_ATTRIBUTE),
                "",
            )

    return (
        "",
        "revision-only fallback",
        "{0} is blank or unavailable; PDF exported with the "
        "revision-only watermark.".format(WAE_VERSION_ATTRIBUTE),
    )


def sheet_uses_inches(sheet):
    units = getattr(sheet, "Units", None)

    drawings = getattr(NXOpen, "Drawings", None)
    for sheet_type_name in ("DrawingSheet", "DraftingDrawingSheet"):
        sheet_type = getattr(drawings, sheet_type_name, None)
        unit_enum = getattr(sheet_type, "Unit", None)
        for member_name in ("Inches", "UnitInches"):
            member = getattr(unit_enum, member_name, None)
            if member is not None and units == member:
                return True
        for member_name in ("Millimeters", "UnitMillimeters"):
            member = getattr(unit_enum, member_name, None)
            if member is not None and units == member:
                return False

    text = normalize_text(getattr(units, "name", units)).upper()
    if "INCH" in text or "ENGLISH" in text:
        return True
    if "MILLIM" in text or "METRIC" in text:
        return False

    numeric_units = getattr(
        units,
        "value",
        getattr(units, "Value", units),
    )
    numeric_text = normalize_text(numeric_units)
    if numeric_units == 1 or numeric_text == "1":
        return True
    if numeric_units == 2 or numeric_text == "2":
        return False

    raise RuntimeError(
        "Unsupported or unavailable drawing-sheet units: {0}".format(
            units
        )
    )


def millimeters_to_sheet_units(value_mm, sheet):
    return float(value_mm) / 25.4 if sheet_uses_inches(sheet) else float(value_mm)


def current_drawing_sheet(drawing_part):
    try:
        return drawing_part.DrawingSheets.CurrentDrawingSheet
    except Exception:
        return None


def part_modified_state(part):
    try:
        return bool(part.IsModified)
    except Exception:
        return None


def set_note_text(note_builder, text):
    text_block = note_builder.Text.TextBlock
    try:
        text_block.SetText([text])
    except TypeError:
        text_block.SetText(text)


def create_pdf_timestamp_note(drawing_part, sheet, timestamp_text):
    sheet.Open()

    note_builder = drawing_part.Annotations.CreateDraftingNoteBuilder(None)
    try:
        text_height = millimeters_to_sheet_units(
            PDF_TIMESTAMP_TEXT_HEIGHT_MM,
            sheet,
        )
        right_margin = millimeters_to_sheet_units(
            PDF_TIMESTAMP_RIGHT_MARGIN_MM,
            sheet,
        )
        bottom_margin = millimeters_to_sheet_units(
            PDF_TIMESTAMP_BOTTOM_MARGIN_MM,
            sheet,
        )
        sheet_length = float(sheet.Length)
        sheet_height = float(sheet.Height)

        if sheet_length <= right_margin or sheet_height <= bottom_margin:
            raise RuntimeError(
                "Sheet is too small for the configured PDF timestamp "
                "margins: length={0}, height={1}".format(
                    sheet_length,
                    sheet_height,
                )
            )

        note_builder.Origin.Anchor = (
            NXOpen.Annotations.OriginBuilder.AlignmentPosition.BottomRight
        )
        note_builder.Origin.OriginPoint = NXOpen.Point3d(
            sheet_length - right_margin,
            bottom_margin,
            0.0,
        )
        note_builder.Style.LetteringStyle.GeneralTextSize = text_height
        try:
            note_builder.Style.LetteringStyle.GeneralTextLineWidth = (
                NXOpen.Annotations.LineWidth.Normal
            )
            note_builder.Style.LetteringStyle.HorizontalTextJustification = (
                NXOpen.Annotations.TextJustification.Right
            )
        except Exception:
            # Some NX releases do not expose this drafting preference through
            # Python. The inherited normal drafting width remains acceptable.
            pass
        set_note_text(note_builder, timestamp_text)
        return note_builder.Commit()
    finally:
        note_builder.Destroy()


def create_timestamp_notes(
    session,
    drawing_part,
    sheets,
    timestamp_text,
    undo_mark,
):
    notes = []
    for sheet in sheets:
        notes.append(
            create_pdf_timestamp_note(
                drawing_part,
                sheet,
                timestamp_text,
            )
        )

    update_manager = getattr(session, "UpdateManager", None)
    do_update = getattr(update_manager, "DoUpdate", None)
    if callable(do_update):
        error_count = int(do_update(undo_mark) or 0)
        if error_count:
            raise RuntimeError(
                "NX reported {0} error(s) while updating temporary "
                "PDF timestamp notes.".format(error_count)
            )
    return notes


def restore_original_sheet(original_sheet):
    if original_sheet is not None:
        original_sheet.Open()


def undo_timestamp_notes(
    session,
    undo_mark,
    undo_mark_name,
    drawing_part,
    initially_modified,
    original_sheet,
):
    errors = []

    try:
        session.UndoToMark(undo_mark, undo_mark_name)
    except Exception as error:
        errors.append("undo failed: {0}".format(error))

    try:
        restore_original_sheet(original_sheet)
    except Exception as error:
        errors.append("active-sheet restore failed: {0}".format(error))

    try:
        session.DeleteUndoMark(undo_mark, undo_mark_name)
    except Exception as error:
        errors.append("undo-mark deletion failed: {0}".format(error))

    final_modified = part_modified_state(drawing_part)
    if (
        initially_modified is not None
        and final_modified is not None
        and final_modified != initially_modified
    ):
        errors.append(
            "drawing modified state changed from {0} to {1}".format(
                initially_modified,
                final_modified,
            )
        )

    if errors:
        raise TimestampCleanupError(
            "Temporary PDF timestamp cleanup could not be proven; "
            "discard unsaved drawing changes before retrying: "
            + " | ".join(errors)
        )


def export_drawing_pdf(
    session,
    drawing_part,
    sheets,
    output_path,
    watermark,
    export_timestamp_text,
):
    if not sheets:
        raise RuntimeError("Drawing contains no sheets to export.")
    if PDF_APPLY_DRAFT_WATERMARK and not normalize_text(watermark):
        raise RuntimeError("A PDF draft watermark is required.")

    metrics = {
        "timestamp_prepare_seconds": 0.0,
        "pdf_commit_seconds": 0.0,
        "timestamp_cleanup_seconds": 0.0,
        "pdf_total_seconds": 0.0,
    }
    total_started = time.perf_counter()
    original_sheet = current_drawing_sheet(drawing_part)
    initially_modified = part_modified_state(drawing_part)
    undo_mark_name = "J36 temporary PDF export timestamp"
    undo_mark = session.SetUndoMark(
        NXOpen.Session.MarkVisibility.Invisible,
        undo_mark_name,
    )
    export_error = None

    try:
        prepare_started = time.perf_counter()
        if PDF_ADD_EXPORT_TIMESTAMP_NOTE:
            create_timestamp_notes(
                session,
                drawing_part,
                sheets,
                export_timestamp_text,
                undo_mark,
            )
        metrics["timestamp_prepare_seconds"] = elapsed_seconds(
            prepare_started
        )

        sheets[0].Open()
        builder = drawing_part.PlotManager.CreatePrintPdfbuilder()
        try:
            builder.Action = NXOpen.PrintPDFBuilder.ActionOption.Native
            builder.Filename = output_path
            builder.Append = False
            try:
                builder.OutputText = (
                    NXOpen.PrintPDFBuilder.OutputTextOption.Text
                )
            except Exception as error:
                raise RuntimeError(
                    "NX PrintPDFBuilder could not apply required searchable "
                    "text output: {0}".format(error)
                )
            if PDF_APPLY_DRAFT_WATERMARK:
                try:
                    builder.AddWatermark = True
                    builder.Watermark = watermark
                    builder.CustomSymbolsInForeground = True
                except Exception as error:
                    raise RuntimeError(
                        "NX PrintPDFBuilder could not apply required native "
                        "watermark {0} and foreground symbols: {1}".format(
                            watermark,
                            error,
                        )
                    )
            builder.SourceBuilder.SetSheets(sheets)
            commit_started = time.perf_counter()
            builder.Commit()
            metrics["pdf_commit_seconds"] = elapsed_seconds(commit_started)
        finally:
            builder.Destroy()
    except Exception as error:
        export_error = error
    finally:
        cleanup_started = time.perf_counter()
        try:
            undo_timestamp_notes(
                session,
                undo_mark,
                undo_mark_name,
                drawing_part,
                initially_modified,
                original_sheet,
            )
        except TimestampCleanupError as cleanup_error:
            if export_error is not None:
                raise TimestampCleanupError(
                    "{0} Original PDF export error: {1}".format(
                        cleanup_error,
                        export_error,
                    )
                )
            raise
        finally:
            metrics["timestamp_cleanup_seconds"] = elapsed_seconds(
                cleanup_started
            )
            metrics["pdf_total_seconds"] = elapsed_seconds(total_started)

    if export_error is not None:
        raise export_error

    commit_seconds = metrics["pdf_commit_seconds"]
    timestamp_seconds = (
        metrics["timestamp_prepare_seconds"]
        + metrics["timestamp_cleanup_seconds"]
    )
    metrics["timestamp_overhead_percent"] = (
        (timestamp_seconds / commit_seconds) * 100.0
        if commit_seconds > 0.0
        else None
    )
    return metrics


def planned_pdf_exports(number, revision, wae_version, candidates):
    """Return [(token, output_path)] for every drawing of one target."""
    output_tokens = unique_drawing_tokens(candidates)
    drawing_count = len(candidates)
    return [
        (
            token,
            build_pdf_filename(
                number,
                revision,
                wae_version,
                token,
                drawing_count,
            ),
        )
        for token in output_tokens
    ]


def export_pdfs_for_target(
    session,
    output_folder,
    number,
    revision,
    target_part,
    candidates,
    export_timestamp_text,
    original_display,
    original_work,
    log_buffer,
):
    """Export one PDF per resolved drawing specification of one target."""
    if not candidates:
        return {
            "result": "SKIPPED_NO_DRAWING",
            "paths": [],
            "message": (
                "No drawing specification could be resolved or opened."
            ),
            "failures": [],
            "halt_pdf_batch": False,
            "source": "",
        }

    wae_version, watermark_source, watermark_warning = resolve_pdf_watermark(
        target_part,
        candidates,
    )
    watermark = build_pdf_watermark(revision, wae_version)
    messages = []
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
        log_line(session, "  WARNING: " + watermark_warning, log_buffer)

    successful_paths = []
    skipped_paths = []
    failures = []
    sources = []
    plans = planned_pdf_exports(
        number,
        revision,
        wae_version,
        candidates,
    )

    try:
        for candidate, (token, filename) in zip(candidates, plans):
            drawing_part = candidate["part"]
            sources.append(
                "{0}={1}".format(
                    token,
                    candidate.get("source") or "unknown",
                )
            )

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
                        "message": (
                            "{0}: could not enumerate drawing sheets: {1}"
                        ).format(token, error),
                        "traceback": traceback.format_exc(),
                    }
                )
                continue

            if not sheets:
                failures.append(
                    {
                        "kind": "NOT_DRAWING",
                        "message": (
                            "{0}: opened part contains no drawing sheets"
                        ).format(token),
                    }
                )
                continue

            output_path = os.path.join(output_folder, filename)
            if os.path.exists(output_path):
                skipped_paths.append(output_path)
                log_line(
                    session,
                    "    PDF skipped, already exists: {0}".format(output_path),
                    log_buffer,
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

            try:
                pdf_metrics = export_drawing_pdf(
                    session,
                    drawing_part,
                    sheets,
                    output_path,
                    watermark,
                    export_timestamp_text,
                )
                overhead = pdf_metrics.get("timestamp_overhead_percent")
                if (
                    overhead is not None
                    and overhead > PDF_TIMESTAMP_OVERHEAD_TARGET_PERCENT
                ):
                    log_line(
                        session,
                        (
                            "    PERFORMANCE WARNING: timestamp overhead "
                            "{0:.1f}% exceeds the NX acceptance target "
                            "of {1:.1f}% for this drawing."
                        ).format(
                            overhead,
                            PDF_TIMESTAMP_OVERHEAD_TARGET_PERCENT,
                        ),
                        log_buffer,
                    )

                if VERIFY_OUTPUT_FILES and not os.path.isfile(output_path):
                    failures.append(
                        {
                            "kind": "NO_OUTPUT",
                            "message": (
                                "{0}: PDF builder committed but no file "
                                "was created"
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
            except TimestampCleanupError as error:
                failures.append(
                    {
                        "kind": "TIMESTAMP_CLEANUP",
                        "message": "{0}: {1}".format(token, error),
                        "traceback": traceback.format_exc(),
                    }
                )
                break
            except Exception as error:
                failures.append(
                    {
                        "kind": "ERROR",
                        "message": "{0}: {1}".format(token, error),
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

    halt_pdf_batch = any(
        failure["kind"] == "TIMESTAMP_CLEANUP" for failure in failures
    )

    if halt_pdf_batch:
        result = "FAILED_TIMESTAMP_CLEANUP"
    elif successful_paths and not failures and not skipped_paths:
        result = "SUCCESS"
    elif successful_paths:
        result = "PARTIAL_SUCCESS"
    elif skipped_paths and not failures:
        result = "SKIPPED_EXISTS"
    elif failures and all(
        failure["kind"] == "NO_OUTPUT" for failure in failures
    ):
        result = "FAILED_NO_OUTPUT_FILE"
    elif failures and all(
        failure["kind"] == "NOT_DRAWING" for failure in failures
    ):
        result = "SKIPPED_NO_DRAWING"
    else:
        result = "FAILED"

    return {
        "result": result,
        "paths": successful_paths,
        "skipped": skipped_paths,
        "message": " | ".join(
            messages
            + [failure["message"] for failure in failures]
        ),
        "failures": failures,
        "halt_pdf_batch": halt_pdf_batch,
        "source": " | ".join(sources),
        "wae_version": wae_version,
    }


# ---------------------------------------------------------------------------
# STEP export
# ---------------------------------------------------------------------------


def step_body_signature_count(path):
    signatures = 0
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
            for token in STEP_BODY_TOKENS:
                signatures += upper.count(token)
    return signatures


def export_step_from_part(
    session,
    part,
    output_folder,
    number,
    revision,
    wae_version,
):
    output_path = os.path.join(
        output_folder,
        build_versioned_base(
            number,
            revision,
            wae_version,
        )
        + ".stp",
    )

    if os.path.exists(output_path):
        return {
            "result": "SKIPPED_EXISTS",
            "path": output_path,
            "size": "",
            "message": "STEP output already exists; left unchanged.",
        }

    set_display_part(session, part)
    session.Parts.SetWork(part)

    exporter = session.DexManager.CreateStepCreator()
    try:
        exporter.OutputFile = output_path
        # Journals 10 and 07 proved this exact display/scope/layer
        # combination: the gasket changes from zero input solids to one
        # processed solid.
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

    if VERIFY_OUTPUT_FILES:
        signatures = step_body_signature_count(output_path)
        if signatures <= 0:
            return {
                "result": "FAILED_ZERO_GEOMETRY",
                "path": output_path,
                "size": file_size,
                "message": (
                    "STEP output contains no body geometry signatures; "
                    "the header-only file was retained for diagnosis"
                ),
            }

    return {
        "result": "SUCCESS",
        "path": output_path,
        "size": file_size,
        "message": "",
    }


# ---------------------------------------------------------------------------
# Result rows
# ---------------------------------------------------------------------------


def step_requested_for_target(has_drawing, step_scope):
    if step_scope == "ALL":
        return True
    return bool(has_drawing)


def new_result(timestamp, write_mode, scope_filter, load_mode, target):
    identity_number, identity_revision = get_part_identity(target["part"])
    return {
        "ROW_TYPE": "TARGET",
        "RUN_TIMESTAMP": timestamp,
        "JOURNAL_BUILD": JOURNAL_BUILD_ID,
        "WRITE_MODE": write_mode,
        "SCOPE_FILTER": scope_filter,
        "LOAD_MODE": load_mode,
        "LEVEL": target["level"],
        "DEEPEST_LEVEL": target.get("deepest_level", target["level"]),
        "OCCURRENCE_COUNT": target.get("occurrence_count", 1),
        "COMPONENT_PATH": target["component_path"],
        "PART_NAME": safe_part_name(target["part"]),
        "DB_PART_NO": identity_number,
        "DB_PART_REV": identity_revision,
        "WAE_VERSION": "",
        "TARGET_KIND": (
            "WORK_PART" if target.get("is_work_part") else "COMPONENT_PROTOTYPE"
        ),
        "PART_KIND": part_kind(target["part"]),
        "INITIAL_LOAD_STATE": "",
        "LOAD_ACTION": "",
        "FINAL_LOAD_STATE": "",
        "LOAD_STATUS": "",
        "DRAWING_SOURCE": "",
        "DRAWING_COUNT": 0,
        "HAS_DRAWING": "NO",
        "PDF_REQUESTED": "YES",
        "PDF_RESULT": "PENDING",
        "PDF_FILE_COUNT": 0,
        "PDF_FILES": "",
        "PDF_SKIPPED_FILES": "",
        "STEP_REQUESTED": "YES",
        "STEP_RESULT": "PENDING",
        "STEP_FILE": "",
        "STEP_FILE_SIZE_BYTES": "",
        "OVERALL_RESULT": "PENDING",
        "MESSAGE": "",
        "DURATION_SECONDS": "",
    }


def diagnostic_row(timestamp, write_mode, scope_filter, load_mode, diagnostic):
    row = {
        "ROW_TYPE": "DIAGNOSTIC",
        "RUN_TIMESTAMP": timestamp,
        "JOURNAL_BUILD": JOURNAL_BUILD_ID,
        "WRITE_MODE": write_mode,
        "SCOPE_FILTER": scope_filter,
        "LOAD_MODE": load_mode,
    }
    for column in _RESULT_COLUMNS:
        row.setdefault(column, "")
    row["LEVEL"] = diagnostic.get("level", "")
    row["COMPONENT_PATH"] = diagnostic.get("component_path", "")
    row["OVERALL_RESULT"] = diagnostic.get("code", "DIAGNOSTIC")
    row["MESSAGE"] = diagnostic.get("message", "")
    return row


def summary_row(timestamp, write_mode, scope_filter, load_mode, counts, totals):
    row = {
        "ROW_TYPE": "SUMMARY",
        "RUN_TIMESTAMP": timestamp,
        "JOURNAL_BUILD": JOURNAL_BUILD_ID,
        "WRITE_MODE": write_mode,
        "SCOPE_FILTER": scope_filter,
        "LOAD_MODE": load_mode,
    }
    for column in _RESULT_COLUMNS:
        row.setdefault(column, "")
    row["PDF_FILE_COUNT"] = totals.get("pdf_files", 0)
    row["STEP_FILE_SIZE_BYTES"] = totals.get("step_bytes", 0)
    row["OVERALL_RESULT"] = " | ".join(
        "{0}={1}".format(status, count)
        for status, count in sorted(counts.items())
    )
    row["MESSAGE"] = " | ".join(
        "{0}={1}".format(key, value)
        for key, value in sorted(totals.items())
    )
    return row


def overall_result(pdf_result, step_result, write_mode="APPLY"):
    if write_mode == "DRY_RUN":
        if pdf_result == "SKIPPED_NO_DRAWING" and step_result == "SKIPPED_NO_DRAWING":
            return "SKIPPED_NO_DRAWING"
        return "DRY_RUN"

    results = (pdf_result, step_result)
    if all(result == "SUCCESS" for result in results):
        return "SUCCESS"
    if any(result in ("SUCCESS", "PARTIAL_SUCCESS") for result in results):
        return "PARTIAL_SUCCESS"
    if all(result == "SKIPPED_NO_DRAWING" for result in results):
        return "SKIPPED_NO_DRAWING"
    if all(result == "SKIPPED_EXISTS" for result in results):
        return "SKIPPED_EXISTS"
    if all(
        result in ("SKIPPED_NO_DRAWING", "SKIPPED_EXISTS")
        for result in results
    ):
        return "SKIPPED_PARTIAL"
    return "FAILED"


def write_result_csv(path, results):
    with open(path, "w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=_RESULT_COLUMNS)
        writer.writeheader()
        for result in results:
            writer.writerow(
                {column: result.get(column, "") for column in _RESULT_COLUMNS}
            )


# ---------------------------------------------------------------------------
# Per-target processing
# ---------------------------------------------------------------------------


def process_target(
    session,
    target,
    folders,
    timestamp,
    export_timestamp_text,
    write_mode,
    scope_filter,
    load_mode,
    step_scope,
    load_record,
    pdf_batch_state,
    original_display,
    original_work,
    log_buffer,
):
    started = datetime.datetime.now()
    result = new_result(timestamp, write_mode, scope_filter, load_mode, target)
    messages = []
    part = target["part"]

    if load_record:
        result["INITIAL_LOAD_STATE"] = load_record.get(
            "initial_load_state", ""
        )
        result["LOAD_ACTION"] = load_record.get("load_action", "")
        result["FINAL_LOAD_STATE"] = load_record.get("final_load_state", "")
        result["LOAD_STATUS"] = load_record.get("load_status", "")
        if load_record.get("load_status") not in ("SUCCESS", ""):
            append_unique(
                messages,
                "LOAD {0}: {1}".format(
                    load_record.get("load_status"),
                    load_record.get("load_message", ""),
                ),
            )

    raw_number, revision = get_part_identity(part)
    has_identity = bool(raw_number)
    number = raw_number or safe_part_name(part)
    # The report shows the token actually used for output naming, so a
    # local/unmanaged part is still identifiable in the result CSV.
    result["DB_PART_NO"] = number
    if not has_identity:
        append_unique(
            messages,
            "DB_PART_NO is unavailable; using the part name for output "
            "naming and skipping the Teamcenter drawing fallback.",
        )

    # --- drawings -------------------------------------------------------
    candidates = []
    attempts = []

    if has_identity:
        try:
            candidates, attempts = resolve_drawing_candidates(
                session,
                number,
                revision,
                log_buffer,
            )
        except Exception as error:
            append_unique(
                messages,
                "Drawing resolution failed: {0}".format(error),
            )
            log_line(session, traceback.format_exc(), log_buffer)
            candidates = []

    result["DRAWING_SOURCE"] = " | ".join(
        "{0}={1}".format(
            drawing_token(
                candidate["part"],
                candidate.get("drawing_index"),
            ),
            candidate.get("source") or "unknown",
        )
        for candidate in candidates
    )
    result["DRAWING_COUNT"] = len(candidates)
    result["HAS_DRAWING"] = "YES" if candidates else "NO"

    has_drawing = bool(candidates)
    step_requested = step_requested_for_target(has_drawing, step_scope)
    result["STEP_REQUESTED"] = "YES" if step_requested else "NO"
    wae_version, _watermark_source, _watermark_warning = resolve_pdf_watermark(
        part, candidates
    )
    result["WAE_VERSION"] = wae_version

    if not candidates:
        result["PDF_RESULT"] = "SKIPPED_NO_DRAWING"
        if not step_requested:
            result["STEP_RESULT"] = "SKIPPED_NO_DRAWING"
        append_unique(
            messages,
            "No drawing specification resolved for this level"
            + (
                "; STEP skipped because NX_J36_STEP_SCOPE=DRAWING_ONLY."
                if not step_requested
                else "."
            ),
        )
        if attempts:
            log_line(
                session,
                "  Drawing attempts ({0}): {1}".format(
                    len(attempts), " | ".join(attempts)
                ),
                log_buffer,
            )

    # --- DRY_RUN plan ---------------------------------------------------
    if write_mode == "DRY_RUN":
        pdf_paths = []
        if candidates:
            pdf_paths = [
                os.path.join(folders["pdf"], filename)
                for _token, filename in planned_pdf_exports(
                    number, revision, wae_version, candidates
                )
            ]
            existing = [
                path for path in pdf_paths if os.path.exists(path)
            ]
            if existing:
                result["PDF_RESULT"] = "PLANNED_SKIPPED_EXISTS"
                result["PDF_SKIPPED_FILES"] = ";".join(existing)
            else:
                result["PDF_RESULT"] = "PLANNED"
            result["PDF_FILES"] = ";".join(pdf_paths)
            result["PDF_FILE_COUNT"] = 0

        if step_requested:
            step_path = os.path.join(
                folders["step"],
                build_versioned_base(number, revision, wae_version) + ".stp",
            )
            if os.path.exists(step_path):
                result["STEP_RESULT"] = "PLANNED_SKIPPED_EXISTS"
            else:
                result["STEP_RESULT"] = "PLANNED"
            result["STEP_FILE"] = step_path
        else:
            result["STEP_RESULT"] = "SKIPPED_NO_DRAWING"

        if candidates and step_requested:
            append_unique(
                messages,
                "DRY_RUN would export {0} PDF(s) and 1 STEP.".format(
                    len(pdf_paths)
                ),
            )
        elif candidates:
            append_unique(
                messages,
                "DRY_RUN would export {0} PDF(s); STEP not requested "
                "for this level.".format(len(pdf_paths)),
            )
        elif step_requested:
            append_unique(
                messages,
                "DRY_RUN would export 1 STEP; no drawing was resolved for "
                "this level.",
            )

        result["OVERALL_RESULT"] = overall_result(
            result["PDF_RESULT"],
            result["STEP_RESULT"],
            write_mode="DRY_RUN",
        )
        result["MESSAGE"] = " | ".join(messages)
        result["DURATION_SECONDS"] = "{0:.3f}".format(
            (datetime.datetime.now() - started).total_seconds()
        )
        # DRY_RUN resolved (and therefore may have opened) Teamcenter
        # drawings; leave the session exactly as it was found.
        restore_parts(session, original_display, original_work, log_buffer)
        close_journal_opened_drawings(session, candidates, log_buffer)
        return result

    # --- APPLY: PDF -----------------------------------------------------
    if candidates:
        if pdf_batch_state["halted"]:
            result["PDF_RESULT"] = "FAILED_TIMESTAMP_CLEANUP"
            append_unique(
                messages,
                (
                    "PDF export skipped because temporary timestamp cleanup "
                    "failed earlier in this run. Discard unsaved drawing "
                    "changes before retrying. Cause: {0}"
                ).format(pdf_batch_state["reason"]),
            )
        else:
            try:
                pdf_export = export_pdfs_for_target(
                    session,
                    folders["pdf"],
                    number,
                    revision,
                    part,
                    candidates,
                    export_timestamp_text,
                    original_display,
                    original_work,
                    log_buffer,
                )
                result["PDF_RESULT"] = pdf_export["result"]
                result["PDF_FILE_COUNT"] = len(pdf_export["paths"])
                result["PDF_FILES"] = ";".join(pdf_export["paths"])
                result["PDF_SKIPPED_FILES"] = ";".join(
                    pdf_export.get("skipped", [])
                )
                append_unique(messages, pdf_export.get("message", ""))
                if pdf_export.get("halt_pdf_batch"):
                    pdf_batch_state["halted"] = True
                    pdf_batch_state["reason"] = (
                        pdf_export.get("message")
                        or "temporary timestamp cleanup failed"
                    )

                for failure in pdf_export.get("failures", []):
                    if failure.get("traceback"):
                        log_line(
                            session, failure["traceback"], log_buffer
                        )
            except Exception as error:
                result["PDF_RESULT"] = "FAILED"
                append_unique(
                    messages,
                    "PDF export failed: {0}".format(error),
                )
                log_line(session, traceback.format_exc(), log_buffer)

    # --- APPLY: STEP ----------------------------------------------------
    if step_requested:
        try:
            step_export = export_step_from_part(
                session,
                part,
                folders["step"],
                number,
                revision,
                wae_version,
            )
            result["STEP_RESULT"] = step_export["result"]
            result["STEP_FILE"] = step_export["path"]
            append_unique(messages, step_export.get("message", ""))
            if step_export.get("size") != "":
                result["STEP_FILE_SIZE_BYTES"] = step_export["size"]
                append_unique(
                    messages,
                    "STEP file size: {0} bytes".format(step_export["size"]),
                )
        except Exception as error:
            result["STEP_RESULT"] = "FAILED"
            append_unique(
                messages,
                "STEP export failed: {0}".format(error),
            )
            log_line(session, traceback.format_exc(), log_buffer)

    if result["PDF_RESULT"] == "PENDING":
        result["PDF_RESULT"] = "FAILED"
    if result["STEP_RESULT"] == "PENDING":
        result["STEP_RESULT"] = "FAILED"

    close_journal_opened_drawings(session, candidates, log_buffer)
    result["OVERALL_RESULT"] = overall_result(
        result["PDF_RESULT"],
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
    pdf_batch_state = {
        "halted": False,
        "reason": "",
    }
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

        write_mode = resolve_write_mode()
        scope_filter = resolve_scope_filter()
        load_mode = resolve_load_mode()
        step_scope = resolve_step_scope()

        io_root = resolve_io_root()
        run_datetime = datetime.datetime.now(MYT_TIMEZONE)
        timestamp = run_datetime.strftime("%Y%m%d_%H%M%S")
        export_timestamp_text = build_export_timestamp_text(run_datetime)
        folders = create_run_folders(io_root, timestamp)
        report_path = os.path.join(
            folders["reports"],
            "EXPORT_RESULT_{0}.csv".format(timestamp),
        )
        log_path = os.path.join(
            folders["logs"],
            "EXPORT_LOG_{0}.txt".format(timestamp),
        )

        log_line(
            session,
            "Journal 36 - Session-tree PDF + STEP packager",
            log_buffer,
        )
        log_line(session, "Journal build: " + JOURNAL_BUILD_ID, log_buffer)
        log_line(
            session,
            "Journal source: " + runtime_source_path(),
            log_buffer,
        )
        log_line(
            session,
            "Mode: {0}; scope: {1}; load: {2}; step scope: {3}".format(
                write_mode, scope_filter, load_mode, step_scope
            ),
            log_buffer,
        )
        log_line(session, "I/O root: " + io_root, log_buffer)
        log_line(session, "Run folder: " + folders["run"], log_buffer)
        log_line(
            session,
            "PDF export timestamp: " + export_timestamp_text,
            log_buffer,
        )
        log_line(
            session,
            (
                "Drawing resolver: loaded session sheets, then "
                "@DB/<part>/<rev>/specification/<part>-<rev>-dwg<n> "
                "via session.Parts.OpenDisplay"
            ),
            log_buffer,
        )
        log_line(
            session,
            (
                "Managed-mode flag: {0} (informational only; @DB opens are "
                "always attempted)"
            ).format(session_is_managed(session)),
            log_buffer,
        )

        work_part = session.Parts.Work
        if work_part is None:
            raise RuntimeError(
                "No work part is loaded. Open the assembly that should be "
                "packaged, make it the work part, and run the journal again."
            )
        log_line(
            session,
            "Assembly root: {0}".format(safe_part_name(work_part)),
            log_buffer,
        )

        # --- traverse + load gate ---------------------------------------
        assembly_root_label = safe_part_name(work_part)
        logger = lambda message: log_line(session, message, log_buffer)
        loaded_ok, targets, load_records, diagnostics = load_session_scope(
            work_part,
            scope_filter=scope_filter,
            load_mode=load_mode,
            logger=logger,
        )

        log_line(
            session,
            "Scope: {0} unique target(s); fully loaded: {1}; diagnostics: "
            "{2}".format(
                len(targets),
                "YES" if loaded_ok else "NO",
                len(diagnostics),
            ),
            log_buffer,
        )

        for diagnostic in diagnostics:
            log_line(
                session,
                "  {0} [{1}]: {2}".format(
                    diagnostic.get("code", "DIAGNOSTIC"),
                    diagnostic.get("component_path", ""),
                    diagnostic.get("message", ""),
                ),
                log_buffer,
            )

        for index, target in enumerate(targets, start=1):
            identity_number, identity_revision = get_part_identity(
                target["part"]
            )
            log_line(
                session,
                "[{0}/{1}] L{2} {3} / {4} ({5})".format(
                    index,
                    len(targets),
                    target["level"],
                    identity_number or safe_part_name(target["part"]),
                    identity_revision or "-",
                    part_kind(target["part"]),
                ),
                log_buffer,
            )

            result = process_target(
                session,
                target,
                folders,
                timestamp,
                export_timestamp_text,
                write_mode,
                scope_filter,
                load_mode,
                step_scope,
                load_records.get(target["key"]),
                pdf_batch_state,
                original_display,
                original_work,
                log_buffer,
            )
            results.append(result)

            log_line(
                session,
                "  PDF: {0} ({1} file(s)); STEP: {2}; overall: {3}".format(
                    result["PDF_RESULT"],
                    result["PDF_FILE_COUNT"],
                    result["STEP_RESULT"],
                    result["OVERALL_RESULT"],
                ),
                log_buffer,
            )

        restore_parts(
            session,
            original_display,
            original_work,
            log_buffer,
        )

        counts = {}
        for result in results:
            status = result["OVERALL_RESULT"]
            counts[status] = counts.get(status, 0) + 1

        level_counts = {}
        for result in results:
            level_key = "L{0}".format(result["LEVEL"])
            level_counts[level_key] = level_counts.get(level_key, 0) + 1

        totals = {
            "targets": len(results),
            "pdf_files": sum(
                int(result["PDF_FILE_COUNT"]) for result in results
            ),
            "step_files": sum(
                1 for result in results if result["STEP_RESULT"] == "SUCCESS"
            ),
            "step_bytes": sum(
                int(result["STEP_FILE_SIZE_BYTES"])
                for result in results
                if str(result["STEP_FILE_SIZE_BYTES"]).isdigit()
            ),
            "pdf_skipped_exists": sum(
                1
                for result in results
                if result["PDF_RESULT"] == "SKIPPED_EXISTS"
            ),
            "step_skipped_exists": sum(
                1
                for result in results
                if result["STEP_RESULT"] == "SKIPPED_EXISTS"
            ),
            "no_drawing": sum(
                1
                for result in results
                if result["HAS_DRAWING"] == "NO"
            ),
            "load_failures": sum(
                1
                for result in results
                if result["LOAD_STATUS"] not in ("SUCCESS", "")
            ),
        }

        for row in [
            diagnostic_row(
                timestamp, write_mode, scope_filter, load_mode, diagnostic
            )
            for diagnostic in diagnostics
        ]:
            results.append(row)
        results.append(
            summary_row(
                timestamp,
                write_mode,
                scope_filter,
                load_mode,
                counts,
                totals,
            )
        )

        write_result_csv(report_path, results)
        report_written = True

        log_line(session, "Packaging complete", log_buffer)
        log_line(
            session,
            "Assembly: {0}".format(assembly_root_label),
            log_buffer,
        )
        for status in sorted(counts):
            log_line(
                session,
                "  {0}: {1}".format(status, counts[status]),
                log_buffer,
            )
        for level_key in sorted(level_counts):
            log_line(
                session,
                "  {0}: {1} target(s)".format(
                    level_key, level_counts[level_key]
                ),
                log_buffer,
            )
        log_line(
            session,
            "  PDF files: {0}".format(totals["pdf_files"]),
            log_buffer,
        )
        log_line(
            session,
            "  STEP files: {0}".format(totals["step_files"]),
            log_buffer,
        )
        if write_mode == "DRY_RUN":
            log_line(
                session,
                "  DRY_RUN: no STEP/PDF file was written.",
                log_buffer,
            )
        log_line(session, "Result report: " + report_path, log_buffer)

        write_text_log(log_path, log_buffer)

    except Exception:
        log_line(
            session,
            "ERROR: Unhandled journal exception.",
            log_buffer,
        )
        log_line(session, traceback.format_exc(), log_buffer)

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
