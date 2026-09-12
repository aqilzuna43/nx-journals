"""
Journal 37 - Display Performance Triage (model view)

Read-only root-cause triage for a slow model view: "a simple view change
takes 10 minutes". It inventories everything NX must evaluate when the model
view regenerates, measures the actual view-change cost, and prints a ranked
suspect list where every line cites a measured value.

Run it on the assembly that is already open and fully loaded. It never loads
anything, so a partial load is itself reported as a finding instead of being
silently corrected.

What it measures
----------------
1. Visible-object census of the active model view (the strongest direct
   metric): View.AskVisibleObjects() counted by NX type, plus the view's
   rendering style.
2. Load truth: PartLoadState / IsFullyLoaded per prototype and
   HasAnyMinimallyLoadedChildren() / GetMinimallyLoadedParts() per part, so a
   Full Load that silently did not take is found immediately.
3. Per-occurrence display configuration for EVERY occurrence (not only
   BoM-visible ones, because reference-only occurrences still draw):
   RepresentationMode (Lightweight / Partial / Exact / None), ReferenceSet
   and the definitive "displays the entire part" test
   (ReferenceSet == EntirePartRefsetName), IsSuppressed, component layer,
   IsBlanked, UsedArrangement.
4. Geometry complexity per unique prototype: bodies (solid / sheet /
   convergent), faces, edges, facet count, vertex count, features,
   expressions.
5. Density per body (Body.Density): min / max / average, zero-density body
   count, unavailable count, material attribute.
6. Non-solid clutter: datums, coordinate systems, curves, lines, points,
   annotations.
7. Heavy display objects: True Shading, True Studio, point clouds, decals,
   cameras, dynamic sections, images, drawing sheets.
8. Layer census per part: state of every layer, and how many known objects
   sit on each layer, so "geometry spread across many layers" or "hidden
   layers still holding thousands of objects" is visible.
9. Display preferences that change how often a view update reloads or
   re-tessellates: LoadComponentOnFacetedViewUpdate / ...OnSelection,
   SmartlightweightViewsLoadComponentOnDemand, WorkPartDisplayAsEntirePart,
   RenderSolidsUsingStoredFacets, SaveAdvancedDisplayFacets, shading
   tolerances, ShowFacetEdges, DisplayUpdateReport.

Modes
-----
NX_J37_MODE=PROBE (default)  inventory only; the model view is never moved.
NX_J37_MODE=TIMED            PROBE plus measured view-change timing:
                             Regenerate / UpdateDisplay / N x Rotate / Fit,
                             each timed in seconds. The active view is
                             restored to its exact Matrix, Origin and Scale
                             afterwards.

Read-only contract
------------------
J37 does not:
- load, fully load, or open any part (no @DB access, no Teamcenter calls);
- change visibility, blanking, suppression, reference sets, layers,
  arrangements, or any preference;
- update, save, check out, or check in;
- create or modify attributes, geometry, or Teamcenter data.

TIMED mode rotates the active view and restores it exactly. Nothing is saved.
It deliberately does not call UpdateManager.DoUpdate: an update may apply
pending model changes, and this journal must not touch the model.

Evidence rules (same as J23)
----------------------------
A probe that is unavailable, unsupported on this NX build, or throws is
recorded as UNAVAILABLE or ERROR with the exact message. It is never
converted into 0, False, or a root cause. Suspect lines only ever quote a
measured value that was successfully read.

Target: NX 2312 and NX X 2506 embedded Python
Run via: NX > Tools > Journal > Play
"""

import csv
import datetime
import json
import os
import time
import traceback

import NXOpen


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

OUTPUT_ROOT_FOLDER = "NX_DISPLAY_PERF"
JOURNAL_BUILD_ID = "J37-NX2506-DISPLAY-PERF-TRIAGE-V2"

MODE = "PROBE"  # PROBE / TIMED
SCOPE_FILTER = "ALL"  # ALL / BOM
VISIBLE_SCAN = True
ROTATION_COUNT = 2
ROTATION_ANGLE_DEGREES = 15.0
ROTATION_AXIS = (0.0, 1.0, 0.0)
MAX_BODIES_PER_PART = 3000
MAX_VISIBLE_OBJECTS = 200000
MAX_LAYERS = 256
TOP_SUSPECTS = 25

MYT_TIMEZONE = datetime.timezone(
    datetime.timedelta(hours=8), name="MYT"
)

# --- suspect thresholds (measured values are always printed next to them) ---
THRESHOLD_FACE_COUNT = 5000
THRESHOLD_DATUM_CURVE_COUNT = 500
THRESHOLD_BLANKED_BODY_COUNT = 100
THRESHOLD_VISIBLE_OBJECTS = 50000
THRESHOLD_VISIBLE_LAYERS = 20
THRESHOLD_DENSITY_RATIO = 1000.0

# --- BoM visibility (mirrors NXOpenBoMExtended.py, Journal 04, 21, 36) ---
IGNORE_KEYWORDS = ("CSYS", "COORDINATE", "DATUM", "REFERENCE", "SKELETON")
BOM_REFERENCE_ATTRIBUTES = ("REFERENCE_COMPONENT", "PLIST_IGNORE_MEMBER")
BOM_EXCLUSION_ATTRIBUTE = "CELESTICA_BOM_EXCLUDE_SUBTREE"
BOM_EXCLUSION_VALUE = "YES"
BOM_REFERENCE_FLAG_VALUES = ("", "YES", "1", "True", "true", "yes")

MAX_OCCURRENCES = 100000

TRUE_VALUES = {"YES", "Y", "TRUE", "1", "X"}

OCCURRENCE_COLUMNS = (
    "RUN_TIMESTAMP",
    "JOURNAL_BUILD",
    "LEVEL",
    "COMPONENT_PATH",
    "COMPONENT_NAME",
    "PROTOTYPE_DB_PART_NO",
    "PROTOTYPE_DB_PART_REV",
    "PROTOTYPE_NAME",
    "PROTOTYPE_LOAD_STATE",
    "PROTOTYPE_NX_TYPE",
    "IS_SUPPRESSED",
    "IS_BLANKED",
    "COMPONENT_LAYER",
    "REFERENCE_SET",
    "ENTIRE_PART_REFSET_NAME",
    "IS_ENTIRE_PART_REFSET",
    "REPRESENTATION_MODE",
    "USED_ARRANGEMENT",
    "BOM_VISIBLE",
    "COUNTS_FOR_DISPLAY",
)

TARGET_COLUMNS = (
    "RUN_TIMESTAMP",
    "JOURNAL_BUILD",
    "MODE",
    "SCOPE_FILTER",
    "DB_PART_NO",
    "DB_PART_REV",
    "PART_NAME",
    "PART_KIND",
    "PART_NX_TYPE",
    "IS_WORK_PART",
    "LEVEL",
    "DEEPEST_LEVEL",
    "OCCURRENCE_COUNT",
    "GLOBAL_OCCURRENCE_COUNT",
    "SAMPLE_PATH",
    "LOAD_STATE",
    "IS_FULLY_LOADED",
    "HAS_MINIMALLY_LOADED_CHILDREN",
    "MINIMALLY_LOADED_CHILDREN",
    "BODY_COUNT",
    "SOLID_BODY_COUNT",
    "SHEET_BODY_COUNT",
    "CONVERGENT_BODY_COUNT",
    "FACE_COUNT",
    "EDGE_COUNT",
    "FACET_COUNT",
    "VERTEX_COUNT",
    "BODIES_ENUMERATED",
    "GEOMETRY_TRUNCATED",
    "BLANKED_BODY_COUNT",
    "DATUM_COUNT",
    "COORDINATE_SYSTEM_COUNT",
    "CURVE_COUNT",
    "LINE_COUNT",
    "POINT_COUNT",
    "ANNOTATION_COUNT",
    "FEATURE_COUNT",
    "EXPRESSION_COUNT",
    "TRUE_SHADING_COUNT",
    "TRUE_STUDIO_COUNT",
    "POINT_CLOUD_COUNT",
    "DECAL_COUNT",
    "CAMERA_COUNT",
    "DYNAMIC_SECTION_COUNT",
    "IMAGE_COUNT",
    "DRAWING_SHEET_COUNT",
    "SAVE_DISPLAY_FACETS",
    "PART_PREVIEW_MODE",
    "IS_DESIGN_REVIEW_PART",
    "IS_DISPLAYED",
    "DENSITY_MIN",
    "DENSITY_MAX",
    "DENSITY_AVERAGE",
    "DENSITY_ZERO_COUNT",
    "DENSITY_UNAVAILABLE_COUNT",
    "MATERIAL",
    "VISIBLE_LAYER_COUNT",
    "LAYER_WITH_OBJECT_COUNT",
    "HIDDEN_LAYER_WITH_OBJECT_COUNT",
    "LAYERS_TRUNCATED",
    "PROBE_STATUS",
    "PROBE_ERROR",
)

LAYER_COLUMNS = (
    "RUN_TIMESTAMP",
    "DB_PART_NO",
    "DB_PART_REV",
    "PART_NAME",
    "LAYER",
    "STATE",
    "IS_HIDDEN",
    "BODY_COUNT",
    "DATUM_COUNT",
    "COORDINATE_SYSTEM_COUNT",
    "CURVE_COUNT",
    "LINE_COUNT",
    "POINT_COUNT",
    "OTHER_COUNT",
    "TOTAL_COUNT",
)

VISIBLE_COLUMNS = (
    "RUN_TIMESTAMP",
    "VIEW_NAME",
    "VIEW_RENDERING_STYLE",
    "TOTAL_VISIBLE_OBJECTS",
    "TYPE_NAME",
    "TYPE_COUNT",
    "TYPE_PERCENT",
    "SCAN_SECONDS",
    "TRUNCATED",
)

PREF_COLUMNS = (
    "RUN_TIMESTAMP",
    "SCOPE",
    "OWNER",
    "KEY",
    "STATUS",
    "VALUE",
    "ERROR",
)

SUSPECT_COLUMNS = (
    "RANK",
    "RUN_TIMESTAMP",
    "SEVERITY",
    "CODE",
    "IDENTITY",
    "OCCURRENCE_COUNT",
    "MEASURED_VALUE",
    "THRESHOLD",
    "IMPACT_FACES",
    "MESSAGE",
    "EVIDENCE_IDS",
)

TIMING_COLUMNS = (
    "RUN_TIMESTAMP",
    "STEP",
    "OPERATION",
    "SECONDS",
    "CUMULATIVE_SECONDS",
    "NOTE",
)

DISCOVER_COLUMNS = (
    "RUN_TIMESTAMP",
    "SCOPE",
    "OWNER",
    "MEMBER",
    "STATUS",
    "VALUE",
)


# ---------------------------------------------------------------------------
# Mode resolution
# ---------------------------------------------------------------------------


def resolve_mode():
    value = normalize_text(os.environ.get("NX_J37_MODE")).upper() or MODE
    if value not in ("PROBE", "TIMED"):
        raise ValueError(
            "NX_J37_MODE must be PROBE or TIMED, got: {0}".format(value)
        )
    return value


def resolve_scope_filter():
    value = (
        normalize_text(os.environ.get("NX_J37_SCOPE")).upper()
        or SCOPE_FILTER
    )
    if value not in ("ALL", "BOM"):
        raise ValueError(
            "NX_J37_SCOPE must be ALL or BOM, got: {0}".format(value)
        )
    return value


def resolve_visible_scan():
    value = (
        normalize_text(os.environ.get("NX_J37_VISIBLE_SCAN")).upper()
    )
    if not value:
        return VISIBLE_SCAN
    return value in TRUE_VALUES


def resolve_rotation_count():
    value = normalize_text(os.environ.get("NX_J37_ROTATIONS"))
    if not value:
        return ROTATION_COUNT
    try:
        count = int(value)
    except Exception:
        raise ValueError(
            "NX_J37_ROTATIONS must be an integer, got: {0}".format(value)
        )
    if count < 0:
        raise ValueError("NX_J37_ROTATIONS must be >= 0")
    return count


def resolve_max_bodies():
    value = normalize_text(os.environ.get("NX_J37_MAX_BODIES"))
    if not value:
        return MAX_BODIES_PER_PART
    try:
        count = int(value)
    except Exception:
        raise ValueError(
            "NX_J37_MAX_BODIES must be an integer, got: {0}".format(value)
        )
    if count < 1:
        raise ValueError("NX_J37_MAX_BODIES must be >= 1")
    return count


def resolve_discover():
    value = normalize_text(os.environ.get("NX_J37_DISCOVER")).upper()
    if not value:
        return False
    return value in TRUE_VALUES


# ---------------------------------------------------------------------------
# Generic helpers
# ---------------------------------------------------------------------------


def normalize_text(value):
    return "" if value is None else str(value).strip()


def clean(value):
    return normalize_text(value)


def error_text(error):
    code = clean(getattr(error, "ErrorCode", ""))
    suffix = " [{0}]".format(code) if code else ""
    return "{0}{1}".format(clean(error) or type(error).__name__, suffix)


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
    for name in ("REPORTS", "LOGS"):
        path = os.path.join(run, name)
        os.makedirs(path, exist_ok=False)
        folders[name.lower()] = path
    return folders


def desktop_folder():
    profile = normalize_text(os.environ.get("USERPROFILE"))
    if profile:
        return os.path.join(profile, "Desktop")

    home = os.path.expanduser("~")
    if home and home != "~":
        return os.path.join(home, "Desktop")

    return os.getcwd()


def elapsed_seconds(started):
    return max(0.0, time.perf_counter() - started)


def seconds_text(value):
    try:
        return "{0:.3f}".format(float(value))
    except Exception:
        return ""


def safe_part_name(part, fallback="part"):
    for property_name in ("Name", "Leaf", "FullPath"):
        try:
            value = normalize_text(getattr(part, property_name))
            if value:
                return value
        except Exception:
            pass
    return fallback


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


def part_identity(part):
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


def part_kind(part):
    assembly = safe_property(part, "ComponentAssembly")
    if assembly is None:
        return "PART"
    root = safe_property(assembly, "RootComponent")
    return "ASSEMBLY" if root is not None else "PART"


def material_text(part):
    for title in ("Material", "MATERIAL", "NX_Material", "FZ_MATERIAL"):
        value = get_string_attribute(part, title)
        if value:
            return value
    return ""


def object_key(nx_object):
    tag = getattr(nx_object, "Tag", None)
    if tag is not None:
        return ("TAG", normalize_text(tag))
    return ("PY", id(nx_object))


# ---------------------------------------------------------------------------
# Evidence ledger and availability-aware probes
# ---------------------------------------------------------------------------


class EvidenceLedger:
    """Ordered fact ledger; every finding can cite stable fact IDs."""

    def __init__(self, build_id, timestamp, mode, scope_filter):
        self.build_id = build_id
        self.timestamp = timestamp
        self.mode = mode
        self.scope_filter = scope_filter
        self.facts = []
        self._index = {}

    def add(self, probe, owner, status, value="", error=""):
        fact_id = "F{0:05d}".format(len(self.facts) + 1)
        self.facts.append(
            {
                "id": fact_id,
                "probe": probe,
                "owner": owner,
                "status": status,
                "value": value,
                "error": error,
            }
        )
        self._index.setdefault(probe, []).append(fact_id)
        return fact_id

    def ids_for(self, probe):
        return list(self._index.get(probe, []))

    def to_dict(self):
        statuses = {}
        for fact in self.facts:
            statuses[fact["status"]] = statuses.get(fact["status"], 0) + 1
        return {
            "build": self.build_id,
            "timestamp": self.timestamp,
            "mode": self.mode,
            "scope_filter": self.scope_filter,
            "fact_count": len(self.facts),
            "status_counts": statuses,
            "facts": self.facts,
        }


def probe_value(owner, name):
    """Read one property defensively.

    Returns (status, value, error) where status is OK, UNAVAILABLE, or ERROR.
    A missing property is UNAVAILABLE (not False, not zero). A property whose
    getter raises is ERROR with the exact message.
    """
    if owner is None:
        return ("UNAVAILABLE", None, "owner is None")
    try:
        value = getattr(owner, name)
    except AttributeError as error:
        return ("UNAVAILABLE", None, "property not exposed: {0}".format(error))
    except Exception as error:
        return ("ERROR", None, error_text(error))

    if callable(value):
        return ("UNAVAILABLE", None, "member is a method, not a property")
    return ("OK", value, "")


def _list_result(value):
    if value is None:
        return ("UNAVAILABLE", [], "returned None")
    to_array = getattr(value, "ToArray", None)
    if callable(to_array):
        try:
            return ("OK", list(to_array()), "")
        except Exception as error:
            return ("ERROR", [], error_text(error))
    # Some NX managers (for example Annotations.AnnotationManager) are not
    # iterable and have no ToArray; try the documented enumeration methods
    # before declaring the probe unavailable.
    try:
        return ("OK", list(value), "")
    except TypeError:
        pass
    except Exception as error:
        return ("ERROR", [], error_text(error))
    for getter_name in ("GetObjects", "GetAnnotations", "AskAnnotations", "GetList"):
        getter = getattr(value, getter_name, None)
        if callable(getter):
            try:
                return ("OK", list(getter()), "")
            except Exception as error:
                return ("ERROR", [], error_text(error))
    return (
        "UNAVAILABLE",
        [],
        "{0} is not iterable and exposes no enumeration method".format(
            type(value).__name__
        ),
    )


def collection_items(owner, name):
    """Return (status, list, error) for a collection property or method."""
    if owner is None:
        return ("UNAVAILABLE", [], "owner is None")
    try:
        member = getattr(owner, name)
    except AttributeError as error:
        return ("UNAVAILABLE", [], "not exposed: {0}".format(error))
    except Exception as error:
        return ("ERROR", [], error_text(error))

    if callable(member):
        try:
            member = member()
        except Exception as error:
            return ("ERROR", [], error_text(error))
    return _list_result(member)


def call_items(owner, *names):
    """Return the first readable collection from any of names.

    Accepts both properties (part.Bodies) and zero-argument methods
    (body.GetFaces()), which is required because NX exposes geometry
    collections as methods on some classes and properties on others.
    """
    last = ("UNAVAILABLE", [], "no candidate name exposed")
    for name in names:
        status, items, error = collection_items(owner, name)
        if status == "OK":
            return (status, items, error)
        last = (status, items, error)
    return last


def call_int(owner, *names):
    """Read the first callable-or-property integer from any of names."""
    if owner is None:
        return ("UNAVAILABLE", None, "owner is None")
    last = ("UNAVAILABLE", None, "no candidate name exposed")
    for name in names:
        try:
            member = getattr(owner, name)
        except AttributeError as error:
            last = ("UNAVAILABLE", None, "not exposed: {0}".format(error))
            continue
        except Exception as error:
            last = ("ERROR", None, error_text(error))
            continue
        try:
            value = member() if callable(member) else member
            if value is None:
                last = ("UNAVAILABLE", None, "returned None")
                continue
            return ("OK", int(value), "")
        except Exception as error:
            last = ("ERROR", None, error_text(error))
    return last


def collection_count(owner, name):
    status, items, error = collection_items(owner, name)
    if status != "OK":
        return (status, None, error)
    return ("OK", len(items), "")


def nx_type_name(nx_object):
    if nx_object is None:
        return "None"
    try:
        get_type = getattr(nx_object, "GetType", None)
        if callable(get_type):
            type_object = get_type()
            name = getattr(type_object, "Name", None) or getattr(
                type_object, "__name__", None
            )
            if name:
                return str(name)
    except Exception:
        pass
    return type(nx_object).__name__


# ---------------------------------------------------------------------------
# Traversal
# ---------------------------------------------------------------------------


def component_name(component):
    for property_name in ("DisplayName", "Name"):
        value = clean(safe_property(component, property_name))
        if value:
            return value
    return "<unknown>"


def component_children(component):
    """Return (children, error) so an unreadable branch cannot look empty."""
    try:
        return (list(component.GetChildren()), "")
    except Exception as error:
        return ([], error_text(error))


def component_string_attribute(component, title):
    try:
        return component.GetStringAttribute(title)
    except Exception:
        return None


def is_bom_visible_component(component):
    name = clean(safe_property(component, "Name"))
    display_name = clean(safe_property(component, "DisplayName"))
    combined = " ".join((name, display_name)).upper()
    for keyword in IGNORE_KEYWORDS:
        if keyword in combined:
            return False
    for title in BOM_REFERENCE_ATTRIBUTES:
        raw = component_string_attribute(component, title)
        if raw is not None and clean(raw) in BOM_REFERENCE_FLAG_VALUES:
            return False
    custom = component_string_attribute(component, BOM_EXCLUSION_ATTRIBUTE)
    if custom is not None and clean(custom) == BOM_EXCLUSION_VALUE:
        return False
    return True


def collect_occurrences(work_part, scope_filter="ALL"):
    """Walk the whole assembly tree.

    Returns (occurrences, targets, diagnostics, global_occurrence_count).

    occurrences: one dict per occurrence (suppressed included, because the
    caller needs to know why an occurrence does not cost display time).
    targets: one dict per unique prototype, with the topmost level/path,
    deepest level, and the occurrence count inside the selected scope.
    """
    occurrences = []
    targets = {}
    diagnostics = []
    occurrence_count = 0

    root_number, _root_revision = part_identity(work_part)
    root_path = root_number or safe_part_name(work_part)

    def add_target(part, level, path, counts_for_display):
        key = object_key(part)
        record = targets.get(key)
        if record is None:
            targets[key] = {
                "key": key,
                "part": part,
                "level": level,
                "deepest_level": level,
                "sample_path": path,
                "occurrence_count": 1 if counts_for_display else 0,
                "is_work_part": False,
            }
            return
        record["deepest_level"] = max(record["deepest_level"], level)
        if counts_for_display:
            record["occurrence_count"] += 1
        if level < record["level"]:
            record["level"] = level
            record["sample_path"] = path

    add_target(work_part, 0, root_path, True)
    targets[object_key(work_part)]["is_work_part"] = True

    root_component = safe_property(
        safe_property(work_part, "ComponentAssembly"), "RootComponent"
    )
    if root_component is None:
        return occurrences, sorted_targets(targets), diagnostics, occurrence_count

    root_children, root_error = component_children(root_component)
    if root_error:
        diagnostics.append(
            {
                "code": "ROOT_CHILDREN_UNREADABLE",
                "message": "Assembly root children could not be read: "
                + root_error,
                "component_path": root_path,
                "level": 0,
            }
        )
        return occurrences, sorted_targets(targets), diagnostics, occurrence_count

    stack = [
        (component, 1, root_path) for component in reversed(root_children)
    ]
    while stack:
        component, level, parent_path = stack.pop()
        occurrence_count += 1
        component_path = "{0} / {1}".format(
            parent_path, component_name(component)
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

        prototype = safe_property(component, "Prototype")
        prototype_error = ""
        if prototype is None:
            probe_status, probe_error = component_prototype_status(component)
            prototype_error = probe_error
            if probe_status == "ERROR":
                diagnostics.append(
                    {
                        "code": "PROTOTYPE_UNAVAILABLE",
                        "message": "Component prototype could not be read: "
                        + probe_error,
                        "component_path": component_path,
                        "level": level,
                    }
                )

        suppressed = bool(safe_property(component, "IsSuppressed", False))
        blanked = bool(safe_property(component, "IsBlanked", False))
        reference_set = clean(safe_property(component, "ReferenceSet"))
        entire_part_refset = clean(
            safe_property(component, "EntirePartRefsetName")
        )
        representation = safe_property(component, "RepresentationMode")
        used_arrangement = safe_property(component, "UsedArrangement")
        bom_visible = is_bom_visible_component(component)

        counts_for_display = not suppressed
        if scope_filter == "BOM" and not bom_visible:
            counts_for_display = False

        if prototype is not None:
            add_target(
                prototype,
                level,
                component_path,
                counts_for_display,
            )

        occurrences.append(
            {
                "LEVEL": level,
                "COMPONENT_PATH": component_path,
                "COMPONENT_NAME": component_name(component),
                "PROTOTYPE_DB_PART_NO": (
                    part_identity(prototype)[0] if prototype is not None else ""
                ),
                "PROTOTYPE_DB_PART_REV": (
                    part_identity(prototype)[1] if prototype is not None else ""
                ),
                "PROTOTYPE_NAME": (
                    safe_part_name(prototype) if prototype is not None else ""
                ),
                "PROTOTYPE_LOAD_STATE": (
                    load_state_text(prototype) if prototype is not None else ""
                ),
                "PROTOTYPE_NX_TYPE": (
                    nx_type_name(prototype) if prototype is not None else ""
                ),
                "IS_SUPPRESSED": "YES" if suppressed else "NO",
                "IS_BLANKED": "YES" if blanked else "NO",
                "COMPONENT_LAYER": safe_property(component, "Layer", ""),
                "REFERENCE_SET": reference_set,
                "ENTIRE_PART_REFSET_NAME": entire_part_refset,
                "IS_ENTIRE_PART_REFSET": (
                    "YES"
                    if (
                        reference_set
                        and entire_part_refset
                        and reference_set.upper()
                        == entire_part_refset.upper()
                    )
                    else "NO"
                ),
                "REPRESENTATION_MODE": enum_text(representation),
                "USED_ARRANGEMENT": arrangement_text(used_arrangement),
                "BOM_VISIBLE": "YES" if bom_visible else "NO",
                "COUNTS_FOR_DISPLAY": (
                    "YES" if counts_for_display else "NO"
                ),
                "PROTOTYPE_READ_ERROR": prototype_error,
            }
        )

        children, children_error = component_children(component)
        if children_error:
            diagnostics.append(
                {
                    "code": "CHILDREN_UNREADABLE",
                    "message": (
                        "Component children could not be read for {0}: {1}"
                    ).format(component_name(component), children_error),
                    "component_path": component_path,
                    "level": level,
                }
            )
        stack.extend(
            (child, level + 1, component_path)
            for child in reversed(children)
        )

    return (
        occurrences,
        sorted_targets(targets),
        diagnostics,
        occurrence_count,
    )


def component_prototype_status(component):
    """Probe Component.Prototype without hasattr: an NXOpen property can
    raise a non-AttributeError exception, which hasattr would propagate."""
    try:
        component.Prototype
        return ("OK", "")
    except Exception as error:
        return ("ERROR", error_text(error))


def sorted_targets(targets):
    return sorted(
        targets.values(),
        key=lambda item: (
            item["level"],
            part_identity(item["part"])[0].upper(),
            safe_part_name(item["part"]).upper(),
        ),
    )


def safe_property(nx_object, name, fallback=None):
    try:
        value = getattr(nx_object, name)
        return value() if callable(value) else value
    except Exception:
        return fallback


def enum_text(value):
    if value is None:
        return ""
    name = getattr(value, "name", None)
    if name:
        return str(name)
    text = clean(value)
    if text:
        return text
    return type(value).__name__


def arrangement_text(arrangement):
    if arrangement is None:
        return ""
    for property_name in ("Name", "JournalIdentifier"):
        value = clean(safe_property(arrangement, property_name))
        if value:
            return value
    return enum_text(arrangement)


# ---------------------------------------------------------------------------
# Load state
# ---------------------------------------------------------------------------


# PartLoadState enum values (NX returns the NAME on some builds and the
# NUMERIC VALUE on others: NX 2506 returned '1' for FullyLoaded).
LOAD_STATE_BY_VALUE = {
    0: "NotLoaded",
    1: "FullyLoaded",
    2: "PartiallyLoaded",
    3: "MinimallyLoaded",
}
LOAD_STATE_BY_NAME = {
    "NOTLOADED": "NotLoaded",
    "FULLYLOADED": "FullyLoaded",
    "PARTIALLYLOADED": "PartiallyLoaded",
    "MINIMALLYLOADED": "MinimallyLoaded",
    "LOADED": "FullyLoaded",
}


def canonical_load_state(raw):
    """Return (canonical_name, raw_text) for a PartLoadState value.

    Accepts an enum (uses .name), a number, a numeric string, or an enum
    repr such as 'PartLoadState.FullyLoaded'. Unknown input is returned as-is
    rather than being forced into a load state.
    """
    if raw is None:
        return ("", "")
    name = getattr(raw, "name", None)
    text = clean(name if name is not None else raw)
    if not text:
        return ("", "")

    key = text.replace("_", "").replace(" ", "").upper()
    if key in LOAD_STATE_BY_NAME:
        return (LOAD_STATE_BY_NAME[key], text)

    tail = key.split(".")[-1]
    if tail in LOAD_STATE_BY_NAME:
        return (LOAD_STATE_BY_NAME[tail], text)

    try:
        value = int(float(text))
    except Exception:
        value = None
    if value is not None and value in LOAD_STATE_BY_VALUE:
        return (LOAD_STATE_BY_VALUE[value], text)

    return (text, text)


def load_state_text(part):
    status, raw_state = part_load_state(part)
    return raw_state or status


def part_load_state(part):
    """Return (status, raw_state).

    status is FULLY_LOADED, NOT_FULLY_LOADED, or UNKNOWN and is the only value
    callers may compare against. raw_state is the human-readable canonical
    name (FullyLoaded / PartiallyLoaded / ...) for the report.
    """
    fully_loaded = safe_property(part, "IsFullyLoaded")
    canonical, raw_text = canonical_load_state(
        safe_property(part, "PartLoadState")
    )

    if canonical in LOAD_STATE_BY_VALUE.values():
        status = (
            "FULLY_LOADED"
            if canonical == "FullyLoaded"
            else "NOT_FULLY_LOADED"
        )
        return (status, canonical)

    if fully_loaded is None:
        return ("UNKNOWN", canonical or raw_text or "UNKNOWN")
    try:
        return (
            "FULLY_LOADED" if bool(fully_loaded) else "NOT_FULLY_LOADED",
            canonical or raw_text or "UNKNOWN",
        )
    except Exception:
        return ("UNKNOWN", canonical or raw_text or "UNKNOWN")


def minimally_loaded_children(part):
    """Return (status, names, error) for a part's minimally loaded children."""
    has_any = safe_property(part, "HasAnyMinimallyLoadedChildren")
    if has_any is None:
        return ("UNAVAILABLE", [], "HasAnyMinimallyLoadedChildren not exposed")
    try:
        if not bool(has_any):
            return ("OK", [], "")
    except Exception as error:
        return ("ERROR", [], error_text(error))

    getter = getattr(part, "GetMinimallyLoadedParts", None)
    if not callable(getter):
        return ("OK", ["<has minimally loaded children>"], "")
    try:
        container = []
        getter(container)
        names = []
        for item in container:
            number, _revision = part_identity(item)
            names.append(number or safe_part_name(item))
        return ("OK", names, "")
    except Exception as error:
        return ("ERROR", [], error_text(error))


# ---------------------------------------------------------------------------
# Geometry census
# ---------------------------------------------------------------------------


def empty_geometry():
    return {
        "status": "UNAVAILABLE",
        "error": "",
        "body_count": "",
        "solid_body_count": "",
        "sheet_body_count": "",
        "convergent_body_count": "",
        "face_count": "",
        "edge_count": "",
        "facet_count": "",
        "vertex_count": "",
        "bodies_enumerated": "",
        "truncated": "NO",
        "blanked_body_count": "",
        "density_min": "",
        "density_max": "",
        "density_average": "",
        "density_zero_count": "",
        "density_unavailable_count": "",
        "max_face_body_density": "",
        "max_face_body_layer": "",
        "datum_count": "",
        "coordinate_system_count": "",
        "curve_count": "",
        "line_count": "",
        "point_count": "",
        "annotation_count": "",
        "feature_count": "",
        "expression_count": "",
        "true_shading_count": "",
        "true_studio_count": "",
        "point_cloud_count": "",
        "decal_count": "",
        "camera_count": "",
        "dynamic_section_count": "",
        "image_count": "",
        "drawing_sheet_count": "",
        "collection_status": {},
    }


PART_COLLECTION_SPECS = (
    ("Datums", "datum_count"),
    ("CoordinateSystems", "coordinate_system_count"),
    ("Curves", "curve_count"),
    ("Lines", "line_count"),
    ("Points", "point_count"),
    ("Annotations", "annotation_count"),
    ("Features", "feature_count"),
    ("Expressions", "expression_count"),
    ("SHEDObjs", "true_shading_count"),
    ("TrueStudioObjs", "true_studio_count"),
    ("PointClouds", "point_cloud_count"),
    ("Decals", "decal_count"),
    ("Cameras", "camera_count"),
    ("DynamicSections", "dynamic_section_count"),
    ("Images", "image_count"),
    ("DrawingSheets", "drawing_sheet_count"),
)


def identity_label(part, sample_path=""):
    """Identity for reports, with the component path when the part is unnamed.

    NX 2506 returned 46 prototypes with no readable Name/Leaf/FullPath and no
    DB_PART_NO; without the path they all collapse into one useless label.
    """
    number, revision = part_identity(part)
    name = safe_part_name(part, fallback="")
    if number or name:
        return "{0}/{1}".format(number or name, revision or "-")
    return "<unnamed prototype> @ {0}".format(sample_path or "<unknown path>")


def part_display_flags(part):
    """Cheap part-level display facts that need no geometry enumeration."""
    flags = {}
    for key, probe in (
        ("save_display_facets", "SaveDisplayFacets"),
        ("part_preview_mode", "PartPreviewMode"),
        ("is_design_review_part", "IsDesignReviewPart"),
        ("is_displayed", "Displayed"),
    ):
        status, value, _error = probe_value(part, probe)
        flags[key] = (
            enum_text(value)
            if status == "OK" and key != "save_display_facets"
            else (value if status == "OK" else "")
        )
    return flags


def discover_attributes(session, parts, max_names=400):
    """Read-only API discovery for probes this NX build does not expose.

    dumps the member names of the preference containers and of one
    representative component/prototype so the probe paths can be corrected
    from real evidence instead of guesswork. Nothing is modified.
    """
    rows = []

    def dump(scope, owner_label, owner, name_filter=None):
        if owner is None:
            rows.append(
                {
                    "SCOPE": scope,
                    "OWNER": owner_label,
                    "MEMBER": "<container>",
                    "STATUS": "UNAVAILABLE",
                    "VALUE": "owner is None",
                }
            )
            return
        names = []
        for name in dir(owner):
            if name.startswith("_"):
                continue
            if name_filter is not None and not any(
                token in name.lower() for token in name_filter
            ):
                continue
            names.append(name)
        for name in sorted(names)[: int(max_names)]:
            status, value, error = probe_value(owner, name)
            rows.append(
                {
                    "SCOPE": scope,
                    "OWNER": owner_label,
                    "MEMBER": name,
                    "STATUS": status,
                    "VALUE": (
                        enum_text(value)
                        if status == "OK"
                        else (error or "")[:120]
                    ),
                }
            )

    session_prefs = safe_property(session, "Preferences")
    dump("SESSION_PREFERENCES", "session.Preferences", session_prefs)
    for container_name in (
        "PerformanceVisualization",
        "Visualization",
        "VisualizationVisual",
        "Assembly",
        "SessionVisualizationPerformance",
        "ShadeVisualization",
    ):
        container = safe_property(session_prefs, container_name)
        if container is not None:
            dump(
                "SESSION_PREFERENCES",
                "session.Preferences.{0}".format(container_name),
                container,
            )

    for part in parts[:1]:
        dump("PART_PREFERENCES", "part.Preferences", safe_property(part, "Preferences"))
        for container_name in (
            "PerformanceVisualization",
            "ShadeVisualization",
            "VisualVisualization",
        ):
            container = safe_property(
                safe_property(part, "Preferences"), container_name
            )
            if container is not None:
                dump(
                    "PART_PREFERENCES",
                    "part.Preferences.{0}".format(container_name),
                    container,
                )
        dump(
            "PROTOTYPE",
            "part",
            part,
            name_filter=(
                "facet",
                "display",
                "lightweight",
                "preview",
                "annotat",
                "represent",
            ),
        )
        root = safe_property(
            safe_property(part, "ComponentAssembly"), "RootComponent"
        )
        if root is not None:
            dump(
                "COMPONENT",
                "root component",
                root,
                name_filter=(
                    "represent",
                    "lightweight",
                    "refset",
                    "reference",
                    "facet",
                    "load",
                ),
            )

    return rows


def census_part_collections(part, result):
    """Fill the non-body displayable-object counts on a census result."""
    statuses = {}
    for collection_name, key in PART_COLLECTION_SPECS:
        status, count, error = collection_count(part, collection_name)
        statuses[collection_name] = (status, count, error)
        result[key] = count if status == "OK" else ""
    result["collection_status"] = statuses
    return statuses


def census_geometry(part, max_bodies=MAX_BODIES_PER_PART):
    """Count bodies, faces, edges, facets and density for one part.

    Geometry is only enumerated for fully loaded parts: a partially loaded
    part would report misleading zeros, so it is reported UNAVAILABLE.
    """
    result = empty_geometry()

    # Non-body displayable objects are part data and are read for every part;
    # body geometry is only enumerated for fully loaded parts, because a
    # partially loaded part would report misleading zeros.
    census_part_collections(part, result)

    status, _raw_state = part_load_state(part)
    if status == "NOT_FULLY_LOADED":
        result["status"] = "UNAVAILABLE"
        result["error"] = (
            "part is not fully loaded; body geometry not enumerated "
            "(displayable-object counts are still reported)"
        )
        return result

    collection_status, bodies, collection_error = collection_items(
        part, "Bodies"
    )
    if collection_status != "OK":
        result["status"] = collection_status
        result["error"] = collection_error
        return result

    result["status"] = "OK"
    result["body_count"] = len(bodies)

    enumerated = bodies[: max(1, int(max_bodies))]
    result["bodies_enumerated"] = len(enumerated)
    if len(bodies) > len(enumerated):
        result["truncated"] = "YES"

    solid = 0
    sheet = 0
    convergent = 0
    faces = 0
    edges = 0
    facets = 0
    vertices = 0
    blanked = 0
    densities = []
    density_unavailable = 0
    max_face_body = None
    max_face_count = -1

    for body in enumerated:
        if body is None:
            continue
        if bool(safe_property(body, "IsSolidBody", False)):
            solid += 1
        if bool(safe_property(body, "IsSheetBody", False)):
            sheet += 1
        if bool(safe_property(body, "IsConvergentBody", False)):
            convergent += 1
        if bool(safe_property(body, "IsBlanked", False)):
            blanked += 1

        body_faces = -1
        face_status, face_items, _face_error = call_items(body, "GetFaces", "Faces")
        if face_status == "OK":
            body_faces = len(face_items)
            faces += body_faces

        edge_status, edge_items, _edge_error = call_items(body, "GetEdges", "Edges")
        if edge_status == "OK":
            edges += len(edge_items)

        facet_status, facet_value, _facet_error = call_int(
            body, "GetNumberOfFacets", "FacetCount"
        )
        if facet_status == "OK" and facet_value is not None:
            facets += facet_value

        vertex_status, vertex_value, _vertex_error = call_int(
            body, "GetNumberOfVertices", "VertexCount"
        )
        if vertex_status == "OK" and vertex_value is not None:
            vertices += vertex_value

        density_status, density_value, _density_error = probe_value(
            body, "Density"
        )
        if density_status == "OK":
            try:
                numeric = float(density_value)
                densities.append(numeric)
            except Exception:
                density_unavailable += 1
        else:
            density_unavailable += 1

        if body_faces > max_face_count:
            max_face_count = body_faces
            max_face_body = body

    result.update(
        {
            "solid_body_count": solid,
            "sheet_body_count": sheet,
            "convergent_body_count": convergent,
            "face_count": faces,
            "edge_count": edges,
            "facet_count": facets,
            "vertex_count": vertices,
            "blanked_body_count": blanked,
            "density_unavailable_count": density_unavailable,
        }
    )

    if densities:
        result["density_min"] = min(densities)
        result["density_max"] = max(densities)
        result["density_average"] = sum(densities) / len(densities)
        result["density_zero_count"] = sum(
            1 for value in densities if value == 0.0
        )
    else:
        result["density_zero_count"] = ""

    if max_face_body is not None:
        if max_face_count >= 0:
            result["max_face_body_face_count"] = max_face_count
        result["max_face_body_layer"] = safe_property(
            max_face_body, "Layer", ""
        )
        density_status, density_value, _error = probe_value(
            max_face_body, "Density"
        )
        if density_status == "OK":
            result["max_face_body_density"] = density_value

    return result


# ---------------------------------------------------------------------------
# Layer census
# ---------------------------------------------------------------------------


def layer_state_text(layer_manager, layer_number):
    getter = getattr(layer_manager, "GetState", None)
    if not callable(getter):
        return ("UNAVAILABLE", "")
    try:
        return ("OK", enum_text(getter(layer_number)))
    except Exception as error:
        return ("ERROR", error_text(error))


def object_layer(nx_object):
    value = safe_property(nx_object, "Layer")
    if value is None:
        return None
    try:
        return int(value)
    except Exception:
        return None


def census_layers(part, max_layers=MAX_LAYERS):
    """Return (rows, status, error).

    Known displayable objects are bucketed by their Layer property; every
    layer's state is read so hidden layers holding geometry are visible in the
    report. Objects NX owns that are not in those collections are not
    enumerated (reported as OTHER_COUNT=0, never as a proven zero).
    """
    layer_manager = safe_property(part, "Layers")
    if layer_manager is None:
        return ([], "UNAVAILABLE", "part.Layers not exposed", 0)

    buckets = {}

    def bucket(nx_object, bucket_name):
        layer = object_layer(nx_object)
        if layer is None:
            return
        entry = buckets.setdefault(
            layer,
            {
                "BODY_COUNT": 0,
                "DATUM_COUNT": 0,
                "COORDINATE_SYSTEM_COUNT": 0,
                "CURVE_COUNT": 0,
                "LINE_COUNT": 0,
                "POINT_COUNT": 0,
                "OTHER_COUNT": 0,
            },
        )
        entry[bucket_name] = entry.get(bucket_name, 0) + 1

    collections = (
        ("Bodies", "BODY_COUNT"),
        ("Datums", "DATUM_COUNT"),
        ("CoordinateSystems", "COORDINATE_SYSTEM_COUNT"),
        ("Curves", "CURVE_COUNT"),
        ("Lines", "LINE_COUNT"),
        ("Points", "POINT_COUNT"),
    )
    errors = []
    for collection_name, bucket_name in collections:
        status, items, error = collection_items(part, collection_name)
        if status != "OK":
            errors.append("{0}: {1}".format(collection_name, error))
            continue
        for item in items:
            bucket(item, bucket_name)
    rows = []
    visible_layers = 0
    for layer_number in range(1, int(max_layers) + 1):
        state_status, state = layer_state_text(layer_manager, layer_number)
        entry = buckets.get(
            layer_number,
            {
                "BODY_COUNT": 0,
                "DATUM_COUNT": 0,
                "COORDINATE_SYSTEM_COUNT": 0,
                "CURVE_COUNT": 0,
                "LINE_COUNT": 0,
                "POINT_COUNT": 0,
                "OTHER_COUNT": 0,
            },
        )
        total = sum(
            int(entry[key] or 0)
            for key in (
                "BODY_COUNT",
                "DATUM_COUNT",
                "COORDINATE_SYSTEM_COUNT",
                "CURVE_COUNT",
                "LINE_COUNT",
                "POINT_COUNT",
                "OTHER_COUNT",
            )
        )
        is_hidden = state.upper() == "HIDDEN" if state_status == "OK" else ""
        if state_status == "OK" and not is_hidden:
            visible_layers += 1
        if total == 0 and state_status != "OK":
            continue
        rows.append(
            {
                "LAYER": layer_number,
                "STATE": state if state_status == "OK" else state_status,
                "IS_HIDDEN": "YES" if is_hidden is True else ("NO" if is_hidden is False else ""),
                "TOTAL_COUNT": total,
                **entry,
            }
        )

    status = "OK" if not errors else "PARTIAL"
    return (rows, status, " | ".join(errors), visible_layers)


# ---------------------------------------------------------------------------
# Visible-object census
# ---------------------------------------------------------------------------


def active_view(session):
    """Return (status, view, error) for the active model view.

    ViewCollection exposes GetActiveViews() (the graphics-window views) and
    has no single "current view" property, so the active list is used first
    and the collection order is the fallback.
    """
    display_part = safe_property(session.Parts, "Display")
    if display_part is None:
        display_part = safe_property(session.Parts, "Work")
    if display_part is None:
        return ("UNAVAILABLE", None, "no display or work part")

    views = safe_property(display_part, "Views")
    if views is None:
        return ("UNAVAILABLE", None, "display part has no Views collection")

    getter = getattr(views, "GetActiveViews", None)
    if callable(getter):
        try:
            active = list(getter())
        except Exception as error:
            active = []
            error_message = error_text(error)
        else:
            error_message = ""
        if active:
            return ("OK", active[0], "")
        if error_message:
            return ("ERROR", None, error_message)

    status, items, error = collection_items(views, "ToArray")
    if status == "OK" and items:
        return ("OK", items[0], "first view in collection")
    if status != "OK":
        return (status, None, error)
    return ("UNAVAILABLE", None, "no view available")


def census_visible_objects(view, max_objects=MAX_VISIBLE_OBJECTS):
    """Count visible objects in a view by NX type.

    Returns (status, rows, total, seconds, truncated, error).
    """
    ask = getattr(view, "AskVisibleObjects", None)
    if not callable(ask):
        return ("UNAVAILABLE", [], "", "", "NO", "AskVisibleObjects not exposed")

    started = time.perf_counter()
    try:
        objects = list(ask())
    except Exception as error:
        return ("ERROR", [], "", "", "NO", error_text(error))
    scan_seconds = elapsed_seconds(started)

    truncated = "NO"
    if len(objects) > max_objects:
        objects = objects[: int(max_objects)]
        truncated = "YES"

    counts = {}
    for nx_object in objects:
        name = nx_type_name(nx_object)
        counts[name] = counts.get(name, 0) + 1

    total = len(objects)
    rows = []
    for name in sorted(counts, key=lambda key: (-counts[key], key)):
        rows.append(
            {
                "TYPE_NAME": name,
                "TYPE_COUNT": counts[name],
                "TYPE_PERCENT": (
                    round((counts[name] / total) * 100.0, 3) if total else ""
                ),
            }
        )
    return ("OK", rows, total, scan_seconds, truncated, "")


# ---------------------------------------------------------------------------
# Preference census
# ---------------------------------------------------------------------------


SESSION_PREF_KEYS = (
    "LoadComponentOnFacetedViewUpdate",
    "LoadComponentOnFacetedViewSelection",
    "SmartlightweightViewsLoadComponentOnDemand",
    "WorkPartDisplayAsEntirePart",
    "DisplayUpdateReport",
    "RenderSolidsUsingStoredFacets",
    "ShowFacetEdges",
    "ShadingTolerance",
    "CustomFaceTolerance",
    "CustomEdgeTolerance",
    "CustomAngleTolerance",
)

PART_PREF_KEYS = (
    "SaveAdvancedDisplayFacets",
    "RenderSolidsUsingStoredFacets",
    "ShowFacetEdges",
    "ShadingTolerance",
    "CustomFaceTolerance",
    "CustomEdgeTolerance",
    "CustomAngleTolerance",
)


def census_preferences(session, parts):
    """Return preference rows for the display-performance probes.

    Every probe is availability-aware: an unsupported property on this NX
    build is reported UNAVAILABLE with the reason, never as False.
    """
    rows = []

    def add_row(scope, owner_label, owner, key):
        status, value, error = probe_value(owner, key)
        rows.append(
            {
                "SCOPE": scope,
                "OWNER": owner_label,
                "KEY": key,
                "STATUS": status,
                "VALUE": enum_text(value) if status == "OK" else "",
                "ERROR": error,
            }
        )
        return status == "OK"

    def add_container_row(scope, owner_label, error):
        rows.append(
            {
                "SCOPE": scope,
                "OWNER": owner_label,
                "KEY": "<container>",
                "STATUS": "UNAVAILABLE",
                "VALUE": "",
                "ERROR": error,
            }
        )

    session_prefs = safe_property(session, "Preferences")
    session_owners = (
        ("session", session_prefs),
        (
            "session.PerformanceVisualization",
            safe_property(session_prefs, "PerformanceVisualization"),
        ),
        (
            "session.ShadeVisualization",
            safe_property(session_prefs, "ShadeVisualization"),
        ),
    )
    for owner_label, owner in session_owners:
        if owner is None:
            add_container_row(
                "SESSION",
                owner_label,
                "container not exposed on this NX build; keys not probed",
            )
            continue
        for key in SESSION_PREF_KEYS:
            add_row("SESSION", owner_label, owner, key)

    for part in parts:
        part_name = safe_part_name(part)
        number, revision = part_identity(part)
        identity = "{0}/{1}".format(number or part_name, revision or "-")
        part_prefs = safe_property(part, "Preferences")
        if part_prefs is None:
            add_container_row(
                "PART", identity, "part.Preferences not exposed"
            )
            continue
        part_owners = (
            (
                "{0} PerformanceVisualization".format(identity),
                safe_property(part_prefs, "PerformanceVisualization"),
            ),
            (
                "{0} ShadeVisualization".format(identity),
                safe_property(part_prefs, "ShadeVisualization"),
            ),
        )
        for owner_label, owner in part_owners:
            if owner is None:
                add_container_row(
                    "PART",
                    owner_label,
                    "container not exposed on this NX build; keys not probed",
                )
                continue
            for key in PART_PREF_KEYS:
                add_row("PART", owner_label, owner, key)

    return rows


def pref_lookup(pref_rows, scope, key_fragment):
    for row in pref_rows:
        if row["SCOPE"] != scope:
            continue
        if key_fragment.lower() not in row["KEY"].lower():
            continue
        if row["STATUS"] != "OK":
            continue
        return row
    return None


def pref_is_true(row):
    if row is None:
        return None
    return clean(row["VALUE"]).upper() in TRUE_VALUES


# ---------------------------------------------------------------------------
# Suspect building
# ---------------------------------------------------------------------------


def build_suspects(
    ledger,
    targets,
    occurrences,
    geometry_by_key,
    layer_info_by_key,
    minimal_by_key,
    visible_total,
    pref_rows,
    diagnostics,
):
    """Build the ranked suspect list. Only measured values are quoted."""
    suspects = []

    def add(
        code,
        severity,
        identity,
        occurrence_count,
        measured,
        threshold,
        impact_faces,
        message,
        probes,
    ):
        evidence_ids = []
        for probe in probes:
            evidence_ids.extend(ledger.ids_for(probe))
        suspects.append(
            {
                "SEVERITY": severity,
                "CODE": code,
                "IDENTITY": identity,
                "OCCURRENCE_COUNT": occurrence_count,
                "MEASURED_VALUE": measured,
                "THRESHOLD": threshold,
                "IMPACT_FACES": impact_faces,
                "MESSAGE": message,
                "EVIDENCE_IDS": " ".join(evidence_ids),
            }
        )

    # --- load truth -----------------------------------------------------
    not_fully_loaded = [
        target
        for target in targets
        if part_load_state(target["part"])[0] != "FULLY_LOADED"
    ]
    for target in not_fully_loaded:
        number, revision = part_identity(target["part"])
        add(
            "NOT_FULLY_LOADED",
            "HIGH",
            "{0}/{1}".format(number or safe_part_name(target["part"]), revision or "-"),
            target["occurrence_count"],
            load_state_text(target["part"]),
            "FULLY_LOADED",
            "",
            (
                "Part is not fully loaded. NX may load geometry during a view "
                "change, which is a direct cause of very slow view updates."
            ),
            ["PartLoadState"],
        )

    for target in targets:
        status, names, error = minimal_by_key.get(
            target["key"], ("UNAVAILABLE", [], "not probed")
        )
        if status != "OK" or not names:
            continue
        number, revision = part_identity(target["part"])
        add(
            "MINIMALLY_LOADED_CHILDREN",
            "HIGH",
            "{0}/{1}".format(number or safe_part_name(target["part"]), revision or "-"),
            target["occurrence_count"],
            len(names),
            "0",
            "",
            (
                "HasAnyMinimallyLoadedChildren is True ({0}). A view update can "
                "pull these in. First entries: {1}"
            ).format(len(names), "; ".join(names[:5])),
            ["HasAnyMinimallyLoadedChildren", "GetMinimallyLoadedParts"],
        )

    # --- occurrence display configuration -------------------------------
    entire_part_by_prototype = {}
    exact_by_prototype = {}
    non_lightweight_by_prototype = {}
    for occurrence in occurrences:
        if occurrence["COUNTS_FOR_DISPLAY"] != "YES":
            continue
        key = "{0}/{1}".format(
            occurrence["PROTOTYPE_DB_PART_NO"] or occurrence["PROTOTYPE_NAME"],
            occurrence["PROTOTYPE_DB_PART_REV"] or "-",
        )
        if occurrence["IS_ENTIRE_PART_REFSET"] == "YES":
            record = entire_part_by_prototype.setdefault(
                key, {"count": 0, "refset": occurrence["REFERENCE_SET"]}
            )
            record["count"] += 1
        mode = occurrence["REPRESENTATION_MODE"].upper()
        if mode == "EXACT":
            exact_by_prototype[key] = exact_by_prototype.get(key, 0) + 1
        if mode not in ("LIGHTWEIGHT", ""):
            non_lightweight_by_prototype[key] = (
                non_lightweight_by_prototype.get(key, 0) + 1
            )

    faces_by_identity = {}
    for target in targets:
        number, revision = part_identity(target["part"])
        identity = identity_label(target["part"], target["sample_path"])
        geometry = geometry_by_key.get(target["key"], {})
        try:
            faces = int(geometry.get("face_count") or 0)
        except Exception:
            faces = 0
        faces_by_identity[identity] = faces

    for identity, record in entire_part_by_prototype.items():
        add(
            "ENTIRE_PART_REFSET",
            "HIGH",
            identity,
            record["count"],
            record["count"],
            "0",
            faces_by_identity.get(identity, ""),
            (
                "Reference set '{0}' equals the component's EntirePartRefsetName, "
                "so the occurrence displays every object in the part instead of a "
                "model reference set."
            ).format(record["refset"]),
            ["ReferenceSet", "EntirePartRefsetName"],
        )

    for identity, count in exact_by_prototype.items():
        add(
            "EXACT_REPRESENTATION",
            "HIGH" if count >= 10 else "MEDIUM",
            identity,
            count,
            count,
            "0",
            faces_by_identity.get(identity, ""),
            (
                "RepresentationMode is Exact on {0} occurrence(s): full geometry "
                "is used instead of a lightweight representation."
            ).format(count),
            ["RepresentationMode"],
        )

    for identity, count in non_lightweight_by_prototype.items():
        if identity in exact_by_prototype:
            continue
        add(
            "NON_LIGHTWEIGHT_REPRESENTATION",
            "MEDIUM",
            identity,
            count,
            count,
            "0",
            faces_by_identity.get(identity, ""),
            (
                "RepresentationMode is not Lightweight on {0} occurrence(s)."
            ).format(count),
            ["RepresentationMode"],
        )

    # --- geometry complexity --------------------------------------------
    ranked = sorted(
        targets,
        key=lambda target: -(
            faces_by_identity.get(
                "{0}/{1}".format(
                    part_identity(target["part"])[0]
                    or safe_part_name(target["part"]),
                    part_identity(target["part"])[1] or "-",
                ),
                0,
            )
        ),
    )
    for target in ranked[:TOP_SUSPECTS]:
        number, revision = part_identity(target["part"])
        identity = identity_label(target["part"], target["sample_path"])
        geometry = geometry_by_key.get(target["key"], {})
        faces = faces_by_identity.get(identity, 0)
        if not faces or faces < THRESHOLD_FACE_COUNT:
            continue
        add(
            "HIGH_FACE_GEOMETRY",
            "HIGH" if faces >= THRESHOLD_FACE_COUNT * 4 else "MEDIUM",
            identity,
            target["occurrence_count"],
            faces,
            THRESHOLD_FACE_COUNT,
            faces * max(1, target["occurrence_count"]),
            (
                "Part carries {0} faces over {1} bodies and appears {2} time(s) "
                "in the assembly."
            ).format(
                faces,
                geometry.get("body_count", ""),
                target["occurrence_count"],
            ),
            ["Bodies", "GetFaces"],
        )

    for target in targets:
        number, revision = part_identity(target["part"])
        identity = identity_label(target["part"], target["sample_path"])
        geometry = geometry_by_key.get(target["key"], {})
        try:
            convergent = int(geometry.get("convergent_body_count") or 0)
            facets = int(geometry.get("facet_count") or 0)
        except Exception:
            convergent, facets = 0, 0
        if convergent or facets:
            add(
                "CONVERGENT_OR_FACETED_GEOMETRY",
                "MEDIUM" if not facets else "HIGH",
                identity,
                target["occurrence_count"],
                "bodies={0}; facets={1}".format(convergent, facets),
                "bodies=0; facets=0",
                "",
                (
                    "Faceted/convergent geometry must be drawn from its facet "
                    "representation; large facet counts are re-evaluated on view "
                    "changes."
                ),
                ["IsConvergentBody", "GetNumberOfFacets"],
            )

    for target in targets:
        number, revision = part_identity(target["part"])
        identity = identity_label(target["part"], target["sample_path"])
        geometry = geometry_by_key.get(target["key"], {})
        zero = geometry.get("density_zero_count", "")
        if zero == "":
            continue
        try:
            zero_count = int(zero)
        except Exception:
            continue
        if zero_count <= 0:
            continue
        add(
            "DENSITY_ZERO_OR_MISSING",
            "MEDIUM",
            identity,
            target["occurrence_count"],
            "bodies_with_zero_density={0}; min={1}; avg={2}".format(
                zero_count,
                geometry.get("density_min", ""),
                geometry.get("density_average", ""),
            ),
            "0",
            "",
            (
                "Body density is exactly zero on {0} body(ies). Mass properties "
                "and any density-driven update cannot be trusted; confirm the "
                "material/density assignment."
            ).format(zero_count),
            ["Density"],
        )

    for target in targets:
        number, revision = part_identity(target["part"])
        identity = identity_label(target["part"], target["sample_path"])
        geometry = geometry_by_key.get(target["key"], {})
        try:
            minimum = float(geometry.get("density_min"))
            maximum = float(geometry.get("density_max"))
        except Exception:
            continue
        if minimum <= 0.0 or maximum <= 0.0:
            continue
        ratio = maximum / minimum
        if ratio < THRESHOLD_DENSITY_RATIO:
            continue
        add(
            "DENSITY_OUTLIER_RANGE",
            "INFO",
            identity,
            target["occurrence_count"],
            "min={0}; max={1}; ratio={2:.1f}".format(minimum, maximum, ratio),
            THRESHOLD_DENSITY_RATIO,
            "",
            (
                "Body densities inside this one part span a factor of {0:.1f}. "
                "That is legal for multi-material parts, but it is worth "
                "confirming the assignment is intentional."
            ).format(ratio),
            ["Density"],
        )

    # --- clutter, blanking, layers --------------------------------------
    for target in targets:
        number, revision = part_identity(target["part"])
        identity = identity_label(target["part"], target["sample_path"])
        geometry = geometry_by_key.get(target["key"], {})

        clutter = 0
        for key in ("datum_count", "curve_count", "line_count", "point_count"):
            value = geometry.get(key, "")
            try:
                clutter += int(value)
            except Exception:
                continue
        if clutter >= THRESHOLD_DATUM_CURVE_COUNT:
            add(
                "NON_SOLID_CLUTTER",
                "MEDIUM",
                identity,
                target["occurrence_count"],
                clutter,
                THRESHOLD_DATUM_CURVE_COUNT,
                "",
                (
                    "Datums + curves + lines + points total {0} displayable "
                    "objects in this part."
                ).format(clutter),
                ["Datums", "Curves", "Lines", "Points"],
            )

        try:
            blanked = int(geometry.get("blanked_body_count") or 0)
        except Exception:
            blanked = 0
        if blanked >= THRESHOLD_BLANKED_BODY_COUNT:
            add(
                "BLANKED_OBJECT_BULK",
                "INFO",
                identity,
                target["occurrence_count"],
                blanked,
                THRESHOLD_BLANKED_BODY_COUNT,
                "",
                (
                    "Individually blanked bodies: {0}. Blank state is stored per "
                    "object and is evaluated per view."
                ).format(blanked),
                ["IsBlanked"],
            )

        heavy = []
        heavy_labels = []
        for key, label in (
            ("true_shading_count", "True Shading"),
            ("true_studio_count", "True Studio"),
            ("point_cloud_count", "point clouds"),
            ("decal_count", "decals"),
            ("dynamic_section_count", "dynamic sections"),
        ):
            value = geometry.get(key, "")
            try:
                count = int(value)
            except Exception:
                continue
            if count:
                heavy.append("{0}={1}".format(label, count))
                heavy_labels.append(label)
        if heavy:
            expensive = any(
                label in ("True Shading", "point clouds")
                for label in heavy_labels
            )
            add(
                "HEAVY_DISPLAY_OBJECTS",
                "HIGH" if expensive else "MEDIUM",
                identity,
                target["occurrence_count"],
                " | ".join(heavy),
                "0",
                "",
                (
                    "Heavy display objects are present. True Shading / studio "
                    "environments and point clouds are re-evaluated on every view "
                    "change."
                ),
                ["SHEDObjs", "TrueStudioObjs", "PointClouds", "Decals"],
            )

    layer_rows = []
    for target in targets:
        info = layer_info_by_key.get(target["key"])
        if not info:
            continue
        number, revision = part_identity(target["part"])
        identity = identity_label(target["part"], target["sample_path"])
        visible_layers = info.get("visible_layers", 0)
        populated = info.get("populated_layers", 0)
        hidden_populated = info.get("hidden_populated_layers", 0)
        layer_rows.append(
            (identity, visible_layers, populated, hidden_populated)
        )

    for identity, visible_layers, populated, hidden_populated in layer_rows:
        if visible_layers >= THRESHOLD_VISIBLE_LAYERS:
            add(
                "MANY_VISIBLE_LAYERS",
                "MEDIUM",
                identity,
                "",
                "visible_layers={0}; populated_layers={1}".format(
                    visible_layers, populated
                ),
                THRESHOLD_VISIBLE_LAYERS,
                "",
                (
                    "Geometry is spread over {0} visible layers. Consolidating "
                    "display layers reduces per-view visibility work."
                ).format(visible_layers),
                ["Layers", "GetState"],
            )
        if hidden_populated:
            add(
                "HIDDEN_LAYERS_WITH_OBJECTS",
                "INFO",
                identity,
                "",
                "hidden_layers_with_objects={0}".format(hidden_populated),
                "0",
                "",
                (
                    "{0} hidden layer(s) still hold known objects. Hidden layers "
                    "do not draw, but they explain where geometry lives."
                ).format(hidden_populated),
                ["Layers", "GetState"],
            )

    # --- preferences ----------------------------------------------------
    for key, severity, message in (
        (
            "LoadComponentOnFacetedViewUpdate",
            "HIGH",
            "NX is configured to load components when a faceted view updates. "
            "Combined with a partial load this makes every view change load "
            "geometry.",
        ),
        (
            "LoadComponentOnFacetedViewSelection",
            "MEDIUM",
            "NX is configured to load components when a faceted view selection "
            "happens.",
        ),
        (
            "SmartlightweightViewsLoadComponentOnDemand",
            "MEDIUM",
            "Smart lightweight views are configured to load components on demand.",
        ),
        (
            "WorkPartDisplayAsEntirePart",
            "MEDIUM",
            "The work part is displayed as the entire part rather than the work "
            "part reference, which can display far more geometry.",
        ),
        (
            "DisplayUpdateReport",
            "INFO",
            "Display Update Report is enabled; NX will report update timings "
            "(useful evidence, but it also blocks the update).",
        ),
    ):
        row = pref_lookup(pref_rows, "SESSION", key)
        if pref_is_true(row) is True:
            add(
                "PREF_{0}_TRUE".format(key.upper()),
                severity,
                "SESSION",
                "",
                row["VALUE"],
                "FALSE",
                "",
                message,
                ["PREF_{0}".format(key)],
            )

    for key, severity, message in (
        (
            "RenderSolidsUsingStoredFacets",
            "MEDIUM",
            "Rendering solids from stored facets is OFF or unavailable, so "
            "facets may be regenerated instead of reused.",
        ),
        (
            "ShowFacetEdges",
            "INFO",
            "Facet edges are shown, which adds drawing work in shaded views.",
        ),
    ):
        for scope in ("SESSION", "PART"):
            for row in pref_rows:
                if row["SCOPE"] != scope or row["KEY"] != key:
                    continue
                if row["STATUS"] != "OK":
                    continue
                value = clean(row["VALUE"]).upper()
                if key == "RenderSolidsUsingStoredFacets" and value in TRUE_VALUES:
                    continue
                if key == "ShowFacetEdges" and value not in TRUE_VALUES:
                    continue
                add(
                    "PREF_{0}_{1}".format(key.upper(), scope),
                    severity,
                    row["OWNER"],
                    "",
                    row["VALUE"],
                    "FALSE" if key != "ShowFacetEdges" else "TRUE",
                    "",
                    message,
                    ["PREF_{0}".format(key)],
                )
                break

    for row in pref_rows:
        if row["KEY"] != "SaveAdvancedDisplayFacets" or row["STATUS"] != "OK":
            continue
        if clean(row["VALUE"]).upper() != "TRUE":
            continue
        add(
            "PREF_SAVE_ADVANCED_DISPLAY_FACETS_TRUE",
            "INFO",
            row["OWNER"],
            "",
            row["VALUE"],
            "FALSE",
            "",
            "Advanced display facets are saved with this part, increasing part "
            "size and memory use.",
            ["PREF_SaveAdvancedDisplayFacets"],
        )

    for row in pref_rows:
        if row["KEY"] not in (
            "ShadingTolerance",
            "CustomFaceTolerance",
            "CustomEdgeTolerance",
            "CustomAngleTolerance",
        ):
            continue
        if row["STATUS"] != "OK":
            continue
        value = clean(row["VALUE"])
        if not value:
            continue
        if "CUSTOM" not in value.upper() and "USER" not in value.upper():
            continue
        add(
            "PREF_SHADING_TOLERANCE_{0}".format(value.upper()),
            "INFO",
            row["OWNER"],
            "",
            value,
            "",
            "",
            (
                "Shading tolerance uses a custom setting. Very fine face/edge "
                "tolerances create far more facets per view regeneration."
            ),
            ["PREF_{0}".format(row["KEY"])],
        )

    # --- visible objects -------------------------------------------------
    if visible_total not in ("", None):
        try:
            total = int(visible_total)
        except Exception:
            total = 0
        if total >= THRESHOLD_VISIBLE_OBJECTS:
            add(
                "VISIBLE_OBJECT_COUNT_HIGH",
                "HIGH",
                "ACTIVE_VIEW",
                "",
                total,
                THRESHOLD_VISIBLE_OBJECTS,
                total,
                (
                    "View.AskVisibleObjects returns {0} visible objects. This is "
                    "the number NX must consider for every view regeneration."
                ).format(total),
                ["AskVisibleObjects"],
            )

    # --- traversal findings ----------------------------------------------
    for diagnostic in diagnostics:
        add(
            "TRAVERSAL_{0}".format(diagnostic.get("code", "DIAGNOSTIC")),
            "MEDIUM",
            diagnostic.get("component_path", ""),
            "",
            diagnostic.get("level", ""),
            "",
            "",
            diagnostic.get("message", ""),
            ["Component.GetChildren"],
        )

    severity_order = {"HIGH": 0, "MEDIUM": 1, "INFO": 2}

    def sort_key(item):
        try:
            impact = int(item["IMPACT_FACES"] or 0)
        except Exception:
            impact = 0
        return (
            severity_order.get(item["SEVERITY"], 3),
            -impact,
            item["CODE"],
            item["IDENTITY"],
        )

    suspects.sort(key=sort_key)
    return suspects


# ---------------------------------------------------------------------------
# Timing
# ---------------------------------------------------------------------------


def save_view_state(view):
    return {
        "matrix": safe_property(view, "Matrix"),
        "origin": safe_property(view, "Origin"),
        "scale": safe_property(view, "Scale"),
    }


def restore_view_state(view, state):
    setter = getattr(view, "SetRotationTranslationScale", None)
    if (
        callable(setter)
        and state.get("matrix") is not None
        and state.get("origin") is not None
        and state.get("scale") is not None
    ):
        setter(state["matrix"], state["origin"], state["scale"])
        return "SetRotationTranslationScale"
    return ""


def run_timing(session, view, rotation_count):
    """Measure view-change cost. Read-only: the view is restored afterwards."""
    rows = []
    cumulative = 0.0
    state = save_view_state(view)
    restore_method = ""
    notes = []

    def timed(step, operation, function, note=""):
        nonlocal cumulative
        started = time.perf_counter()
        error = ""
        try:
            function()
        except Exception as exception:
            error = error_text(exception)
        seconds = elapsed_seconds(started)
        cumulative += seconds
        rows.append(
            {
                "STEP": step,
                "OPERATION": operation,
                "SECONDS": seconds,
                "CUMULATIVE_SECONDS": cumulative,
                "NOTE": note if not error else "ERROR: " + error,
            }
        )
        return seconds

    regenerate = getattr(view, "Regenerate", None)
    if callable(regenerate):
        timed("1", "View.Regenerate", regenerate, "explicit regeneration")
    else:
        rows.append(
            {
                "STEP": "1",
                "OPERATION": "View.Regenerate",
                "SECONDS": "",
                "CUMULATIVE_SECONDS": "",
                "NOTE": "UNAVAILABLE: Regenerate not exposed",
            }
        )

    update_display = getattr(view, "UpdateDisplay", None)
    if callable(update_display):
        timed("2", "View.UpdateDisplay", update_display, "display update")
    else:
        rows.append(
            {
                "STEP": "2",
                "OPERATION": "View.UpdateDisplay",
                "SECONDS": "",
                "CUMULATIVE_SECONDS": "",
                "NOTE": "UNAVAILABLE: UpdateDisplay not exposed",
            }
        )

    rotate = getattr(view, "Rotate", None)
    origin = state.get("origin")
    if callable(rotate) and origin is not None:
        axis = NXOpen.Vector3d(*ROTATION_AXIS)
        for index in range(1, int(rotation_count) + 1):
            timed(
                "3.{0}".format(index),
                "View.Rotate (+{0} deg)".format(ROTATION_ANGLE_DEGREES),
                lambda index=index: rotate(origin, axis, ROTATION_ANGLE_DEGREES),
                "view change {0} of {1}; this is the measured symptom".format(
                    index, rotation_count
                ),
            )
        for index in range(1, int(rotation_count) + 1):
            timed(
                "4.{0}".format(index),
                "View.Rotate (-{0} deg)".format(ROTATION_ANGLE_DEGREES),
                lambda index=index: rotate(origin, axis, -ROTATION_ANGLE_DEGREES),
                "reverse rotation {0} of {1}".format(index, rotation_count),
            )
    else:
        rows.append(
            {
                "STEP": "3",
                "OPERATION": "View.Rotate",
                "SECONDS": "",
                "CUMULATIVE_SECONDS": "",
                "NOTE": (
                    "UNAVAILABLE: Rotate not exposed"
                    if not callable(rotate)
                    else "UNAVAILABLE: view Origin could not be read"
                ),
            }
        )

    fit = getattr(view, "Fit", None)
    if callable(fit):
        timed("5", "View.Fit", fit, "fit; changes zoom, orientation preserved")
    else:
        rows.append(
            {
                "STEP": "5",
                "OPERATION": "View.Fit",
                "SECONDS": "",
                "CUMULATIVE_SECONDS": "",
                "NOTE": "UNAVAILABLE: Fit not exposed",
            }
        )

    started = time.perf_counter()
    try:
        restore_method = restore_view_state(view, state)
    except Exception as error:
        notes.append("view restore failed: " + error_text(error))
    restore_seconds = elapsed_seconds(started)
    cumulative += restore_seconds
    rows.append(
        {
            "STEP": "6",
            "OPERATION": "Restore view state",
            "SECONDS": restore_seconds,
            "CUMULATIVE_SECONDS": cumulative,
            "NOTE": restore_method or "UNAVAILABLE: view state not restored",
        }
    )

    return rows, cumulative, notes


# ---------------------------------------------------------------------------
# Reports
# ---------------------------------------------------------------------------


def write_csv(path, columns, rows):
    with open(path, "w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(
            handle, fieldnames=columns, extrasaction="ignore"
        )
        writer.writeheader()
        for row in rows:
            writer.writerow(
                {column: row.get(column, "") for column in columns}
            )


def write_json(path, payload):
    with open(path, "w", encoding="utf-8", newline="") as handle:
        json.dump(payload, handle, indent=2, default=str)


def format_number(value):
    if value in ("", None):
        return ""
    try:
        numeric = float(value)
    except Exception:
        return clean(value)
    if numeric == int(numeric):
        return str(int(numeric))
    return "{0:.6g}".format(numeric)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------


def main():
    session = NXOpen.Session.GetSession()
    log_buffer = []
    folders = None
    ledger = None
    report_paths = []

    try:
        mode = resolve_mode()
        scope_filter = resolve_scope_filter()
        visible_scan = resolve_visible_scan()
        rotation_count = resolve_rotation_count()
        max_bodies = resolve_max_bodies()
        discover = resolve_discover()

        io_root = resolve_io_root()
        run_datetime = datetime.datetime.now(MYT_TIMEZONE)
        timestamp = run_datetime.strftime("%Y%m%d_%H%M%S")
        folders = create_run_folders(io_root, timestamp)
        ledger = EvidenceLedger(
            JOURNAL_BUILD_ID, timestamp, mode, scope_filter
        )

        log_line(session, "Journal 37 - Display performance triage", log_buffer)
        log_line(session, "Journal build: " + JOURNAL_BUILD_ID, log_buffer)
        log_line(
            session, "Journal source: " + runtime_source_path(), log_buffer
        )
        log_line(
            session,
            "Mode: {0}; scope: {1}; visible scan: {2}; rotations: {3}; "
            "discover: {4}".format(
                mode,
                scope_filter,
                "YES" if visible_scan else "NO",
                rotation_count if mode == "TIMED" else "n/a",
                "YES" if discover else "NO",
            ),
            log_buffer,
        )
        log_line(session, "I/O root: " + io_root, log_buffer)
        log_line(session, "Run folder: " + folders["run"], log_buffer)
        log_line(
            session,
            "Read-only: no load, no visibility/suppression/reference-set/"
            "layer change, no update, no save, no checkout.",
            log_buffer,
        )

        work_part = safe_property(session.Parts, "Work")
        if work_part is None:
            raise RuntimeError(
                "No work part is loaded. Open the slow assembly, make it the "
                "work part, fully load it, and run the journal again."
            )

        log_line(
            session,
            "Assembly root: {0}".format(safe_part_name(work_part)),
            log_buffer,
        )

        # --- traversal ---------------------------------------------------
        started = time.perf_counter()
        occurrences, targets, diagnostics, global_occurrences = (
            collect_occurrences(work_part, scope_filter)
        )
        traversal_seconds = elapsed_seconds(started)
        log_line(
            session,
            "Traversal: {0} occurrence(s), {1} unique prototype(s), "
            "{2} diagnostic(s), {3:.3f} s".format(
                global_occurrences,
                len(targets),
                len(diagnostics),
                traversal_seconds,
            ),
            log_buffer,
        )
        ledger.add(
            "Component.GetChildren",
            "assembly root",
            "OK",
            "occurrences={0}; prototypes={1}".format(
                global_occurrences, len(targets)
            ),
        )
        for diagnostic in diagnostics:
            log_line(
                session,
                "  {0} [{1}]: {2}".format(
                    diagnostic.get("code"),
                    diagnostic.get("component_path"),
                    diagnostic.get("message"),
                ),
                log_buffer,
            )

        # --- load truth ---------------------------------------------------
        log_line(session, "Load state per prototype:", log_buffer)
        for target in targets:
            state = load_state_text(target["part"])
            load_status = part_load_state(target["part"])[0]
            number, revision = part_identity(target["part"])
            identity = identity_label(
                target["part"], target["sample_path"]
            )
            ledger.add(
                "PartLoadState",
                identity,
                "OK" if load_status != "UNKNOWN" else "UNAVAILABLE",
                state,
            )
            if load_status != "FULLY_LOADED":
                log_line(
                    session,
                    "  NOT FULLY LOADED: {0} -> {1}".format(identity, state),
                    log_buffer,
                )
        minimal_total = 0
        for target in targets:
            status, names, error = minimally_loaded_children(target["part"])
            number, revision = part_identity(target["part"])
            identity = "{0}/{1}".format(
                number or safe_part_name(target["part"]), revision or "-"
            )
            ledger.add(
                "HasAnyMinimallyLoadedChildren",
                identity_label(target["part"], target["sample_path"]),
                status,
                len(names) if status == "OK" else "",
                error,
            )
            if names:
                minimal_total += len(names)
                log_line(
                    session,
                    "  MINIMALLY LOADED CHILDREN under {0}: {1}".format(
                        identity, "; ".join(names[:5])
                    ),
                    log_buffer,
                )
        log_line(
            session,
            "Minimally loaded children reported: {0}".format(minimal_total),
            log_buffer,
        )

        # --- geometry census ---------------------------------------------
        geometry_by_key = {}
        minimal_by_key = {}
        target_rows = []
        target_row_by_key = {}
        log_line(session, "Geometry census:", log_buffer)
        for target in targets:
            part = target["part"]
            number, revision = part_identity(part)
            identity = identity_label(part, target["sample_path"])
            started = time.perf_counter()
            geometry = census_geometry(part, max_bodies=max_bodies)
            census_seconds = elapsed_seconds(started)

            ledger.add(
                "Bodies",
                identity,
                geometry["status"],
                geometry["body_count"],
                geometry["error"],
            )
            ledger.add(
                "GetFaces",
                identity,
                geometry["status"],
                geometry["face_count"],
                geometry["error"],
            )
            ledger.add(
                "GetNumberOfFacets",
                identity,
                geometry["status"],
                geometry["facet_count"],
                geometry["error"],
            )
            ledger.add(
                "Density",
                identity,
                geometry["status"],
                "min={0}; max={1}; zero={2}".format(
                    format_number(geometry.get("density_min")),
                    format_number(geometry.get("density_max")),
                    geometry.get("density_zero_count", ""),
                ),
                geometry["error"],
            )

            for collection_name, key in PART_COLLECTION_SPECS:
                status, count, error = geometry["collection_status"].get(
                    collection_name, ("UNAVAILABLE", "", "not probed")
                )
                geometry[key] = count if status == "OK" else ""
                ledger.add(collection_name, identity, status, count, error)

            geometry_by_key[target["key"]] = geometry

            log_line(
                session,
                (
                    "  {0}: bodies={1} faces={2} facets={3} density("
                    "min={4} zero={5}) blanked={6} datums={7} curves={8} "
                    "{9:.3f} s"
                ).format(
                    identity,
                    format_number(geometry.get("body_count")),
                    format_number(geometry.get("face_count")),
                    format_number(geometry.get("facet_count")),
                    format_number(geometry.get("density_min")),
                    geometry.get("density_zero_count", ""),
                    format_number(geometry.get("blanked_body_count")),
                    format_number(geometry.get("datum_count")),
                    format_number(geometry.get("curve_count")),
                    census_seconds,
                ),
                log_buffer,
            )

            row = {
                "RUN_TIMESTAMP": timestamp,
                "JOURNAL_BUILD": JOURNAL_BUILD_ID,
                "MODE": mode,
                "SCOPE_FILTER": scope_filter,
                "DB_PART_NO": number,
                "DB_PART_REV": revision,
                "PART_NAME": safe_part_name(part),
                "PART_KIND": part_kind(part),
                "PART_NX_TYPE": nx_type_name(part),
                "IS_WORK_PART": "YES" if target["is_work_part"] else "NO",
                "LEVEL": target["level"],
                "DEEPEST_LEVEL": target["deepest_level"],
                "OCCURRENCE_COUNT": target["occurrence_count"],
                "GLOBAL_OCCURRENCE_COUNT": global_occurrences,
                "SAMPLE_PATH": target["sample_path"],
                "LOAD_STATE": load_state_text(part),
                "IS_FULLY_LOADED": {
                    "FULLY_LOADED": "YES",
                    "NOT_FULLY_LOADED": "NO",
                }.get(part_load_state(part)[0], "UNKNOWN"),
                "MATERIAL": material_text(part),
                "PROBE_STATUS": geometry["status"],
                "PROBE_ERROR": geometry["error"],
            }
            flags = part_display_flags(part)
            row["SAVE_DISPLAY_FACETS"] = flags.get("save_display_facets", "")
            row["PART_PREVIEW_MODE"] = flags.get("part_preview_mode", "")
            row["IS_DESIGN_REVIEW_PART"] = flags.get(
                "is_design_review_part", ""
            )
            row["IS_DISPLAYED"] = flags.get("is_displayed", "")
            for column, key in (
                ("BODY_COUNT", "body_count"),
                ("SOLID_BODY_COUNT", "solid_body_count"),
                ("SHEET_BODY_COUNT", "sheet_body_count"),
                ("CONVERGENT_BODY_COUNT", "convergent_body_count"),
                ("FACE_COUNT", "face_count"),
                ("EDGE_COUNT", "edge_count"),
                ("FACET_COUNT", "facet_count"),
                ("VERTEX_COUNT", "vertex_count"),
                ("BODIES_ENUMERATED", "bodies_enumerated"),
                ("GEOMETRY_TRUNCATED", "truncated"),
                ("BLANKED_BODY_COUNT", "blanked_body_count"),
                ("DATUM_COUNT", "datum_count"),
                ("COORDINATE_SYSTEM_COUNT", "coordinate_system_count"),
                ("CURVE_COUNT", "curve_count"),
                ("LINE_COUNT", "line_count"),
                ("POINT_COUNT", "point_count"),
                ("ANNOTATION_COUNT", "annotation_count"),
                ("FEATURE_COUNT", "feature_count"),
                ("EXPRESSION_COUNT", "expression_count"),
                ("TRUE_SHADING_COUNT", "true_shading_count"),
                ("TRUE_STUDIO_COUNT", "true_studio_count"),
                ("POINT_CLOUD_COUNT", "point_cloud_count"),
                ("DECAL_COUNT", "decal_count"),
                ("CAMERA_COUNT", "camera_count"),
                ("DYNAMIC_SECTION_COUNT", "dynamic_section_count"),
                ("IMAGE_COUNT", "image_count"),
                ("DRAWING_SHEET_COUNT", "drawing_sheet_count"),
                ("DENSITY_MIN", "density_min"),
                ("DENSITY_MAX", "density_max"),
                ("DENSITY_AVERAGE", "density_average"),
                ("DENSITY_ZERO_COUNT", "density_zero_count"),
                ("DENSITY_UNAVAILABLE_COUNT", "density_unavailable_count"),
            ):
                value = geometry.get(key, "")
                row[column] = (
                    format_number(value)
                    if column.startswith("DENSITY") and isinstance(value, float)
                    else value
                )

            minimal_status, minimal_names, minimal_error = (
                minimally_loaded_children(part)
            )
            minimal_by_key[target["key"]] = (
                minimal_status,
                minimal_names,
                minimal_error,
            )
            if minimal_names:
                ledger.add(
                    "GetMinimallyLoadedParts",
                    identity,
                    minimal_status,
                    "; ".join(minimal_names[:10]),
                    minimal_error,
                )
            row["HAS_MINIMALLY_LOADED_CHILDREN"] = (
                "YES"
                if minimal_names
                else ("NO" if minimal_status == "OK" else minimal_status)
            )
            row["MINIMALLY_LOADED_CHILDREN"] = "; ".join(minimal_names[:10])

            target_rows.append(row)
            target_row_by_key[target["key"]] = row

        # --- layer census -------------------------------------------------
        layer_rows = []
        layer_info_by_key = {}
        for target in targets:
            part = target["part"]
            number, revision = part_identity(part)
            identity = identity_label(part, target["sample_path"])
            rows, status, error, visible_layers = census_layers(part)
            ledger.add("Layers", identity, status, visible_layers, error)
            populated = sum(1 for row in rows if int(row["TOTAL_COUNT"]) > 0)
            hidden_populated = sum(
                1
                for row in rows
                if int(row["TOTAL_COUNT"]) > 0 and row["IS_HIDDEN"] == "YES"
            )
            layer_info_by_key[target["key"]] = {
                "visible_layers": visible_layers,
                "populated_layers": populated,
                "hidden_populated_layers": hidden_populated,
            }
            for row in rows:
                layer_rows.append(
                    {
                        "RUN_TIMESTAMP": timestamp,
                        "DB_PART_NO": number,
                        "DB_PART_REV": revision,
                        "PART_NAME": safe_part_name(part),
                        **row,
                    }
                )
            if populated:
                log_line(
                    session,
                    "  {0}: visible layers={1}; layers with objects={2}; "
                    "hidden layers with objects={3}".format(
                        identity, visible_layers, populated, hidden_populated
                    ),
                    log_buffer,
                )

        # Layer totals live in the per-part row; fill them after the census.
        for target in targets:
            row = target_row_by_key.get(target["key"])
            if row is None:
                continue
            info = layer_info_by_key.get(target["key"], {})
            row["VISIBLE_LAYER_COUNT"] = info.get("visible_layers", "")
            row["LAYER_WITH_OBJECT_COUNT"] = info.get("populated_layers", "")
            row["HIDDEN_LAYER_WITH_OBJECT_COUNT"] = info.get(
                "hidden_populated_layers", ""
            )
            row["LAYERS_TRUNCATED"] = "NO"

        # --- visible objects ----------------------------------------------
        visible_rows = []
        visible_total = ""
        visible_status = "SKIPPED"
        visible_error = ""
        visible_seconds = ""
        if visible_scan:
            view_status, view, view_error = active_view(session)
            if view_status != "OK" or view is None:
                visible_status = view_status
                visible_error = view_error
            else:
                (
                    visible_status,
                    rows,
                    visible_total,
                    visible_seconds,
                    truncated,
                    visible_error,
                ) = census_visible_objects(view)
                ledger.add(
                    "AskVisibleObjects",
                    safe_property(view, "Name") or "active view",
                    visible_status,
                    visible_total,
                    visible_error,
                )
                view_name = clean(safe_property(view, "Name")) or "<active view>"
                rendering_style = enum_text(
                    safe_property(view, "RenderingStyle")
                )
                for row in rows:
                    visible_rows.append(
                        {
                            "RUN_TIMESTAMP": timestamp,
                            "VIEW_NAME": view_name,
                            "VIEW_RENDERING_STYLE": rendering_style,
                            "TOTAL_VISIBLE_OBJECTS": visible_total,
                            "SCAN_SECONDS": seconds_text(visible_seconds),
                            "TRUNCATED": truncated,
                            **row,
                        }
                    )
                if visible_status == "OK":
                    log_line(
                        session,
                        "Active view '{0}': {1} visible object(s) in {2} s "
                        "(rendering style: {3})".format(
                            view_name,
                            visible_total,
                            seconds_text(visible_seconds),
                            rendering_style or "unavailable",
                        ),
                        log_buffer,
                    )
                    for row in rows[:10]:
                        log_line(
                            session,
                            "  {0}: {1} ({2}%)".format(
                                row["TYPE_NAME"],
                                row["TYPE_COUNT"],
                                row["TYPE_PERCENT"],
                            ),
                            log_buffer,
                        )
                else:
                    log_line(
                        session,
                        "Visible-object scan: {0} {1}".format(
                            visible_status, visible_error
                        ),
                        log_buffer,
                    )
        else:
            log_line(
                session,
                "Visible-object scan disabled (NX_J37_VISIBLE_SCAN=NO).",
                log_buffer,
            )

        # --- preferences ---------------------------------------------------
        pref_rows = census_preferences(
            session, [t["part"] for t in targets]
        )
        ledger.add(
            "Preferences",
            "session",
            "OK",
            "probes={0}".format(len(pref_rows)),
        )
        for row in pref_rows:
            ledger.add(
                "PREF_{0}".format(row["KEY"]),
                row["OWNER"],
                row["STATUS"],
                row["VALUE"],
                row["ERROR"],
            )
        log_line(session, "Display preferences:", log_buffer)
        for row in pref_rows:
            if row["SCOPE"] != "SESSION":
                continue
            log_line(
                session,
                "  {0} = {1} [{2}]".format(
                    row["KEY"],
                    row["VALUE"] if row["STATUS"] == "OK" else row["STATUS"],
                    row["ERROR"] if row["STATUS"] != "OK" else "OK",
                ),
                log_buffer,
            )

        # --- timing --------------------------------------------------------
        timing_rows = []
        timing_total = ""
        if mode == "TIMED":
            view_status, view, view_error = active_view(session)
            if view_status != "OK" or view is None:
                log_line(
                    session,
                    "TIMED requested but no view is available: {0}".format(
                        view_error
                    ),
                    log_buffer,
                )
            else:
                view_name = clean(safe_property(view, "Name")) or "<active view>"
                log_line(
                    session,
                    "Measuring view change on '{0}' ({1} rotation pair(s)); "
                    "this can take as long as the reported symptom.".format(
                        view_name, rotation_count
                    ),
                    log_buffer,
                )
                rows, total, notes = run_timing(
                    session, view, rotation_count
                )
                timing_total = total
                for note in notes:
                    log_line(session, "  WARNING: " + note, log_buffer)
                for row in rows:
                    timing_rows.append(
                        {"RUN_TIMESTAMP": timestamp, **row}
                    )
                    log_line(
                        session,
                        "  {0} {1}: {2} s".format(
                            row["STEP"],
                            row["OPERATION"],
                            seconds_text(row["SECONDS"]),
                        ),
                        log_buffer,
                    )
                log_line(
                    session,
                    "Timing total: {0} s".format(seconds_text(total)),
                    log_buffer,
                )
        else:
            log_line(
                session,
                "View-change timing skipped (PROBE mode). Rerun with "
                "NX_J37_MODE=TIMED to measure the symptom.",
                log_buffer,
            )

        # --- optional API discovery ----------------------------------------
        discover_path = ""
        discover_rows = []
        if discover:
            try:
                discover_rows = discover_attributes(
                    session, [t["part"] for t in targets]
                )
            except Exception:
                log_line(session, traceback.format_exc(), log_buffer)
                discover_rows = []
            discover_path = os.path.join(
                folders["reports"], "J37_DISCOVER_{0}.csv".format(timestamp)
            )
            write_csv(discover_path, DISCOVER_COLUMNS, discover_rows)
            log_line(
                session,
                "API discovery rows: {0} -> {1}".format(
                    len(discover_rows), discover_path
                ),
                log_buffer,
            )
            for row in discover_rows:
                if row["STATUS"] == "OK" and any(
                    token in row["MEMBER"].lower()
                    for token in (
                        "lightweight",
                        "faceted",
                        "facet",
                        "represent",
                        "preview",
                    )
                ):
                    log_line(
                        session,
                        "  {0}.{1} = {2}".format(
                            row["OWNER"], row["MEMBER"], row["VALUE"]
                        ),
                        log_buffer,
                    )

        # --- suspects ------------------------------------------------------
        # Aggregate occurrence-level probes so every suspect line can cite a
        # fact that was actually recorded.
        suppressed_count = sum(
            1 for row in occurrences if row["IS_SUPPRESSED"] == "YES"
        )
        entire_part_count = sum(
            1 for row in occurrences if row["IS_ENTIRE_PART_REFSET"] == "YES"
        )
        representation_counts = {}
        for row in occurrences:
            mode_name = row["REPRESENTATION_MODE"] or "<unavailable>"
            representation_counts[mode_name] = (
                representation_counts.get(mode_name, 0) + 1
            )
        ledger.add(
            "ReferenceSet",
            "occurrences={0}".format(len(occurrences)),
            "OK",
            "entire_part_refset_occurrences={0}".format(entire_part_count),
        )
        ledger.add(
            "EntirePartRefsetName",
            "occurrences={0}".format(len(occurrences)),
            "OK",
            "entire_part_refset_occurrences={0}".format(entire_part_count),
        )
        ledger.add(
            "RepresentationMode",
            "occurrences={0}".format(len(occurrences)),
            "OK",
            " | ".join(
                "{0}={1}".format(key, representation_counts[key])
                for key in sorted(representation_counts)
            ),
        )
        ledger.add(
            "IsSuppressed",
            "occurrences={0}".format(len(occurrences)),
            "OK",
            "suppressed={0}".format(suppressed_count),
        )
        ledger.add(
            "GetState",
            "layers",
            "OK",
            "layer state read for parts with a Layers collection",
        )
        convergent_total = sum(
            1
            for geometry in geometry_by_key.values()
            if clean(str(geometry.get("convergent_body_count", ""))) not in ("", "0")
        )
        blanked_total = sum(
            1
            for geometry in geometry_by_key.values()
            if clean(str(geometry.get("blanked_body_count", ""))) not in ("", "0")
        )
        ledger.add(
            "IsConvergentBody",
            "prototypes_with_convergent_bodies",
            "OK",
            convergent_total,
        )
        ledger.add(
            "IsBlanked",
            "prototypes_with_blanked_bodies",
            "OK",
            blanked_total,
        )

        suspects = build_suspects(
            ledger,
            targets,
            occurrences,
            geometry_by_key,
            layer_info_by_key,
            minimal_by_key,
            visible_total,
            pref_rows,
            diagnostics,
        )
        log_line(session, "Ranked suspects:", log_buffer)
        if not suspects:
            log_line(
                session,
                "  No suspect crossed a configured threshold.",
                log_buffer,
            )
        for index, suspect in enumerate(suspects, start=1):
            suspect["RANK"] = index
        for suspect in suspects[:TOP_SUSPECTS]:
            log_line(
                session,
                "  {0}. [{1}] {2} | {3} | measured={4} threshold={5}".format(
                    suspect["RANK"],
                    suspect["SEVERITY"],
                    suspect["CODE"],
                    suspect["IDENTITY"],
                    suspect["MEASURED_VALUE"],
                    suspect["THRESHOLD"],
                ),
                log_buffer,
            )
            log_line(session, "     " + suspect["MESSAGE"], log_buffer)

        # --- reports -------------------------------------------------------
        reports = folders["reports"]
        target_path = os.path.join(reports, "J37_TARGETS_{0}.csv".format(timestamp))
        occurrence_path = os.path.join(
            reports, "J37_OCCURRENCES_{0}.csv".format(timestamp)
        )
        layer_path = os.path.join(reports, "J37_LAYERS_{0}.csv".format(timestamp))
        visible_path = os.path.join(
            reports, "J37_VISIBLE_{0}.csv".format(timestamp)
        )
        pref_path = os.path.join(reports, "J37_PREFS_{0}.csv".format(timestamp))
        suspect_path = os.path.join(
            reports, "J37_SUSPECTS_{0}.csv".format(timestamp)
        )
        timing_path = os.path.join(reports, "J37_TIMING_{0}.csv".format(timestamp))
        evidence_path = os.path.join(
            reports, "J37_EVIDENCE_{0}.json".format(timestamp)
        )
        log_path = os.path.join(folders["logs"], "J37_LOG_{0}.txt".format(timestamp))

        write_csv(target_path, TARGET_COLUMNS, target_rows)
        write_csv(occurrence_path, OCCURRENCE_COLUMNS, occurrences)
        write_csv(layer_path, LAYER_COLUMNS, layer_rows)
        write_csv(visible_path, VISIBLE_COLUMNS, visible_rows)
        write_csv(pref_path, PREF_COLUMNS, pref_rows)
        write_csv(suspect_path, SUSPECT_COLUMNS, suspects)
        write_csv(timing_path, TIMING_COLUMNS, timing_rows)
        report_paths = [
            target_path,
            occurrence_path,
            layer_path,
            visible_path,
            pref_path,
            suspect_path,
            timing_path,
        ]

        high = sum(1 for s in suspects if s["SEVERITY"] == "HIGH")
        medium = sum(1 for s in suspects if s["SEVERITY"] == "MEDIUM")
        info = sum(1 for s in suspects if s["SEVERITY"] == "INFO")

        def total_faces():
            total = 0
            for geometry in geometry_by_key.values():
                try:
                    total += int(geometry.get("face_count") or 0)
                except Exception:
                    continue
            return total

        def total_bodies():
            total = 0
            for geometry in geometry_by_key.values():
                try:
                    total += int(geometry.get("body_count") or 0)
                except Exception:
                    continue
            return total

        unnamed_prototypes = sum(
            1
            for target in targets
            if not part_identity(target["part"])[0]
            and not safe_part_name(target["part"], fallback="")
        )
        not_fully_loaded_count = sum(
            1
            for target in targets
            if part_load_state(target["part"])[0] != "FULLY_LOADED"
        )
        totals = {
            "occurrences": global_occurrences,
            "prototypes": len(targets),
            "faces_total": total_faces(),
            "bodies_total": total_bodies(),
            "entire_part_refset_occurrences": entire_part_count,
            "entire_part_refset_parts": sum(
                1
                for suspect in suspects
                if suspect["CODE"] == "ENTIRE_PART_REFSET"
            ),
            "not_fully_loaded_prototypes": not_fully_loaded_count,
            "unnamed_prototypes": unnamed_prototypes,
            "visible_objects": visible_total,
            "suspects": len(suspects),
        }

        payload = ledger.to_dict()
        payload.update(
            {
                "assembly_root": safe_part_name(work_part),
                "occurrence_count": global_occurrences,
                "prototype_count": len(targets),
                "diagnostics": diagnostics,
                "suspect_counts": {
                    "HIGH": high,
                    "MEDIUM": medium,
                    "INFO": info,
                },
                "visible_objects": visible_total,
                "visible_scan_status": visible_status,
                "visible_scan_error": visible_error,
                "visible_scan_seconds": visible_seconds,
                "timing_total_seconds": timing_total,
                "timing_rows": timing_rows,
                "suspects": suspects,
                "report_paths": report_paths,
                "totals": totals,
                "visible_rows": visible_rows,
                "discover_rows": discover_rows,
                "discover_report": discover_path,
                "reports": {
                    "targets": target_path,
                    "occurrences": occurrence_path,
                    "layers": layer_path,
                    "visible": visible_path,
                    "preferences": pref_path,
                    "suspects": suspect_path,
                    "timing": timing_path,
                },
            }
        )
        write_json(evidence_path, payload)

        log_line(session, "Triage complete", log_buffer)
        log_line(
            session,
            "Totals: occurrences={0}; prototypes={1}; faces={2}; bodies={3}; "
            "entire-part-refset occurrences={4}; not fully loaded={5}; "
            "unnamed prototypes={6}".format(
                totals["occurrences"],
                totals["prototypes"],
                totals["faces_total"],
                totals["bodies_total"],
                totals["entire_part_refset_occurrences"],
                totals["not_fully_loaded_prototypes"],
                totals["unnamed_prototypes"],
            ),
            log_buffer,
        )
        log_line(
            session,
            "Suspects: HIGH={0}; MEDIUM={1}; INFO={2}".format(
                high, medium, info
            ),
            log_buffer,
        )
        log_line(session, "Evidence JSON: " + evidence_path, log_buffer)
        log_line(session, "Reports: " + reports, log_buffer)
        write_text_log(log_path, log_buffer)

    except Exception:
        log_line(session, "ERROR: Unhandled journal exception.", log_buffer)
        log_line(session, traceback.format_exc(), log_buffer)

    finally:
        if folders is not None:
            try:
                fallback_log = os.path.join(
                    folders["logs"],
                    "J37_LOG_{0}.txt".format(
                        os.path.basename(folders["run"])
                    ),
                )
                write_text_log(fallback_log, log_buffer)
            except Exception:
                pass

        if ledger is not None:
            try:
                fallback_evidence = os.path.join(
                    folders["reports"],
                    "J37_EVIDENCE_{0}.json".format(
                        os.path.basename(folders["run"])
                    ),
                )
                if not os.path.exists(fallback_evidence):
                    write_json(fallback_evidence, ledger.to_dict())
            except Exception:
                pass


if __name__ == "__main__":
    main()
