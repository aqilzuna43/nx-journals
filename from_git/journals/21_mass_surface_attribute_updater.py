"""Journal 21 - Assembly Mass & Surface Area Attribute Updater (NX 2506)

Drives NX's NATIVE Mass Properties update on the open assembly - the same
engine behind Tools > Measure Mass Properties with Update On Save.  NX itself
computes and writes its standard attributes on every component:

    NX_MassPropRollupMass  roll-up mass (kg)      [Rolled-Up Mass Properties]
    NX_MassPropRollupArea  roll-up area (mm^2)    [Rolled-Up Mass Properties]

The journal does NOT create, compute, or write attributes itself.  In APPLY
mode it:
  1. fully loads the complete BoM-visible subtree in memory (same filter as
     NXOpenBoMExtended.py / Journal 04: suppressed, reference-only, and
     keyword-named occurrences are excluded),
  2. scans every unique prototype for NX_MassPropRollupMass,
  3. processes only prototypes where that attribute is absent, bottom-up,
  4. makes each selected target the work part, triggers NX's native
     mass-properties update, and saves that target.

The load phase never checks out, saves, or updates mass.  APPLY continues on
independent branches when a prototype cannot load, while blocking a missing
assembly whose roll-up depends on that failed branch.  REFRESH_ALL preserves
the V5 all-or-nothing full rebuild.  Successfully loaded parts remain loaded.

J21 never checks a part out.  In Teamcenter managed mode it reports checkout
state and owner, skips parts that are checked in, checked out by somebody
else, or read-only, and continues with everything it can update.  The
original display/work context is restored after the run.

WRITE_MODE defaults to "APPLY".  APPLY fills missing roll-up mass only.  Use
"REFRESH_ALL" for the V5 full bottom-up rebuild, "DRY_RUN" to report without
loading/updating/saving, "SMOKE" to force one active-part update, or "PROBE"
to dump the MassPropertiesBuilder API surface.

Why native and not direct write: NX_MassPropRollupMass / NX_MassPropRollupArea
are RESERVED NX attribute titles - a journal cannot write them with
AttributePropertiesBuilder (NX raises "This is a reserved attribute title.
[512006]").  Only NX's own mass-properties update can populate them, so this
journal only triggers that native update and reports what NX wrote.

Note: SMOKE measures only the active work part.  APPLY trusts every existing
roll-up mass value, including 0.0; only an absent exact attribute title is
selected.  Missing roll-up area alone never selects a target.

Target: NX X 2506 embedded Python only
Run via: NX > Tools > Journal > Play
"""

import csv
import datetime
import os
import traceback

import NXOpen


BUILD = "J21-NX2506-SCAN-MISSING-ROLLUP-MASS-V6"
WRITE_MODE = "APPLY"  # APPLY / REFRESH_ALL / DRY_RUN / SMOKE / PROBE
OUTPUT_FOLDER = "NX_MASS_SURFACE_UPDATE"
MEASUREMENT_ACCURACY = 0.99
MASS_DECIMAL_PLACES = 6
AREA_DECIMAL_PLACES = 2
AREA_M2_DECIMAL_PLACES = 4
# NX stores the roll-up area in square millimetres (PDM template); the report
# also presents it in square metres for readability on large systems.
SQUARE_METRES_PER_SQUARE_MILLIMETRE = 1e-6
MAX_OCCURRENCES = 100000
MAX_LOAD_PASSES = 100

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

# Standard NX roll-up attributes (category "Rolled-Up Mass Properties").
ROLLUP_MASS_ATTRIBUTE = "NX_MassPropRollupMass"
ROLLUP_AREA_ATTRIBUTE = "NX_MassPropRollupArea"

# --- BOM VISIBILITY (mirrors NXOpenBoMExtended.py and Journal 04) ---
IGNORE_KEYWORDS = ["CSYS", "COORDINATE", "DATUM", "REFERENCE", "SKELETON"]
BOM_REFERENCE_ATTRIBUTES = ("REFERENCE_COMPONENT", "PLIST_IGNORE_MEMBER")
BOM_EXCLUSION_ATTRIBUTE = "CELESTICA_BOM_EXCLUDE_SUBTREE"
BOM_EXCLUSION_VALUE = "YES"
# NX marks native reference components with an empty string; manual overrides
# use YES/1/True/true/yes.
BOM_REFERENCE_FLAG_VALUES = ("", "YES", "1", "True", "true", "yes")

RESULT_COLUMNS = (
    "ROW_TYPE",
    "RUN_TIMESTAMP",
    "JOURNAL_BUILD",
    "WRITE_MODE",
    "DB_PART_NO",
    "DB_PART_REV",
    "PART_NAME",
    "COMPONENT_PATH",
    "LEVEL",
    "PROCESS_ORDER",
    "PART_KIND",
    "INITIAL_LOAD_STATE",
    "LOAD_ACTION",
    "FINAL_LOAD_STATE",
    "LOAD_STATUS",
    "LOAD_MESSAGE",
    "LOAD_STATE",
    "CHECKOUT_STATE",
    "CHECKOUT_OWNER",
    "CURRENT_USER",
    "READ_ONLY",
    "UPDATE",
    "INITIAL_ROLLUP_MASS_KG",
    "MASS_SCAN_STATUS",
    "SELECTION",
    "SELECTION_REASON",
    "DEPENDENCY_STATUS",
    "ROLLUP_MASS_KG",
    "ROLLUP_AREA_MM2",
    "ROLLUP_AREA_M2",
    "ROLLUP_MASS_ATTRIBUTE",
    "ROLLUP_AREA_ATTRIBUTE",
    "SAVED",
    "STATUS",
    "MESSAGE",
)

_INVALID_FILENAME_CHARS = '<>:"/\\|?*'


def _text(value):
    return "" if value is None else str(value)


def clean(value):
    return _text(value).strip()


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


def log_line(session, message):
    text = str(message)
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
    return os.path.abspath(
        os.path.expanduser(configured or desktop_folder())
    )


def clean_filename_token(value, fallback="UNKNOWN"):
    text = clean(value)
    if not text:
        return fallback
    result = "".join(
        "_" if char in _INVALID_FILENAME_CHARS or ord(char) < 32 else char
        for char in text
    ).strip(" .")
    return result or fallback


def get_string_attribute(nx_object, name):
    try:
        return clean(nx_object.GetStringAttribute(name))
    except Exception:
        pass

    try:
        attribute = nx_object.GetUserAttribute(
            name,
            NXOpen.NXObject.AttributeType.String,
            -1,
        )
        return clean(attribute.StringValue)
    except Exception:
        return ""


def safe_part_name(part):
    for property_name in ("Name", "Leaf", "FullPath"):
        try:
            value = clean(getattr(part, property_name))
            if value:
                return value
        except Exception:
            pass
    return "UNKNOWN"


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
    return {
        "number": number,
        "revision": revision,
        "name": safe_part_name(part),
    }


def _object_key(nx_object):
    tag = getattr(nx_object, "Tag", None)
    return ("TAG", _text(tag)) if tag is not None else ("PY", id(nx_object))


def _safe_property(nx_object, name, fallback=None):
    try:
        value = getattr(nx_object, name)
        return value() if callable(value) else value
    except Exception:
        return fallback


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
    """Mirror NXOpenBoMExtended.py: only BoM-visible components are pulled.

    Suppression is handled separately by the caller.  Keyword-named and
    reference-flagged occurrences are excluded together with their subtrees.
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


def collect_bom_scope(work_part):
    """Return unique targets, diagnostics, and prototype dependencies.

    ``dependencies[parent_key]`` contains the unique prototype keys directly
    used by that assembly prototype.  Diagnostics carry ``blocked_keys`` for
    the known assembly ancestors whose roll-up cannot be trusted when an
    occurrence/prototype is unresolved.  This lets APPLY continue on sibling
    branches without updating a dependent assembly.
    """
    unique = {}
    diagnostics = []
    dependencies = {}
    occurrence_count = 0
    root_identity = part_identity(work_part)
    root_path = root_identity["number"] or root_identity["name"]
    root_key = _object_key(work_part)

    def add_part(part, level, path):
        key = _object_key(part)
        dependencies.setdefault(key, set())
        if key not in unique:
            unique[key] = (part, level, path)
        elif level > unique[key][1]:
            # Shared prototypes are processed once, at their deepest observed
            # level, so sorting by descending level remains bottom-up.
            unique[key] = (part, level, path)

    add_part(work_part, 0, root_path)
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
        diagnostics[-1]["blocked_keys"] = [root_key]
        return list(unique.values()), diagnostics, dependencies
    if root_component is None:
        return list(unique.values()), diagnostics, dependencies

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
        diagnostics[-1]["blocked_keys"] = [root_key]
        return list(unique.values()), diagnostics, dependencies

    stack = [
        (component, 1, root_path, root_key, (root_key,))
        for component in reversed(root_children)
    ]
    while stack:
        component, level, parent_path, parent_key, ancestor_keys = stack.pop()
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
                    "blocked_keys": list(ancestor_keys),
                }
            )
            break
        if not _is_active_visible(component):
            continue
        if not _is_bom_visible(component):
            continue
        prototype_error = ""
        try:
            prototype = getattr(component, "Prototype", None)
        except Exception as error:
            prototype = None
            prototype_error = error_text(error)
        if prototype_error:
            diagnostics.append(
                {
                    "code": classify_load_failure(
                        prototype_error, "PROTOTYPE_UNAVAILABLE"
                    ),
                    "message": "Component prototype could not be read: "
                    + prototype_error,
                    "component_path": component_path,
                    "level": level,
                    "blocked_keys": list(ancestor_keys),
                }
            )
            child_parent_key = parent_key
            child_ancestor_keys = ancestor_keys
        elif prototype is None:
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
                    "blocked_keys": list(ancestor_keys),
                }
            )
            child_parent_key = parent_key
            child_ancestor_keys = ancestor_keys
        else:
            add_part(prototype, level, component_path)
            prototype_key = _object_key(prototype)
            dependencies.setdefault(parent_key, set()).add(prototype_key)
            child_parent_key = prototype_key
            child_ancestor_keys = ancestor_keys + (prototype_key,)

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
                    "blocked_keys": list(child_ancestor_keys),
                }
            )
        stack.extend(
            (
                child,
                level + 1,
                component_path,
                child_parent_key,
                child_ancestor_keys,
            )
            for child in reversed(children)
        )
    return list(unique.values()), diagnostics, dependencies


def collect_unique_parts(work_part):
    """Compatibility wrapper returning BoM-visible targets and diagnostics."""
    parts, diagnostics, _dependencies = collect_bom_scope(work_part)
    return parts, diagnostics


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
    identity = part_identity(part)
    label = identity["number"] or identity["name"]
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


def auto_load_bom_visible(work_part, logger=None):
    """Load visible unique targets, re-traversing until scope is stable."""
    records = {}
    final_parts = []
    final_diagnostics = []
    final_dependencies = {}

    for pass_index in range(1, MAX_LOAD_PASSES + 1):
        parts, _diagnostics, _dependencies = collect_bom_scope(work_part)
        for parent_key, child_keys in _dependencies.items():
            final_dependencies.setdefault(parent_key, set()).update(child_keys)
        attempted = False
        for part, level, component_path in parts:
            key = _object_key(part)
            if key in records:
                if level > records[key]["level"]:
                    records[key]["level"] = level
                    records[key]["component_path"] = component_path
                continue
            record = load_target(
                part, level, component_path, logger=logger
            )
            records[key] = record
            if record["load_action"] == "LOAD_THIS_PART_FULLY":
                attempted = True

        (
            final_parts,
            final_diagnostics,
            current_dependencies,
        ) = collect_bom_scope(work_part)
        for parent_key, child_keys in current_dependencies.items():
            final_dependencies.setdefault(parent_key, set()).update(child_keys)
        final_keys = {_object_key(part) for part, _level, _path in final_parts}
        known_keys = set(records)
        all_recorded = final_keys.issubset(known_keys)
        all_loaded = all(
            records[key]["load_status"] == "SUCCESS"
            and part_load_state(records[key]["part"])[0] == "FULLY_LOADED"
            for key in final_keys
            if key in records
        )
        if all_recorded and all_loaded and not final_diagnostics:
            return True, final_parts, records, [], final_dependencies

        new_targets_exist = not all_recorded
        if not attempted and not new_targets_exist:
            break

        if logger:
            logger(
                "FULL LOAD PASS {0}: targets={1}; new_targets={2}".format(
                    pass_index, len(final_parts), "YES" if new_targets_exist else "NO"
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
            identity = part_identity(record["part"])
            label = identity["number"] or identity["name"]
            failures.append(
                {
                    "code": record["load_status"],
                    "message": "{0}: {1}".format(
                        label, record["load_message"]
                    ),
                    "component_path": record["component_path"],
                    "level": record["level"],
                    "blocked_keys": [key],
                }
            )
    final_keys = {
        _object_key(part) for part, _level, _path in final_parts
    }
    for key, record in records.items():
        if key not in final_keys:
            final_parts.append(
                (
                    record["part"],
                    record["level"],
                    record["component_path"],
                )
            )
    return (
        False,
        final_parts,
        records,
        final_diagnostics + failures,
        final_dependencies,
    )


def _update_on_save_yes(builder):
    """Resolve the UpdateOnSave=Yes member from the builder instance.

    NX 2506 exposes the nested UpdateOptions enum on the builder; module
    namespace lookups are unreliable across builds.
    """
    try:
        options = getattr(builder, "UpdateOptions", None)
        if options is not None:
            return getattr(options, "Yes", None)
    except Exception:
        pass
    return None


def _create_mass_properties_builder(work_part, objects):
    """Create the native MassPropertiesBuilder from the correct manager.

    NX places CreateMassPropertiesBuilder on PropertiesManager (NX12+) and,
    on some builds, also on MeasureManager.  Returns (builder, manager_name).
    """
    attempts = ()
    properties_manager = getattr(work_part, "PropertiesManager", None)
    if properties_manager is not None:
        attempts += ((properties_manager, "PropertiesManager"),)
    measure_manager = getattr(work_part, "MeasureManager", None)
    if measure_manager is not None:
        attempts += ((measure_manager, "MeasureManager"),)
    last_error = None
    for manager, name in attempts:
        create = getattr(manager, "CreateMassPropertiesBuilder", None)
        if create is None:
            last_error = "{0} has no CreateMassPropertiesBuilder".format(name)
            continue
        try:
            return create(objects), name
        except Exception as error:
            last_error = "{0}: {1}".format(name, error_text(error))
    raise RuntimeError(
        "No MassPropertiesBuilder factory found: {0}".format(
            last_error or "no PropertiesManager/MeasureManager"
        )
    )


def run_native_mass_property_update(work_part, objects=None):
    """Trigger NX's native roll-up mass property update on one target part.

    NX itself computes and writes NX_MassPropRollupMass / NX_MassPropRollupArea
    (and the rest of the standard family) on the measured object.  APPLY calls
    this function separately for every unique prototype.  Returns a status
    message; raises nothing unless the update cannot be started.
    """
    warnings = []
    builder = None
    try:
        if objects is None:
            root_component = getattr(
                getattr(work_part, "ComponentAssembly", None),
                "RootComponent",
                None,
            )
            objects = (
                [root_component]
                if root_component is not None
                else [work_part]
            )
        builder, manager_name = _create_mass_properties_builder(
            work_part, objects
        )
        warnings.append("factory: {0}".format(manager_name))

        builder.Accuracy = MEASUREMENT_ACCURACY
        # NX 2506 has no RollUp builder option: roll-up is implicit because
        # the assembly root component is the measured object.
        if getattr(builder, "UpdateOnSave", None) is not None:
            yes = _update_on_save_yes(builder)
            if yes is not None:
                builder.UpdateOnSave = yes
            else:
                warnings.append("UpdateOnSave Yes value unavailable")
        else:
            warnings.append("UpdateOnSave option unavailable")
        update_now = getattr(builder, "UpdateNow", None)
        commit = getattr(builder, "Commit", None)
        if update_now is None and commit is None:
            raise RuntimeError(
                "MassPropertiesBuilder has neither UpdateNow nor Commit."
            )
        # Compute immediately, then Commit creates the mass-property update
        # feature; without Commit no attributes are written at save time.
        if update_now is not None:
            update_now()
        if commit is not None:
            commit()
        else:
            warnings.append("Commit unavailable (update feature not created)")
        if warnings:
            return "NATIVE_UPDATE_OK ({0})".format(
                "; ".join(warnings)
            )
        return "NATIVE_UPDATE_OK"
    except Exception as error:
        return "NATIVE_UPDATE_FAILED: " + error_text(error)
    finally:
        dispose(builder)


def probe_builder_api(work_part):
    """Dump the MassPropertiesBuilder API surface of this NX build."""
    lines = []
    for manager_name in ("PropertiesManager", "MeasureManager"):
        manager = getattr(work_part, manager_name, None)
        lines.append("{0} members:".format(manager_name))
        if manager is None:
            lines.append("  <unavailable>")
            continue
        for member in sorted(
            name
            for name in dir(manager)
            if not name.startswith("_")
        ):
            lines.append("  " + member)

    builder = None
    try:
        builder, manager_name = _create_mass_properties_builder(
            work_part, [work_part]
        )
        lines.append("MassPropertiesBuilder via {0}:".format(manager_name))
        for member in sorted(
            name
            for name in dir(builder)
            if not name.startswith("_")
        ):
            lines.append("  " + member)
        options = getattr(builder, "UpdateOptions", None)
        if options is not None:
            lines.append(
                "UpdateOptions members: {0}".format(
                    [
                        name
                        for name in dir(options)
                        if not name.startswith("_")
                    ]
                )
            )
        else:
            lines.append("UpdateOptions: <unavailable>")
        if getattr(builder, "RollUp", None) is None:
            lines.append(
                "RollUp option: <absent on this build; roll-up is implicit "
                "when measuring the assembly root>"
            )
    except Exception as error:
        lines.append("PROBE FAILED: " + error_text(error))
    finally:
        dispose(builder)
    return lines


def _get_real_attribute(part, title):
    try:
        return float(part.GetRealAttribute(title))
    except Exception:
        return None


def read_rollup_attributes(part):
    """Read back the standard NX roll-up attributes (kg and mm^2)."""
    return {
        "mass": _get_real_attribute(part, ROLLUP_MASS_ATTRIBUTE),
        "area": _get_real_attribute(part, ROLLUP_AREA_ATTRIBUTE),
    }


def scan_rollup_mass(part):
    """Return a fail-closed tri-state scan of the reserved roll-up mass.

    An exception from ``GetRealAttribute`` alone cannot distinguish a missing
    attribute from an NX/API failure.  Enumerating attributes first makes
    ``MISSING`` an affirmative result: enumeration succeeded and the exact
    title was absent.  An existing numeric 0.0 is therefore ``PRESENT``.
    """
    iterator = None
    try:
        get_attributes = getattr(part, "GetUserAttributes", None)
        if not callable(get_attributes):
            raise RuntimeError("GetUserAttributes is unavailable.")

        create_iterator = getattr(part, "CreateAttributeIterator", None)
        if callable(create_iterator):
            iterator = create_iterator()
            try:
                attributes = list(get_attributes(iterator))
            except TypeError:
                attributes = list(get_attributes())
        else:
            attributes = list(get_attributes())

        matches = [
            info
            for info in attributes
            if clean(getattr(info, "Title", "")) == ROLLUP_MASS_ATTRIBUTE
        ]
        if not matches:
            return {
                "status": "MISSING",
                "value": None,
                "message": "Exact roll-up mass attribute title is absent.",
            }

        info = matches[0]
        if bool(getattr(info, "Unset", False)):
            return {
                "status": "READ_FAILED",
                "value": None,
                "message": "Roll-up mass attribute exists but is unset.",
            }
        value = getattr(info, "RealValue", None)
        if value is None:
            value = part.GetRealAttribute(ROLLUP_MASS_ATTRIBUTE)
        return {
            "status": "PRESENT",
            "value": float(value),
            "message": "Exact roll-up mass attribute exists.",
        }
    except Exception as error:
        return {
            "status": "READ_FAILED",
            "value": None,
            "message": "Roll-up mass scan failed: " + error_text(error),
        }
    finally:
        dispose(iterator)


def number_text(value, decimal_places):
    if value is None:
        return ""
    return ("{0:." + str(decimal_places) + "f}").format(value)


def _read_only(part):
    value = _safe_property(part, "IsReadOnly")
    return None if value is None else bool(value)


def _managed_mode(session, part):
    managed = _safe_property(session, "IsManagedMode", False)
    if bool(managed):
        return True
    identifier = clean(_safe_property(part, "FullPath"))
    return identifier.upper().startswith("@DB/")


def _pdm_part(part):
    return _safe_property(part, "PDMPart")


def _checkout_result(raw):
    checked = None
    owner = ""
    if isinstance(raw, dict):
        for key in ("isCheckedOut", "is_checked_out", "checkedOut", "checked_out"):
            if isinstance(raw.get(key), bool):
                checked = raw[key]
                break
        for key in ("checkedOutBy", "checked_out_by", "owner", "user"):
            if key in raw:
                owner = clean(raw[key])
                break
    elif isinstance(raw, (tuple, list)):
        for value in raw:
            if checked is None and isinstance(value, bool):
                checked = value
            elif isinstance(value, str) and not owner:
                owner = clean(value)
            elif checked is None:
                normalized = clean(
                    getattr(value, "name", value)
                ).upper().replace("_", "")
                if normalized.startswith("NOT") and "CHECKEDOUT" in normalized:
                    checked = False
                elif "CHECKEDOUT" in normalized:
                    checked = True
    elif isinstance(raw, bool):
        checked = raw
    else:
        for name in ("IsCheckedOut", "isCheckedOut", "CheckedOut"):
            value = getattr(raw, name, None)
            if isinstance(value, bool):
                checked = value
                break
        for name in ("CheckedOutBy", "checkedOutBy", "Owner", "User"):
            value = getattr(raw, name, None)
            if value is not None:
                owner = clean(value)
                break
    return (
        "CHECKED_OUT" if checked is True else
        "CHECKED_IN" if checked is False else
        "UNKNOWN",
        owner,
    )


def checkout_status(part):
    pdm_part = _pdm_part(part)
    method = getattr(pdm_part, "GetCheckedoutStatusAndUser", None)
    if not callable(method):
        return "UNKNOWN", "", "GetCheckedoutStatusAndUser unavailable"
    try:
        raw = method()
    except TypeError:
        try:
            raw = method(False, "")
        except Exception as error:
            return "UNKNOWN", "", error_text(error)
    except Exception as error:
        return "UNKNOWN", "", error_text(error)
    state, owner = _checkout_result(raw)
    return state, owner, clean(repr(raw))[:2000]


def current_teamcenter_user(session):
    pdm_session = _safe_property(session, "PdmSession")
    method = getattr(pdm_session, "GetUserName", None)
    if not callable(method):
        return ""
    try:
        return clean(method())
    except Exception:
        return ""


def inspect_write_access(session, part):
    """Read-only gate.  This function never checks a part out."""
    read_only = _read_only(part)
    current_user = current_teamcenter_user(session)
    if not _managed_mode(session, part):
        return {
            "allowed": read_only is not True,
            "checkout_state": "NATIVE",
            "checkout_owner": "",
            "current_user": current_user,
            "read_only": read_only,
            "message": "Part is read-only." if read_only is True else "",
        }

    state, owner, raw = checkout_status(part)
    owner_is_other = bool(
        state == "CHECKED_OUT"
        and owner
        and current_user
        and owner.casefold() != current_user.casefold()
    )
    messages = []
    if state == "CHECKED_IN":
        messages.append("Part is not checked out; J21 does not perform checkout.")
    elif owner_is_other:
        messages.append("Part is checked out by another user: {0}.".format(owner))
    elif state == "UNKNOWN":
        messages.append("Checkout state is unknown: {0}.".format(raw or "<none>"))
    if read_only is True:
        messages.append("Part is read-only in this NX session.")

    allowed = (
        state != "CHECKED_IN"
        and not owner_is_other
        and read_only is not True
    )
    return {
        "allowed": allowed,
        "checkout_state": state,
        "checkout_owner": owner,
        "current_user": current_user,
        "read_only": read_only,
        "message": " ".join(messages),
    }


def read_only_text(value):
    return "UNKNOWN" if value is None else "YES" if value else "NO"


def measurement_objects(part):
    try:
        root_component = getattr(
            getattr(part, "ComponentAssembly", None), "RootComponent", None
        )
    except Exception:
        root_component = None
    return [root_component] if root_component is not None else [part]


def part_kind(part):
    return "ASSEMBLY" if measurement_objects(part)[0] is not part else "LEAF"


def set_work_part(session, part):
    setter = getattr(getattr(session, "Parts", None), "SetWork", None)
    if not callable(setter):
        current = getattr(getattr(session, "Parts", None), "Work", None)
        if current is part:
            return
        raise RuntimeError("Session.Parts.SetWork is unavailable.")
    setter(part)


def _same_object(left, right):
    if left is right:
        return True
    if left is None or right is None:
        return False
    return _object_key(left) == _object_key(right)


def _set_display_part(session, part):
    setter = getattr(getattr(session, "Parts", None), "SetDisplay", None)
    if not callable(setter):
        raise RuntimeError("Session.Parts.SetDisplay is unavailable.")
    result = setter(part, False, True)
    if isinstance(result, (tuple, list)) and len(result) > 1:
        dispose(result[1])


def restore_part_context(session, original_display, original_work):
    messages = []
    parts = getattr(session, "Parts", None)
    current_display = getattr(parts, "Display", None)
    if original_display is not None and not _same_object(
        current_display, original_display
    ):
        try:
            _set_display_part(session, original_display)
        except Exception as error:
            messages.append("DISPLAY RESTORE: " + error_text(error))
    current_work = getattr(parts, "Work", None)
    if original_work is not None and not _same_object(current_work, original_work):
        try:
            set_work_part(session, original_work)
        except Exception as error:
            messages.append("WORK RESTORE: " + error_text(error))
    return messages


def save_part(part):
    status = None
    try:
        status = part.Save(
            NXOpen.BasePart.SaveComponents.FalseValue,
            NXOpen.BasePart.CloseAfterSave.FalseValue,
        )
        unsaved_parts = int(getattr(status, "NumberUnsavedParts", 0))
        unsaved_objects = int(getattr(status, "NumberUnsavedObjects", 0))
        if unsaved_parts or unsaved_objects:
            raise RuntimeError(
                "NX reported {0} unsaved part(s) and {1} unsaved "
                "object(s).".format(unsaved_parts, unsaved_objects)
            )
        return True, ""
    except Exception as error:
        return False, error_text(error)
    finally:
        dispose(status)


def empty_access():
    return {
        "checkout_state": "NOT_INSPECTED",
        "checkout_owner": "",
        "current_user": "",
        "read_only": None,
    }


def current_load_record(part, level, component_path, dry_run=False):
    status, raw_state = part_load_state(part)
    fully_loaded = status == "FULLY_LOADED"
    return {
        "part": part,
        "level": level,
        "component_path": component_path,
        "initial_load_state": raw_state or status,
        "load_action": (
            "NOT_REQUIRED" if fully_loaded else
            "WOULD_LOAD" if dry_run else
            "NOT_RUN"
        ),
        "final_load_state": raw_state or status,
        "load_status": (
            "SUCCESS" if fully_loaded else
            "LOAD_REQUIRED" if dry_run else
            status
        ),
        "load_message": (
            "Part is fully loaded." if fully_loaded else
            "APPLY would call LoadThisPartFully. Descendants may be hidden "
            "until this part is loaded." if dry_run else
            "Load was not run."
        ),
    }


def _result_row(
    timestamp,
    mode,
    part,
    level,
    component_path,
    process_order,
    access,
    load_record,
    update,
    saved,
    status,
    messages,
    stored_label="POPULATED",
    read_attributes=True,
    initial_mass=None,
    mass_scan_status="",
    selection="",
    selection_reason="",
    dependency_status="",
    require_area=True,
):
    identity = part_identity(part)
    attributes = (
        read_rollup_attributes(part)
        if read_attributes else
        {"mass": None, "area": None}
    )
    if mass_scan_status == "PRESENT" and attributes["mass"] is None:
        attributes["mass"] = initial_mass
    mass_status = (
        stored_label if attributes["mass"] is not None else
        "BLANK" if read_attributes else
        "NOT_READ"
    )
    area_status = (
        stored_label if attributes["area"] is not None else
        "BLANK" if read_attributes else
        "NOT_READ"
    )
    if (
        status == "SUCCESS"
        and (
            mass_status == "BLANK"
            or (require_area and area_status == "BLANK")
        )
    ):
        status = "PARTIAL"
    if update == "UPDATED":
        if mass_status == "BLANK":
            messages.append(
                "MASS ATTRIBUTE: NX did not write {0} for this part.".format(
                    ROLLUP_MASS_ATTRIBUTE
                )
            )
        if area_status == "BLANK":
            messages.append(
                "AREA ATTRIBUTE: NX did not write {0} for this part.".format(
                    ROLLUP_AREA_ATTRIBUTE
                )
            )

    return {
        "ROW_TYPE": "PART",
        "RUN_TIMESTAMP": timestamp,
        "JOURNAL_BUILD": BUILD,
        "WRITE_MODE": mode,
        "DB_PART_NO": identity["number"],
        "DB_PART_REV": identity["revision"],
        "PART_NAME": identity["name"],
        "COMPONENT_PATH": component_path,
        "LEVEL": level,
        "PROCESS_ORDER": process_order,
        "PART_KIND": part_kind(part),
        "INITIAL_LOAD_STATE": load_record["initial_load_state"],
        "LOAD_ACTION": load_record["load_action"],
        "FINAL_LOAD_STATE": load_record["final_load_state"],
        "LOAD_STATUS": load_record["load_status"],
        "LOAD_MESSAGE": load_record["load_message"],
        "LOAD_STATE": load_record["final_load_state"],
        "CHECKOUT_STATE": access["checkout_state"],
        "CHECKOUT_OWNER": access["checkout_owner"],
        "CURRENT_USER": access["current_user"],
        "READ_ONLY": read_only_text(access["read_only"]),
        "UPDATE": update,
        "INITIAL_ROLLUP_MASS_KG": number_text(
            initial_mass, MASS_DECIMAL_PLACES
        ),
        "MASS_SCAN_STATUS": mass_scan_status,
        "SELECTION": selection,
        "SELECTION_REASON": selection_reason,
        "DEPENDENCY_STATUS": dependency_status,
        "ROLLUP_MASS_KG": number_text(
            attributes["mass"], MASS_DECIMAL_PLACES
        ),
        "ROLLUP_AREA_MM2": number_text(
            attributes["area"], AREA_DECIMAL_PLACES
        ),
        "ROLLUP_AREA_M2": number_text(
            (
                attributes["area"] * SQUARE_METRES_PER_SQUARE_MILLIMETRE
                if attributes["area"] is not None
                else None
            ),
            AREA_M2_DECIMAL_PLACES,
        ),
        "ROLLUP_MASS_ATTRIBUTE": mass_status,
        "ROLLUP_AREA_ATTRIBUTE": area_status,
        "SAVED": saved,
        "STATUS": status,
        "MESSAGE": " | ".join(message for message in messages if message),
    }


def bottom_up_parts(parts):
    """Deepest prototypes first; stable discovery order breaks equal-depth ties."""
    return sorted(parts, key=lambda item: -item[1])


def diagnostic_row(timestamp, mode, diagnostic):
    dry_run = mode == "DRY_RUN"
    row = {column: "" for column in RESULT_COLUMNS}
    row.update(
        {
            "ROW_TYPE": "LOAD_DIAGNOSTIC",
            "RUN_TIMESTAMP": timestamp,
            "JOURNAL_BUILD": BUILD,
            "WRITE_MODE": mode,
            "COMPONENT_PATH": diagnostic.get("component_path", ""),
            "LEVEL": diagnostic.get("level", ""),
            "LOAD_ACTION": "NOT_RUN" if dry_run else "NOT_AVAILABLE",
            "LOAD_STATUS": diagnostic.get("code", "LOAD_FAILED"),
            "LOAD_MESSAGE": diagnostic.get("message", ""),
            "UPDATE": "DRY_RUN" if dry_run else "NOT_RUN_LOAD_FAILED",
            "SAVED": "DRY_RUN" if dry_run else "NOT_RUN_LOAD_FAILED",
            "STATUS": "DRY_RUN_WARNING" if dry_run else "LOAD_FAILED",
            "MESSAGE": diagnostic.get("message", ""),
        }
    )
    return row


def summary_row(timestamp, mode, rows, load_gate):
    part_rows = [row for row in rows if row.get("ROW_TYPE") == "PART"]
    updated = sum(row.get("UPDATE") == "UPDATED" for row in part_rows)
    present = sum(
        row.get("MASS_SCAN_STATUS") == "PRESENT" for row in part_rows
    )
    missing = sum(
        row.get("MASS_SCAN_STATUS") == "MISSING" for row in part_rows
    )
    selected = sum(
        row.get("SELECTION") == "SELECTED_MISSING_MASS" for row in part_rows
    )
    not_writable = sum(
        row.get("UPDATE") == "SKIPPED_NOT_WRITABLE" for row in part_rows
    )
    blocked = sum(
        row.get("UPDATE") == "BLOCKED_LOAD_DEPENDENCY" for row in part_rows
    )
    read_failed = sum(
        row.get("MASS_SCAN_STATUS") in (
            "READ_FAILED",
            "NOT_SCANNED_LOAD_FAILED",
        )
        for row in part_rows
    )
    update_failed = sum(
        row.get("UPDATE") == "UPDATE_FAILED" for row in part_rows
    )
    save_failed = sum(
        row.get("SAVED") == "SAVE_FAILED" for row in part_rows
    )
    partial_rows = sum(
        row.get("STATUS") not in ("SUCCESS", "DRY_RUN")
        for row in part_rows
        if row.get("UPDATE") != "SKIPPED_MASS_PRESENT"
    )
    row = {column: "" for column in RESULT_COLUMNS}
    if load_gate == "FAILED":
        status = "MASS_UPDATE_ABORTED"
        update = "NOT_RUN_LOAD_FAILED"
    elif mode == "DRY_RUN":
        status = "DRY_RUN"
        update = "DRY_RUN"
    elif load_gate == "PARTIAL" or partial_rows:
        status = "PARTIAL"
        update = "COMPLETED_WITH_EXCEPTIONS"
    else:
        status = "SUCCESS"
        update = "COMPLETED"
    row.update(
        {
            "ROW_TYPE": "RUN_SUMMARY",
            "RUN_TIMESTAMP": timestamp,
            "JOURNAL_BUILD": BUILD,
            "WRITE_MODE": mode,
            "LOAD_STATUS": load_gate,
            "UPDATE": update,
            "SAVED": "SUMMARY",
            "STATUS": status,
            "MESSAGE": (
                "discovered={0}; present={1}; missing={2}; selected={3}; "
                "updated={4}; not_writable={5}; blocked={6}; read_failed={7}; "
                "update_failed={8}; save_failed={9}".format(
                    len(part_rows),
                    present,
                    missing,
                    selected,
                    updated,
                    not_writable,
                    blocked,
                    read_failed,
                    update_failed,
                    save_failed,
                )
            ),
        }
    )
    return row


def build_dry_run_rows(session, timestamp, parts, diagnostics):
    rows = []
    for process_order, (part, level, component_path) in enumerate(
        bottom_up_parts(parts), start=1
    ):
        scan = scan_rollup_mass(part)
        if scan["status"] == "PRESENT":
            selection = "SKIPPED_MASS_PRESENT"
            selection_reason = "Existing roll-up mass would be trusted."
        elif scan["status"] == "MISSING":
            selection = "WOULD_UPDATE_MISSING_MASS"
            selection_reason = "Roll-up mass is absent."
        else:
            selection = "SKIPPED_READ_FAILED"
            selection_reason = scan["message"]
        access = inspect_write_access(session, part)
        messages = [access["message"]] if access["message"] else []
        if scan["status"] == "READ_FAILED":
            messages.append(scan["message"])
        load_record = current_load_record(
            part, level, component_path, dry_run=True
        )
        rows.append(
            _result_row(
                timestamp,
                "DRY_RUN",
                part,
                level,
                component_path,
                process_order,
                access,
                load_record,
                "DRY_RUN",
                "DRY_RUN",
                "DRY_RUN",
                messages,
                stored_label="STORED",
                initial_mass=scan["value"],
                mass_scan_status=scan["status"],
                selection=selection,
                selection_reason=selection_reason,
                dependency_status="NOT_EVALUATED_DRY_RUN",
                require_area=False,
            )
        )
    rows.extend(
        diagnostic_row(timestamp, "DRY_RUN", item)
        for item in diagnostics
    )
    rows.append(summary_row(timestamp, "DRY_RUN", rows, "NOT_RUN"))
    return rows


def build_load_aborted_rows(timestamp, mode, parts, load_records, diagnostics):
    rows = []
    targets = list(parts)
    target_keys = {_object_key(part) for part, _level, _path in targets}
    for key, record in load_records.items():
        if key not in target_keys:
            targets.append(
                (
                    record["part"],
                    record["level"],
                    record["component_path"],
                )
            )
    for process_order, (part, level, component_path) in enumerate(
        bottom_up_parts(targets), start=1
    ):
        record = load_records.get(
            _object_key(part),
            current_load_record(part, level, component_path),
        )
        rows.append(
            _result_row(
                timestamp,
                mode,
                part,
                level,
                component_path,
                process_order,
                empty_access(),
                record,
                "NOT_RUN_LOAD_FAILED",
                "NOT_RUN_LOAD_FAILED",
                "LOAD_FAILED",
                ["Mass update aborted because the full-load gate failed."],
                read_attributes=False,
                mass_scan_status="NOT_SCANNED_LOAD_FAILED",
                selection="NOT_RUN_LOAD_FAILED",
                selection_reason="Full refresh requires a complete load gate.",
                dependency_status="BLOCKED_BY_LOAD_FAILURE",
            )
        )
    rows.extend(
        diagnostic_row(timestamp, mode, item)
        for item in diagnostics
    )
    rows.append(summary_row(timestamp, mode, rows, "FAILED"))
    return rows


def dependency_blocked_keys(parts, load_records, diagnostics, dependencies):
    """Return failed prototypes plus every assembly prototype above them."""
    all_keys = {_object_key(part) for part, _level, _path in parts}
    blocked = {
        key
        for key, record in load_records.items()
        if record.get("load_status") != "SUCCESS"
    }
    for diagnostic in diagnostics:
        diagnostic_keys = diagnostic.get("blocked_keys", ())
        if diagnostic_keys:
            blocked.update(diagnostic_keys)
        elif diagnostic.get("code") in (
            "LOAD_PASS_LIMIT",
            "OCCURRENCE_LIMIT",
            "ROOT_COMPONENT_UNREADABLE",
        ):
            blocked.update(all_keys)

    parents = {}
    for parent_key, child_keys in dependencies.items():
        for child_key in child_keys:
            parents.setdefault(child_key, set()).add(parent_key)

    pending = list(blocked)
    while pending:
        child_key = pending.pop()
        for parent_key in parents.get(child_key, ()):
            if parent_key not in blocked:
                blocked.add(parent_key)
                pending.append(parent_key)
    return blocked


def apply_missing_mass_bottom_up(
    session,
    timestamp,
    parts,
    load_records,
    load_diagnostics,
    dependencies,
):
    """Scan all targets and mutate only safe prototypes missing roll-up mass."""
    rows = []
    diagnostics = []
    blocked_keys = dependency_blocked_keys(
        parts, load_records, load_diagnostics, dependencies
    )
    part_collection = getattr(session, "Parts", None)
    original_work = getattr(part_collection, "Work", None)
    original_display = getattr(part_collection, "Display", None)

    try:
        ordered = bottom_up_parts(parts)
        for process_order, (part, level, component_path) in enumerate(
            ordered, start=1
        ):
            key = _object_key(part)
            identity = part_identity(part)
            label = identity["number"] or identity["name"]
            load_record = load_records.get(
                key, current_load_record(part, level, component_path)
            )
            dependency_status = (
                "BLOCKED_BY_LOAD_FAILURE" if key in blocked_keys else "READY"
            )

            if load_record["load_status"] != "SUCCESS":
                messages = [
                    "Mass scan/update skipped because this prototype did not "
                    "fully load: " + load_record["load_message"]
                ]
                rows.append(
                    _result_row(
                        timestamp,
                        "APPLY",
                        part,
                        level,
                        component_path,
                        process_order,
                        empty_access(),
                        load_record,
                        "BLOCKED_LOAD_DEPENDENCY",
                        "NOT_SAVED",
                        "BLOCKED",
                        messages,
                        read_attributes=False,
                        mass_scan_status="NOT_SCANNED_LOAD_FAILED",
                        selection="BLOCKED_LOAD_DEPENDENCY",
                        selection_reason="Prototype did not fully load.",
                        dependency_status=dependency_status,
                        require_area=False,
                    )
                )
                continue

            scan = scan_rollup_mass(part)
            if scan["status"] == "PRESENT":
                log_line(session, "SKIP PRESENT {0}".format(label))
                rows.append(
                    _result_row(
                        timestamp,
                        "APPLY",
                        part,
                        level,
                        component_path,
                        process_order,
                        empty_access(),
                        load_record,
                        "SKIPPED_MASS_PRESENT",
                        "NOT_SAVED",
                        "SUCCESS",
                        [],
                        stored_label="EXISTING",
                        initial_mass=scan["value"],
                        mass_scan_status="PRESENT",
                        selection="SKIPPED_MASS_PRESENT",
                        selection_reason="Existing roll-up mass is trusted.",
                        dependency_status=dependency_status,
                        require_area=False,
                    )
                )
                continue

            if scan["status"] == "READ_FAILED":
                log_line(session, "SKIP READ FAILED {0}".format(label))
                rows.append(
                    _result_row(
                        timestamp,
                        "APPLY",
                        part,
                        level,
                        component_path,
                        process_order,
                        empty_access(),
                        load_record,
                        "SKIPPED_READ_FAILED",
                        "NOT_SAVED",
                        "READ_FAILED",
                        [scan["message"]],
                        read_attributes=False,
                        mass_scan_status="READ_FAILED",
                        selection="SKIPPED_READ_FAILED",
                        selection_reason=scan["message"],
                        dependency_status=dependency_status,
                        require_area=False,
                    )
                )
                continue

            if key in blocked_keys:
                log_line(session, "BLOCK DEPENDENCY {0}".format(label))
                rows.append(
                    _result_row(
                        timestamp,
                        "APPLY",
                        part,
                        level,
                        component_path,
                        process_order,
                        empty_access(),
                        load_record,
                        "BLOCKED_LOAD_DEPENDENCY",
                        "NOT_SAVED",
                        "BLOCKED",
                        [
                            "Missing mass was not updated because this assembly "
                            "depends on an unresolved or unloaded branch."
                        ],
                        mass_scan_status="MISSING",
                        selection="BLOCKED_LOAD_DEPENDENCY",
                        selection_reason="A required descendant branch failed.",
                        dependency_status=dependency_status,
                        require_area=False,
                    )
                )
                continue

            access = inspect_write_access(session, part)
            messages = [access["message"]] if access["message"] else []
            if not access["allowed"]:
                log_line(
                    session,
                    "SKIP {0}: {1}".format(
                        label, access["message"] or "target is not writable"
                    ),
                )
                rows.append(
                    _result_row(
                        timestamp,
                        "APPLY",
                        part,
                        level,
                        component_path,
                        process_order,
                        access,
                        load_record,
                        "SKIPPED_NOT_WRITABLE",
                        "NOT_SAVED",
                        "SKIPPED",
                        messages,
                        mass_scan_status="MISSING",
                        selection="SELECTED_MISSING_MASS",
                        selection_reason="Roll-up mass is absent.",
                        dependency_status="READY",
                        require_area=False,
                    )
                )
                continue

            log_line(
                session,
                "UPDATE MISSING {0} ({1}/{2}, level {3}, {4})".format(
                    label, process_order, len(ordered), level, part_kind(part)
                ),
            )
            try:
                set_work_part(session, part)
            except Exception as error:
                messages.append("SET WORK: " + error_text(error))
                rows.append(
                    _result_row(
                        timestamp,
                        "APPLY",
                        part,
                        level,
                        component_path,
                        process_order,
                        access,
                        load_record,
                        "UPDATE_FAILED",
                        "NOT_SAVED",
                        "UPDATE_FAILED",
                        messages,
                        mass_scan_status="MISSING",
                        selection="SELECTED_MISSING_MASS",
                        selection_reason="Roll-up mass is absent.",
                        dependency_status="READY",
                        require_area=False,
                    )
                )
                continue

            update_status = run_native_mass_property_update(
                part, objects=measurement_objects(part)
            )
            if not update_status.startswith("NATIVE_UPDATE_OK"):
                messages.append("UPDATE: " + update_status)
                rows.append(
                    _result_row(
                        timestamp,
                        "APPLY",
                        part,
                        level,
                        component_path,
                        process_order,
                        access,
                        load_record,
                        "UPDATE_FAILED",
                        "NOT_SAVED",
                        "UPDATE_FAILED",
                        messages,
                        mass_scan_status="MISSING",
                        selection="SELECTED_MISSING_MASS",
                        selection_reason="Roll-up mass is absent.",
                        dependency_status="READY",
                        require_area=False,
                    )
                )
                continue

            saved_ok, save_message = save_part(part)
            if not saved_ok:
                messages.append("SAVE: " + save_message)
            rows.append(
                _result_row(
                    timestamp,
                    "APPLY",
                    part,
                    level,
                    component_path,
                    process_order,
                    access,
                    load_record,
                    "UPDATED",
                    "SAVED" if saved_ok else "SAVE_FAILED",
                    "SUCCESS" if saved_ok else "SAVE_FAILED",
                    messages,
                    mass_scan_status="MISSING",
                    selection="SELECTED_MISSING_MASS",
                    selection_reason="Roll-up mass is absent.",
                    dependency_status="READY",
                    require_area=False,
                )
            )
    finally:
        for message in restore_part_context(
            session, original_display, original_work
        ):
            diagnostics.append(
                {"code": "CONTEXT_RESTORE_FAILED", "message": message}
            )
    return rows, diagnostics


def apply_parts_bottom_up(session, timestamp, mode, parts, load_records):
    rows = []
    diagnostics = []
    part_collection = getattr(session, "Parts", None)
    original_work = getattr(part_collection, "Work", None)
    original_display = getattr(part_collection, "Display", None)

    try:
        ordered = bottom_up_parts(parts)
        for process_order, (part, level, component_path) in enumerate(
            ordered, start=1
        ):
            identity = part_identity(part)
            label = identity["number"] or identity["name"]
            scan = scan_rollup_mass(part)
            forced_selection = (
                "REFRESH_ALL" if mode == "REFRESH_ALL" else "SMOKE_FORCED"
            )
            access = inspect_write_access(session, part)
            messages = [access["message"]] if access["message"] else []
            load_record = load_records.get(
                _object_key(part),
                current_load_record(part, level, component_path),
            )

            if not access["allowed"]:
                log_line(
                    session,
                    "SKIP {0}: {1}".format(
                        label, access["message"] or "target is not writable"
                    ),
                )
                rows.append(
                    _result_row(
                        timestamp,
                        mode,
                        part,
                        level,
                        component_path,
                        process_order,
                        access,
                        load_record,
                        "SKIPPED_NOT_WRITABLE",
                        "NOT_SAVED",
                        "SKIPPED",
                        messages,
                        initial_mass=scan["value"],
                        mass_scan_status=scan["status"],
                        selection=forced_selection,
                        selection_reason="Mode forces a native update.",
                        dependency_status="READY",
                    )
                )
                continue

            log_line(
                session,
                "UPDATE {0} ({1}/{2}, level {3}, {4})".format(
                    label, process_order, len(ordered), level, part_kind(part)
                ),
            )
            try:
                set_work_part(session, part)
            except Exception as error:
                messages.append("SET WORK: " + error_text(error))
                rows.append(
                    _result_row(
                        timestamp,
                        mode,
                        part,
                        level,
                        component_path,
                        process_order,
                        access,
                        load_record,
                        "UPDATE_FAILED",
                        "NOT_SAVED",
                        "UPDATE_FAILED",
                        messages,
                        initial_mass=scan["value"],
                        mass_scan_status=scan["status"],
                        selection=forced_selection,
                        selection_reason="Mode forces a native update.",
                        dependency_status="READY",
                    )
                )
                continue

            update_status = run_native_mass_property_update(
                part, objects=measurement_objects(part)
            )
            if not update_status.startswith("NATIVE_UPDATE_OK"):
                messages.append("UPDATE: " + update_status)
                rows.append(
                    _result_row(
                        timestamp,
                        mode,
                        part,
                        level,
                        component_path,
                        process_order,
                        access,
                        load_record,
                        "UPDATE_FAILED",
                        "NOT_SAVED",
                        "UPDATE_FAILED",
                        messages,
                        initial_mass=scan["value"],
                        mass_scan_status=scan["status"],
                        selection=forced_selection,
                        selection_reason="Mode forces a native update.",
                        dependency_status="READY",
                    )
                )
                continue

            saved_ok, save_message = save_part(part)
            if not saved_ok:
                messages.append("SAVE: " + save_message)
            rows.append(
                _result_row(
                    timestamp,
                    mode,
                    part,
                    level,
                    component_path,
                    process_order,
                    access,
                    load_record,
                    "UPDATED",
                    "SAVED" if saved_ok else "SAVE_FAILED",
                    "SUCCESS" if saved_ok else "SAVE_FAILED",
                    messages,
                    initial_mass=scan["value"],
                    mass_scan_status=scan["status"],
                    selection=forced_selection,
                    selection_reason="Mode forces a native update.",
                    dependency_status="READY",
                )
            )
    finally:
        for message in restore_part_context(
            session, original_display, original_work
        ):
            diagnostics.append({"code": "CONTEXT_RESTORE_FAILED", "message": message})
    return rows, diagnostics


def output_path(identity, file_timestamp):
    folder = os.path.join(io_root(), OUTPUT_FOLDER)
    os.makedirs(folder, exist_ok=True)
    token = (
        identity.get("number")
        or identity.get("name")
        or "UNKNOWN"
    )
    filename = "J21_MASS_SURFACE_{0}_{1}.csv".format(
        clean_filename_token(token),
        file_timestamp,
    )
    return os.path.join(folder, filename)


def write_csv(path, rows):
    with open(path, "w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=RESULT_COLUMNS,
            extrasaction="ignore",
        )
        writer.writeheader()
        writer.writerows(rows)


def run(session, run_datetime=None):
    now = run_datetime or datetime.datetime.now().astimezone()
    timestamp = now.isoformat(timespec="seconds")
    file_timestamp = now.strftime("%Y%m%d_%H%M%S")
    mode = clean(os.environ.get("NX_J21_MODE")) or WRITE_MODE
    if mode not in ("APPLY", "REFRESH_ALL", "DRY_RUN", "PROBE", "SMOKE"):
        raise RuntimeError(
            "NX_J21_MODE must be APPLY, REFRESH_ALL, DRY_RUN, PROBE, or "
            "SMOKE, got: {0}".format(mode)
        )

    try:
        work_part = session.Parts.Work
    except Exception:
        work_part = None

    if work_part is None:
        raise RuntimeError("Open an NX 3D master part or assembly first.")

    if mode == "PROBE":
        return None, probe_builder_api(work_part), []

    path = output_path(part_identity(work_part), file_timestamp)
    parts, traversal_diagnostics = collect_unique_parts(work_part)
    if mode == "APPLY":
        (
            load_ok,
            parts,
            load_records,
            diagnostics,
            dependencies,
        ) = auto_load_bom_visible(
            work_part,
            logger=lambda message: log_line(session, message),
        )
        rows, context_diagnostics = apply_missing_mass_bottom_up(
            session,
            timestamp,
            parts,
            load_records,
            diagnostics,
            dependencies,
        )
        diagnostics.extend(context_diagnostics)
        rows.extend(diagnostic_row(timestamp, mode, item) for item in diagnostics)
        rows.append(
            summary_row(
                timestamp,
                mode,
                rows,
                "SUCCESS" if load_ok and not diagnostics else "PARTIAL",
            )
        )
    elif mode == "REFRESH_ALL":
        (
            load_ok,
            parts,
            load_records,
            diagnostics,
            _dependencies,
        ) = auto_load_bom_visible(
            work_part,
            logger=lambda message: log_line(session, message),
        )
        if not load_ok:
            rows = build_load_aborted_rows(
                timestamp, mode, parts, load_records, diagnostics
            )
        else:
            rows, context_diagnostics = apply_parts_bottom_up(
                session, timestamp, mode, parts, load_records
            )
            diagnostics.extend(context_diagnostics)
            rows.append(summary_row(timestamp, mode, rows, "SUCCESS"))
    elif mode == "SMOKE":
        root_identity = part_identity(work_part)
        root_path = root_identity["number"] or root_identity["name"]
        smoke_parts = [(work_part, 0, root_path)]
        record = load_target(
            work_part,
            0,
            root_path,
            logger=lambda message: log_line(session, message),
        )
        load_records = {_object_key(work_part): record}
        if record["load_status"] != "SUCCESS":
            diagnostics = [
                {
                    "code": record["load_status"],
                    "message": record["load_message"],
                    "component_path": root_path,
                    "level": 0,
                }
            ]
            rows = build_load_aborted_rows(
                timestamp, mode, smoke_parts, load_records, diagnostics
            )
        else:
            rows, diagnostics = apply_parts_bottom_up(
                session, timestamp, mode, smoke_parts, load_records
            )
            rows.append(summary_row(timestamp, mode, rows, "SUCCESS"))
    else:
        rows = build_dry_run_rows(
            session, timestamp, parts, traversal_diagnostics
        )
        diagnostics = traversal_diagnostics
    write_csv(path, rows)
    return path, rows, diagnostics


def main():
    session = NXOpen.Session.GetSession()
    mode = clean(os.environ.get("NX_J21_MODE")) or WRITE_MODE
    log_line(session, "=" * 72)
    log_line(session, "J21 ASSEMBLY MASS & SURFACE AREA ATTRIBUTE UPDATER")
    log_line(session, "Build: " + BUILD)
    log_line(session, "Mode: " + mode)
    log_line(
        session,
        "Mechanism: NX native mass-properties update (Update On Save + Commit)",
    )
    log_line(
        session,
        "APPLY: auto-load and scan all BoM-visible targets; update/save only "
        "missing roll-up mass; checkout: inspect selected targets only",
    )
    log_line(
        session,
        "REFRESH_ALL: force the V5 full bottom-up rebuild with an "
        "all-or-nothing load gate",
    )
    log_line(
        session,
        "Attributes (standard NX, Rolled-Up Mass Properties): "
        "{0} (kg), {1} (mm^2)".format(
            ROLLUP_MASS_ATTRIBUTE, ROLLUP_AREA_ATTRIBUTE
        ),
    )
    log_line(session, "=" * 72)

    try:
        path, rows, diagnostics = run(session)
        if mode == "PROBE":
            for line in rows:
                log_line(session, line)
            log_line(
                session,
                "Send this probe output to confirm the exact NX 2506 "
                "MassPropertiesBuilder option names.",
            )
            return

        for row in rows:
            if row["ROW_TYPE"] == "RUN_SUMMARY":
                log_line(
                    session,
                    "RUN SUMMARY | load={0} | update={1} | status={2}".format(
                        row["LOAD_STATUS"], row["UPDATE"], row["STATUS"]
                    ),
                )
                if row["MESSAGE"]:
                    log_line(session, "    " + row["MESSAGE"])
                continue
            if row["ROW_TYPE"] == "LOAD_DIAGNOSTIC":
                log_line(
                    session,
                    "LOAD DIAGNOSTIC | {0} | {1} | {2}".format(
                        row["COMPONENT_PATH"] or "<path unavailable>",
                        row["LOAD_STATUS"],
                        row["LOAD_MESSAGE"],
                    ),
                )
                continue
            log_line(
                session,
                "{0} | level={1} | order={2} | {3} | load={4}/{5} | "
                "scan={6} | selection={7} | dependency={8} | checkout={9} "
                "({10}) | update={11} | rollup mass={12} kg [{13}] | "
                "rollup area={14} m^2 [{15}] | saved={16} | {17}".format(
                    row["DB_PART_NO"] or row["PART_NAME"],
                    row["LEVEL"],
                    row["PROCESS_ORDER"],
                    row["PART_KIND"],
                    row["LOAD_ACTION"],
                    row["LOAD_STATUS"],
                    row["MASS_SCAN_STATUS"],
                    row["SELECTION"],
                    row["DEPENDENCY_STATUS"],
                    row["CHECKOUT_STATE"],
                    row["CHECKOUT_OWNER"] or "<blank>",
                    row["UPDATE"],
                    row["ROLLUP_MASS_KG"] or "<blank>",
                    row["ROLLUP_MASS_ATTRIBUTE"],
                    row["ROLLUP_AREA_M2"] or "<blank>",
                    row["ROLLUP_AREA_ATTRIBUTE"],
                    row["SAVED"],
                    row["STATUS"],
                ),
            )
            if row["MESSAGE"]:
                log_line(session, "    " + row["MESSAGE"])
        log_line(
            session,
            "Parts reported: {0}".format(
                sum(row["ROW_TYPE"] == "PART" for row in rows)
            ),
        )
        if diagnostics:
            log_line(
                session,
                "Diagnostics: {0}".format(len(diagnostics)),
            )
            for item in diagnostics:
                log_line(
                    session,
                    "  {0}: {1}".format(item["code"], item["message"]),
                )
        log_line(session, "CSV: " + path)
    except Exception as error:
        log_line(session, "J21 FAILED: " + error_text(error))
        log_line(session, traceback.format_exc())
        raise


if __name__ == "__main__":
    main()
