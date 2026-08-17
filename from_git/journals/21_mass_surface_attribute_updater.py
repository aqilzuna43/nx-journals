"""Journal 21 - Assembly Mass & Surface Area Attribute Updater (NX 2506)

Drives NX's NATIVE Mass Properties update on the open assembly - the same
engine behind Tools > Measure Mass Properties with Update On Save.  NX itself
computes and writes its standard attributes on every component:

    NX_MassPropRollupMass  roll-up mass (kg)      [Rolled-Up Mass Properties]
    NX_MassPropRollupArea  roll-up area (mm^2)    [Rolled-Up Mass Properties]

The journal does NOT create, compute, or write attributes itself.  It:
  1. confirms the complete BoM-visible subtree is already fully loaded (same
     filter as NXOpenBoMExtended.py / Journal 04: suppressed, reference-only,
     and keyword-named occurrences are excluded),
  2. processes every unique prototype bottom-up: leaf parts first,
     subassemblies next, and the active assembly last,
  3. makes each target the work part and triggers NX's native mass-properties
     update on that target (PropertiesManager.CreateMassPropertiesBuilder
     with UpdateOnSave=Yes, then UpdateNow and Commit),
  4. saves every writable target and reads the reserved attributes back.

J21 never checks a part out.  In Teamcenter managed mode it reports checkout
state and owner, skips parts that are checked in, checked out by somebody
else, or read-only, and continues with everything it can update.  The
original display/work context is restored after the run.

WRITE_MODE defaults to "APPLY".  Set WRITE_MODE = "DRY_RUN" (or the
NX_J21_MODE environment variable) to report the current stored attribute
values and scope without updating or saving anything.  Set WRITE_MODE =
"SMOKE" to run the native update on the active work part only (fast
iteration to verify the mechanism before a full-assembly run).  Set
WRITE_MODE = "PROBE" to dump the PropertiesManager/MeasureManager and
builder API surface of this NX build to the Listing Window (useful once,
to confirm option names).

Why native and not direct write: NX_MassPropRollupMass / NX_MassPropRollupArea
are RESERVED NX attribute titles - a journal cannot write them with
AttributePropertiesBuilder (NX raises "This is a reserved attribute title.
[512006]").  Only NX's own mass-properties update can populate them, so this
journal only triggers that native update and reports what NX wrote.

Note: SMOKE measures only the active work part.  APPLY is the recursive,
bottom-up mode required to refresh both child parts and assembly roll-ups.

Target: NX X 2506 embedded Python only
Run via: NX > Tools > Journal > Play
"""

import csv
import datetime
import os
import traceback

import NXOpen


BUILD = "J21-NX2506-BOTTOM-UP-MASS-PROP-UPDATE-V4"
WRITE_MODE = "APPLY"  # APPLY / DRY_RUN / SMOKE / PROBE; NX_J21_MODE overrides
OUTPUT_FOLDER = "NX_MASS_SURFACE_UPDATE"
MEASUREMENT_ACCURACY = 0.99
MASS_DECIMAL_PLACES = 6
AREA_DECIMAL_PLACES = 2
AREA_M2_DECIMAL_PLACES = 4
# NX stores the roll-up area in square millimetres (PDM template); the report
# also presents it in square metres for readability on large systems.
SQUARE_METRES_PER_SQUARE_MILLIMETRE = 1e-6

# Standard NX roll-up attributes (category "Rolled-Up Mass Properties").
ROLLUP_MASS_ATTRIBUTE = "NX_MassPropRollupMass"
ROLLUP_AREA_ATTRIBUTE = "NX_MassPropRollupArea"

# --- BOM VISIBILITY (mirrors NXOpenBoMExtended.py and Journal 04) ---
IGNORE_KEYWORDS = ["CSYS", "COORDINATE", "DATUM", "REFERENCE", "SKELETON"]
BOM_REFERENCE_ATTRIBUTES = ("REFERENCE_COMPONENT", "PLIST_IGNORE_MEMBER")
# NX marks native reference components with an empty string; manual overrides
# use YES/1/True/true/yes.
BOM_REFERENCE_FLAG_VALUES = ("", "YES", "1", "True", "true", "yes")

RESULT_COLUMNS = (
    "RUN_TIMESTAMP",
    "JOURNAL_BUILD",
    "WRITE_MODE",
    "DB_PART_NO",
    "DB_PART_REV",
    "PART_NAME",
    "LEVEL",
    "PROCESS_ORDER",
    "PART_KIND",
    "LOAD_STATE",
    "CHECKOUT_STATE",
    "CHECKOUT_OWNER",
    "CURRENT_USER",
    "READ_ONLY",
    "UPDATE",
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
    return True


def _is_active_visible(component):
    """Suppression state; unreadable suppression is treated as not active."""
    try:
        return not bool(component.IsSuppressed)
    except Exception:
        return False


def collect_unique_parts(work_part):
    """Return BoM-visible unique 3D masters and traversal diagnostics.

    The work part is always included.  Each unique child prototype is included
    once, at its first-seen level.  Suppressed, reference-flagged, and
    keyword-named occurrences are excluded together with their subtrees.
    """
    unique = {}
    diagnostics = []

    def add_part(part, level):
        key = _object_key(part)
        if key not in unique:
            unique[key] = (part, level)
        elif level > unique[key][1]:
            # Shared prototypes are processed once, at their deepest observed
            # level, so sorting by descending level remains bottom-up.
            unique[key] = (part, level)

    add_part(work_part, 0)
    root_component = getattr(
        getattr(work_part, "ComponentAssembly", None), "RootComponent", None
    )
    if root_component is None:
        return list(unique.values()), diagnostics

    root_children, root_error = _children(root_component)
    if root_error:
        diagnostics.append(
            {
                "code": "CHILDREN_UNREADABLE",
                "message": "Assembly root children could not be read: " + root_error,
            }
        )
        return list(unique.values()), diagnostics

    stack = [(component, 1) for component in reversed(root_children)]
    while stack:
        component, level = stack.pop()
        if not _is_active_visible(component):
            continue
        if not _is_bom_visible(component):
            continue
        prototype = getattr(component, "Prototype", None)
        if prototype is None:
            diagnostics.append(
                {
                    "code": "MISSING_MODEL",
                    "message": (
                        "Component has no loaded prototype: {0}".format(
                            _component_name(component)
                        )
                    ),
                }
            )
        else:
            add_part(prototype, level)

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
                }
            )
        stack.extend(
            (child, level + 1)
            for child in reversed(children)
        )
    return list(unique.values()), diagnostics


def part_load_state(part):
    fully_loaded = _safe_property(part, "IsFullyLoaded")
    state = clean(_safe_property(part, "PartLoadState"))
    if fully_loaded is True:
        return "FULLY_LOADED", state
    if fully_loaded is False:
        return "NOT_FULLY_LOADED", state
    return "UNKNOWN", state


def full_load_preflight(parts, traversal_diagnostics):
    """Return every issue that must abort APPLY before the first mutation."""
    issues = list(traversal_diagnostics)
    for part, _level in parts:
        load_status, raw_state = part_load_state(part)
        if load_status == "FULLY_LOADED":
            continue
        identity = part_identity(part)
        label = identity["number"] or identity["name"]
        issues.append(
            {
                "code": load_status,
                "message": "{0}: IsFullyLoaded is {1}; PartLoadState={2}".format(
                    label,
                    "False" if load_status == "NOT_FULLY_LOADED" else "unavailable",
                    raw_state or "<unavailable>",
                ),
            }
        )
    return issues


def require_fully_loaded(parts, traversal_diagnostics):
    issues = full_load_preflight(parts, traversal_diagnostics)
    if not issues:
        return
    details = " | ".join(
        "{0}: {1}".format(item["code"], item["message"])
        for item in issues
    )
    raise RuntimeError(
        "FULL_LOAD_REQUIRED: no parts were changed. Fully load the complete "
        "BoM-visible subtree (J20 can diagnose load failures), then rerun J21. "
        + details
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
    root_component = getattr(
        getattr(part, "ComponentAssembly", None), "RootComponent", None
    )
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


def _result_row(
    timestamp,
    mode,
    part,
    level,
    process_order,
    access,
    update,
    saved,
    status,
    messages,
    stored_label="POPULATED",
):
    identity = part_identity(part)
    attributes = read_rollup_attributes(part)
    mass_status = stored_label if attributes["mass"] is not None else "BLANK"
    area_status = stored_label if attributes["area"] is not None else "BLANK"
    if status == "SUCCESS" and (mass_status == "BLANK" or area_status == "BLANK"):
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

    load_status, raw_load_state = part_load_state(part)
    return {
        "RUN_TIMESTAMP": timestamp,
        "JOURNAL_BUILD": BUILD,
        "WRITE_MODE": mode,
        "DB_PART_NO": identity["number"],
        "DB_PART_REV": identity["revision"],
        "PART_NAME": identity["name"],
        "LEVEL": level,
        "PROCESS_ORDER": process_order,
        "PART_KIND": part_kind(part),
        "LOAD_STATE": raw_load_state or load_status,
        "CHECKOUT_STATE": access["checkout_state"],
        "CHECKOUT_OWNER": access["checkout_owner"],
        "CURRENT_USER": access["current_user"],
        "READ_ONLY": read_only_text(access["read_only"]),
        "UPDATE": update,
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


def build_dry_run_rows(session, timestamp, parts):
    rows = []
    for process_order, (part, level) in enumerate(
        bottom_up_parts(parts), start=1
    ):
        access = inspect_write_access(session, part)
        messages = [access["message"]] if access["message"] else []
        rows.append(
            _result_row(
                timestamp,
                "DRY_RUN",
                part,
                level,
                process_order,
                access,
                "DRY_RUN",
                "DRY_RUN",
                "SUCCESS",
                messages,
                stored_label="STORED",
            )
        )
    return rows


def apply_parts_bottom_up(session, timestamp, mode, parts):
    rows = []
    diagnostics = []
    part_collection = getattr(session, "Parts", None)
    original_work = getattr(part_collection, "Work", None)
    original_display = getattr(part_collection, "Display", None)

    try:
        ordered = bottom_up_parts(parts)
        for process_order, (part, level) in enumerate(ordered, start=1):
            identity = part_identity(part)
            label = identity["number"] or identity["name"]
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
                        mode,
                        part,
                        level,
                        process_order,
                        access,
                        "SKIPPED_NOT_WRITABLE",
                        "NOT_SAVED",
                        "SKIPPED",
                        messages,
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
                        process_order,
                        access,
                        "UPDATE_FAILED",
                        "NOT_SAVED",
                        "UPDATE_FAILED",
                        messages,
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
                        process_order,
                        access,
                        "UPDATE_FAILED",
                        "NOT_SAVED",
                        "UPDATE_FAILED",
                        messages,
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
                    process_order,
                    access,
                    "UPDATED",
                    "SAVED" if saved_ok else "SAVE_FAILED",
                    "SUCCESS" if saved_ok else "SAVE_FAILED",
                    messages,
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
    if mode not in ("APPLY", "DRY_RUN", "PROBE", "SMOKE"):
        raise RuntimeError(
            "NX_J21_MODE must be APPLY, DRY_RUN, PROBE, or SMOKE, got: {0}".format(
                mode
            )
        )

    try:
        work_part = session.Parts.Work
    except Exception:
        work_part = None

    if work_part is None:
        raise RuntimeError("Open an NX 3D master part or assembly first.")

    if mode == "PROBE":
        return None, probe_builder_api(work_part), []

    parts, traversal_diagnostics = collect_unique_parts(work_part)
    if mode == "APPLY":
        # Fail before the first builder/checkout/save operation when any
        # BoM-visible branch is missing, unreadable, or not fully loaded.
        require_fully_loaded(parts, traversal_diagnostics)
        rows, diagnostics = apply_parts_bottom_up(
            session, timestamp, mode, parts
        )
    elif mode == "SMOKE":
        smoke_parts = [(work_part, 0)]
        require_fully_loaded(smoke_parts, [])
        rows, diagnostics = apply_parts_bottom_up(
            session, timestamp, mode, smoke_parts
        )
    else:
        rows = build_dry_run_rows(session, timestamp, parts)
        diagnostics = traversal_diagnostics
    path = output_path(part_identity(work_part), file_timestamp)
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
        "APPLY order: deepest leaf first, active assembly last; checkout: inspect only",
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
            log_line(
                session,
                "{0} | level={1} | order={2} | {3} | checkout={4} ({5}) | "
                "update={6} | rollup mass={7} kg [{8}] | rollup area={9} m^2 [{10}] | "
                "saved={11} | {12}".format(
                    row["DB_PART_NO"] or row["PART_NAME"],
                    row["LEVEL"],
                    row["PROCESS_ORDER"],
                    row["PART_KIND"],
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
            "Parts reported: {0}".format(len(rows)),
        )
        if diagnostics:
            log_line(
                session,
                "Traversal diagnostics: {0}".format(len(diagnostics)),
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
