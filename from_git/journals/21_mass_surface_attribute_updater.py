"""Journal 21 - Assembly Mass & Surface Area Attribute Updater (NX 2506)

Drives NX's NATIVE Mass Properties update on the open assembly - the same
engine behind Tools > Measure Mass Properties with Roll Up + Update On Save.
NX itself computes and writes its standard attributes on every component:

    NX_MassPropRollupMass  roll-up mass (kg)      [Rolled-Up Mass Properties]
    NX_MassPropRollupArea  roll-up area (mm^2)    [Rolled-Up Mass Properties]

The journal does NOT create, compute, or write attributes itself.  It:
  1. confirms the BoM-visible 3D masters in the open assembly (same filter
     as NXOpenBoMExtended.py / Journal 04: suppressed, reference-only, and
     keyword-named occurrences are excluded),
  2. triggers the native mass-properties update once on the assembly,
  3. saves each BoM-visible part,
  4. reads the standard attributes back for the CSV/Listing report, so the
     run proves what NX wrote.  A blank read-back means the native update
     did not engage for that part and must be investigated.

WRITE_MODE defaults to "APPLY".  Set WRITE_MODE = "DRY_RUN" (or the
NX_J21_MODE environment variable) to report the current stored attribute
values and scope without updating or saving anything.  Set WRITE_MODE =
"PROBE" to dump the MassPropertiesBuilder API surface of this NX build to
the Listing Window (useful once, to confirm option names).

Target: NX X 2506 embedded Python only
Run via: NX > Tools > Journal > Play

========================================================================
J21 ASSEMBLY MASS & SURFACE AREA ATTRIBUTE UPDATER
Build: J21-NX2506-NATIVE-MASS-PROP-UPDATE-V3
Mode: PROBE
Mechanism: NX native mass-properties update (Roll Up + Update On Save)
Attributes (standard NX, Rolled-Up Mass Properties): NX_MassPropRollupMass (kg), NX_MassPropRollupArea (mm^2)
========================================================================
PROBE FAILED: 'NXOpen.MeasureManager' object has no attribute 'CreateMassPropertiesBuilder'
Send this probe output to confirm the exact NX 2506 MassPropertiesBuilder option names.
"""

import csv
import datetime
import os
import traceback

import NXOpen


BUILD = "J21-NX2506-NATIVE-MASS-PROP-UPDATE-V3"
WRITE_MODE = "PROBE"  # "APPLY", "DRY_RUN", or "PROBE"; NX_J21_MODE overrides
OUTPUT_FOLDER = "NX_MASS_SURFACE_UPDATE"
MEASUREMENT_ACCURACY = 0.99
MASS_DECIMAL_PLACES = 6
AREA_DECIMAL_PLACES = 2

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
    "ROLLUP_MASS_KG",
    "ROLLUP_AREA_MM2",
    "ROLLUP_MASS_ATTRIBUTE",
    "ROLLUP_AREA_ATTRIBUTE",
    "SAVED",
    "STATUS",
    "MESSAGE",
)

_INVALID_FILENAME_CHARS = '<>:"/\\|?*'


def _resolve_update_on_save_yes():
    """Resolve the UpdateOnSave=Yes enum member at import time."""
    for base in ("NXOpen.Measure", "NXOpen"):
        try:
            current = __import__("NXOpen", fromlist=["*"])
            for part in base.split(".")[1:] + [
                "MassPropertiesBuilder",
                "UpdateOptions",
            ]:
                current = getattr(current, part)
            return getattr(current, "Yes")
        except Exception:
            continue
    return None


UPDATE_ON_SAVE_YES = _resolve_update_on_save_yes()


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


def _children(component):
    try:
        return list(component.GetChildren())
    except Exception:
        return []


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

    add_part(work_part, 0)
    root_component = getattr(
        getattr(work_part, "ComponentAssembly", None), "RootComponent", None
    )
    if root_component is None:
        return list(unique.values()), diagnostics

    stack = [(component, 1) for component in reversed(_children(root_component))]
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
                            clean(getattr(component, "DisplayName", ""))
                            or clean(getattr(component, "Name", ""))
                            or "<unknown>"
                        )
                    ),
                }
            )
        else:
            add_part(prototype, level)
        stack.extend(
            (child, level + 1)
            for child in reversed(_children(component))
        )
    return list(unique.values()), diagnostics


def _nx_enum(enum_path, member_name):
    """Resolve an NXOpen enum member defensively across package layouts."""
    for base in ("NXOpen.Measure", "NXOpen"):
        try:
            current = __import__("NXOpen", fromlist=["*"])
            for part in base.split(".")[1:] + enum_path.split("."):
                current = getattr(current, part)
            return getattr(current, member_name)
        except Exception:
            continue
    raise RuntimeError(
        "NX enum {0}.{1} is unavailable on this build.".format(
            enum_path, member_name
        )
    )


def run_native_mass_property_update(work_part):
    """Trigger NX's native roll-up mass property update on the assembly.

    NX itself computes and writes NX_MassPropRollupMass / NX_MassPropRollupArea
    (and the rest of the standard family) on every component.  Returns a
    status message; raises nothing unless the update cannot be started.
    """
    warnings = []
    try:
        root_component = getattr(
            getattr(work_part, "ComponentAssembly", None),
            "RootComponent",
            None,
        )
        objects = [root_component] if root_component is not None else [work_part]
        manager = work_part.MeasureManager
        builder = manager.CreateMassPropertiesBuilder(objects)

        builder.Accuracy = MEASUREMENT_ACCURACY
        # NX X 2506 builder options; skipped defensively if a name differs.
        if getattr(builder, "RollUp", None) is not None:
            builder.RollUp = True
        else:
            warnings.append("RollUp option unavailable")
        if UPDATE_ON_SAVE_YES is not None and getattr(
            builder, "UpdateOnSave", None
        ) is not None:
            builder.UpdateOnSave = UPDATE_ON_SAVE_YES
        else:
            warnings.append("UpdateOnSave option unavailable")
        update_now = getattr(builder, "UpdateNow", None)
        if update_now is None:
            raise RuntimeError(
                "MassPropertiesBuilder.UpdateNow is unavailable on this build."
            )
        update_now()
        if warnings:
            return "NATIVE_UPDATE_OK (skipped: {0})".format(
                "; ".join(warnings)
            )
        return "NATIVE_UPDATE_OK"
    except Exception as error:
        return "NATIVE_UPDATE_FAILED: " + error_text(error)
    finally:
        dispose(builder)


def probe_builder_api(work_part):
    """Dump the MassPropertiesBuilder API surface of this NX build."""
    builder = None
    lines = []
    try:
        builder = work_part.MeasureManager.CreateMassPropertiesBuilder(
            [work_part]
        )
        lines.append("MassPropertiesBuilder members:")
        for member in sorted(
            name
            for name in dir(builder)
            if not name.startswith("_")
        ):
            lines.append("  " + member)
        for path in (
            "MassPropertiesBuilder.UpdateOptions",
            "MassPropertiesBuilder.MeasurementType",
        ):
            try:
                enum_type = _nx_enum(path, "__members__")
                lines.append(
                    "{0} = {1}".format(
                        path,
                        [member for member in enum_type],
                    )
                )
            except Exception as error:
                lines.append(
                    "{0}: unavailable ({1})".format(path, error_text(error))
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


def build_result_rows(work_part, timestamp, mode):
    parts, diagnostics = collect_unique_parts(work_part)
    rows = []

    for part, level in parts:
        identity = part_identity(part)
        messages = []

        if mode == "APPLY":
            attributes = read_rollup_attributes(part)
            mass_status = (
                "POPULATED" if attributes["mass"] is not None else "BLANK"
            )
            area_status = (
                "POPULATED" if attributes["area"] is not None else "BLANK"
            )
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
            saved_ok, save_message = save_part(part)
            saved = "SAVED" if saved_ok else "SAVE_FAILED"
            if not saved_ok:
                messages.append("SAVE: " + save_message)
        else:
            # DRY_RUN: report the currently stored attributes without updating.
            attributes = read_rollup_attributes(part)
            mass_status = (
                "STORED" if attributes["mass"] is not None else "BLANK"
            )
            area_status = (
                "STORED" if attributes["area"] is not None else "BLANK"
            )
            saved = "DRY_RUN"

        row_status = "SUCCESS"
        if "SAVE_FAILED" in saved:
            row_status = "SAVE_FAILED"
        elif "BLANK" in mass_status or "BLANK" in area_status:
            row_status = "PARTIAL"
        elif messages:
            row_status = "PARTIAL"

        rows.append(
            {
                "RUN_TIMESTAMP": timestamp,
                "JOURNAL_BUILD": BUILD,
                "WRITE_MODE": mode,
                "DB_PART_NO": identity["number"],
                "DB_PART_REV": identity["revision"],
                "PART_NAME": identity["name"],
                "LEVEL": level,
                "ROLLUP_MASS_KG": number_text(
                    attributes["mass"], MASS_DECIMAL_PLACES
                ),
                "ROLLUP_AREA_MM2": number_text(
                    attributes["area"], AREA_DECIMAL_PLACES
                ),
                "ROLLUP_MASS_ATTRIBUTE": mass_status,
                "ROLLUP_AREA_ATTRIBUTE": area_status,
                "SAVED": saved,
                "STATUS": row_status,
                "MESSAGE": " | ".join(messages),
            }
        )

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
    if mode not in ("APPLY", "DRY_RUN", "PROBE"):
        raise RuntimeError(
            "NX_J21_MODE must be APPLY, DRY_RUN, or PROBE, got: {0}".format(
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

    if mode == "APPLY":
        update_status = run_native_mass_property_update(work_part)
        if not update_status.startswith("NATIVE_UPDATE_OK"):
            raise RuntimeError(update_status)

    rows, diagnostics = build_result_rows(work_part, timestamp, mode)
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
        "Mechanism: NX native mass-properties update (Roll Up + Update On Save)",
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
                "{0} | {1} | rollup mass={2} kg [{3}] | rollup area={4} mm^2 [{5}] | "
                "saved={6} | {7}".format(
                    row["DB_PART_NO"] or row["PART_NAME"],
                    row["LEVEL"],
                    row["ROLLUP_MASS_KG"] or "<blank>",
                    row["ROLLUP_MASS_ATTRIBUTE"],
                    row["ROLLUP_AREA_MM2"] or "<blank>",
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
