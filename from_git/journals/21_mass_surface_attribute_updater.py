"""Journal 21 - Assembly Mass & Surface Area Attribute Updater (NX 2506)

Measures every direct traditional solid body of each unique BoM-visible 3D
master in the open assembly (the active work part plus all child prototypes)
and writes two Number attributes on every such part in one run, using the
standard NX roll-up attribute titles:

    NX_MassPropRollupMass  roll-up mass of the part including every
                           BoM-visible descendant (kg)
    NX_MassPropRollupArea  roll-up surface area of the part including every
                           BoM-visible descendant (square millimetres)

Values are computed by the journal with the classic, proven NX measure APIs
(NewFaceProperties for area, NewMassProperties for mass) and written with
AttributePropertiesBuilder - the same write path that was verified working.
The category defaults to the NX-native "Rolled-Up Mass Properties" and falls
back to "Materials" if the PDM template rejects the write, so the run always
reports exactly what happened per attribute.

Scope follows the BoM exactly (same filter as NXOpenBoMExtended.py and
Journal 04): suppressed occurrences, reference-only members
(REFERENCE_COMPONENT / PLIST_IGNORE_MEMBER), and keyword-named occurrences
(CSYS, COORDINATE, DATUM, REFERENCE, SKELETON) are excluded together with
their subtrees, so the journal never touches the same noise parts that the
BoM export hides.

WRITE_MODE defaults to "APPLY".  Set WRITE_MODE = "DRY_RUN" (or the
NX_J21_MODE environment variable) to compute and report without writing or
saving.  Set WRITE_MODE = "SMOKE" to run on the active work part only (fast
iteration).  Teamcenter parts must be writable (checked out) before APPLY;
otherwise the per-part save reports SAVE_FAILED.

Target: NX X 2506 embedded Python
Run via: NX > Tools > Journal > Play

========================================================================
J21 ASSEMBLY MASS & SURFACE AREA ATTRIBUTE UPDATER
Build: J21-NX2506-MASS-SURFACE-ATTRIBUTE-UPDATER-V4
Mode: APPLY
Mechanism: classic measure APIs + AttributePropertiesBuilder write
Attributes (standard titles): NX_MassPropRollupMass (kg), NX_MassPropRollupArea (mm^2) in Rolled-Up Mass Properties / Materials
========================================================================
264MN032797A01 | 0 | rollup mass=5.501863 kg [WRITE_FAILED] | rollup area=1.2622 m^2 [WRITE_FAILED] | saved=SAVED | PARTIAL
    AREA ATTRIBUTE: This is a reserved attribute title. [512006] | MASS ATTRIBUTE: This is a reserved attribute title. [512006] | VERIFY: NX_MassPropRollupMass not readable after write | VERIFY: NX_MassPropRollupArea not readable after write
264MN025224A01 | 1 | rollup mass=2.498485 kg [WRITE_FAILED] | rollup area=0.3143 m^2 [WRITE_FAILED] | saved=SAVED | PARTIAL
    AREA ATTRIBUTE: This is a reserved attribute title. [512006] | MASS ATTRIBUTE: This is a reserved attribute title. [512006] | VERIFY: NX_MassPropRollupMass not readable after write | VERIFY: NX_MassPropRollupArea not readable after write
264MN034376A01 | 2 | rollup mass=0.022828 kg [WRITE_FAILED] | rollup area=0.0020 m^2 [WRITE_FAILED] | saved=SAVED | PARTIAL
    AREA ATTRIBUTE: This is a reserved attribute title. [512006] | MASS ATTRIBUTE: This is a reserved attribute title. [512006] | VERIFY: NX_MassPropRollupMass not readable after write | VERIFY: NX_MassPropRollupArea not readable after write
264MN032800A01 | 2 | rollup mass=0.022351 kg [WRITE_FAILED] | rollup area=0.0055 m^2 [WRITE_FAILED] | saved=SAVED | PARTIAL
    AREA ATTRIBUTE: This is a reserved attribute title. [512006] | MASS ATTRIBUTE: This is a reserved attribute title. [512006] | VERIFY: NX_MassPropRollupMass not readable after write | VERIFY: NX_MassPropRollupArea not readable after write
264MN025225A01 | 2 | rollup mass=0.011399 kg [WRITE_FAILED] | rollup area=0.1315 m^2 [WRITE_FAILED] | saved=SAVED | PARTIAL
    AREA ATTRIBUTE: This is a reserved attribute title. [512006] | MASS ATTRIBUTE: This is a reserved attribute title. [512006] | VERIFY: NX_MassPropRollupMass not readable after write | VERIFY: NX_MassPropRollupArea not readable after write
264MN025226A01 | 2 | rollup mass=2.259759 kg [WRITE_FAILED] | rollup area=0.1555 m^2 [WRITE_FAILED] | saved=SAVED | PARTIAL
    AREA ATTRIBUTE: This is a reserved attribute title. [512006] | MASS ATTRIBUTE: This is a reserved attribute title. [512006] | VERIFY: NX_MassPropRollupMass not readable after write | VERIFY: NX_MassPropRollupArea not readable after write
264MN032756A01 | 3 | rollup mass=0.015913 kg [WRITE_FAILED] | rollup area=0.0012 m^2 [WRITE_FAILED] | saved=SAVED | PARTIAL
    AREA ATTRIBUTE: This is a reserved attribute title. [512006] | MASS ATTRIBUTE: This is a reserved attribute title. [512006] | VERIFY: NX_MassPropRollupMass not readable after write | VERIFY: NX_MassPropRollupArea not readable after write
264MN032237A01 | 3 | rollup mass=0.531714 kg [WRITE_FAILED] | rollup area=0.0166 m^2 [WRITE_FAILED] | saved=SAVED | PARTIAL
    AREA ATTRIBUTE: This is a reserved attribute title. [512006] | MASS ATTRIBUTE: This is a reserved attribute title. [512006] | VERIFY: NX_MassPropRollupMass not readable after write | VERIFY: NX_MassPropRollupArea not readable after write
264MN032635A01 | 4 | rollup mass=0.001192 kg [WRITE_FAILED] | rollup area=0.0006 m^2 [WRITE_FAILED] | saved=SAVED | PARTIAL
    AREA ATTRIBUTE: This is a reserved attribute title. [512006] | MASS ATTRIBUTE: This is a reserved attribute title. [512006] | VERIFY: NX_MassPropRollupMass not readable after write | VERIFY: NX_MassPropRollupArea not readable after write
264MN025162A01 | 4 | rollup mass=0.530521 kg [WRITE_FAILED] | rollup area=0.0160 m^2 [WRITE_FAILED] | saved=SAVED | PARTIAL
    AREA ATTRIBUTE: This is a reserved attribute title. [512006] | MASS ATTRIBUTE: This is a reserved attribute title. [512006] | VERIFY: NX_MassPropRollupMass not readable after write | VERIFY: NX_MassPropRollupArea not readable after write
264MN034348A01 | 5 | rollup mass=<blank> kg [NO_SOLIDS] | rollup area=<blank> m^2 [NO_SOLIDS] | saved=NOT_APPLICABLE | PARTIAL
    AREA: No direct traditional solid bodies to measure. | MASS: No direct traditional solid bodies in the roll-up scope.
264MN032801A01 | 3 | rollup mass=0.027436 kg [WRITE_FAILED] | rollup area=0.0021 m^2 [WRITE_FAILED] | saved=SAVED | PARTIAL
    AREA ATTRIBUTE: This is a reserved attribute title. [512006] | MASS ATTRIBUTE: This is a reserved attribute title. [512006] | VERIFY: NX_MassPropRollupMass not readable after write | VERIFY: NX_MassPropRollupArea not readable after write
264MN032789A01 | 3 | rollup mass=0.020953 kg [WRITE_FAILED] | rollup area=0.0017 m^2 [WRITE_FAILED] | saved=SAVED | PARTIAL
    AREA ATTRIBUTE: This is a reserved attribute title. [512006] | MASS ATTRIBUTE: This is a reserved attribute title. [512006] | VERIFY: NX_MassPropRollupMass not readable after write | VERIFY: NX_MassPropRollupArea not readable after write
264MN025210A01 | 1 | rollup mass=0.504893 kg [WRITE_FAILED] | rollup area=0.6337 m^2 [WRITE_FAILED] | saved=SAVED | PARTIAL
    AREA ATTRIBUTE: This is a reserved attribute title. [512006] | MASS ATTRIBUTE: This is a reserved attribute title. [512006] | VERIFY: NX_MassPropRollupMass not readable after write | VERIFY: NX_MassPropRollupArea not readable after write
264MN025211A01 | 2 | rollup mass=0.025120 kg [WRITE_FAILED] | rollup area=0.2896 m^2 [WRITE_FAILED] | saved=SAVED | PARTIAL
    AREA ATTRIBUTE: This is a reserved attribute title. [512006] | MASS ATTRIBUTE: This is a reserved attribute title. [512006] | VERIFY: NX_MassPropRollupMass not readable after write | VERIFY: NX_MassPropRollupArea not readable after write
Parts reported: 15
CSV: C:\Users\my62022696\Desktop\NX_MASS_SURFACE_UPDATE\J21_MASS_SURFACE_264MN032797A01_20260813_003419.csv

========================================================================
J22 MASS ATTRIBUTE WRITE DIAGNOSTIC
Build: J22-NX2506-MASS-ATTRIBUTE-WRITE-DIAGNOSTIC-V1
Scope: active work part only; writes test values.
========================================================================
Work part: 264MN032797A01 (solid bodies: 0)
[A_classic_compute]  -> NO_SOLID_BODIES
[B_native_builder]  -> OK
[C_direct_write] Rolled-Up Mass Properties -> FAILED | This is a reserved attribute title. [512006]
[C_direct_write] Materials -> FAILED | This is a reserved attribute title. [512006]
[C_direct_write] Rolled-Up Mass Properties -> FAILED | This is a reserved attribute title. [512006]
[C_direct_write] Materials -> FAILED | This is a reserved attribute title. [512006]
[save]  -> OK
CSV: C:\Users\my62022696\Desktop\NX_MASS_SURFACE_UPDATE\J22_DIAGNOSTIC_264MN032797A01_20260813_003737.csv
JSON: C:\Users\my62022696\Desktop\NX_MASS_SURFACE_UPDATE\J22_DIAGNOSTIC_264MN032797A01_20260813_003737.json
Send the JSON (or Listing output) to confirm which write mechanism and category work on this NX build.
"""

import csv
import datetime
import os
import re
import traceback

import NXOpen


BUILD = "J21-NX2506-MASS-SURFACE-ATTRIBUTE-UPDATER-V4"
WRITE_MODE = "APPLY"  # "APPLY", "DRY_RUN", or "SMOKE"; NX_J21_MODE overrides
OUTPUT_FOLDER = "NX_MASS_SURFACE_UPDATE"
MEASUREMENT_ACCURACY = 0.99
AREA_DECIMAL_PLACES = 2   # mm^2 column
AREA_M2_DECIMAL_PLACES = 4
MASS_DECIMAL_PLACES = 6

# Standard NX roll-up attribute titles.  NX defines these under
# "Rolled-Up Mass Properties"; "Materials" is the fallback that was verified
# writable on NX 2506.
ROLLUP_MASS_ATTRIBUTE = "NX_MassPropRollupMass"
ROLLUP_AREA_ATTRIBUTE = "NX_MassPropRollupArea"
ATTRIBUTE_CATEGORIES = ("Rolled-Up Mass Properties", "Materials")
# NX defines NX_MassPropRollupArea in square millimetres; the journal
# measures in square metres and converts at write time.
SQUARE_MILLIMETRES_PER_SQUARE_METRE = 1_000_000.0

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
    "OWN_SOLID_BODY_COUNT",
    "ROLLUP_SOLID_BODY_COUNT",
    "ROLLUP_AREA_MM2",
    "ROLLUP_AREA_M2",
    "ROLLUP_MASS_KG",
    "ROLLUP_AREA_ATTRIBUTE",
    "ROLLUP_MASS_ATTRIBUTE",
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


def body_flag(body, property_name):
    value = getattr(body, property_name)
    return bool(value() if callable(value) else value)


def classify_bodies(part):
    """Return direct traditional solid bodies plus skipped-body counts."""
    included = []
    skipped_sheet = 0
    skipped_convergent = 0

    for body in list(getattr(part, "Bodies", [])):
        if body_flag(body, "IsConvergentBody"):
            skipped_convergent += 1
        elif body_flag(body, "IsSolidBody"):
            included.append(body)
        elif body_flag(body, "IsSheetBody"):
            skipped_sheet += 1

    return {
        "included": included,
        "skipped_sheet": skipped_sheet,
        "skipped_convergent": skipped_convergent,
    }


def normalized_unit_token(value):
    text = clean(value).upper()
    text = text.replace("²", "2").replace("^", "")
    return re.sub(r"[^A-Z0-9]", "", text)


def unit_tokens(unit):
    values = []
    for property_name in ("Name", "Symbol", "Abbreviation", "TypeName"):
        try:
            values.append(normalized_unit_token(getattr(unit, property_name)))
        except Exception:
            pass
    return {value for value in values if value}


def unit_matches(unit, wanted_tokens):
    return bool(unit_tokens(unit).intersection(wanted_tokens))


def resolve_measure_unit(unit_collection, measure_name, object_names, tokens):
    for object_name in object_names:
        try:
            unit = unit_collection.FindObject(object_name)
            if unit is not None:
                return unit
        except Exception:
            pass

    try:
        candidates = list(unit_collection.GetMeasureTypes(measure_name))
    except Exception as error:
        raise RuntimeError(
            "NX could not enumerate {0} units: {1}".format(
                measure_name, error_text(error)
            )
        )

    for unit in candidates:
        if unit_matches(unit, tokens):
            return unit

    available = sorted(
        {
            token
            for unit in candidates
            for token in unit_tokens(unit)
        }
    )
    raise RuntimeError(
        "NX {0} units are unavailable for {1}. Available unit tokens: {2}".format(
            measure_name, measure_name, ", ".join(available) or "<none>"
        )
    )


def resolve_units(work_part):
    """Resolve square-metre area, metre length, and kilogram mass units once."""
    units = work_part.UnitCollection
    area_unit = resolve_measure_unit(
        units,
        "Area",
        ("SquareMeter", "SquareMetre"),
        {"SQUAREMETER", "SQUAREMETRE", "M2"},
    )
    length_unit = resolve_measure_unit(
        units,
        "Length",
        ("Meter", "Metre"),
        {"METER", "METRE", "M"},
    )
    mass_unit = resolve_measure_unit(
        units,
        "Mass",
        ("Kilogram",),
        {"KILOGRAM", "KG"},
    )
    return area_unit, length_unit, mass_unit


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
    """Return BoM-visible unique 3D masters and traversal diagnostics."""
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


def rollup_bodies(part, cache):
    """All BoM-visible solid bodies owned by the part or its descendants."""
    key = _object_key(part)
    if key in cache:
        return cache[key]
    total = list(classify_bodies(part)["included"])
    root_component = getattr(
        getattr(part, "ComponentAssembly", None), "RootComponent", None
    )
    if root_component is not None:
        for child in _children(root_component):
            if not _is_active_visible(child):
                continue
            if not _is_bom_visible(child):
                continue
            prototype = getattr(child, "Prototype", None)
            if prototype is not None:
                total.extend(rollup_bodies(prototype, cache))
    cache[key] = total
    return total


def measure_surface_area_m2(measure_manager, area_unit, length_unit, bodies):
    """Sum of NewFaceProperties areas; fail-closed with a message."""
    if not bodies:
        return None, "No direct traditional solid bodies to measure."
    total = 0.0
    failures = []
    for body in bodies:
        faces = list(body.GetFaces())
        if not faces:
            failures.append("{0}: no measurable faces".format(body.Name))
            continue
        measurement = None
        try:
            measurement = measure_manager.NewFaceProperties(
                area_unit,
                length_unit,
                MEASUREMENT_ACCURACY,
                faces,
            )
            area = float(measurement.Area)
            if area < 0.0:
                raise RuntimeError(
                    "NX returned a negative surface area: {0}".format(area)
                )
            total += area
        except Exception as error:
            failures.append("{0}: {1}".format(body.Name, error_text(error)))
        finally:
            dispose(measurement)
    if failures:
        return None, " | ".join(failures)
    return total, ""


def measure_rollup_mass_kg(measure_manager, mass_unit, bodies):
    """NewMassProperties over the whole roll-up body set; fail-closed."""
    if not bodies:
        return None, "No direct traditional solid bodies in the roll-up scope."
    measurement = None
    try:
        measurement = measure_manager.NewMassProperties(
            [mass_unit],
            MEASUREMENT_ACCURACY,
            bodies,
        )
        mass = float(measurement.Mass)
        if mass < 0.0:
            raise RuntimeError(
                "NX returned a negative roll-up mass: {0}".format(mass)
            )
        return mass, ""
    except Exception as error:
        return None, error_text(error)
    finally:
        dispose(measurement)


def number_text(value, decimal_places):
    if value is None:
        return ""
    return ("{0:." + str(decimal_places) + "f}").format(value)


def _builder_data_type():
    enum = NXOpen.AttributePropertiesBaseBuilder.DataTypeOptions
    return enum.Number


def write_number_attribute(session, part, title, value):
    """Write a Number attribute; returns (ok, category, message).

    Tries the NX-native category first, then the verified-writable fallback.
    """
    last_error = ""
    for category in ATTRIBUTE_CATEGORIES:
        builder = None
        try:
            builder = session.AttributeManager.CreateAttributePropertiesBuilder(
                part,
                [part],
                NXOpen.AttributePropertiesBuilder.OperationType.Save,
            )
            builder.Category = category
            builder.Title = title
            builder.DataType = _builder_data_type()
            builder.NumberValue = float(value)
            builder.Commit()
            return True, category, ""
        except Exception as error:
            last_error = error_text(error)
        finally:
            dispose(builder)
    return False, "", last_error


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


def _get_real_attribute(part, title):
    try:
        return float(part.GetRealAttribute(title))
    except Exception:
        return None


def build_result_rows(
    session,
    work_part,
    timestamp,
    mode,
    area_unit,
    length_unit,
    mass_unit,
    parts=None,
):
    if parts is None:
        parts, diagnostics = collect_unique_parts(work_part)
    else:
        diagnostics = []
    measure_manager = work_part.MeasureManager
    body_cache = {}
    rows = []

    for part, level in parts:
        identity = part_identity(part)
        own_bodies = classify_bodies(part)["included"]
        rollup = rollup_bodies(part, body_cache)
        area, area_message = measure_surface_area_m2(
            measure_manager, area_unit, length_unit, rollup
        )
        mass, mass_message = measure_rollup_mass_kg(
            measure_manager, mass_unit, rollup
        )

        issues = []
        notes = []
        if area_message:
            issues.append("AREA: " + area_message)
        if mass_message:
            issues.append("MASS: " + mass_message)

        area_attr_status = ""
        mass_attr_status = ""
        if mode in ("APPLY", "SMOKE"):
            if area is not None:
                ok, category, write_message = write_number_attribute(
                    session,
                    part,
                    ROLLUP_AREA_ATTRIBUTE,
                    area * SQUARE_MILLIMETRES_PER_SQUARE_METRE,
                )
                area_attr_status = "WRITTEN" if ok else "WRITE_FAILED"
                if not ok:
                    issues.append(
                        "AREA ATTRIBUTE: " + write_message
                    )
                elif category != ATTRIBUTE_CATEGORIES[0]:
                    notes.append(
                        "AREA ATTRIBUTE: fallback category {0}".format(category)
                    )
            else:
                area_attr_status = "NO_SOLIDS" if not rollup else "FAILED"
            if mass is not None:
                ok, category, write_message = write_number_attribute(
                    session, part, ROLLUP_MASS_ATTRIBUTE, mass
                )
                mass_attr_status = "WRITTEN" if ok else "WRITE_FAILED"
                if not ok:
                    issues.append(
                        "MASS ATTRIBUTE: " + write_message
                    )
                elif category != ATTRIBUTE_CATEGORIES[0]:
                    notes.append(
                        "MASS ATTRIBUTE: fallback category {0}".format(category)
                    )
            else:
                mass_attr_status = "NO_SOLIDS" if not rollup else "FAILED"
        else:
            area_attr_status = "DRY_RUN" if area is not None else (
                "NO_SOLIDS" if not rollup else "FAILED"
            )
            mass_attr_status = "DRY_RUN" if mass is not None else (
                "NO_SOLIDS" if not rollup else "FAILED"
            )

        saved = ""
        if mode in ("APPLY", "SMOKE") and (area is not None or mass is not None):
            saved_ok, save_message = save_part(part)
            saved = "SAVED" if saved_ok else "SAVE_FAILED"
            if not saved_ok:
                issues.append("SAVE: " + save_message)
        elif mode in ("APPLY", "SMOKE"):
            saved = "NOT_APPLICABLE"
        else:
            saved = "DRY_RUN"

        # Verify by reading the standard titles back after the write+save.
        # Only meaningful when attributes were actually written.
        if mode in ("APPLY", "SMOKE"):
            read_mass = _get_real_attribute(part, ROLLUP_MASS_ATTRIBUTE)
            read_area = _get_real_attribute(part, ROLLUP_AREA_ATTRIBUTE)
            if read_mass is None and mass is not None:
                issues.append(
                    "VERIFY: {0} not readable after write".format(
                        ROLLUP_MASS_ATTRIBUTE
                    )
                )
            if read_area is None and area is not None:
                issues.append(
                    "VERIFY: {0} not readable after write".format(
                        ROLLUP_AREA_ATTRIBUTE
                    )
                )

        row_status = "SUCCESS"
        if "SAVE_FAILED" in saved:
            row_status = "SAVE_FAILED"
        elif "WRITE_FAILED" in area_attr_status or "WRITE_FAILED" in mass_attr_status:
            row_status = "PARTIAL"
        elif "FAILED" in area_attr_status or "FAILED" in mass_attr_status:
            row_status = "PARTIAL"
        elif issues:
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
                "OWN_SOLID_BODY_COUNT": len(own_bodies),
                "ROLLUP_SOLID_BODY_COUNT": len(rollup),
                "ROLLUP_AREA_MM2": number_text(
                    (
                        area * SQUARE_MILLIMETRES_PER_SQUARE_METRE
                        if area is not None
                        else None
                    ),
                    AREA_DECIMAL_PLACES,
                ),
                "ROLLUP_AREA_M2": number_text(
                    area, AREA_M2_DECIMAL_PLACES
                ),
                "ROLLUP_MASS_KG": number_text(
                    mass, MASS_DECIMAL_PLACES
                ),
                "ROLLUP_AREA_ATTRIBUTE": area_attr_status,
                "ROLLUP_MASS_ATTRIBUTE": mass_attr_status,
                "SAVED": saved,
                "STATUS": row_status,
                "MESSAGE": " | ".join(notes + issues),
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
    if mode not in ("APPLY", "DRY_RUN", "SMOKE"):
        raise RuntimeError(
            "NX_J21_MODE must be APPLY, DRY_RUN, or SMOKE, got: {0}".format(
                mode
            )
        )

    try:
        work_part = session.Parts.Work
    except Exception:
        work_part = None

    if work_part is None:
        raise RuntimeError("Open an NX 3D master part or assembly first.")

    area_unit, length_unit, mass_unit = resolve_units(work_part)
    if mode == "SMOKE":
        parts = [(work_part, 0)]
    else:
        parts = None
    rows, diagnostics = build_result_rows(
        session,
        work_part,
        timestamp,
        mode,
        area_unit,
        length_unit,
        mass_unit,
        parts=parts,
    )
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
        "Mechanism: classic measure APIs + AttributePropertiesBuilder write",
    )
    log_line(
        session,
        "Attributes (standard titles): {0} (kg), {1} (mm^2) in {2}".format(
            ROLLUP_MASS_ATTRIBUTE,
            ROLLUP_AREA_ATTRIBUTE,
            " / ".join(ATTRIBUTE_CATEGORIES),
        ),
    )
    log_line(session, "=" * 72)

    try:
        path, rows, diagnostics = run(session)
        for row in rows:
            log_line(
                session,
                "{0} | {1} | rollup mass={2} kg [{3}] | rollup area={4} m^2 [{5}] | "
                "saved={6} | {7}".format(
                    row["DB_PART_NO"] or row["PART_NAME"],
                    row["LEVEL"],
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
        log_line(session, "Parts reported: {0}".format(len(rows)))
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
