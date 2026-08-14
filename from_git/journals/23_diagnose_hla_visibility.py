"""J23 - Read-only HLA assembly visibility diagnostic.

Use this journal when geometry is visible after a component prototype is
opened in its own window, but the same component is missing from the top-level
HLA assembly window.  J23 inventories every occurrence and ranks assembly-only
visibility causes.  It does not repair or change the assembly.

Evidence includes occurrence/ancestor blanking, layer state, active-arrangement
suppression, non-geometric state, reference-set membership, representation and
load state, prototype geometry, mapped occurrence geometry, work-view visible
objects, and dynamic-section context.

Optional: set NX_J23_TARGET to a component name, part number, or assembly-path
substring before starting NX.  All occurrences are still captured; matching
rows are marked and printed first.

Target: NX 2312 and NX X 2506 embedded Python.
Run via: NX > Tools > Journal > Play
"""

import csv
import datetime
import json
import os
import traceback

import NXOpen


BUILD = "J23-NX2506-HLA-VISIBILITY-DIAGNOSTIC-V1"
OUTPUT_FOLDER = "NX_HLA_VISIBILITY_DIAGNOSTIC"
MAX_OCCURRENCES = 100000
MAX_MEMBER_PROBES_PER_OCCURRENCE = 500

CSV_COLUMNS = (
    "RUN_TIMESTAMP",
    "JOURNAL_BUILD",
    "ROOT_ASSEMBLY",
    "TARGET_MATCH",
    "LEVEL",
    "ASSEMBLY_PATH",
    "COMPONENT_NAME",
    "PARENT_COMPONENT",
    "PROTOTYPE_NAME",
    "PART_NUMBER",
    "REVISION",
    "REFERENCE_SET",
    "REFERENCE_SET_FOUND",
    "REFERENCE_SET_MEMBER_COUNT",
    "REFERENCE_SET_BODY_COUNT",
    "REFERENCE_SET_COMPONENT_COUNT",
    "SUPPRESSED",
    "SUPPRESSED_STATE",
    "SUPPRESSION_EXPRESSION",
    "SUPPRESSING_ARRANGEMENT",
    "ANCESTOR_SUPPRESSED",
    "IS_BLANKED",
    "ANCESTOR_BLANKED",
    "COMPONENT_LAYER",
    "COMPONENT_LAYER_STATE",
    "ANCESTOR_HIDDEN_LAYER",
    "NON_GEOMETRIC",
    "REPRESENTATION_MODE",
    "USED_ARRANGEMENT",
    "PROTOTYPE_LOAD_STATE",
    "PROTOTYPE_FULLY_LOADED",
    "PROTOTYPE_BODY_COUNT",
    "PROTOTYPE_SOLID_BODY_COUNT",
    "PROTOTYPE_CHILD_COMPONENT_COUNT",
    "PROTOTYPE_BLANKED_BODY_COUNT",
    "PROTOTYPE_HIDDEN_LAYER_BODY_COUNT",
    "GEOMETRY_MEMBERS_PROBED",
    "OCCURRENCE_MEMBERS_FOUND",
    "OCCURRENCE_MEMBERS_BLANKED",
    "OCCURRENCE_BODY_TAGS_TESTED_IN_WORK_VIEW",
    "OCCURRENCE_MEMBERS_VISIBLE_IN_WORK_VIEW",
    "PROBE_LIMIT_REACHED",
    "ISSUE_CODES",
    "ROOT_CAUSE",
    "CONFIDENCE",
    "RECOMMENDATION",
    "PROBE_ERRORS",
)

ENTIRE_REFSET_NAMES = ("ENTIRE PART", "ENTIRE_PART", "ENTIRE")
EMPTY_REFSET_NAMES = ("EMPTY", "EMPTY PART", "EMPTY_PART")
HIDDEN_LAYER_STATES = ("HIDDEN", "INVISIBLE")


def clean(value):
    if value is None:
        return ""
    try:
        return str(value).strip()
    except Exception:
        return ""


def enum_text(value):
    if value is None:
        return ""
    try:
        return clean(value.name)
    except Exception:
        return clean(value)


def yes_no(value):
    return "YES" if bool(value) else "NO"


def error_text(error):
    message = clean(error) or type(error).__name__
    code = clean(getattr(error, "ErrorCode", ""))
    if code:
        return "{0} [NX error {1}]".format(message, code)
    return message


def safe_value(value, property_name, fallback=""):
    if value is None:
        return fallback
    try:
        result = getattr(value, property_name)
        if callable(result):
            result = result()
        return result
    except Exception:
        return fallback


def safe_name(value, fallback="<unavailable>"):
    for property_name in (
        "DisplayName",
        "Name",
        "Leaf",
        "JournalIdentifier",
        "FullPath",
    ):
        result = clean(safe_value(value, property_name))
        if result:
            return result
    return fallback


def object_tag(value):
    try:
        tag = clean(value.Tag)
        return tag if tag and tag != "0" else ""
    except Exception:
        return ""


def same_nx_object(first, second):
    if first is second:
        return True
    first_tag = object_tag(first)
    second_tag = object_tag(second)
    return bool(first_tag and second_tag and first_tag == second_tag)


def object_kind(value):
    if value is None:
        return ""
    try:
        runtime_type = value.GetType()
        name = clean(getattr(runtime_type, "Name", ""))
        if name:
            return name
    except Exception:
        pass
    return clean(type(value).__name__)


def is_body(value):
    return "body" in object_kind(value).lower()


def is_component(value):
    kind = object_kind(value).lower()
    return "component" in kind and "assembly" not in kind


def bool_property(value, property_name):
    try:
        return bool(getattr(value, property_name)), ""
    except Exception as error:
        return False, "{0}: {1}".format(property_name, error_text(error))


def get_string_attribute(nx_object, names):
    for name in names:
        try:
            result = clean(nx_object.GetStringAttribute(name))
            if result:
                return result
        except Exception:
            pass
        try:
            info = nx_object.GetUserAttribute(
                name,
                NXOpen.NXObject.AttributeType.String,
                -1,
            )
            result = clean(getattr(info, "StringValue", ""))
            if result:
                return result
        except Exception:
            pass
    return ""


def log_line(session, message):
    text = clean(message)
    try:
        window = session.ListingWindow
        window.Open()
        writer = getattr(window, "WriteFullline", None)
        if not callable(writer):
            writer = getattr(window, "WriteLine", None)
        if callable(writer):
            for line in text.splitlines() or [""]:
                writer(line)
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
    return os.path.abspath(os.path.expanduser(configured or desktop_folder()))


def filename_token(value):
    text = clean(value) or "UNKNOWN"
    result = "".join(char if char.isalnum() or char in "-_" else "_" for char in text)
    return result[:80] or "UNKNOWN"


def get_layer_state(part, layer):
    try:
        number = int(layer)
        return enum_text(part.Layers.GetState(number)), ""
    except Exception as error:
        return "", "Layer {0}: {1}".format(clean(layer), error_text(error))


def collection_items(value):
    if value is None:
        return []
    try:
        return list(value)
    except Exception:
        pass
    try:
        return list(value.ToArray())
    except Exception:
        return []


def direct_bodies(prototype):
    try:
        return list(prototype.Bodies), ""
    except Exception as error:
        return [], "Prototype.Bodies: {0}".format(error_text(error))


def direct_child_count(prototype):
    try:
        root = prototype.ComponentAssembly.RootComponent
        if root is None:
            return 0, ""
        return len(list(root.GetChildren())), ""
    except Exception as error:
        return 0, "Prototype child components: {0}".format(error_text(error))


def reference_set_details(prototype, selected_name, bodies):
    """Return the selected reference-set members without modifying it."""
    result = {
        "found": False,
        "members": [],
        "body_count": 0,
        "component_count": 0,
        "errors": [],
    }
    normalized = clean(selected_name).upper()
    entire_name = clean(safe_value(prototype, "EntirePartRefsetName")).upper()
    empty_name = clean(safe_value(prototype, "EmptyPartRefsetName")).upper()

    if normalized in ENTIRE_REFSET_NAMES or (entire_name and normalized == entire_name):
        result["found"] = True
        result["members"] = list(bodies)
        result["body_count"] = len(bodies)
        return result
    if normalized in EMPTY_REFSET_NAMES or (empty_name and normalized == empty_name):
        result["found"] = True
        return result

    try:
        reference_sets = list(prototype.GetAllReferenceSets())
    except Exception as error:
        result["errors"].append(
            "Prototype.GetAllReferenceSets: {0}".format(error_text(error))
        )
        return result

    selected = None
    for reference_set in reference_sets:
        if safe_name(reference_set, "").upper() == normalized:
            selected = reference_set
            break
    if selected is None:
        return result

    result["found"] = True
    try:
        result["members"] = list(selected.AskMembersInReferenceSet())
    except Exception as error:
        result["errors"].append(
            "ReferenceSet.AskMembersInReferenceSet: {0}".format(error_text(error))
        )
        try:
            result["members"] = list(selected.AskAllDirectMembers())
        except Exception as fallback_error:
            result["errors"].append(
                "ReferenceSet.AskAllDirectMembers: {0}".format(
                    error_text(fallback_error)
                )
            )
    result["body_count"] = sum(1 for item in result["members"] if is_body(item))
    result["component_count"] = sum(
        1 for item in result["members"] if is_component(item)
    )
    return result


def expression_text(expression):
    if expression is None:
        return ""
    pieces = []
    for property_name in ("Name", "RightHandSide", "Equation", "Value"):
        value = clean(safe_value(expression, property_name))
        if value and value not in pieces:
            pieces.append(value)
    return " | ".join(pieces) or safe_name(expression, "")


def suppression_details(component_assembly, component):
    result = {"state": "", "expression": "", "error": ""}
    try:
        result["state"] = enum_text(
            component_assembly.GetSuppressedState(component, False)
        )
    except Exception as error:
        result["error"] = "GetSuppressedState: {0}".format(error_text(error))
    try:
        result["expression"] = expression_text(
            component_assembly.GetSuppressionExpression(component)
        )
    except Exception:
        # Most ordinary components have no controlling expression.
        pass
    return result


def work_view_snapshot(work_part):
    result = {
        "available": False,
        "name": "",
        "visible_object_count": 0,
        "visible_tags": set(),
        "error": "",
    }
    try:
        view = work_part.ModelingViews.WorkView
        result["name"] = safe_name(view, "<work view>")
        objects = list(view.AskVisibleObjects())
        result["available"] = True
        result["visible_object_count"] = len(objects)
        result["visible_tags"] = set(
            tag for tag in (object_tag(item) for item in objects) if tag
        )
    except Exception as error:
        result["error"] = "WorkView.AskVisibleObjects: {0}".format(error_text(error))
    return result


def dynamic_section_snapshot(work_part):
    result = {"defined_count": 0, "clip_enabled_count": 0, "sections": [], "errors": []}
    try:
        collection = work_part.DynamicSections
        sections = collection_items(collection)
    except Exception as error:
        result["errors"].append("DynamicSections: {0}".format(error_text(error)))
        return result
    result["defined_count"] = len(sections)
    try:
        view = work_part.ModelingViews.WorkView
    except Exception:
        view = None
    for section in sections:
        item = {"name": safe_name(section, "<dynamic section>"), "show_clip": "UNKNOWN"}
        builder = None
        try:
            builder = collection.CreateSectionBuilder(section, view)
            show_clip = bool(safe_value(builder, "ShowClip", False))
            item["show_clip"] = yes_no(show_clip)
            if show_clip:
                result["clip_enabled_count"] += 1
        except Exception as error:
            result["errors"].append(
                "Dynamic section {0}: {1}".format(item["name"], error_text(error))
            )
        finally:
            if builder is not None:
                try:
                    builder.Destroy()
                except Exception:
                    pass
        result["sections"].append(item)
    return result


def prototype_geometry(prototype):
    result = {
        "bodies": [],
        "body_count": 0,
        "solid_count": 0,
        "blanked_count": 0,
        "hidden_layer_count": 0,
        "child_count": 0,
        "errors": [],
    }
    bodies, error = direct_bodies(prototype)
    result["bodies"] = bodies
    result["body_count"] = len(bodies)
    if error:
        result["errors"].append(error)
    child_count, child_error = direct_child_count(prototype)
    result["child_count"] = child_count
    if child_error:
        result["errors"].append(child_error)

    for body in bodies:
        try:
            if bool(body.IsSolidBody):
                result["solid_count"] += 1
        except Exception:
            pass
        try:
            if bool(body.IsBlanked):
                result["blanked_count"] += 1
        except Exception as error:
            result["errors"].append("Body.IsBlanked: {0}".format(error_text(error)))
        layer = safe_value(body, "Layer", "")
        state, layer_error = get_layer_state(prototype, layer)
        if state.upper() in HIDDEN_LAYER_STATES:
            result["hidden_layer_count"] += 1
        if layer_error:
            result["errors"].append(layer_error)
    return result


def occurrence_geometry(component, members, view_snapshot):
    result = {
        "probed": 0,
        "found": 0,
        "blanked": 0,
        "view_tested": 0,
        "visible": 0,
        "limit_reached": False,
        "errors": [],
    }
    candidates = [item for item in members if is_body(item) or is_component(item)]
    if len(candidates) > MAX_MEMBER_PROBES_PER_OCCURRENCE:
        result["limit_reached"] = True
        candidates = candidates[:MAX_MEMBER_PROBES_PER_OCCURRENCE]
    for member in candidates:
        result["probed"] += 1
        try:
            occurrence = component.FindOccurrence(member)
        except Exception as error:
            result["errors"].append(
                "Component.FindOccurrence({0}): {1}".format(
                    object_kind(member), error_text(error)
                )
            )
            continue
        if occurrence is None:
            continue
        result["found"] += 1
        try:
            if bool(occurrence.IsBlanked):
                result["blanked"] += 1
        except Exception:
            pass
        tag = object_tag(occurrence)
        if view_snapshot["available"] and tag and is_body(member):
            result["view_tested"] += 1
            if tag in view_snapshot["visible_tags"]:
                result["visible"] += 1
    return result


CAUSE_GUIDANCE = {
    "ANCESTOR_SUPPRESSED": (
        "HIGH",
        "A parent occurrence is suppressed in the active assembly state.",
        "Inspect the first suppressed parent in the reported path and its active-arrangement suppression control.",
    ),
    "SUPPRESSED_CURRENT_ARRANGEMENT": (
        "HIGH",
        "The occurrence is suppressed in the active arrangement.",
        "Review the active arrangement and the reported suppression expression/arrangement; unsuppress only after confirming design intent.",
    ),
    "NON_GEOMETRIC_OCCURRENCE": (
        "HIGH",
        "The occurrence is marked non-geometric, so NX does not display model geometry at HLA level.",
        "In Assembly Navigator, review the component's non-geometric state and restore geometric status if this was unintended.",
    ),
    "ANCESTOR_BLANKED": (
        "HIGH",
        "A parent occurrence is blanked, which hides its complete subtree.",
        "Unblank/show the first blanked parent occurrence in the reported assembly path.",
    ),
    "COMPONENT_BLANKED": (
        "HIGH",
        "The component occurrence itself is blanked in the HLA display.",
        "Show the occurrence in the HLA window and verify no parent remains blanked.",
    ),
    "ANCESTOR_LAYER_HIDDEN": (
        "HIGH",
        "A parent occurrence is on a hidden HLA layer.",
        "Make the parent component layer visible/selectable in the top-level assembly layer settings.",
    ),
    "COMPONENT_LAYER_HIDDEN": (
        "HIGH",
        "The occurrence is placed on a hidden layer in the top-level assembly.",
        "Change the HLA layer state for the reported component layer; prototype-window layer settings are separate.",
    ),
    "EMPTY_REFERENCE_SET": (
        "HIGH",
        "The occurrence explicitly uses the Empty reference set.",
        "Assign the intended MODEL or Entire Part reference set to this occurrence.",
    ),
    "REFERENCE_SET_NOT_FOUND": (
        "HIGH",
        "The occurrence names a reference set that is not present on the resolved prototype.",
        "Correct the occurrence reference set or restore the same-named reference set on the exact loaded revision.",
    ),
    "REFERENCE_SET_HAS_NO_GEOMETRY": (
        "HIGH",
        "The selected reference set has no body/component members although the prototype contains geometry.",
        "Add the intended geometry to that reference set or use the correct populated reference set.",
    ),
    "PROTOTYPE_UNAVAILABLE": (
        "HIGH",
        "NX could not resolve a prototype object for this occurrence.",
        "Check Teamcenter access, revision rule, dataset availability, and assembly load/search options.",
    ),
    "NO_PROTOTYPE_GEOMETRY": (
        "HIGH",
        "The resolved prototype has neither direct bodies nor direct child components.",
        "Confirm that the occurrence resolves to the intended Item Revision and model dataset.",
    ),
    "NO_OCCURRENCE_GEOMETRY": (
        "HIGH",
        "Reference-set geometry exists in the prototype but NX returned no mapped HLA occurrences.",
        "This indicates stale/corrupt occurrence or representation data; replace/re-add only after reviewing this evidence and the exact revision.",
    ),
    "ALL_OCCURRENCE_GEOMETRY_BLANKED": (
        "HIGH",
        "All mapped HLA geometry members are blanked.",
        "Use Show in the HLA window on the mapped occurrence geometry and check for view-dependent hiding.",
    ),
    "NOT_VISIBLE_IN_WORK_VIEW": (
        "MEDIUM",
        "Mapped occurrence geometry is absent from the active work view's visible-object inventory.",
        "Clear any isolate/view-dependent hide state or test a new modeling work view, then rerun J23.",
    ),
    "ALL_PROTOTYPE_BODY_LAYERS_HIDDEN": (
        "MEDIUM",
        "Every direct prototype body is on a hidden prototype layer.",
        "Compare layer states between the standalone part window and the HLA display context.",
    ),
    "PROTOTYPE_BODIES_BLANKED": (
        "MEDIUM",
        "Every direct prototype body reports blanked.",
        "Check body-level blanking inside the prototype and reference set, then update the HLA display.",
    ),
    "PROTOTYPE_NOT_FULLY_LOADED": (
        "MEDIUM",
        "The component prototype does not report fully loaded.",
        "Review J20/load-status evidence for this exact prototype even if assembly-wide Full Load was already requested.",
    ),
    "LIGHTWEIGHT_OR_PARTIAL_REPRESENTATION": (
        "MEDIUM",
        "The occurrence uses a lightweight or partial representation.",
        "Display the component Exact and rerun J23; if geometry remains absent, use the higher-ranked evidence.",
    ),
    "ACTIVE_DYNAMIC_SECTION": (
        "LOW",
        "At least one dynamic-section definition reports clipping enabled in the active work view.",
        "Temporarily deactivate dynamic section clipping and rerun J23 to confirm or eliminate it.",
    ),
}


def diagnose_record(row, dynamic_sections):
    issues = []
    refset = clean(row["REFERENCE_SET"]).upper()
    suppressed_state = clean(row["SUPPRESSED_STATE"]).upper()
    representation = clean(row["REPRESENTATION_MODE"]).upper()

    if row["ANCESTOR_SUPPRESSED"] == "YES":
        issues.append("ANCESTOR_SUPPRESSED")
    if row["SUPPRESSED"] == "YES" or suppressed_state in (
        "SUPPRESSED",
        "SUPPRESSEDBYEXP",
    ):
        issues.append("SUPPRESSED_CURRENT_ARRANGEMENT")
    if row["NON_GEOMETRIC"] == "YES":
        issues.append("NON_GEOMETRIC_OCCURRENCE")
    if row["ANCESTOR_BLANKED"] == "YES":
        issues.append("ANCESTOR_BLANKED")
    if row["IS_BLANKED"] == "YES":
        issues.append("COMPONENT_BLANKED")
    if row["ANCESTOR_HIDDEN_LAYER"] == "YES":
        issues.append("ANCESTOR_LAYER_HIDDEN")
    if clean(row["COMPONENT_LAYER_STATE"]).upper() in HIDDEN_LAYER_STATES:
        issues.append("COMPONENT_LAYER_HIDDEN")
    if refset in EMPTY_REFSET_NAMES:
        issues.append("EMPTY_REFERENCE_SET")
    elif (
        row["REFERENCE_SET_FOUND"] == "NO"
        and row["PROTOTYPE_NAME"] != "<unavailable>"
    ):
        issues.append("REFERENCE_SET_NOT_FOUND")
    elif (
        int(row["REFERENCE_SET_BODY_COUNT"])
        + int(row["REFERENCE_SET_COMPONENT_COUNT"])
        == 0
        and int(row["PROTOTYPE_BODY_COUNT"]) + int(row["PROTOTYPE_CHILD_COMPONENT_COUNT"]) > 0
        and refset not in ENTIRE_REFSET_NAMES
    ):
        issues.append("REFERENCE_SET_HAS_NO_GEOMETRY")
    if row["PROTOTYPE_NAME"] == "<unavailable>":
        issues.append("PROTOTYPE_UNAVAILABLE")
    elif (
        int(row["PROTOTYPE_BODY_COUNT"]) == 0
        and int(row["PROTOTYPE_CHILD_COMPONENT_COUNT"]) == 0
    ):
        issues.append("NO_PROTOTYPE_GEOMETRY")
    if (
        int(row["GEOMETRY_MEMBERS_PROBED"]) > 0
        and int(row["OCCURRENCE_MEMBERS_FOUND"]) == 0
    ):
        issues.append("NO_OCCURRENCE_GEOMETRY")
    if (
        int(row["OCCURRENCE_MEMBERS_FOUND"]) > 0
        and int(row["OCCURRENCE_MEMBERS_FOUND"])
        == int(row["OCCURRENCE_MEMBERS_BLANKED"])
    ):
        issues.append("ALL_OCCURRENCE_GEOMETRY_BLANKED")
    if (
        int(row["OCCURRENCE_BODY_TAGS_TESTED_IN_WORK_VIEW"]) > 0
        and int(row["OCCURRENCE_MEMBERS_VISIBLE_IN_WORK_VIEW"]) == 0
    ):
        issues.append("NOT_VISIBLE_IN_WORK_VIEW")
    if (
        int(row["PROTOTYPE_BODY_COUNT"]) > 0
        and int(row["PROTOTYPE_BODY_COUNT"])
        == int(row["PROTOTYPE_HIDDEN_LAYER_BODY_COUNT"])
    ):
        issues.append("ALL_PROTOTYPE_BODY_LAYERS_HIDDEN")
    if (
        int(row["PROTOTYPE_BODY_COUNT"]) > 0
        and int(row["PROTOTYPE_BODY_COUNT"])
        == int(row["PROTOTYPE_BLANKED_BODY_COUNT"])
    ):
        issues.append("PROTOTYPE_BODIES_BLANKED")
    if row["PROTOTYPE_FULLY_LOADED"] == "NO":
        issues.append("PROTOTYPE_NOT_FULLY_LOADED")
    if "LIGHTWEIGHT" in representation or "PARTIAL" in representation:
        issues.append("LIGHTWEIGHT_OR_PARTIAL_REPRESENTATION")
    if int(dynamic_sections.get("clip_enabled_count", 0)) > 0:
        issues.append("ACTIVE_DYNAMIC_SECTION")

    # Preserve ranking order and remove duplicates.
    issues = list(dict.fromkeys(issues))
    row["ISSUE_CODES"] = " | ".join(issues) if issues else "NO_DIRECT_CAUSE_FOUND"
    if issues:
        confidence, cause, recommendation = CAUSE_GUIDANCE[issues[0]]
    else:
        confidence = "LOW"
        cause = "No direct visibility cause was exposed by the read-only NXOpen probes."
        recommendation = (
            "Use the target filter, confirm the exact occurrence path, capture a screenshot, "
            "and return the J23 JSON so the next probe can focus on view/isolate or corrupt display data."
        )
    row["ROOT_CAUSE"] = cause
    row["CONFIDENCE"] = confidence
    row["RECOMMENDATION"] = recommendation
    return row


def initial_row(root_name, timestamp, component, parent_name, path, level, target):
    component_name = safe_name(component, "<component unavailable>")
    full_path = "{0} / {1}".format(path, component_name)
    row = {column: "" for column in CSV_COLUMNS}
    row.update(
        {
            "RUN_TIMESTAMP": timestamp,
            "JOURNAL_BUILD": BUILD,
            "ROOT_ASSEMBLY": root_name,
            "LEVEL": level,
            "ASSEMBLY_PATH": full_path,
            "COMPONENT_NAME": component_name,
            "PARENT_COMPONENT": parent_name,
            "REFERENCE_SET_MEMBER_COUNT": 0,
            "REFERENCE_SET_BODY_COUNT": 0,
            "REFERENCE_SET_COMPONENT_COUNT": 0,
            "PROTOTYPE_BODY_COUNT": 0,
            "PROTOTYPE_SOLID_BODY_COUNT": 0,
            "PROTOTYPE_CHILD_COMPONENT_COUNT": 0,
            "PROTOTYPE_BLANKED_BODY_COUNT": 0,
            "PROTOTYPE_HIDDEN_LAYER_BODY_COUNT": 0,
            "GEOMETRY_MEMBERS_PROBED": 0,
            "OCCURRENCE_MEMBERS_FOUND": 0,
            "OCCURRENCE_MEMBERS_BLANKED": 0,
            "OCCURRENCE_BODY_TAGS_TESTED_IN_WORK_VIEW": 0,
            "OCCURRENCE_MEMBERS_VISIBLE_IN_WORK_VIEW": 0,
            "PROBE_LIMIT_REACHED": "NO",
            "TARGET_MATCH": "YES" if target and target in full_path.upper() else "NO",
        }
    )
    return row


def inspect_occurrence(
    work_part,
    component_assembly,
    component,
    row,
    inherited,
    view_snapshot,
    prototype_cache,
    dynamic_sections,
    target,
):
    errors = []
    suppressed, error = bool_property(component, "IsSuppressed")
    if error:
        errors.append(error)
    blanked, error = bool_property(component, "IsBlanked")
    if error:
        errors.append(error)
    layer = safe_value(component, "Layer", "")
    layer_state, error = get_layer_state(work_part, layer)
    if error:
        errors.append(error)
    layer_hidden = layer_state.upper() in HIDDEN_LAYER_STATES

    suppression = suppression_details(component_assembly, component)
    if suppression["error"]:
        errors.append(suppression["error"])
    try:
        non_geometric = bool(component_assembly.GetNonGeometricState(component))
    except Exception as error:
        non_geometric = False
        errors.append("GetNonGeometricState: {0}".format(error_text(error)))
    try:
        representation = enum_text(component.GetComponentRepresentationMode())
    except Exception as error:
        representation = ""
        errors.append("GetComponentRepresentationMode: {0}".format(error_text(error)))

    row.update(
        {
            "REFERENCE_SET": clean(safe_value(component, "ReferenceSet")),
            "SUPPRESSED": yes_no(suppressed),
            "SUPPRESSED_STATE": suppression["state"],
            "SUPPRESSION_EXPRESSION": suppression["expression"],
            "SUPPRESSING_ARRANGEMENT": safe_name(
                safe_value(component, "SuppressingArrangement", None), ""
            ),
            "ANCESTOR_SUPPRESSED": yes_no(inherited["suppressed"]),
            "IS_BLANKED": yes_no(blanked),
            "ANCESTOR_BLANKED": yes_no(inherited["blanked"]),
            "COMPONENT_LAYER": clean(layer),
            "COMPONENT_LAYER_STATE": layer_state,
            "ANCESTOR_HIDDEN_LAYER": yes_no(inherited["hidden_layer"]),
            "NON_GEOMETRIC": yes_no(non_geometric),
            "REPRESENTATION_MODE": representation,
            "USED_ARRANGEMENT": safe_name(
                safe_value(component, "UsedArrangement", None), ""
            ),
        }
    )

    try:
        prototype = component.Prototype
    except Exception as error:
        prototype = None
        errors.append("Component.Prototype: {0}".format(error_text(error)))
    row["PROTOTYPE_NAME"] = safe_name(prototype, "<unavailable>")
    if prototype is not None:
        row["PART_NUMBER"] = get_string_attribute(
            prototype, ("DB_PART_NO", "ITEM_ID", "PART_NUMBER")
        )
        row["REVISION"] = get_string_attribute(
            prototype, ("DB_PART_REV", "ITEM_REVISION", "REVISION")
        )
        if target:
            target_fields = " | ".join(
                (
                    row["ASSEMBLY_PATH"],
                    row["COMPONENT_NAME"],
                    row["PROTOTYPE_NAME"],
                    row["PART_NUMBER"],
                    row["REVISION"],
                )
            ).upper()
            if target in target_fields:
                row["TARGET_MATCH"] = "YES"
        row["PROTOTYPE_LOAD_STATE"] = enum_text(
            safe_value(prototype, "PartLoadState")
        )
        fully_loaded, error = bool_property(prototype, "IsFullyLoaded")
        row["PROTOTYPE_FULLY_LOADED"] = yes_no(fully_loaded)
        if error:
            errors.append(error)

        key = object_tag(prototype) or safe_name(prototype)
        if key not in prototype_cache:
            prototype_cache[key] = prototype_geometry(prototype)
        geometry = prototype_cache[key]
        errors.extend(geometry["errors"])
        row.update(
            {
                "PROTOTYPE_BODY_COUNT": geometry["body_count"],
                "PROTOTYPE_SOLID_BODY_COUNT": geometry["solid_count"],
                "PROTOTYPE_CHILD_COMPONENT_COUNT": geometry["child_count"],
                "PROTOTYPE_BLANKED_BODY_COUNT": geometry["blanked_count"],
                "PROTOTYPE_HIDDEN_LAYER_BODY_COUNT": geometry["hidden_layer_count"],
            }
        )

        refset = reference_set_details(
            prototype, row["REFERENCE_SET"], geometry["bodies"]
        )
        errors.extend(refset["errors"])
        row.update(
            {
                "REFERENCE_SET_FOUND": yes_no(refset["found"]),
                "REFERENCE_SET_MEMBER_COUNT": len(refset["members"]),
                "REFERENCE_SET_BODY_COUNT": refset["body_count"],
                "REFERENCE_SET_COMPONENT_COUNT": refset["component_count"],
            }
        )
        occurrence = occurrence_geometry(component, refset["members"], view_snapshot)
        errors.extend(occurrence["errors"])
        row.update(
            {
                "GEOMETRY_MEMBERS_PROBED": occurrence["probed"],
                "OCCURRENCE_MEMBERS_FOUND": occurrence["found"],
                "OCCURRENCE_MEMBERS_BLANKED": occurrence["blanked"],
                "OCCURRENCE_BODY_TAGS_TESTED_IN_WORK_VIEW": occurrence[
                    "view_tested"
                ],
                "OCCURRENCE_MEMBERS_VISIBLE_IN_WORK_VIEW": occurrence["visible"],
                "PROBE_LIMIT_REACHED": yes_no(occurrence["limit_reached"]),
            }
        )
    else:
        row["REFERENCE_SET_FOUND"] = "NO"
        row["PROTOTYPE_FULLY_LOADED"] = "NO"

    row["PROBE_ERRORS"] = " | ".join(list(dict.fromkeys(errors)))
    diagnose_record(row, dynamic_sections)
    child_inherited = {
        "suppressed": inherited["suppressed"] or suppressed,
        "blanked": inherited["blanked"] or blanked,
        "hidden_layer": inherited["hidden_layer"] or layer_hidden,
    }
    return row, child_inherited


def collect_records(work_part, timestamp, target, view_snapshot, dynamic_sections):
    component_assembly = work_part.ComponentAssembly
    root = component_assembly.RootComponent
    if root is None:
        raise RuntimeError("The active work part has no assembly root component.")
    root_name = safe_name(work_part, "<HLA assembly>")
    try:
        children = list(root.GetChildren())
    except Exception as error:
        raise RuntimeError("Cannot read HLA root children: {0}".format(error_text(error)))

    records = []
    traversal_errors = []
    prototype_cache = {}
    stack = [
        (child, 1, root_name, root_name, {"suppressed": False, "blanked": False, "hidden_layer": False})
        for child in reversed(children)
    ]
    while stack:
        if len(records) >= MAX_OCCURRENCES:
            traversal_errors.append(
                "Traversal stopped at the safety limit of {0} occurrences.".format(
                    MAX_OCCURRENCES
                )
            )
            break
        component, level, parent_name, parent_path, inherited = stack.pop()
        row = initial_row(
            root_name, timestamp, component, parent_name, parent_path, level, target
        )
        row, child_inherited = inspect_occurrence(
            work_part,
            component_assembly,
            component,
            row,
            inherited,
            view_snapshot,
            prototype_cache,
            dynamic_sections,
            target,
        )
        records.append(row)
        try:
            children = list(component.GetChildren())
        except Exception as error:
            traversal_errors.append(
                "{0}: Component.GetChildren: {1}".format(
                    row["ASSEMBLY_PATH"], error_text(error)
                )
            )
            children = []
        for child in reversed(children):
            stack.append(
                (
                    child,
                    level + 1,
                    row["COMPONENT_NAME"],
                    row["ASSEMBLY_PATH"],
                    child_inherited,
                )
            )
    return records, traversal_errors


def public_view_snapshot(snapshot):
    return {
        "available": snapshot["available"],
        "name": snapshot["name"],
        "visible_object_count": snapshot["visible_object_count"],
        "visible_tag_count": len(snapshot["visible_tags"]),
        "error": snapshot["error"],
    }


def build_report(work_part, timestamp, target, records, traversal_errors, view, sections):
    active_arrangement = safe_name(
        safe_value(work_part.ComponentAssembly, "ActiveArrangement", None), ""
    )
    suspects = sorted(
        records,
        key=lambda row: (
            0 if row["TARGET_MATCH"] == "YES" else 1,
            {"HIGH": 0, "MEDIUM": 1, "LOW": 2}.get(row["CONFIDENCE"], 3),
            row["LEVEL"],
            row["ASSEMBLY_PATH"],
        ),
    )
    return {
        "journal_build": BUILD,
        "run_timestamp": timestamp,
        "scope": "READ_ONLY_HLA_VISIBILITY_DIAGNOSTIC",
        "root_assembly": safe_name(work_part, "<HLA assembly>"),
        "active_arrangement": active_arrangement,
        "target_filter": target,
        "work_view": public_view_snapshot(view),
        "dynamic_sections": sections,
        "occurrence_count": len(records),
        "high_confidence_count": sum(1 for row in records if row["CONFIDENCE"] == "HIGH"),
        "target_match_count": sum(1 for row in records if row["TARGET_MATCH"] == "YES"),
        "traversal_errors": traversal_errors,
        "ranked_occurrences": suspects,
    }


def write_csv(path, records):
    with open(path, "w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(handle, fieldnames=CSV_COLUMNS, extrasaction="ignore")
        writer.writeheader()
        for row in records:
            writer.writerow({column: row.get(column, "") for column in CSV_COLUMNS})


def write_json(path, report):
    with open(path, "w", encoding="utf-8") as handle:
        json.dump(report, handle, indent=2, ensure_ascii=False)


def run(session, run_datetime=None):
    now = run_datetime or datetime.datetime.now().astimezone()
    timestamp = now.isoformat(timespec="seconds")
    file_timestamp = now.strftime("%Y%m%d_%H%M%S")
    try:
        work_part = session.Parts.Work
        display_part = session.Parts.Display
    except Exception:
        work_part = None
        display_part = None
    if work_part is None:
        raise RuntimeError("Open the affected top-level HLA assembly first.")
    if not same_nx_object(display_part, work_part):
        raise RuntimeError(
            "Make the affected HLA both the displayed part and work part, then rerun J23."
        )
    root = safe_value(safe_value(work_part, "ComponentAssembly", None), "RootComponent", None)
    if root is None:
        raise RuntimeError("The active work/display part is not an HLA assembly.")

    target = clean(os.environ.get("NX_J23_TARGET")).upper()
    view = work_view_snapshot(work_part)
    sections = dynamic_section_snapshot(work_part)
    records, traversal_errors = collect_records(
        work_part, timestamp, target, view, sections
    )
    report = build_report(
        work_part, timestamp, target, records, traversal_errors, view, sections
    )

    folder = os.path.join(io_root(), OUTPUT_FOLDER)
    os.makedirs(folder, exist_ok=True)
    root_token = filename_token(
        get_string_attribute(work_part, ("DB_PART_NO", "ITEM_ID", "PART_NUMBER"))
        or safe_name(work_part)
    )
    stem = "J23_HLA_VISIBILITY_{0}_{1}".format(root_token, file_timestamp)
    csv_path = os.path.join(folder, stem + ".csv")
    json_path = os.path.join(folder, stem + ".json")
    write_csv(csv_path, records)
    write_json(json_path, report)
    return csv_path, json_path, report


def main():
    session = NXOpen.Session.GetSession()
    log_line(session, "=" * 72)
    log_line(session, "J23 HLA ASSEMBLY VISIBILITY DIAGNOSTIC")
    log_line(session, "Build: " + BUILD)
    log_line(session, "Scope: read-only; no visibility, assembly, load, or save changes.")
    log_line(session, "=" * 72)
    try:
        csv_path, json_path, report = run(session)
        log_line(
            session,
            "HLA: {0} | active arrangement: {1}".format(
                report["root_assembly"], report["active_arrangement"] or "<none>"
            ),
        )
        log_line(
            session,
            "Occurrences: {0} | high-confidence flags: {1} | target matches: {2}".format(
                report["occurrence_count"],
                report["high_confidence_count"],
                report["target_match_count"],
            ),
        )
        ranked = report["ranked_occurrences"]
        priority = [
            row
            for row in ranked
            if row["TARGET_MATCH"] == "YES" or row["CONFIDENCE"] == "HIGH"
        ][:20]
        if not priority:
            priority = ranked[:10]
        for row in priority:
            log_line(
                session,
                "[{0}] {1}\n  Issues: {2}\n  Root cause: {3}".format(
                    row["CONFIDENCE"],
                    row["ASSEMBLY_PATH"],
                    row["ISSUE_CODES"],
                    row["ROOT_CAUSE"],
                ),
            )
        if report["traversal_errors"]:
            log_line(session, "Traversal warnings: " + " | ".join(report["traversal_errors"]))
        log_line(session, "CSV: " + csv_path)
        log_line(session, "JSON: " + json_path)
        log_line(session, "Return the JSON plus the exact missing component name/path.")
    except Exception as error:
        log_line(session, "J23 FAILED: " + error_text(error))
        log_line(session, traceback.format_exc())
        raise


def get_unload_option(dummy):
    return NXOpen.Session.LibraryUnloadOption.Immediately


if __name__ == "__main__":
    main()
