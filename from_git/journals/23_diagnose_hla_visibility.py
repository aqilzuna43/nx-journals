"""J23 V2 - target-focused, read-only HLA visibility root-cause proof.

Select one missing component in Assembly Navigator before playing the journal,
or set NX_J23_TARGET to its exact part number.  USER_TARGET is the fallback for
the current investigation.  V2 never converts a failed probe into False: every
value is OBSERVED, ERROR, UNAVAILABLE, or NOT_APPLICABLE.  Conclusions cite the
fact IDs that prove them; all other hypotheses are explicitly ruled out or
left inconclusive.

The journal changes no NX object, view, load state, arrangement, or Teamcenter
data and never saves.

Target: NX 2312 and NX X 2506 embedded Python.
Run via: NX > Tools > Journal > Play
"""

import csv
import datetime
import json
import os
import traceback

import NXOpen


BUILD = "J23-NX2506-HLA-VISIBILITY-EVIDENCE-V2"
SCHEMA_VERSION = 2
USER_TARGET = "264MN031978A01"
OUTPUT_FOLDER = "NX_HLA_VISIBILITY_DIAGNOSTIC"
MAX_OCCURRENCES = 100000
MAX_MEMBER_PROBES = 500

OBSERVED = "OBSERVED"
ERROR = "ERROR"
UNAVAILABLE = "UNAVAILABLE"
NOT_APPLICABLE = "NOT_APPLICABLE"
HIDDEN_LAYER_NAMES = ("HIDDEN", "INVISIBLE")
ENTIRE_REFSETS = ("ENTIRE PART", "ENTIRE_PART", "ENTIRE")
EMPTY_REFSETS = ("EMPTY", "EMPTY PART", "EMPTY_PART")

CSV_COLUMNS = (
    "ROLE",
    "ASSEMBLY_PATH",
    "LEVEL",
    "COMPONENT_TAG",
    "PART_NUMBER",
    "REVISION",
    "REFERENCE_SET",
    "SUPPRESSED_STATUS",
    "SUPPRESSED_VALUE",
    "BLANKED_STATUS",
    "BLANKED_VALUE",
    "NON_GEOMETRIC_STATUS",
    "NON_GEOMETRIC_VALUE",
    "COMPONENT_LAYER",
    "COMPONENT_LAYER_STATE_STATUS",
    "COMPONENT_LAYER_STATE_VALUE",
    "PROTOTYPE_TYPE",
    "LOAD_STATUS",
    "FULLY_LOADED_STATUS",
    "FULLY_LOADED_VALUE",
    "REFSET_PROBE_STATUS",
    "REFSET_FOUND",
    "REFSET_BODY_MEMBERS",
    "REFSET_COMPONENT_MEMBERS",
    "MAPPED_BODY_OCCURRENCES",
    "MAPPED_COMPONENT_OCCURRENCES",
    "MAPPED_BODIES_VISIBLE_CURRENT_VIEW",
    "CURRENT_VIEW_NAME",
    "PROBE_ERRORS",
)


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


def error_text(error):
    message = clean(error) or type(error).__name__
    code = clean(getattr(error, "ErrorCode", ""))
    return "{0}{1}".format(message, " [NX error {0}]".format(code) if code else "")


def probe(status, value=None, source="", error=""):
    return {"status": status, "value": value, "source": source, "error": error}


def observed(value, source):
    return probe(OBSERVED, value=value, source=source)


def failed(source, error):
    return probe(ERROR, value=None, source=source, error=error_text(error))


def unavailable(source, reason):
    return probe(UNAVAILABLE, value=None, source=source, error=clean(reason))


def safe_value(value, property_name, fallback=""):
    if value is None:
        return fallback
    try:
        result = getattr(value, property_name)
        return result() if callable(result) else result
    except Exception:
        return fallback


def safe_name(value, fallback="<unavailable>"):
    for name in ("DisplayName", "Name", "Leaf", "JournalIdentifier", "FullPath"):
        text = clean(safe_value(value, name))
        if text:
            return text
    return fallback


def object_tag(value):
    try:
        tag = clean(value.Tag)
        return tag if tag and tag != "0" else ""
    except Exception:
        return ""


def object_kind(value):
    if value is None:
        return ""
    try:
        name = clean(getattr(value.GetType(), "Name", ""))
        if name:
            return name
    except Exception:
        pass
    return clean(type(value).__name__)


def is_body(value):
    return "body" in object_kind(value).lower()


def is_component(value):
    name = object_kind(value).lower()
    return "component" in name and "assembly" not in name


def same_object(first, second):
    if first is second:
        return True
    left, right = object_tag(first), object_tag(second)
    return bool(left and right and left == right)


def property_probe(value, property_name):
    source = "{0}.{1}".format(object_kind(value) or "object", property_name)
    if value is None:
        return unavailable(source, "Object is unavailable.")
    try:
        result = getattr(value, property_name)
    except AttributeError:
        return unavailable(source, "Property is not exposed by this runtime object type.")
    except Exception as error:
        return failed(source, error)
    try:
        return observed(result() if callable(result) else result, source)
    except Exception as error:
        return failed(source, error)


def method_probe(value, method_name, *args):
    source = "{0}.{1}".format(object_kind(value) or "object", method_name)
    if value is None:
        return unavailable(source, "Object is unavailable.")
    try:
        method = getattr(value, method_name)
    except AttributeError:
        return unavailable(source, "Method is not exposed by this runtime object type.")
    except Exception as error:
        return failed(source, error)
    try:
        return observed(method(*args), source)
    except Exception as error:
        return failed(source, error)


def get_string_attribute(nx_object, names):
    for name in names:
        try:
            text = clean(nx_object.GetStringAttribute(name))
            if text:
                return text
        except Exception:
            pass
        try:
            info = nx_object.GetUserAttribute(
                name, NXOpen.NXObject.AttributeType.String, -1
            )
            text = clean(getattr(info, "StringValue", ""))
            if text:
                return text
        except Exception:
            pass
    return ""


def layer_state_probe(work_part, layer):
    source = "work_part.Layers.GetState({0})".format(clean(layer))
    try:
        raw = work_part.Layers.GetState(int(layer))
        text = enum_text(raw)
        numeric = clean(raw)
        if text == numeric:
            text = {"0": "WorkLayer", "1": "Selectable", "2": "Visible", "3": "Hidden"}.get(
                numeric, text
            )
        return observed(text, source)
    except Exception as error:
        return failed(source, error)


def list_probe(value, property_name):
    source = "{0}.{1}".format(object_kind(value) or "object", property_name)
    try:
        return observed(list(getattr(value, property_name)), source)
    except AttributeError:
        return unavailable(source, "Collection is not exposed by this runtime object type.")
    except Exception as error:
        return failed(source, error)


def collection_items(value):
    try:
        return list(value)
    except Exception:
        pass
    try:
        return list(value.ToArray())
    except Exception:
        return []


def log_line(session, message):
    text = clean(message)
    try:
        window = session.ListingWindow
        window.Open()
        writer = getattr(window, "WriteFullline", None) or getattr(window, "WriteLine", None)
        if callable(writer):
            for line in text.splitlines() or [""]:
                writer(line)
    except Exception:
        pass
    try:
        print(text)
    except Exception:
        pass


def io_root():
    configured = clean(os.environ.get("NX_JOURNALS_IO_DIR"))
    if configured:
        return os.path.abspath(os.path.expanduser(configured))
    profile = clean(os.environ.get("USERPROFILE"))
    return os.path.join(profile, "Desktop") if profile else os.getcwd()


def filename_token(value):
    text = clean(value) or "UNKNOWN"
    return "".join(c if c.isalnum() or c in "-_" else "_" for c in text)[:80]


class EvidenceLedger:
    def __init__(self):
        self.items = []

    def add(self, category, statement, source, value=None, status=OBSERVED):
        item = {
            "id": "F{0:03d}".format(len(self.items) + 1),
            "status": status,
            "category": category,
            "statement": statement,
            "source": source,
            "value": value,
        }
        self.items.append(item)
        return item["id"]


def component_identity(component, prototype, path, level, parent_tag):
    return {
        "assembly_path": path,
        "level": level,
        "component_tag": object_tag(component),
        "component_name": safe_name(component, "<component>"),
        "parent_component_tag": parent_tag,
        "prototype_tag": object_tag(prototype),
        "prototype_type": object_kind(prototype),
        "prototype_name": safe_name(prototype, "<unavailable>"),
        "part_number": get_string_attribute(
            prototype, ("DB_PART_NO", "ITEM_ID", "PART_NUMBER")
        ) if prototype is not None else "",
        "revision": get_string_attribute(
            prototype, ("DB_PART_REV", "ITEM_REVISION", "REVISION")
        ) if prototype is not None else "",
        "reference_set": clean(safe_value(component, "ReferenceSet")),
    }


def collect_nodes(work_part):
    root = work_part.ComponentAssembly.RootComponent
    children_probe = method_probe(root, "GetChildren")
    if children_probe["status"] != OBSERVED:
        raise RuntimeError("Cannot read HLA root children: " + children_probe["error"])
    root_name = safe_name(work_part, "<HLA>")
    stack = [
        (child, 1, root_name, [], "")
        for child in reversed(list(children_probe["value"]))
    ]
    nodes, errors = [], []
    while stack:
        if len(nodes) >= MAX_OCCURRENCES:
            errors.append("Safety limit reached: {0}".format(MAX_OCCURRENCES))
            break
        component, level, parent_path, ancestor_tags, parent_tag = stack.pop()
        name = safe_name(component, "<component>")
        path = "{0} / {1}".format(parent_path, name)
        prototype_probe = property_probe(component, "Prototype")
        prototype = prototype_probe["value"] if prototype_probe["status"] == OBSERVED else None
        node = {
            "_component": component,
            "_prototype": prototype,
            "_ancestor_tags": list(ancestor_tags),
            "identity": component_identity(component, prototype, path, level, parent_tag),
            "prototype_probe": prototype_probe,
        }
        nodes.append(node)
        child_probe = method_probe(component, "GetChildren")
        if child_probe["status"] != OBSERVED:
            errors.append("{0}: {1}".format(path, child_probe["error"]))
            children = []
        else:
            children = list(child_probe["value"])
        next_ancestors = ancestor_tags + [object_tag(component)]
        for child in reversed(children):
            stack.append((child, level + 1, path, next_ancestors, object_tag(component)))
    return nodes, errors


def preselected_component():
    try:
        manager = NXOpen.UI.GetUI().SelectionManager
        count = int(manager.GetNumSelectedObjects())
    except Exception:
        return None
    candidates = []
    for index in range(count):
        try:
            selected = manager.GetSelectedTaggedObject(index)
        except Exception:
            continue
        if is_component(selected):
            candidates.append(selected)
            continue
        owner = safe_value(selected, "OwningComponent", None)
        if owner is not None and is_component(owner):
            candidates.append(owner)
    unique = {object_tag(item): item for item in candidates if object_tag(item)}
    return list(unique.values())[0] if len(unique) == 1 else None


def resolve_targets(nodes):
    selected = preselected_component()
    if selected is not None:
        tag = object_tag(selected)
        matches = [node for node in nodes if node["identity"]["component_tag"] == tag]
        if matches:
            return {"source": "ASSEMBLY_NAVIGATOR_PRESELECTION", "value": tag}, matches
    requested = clean(os.environ.get("NX_J23_TARGET") or USER_TARGET).upper()
    if not requested:
        raise RuntimeError(
            "Select one missing component in Assembly Navigator or set NX_J23_TARGET."
        )
    exact = [
        node for node in nodes
        if node["identity"]["part_number"].upper() == requested
    ]
    if exact:
        return {"source": "EXACT_PART_NUMBER", "value": requested}, exact
    fallback = [
        node for node in nodes
        if requested in node["identity"]["assembly_path"].upper()
        or requested in node["identity"]["prototype_name"].upper()
    ]
    return {"source": "NAME_OR_PATH_SUBSTRING", "value": requested}, fallback


def component_state(component, work_part):
    suppressed = property_probe(component, "IsSuppressed")
    blanked = property_probe(component, "IsBlanked")
    layer = property_probe(component, "Layer")
    layer_state = (
        layer_state_probe(work_part, layer["value"])
        if layer["status"] == OBSERVED
        else unavailable("work_part.Layers.GetState", "Component layer is unknown.")
    )
    non_geometric = method_probe(component, "GetNonGeometricState")
    representation = method_probe(component, "GetComponentRepresentationMode")
    if representation["status"] == OBSERVED:
        representation["value"] = enum_text(representation["value"])
    return {
        "suppressed": suppressed,
        "blanked": blanked,
        "layer": layer,
        "layer_state": layer_state,
        "non_geometric": non_geometric,
        "representation": representation,
    }


def prototype_state(prototype):
    load_state = property_probe(prototype, "PartLoadState")
    if load_state["status"] == OBSERVED:
        load_state["value"] = enum_text(load_state["value"])
    fully_loaded = property_probe(prototype, "IsFullyLoaded")
    bodies = list_probe(prototype, "Bodies")
    child_probe = unavailable("Prototype.ComponentAssembly.RootComponent.GetChildren", "Prototype is unavailable.")
    if prototype is not None:
        assembly = safe_value(prototype, "ComponentAssembly", None)
        root = safe_value(assembly, "RootComponent", None)
        if root is None:
            child_probe = observed([], "Prototype has no component root")
        else:
            child_probe = method_probe(root, "GetChildren")
            if child_probe["status"] == OBSERVED:
                child_probe["value"] = list(child_probe["value"])
    return {
        "runtime_type": object_kind(prototype),
        "load_state": load_state,
        "fully_loaded": fully_loaded,
        "bodies": bodies,
        "direct_children": child_probe,
    }


def reference_set_probe(component, prototype, prototype_info):
    selected = clean(safe_value(component, "ReferenceSet")).upper()
    result = {
        "status": OBSERVED,
        "source": "Component.ReferenceSet + Prototype.GetAllReferenceSets",
        "selected": selected,
        "found": False,
        "kind": "NAMED",
        "members": [],
        "body_members": [],
        "component_members": [],
        "error": "",
    }
    if selected in EMPTY_REFSETS:
        result.update({"found": True, "kind": "EMPTY"})
        return result
    entire_name = clean(safe_value(prototype, "EntirePartRefsetName")).upper()
    if selected in ENTIRE_REFSETS or (entire_name and selected == entire_name):
        result.update({"found": True, "kind": "ENTIRE_PART"})
        bodies = prototype_info["bodies"]
        children = prototype_info["direct_children"]
        if bodies["status"] != OBSERVED or children["status"] != OBSERVED:
            result["status"] = UNAVAILABLE
            result["error"] = "Entire Part members require both body and child-component probes."
            return result
        result["body_members"] = list(bodies["value"])
        result["component_members"] = list(children["value"])
        result["members"] = result["body_members"] + result["component_members"]
        return result
    refs_probe = method_probe(prototype, "GetAllReferenceSets")
    if refs_probe["status"] != OBSERVED:
        result.update({"status": refs_probe["status"], "error": refs_probe["error"]})
        return result
    selected_ref = None
    for refset in list(refs_probe["value"]):
        if safe_name(refset, "").upper() == selected:
            selected_ref = refset
            break
    if selected_ref is None:
        return result
    result["found"] = True
    members_probe = method_probe(selected_ref, "AskMembersInReferenceSet")
    if members_probe["status"] != OBSERVED:
        result.update({"status": members_probe["status"], "error": members_probe["error"]})
        return result
    result["members"] = list(members_probe["value"])
    result["body_members"] = [item for item in result["members"] if is_body(item)]
    result["component_members"] = [item for item in result["members"] if is_component(item)]
    return result


def map_occurrence_members(component, refset):
    result = {
        "status": OBSERVED,
        "source": "Component.FindOccurrence(reference-set member)",
        "body_occurrence_tags": [],
        "component_occurrence_tags": [],
        "body_blanked": [],
        "errors": [],
        "limit_reached": False,
    }
    members = list(refset.get("members", []))
    if refset.get("status") != OBSERVED:
        result.update({"status": UNAVAILABLE, "source": refset.get("source", "reference set")})
        return result
    if len(members) > MAX_MEMBER_PROBES:
        members = members[:MAX_MEMBER_PROBES]
        result["limit_reached"] = True
    for member in members:
        mapped = method_probe(component, "FindOccurrence", member)
        if mapped["status"] != OBSERVED:
            result["errors"].append(mapped["error"])
            continue
        occurrence = mapped["value"]
        if occurrence is None:
            continue
        tag = object_tag(occurrence)
        if is_body(member):
            if tag:
                result["body_occurrence_tags"].append(tag)
            result["body_blanked"].append(property_probe(occurrence, "IsBlanked"))
        elif is_component(member) and tag:
            result["component_occurrence_tags"].append(tag)
    if result["errors"]:
        result["status"] = ERROR
    return result


def view_snapshots(work_part, mapped_tags):
    current = work_part.ModelingViews.WorkView
    views = collection_items(work_part.ModelingViews)
    if not any(same_object(view, current) for view in views):
        views.insert(0, current)
    snapshots = []
    for view in views:
        visible = method_probe(view, "AskVisibleObjects")
        row = {
            "name": safe_name(view, "<view>"),
            "tag": object_tag(view),
            "is_work_view": same_object(view, current),
            "probe_status": visible["status"],
            "visible_object_count": None,
            "mapped_target_tags_visible": [],
            "error": visible["error"],
        }
        if visible["status"] == OBSERVED:
            objects = list(visible["value"])
            visible_tags = {object_tag(item) for item in objects if object_tag(item)}
            row["visible_object_count"] = len(objects)
            row["mapped_target_tags_visible"] = sorted(set(mapped_tags) & visible_tags)
        snapshots.append(row)
    return snapshots


def dynamic_section_evidence(work_part):
    view = work_part.ModelingViews.WorkView
    sections = collection_items(safe_value(work_part, "DynamicSections", None))
    rows = []
    for section in sections:
        state = method_probe(view, "IsDynamicSectionVisible", section)
        rows.append(
            {
                "name": safe_name(section, "<section>"),
                "status": state["status"],
                "visible_in_work_view": bool(state["value"]) if state["status"] == OBSERVED else None,
                "source": state["source"],
                "error": state["error"],
            }
        )
    return rows


def analyze_node(node, work_part, current_visible_tags):
    component, prototype = node["_component"], node["_prototype"]
    state = component_state(component, work_part)
    proto = prototype_state(prototype)
    refset = reference_set_probe(component, prototype, proto)
    mapping = map_occurrence_members(component, refset)
    body_tags = mapping["body_occurrence_tags"]
    visible = sorted(set(body_tags) & current_visible_tags)
    errors = []
    for item in list(state.values()) + [proto["load_state"], proto["fully_loaded"], proto["bodies"], proto["direct_children"]]:
        if item.get("status") == ERROR:
            errors.append(item.get("error", ""))
    errors.extend(mapping["errors"])
    if refset["status"] == ERROR:
        errors.append(refset["error"])
    return {
        "identity": node["identity"],
        "component_state": state,
        "prototype": {
            "runtime_type": proto["runtime_type"],
            "load_state": proto["load_state"],
            "fully_loaded": proto["fully_loaded"],
            "direct_body_count": len(proto["bodies"]["value"]) if proto["bodies"]["status"] == OBSERVED else None,
            "direct_child_count": len(proto["direct_children"]["value"]) if proto["direct_children"]["status"] == OBSERVED else None,
            "body_probe_status": proto["bodies"]["status"],
            "child_probe_status": proto["direct_children"]["status"],
        },
        "reference_set": {
            "status": refset["status"],
            "selected": refset["selected"],
            "found": refset["found"] if refset["status"] == OBSERVED else None,
            "kind": refset["kind"],
            "member_count": len(refset["members"]),
            "body_member_count": len(refset["body_members"]),
            "component_member_count": len(refset["component_members"]),
            "error": refset["error"],
        },
        "mapping": {
            "status": mapping["status"],
            "mapped_body_occurrence_tags": body_tags,
            "mapped_component_occurrence_tags": mapping["component_occurrence_tags"],
            "mapped_body_count": len(body_tags),
            "mapped_component_count": len(mapping["component_occurrence_tags"]),
            "mapped_bodies_visible_current_view": visible,
            "mapped_bodies_visible_current_view_count": len(visible),
            "body_blanked_probes": mapping["body_blanked"],
            "limit_reached": mapping["limit_reached"],
        },
        "probe_errors": list(dict.fromkeys(error for error in errors if error)),
    }


def boolean_observed(item, expected):
    return item["status"] == OBSERVED and bool(item["value"]) is expected


def layer_hidden(item):
    return item["status"] == OBSERVED and clean(item["value"]).upper() in HIDDEN_LAYER_NAMES


def hypothesis(code, verdict, statement, evidence, missing=None):
    return {
        "code": code,
        "verdict": verdict,
        "statement": statement,
        "evidence_ids": list(evidence),
        "missing_evidence": list(missing or []),
    }


def build_hypotheses(target, subtree, views, sections, controls, ledger):
    hypotheses = []
    mapped = sum(row["mapping"]["mapped_body_count"] for row in subtree)
    visible_current = sum(row["mapping"]["mapped_bodies_visible_current_view_count"] for row in subtree)
    unsuppressed_absent = [
        row for row in subtree
        if boolean_observed(row["component_state"]["suppressed"], False)
        and row["mapping"]["mapped_body_count"] > 0
        and row["mapping"]["mapped_bodies_visible_current_view_count"] == 0
    ]
    unblanked_absent = [
        row for row in subtree
        if boolean_observed(row["component_state"]["blanked"], False)
        and row["mapping"]["mapped_body_count"] > 0
        and row["mapping"]["mapped_bodies_visible_current_view_count"] == 0
    ]
    mapped_id = ledger.add(
        "OCCURRENCE_MAPPING",
        "Reference-set body geometry mapped to HLA occurrence objects.",
        "Component.FindOccurrence across target subtree",
        mapped,
    )
    current_id = ledger.add(
        "CURRENT_VIEW",
        "Mapped target-subtree body occurrences present in the current work view.",
        "WorkView.AskVisibleObjects tag intersection",
        visible_current,
    )
    hypotheses.append(hypothesis(
        "OCCURRENCE_MAPPING_FAILURE",
        "RULED_OUT" if mapped > 0 else "INCONCLUSIVE",
        "Occurrence mapping is not the cause." if mapped > 0 else "No mapped bodies were proven.",
        [mapped_id],
    ))

    suppression_id = ledger.add(
        "SUPPRESSION",
        "Unsuppressed subtree occurrences with mapped geometry are still absent.",
        "Component.IsSuppressed + current-view intersection",
        len(unsuppressed_absent),
    )
    hypotheses.append(hypothesis(
        "SUPPRESSION_AS_PRIMARY_CAUSE",
        "RULED_OUT" if unsuppressed_absent else "INCONCLUSIVE",
        "Suppression cannot explain the whole missing subtree." if unsuppressed_absent else "Suppression evidence is incomplete.",
        [suppression_id],
    ))
    blanking_id = ledger.add(
        "BLANKING",
        "Unblanked subtree occurrences with mapped geometry are still absent.",
        "Component.IsBlanked + current-view intersection",
        len(unblanked_absent),
    )
    hypotheses.append(hypothesis(
        "BLANKING_AS_PRIMARY_CAUSE",
        "RULED_OUT" if unblanked_absent else "INCONCLUSIVE",
        "Blanking cannot explain the whole missing subtree." if unblanked_absent else "Blanking evidence is incomplete.",
        [blanking_id],
    ))

    valid_refsets = [
        row for row in subtree
        if row["reference_set"]["status"] == OBSERVED
        and row["reference_set"]["found"] is True
        and row["mapping"]["mapped_body_count"] > 0
    ]
    refset_id = ledger.add(
        "REFERENCE_SET",
        "Subtree occurrences have found reference sets whose bodies map into the HLA.",
        "ReferenceSet members + Component.FindOccurrence",
        len(valid_refsets),
    )
    hypotheses.append(hypothesis(
        "REFERENCE_SET_AS_PRIMARY_CAUSE",
        "RULED_OUT" if valid_refsets else "INCONCLUSIVE",
        "Reference-set failure cannot explain the mapped-but-absent bodies." if valid_refsets else "Reference-set evidence is incomplete.",
        [refset_id, mapped_id],
    ))

    hidden_layers = [row for row in subtree if layer_hidden(row["component_state"]["layer_state"])]
    known_layers = [row for row in subtree if row["component_state"]["layer_state"]["status"] == OBSERVED]
    layer_id = ledger.add(
        "HLA_LAYER",
        "Hidden component layers among target-subtree occurrence rows.",
        "Displayed HLA work_part.Layers.GetState(component.Layer)",
        len(hidden_layers),
    )
    layer_verdict = "RULED_OUT" if known_layers and not hidden_layers else ("CONFIRMED" if hidden_layers else "INCONCLUSIVE")
    hypotheses.append(hypothesis(
        "HIDDEN_HLA_COMPONENT_LAYER",
        layer_verdict,
        "No target-subtree component is on a hidden HLA layer." if layer_verdict == "RULED_OUT" else "Hidden HLA component-layer evidence exists." if hidden_layers else "Layer evidence is unavailable.",
        [layer_id],
    ))

    section_visible = [row for row in sections if row["status"] == OBSERVED and row["visible_in_work_view"]]
    section_known = all(row["status"] == OBSERVED for row in sections)
    section_id = ledger.add(
        "DYNAMIC_SECTION",
        "Dynamic sections visible in the current work view.",
        "ModelingView.IsDynamicSectionVisible",
        [row["name"] for row in section_visible],
    )
    hypotheses.append(hypothesis(
        "DYNAMIC_SECTION_CLIPPING",
        "RULED_OUT" if section_known and not section_visible else ("POSSIBLE" if section_visible else "INCONCLUSIVE"),
        "No dynamic section is visible in the current work view." if section_known and not section_visible else "Dynamic-section evidence does not rule clipping out.",
        [section_id],
    ))

    other_visible = [
        view for view in views
        if not view["is_work_view"] and len(view["mapped_target_tags_visible"]) > 0
    ]
    alternate_id = ledger.add(
        "VIEW_COMPARISON",
        "Non-work modeling views containing mapped target-body occurrence tags.",
        "ModelingView.AskVisibleObjects across saved views",
        [{"name": row["name"], "count": len(row["mapped_target_tags_visible"])} for row in other_visible],
    )
    control_id = ledger.add(
        "SAME_PROTOTYPE_CONTROL",
        "Same part/revision controls outside the target subtree visible in the current view.",
        "Current WorkView tag intersection",
        len(controls),
    )
    view_confirmed = mapped > 0 and visible_current == 0 and bool(other_visible)
    view_supported = mapped > 0 and visible_current == 0 and bool(controls)
    verdict = "CONFIRMED" if view_confirmed else ("STRONGLY_SUPPORTED" if view_supported else "INCONCLUSIVE")
    hypotheses.append(hypothesis(
        "CURRENT_WORK_VIEW_EXCLUSION",
        verdict,
        "Mapped target geometry is absent from the work view but present in another modeling view." if view_confirmed else "Mapped target geometry is absent while same-prototype controls are visible in the work view." if view_supported else "The current view excludes mapped target geometry, but no independent view/control completed the proof.",
        [mapped_id, current_id, alternate_id, control_id],
        [] if verdict == "CONFIRMED" else ["A readable alternate modeling view containing the exact target occurrence tags."],
    ))

    work_view = next((row for row in views if row["is_work_view"]), None)
    isolate_named = bool(work_view and clean(work_view["name"]).upper() == "ISOLATE")
    isolate_id = ledger.add(
        "ISOLATE_CONTEXT",
        "The active NX work-view name is exactly 'Isolate'.",
        "work_part.ModelingViews.WorkView.Name",
        isolate_named,
    )
    isolate_verdict = "STRONGLY_SUPPORTED" if isolate_named and verdict in ("CONFIRMED", "STRONGLY_SUPPORTED") else "INCONCLUSIVE"
    hypotheses.append(hypothesis(
        "ISOLATE_VIEW_MECHANISM",
        isolate_verdict,
        "Isolation is strongly supported by the work-view name plus independent view-exclusion evidence; the public read API does not expose isolate membership directly." if isolate_verdict == "STRONGLY_SUPPORTED" else "A view name alone cannot prove isolate membership.",
        [isolate_id, mapped_id, current_id, alternate_id, control_id],
        ["NXOpen exposes commands to create/change isolate membership but no corresponding read-only membership query."],
    ))

    # Mapped geometry proves that incomplete loading is not the reason those exact objects are absent.
    load_id = ledger.add(
        "LOAD_STATE",
        "Mapped occurrence bodies exist even where load-state properties are partial or unavailable.",
        "Component.FindOccurrence",
        mapped,
    )
    hypotheses.append(hypothesis(
        "INCOMPLETE_LOAD_AS_PRIMARY_CAUSE",
        "RULED_OUT" if mapped > 0 else "INCONCLUSIVE",
        "Incomplete loading cannot explain absence of the already-mapped occurrence bodies." if mapped > 0 else "Load evidence is incomplete.",
        [load_id, mapped_id],
    ))

    confirmed = next((item for item in hypotheses if item["code"] == "CURRENT_WORK_VIEW_EXCLUSION" and item["verdict"] == "CONFIRMED"), None)
    supported = next((item for item in hypotheses if item["code"] == "CURRENT_WORK_VIEW_EXCLUSION" and item["verdict"] == "STRONGLY_SUPPORTED"), None)
    if confirmed:
        conclusion = {
            "status": "CONFIRMED",
            "root_cause_code": "CURRENT_WORK_VIEW_EXCLUSION",
            "statement": confirmed["statement"],
            "evidence_ids": confirmed["evidence_ids"],
        }
    elif supported:
        conclusion = {
            "status": "STRONGLY_SUPPORTED",
            "root_cause_code": "CURRENT_WORK_VIEW_OR_OCCURRENCE_DISPLAY_EXCLUSION",
            "statement": supported["statement"],
            "evidence_ids": supported["evidence_ids"],
        }
    else:
        conclusion = {
            "status": "INCONCLUSIVE",
            "root_cause_code": "UNRESOLVED",
            "statement": "The available read-only probes do not yet prove one root cause.",
            "evidence_ids": [mapped_id, current_id],
        }
    return hypotheses, conclusion


def analyze_target(target, nodes, work_part):
    ledger = EvidenceLedger()
    target_tag = target["identity"]["component_tag"]
    subtree_nodes = [
        node for node in nodes
        if node is target or target_tag in node["_ancestor_tags"]
    ]
    current_view = work_part.ModelingViews.WorkView
    current_probe = method_probe(current_view, "AskVisibleObjects")
    if current_probe["status"] != OBSERVED:
        raise RuntimeError("Cannot inventory current work view: " + current_probe["error"])
    current_objects = list(current_probe["value"])
    current_tags = {object_tag(item) for item in current_objects if object_tag(item)}
    subtree = [analyze_node(node, work_part, current_tags) for node in subtree_nodes]
    mapped_tags = sorted({
        tag for row in subtree for tag in row["mapping"]["mapped_body_occurrence_tags"]
    })
    views = view_snapshots(work_part, mapped_tags)
    sections = dynamic_section_evidence(work_part)

    subtree_keys = {
        (row["identity"]["part_number"], row["identity"]["revision"])
        for row in subtree if row["identity"]["part_number"]
    }
    subtree_component_tags = {row["identity"]["component_tag"] for row in subtree}
    controls = []
    for node in nodes:
        identity = node["identity"]
        key = (identity["part_number"], identity["revision"])
        if identity["component_tag"] in subtree_component_tags or key not in subtree_keys:
            continue
        control = analyze_node(node, work_part, current_tags)
        if control["mapping"]["mapped_bodies_visible_current_view_count"] > 0:
            controls.append({
                "assembly_path": identity["assembly_path"],
                "part_number": identity["part_number"],
                "revision": identity["revision"],
                "visible_mapped_body_count": control["mapping"]["mapped_bodies_visible_current_view_count"],
            })
    hypotheses, conclusion = build_hypotheses(
        target, subtree, views, sections, controls, ledger
    )
    summary = {
        "occurrence_rows": len(subtree),
        "mapped_body_occurrences": len(mapped_tags),
        "mapped_bodies_visible_current_view": len(
            set(mapped_tags) & current_tags
        ),
        "suppressed_rows": sum(
            1 for row in subtree if boolean_observed(row["component_state"]["suppressed"], True)
        ),
        "blanked_rows": sum(
            1 for row in subtree if boolean_observed(row["component_state"]["blanked"], True)
        ),
        "unsuppressed_mapped_absent_rows": sum(
            1 for row in subtree
            if boolean_observed(row["component_state"]["suppressed"], False)
            and row["mapping"]["mapped_body_count"] > 0
            and row["mapping"]["mapped_bodies_visible_current_view_count"] == 0
        ),
        "unblanked_mapped_absent_rows": sum(
            1 for row in subtree
            if boolean_observed(row["component_state"]["blanked"], False)
            and row["mapping"]["mapped_body_count"] > 0
            and row["mapping"]["mapped_bodies_visible_current_view_count"] == 0
        ),
    }
    return {
        "target": target["identity"],
        "subtree_summary": summary,
        "current_work_view": {
            "name": safe_name(current_view, "<work view>"),
            "tag": object_tag(current_view),
            "visible_object_count": len(current_objects),
        },
        "subtree_occurrences": subtree,
        "view_comparison": views,
        "dynamic_sections": sections,
        "same_prototype_controls": controls,
        "evidence_ledger": ledger.items,
        "hypotheses": hypotheses,
        "conclusion": conclusion,
    }


def csv_row(role, row, current_view_name):
    state, proto, refset, mapping = (
        row["component_state"], row["prototype"], row["reference_set"], row["mapping"]
    )
    identity = row["identity"]
    values = {
        "ROLE": role,
        "ASSEMBLY_PATH": identity["assembly_path"],
        "LEVEL": identity["level"],
        "COMPONENT_TAG": identity["component_tag"],
        "PART_NUMBER": identity["part_number"],
        "REVISION": identity["revision"],
        "REFERENCE_SET": identity["reference_set"],
        "SUPPRESSED_STATUS": state["suppressed"]["status"],
        "SUPPRESSED_VALUE": state["suppressed"]["value"],
        "BLANKED_STATUS": state["blanked"]["status"],
        "BLANKED_VALUE": state["blanked"]["value"],
        "NON_GEOMETRIC_STATUS": state["non_geometric"]["status"],
        "NON_GEOMETRIC_VALUE": state["non_geometric"]["value"],
        "COMPONENT_LAYER": state["layer"]["value"],
        "COMPONENT_LAYER_STATE_STATUS": state["layer_state"]["status"],
        "COMPONENT_LAYER_STATE_VALUE": state["layer_state"]["value"],
        "PROTOTYPE_TYPE": proto["runtime_type"],
        "LOAD_STATUS": proto["load_state"]["status"],
        "FULLY_LOADED_STATUS": proto["fully_loaded"]["status"],
        "FULLY_LOADED_VALUE": proto["fully_loaded"]["value"],
        "REFSET_PROBE_STATUS": refset["status"],
        "REFSET_FOUND": refset["found"],
        "REFSET_BODY_MEMBERS": refset["body_member_count"],
        "REFSET_COMPONENT_MEMBERS": refset["component_member_count"],
        "MAPPED_BODY_OCCURRENCES": mapping["mapped_body_count"],
        "MAPPED_COMPONENT_OCCURRENCES": mapping["mapped_component_count"],
        "MAPPED_BODIES_VISIBLE_CURRENT_VIEW": mapping["mapped_bodies_visible_current_view_count"],
        "CURRENT_VIEW_NAME": current_view_name,
        "PROBE_ERRORS": " | ".join(row["probe_errors"]),
    }
    return values


def write_outputs(folder, stem, report):
    csv_path = os.path.join(folder, stem + ".csv")
    json_path = os.path.join(folder, stem + ".json")
    with open(csv_path, "w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(handle, fieldnames=CSV_COLUMNS)
        writer.writeheader()
        for analysis in report["target_analyses"]:
            for index, row in enumerate(analysis["subtree_occurrences"]):
                writer.writerow(csv_row(
                    "TARGET" if index == 0 else "TARGET_DESCENDANT",
                    row,
                    analysis["current_work_view"]["name"],
                ))
    with open(json_path, "w", encoding="utf-8") as handle:
        json.dump(report, handle, indent=2, ensure_ascii=False)
    return csv_path, json_path


def run(session, run_datetime=None):
    now = run_datetime or datetime.datetime.now().astimezone()
    try:
        work_part, display_part = session.Parts.Work, session.Parts.Display
    except Exception:
        work_part, display_part = None, None
    if work_part is None:
        raise RuntimeError("Open the affected top-level HLA first.")
    if not same_object(work_part, display_part):
        raise RuntimeError("Make the affected HLA both work and displayed part.")
    if safe_value(safe_value(work_part, "ComponentAssembly", None), "RootComponent", None) is None:
        raise RuntimeError("The active part is not an HLA assembly.")

    nodes, traversal_errors = collect_nodes(work_part)
    request, targets = resolve_targets(nodes)
    if not targets:
        raise RuntimeError("J23 target was not found: {0}".format(request["value"]))
    analyses = [analyze_target(target, nodes, work_part) for target in targets]
    report = {
        "schema_version": SCHEMA_VERSION,
        "journal_build": BUILD,
        "run_timestamp": now.isoformat(timespec="seconds"),
        "scope": "READ_ONLY_EXACT_TARGET_ROOT_CAUSE_PROOF",
        "root_assembly": safe_name(work_part, "<HLA>"),
        "target_request": request,
        "target_match_count": len(targets),
        "assembly_occurrence_count": len(nodes),
        "traversal_errors": traversal_errors,
        "truth_policy": {
            "OBSERVED": "NX returned a value successfully.",
            "ERROR": "NX exposed the probe but it failed; no boolean default is inferred.",
            "UNAVAILABLE": "The runtime object does not expose the required probe.",
            "RULED_OUT": "Observed counter-evidence disproves the hypothesis as the primary cause.",
            "CONFIRMED": "An independent comparison completes the causal evidence chain.",
            "STRONGLY_SUPPORTED": "Evidence points to the cause, but a named missing read API prevents direct confirmation.",
        },
        "target_analyses": analyses,
    }
    folder = os.path.join(io_root(), OUTPUT_FOLDER)
    os.makedirs(folder, exist_ok=True)
    token = filename_token(request["value"])
    stem = "J23_EVIDENCE_{0}_{1}".format(token, now.strftime("%Y%m%d_%H%M%S"))
    csv_path, json_path = write_outputs(folder, stem, report)
    return csv_path, json_path, report


def main():
    session = NXOpen.Session.GetSession()
    log_line(session, "=" * 72)
    log_line(session, "J23 V2 TARGET-FOCUSED HLA VISIBILITY EVIDENCE")
    log_line(session, "Build: " + BUILD)
    log_line(session, "Read-only: no display, view, load, assembly, or save changes.")
    log_line(session, "=" * 72)
    try:
        csv_path, json_path, report = run(session)
        for analysis in report["target_analyses"]:
            conclusion = analysis["conclusion"]
            log_line(session, "Target: " + analysis["target"]["assembly_path"])
            log_line(session, "Conclusion: {0} / {1}".format(
                conclusion["status"], conclusion["root_cause_code"]
            ))
            log_line(session, conclusion["statement"])
            for item in analysis["hypotheses"]:
                log_line(session, "  {0}: {1}".format(item["code"], item["verdict"]))
        log_line(session, "CSV: " + csv_path)
        log_line(session, "JSON: " + json_path)
        log_line(session, "Return the JSON; every conclusion cites its fact IDs.")
    except Exception as error:
        log_line(session, "J23 V2 FAILED: " + error_text(error))
        log_line(session, traceback.format_exc())
        raise


def get_unload_option(dummy):
    return NXOpen.Session.LibraryUnloadOption.Immediately


if __name__ == "__main__":
    main()
