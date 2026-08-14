"""J24 - guarded repair and causal test for missing HLA isolate-view geometry.

Select exactly one missing component occurrence in the Assembly Navigator, then
play this journal in the affected top-level HLA window. J24 first adds the exact
selected parent and, only if necessary, then adds its unsuppressed descendants
that have mapped body occurrences. It uses the NX Python one-input form of
ComponentAssembly.ShowComponentsInIsolateView and records the returned output
objects instead of assuming the C# out-parameter is a Python input.

The operation is display-only and is never saved by this journal.  A visible NX
undo mark named "J24 Show Target In Isolate View" is created before mutation.
If the call fails or produces no target-body visibility change, J24 attempts to
roll back to that mark automatically.

Target: NX 2312 and NX X 2506 embedded Python.
Run via: NX > Tools > Journal > Play
"""

import datetime
import importlib.util
import json
import os
import traceback

import NXOpen


BUILD = "J24-NX2506-HLA-ISOLATE-REPAIR-V2"
SCHEMA_VERSION = 2
UNDO_MARK_NAME = "J24 Show Target In Isolate View"
ISOLATE_NAME_TOKEN = "ISOLATE"


def load_j23():
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "23_diagnose_hla_visibility.py")
    spec = importlib.util.spec_from_file_location("j23_hla_visibility_dependency", path)
    if spec is None or spec.loader is None:
        raise RuntimeError("Cannot load J23 visibility dependency: " + path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def error_list_evidence(error_list, j23):
    if error_list is None:
        return []
    rows = []
    try:
        length = int(getattr(error_list, "Length", 0) or 0)
    except Exception as error:
        rows.append({
            "index": None,
            "error_code": "",
            "description": "Cannot read ErrorList.Length: " + j23.error_text(error),
            "error_object_tag": "",
            "error_object_description": "",
        })
        length = 0
    for index in range(length):
        try:
            info = error_list.GetErrorInfo(index)
            rows.append({
                "index": index,
                "error_code": j23.clean(getattr(info, "ErrorCode", "")),
                "description": j23.clean(getattr(info, "Description", "")),
                "error_object_tag": j23.object_tag(getattr(info, "ErrorObject", None)),
                "error_object_description": j23.clean(
                    getattr(info, "ErrorObjectDescription", "")
                ),
            })
        except Exception as error:
            rows.append({
                "index": index,
                "error_code": "",
                "description": "Cannot read error entry: " + j23.error_text(error),
                "error_object_tag": "",
                "error_object_description": "",
            })
    try:
        error_list.FreeResource()
    except Exception:
        pass
    return rows


def exposes(value, member_name):
    if value is None:
        return False
    try:
        getattr(value, member_name)
        return True
    except Exception:
        return False


def returned_object_evidence(value, j23):
    return {
        "runtime_type": j23.object_kind(value),
        "tag": j23.object_tag(value),
        "name": j23.safe_name(value, ""),
        "has_error_list_shape": exposes(value, "GetErrorInfo") and exposes(value, "Length"),
        "has_view_shape": exposes(value, "AskVisibleObjects"),
    }


def normalize_isolate_result(result, j23):
    """Identify ErrorList- and View-shaped values without assuming tuple order."""
    if isinstance(result, (tuple, list)):
        values = list(result)
        shape = "SEQUENCE"
    else:
        values = [result]
        shape = "SINGLE"
    error_list = None
    output_view = None
    for value in values:
        if value is None:
            continue
        if error_list is None and exposes(value, "GetErrorInfo") and exposes(value, "Length"):
            error_list = value
        if output_view is None and exposes(value, "AskVisibleObjects"):
            output_view = value
    for container in values:
        if container is None:
            continue
        if error_list is None:
            for member_name in ("ErrorList", "errorList", "Errors", "errors"):
                candidate = j23.safe_value(container, member_name, None)
                if candidate is not None and exposes(candidate, "GetErrorInfo"):
                    error_list = candidate
                    break
        if output_view is None:
            for member_name in ("View", "view", "OutputView", "outputView"):
                candidate = j23.safe_value(container, member_name, None)
                if candidate is not None and exposes(candidate, "AskVisibleObjects"):
                    output_view = candidate
                    break
    evidence = {
        "shape": shape,
        "runtime_type": j23.object_kind(result),
        "item_count": len(values),
        "items": [returned_object_evidence(value, j23) for value in values],
        "error_list_detected": error_list is not None,
        "output_view_detected": output_view is not None,
    }
    return error_list, output_view, evidence


def view_identity(view, j23):
    if view is None:
        return {"available": False, "name": "", "tag": ""}
    return {
        "available": True,
        "name": j23.safe_name(view, "<view>"),
        "tag": j23.object_tag(view),
    }


def displayed_view_context(work_part, j23):
    result = {
        "work_view": view_identity(
            getattr(getattr(work_part, "ModelingViews", None), "WorkView", None), j23
        ),
        "active_views_status": j23.UNAVAILABLE,
        "active_views": [],
        "active_views_error": "",
        "layout_status": j23.UNAVAILABLE,
        "current_layout": {"name": "", "tag": ""},
        "layout_views": [],
        "layout_error": "",
    }
    views = j23.safe_value(work_part, "Views", None)
    if views is not None:
        active = j23.method_probe(views, "GetActiveViews")
        result["active_views_status"] = active["status"]
        result["active_views_error"] = active["error"]
        if active["status"] == j23.OBSERVED:
            result["active_views"] = [
                view_identity(view, j23) for view in list(active["value"])
            ]
    layouts = j23.safe_value(work_part, "Layouts", None)
    current = j23.safe_value(layouts, "Current", None) if layouts is not None else None
    if current is not None:
        result["current_layout"] = {
            "name": j23.safe_name(current, "<layout>"),
            "tag": j23.object_tag(current),
        }
        layout_views = j23.method_probe(current, "GetViews")
        result["layout_status"] = layout_views["status"]
        result["layout_error"] = layout_views["error"]
        if layout_views["status"] == j23.OBSERVED:
            result["layout_views"] = [
                view_identity(view, j23) for view in list(layout_views["value"])
            ]
    return result


def visible_snapshot(view, mapped_tags, j23):
    result = j23.method_probe(view, "AskVisibleObjects")
    if result["status"] != j23.OBSERVED:
        return {
            "probe_status": result["status"],
            "visible_object_count": None,
            "mapped_target_tags_visible": [],
            "mapped_target_count_visible": None,
            "error": result["error"],
        }
    visible_tags = {
        j23.object_tag(item) for item in result["value"] if j23.object_tag(item)
    }
    target_visible = sorted(set(mapped_tags) & visible_tags)
    return {
        "probe_status": j23.OBSERVED,
        "visible_object_count": len(result["value"]),
        "mapped_target_tags_visible": target_visible,
        "mapped_target_count_visible": len(target_visible),
        "error": "",
    }


def refresh_display(session, work_part, view, j23):
    probes = []
    candidates = (
        (getattr(session, "DisplayManager", None), "MakeUpToDate"),
        (view, "Regenerate"),
        (getattr(work_part, "Views", None), "UpdateDisplay"),
    )
    for value, method_name in candidates:
        if value is None:
            probes.append({
                "status": j23.UNAVAILABLE,
                "source": method_name,
                "error": "Runtime object is unavailable.",
            })
            continue
        result = j23.method_probe(value, method_name)
        probes.append({
            "status": result["status"],
            "source": result["source"],
            "error": result["error"],
        })
    return probes


def run_show_stage(label, components, component_assembly, session, work_part, mapped_tags, j23):
    stage = {
        "label": label,
        "api": "ComponentAssembly.ShowComponentsInIsolateView",
        "python_signature": "ShowComponentsInIsolateView(components)",
        "attempted": False,
        "component_count": len(components),
        "component_tags": [j23.object_tag(component) for component in components],
        "return_evidence": None,
        "error_list": [],
        "exception": "",
        "refresh_probes": [],
        "work_view_after": None,
        "returned_view": None,
        "returned_view_visibility": None,
        "maximum_mapped_target_count_visible": None,
    }
    if not components:
        stage["exception"] = "No eligible components for this stage."
        return stage
    stage["attempted"] = True
    try:
        # The real NX X 2506 V1 artifact proved that passing the documented
        # C# `out View` as a second Python input raises "Function takes 1
        # arguments, 2 passed" before the operation executes. Inspect the
        # one-input call's returned runtime values without assuming their order.
        raw_result = component_assembly.ShowComponentsInIsolateView(components)
        error_list, returned_view, return_evidence = normalize_isolate_result(
            raw_result, j23
        )
        stage["return_evidence"] = return_evidence
        stage["error_list"] = error_list_evidence(error_list, j23)
        current_view = work_part.ModelingViews.WorkView
        stage["refresh_probes"] = refresh_display(
            session, work_part, current_view, j23
        )
        current_view = work_part.ModelingViews.WorkView
        stage["work_view_after"] = {
            "identity": view_identity(current_view, j23),
            "visibility": visible_snapshot(current_view, mapped_tags, j23),
        }
        stage["returned_view"] = view_identity(returned_view, j23)
        if returned_view is not None:
            stage["returned_view_visibility"] = visible_snapshot(
                returned_view, mapped_tags, j23
            )
        counts = []
        work_count = stage["work_view_after"]["visibility"][
            "mapped_target_count_visible"
        ]
        if work_count is not None:
            counts.append(work_count)
        if stage["returned_view_visibility"] is not None:
            returned_count = stage["returned_view_visibility"][
                "mapped_target_count_visible"
            ]
            if returned_count is not None:
                counts.append(returned_count)
        stage["maximum_mapped_target_count_visible"] = max(counts) if counts else None
    except Exception as error:
        stage["exception"] = j23.error_text(error)
    return stage


def mapped_unsuppressed_descendants(target, subtree_nodes, target_analysis, j23):
    target_tag = target["identity"]["component_tag"]
    eligible_tags = {
        row["identity"]["component_tag"]
        for row in target_analysis["subtree_occurrences"]
        if row["identity"]["component_tag"] != target_tag
        and row["mapping"]["mapped_body_count"] > 0
        and j23.boolean_observed(row["component_state"]["suppressed"], False)
    }
    return [
        node["_component"]
        for node in subtree_nodes
        if node["identity"]["component_tag"] in eligible_tags
    ]


def rollback(session, undo_mark, j23):
    try:
        session.UndoToMark(undo_mark, UNDO_MARK_NAME)
        return {"attempted": True, "status": "ROLLED_BACK", "error": ""}
    except Exception as error:
        return {
            "attempted": True,
            "status": "ROLLBACK_FAILED",
            "error": j23.error_text(error),
        }


def capture_after_state(report, work_part, mapped_tags, j23):
    current_view = work_part.ModelingViews.WorkView
    report["after"] = visible_snapshot(current_view, mapped_tags, j23)
    report["view_context_after"] = displayed_view_context(work_part, j23)


def write_report(report, request_value, now, j23):
    folder = os.path.join(j23.io_root(), j23.OUTPUT_FOLDER)
    os.makedirs(folder, exist_ok=True)
    stem = "J24_ISOLATE_REPAIR_{0}_{1}".format(
        j23.filename_token(request_value), now.strftime("%Y%m%d_%H%M%S")
    )
    path = os.path.join(folder, stem + ".json")
    with open(path, "w", encoding="utf-8") as handle:
        json.dump(report, handle, indent=2, ensure_ascii=False)
    return path


def run(session, run_datetime=None, dependency=None):
    j23 = dependency or load_j23()
    now = run_datetime or datetime.datetime.now().astimezone()
    work_part = getattr(session.Parts, "Work", None)
    display_part = getattr(session.Parts, "Display", None)
    if work_part is None or not j23.same_object(work_part, display_part):
        raise RuntimeError("Make the affected top-level HLA both work and displayed part.")
    component_assembly = getattr(work_part, "ComponentAssembly", None)
    if component_assembly is None or getattr(component_assembly, "RootComponent", None) is None:
        raise RuntimeError("The active part is not an HLA assembly.")

    nodes, traversal_errors = j23.collect_nodes(work_part)
    request, targets = j23.resolve_targets(nodes)
    if len(targets) != 1:
        raise RuntimeError(
            "J24 requires exactly one target occurrence; resolved {0}. "
            "Preselect one missing component in Assembly Navigator.".format(len(targets))
        )
    target = targets[0]
    target_tag = target["identity"]["component_tag"]
    subtree_nodes = [
        node for node in nodes
        if node is target or target_tag in node["_ancestor_tags"]
    ]
    target_analysis = j23.analyze_target(target, nodes, work_part)
    mapped_tags = sorted({
        tag
        for row in target_analysis["subtree_occurrences"]
        for tag in row["mapping"]["mapped_body_occurrence_tags"]
    })
    work_view = work_part.ModelingViews.WorkView
    view_name = j23.safe_name(work_view, "<work view>")
    before = visible_snapshot(work_view, mapped_tags, j23)
    descendant_components = mapped_unsuppressed_descendants(
        target, subtree_nodes, target_analysis, j23
    )
    report = {
        "schema_version": SCHEMA_VERSION,
        "journal_build": BUILD,
        "run_timestamp": now.isoformat(timespec="seconds"),
        "scope": "DISPLAY_ONLY_CURRENT_ISOLATE_VIEW_MUTATION_NO_SAVE",
        "root_assembly": j23.safe_name(work_part, "<HLA>"),
        "target_request": request,
        "target": target["identity"],
        "traversal_errors": traversal_errors,
        "view_context_before": displayed_view_context(work_part, j23),
        "work_view": {"name": view_name, "tag": j23.object_tag(work_view)},
        "mapped_target_body_count": len(mapped_tags),
        "before": before,
        "action": {
            "api": "ComponentAssembly.ShowComponentsInIsolateView",
            "attempted": False,
            "binding_evidence": {
                "failed_v1_signature": "ShowComponentsInIsolateView(components, work_view)",
                "failed_v1_error": "Function takes 1 arguments, 2 passed.",
                "v2_signature": "ShowComponentsInIsolateView(components)",
                "out_view_policy": "Inspect possible returned view-shaped values; never pass the C# out View as a Python input.",
            },
            "strategy": "TARGET_PARENT_THEN_UNSUPPRESSED_MAPPED_DESCENDANTS",
            "eligible_descendant_count": len(descendant_components),
            "eligible_descendant_tags": [
                j23.object_tag(component) for component in descendant_components
            ],
            "stages": [],
            "exception": "",
            "undo_mark_name": UNDO_MARK_NAME,
        },
        "after": None,
        "view_context_after": None,
        "rollback": {"attempted": False, "status": "NOT_REQUIRED", "error": ""},
        "verdict": None,
    }

    if ISOLATE_NAME_TOKEN not in view_name.upper():
        report["verdict"] = {
            "status": "NOT_APPLIED",
            "root_cause_code": "ACTIVE_VIEW_NOT_ISOLATE_NAMED",
            "statement": "J24 refused to mutate a work view whose name does not contain 'Isolate'.",
        }
        return write_report(report, request["value"], now, j23), report
    if before["probe_status"] != j23.OBSERVED:
        report["verdict"] = {
            "status": "NOT_APPLIED",
            "root_cause_code": "BEFORE_VISIBILITY_PROBE_FAILED",
            "statement": "J24 could not establish the target visibility baseline.",
        }
        return write_report(report, request["value"], now, j23), report
    if not mapped_tags:
        report["verdict"] = {
            "status": "NOT_APPLIED",
            "root_cause_code": "NO_MAPPED_TARGET_GEOMETRY",
            "statement": "No mapped target body tags exist, so isolate membership cannot be tested.",
        }
        return write_report(report, request["value"], now, j23), report
    if before["mapped_target_count_visible"] > 0:
        report["verdict"] = {
            "status": "NOT_APPLIED",
            "root_cause_code": "TARGET_ALREADY_VISIBLE",
            "statement": "Mapped target geometry is already present in the current work view.",
        }
        return write_report(report, request["value"], now, j23), report

    try:
        undo_mark = session.SetUndoMark(
            NXOpen.Session.MarkVisibility.Visible, UNDO_MARK_NAME
        )
    except Exception as error:
        report["action"]["exception"] = j23.error_text(error)
        report["verdict"] = {
            "status": "NOT_APPLIED",
            "root_cause_code": "UNDO_GUARD_UNAVAILABLE",
            "statement": "J24 refused to mutate the view because it could not create an NX undo mark.",
        }
        return write_report(report, request["value"], now, j23), report

    report["action"]["attempted"] = True
    parent_stage = run_show_stage(
        "TARGET_PARENT",
        [target["_component"]],
        component_assembly,
        session,
        work_part,
        mapped_tags,
        j23,
    )
    report["action"]["stages"].append(parent_stage)
    parent_count = parent_stage["maximum_mapped_target_count_visible"]
    if parent_stage["exception"]:
        report["action"]["exception"] = parent_stage["exception"]
        report["rollback"] = rollback(session, undo_mark, j23)
        capture_after_state(report, work_part, mapped_tags, j23)
        report["verdict"] = {
            "status": "API_ERROR",
            "root_cause_code": "SHOW_COMPONENTS_IN_ISOLATE_VIEW_FAILED",
            "statement": "The NX Python one-input isolate-view show call failed before a visibility comparison; rollback was attempted.",
        }
        return write_report(report, request["value"], now, j23), report

    if not parent_count:
        descendant_stage = run_show_stage(
            "UNSUPPRESSED_MAPPED_DESCENDANTS",
            descendant_components,
            component_assembly,
            session,
            work_part,
            mapped_tags,
            j23,
        )
        report["action"]["stages"].append(descendant_stage)
        if descendant_stage["attempted"] and descendant_stage["exception"]:
            report["action"]["exception"] = descendant_stage["exception"]
            report["rollback"] = rollback(session, undo_mark, j23)
            capture_after_state(report, work_part, mapped_tags, j23)
            report["verdict"] = {
                "status": "API_ERROR",
                "root_cause_code": "SHOW_MAPPED_DESCENDANTS_IN_ISOLATE_VIEW_FAILED",
                "statement": "The parent call completed without restoring geometry, then the mapped-descendant call failed; rollback was attempted.",
            }
            return write_report(report, request["value"], now, j23), report

    capture_after_state(report, work_part, mapped_tags, j23)
    stage_counts = [
        stage["maximum_mapped_target_count_visible"]
        for stage in report["action"]["stages"]
        if stage["maximum_mapped_target_count_visible"] is not None
    ]
    after_count = max(stage_counts + [
        report["after"]["mapped_target_count_visible"]
        if report["after"]["mapped_target_count_visible"] is not None else 0
    ])
    if report["after"]["probe_status"] == j23.OBSERVED and after_count > 0:
        error_count = sum(
            len(stage["error_list"]) for stage in report["action"]["stages"]
        )
        successful_stage = next(
            stage["label"] for stage in report["action"]["stages"]
            if (stage["maximum_mapped_target_count_visible"] or 0) > 0
        )
        report["verdict"] = {
            "status": "CONFIRMED" if error_count == 0 else "CONFIRMED_WITH_API_WARNINGS",
            "root_cause_code": "ISOLATE_VIEW_MEMBERSHIP_EXCLUDED_TARGET",
            "statement": (
                "The {0} stage changed exact mapped-body visibility from 0 to {1}; "
                "isolation membership is the confirmed cause."
            ).format(successful_stage, after_count),
        }
    else:
        report["rollback"] = rollback(session, undo_mark, j23)
        report["verdict"] = {
            "status": "INCONCLUSIVE",
            "root_cause_code": "ISOLATE_SHOW_DID_NOT_RESTORE_MAPPED_GEOMETRY",
            "statement": "The isolate-view show API did not make any mapped target bodies visible; J24 attempted rollback.",
        }
    return write_report(report, request["value"], now, j23), report


def main():
    session = NXOpen.Session.GetSession()
    j23 = load_j23()
    j23.log_line(session, "=" * 72)
    j23.log_line(session, "J24 GUARDED HLA ISOLATE-VIEW TARGET REPAIR")
    j23.log_line(session, "Build: " + BUILD)
    j23.log_line(session, "Display-only; no save. Successful changes remain under one visible undo mark.")
    j23.log_line(session, "=" * 72)
    try:
        json_path, report = run(session, dependency=j23)
        verdict = report["verdict"]
        j23.log_line(session, "Target: " + report["target"]["assembly_path"])
        j23.log_line(session, "View: " + report["work_view"]["name"])
        j23.log_line(session, "Verdict: {0} / {1}".format(
            verdict["status"], verdict["root_cause_code"]
        ))
        j23.log_line(session, verdict["statement"])
        j23.log_line(session, "JSON: " + json_path)
        j23.log_line(session, "If needed, Undo once to revert: " + UNDO_MARK_NAME)
    except Exception as error:
        j23.log_line(session, "J24 FAILED: " + j23.error_text(error))
        j23.log_line(session, traceback.format_exc())
        raise


def get_unload_option(dummy):
    return NXOpen.Session.LibraryUnloadOption.Immediately


if __name__ == "__main__":
    main()
