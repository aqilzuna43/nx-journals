"""J24 - guarded repair and causal test for missing HLA isolate-view geometry.

Select exactly one missing component occurrence in the Assembly Navigator, then
play this journal in the affected top-level HLA window.  J24 adds that exact
occurrence and its subtree to the current NX isolate view through the supported
ComponentAssembly.ShowComponentsInIsolateView API.  It records exact mapped-body
visibility before and after the call.

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


BUILD = "J24-NX2506-HLA-ISOLATE-REPAIR-V1"
SCHEMA_VERSION = 1
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
    report = {
        "schema_version": SCHEMA_VERSION,
        "journal_build": BUILD,
        "run_timestamp": now.isoformat(timespec="seconds"),
        "scope": "DISPLAY_ONLY_CURRENT_ISOLATE_VIEW_MUTATION_NO_SAVE",
        "root_assembly": j23.safe_name(work_part, "<HLA>"),
        "target_request": request,
        "target": target["identity"],
        "traversal_errors": traversal_errors,
        "work_view": {"name": view_name, "tag": j23.object_tag(work_view)},
        "mapped_target_body_count": len(mapped_tags),
        "before": before,
        "action": {
            "api": "ComponentAssembly.ShowComponentsInIsolateView",
            "attempted": False,
            "component_scope": "EXACT_TARGET_AND_COMPLETE_SUBTREE",
            "component_count": len(subtree_nodes),
            "component_tags": [node["identity"]["component_tag"] for node in subtree_nodes],
            "error_list": [],
            "exception": "",
            "refresh_probes": [],
            "undo_mark_name": UNDO_MARK_NAME,
        },
        "after": None,
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
    components = [node["_component"] for node in subtree_nodes]
    try:
        errors = component_assembly.ShowComponentsInIsolateView(components, work_view)
        report["action"]["error_list"] = error_list_evidence(errors, j23)
        report["action"]["refresh_probes"] = refresh_display(
            session, work_part, work_view, j23
        )
        after = visible_snapshot(work_view, mapped_tags, j23)
        report["after"] = after
    except Exception as error:
        report["action"]["exception"] = j23.error_text(error)
        report["rollback"] = rollback(session, undo_mark, j23)
        report["verdict"] = {
            "status": "API_ERROR",
            "root_cause_code": "SHOW_COMPONENTS_IN_ISOLATE_VIEW_FAILED",
            "statement": "NX rejected the supported isolate-view show operation; the undo rollback was attempted.",
        }
        return write_report(report, request["value"], now, j23), report

    after_count = report["after"]["mapped_target_count_visible"]
    if report["after"]["probe_status"] == j23.OBSERVED and after_count > 0:
        error_count = len(report["action"]["error_list"])
        report["verdict"] = {
            "status": "CONFIRMED" if error_count == 0 else "CONFIRMED_WITH_API_WARNINGS",
            "root_cause_code": "ISOLATE_VIEW_MEMBERSHIP_EXCLUDED_TARGET",
            "statement": (
                "Adding the exact target subtree to the current isolate view changed mapped-body visibility "
                "from 0 to {0}; isolation membership is the confirmed cause."
            ).format(after_count),
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
