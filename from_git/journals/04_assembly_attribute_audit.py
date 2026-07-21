"""Journal 04 - NX-authoritative attribute and BOM certification gate.

The journal never changes or saves NX data.  It audits the active model tree,
applicable Teamcenter drawing specifications, FZ-required metadata, and an
optional downstream MASTER export.  A certified BOM is emitted only with zero
blocking findings.
"""

import os
import sys
from collections import Counter
from datetime import datetime

import NXOpen
import NXOpen.UF

_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _REPO_ROOT not in sys.path:
    sys.path.insert(0, _REPO_ROOT)

from utils.attribute_reconciliation import (  # noqa: E402
    FZ_BOM_COLUMNS,
    ReconciliationError,
    bom_export_rows,
    clean_text,
    collect_bom_nodes,
    compare_master_reference,
    config_path,
    config_sha256,
    drawing_decision,
    exact_model_dimensions,
    file_sha256,
    identity_from_attributes,
    load_config,
    load_drawing_scope,
    mass_density_volume_consistent,
    new_run_id,
    normalize_value,
    normalized_text,
    object_identifier,
    object_key,
    read_attributes,
    rules_by_logical,
    validate_attribute_value,
    write_csv,
)
from utils.nx_helpers import (  # noqa: E402
    get_input_folder,
    get_output_folder,
    log_info,
    require_work_part,
    run_journal,
)


DETAIL_COLUMNS = [
    "RUN_ID",
    "RUN_TIMESTAMP",
    "CONFIG_SHA256",
    "DRAWING_SCOPE_SHA256",
    "MASTER_REFERENCE_SHA256",
    "OCCURRENCE_PATH",
    "PARENT_PART_NUMBER",
    "PARENT_REVISION",
    "PART_NUMBER",
    "REVISION",
    "BOM_LEVEL",
    "NX_QUANTITY",
    "MODEL_IDENTIFIER",
    "DRAWING_REQUIRED",
    "DRAWING_INDEX",
    "DRAWING_IDENTIFIER",
    "LOGICAL_ATTRIBUTE",
    "CATEGORY",
    "NX_ATTRIBUTE_NAME",
    "ATTRIBUTE_TYPE",
    "AUTHORITATIVE_SOURCE",
    "EXPECTED_VALUE",
    "MODEL_VALUE",
    "DRAWING_VALUE",
    "NORMALIZED_EXPECTED",
    "NORMALIZED_MODEL",
    "NORMALIZED_DRAWING",
    "SEVERITY",
    "COMPARISON_RESULT",
    "FAILURE_CODE",
    "NX_EXCEPTION_TYPE",
    "NX_ERROR_CODE",
    "MESSAGE",
]


def _env_path(name, default_name):
    value = clean_text(os.environ.get(name))
    return value if value else os.path.join(get_input_folder(), default_name)


def _dispose(value):
    if value is None:
        return
    for method_name in ("Dispose", "FreeResource"):
        try:
            getattr(value, method_name)()
            return
        except Exception:
            pass


def _unwrap(value):
    if isinstance(value, tuple):
        first = value[0] if value else None
        for extra in value[1:]:
            _dispose(extra)
        return first
    return value


def _exception_details(error):
    code = getattr(error, "ErrorCode", None)
    if code is None:
        return "{0}: {1}".format(type(error).__name__, error)
    return "{0}: {1} (ErrorCode={2})".format(type(error).__name__, error, code)


def _exception_fields(error):
    code = getattr(error, "ErrorCode", "")
    return {
        "exception_type": type(error).__name__,
        "error_code": "" if code is None else code,
    }


def _loaded_parts(session):
    try:
        return list(session.Parts)
    except Exception:
        try:
            return list(session.Parts.ToArray())
        except Exception:
            return []


def _journal_identifier(part):
    try:
        return clean_text(part.JournalIdentifier)
    except Exception:
        return object_identifier(part)


def _find_loaded(session, expected_identifier):
    expected = normalized_text(expected_identifier)
    for part in _loaded_parts(session):
        if normalized_text(_journal_identifier(part)) == expected:
            return part
    return None


def _drawing_identifier(config, part_number, revision, index):
    return config["drawing"]["identifier_template"].format(
        part_number=part_number, revision=revision, index=index
    )


def _open_drawing(session, config, part_number, revision, index):
    expected = _drawing_identifier(config, part_number, revision, index)
    loaded = _find_loaded(session, expected)
    if loaded is not None:
        return loaded, False, ""
    try:
        opened = _unwrap(session.Parts.OpenDisplay(expected))
        if opened is None:
            return None, False, "OpenDisplay returned no part for {0}.".format(expected)
        if normalized_text(_journal_identifier(opened)) != normalized_text(expected):
            return opened, True, "OpenDisplay returned a different JournalIdentifier."
        return opened, True, ""
    except Exception as exc:
        return None, False, _exception_details(exc)


def _sheet_count(part):
    try:
        return len(list(part.DrawingSheets))
    except Exception:
        try:
            return int(part.DrawingSheets.Count)
        except Exception:
            return -1


def _restore_state(session, display_part, work_part):
    messages = []
    if display_part is not None:
        try:
            _dispose(session.Parts.SetDisplay(display_part, False, True))
        except Exception as exc:
            messages.append("Display restore failed: " + _exception_details(exc))
    if work_part is not None:
        try:
            session.Parts.SetWork(work_part)
        except Exception as exc:
            messages.append("Work-part restore failed: " + _exception_details(exc))
    return messages


def _close_opened(part):
    try:
        part.Close(
            NXOpen.BasePart.CloseWholeTree.FalseValue,
            NXOpen.BasePart.CloseModified.CloseModified,
            None,
        )
        return ""
    except Exception as exc:
        return _exception_details(exc)


def _base_detail(context, node=None):
    node = node or {}
    return {
        "RUN_ID": context["run_id"],
        "RUN_TIMESTAMP": context["timestamp"],
        "CONFIG_SHA256": context["config_hash"],
        "DRAWING_SCOPE_SHA256": context.get("scope_hash", ""),
        "MASTER_REFERENCE_SHA256": context.get("master_hash", ""),
        "OCCURRENCE_PATH": node.get("occurrence_path", ""),
        "PARENT_PART_NUMBER": node.get("parent_part_number", ""),
        "PARENT_REVISION": node.get("parent_revision", ""),
        "PART_NUMBER": node.get("part_number", ""),
        "REVISION": node.get("revision", ""),
        "BOM_LEVEL": node.get("level", ""),
        "NX_QUANTITY": node.get("quantity", ""),
        "MODEL_IDENTIFIER": object_identifier(node.get("part")) if node.get("part") is not None else "",
        "DRAWING_REQUIRED": "",
        "DRAWING_INDEX": "",
        "DRAWING_IDENTIFIER": "",
        "LOGICAL_ATTRIBUTE": "",
        "CATEGORY": "",
        "NX_ATTRIBUTE_NAME": "",
        "ATTRIBUTE_TYPE": "",
        "AUTHORITATIVE_SOURCE": "",
        "EXPECTED_VALUE": "",
        "MODEL_VALUE": "",
        "DRAWING_VALUE": "",
        "NORMALIZED_EXPECTED": "",
        "NORMALIZED_MODEL": "",
        "NORMALIZED_DRAWING": "",
        "SEVERITY": "BLOCK",
        "COMPARISON_RESULT": "FAIL",
        "FAILURE_CODE": "ERROR",
        "NX_EXCEPTION_TYPE": "",
        "NX_ERROR_CODE": "",
        "MESSAGE": "",
    }


def _finding_detail(context, finding, node=None):
    row = _base_detail(context, node)
    row.update(
        {
            "OCCURRENCE_PATH": finding.get("occurrence_path", row["OCCURRENCE_PATH"]),
            "PART_NUMBER": finding.get("part_number", row["PART_NUMBER"]),
            "REVISION": finding.get("revision", row["REVISION"]),
            "FAILURE_CODE": finding.get("code", "ERROR"),
            "NX_EXCEPTION_TYPE": finding.get("exception_type", ""),
            "NX_ERROR_CODE": finding.get("error_code", ""),
            "MESSAGE": finding.get("message", ""),
        }
    )
    return row


def _attribute_detail(context, node, rule, result):
    severity, code, message = validate_attribute_value(result, rule, context["config"])
    raw = result.get("raw_value", "")
    row = _base_detail(context, node)
    row.update(
        {
            "LOGICAL_ATTRIBUTE": rule["logical_name"],
            "CATEGORY": rule["category"],
            "NX_ATTRIBUTE_NAME": rule["attribute"],
            "ATTRIBUTE_TYPE": result.get("type", rule["type"]),
            "AUTHORITATIVE_SOURCE": rule["source_owner"],
            "EXPECTED_VALUE": raw,
            "MODEL_VALUE": raw,
            "NORMALIZED_EXPECTED": normalize_value(raw, rule, context["config"]),
            "NORMALIZED_MODEL": normalize_value(raw, rule, context["config"]),
            "SEVERITY": severity,
            "COMPARISON_RESULT": "PASS" if severity == "PASS" else severity,
            "FAILURE_CODE": code,
            "MESSAGE": message,
        }
    )
    return row


def _drawing_attribute_detail(context, node, rule, model_result, drawing_result, decision, index, drawing):
    row = _base_detail(context, node)
    model_raw = model_result.get("raw_value", "")
    drawing_raw = drawing_result.get("raw_value", "")
    row.update(
        {
            "DRAWING_REQUIRED": decision,
            "DRAWING_INDEX": index,
            "DRAWING_IDENTIFIER": _journal_identifier(drawing),
            "LOGICAL_ATTRIBUTE": rule["logical_name"],
            "CATEGORY": rule["category"],
            "NX_ATTRIBUTE_NAME": rule["attribute"],
            "ATTRIBUTE_TYPE": drawing_result.get("type", rule["type"]),
            "AUTHORITATIVE_SOURCE": rule["source_owner"],
            "EXPECTED_VALUE": model_raw,
            "MODEL_VALUE": model_raw,
            "DRAWING_VALUE": drawing_raw,
            "NORMALIZED_EXPECTED": normalize_value(model_raw, rule, context["config"]),
            "NORMALIZED_MODEL": normalize_value(model_raw, rule, context["config"]),
            "NORMALIZED_DRAWING": normalize_value(drawing_raw, rule, context["config"]),
        }
    )
    if drawing_result.get("status") in ("MISSING", "UNSET", "BLANK"):
        if rule.get("required_for_certification"):
            row.update(
                SEVERITY="BLOCK",
                COMPARISON_RESULT="FAIL",
                FAILURE_CODE="ATTRIBUTE_MISSING_DRAWING",
                MESSAGE="Required drawing attribute is not populated.",
            )
        else:
            row.update(
                SEVERITY="INFO",
                COMPARISON_RESULT="NOT_APPLICABLE",
                FAILURE_CODE="NOT_APPLICABLE",
                MESSAGE="Optional drawing attribute is not populated.",
            )
        return row
    if drawing_result.get("status") == "UNREADABLE":
        row.update(
            SEVERITY="BLOCK",
            COMPARISON_RESULT="FAIL",
            FAILURE_CODE="UNREADABLE_ATTRIBUTE",
            MESSAGE=drawing_result.get("message", "Drawing attribute could not be read."),
        )
        return row
    if normalize_value(model_raw, rule, context["config"]) != normalize_value(
        drawing_raw, rule, context["config"]
    ):
        row.update(
            SEVERITY="BLOCK",
            COMPARISON_RESULT="FAIL",
            FAILURE_CODE="MODEL_DRAWING_MISMATCH",
            MESSAGE="Drawing value does not match the NX-authoritative model value.",
        )
        return row
    row.update(
        SEVERITY="PASS",
        COMPARISON_RESULT="PASS",
        FAILURE_CODE="PASS",
        MESSAGE="Drawing value matches the NX-authoritative model value.",
    )
    return row


def _audit_models(context, nodes, uf_session):
    rows = []
    config = context["config"]
    rules = rules_by_logical(config)
    unique = OrderedDict()
    for node in nodes:
        key = (normalized_text(node["part_number"]), normalized_text(node["revision"]))
        unique.setdefault(key, node)
        for rule in config["attributes"]:
            if "MODEL" in rule.get("required_on", []):
                rows.append(_attribute_detail(context, node, rule, node["attributes"][rule["logical_name"]]))

    tolerance = config["release_policy"]["mass_relative_tolerance"]
    for node in unique.values():
        values = node["attributes"]
        consistent, calculated = mass_density_volume_consistent(
            values.get("mass_kg", {}).get("raw_value"),
            values.get("density_kg_per_mm3", {}).get("raw_value"),
            values.get("volume_mm3", {}).get("raw_value"),
            tolerance,
        )
        row = _base_detail(context, node)
        row.update(
            LOGICAL_ATTRIBUTE="mass_density_volume_consistency",
            CATEGORY="DERIVED_NX",
            NX_ATTRIBUTE_NAME="NX_Mass ~= NX_Density * NX_Volume",
            ATTRIBUTE_TYPE="Number",
            AUTHORITATIVE_SOURCE="DERIVED_NX",
            EXPECTED_VALUE=calculated,
            MODEL_VALUE=values.get("mass_kg", {}).get("raw_value", ""),
            SEVERITY="PASS" if consistent else "BLOCK",
            COMPARISON_RESULT="PASS" if consistent else "FAIL",
            FAILURE_CODE="PASS" if consistent else "ATTRIBUTE_MISMATCH",
            MESSAGE="Mass is consistent with density and volume." if consistent else (
                "Mass is not consistent with density and volume."
            ),
        )
        rows.append(row)
        try:
            dimensions = exact_model_dimensions(node["part"], uf_session)
            for axis, value in zip(("length_x_mm", "width_y_mm", "height_z_mm"), dimensions):
                dimension_row = _base_detail(context, node)
                dimension_row.update(
                    LOGICAL_ATTRIBUTE=axis,
                    CATEGORY="DERIVED_NX",
                    NX_ATTRIBUTE_NAME="EXACT_BOUNDING_BOX_" + axis.split("_")[1].upper(),
                    ATTRIBUTE_TYPE="Number",
                    AUTHORITATIVE_SOURCE="DERIVED_NX",
                    EXPECTED_VALUE=value,
                    MODEL_VALUE=value,
                    NORMALIZED_EXPECTED=format(value, ".15g"),
                    NORMALIZED_MODEL=format(value, ".15g"),
                    SEVERITY="PASS" if value > 0 else "BLOCK",
                    COMPARISON_RESULT="PASS" if value > 0 else "FAIL",
                    FAILURE_CODE="PASS" if value > 0 else "DIMENSION_DERIVATION_FAILED",
                    MESSAGE="Exact model-coordinate bounding-box extent." if value > 0 else (
                        "Bounding-box extent must be greater than zero."
                    ),
                )
                rows.append(dimension_row)
        except Exception as exc:
            dimension_row = _base_detail(context, node)
            dimension_row.update(
                LOGICAL_ATTRIBUTE="dimensions_xyz_mm",
                CATEGORY="DERIVED_NX",
                NX_ATTRIBUTE_NAME="EXACT_BOUNDING_BOX_XYZ",
                ATTRIBUTE_TYPE="Number",
                AUTHORITATIVE_SOURCE="DERIVED_NX",
                SEVERITY="BLOCK",
                COMPARISON_RESULT="FAIL",
                FAILURE_CODE="DIMENSION_DERIVATION_FAILED",
                NX_EXCEPTION_TYPE=type(exc).__name__,
                NX_ERROR_CODE="" if getattr(exc, "ErrorCode", None) is None else getattr(exc, "ErrorCode"),
                MESSAGE=_exception_details(exc),
            )
            rows.append(dimension_row)
    return rows, list(unique.values())


def _audit_drawings(session, context, unique_nodes, scope, opened_parts):
    rows = []
    config = context["config"]
    drawing_rules = [rule for rule in config["attributes"] if "DRAWING" in rule.get("required_on", [])]
    for node in unique_nodes:
        decision = drawing_decision(scope, node["part_number"], node["revision"])
        if decision != "YES":
            row = _base_detail(context, node)
            row.update(
                DRAWING_REQUIRED=decision,
                LOGICAL_ATTRIBUTE="drawing_scope",
                CATEGORY="GOVERNANCE",
                NX_ATTRIBUTE_NAME="Drawing Required",
                ATTRIBUTE_TYPE="String",
                AUTHORITATIVE_SOURCE="DATAPACK_SCOPE",
                EXPECTED_VALUE=decision,
                SEVERITY="INFO" if decision == "NO" else "BLOCK",
                COMPARISON_RESULT="NOT_APPLICABLE" if decision == "NO" else "FAIL",
                FAILURE_CODE="NOT_APPLICABLE" if decision == "NO" else "DRAWING_SCOPE_REVIEW",
                MESSAGE="Drawing is not required." if decision == "NO" else (
                    "Drawing applicability is missing or requires review."
                ),
            )
            rows.append(row)
            continue
        found = []
        open_errors = []
        for index in range(1, int(config["drawing"]["max_index"]) + 1):
            drawing, newly_opened, error = _open_drawing(
                session, config, node["part_number"], node["revision"], index
            )
            if newly_opened and drawing is not None and object_key(drawing) not in {
                object_key(part) for part in opened_parts
            }:
                opened_parts.append(drawing)
            if error:
                open_errors.append("dwg{0}: {1}".format(index, error))
                if drawing is not None:
                    identity_row = _base_detail(context, node)
                    identity_row.update(
                        DRAWING_REQUIRED="YES",
                        DRAWING_INDEX=index,
                        DRAWING_IDENTIFIER=_journal_identifier(drawing),
                        LOGICAL_ATTRIBUTE="drawing_resolution",
                        CATEGORY="DRAWING",
                        NX_ATTRIBUTE_NAME="JournalIdentifier",
                        ATTRIBUTE_TYPE="String",
                        AUTHORITATIVE_SOURCE="NX_TEAMCENTER_IDENTITY",
                        EXPECTED_VALUE=_drawing_identifier(
                            config, node["part_number"], node["revision"], index
                        ),
                        DRAWING_VALUE=_journal_identifier(drawing),
                        SEVERITY="BLOCK",
                        COMPARISON_RESULT="FAIL",
                        FAILURE_CODE="DRAWING_IDENTITY_MISMATCH",
                        MESSAGE=error,
                    )
                    rows.append(identity_row)
                continue
            if drawing is None:
                continue
            found.append((index, drawing))
        if not found:
            row = _base_detail(context, node)
            row.update(
                DRAWING_REQUIRED="YES",
                LOGICAL_ATTRIBUTE="drawing_resolution",
                CATEGORY="DRAWING",
                NX_ATTRIBUTE_NAME="JournalIdentifier",
                ATTRIBUTE_TYPE="String",
                AUTHORITATIVE_SOURCE="DATAPACK_SCOPE",
                SEVERITY="BLOCK",
                COMPARISON_RESULT="FAIL",
                FAILURE_CODE="MISSING_DRAWING",
                MESSAGE="No canonical drawing opened. " + " | ".join(open_errors),
            )
            rows.append(row)
            continue
        for index, drawing in found:
            sheet_count = _sheet_count(drawing)
            sheet_row = _base_detail(context, node)
            sheet_row.update(
                DRAWING_REQUIRED="YES",
                DRAWING_INDEX=index,
                DRAWING_IDENTIFIER=_journal_identifier(drawing),
                LOGICAL_ATTRIBUTE="drawing_sheet_count",
                CATEGORY="DRAWING",
                NX_ATTRIBUTE_NAME="DrawingSheets",
                ATTRIBUTE_TYPE="Integer",
                AUTHORITATIVE_SOURCE="DRAWING",
                EXPECTED_VALUE=">=1",
                DRAWING_VALUE=sheet_count,
                SEVERITY="PASS" if sheet_count > 0 else "BLOCK",
                COMPARISON_RESULT="PASS" if sheet_count > 0 else "FAIL",
                FAILURE_CODE="PASS" if sheet_count > 0 else "MISSING_DRAWING",
                MESSAGE="Drawing contains at least one sheet." if sheet_count > 0 else (
                    "Drawing opened but contains no readable sheets."
                ),
            )
            rows.append(sheet_row)
            drawing_values = read_attributes(drawing, config)
            for rule in drawing_rules:
                rows.append(
                    _drawing_attribute_detail(
                        context,
                        node,
                        rule,
                        node["attributes"][rule["logical_name"]],
                        drawing_values[rule["logical_name"]],
                        "YES",
                        index,
                        drawing,
                    )
                )
    return rows


def main(session):
    work_part = require_work_part(session)
    if work_part is None:
        return
    original_display = session.Parts.Display
    original_work = session.Parts.Work
    opened_parts = []
    detail_rows = []
    restore_messages = []
    config = load_config(_REPO_ROOT)
    nodes, traversal_findings = collect_bom_nodes(work_part, config)
    root_part_number = nodes[0]["part_number"] if nodes else "UNKNOWN"
    timestamp = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
    context = {
        "config": config,
        "run_id": new_run_id(root_part_number),
        "timestamp": timestamp,
        "config_hash": config_sha256(config),
        "scope_hash": "",
        "master_hash": "",
    }
    scope_path = _env_path(
        "NX_DRAWING_SCOPE_FILE", config["inputs"]["drawing_scope_filename"]
    )
    master_override = clean_text(os.environ.get("NX_ATTRIBUTE_MASTER_FILE"))
    master_path = _env_path(
        "NX_ATTRIBUTE_MASTER_FILE", config["inputs"]["master_reference_filename"]
    )
    scope = {}
    scope_available = False

    try:
        for finding in traversal_findings:
            detail_rows.append(_finding_detail(context, finding))
        if os.path.exists(scope_path):
            context["scope_hash"] = file_sha256(scope_path)
            scope, scope_findings = load_drawing_scope(scope_path)
            scope_available = True
            for finding in scope_findings:
                detail_rows.append(_finding_detail(context, finding))
        else:
            detail_rows.append(
                _finding_detail(
                    context,
                    {
                        "code": "DRAWING_SCOPE_REVIEW",
                        "message": "Required drawing-scope CSV not found: {0}".format(scope_path),
                    },
                )
            )

        uf_session = NXOpen.UF.UFSession.GetUFSession()
        model_rows, unique_nodes = _audit_models(context, nodes, uf_session)
        detail_rows.extend(model_rows)
        if scope_available:
            detail_rows.extend(_audit_drawings(session, context, unique_nodes, scope, opened_parts))

        nx_bom_rows = bom_export_rows(nodes)
        if os.path.exists(master_path):
            context["master_hash"] = file_sha256(master_path)
            for finding in compare_master_reference(master_path, root_part_number, nx_bom_rows):
                detail_rows.append(_finding_detail(context, finding))
        elif master_override:
            detail_rows.append(
                _finding_detail(
                    context,
                    {
                        "code": "DOWNSTREAM_BOM_DRIFT",
                        "message": "Explicit MASTER reference file was not found: {0}".format(
                            master_path
                        ),
                    },
                )
            )
    except Exception as exc:
        exception_fields = _exception_fields(exc)
        detail_rows.append(
            _finding_detail(
                context,
                {
                    "code": "ERROR",
                    "message": "Unhandled audit stage: " + _exception_details(exc),
                    "exception_type": exception_fields["exception_type"],
                    "error_code": exception_fields["error_code"],
                },
            )
        )
    finally:
        restore_messages.extend(_restore_state(session, original_display, original_work))
        for opened in reversed(opened_parts):
            close_error = _close_opened(opened)
            if close_error:
                restore_messages.append(
                    "Unable to close journal-opened drawing {0}: {1}".format(
                        _journal_identifier(opened), close_error
                    )
                )

    for message in restore_messages:
        detail_rows.append(
            _finding_detail(context, {"code": "ERROR", "message": message})
        )

    # Hashes are finalized after optional inputs are processed.  Backfill all
    # rows so each detail line independently identifies the complete run input.
    for row in detail_rows:
        row["CONFIG_SHA256"] = context["config_hash"]
        row["DRAWING_SCOPE_SHA256"] = context["scope_hash"]
        row["MASTER_REFERENCE_SHA256"] = context["master_hash"]

    output_folder = get_output_folder()
    detail_path = os.path.join(
        output_folder, "RECONCILIATION_{0}.csv".format(context["run_id"])
    )
    summary_path = os.path.join(
        output_folder, "RECONCILIATION_SUMMARY_{0}.csv".format(context["run_id"])
    )
    write_csv(detail_path, DETAIL_COLUMNS, detail_rows)
    counts = Counter(row.get("FAILURE_CODE", "") for row in detail_rows)
    blocker_count = sum(1 for row in detail_rows if row.get("SEVERITY") == "BLOCK")
    warning_count = sum(1 for row in detail_rows if row.get("SEVERITY") == "WARN")
    summary_rows = [
        ["RUN_ID", context["run_id"]],
        ["ROOT_PART_NUMBER", root_part_number],
        ["AUTHORITY", "NX_TEAMCENTER"],
        ["CONFIG_FILE", config_path(_REPO_ROOT)],
        ["CONFIG_SHA256", context["config_hash"]],
        ["DRAWING_SCOPE_FILE", scope_path],
        ["DRAWING_SCOPE_SHA256", context["scope_hash"]],
        ["MASTER_REFERENCE_FILE", master_path if os.path.exists(master_path) else "NOT_SUPPLIED"],
        ["MASTER_REFERENCE_SHA256", context["master_hash"]],
        ["BOM_ROWS", len(nodes)],
        ["DETAIL_ROWS", len(detail_rows)],
        ["BLOCKING_FINDINGS", blocker_count],
        ["WARNINGS", warning_count],
        ["CERTIFICATION", "PASS" if blocker_count == 0 else "FAIL"],
    ]
    summary_rows.extend([["FAILURE_CODE_" + code, count] for code, count in sorted(counts.items())])
    write_csv(summary_path, ["Metric", "Value"], summary_rows)

    certified_path = ""
    if blocker_count == 0:
        certified_path = os.path.join(
            output_folder, "NX_CERTIFIED_BOM_{0}.csv".format(context["run_id"])
        )
        write_csv(certified_path, FZ_BOM_COLUMNS, bom_export_rows(nodes))

    log_info(
        session,
        "\n".join(
            [
                "J04 NX-authoritative reconciliation complete.",
                "  Run ID             : {0}".format(context["run_id"]),
                "  BOM rows           : {0}".format(len(nodes)),
                "  Blocking findings  : {0}".format(blocker_count),
                "  Warnings           : {0}".format(warning_count),
                "  Certification      : {0}".format("PASS" if blocker_count == 0 else "FAIL"),
                "  Detail report      : {0}".format(detail_path),
                "  Summary report     : {0}".format(summary_path),
                "  Certified BOM      : {0}".format(certified_path or "WITHHELD"),
            ]
        ),
    )


if __name__ == "__main__":
    run_journal(main)
