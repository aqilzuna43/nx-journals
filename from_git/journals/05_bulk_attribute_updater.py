"""Journal 05 - approved business-attribute updates for 3D master models.

Input is the wide CSV emitted by Journal 04 plus its .baseline.json sidecar.
DRY_RUN is always non-mutating. APPLY_APPROVED requires APPROVED=YES, a clean
stale-value preflight, explicit Teamcenter checkout, verification, and the
SAVE_CHANGED_PARTS configuration gate.

Managed targets are deduplicated and checked with one session-wide status
snapshot. Already checked-out targets are reused; all remaining targets are
checked out in one batch before any attribute write begins.
"""

import csv
import json
import os
import time
import traceback
from collections import OrderedDict
from datetime import datetime

import NXOpen
import NXOpen.PDM


# ============================================================================
# USER SETTINGS - EDIT ONLY THESE TWO LINES FOR NORMAL NX USE
# Paste the full path of the CSV created by Journal 04 between the quotes.
USER_UPDATE_CSV = r""
# Keep DRY_RUN for validation. The only other valid value is APPLY_APPROVED.
USER_MODE = "DRY_RUN"
# PowerShell variables NX_ATTRIBUTE_UPDATE_FILE and NX_J05_MODE, when set,
# override these values for automated or advanced use.
# ============================================================================


CONTROL_COLUMNS = [
    "AUDIT_RUN_ID",
    "APPROVED",
    "ENGINEER",
    "APPROVAL_NOTE",
    "PULL_STATUS",
    "PULL_MESSAGE",
]
VALID_MODES = ("DRY_RUN", "APPLY_APPROVED")
COMPATIBILITY_ALLOWED_VALUES = {
    "stocking_type": ("BUY/REF",),
}
# Journal 05 intentionally accepts future commodity names without requiring a
# shared reconciliation-config update. Blank values are still rejected before
# this exception is considered.
OPEN_TEXT_UPDATE_FIELDS = ("commodity_type",)
# CSV display names can change without changing the underlying NX attribute.
# Keep previously emitted J04 update packages usable when that happens.
LEGACY_COLUMN_ALIASES = {
    "Traceability": ("SERIAL_NUMBERED_PART",),
}
REPORT_COLUMNS = [
    "RUN_TIMESTAMP",
    "MODE",
    "AUDIT_RUN_ID",
    "CSV_ROW",
    "PART_NUMBER",
    "REVISION",
    "TARGET_IDENTIFIER",
    "CSV_COLUMN",
    "LOGICAL_ATTRIBUTE",
    "CATEGORY",
    "NX_ATTRIBUTE_NAME",
    "NX_ATTRIBUTE_TYPE",
    "BASELINE_VALUE",
    "ACTUAL_CURRENT_VALUE",
    "EXPECTED_VALUE",
    "APPROVED",
    "ENGINEER",
    "ACTION",
    "CHECKOUT_BEFORE",
    "CHECKOUT_ACTION",
    "CHECKOUT_RESULT",
    "READ_ONLY_BEFORE",
    "READ_ONLY_AFTER",
    "WRITE_ATTEMPTED",
    "ROLLBACK_RESULT",
    "REREAD_VALUE",
    "VERIFICATION_RESULT",
    "SAVE_RESULT",
    "NX_EXCEPTION_TYPE",
    "NX_ERROR_CODE",
    "MESSAGE",
]


def _text(value):
    return "" if value is None else str(value)


def _clean(value):
    return _text(value).strip()


def _normalized(value):
    return " ".join(_clean(value).split()).upper()


def configured_input_path():
    return _clean(
        os.environ.get("NX_ATTRIBUTE_UPDATE_FILE") or USER_UPDATE_CSV
    )


def configured_mode():
    return _normalized(os.environ.get("NX_J05_MODE") or USER_MODE or "DRY_RUN")


def _enum_name(value):
    if value is None:
        return ""
    name = getattr(value, "name", None)
    return _text(name if name is not None else value).split(".")[-1]


def _dispose(value):
    if value is None:
        return
    for name in ("Dispose", "FreeResource"):
        method = getattr(value, name, None)
        if callable(method):
            try:
                method()
            except Exception:
                pass
            return


def _exception_fields(error):
    return type(error).__name__, _text(getattr(error, "ErrorCode", ""))


def _exception_details(error):
    exception_type, error_code = _exception_fields(error)
    code = ": {0}".format(error_code) if error_code else ""
    return "{0}{1} - {2}".format(exception_type, code, _text(error))


def _journal_path():
    return os.path.abspath(__file__)


def _runtime_candidates():
    script_parent = os.path.dirname(_journal_path())
    configured = _clean(os.environ.get("NX_JOURNALS_ROOT"))
    candidates = [
        configured,
        os.path.join(configured, "from_git") if configured else "",
        os.path.dirname(script_parent),
        os.getcwd(),
        os.path.join(os.getcwd(), "from_git"),
    ]
    result = []
    for candidate in candidates:
        if not candidate:
            continue
        absolute = os.path.abspath(candidate)
        if absolute not in result:
            result.append(absolute)
    return result


def _runtime_root():
    attempted = []
    for candidate in _runtime_candidates():
        attempted.append(candidate)
        if os.path.isfile(
            os.path.join(candidate, "config", "attribute_reconciliation.json")
        ):
            return candidate
    raise RuntimeError(
        "Journal 05 configuration was not found. Deploy config beside journals "
        "or set NX_JOURNALS_ROOT. Attempted: {0}".format(
            " | ".join(attempted)
        )
    )


def _io_root():
    configured = _clean(os.environ.get("NX_JOURNALS_IO_DIR"))
    if configured:
        return os.path.abspath(configured)
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    return desktop if os.path.isdir(desktop) else os.getcwd()


def _load_config():
    path = os.path.join(
        _runtime_root(), "config", "attribute_reconciliation.json"
    )
    with open(path, "r", encoding="utf-8-sig") as handle:
        config = json.load(handle)
    if config.get("authority") != "NX_TEAMCENTER":
        raise RuntimeError("Unsupported reconciliation authority.")
    workflow = config.get("update_workflow", {})
    if workflow.get("schema_version") != 1:
        raise RuntimeError("Unsupported or missing update_workflow schema.")
    if config.get("save_policy") not in (
        "NO_SAVE",
        "SAVE_CHANGED_PARTS",
    ):
        raise RuntimeError("Invalid save_policy.")
    _business_specs(config)
    _identity_specs(config)
    return config


def _rule_map(config):
    return {
        rule["logical_name"]: rule for rule in config.get("attributes", [])
    }


def _mapped_specs(config, key):
    rules = _rule_map(config)
    specs = []
    seen = set()
    for mapping in config["update_workflow"][key]:
        column = _clean(mapping.get("csv_column"))
        logical_name = _clean(mapping.get("logical_name"))
        rule = rules.get(logical_name)
        if not column or rule is None:
            raise RuntimeError(
                "Invalid {0} mapping: {1}".format(key, logical_name)
            )
        if column in seen:
            raise RuntimeError("Duplicate update CSV column: " + column)
        seen.add(column)
        specs.append((column, rule))
    return specs


def _identity_specs(config):
    return _mapped_specs(config, "identity_columns")


def _business_specs(config):
    specs = _mapped_specs(config, "business_columns")
    for _column, rule in specs:
        if (
            not rule.get("writable")
            or "MODEL" not in rule.get("write_targets", [])
        ):
            raise RuntimeError(
                "Business mapping is not model-writable: {0}".format(
                    rule["logical_name"]
                )
            )
    return specs


def update_columns(config):
    return (
        list(CONTROL_COLUMNS)
        + [column for column, _rule in _identity_specs(config)]
        + [column for column, _rule in _business_specs(config)]
    )


def _input_column_sources(headers, required):
    sources = {}
    missing = []
    for column in required:
        candidates = [column] + list(LEGACY_COLUMN_ALIASES.get(column, ()))
        present = [candidate for candidate in candidates if candidate in headers]
        if not present:
            missing.append(column)
            continue
        if len(present) > 1:
            raise RuntimeError(
                "Update CSV has ambiguous columns for {0}: {1}".format(
                    column, ", ".join(present)
                )
            )
        sources[column] = present[0]
    if missing:
        raise RuntimeError(
            "Update CSV is missing columns: {0}".format(", ".join(missing))
        )
    return sources


def _read_csv(path, config):
    required = update_columns(config)
    last_error = None
    for encoding in ("utf-8-sig", "utf-8", "cp1252"):
        try:
            with open(path, "r", encoding=encoding, newline="") as handle:
                reader = csv.DictReader(handle)
                headers = [_clean(name) for name in (reader.fieldnames or [])]
                sources = _input_column_sources(headers, required)
                rows = []
                for row_number, source in enumerate(reader, 2):
                    row = {
                        _clean(key): _clean(value)
                        for key, value in source.items()
                        if key is not None
                    }
                    for column, source_column in sources.items():
                        row[column] = row.get(source_column, "")
                    row["_CSV_ROW"] = row_number
                    rows.append(row)
                return rows
        except UnicodeDecodeError as exc:
            last_error = exc
    raise RuntimeError(
        "Unable to decode update CSV: {0}".format(last_error or path)
    )


def _baseline_path(csv_path):
    return os.path.splitext(csv_path)[0] + ".baseline.json"


def _load_baseline(csv_path):
    path = _baseline_path(csv_path)
    if not os.path.isfile(path):
        raise RuntimeError(
            "Journal 04 baseline sidecar not found: {0}".format(path)
        )
    with open(path, "r", encoding="utf-8-sig") as handle:
        baseline = json.load(handle)
    if baseline.get("schema_version") != 1:
        raise RuntimeError("Unsupported Journal 04 baseline schema.")
    return baseline, path


def _baseline_contract(config, key):
    return [
        {
            "csv_column": column,
            "logical_name": rule["logical_name"],
            "category": rule["category"],
            "attribute": rule["attribute"],
            "type": rule["type"],
        }
        for column, rule in _mapped_specs(config, key)
    ]


def _canonical_contract(contract):
    aliases = {
        alias: canonical
        for canonical, legacy_names in LEGACY_COLUMN_ALIASES.items()
        for alias in legacy_names
    }
    result = []
    for source in contract or []:
        item = dict(source)
        item["csv_column"] = aliases.get(
            item.get("csv_column"), item.get("csv_column")
        )
        result.append(item)
    return result


def _validate_baseline_contract(baseline, config):
    if baseline.get("identity_columns") != _baseline_contract(
        config, "identity_columns"
    ):
        raise RuntimeError(
            "Baseline identity mapping does not match the deployed config."
        )
    if _canonical_contract(
        baseline.get("business_columns")
    ) != _baseline_contract(config, "business_columns"):
        raise RuntimeError(
            "Baseline business mapping does not match the deployed config."
        )


def _baseline_business_value(baseline_part, column):
    values = baseline_part.get("business_values", {})
    if column in values:
        return values[column]
    for alias in LEGACY_COLUMN_ALIASES.get(column, ()):
        if alias in values:
            return values[alias]
    return {}


def _attribute_value(info):
    kind = _enum_name(getattr(info, "Type", ""))
    numeric_kind = getattr(info, "Type", None)
    if kind in ("String", "5") or numeric_kind == 5:
        return getattr(info, "StringValue", ""), "String"
    if kind in ("Real", "Number", "4") or numeric_kind == 4:
        return getattr(info, "RealValue", None), "Number"
    if kind in ("Integer", "3") or numeric_kind == 3:
        return getattr(info, "IntegerValue", None), "Integer"
    if kind in ("Boolean", "1") or numeric_kind == 1:
        return getattr(info, "BooleanValue", None), "Boolean"
    return getattr(info, "StringValue", ""), kind or "String"


def _read_attribute(nx_object, rule):
    iterator = None
    try:
        iterator = nx_object.CreateAttributeIterator()
        iterator.SetIncludeOnlyCategory(rule["category"])
        iterator.SetIncludeOnlyTitle(rule["attribute"])
        iterator.SetIncludeAlsoUnset(True)
        matches = [
            info
            for info in nx_object.GetUserAttributes(iterator)
            if _clean(getattr(info, "Category", "")) == rule["category"]
            and _clean(getattr(info, "Title", "")) == rule["attribute"]
        ]
        if not matches:
            return {
                "status": "MISSING",
                "raw": "",
                "type": rule["type"],
                "flags": {},
            }
        if len(matches) > 1:
            return {
                "status": "AMBIGUOUS",
                "raw": "",
                "type": rule["type"],
                "flags": {},
                "message": "Multiple category/title matches.",
            }
        info = matches[0]
        raw, actual_type = _attribute_value(info)
        status = (
            "UNSET"
            if bool(getattr(info, "Unset", False))
            else ("BLANK" if _clean(raw) == "" else "POPULATED")
        )
        return {
            "status": status,
            "raw": raw,
            "type": actual_type,
            "flags": {
                "locked": bool(getattr(info, "Locked", False)),
                "owned_by_system": bool(
                    getattr(info, "OwnedBySystem", False)
                ),
                "pdm_based": bool(getattr(info, "PdmBased", False)),
                "not_saved": bool(getattr(info, "NotSaved", False)),
            },
        }
    except Exception as exc:
        return {
            "status": "UNREADABLE",
            "raw": "",
            "type": rule.get("type", ""),
            "flags": {},
            "message": _exception_details(exc),
        }
    finally:
        _dispose(iterator)


def _read_identity_attribute(nx_object, rule):
    """Read hard-coded NX/CAD identity by title, independent of category."""
    try:
        raw = nx_object.GetStringAttribute(rule["attribute"])
        return {
            "status": "BLANK" if _clean(raw) == "" else "POPULATED",
            "raw": raw,
            "type": "String",
            "flags": {},
        }
    except AttributeError:
        return _read_attribute(nx_object, rule)
    except Exception:
        fallback = _read_attribute(nx_object, rule)
        return fallback


def _object_key(nx_object):
    tag = getattr(nx_object, "Tag", None)
    return ("TAG", _text(tag)) if tag is not None else ("PY", id(nx_object))


def _object_identifier(nx_object):
    for name in ("JournalIdentifier", "FullPath", "Name", "Leaf"):
        value = getattr(nx_object, name, "")
        if callable(value):
            try:
                value = value()
            except Exception:
                value = ""
        if _clean(value):
            return _clean(value)
    return _text(_object_key(nx_object))


def _children(component):
    try:
        return list(component.GetChildren())
    except Exception:
        return []


def _is_suppressed(component):
    try:
        return bool(component.IsSuppressed)
    except Exception:
        return False


def _unique_model_parts(work_part):
    unique = OrderedDict()

    def add(part):
        if part is not None:
            unique.setdefault(_object_key(part), part)

    add(work_part)
    root = getattr(
        getattr(work_part, "ComponentAssembly", None), "RootComponent", None
    )
    if root is None:
        return list(unique.values())
    stack = list(reversed(_children(root)))
    while stack:
        component = stack.pop()
        if _is_suppressed(component):
            continue
        add(getattr(component, "Prototype", None))
        stack.extend(reversed(_children(component)))
    return list(unique.values())


def _model_index(work_part, config):
    identity = {
        rule["logical_name"]: rule for _column, rule in _identity_specs(config)
    }
    by_identity = {}
    for part in _unique_model_parts(work_part):
        part_number = _clean(
            _read_identity_attribute(part, identity["part_number"])["raw"]
        )
        revision = _clean(
            _read_identity_attribute(part, identity["revision"])["raw"]
        )
        if part_number and revision:
            by_identity.setdefault(
                (_normalized(part_number), _normalized(revision)), []
            ).append(part)
    return by_identity


def _normalize_for_rule(value, rule, config):
    comparison = _normalized(rule.get("comparison"))
    if comparison == "BOOLEAN_ALIAS":
        aliases = config.get("release_policy", {}).get(
            "boolean_aliases", {}
        )
        normalized = _normalized(value)
        for canonical, candidates in aliases.items():
            if normalized in [_normalized(item) for item in candidates]:
                return _normalized(canonical)
    if comparison == "NUMBER":
        try:
            return "{0:.15g}".format(float(value))
        except (TypeError, ValueError):
            return _clean(value)
    if comparison in (
        "TRIMMED_CASE_INSENSITIVE",
        "NORMALIZED_TEXT",
        "BOOLEAN_ALIAS",
    ):
        return _normalized(value)
    return _clean(value)


def _validate_expected(value, rule, config):
    if _clean(value) == "":
        return "Populated-to-blank updates are not supported."
    kind = _normalized(rule.get("type"))
    if kind in ("NUMBER", "REAL"):
        try:
            float(value)
        except ValueError:
            return "Expected value is not a valid number."
    elif kind == "INTEGER":
        try:
            int(value)
        except ValueError:
            return "Expected value is not a valid integer."
    if rule.get("logical_name") in OPEN_TEXT_UPDATE_FIELDS:
        return ""
    allowed = list(rule.get("allowed_values") or [])
    allowed.extend(
        COMPATIBILITY_ALLOWED_VALUES.get(rule.get("logical_name"), ())
    )
    if allowed:
        expected = _normalize_for_rule(value, rule, config)
        allowed_normalized = [
            _normalize_for_rule(item, rule, config) for item in allowed
        ]
        if expected not in allowed_normalized:
            return "Expected value is outside the controlled value set."
    return ""


def _approved(row):
    return _normalized(row.get("APPROVED")) in ("YES", "Y", "TRUE", "1")


def _base_report(timestamp, mode, row, baseline_part=None):
    baseline_part = baseline_part or {}
    return {
        column: "" for column in REPORT_COLUMNS
    } | {
        "RUN_TIMESTAMP": timestamp,
        "MODE": mode,
        "AUDIT_RUN_ID": row.get("AUDIT_RUN_ID", ""),
        "CSV_ROW": row.get("_CSV_ROW", ""),
        "PART_NUMBER": baseline_part.get(
            "part_number", row.get("Item Number", "")
        ),
        "REVISION": baseline_part.get(
            "revision", row.get("Item Rev", "")
        ),
        "APPROVED": row.get("APPROVED", ""),
        "ENGINEER": row.get("ENGINEER", ""),
        "WRITE_ATTEMPTED": "NO",
    }


def _baseline_index(baseline):
    result = {}
    for part in baseline.get("parts", []):
        key = (
            _normalized(part.get("part_number")),
            _normalized(part.get("revision")),
        )
        result.setdefault(key, []).append(part)
    return result


def _read_only(part):
    value = getattr(part, "IsReadOnly", None)
    try:
        value = value() if callable(value) else value
    except Exception:
        return None
    return None if value is None else bool(value)


def _managed_mode(session):
    value = getattr(session, "IsManagedMode", False)
    try:
        value = value() if callable(value) else value
    except Exception:
        return False
    return bool(value)


def _is_teamcenter_target(session, target):
    if _managed_mode(session):
        return True
    return _object_identifier(target).upper().startswith("@DB/")


def _pdm_part(target):
    value = getattr(target, "PDMPart", None)
    return value() if callable(value) else value


def _pdm_checkout_state(target):
    pdm_part = _pdm_part(target)
    method = getattr(pdm_part, "GetCheckedoutStatusAndUser", None)
    if not callable(method):
        return None, ""
    try:
        result = method()
        values = result if isinstance(result, tuple) else (result,)
        status = values[0] if values else None
        user = _clean(values[1]) if len(values) > 1 else ""
        if isinstance(status, bool):
            state = "CHECKED_OUT" if status else "NOT_CHECKED_OUT"
        else:
            status_name = _normalized(_enum_name(status)).replace("_", "")
            if status_name.startswith("NOT") and "CHECKEDOUT" in status_name:
                state = "NOT_CHECKED_OUT"
            elif "CHECKEDOUT" in status_name:
                state = "CHECKED_OUT"
            else:
                return None, "PDM checkout status: " + repr(result)
        detail = "Checkout user: " + user if user else ""
        return state, detail
    except Exception as exc:
        return None, "PDM checkout status unavailable: " + _exception_details(
            exc
        )


def _session_checkout_snapshot(session):
    """Return one session-wide checkout snapshot or a fail-closed error."""
    pdm_session = getattr(session, "PdmSession", None)
    method = getattr(
        pdm_session, "GetCheckedoutStatusOfAllObjectsInSession", None
    )
    if not callable(method):
        return None, None, "API_UNAVAILABLE"
    try:
        result = method()
        if isinstance(result, tuple) and len(result) >= 2:
            checked, unchecked = result[0], result[1]
            return (
                {_object_key(item) for item in checked},
                {_object_key(item) for item in unchecked},
                "",
            )
        return None, None, "UNRECOGNIZED_API_RESULT"
    except Exception as exc:
        return None, None, _exception_details(exc)


def _checkout_result(target, before, action="NONE"):
    read_only = _read_only(target)
    return {
        "success": False,
        "before": before,
        "action": action,
        "result": "",
        "read_only_before": read_only,
        "read_only_after": read_only,
        "message": "",
        "exception_type": "",
        "error_code": "",
    }


def _batch_checkout(targets):
    """Explicitly check out loaded targets in one managed-NX operation."""
    if not targets:
        return
    pdm_part = _pdm_part(targets[0])
    checkout = getattr(pdm_part, "CheckoutParts", None)
    if not callable(checkout):
        raise RuntimeError("PDMPart.CheckoutParts is unavailable.")
    checkout_input = NXOpen.PDM.PdmPart.CheckoutInput(
        "J05 approved business-attribute update",
        "",
        True,
        True,
        False,
    )
    operation_errors = None
    try:
        operation_errors = checkout(targets, checkout_input)
    finally:
        _dispose(operation_errors)


def prepare_updates(
    session, work_part, config, rows, baseline, timestamp, mode
):
    _validate_baseline_contract(baseline, config)
    reports = []
    proposals = []
    baseline_by_identity = _baseline_index(baseline)
    models = _model_index(work_part, config)
    business_specs = _business_specs(config)
    expected_audit = _clean(baseline.get("audit_run_id"))

    csv_counts = {}
    for row in rows:
        key = (
            _normalized(row.get("Item Number")),
            _normalized(row.get("Item Rev")),
        )
        csv_counts[key] = csv_counts.get(key, 0) + 1

    for row in rows:
        key = (
            _normalized(row.get("Item Number")),
            _normalized(row.get("Item Rev")),
        )
        baseline_matches = baseline_by_identity.get(key, [])
        base = _base_report(timestamp, mode, row)
        if csv_counts.get(key, 0) > 1:
            base.update(
                ACTION="ERROR_DUPLICATE_CSV_IDENTITY",
                MESSAGE="Update CSV contains duplicate part/revision rows.",
            )
            reports.append(base)
            continue
        if len(baseline_matches) != 1:
            base.update(
                ACTION="ERROR_BASELINE_IDENTITY",
                MESSAGE=(
                    "Part/revision was edited or has no unique Journal 04 "
                    "baseline."
                ),
            )
            reports.append(base)
            continue
        baseline_part = baseline_matches[0]
        base = _base_report(timestamp, mode, row, baseline_part)
        if _clean(row.get("AUDIT_RUN_ID")) != expected_audit:
            base.update(
                ACTION="ERROR_BASELINE_RUN",
                MESSAGE="AUDIT_RUN_ID does not match the baseline sidecar.",
            )
            reports.append(base)
            continue
        if _normalized(row.get("PULL_STATUS")) != "READY":
            base.update(
                ACTION="ERROR_PULL_REVIEW",
                MESSAGE="Journal 04 marked this prototype for review.",
            )
            reports.append(base)
            continue
        if _normalized(row.get("Part Description")) != _normalized(
            baseline_part.get("part_name")
        ):
            base.update(
                ACTION="ERROR_PROTECTED_IDENTITY_EDIT",
                MESSAGE="Part Description is read-only and was changed.",
            )
            reports.append(base)
            continue

        target_matches = models.get(key, [])
        if len(target_matches) != 1:
            base.update(
                ACTION="ERROR_TARGET_IDENTITY",
                MESSAGE="No unique loaded 3D master prototype matches the row.",
            )
            reports.append(base)
            continue
        target = target_matches[0]
        approved = _approved(row)
        row_change_count = 0
        for column, rule in business_specs:
            baseline_value = _baseline_business_value(
                baseline_part, column
            ).get("raw_value", "")
            expected = row.get(column, "")
            if _normalize_for_rule(
                baseline_value, rule, config
            ) == _normalize_for_rule(expected, rule, config):
                continue
            row_change_count += 1
            report = dict(base)
            report.update(
                TARGET_IDENTIFIER=_object_identifier(target),
                CSV_COLUMN=column,
                LOGICAL_ATTRIBUTE=rule["logical_name"],
                CATEGORY=rule["category"],
                NX_ATTRIBUTE_NAME=rule["attribute"],
                NX_ATTRIBUTE_TYPE=rule["type"],
                BASELINE_VALUE=_text(baseline_value),
                EXPECTED_VALUE=_text(expected),
            )
            if not approved:
                report.update(
                    ACTION="SKIPPED_NOT_APPROVED",
                    MESSAGE="Row changes are not approved.",
                )
                reports.append(report)
                continue
            if not _clean(row.get("ENGINEER")):
                report.update(
                    ACTION="ERROR_APPROVAL",
                    MESSAGE="ENGINEER is required for an approved row.",
                )
                reports.append(report)
                continue
            expected_error = _validate_expected(expected, rule, config)
            if expected_error:
                report.update(ACTION="ERROR_VALUE", MESSAGE=expected_error)
                reports.append(report)
                continue

            actual = _read_attribute(target, rule)
            report["ACTUAL_CURRENT_VALUE"] = _text(actual.get("raw", ""))
            if actual["status"] in ("UNREADABLE", "AMBIGUOUS"):
                report.update(
                    ACTION="ERROR_ATTRIBUTE_READ",
                    MESSAGE=actual.get("message", actual["status"]),
                )
                reports.append(report)
                continue
            flags = actual.get("flags", {})
            if (
                flags.get("locked")
                or flags.get("owned_by_system")
                or flags.get("pdm_based")
            ):
                report.update(
                    ACTION="ERROR_ATTRIBUTE_NOT_WRITABLE",
                    MESSAGE=(
                        "Runtime attribute flags prohibit this business "
                        "attribute write."
                    ),
                )
                reports.append(report)
                continue
            actual_value = actual.get("raw", "")
            actual_normalized = _normalize_for_rule(
                actual_value, rule, config
            )
            expected_normalized = _normalize_for_rule(
                expected, rule, config
            )
            baseline_normalized = _normalize_for_rule(
                baseline_value, rule, config
            )
            if actual_normalized == expected_normalized:
                report.update(
                    ACTION="ALREADY_AT_EXPECTED_VALUE",
                    REREAD_VALUE=_text(actual_value),
                    VERIFICATION_RESULT="ALREADY_MATCHED",
                    SAVE_RESULT="NOT_REQUIRED",
                    MESSAGE=(
                        "Current NX value already matches the approved value; "
                        "no write or checkout is required."
                    ),
                )
                reports.append(report)
                continue
            if actual_normalized != baseline_normalized:
                report.update(
                    ACTION="STALE_BASELINE_VALUE",
                    MESSAGE=(
                        "Current NX value differs from the Journal 04 "
                        "baseline."
                    ),
                )
                reports.append(report)
                continue

            report.update(
                ACTION="PROPOSED_UPDATE",
                READ_ONLY_BEFORE=_text(_read_only(target)),
                MESSAGE="Approved change passed preflight.",
            )
            reports.append(report)
            proposals.append(
                {
                    "source_row": row,
                    "report": report,
                    "rule": rule,
                    "target": target,
                    "expected": expected,
                }
            )

        if row_change_count == 0:
            no_change = dict(base)
            no_change.update(
                TARGET_IDENTIFIER=_object_identifier(target),
                ACTION="NO_CHANGE",
                SAVE_RESULT="NOT_REQUIRED",
                MESSAGE="CSV business values match the Journal 04 baseline.",
            )
            reports.append(no_change)
    return reports, proposals


def _hard_preflight_error(report):
    action = _clean(report.get("ACTION"))
    return (
        _approved(report)
        and action not in (
            "PROPOSED_UPDATE",
            "NO_CHANGE",
            "ALREADY_AT_EXPECTED_VALUE",
        )
    )


def _apply_checkout_results(proposals, results):
    for proposal in proposals:
        result = results[_object_key(proposal["target"])]
        proposal["report"].update(
            CHECKOUT_BEFORE=result["before"],
            CHECKOUT_ACTION=result["action"],
            CHECKOUT_RESULT=result["result"],
            READ_ONLY_BEFORE=_text(result["read_only_before"]),
            READ_ONLY_AFTER=_text(result["read_only_after"]),
        )
        if result.get("exception_type"):
            proposal["report"]["NX_EXCEPTION_TYPE"] = result[
                "exception_type"
            ]
            proposal["report"]["NX_ERROR_CODE"] = result["error_code"]


def checkout_targets(session, work_part, proposals, progress=None):
    del work_part  # Kept in the signature for journal/test compatibility.
    targets = OrderedDict()
    for proposal in proposals:
        targets.setdefault(_object_key(proposal["target"]), proposal["target"])

    results = OrderedDict()
    managed_targets = OrderedDict()
    for key, target in targets.items():
        if _is_teamcenter_target(session, target):
            managed_targets[key] = target
            continue
        result = _checkout_result(target, "NATIVE_MODE")
        result["action"] = "NATIVE_MODE_NO_CHECKOUT"
        result["success"] = result["read_only_before"] is not True
        result["result"] = (
            "WRITABLE" if result["success"] else "READ_ONLY"
        )
        if not result["success"]:
            result["message"] = "Native target is read-only."
        results[key] = result

    if managed_targets:
        snapshot_start = time.perf_counter()
        checked, unchecked, snapshot_error = _session_checkout_snapshot(
            session
        )
        if progress is not None:
            progress(
                "  Checkout status snapshot: {0} target(s), {1:.3f} s".format(
                    len(managed_targets), time.perf_counter() - snapshot_start
                )
            )

        pending = OrderedDict()
        for key, target in managed_targets.items():
            read_only = _read_only(target)
            before = (
                "CHECKED_OUT"
                if checked is not None and key in checked
                else (
                    "NOT_CHECKED_OUT"
                    if unchecked is not None and key in unchecked
                    else "UNKNOWN"
                )
            )
            result = _checkout_result(target, before)
            if snapshot_error:
                result.update(
                    result="FAILED",
                    message=(
                        "Session checkout status is unavailable: "
                        + snapshot_error
                    ),
                )
            elif before == "CHECKED_OUT" and read_only is not True:
                result.update(
                    success=True,
                    result="ALREADY_CHECKED_OUT",
                )
            elif before == "CHECKED_OUT":
                pdm_state, detail = _pdm_checkout_state(target)
                result.update(
                    result="FAILED",
                    message=(
                        "Target is checked out but remains read-only."
                        + (" " + detail if detail else "")
                    ),
                )
                if pdm_state == "CHECKED_OUT" and not detail:
                    result["message"] += " Checkout ownership is unavailable."
            elif before == "NOT_CHECKED_OUT":
                result["action"] = "BATCH_CHECKOUT"
                pending[key] = target
            else:
                pdm_state, detail = _pdm_checkout_state(target)
                result.update(
                    result="FAILED",
                    message=(
                        "Target was not represented in the session checkout "
                        "snapshot."
                        + (" " + detail if detail else "")
                    ),
                )
                if pdm_state == "CHECKED_OUT" and read_only is not True:
                    result.update(
                        success=True,
                        result="ALREADY_CHECKED_OUT",
                        message=detail,
                    )
            results[key] = result

        if pending and not snapshot_error:
            batch_error = None
            batch_start = time.perf_counter()
            if progress is not None:
                progress(
                    "  Batch checkout: {0} target(s)".format(len(pending))
                )
            try:
                _batch_checkout(list(pending.values()))
            except Exception as exc:
                batch_error = exc

            post_checked, _post_unchecked, post_error = (
                _session_checkout_snapshot(session)
            )
            if progress is not None:
                progress(
                    "  Batch checkout verification: {0:.3f} s".format(
                        time.perf_counter() - batch_start
                    )
                )

            for key, target in pending.items():
                result = results[key]
                read_only_after = _read_only(target)
                result["read_only_after"] = read_only_after
                if (
                    batch_error is None
                    and not post_error
                    and post_checked is not None
                    and key in post_checked
                    and read_only_after is not True
                ):
                    result.update(
                        success=True,
                        result="BATCH_CHECKOUT",
                        message="Batch checkout verified.",
                    )
                    continue

                details = []
                if batch_error is not None:
                    details.append(_exception_details(batch_error))
                    exception_type, error_code = _exception_fields(
                        batch_error
                    )
                    result["exception_type"] = exception_type
                    result["error_code"] = error_code
                if post_error:
                    details.append(
                        "Post-checkout session status unavailable: "
                        + post_error
                    )
                pdm_state, pdm_detail = _pdm_checkout_state(target)
                if pdm_detail:
                    details.append(pdm_detail)
                if pdm_state == "CHECKED_OUT" and read_only_after is True:
                    details.append("Target remains read-only.")
                result.update(
                    success=False,
                    result="FAILED",
                    message=" | ".join(details)
                    or "Batch checkout postconditions were not satisfied.",
                )

    _apply_checkout_results(proposals, results)
    return results


def _builder_data_type(rule):
    enum = NXOpen.AttributePropertiesBaseBuilder.DataTypeOptions
    kind = _normalized(rule["type"])
    if kind == "BOOLEAN":
        return enum.Boolean
    if kind == "INTEGER":
        return enum.Integer
    if kind in ("NUMBER", "REAL"):
        return enum.Number
    if kind in ("DATE", "TIME"):
        return enum.Date
    return enum.String


def _set_builder_value(builder, rule, value):
    kind = _normalized(rule["type"])
    if kind == "BOOLEAN":
        builder.BooleanValue = (
            NXOpen.AttributePropertiesBaseBuilder.BooleanValueOptions.TrueValue
            if _normalized(value) in ("Y", "YES", "TRUE", "1")
            else NXOpen.AttributePropertiesBaseBuilder.BooleanValueOptions.FalseValue
        )
    elif kind == "INTEGER":
        builder.IntegerValue = int(value)
    elif kind in ("NUMBER", "REAL"):
        builder.NumberValue = float(value)
    else:
        builder.StringValue = _text(value)


def _write_attribute(session, target, rule, expected):
    builder = None
    try:
        builder = session.AttributeManager.CreateAttributePropertiesBuilder(
            target,
            [target],
            NXOpen.AttributePropertiesBuilder.OperationType.Save,
        )
        builder.Category = rule["category"]
        builder.Title = rule["attribute"]
        builder.DataType = _builder_data_type(rule)
        _set_builder_value(builder, rule, expected)
        builder.Commit()
    finally:
        _dispose(builder)


def _save_target(target):
    status = None
    try:
        status = target.Save(
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
    finally:
        _dispose(status)


def apply_groups(session, proposals, config, progress=None):
    grouped = OrderedDict()
    for proposal in proposals:
        grouped.setdefault(_object_key(proposal["target"]), []).append(
            proposal
        )
    unsaved_modified = set()
    stop_saves = False
    target_count = len(grouped)
    for target_index, group in enumerate(grouped.values(), 1):
        target = group[0]["target"]
        target_start = time.perf_counter()
        if progress is not None:
            progress(
                "  Updating target {0}/{1}: {2} ({3} attribute(s))".format(
                    target_index,
                    target_count,
                    _object_identifier(target),
                    len(group),
                )
            )
        mark_name = "J05 {0}".format(_object_identifier(target))
        mark = session.SetUndoMark(
            NXOpen.Session.MarkVisibility.Invisible, mark_name
        )
        failed = None
        try:
            for proposal in group:
                report = proposal["report"]
                report["WRITE_ATTEMPTED"] = "YES"
                _write_attribute(
                    session,
                    target,
                    proposal["rule"],
                    proposal["expected"],
                )
                reread = _read_attribute(target, proposal["rule"])
                report["REREAD_VALUE"] = _text(reread.get("raw", ""))
                if _normalize_for_rule(
                    reread.get("raw", ""), proposal["rule"], config
                ) != _normalize_for_rule(
                    proposal["expected"], proposal["rule"], config
                ):
                    raise RuntimeError(
                        "Immediate reread did not match expected value."
                    )
                report.update(
                    ACTION="UPDATED_VERIFIED",
                    VERIFICATION_RESULT="PASS",
                    MESSAGE="Attribute updated and reread successfully.",
                )
            unsaved_modified.add(_object_key(target))
        except Exception as exc:
            failed = exc
            try:
                session.UndoToMark(mark, mark_name)
                rollback = "PASS"
            except Exception as rollback_exc:
                rollback = "FAILED: " + _exception_details(rollback_exc)
            exception_type, error_code = _exception_fields(exc)
            for proposal in group:
                proposal["report"].update(
                    ACTION="UPDATED_VERIFICATION_FAILED",
                    ROLLBACK_RESULT=rollback,
                    VERIFICATION_RESULT="FAIL",
                    SAVE_RESULT="NOT_ATTEMPTED",
                    NX_EXCEPTION_TYPE=exception_type,
                    NX_ERROR_CODE=error_code,
                    MESSAGE=_exception_details(exc),
                )
            if rollback == "PASS":
                unsaved_modified.discard(_object_key(target))
        finally:
            try:
                session.DeleteUndoMark(mark, mark_name)
            except Exception:
                pass
        if failed:
            if progress is not None:
                progress(
                    "    Verification failed after {0:.3f} s".format(
                        time.perf_counter() - target_start
                    )
                )
            continue
        if stop_saves:
            for proposal in group:
                proposal["report"].update(
                    ACTION="ERROR_SAVE_NOT_ATTEMPTED",
                    SAVE_RESULT="NOT_ATTEMPTED",
                    MESSAGE="A previous save failed; later saves were stopped.",
                )
            if progress is not None:
                progress("    Save skipped after an earlier save failure.")
            continue
        try:
            _save_target(target)
            unsaved_modified.discard(_object_key(target))
            for proposal in group:
                proposal["report"]["SAVE_RESULT"] = "SAVED"
            if progress is not None:
                progress(
                    "    Verified and saved in {0:.3f} s".format(
                        time.perf_counter() - target_start
                    )
                )
        except Exception as exc:
            stop_saves = True
            exception_type, error_code = _exception_fields(exc)
            for proposal in group:
                proposal["report"].update(
                    ACTION="ERROR_SAVE_FAILED",
                    SAVE_RESULT="SAVE_FAILED_PART_LEFT_MODIFIED",
                    NX_EXCEPTION_TYPE=exception_type,
                    NX_ERROR_CODE=error_code,
                    MESSAGE=(
                        "Save failed; target remains checked out and visibly "
                        "modified: " + _exception_details(exc)
                    ),
                )
            if progress is not None:
                progress(
                    "    Save failed after {0:.3f} s".format(
                        time.perf_counter() - target_start
                    )
                )
    return unsaved_modified


def execute(
    session,
    work_part,
    config,
    rows,
    baseline,
    timestamp,
    mode,
    progress=None,
):
    preflight_start = time.perf_counter()
    reports, proposals = prepare_updates(
        session, work_part, config, rows, baseline, timestamp, mode
    )
    unique_target_count = len(
        {_object_key(proposal["target"]) for proposal in proposals}
    )
    if progress is not None:
        progress(
            "  Preflight: {0} proposed change(s) across {1} target(s), "
            "{2:.3f} s".format(
                len(proposals),
                unique_target_count,
                time.perf_counter() - preflight_start,
            )
        )
    if mode == "DRY_RUN" or not proposals:
        return reports, set()
    if config.get("save_policy") != "SAVE_CHANGED_PARTS":
        for proposal in proposals:
            proposal["report"].update(
                ACTION="SAVE_GATE_DISABLED",
                SAVE_RESULT="NOT_ATTEMPTED",
                MESSAGE=(
                    "Production save gate is NO_SAVE. Pass Journal 11 runtime "
                    "acceptance before enabling SAVE_CHANGED_PARTS."
                ),
            )
        return reports, set()
    if any(_hard_preflight_error(report) for report in reports):
        for proposal in proposals:
            proposal["report"].update(
                ACTION="BATCH_ABORTED_PREFLIGHT",
                SAVE_RESULT="NOT_ATTEMPTED",
                MESSAGE="An approved row failed preflight; no checkout occurred.",
            )
        return reports, set()

    checkout_start = time.perf_counter()
    checkout_results = checkout_targets(
        session, work_part, proposals, progress=progress
    )
    if progress is not None:
        progress(
            "  Checkout phase complete: {0:.3f} s".format(
                time.perf_counter() - checkout_start
            )
        )
    writable_proposals = []
    blocked_target_keys = {
        key
        for key, result in checkout_results.items()
        if not result["success"]
    }
    for proposal in proposals:
        key = _object_key(proposal["target"])
        if key not in blocked_target_keys:
            writable_proposals.append(proposal)
            continue
        result = checkout_results[key]
        proposal["report"].update(
            ACTION="CHECKOUT_FAILED",
            SAVE_RESULT="NOT_ATTEMPTED",
            MESSAGE=result["message"],
        )
    if progress is not None and blocked_target_keys:
        progress(
            "  Checkout gate: {0} writable target(s), {1} blocked target(s); "
            "continuing with writable targets only.".format(
                len(checkout_results) - len(blocked_target_keys),
                len(blocked_target_keys),
            )
        )
    if not writable_proposals:
        return reports, set()
    apply_start = time.perf_counter()
    unsaved = apply_groups(
        session, writable_proposals, config, progress=progress
    )
    if progress is not None:
        progress(
            "  Write/verify/save phase complete: {0:.3f} s".format(
                time.perf_counter() - apply_start
            )
        )
    return reports, unsaved


def _write_csv(path, rows):
    with open(path, "w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=REPORT_COLUMNS)
        writer.writeheader()
        for row in rows:
            writer.writerow(
                {column: row.get(column, "") for column in REPORT_COLUMNS}
            )


def _listing_window(session):
    listing = getattr(session, "ListingWindow", None)
    if listing is not None:
        try:
            listing.Open()
        except Exception:
            pass
    return listing


def _log(listing, message):
    if listing is not None:
        try:
            listing.WriteLine(_text(message))
            return
        except Exception:
            pass
    print(_text(message))


def main(session):
    run_start = time.perf_counter()
    work_part = getattr(getattr(session, "Parts", None), "Work", None)
    if work_part is None:
        raise RuntimeError("Open the source 3D assembly before Journal 05.")
    config = _load_config()
    mode = configured_mode()
    if mode not in VALID_MODES:
        raise RuntimeError(
            "USER_MODE (or NX_J05_MODE) must be DRY_RUN or APPLY_APPROVED."
        )
    input_path = configured_input_path()
    if not input_path:
        raise RuntimeError(
            "Edit USER_UPDATE_CSV near the top of Journal 05 and paste the "
            "full path of the edited Journal 04 CSV. Advanced users may "
            "instead set NX_ATTRIBUTE_UPDATE_FILE."
        )
    input_path = os.path.abspath(input_path)
    if not os.path.isfile(input_path):
        raise RuntimeError("Update CSV not found: " + input_path)

    baseline, baseline_path = _load_baseline(input_path)
    rows = _read_csv(input_path, config)
    timestamp = datetime.now().isoformat(timespec="seconds")
    listing = _listing_window(session)
    _log(listing, "Journal 05 started.")
    _log(listing, "  Mode: " + mode)
    _log(listing, "  Input rows: {0}".format(len(rows)))
    reports, unsaved = execute(
        session,
        work_part,
        config,
        rows,
        baseline,
        timestamp,
        mode,
        progress=lambda message: _log(listing, message),
    )
    report_stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_root = _io_root()
    os.makedirs(output_root, exist_ok=True)
    report_path = os.path.join(
        output_root, "J05_{0}_{1}.csv".format(mode, report_stamp)
    )
    _write_csv(report_path, reports)

    _log(listing, "Journal 05 complete.")
    _log(listing, "  Mode: " + mode)
    _log(listing, "  Input: " + input_path)
    _log(listing, "  Baseline: " + baseline_path)
    _log(listing, "  Save gate: " + config["save_policy"])
    _log(listing, "  Report: " + report_path)
    _log(
        listing,
        "  Total runtime: {0:.3f} s".format(time.perf_counter() - run_start),
    )
    if unsaved:
        _log(
            listing,
            "  WARNING: {0} target(s) may contain unsaved changes.".format(
                len(unsaved)
            ),
        )
    _log(
        listing,
        "Journal 05 never checks Teamcenter parts in automatically.",
    )
    return report_path


def _run_journal():
    session = NXOpen.Session.GetSession()
    listing = _listing_window(session)
    try:
        main(session)
    except Exception as exc:
        _log(listing, "JOURNAL 05 FAILED: " + _exception_details(exc))
        _log(listing, traceback.format_exc())
        raise


if __name__ == "__main__":
    _run_journal()
