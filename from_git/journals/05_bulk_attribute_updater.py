"""Journal 05 - Controlled NX/Teamcenter attribute correction.

Self-contained for NX 2312 deployment.  Modes are selected with NX_J05_MODE:
PULL, DRY_RUN (default), or APPLY_APPROVED.  MASTER/BOM values are never an
authoritative correction source.
"""

import csv
import json
import os
import sys
import traceback
from collections import OrderedDict
from datetime import datetime

import NXOpen
import NXOpen.UF


CORRECTION_COLUMNS = [
    "APPROVED",
    "PART_NUMBER",
    "REVISION",
    "TARGET_OBJECT",
    "DRAWING_INDEX",
    "CATEGORY",
    "NX_ATTRIBUTE_NAME",
    "NX_ATTRIBUTE_TYPE",
    "CURRENT_VALUE_FROM_AUDIT",
    "EXPECTED_VALUE",
    "AUTHORITATIVE_SOURCE",
    "AUDIT_RUN_ID",
    "ENGINEER",
    "EVIDENCE_REFERENCE",
    "APPROVAL_NOTE",
]
REPORT_COLUMNS = [
    "RUN_TIMESTAMP",
    "MODE",
    "AUDIT_RUN_ID",
    "PART_NUMBER",
    "REVISION",
    "TARGET_OBJECT",
    "DRAWING_INDEX",
    "TARGET_IDENTIFIER",
    "LOGICAL_ATTRIBUTE",
    "CATEGORY",
    "NX_ATTRIBUTE_NAME",
    "NX_ATTRIBUTE_TYPE",
    "AUDITED_CURRENT_VALUE",
    "ACTUAL_CURRENT_VALUE",
    "EXPECTED_VALUE",
    "AUTHORITATIVE_SOURCE",
    "APPROVED",
    "ENGINEER",
    "EVIDENCE_REFERENCE",
    "ACTION",
    "WRITE_ATTEMPTED",
    "ROLLBACK_RESULT",
    "REREAD_VALUE",
    "VERIFICATION_RESULT",
    "SAVE_RESULT",
    "MESSAGE",
]
PULL_COLUMNS = [
    "RUN_TIMESTAMP",
    "PART_NUMBER",
    "REVISION",
    "TARGET_OBJECT",
    "DRAWING_INDEX",
    "TARGET_IDENTIFIER",
    "LOGICAL_ATTRIBUTE",
    "CATEGORY",
    "NX_ATTRIBUTE_NAME",
    "NX_ATTRIBUTE_TYPE",
    "ATTRIBUTE_STATUS",
    "RAW_VALUE",
    "NORMALIZED_VALUE",
    "LOCKED",
    "OWNED_BY_SYSTEM",
    "PDM_BASED",
    "NOT_SAVED",
]
VALID_MODES = ("PULL", "DRY_RUN", "APPLY_APPROVED")
VALID_SOURCES = ("ENGINEERING_APPROVAL", "MODEL")
PROTECTED_LOGICAL_WRITES = {
    "part_number",
    "revision",
    "part_name",
    "part_description",
    "part_type",
    "reference_set",
    "mass_kg",
    "volume_mm3",
    "density_kg_per_mm3",
    "material",
    "material_missing_assignments",
    "material_multiple_assigned",
    "length_x_mm",
    "width_y_mm",
    "height_z_mm",
}


def _text(value):
    return "" if value is None else str(value)


def _clean(value):
    return _text(value).strip()


def _normalized_text(value):
    return " ".join(_clean(value).split()).upper()


def _enum_name(value):
    name = getattr(value, "name", None)
    return str(name) if name else _text(value).rsplit(".", 1)[-1]


def _session():
    return NXOpen.Session.GetSession()


def _listing_window(session):
    window = session.ListingWindow
    window.Open()
    return window


def _log(session, message):
    window = _listing_window(session)
    for line in _text(message).splitlines() or [""]:
        print(line)
        window.WriteFullline(line)


def _exception_details(error):
    code = getattr(error, "ErrorCode", None)
    if code is None:
        return "{0}: {1}".format(type(error).__name__, error)
    return "{0}: {1} (ErrorCode={2})".format(type(error).__name__, error, code)


def _journal_path():
    try:
        return os.path.abspath(__file__)
    except Exception:
        return ""


def _runtime_root():
    path = _journal_path()
    return os.path.dirname(os.path.dirname(path)) if path else os.getcwd()


def _desktop():
    profile = os.environ.get("USERPROFILE")
    if profile:
        return os.path.join(profile, "Desktop")
    home = os.path.expanduser("~")
    return os.path.join(home, "Desktop") if home and home != "~" else os.getcwd()


def _io_root():
    root = _clean(os.environ.get("NX_JOURNALS_IO_DIR")) or _desktop()
    os.makedirs(root, exist_ok=True)
    return root


def _config_path():
    return os.path.join(_runtime_root(), "config", "attribute_reconciliation.json")


def _load_config():
    path = _config_path()
    try:
        with open(path, "r", encoding="utf-8") as handle:
            config = json.load(handle)
    except (OSError, ValueError) as exc:
        raise RuntimeError("Unable to load reconciliation config {0}: {1}".format(path, exc))
    if config.get("schema_version") != 1 or config.get("authority") != "NX_TEAMCENTER":
        raise RuntimeError("Unsupported reconciliation config schema or authority.")
    if config.get("save_policy") not in ("NO_SAVE", "SAVE_CHANGED_PARTS"):
        raise RuntimeError("Invalid save_policy in reconciliation config.")
    rules = config.get("attributes")
    if not isinstance(rules, list) or not rules:
        raise RuntimeError("Reconciliation config has no attribute rules.")
    seen_logical = set()
    seen_keys = set()
    for rule in rules:
        required = ("logical_name", "category", "attribute", "type", "source_owner")
        if any(not _clean(rule.get(name)) for name in required):
            raise RuntimeError("Malformed reconciliation attribute rule.")
        key = (rule["category"], rule["attribute"])
        if rule["logical_name"] in seen_logical or key in seen_keys:
            raise RuntimeError("Duplicate reconciliation attribute rule.")
        seen_logical.add(rule["logical_name"])
        seen_keys.add(key)
    return config


def _write_csv(path, headers, rows):
    with open(path, "w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(headers)
        for row in rows:
            writer.writerow([row.get(header, "") for header in headers])
    return path


def _read_csv(path):
    last_error = None
    for encoding in ("utf-8-sig", "utf-8", "cp1252"):
        try:
            with open(path, "r", encoding=encoding, newline="") as handle:
                reader = csv.DictReader(handle)
                if reader.fieldnames is None:
                    raise RuntimeError("CSV has no header row: {0}".format(path))
                headers = [_clean(name) for name in reader.fieldnames]
                missing = [name for name in CORRECTION_COLUMNS if name not in headers]
                if missing:
                    raise RuntimeError("Correction CSV missing columns: {0}".format(", ".join(missing)))
                rows = []
                for row_number, source in enumerate(reader, 2):
                    row = {_clean(key): _clean(value) for key, value in source.items() if key is not None}
                    if not any(row.values()):
                        continue
                    row["__ROW_NUMBER__"] = row_number
                    rows.append(row)
                return rows
        except UnicodeDecodeError as exc:
            last_error = exc
    raise RuntimeError("Unable to decode correction CSV: {0}".format(last_error))


def _dispose(value):
    if value is None:
        return
    for name in ("Destroy", "Dispose", "FreeResource"):
        try:
            getattr(value, name)()
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


def _object_identifier(nx_object):
    for name in ("JournalIdentifier", "FullPath", "Leaf", "Name"):
        try:
            value = _clean(getattr(nx_object, name))
            if value:
                return value
        except Exception:
            pass
    return ""


def _object_key(nx_object):
    try:
        tag = getattr(nx_object, "Tag")
        if tag:
            return ("TAG", str(tag))
    except Exception:
        pass
    return ("ID", _object_identifier(nx_object) or str(id(nx_object)))


def _attribute_value(info):
    kind = _enum_name(getattr(info, "Type", ""))
    field = {
        "String": "StringValue",
        "Real": "RealValue",
        "Number": "RealValue",
        "Integer": "IntegerValue",
        "Boolean": "BooleanValue",
        "Time": "TimeValue",
        "Reference": "ReferenceValue",
    }.get(kind)
    return (getattr(info, field, "") if field else ""), kind


def _read_attribute(nx_object, rule):
    iterator = None
    try:
        iterator = nx_object.CreateAttributeIterator()
        iterator.SetIncludeOnlyCategory(rule["category"])
        iterator.SetIncludeOnlyTitle(rule["attribute"])
        iterator.SetIncludeAlsoUnset(True)
        infos = list(nx_object.GetUserAttributes(iterator))
        matches = [
            info
            for info in infos
            if _clean(getattr(info, "Category", "")) == rule["category"]
            and _clean(getattr(info, "Title", "")) == rule["attribute"]
        ]
        if not matches:
            return {"status": "MISSING", "raw": "", "type": rule["type"], "flags": {}}
        if len(matches) > 1:
            return {"status": "AMBIGUOUS", "raw": "", "type": rule["type"], "flags": {}}
        info = matches[0]
        raw, kind = _attribute_value(info)
        unset = bool(getattr(info, "Unset", False))
        return {
            "status": "UNSET" if unset else ("BLANK" if _clean(raw) == "" else "POPULATED"),
            "raw": raw,
            "type": kind,
            "flags": {
                "locked": bool(getattr(info, "Locked", False)),
                "owned_by_system": bool(getattr(info, "OwnedBySystem", False)),
                "pdm_based": bool(getattr(info, "PdmBased", False)),
                "not_saved": bool(getattr(info, "NotSaved", False)),
            },
        }
    except Exception as exc:
        return {
            "status": "UNREADABLE",
            "raw": "",
            "type": rule["type"],
            "flags": {},
            "message": _exception_details(exc),
        }
    finally:
        _dispose(iterator)


def _normalize(value, rule, config):
    raw = _clean(value)
    mode = rule.get("comparison", "NORMALIZED_TEXT")
    if mode == "EXACT":
        return raw
    if mode == "NUMBER":
        try:
            return format(float(raw), ".15g") if raw else ""
        except (TypeError, ValueError):
            return raw
    if mode == "BOOLEAN_ALIAS":
        upper = raw.upper()
        for canonical, aliases in config.get("release_policy", {}).get("boolean_aliases", {}).items():
            if upper in {_clean(alias).upper() for alias in aliases}:
                return canonical
        return upper
    if mode == "TRIMMED_CASE_INSENSITIVE":
        return raw.upper()
    return _normalized_text(raw)


def _validate_expected(value, rule, config):
    normalized = _normalize(value, rule, config)
    placeholders = {
        _normalized_text(item)
        for item in config.get("release_policy", {}).get("placeholder_values", [])
    }
    if _normalized_text(value) in placeholders:
        return "Expected value is blank or a configured placeholder."
    allowed = rule.get("allowed_values") or []
    if allowed and normalized not in {_normalize(item, rule, config) for item in allowed}:
        return "Expected value is outside the configured controlled vocabulary."
    if rule.get("validation") == "GREATER_THAN_ZERO":
        try:
            if float(value) <= 0:
                raise ValueError()
        except (TypeError, ValueError):
            return "Expected value must be numeric and greater than zero."
    return ""


def _root_component(part):
    try:
        return part.ComponentAssembly.RootComponent
    except Exception:
        return None


def _traverse(component):
    if component is None:
        return
    for child in component.GetChildren():
        try:
            suppressed = bool(child.IsSuppressed)
        except Exception:
            suppressed = False
        if suppressed:
            continue
        yield child
        yield from _traverse(child)


def _unique_model_parts(work_part):
    result = []
    seen = set()

    def add(part):
        if part is None:
            return
        key = _object_key(part)
        if key not in seen:
            seen.add(key)
            result.append(part)

    add(work_part)
    for component in _traverse(_root_component(work_part)):
        try:
            add(component.Prototype)
        except Exception:
            pass
    return result


def _rule_map(config):
    return {rule["logical_name"]: rule for rule in config["attributes"]}


def _identity(part, rule_map):
    pn = _clean(_read_attribute(part, rule_map["part_number"])["raw"])
    rev = _clean(_read_attribute(part, rule_map["revision"])["raw"])
    return pn, rev


def _model_index(work_part, rule_map):
    by_identity = {}
    by_part_number = {}
    for part in _unique_model_parts(work_part):
        pn, rev = _identity(part, rule_map)
        if not pn:
            continue
        by_identity.setdefault((_normalized_text(pn), _normalized_text(rev)), []).append(part)
        by_part_number.setdefault(_normalized_text(pn), set()).add(_normalized_text(rev))
    return by_identity, by_part_number


def _loaded_parts(session):
    try:
        return list(session.Parts)
    except Exception:
        try:
            return list(session.Parts.ToArray())
        except Exception:
            return []


def _drawing_identifier(config, pn, rev, index):
    return config["drawing"]["identifier_template"].format(
        part_number=pn, revision=rev, index=index
    )


def _find_loaded(session, identifier):
    expected = _normalized_text(identifier)
    for part in _loaded_parts(session):
        if _normalized_text(_object_identifier(part)) == expected:
            return part
    return None


def _open_drawing(session, config, pn, rev, index):
    identifier = _drawing_identifier(config, pn, rev, index)
    part = _find_loaded(session, identifier)
    if part is not None:
        return part, False, ""
    try:
        part = _unwrap(session.Parts.OpenDisplay(identifier))
        if part is None:
            return None, False, "OpenDisplay returned no part."
        if _normalized_text(_object_identifier(part)) != _normalized_text(identifier):
            return part, True, "OpenDisplay returned a different JournalIdentifier."
        return part, True, ""
    except Exception as exc:
        return None, False, _exception_details(exc)


def _close_part(part):
    try:
        part.Close(
            NXOpen.BasePart.CloseWholeTree.FalseValue,
            NXOpen.BasePart.CloseModified.CloseModified,
            None,
        )
        return ""
    except Exception as exc:
        return _exception_details(exc)


def _restore(session, display, work):
    errors = []
    if display is not None:
        try:
            _dispose(session.Parts.SetDisplay(display, False, True))
        except Exception as exc:
            errors.append("Display restore failed: " + _exception_details(exc))
    if work is not None:
        try:
            session.Parts.SetWork(work)
        except Exception as exc:
            errors.append("Work-part restore failed: " + _exception_details(exc))
    return errors


def _drawing_scope(config):
    default_name = config["inputs"]["drawing_scope_filename"]
    path = _clean(os.environ.get("NX_DRAWING_SCOPE_FILE")) or os.path.join(_io_root(), default_name)
    if not os.path.exists(path):
        return {}
    scope = {}
    for encoding in ("utf-8-sig", "utf-8", "cp1252"):
        try:
            with open(path, "r", encoding=encoding, newline="") as handle:
                for row in csv.DictReader(handle):
                    pn = _normalized_text(row.get("Item Number"))
                    rev = _normalized_text(row.get("Item Rev"))
                    decision = _normalized_text(row.get("Drawing Required"))
                    if decision in ("Y", "TRUE", "1"):
                        decision = "YES"
                    if decision in ("N", "FALSE", "0"):
                        decision = "NO"
                    if decision in ("YES", "NO"):
                        scope[(pn, rev)] = decision
            return scope
        except UnicodeDecodeError:
            continue
    return {}


def _base_report(timestamp, mode, row):
    return {
        "RUN_TIMESTAMP": timestamp,
        "MODE": mode,
        "AUDIT_RUN_ID": row.get("AUDIT_RUN_ID", ""),
        "PART_NUMBER": row.get("PART_NUMBER", ""),
        "REVISION": row.get("REVISION", ""),
        "TARGET_OBJECT": row.get("TARGET_OBJECT", ""),
        "DRAWING_INDEX": row.get("DRAWING_INDEX", ""),
        "TARGET_IDENTIFIER": "",
        "LOGICAL_ATTRIBUTE": "",
        "CATEGORY": row.get("CATEGORY", ""),
        "NX_ATTRIBUTE_NAME": row.get("NX_ATTRIBUTE_NAME", ""),
        "NX_ATTRIBUTE_TYPE": row.get("NX_ATTRIBUTE_TYPE", ""),
        "AUDITED_CURRENT_VALUE": row.get("CURRENT_VALUE_FROM_AUDIT", ""),
        "ACTUAL_CURRENT_VALUE": "",
        "EXPECTED_VALUE": row.get("EXPECTED_VALUE", ""),
        "AUTHORITATIVE_SOURCE": row.get("AUTHORITATIVE_SOURCE", ""),
        "APPROVED": row.get("APPROVED", ""),
        "ENGINEER": row.get("ENGINEER", ""),
        "EVIDENCE_REFERENCE": row.get("EVIDENCE_REFERENCE", ""),
        "ACTION": "ERROR",
        "WRITE_ATTEMPTED": "NO",
        "ROLLBACK_RESULT": "NOT_REQUIRED",
        "REREAD_VALUE": "",
        "VERIFICATION_RESULT": "NOT_RUN",
        "SAVE_RESULT": "NOT_RUN",
        "MESSAGE": "",
    }


def _resolve_rule(row, config):
    matches = [
        rule
        for rule in config["attributes"]
        if rule["category"] == row["CATEGORY"] and rule["attribute"] == row["NX_ATTRIBUTE_NAME"]
    ]
    if len(matches) != 1:
        return None, "Correction does not resolve to exactly one configured category/title rule."
    rule = matches[0]
    if _normalized_text(row["NX_ATTRIBUTE_TYPE"]) != _normalized_text(rule["type"]):
        return None, "Correction attribute type does not match configuration."
    return rule, ""


def _resolve_target(session, row, config, model_by_identity, revisions, opened):
    pn_key = _normalized_text(row["PART_NUMBER"])
    rev_key = _normalized_text(row["REVISION"])
    if row["TARGET_OBJECT"] == "MODEL":
        matches = model_by_identity.get((pn_key, rev_key), [])
        if len(matches) == 1:
            return matches[0], False, "", ""
        if len(matches) > 1:
            return None, False, "AMBIGUOUS_MATCH", "Multiple loaded models match part/revision."
        if pn_key in revisions:
            return None, False, "SKIPPED_REVISION_MISMATCH", "Part is loaded at a different revision."
        return None, False, "SKIPPED_TARGET_NOT_FOUND", "Exact model target was not found."
    try:
        index = int(row["DRAWING_INDEX"])
    except (TypeError, ValueError):
        return None, False, "SKIPPED_TARGET_NOT_FOUND", "Drawing target requires a numeric DRAWING_INDEX."
    if index < 1 or index > int(config["drawing"]["max_index"]):
        return None, False, "SKIPPED_TARGET_NOT_FOUND", "DRAWING_INDEX is outside configured range."
    target, newly_opened, error = _open_drawing(
        session, config, row["PART_NUMBER"], row["REVISION"], index
    )
    if target is None:
        return None, False, "SKIPPED_TARGET_NOT_FOUND", error
    if newly_opened and _object_key(target) not in {_object_key(part) for part in opened}:
        opened.append(target)
    if error:
        return target, newly_opened, "SKIPPED_TARGET_NOT_FOUND", error
    return target, newly_opened, "", ""


def _validate_row(session, timestamp, mode, row, config, rule_map, model_by_identity, revisions, opened):
    report = _base_report(timestamp, mode, row)
    row["TARGET_OBJECT"] = _normalized_text(row["TARGET_OBJECT"])
    row["AUTHORITATIVE_SOURCE"] = _normalized_text(row["AUTHORITATIVE_SOURCE"])
    report["TARGET_OBJECT"] = row["TARGET_OBJECT"]
    report["AUTHORITATIVE_SOURCE"] = row["AUTHORITATIVE_SOURCE"]
    if _normalized_text(row["APPROVED"]) != "YES":
        report.update(ACTION="SKIPPED_NOT_APPROVED", MESSAGE="APPROVED must be exactly YES.")
        return report, None
    if row["TARGET_OBJECT"] not in ("MODEL", "DRAWING"):
        report.update(ACTION="SKIPPED_TARGET_NOT_FOUND", MESSAGE="TARGET_OBJECT must be MODEL or DRAWING.")
        return report, None
    if row["AUTHORITATIVE_SOURCE"] not in VALID_SOURCES:
        report.update(
            ACTION="ERROR",
            MESSAGE="AUTHORITATIVE_SOURCE must be ENGINEERING_APPROVAL or MODEL; BOM/MASTER is prohibited.",
        )
        return report, None
    if row["AUTHORITATIVE_SOURCE"] == "MODEL" and row["TARGET_OBJECT"] != "DRAWING":
        report.update(ACTION="ERROR", MESSAGE="MODEL authority is valid only for a DRAWING target.")
        return report, None
    if not row["ENGINEER"] or not row["EVIDENCE_REFERENCE"] or not row["AUDIT_RUN_ID"]:
        report.update(ACTION="ERROR", MESSAGE="ENGINEER, EVIDENCE_REFERENCE, and AUDIT_RUN_ID are required.")
        return report, None
    if row["EXPECTED_VALUE"] == "":
        report.update(ACTION="ERROR", MESSAGE="Blank expected values are not supported.")
        return report, None
    rule, rule_error = _resolve_rule(row, config)
    if rule is None:
        report.update(ACTION="SKIPPED_NOT_WRITABLE", MESSAGE=rule_error)
        return report, None
    report["LOGICAL_ATTRIBUTE"] = rule["logical_name"]
    if (
        rule["logical_name"] in PROTECTED_LOGICAL_WRITES
        or not rule.get("writable")
        or row["TARGET_OBJECT"] not in rule.get("write_targets", [])
    ):
        report.update(ACTION="SKIPPED_NOT_WRITABLE", MESSAGE="Configuration prohibits this target/attribute write.")
        return report, None
    expected_error = _validate_expected(row["EXPECTED_VALUE"], rule, config)
    if expected_error:
        report.update(ACTION="ERROR", MESSAGE=expected_error)
        return report, None
    target, newly_opened, target_code, target_error = _resolve_target(
        session, row, config, model_by_identity, revisions, opened
    )
    if target is None or target_code:
        report.update(ACTION=target_code or "SKIPPED_TARGET_NOT_FOUND", MESSAGE=target_error)
        return report, None
    report["TARGET_IDENTIFIER"] = _object_identifier(target)
    actual = _read_attribute(target, rule)
    report["ACTUAL_CURRENT_VALUE"] = actual["raw"]
    if actual["status"] in ("UNREADABLE", "AMBIGUOUS"):
        report.update(ACTION="ERROR", MESSAGE=actual.get("message", actual["status"]))
        return report, None
    flags = actual.get("flags", {})
    if flags.get("locked") or flags.get("owned_by_system") or flags.get("pdm_based"):
        report.update(ACTION="SKIPPED_NOT_WRITABLE", MESSAGE="Runtime attribute flags prohibit writing.")
        return report, None
    if _text(actual["raw"]) != _text(row["CURRENT_VALUE_FROM_AUDIT"]):
        report.update(ACTION="STALE_AUDIT_VALUE", MESSAGE="Current raw value differs from the audited raw value.")
        return report, None
    if row["AUTHORITATIVE_SOURCE"] == "MODEL":
        models = model_by_identity.get(
            (_normalized_text(row["PART_NUMBER"]), _normalized_text(row["REVISION"])), []
        )
        if len(models) != 1:
            report.update(ACTION="SKIPPED_TARGET_NOT_FOUND", MESSAGE="Exact authoritative model was not found.")
            return report, None
        model_value = _read_attribute(models[0], rule)
        if _normalize(model_value["raw"], rule, config) != _normalize(row["EXPECTED_VALUE"], rule, config):
            report.update(ACTION="STALE_AUDIT_VALUE", MESSAGE="Expected value no longer matches the model authority.")
            return report, None
    if _normalize(actual["raw"], rule, config) == _normalize(row["EXPECTED_VALUE"], rule, config):
        report.update(
            ACTION="NO_CHANGE_ALREADY_MATCHES",
            REREAD_VALUE=actual["raw"],
            VERIFICATION_RESULT="PASS",
            SAVE_RESULT="NOT_REQUIRED",
            MESSAGE="Target already matches expected value.",
        )
        return report, None
    report.update(
        ACTION="PROPOSED_UPDATE",
        MESSAGE="Approved correction passed preflight." if mode == "DRY_RUN" else "Ready to apply.",
    )
    return report, {
        "source_row": row,
        "report": report,
        "rule": rule,
        "target": target,
        "newly_opened": newly_opened,
    }


def _builder_data_type(rule):
    name = _normalized_text(rule["type"])
    enum = NXOpen.AttributePropertiesBaseBuilder.DataTypeOptions
    if name == "BOOLEAN":
        return enum.Boolean
    if name == "INTEGER":
        return enum.Integer
    if name in ("NUMBER", "REAL"):
        return enum.Number
    if name in ("DATE", "TIME"):
        return enum.Date
    return enum.String


def _set_builder_value(builder, rule, value):
    kind = _normalized_text(rule["type"])
    if kind == "BOOLEAN":
        canonical = _normalized_text(value)
        builder.BooleanValue = (
            NXOpen.AttributePropertiesBaseBuilder.BooleanValueOptions.TrueValue
            if canonical in ("Y", "YES", "TRUE", "1")
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
                "NX reported {0} unsaved part(s) and {1} unsaved object(s).".format(
                    unsaved_parts, unsaved_objects
                )
            )
    finally:
        _dispose(status)


def _apply_groups(session, proposals, config):
    grouped = OrderedDict()
    for proposal in proposals:
        grouped.setdefault(_object_key(proposal["target"]), []).append(proposal)
    stop_saves = False
    # Keys returned by this function identify targets that may still contain
    # unsaved journal changes.  Journal-opened targets in this set must remain
    # open so that a save or rollback failure is visible and recoverable.
    unsaved_modified_keys = set()
    for group in grouped.values():
        target = group[0]["target"]
        mark_name = "J05 {0}".format(_object_identifier(target))
        mark = session.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, mark_name)
        failed = None
        try:
            for proposal in group:
                report = proposal["report"]
                report["WRITE_ATTEMPTED"] = "YES"
                _write_attribute(
                    session,
                    target,
                    proposal["rule"],
                    proposal["source_row"]["EXPECTED_VALUE"],
                )
                reread = _read_attribute(target, proposal["rule"])
                report["REREAD_VALUE"] = reread["raw"]
                if _normalize(reread["raw"], proposal["rule"], config) != _normalize(
                    proposal["source_row"]["EXPECTED_VALUE"], proposal["rule"], config
                ):
                    raise RuntimeError("Immediate reread did not match expected value.")
                report.update(
                    ACTION="UPDATED_VERIFIED",
                    VERIFICATION_RESULT="PASS",
                    MESSAGE="Attribute updated and reread successfully.",
                )
            unsaved_modified_keys.add(_object_key(target))
        except Exception as exc:
            failed = _exception_details(exc)
            try:
                session.UndoToMark(mark, mark_name)
                rollback = "PASS"
            except Exception as rollback_exc:
                rollback = "FAILED: " + _exception_details(rollback_exc)
            for proposal in group:
                proposal["report"].update(
                    ACTION="UPDATED_VERIFICATION_FAILED",
                    ROLLBACK_RESULT=rollback,
                    VERIFICATION_RESULT="FAIL",
                    SAVE_RESULT="NOT_ATTEMPTED",
                    MESSAGE=failed,
                )
            if rollback == "PASS":
                unsaved_modified_keys.discard(_object_key(target))
            else:
                unsaved_modified_keys.add(_object_key(target))
        finally:
            try:
                session.DeleteUndoMark(mark, mark_name)
            except Exception:
                pass
        if failed:
            continue
        if config["save_policy"] == "NO_SAVE":
            for proposal in group:
                proposal["report"]["SAVE_RESULT"] = "NO_SAVE"
            continue
        if stop_saves:
            for proposal in group:
                proposal["report"].update(
                    SAVE_RESULT="NOT_ATTEMPTED",
                    MESSAGE="A previous save failed; later saves were stopped.",
                )
            continue
        try:
            _save_target(target)
            unsaved_modified_keys.discard(_object_key(target))
            for proposal in group:
                proposal["report"]["SAVE_RESULT"] = "SAVED"
        except Exception as exc:
            stop_saves = True
            for proposal in group:
                proposal["report"].update(
                    ACTION="ERROR",
                    SAVE_RESULT="SAVE_FAILED_PART_LEFT_MODIFIED",
                    MESSAGE="Save failed; target remains visibly modified: " + _exception_details(exc),
                )
    return unsaved_modified_keys


def _run_pull(session, work_part, config, timestamp, opened):
    rule_map = _rule_map(config)
    models = _unique_model_parts(work_part)
    rows = []

    def add_target(part, target_object, drawing_index=""):
        pn, rev = _identity(part, rule_map)
        for rule in config["attributes"]:
            if target_object not in rule.get("required_on", []) and target_object not in rule.get(
                "write_targets", []
            ):
                continue
            result = _read_attribute(part, rule)
            flags = result.get("flags", {})
            rows.append(
                {
                    "RUN_TIMESTAMP": timestamp,
                    "PART_NUMBER": pn,
                    "REVISION": rev,
                    "TARGET_OBJECT": target_object,
                    "DRAWING_INDEX": drawing_index,
                    "TARGET_IDENTIFIER": _object_identifier(part),
                    "LOGICAL_ATTRIBUTE": rule["logical_name"],
                    "CATEGORY": rule["category"],
                    "NX_ATTRIBUTE_NAME": rule["attribute"],
                    "NX_ATTRIBUTE_TYPE": result["type"],
                    "ATTRIBUTE_STATUS": result["status"],
                    "RAW_VALUE": result["raw"],
                    "NORMALIZED_VALUE": _normalize(result["raw"], rule, config),
                    "LOCKED": flags.get("locked", ""),
                    "OWNED_BY_SYSTEM": flags.get("owned_by_system", ""),
                    "PDM_BASED": flags.get("pdm_based", ""),
                    "NOT_SAVED": flags.get("not_saved", ""),
                }
            )

    scope = _drawing_scope(config)
    for model in models:
        add_target(model, "MODEL")
        pn, rev = _identity(model, rule_map)
        if scope.get((_normalized_text(pn), _normalized_text(rev))) != "YES":
            continue
        for index in range(1, int(config["drawing"]["max_index"]) + 1):
            drawing, newly_opened, error = _open_drawing(session, config, pn, rev, index)
            if drawing is None or error:
                continue
            if newly_opened and _object_key(drawing) not in {_object_key(part) for part in opened}:
                opened.append(drawing)
            add_target(drawing, "DRAWING", index)
    output = os.path.join(_io_root(), "PULL_{0}.csv".format(timestamp.replace(":", "").replace("-", "")))
    _write_csv(output, PULL_COLUMNS, rows)
    _log(session, "PULL complete.\n  Rows: {0}\n  Report: {1}".format(len(rows), output))
    return set()


def main(session):
    work_part = session.Parts.Work
    if work_part is None:
        _log(session, "ERROR: No work part loaded.")
        return
    mode = _normalized_text(os.environ.get("NX_J05_MODE") or "DRY_RUN")
    if mode not in VALID_MODES:
        raise RuntimeError("NX_J05_MODE must be PULL, DRY_RUN, or APPLY_APPROVED.")
    config = _load_config()
    timestamp = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
    original_display = session.Parts.Display
    original_work = session.Parts.Work
    opened = []
    unsaved_modified_keys = set()
    reports = []
    try:
        if mode == "PULL":
            unsaved_modified_keys = _run_pull(session, work_part, config, timestamp, opened)
            return
        correction_name = config["inputs"]["corrections_filename"]
        correction_path = _clean(os.environ.get("NX_ATTRIBUTE_CORRECTIONS_FILE")) or os.path.join(
            _io_root(), correction_name
        )
        if not os.path.exists(correction_path):
            raise RuntimeError("Correction CSV not found: {0}".format(correction_path))
        rows = _read_csv(correction_path)
        rule_map = _rule_map(config)
        model_by_identity, revisions = _model_index(work_part, rule_map)
        proposals = []
        for row in rows:
            report, proposal = _validate_row(
                session,
                timestamp,
                mode,
                row,
                config,
                rule_map,
                model_by_identity,
                revisions,
                opened,
            )
            reports.append(report)
            if proposal is not None:
                proposals.append(proposal)
        if mode == "APPLY_APPROVED":
            unsaved_modified_keys = _apply_groups(session, proposals, config)
        report_stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_path = os.path.join(_io_root(), "J05_{0}_{1}.csv".format(mode, report_stamp))
        _write_csv(report_path, REPORT_COLUMNS, reports)
        counts = OrderedDict()
        for report in reports:
            action = report["ACTION"]
            counts[action] = counts.get(action, 0) + 1
        lines = [
            "J05 {0} complete.".format(mode),
            "  Save policy : {0}".format(config["save_policy"]),
            "  Input rows  : {0}".format(len(rows)),
            "  Report      : {0}".format(report_path),
        ]
        lines.extend("  {0}: {1}".format(action, count) for action, count in counts.items())
        _log(session, "\n".join(lines))
    finally:
        restore_errors = _restore(session, original_display, original_work)
        for part in reversed(opened):
            if _object_key(part) in unsaved_modified_keys:
                restore_errors.append(
                    "Journal-opened target left open because it may contain unsaved changes: {0}".format(
                        _object_identifier(part)
                    )
                )
                continue
            error = _close_part(part)
            if error:
                restore_errors.append("Unable to close {0}: {1}".format(_object_identifier(part), error))
        if restore_errors:
            _log(session, "\n".join("WARNING: " + message for message in restore_errors))


def _run_journal():
    session = None
    try:
        session = _session()
        _log(
            session,
            "\n".join(
                [
                    "J05 startup diagnostics",
                    "  Journal path : {0}".format(_journal_path() or "(unavailable)"),
                    "  Runtime root : {0}".format(_runtime_root()),
                    "  Config path  : {0}".format(_config_path()),
                    "  IO root      : {0}".format(_io_root()),
                    "  Python       : {0}".format(sys.version),
                ]
            ),
        )
        main(session)
    except Exception:
        message = "ERROR: Unhandled J05 exception.\n" + traceback.format_exc()
        if session is not None:
            _log(session, message)
        else:
            print(message)


if __name__ == "__main__":
    _run_journal()
