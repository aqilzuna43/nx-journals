"""Shared NX X 2506 implementation for J34/J35 CSV administrative freeze."""

import csv
import datetime
import hashlib
import json
import os
import re
import traceback

import NXOpen
import NXOpen.PDM


COMMON_BUILD = "NX-ADMIN-FREEZE-V1"
INPUT_FILENAME = "NX_ADMIN_FREEZE_SCOPE.csv"
MANIFEST_FILENAME = "NX_ADMIN_FREEZE_VALIDATION.json"
REPORT_FOLDER = "reports"
FREEZE_WORKFLOW = "Part_Freeze_Process"
ATTRIBUTE_CATEGORY = "WAEItem"
WAE_VERSION_TITLE = "WAE_VERSION"
VALID_YES = ("YES", "Y", "TRUE", "1")
VALID_NO = ("", "NO", "N", "FALSE", "0")
INPUT_ALIASES = {
    "freeze": ("FREEZE", "ENABLED", "ENABLE"),
    "part_number": ("DB_PART_NO", "ITEM NUMBER", "PART_NUMBER", "PART NUMBER"),
    "revision": ("DB_PART_REV", "ITEM REV", "REVISION"),
    "wae_version": ("WAE_VERSION",),
}
REPORT_COLUMNS = (
    "CSV_ROWS", "FREEZE", "DB_PART_NO", "DB_PART_REV", "CSV_WAE_VERSION",
    "ACTUAL_WAE_VERSION", "WAE_CLASS", "CHECKOUT_STATE", "CHECKOUT_OWNER",
    "RELEASE_STATUS", "RESULT", "MESSAGE", "RESOLUTION_HINT",
    "OPEN_SOURCE", "WORKFLOW_ERROR", "CLOSE_WARNING",
)


def clean(value):
    return "" if value is None else str(value).strip()


def error_text(error):
    code = clean(getattr(error, "ErrorCode", ""))
    suffix = " [{0}]".format(code) if code else ""
    return "{0}{1}".format(clean(error) or type(error).__name__, suffix)


def safe_property(value, name, default=None):
    try:
        result = getattr(value, name)
        return result() if callable(result) else result
    except Exception:
        return default


def dispose(value):
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


def package_root():
    return os.path.dirname(os.path.abspath(__file__))


def input_path():
    return os.path.join(package_root(), INPUT_FILENAME)


def manifest_path():
    return os.path.join(package_root(), MANIFEST_FILENAME)


def report_root():
    path = os.path.join(package_root(), REPORT_FOLDER)
    os.makedirs(path, exist_ok=True)
    return path


def file_sha256(path):
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        while True:
            block = handle.read(1024 * 1024)
            if not block:
                break
            digest.update(block)
    return digest.hexdigest()


def normalized_headers(fieldnames):
    lookup = {}
    for name in fieldnames or []:
        normalized = clean(name).upper()
        if normalized in lookup:
            raise RuntimeError("CSV contains duplicate header: " + clean(name))
        lookup[normalized] = name
    resolved = {}
    for logical, aliases in INPUT_ALIASES.items():
        matches = [lookup[alias] for alias in aliases if alias in lookup]
        if len(matches) > 1:
            raise RuntimeError(
                "CSV has ambiguous columns for {0}: {1}".format(
                    logical, ", ".join(matches)
                )
            )
        resolved[logical] = matches[0] if matches else None
    missing = [
        logical for logical in ("freeze", "part_number", "revision")
        if not resolved[logical]
    ]
    if missing:
        raise RuntimeError("CSV is missing required column(s): " + ", ".join(missing))
    return resolved


def read_scope(path):
    with open(path, "r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        if reader.fieldnames is None:
            raise RuntimeError("CSV does not contain a header row.")
        columns = normalized_headers(reader.fieldnames)
        rows = []
        for row_number, row in enumerate(reader, start=2):
            if not any(clean(value) for value in row.values()):
                continue
            rows.append({
                "csv_row": row_number,
                "freeze": clean(row.get(columns["freeze"])).upper(),
                "part_number": clean(row.get(columns["part_number"])),
                "revision": clean(row.get(columns["revision"])),
                "csv_wae_version": (
                    clean(row.get(columns["wae_version"]))
                    if columns["wae_version"] else ""
                ),
            })
    if not rows:
        raise RuntimeError("CSV contains no data rows.")
    return rows


def base_result(rows, freeze, number, revision, csv_wae):
    return {
        "csv_rows": list(rows),
        "freeze": freeze,
        "part_number": number,
        "revision": revision,
        "csv_wae_version": csv_wae,
        "actual_wae_version": "",
        "wae_class": "",
        "checkout_state": "",
        "checkout_owner": "",
        "release_status": [],
        "result": "PENDING",
        "message": "",
        "resolution_hint": "",
        "open_source": "",
        "workflow_error": "",
        "close_warning": "",
        "snapshot": {},
    }


def plan_scope(rows):
    grouped = {}
    malformed = []
    for row in rows:
        number = row["part_number"]
        revision = row["revision"]
        if not number or not revision:
            result = base_result(
                [row["csv_row"]], row["freeze"], number, revision,
                row["csv_wae_version"],
            )
            result.update({
                "result": "BLOCKED_MISSING_IDENTITY",
                "message": "DB_PART_NO and DB_PART_REV are required.",
                "resolution_hint": "Complete or disable this CSV row, then rerun J34.",
            })
            malformed.append(result)
            continue
        key = (number.casefold(), revision.casefold())
        grouped.setdefault(key, []).append(row)

    planned = list(malformed)
    for key in sorted(grouped):
        members = grouped[key]
        flags = {row["freeze"] for row in members}
        number = members[0]["part_number"]
        revision = members[0]["revision"]
        csv_rows = [row["csv_row"] for row in members]
        wae_by_normalized_value = {}
        for row in members:
            value = row["csv_wae_version"]
            if value:
                wae_by_normalized_value.setdefault(value.casefold(), value)
        nonblank_wae = list(wae_by_normalized_value.values())
        result = base_result(
            csv_rows, ";".join(sorted(flags)), number, revision,
            nonblank_wae[0] if len(nonblank_wae) == 1 else "",
        )
        if not flags.issubset(set(VALID_YES + VALID_NO)):
            result.update({
                "result": "BLOCKED_INVALID_FREEZE_FLAG",
                "message": "FREEZE must be YES or NO.",
                "resolution_hint": "Correct the FREEZE value and rerun J34.",
            })
        elif any(flag in VALID_YES for flag in flags) and any(
            flag in VALID_NO for flag in flags
        ):
            result.update({
                "result": "BLOCKED_CONFLICTING_FREEZE_FLAGS",
                "message": "Duplicate identity has both enabled and disabled rows.",
                "resolution_hint": "Make every duplicate row use the same FREEZE value.",
            })
        elif all(flag in VALID_NO for flag in flags):
            result.update({
                "result": "SKIPPED_DISABLED",
                "message": "All duplicate occurrences are disabled.",
                "resolution_hint": "No action required.",
            })
        elif len(nonblank_wae) > 1:
            result.update({
                "result": "BLOCKED_CONFLICTING_CSV_WAE",
                "message": "Duplicate identity has conflicting WAE_VERSION values: {0}.".format(
                    ", ".join(sorted(nonblank_wae))
                ),
                "resolution_hint": "Correct the CSV so the identity has one expected WAE value.",
            })
        planned.append(result)
    return planned


def classify_wae(value, revision):
    raw = clean(value)
    rev = clean(revision)
    if not raw:
        return "", "WAE_VERSION is blank."
    if re.fullmatch(r"[1-9][0-9]*", raw):
        return "NUMERIC_WORKING", ""
    if re.fullmatch(r"[A-Za-z]+", raw):
        if raw.casefold() == rev.casefold():
            return "ALPHABETIC_FINAL", ""
        return "", (
            "Alphabetic WAE_VERSION {0!r} does not match DB_PART_REV {1!r}.".format(
                raw, rev
            )
        )
    return "", "WAE_VERSION is neither a positive whole number nor a matching alphabetic revision."


def read_identity(part, title):
    method = getattr(part, "GetStringAttribute", None)
    if not callable(method):
        return ""
    try:
        return clean(method(title))
    except Exception:
        return ""


def attribute_type_name(info):
    raw = clean(safe_property(info, "Type"))
    normalized = raw.split(".")[-1].upper()
    return {"5": "STRING"}.get(normalized, normalized or "UNKNOWN")


def read_wae_attribute(part):
    iterator = None
    try:
        iterator = part.CreateAttributeIterator()
        iterator.SetIncludeOnlyCategory(ATTRIBUTE_CATEGORY)
        iterator.SetIncludeOnlyTitle(WAE_VERSION_TITLE)
        iterator.SetIncludeAlsoUnset(True)
        matches = []
        for info in part.GetUserAttributes(iterator):
            if (
                clean(safe_property(info, "Category")) == ATTRIBUTE_CATEGORY
                and clean(safe_property(info, "Title")) == WAE_VERSION_TITLE
            ):
                matches.append(info)
        if len(matches) != 1:
            raise RuntimeError(
                "Expected exactly one WAEItem/WAE_VERSION attribute; found {0}.".format(
                    len(matches)
                )
            )
        info = matches[0]
        value = clean(safe_property(info, "StringValue", ""))
        if attribute_type_name(info) != "STRING":
            raise RuntimeError("WAE_VERSION is not a string attribute.")
        if bool(safe_property(info, "Unset", False)) or not value:
            raise RuntimeError("WAE_VERSION is blank.")
        return value
    finally:
        dispose(iterator)


def checkout_result(raw):
    checked = None
    owner = ""
    if isinstance(raw, (tuple, list)):
        for value in raw:
            if checked is None and isinstance(value, bool):
                checked = value
            elif isinstance(value, str) and not owner:
                owner = clean(value)
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
        "CHECKED_OUT" if checked is True else "CHECKED_IN" if checked is False else "UNKNOWN",
        owner,
    )


def checkout_snapshot(part):
    pdm_part = safe_property(part, "PDMPart")
    method = getattr(pdm_part, "GetCheckedoutStatusAndUser", None)
    if not callable(method):
        return {"state": "UNKNOWN", "owner": "", "raw": "API unavailable"}
    try:
        raw = method()
    except TypeError:
        raw = method(False, "")
    state, owner = checkout_result(raw)
    return {"state": state, "owner": owner, "raw": repr(raw)[:2000]}


def release_status_snapshot(part):
    pdm_part = safe_property(part, "PDMPart")
    result = {"display": "", "internal": [], "errors": []}
    display_method = getattr(pdm_part, "GetReleaseStatus", None)
    internal_method = getattr(pdm_part, "GetInternalReleaseStatus", None)
    if not callable(display_method):
        result["errors"].append("PDMPart.GetReleaseStatus unavailable")
    else:
        try:
            result["display"] = clean(display_method())
        except Exception as error:
            result["errors"].append("GetReleaseStatus: " + error_text(error))
    if not callable(internal_method):
        result["errors"].append("PDMPart.GetInternalReleaseStatus unavailable")
    else:
        try:
            raw = internal_method([part])
            values = [raw] if isinstance(raw, str) else list(raw)
            result["internal"] = [clean(value) for value in values if clean(value)]
        except Exception as error:
            result["errors"].append("GetInternalReleaseStatus: " + error_text(error))
    return result


def status_values(snapshot):
    status = snapshot.get("release_status") or {}
    values = [clean(status.get("display"))] + list(status.get("internal") or [])
    return [clean(value) for value in values if clean(value)]


def is_frozen(snapshot):
    for value in status_values(snapshot):
        normalized = value.upper().replace(" ", "_").replace("-", "_")
        if "UNFREEZ" not in normalized and "UNFROZ" not in normalized and (
            "FREEZ" in normalized or "FROZ" in normalized
        ):
            return True
    return False


def has_other_status(snapshot):
    for value in status_values(snapshot):
        normalized = value.upper().replace(" ", "_").replace("-", "_")
        if "FREEZ" not in normalized and "FROZ" not in normalized:
            return True
        if "RELEAS" in normalized:
            return True
    return False


def part_snapshot(part):
    pdm_part = safe_property(part, "PDMPart")
    modifiable = None
    mod_error = ""
    method = getattr(pdm_part, "IsModifiable", None)
    if callable(method):
        try:
            modifiable = bool(method())
        except Exception as error:
            mod_error = error_text(error)
    else:
        mod_error = "PDMPart.IsModifiable unavailable"
    return {
        "part_number": read_identity(part, "DB_PART_NO"),
        "revision": read_identity(part, "DB_PART_REV"),
        "wae_version": read_wae_attribute(part),
        "checkout": checkout_snapshot(part),
        "release_status": release_status_snapshot(part),
        "read_only": safe_property(part, "IsReadOnly"),
        "pdm_modifiable": modifiable,
        "pdm_modifiable_error": mod_error,
        "modified": bool(safe_property(part, "IsModified", False)),
    }


def object_identity(value):
    tag = safe_property(value, "Tag")
    return ("TAG", clean(tag)) if tag is not None else ("PY", id(value))


def session_part_identities(session):
    try:
        return {object_identity(part) for part in list(session.Parts)}
    except Exception:
        return set()


def unwrap_open_result(value):
    if isinstance(value, (tuple, list)):
        return (value[0] if value else None, value[1] if len(value) > 1 else None)
    return value, None


def open_exact_part(session, number, revision):
    preloaded = session_part_identities(session)
    attempts = [
        "@DB/{0}/{1}".format(number, revision),
        "@DB/{0}/{1}/master".format(number, revision),
    ]
    errors = []
    for specification in attempts:
        part = None
        status = None
        try:
            part, status = unwrap_open_result(session.Parts.OpenBase(specification))
        except Exception as error:
            errors.append("{0}: {1}".format(specification, error_text(error)))
        finally:
            dispose(status)
        if part is not None:
            return {
                "part": part,
                "opened_by_journal": object_identity(part) not in preloaded,
                "source": specification,
            }
    raise RuntimeError("Could not open exact Teamcenter master. " + " | ".join(errors))


def close_opened_part(opened):
    if not opened or not opened.get("opened_by_journal"):
        return ""
    part = opened["part"]
    if bool(safe_property(part, "IsModified", False)):
        return "Journal-opened part became modified and was left open."
    try:
        part.Close(
            NXOpen.BasePart.CloseWholeTree.FalseValue,
            NXOpen.BasePart.CloseModified.CloseModified,
            None,
        )
        return ""
    except Exception as error:
        return "Could not close journal-opened part: " + error_text(error)


def get_workflows(session, part):
    method = getattr(safe_property(session, "PdmSession"), "GetNXWorkflows", None)
    if not callable(method):
        raise RuntimeError("PdmSession.GetNXWorkflows is unavailable.")
    errors = None
    try:
        raw = method([part])
        if not isinstance(raw, (tuple, list)) or len(raw) < 2:
            raise RuntimeError("GetNXWorkflows returned an unexpected result.")
        errors = raw[0]
        names = [raw[1]] if isinstance(raw[1], str) else list(raw[1])
        return [clean(name) for name in names if clean(name)]
    finally:
        dispose(errors)


def assign_freeze(session, part):
    method = getattr(safe_property(session, "PdmSession"), "AssignFreezeStatus", None)
    if not callable(method):
        raise RuntimeError("PdmSession.AssignFreezeStatus is unavailable.")
    errors = None
    try:
        errors = method([part], FREEZE_WORKFLOW)
        return repr(errors)[:2000]
    finally:
        dispose(errors)


def set_observed(result, snapshot):
    result["actual_wae_version"] = snapshot.get("wae_version", "")
    result["checkout_state"] = (snapshot.get("checkout") or {}).get("state", "")
    result["checkout_owner"] = (snapshot.get("checkout") or {}).get("owner", "")
    result["release_status"] = status_values(snapshot)
    result["snapshot"] = snapshot


def frozen_postcondition(snapshot, number, revision, wae):
    return (
        snapshot.get("part_number", "").casefold() == number.casefold()
        and snapshot.get("revision", "").casefold() == revision.casefold()
        and snapshot.get("wae_version", "").casefold() == wae.casefold()
        and (snapshot.get("checkout") or {}).get("state") == "CHECKED_IN"
        and snapshot.get("read_only") is True
        and snapshot.get("pdm_modifiable") is False
        and is_frozen(snapshot)
        and not has_other_status(snapshot)
    )


def validate_one(session, result):
    opened = None
    try:
        opened = open_exact_part(session, result["part_number"], result["revision"])
        result["open_source"] = opened["source"]
        snapshot = part_snapshot(opened["part"])
        set_observed(result, snapshot)
        if (
            snapshot["part_number"].casefold() != result["part_number"].casefold()
            or snapshot["revision"].casefold() != result["revision"].casefold()
        ):
            result.update({
                "result": "BLOCKED_IDENTITY_MISMATCH",
                "message": "Opened Teamcenter identity does not match the CSV.",
                "resolution_hint": "Correct the CSV identity or Teamcenter data, then rerun J34.",
            })
            return result
        wae_class, wae_error = classify_wae(snapshot["wae_version"], snapshot["revision"])
        result["wae_class"] = wae_class
        if wae_error:
            result.update({
                "result": (
                    "BLOCKED_MISSING_WAE_VERSION"
                    if not snapshot["wae_version"] else "BLOCKED_INVALID_WAE_VERSION"
                ),
                "message": wae_error,
                "resolution_hint": (
                    "Initialize an approved numeric WAE through J5, or disable the row."
                    if not snapshot["wae_version"] else
                    "Correct the lifecycle data; alphabetic WAE must match DB_PART_REV."
                ),
            })
            return result
        if result["csv_wae_version"] and (
            result["csv_wae_version"].casefold() != snapshot["wae_version"].casefold()
        ):
            result.update({
                "result": "BLOCKED_STALE_CSV_WAE",
                "message": "CSV WAE_VERSION {0!r} differs from live value {1!r}.".format(
                    result["csv_wae_version"], snapshot["wae_version"]
                ),
                "resolution_hint": "Regenerate/update the CSV, or correct editable live data through J5.",
            })
            return result
        status_errors = (snapshot.get("release_status") or {}).get("errors") or []
        if status_errors:
            result.update({
                "result": "BLOCKED_STATUS_QUERY",
                "message": " | ".join(status_errors),
                "resolution_hint": "Resolve the NX/Teamcenter status-query error and rerun J34.",
            })
            return result
        if snapshot["checkout"]["state"] != "CHECKED_IN":
            result.update({
                "result": "BLOCKED_CHECKED_OUT",
                "message": "Target is checked out by {0}.".format(
                    snapshot["checkout"].get("owner") or "an unknown user"
                ),
                "resolution_hint": "Have the owner review and check in the CAD, then rerun J34.",
            })
            return result
        if has_other_status(snapshot):
            result.update({
                "result": "BLOCKED_OTHER_RELEASE_STATUS",
                "message": "Target already has another controlled status: {0}.".format(
                    ", ".join(status_values(snapshot))
                ),
                "resolution_hint": "Resolve through the authorized Teamcenter lifecycle process.",
            })
            return result
        if is_frozen(snapshot):
            if frozen_postcondition(
                snapshot, result["part_number"], result["revision"],
                snapshot["wae_version"],
            ):
                result.update({
                    "result": "ALREADY_FROZEN",
                    "message": "Target is already a verified Frozen baseline.",
                    "resolution_hint": "No action required.",
                })
            else:
                result.update({
                    "result": "BLOCKED_INCONSISTENT_FROZEN_STATE",
                    "message": "Frozen status conflicts with checkout/read-only/modifiable state.",
                    "resolution_hint": "Inspect and repair the Teamcenter lifecycle state.",
                })
            return result
        if snapshot.get("pdm_modifiable_error"):
            result.update({
                "result": "BLOCKED_MODIFIABILITY_QUERY",
                "message": snapshot["pdm_modifiable_error"],
                "resolution_hint": "Resolve the NX/Teamcenter query error and rerun J34.",
            })
            return result
        workflows = get_workflows(session, opened["part"])
        if FREEZE_WORKFLOW not in workflows:
            result.update({
                "result": "BLOCKED_WORKFLOW_UNAVAILABLE",
                "message": "Required workflow is unavailable; found: {0}.".format(
                    ", ".join(workflows)
                ),
                "resolution_hint": "Ask the Teamcenter administrator to verify workflow availability.",
            })
            return result
        result.update({
            "result": "READY",
            "message": "Exact identity and live WAE are ready for administrative freeze.",
            "resolution_hint": "Run J35 after reviewing the validation report.",
        })
        return result
    except Exception as error:
        message = error_text(error)
        result.update({
            "result": (
                "BLOCKED_MISSING_WAE_VERSION"
                if "WAE_VERSION" in message and ("blank" in message or "found 0" in message)
                else "NOT_FOUND" if "Could not open exact" in message
                else "BLOCKED_VALIDATION_ERROR"
            ),
            "message": message,
            "resolution_hint": (
                "Initialize an approved numeric WAE through J5, or disable the row."
                if "WAE_VERSION" in message else
                "Review the exact identity and NX/Teamcenter error, then rerun J34."
            ),
        })
        return result
    finally:
        result["close_warning"] = close_opened_part(opened)


def result_count_map(results):
    counts = {}
    for result in results:
        key = result["result"]
        counts[key] = counts.get(key, 0) + 1
    return counts


def report_payload(mode, source_path, source_hash, results):
    return {
        "build": COMMON_BUILD,
        "mode": mode,
        "timestamp": datetime.datetime.now().astimezone().isoformat(timespec="seconds"),
        "input_csv": source_path,
        "input_sha256": source_hash,
        "workflow": FREEZE_WORKFLOW,
        "counts": result_count_map(results),
        "results": results,
    }


def csv_result_row(result):
    return {
        "CSV_ROWS": ";".join(str(value) for value in result.get("csv_rows") or []),
        "FREEZE": result.get("freeze", ""),
        "DB_PART_NO": result.get("part_number", ""),
        "DB_PART_REV": result.get("revision", ""),
        "CSV_WAE_VERSION": result.get("csv_wae_version", ""),
        "ACTUAL_WAE_VERSION": result.get("actual_wae_version", ""),
        "WAE_CLASS": result.get("wae_class", ""),
        "CHECKOUT_STATE": result.get("checkout_state", ""),
        "CHECKOUT_OWNER": result.get("checkout_owner", ""),
        "RELEASE_STATUS": ";".join(result.get("release_status") or []),
        "RESULT": result.get("result", ""),
        "MESSAGE": result.get("message", ""),
        "RESOLUTION_HINT": result.get("resolution_hint", ""),
        "OPEN_SOURCE": result.get("open_source", ""),
        "WORKFLOW_ERROR": result.get("workflow_error", ""),
        "CLOSE_WARNING": result.get("close_warning", ""),
    }


def write_outputs(payload):
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    prefix = "J34_VALIDATE" if payload["mode"] == "VALIDATE" else "J35_APPLY"
    root = report_root()
    json_path = os.path.join(root, "{0}_{1}.json".format(prefix, stamp))
    csv_path = os.path.join(root, "{0}_{1}.csv".format(prefix, stamp))
    with open(json_path, "w", encoding="utf-8") as handle:
        json.dump(payload, handle, indent=2, sort_keys=True)
    with open(csv_path, "w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=REPORT_COLUMNS)
        writer.writeheader()
        writer.writerows(csv_result_row(result) for result in payload["results"])
    return csv_path, json_path


def write_manifest(payload):
    with open(manifest_path(), "w", encoding="utf-8") as handle:
        json.dump(payload, handle, indent=2, sort_keys=True)


def read_manifest():
    path = manifest_path()
    if not os.path.isfile(path):
        raise RuntimeError("J34 validation manifest not found; run J34 first.")
    with open(path, "r", encoding="utf-8-sig") as handle:
        payload = json.load(handle)
    if payload.get("build") != COMMON_BUILD or payload.get("mode") != "VALIDATE":
        raise RuntimeError("Validation manifest build/mode is incompatible; rerun J34.")
    return payload


def clone_manifest_result(source):
    result = base_result(
        source.get("csv_rows") or [], source.get("freeze", ""),
        source.get("part_number", ""), source.get("revision", ""),
        source.get("csv_wae_version", ""),
    )
    for name in (
        "actual_wae_version", "wae_class", "checkout_state", "checkout_owner",
        "release_status", "result", "message", "resolution_hint",
    ):
        result[name] = source.get(name, result.get(name))
    return result


def apply_one(session, result, validated):
    if validated.get("result") not in ("READY", "ALREADY_FROZEN"):
        result["result"] = validated.get("result", "BLOCKED_NOT_VALIDATED")
        result["message"] = validated.get("message", "Row was not validated for apply.")
        result["resolution_hint"] = validated.get("resolution_hint", "Run J34 after correction.")
        return result

    opened = None
    try:
        opened = open_exact_part(session, result["part_number"], result["revision"])
        result["open_source"] = opened["source"]
        before = part_snapshot(opened["part"])
        set_observed(result, before)
        expected_wae = clean(validated.get("actual_wae_version"))
        if not expected_wae or before["wae_version"].casefold() != expected_wae.casefold():
            result.update({
                "result": "BLOCKED_LIVE_STATE_CHANGED",
                "message": "Live WAE_VERSION changed after J34 validation.",
                "resolution_hint": "Run J34 again against the current Teamcenter state.",
            })
            return result
        wae_class, wae_error = classify_wae(before["wae_version"], before["revision"])
        result["wae_class"] = wae_class
        if wae_error:
            result.update({
                "result": "BLOCKED_INVALID_WAE_VERSION",
                "message": wae_error,
                "resolution_hint": "Correct the lifecycle data and rerun J34.",
            })
            return result
        if before["checkout"]["state"] != "CHECKED_IN":
            result.update({
                "result": "BLOCKED_CHECKED_OUT",
                "message": "Target is checked out by {0}.".format(
                    before["checkout"].get("owner") or "an unknown user"
                ),
                "resolution_hint": "Have the owner review/check in the CAD, then rerun J34.",
            })
            return result
        if has_other_status(before):
            result.update({
                "result": "BLOCKED_OTHER_RELEASE_STATUS",
                "message": "Target has another controlled status: {0}.".format(
                    ", ".join(status_values(before))
                ),
                "resolution_hint": "Resolve through the authorized Teamcenter lifecycle process.",
            })
            return result
        if is_frozen(before):
            if frozen_postcondition(
                before, result["part_number"], result["revision"], expected_wae
            ):
                result.update({
                    "result": "ALREADY_FROZEN",
                    "message": "Target is already a verified Frozen baseline.",
                    "resolution_hint": "No action required.",
                })
            else:
                result.update({
                    "result": "BLOCKED_INCONSISTENT_FROZEN_STATE",
                    "message": "Frozen status conflicts with checkout/read-only/modifiable state.",
                    "resolution_hint": "Inspect and repair the Teamcenter lifecycle state.",
                })
            return result
        workflows = get_workflows(session, opened["part"])
        if FREEZE_WORKFLOW not in workflows:
            result.update({
                "result": "BLOCKED_WORKFLOW_UNAVAILABLE",
                "message": "Part_Freeze_Process is unavailable for this target.",
                "resolution_hint": "Ask the Teamcenter administrator to verify workflow availability.",
            })
            return result

        workflow_error = ""
        operation_raw = ""
        try:
            operation_raw = assign_freeze(session, opened["part"])
        except Exception as error:
            workflow_error = error_text(error)
        result["workflow_error"] = workflow_error or operation_raw

        try:
            after = part_snapshot(opened["part"])
            set_observed(result, after)
        except Exception as error:
            result.update({
                "result": "FAILED_FREEZE_VERIFICATION",
                "message": "Could not verify post-freeze state: " + error_text(error),
                "resolution_hint": "Inspect this part manually in NX/Teamcenter.",
            })
            return result
        if frozen_postcondition(
            after, result["part_number"], result["revision"], expected_wae
        ):
            result.update({
                "result": "FROZEN_WITH_WARNING" if workflow_error else "FROZEN",
                "message": (
                    "Verified Frozen despite workflow warning: " + workflow_error
                    if workflow_error else "Administrative freeze completed and verified."
                ),
                "resolution_hint": (
                    "Review the recorded warning; no repeat freeze is required."
                    if workflow_error else "No action required."
                ),
            })
        else:
            result.update({
                "result": "FAILED_FREEZE_WORKFLOW",
                "message": workflow_error or "Freeze workflow returned without a valid Frozen state.",
                "resolution_hint": "Review the Teamcenter error and freeze this identity manually if appropriate.",
            })
        return result
    except Exception as error:
        result.update({
            "result": "FAILED_APPLY_ERROR",
            "message": error_text(error),
            "resolution_hint": "Review the identity and NX/Teamcenter error; other rows continued.",
        })
        return result
    finally:
        result["close_warning"] = close_opened_part(opened)


def run_validation(session):
    path = input_path()
    if not os.path.isfile(path):
        raise RuntimeError("Input CSV not found beside J34/J35: " + path)
    source_hash = file_sha256(path)
    results = plan_scope(read_scope(path))
    for result in results:
        if result["result"] == "PENDING":
            validate_one(session, result)
    payload = report_payload("VALIDATE", path, source_hash, results)
    csv_path, json_path = write_outputs(payload)
    payload["report_csv"] = csv_path
    payload["report_json"] = json_path
    write_manifest(payload)
    return payload


def run_apply(session):
    path = input_path()
    if not os.path.isfile(path):
        raise RuntimeError("Input CSV not found beside J34/J35: " + path)
    source_hash = file_sha256(path)
    manifest = read_manifest()
    if source_hash != manifest.get("input_sha256"):
        raise RuntimeError("Input CSV changed after J34 validation; run J34 again.")
    results = []
    for validated in manifest.get("results") or []:
        result = clone_manifest_result(validated)
        apply_one(session, result, validated)
        results.append(result)
    payload = report_payload("APPLY", path, source_hash, results)
    csv_path, json_path = write_outputs(payload)
    payload["report_csv"] = csv_path
    payload["report_json"] = json_path
    return payload


def summary_text(payload):
    counts = payload.get("counts") or {}
    lines = [
        "Administrative freeze {0} complete".format(payload.get("mode", "").lower()),
        "",
    ]
    for name in sorted(counts):
        lines.append("{0}: {1}".format(name, counts[name]))
    lines.extend(["", "Report: " + payload.get("report_csv", "")])
    return "\n".join(lines)


def log_line(session, message):
    try:
        window = session.ListingWindow
        window.Open()
        window.WriteLine(str(message))
    except Exception:
        try:
            print(message)
        except Exception:
            pass


def show_summary(title, message, error=False):
    try:
        dialog_type = (
            NXOpen.NXMessageBox.DialogType.Error
            if error else NXOpen.NXMessageBox.DialogType.Information
        )
        NXOpen.UI.GetUI().NXMessageBox.Show(title, dialog_type, message)
    except Exception:
        pass


def run_ui(mode, wrapper_build):
    session = NXOpen.Session.GetSession()
    log_line(session, "=" * 72)
    log_line(session, "{0} | helper {1}".format(wrapper_build, COMMON_BUILD))
    log_line(session, "Input CSV: " + input_path())
    try:
        payload = run_validation(session) if mode == "VALIDATE" else run_apply(session)
        message = summary_text(payload)
        log_line(session, message)
        show_summary("NX Administrative Freeze", message)
        return payload
    except Exception as error:
        message = error_text(error)
        log_line(session, "FAILED: " + message)
        log_line(session, traceback.format_exc())
        show_summary("NX Administrative Freeze", message, error=True)
        return {"build": COMMON_BUILD, "mode": mode, "result": "FAILED", "message": message}
