"""Journal 14 - approved Teamcenter Item Name updates through UF_UGMGR.

This journal is intentionally separate from Journal 05.  It changes Teamcenter
Item Name only by calling UF_UGMGR SetPartNameDesc on the Item database tag.
It never changes Item ID / Part Number, revision, description, geometry, or
business attributes.

Safety model:
- DRY_RUN is the default and never writes.
- APPLY_APPROVED writes only rows with APPROVED=YES and a nonblank ENGINEER.
- CURRENT_PART_NAME is an optimistic-lock baseline; stale rows are rejected.
- Duplicate PART_NUMBER rows are rejected.
- All approved rows must pass preflight before any write occurs.
- After each write, Item Name and description are reread and verified.
- A runtime write/verification failure stops all remaining writes.
- No automatic rollback is attempted.
"""

import csv
import os
import traceback
from collections import Counter
from datetime import datetime

import NXOpen
import NXOpen.UF


# ============================================================================
# USER SETTINGS - EDIT ONLY THESE TWO LINES FOR NORMAL NX USE
# Paste the full path of the Journal 14 input CSV between the quotes.
USER_PART_NAME_CSV = r"C:\Users\my62022696\Desktop\NX_PART_NAME_UPDATE.csv"
# Keep DRY_RUN for validation. The only other valid value is APPLY_APPROVED.
USER_MODE = "APPLY_APPROVED"
# Optional environment overrides:
#   NX_PART_NAME_UPDATE_FILE=<full CSV path>
#   NX_J14_MODE=DRY_RUN or APPLY_APPROVED
# ============================================================================

INPUT_COLUMNS = [
    "PART_NUMBER",
    "CURRENT_PART_NAME",
    "NEW_PART_NAME",
    "APPROVED",
    "ENGINEER",
    "APPROVAL_NOTE",
]

REPORT_COLUMNS = [
    "RUN_TIMESTAMP",
    "MODE",
    "CSV_ROW",
    "PART_NUMBER",
    "DATABASE_PART_TAG",
    "EXPECTED_CURRENT_PART_NAME",
    "ACTUAL_CURRENT_PART_NAME",
    "NEW_PART_NAME",
    "DESCRIPTION_BEFORE",
    "DESCRIPTION_AFTER",
    "APPROVED",
    "ENGINEER",
    "APPROVAL_NOTE",
    "METHOD",
    "ACTION",
    "WRITE_ATTEMPTED",
    "VERIFICATION_RESULT",
    "NX_EXCEPTION_TYPE",
    "NX_ERROR_CODE",
    "MESSAGE",
]

VALID_MODES = ("DRY_RUN", "APPLY_APPROVED")
METHOD_NAME = "UF_UGMGR.SetPartNameDesc"


def _text(value):
    return "" if value is None else str(value)


def _clean(value):
    return _text(value).strip()


def _normalized(value):
    return " ".join(_clean(value).split()).upper()


def _exception_fields(error):
    return type(error).__name__, _text(getattr(error, "ErrorCode", ""))


def _exception_details(error):
    exception_type, error_code = _exception_fields(error)
    code = ": {0}".format(error_code) if error_code else ""
    return "{0}{1} - {2}".format(exception_type, code, _text(error))


def configured_input_path():
    return _clean(os.environ.get("NX_PART_NAME_UPDATE_FILE") or USER_PART_NAME_CSV)


def configured_mode():
    return _normalized(os.environ.get("NX_J14_MODE") or USER_MODE or "DRY_RUN")


def _io_root():
    configured = _clean(os.environ.get("NX_JOURNALS_IO_DIR"))
    if configured:
        return os.path.abspath(configured)
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    return desktop if os.path.isdir(desktop) else os.getcwd()


def _read_csv(path):
    last_error = None
    for encoding in ("utf-8-sig", "utf-8", "cp1252"):
        try:
            with open(path, "r", encoding=encoding, newline="") as handle:
                reader = csv.DictReader(handle)
                headers = [_clean(name) for name in (reader.fieldnames or [])]
                missing = [column for column in INPUT_COLUMNS if column not in headers]
                if missing:
                    raise RuntimeError(
                        "Part Name CSV is missing columns: {0}".format(
                            ", ".join(missing)
                        )
                    )
                rows = []
                for row_number, source in enumerate(reader, 2):
                    row = {
                        _clean(key): _clean(value)
                        for key, value in source.items()
                        if key is not None
                    }
                    row["_CSV_ROW"] = row_number
                    rows.append(row)
                return rows
        except UnicodeDecodeError as exc:
            last_error = exc
    raise RuntimeError(
        "Unable to decode Part Name CSV: {0}".format(last_error or path)
    )


def _write_csv(path, rows):
    with open(path, "w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=REPORT_COLUMNS)
        writer.writeheader()
        for row in rows:
            writer.writerow({column: row.get(column, "") for column in REPORT_COLUMNS})


def _is_null_tag(tag_value):
    if tag_value is None:
        return True
    try:
        return int(tag_value) == 0
    except (TypeError, ValueError):
        return False


def _ask_part_tag(ugmgr, item_id):
    result = ugmgr.AskPartTag(item_id)
    if isinstance(result, tuple):
        return result[-1]
    return result


def _ask_part_name_desc(ugmgr, database_part_tag):
    result = ugmgr.AskPartNameDesc(database_part_tag)
    if isinstance(result, tuple) and len(result) >= 2:
        return _text(result[0]), _text(result[1])
    raise RuntimeError(
        "UF_UGMGR AskPartNameDesc returned an unexpected result: {0}".format(
            _text(result)
        )
    )


def _base_report(row, timestamp, mode):
    return {
        "RUN_TIMESTAMP": timestamp,
        "MODE": mode,
        "CSV_ROW": row.get("_CSV_ROW", ""),
        "PART_NUMBER": row.get("PART_NUMBER", ""),
        "DATABASE_PART_TAG": "",
        "EXPECTED_CURRENT_PART_NAME": row.get("CURRENT_PART_NAME", ""),
        "ACTUAL_CURRENT_PART_NAME": "",
        "NEW_PART_NAME": row.get("NEW_PART_NAME", ""),
        "DESCRIPTION_BEFORE": "",
        "DESCRIPTION_AFTER": "",
        "APPROVED": row.get("APPROVED", ""),
        "ENGINEER": row.get("ENGINEER", ""),
        "APPROVAL_NOTE": row.get("APPROVAL_NOTE", ""),
        "METHOD": METHOD_NAME,
        "ACTION": "",
        "WRITE_ATTEMPTED": "NO",
        "VERIFICATION_RESULT": "NOT_RUN",
        "NX_EXCEPTION_TYPE": "",
        "NX_ERROR_CODE": "",
        "MESSAGE": "",
    }


def _set_error(report, action, message, error=None):
    report["ACTION"] = action
    report["MESSAGE"] = message
    if error is not None:
        exception_type, error_code = _exception_fields(error)
        report["NX_EXCEPTION_TYPE"] = exception_type
        report["NX_ERROR_CODE"] = error_code
    return report


def _approval_state(row):
    value = _normalized(row.get("APPROVED", ""))
    if value == "YES":
        return "YES"
    if value in ("", "NO"):
        return "NO"
    return "INVALID"


def _validate_text_field(value, field_name):
    if not _clean(value):
        return "{0} is required.".format(field_name)
    if "\r" in _text(value) or "\n" in _text(value):
        return "{0} must not contain line breaks.".format(field_name)
    return ""


def _duplicate_part_numbers(rows):
    keys = [_normalized(row.get("PART_NUMBER", "")) for row in rows]
    counts = Counter(key for key in keys if key)
    return {key for key, count in counts.items() if count > 1}


def preflight_rows(ugmgr, rows, timestamp, mode):
    """Resolve Item tags and current names without making any changes."""
    duplicate_ids = _duplicate_part_numbers(rows)
    reports = []
    proposals = []

    for row in rows:
        report = _base_report(row, timestamp, mode)
        reports.append(report)

        part_number = _clean(row.get("PART_NUMBER", ""))
        expected_current = _clean(row.get("CURRENT_PART_NAME", ""))
        new_name = _clean(row.get("NEW_PART_NAME", ""))
        engineer = _clean(row.get("ENGINEER", ""))
        approval = _approval_state(row)

        error = _validate_text_field(part_number, "PART_NUMBER")
        if error:
            _set_error(report, "ERROR_INPUT", error)
            continue

        if _normalized(part_number) in duplicate_ids:
            _set_error(
                report,
                "ERROR_DUPLICATE_PART_NUMBER",
                "PART_NUMBER appears more than once in the input CSV. "
                "Each Teamcenter Item may appear only once per J14 run.",
            )
            continue

        error = _validate_text_field(expected_current, "CURRENT_PART_NAME")
        if error:
            _set_error(report, "ERROR_INPUT", error)
            continue

        error = _validate_text_field(new_name, "NEW_PART_NAME")
        if error:
            _set_error(report, "ERROR_INPUT", error)
            continue

        if approval == "INVALID":
            _set_error(
                report,
                "ERROR_APPROVAL_VALUE",
                "APPROVED must be YES, NO, or blank. Only YES authorizes a write.",
            )
            continue

        if approval == "YES" and not engineer:
            _set_error(
                report,
                "ERROR_ENGINEER_REQUIRED",
                "ENGINEER is required when APPROVED=YES.",
            )
            continue

        try:
            database_part_tag = _ask_part_tag(ugmgr, part_number)
            report["DATABASE_PART_TAG"] = _text(database_part_tag)
        except Exception as exc:
            _set_error(
                report,
                "ERROR_ITEM_LOOKUP",
                "UF_UGMGR could not resolve PART_NUMBER '{0}': {1}".format(
                    part_number, _exception_details(exc)
                ),
                exc,
            )
            continue

        if _is_null_tag(database_part_tag):
            _set_error(
                report,
                "ERROR_ITEM_NOT_FOUND",
                "AskPartTag returned a null database part tag for '{0}'.".format(
                    part_number
                ),
            )
            continue

        try:
            actual_current, description_before = _ask_part_name_desc(
                ugmgr, database_part_tag
            )
            report["ACTUAL_CURRENT_PART_NAME"] = actual_current
            report["DESCRIPTION_BEFORE"] = description_before
        except Exception as exc:
            _set_error(
                report,
                "ERROR_NAME_READ",
                "Could not read the current Teamcenter Item Name: {0}".format(
                    _exception_details(exc)
                ),
                exc,
            )
            continue

        # Idempotent rerun: if the requested new name is already present, do not
        # treat the old baseline as stale and never write again.
        if _clean(actual_current) == new_name:
            report["ACTION"] = "NO_CHANGE_ALREADY_AT_NEW_NAME"
            report["VERIFICATION_RESULT"] = "ALREADY_MATCHES"
            report["DESCRIPTION_AFTER"] = description_before
            report["MESSAGE"] = (
                "Teamcenter already contains NEW_PART_NAME; no write is required."
            )
            continue

        if _clean(actual_current) != expected_current:
            _set_error(
                report,
                "STALE_CURRENT_NAME",
                "Teamcenter Item Name does not match CURRENT_PART_NAME. "
                "Refresh the CSV before applying this rename.",
            )
            continue

        if approval != "YES":
            report["ACTION"] = "NOT_APPROVED"
            report["MESSAGE"] = "Validated successfully but APPROVED is not YES."
            continue

        report["ACTION"] = "READY_TO_UPDATE"
        report["MESSAGE"] = (
            "Approved row passed preflight. Item Name is an Item-level property "
            "and the rename can affect how all revisions of this Item are displayed."
        )
        proposals.append(
            {
                "row": row,
                "report": report,
                "tag": database_part_tag,
                "part_number": part_number,
                "old_name": actual_current,
                "old_description": description_before,
                "new_name": new_name,
            }
        )

    return reports, proposals


def _approved_preflight_failure(report):
    if _approval_state(report) != "YES":
        return False
    return report.get("ACTION") not in (
        "READY_TO_UPDATE",
        "NO_CHANGE_ALREADY_AT_NEW_NAME",
    )


def apply_proposals(ugmgr, reports, proposals):
    """Apply verified proposals sequentially; stop after the first runtime failure."""
    for index, proposal in enumerate(proposals):
        report = proposal["report"]
        report["WRITE_ATTEMPTED"] = "YES"

        try:
            # Verified by Journal 13 on Teamcenter X: blank description preserves
            # the existing description while changing Item Name.
            ugmgr.SetPartNameDesc(proposal["tag"], proposal["new_name"], "")
        except Exception as exc:
            _set_error(
                report,
                "ERROR_WRITE",
                "Teamcenter rejected the Item Name update: {0}".format(
                    _exception_details(exc)
                ),
                exc,
            )
            report["VERIFICATION_RESULT"] = "WRITE_FAILED"
            _mark_remaining_stopped(proposals[index + 1 :])
            break

        try:
            verified_name, verified_description = _ask_part_name_desc(
                ugmgr, proposal["tag"]
            )
            report["ACTUAL_CURRENT_PART_NAME"] = verified_name
            report["DESCRIPTION_AFTER"] = verified_description
        except Exception as exc:
            _set_error(
                report,
                "ERROR_VERIFICATION_READ",
                "Rename call returned, but Teamcenter read-back failed: {0}".format(
                    _exception_details(exc)
                ),
                exc,
            )
            report["VERIFICATION_RESULT"] = "READBACK_FAILED"
            _mark_remaining_stopped(proposals[index + 1 :])
            break

        name_ok = _clean(verified_name) == proposal["new_name"]
        description_ok = _text(verified_description) == _text(
            proposal["old_description"]
        )

        if name_ok and description_ok:
            report["ACTION"] = "UPDATED_VERIFIED"
            report["VERIFICATION_RESULT"] = "PASS"
            report["MESSAGE"] = (
                "Item Name changed and read-back verified. Description remained unchanged."
            )
            continue

        if name_ok and not description_ok:
            report["ACTION"] = "UPDATED_NAME_DESCRIPTION_CHANGED"
            report["VERIFICATION_RESULT"] = "FAIL_DESCRIPTION_CHANGED"
            report["MESSAGE"] = (
                "Item Name changed, but description did not remain unchanged. "
                "J14 stopped further writes. Review this Item in Teamcenter."
            )
        else:
            report["ACTION"] = "UPDATED_VERIFICATION_FAILED"
            report["VERIFICATION_RESULT"] = "FAIL_NAME_MISMATCH"
            report["MESSAGE"] = (
                "SetPartNameDesc returned, but read-back does not match NEW_PART_NAME. "
                "J14 stopped further writes."
            )

        _mark_remaining_stopped(proposals[index + 1 :])
        break

    return reports


def _mark_remaining_stopped(proposals):
    for proposal in proposals:
        report = proposal["report"]
        if report.get("ACTION") == "READY_TO_UPDATE":
            report["ACTION"] = "BATCH_STOPPED_AFTER_RUNTIME_FAILURE"
            report["MESSAGE"] = (
                "A previous approved rename failed during write/verification. "
                "No write was attempted for this row."
            )


def execute(ugmgr, rows, timestamp, mode):
    reports, proposals = preflight_rows(ugmgr, rows, timestamp, mode)

    if mode == "DRY_RUN":
        for proposal in proposals:
            proposal["report"]["ACTION"] = "DRY_RUN_READY"
            proposal["report"]["MESSAGE"] = (
                "Approved row passed preflight. DRY_RUN performed no write."
            )
        return reports

    # Fail closed before any write if any approved row is invalid or stale.
    hard_errors = [report for report in reports if _approved_preflight_failure(report)]
    if hard_errors:
        for proposal in proposals:
            report = proposal["report"]
            report["ACTION"] = "BATCH_ABORTED_PREFLIGHT"
            report["MESSAGE"] = (
                "Another APPROVED=YES row failed J14 preflight. No Item Names were changed."
            )
        return reports

    return apply_proposals(ugmgr, reports, proposals)


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


def _summary_counts(reports):
    counts = Counter(report.get("ACTION", "") or "<blank>" for report in reports)
    return counts


def main(session):
    mode = configured_mode()
    if mode not in VALID_MODES:
        raise RuntimeError(
            "USER_MODE (or NX_J14_MODE) must be DRY_RUN or APPLY_APPROVED."
        )

    input_path = configured_input_path()
    if not input_path:
        raise RuntimeError(
            "Edit USER_PART_NAME_CSV near the top of Journal 14 and paste the full "
            "path of the input CSV. Required columns: {0}. Advanced users may set "
            "NX_PART_NAME_UPDATE_FILE instead.".format(", ".join(INPUT_COLUMNS))
        )
    input_path = os.path.abspath(input_path)
    if not os.path.isfile(input_path):
        raise RuntimeError("Part Name update CSV not found: " + input_path)

    rows = _read_csv(input_path)
    if not rows:
        raise RuntimeError("Part Name update CSV contains no data rows: " + input_path)

    uf_session = NXOpen.UF.UFSession.GetUFSession()
    ugmgr = uf_session.Ugmgr

    timestamp = datetime.now().isoformat(timespec="seconds")
    reports = execute(ugmgr, rows, timestamp, mode)

    report_stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_root = _io_root()
    os.makedirs(output_root, exist_ok=True)
    report_path = os.path.join(
        output_root, "J14_PART_NAME_{0}_{1}.csv".format(mode, report_stamp)
    )
    _write_csv(report_path, reports)

    listing = _listing_window(session)
    _log(listing, "Journal 14 Teamcenter Item Name workflow complete.")
    _log(listing, "  Mode: " + mode)
    _log(listing, "  Input: " + input_path)
    _log(listing, "  Method: " + METHOD_NAME)
    _log(listing, "  Report: " + report_path)
    _log(listing, "  Rows: {0}".format(len(reports)))
    for action, count in sorted(_summary_counts(reports).items()):
        _log(listing, "    {0}: {1}".format(action, count))
    _log(listing, "Journal 14 never changes Item ID / Part Number or Revision.")
    _log(listing, "Journal 14 never changes description intentionally.")
    _log(listing, "Journal 14 performs no checkout, geometry save, check-in, or delete.")
    return report_path


def _run_journal():
    session = NXOpen.Session.GetSession()
    listing = _listing_window(session)
    try:
        main(session)
    except Exception as exc:
        _log(listing, "JOURNAL 14 FAILED: " + _exception_details(exc))
        _log(listing, traceback.format_exc())
        raise


if __name__ == "__main__":
    _run_journal()
