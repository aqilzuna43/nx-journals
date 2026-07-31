"""J16 - standalone Teamcenter X drawing import.

NX X 2506 only.

Purpose:
- Import a locally edited native NX drawing back to the existing Teamcenter
  drawing specification.
- J16 does NOT require J15 to have performed the export.
- A J15 manifest may also be used directly because J16 requires only a subset
  of its columns.

Core safety rule:
- Every discovered object defaults to UseExisting.
- Only the exact local drawing is set to Overwrite.
- Item ID, revision, and 3D/reference parts are never intentionally replaced.

Run DRY_RUN first, approve its local and Teamcenter baseline hashes, then use
APPLY_ONE_APPROVED for one drawing or APPLY_APPROVED for a controlled batch in
a fresh NX session.
"""

import csv
import datetime
import hashlib
import os
import re
import shutil
import traceback
import zipfile
from collections import Counter

import NXOpen
import NXOpen.UF


# ============================================================================
# USER SETTINGS
# ============================================================================
USER_MODE = "DRY_RUN"  # DRY_RUN | APPLY_ONE_APPROVED | APPLY_APPROVED
USER_IMPORT_CSV = r""  # blank => <I/O root>\NX_TC_DRAWING_IMPORT.csv
# Optional environment overrides:
#   NX_TC_DRAWING_IMPORT_FILE=<full CSV path>
#   NX_J16_MODE=DRY_RUN, APPLY_ONE_APPROVED, or APPLY_APPROVED
#   NX_J16_MAX_APPROVED_WRITES=1..100 (APPLY_APPROVED only; default 25)
# ============================================================================

BUILD = "J16-TCX-DRAWING-IMPORT-NX2506-V10-CONTROLLED-BULK"
DEFAULT_INPUT = "NX_TC_DRAWING_IMPORT.csv"
VALID_MODES = ("DRY_RUN", "APPLY_ONE_APPROVED", "APPLY_APPROVED")
BATCH_APPLY_ENABLED = True
DEFAULT_MAX_APPROVED_WRITES = 25
DEFAULT_DATASET_TYPE = "UGPART"
DEFAULT_EXPORT_TOOL = "UGII V10-ALL"
RELATION_CANDIDATES = (
    "has specification",
    "specification",
    "IMAN_specification",
)

REQUIRED_COLUMNS = [
    "PART_NUMBER",
    "REVISION",
    "DWG_INDEX",
    "DRAWING_FILE",
    "APPROVED",
    "ENGINEER",
]

REPORT_COLUMNS = [
    "RUN_TIMESTAMP",
    "MODE",
    "CSV_ROW",
    "PART_NUMBER",
    "REVISION",
    "DWG_INDEX",
    "DRAWING_IDENTIFIER",
    "DRAWING_FILE",
    "BASELINE_SHA256",
    "APPROVED_LOCAL_SHA256",
    "APPROVED_TC_BASELINE_SHA256",
    "PREFLIGHT_SHA256",
    "CHANGED_FROM_BASELINE",
    "APPROVED",
    "ENGINEER",
    "DEFAULT_IMPORT_ACTION",
    "DRAWING_IMPORT_ACTION",
    "OPENED_IDENTIFIER",
    "CHECKOUT_STATE",
    "CHECKOUT_OWNER",
    "CHECKOUT_RAW",
    "VERIFICATION_CHANNEL",
    "RELATION_TYPE",
    "BASELINE_ASSOCIATED_FILES",
    "BASELINE_DOWNLOAD_CWD",
    "BASELINE_EXPORT_PDI_CODE",
    "BASELINE_EXPORT_FILE",
    "TC_BASELINE_SHA256",
    "CLONE_PREFLIGHT",
    "CHECKOUT_RECHECK_STATE",
    "CHECKOUT_RECHECK_OWNER",
    "PREWRITE_ASSOCIATED_FILES",
    "PREWRITE_DOWNLOAD_CWD",
    "PREWRITE_EXPORT_PDI_CODE",
    "PREWRITE_EXPORT_FILE",
    "PREWRITE_TC_SHA256",
    "POST_IMPORT_ASSOCIATED_FILES",
    "POST_IMPORT_DOWNLOAD_CWD",
    "POST_IMPORT_EXPORT_PDI_CODE",
    "POST_IMPORT_EXPORT_FILE",
    "POST_IMPORT_TC_SHA256",
    "POST_IMPORT_OPENED_IDENTIFIER",
    "POST_IMPORT_CHECKOUT_STATE",
    "POST_IMPORT_CHECKOUT_OWNER",
    "POST_IMPORT_CHECKOUT_RAW",
    "POST_IMPORT_VERIFICATION",
    "WRITE_ATTEMPTED",
    "RESULT",
    "MESSAGE",
    "CLONE_PREFLIGHT_LOG",
    "CLONE_APPLY_LOG",
    "CLONE_LOG",
]


def text(value):
    return "" if value is None else str(value)


def clean(value):
    return text(value).strip()


def upper(value):
    return clean(value).upper()


def env(name):
    return clean(os.environ.get(name))


def stamp():
    return datetime.datetime.now().strftime("%Y%m%d_%H%M%S")


def io_root():
    configured = env("NX_JOURNALS_IO_DIR")
    if configured:
        return os.path.abspath(os.path.expanduser(configured))
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    return desktop if os.path.isdir(desktop) else os.getcwd()


def configured_mode():
    return upper(env("NX_J16_MODE") or USER_MODE or "DRY_RUN")


def is_apply_mode(mode):
    return mode in ("APPLY_ONE_APPROVED", "APPLY_APPROVED")


def configured_max_approved_writes():
    raw = env("NX_J16_MAX_APPROVED_WRITES")
    try:
        value = int(raw) if raw else DEFAULT_MAX_APPROVED_WRITES
    except Exception:
        raise RuntimeError("NX_J16_MAX_APPROVED_WRITES must be an integer.")
    if value < 1 or value > 100:
        raise RuntimeError(
            "NX_J16_MAX_APPROVED_WRITES must be between 1 and 100."
        )
    return value


def configured_input_path():
    configured = env("NX_TC_DRAWING_IMPORT_FILE") or clean(USER_IMPORT_CSV)
    if configured:
        return os.path.abspath(os.path.expanduser(configured))
    return os.path.join(io_root(), DEFAULT_INPUT)


def configured_dataset_type():
    return clean(env("NX_J16_DATASET_TYPE") or DEFAULT_DATASET_TYPE)


def configured_export_tool():
    return clean(env("NX_J16_EXPORT_TOOL") or DEFAULT_EXPORT_TOOL)


def configured_relation_candidates():
    configured = clean(env("NX_J16_RELATION_TYPE"))
    return (configured,) if configured else RELATION_CANDIDATES


def error_text(error):
    code = clean(getattr(error, "ErrorCode", ""))
    suffix = ":{0}".format(code) if code else ""
    return "{0}{1} - {2}".format(type(error).__name__, suffix, text(error))


class Log:
    def __init__(self, session):
        self.lines = []
        try:
            self.window = session.ListingWindow
            self.window.Open()
        except Exception:
            self.window = None

    def write(self, message=""):
        message = text(message)
        self.lines.append(message)
        if self.window is not None:
            try:
                self.window.WriteFullline(message)
            except Exception:
                try:
                    self.window.WriteLine(message)
                except Exception:
                    pass
        try:
            print(message)
        except Exception:
            pass


def dispose(value):
    if value is not None:
        try:
            value.Dispose()
        except Exception:
            pass


def read_csv(path):
    last_decode_error = None
    for encoding in ("utf-8-sig", "utf-8", "cp1252"):
        try:
            with open(path, "r", encoding=encoding, newline="") as handle:
                reader = csv.DictReader(handle)
                headers = [clean(name) for name in (reader.fieldnames or [])]
                missing = [name for name in REQUIRED_COLUMNS if name not in headers]
                if missing:
                    raise RuntimeError(
                        "Import CSV is missing column(s): {0}".format(", ".join(missing))
                    )
                rows = []
                for row_number, source in enumerate(reader, 2):
                    row = {
                        clean(key): clean(value)
                        for key, value in source.items()
                        if key is not None
                    }
                    row["_CSV_ROW"] = row_number
                    rows.append(row)
                return rows
        except UnicodeDecodeError as exc:
            last_decode_error = exc
    raise RuntimeError(
        "Unable to decode import CSV: {0}: {1}".format(path, last_decode_error)
    )


def write_csv(path, rows):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=REPORT_COLUMNS)
        writer.writeheader()
        for row in rows:
            writer.writerow({column: row.get(column, "") for column in REPORT_COLUMNS})


def write_log(path, lines):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8-sig") as handle:
        handle.write("\n".join(lines) + "\n")


def zip_artifacts(zip_path, evidence_root, artifact_paths):
    """Create one portable evidence ZIP without including the source drawing."""
    os.makedirs(os.path.dirname(zip_path), exist_ok=True)
    temporary = zip_path + ".tmp"
    added = set()
    with zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as archive:
        if os.path.isdir(evidence_root):
            parent = os.path.dirname(evidence_root)
            for folder, directories, files in os.walk(evidence_root):
                directories.sort(key=lambda value: value.lower())
                files.sort(key=lambda value: value.lower())
                for name in files:
                    path = os.path.abspath(os.path.join(folder, name))
                    arcname = os.path.relpath(path, parent)
                    archive.write(path, arcname)
                    added.add(os.path.normcase(path))
        for value in artifact_paths:
            path = os.path.abspath(clean(value)) if clean(value) else ""
            if (
                path
                and os.path.isfile(path)
                and os.path.normcase(path) not in added
            ):
                archive.write(path, os.path.basename(path))
                added.add(os.path.normcase(path))
    os.replace(temporary, zip_path)
    return zip_path


def sha256(path):
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        while True:
            block = handle.read(1024 * 1024)
            if not block:
                break
            digest.update(block)
    return digest.hexdigest()


def valid_sha256(value):
    return bool(re.match(r"^[0-9a-fA-F]{64}$", clean(value)))


# ---------------------------------------------------------------------------
# Teamcenter identity / local file validation
# ---------------------------------------------------------------------------
def dataset_name(part_number, revision, drawing_index):
    return "{0}-{1}-dwg{2}".format(part_number, revision, drawing_index)


def drawing_id(part_number, revision, drawing_index):
    return "@DB/{0}/{1}/specification/{2}".format(
        part_number, revision, dataset_name(part_number, revision, drawing_index)
    )


def expected_native(part_number, revision, drawing_index):
    return "{0}_{1}_s_{2}.prt".format(
        part_number, revision, dataset_name(part_number, revision, drawing_index)
    )


def valid_native(path, part_number, revision, drawing_index):
    """Require Teamcenter AutoTranslate identity, while allowing a prefixed variant."""
    name = os.path.basename(path).lower()
    expected = expected_native(part_number, revision, drawing_index).lower()
    if name == expected:
        return True
    suffix = "_s_{0}-{1}-dwg{2}.prt".format(
        part_number, revision, drawing_index
    ).lower()
    return name.endswith(suffix)


def resolve_local_path(csv_path, value):
    path = os.path.expanduser(clean(value))
    if not path:
        return ""
    if not os.path.isabs(path):
        path = os.path.join(os.path.dirname(csv_path), path)
    return os.path.abspath(path)


def parse_target(row):
    part_number = clean(row.get("PART_NUMBER"))
    revision = clean(row.get("REVISION"))
    if not part_number or not revision:
        raise RuntimeError("PART_NUMBER and REVISION are required")
    try:
        drawing_index = int(clean(row.get("DWG_INDEX")))
    except Exception:
        raise RuntimeError("DWG_INDEX must be an integer")
    if drawing_index < 1:
        raise RuntimeError("DWG_INDEX must be >= 1")
    return part_number, revision, drawing_index


def approval_state(row):
    value = upper(row.get("APPROVED"))
    if value == "YES":
        return "YES"
    if value in ("", "NO"):
        return "NO"
    return "INVALID"


def duplicate_target_keys(rows):
    keys = []
    for row in rows:
        try:
            pn, rev, idx = parse_target(row)
            keys.append((upper(pn), upper(rev), idx))
        except Exception:
            continue
    counts = Counter(keys)
    return {key for key, count in counts.items() if count > 1}


def base_report(row, timestamp, mode):
    return {
        "RUN_TIMESTAMP": timestamp,
        "MODE": mode,
        "CSV_ROW": row.get("_CSV_ROW", ""),
        "PART_NUMBER": row.get("PART_NUMBER", ""),
        "REVISION": row.get("REVISION", ""),
        "DWG_INDEX": row.get("DWG_INDEX", ""),
        "DRAWING_IDENTIFIER": "",
        "DRAWING_FILE": row.get("DRAWING_FILE", ""),
        "BASELINE_SHA256": row.get("EXPORT_SHA256", ""),
        "APPROVED_LOCAL_SHA256": (
            row.get("APPROVED_LOCAL_SHA256", "")
            or row.get("PREFLIGHT_SHA256", "")
        ),
        "APPROVED_TC_BASELINE_SHA256": (
            row.get("APPROVED_TC_BASELINE_SHA256", "")
            or row.get("TC_BASELINE_SHA256", "")
        ),
        "PREFLIGHT_SHA256": "",
        "CHANGED_FROM_BASELINE": "",
        # A DRY_RUN report is an approval candidate, never prior authorization.
        # The operator must explicitly set one row to YES and name the engineer.
        "APPROVED": "NO" if mode == "DRY_RUN" else row.get("APPROVED", ""),
        "ENGINEER": "" if mode == "DRY_RUN" else row.get("ENGINEER", ""),
        "DEFAULT_IMPORT_ACTION": "UseExisting",
        "DRAWING_IMPORT_ACTION": "Overwrite",
        "OPENED_IDENTIFIER": "",
        "CHECKOUT_STATE": "NOT_CHECKED",
        "CHECKOUT_OWNER": "",
        "CHECKOUT_RAW": "",
        "VERIFICATION_CHANNEL": "GetAssociatedFiles+DownloadAssociatedFiles",
        "RELATION_TYPE": "",
        "BASELINE_ASSOCIATED_FILES": "",
        "BASELINE_DOWNLOAD_CWD": "",
        "BASELINE_EXPORT_PDI_CODE": "",
        "BASELINE_EXPORT_FILE": "",
        "TC_BASELINE_SHA256": "",
        "CLONE_PREFLIGHT": "NOT_RUN",
        "CHECKOUT_RECHECK_STATE": "NOT_RUN",
        "CHECKOUT_RECHECK_OWNER": "",
        "PREWRITE_ASSOCIATED_FILES": "",
        "PREWRITE_DOWNLOAD_CWD": "",
        "PREWRITE_EXPORT_PDI_CODE": "",
        "PREWRITE_EXPORT_FILE": "",
        "PREWRITE_TC_SHA256": "",
        "POST_IMPORT_ASSOCIATED_FILES": "",
        "POST_IMPORT_DOWNLOAD_CWD": "",
        "POST_IMPORT_EXPORT_PDI_CODE": "",
        "POST_IMPORT_EXPORT_FILE": "",
        "POST_IMPORT_TC_SHA256": "",
        "POST_IMPORT_OPENED_IDENTIFIER": "",
        "POST_IMPORT_CHECKOUT_STATE": "NOT_RUN",
        "POST_IMPORT_CHECKOUT_OWNER": "",
        "POST_IMPORT_CHECKOUT_RAW": "",
        "POST_IMPORT_VERIFICATION": "NOT_RUN",
        "WRITE_ATTEMPTED": "NO",
        "RESULT": "",
        "MESSAGE": "",
        "CLONE_PREFLIGHT_LOG": "",
        "CLONE_APPLY_LOG": "",
        "CLONE_LOG": "",
    }


def set_error(report, result, message, error=None):
    report["RESULT"] = result
    report["MESSAGE"] = message
    if error is not None:
        detail = error_text(error)
        if detail not in message:
            report["MESSAGE"] = "{0} | {1}".format(message, detail)
    return report


def local_preflight(rows, csv_path, timestamp, mode):
    duplicate_keys = duplicate_target_keys(rows)
    reports = []
    proposals = []

    for row in rows:
        report = base_report(row, timestamp, mode)
        reports.append(report)

        try:
            part_number, revision, drawing_index = parse_target(row)
        except Exception as exc:
            set_error(report, "ERROR_INPUT", error_text(exc), exc)
            continue

        approval = approval_state(row)
        if approval == "INVALID":
            set_error(
                report,
                "ERROR_APPROVAL_VALUE",
                "APPROVED must be YES, NO, or blank.",
            )
            continue

        # In an apply mode, unapproved rows are not candidates and cannot
        # block approved rows because of stale/missing local files.
        if is_apply_mode(mode) and approval != "YES":
            report["RESULT"] = "NOT_APPROVED"
            report["MESSAGE"] = "No write authorized for this row."
            continue

        target_key = (upper(part_number), upper(revision), drawing_index)
        if target_key in duplicate_keys:
            set_error(
                report,
                "ERROR_DUPLICATE_TARGET",
                "The same PART_NUMBER/REVISION/DWG_INDEX appears more than once.",
            )
            continue

        if is_apply_mode(mode) and not clean(row.get("ENGINEER")):
            set_error(
                report,
                "ERROR_ENGINEER_REQUIRED",
                "ENGINEER is required when APPROVED=YES.",
            )
            continue

        identifier = drawing_id(part_number, revision, drawing_index)
        supplied_identifier = clean(row.get("DRAWING_IDENTIFIER"))
        if supplied_identifier and upper(supplied_identifier) != upper(identifier):
            set_error(
                report,
                "ERROR_IDENTITY_MISMATCH",
                "DRAWING_IDENTIFIER does not match PART_NUMBER/REVISION/DWG_INDEX.",
            )
            continue
        report["DRAWING_IDENTIFIER"] = identifier

        drawing = resolve_local_path(csv_path, row.get("DRAWING_FILE"))
        report["DRAWING_FILE"] = drawing
        if not drawing or not os.path.isfile(drawing):
            set_error(
                report,
                "ERROR_FILE_NOT_FOUND",
                "DRAWING_FILE was not found: {0}".format(drawing or "<blank>"),
            )
            continue
        if not drawing.lower().endswith(".prt"):
            set_error(report, "ERROR_FILE_TYPE", "DRAWING_FILE must be a native NX .prt file.")
            continue
        if not valid_native(drawing, part_number, revision, drawing_index):
            set_error(
                report,
                "ERROR_NATIVE_FILENAME",
                "DRAWING_FILE does not match the expected Teamcenter AutoTranslate identity. "
                "Do not rename the exported drawing.",
            )
            continue

        try:
            current_sha = sha256(drawing)
        except Exception as exc:
            set_error(
                report,
                "ERROR_HASH_READ",
                "Could not read DRAWING_FILE for SHA-256.",
                exc,
            )
            continue

        report["PREFLIGHT_SHA256"] = current_sha
        if mode == "DRY_RUN":
            report["APPROVED_LOCAL_SHA256"] = current_sha
        elif is_apply_mode(mode):
            approved_local_sha = clean(report.get("APPROVED_LOCAL_SHA256"))
            if not valid_sha256(approved_local_sha):
                set_error(
                    report,
                    "ERROR_APPROVAL_HANDSHAKE_REQUIRED",
                    (
                        "Controlled apply requires APPROVED_LOCAL_SHA256 from "
                        "a successful J16 DRY_RUN report."
                    ),
                )
                continue
            if current_sha.lower() != approved_local_sha.lower():
                set_error(
                    report,
                    "BLOCKED_LOCAL_CHANGED_AFTER_APPROVAL",
                    (
                        "DRAWING_FILE SHA-256 no longer matches the approved "
                        "J16 DRY_RUN local hash."
                    ),
                )
                continue
        baseline = clean(row.get("EXPORT_SHA256"))
        if baseline:
            changed = current_sha.lower() != baseline.lower()
            report["CHANGED_FROM_BASELINE"] = "YES" if changed else "NO"
            if not changed:
                report["RESULT"] = "SKIPPED_UNCHANGED"
                report["MESSAGE"] = (
                    "Current drawing SHA-256 matches EXPORT_SHA256; no import is required."
                )
                continue
        else:
            report["CHANGED_FROM_BASELINE"] = "UNKNOWN"

        report["RESULT"] = "LOCAL_PREFLIGHT_OK"
        report["MESSAGE"] = (
            "Local identity checks passed. Target is the existing Teamcenter "
            "drawing specification; all related objects default to UseExisting."
        )
        proposals.append({
            "row": row,
            "report": report,
            "part_number": part_number,
            "revision": revision,
            "drawing_index": drawing_index,
            "dataset_name": dataset_name(
                part_number, revision, drawing_index
            ),
            "dataset_type": configured_dataset_type(),
            "export_tool": configured_export_tool(),
            "relation_type": "",
            "drawing": drawing,
            "identifier": identifier,
            "preflight_sha": current_sha,
        })

    return reports, proposals


# ---------------------------------------------------------------------------
# Exact Teamcenter target and checkout inspection
# ---------------------------------------------------------------------------
def journal_identifier(part):
    try:
        return clean(part.JournalIdentifier)
    except Exception:
        return ""


def unwrap_open_result(value):
    if isinstance(value, (tuple, list)):
        return (
            value[0] if value else None,
            value[1] if len(value) > 1 else None,
        )
    return value, None


def find_loaded_by_identifier(session, identifier):
    expected = upper(identifier).replace("\\", "/")
    try:
        parts = list(session.Parts)
    except Exception:
        parts = []
    for part in parts:
        actual = upper(journal_identifier(part)).replace("\\", "/")
        if actual == expected:
            return part
    try:
        candidate = session.Parts.FindObject(identifier)
    except Exception:
        candidate = None
    if candidate is not None:
        actual = upper(journal_identifier(candidate)).replace("\\", "/")
        if actual == expected:
            return candidate
    return None


def pdm_part(part):
    value = getattr(part, "PDMPart", None)
    return value() if callable(value) else value


def decode_checkout_result(raw):
    """Decode NX 2506 Python out-parameter shapes without guessing success."""
    checked = None
    owner = ""

    if isinstance(raw, dict):
        for key in ("isCheckedOut", "is_checked_out", "checkedOut", "checked_out"):
            if key in raw and isinstance(raw[key], bool):
                checked = raw[key]
                break
        for key in ("checkedOutBy", "checked_out_by", "owner", "user"):
            if key in raw:
                owner = clean(raw[key])
                break
    elif isinstance(raw, (tuple, list)):
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

    state = (
        "CHECKED_OUT"
        if checked is True
        else "CHECKED_IN"
        if checked is False
        else "UNKNOWN"
    )
    return {
        "state": state,
        "owner": owner,
        "raw": repr(raw)[:2000],
    }


def query_pdm_checkout(part):
    target_pdm = pdm_part(part)
    if target_pdm is None:
        return {
            "state": "UNKNOWN",
            "owner": "",
            "raw": "PDMPart unavailable",
        }
    method = getattr(target_pdm, "GetCheckedoutStatusAndUser", None)
    if not callable(method):
        return {
            "state": "UNKNOWN",
            "owner": "",
            "raw": "PDMPart.GetCheckedoutStatusAndUser unavailable",
        }
    try:
        raw = method()
    except TypeError:
        try:
            raw = method(False, "")
        except Exception as exc:
            return {
                "state": "UNKNOWN",
                "owner": "",
                "raw": error_text(exc),
            }
    except Exception as exc:
        return {
            "state": "UNKNOWN",
            "owner": "",
            "raw": error_text(exc),
        }
    return decode_checkout_result(raw)


def close_opened_part(part, log):
    if part is None:
        return
    try:
        whole_tree, _ = resolve_attr(
            NXOpen.BasePart.CloseWholeTree,
            ("FalseValue", "False_", "False", "CloseWholeTreeFalse"),
            "BasePart.CloseWholeTree false value",
        )
        close_modified, _ = resolve_attr(
            NXOpen.BasePart.CloseModified,
            ("CloseModified", "UseResponses", "UseLatest"),
            "BasePart.CloseModified value",
        )
        part.Close(whole_tree, close_modified, None)
    except Exception as exc:
        log.write(
            "  WARNING: could not close checkout-inspection drawing: {0}".format(
                error_text(exc)
            )
        )


def inspect_target_checkout(session, identifier, log):
    """Open the exact specification, verify identity, and query checkout owner."""
    part = find_loaded_by_identifier(session, identifier)
    load_status = None
    opened_here = False
    result = {
        "state": "UNKNOWN",
        "owner": "",
        "raw": "",
        "opened_identifier": "",
    }
    try:
        if part is None:
            part, load_status = unwrap_open_result(
                session.Parts.OpenBase(identifier)
            )
            opened_here = True
        if part is None:
            result["raw"] = "OpenBase returned no part"
            return result

        actual = journal_identifier(part)
        result["opened_identifier"] = actual
        if upper(actual).replace("\\", "/") != upper(identifier).replace("\\", "/"):
            result["raw"] = (
                "Opened JournalIdentifier does not match exact target: {0}".format(
                    actual or "<blank>"
                )
            )
            return result

        checkout = query_pdm_checkout(part)
        result.update(checkout)
        return result
    except Exception as exc:
        result["raw"] = error_text(exc)
        return result
    finally:
        dispose(load_status)
        if opened_here:
            close_opened_part(part, log)


# ---------------------------------------------------------------------------
# Read-only exact-dataset associated-file retrieval and verification
# ---------------------------------------------------------------------------
def new_file_management(session):
    pdm = getattr(session, "PdmSession", None)
    if pdm is None:
        raise RuntimeError(
            "NXOpen.Session.PdmSession is unavailable. Run J16 in managed mode."
        )
    method = getattr(pdm, "NewFileManagement", None)
    if not callable(method):
        raise RuntimeError("PdmSession.NewFileManagement is unavailable.")
    return pdm, method()


def safe_folder_name(value):
    return re.sub(r"[^A-Za-z0-9_.-]", "_", clean(value))


def managed_native_name(part_number, revision, drawing_index):
    return "{0}_{1}_dwg{2}.prt".format(
        part_number, revision, drawing_index
    )


def collect_pdm_files(value):
    files = []
    seen = set()

    def visit(candidate):
        if isinstance(candidate, (tuple, list)):
            for item in candidate:
                visit(item)
            return
        if callable(getattr(candidate, "GetFileName", None)):
            marker = id(candidate)
            if marker not in seen:
                seen.add(marker)
                files.append(candidate)

    visit(value)
    return files


def pdm_file_name(value):
    method = getattr(value, "GetFileName", None)
    if not callable(method):
        return ""
    try:
        return clean(method())
    except Exception:
        return ""


def release_pdm_files(files):
    for value in files:
        method = getattr(value, "FreeResource", None)
        if not callable(method):
            method = getattr(value, "Dispose", None)
        if callable(method):
            try:
                method()
            except Exception:
                pass


def phase_report_fields(phase):
    fields = {
        "BASELINE": (
            "BASELINE_ASSOCIATED_FILES",
            "BASELINE_DOWNLOAD_CWD",
        ),
        "PREWRITE": (
            "PREWRITE_ASSOCIATED_FILES",
            "PREWRITE_DOWNLOAD_CWD",
        ),
        "POSTIMPORT": (
            "POST_IMPORT_ASSOCIATED_FILES",
            "POST_IMPORT_DOWNLOAD_CWD",
        ),
    }
    if phase not in fields:
        raise RuntimeError("Unknown retrieval phase: {0}".format(phase))
    return fields[phase]


def retrieve_exact_associated_drawing(
    session,
    file_management,
    proposal,
    root,
    phase,
    pdi_field,
    file_field,
):
    """Download and fingerprint the exact native UGPART proven by J19 V2."""
    report = proposal["report"]
    associated_field, cwd_field = phase_report_fields(phase)
    evidence_root = os.path.join(root, "{0}_ASSOCIATED_FILES".format(phase))
    os.makedirs(evidence_root, exist_ok=True)
    expected_name = managed_native_name(
        proposal["part_number"],
        proposal["revision"],
        proposal["drawing_index"],
    )
    part = find_loaded_by_identifier(session, proposal["identifier"])
    load_status = None
    opened_here = False
    pdm_files = []
    original_cwd = os.getcwd()
    download_cwd = original_cwd
    try:
        if part is None:
            part, load_status = unwrap_open_result(
                session.Parts.OpenBase(proposal["identifier"])
            )
            opened_here = True
        if part is None:
            raise RuntimeError("OpenBase returned no part for exact specification.")
        actual = journal_identifier(part)
        if upper(actual).replace("\\", "/") != upper(
            proposal["identifier"]
        ).replace("\\", "/"):
            raise RuntimeError(
                "Opened JournalIdentifier does not match exact target: {0}".format(
                    actual or "<blank>"
                )
            )

        get_method = getattr(file_management, "GetAssociatedFiles", None)
        download_method = getattr(file_management, "DownloadAssociatedFiles", None)
        if not callable(get_method) or not callable(download_method):
            raise RuntimeError(
                "PDM GetAssociatedFiles/DownloadAssociatedFiles is unavailable."
            )

        raw_files = get_method([part], [])
        pdm_files = collect_pdm_files(raw_files)
        names = [pdm_file_name(value) for value in pdm_files]
        report[associated_field] = " | ".join(
            name or "<unreadable>" for name in names
        )
        exact_files = [
            value
            for value, name in zip(pdm_files, names)
            if os.path.basename(name).lower() == expected_name.lower()
        ]
        if len(exact_files) != 1:
            raise RuntimeError(
                "Expected exactly one associated native drawing named {0}; "
                "found {1}. Associated files: {2}".format(
                    expected_name,
                    len(exact_files),
                    report[associated_field] or "<none>",
                )
            )

        download_result = download_method([part], pdm_files)
        returned_files = collect_pdm_files(download_result)
        for value in returned_files:
            if all(value is not existing for existing in pdm_files):
                pdm_files.append(value)
        download_cwd = os.getcwd()
        report[cwd_field] = "{0} -> {1}".format(original_cwd, download_cwd)
        report[pdi_field] = "N/A_ASSOCIATED_FILES"
        report["RELATION_TYPE"] = "N/A_ASSOCIATED_FILES"

        candidate_names = [pdm_file_name(exact_files[0]), expected_name]
        candidate_names.extend(
            pdm_file_name(value)
            for value in returned_files
            if os.path.basename(pdm_file_name(value)).lower()
            == expected_name.lower()
        )
        physical = {}
        for name in candidate_names:
            if not name:
                continue
            candidates = [name] if os.path.isabs(name) else [
                os.path.join(download_cwd, name)
            ]
            for candidate in candidates:
                if os.path.isfile(candidate):
                    absolute = os.path.abspath(candidate)
                    physical[os.path.normcase(absolute)] = absolute
        if len(physical) != 1:
            raise RuntimeError(
                "DownloadAssociatedFiles did not materialize one unambiguous "
                "{0}; found {1} physical matches in download cwd {2}.".format(
                    expected_name, len(physical), download_cwd
                )
            )

        downloaded = next(iter(physical.values()))
        evidence_file = os.path.join(evidence_root, expected_name)
        shutil.copy2(downloaded, evidence_file)
        report[file_field] = evidence_file
        return evidence_file, sha256(evidence_file)
    finally:
        # NX 2506 DownloadAssociatedFiles changes the process working directory.
        # Always restore it; failure is safety-critical and must propagate.
        try:
            os.chdir(original_cwd)
        finally:
            release_pdm_files(pdm_files)
            dispose(load_status)
            if opened_here:
                close_opened_part(part, proposal.get("log") or _NullLog())


class _NullLog:
    def write(self, message=""):
        pass


# ---------------------------------------------------------------------------
# NX X 2506 UF Clone runtime enum resolution
# ---------------------------------------------------------------------------
def normalized_name(value):
    return "".join(ch.lower() for ch in clean(value) if ch.isalnum())


def public_names(value):
    names = []
    try:
        member_map = getattr(value, "__members__", None)
        if member_map:
            try:
                names.extend(list(member_map.keys()))
            except Exception:
                pass
    except Exception:
        pass
    try:
        names.extend([name for name in dir(value) if not name.startswith("_")])
    except Exception:
        pass

    result = []
    seen = set()
    for name in names:
        key = clean(name)
        if key and key not in seen:
            seen.add(key)
            result.append(key)
    return sorted(result)


def resolve_attr(container, candidate_names, label):
    names = public_names(container)
    by_normalized = {normalized_name(name): name for name in names}

    for candidate in candidate_names:
        if hasattr(container, candidate):
            return getattr(container, candidate), candidate

    for candidate in candidate_names:
        actual = by_normalized.get(normalized_name(candidate))
        if actual and hasattr(container, actual):
            return getattr(container, actual), actual

    raise RuntimeError(
        "Could not resolve {0}. Tried: {1}. Available runtime members: {2}".format(
            label,
            ", ".join(candidate_names),
            ", ".join(names) if names else "<none>",
        )
    )


def resolve_clone_api(ufs, log):
    clone_obj = getattr(ufs, "Clone", None)
    if clone_obj is None:
        raise RuntimeError("UFSession.Clone is unavailable in this NX X 2506 session.")

    clone_type = getattr(NXOpen.UF, "Clone", None)
    if clone_type is None:
        raise RuntimeError("NXOpen.UF.Clone is unavailable in this NX X 2506 binding.")

    operation_type, operation_type_name = resolve_attr(
        clone_type,
        ("OperationClass", "Operation", "OperationType"),
        "UF Clone operation enum type",
    )
    family_type, family_type_name = resolve_attr(
        clone_type,
        ("FamilyTreatment", "Family", "FamilyTreatmentType"),
        "UF Clone family-treatment enum type",
    )
    naming_type, naming_type_name = resolve_attr(
        clone_type,
        ("NamingTechnique", "Naming", "NamingType"),
        "UF Clone naming enum type",
    )
    action_type, action_type_name = resolve_attr(
        clone_type,
        ("Action", "CloneAction", "ActionType"),
        "UF Clone action enum type",
    )

    import_operation, import_name = resolve_attr(
        operation_type,
        ("ImportOperation", "Import", "OperationImport", "ImportOp"),
        "UF Clone import operation",
    )
    treat_as_lost, lost_name = resolve_attr(
        family_type,
        ("TreatAsLost", "AsLost", "Lost", "TreatLost"),
        "UF Clone TreatAsLost family treatment",
    )
    autotranslate, naming_name = resolve_attr(
        naming_type,
        ("Autotranslate", "AutoTranslate", "Auto_Translate", "AutomaticTranslate"),
        "UF Clone AutoTranslate naming technique",
    )
    use_existing, use_existing_name = resolve_attr(
        action_type,
        ("UseExisting", "UseExistingPart", "Existing", "UseExistingItem"),
        "UF Clone UseExisting action",
    )
    overwrite, overwrite_name = resolve_attr(
        action_type,
        ("Overwrite", "OverWrite", "Replace", "OverwriteExisting"),
        "UF Clone Overwrite action",
    )

    log.write("UF Clone binding: NXOpen.UF.Clone")
    log.write(
        "UF Clone resolved enums: {0}.{1}; {2}.{3}; {4}.{5}; {6}.{7}, {6}.{8}".format(
            operation_type_name,
            import_name,
            family_type_name,
            lost_name,
            naming_type_name,
            naming_name,
            action_type_name,
            use_existing_name,
            overwrite_name,
        )
    )

    return {
        "clone": clone_obj,
        "import_operation": import_operation,
        "treat_as_lost": treat_as_lost,
        "autotranslate": autotranslate,
        "use_existing": use_existing,
        "overwrite": overwrite,
    }


# ---------------------------------------------------------------------------
# UF Clone call wrappers / import
# ---------------------------------------------------------------------------
def terminate(clone):
    try:
        clone.Terminate()
    except Exception:
        pass


def add_assembly(clone, name):
    try:
        result = clone.AddAssembly(name)
    except TypeError:
        result = clone.AddAssembly(name, None)
    if isinstance(result, (tuple, list)):
        for value in result:
            if hasattr(value, "Dispose"):
                return value
    return None


def naming_failures(clone):
    try:
        result = clone.InitNamingFailures()
        if isinstance(result, (tuple, list)) and result:
            return result[-1]
        return result
    except Exception:
        return None


def perform_clone(clone, failures):
    try:
        return clone.PerformClone(failures)
    except TypeError:
        return clone.PerformClone(None)


def set_action(clone, part_name, action):
    try:
        return clone.SetAction(part_name, action, "")
    except TypeError:
        return clone.SetAction(part_name, action)


def iterate_parts(clone):
    parts = []
    try:
        clone.StartIteration()
    except Exception:
        return parts

    while True:
        try:
            result = clone.Iterate()
        except TypeError:
            try:
                result = clone.Iterate(None)
            except Exception:
                break
        except Exception:
            break

        if isinstance(result, (tuple, list)):
            part_name = ""
            for value in result:
                if isinstance(value, str):
                    part_name = value
        else:
            part_name = clean(result)

        if not part_name:
            break
        parts.append(part_name)
    return parts


def same_part(candidate, target):
    if not clean(candidate):
        return False
    try:
        if os.path.normcase(os.path.abspath(candidate)) == os.path.normcase(
            os.path.abspath(target)
        ):
            return True
    except Exception:
        pass
    return os.path.basename(candidate).lower() == os.path.basename(target).lower()


def import_one(api, drawing, logfile, dry_run, log):
    """Run one UF Clone import with refs UseExisting and exact drawing Overwrite."""
    clone = api["clone"]
    load_status = None
    folder = os.path.dirname(drawing)
    try:
        terminate(clone)
        clone.Initialise(api["import_operation"])
        clone.SetFamilyTreatment(api["treat_as_lost"])
        clone.SetDefNaming(api["autotranslate"])
        clone.SetDefItemType("")
        clone.SetDefDirectory(folder)
        try:
            clone.SetAssocFileRootDir(folder)
        except Exception:
            pass

        # Safety invariant: related 3D/reference parts are never defaulted to
        # overwrite. Only the exact drawing is promoted to Overwrite below.
        clone.SetDefAction(api["use_existing"])
        clone.SetDefAssocFileCopy(True)
        clone.SetLogfile(logfile)
        try:
            clone.SetPropagateActions(False)
        except Exception:
            pass

        load_status = add_assembly(clone, drawing)
        discovered_parts = iterate_parts(clone)
        drawing_action_set = False

        for part_name in discovered_parts:
            if same_part(part_name, drawing):
                set_action(clone, part_name, api["overwrite"])
                drawing_action_set = True
            else:
                # Default is already UseExisting, so failure of this redundant
                # per-object call does not weaken reference protection.
                try:
                    set_action(clone, part_name, api["use_existing"])
                except Exception:
                    pass

        if not drawing_action_set:
            # Same fallback used by the current J15 import implementation.
            set_action(clone, drawing, api["overwrite"])

        failures = naming_failures(clone)
        clone.SetDryrun(bool(dry_run))
        try:
            clone.GenerateReport()
        except Exception:
            pass
        perform_clone(clone, failures)

        log.write(
            "  UF Clone completed: discovered={0}; default=UseExisting; "
            "drawing=Overwrite; dry_run={1}".format(len(discovered_parts), dry_run)
        )
        return discovered_parts
    finally:
        dispose(load_status)
        terminate(clone)


def clone_log_path(proposal, mode, phase):
    return os.path.join(
        os.path.dirname(proposal["drawing"]),
        "J16_{0}_{1}_{2}_{3}_DWG{4}.clone".format(
            phase,
            mode,
            proposal["part_number"],
            proposal["revision"],
            proposal["drawing_index"],
        ),
    )


def target_evidence_root(work_root, proposal):
    return os.path.join(
        work_root,
        safe_folder_name(
            "{0}_{1}_DWG{2}".format(
                proposal["part_number"],
                proposal["revision"],
                proposal["drawing_index"],
            )
        ),
    )


def record_checkout(report, checkout, recheck=False):
    if recheck:
        report["CHECKOUT_RECHECK_STATE"] = checkout.get("state", "UNKNOWN")
        report["CHECKOUT_RECHECK_OWNER"] = checkout.get("owner", "")
    else:
        report["OPENED_IDENTIFIER"] = checkout.get("opened_identifier", "")
        report["CHECKOUT_STATE"] = checkout.get("state", "UNKNOWN")
        report["CHECKOUT_OWNER"] = checkout.get("owner", "")
        report["CHECKOUT_RAW"] = checkout.get("raw", "")


def record_post_import_checkout(report, checkout):
    report["POST_IMPORT_OPENED_IDENTIFIER"] = checkout.get(
        "opened_identifier", ""
    )
    report["POST_IMPORT_CHECKOUT_STATE"] = checkout.get("state", "UNKNOWN")
    report["POST_IMPORT_CHECKOUT_OWNER"] = checkout.get("owner", "")
    report["POST_IMPORT_CHECKOUT_RAW"] = checkout.get("raw", "")


def block_for_checkout(report, checkout, phase):
    state = checkout.get("state", "UNKNOWN")
    owner = checkout.get("owner", "")
    if state == "CHECKED_OUT":
        set_error(
            report,
            "BLOCKED_CHECKED_OUT",
            (
                "Exact drawing specification is checked out by {0}; "
                "J16 blocks every existing checkout before {1}."
            ).format(owner or "<owner unavailable>", phase),
        )
        return True
    if state != "CHECKED_IN":
        set_error(
            report,
            "BLOCKED_CHECKOUT_UNKNOWN",
            (
                "Exact drawing specification checkout state could not be "
                "proven CHECKED_IN before {0}. Raw status: {1}"
            ).format(phase, checkout.get("raw", "") or "<none>"),
        )
        return True
    return False


def run_managed_preflight(
    session,
    file_management,
    proposals,
    work_root,
    log,
):
    for proposal in proposals:
        report = proposal["report"]
        proposal["log"] = log
        if report.get("RESULT") != "LOCAL_PREFLIGHT_OK":
            continue

        checkout = inspect_target_checkout(
            session, proposal["identifier"], log
        )
        record_checkout(report, checkout)
        log.write(
            "  CHECKOUT {0}: state={1}; owner={2}; opened={3}".format(
                proposal["identifier"],
                checkout.get("state", "UNKNOWN"),
                checkout.get("owner", "") or "<blank>",
                checkout.get("opened_identifier", "") or "<blank>",
            )
        )
        if block_for_checkout(report, checkout, "clone preflight"):
            log.write(
                "  BLOCKED {0}: {1}".format(
                    proposal["identifier"], report["MESSAGE"]
                )
            )
            continue

        try:
            _, baseline_sha = retrieve_exact_associated_drawing(
                session,
                file_management,
                proposal,
                target_evidence_root(work_root, proposal),
                "BASELINE",
                "BASELINE_EXPORT_PDI_CODE",
                "BASELINE_EXPORT_FILE",
            )
            report["TC_BASELINE_SHA256"] = baseline_sha
            proposal["tc_baseline_sha"] = baseline_sha
            if report.get("MODE") == "DRY_RUN":
                report["APPROVED_TC_BASELINE_SHA256"] = baseline_sha
            log.write(
                "  BASELINE {0}: channel=associated-files; sha256={1}".format(
                    proposal["identifier"],
                    baseline_sha,
                )
            )
        except Exception as exc:
            set_error(
                report,
                "FAILED_TARGET_BASELINE_RETRIEVAL",
                (
                    "Could not retrieve and fingerprint the exact Teamcenter "
                    "native drawing before clone preflight."
                ),
                exc,
            )
            log.write(
                "  BLOCKED {0}: {1}".format(
                    proposal["identifier"], report["MESSAGE"]
                )
            )
            continue

        if is_apply_mode(report.get("MODE")):
            approved_tc_sha = clean(
                report.get("APPROVED_TC_BASELINE_SHA256")
            )
            if not valid_sha256(approved_tc_sha):
                set_error(
                    report,
                    "ERROR_APPROVAL_HANDSHAKE_REQUIRED",
                    (
                        "Controlled apply requires "
                        "APPROVED_TC_BASELINE_SHA256 from a successful J16 "
                        "DRY_RUN report."
                    ),
                )
                log.write(
                    "  BLOCKED {0}: {1}".format(
                        proposal["identifier"], report["MESSAGE"]
                    )
                )
                continue
            if approved_tc_sha.lower() != baseline_sha.lower():
                set_error(
                    report,
                    "BLOCKED_STALE_TARGET",
                    (
                        "The exact Teamcenter drawing changed after the approved "
                        "J16 DRY_RUN. No write was attempted."
                    ),
                )
                log.write(
                    "  BLOCKED {0}: {1}".format(
                        proposal["identifier"], report["MESSAGE"]
                    )
                )
                continue

        if proposal["preflight_sha"].lower() == baseline_sha.lower():
            report["RESULT"] = "SKIPPED_ALREADY_CURRENT"
            report["MESSAGE"] = (
                "The exact Teamcenter drawing already matches the approved "
                "local drawing; no import is required."
            )
            continue

        report["RESULT"] = "MANAGED_PREFLIGHT_OK"
        report["MESSAGE"] = (
            "Exact target identity and CHECKED_IN state were proven; current "
            "Teamcenter native drawing was downloaded and fingerprinted."
        )


def run_dry_run(api, proposals, log, mode):
    for proposal in proposals:
        report = proposal["report"]
        if report.get("RESULT") != "MANAGED_PREFLIGHT_OK":
            continue
        logfile = clone_log_path(proposal, mode, "PREFLIGHT")
        report["CLONE_LOG"] = logfile
        report["CLONE_PREFLIGHT_LOG"] = logfile
        try:
            import_one(api, proposal["drawing"], logfile, True, log)
            report["CLONE_PREFLIGHT"] = "PASS"
            report["RESULT"] = "DRY_RUN_OK" if mode == "DRY_RUN" else "CLONE_PREFLIGHT_OK"
            report["MESSAGE"] = (
                "UF Clone dry run passed. All related objects default to UseExisting; "
                "only the exact drawing is Overwrite."
            )
        except Exception as exc:
            report["CLONE_PREFLIGHT"] = "FAIL"
            set_error(
                report,
                "FAILED_CLONE_PREFLIGHT",
                "UF Clone dry run failed.",
                exc,
            )
            log.write("  FAILED dry run: {0}".format(error_text(exc)))
            log.write(traceback.format_exc())
    return proposals


def mark_remaining_after_stopped_write(proposals, start, review_required):
    for proposal in proposals[start:]:
        report = proposal["report"]
        if report.get("RESULT") == "CLONE_PREFLIGHT_OK":
            if review_required:
                report["RESULT"] = "REVIEW_NOT_ATTEMPTED_AFTER_PRIOR_WRITE"
                report["MESSAGE"] = (
                    "A previous write requires manual acceptance. No write "
                    "was attempted for this row."
                )
            else:
                report["RESULT"] = "BATCH_STOPPED_AFTER_UNVERIFIED_WRITE"
                report["MESSAGE"] = (
                    "A previous write was attempted but could not be verified. "
                    "No write was attempted for this row."
                )


def classify_post_import(
    proposal,
    prewrite_sha,
    post_sha,
    checkout,
    log,
):
    """Classify persistence separately from managed-mode byte transformation."""
    report = proposal["report"]
    source_sha = proposal["preflight_sha"]
    record_post_import_checkout(report, checkout)
    state = checkout.get("state", "UNKNOWN")
    owner = checkout.get("owner", "")

    log.write(
        "  POSTIMPORT CHECKOUT {0}: state={1}; owner={2}; opened={3}".format(
            proposal["identifier"],
            state,
            owner or "<blank>",
            checkout.get("opened_identifier", "") or "<blank>",
        )
    )

    if post_sha.lower() == prewrite_sha.lower():
        report["POST_IMPORT_VERIFICATION"] = "FAILED_UNCHANGED_FROM_PREWRITE"
        set_error(
            report,
            "FAILED_IMPORT_UNVERIFIED",
            (
                "UF Clone returned without an exception, but the exact managed "
                "drawing payload is unchanged from immediately before the write."
            ),
        )
        return False

    exact_source = post_sha.lower() == source_sha.lower()
    if state not in ("CHECKED_IN", "CHECKED_OUT"):
        report["POST_IMPORT_VERIFICATION"] = "FAILED_POST_CHECKOUT_UNKNOWN"
        set_error(
            report,
            "FAILED_IMPORT_UNVERIFIED",
            (
                "The exact managed payload changed after UF Clone, but its "
                "post-import checkout state could not be proven. Raw status: {0}"
            ).format(checkout.get("raw", "") or "<none>"),
        )
        return False

    if state == "CHECKED_OUT":
        report["POST_IMPORT_VERIFICATION"] = (
            "VERIFIED_SHA256_CHECKIN_REQUIRED"
            if exact_source
            else "MANUAL_CONTENT_AND_CHECKIN_REQUIRED"
        )
        report["RESULT"] = "MANUAL_CHECKIN_REQUIRED"
        report["MESSAGE"] = (
            "The exact managed drawing payload changed after UF Clone but is "
            "still checked out by {0}. If this is your checkout, verify the "
            "drawing and then check it in manually. If another user owns it or "
            "the owner is unclear, stop and escalate. J16 will not call check-in."
        ).format(owner or "<owner unavailable>")
        return True

    if exact_source:
        report["POST_IMPORT_VERIFICATION"] = "VERIFIED_SHA256"
        report["RESULT"] = "IMPORT_VERIFIED"
        report["MESSAGE"] = (
            "Exact Teamcenter drawing replacement and CHECKED_IN state were "
            "verified by post-import associated-file SHA-256."
        )
        return False

    report["POST_IMPORT_VERIFICATION"] = "VERIFIED_MANAGED_TRANSFORM"
    report["RESULT"] = "IMPORT_VERIFIED_MANAGED_TRANSFORM"
    report["MESSAGE"] = (
        "The exact CHECKED_IN managed drawing payload changed after UF Clone, "
        "and no longer matches its prewrite payload. Teamcenter's expected "
        "managed-mode byte transformation is accepted by the production "
        "contract validated in the controlled J16 trials."
    )
    return False


def require_fresh_apply_session(session):
    try:
        loaded = list(session.Parts)
    except Exception as exc:
        raise RuntimeError(
            "Controlled apply could not prove that the NX session has no loaded parts: {0}"
            .format(error_text(exc))
        )
    if loaded:
        identifiers = [journal_identifier(part) or "<unidentified>" for part in loaded]
        raise RuntimeError(
            "Controlled apply requires a fresh NX managed session with no loaded parts; "
            "found {0}: {1}".format(len(loaded), " | ".join(identifiers))
        )


def require_approved_rows(rows, mode, maximum):
    invalid = [row for row in rows if approval_state(row) == "INVALID"]
    if invalid:
        raise RuntimeError(
            "Controlled apply requires every APPROVED value to be YES, NO, "
            "or blank; found {0} invalid row(s).".format(len(invalid))
        )
    approved = [row for row in rows if approval_state(row) == "YES"]
    if mode == "APPLY_ONE_APPROVED" and len(approved) != 1:
        raise RuntimeError(
            "APPLY_ONE_APPROVED requires exactly one APPROVED=YES CSV row; "
            "found {0}.".format(
                len(approved),
            )
        )
    if mode == "APPLY_APPROVED" and not approved:
        raise RuntimeError(
            "APPLY_APPROVED requires at least one APPROVED=YES CSV row."
        )
    if len(approved) > maximum:
        raise RuntimeError(
            "Approved row count {0} exceeds the controlled write limit {1}."
            .format(len(approved), maximum)
        )
    return len(approved)


def abort_batch_before_writes(reports):
    safe_no_write_results = (
        "CLONE_PREFLIGHT_OK",
        "SKIPPED_ALREADY_CURRENT",
        "SKIPPED_UNCHANGED",
        "NOT_APPROVED",
    )
    blockers = [
        report
        for report in reports
        if upper(report.get("APPROVED")) == "YES"
        and report.get("RESULT") not in safe_no_write_results
    ]
    if not blockers:
        return False
    for report in reports:
        if report.get("RESULT") == "CLONE_PREFLIGHT_OK":
            report["RESULT"] = "BATCH_ABORTED_PREWRITE_VALIDATION"
            report["MESSAGE"] = (
                "Another approved row failed controlled preflight. The batch "
                "was aborted before every Teamcenter write."
            )
    return True


def execute(
    session,
    file_management,
    api,
    rows,
    csv_path,
    timestamp,
    mode,
    log,
    work_root,
):
    write_limit = 0
    if is_apply_mode(mode):
        write_limit = (
            1 if mode == "APPLY_ONE_APPROVED" else configured_max_approved_writes()
        )
        require_approved_rows(rows, mode, write_limit)
        require_fresh_apply_session(session)
    reports, proposals = local_preflight(rows, csv_path, timestamp, mode)
    run_managed_preflight(
        session,
        file_management,
        proposals,
        work_root,
        log,
    )
    run_dry_run(api, proposals, log, mode)

    if mode == "DRY_RUN":
        return reports

    if mode == "APPLY_APPROVED" and abort_batch_before_writes(reports):
        log.write(
            "  BATCH ABORTED: at least one approved row failed prewrite "
            "validation; zero Teamcenter writes were attempted."
        )
        return reports

    writes_attempted = 0
    for index, proposal in enumerate(proposals):
        report = proposal["report"]
        if report.get("RESULT") != "CLONE_PREFLIGHT_OK":
            continue

        if writes_attempted >= write_limit:
            set_error(
                report,
                "BLOCKED_PRODUCTION_WRITE_LIMIT",
                "Controlled apply reached its maximum write count {0}.".format(
                    write_limit
                ),
            )
            continue

        try:
            apply_sha = sha256(proposal["drawing"])
        except Exception as exc:
            set_error(
                report,
                "ERROR_FILE_CHANGED_CHECK",
                "Could not reread drawing immediately before import.",
                exc,
            )
            continue

        if apply_sha.lower() != proposal["preflight_sha"].lower():
            set_error(
                report,
                "ERROR_FILE_CHANGED_AFTER_PREFLIGHT",
                "DRAWING_FILE changed after clone preflight. J16 stopped before writing this row.",
            )
            continue

        checkout = inspect_target_checkout(
            session, proposal["identifier"], log
        )
        record_checkout(report, checkout, recheck=True)
        log.write(
            "  CHECKOUT RECHECK {0}: state={1}; owner={2}".format(
                proposal["identifier"],
                checkout.get("state", "UNKNOWN"),
                checkout.get("owner", "") or "<blank>",
            )
        )
        if block_for_checkout(report, checkout, "UF Clone apply"):
            log.write(
                "  BLOCKED {0}: {1}".format(
                    proposal["identifier"], report["MESSAGE"]
                )
            )
            continue

        try:
            _, prewrite_sha = retrieve_exact_associated_drawing(
                session,
                file_management,
                proposal,
                target_evidence_root(work_root, proposal),
                "PREWRITE",
                "PREWRITE_EXPORT_PDI_CODE",
                "PREWRITE_EXPORT_FILE",
            )
            report["PREWRITE_TC_SHA256"] = prewrite_sha
            log.write(
                "  PREWRITE {0}: sha256={1}".format(
                    proposal["identifier"], prewrite_sha
                )
            )
        except Exception as exc:
            set_error(
                report,
                "FAILED_PREWRITE_TARGET_RETRIEVAL",
                (
                    "Could not re-download the exact Teamcenter drawing "
                    "immediately before import."
                ),
                exc,
            )
            log.write(
                "  BLOCKED {0}: {1}".format(
                    proposal["identifier"], report["MESSAGE"]
                )
            )
            continue

        if prewrite_sha.lower() != proposal["tc_baseline_sha"].lower():
            set_error(
                report,
                "BLOCKED_STALE_TARGET",
                (
                    "The exact Teamcenter drawing changed after managed "
                    "preflight. No write was attempted."
                ),
            )
            log.write(
                "  BLOCKED {0}: {1}".format(
                    proposal["identifier"], report["MESSAGE"]
                )
            )
            continue

        logfile = clone_log_path(proposal, mode, "APPLY")
        report["CLONE_LOG"] = logfile
        report["CLONE_APPLY_LOG"] = logfile
        report["WRITE_ATTEMPTED"] = "YES"
        writes_attempted += 1
        try:
            import_one(api, proposal["drawing"], logfile, False, log)

            _, post_sha = retrieve_exact_associated_drawing(
                session,
                file_management,
                proposal,
                target_evidence_root(work_root, proposal),
                "POSTIMPORT",
                "POST_IMPORT_EXPORT_PDI_CODE",
                "POST_IMPORT_EXPORT_FILE",
            )
            report["POST_IMPORT_TC_SHA256"] = post_sha
            post_checkout = inspect_target_checkout(
                session, proposal["identifier"], log
            )
            review_required = classify_post_import(
                proposal,
                prewrite_sha,
                post_sha,
                post_checkout,
                log,
            )
            if report["RESULT"] == "FAILED_IMPORT_UNVERIFIED":
                log.write(
                    "  UNVERIFIED {0}: source_sha256={1}; prewrite_sha256={2}; "
                    "post_sha256={3}".format(
                        proposal["identifier"],
                        proposal["preflight_sha"],
                        prewrite_sha,
                        post_sha,
                    )
                )
                mark_remaining_after_stopped_write(
                    proposals, index + 1, False
                )
                break
            if review_required:
                log.write(
                    "  REVIEW REQUIRED {0}: source_sha256={1}; "
                    "prewrite_sha256={2}; post_sha256={3}; result={4}".format(
                        proposal["identifier"],
                        proposal["preflight_sha"],
                        prewrite_sha,
                        post_sha,
                        report["RESULT"],
                    )
                )
                mark_remaining_after_stopped_write(
                    proposals, index + 1, True
                )
                break
            log.write(
                "  VERIFIED {0}: sha256={1}; checkout=CHECKED_IN".format(
                    proposal["identifier"], post_sha
                )
            )
        except Exception as exc:
            report["POST_IMPORT_VERIFICATION"] = "FAILED"
            set_error(
                report,
                "FAILED_IMPORT_UNVERIFIED",
                (
                    "A write was attempted, but UF Clone or exact-target "
                    "post-import verification failed."
                ),
                exc,
            )
            log.write("  FAILED or unverified apply: {0}".format(error_text(exc)))
            log.write(traceback.format_exc())
            mark_remaining_after_stopped_write(proposals, index + 1, False)
            break

    return reports


def summary_counts(reports):
    return Counter(report.get("RESULT", "") or "<blank>" for report in reports)


def has_failure(reports, mode):
    failure_prefixes = (
        "ERROR_",
        "FAILED_",
        "BLOCKED_",
        "BATCH_ABORTED_",
        "BATCH_STOPPED_",
    )
    for report in reports:
        result = report.get("RESULT", "")
        if result.startswith(failure_prefixes):
            if is_apply_mode(mode):
                if result == "ERROR_APPROVAL_VALUE" or upper(report.get("APPROVED")) == "YES":
                    return True
                continue
            return True
    return False


def has_review_required(reports):
    review_results = (
        "MANUAL_CHECKIN_REQUIRED",
        "REVIEW_NOT_ATTEMPTED_AFTER_PRIOR_WRITE",
    )
    return any(report.get("RESULT", "") in review_results for report in reports)


def main():
    session = NXOpen.Session.GetSession()
    ufs = NXOpen.UF.UFSession.GetUFSession()
    log = Log(session)
    current_mode = configured_mode()
    input_path = configured_input_path()
    timestamp = stamp()
    work_root = os.path.join(
        os.path.dirname(input_path) if input_path else io_root(),
        "J16_EVIDENCE_{0}".format(timestamp),
    )
    evidence_zip = work_root + ".zip"

    log.write("=" * 72)
    log.write("J16 TEAMCENTER X STANDALONE DRAWING IMPORT")
    log.write("Build: {0} | Mode: {1}".format(BUILD, current_mode))
    log.write("Runtime target: NX X 2506 only")
    log.write("Checkout rule: every existing checkout blocks that row")
    log.write("Verification: GetAssociatedFiles + DownloadAssociatedFiles")
    log.write(
        "Persistence rule: post payload must differ from prewrite; managed byte "
        "transformations are accepted only for the exact CHECKED_IN target"
    )
    if current_mode == "APPLY_ONE_APPROVED":
        log.write(
            "PRODUCTION CONTROL: generic target; fresh session; exactly one "
            "approved row; maximum one write"
        )
        log.write(
            "Approval handshake: local and Teamcenter hashes must come from "
            "the accepted J16 DRY_RUN report"
        )
    elif current_mode == "APPLY_APPROVED":
        log.write(
            "BULK PRODUCTION CONTROL: fresh session; all approved rows must "
            "pass preflight before the first write; maximum writes={0}".format(
                env("NX_J16_MAX_APPROVED_WRITES")
                or DEFAULT_MAX_APPROVED_WRITES
            )
        )
        log.write(
            "Approval handshake: every approved row requires local and "
            "Teamcenter hashes from the accepted J16 DRY_RUN report"
        )
    else:
        log.write(
            "DRY_RUN output supplies APPROVED_LOCAL_SHA256 and "
            "APPROVED_TC_BASELINE_SHA256 for controlled apply"
        )
        log.write(
            "Approval step: edit the DRY_RUN report, set exactly one row to "
            "APPROVED=YES, and enter ENGINEER"
        )
    log.write("Input: {0}".format(input_path))
    log.write("Evidence: {0}".format(work_root))
    log.write("Evidence ZIP: {0}".format(evidence_zip))
    log.write("=" * 72)

    report_path = ""
    reports = []
    log_path = ""
    file_management = None
    try:
        if current_mode not in VALID_MODES:
            raise RuntimeError(
                "USER_MODE/NX_J16_MODE must be DRY_RUN, APPLY_ONE_APPROVED, "
                "or APPLY_APPROVED."
            )
        if current_mode == "APPLY_APPROVED" and not BATCH_APPLY_ENABLED:
            raise RuntimeError(
                "APPLY_APPROVED batch mode is disabled. Use DRY_RUN followed "
                "by APPLY_ONE_APPROVED for one approved drawing."
            )
        if not os.path.isfile(input_path):
            raise RuntimeError("Import CSV not found: {0}".format(input_path))

        rows = read_csv(input_path)
        if not rows:
            raise RuntimeError("Import CSV contains no data rows: {0}".format(input_path))

        api = resolve_clone_api(ufs, log)
        _, file_management = new_file_management(session)
        os.makedirs(work_root, exist_ok=True)
        reports = execute(
            session,
            file_management,
            api,
            rows,
            input_path,
            timestamp,
            current_mode,
            log,
            work_root,
        )

        report_path = os.path.join(
            os.path.dirname(input_path),
            "J16_{0}_{1}.csv".format(current_mode, timestamp),
        )
        write_csv(report_path, reports)

        log.write("Report: {0}".format(report_path))
        for result, count in sorted(summary_counts(reports).items()):
            log.write("  {0}: {1}".format(result, count))

        if has_failure(reports, current_mode):
            log.write("FINAL STATUS: FAILED")
            raise RuntimeError(
                "J16 completed with one or more failed safety/import rows. "
                "Review: {0}".format(report_path)
            )

        if has_review_required(reports):
            log.write("FINAL STATUS: REVIEW_REQUIRED")
            log.write(
                "Do not rerun the import. Manually verify drawing content, "
                "checkout state, unchanged 3D master, and managed object count."
            )
        else:
            log.write("FINAL STATUS: SUCCESS")

    except Exception as exc:
        if "FINAL STATUS: FAILED" not in log.lines:
            log.write("FINAL STATUS: FAILED")
        log.write(error_text(exc))
        log.write(traceback.format_exc())
        raise

    finally:
        dispose(file_management)
        try:
            log_dir = os.path.dirname(input_path) if input_path else io_root()
            if not log_dir:
                log_dir = io_root()
            os.makedirs(log_dir, exist_ok=True)
            log_path = os.path.join(
                log_dir, "J16_RUN_{0}_{1}.log".format(current_mode, timestamp)
            )
            write_log(log_path, log.lines)
            clone_logs = []
            for report in reports:
                clone_logs.extend(
                    [
                        report.get("CLONE_PREFLIGHT_LOG", ""),
                        report.get("CLONE_APPLY_LOG", ""),
                    ]
                )
            zip_artifacts(
                evidence_zip,
                work_root,
                [report_path, log_path] + clone_logs,
            )
        except Exception as exc:
            log.write(
                "WARNING: automatic evidence ZIP failed: {0}".format(
                    error_text(exc)
                )
            )
            try:
                if log_path:
                    write_log(log_path, log.lines)
            except Exception:
                pass

    return report_path


if __name__ == "__main__":
    main()
