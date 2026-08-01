"""J17 - create a new drawing specification beneath an existing 3D master.

NX X 2506 managed mode only.

J17 imports a locally created NX drawing as a NEW /specification/ dataset under
an existing 3D Item/Revision. J16 owns replacement of an existing drawing
specification; J17 refuses to overwrite one.

Production contract:
- APPROVED=YES plus ENGINEER authorizes one automatic preflight-and-apply run.
- The exact destination /specification/ must not be openable before the write.
- The parent 3D Item/Revision must exist, be CHECKED_IN, be discovered by UF
  Clone, and remain UseExisting.
- Every discovered object defaults to UseExisting. Only the staged local
  encoded specification drawing receives the UF Clone Overwrite action used by
  AutoTranslate to create its new managed dataset.
- Immediately before apply, J17 repeats destination, source-file, and 3D
  payload checks.
- After apply, the exact new drawing specification must open, contain sheets,
  expose one native .prt, and the preserved 3D native SHA-256 must be unchanged.
- J17 never checks in, saves, revises, deletes, or force-unlocks an object.

DRY_RUN remains available as a diagnostic mode. APPLY_APPROVED performs the
read-only managed checks and UF Clone dry run internally; a separate operator
dry-run pass is not required.
"""

import csv
import importlib.util
import os
import re
import shutil
import traceback
from collections import Counter

import NXOpen
import NXOpen.UF


# ============================================================================
# USER SETTINGS
# ============================================================================
USER_IMPORT_CSV = r""  # blank => <I/O root>\NX_TC_NEW_DRAWING_SPECIFICATION.csv
USER_MODE = "APPLY_APPROVED"  # production default; DRY_RUN remains diagnostic
# Optional environment overrides:
#   NX_TC_NEW_DRAWING_SPECIFICATION_FILE=<full CSV path>
#   NX_J17_MODE=DRY_RUN or APPLY_APPROVED
#   NX_J17_MAX_APPROVED_WRITES=1..100 (default 25)
# ============================================================================

BUILD = "J17-TCX-NEW-DRAWING-SPECIFICATION-NX2506-V3"
DEFAULT_INPUT = "NX_TC_NEW_DRAWING_SPECIFICATION.csv"
VALID_MODES = ("DRY_RUN", "APPLY_APPROVED")
DEFAULT_MAX_APPROVED_WRITES = 25

REQUIRED_COLUMNS = (
    "PART_NUMBER",
    "REVISION",
    "DWG_INDEX",
    "SOURCE_DRAWING_FILE",
    "APPROVED",
    "ENGINEER",
)

REPORT_COLUMNS = (
    "RUN_TIMESTAMP",
    "MODE",
    "CSV_ROW",
    "PART_NUMBER",
    "REVISION",
    "DWG_INDEX",
    "DRAWING_IDENTIFIER",
    "SOURCE_DRAWING_FILE",
    "SOURCE_SHA256",
    "STAGED_SPECIFICATION_FILE",
    "STAGED_SHA256",
    "PRESERVE_3D_PART_NUMBER",
    "PRESERVE_3D_REVISION",
    "PRESERVE_3D_IDENTIFIER",
    "PRESERVE_3D_OPENED_IDENTIFIER",
    "PRESERVE_3D_CHECKOUT_STATE",
    "PRESERVE_3D_CHECKOUT_OWNER",
    "PRESERVE_3D_CHECKOUT_RAW",
    "PRESERVE_3D_ASSOCIATED_FILES",
    "PRESERVE_3D_NATIVE_FILE",
    "PRESERVE_3D_BASELINE_SHA256",
    "PRESERVE_3D_RECHECK_STATE",
    "PRESERVE_3D_RECHECK_OWNER",
    "PRESERVE_3D_PREWRITE_SHA256",
    "PRESERVE_3D_POST_SHA256",
    "PRESERVE_3D_UNCHANGED",
    "TARGET_INITIAL_STATE",
    "TARGET_INITIAL_DETAIL",
    "TARGET_RECHECK_STATE",
    "TARGET_RECHECK_DETAIL",
    "DEFAULT_IMPORT_ACTION",
    "DRAWING_SPECIFICATION_ACTION",
    "PRESERVE_3D_ACTION",
    "PRESERVE_3D_DISCOVERED",
    "PRESERVE_3D_DISCOVERED_NAME",
    "DISCOVERED_PARTS",
    "NAMING_FAILURE_EVIDENCE",
    "CLONE_PREFLIGHT",
    "CLONE_PREFLIGHT_LOG",
    "CLONE_APPLY_LOG",
    "POST_IMPORT_OPENED_IDENTIFIER",
    "POST_IMPORT_CHECKOUT_STATE",
    "POST_IMPORT_CHECKOUT_OWNER",
    "POST_IMPORT_CHECKOUT_RAW",
    "POST_IMPORT_DRAWING_SHEET_COUNT",
    "POST_IMPORT_ASSOCIATED_FILES",
    "POST_IMPORT_NATIVE_FILE",
    "POST_IMPORT_TC_SHA256",
    "POST_IMPORT_VERIFICATION",
    "WRITE_ATTEMPTED",
    "DISPOSITION",
    "QUARANTINE_REASON",
    "APPROVED",
    "ENGINEER",
    "RESULT",
    "MESSAGE",
)


def load_j16():
    """Load only the current, tested J16 utility boundary used by J17."""
    path = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "16_tc_offline_drawing_import.py",
    )
    if not os.path.isfile(path):
        raise RuntimeError("J16 dependency not found beside J17: {0}".format(path))
    spec = importlib.util.spec_from_file_location("nx_journal_16_for_j17", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    required = (
        "Log",
        "ProcessStateError",
        "clean",
        "upper",
        "env",
        "stamp",
        "io_root",
        "error_text",
        "dispose",
        "sha256",
        "write_log",
        "zip_artifacts",
        "safe_folder_name",
        "resolve_clone_api",
        "terminate",
        "add_assembly",
        "naming_failures",
        "perform_clone",
        "set_action",
        "iterate_parts",
        "same_part",
        "require_fresh_apply_session",
        "new_file_management",
        "unwrap_open_result",
        "find_loaded_by_identifier",
        "journal_identifier",
        "query_pdm_checkout",
        "close_opened_part",
        "collect_pdm_files",
        "pdm_file_name",
        "locate_downloaded_files",
        "release_pdm_files",
        "managed_native_name",
    )
    missing = [name for name in required if not hasattr(module, name)]
    if missing:
        raise RuntimeError(
            "J16 is incompatible with J17; missing: {0}".format(
                ", ".join(missing)
            )
        )
    return module


J16 = load_j16()
clean = J16.clean
upper = J16.upper


def configured_mode():
    return upper(J16.env("NX_J17_MODE") or USER_MODE or "DRY_RUN")


def configured_input_path():
    value = J16.env("NX_TC_NEW_DRAWING_SPECIFICATION_FILE") or clean(
        USER_IMPORT_CSV
    )
    if value:
        return os.path.abspath(os.path.expanduser(value))
    return os.path.join(J16.io_root(), DEFAULT_INPUT)


def configured_max_approved_writes():
    raw = J16.env("NX_J17_MAX_APPROVED_WRITES")
    try:
        value = int(raw) if raw else DEFAULT_MAX_APPROVED_WRITES
    except Exception:
        raise RuntimeError("NX_J17_MAX_APPROVED_WRITES must be an integer.")
    if value < 1 or value > 100:
        raise RuntimeError(
            "NX_J17_MAX_APPROVED_WRITES must be between 1 and 100."
        )
    return value


def master_id(part_number, revision):
    return "@DB/{0}/{1}".format(part_number, revision)


def dataset_name(part_number, revision, drawing_index):
    return "{0}-{1}-dwg{2}".format(part_number, revision, drawing_index)


def drawing_id(part_number, revision, drawing_index):
    return "@DB/{0}/{1}/specification/{2}".format(
        part_number,
        revision,
        dataset_name(part_number, revision, drawing_index),
    )


def expected_specification_import_native(part_number, revision, drawing_index):
    """Proven UF Clone AutoTranslate encoding for a specification dataset."""
    return "{0}_{1}_s_{2}.prt".format(
        part_number,
        revision,
        dataset_name(part_number, revision, drawing_index),
    )


def resolve_local_path(csv_path, value):
    path = os.path.expanduser(clean(value))
    if not path:
        return ""
    if not os.path.isabs(path):
        path = os.path.join(os.path.dirname(csv_path), path)
    return os.path.abspath(path)


def read_csv(path):
    last_error = None
    for encoding in ("utf-8-sig", "utf-8", "cp1252"):
        try:
            with open(path, "r", encoding=encoding, newline="") as handle:
                reader = csv.DictReader(handle)
                headers = [clean(name) for name in (reader.fieldnames or [])]
                missing = [name for name in REQUIRED_COLUMNS if name not in headers]
                if missing:
                    raise RuntimeError(
                        "Missing CSV column(s): {0}".format(", ".join(missing))
                    )
                rows = []
                for number, source in enumerate(reader, 2):
                    row = {
                        clean(key): clean(value)
                        for key, value in source.items()
                        if key is not None
                    }
                    row["_CSV_ROW"] = number
                    rows.append(row)
                return rows
        except UnicodeDecodeError as exc:
            last_error = exc
    raise RuntimeError("Unable to decode CSV {0}: {1}".format(path, last_error))


def write_csv(path, rows):
    with open(path, "w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=REPORT_COLUMNS)
        writer.writeheader()
        for row in rows:
            writer.writerow({key: row.get(key, "") for key in REPORT_COLUMNS})


def approval_state(row):
    value = upper(row.get("APPROVED"))
    return value if value in ("YES", "NO") else ("NO" if not value else "INVALID")


def parse_row(row):
    part_number = clean(row.get("PART_NUMBER"))
    revision = clean(row.get("REVISION"))
    raw_index = clean(row.get("DWG_INDEX"))
    if not part_number or not revision or not raw_index:
        raise RuntimeError("PART_NUMBER, REVISION, and DWG_INDEX are required.")
    try:
        drawing_index = int(raw_index)
    except Exception:
        raise RuntimeError("DWG_INDEX must be an integer from 1 through 99.")
    if drawing_index < 1 or drawing_index > 99:
        raise RuntimeError("DWG_INDEX must be an integer from 1 through 99.")
    return part_number, revision, drawing_index


def base_report(row, timestamp, mode):
    report = {name: "" for name in REPORT_COLUMNS}
    report.update(
        RUN_TIMESTAMP=timestamp,
        MODE=mode,
        CSV_ROW=row.get("_CSV_ROW", ""),
        PART_NUMBER=row.get("PART_NUMBER", ""),
        REVISION=row.get("REVISION", ""),
        DWG_INDEX=row.get("DWG_INDEX", ""),
        SOURCE_DRAWING_FILE=row.get("SOURCE_DRAWING_FILE", ""),
        DEFAULT_IMPORT_ACTION="UseExisting",
        DRAWING_SPECIFICATION_ACTION="Overwrite",
        PRESERVE_3D_ACTION="UseExisting",
        PRESERVE_3D_DISCOVERED="NO",
        CLONE_PREFLIGHT="NOT_RUN",
        POST_IMPORT_VERIFICATION="NOT_RUN",
        WRITE_ATTEMPTED="NO",
        DISPOSITION="PENDING",
        APPROVED=row.get("APPROVED", ""),
        ENGINEER=row.get("ENGINEER", ""),
    )
    return report


def set_result(report, result, message, disposition="", error=None):
    report["RESULT"] = result
    report["MESSAGE"] = message
    if disposition:
        report["DISPOSITION"] = disposition
    if error is not None:
        detail = J16.error_text(error)
        if detail not in report["MESSAGE"]:
            report["MESSAGE"] += " | " + detail


def quarantine(report, result, message, error=None):
    set_result(report, result, message, "QUARANTINED", error)
    report["WRITE_ATTEMPTED"] = "NO"
    report["QUARANTINE_REASON"] = report["MESSAGE"]


def stage_copy(source, root, part_number, revision, drawing_index):
    folder = os.path.join(
        root,
        J16.safe_folder_name(
            "{0}_{1}_DWG{2}".format(part_number, revision, drawing_index)
        ),
    )
    os.makedirs(folder, exist_ok=True)
    target = os.path.join(
        folder,
        expected_specification_import_native(
            part_number, revision, drawing_index
        ),
    )
    shutil.copy2(source, target)
    return target


def duplicate_target_keys(rows):
    keys = []
    for row in rows:
        try:
            part_number, revision, drawing_index = parse_row(row)
            keys.append((upper(part_number), upper(revision), drawing_index))
        except Exception:
            pass
    counts = Counter(keys)
    return {key for key, count in counts.items() if count > 1}


def local_preflight(rows, csv_path, stage_root, timestamp, mode):
    reports = []
    proposals = []
    duplicates = duplicate_target_keys(rows)

    for row in rows:
        report = base_report(row, timestamp, mode)
        reports.append(report)
        approval = approval_state(row)
        if approval == "INVALID":
            quarantine(
                report,
                "ERROR_APPROVAL_VALUE",
                "APPROVED must be YES, NO, or blank.",
            )
            continue
        if mode == "APPLY_APPROVED" and approval != "YES":
            set_result(
                report,
                "NOT_APPROVED",
                "No Teamcenter creation was authorized for this row.",
                "NOT_APPROVED",
            )
            continue
        if mode == "APPLY_APPROVED" and not clean(row.get("ENGINEER")):
            quarantine(
                report,
                "ERROR_ENGINEER_REQUIRED",
                "ENGINEER is required when APPROVED=YES.",
            )
            continue

        try:
            part_number, revision, drawing_index = parse_row(row)
            if (upper(part_number), upper(revision), drawing_index) in duplicates:
                raise RuntimeError(
                    "Duplicate PART_NUMBER/REVISION/DWG_INDEX destination."
                )
            source = resolve_local_path(csv_path, row.get("SOURCE_DRAWING_FILE"))
            if not source or not os.path.isfile(source):
                raise RuntimeError(
                    "SOURCE_DRAWING_FILE not found: {0}".format(source or "<blank>")
                )
            if not source.lower().endswith(".prt"):
                raise RuntimeError("SOURCE_DRAWING_FILE must be an NX .prt file.")

            identifier = drawing_id(part_number, revision, drawing_index)
            supplied_identifier = clean(row.get("DRAWING_IDENTIFIER"))
            if supplied_identifier and upper(supplied_identifier).replace(
                "\\", "/"
            ) != upper(identifier).replace("\\", "/"):
                raise RuntimeError(
                    "Optional DRAWING_IDENTIFIER does not match the requested specification."
                )

            staged = stage_copy(
                source,
                stage_root,
                part_number,
                revision,
                drawing_index,
            )
            source_sha = J16.sha256(source)
            staged_sha = J16.sha256(staged)
            if source_sha.lower() != staged_sha.lower():
                raise RuntimeError("Staged drawing is not byte-identical to source.")

            model_identifier = master_id(part_number, revision)
            report.update(
                PART_NUMBER=part_number,
                REVISION=revision,
                DWG_INDEX=drawing_index,
                DRAWING_IDENTIFIER=identifier,
                SOURCE_DRAWING_FILE=source,
                SOURCE_SHA256=source_sha,
                STAGED_SPECIFICATION_FILE=staged,
                STAGED_SHA256=staged_sha,
                PRESERVE_3D_PART_NUMBER=part_number,
                PRESERVE_3D_REVISION=revision,
                PRESERVE_3D_IDENTIFIER=model_identifier,
                RESULT="LOCAL_PREFLIGHT_OK",
                MESSAGE=(
                    "Local drawing was staged with the proven encoded name for "
                    "a new specification beneath the existing 3D revision."
                ),
            )
            proposals.append(
                {
                    "row": row,
                    "report": report,
                    "part_number": part_number,
                    "revision": revision,
                    "drawing_index": drawing_index,
                    "identifier": identifier,
                    "model_part_number": part_number,
                    "model_revision": revision,
                    "model_identifier": model_identifier,
                    "source": source,
                    "source_sha": source_sha,
                    "staged": staged,
                    "staged_sha": staged_sha,
                }
            )
        except Exception as exc:
            quarantine(
                report,
                "ERROR_LOCAL_PREFLIGHT",
                "Local master-drawing safety preflight failed.",
                exc,
            )
    return reports, proposals


def collection_count(value):
    try:
        return int(value.Count)
    except Exception:
        try:
            return len(list(value))
        except Exception:
            return -1


def inspect_exact_part(session, identifier, log):
    """Read-only exact identity, checkout, and drawing-sheet inspection."""
    part = J16.find_loaded_by_identifier(session, identifier)
    load_status = None
    opened_here = False
    result = {
        "state": "NOT_OPENABLE",
        "opened_identifier": "",
        "checkout_state": "UNKNOWN",
        "checkout_owner": "",
        "checkout_raw": "",
        "drawing_sheet_count": -1,
        "detail": "",
        "error_code": "",
    }
    try:
        if part is None:
            part, load_status = J16.unwrap_open_result(
                session.Parts.OpenBase(identifier)
            )
            opened_here = True
        if part is None:
            result["detail"] = "OpenBase returned no part."
            return result

        actual = J16.journal_identifier(part)
        result["opened_identifier"] = actual
        if upper(actual).replace("\\", "/") != upper(identifier).replace("\\", "/"):
            result["state"] = "IDENTITY_MISMATCH"
            result["detail"] = (
                "Opened JournalIdentifier does not match exact identifier: {0}"
                .format(actual or "<blank>")
            )
            return result

        checkout = J16.query_pdm_checkout(part)
        result.update(
            state="EXISTS",
            checkout_state=checkout.get("state", "UNKNOWN"),
            checkout_owner=checkout.get("owner", ""),
            checkout_raw=checkout.get("raw", ""),
            drawing_sheet_count=collection_count(getattr(part, "DrawingSheets", [])),
            detail="Exact managed Item/Revision opened successfully.",
        )
        return result
    except Exception as exc:
        result["detail"] = J16.error_text(exc)
        result["error_code"] = clean(getattr(exc, "ErrorCode", ""))
        return result
    finally:
        J16.dispose(load_status)
        if opened_here and part is not None:
            J16.close_opened_part(part, log)


def record_model_inspection(report, inspection, recheck=False):
    if recheck:
        report["PRESERVE_3D_RECHECK_STATE"] = inspection.get(
            "checkout_state", "UNKNOWN"
        )
        report["PRESERVE_3D_RECHECK_OWNER"] = inspection.get(
            "checkout_owner", ""
        )
        return
    report["PRESERVE_3D_OPENED_IDENTIFIER"] = inspection.get(
        "opened_identifier", ""
    )
    report["PRESERVE_3D_CHECKOUT_STATE"] = inspection.get(
        "checkout_state", "UNKNOWN"
    )
    report["PRESERVE_3D_CHECKOUT_OWNER"] = inspection.get(
        "checkout_owner", ""
    )
    report["PRESERVE_3D_CHECKOUT_RAW"] = inspection.get("checkout_raw", "")


def require_clear_existing_model(inspection):
    if inspection.get("state") != "EXISTS":
        raise RuntimeError(
            "Preserved 3D Item/Revision could not be opened exactly: {0}".format(
                inspection.get("detail", "") or "<no detail>"
            )
        )
    if inspection.get("checkout_state") != "CHECKED_IN":
        raise RuntimeError(
            "Preserved 3D Item/Revision is not proven CHECKED_IN; state={0}, "
            "owner={1}, raw={2}".format(
                inspection.get("checkout_state", "UNKNOWN"),
                inspection.get("checkout_owner", "") or "<blank>",
                inspection.get("checkout_raw", "") or "<blank>",
            )
        )


def target_absence_state(inspection):
    """Existing exact targets always block; non-openable targets may be created."""
    if inspection.get("state") == "EXISTS":
        return "EXISTS"
    if inspection.get("state") == "IDENTITY_MISMATCH":
        return "UNKNOWN"
    return "NOT_OPENABLE"


def record_target_inspection(report, inspection, recheck=False):
    state_field = "TARGET_RECHECK_STATE" if recheck else "TARGET_INITIAL_STATE"
    detail_field = "TARGET_RECHECK_DETAIL" if recheck else "TARGET_INITIAL_DETAIL"
    report[state_field] = target_absence_state(inspection)
    report[detail_field] = inspection.get("detail", "")


def open_exact_for_retrieval(session, identifier, log):
    part = J16.find_loaded_by_identifier(session, identifier)
    load_status = None
    opened_here = False
    if part is None:
        part, load_status = J16.unwrap_open_result(session.Parts.OpenBase(identifier))
        opened_here = True
    if part is None:
        J16.dispose(load_status)
        raise RuntimeError("OpenBase returned no part for {0}.".format(identifier))
    actual = J16.journal_identifier(part)
    if upper(actual).replace("\\", "/") != upper(identifier).replace("\\", "/"):
        J16.dispose(load_status)
        if opened_here:
            J16.close_opened_part(part, log)
        raise RuntimeError(
            "Opened JournalIdentifier does not match {0}: {1}".format(
                identifier, actual or "<blank>"
            )
        )
    return part, load_status, opened_here


def retrieve_single_native(
    session,
    file_management,
    identifier,
    evidence_root,
    log,
    expected_basename="",
):
    """Download exactly one native .prt attached to an exact Item/Revision."""
    os.makedirs(evidence_root, exist_ok=True)
    part = None
    load_status = None
    opened_here = False
    pdm_files = []
    original_cwd = os.getcwd()
    try:
        part, load_status, opened_here = open_exact_for_retrieval(
            session, identifier, log
        )
        get_method = getattr(file_management, "GetAssociatedFiles", None)
        download_method = getattr(file_management, "DownloadAssociatedFiles", None)
        if not callable(get_method) or not callable(download_method):
            raise RuntimeError(
                "PDM GetAssociatedFiles/DownloadAssociatedFiles is unavailable."
            )

        raw_files = get_method([part], [])
        pdm_files = J16.collect_pdm_files(raw_files)
        names = [J16.pdm_file_name(value) for value in pdm_files]
        native_pairs = [
            (value, name)
            for value, name in zip(pdm_files, names)
            if os.path.basename(name).lower().endswith(".prt")
        ]
        if expected_basename:
            native_pairs = [
                (value, name)
                for value, name in native_pairs
                if os.path.basename(name).lower()
                == os.path.basename(expected_basename).lower()
            ]
        if len(native_pairs) != 1:
            raise RuntimeError(
                "Expected exactly one attached native {0} on {1}; found {2}. "
                "Associated files: {3}".format(
                    expected_basename or ".prt",
                    identifier,
                    len(native_pairs),
                    " | ".join(name or "<unreadable>" for name in names)
                    or "<none>",
                )
            )

        download_result = download_method([part], pdm_files)
        returned_files = J16.collect_pdm_files(download_result)
        for value in returned_files:
            if all(value is not existing for existing in pdm_files):
                pdm_files.append(value)
        download_cwd = os.getcwd()
        native_name = native_pairs[0][1]
        candidates = [native_name, os.path.basename(native_name)]
        candidates.extend(
            J16.pdm_file_name(value)
            for value in returned_files
            if os.path.basename(J16.pdm_file_name(value)).lower()
            == os.path.basename(native_name).lower()
        )
        physical = J16.locate_downloaded_files(candidates, download_cwd)
        if len(physical) != 1:
            raise RuntimeError(
                "DownloadAssociatedFiles did not materialize one unambiguous "
                "native {0}; found {1} physical matches in {2}.".format(
                    os.path.basename(native_name), len(physical), download_cwd
                )
            )

        downloaded = next(iter(physical.values()))
        evidence_name = J16.safe_folder_name(os.path.basename(native_name))
        evidence_file = os.path.join(evidence_root, evidence_name)
        shutil.copy2(downloaded, evidence_file)
        return {
            "associated_files": " | ".join(
                name or "<unreadable>" for name in names
            ),
            "native_name": native_name,
            "evidence_file": evidence_file,
            "sha256": J16.sha256(evidence_file),
            "download_cwd": download_cwd,
        }
    finally:
        try:
            os.chdir(original_cwd)
        except Exception as exc:
            raise J16.ProcessStateError(
                "Could not restore process working directory after associated-file "
                "retrieval: {0}".format(J16.error_text(exc))
            )
        finally:
            J16.release_pdm_files(pdm_files)
            J16.dispose(load_status)
            if opened_here and part is not None:
                J16.close_opened_part(part, log)


def normalized(value):
    return clean(value).lower().replace("\\", "/")


def model_tokens(part_number, revision):
    stem = "{0}_{1}".format(part_number, revision).lower()
    return (
        master_id(part_number, revision).lower(),
        stem,
        stem + ".prt",
    )


def find_model_references(parts, staged, part_number, revision):
    canonical, stem, native = model_tokens(part_number, revision)
    matches = []
    for part in parts:
        if J16.same_part(part, staged):
            continue
        value = normalized(part).rstrip("/")
        leaf = os.path.basename(value)
        if value == canonical or leaf in (stem, native):
            matches.append(part)
    return matches


def naming_failure_evidence(value):
    """Capture NX binding output without assuming one runtime result shape."""
    details = [
        "type={0}".format(type(value).__name__),
        "raw={0}".format(repr(value)),
    ]
    try:
        names = [
            name
            for name in dir(value)
            if not name.startswith("_")
            and any(
                token in name.lower()
                for token in ("error", "fail", "name", "part", "status")
            )
        ]
    except Exception:
        names = []
    for name in names:
        try:
            member = getattr(value, name)
            if not callable(member):
                details.append("{0}={1}".format(name, repr(member)))
        except Exception:
            pass
    return " | ".join(details)[:4000]


def clone_log_path(proposal, mode, phase):
    return os.path.join(
        os.path.dirname(proposal["staged"]),
        "J17_{0}_{1}_{2}_{3}_DWG{4}.clone".format(
            phase,
            mode,
            proposal["part_number"],
            proposal["revision"],
            proposal["drawing_index"],
        ),
    )


def import_one(api, proposal, logfile, dry_run, log):
    """UF Clone import: references UseExisting, exact staged master Overwrite."""
    clone = api["clone"]
    load_status = None
    discovered_parts = []
    try:
        J16.terminate(clone)
        clone.Initialise(api["import_operation"])
        clone.SetFamilyTreatment(api["treat_as_lost"])
        clone.SetDefNaming(api["autotranslate"])
        clone.SetDefItemType("")
        clone.SetDefDirectory(os.path.dirname(proposal["staged"]))
        try:
            clone.SetAssocFileRootDir(os.path.dirname(proposal["staged"]))
        except Exception:
            pass

        clone.SetDefAction(api["use_existing"])
        clone.SetDefAssocFileCopy(True)
        clone.SetLogfile(logfile)
        try:
            clone.SetPropagateActions(False)
        except Exception:
            pass

        load_status = J16.add_assembly(clone, proposal["staged"])
        discovered_parts = J16.iterate_parts(clone)
        if not discovered_parts:
            raise RuntimeError("UF Clone discovered no parts.")

        target_action_set = False
        for part_name in discovered_parts:
            if J16.same_part(part_name, proposal["staged"]):
                J16.set_action(clone, part_name, api["overwrite"])
                target_action_set = True
            else:
                try:
                    J16.set_action(clone, part_name, api["use_existing"])
                except Exception:
                    # The already-established default remains UseExisting.
                    pass
        if not target_action_set:
            J16.set_action(clone, proposal["staged"], api["overwrite"])

        model_matches = find_model_references(
            discovered_parts,
            proposal["staged"],
            proposal["model_part_number"],
            proposal["model_revision"],
        )
        if not model_matches:
            raise RuntimeError(
                "Required preserved 3D reference was not discovered by UF Clone: {0}"
                .format(proposal["model_identifier"])
            )
        for model_name in model_matches:
            try:
                J16.set_action(clone, model_name, api["use_existing"])
            except Exception:
                pass

        report = proposal["report"]
        report["PRESERVE_3D_DISCOVERED"] = "YES"
        report["PRESERVE_3D_DISCOVERED_NAME"] = " | ".join(model_matches)
        report["DISCOVERED_PARTS"] = " | ".join(discovered_parts)

        failures = J16.naming_failures(clone)
        clone.SetDryrun(bool(dry_run))
        try:
            clone.GenerateReport()
        except Exception:
            pass
        try:
            J16.perform_clone(clone, failures)
        except Exception:
            proposal["report"]["NAMING_FAILURE_EVIDENCE"] = (
                naming_failure_evidence(failures)
            )
            raise
        log.write(
            "  UF Clone completed: discovered={0}; default=UseExisting; "
            "new drawing specification=Overwrite; 3D master=UseExisting; dry_run={1}"
            .format(len(discovered_parts), dry_run)
        )
        return discovered_parts
    finally:
        J16.dispose(load_status)
        J16.terminate(clone)


def evidence_path(work_root, proposal, phase):
    return os.path.join(
        work_root,
        J16.safe_folder_name(
            "{0}_{1}_DWG{2}".format(
                proposal["part_number"],
                proposal["revision"],
                proposal["drawing_index"],
            )
        ),
        phase,
    )


def managed_preflight(session, file_management, api, proposals, mode, work_root, log):
    for proposal in proposals:
        report = proposal["report"]
        if report.get("RESULT") != "LOCAL_PREFLIGHT_OK":
            continue
        try:
            model = inspect_exact_part(session, proposal["model_identifier"], log)
            record_model_inspection(report, model)
            require_clear_existing_model(model)
            log.write(
                "  PRESERVE 3D {0}: state={1}; owner={2}; opened={3}".format(
                    proposal["model_identifier"],
                    model.get("checkout_state", "UNKNOWN"),
                    model.get("checkout_owner", "") or "<blank>",
                    model.get("opened_identifier", "") or "<blank>",
                )
            )

            target = inspect_exact_part(session, proposal["identifier"], log)
            record_target_inspection(report, target)
            if target_absence_state(target) != "NOT_OPENABLE":
                quarantine(
                    report,
                    "QUARANTINED_TARGET_ALREADY_EXISTS",
                    "The exact destination is already present or its identity is "
                    "ambiguous. J17 creates only missing specifications and will not "
                    "overwrite it. Detail: {0}".format(target.get("detail", "")),
                )
                continue

            baseline = retrieve_single_native(
                session,
                file_management,
                proposal["model_identifier"],
                evidence_path(work_root, proposal, "PRESERVE_3D_BASELINE"),
                log,
            )
            proposal["model_baseline_sha"] = baseline["sha256"]
            report.update(
                PRESERVE_3D_ASSOCIATED_FILES=baseline["associated_files"],
                PRESERVE_3D_NATIVE_FILE=baseline["native_name"],
                PRESERVE_3D_BASELINE_SHA256=baseline["sha256"],
            )

            preflight_log = clone_log_path(proposal, mode, "PREFLIGHT")
            report["CLONE_PREFLIGHT_LOG"] = preflight_log
            import_one(api, proposal, preflight_log, True, log)
            report["CLONE_PREFLIGHT"] = "PASS"
            report["RESULT"] = (
                "DRY_RUN_OK" if mode == "DRY_RUN" else "CLONE_PREFLIGHT_OK"
            )
            report["DISPOSITION"] = "PREFLIGHT_CLEAR"
            report["MESSAGE"] = (
                "Destination was not openable, exact preserved 3D was CHECKED_IN, "
                "its native payload was fingerprinted, and UF Clone dry run passed."
            )
        except J16.ProcessStateError:
            raise
        except Exception as exc:
            report["CLONE_PREFLIGHT"] = "FAIL"
            quarantine(
                report,
                "QUARANTINED_PREFLIGHT",
                "Managed identity, preserved-3D, or UF Clone preflight failed.",
                exc,
            )
            log.write("  QUARANTINED: {0}".format(report["MESSAGE"]))


def mark_later_stopped(proposals, start, review_required=False):
    for proposal in proposals[start:]:
        report = proposal["report"]
        if report.get("RESULT") != "CLONE_PREFLIGHT_OK":
            continue
        if review_required:
            set_result(
                report,
                "REVIEW_NOT_ATTEMPTED_AFTER_PRIOR_WRITE",
                "A prior new drawing specification requires manual check-in/review; no "
                "write was attempted for this row.",
                "STOPPED",
            )
        else:
            set_result(
                report,
                "BATCH_STOPPED_AFTER_UNVERIFIED_WRITE",
                "A prior creation could not be verified; no write was attempted "
                "for this row.",
                "STOPPED",
            )


def prewrite_checks(session, file_management, proposal, work_root, log):
    report = proposal["report"]
    if J16.sha256(proposal["source"]).lower() != proposal["source_sha"].lower():
        raise RuntimeError("SOURCE_DRAWING_FILE changed after local preflight.")
    if J16.sha256(proposal["staged"]).lower() != proposal["staged_sha"].lower():
        raise RuntimeError("Staged specification drawing changed after local preflight.")

    model = inspect_exact_part(session, proposal["model_identifier"], log)
    record_model_inspection(report, model, recheck=True)
    require_clear_existing_model(model)

    target = inspect_exact_part(session, proposal["identifier"], log)
    record_target_inspection(report, target, recheck=True)
    if target_absence_state(target) != "NOT_OPENABLE":
        raise RuntimeError(
            "Destination became present or ambiguous after preflight. J17 will "
            "not overwrite it. Detail: {0}".format(target.get("detail", ""))
        )

    model_now = retrieve_single_native(
        session,
        file_management,
        proposal["model_identifier"],
        evidence_path(work_root, proposal, "PRESERVE_3D_PREWRITE"),
        log,
    )
    report["PRESERVE_3D_PREWRITE_SHA256"] = model_now["sha256"]
    if model_now["sha256"].lower() != proposal["model_baseline_sha"].lower():
        raise RuntimeError(
            "Preserved 3D native payload changed after managed preflight."
        )


def verify_after_import(session, file_management, proposal, work_root, log):
    report = proposal["report"]
    target = inspect_exact_part(session, proposal["identifier"], log)
    report.update(
        POST_IMPORT_OPENED_IDENTIFIER=target.get("opened_identifier", ""),
        POST_IMPORT_CHECKOUT_STATE=target.get("checkout_state", "UNKNOWN"),
        POST_IMPORT_CHECKOUT_OWNER=target.get("checkout_owner", ""),
        POST_IMPORT_CHECKOUT_RAW=target.get("checkout_raw", ""),
        POST_IMPORT_DRAWING_SHEET_COUNT=target.get("drawing_sheet_count", -1),
    )
    if target.get("state") != "EXISTS":
        raise RuntimeError(
            "The exact new drawing specification could not be opened after "
            "UF Clone: {0}".format(target.get("detail", ""))
        )
    if target.get("drawing_sheet_count", -1) < 1:
        raise RuntimeError(
            "The exact new managed master opened but no drawing sheets were proven."
        )

    imported = retrieve_single_native(
        session,
        file_management,
        proposal["identifier"],
        evidence_path(work_root, proposal, "NEW_DRAWING_SPECIFICATION_POSTIMPORT"),
        log,
        J16.managed_native_name(
            proposal["part_number"],
            proposal["revision"],
            proposal["drawing_index"],
        ),
    )
    report.update(
        POST_IMPORT_ASSOCIATED_FILES=imported["associated_files"],
        POST_IMPORT_NATIVE_FILE=imported["native_name"],
        POST_IMPORT_TC_SHA256=imported["sha256"],
    )

    model = inspect_exact_part(session, proposal["model_identifier"], log)
    require_clear_existing_model(model)
    preserved = retrieve_single_native(
        session,
        file_management,
        proposal["model_identifier"],
        evidence_path(work_root, proposal, "PRESERVE_3D_POSTIMPORT"),
        log,
    )
    report["PRESERVE_3D_POST_SHA256"] = preserved["sha256"]
    unchanged = preserved["sha256"].lower() == proposal["model_baseline_sha"].lower()
    report["PRESERVE_3D_UNCHANGED"] = "YES" if unchanged else "NO"
    if not unchanged:
        raise RuntimeError(
            "Preserved 3D native SHA-256 changed across the J17 creation."
        )

    checkout_state = target.get("checkout_state", "UNKNOWN")
    checkout_owner = target.get("checkout_owner", "")
    if checkout_state == "CHECKED_OUT":
        report["POST_IMPORT_VERIFICATION"] = (
            "CREATED_EXACT_TARGET_MANUAL_CHECKIN_REQUIRED"
        )
        set_result(
            report,
            "MANUAL_CHECKIN_REQUIRED",
            "The exact new drawing specification, drawing sheets, native payload, and "
            "unchanged 3D were verified, but the new object remains checked out "
            "by {0}. Verify it and check it in manually; J17 will not check in."
            .format(checkout_owner or "<owner unavailable>"),
            "REVIEW_REQUIRED",
        )
        return True
    if checkout_state != "CHECKED_IN":
        raise RuntimeError(
            "The new drawing specification exists, but post-import checkout state is "
            "unknown: {0}".format(target.get("checkout_raw", "") or "<none>")
        )

    if imported["sha256"].lower() == proposal["source_sha"].lower():
        report["POST_IMPORT_VERIFICATION"] = "VERIFIED_EXACT_SOURCE_SHA256"
        result = "SPECIFICATION_CREATED_VERIFIED"
        message = (
            "New exact CHECKED_IN drawing specification was created with source-matching "
            "native SHA-256; preserved 3D remained byte-identical."
        )
    else:
        report["POST_IMPORT_VERIFICATION"] = "VERIFIED_MANAGED_TRANSFORM"
        result = "SPECIFICATION_CREATED_VERIFIED_MANAGED_TRANSFORM"
        message = (
            "New exact CHECKED_IN drawing specification was created, drawing sheets and "
            "one managed native payload were proven, and preserved 3D remained "
            "byte-identical. Teamcenter rewrote the managed native payload, as "
            "validated by the J16 production contract."
        )
    set_result(report, result, message, "CREATED")
    return False


def validate_approved_rows(rows, mode, maximum):
    invalid = [row for row in rows if approval_state(row) == "INVALID"]
    if invalid:
        raise RuntimeError(
            "Every APPROVED value must be YES, NO, or blank; found {0} invalid "
            "row(s).".format(len(invalid))
        )
    if mode != "APPLY_APPROVED":
        return 0
    approved = [row for row in rows if approval_state(row) == "YES"]
    if not approved:
        raise RuntimeError("APPLY_APPROVED requires at least one APPROVED=YES row.")
    if len(approved) > maximum:
        raise RuntimeError(
            "Approved row count {0} exceeds the controlled J17 write limit {1}."
            .format(len(approved), maximum)
        )
    return len(approved)


def execute(
    session,
    file_management,
    api,
    rows,
    csv_path,
    stage_root,
    work_root,
    timestamp,
    mode,
    log,
):
    maximum = configured_max_approved_writes()
    validate_approved_rows(rows, mode, maximum)
    if mode == "APPLY_APPROVED":
        J16.require_fresh_apply_session(session)

    reports, proposals = local_preflight(
        rows, csv_path, stage_root, timestamp, mode
    )
    try:
        managed_preflight(
            session, file_management, api, proposals, mode, work_root, log
        )
    except J16.ProcessStateError as exc:
        for report in reports:
            if report.get("WRITE_ATTEMPTED") == "YES" or report.get(
                "DISPOSITION"
            ) in ("QUARANTINED", "NOT_APPROVED", "CREATED"):
                continue
            set_result(
                report,
                "FAILED_PROCESS_STATE",
                "Process-wide managed preflight failure: {0}".format(
                    J16.error_text(exc)
                ),
                "ABORTED",
            )
        return reports

    if mode == "DRY_RUN":
        return reports

    writes_attempted = 0
    for index, proposal in enumerate(proposals):
        report = proposal["report"]
        if report.get("RESULT") != "CLONE_PREFLIGHT_OK":
            continue
        if writes_attempted >= maximum:
            quarantine(
                report,
                "QUARANTINED_PRODUCTION_WRITE_LIMIT",
                "Controlled J17 write limit {0} was reached.".format(maximum),
            )
            continue

        try:
            prewrite_checks(session, file_management, proposal, work_root, log)
        except J16.ProcessStateError as exc:
            set_result(
                report,
                "FAILED_PROCESS_STATE",
                "Process-wide prewrite failure: {0}".format(J16.error_text(exc)),
                "ABORTED",
            )
            mark_later_stopped(proposals, index + 1)
            break
        except Exception as exc:
            quarantine(
                report,
                "QUARANTINED_PREWRITE",
                "Immediate prewrite checks failed; no Teamcenter write occurred.",
                exc,
            )
            continue

        apply_log = clone_log_path(proposal, mode, "APPLY")
        report["CLONE_APPLY_LOG"] = apply_log
        report["WRITE_ATTEMPTED"] = "YES"
        report["DISPOSITION"] = "WRITE_ATTEMPTED"
        writes_attempted += 1
        try:
            import_one(api, proposal, apply_log, False, log)
            review_required = verify_after_import(
                session, file_management, proposal, work_root, log
            )
            if review_required:
                mark_later_stopped(proposals, index + 1, True)
                break
        except Exception as exc:
            report["POST_IMPORT_VERIFICATION"] = "FAILED"
            set_result(
                report,
                "FAILED_IMPORT_UNVERIFIED",
                "A creation was attempted, but exact post-import drawing/3D "
                "verification failed.",
                "FAILED_AFTER_WRITE",
                exc,
            )
            log.write("  FAILED or unverified J17 creation: {0}".format(J16.error_text(exc)))
            log.write(traceback.format_exc())
            mark_later_stopped(proposals, index + 1)
            break
    return reports


def summary_counts(reports):
    return Counter(report.get("RESULT", "") or "<blank>" for report in reports)


def final_run_status(reports, mode):
    approved = [
        report
        for report in reports
        if mode == "DRY_RUN" or upper(report.get("APPROVED")) == "YES"
    ]
    if any(report.get("RESULT") == "MANUAL_CHECKIN_REQUIRED" for report in approved):
        return "REVIEW_REQUIRED"
    failed_after_write = any(
        report.get("RESULT", "").startswith(("FAILED_", "BATCH_STOPPED_"))
        for report in approved
    )
    if failed_after_write:
        return "FAILED"
    quarantined = any(
        report.get("DISPOSITION") == "QUARANTINED" for report in approved
    )
    completed = any(
        report.get("RESULT")
        in (
            "SPECIFICATION_CREATED_VERIFIED",
            "SPECIFICATION_CREATED_VERIFIED_MANAGED_TRANSFORM",
            "DRY_RUN_OK",
        )
        for report in approved
    )
    if quarantined:
        return "COMPLETED_WITH_QUARANTINE" if completed else "FAILED"
    return "SUCCESS"


def main():
    session = NXOpen.Session.GetSession()
    ufs = NXOpen.UF.UFSession.GetUFSession()
    log = J16.Log(session)
    mode = configured_mode()
    input_path = configured_input_path()
    timestamp = J16.stamp()
    output_dir = os.path.dirname(input_path) if input_path else J16.io_root()
    stage_root = os.path.join(
        output_dir, "J17_NEW_DRAWING_SPECIFICATION_STAGE_" + timestamp
    )
    work_root = os.path.join(output_dir, "J17_EVIDENCE_" + timestamp)
    evidence_zip = work_root + ".zip"
    report_path = ""
    run_log_path = ""
    reports = []
    file_management = None

    log.write("=" * 72)
    log.write("J17 CREATE NEW DRAWING SPECIFICATION UNDER EXISTING 3D MASTER")
    log.write("Build: {0} | Mode: {1}".format(BUILD, mode))
    log.write("Runtime target: NX X 2506 managed mode only")
    log.write("Destination rule: exact /specification/ must not exist; never overwrite")
    log.write("Reference rule: exact 3D must stay CHECKED_IN and UseExisting")
    log.write("Production: internal managed checks plus UF Clone dry run are automatic")
    log.write("Input: {0}".format(input_path))
    log.write("Stage: {0}".format(stage_root))
    log.write("Evidence: {0}".format(work_root))
    log.write("Evidence ZIP: {0}".format(evidence_zip))
    log.write("=" * 72)

    try:
        if mode not in VALID_MODES:
            raise RuntimeError("NX_J17_MODE must be DRY_RUN or APPLY_APPROVED.")
        if not os.path.isfile(input_path):
            raise RuntimeError("Import CSV not found: {0}".format(input_path))
        rows = read_csv(input_path)
        if not rows:
            raise RuntimeError("Import CSV contains no data rows: {0}".format(input_path))

        os.makedirs(stage_root, exist_ok=True)
        os.makedirs(work_root, exist_ok=True)
        api = J16.resolve_clone_api(ufs, log)
        _, file_management = J16.new_file_management(session)
        reports = execute(
            session,
            file_management,
            api,
            rows,
            input_path,
            stage_root,
            work_root,
            timestamp,
            mode,
            log,
        )
        report_path = os.path.join(
            output_dir, "J17_{0}_{1}.csv".format(mode, timestamp)
        )
        write_csv(report_path, reports)
        log.write("Report: {0}".format(report_path))
        for result, count in sorted(summary_counts(reports).items()):
            log.write("  {0}: {1}".format(result, count))

        status = final_run_status(reports, mode)
        log.write("FINAL STATUS: {0}".format(status))
        if status == "FAILED":
            raise RuntimeError(
                "J17 completed with a failed or wholly quarantined approved row; "
                "review {0}".format(report_path)
            )
        if status == "REVIEW_REQUIRED":
            log.write(
                "Do not rerun. Verify the new drawing specification and check it in "
                "manually only when the checkout belongs to you."
            )
        return report_path
    except Exception as exc:
        if not any(line == "FINAL STATUS: FAILED" for line in log.lines):
            log.write("FINAL STATUS: FAILED")
        log.write(J16.error_text(exc))
        log.write(traceback.format_exc())
        raise
    finally:
        J16.dispose(file_management)
        try:
            os.makedirs(output_dir, exist_ok=True)
            run_log_path = os.path.join(
                output_dir, "J17_RUN_{0}_{1}.log".format(mode, timestamp)
            )
            J16.write_log(run_log_path, log.lines)
            artifacts = [report_path, run_log_path]
            for report in reports:
                artifacts.extend(
                    [
                        report.get("CLONE_PREFLIGHT_LOG", ""),
                        report.get("CLONE_APPLY_LOG", ""),
                    ]
                )
            J16.zip_artifacts(evidence_zip, work_root, artifacts)
        except Exception:
            pass


if __name__ == "__main__":
    main()
