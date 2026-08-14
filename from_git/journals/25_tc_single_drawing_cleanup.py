"""J25 - reduce one managed 3D revision to one drawing specification.

NX X 2506 managed mode only.

J07 can discover drawing specifications named dwg1..dwg9 beneath one 3D
Item/Revision. J25 keeps the explicitly selected drawing and removes every
explicitly approved extra drawing so the revision has one final DWG.

Important Teamcenter semantics:
- NXOpen does not expose a supported relation-only detach operation here.
- APPLY_APPROVED calls PDM FileManagement.DeleteExistingAttachedFiles with
  keepEmptyDataset=False. This removes every associated file from an extra
  drawing dataset and then removes the empty dataset. The specification
  relation disappears because the dataset is removed; the dataset is NOT
  retained as an orphan.
- Before each removal, J25 downloads every associated file into the run's
  BACKUP folder and records SHA-256 evidence.
- J25 never modifies or deletes the selected KEEP_DWG_INDEX drawing or the 3D
  master. It never checks out, checks in, revises, or saves an NX part.

DRY_RUN is the default. APPLY_APPROVED requires APPROVED=YES, a nonblank
ENGINEER, CONFIRMATION=REMOVE_EXTRA_DRAWINGS, and an exact
EXPECTED_REMOVE_DWG_INDICES list matching the live dwg1..dwg9 inventory.

Target: NX X 2506 embedded Python.
Run via: NX > Tools > Journal > Play
"""

import csv
import datetime
import hashlib
import importlib.util
import json
import os
import re
import shutil
import traceback

import NXOpen


# ============================================================================
# USER SETTINGS
# ============================================================================
USER_INPUT_CSV = r""  # blank => <I/O root>\NX_TC_SINGLE_DRAWING_SCOPE.csv
USER_MODE = "DRY_RUN"
# Optional environment overrides:
#   NX_TC_SINGLE_DRAWING_FILE=<full CSV path>
#   NX_J25_MODE=DRY_RUN or APPLY_APPROVED
#   NX_J25_MAX_DELETIONS=1..100 (default 25)
# ============================================================================

BUILD = "J25-TCX-SINGLE-DRAWING-CLEANUP-NX2506-V1"
DEFAULT_INPUT = "NX_TC_SINGLE_DRAWING_SCOPE.csv"
OUTPUT_FOLDER = "NX_TC_SINGLE_DRAWING_CLEANUP"
VALID_MODES = ("DRY_RUN", "APPLY_APPROVED")
CONFIRMATION_TEXT = "REMOVE_EXTRA_DRAWINGS"
MAX_DRAWING_INDEX = 9
DEFAULT_MAX_DELETIONS = 25

REQUIRED_COLUMNS = (
    "PART_NUMBER",
    "REVISION",
    "KEEP_DWG_INDEX",
    "EXPECTED_REMOVE_DWG_INDICES",
    "APPROVED",
    "ENGINEER",
    "CONFIRMATION",
)

REPORT_COLUMNS = (
    "RUN_TIMESTAMP", "MODE", "CSV_ROW", "PART_NUMBER", "REVISION",
    "MASTER_IDENTIFIER", "KEEP_DWG_INDEX", "KEEP_IDENTIFIER",
    "DISCOVERED_DWG_INDICES", "EXPECTED_REMOVE_DWG_INDICES",
    "LIVE_REMOVE_DWG_INDICES", "MASTER_CHECKOUT_STATE",
    "KEEP_CHECKOUT_STATE", "KEEP_DRAWING_SHEET_COUNT",
    "EXTRA_CHECKOUT_STATES", "EXTRA_DRAWING_SHEET_COUNTS",
    "EXTRA_LOADED_AT_START", "BACKUP_FILES", "BACKUP_SHA256",
    "DELETE_API_RESULTS", "REMOVED_DWG_INDICES", "POSTCHECK_DWG_INDICES",
    "APPROVED", "ENGINEER", "CONFIRMATION", "WRITE_ATTEMPTED", "RESULT",
    "MESSAGE",
)


def load_j16():
    """Load the tested managed-file and identity helpers beside this journal."""
    path = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "16_tc_offline_drawing_import.py",
    )
    if not os.path.isfile(path):
        raise RuntimeError("J16 dependency not found beside J25: {0}".format(path))
    spec = importlib.util.spec_from_file_location("nx_journal_16_for_j25", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    required = (
        "Log", "clean", "upper", "env", "io_root", "error_text", "dispose",
        "unwrap_open_result", "find_loaded_by_identifier", "journal_identifier",
        "query_pdm_checkout", "close_opened_part", "new_file_management",
        "collect_pdm_files", "pdm_file_name", "locate_downloaded_files",
        "release_pdm_files", "safe_folder_name",
    )
    missing = [name for name in required if not hasattr(module, name)]
    if missing:
        raise RuntimeError(
            "J16 is incompatible with J25; missing: {0}".format(", ".join(missing))
        )
    return module


J16 = load_j16()
clean = J16.clean
upper = J16.upper


def configured_mode():
    mode = upper(J16.env("NX_J25_MODE") or USER_MODE or "DRY_RUN")
    if mode not in VALID_MODES:
        raise RuntimeError(
            "NX_J25_MODE must be DRY_RUN or APPLY_APPROVED; received {0}.".format(
                mode or "<blank>"
            )
        )
    return mode


def configured_input_path():
    value = J16.env("NX_TC_SINGLE_DRAWING_FILE") or clean(USER_INPUT_CSV)
    if value:
        return os.path.abspath(os.path.expanduser(value))
    return os.path.join(J16.io_root(), DEFAULT_INPUT)


def configured_max_deletions():
    raw = J16.env("NX_J25_MAX_DELETIONS")
    try:
        value = int(raw) if raw else DEFAULT_MAX_DELETIONS
    except Exception:
        raise RuntimeError("NX_J25_MAX_DELETIONS must be an integer.")
    if value < 1 or value > 100:
        raise RuntimeError("NX_J25_MAX_DELETIONS must be between 1 and 100.")
    return value


def master_id(part_number, revision):
    return "@DB/{0}/{1}".format(part_number, revision)


def dataset_name(part_number, revision, drawing_index):
    return "{0}-{1}-dwg{2}".format(part_number, revision, drawing_index)


def drawing_id(part_number, revision, drawing_index):
    return "@DB/{0}/{1}/specification/{2}".format(
        part_number, revision, dataset_name(part_number, revision, drawing_index)
    )


def parse_index(value, label="DWG index"):
    try:
        index = int(clean(value))
    except Exception:
        raise RuntimeError(
            "{0} must be an integer from 1 to {1}.".format(label, MAX_DRAWING_INDEX)
        )
    if index < 1 or index > MAX_DRAWING_INDEX:
        raise RuntimeError(
            "{0} must be from 1 to {1}.".format(label, MAX_DRAWING_INDEX)
        )
    return index


def parse_index_list(value):
    text = clean(value)
    if not text:
        return []
    tokens = [item for item in re.split(r"[\s,;|]+", text) if item]
    indices = [parse_index(item, "EXPECTED_REMOVE_DWG_INDICES") for item in tokens]
    if len(set(indices)) != len(indices):
        raise RuntimeError("EXPECTED_REMOVE_DWG_INDICES contains a duplicate.")
    return sorted(indices)


def indices_text(values):
    return "|".join(str(value) for value in sorted(values))


def validate_identity_token(value, label):
    text = clean(value)
    if not text:
        raise RuntimeError("{0} is blank.".format(label))
    if any(char in text for char in "/\\") or text in (".", ".."):
        raise RuntimeError("{0} contains an invalid identity separator.".format(label))
    return text


def read_input(path):
    last_error = None
    for encoding in ("utf-8-sig", "utf-8", "cp1252"):
        try:
            with open(path, "r", encoding=encoding, newline="") as handle:
                reader = csv.DictReader(handle)
                headers = [clean(name) for name in (reader.fieldnames or [])]
                missing = [name for name in REQUIRED_COLUMNS if name not in headers]
                if missing:
                    raise RuntimeError(
                        "Input CSV is missing column(s): {0}".format(", ".join(missing))
                    )
                rows = []
                for number, source in enumerate(reader, 2):
                    row = {
                        clean(key): clean(value)
                        for key, value in source.items()
                        if key is not None
                    }
                    if not any(row.values()):
                        continue
                    row["_CSV_ROW"] = number
                    rows.append(row)
                return rows
        except UnicodeDecodeError as error:
            last_error = error
    raise RuntimeError("Unable to decode {0}: {1}".format(path, last_error))


def blank_report(row, mode, timestamp):
    report = {name: "" for name in REPORT_COLUMNS}
    report.update(
        RUN_TIMESTAMP=timestamp,
        MODE=mode,
        CSV_ROW=row.get("_CSV_ROW", ""),
        PART_NUMBER=clean(row.get("PART_NUMBER")),
        REVISION=clean(row.get("REVISION")),
        KEEP_DWG_INDEX=clean(row.get("KEEP_DWG_INDEX")),
        EXPECTED_REMOVE_DWG_INDICES=clean(row.get("EXPECTED_REMOVE_DWG_INDICES")),
        APPROVED=clean(row.get("APPROVED")),
        ENGINEER=clean(row.get("ENGINEER")),
        CONFIRMATION=clean(row.get("CONFIRMATION")),
        WRITE_ATTEMPTED="NO",
        RESULT="PENDING",
    )
    return report


def collection_count(value):
    try:
        return int(value.Count)
    except Exception:
        try:
            return len(list(value))
        except Exception:
            return -1


def normalized_identifier(value):
    return upper(value).replace("\\", "/")


def inspect_exact(session, identifier, log):
    """Open, identity-check, inspect, and close one exact managed part."""
    part = J16.find_loaded_by_identifier(session, identifier)
    loaded_at_start = part is not None
    load_status = None
    opened_here = False
    result = {
        "state": "NOT_OPENABLE", "identifier": identifier,
        "opened_identifier": "", "checkout_state": "UNKNOWN",
        "checkout_owner": "", "checkout_raw": "", "drawing_sheet_count": -1,
        "loaded_at_start": loaded_at_start, "detail": "", "error_code": "",
    }
    try:
        if part is None:
            part, load_status = J16.unwrap_open_result(session.Parts.OpenBase(identifier))
            opened_here = True
        if part is None:
            result["detail"] = "OpenBase returned no part."
            return result
        actual = J16.journal_identifier(part)
        result["opened_identifier"] = actual
        if normalized_identifier(actual) != normalized_identifier(identifier):
            result["state"] = "IDENTITY_MISMATCH"
            result["detail"] = "Opened a different managed identity: {0}".format(
                actual or "<blank>"
            )
            return result
        checkout = J16.query_pdm_checkout(part)
        result.update(
            state="EXISTS",
            checkout_state=checkout.get("state", "UNKNOWN"),
            checkout_owner=checkout.get("owner", ""),
            checkout_raw=checkout.get("raw", ""),
            drawing_sheet_count=collection_count(getattr(part, "DrawingSheets", [])),
            detail="Exact managed identity opened.",
        )
        return result
    except Exception as error:
        result["detail"] = J16.error_text(error)
        result["error_code"] = clean(getattr(error, "ErrorCode", ""))
        return result
    finally:
        J16.dispose(load_status)
        if opened_here and part is not None:
            J16.close_opened_part(part, log)


def discover_drawings(session, part_number, revision, log):
    inspections = {}
    for index in range(1, MAX_DRAWING_INDEX + 1):
        result = inspect_exact(session, drawing_id(part_number, revision, index), log)
        if result["state"] == "EXISTS":
            inspections[index] = result
        elif result["state"] == "IDENTITY_MISMATCH":
            raise RuntimeError(
                "DWG{0} identity mismatch: {1}".format(index, result["detail"])
            )
    return inspections


def sha256(path):
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        while True:
            block = handle.read(1024 * 1024)
            if not block:
                break
            digest.update(block)
    return digest.hexdigest()


def unique_pdm_files(value):
    files = []
    for candidate in J16.collect_pdm_files(value):
        if all(candidate is not existing for existing in files):
            files.append(candidate)
    return files


def flatten_int_statuses(value):
    statuses = []

    def visit(candidate):
        if isinstance(candidate, (tuple, list)):
            for item in candidate:
                visit(item)
        elif isinstance(candidate, int) and not isinstance(candidate, bool):
            statuses.append(candidate)

    visit(value)
    return statuses


def backup_and_delete_target(
    session, file_management, part_number, revision, drawing_index,
    backup_root, log,
):
    """Back up one exact drawing, then delete all files and its empty dataset."""
    identifier = drawing_id(part_number, revision, drawing_index)
    if J16.find_loaded_by_identifier(session, identifier) is not None:
        raise RuntimeError("{0} is loaded; close it before APPLY_APPROVED.".format(identifier))
    get_files = getattr(file_management, "GetAssociatedFiles", None)
    download_files = getattr(file_management, "DownloadAssociatedFiles", None)
    delete_files = getattr(file_management, "DeleteExistingAttachedFiles", None)
    if not all(callable(value) for value in (get_files, download_files, delete_files)):
        raise RuntimeError(
            "PDM GetAssociatedFiles, DownloadAssociatedFiles, or "
            "DeleteExistingAttachedFiles is unavailable."
        )

    part = None
    load_status = None
    pdm_files = []
    resource_files = []
    original_cwd = os.getcwd()
    backup_rows = []
    try:
        part, load_status = J16.unwrap_open_result(session.Parts.OpenBase(identifier))
        if part is None:
            raise RuntimeError("OpenBase returned no part for {0}.".format(identifier))
        actual = J16.journal_identifier(part)
        if normalized_identifier(actual) != normalized_identifier(identifier):
            raise RuntimeError(
                "Opened identity does not match delete target: {0}".format(
                    actual or "<blank>"
                )
            )
        pdm_files = unique_pdm_files(get_files([part], []))
        resource_files.extend(pdm_files)
        names = [J16.pdm_file_name(value) for value in pdm_files]
        if not pdm_files or any(not name for name in names):
            raise RuntimeError("Could not prove all associated file names for {0}.".format(identifier))
        native_names = [name for name in names if name.lower().endswith(".prt")]
        if len(native_names) != 1:
            raise RuntimeError(
                "Expected exactly one native .prt for {0}; found {1}: {2}".format(
                    identifier, len(native_names), " | ".join(names)
                )
            )

        returned = unique_pdm_files(download_files([part], pdm_files))
        for value in returned:
            if all(value is not existing for existing in resource_files):
                resource_files.append(value)
        download_cwd = os.getcwd()
        all_names = list(names) + [
            J16.pdm_file_name(value)
            for value in returned
            if J16.pdm_file_name(value)
        ]
        physical = J16.locate_downloaded_files(all_names, download_cwd)
        expected_basenames = {os.path.basename(name).lower() for name in names}
        found_basenames = {os.path.basename(path).lower() for path in physical.values()}
        missing = sorted(expected_basenames - found_basenames)
        if missing:
            raise RuntimeError(
                "Backup download did not materialize: {0}".format(", ".join(missing))
            )

        target_backup = os.path.join(
            backup_root,
            J16.safe_folder_name(
                "{0}_{1}_DWG{2}".format(part_number, revision, drawing_index)
            ),
        )
        os.makedirs(target_backup, exist_ok=True)
        for source in sorted(physical.values(), key=lambda value: value.lower()):
            destination = os.path.join(
                target_backup, J16.safe_folder_name(os.path.basename(source))
            )
            shutil.copy2(source, destination)
            backup_rows.append({"file": destination, "sha256": sha256(destination)})
        if not backup_rows:
            raise RuntimeError("No backup files were copied for {0}.".format(identifier))

        os.chdir(original_cwd)
        J16.dispose(load_status)
        load_status = None
        J16.close_opened_part(part, log)
        part = None
        raw_delete = delete_files(pdm_files, False)
        return {
            "identifier": identifier,
            "backup": backup_rows,
            "delete_result": repr(raw_delete)[:2000],
            "delete_statuses": flatten_int_statuses(raw_delete),
            "keep_empty_dataset": False,
        }
    finally:
        try:
            os.chdir(original_cwd)
        finally:
            J16.dispose(load_status)
            if part is not None:
                J16.close_opened_part(part, log)
            J16.release_pdm_files(resource_files)


def validate_plan(row, session, log):
    part_number = validate_identity_token(row.get("PART_NUMBER"), "PART_NUMBER")
    revision = validate_identity_token(row.get("REVISION"), "REVISION")
    keep = parse_index(row.get("KEEP_DWG_INDEX"), "KEEP_DWG_INDEX")
    expected_remove = parse_index_list(row.get("EXPECTED_REMOVE_DWG_INDICES"))
    if keep in expected_remove:
        raise RuntimeError("KEEP_DWG_INDEX is also listed for removal.")
    master = inspect_exact(session, master_id(part_number, revision), log)
    if master["state"] != "EXISTS":
        raise RuntimeError("Exact 3D master could not be opened: {0}".format(master["detail"]))
    drawings = discover_drawings(session, part_number, revision, log)
    discovered = sorted(drawings)
    if keep not in drawings:
        raise RuntimeError("Selected KEEP_DWG_INDEX DWG{0} does not exist.".format(keep))
    if drawings[keep]["drawing_sheet_count"] < 1:
        raise RuntimeError("Selected final DWG{0} has no drawing sheets.".format(keep))
    live_remove = sorted(index for index in drawings if index != keep)
    if expected_remove != live_remove:
        raise RuntimeError(
            "Live extras are [{0}], but EXPECTED_REMOVE_DWG_INDICES is [{1}].".format(
                indices_text(live_remove), indices_text(expected_remove)
            )
        )
    return {
        "part_number": part_number, "revision": revision, "keep": keep,
        "expected_remove": expected_remove, "live_remove": live_remove,
        "master": master, "drawings": drawings, "discovered": discovered,
    }


def require_apply_authorization(row, plan):
    if upper(row.get("APPROVED")) != "YES":
        raise RuntimeError("APPLY_APPROVED requires APPROVED=YES.")
    if not clean(row.get("ENGINEER")):
        raise RuntimeError("APPLY_APPROVED requires a nonblank ENGINEER.")
    if upper(row.get("CONFIRMATION")) != CONFIRMATION_TEXT:
        raise RuntimeError(
            "APPLY_APPROVED requires CONFIRMATION={0}.".format(CONFIRMATION_TEXT)
        )
    if not plan["live_remove"]:
        return
    if plan["master"]["checkout_state"] != "CHECKED_IN":
        raise RuntimeError("The 3D master is not proven CHECKED_IN.")
    for index, inspection in plan["drawings"].items():
        if inspection["checkout_state"] != "CHECKED_IN":
            raise RuntimeError(
                "DWG{0} is not proven CHECKED_IN; state={1}, owner={2}.".format(
                    index, inspection["checkout_state"],
                    inspection["checkout_owner"] or "<blank>",
                )
            )
    loaded_extras = [
        index for index in plan["live_remove"]
        if plan["drawings"][index]["loaded_at_start"]
    ]
    if loaded_extras:
        raise RuntimeError(
            "Close these extra drawing parts before apply: {0}.".format(
                indices_text(loaded_extras)
            )
        )


def fill_plan_report(report, plan):
    drawings = plan["drawings"]
    report.update(
        MASTER_IDENTIFIER=master_id(plan["part_number"], plan["revision"]),
        KEEP_DWG_INDEX=str(plan["keep"]),
        KEEP_IDENTIFIER=drawing_id(plan["part_number"], plan["revision"], plan["keep"]),
        DISCOVERED_DWG_INDICES=indices_text(plan["discovered"]),
        EXPECTED_REMOVE_DWG_INDICES=indices_text(plan["expected_remove"]),
        LIVE_REMOVE_DWG_INDICES=indices_text(plan["live_remove"]),
        MASTER_CHECKOUT_STATE=plan["master"]["checkout_state"],
        KEEP_CHECKOUT_STATE=drawings[plan["keep"]]["checkout_state"],
        KEEP_DRAWING_SHEET_COUNT=str(drawings[plan["keep"]]["drawing_sheet_count"]),
        EXTRA_CHECKOUT_STATES=" | ".join(
            "DWG{0}:{1}".format(index, drawings[index]["checkout_state"])
            for index in plan["live_remove"]
        ),
        EXTRA_DRAWING_SHEET_COUNTS=" | ".join(
            "DWG{0}:{1}".format(index, drawings[index]["drawing_sheet_count"])
            for index in plan["live_remove"]
        ),
        EXTRA_LOADED_AT_START=indices_text(
            index for index in plan["live_remove"]
            if drawings[index]["loaded_at_start"]
        ),
    )


def execute(rows, session, file_management, mode, backup_root, timestamp, log):
    reports = []
    prepared = []
    seen = set()
    max_deletions = configured_max_deletions()

    # Phase 1 is read-only for the complete CSV. In apply mode, one bad row
    # blocks every row so a batch cannot be partly changed due to a later typo.
    for row in rows:
        report = blank_report(row, mode, timestamp)
        reports.append(report)
        try:
            key = (upper(row.get("PART_NUMBER")), upper(row.get("REVISION")))
            if key in seen:
                raise RuntimeError("Duplicate PART_NUMBER + REVISION input row.")
            seen.add(key)
            plan = validate_plan(row, session, log)
            fill_plan_report(report, plan)
            if mode == "APPLY_APPROVED":
                require_apply_authorization(row, plan)
            prepared.append((row, report, plan))
            if mode == "DRY_RUN":
                report["RESULT"] = "DRY_RUN_READY" if plan["live_remove"] else "ALREADY_SINGLE_DWG"
                report["MESSAGE"] = (
                    "Exact live inventory matches the requested keep/remove plan. No Teamcenter data was changed."
                    if plan["live_remove"]
                    else "Only the selected final drawing is associated; no removal is needed."
                )
            else:
                report["RESULT"] = "APPLY_PREFLIGHT_READY"
                report["MESSAGE"] = "Complete-row apply preflight passed; no write attempted yet."
        except Exception as error:
            report["RESULT"] = "BLOCKED"
            report["MESSAGE"] = J16.error_text(error)
            log.write(
                "  ROW {0} BLOCKED: {1}".format(report["CSV_ROW"], report["MESSAGE"])
            )

    if mode == "DRY_RUN":
        return reports
    if any(report["RESULT"] == "BLOCKED" for report in reports):
        for _, report, _ in prepared:
            report["RESULT"] = "BLOCKED_BY_BATCH_PREFLIGHT"
            report["MESSAGE"] = "Another input row failed; J25 performed no Teamcenter writes."
        return reports

    planned_deletions = sum(len(plan["live_remove"]) for _, _, plan in prepared)
    if planned_deletions > max_deletions:
        for _, report, _ in prepared:
            report["RESULT"] = "BLOCKED"
            report["MESSAGE"] = (
                "Approved batch requests {0} deletions, exceeding "
                "NX_J25_MAX_DELETIONS={1}; no writes were attempted."
            ).format(planned_deletions, max_deletions)
        return reports

    # Phase 2 mutates only after every row and the whole-batch cap have passed.
    for position, (row, report, plan) in enumerate(prepared):
        if not plan["live_remove"]:
            report["RESULT"] = "ALREADY_SINGLE_DWG"
            report["MESSAGE"] = "No extra drawing dataset exists."
            continue
        try:
            recheck = validate_plan(row, session, log)
            require_apply_authorization(row, recheck)
            report["WRITE_ATTEMPTED"] = "YES"
            removed, backup_files, backup_hashes, api_results = [], [], [], []
            for index in recheck["live_remove"]:
                outcome = backup_and_delete_target(
                    session, file_management, recheck["part_number"],
                    recheck["revision"], index, backup_root, log,
                )
                backup_files.extend(item["file"] for item in outcome["backup"])
                backup_hashes.extend(item["sha256"] for item in outcome["backup"])
                api_results.append("DWG{0}:{1}".format(index, outcome["delete_result"]))
                report["BACKUP_FILES"] = " | ".join(backup_files)
                report["BACKUP_SHA256"] = " | ".join(backup_hashes)
                report["DELETE_API_RESULTS"] = " | ".join(api_results)
                statuses = outcome["delete_statuses"]
                if not statuses or any(status != 0 for status in statuses):
                    raise RuntimeError(
                        "DWG{0} delete API did not return all-zero status evidence: {1}."
                        .format(index, statuses)
                    )
                post = inspect_exact(
                    session,
                    drawing_id(recheck["part_number"], recheck["revision"], index),
                    log,
                )
                if post["state"] == "EXISTS":
                    raise RuntimeError(
                        "DWG{0} still opens after DeleteExistingAttachedFiles.".format(index)
                    )
                if post["state"] == "IDENTITY_MISMATCH":
                    raise RuntimeError("DWG{0} post-delete identity mismatch.".format(index))
                removed.append(index)
                report["REMOVED_DWG_INDICES"] = indices_text(removed)

            final_drawings = discover_drawings(
                session, recheck["part_number"], recheck["revision"], log
            )
            report["POSTCHECK_DWG_INDICES"] = indices_text(final_drawings)
            if sorted(final_drawings) != [recheck["keep"]]:
                raise RuntimeError(
                    "Postcheck expected only DWG{0}; found [{1}].".format(
                        recheck["keep"], indices_text(final_drawings)
                    )
                )
            if final_drawings[recheck["keep"]]["drawing_sheet_count"] < 1:
                raise RuntimeError("Retained drawing no longer proves drawing sheets.")
            report["RESULT"] = "SINGLE_DWG_VERIFIED"
            report["MESSAGE"] = (
                "Extra drawing datasets were backed up and removed; only the selected final drawing remains openable."
            )
        except Exception as error:
            report["RESULT"] = (
                "PARTIAL_FAILURE" if report["REMOVED_DWG_INDICES"]
                else "FAILED"
            )
            report["MESSAGE"] = J16.error_text(error)
            log.write(
                "  ROW {0} {1}: {2}".format(
                    report["CSV_ROW"], report["RESULT"], report["MESSAGE"]
                )
            )
            for _, later_report, _ in prepared[position + 1:]:
                later_report["RESULT"] = "SKIPPED_AFTER_FAILURE"
                later_report["MESSAGE"] = (
                    "A prior apply row failed; no write was attempted for this row."
                )
            break
    return reports


def write_csv(path, reports):
    with open(path, "w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=REPORT_COLUMNS)
        writer.writeheader()
        for report in reports:
            writer.writerow({name: report.get(name, "") for name in REPORT_COLUMNS})


def write_json(path, reports):
    payload = {
        "schema_version": 1,
        "journal_build": BUILD,
        "semantics": (
            "DeleteExistingAttachedFiles(files, keepEmptyDataset=False): "
            "backup files, delete extra drawing dataset, retain selected drawing"
        ),
        "reports": reports,
    }
    with open(path, "w", encoding="utf-8") as handle:
        json.dump(payload, handle, indent=2, ensure_ascii=False)


def main():
    session = NXOpen.Session.GetSession()
    log = J16.Log(session)
    mode = configured_mode()
    input_path = configured_input_path()
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    run_root = os.path.join(J16.io_root(), OUTPUT_FOLDER, timestamp)
    backup_root = os.path.join(run_root, "BACKUP")
    os.makedirs(backup_root, exist_ok=False)
    csv_path = os.path.join(run_root, "J25_RESULT_{0}.csv".format(timestamp))
    json_path = os.path.join(run_root, "J25_RESULT_{0}.json".format(timestamp))
    log_path = os.path.join(run_root, "J25_LOG_{0}.txt".format(timestamp))

    log.write("=" * 72)
    log.write("J25 TEAMCENTER SINGLE-DRAWING CLEANUP")
    log.write("Build: " + BUILD)
    log.write("Mode: " + mode)
    log.write("Input: " + input_path)
    log.write(
        "Mutation: backup, then delete extra dataset via "
        "DeleteExistingAttachedFiles(..., keepEmptyDataset=False)"
    )
    log.write("This is not a relation-only detach and does not retain an orphan dataset.")
    log.write("=" * 72)

    file_management = None
    try:
        if not os.path.isfile(input_path):
            raise RuntimeError("Input CSV not found: {0}".format(input_path))
        rows = read_input(input_path)
        if not rows:
            raise RuntimeError("Input CSV contains no data rows.")
        _, file_management = J16.new_file_management(session)
        reports = execute(
            rows, session, file_management, mode, backup_root, timestamp, log
        )
        write_csv(csv_path, reports)
        write_json(json_path, reports)
        log.write("CSV: " + csv_path)
        log.write("JSON: " + json_path)
        counts = {}
        for report in reports:
            result = report["RESULT"]
            counts[result] = counts.get(result, 0) + 1
        log.write(
            "Results: " + ", ".join(
                "{0}={1}".format(key, counts[key]) for key in sorted(counts)
            )
        )
    except Exception as error:
        log.write("J25 FAILED: " + J16.error_text(error))
        log.write(traceback.format_exc())
        raise
    finally:
        J16.dispose(file_management)
        try:
            with open(log_path, "w", encoding="utf-8-sig") as handle:
                handle.write("\n".join(log.lines) + "\n")
        except Exception:
            pass


def get_unload_option(dummy):
    return NXOpen.Session.LibraryUnloadOption.Immediately


if __name__ == "__main__":
    main()
