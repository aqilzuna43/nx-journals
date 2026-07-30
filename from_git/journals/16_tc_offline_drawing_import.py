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

Run DRY_RUN first. APPLY_APPROVED writes only APPROVED=YES rows with ENGINEER.
"""

import csv
import datetime
import hashlib
import os
import traceback
from collections import Counter

import NXOpen
import NXOpen.UF


# ============================================================================
# USER SETTINGS
# ============================================================================
USER_IMPORT_CSV = r"C:\Users\my62022696\Desktop\TEST_IMPORT\NX_TC_DRAWING_IMPORT.csv"  # blank => <I/O root>\NX_TC_DRAWING_IMPORT.csv
USER_MODE = "APPLY_APPROVED"  # DRY_RUN | APPLY_APPROVED
# Optional environment overrides:
#   NX_TC_DRAWING_IMPORT_FILE=<full CSV path>
#   NX_J16_MODE=DRY_RUN or APPLY_APPROVED
# ============================================================================

BUILD = "J16-TCX-DRAWING-IMPORT-NX2506-V1"
DEFAULT_INPUT = "NX_TC_DRAWING_IMPORT.csv"
VALID_MODES = ("DRY_RUN", "APPLY_APPROVED")

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
    "PREFLIGHT_SHA256",
    "CHANGED_FROM_BASELINE",
    "APPROVED",
    "ENGINEER",
    "DEFAULT_IMPORT_ACTION",
    "DRAWING_IMPORT_ACTION",
    "CLONE_PREFLIGHT",
    "WRITE_ATTEMPTED",
    "RESULT",
    "MESSAGE",
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


def configured_input_path():
    configured = env("NX_TC_DRAWING_IMPORT_FILE") or clean(USER_IMPORT_CSV)
    if configured:
        return os.path.abspath(os.path.expanduser(configured))
    return os.path.join(io_root(), DEFAULT_INPUT)


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


def sha256(path):
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        while True:
            block = handle.read(1024 * 1024)
            if not block:
                break
            digest.update(block)
    return digest.hexdigest()


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
        "PREFLIGHT_SHA256": "",
        "CHANGED_FROM_BASELINE": "",
        "APPROVED": row.get("APPROVED", ""),
        "ENGINEER": row.get("ENGINEER", ""),
        "DEFAULT_IMPORT_ACTION": "UseExisting",
        "DRAWING_IMPORT_ACTION": "Overwrite",
        "CLONE_PREFLIGHT": "NOT_RUN",
        "WRITE_ATTEMPTED": "NO",
        "RESULT": "",
        "MESSAGE": "",
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

        approval = approval_state(row)
        if approval == "INVALID":
            set_error(
                report,
                "ERROR_APPROVAL_VALUE",
                "APPROVED must be YES, NO, or blank.",
            )
            continue

        # In apply mode, unapproved rows are not candidates and cannot block
        # approved rows because of stale/missing local files.
        if mode == "APPLY_APPROVED" and approval != "YES":
            report["RESULT"] = "NOT_APPROVED"
            report["MESSAGE"] = "No write authorized for this row."
            continue

        try:
            part_number, revision, drawing_index = parse_target(row)
        except Exception as exc:
            set_error(report, "ERROR_INPUT", error_text(exc), exc)
            continue

        target_key = (upper(part_number), upper(revision), drawing_index)
        if target_key in duplicate_keys:
            set_error(
                report,
                "ERROR_DUPLICATE_TARGET",
                "The same PART_NUMBER/REVISION/DWG_INDEX appears more than once.",
            )
            continue

        if mode == "APPLY_APPROVED" and not clean(row.get("ENGINEER")):
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
            "drawing": drawing,
            "identifier": identifier,
            "preflight_sha": current_sha,
        })

    return reports, proposals


def apply_blocking_error(report):
    if report.get("RESULT") == "ERROR_APPROVAL_VALUE":
        return True
    if upper(report.get("APPROVED")) != "YES":
        return False
    return report.get("RESULT", "").startswith("ERROR_")


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


def mark_batch_aborted(proposals, result, message):
    for proposal in proposals:
        report = proposal["report"]
        if report.get("RESULT") in ("LOCAL_PREFLIGHT_OK", "CLONE_PREFLIGHT_OK"):
            report["RESULT"] = result
            report["MESSAGE"] = message


def run_dry_run(api, reports, proposals, log, mode, stop_on_failure):
    failed = False
    for proposal in proposals:
        report = proposal["report"]
        logfile = clone_log_path(proposal, mode, "PREFLIGHT")
        report["CLONE_LOG"] = logfile
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
            failed = True
            if stop_on_failure:
                break
    return failed


def execute(api, rows, csv_path, timestamp, mode, log):
    reports, proposals = local_preflight(rows, csv_path, timestamp, mode)

    if mode == "DRY_RUN":
        run_dry_run(api, reports, proposals, log, mode, False)
        return reports

    # APPLY_APPROVED: fail closed before any Teamcenter write.
    blocking = [report for report in reports if apply_blocking_error(report)]
    if blocking:
        mark_batch_aborted(
            proposals,
            "BATCH_ABORTED_LOCAL_PREFLIGHT",
            "At least one approved row failed local preflight. No Teamcenter write was attempted.",
        )
        return reports

    # Dry-run every approved changed row before the first actual Teamcenter write.
    clone_preflight_failed = run_dry_run(api, reports, proposals, log, mode, True)
    if clone_preflight_failed:
        mark_batch_aborted(
            proposals,
            "BATCH_ABORTED_CLONE_PREFLIGHT",
            "UF Clone preflight failed for the approved batch. No Teamcenter write was attempted.",
        )
        return reports

    for index, proposal in enumerate(proposals):
        report = proposal["report"]

        # Guard against the local drawing changing after clone preflight but
        # before the actual import in the same J16 run.
        try:
            apply_sha = sha256(proposal["drawing"])
        except Exception as exc:
            set_error(
                report,
                "ERROR_FILE_CHANGED_CHECK",
                "Could not reread drawing immediately before import.",
                exc,
            )
            mark_batch_aborted(
                proposals[index + 1 :],
                "BATCH_STOPPED_BEFORE_WRITE",
                "A previous row failed the final local-file check. No write was attempted for this row.",
            )
            break

        if apply_sha.lower() != proposal["preflight_sha"].lower():
            set_error(
                report,
                "ERROR_FILE_CHANGED_AFTER_PREFLIGHT",
                "DRAWING_FILE changed after clone preflight. J16 stopped before writing this row.",
            )
            mark_batch_aborted(
                proposals[index + 1 :],
                "BATCH_STOPPED_BEFORE_WRITE",
                "A previous row changed after preflight. No write was attempted for this row.",
            )
            break

        logfile = clone_log_path(proposal, mode, "APPLY")
        report["CLONE_LOG"] = logfile
        report["WRITE_ATTEMPTED"] = "YES"
        try:
            import_one(api, proposal["drawing"], logfile, False, log)
            report["RESULT"] = "IMPORT_APPLIED"
            report["MESSAGE"] = (
                "UF Clone apply returned successfully. Exact drawing=Overwrite; "
                "related 3D/reference objects=UseExisting. Reopen the managed drawing "
                "from Teamcenter to confirm final persistence."
            )
        except Exception as exc:
            set_error(
                report,
                "FAILED_IMPORT_APPLY",
                "UF Clone apply failed.",
                exc,
            )
            log.write("  FAILED apply: {0}".format(error_text(exc)))
            log.write(traceback.format_exc())
            mark_batch_aborted(
                proposals[index + 1 :],
                "BATCH_STOPPED_AFTER_RUNTIME_FAILURE",
                "A previous approved import failed. No write was attempted for this row.",
            )
            break

    return reports


def summary_counts(reports):
    return Counter(report.get("RESULT", "") or "<blank>" for report in reports)


def has_failure(reports, mode):
    failure_prefixes = (
        "ERROR_",
        "FAILED_",
        "BATCH_ABORTED_",
        "BATCH_STOPPED_",
    )
    for report in reports:
        result = report.get("RESULT", "")
        if result.startswith(failure_prefixes):
            if mode == "APPLY_APPROVED":
                if result == "ERROR_APPROVAL_VALUE" or upper(report.get("APPROVED")) == "YES":
                    return True
                continue
            return True
    return False


def main():
    session = NXOpen.Session.GetSession()
    ufs = NXOpen.UF.UFSession.GetUFSession()
    log = Log(session)
    current_mode = configured_mode()
    input_path = configured_input_path()
    timestamp = stamp()

    log.write("=" * 72)
    log.write("J16 TEAMCENTER X STANDALONE DRAWING IMPORT")
    log.write("Build: {0} | Mode: {1}".format(BUILD, current_mode))
    log.write("Runtime target: NX X 2506 only")
    log.write("Input: {0}".format(input_path))
    log.write("=" * 72)

    report_path = ""
    try:
        if current_mode not in VALID_MODES:
            raise RuntimeError(
                "USER_MODE/NX_J16_MODE must be DRY_RUN or APPLY_APPROVED."
            )
        if not os.path.isfile(input_path):
            raise RuntimeError("Import CSV not found: {0}".format(input_path))

        rows = read_csv(input_path)
        if not rows:
            raise RuntimeError("Import CSV contains no data rows: {0}".format(input_path))

        api = resolve_clone_api(ufs, log)
        reports = execute(api, rows, input_path, timestamp, current_mode, log)

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

        log.write("FINAL STATUS: SUCCESS")

    except Exception as exc:
        if "FINAL STATUS: FAILED" not in log.lines:
            log.write("FINAL STATUS: FAILED")
        log.write(error_text(exc))
        log.write(traceback.format_exc())
        raise

    finally:
        try:
            log_dir = os.path.dirname(input_path) if input_path else io_root()
            if not log_dir:
                log_dir = io_root()
            os.makedirs(log_dir, exist_ok=True)
            log_path = os.path.join(
                log_dir, "J16_RUN_{0}_{1}.log".format(current_mode, timestamp)
            )
            write_log(log_path, log.lines)
        except Exception:
            pass

    return report_path


if __name__ == "__main__":
    main()
