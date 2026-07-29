"""J16 - Teamcenter X specification drawing import using UF Clone.

NX X 2506 managed mode only.

Direction implemented:
- The exact drawing specification must be fully checked in before any write.
  Checkout by the current user or any other user blocks the whole approved batch.
- UF Clone remains the proven specification import method.
- A completed UF Clone apply call without an exception is reported as
  IMPORT_COMPLETED. J16 does not reopen, export, hash, or otherwise verify the
  managed drawing after import.
- The master 3D and all discovered references default to UseExisting. Only the
  exact local drawing is assigned Overwrite.
- This file is self-contained. J17 compatibility helpers remain public here.

Run DRY_RUN first. APPLY_APPROVED writes only APPROVED=YES rows with ENGINEER.
"""

import csv
import datetime
import hashlib
import os
import re
import traceback
from collections import Counter

import NXOpen
import NXOpen.UF


# ============================================================================
# USER SETTINGS
# ============================================================================
USER_IMPORT_CSV = r""  # blank => <I/O root>\NX_TC_DRAWING_IMPORT.csv
USER_MODE = "DRY_RUN"  # DRY_RUN | APPLY_APPROVED
# Optional environment overrides:
#   NX_TC_DRAWING_IMPORT_FILE=<full CSV path>
#   NX_J16_MODE=DRY_RUN or APPLY_APPROVED
# ============================================================================

BUILD = "J16-TCX-SPECIFICATION-UFCLONE-NX2506-V5"
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
    "CHECKOUT_PREFLIGHT",
    "CHECKOUT_RECHECK",
    "CLONE_PREFLIGHT",
    "POST_IMPORT_VERIFICATION",
    "WRITE_ATTEMPTED",
    "RESULT",
    "MESSAGE",
    "CLONE_LOG",
]

# Kept public for J17 compatibility. J16 V5 does not use clone-log text as
# persistence proof.
TARGET_BLOCK_TERMS = (
    "checked out by", "checked-out by", "is checked out", "already checked out",
    "checkout conflict", "reserved by", "is reserved", "reservation conflict",
    "locked by", "is locked", "read-only", "read only", "no write access",
    "write access denied", "write denied", "not writable", "cannot be written",
    "unable to write", "permission denied", "access denied", "cannot overwrite",
    "could not overwrite", "unable to overwrite", "not overwritten",
    "failed to overwrite", "cannot import", "could not import", "unable to import",
    "not imported", "failed to import", "cannot modify", "unable to modify",
    "not modifiable", "cannot be replaced", "unable to replace",
    "not allowed to modify", "write failed",
)
GLOBAL_FAILURE_TERMS = (
    "fatal error", "operation failed", "clone failed", "import failed",
    "errors occurred", "error occurred",
)
TARGET_SUCCESS_TERMS = (
    "successfully overwritten", "was overwritten", "has been overwritten",
    "successfully imported", "was imported", "has been imported",
    "successfully replaced", "was replaced", "has been replaced",
    "successfully updated", "was updated", "has been updated",
    "import completed successfully", "clone completed successfully",
    "operation completed successfully", "overwrite successful", "import successful",
    "successfully completed",
)
NEGATED_FAILURE_PATTERNS = (
    re.compile(r"\b0\s+errors?\b", re.IGNORECASE),
    re.compile(r"\bno\s+errors?\b", re.IGNORECASE),
    re.compile(r"\b0\s+failures?\b", re.IGNORECASE),
    re.compile(r"\bno\s+failures?\b", re.IGNORECASE),
)


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
    if value is None:
        return
    for name in ("Dispose", "Destroy", "FreeResource"):
        method = getattr(value, name, None)
        if method is not None:
            try:
                method()
                return
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


def matching_native_drawings(folder, part_number, revision, drawing_index):
    if not folder or not os.path.isdir(folder):
        return []
    matches = []
    for name in os.listdir(folder):
        path = os.path.join(folder, name)
        if os.path.isfile(path) and valid_native(
            path, part_number, revision, drawing_index
        ):
            matches.append(os.path.abspath(path))
    unique = {
        os.path.normcase(os.path.abspath(path)): os.path.abspath(path)
        for path in matches
    }
    return sorted(unique.values(), key=lambda value: value.lower())


def resolve_drawing_file(csv_path, supplied_value, part_number, revision, drawing_index):
    requested = resolve_local_path(csv_path, supplied_value)
    if requested and os.path.isfile(requested):
        return requested, "EXACT", []
    if not clean(supplied_value):
        return requested, "NOT_FOUND", []
    folder = os.path.dirname(requested) if requested else os.path.dirname(csv_path)
    matches = matching_native_drawings(folder, part_number, revision, drawing_index)
    if len(matches) == 1:
        return matches[0], "AUTO_RESOLVED", matches
    if len(matches) > 1:
        return requested, "MULTIPLE", matches
    return requested, "NOT_FOUND", []


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
        "CHECKOUT_PREFLIGHT": "NOT_RUN",
        "CHECKOUT_RECHECK": "NOT_RUN",
        "CLONE_PREFLIGHT": "NOT_RUN",
        "POST_IMPORT_VERIFICATION": "NOT_REQUESTED",
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
        if detail not in report["MESSAGE"]:
            report["MESSAGE"] = "{0} | {1}".format(report["MESSAGE"], detail)
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
            set_error(report, "ERROR_APPROVAL_VALUE", "APPROVED must be YES, NO, or blank.")
            continue
        if mode == "APPLY_APPROVED" and approval != "YES":
            report["RESULT"] = "NOT_APPROVED"
            report["MESSAGE"] = "No write authorized for this row."
            continue
        if mode == "APPLY_APPROVED" and not clean(row.get("ENGINEER")):
            set_error(report, "ERROR_ENGINEER_REQUIRED", "ENGINEER is required when APPROVED=YES.")
            continue
        try:
            part_number, revision, drawing_index = parse_target(row)
            key = (upper(part_number), upper(revision), drawing_index)
            if key in duplicate_keys:
                raise RuntimeError(
                    "The same PART_NUMBER/REVISION/DWG_INDEX appears more than once."
                )
            identifier = drawing_id(part_number, revision, drawing_index)
            supplied_identifier = clean(row.get("DRAWING_IDENTIFIER"))
            if supplied_identifier and upper(supplied_identifier) != upper(identifier):
                raise RuntimeError(
                    "DRAWING_IDENTIFIER does not match PART_NUMBER/REVISION/DWG_INDEX."
                )
            drawing, resolution, matches = resolve_drawing_file(
                csv_path, row.get("DRAWING_FILE"), part_number, revision, drawing_index
            )
            if resolution == "MULTIPLE":
                raise RuntimeError(
                    "More than one valid AutoTranslate drawing matched: {0}".format(
                        " | ".join(matches)
                    )
                )
            if not drawing or not os.path.isfile(drawing):
                raise RuntimeError(
                    "DRAWING_FILE was not found: {0}".format(drawing or "<blank>")
                )
            if not drawing.lower().endswith(".prt"):
                raise RuntimeError("DRAWING_FILE must be a native NX .prt file.")
            if not valid_native(drawing, part_number, revision, drawing_index):
                raise RuntimeError(
                    "DRAWING_FILE does not match the Teamcenter AutoTranslate identity."
                )
            current_sha = sha256(drawing)
            report.update(
                DRAWING_IDENTIFIER=identifier,
                DRAWING_FILE=drawing,
                PREFLIGHT_SHA256=current_sha,
            )
            baseline = clean(row.get("EXPORT_SHA256"))
            if baseline:
                changed = current_sha.lower() != baseline.lower()
                report["CHANGED_FROM_BASELINE"] = "YES" if changed else "NO"
                if not changed:
                    report["RESULT"] = "SKIPPED_UNCHANGED"
                    report["MESSAGE"] = "Local drawing still matches EXPORT_SHA256; no import required."
                    continue
            else:
                report["CHANGED_FROM_BASELINE"] = "UNKNOWN"
            report["RESULT"] = "LOCAL_PREFLIGHT_OK"
            report["MESSAGE"] = (
                "Local identity checks passed."
                + (" Shortened filename was auto-resolved." if resolution == "AUTO_RESOLVED" else "")
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
        except Exception as exc:
            set_error(report, "ERROR_LOCAL_PREFLIGHT", "Local safety preflight failed.", exc)
    return reports, proposals


def apply_blocking_error(report):
    if report.get("RESULT") == "ERROR_APPROVAL_VALUE":
        return True
    if upper(report.get("APPROVED")) != "YES":
        return False
    return report.get("RESULT", "").startswith(("ERROR_", "FAILED_"))


def normalized_name(value):
    return "".join(ch.lower() for ch in clean(value) if ch.isalnum())


def public_names(value):
    names = []
    try:
        member_map = getattr(value, "__members__", None)
        if member_map:
            names.extend(list(member_map.keys()))
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
        clone_type, ("OperationClass", "Operation", "OperationType"),
        "UF Clone operation enum type",
    )
    family_type, family_type_name = resolve_attr(
        clone_type, ("FamilyTreatment", "Family", "FamilyTreatmentType"),
        "UF Clone family-treatment enum type",
    )
    naming_type, naming_type_name = resolve_attr(
        clone_type, ("NamingTechnique", "Naming", "NamingType"),
        "UF Clone naming enum type",
    )
    action_type, action_type_name = resolve_attr(
        clone_type, ("Action", "CloneAction", "ActionType"),
        "UF Clone action enum type",
    )
    import_operation, import_name = resolve_attr(
        operation_type, ("ImportOperation", "Import", "OperationImport", "ImportOp"),
        "UF Clone import operation",
    )
    treat_as_lost, lost_name = resolve_attr(
        family_type, ("TreatAsLost", "AsLost", "Lost", "TreatLost"),
        "UF Clone TreatAsLost family treatment",
    )
    autotranslate, naming_name = resolve_attr(
        naming_type, ("Autotranslate", "AutoTranslate", "Auto_Translate", "AutomaticTranslate"),
        "UF Clone AutoTranslate naming technique",
    )
    use_existing, use_existing_name = resolve_attr(
        action_type, ("UseExisting", "UseExistingPart", "Existing", "UseExistingItem"),
        "UF Clone UseExisting action",
    )
    overwrite, overwrite_name = resolve_attr(
        action_type, ("Overwrite", "OverWrite", "Replace", "OverwriteExisting"),
        "UF Clone Overwrite action",
    )
    log.write("UF Clone binding: NXOpen.UF.Clone")
    log.write(
        "UF Clone resolved enums: {0}.{1}; {2}.{3}; {4}.{5}; {6}.{7}, {6}.{8}".format(
            operation_type_name, import_name, family_type_name, lost_name,
            naming_type_name, naming_name, action_type_name,
            use_existing_name, overwrite_name,
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


def read_text_file(path):
    if not path or not os.path.isfile(path):
        return "", ""
    last_error = None
    for encoding in ("utf-8-sig", "utf-8", "utf-16", "utf-16-le", "cp1252"):
        try:
            with open(path, "r", encoding=encoding) as handle:
                return handle.read(), encoding
        except (UnicodeDecodeError, UnicodeError) as exc:
            last_error = exc
    if last_error is not None:
        raise RuntimeError("Unable to decode clone log: {0}".format(path))
    return "", ""


def compact_line(value, limit=500):
    value = " ".join(clean(value).split())
    return value if len(value) <= limit else value[: limit - 3] + "..."


def is_negated_failure(line):
    return any(pattern.search(line) for pattern in NEGATED_FAILURE_PATTERNS)


def object_tag(value):
    try:
        return int(value.Tag)
    except Exception:
        try:
            return value.Tag
        except Exception:
            return None


def same_nx_object(left, right):
    if left is right:
        return True
    left_tag = object_tag(left)
    right_tag = object_tag(right)
    return left_tag is not None and right_tag is not None and left_tag == right_tag


def checkedout_arrays(pdm):
    method = getattr(pdm, "GetCheckedoutStatusOfAllObjectsInSession", None)
    if method is None:
        raise RuntimeError(
            "PdmSession.GetCheckedoutStatusOfAllObjectsInSession is unavailable."
        )
    checked_output = []
    checkedin_output = []
    try:
        raw = method()
    except TypeError:
        raw = method(checked_output, checkedin_output)
    if isinstance(raw, (tuple, list)) and len(raw) >= 2:
        return list(raw[0] or []), list(raw[1] or [])
    if checked_output or checkedin_output:
        return list(checked_output), list(checkedin_output)
    raise RuntimeError(
        "Unexpected checkout-status return: {0}".format(type(raw).__name__)
    )


def close_opened_part(part, log):
    if part is None:
        return
    try:
        whole_tree, _ = resolve_attr(
            NXOpen.BasePart.CloseWholeTree,
            ("False_", "False", "CloseWholeTreeFalse"),
            "BasePart.CloseWholeTree false value",
        )
        close_modified, _ = resolve_attr(
            NXOpen.BasePart.CloseModified,
            ("UseResponses", "UseLatest", "CloseModified"),
            "BasePart.CloseModified safe value",
        )
        part.Close(whole_tree, close_modified, None)
    except Exception as exc:
        log.write(
            "  WARNING: could not close checkout-probe drawing: {0}".format(
                error_text(exc)
            )
        )


def check_target_checkout(session, pdm, identifier, log):
    """Return CHECKED_IN only when the exact drawing is proven not checked out."""
    try:
        existing = session.Parts.FindObject(identifier)
    except Exception:
        existing = None
    part = existing
    load_status = None
    opened_here = False
    try:
        if part is None:
            part, load_status = session.Parts.OpenBase(identifier)
            opened_here = True
        dispose(load_status)
        load_status = None
        checked, checkedin = checkedout_arrays(pdm)
        if any(same_nx_object(part, item) for item in checked):
            return "CHECKED_OUT"
        if any(same_nx_object(part, item) for item in checkedin):
            return "CHECKED_IN"
        return "UNKNOWN"
    finally:
        dispose(load_status)
        if opened_here:
            close_opened_part(part, log)


def clone_log_path(proposal, mode, phase):
    return os.path.join(
        os.path.dirname(proposal["drawing"]),
        "J16_{0}_{1}_{2}_{3}_DWG{4}.clone".format(
            phase, mode, proposal["part_number"], proposal["revision"],
            proposal["drawing_index"],
        ),
    )


def import_one(api, drawing, logfile, dry_run, log):
    """Execute one UF Clone operation. No post-import verification is performed."""
    clone = api["clone"]
    load_status = None
    discovered_parts = []
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
                try:
                    set_action(clone, part_name, api["use_existing"])
                except Exception:
                    pass
        if not drawing_action_set:
            set_action(clone, drawing, api["overwrite"])
        failures = naming_failures(clone)
        clone.SetDryrun(bool(dry_run))
        try:
            clone.GenerateReport()
        except Exception:
            pass
        raw_result = perform_clone(clone, failures)
        log.write(
            "  UF Clone completed; returned={0}; discovered={1}; "
            "default=UseExisting; drawing=Overwrite; dry_run={2}".format(
                text(raw_result), len(discovered_parts), bool(dry_run)
            )
        )
        return raw_result, discovered_parts
    finally:
        dispose(load_status)
        terminate(clone)


def mark_batch(proposals, result, message, eligible_results):
    for proposal in proposals:
        report = proposal["report"]
        if report.get("RESULT") in eligible_results:
            report["RESULT"] = result
            report["MESSAGE"] = message


def run_checkout_preflight(session, pdm, proposals, log):
    failed = False
    for proposal in proposals:
        report = proposal["report"]
        try:
            status = check_target_checkout(
                session, pdm, proposal["identifier"], log
            )
            report["CHECKOUT_PREFLIGHT"] = status
            if status == "CHECKED_OUT":
                raise RuntimeError(
                    "The exact drawing specification is checked out. It must be checked in, "
                    "including when the checkout belongs to the current user."
                )
            if status != "CHECKED_IN":
                raise RuntimeError(
                    "The exact drawing specification could not be proven checked in."
                )
            report["RESULT"] = "CHECKOUT_PREFLIGHT_OK"
            report["MESSAGE"] = "Exact drawing specification is checked in."
        except Exception as exc:
            failed = True
            set_error(
                report,
                "FAILED_CHECKOUT_PREFLIGHT",
                "Target drawing is not proven fully checked in.",
                exc,
            )
    return failed


def run_clone_preflight(api, proposals, log, mode):
    failed = False
    for proposal in proposals:
        report = proposal["report"]
        if report.get("RESULT") != "CHECKOUT_PREFLIGHT_OK":
            continue
        logfile = clone_log_path(proposal, mode, "PREFLIGHT")
        report["CLONE_LOG"] = logfile
        try:
            import_one(api, proposal["drawing"], logfile, True, log)
            report["CLONE_PREFLIGHT"] = "PASS"
            report["RESULT"] = "DRY_RUN_OK" if mode == "DRY_RUN" else "CLONE_PREFLIGHT_OK"
            report["MESSAGE"] = (
                "Checkout is clear and UF Clone dry run completed. No write was attempted."
            )
        except Exception as exc:
            failed = True
            report["CLONE_PREFLIGHT"] = "FAIL"
            set_error(
                report,
                "FAILED_CLONE_PREFLIGHT",
                "UF Clone dry run raised an exception.",
                exc,
            )
            log.write("  FAILED dry run: {0}".format(error_text(exc)))
            log.write(traceback.format_exc())
    return failed


def execute(session, pdm, api, rows, csv_path, timestamp, mode, log):
    reports, proposals = local_preflight(rows, csv_path, timestamp, mode)
    if mode == "APPLY_APPROVED":
        blocking = [report for report in reports if apply_blocking_error(report)]
        if blocking:
            mark_batch(
                proposals,
                "BATCH_ABORTED_LOCAL_PREFLIGHT",
                "At least one approved row failed local preflight. No Teamcenter write was attempted.",
                ("LOCAL_PREFLIGHT_OK",),
            )
            return reports

    checkout_failed = run_checkout_preflight(session, pdm, proposals, log)
    if mode == "APPLY_APPROVED" and checkout_failed:
        mark_batch(
            proposals,
            "BATCH_ABORTED_CHECKOUT_PREFLIGHT",
            "At least one approved drawing is checked out or not proven checked in. "
            "The whole batch was stopped before any write.",
            ("CHECKOUT_PREFLIGHT_OK",),
        )
        return reports

    clone_failed = run_clone_preflight(api, proposals, log, mode)
    if mode == "DRY_RUN":
        return reports
    if clone_failed:
        mark_batch(
            proposals,
            "BATCH_ABORTED_CLONE_PREFLIGHT",
            "UF Clone dry run failed for at least one approved drawing. No Teamcenter write was attempted.",
            ("CLONE_PREFLIGHT_OK",),
        )
        return reports

    writable = [
        proposal for proposal in proposals
        if proposal["report"].get("RESULT") == "CLONE_PREFLIGHT_OK"
    ]
    for index, proposal in enumerate(writable):
        report = proposal["report"]
        try:
            current_sha = sha256(proposal["drawing"])
            if current_sha.lower() != proposal["preflight_sha"].lower():
                raise RuntimeError("DRAWING_FILE changed after preflight.")
            checkout = check_target_checkout(
                session, pdm, proposal["identifier"], log
            )
            report["CHECKOUT_RECHECK"] = checkout
            if checkout == "CHECKED_OUT":
                raise RuntimeError(
                    "The exact drawing became checked out before import."
                )
            if checkout != "CHECKED_IN":
                raise RuntimeError(
                    "The exact drawing could not be proven checked in immediately before import."
                )
            logfile = clone_log_path(proposal, mode, "APPLY")
            report["CLONE_LOG"] = logfile
            report["WRITE_ATTEMPTED"] = "YES"
            import_one(api, proposal["drawing"], logfile, False, log)
            report["RESULT"] = "IMPORT_COMPLETED"
            report["POST_IMPORT_VERIFICATION"] = "NOT_REQUESTED"
            report["MESSAGE"] = (
                "UF Clone apply completed without an exception. No post-import open, "
                "export, hash comparison, or persistence verification was requested."
            )
        except Exception as exc:
            before_write = report.get("WRITE_ATTEMPTED") == "NO"
            set_error(
                report,
                "FAILED_BEFORE_WRITE" if before_write else "FAILED_IMPORT_APPLY",
                "Specification import stopped.",
                exc,
            )
            log.write("  STOPPED {0}: {1}".format(
                proposal["identifier"], error_text(exc)
            ))
            log.write(traceback.format_exc())
            mark_batch(
                writable[index + 1 :],
                "BATCH_STOPPED_AFTER_FAILURE",
                "A previous drawing failed or became checked out. No write was attempted for this row.",
                ("CLONE_PREFLIGHT_OK",),
            )
            break
    return reports


def summary_counts(reports):
    return Counter(report.get("RESULT", "") or "<blank>" for report in reports)


def has_failure(reports, mode):
    prefixes = ("ERROR_", "FAILED_", "BATCH_ABORTED_", "BATCH_STOPPED_")
    for report in reports:
        result = report.get("RESULT", "")
        if not result.startswith(prefixes):
            continue
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
    mode = configured_mode()
    input_path = configured_input_path()
    timestamp = stamp()
    log.write("=" * 72)
    log.write("J16 TEAMCENTER X SPECIFICATION DRAWING IMPORT")
    log.write("Build: {0} | Mode: {1}".format(BUILD, mode))
    log.write("Import method: UF Clone")
    log.write("Checkout rule: target drawing must be CHECKED_IN; own checkout also blocks")
    log.write("Post-import verification: NOT REQUESTED")
    log.write("Input: {0}".format(input_path))
    log.write("=" * 72)
    report_path = ""
    try:
        if mode not in VALID_MODES:
            raise RuntimeError(
                "USER_MODE/NX_J16_MODE must be DRY_RUN or APPLY_APPROVED."
            )
        if not os.path.isfile(input_path):
            raise RuntimeError("Import CSV not found: {0}".format(input_path))
        rows = read_csv(input_path)
        if not rows:
            raise RuntimeError("Import CSV contains no data rows: {0}".format(input_path))
        pdm = getattr(session, "PdmSession", None)
        if pdm is None:
            raise RuntimeError(
                "NXOpen.Session.PdmSession is unavailable. Run J16 in Teamcenter managed mode."
            )
        api = resolve_clone_api(ufs, log)
        reports = execute(
            session, pdm, api, rows, input_path, timestamp, mode, log
        )
        report_path = os.path.join(
            os.path.dirname(input_path),
            "J16_{0}_{1}.csv".format(mode, timestamp),
        )
        write_csv(report_path, reports)
        log.write("Report: {0}".format(report_path))
        for result, count in sorted(summary_counts(reports).items()):
            log.write("  {0}: {1}".format(result, count))
        if has_failure(reports, mode):
            log.write("FINAL STATUS: FAILED")
            log.write(
                "Failures are recorded in the CSV report; handled row failures do not raise an NX prompt."
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
        try:
            log_dir = os.path.dirname(input_path) if input_path else io_root()
            if not log_dir:
                log_dir = io_root()
            os.makedirs(log_dir, exist_ok=True)
            write_log(
                os.path.join(
                    log_dir, "J16_RUN_{0}_{1}.log".format(mode, timestamp)
                ),
                log.lines,
            )
        except Exception:
            pass
    return report_path


if __name__ == "__main__":
    main()
