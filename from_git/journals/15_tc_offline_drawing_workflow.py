"""J15 - Teamcenter X offline drawing workflow.

NX X 2506 only.

EXPORT keeps Teamcenter AutoTranslate names, exports model + drawing, marks local
3D/reference .prt files read-only, and writes a SHA-256 manifest.

IMPORT_DRY_RUN / IMPORT_APPLY default every discovered object to UseExisting and
set only the exact drawing to Overwrite.

The UF Clone enum names are resolved from the actual NX X 2506 Python runtime.
This avoids assuming that Python enum member names match generated .NET docs.
"""

import csv
import datetime
import hashlib
import os
import stat
import traceback

import NXOpen
import NXOpen.UF


# ============================================================================
# USER SETTINGS
# ============================================================================
USER_MODE = "EXPORT"  # EXPORT | IMPORT_DRY_RUN | IMPORT_APPLY
USER_SCOPE_CSV = r""   # blank => <I/O root>\NX_TC_OFFLINE_SCOPE.csv
USER_MANIFEST_CSV = r""
# ============================================================================

BUILD = "J15-TCX-OFFLINE-DRAWING-NX2506-V4"
OUT_DIR = "NX_TC_OFFLINE_DRAWINGS"
DEFAULT_SCOPE = "NX_TC_OFFLINE_SCOPE.csv"
MODES = ("EXPORT", "IMPORT_DRY_RUN", "IMPORT_APPLY")

MANIFEST_FIELDS = [
    "RUN_ID", "PART_NUMBER", "REVISION", "DWG_INDEX", "MODEL_IDENTIFIER",
    "DRAWING_IDENTIFIER", "PACKAGE_DIR", "DRAWING_FILE", "EXPORT_LOG",
    "EXPORT_SHA256", "EXPORTED_AT", "REFERENCE_PRT_COUNT", "APPROVED",
    "ENGINEER", "IMPORT_STATUS", "NOTES",
]

REPORT_FIELDS = [
    "RUN_TIMESTAMP", "MODE", "PART_NUMBER", "REVISION", "DWG_INDEX",
    "DRAWING_IDENTIFIER", "DRAWING_FILE", "EXPORTED_SHA256", "CURRENT_SHA256",
    "CHANGED", "APPROVED", "ENGINEER", "DEFAULT_IMPORT_ACTION",
    "DRAWING_IMPORT_ACTION", "DRY_RUN", "RESULT", "MESSAGE", "CLONE_LOG",
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


def mode():
    return upper(env("NX_TC_OFFLINE_MODE") or USER_MODE or "EXPORT")


def scope_path():
    configured = env("NX_TC_OFFLINE_SCOPE_FILE") or clean(USER_SCOPE_CSV)
    if configured:
        return os.path.abspath(os.path.expanduser(configured))
    return os.path.join(io_root(), DEFAULT_SCOPE)


def manifest_path():
    configured = env("NX_TC_OFFLINE_MANIFEST_FILE") or clean(USER_MANIFEST_CSV)
    if configured:
        return os.path.abspath(os.path.expanduser(configured))
    return ""


def error_text(error):
    code = clean(getattr(error, "ErrorCode", ""))
    suffix = ":{0}".format(code) if code else ""
    return "{0}{1} - {2}".format(type(error).__name__, suffix, text(error))


def dispose(value):
    if value is not None:
        try:
            value.Dispose()
        except Exception:
            pass


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


def read_csv(path, required_columns):
    last_decode_error = None
    for encoding in ("utf-8-sig", "utf-8", "cp1252"):
        try:
            with open(path, "r", encoding=encoding, newline="") as handle:
                reader = csv.DictReader(handle)
                headers = [clean(name) for name in (reader.fieldnames or [])]
                missing = [name for name in required_columns if name not in headers]
                if missing:
                    raise RuntimeError(
                        "Missing CSV column(s): {0}".format(", ".join(missing))
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
    raise RuntimeError("Unable to decode CSV: {0}: {1}".format(path, last_decode_error))


def write_csv(path, fields, rows):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields)
        writer.writeheader()
        for row in rows:
            writer.writerow({field: row.get(field, "") for field in fields})


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

    export_op, export_name = resolve_attr(
        operation_type,
        ("ExportOperation", "Export", "OperationExport", "ExportOp"),
        "UF Clone export operation",
    )
    import_op, import_name = resolve_attr(
        operation_type,
        ("ImportOperation", "Import", "OperationImport", "ImportOp"),
        "UF Clone import operation",
    )
    treat_lost, lost_name = resolve_attr(
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
        "UF Clone resolved enums: {0}.{1}, {0}.{2}; {3}.{4}; {5}.{6}; {7}.{8}, {7}.{9}".format(
            operation_type_name, export_name, import_name,
            family_type_name, lost_name,
            naming_type_name, naming_name,
            action_type_name, use_existing_name, overwrite_name,
        )
    )

    return {
        "clone": clone_obj,
        "export_operation": export_op,
        "import_operation": import_op,
        "treat_as_lost": treat_lost,
        "autotranslate": autotranslate,
        "use_existing": use_existing,
        "overwrite": overwrite,
    }


# ---------------------------------------------------------------------------
# Identity and local-file helpers
# ---------------------------------------------------------------------------
def model_id(part_number, revision):
    return "@DB/{0}/{1}".format(part_number, revision)


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
    if name == expected_native(part_number, revision, drawing_index).lower():
        return True
    return "_s_" in name and name.endswith(
        "-{0}-dwg{1}.prt".format(revision, drawing_index).lower()
    )


def find_drawing(folder, part_number, revision, drawing_index):
    exact = os.path.join(folder, expected_native(part_number, revision, drawing_index))
    if os.path.isfile(exact):
        return exact

    matches = []
    for name in os.listdir(folder):
        path = os.path.join(folder, name)
        if os.path.isfile(path) and valid_native(path, part_number, revision, drawing_index):
            matches.append(path)
    if len(matches) == 1:
        return matches[0]
    if not matches:
        return ""
    raise RuntimeError("Multiple native drawing files matched: {0}".format(", ".join(matches)))


def protect_references(folder, drawing):
    target = os.path.normcase(os.path.abspath(drawing))
    count = 0
    for name in os.listdir(folder):
        if not name.lower().endswith(".prt"):
            continue
        path = os.path.join(folder, name)
        if not os.path.isfile(path):
            continue
        current_mode = os.stat(path).st_mode
        if os.path.normcase(os.path.abspath(path)) == target:
            os.chmod(path, current_mode | stat.S_IWRITE)
        else:
            os.chmod(path, current_mode & ~stat.S_IWRITE)
            count += 1
    return count


# ---------------------------------------------------------------------------
# UF Clone call wrappers
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


def setup_export(api, folder, logfile):
    clone = api["clone"]
    clone.Initialise(api["export_operation"])
    clone.SetFamilyTreatment(api["treat_as_lost"])
    clone.SetDefNaming(api["autotranslate"])
    clone.SetDefItemType("")
    clone.SetDefDirectory(folder)
    try:
        clone.SetAssocFileRootDir(folder)
    except Exception:
        pass
    clone.SetDefAction(api["overwrite"])
    clone.SetDefAssocFileCopy(True)
    clone.SetLogfile(logfile)
    try:
        clone.SetCloneRelatedDwgs(False)
    except Exception:
        pass


def export_package(api, folder, model, drawing, logfile, log):
    clone = api["clone"]
    load_status = None
    try:
        terminate(clone)
        setup_export(api, folder, logfile)
        log.write("  Add assembly: {0}".format(model))
        load_status = add_assembly(clone, model)
        log.write("  Add drawing:  {0}".format(drawing))
        clone.AddPart(drawing)
        failures = naming_failures(clone)
        clone.SetDryrun(False)
        try:
            clone.GenerateReport()
        except Exception:
            pass
        perform_clone(clone, failures)
    finally:
        dispose(load_status)
        terminate(clone)


# ---------------------------------------------------------------------------
# Export
# ---------------------------------------------------------------------------
def parse_scope(row):
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


def do_export(api, log):
    path = scope_path()
    log.write("Scope CSV: {0}".format(path))
    if not os.path.isfile(path):
        raise RuntimeError("Scope CSV not found: {0}".format(path))

    rows = read_csv(path, ["PART_NUMBER", "REVISION", "DWG_INDEX"])
    if not rows:
        raise RuntimeError("Scope CSV contains no data rows")

    run_id = stamp()
    root = os.path.join(io_root(), OUT_DIR, run_id)
    os.makedirs(root, exist_ok=True)
    manifest = os.path.join(root, "TCX_OFFLINE_MANIFEST_{0}.csv".format(run_id))

    results = []
    succeeded = 0
    failed = 0

    for source in rows:
        record = {field: "" for field in MANIFEST_FIELDS}
        record.update({"RUN_ID": run_id, "IMPORT_STATUS": "NOT_IMPORTED"})
        try:
            part_number, revision, drawing_index = parse_scope(source)
            model = model_id(part_number, revision)
            drawing = drawing_id(part_number, revision, drawing_index)
            folder = os.path.join(
                root, "{0}_{1}_DWG{2}".format(part_number, revision, drawing_index)
            )
            os.makedirs(folder, exist_ok=True)
            export_log = os.path.join(
                folder,
                "EXPORT_{0}_{1}_DWG{2}.clone".format(part_number, revision, drawing_index),
            )

            record.update({
                "PART_NUMBER": part_number,
                "REVISION": revision,
                "DWG_INDEX": str(drawing_index),
                "MODEL_IDENTIFIER": model,
                "DRAWING_IDENTIFIER": drawing,
                "PACKAGE_DIR": folder,
                "EXPORT_LOG": export_log,
                "EXPORTED_AT": datetime.datetime.now().isoformat(timespec="seconds"),
            })

            log.write("EXPORT {0}/{1}/dwg{2}".format(part_number, revision, drawing_index))
            export_package(api, folder, model, drawing, export_log, log)

            native_drawing = find_drawing(folder, part_number, revision, drawing_index)
            if not native_drawing:
                raise RuntimeError(
                    "Expected native drawing not found: {0}".format(
                        expected_native(part_number, revision, drawing_index)
                    )
                )

            record["DRAWING_FILE"] = native_drawing
            record["EXPORT_SHA256"] = sha256(native_drawing)
            record["REFERENCE_PRT_COUNT"] = str(protect_references(folder, native_drawing))
            record["NOTES"] = "Export OK; all non-drawing .prt files set read-only"
            succeeded += 1
            log.write("  Drawing: {0}".format(os.path.basename(native_drawing)))

        except Exception as error:
            failed += 1
            record["IMPORT_STATUS"] = "EXPORT_FAILED"
            record["NOTES"] = error_text(error)
            log.write("  FAILED: {0}".format(error_text(error)))
            log.write(traceback.format_exc())

        results.append(record)
        write_csv(manifest, MANIFEST_FIELDS, results)

    log.write("Manifest: {0}".format(manifest))
    log.write("Export summary: {0} succeeded, {1} failed".format(succeeded, failed))
    return manifest, succeeded, failed


# ---------------------------------------------------------------------------
# Import
# ---------------------------------------------------------------------------
def validate_target(row):
    part_number = clean(row.get("PART_NUMBER"))
    revision = clean(row.get("REVISION"))
    try:
        drawing_index = int(clean(row.get("DWG_INDEX")))
    except Exception:
        raise RuntimeError("Invalid DWG_INDEX in manifest")

    path = clean(row.get("DRAWING_FILE"))
    identifier = clean(row.get("DRAWING_IDENTIFIER"))
    expected_identifier = drawing_id(part_number, revision, drawing_index)

    if identifier.upper() != expected_identifier.upper() or "/SPECIFICATION/" not in identifier.upper():
        raise RuntimeError("Manifest drawing identity is not the expected /specification/ target")
    if not os.path.isfile(path):
        raise RuntimeError("DRAWING_FILE not found: {0}".format(path))
    if not valid_native(path, part_number, revision, drawing_index):
        raise RuntimeError("Native drawing was renamed or does not match Teamcenter AutoTranslate naming")
    return part_number, revision, drawing_index, path, identifier


def same_part(candidate, target):
    if not clean(candidate):
        return False
    try:
        if os.path.normcase(os.path.abspath(candidate)) == os.path.normcase(os.path.abspath(target)):
            return True
    except Exception:
        pass
    return os.path.basename(candidate).lower() == os.path.basename(target).lower()


def import_one(api, drawing, folder, logfile, dry_run, log):
    clone = api["clone"]
    load_status = None
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
        perform_clone(clone, failures)
        log.write("  Default=UseExisting; drawing=Overwrite; dry_run={0}".format(dry_run))

    finally:
        dispose(load_status)
        terminate(clone)


def report_row(row, current_mode, timestamp):
    return {
        "RUN_TIMESTAMP": timestamp,
        "MODE": current_mode,
        "PART_NUMBER": row.get("PART_NUMBER", ""),
        "REVISION": row.get("REVISION", ""),
        "DWG_INDEX": row.get("DWG_INDEX", ""),
        "DRAWING_IDENTIFIER": row.get("DRAWING_IDENTIFIER", ""),
        "DRAWING_FILE": row.get("DRAWING_FILE", ""),
        "EXPORTED_SHA256": row.get("EXPORT_SHA256", ""),
        "CURRENT_SHA256": "",
        "CHANGED": "",
        "APPROVED": row.get("APPROVED", ""),
        "ENGINEER": row.get("ENGINEER", ""),
        "DEFAULT_IMPORT_ACTION": "UseExisting",
        "DRAWING_IMPORT_ACTION": "Overwrite",
        "DRY_RUN": "YES" if current_mode == "IMPORT_DRY_RUN" else "NO",
        "RESULT": "",
        "MESSAGE": "",
        "CLONE_LOG": "",
    }


def do_import(api, log, current_mode):
    path = manifest_path()
    log.write("Manifest CSV: {0}".format(path or "<not set>"))
    if not path or not os.path.isfile(path):
        raise RuntimeError("Set USER_MANIFEST_CSV/NX_TC_OFFLINE_MANIFEST_FILE to a valid manifest")

    rows = read_csv(path, [
        "PART_NUMBER", "REVISION", "DWG_INDEX", "DRAWING_IDENTIFIER",
        "DRAWING_FILE", "EXPORT_SHA256", "APPROVED", "ENGINEER",
    ])

    timestamp = stamp()
    report = os.path.join(
        os.path.dirname(path), "TCX_OFFLINE_{0}_{1}.csv".format(current_mode, timestamp)
    )
    results = []
    succeeded = 0
    failed = 0

    for row in rows:
        result = report_row(row, current_mode, timestamp)
        results.append(result)
        try:
            part_number, revision, drawing_index, drawing, _identifier = validate_target(row)
            baseline = clean(row.get("EXPORT_SHA256"))
            if not baseline:
                raise RuntimeError("EXPORT_SHA256 is blank")

            current = sha256(drawing)
            result["CURRENT_SHA256"] = current
            changed = current.lower() != baseline.lower()
            result["CHANGED"] = "YES" if changed else "NO"
            log.write("IMPORT {0}/{1}/dwg{2} changed={3}".format(
                part_number, revision, drawing_index, changed
            ))

            if not changed:
                result["RESULT"] = "SKIPPED_UNCHANGED"
                result["MESSAGE"] = "SHA-256 matches export snapshot"
                write_csv(report, REPORT_FIELDS, results)
                continue

            if current_mode == "IMPORT_APPLY" and upper(row.get("APPROVED")) != "YES":
                failed += 1
                result["RESULT"] = "BLOCKED_NOT_APPROVED"
                result["MESSAGE"] = "IMPORT_APPLY requires APPROVED=YES"
                write_csv(report, REPORT_FIELDS, results)
                continue

            if current_mode == "IMPORT_APPLY" and not clean(row.get("ENGINEER")):
                failed += 1
                result["RESULT"] = "BLOCKED_ENGINEER_REQUIRED"
                result["MESSAGE"] = "IMPORT_APPLY requires ENGINEER"
                write_csv(report, REPORT_FIELDS, results)
                continue

            import_log = os.path.join(
                os.path.dirname(drawing),
                "IMPORT_{0}_{1}_{2}_DWG{3}.clone".format(
                    current_mode, part_number, revision, drawing_index
                ),
            )
            result["CLONE_LOG"] = import_log
            import_one(
                api, drawing, os.path.dirname(drawing), import_log,
                current_mode == "IMPORT_DRY_RUN", log,
            )
            result["RESULT"] = "DRY_RUN_OK" if current_mode == "IMPORT_DRY_RUN" else "IMPORT_APPLIED"
            result["MESSAGE"] = "UF Clone completed: default UseExisting, exact drawing Overwrite"
            succeeded += 1

        except Exception as error:
            failed += 1
            result["RESULT"] = "FAILED"
            result["MESSAGE"] = error_text(error)
            log.write("  FAILED: {0}".format(error_text(error)))
            log.write(traceback.format_exc())
            write_csv(report, REPORT_FIELDS, results)
            if current_mode == "IMPORT_APPLY":
                break

        write_csv(report, REPORT_FIELDS, results)

    log.write("Import report: {0}".format(report))
    log.write("Import summary: {0} succeeded/validated, {1} failed/blocked".format(succeeded, failed))
    return report, succeeded, failed


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    session = NXOpen.Session.GetSession()
    ufs = NXOpen.UF.UFSession.GetUFSession()
    log = Log(session)
    current_mode = mode()

    log.write("=" * 72)
    log.write("J15 TEAMCENTER X OFFLINE DRAWING WORKFLOW")
    log.write("Build: {0} | Mode: {1} | I/O: {2}".format(BUILD, current_mode, io_root()))
    log.write("Runtime target: NX X 2506 only")
    log.write("=" * 72)

    try:
        if current_mode not in MODES:
            raise RuntimeError("Invalid USER_MODE: {0}".format(current_mode))

        api = resolve_clone_api(ufs, log)

        if current_mode == "EXPORT":
            output, succeeded, failed = do_export(api, log)
        else:
            output, succeeded, failed = do_import(api, log, current_mode)

        if failed:
            log.write("FINAL STATUS: FAILED")
            log.write("Primary output: {0}".format(output))
            raise RuntimeError(
                "J15 completed with {0} failed/blocked row(s). Review: {1}".format(
                    failed, output
                )
            )

        log.write("FINAL STATUS: SUCCESS")
        log.write("Primary output: {0}".format(output))

    except Exception as error:
        if "FINAL STATUS: FAILED" not in log.lines:
            log.write("FINAL STATUS: FAILED")
        log.write(error_text(error))
        log.write(traceback.format_exc())
        raise

    finally:
        try:
            root = os.path.join(io_root(), OUT_DIR)
            os.makedirs(root, exist_ok=True)
            write_log(
                os.path.join(root, "J15_{0}_{1}.txt".format(current_mode, stamp())),
                log.lines,
            )
        except Exception:
            pass


if __name__ == "__main__":
    main()


def GetUnloadOption(dummy):
    return NXOpen.Session.LibraryUnloadOption.Immediately
