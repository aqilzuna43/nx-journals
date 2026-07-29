"""J17 - Teamcenter X master-drawing import with 3D preservation.

NX X 2506 only.

J17 stages a local NX drawing as the master file of a SEPARATE drawing item.
It never overwrites the preserved 3D Item/Revision. The required 3D reference
must be discovered and is forced to UseExisting; only the staged master drawing
is Overwrite. Afterward, create the final drawing by Teamcenter specification.

Run DRY_RUN first. APPLY_APPROVED writes only APPROVED=YES rows with ENGINEER.
J16 V2 must remain beside this file because J17 reuses its proven NX2506 UF
Clone binding and call wrappers.
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


USER_IMPORT_CSV = r""  # blank => <I/O root>\NX_TC_MASTER_DRAWING_IMPORT.csv
USER_MODE = "DRY_RUN"  # DRY_RUN | APPLY_APPROVED

BUILD = "J17-TCX-MASTER-DRAWING-IMPORT-NX2506-V1"
DEFAULT_INPUT = "NX_TC_MASTER_DRAWING_IMPORT.csv"
VALID_MODES = ("DRY_RUN", "APPLY_APPROVED")
REQUIRED_COLUMNS = (
    "MASTER_PART_NUMBER",
    "MASTER_REVISION",
    "SOURCE_DRAWING_FILE",
    "PRESERVE_3D_PART_NUMBER",
    "PRESERVE_3D_REVISION",
    "APPROVED",
    "ENGINEER",
)
REPORT_COLUMNS = (
    "RUN_TIMESTAMP", "MODE", "CSV_ROW",
    "MASTER_PART_NUMBER", "MASTER_REVISION", "MASTER_IDENTIFIER",
    "SOURCE_DRAWING_FILE", "SOURCE_SHA256",
    "STAGED_MASTER_FILE", "STAGED_SHA256",
    "PRESERVE_3D_PART_NUMBER", "PRESERVE_3D_REVISION",
    "PRESERVE_3D_IDENTIFIER", "PRESERVE_3D_DISCOVERED",
    "PRESERVE_3D_DISCOVERED_NAME",
    "DEFAULT_IMPORT_ACTION", "MASTER_DRAWING_ACTION", "PRESERVE_3D_ACTION",
    "CLONE_PREFLIGHT", "CLONE_LOG_STATUS", "TARGET_RESERVATION_STATUS",
    "POST_IMPORT_VERIFICATION", "WRITE_ATTEMPTED",
    "APPROVED", "ENGINEER", "RESULT", "MESSAGE",
    "CLONE_LOG", "CLONE_LOG_EVIDENCE",
)


def load_j16():
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                        "16_tc_offline_drawing_import.py")
    if not os.path.isfile(path):
        raise RuntimeError("J16 dependency not found beside J17: {0}".format(path))
    spec = importlib.util.spec_from_file_location("nx_journal_16", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    required = (
        "Log", "clean", "upper", "env", "stamp", "io_root", "error_text",
        "dispose", "sha256", "resolve_clone_api", "terminate", "add_assembly",
        "naming_failures", "perform_clone", "set_action", "iterate_parts",
        "same_part", "read_text_file", "compact_line",
        "TARGET_BLOCK_TERMS", "GLOBAL_FAILURE_TERMS", "TARGET_SUCCESS_TERMS",
        "is_negated_failure",
    )
    missing = [name for name in required if not hasattr(module, name)]
    if missing:
        raise RuntimeError("J16 is incompatible; missing: {0}".format(", ".join(missing)))
    return module


J16 = load_j16()
clean = J16.clean
upper = J16.upper


def configured_mode():
    return upper(J16.env("NX_J17_MODE") or USER_MODE or "DRY_RUN")


def configured_input_path():
    value = J16.env("NX_TC_MASTER_DRAWING_IMPORT_FILE") or clean(USER_IMPORT_CSV)
    if value:
        return os.path.abspath(os.path.expanduser(value))
    return os.path.join(J16.io_root(), DEFAULT_INPUT)


def master_id(part_number, revision):
    return "@DB/{0}/{1}".format(part_number, revision)


def expected_master_native(part_number, revision):
    return "{0}_{1}_m.prt".format(part_number, revision)


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
                    raise RuntimeError("Missing CSV column(s): {0}".format(", ".join(missing)))
                rows = []
                for number, source in enumerate(reader, 2):
                    row = {clean(k): clean(v) for k, v in source.items() if k is not None}
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


def parse_row(row):
    target_pn = clean(row.get("MASTER_PART_NUMBER"))
    target_rev = clean(row.get("MASTER_REVISION"))
    model_pn = clean(row.get("PRESERVE_3D_PART_NUMBER"))
    model_rev = clean(row.get("PRESERVE_3D_REVISION"))
    if not target_pn or not target_rev:
        raise RuntimeError("MASTER_PART_NUMBER and MASTER_REVISION are required")
    if not model_pn or not model_rev:
        raise RuntimeError("PRESERVE_3D_PART_NUMBER and PRESERVE_3D_REVISION are required")
    if upper(target_pn) == upper(model_pn) and upper(target_rev) == upper(model_rev):
        raise RuntimeError(
            "The drawing-master target cannot equal the preserved 3D Item/Revision. "
            "J17 refuses to overwrite the 3D master file."
        )
    return target_pn, target_rev, model_pn, model_rev


def approval(row):
    value = upper(row.get("APPROVED"))
    return value if value in ("YES", "NO") else ("NO" if not value else "INVALID")


def base_report(row, timestamp, mode):
    return {
        "RUN_TIMESTAMP": timestamp, "MODE": mode, "CSV_ROW": row.get("_CSV_ROW", ""),
        "MASTER_PART_NUMBER": row.get("MASTER_PART_NUMBER", ""),
        "MASTER_REVISION": row.get("MASTER_REVISION", ""), "MASTER_IDENTIFIER": "",
        "SOURCE_DRAWING_FILE": row.get("SOURCE_DRAWING_FILE", ""), "SOURCE_SHA256": "",
        "STAGED_MASTER_FILE": "", "STAGED_SHA256": "",
        "PRESERVE_3D_PART_NUMBER": row.get("PRESERVE_3D_PART_NUMBER", ""),
        "PRESERVE_3D_REVISION": row.get("PRESERVE_3D_REVISION", ""),
        "PRESERVE_3D_IDENTIFIER": "", "PRESERVE_3D_DISCOVERED": "NO",
        "PRESERVE_3D_DISCOVERED_NAME": "",
        "DEFAULT_IMPORT_ACTION": "UseExisting", "MASTER_DRAWING_ACTION": "Overwrite",
        "PRESERVE_3D_ACTION": "UseExisting", "CLONE_PREFLIGHT": "NOT_RUN",
        "CLONE_LOG_STATUS": "NOT_READ", "TARGET_RESERVATION_STATUS": "UNKNOWN",
        "POST_IMPORT_VERIFICATION": "NOT_RUN", "WRITE_ATTEMPTED": "NO",
        "APPROVED": row.get("APPROVED", ""), "ENGINEER": row.get("ENGINEER", ""),
        "RESULT": "", "MESSAGE": "", "CLONE_LOG": "", "CLONE_LOG_EVIDENCE": "",
    }


def fail(report, result, message, error=None):
    report["RESULT"] = result
    report["MESSAGE"] = message
    if error is not None:
        detail = J16.error_text(error)
        if detail not in report["MESSAGE"]:
            report["MESSAGE"] += " | " + detail


def stage_copy(source, root, part_number, revision):
    folder = os.path.join(root, re.sub(r"[^A-Za-z0-9_.-]", "_",
                                       "{0}_{1}".format(part_number, revision)))
    os.makedirs(folder, exist_ok=True)
    target = os.path.join(folder, expected_master_native(part_number, revision))
    shutil.copy2(source, target)
    return target


def local_preflight(rows, csv_path, stage_root, timestamp, mode):
    reports, proposals = [], []
    keys = []
    for row in rows:
        try:
            pn, rev, _, _ = parse_row(row)
            keys.append((upper(pn), upper(rev)))
        except Exception:
            pass
    duplicate = {key for key, count in Counter(keys).items() if count > 1}

    for row in rows:
        report = base_report(row, timestamp, mode)
        reports.append(report)
        state = approval(row)
        if state == "INVALID":
            fail(report, "ERROR_APPROVAL_VALUE", "APPROVED must be YES, NO, or blank.")
            continue
        if mode == "APPLY_APPROVED" and state != "YES":
            report.update(RESULT="NOT_APPROVED", MESSAGE="No Teamcenter write authorized.")
            continue
        if mode == "APPLY_APPROVED" and not clean(row.get("ENGINEER")):
            fail(report, "ERROR_ENGINEER_REQUIRED", "ENGINEER is required for APPROVED=YES.")
            continue
        try:
            pn, rev, model_pn, model_rev = parse_row(row)
            if (upper(pn), upper(rev)) in duplicate:
                raise RuntimeError("Duplicate MASTER_PART_NUMBER/MASTER_REVISION target")
            source = resolve_local_path(csv_path, row.get("SOURCE_DRAWING_FILE"))
            if not source or not os.path.isfile(source):
                raise RuntimeError("SOURCE_DRAWING_FILE not found: {0}".format(source or "<blank>"))
            if not source.lower().endswith(".prt"):
                raise RuntimeError("SOURCE_DRAWING_FILE must be a native NX .prt file")
            target_id = master_id(pn, rev)
            supplied_id = clean(row.get("MASTER_IDENTIFIER"))
            if supplied_id and upper(supplied_id) != upper(target_id):
                raise RuntimeError("MASTER_IDENTIFIER does not match target Item/Revision")
            staged = stage_copy(source, stage_root, pn, rev)
            source_sha, staged_sha = J16.sha256(source), J16.sha256(staged)
            if source_sha.lower() != staged_sha.lower():
                raise RuntimeError("Staged master file is not byte-identical to source")
            report.update(
                MASTER_IDENTIFIER=target_id,
                SOURCE_DRAWING_FILE=source, SOURCE_SHA256=source_sha,
                STAGED_MASTER_FILE=staged, STAGED_SHA256=staged_sha,
                PRESERVE_3D_IDENTIFIER=master_id(model_pn, model_rev),
                RESULT="LOCAL_PREFLIGHT_OK",
                MESSAGE=("Drawing staged as a separate master item; preserved 3D will "
                         "be required and forced to UseExisting."),
            )
            proposals.append({
                "report": report, "pn": pn, "rev": rev, "target_id": target_id,
                "model_pn": model_pn, "model_rev": model_rev,
                "model_id": master_id(model_pn, model_rev),
                "source": source, "source_sha": source_sha,
                "staged": staged, "staged_sha": staged_sha,
            })
        except Exception as exc:
            fail(report, "ERROR_LOCAL_PREFLIGHT", "Local safety preflight failed.", exc)
    return reports, proposals


def normalize(value):
    return clean(value).lower().replace("\\", "/")


def master_tokens(part_number, revision):
    name = expected_master_native(part_number, revision).lower()
    return (master_id(part_number, revision).lower(), name,
            os.path.splitext(name)[0], "{0}/{1}".format(part_number, revision).lower())


def find_model(parts, staged, part_number, revision):
    tokens = master_tokens(part_number, revision)
    return [part for part in parts
            if not J16.same_part(part, staged)
            and any(token in normalize(part) for token in tokens)]


def classify_log(path, staged, target_id, parts, dry_run):
    content, _ = J16.read_text_file(path)
    if not content.strip():
        return ("LOG_MISSING", "UNKNOWN", "NOT_VERIFIED", "", "Clone log is missing or empty.")
    lines = [J16.compact_line(line) for line in content.splitlines() if clean(line)]
    target_tokens = tuple(set([normalize(staged), os.path.basename(staged).lower(),
                               os.path.splitext(os.path.basename(staged))[0].lower(),
                               target_id.lower()]))
    ref_tokens = []
    for part in parts:
        if not J16.same_part(part, staged):
            ref_tokens.extend([normalize(part), os.path.basename(normalize(part)),
                               os.path.splitext(os.path.basename(normalize(part)))[0]])
    ref_tokens = tuple(set(token for token in ref_tokens if token))
    blockers, ignored, success = [], [], []
    for index, line in enumerate(lines):
        low = line.lower()
        previous = lines[index - 1].lower() if index else ""
        target_here = any(token in low for token in target_tokens)
        ref_here = any(token in low for token in ref_tokens)
        target_prev = any(token in previous for token in target_tokens)
        ref_prev = any(token in previous for token in ref_tokens)
        is_target = target_here or (target_prev and not ref_prev)
        is_ref = ref_here or (ref_prev and not target_prev)
        context = "\n".join(lines[max(0, index - 1):min(len(lines), index + 2)])
        blocked = (any(term in low for term in J16.TARGET_BLOCK_TERMS)
                   or any(term in low for term in J16.GLOBAL_FAILURE_TERMS)
                   or ("error" in low and not J16.is_negated_failure(line))
                   or ("failed" in low and not J16.is_negated_failure(line)))
        if blocked:
            (ignored if is_ref and not is_target else blockers).append(context)
        elif any(term in low for term in J16.TARGET_SUCCESS_TERMS) and is_target:
            success.append(context)
    if blockers:
        evidence = J16.compact_line(blockers[0], 1200)
        reservation = "BLOCKED" if any(term in evidence.lower() for term in
            ("checked out", "reserved", "locked", "write access", "permission", "read-only")) else "WRITE_BLOCKED"
        return ("TARGET_BLOCKED", reservation, "FAILED", evidence,
                "Target/global write blocker found; master drawing was not verified.")
    if dry_run:
        evidence = J16.compact_line(ignored[0], 1200) if ignored else ""
        return ("PREFLIGHT_CLEAR", "NO_TARGET_BLOCKER_IN_CLONE_LOG", "PREFLIGHT_ONLY",
                evidence, "Dry-run log is clear; reference-only warnings are non-blocking.")
    if success:
        return ("TARGET_SUCCESS", "CLEAR_BY_APPLY_EVIDENCE", "VERIFIED_BY_CLONE_LOG",
                J16.compact_line(success[0], 1200), "Master-drawing overwrite verified by clone log.")
    return ("INCONCLUSIVE", "UNKNOWN", "NOT_VERIFIED",
            J16.compact_line(ignored[0], 1200) if ignored else "",
            "Clone returned, but persistence was not proven; J17 fails closed.")


def import_one(api, proposal, logfile, dry_run, log):
    clone, load, parts = api["clone"], None, []
    try:
        J16.terminate(clone)
        clone.Initialise(api["import_operation"])
        clone.SetFamilyTreatment(api["treat_as_lost"])
        clone.SetDefNaming(api["autotranslate"])
        clone.SetDefItemType("")
        clone.SetDefDirectory(os.path.dirname(proposal["staged"]))
        try: clone.SetAssocFileRootDir(os.path.dirname(proposal["staged"]))
        except Exception: pass
        clone.SetDefAction(api["use_existing"])
        clone.SetDefAssocFileCopy(True)
        clone.SetLogfile(logfile)
        try: clone.SetPropagateActions(False)
        except Exception: pass
        load = J16.add_assembly(clone, proposal["staged"])
        parts = J16.iterate_parts(clone)
        if not parts:
            raise RuntimeError("UF Clone discovered no parts")
        target_set = False
        for part in parts:
            if J16.same_part(part, proposal["staged"]):
                J16.set_action(clone, part, api["overwrite"]); target_set = True
            else:
                J16.set_action(clone, part, api["use_existing"])
        if not target_set:
            J16.set_action(clone, proposal["staged"], api["overwrite"])
        model_matches = find_model(parts, proposal["staged"], proposal["model_pn"], proposal["model_rev"])
        if not model_matches:
            raise RuntimeError("Required preserved 3D reference not discovered: {0}".format(proposal["model_id"]))
        for model in model_matches:
            J16.set_action(clone, model, api["use_existing"])
        proposal["report"].update(PRESERVE_3D_DISCOVERED="YES",
                                  PRESERVE_3D_DISCOVERED_NAME=" | ".join(model_matches))
        failures = J16.naming_failures(clone)
        clone.SetDryrun(bool(dry_run))
        try: clone.GenerateReport()
        except Exception: pass
        J16.perform_clone(clone, failures)
    finally:
        J16.dispose(load)
        J16.terminate(clone)
    result = classify_log(logfile, proposal["staged"], proposal["target_id"], parts, dry_run)
    log.write("  discovered={0}; preserved_3d=YES; dry_run={1}; log_status={2}".format(
        len(parts), dry_run, result[0]))
    return result


def apply_check(report, result):
    status, reservation, verification, evidence, message = result
    report.update(CLONE_LOG_STATUS=status, TARGET_RESERVATION_STATUS=reservation,
                  POST_IMPORT_VERIFICATION=verification, CLONE_LOG_EVIDENCE=evidence)
    return status, message


def log_path(proposal, mode, phase):
    return os.path.join(os.path.dirname(proposal["staged"]),
                        "J17_{0}_{1}_{2}_{3}.clone".format(
                            phase, mode, proposal["pn"], proposal["rev"]))


def run_preflight(api, proposals, mode, log, stop):
    failed = False
    for proposal in proposals:
        report = proposal["report"]
        report["CLONE_LOG"] = log_path(proposal, mode, "PREFLIGHT")
        try:
            status, message = apply_check(report, import_one(api, proposal, report["CLONE_LOG"], True, log))
            if status != "PREFLIGHT_CLEAR":
                report["CLONE_PREFLIGHT"] = "FAIL"
                fail(report, "FAILED_CLONE_PREFLIGHT", message)
                failed = True
                if stop: break
            else:
                report.update(CLONE_PREFLIGHT="PASS",
                              RESULT="DRY_RUN_OK" if mode == "DRY_RUN" else "CLONE_PREFLIGHT_OK",
                              MESSAGE=message)
        except Exception as exc:
            report["CLONE_PREFLIGHT"] = "FAIL"
            fail(report, "FAILED_CLONE_PREFLIGHT", "Clone/3D-preservation preflight failed.", exc)
            failed = True
            if stop: break
    return failed


def abort_pending(proposals, result, message):
    for proposal in proposals:
        if proposal["report"].get("RESULT") in ("LOCAL_PREFLIGHT_OK", "CLONE_PREFLIGHT_OK"):
            proposal["report"].update(RESULT=result, MESSAGE=message)


def execute(api, rows, csv_path, stage_root, timestamp, mode, log):
    reports, proposals = local_preflight(rows, csv_path, stage_root, timestamp, mode)
    if mode == "DRY_RUN":
        run_preflight(api, proposals, mode, log, False)
        return reports
    if any(report.get("RESULT", "").startswith("ERROR_") for report in reports):
        abort_pending(proposals, "BATCH_ABORTED_LOCAL_PREFLIGHT", "Approved row failed; no write attempted.")
        return reports
    if run_preflight(api, proposals, mode, log, True):
        abort_pending(proposals, "BATCH_ABORTED_CLONE_PREFLIGHT", "Clone/3D guard failed; no write attempted.")
        return reports
    for index, proposal in enumerate(proposals):
        report = proposal["report"]
        try:
            if (J16.sha256(proposal["source"]).lower() != proposal["source_sha"].lower()
                    or J16.sha256(proposal["staged"]).lower() != proposal["staged_sha"].lower()):
                raise RuntimeError("Source or staged file changed after preflight")
            report["CLONE_LOG"] = log_path(proposal, mode, "APPLY")
            report["WRITE_ATTEMPTED"] = "YES"
            status, message = apply_check(report, import_one(api, proposal, report["CLONE_LOG"], False, log))
            if status == "TARGET_SUCCESS":
                report.update(RESULT="IMPORT_VERIFIED",
                              MESSAGE="Master drawing verified; preserved 3D remained UseExisting.")
            else:
                fail(report, "FAILED_IMPORT_UNVERIFIED", message)
                abort_pending(proposals[index + 1:], "BATCH_STOPPED_AFTER_UNVERIFIED_WRITE",
                              "Previous write was not verified; no further write attempted.")
                break
        except Exception as exc:
            fail(report, "FAILED_IMPORT_APPLY", "Apply failed before reliable verification.", exc)
            abort_pending(proposals[index + 1:], "BATCH_STOPPED_AFTER_RUNTIME_FAILURE",
                          "Previous write failed; no further write attempted.")
            break
    return reports


def main():
    session = NXOpen.Session.GetSession()
    ufs = NXOpen.UF.UFSession.GetUFSession()
    log = J16.Log(session)
    mode, input_path, timestamp = configured_mode(), configured_input_path(), J16.stamp()
    stage_root = os.path.join(os.path.dirname(input_path), "J17_MASTER_DRAWING_STAGE_" + timestamp)
    log.write("=" * 76)
    log.write("J17 MASTER-DRAWING IMPORT WITH 3D PRESERVATION")
    log.write("Build: {0} | Mode: {1}".format(BUILD, mode))
    log.write("Input: {0}".format(input_path))
    log.write("Stage: {0}".format(stage_root))
    log.write("=" * 76)
    try:
        if mode not in VALID_MODES:
            raise RuntimeError("NX_J17_MODE must be DRY_RUN or APPLY_APPROVED")
        if not os.path.isfile(input_path):
            raise RuntimeError("Import CSV not found: {0}".format(input_path))
        rows = read_csv(input_path)
        if not rows:
            raise RuntimeError("Import CSV contains no rows")
        os.makedirs(stage_root, exist_ok=True)
        reports = execute(J16.resolve_clone_api(ufs, log), rows, input_path,
                          stage_root, timestamp, mode, log)
        report_path = os.path.join(os.path.dirname(input_path),
                                   "J17_{0}_{1}.csv".format(mode, timestamp))
        write_csv(report_path, reports)
        log.write("Report: {0}".format(report_path))
        for result, count in sorted(Counter(r.get("RESULT", "") for r in reports).items()):
            log.write("  {0}: {1}".format(result or "<blank>", count))
        failed = any(r.get("RESULT", "").startswith(("ERROR_", "FAILED_", "BATCH_"))
                     and (mode == "DRY_RUN" or upper(r.get("APPROVED")) == "YES")
                     for r in reports)
        if failed:
            raise RuntimeError("J17 safety/import failure; review {0}".format(report_path))
        log.write("FINAL STATUS: SUCCESS")
        return report_path
    except Exception as exc:
        log.write("FINAL STATUS: FAILED")
        log.write(J16.error_text(exc))
        log.write(traceback.format_exc())
        raise
    finally:
        try:
            J16.write_log(os.path.join(os.path.dirname(input_path),
                "J17_RUN_{0}_{1}.log".format(mode, timestamp)), log.lines)
        except Exception:
            pass


if __name__ == "__main__":
    main()
