from _16_tc_specification_import_v4_pdm import *

def local_preflight(rows, csv_path, timestamp, mode, dataset_type, export_tool, stage_root):
    duplicate_keys = duplicate_target_keys(rows)
    reports, proposals = [], []
    for row in rows:
        report = base_report(row, timestamp, mode, dataset_type, export_tool)
        reports.append(report)
        approval = approval_state(row)
        if approval == "INVALID":
            set_result(report, "ERROR_APPROVAL_VALUE", "APPROVED must be YES, NO, or blank.")
            continue
        if mode == "APPLY_APPROVED" and approval != "YES":
            set_result(report, "NOT_APPROVED", "No Teamcenter write authorized.")
            continue
        if mode == "APPLY_APPROVED" and not clean(row.get("ENGINEER")):
            set_result(report, "ERROR_ENGINEER_REQUIRED", "ENGINEER is required when APPROVED=YES.")
            continue
        try:
            pn, rev, idx = parse_target(row)
            if (upper(pn), upper(rev), idx) in duplicate_keys:
                raise RuntimeError("The same PART_NUMBER/REVISION/DWG_INDEX appears more than once.")
            identifier = drawing_id(pn, rev, idx)
            supplied_identifier = clean(row.get("DRAWING_IDENTIFIER"))
            if supplied_identifier and upper(supplied_identifier) != upper(identifier):
                raise RuntimeError("DRAWING_IDENTIFIER does not match target identity.")
            drawing, resolution, matches = resolve_drawing_file(
                csv_path, row.get("DRAWING_FILE"), pn, rev, idx
            )
            if resolution == "MULTIPLE":
                raise RuntimeError("More than one valid drawing matched: {0}".format(" | ".join(matches)))
            if not drawing or not os.path.isfile(drawing):
                raise RuntimeError("DRAWING_FILE was not found: {0}".format(drawing or "<blank>"))
            if not drawing.lower().endswith(".prt") or not valid_native(drawing, pn, rev, idx):
                raise RuntimeError("DRAWING_FILE is not the expected native AutoTranslate drawing.")
            source_sha = sha256(drawing)
            stage_dir, staged = stage_source(drawing, stage_root, pn, rev, idx)
            staged_sha = sha256(staged)
            if source_sha.lower() != staged_sha.lower():
                raise RuntimeError("Staged import copy does not match source SHA-256.")
            report.update(
                DRAWING_IDENTIFIER=identifier, DATASET_NAME=dataset_name(pn, rev, idx),
                DRAWING_FILE=drawing, SOURCE_SHA256=source_sha,
                STAGED_IMPORT_FILE=staged, STAGED_SHA256=staged_sha,
                RESULT="LOCAL_PREFLIGHT_OK",
                MESSAGE="Local identity and isolated staging checks passed."
                + (" Shortened filename was auto-resolved." if resolution == "AUTO_RESOLVED" else ""),
            )
            proposals.append({
                "row": row, "report": report, "part_number": pn, "revision": rev,
                "drawing_index": idx, "identifier": identifier,
                "dataset_name": report["DATASET_NAME"], "dataset_type": dataset_type,
                "export_tool": export_tool, "source": drawing, "source_sha": source_sha,
                "stage_dir": stage_dir, "staged": staged, "staged_sha": staged_sha,
                "baseline_sha": "", "relation_type": "",
            })
        except Exception as exc:
            set_result(report, "ERROR_LOCAL_PREFLIGHT", "Local safety preflight failed.", exc)
    return reports, proposals


def target_root(base, proposal, phase):
    return os.path.join(base, re.sub(
        r"[^A-Za-z0-9_.-]", "_", "{0}_{1}_DWG{2}_{3}".format(
            proposal["part_number"], proposal["revision"], proposal["drawing_index"], phase
        )
    ))


def run_managed_preflight(session, pdm, fm, proposals, export_root, log):
    for proposal in proposals:
        report = proposal["report"]
        try:
            relation, exported, baseline_sha = resolve_relation_and_baseline(
                fm, proposal, target_root(export_root, proposal, "BASELINE")
            )
            proposal["relation_type"] = relation
            proposal["baseline_sha"] = baseline_sha
            report.update(
                RELATION_TYPE=relation, BASELINE_EXPORT_FILE=exported,
                TC_BASELINE_SHA256=baseline_sha,
            )
            csv_baseline = clean(report.get("CSV_EXPORT_SHA256"))
            if csv_baseline and csv_baseline.lower() != baseline_sha.lower():
                raise RuntimeError(
                    "CSV EXPORT_SHA256 does not match the current Teamcenter drawing."
                )
            changed = proposal["source_sha"].lower() != baseline_sha.lower()
            report["CHANGED_FROM_TC_BASELINE"] = "YES" if changed else "NO"
            if not changed:
                set_result(report, "SKIPPED_UNCHANGED", "Local source already matches Teamcenter.")
                continue
            checkout = check_target_checkout(session, pdm, proposal["identifier"], log)
            report["CHECKOUT_STATUS"] = checkout
            if checkout == "CHECKED_OUT":
                raise RuntimeError("The exact drawing specification is checked out.")
            if checkout != "CLEAR":
                raise RuntimeError("The drawing checkout state could not be proven clear.")
            set_result(
                report, "MANAGED_PREFLIGHT_OK",
                "Existing UGPART specification exported; checkout is clear; master 3D is untouched.",
            )
        except Exception as exc:
            set_result(report, "FAILED_MANAGED_PREFLIGHT", "Managed Teamcenter preflight failed.", exc)


def approved_preflight_failure(report):
    result = report.get("RESULT", "")
    if result == "ERROR_APPROVAL_VALUE":
        return True
    if upper(report.get("APPROVED")) != "YES":
        return False
    return result.startswith(("ERROR_", "FAILED_"))


def mark_remaining(proposals, start, result, message):
    for proposal in proposals[start:]:
        report = proposal["report"]
        if report.get("RESULT") == "MANAGED_PREFLIGHT_OK":
            set_result(report, result, message)


def apply_one(session, pdm, fm, proposal, export_root, log):
    report = proposal["report"]
    if sha256(proposal["source"]).lower() != proposal["source_sha"].lower():
        raise RuntimeError("Local source changed after preflight.")
    if sha256(proposal["staged"]).lower() != proposal["staged_sha"].lower():
        raise RuntimeError("Staged import file changed after preflight.")
    checkout = check_target_checkout(session, pdm, proposal["identifier"], log)
    report["CHECKOUT_RECHECK"] = checkout
    if checkout == "CHECKED_OUT":
        raise RuntimeError("The drawing became checked out after preflight.")
    if checkout != "CLEAR":
        raise RuntimeError("Checkout state could not be proven clear immediately before write.")
    pre_file, pre_sha = export_exact_dataset(
        fm, proposal, proposal["relation_type"],
        target_root(export_root, proposal, "PREWRITE"),
        "PREWRITE_EXPORT_PDI_CODE", "PREWRITE_EXPORT_FILE",
    )
    report.update(PREWRITE_EXPORT_FILE=pre_file, PREWRITE_TC_SHA256=pre_sha)
    if pre_sha.lower() != proposal["baseline_sha"].lower():
        raise RuntimeError("Teamcenter drawing changed after preflight.")
    report["WRITE_ATTEMPTED"] = "YES"
    import_code = invoke_import_files(fm, proposal)
    report["IMPORT_PDI_CODE"] = "" if import_code is None else str(import_code)
    if import_code != 0:
        raise RuntimeError(
            "PDM ImportFiles failed with PDI code {0}. Target may be checked out, "
            "released, missing, or not writable.".format(
                "<missing>" if import_code is None else import_code
            )
        )
    post_file, post_sha = export_exact_dataset(
        fm, proposal, proposal["relation_type"],
        target_root(export_root, proposal, "POSTIMPORT"),
        "POST_IMPORT_EXPORT_PDI_CODE", "POST_IMPORT_EXPORT_FILE",
    )
    report.update(POST_IMPORT_EXPORT_FILE=post_file, POST_IMPORT_TC_SHA256=post_sha)
    if post_sha.lower() != proposal["source_sha"].lower():
        raise RuntimeError(
            "Import returned success, but re-exported Teamcenter SHA-256 does not match source."
        )
    set_result(
        report, "IMPORT_VERIFIED",
        "Exact UGPART specification replaced and verified by re-export SHA-256. Master 3D untouched.",
    )


def execute(session, pdm, fm, rows, csv_path, timestamp, mode, log, dataset_type, export_tool, work_root):
    stage_root = os.path.join(work_root, "STAGE")
    export_root = os.path.join(work_root, "EXPORT")
    os.makedirs(stage_root, exist_ok=True)
    os.makedirs(export_root, exist_ok=True)
    reports, proposals = local_preflight(
        rows, csv_path, timestamp, mode, dataset_type, export_tool, stage_root
    )
    local_failures = [
        report for report in reports
        if report.get("RESULT", "").startswith("ERROR_")
        and (mode == "DRY_RUN" or report.get("RESULT") == "ERROR_APPROVAL_VALUE"
             or upper(report.get("APPROVED")) == "YES")
    ]
    if local_failures and mode == "APPLY_APPROVED":
        for proposal in proposals:
            if proposal["report"].get("RESULT") == "LOCAL_PREFLIGHT_OK":
                set_result(
                    proposal["report"], "BATCH_ABORTED_LOCAL_PREFLIGHT",
                    "An approved row failed local preflight. No Teamcenter write was attempted.",
                )
        return reports
    run_managed_preflight(session, pdm, fm, proposals, export_root, log)
    if mode == "DRY_RUN":
        for proposal in proposals:
            if proposal["report"].get("RESULT") == "MANAGED_PREFLIGHT_OK":
                set_result(
                    proposal["report"], "DRY_RUN_OK",
                    "Specification exists, checkout is clear, source differs, and no write was attempted.",
                )
        return reports
    blocking = [report for report in reports if approved_preflight_failure(report)]
    if blocking:
        for proposal in proposals:
            if proposal["report"].get("RESULT") == "MANAGED_PREFLIGHT_OK":
                set_result(
                    proposal["report"], "BATCH_ABORTED_MANAGED_PREFLIGHT",
                    "At least one approved drawing failed managed preflight. No write was attempted.",
                )
        return reports
    writable = [p for p in proposals if p["report"].get("RESULT") == "MANAGED_PREFLIGHT_OK"]
    for index, proposal in enumerate(writable):
        try:
            apply_one(session, pdm, fm, proposal, export_root, log)
            log.write("  VERIFIED {0}".format(proposal["identifier"]))
        except Exception as exc:
            result = "FAILED_BEFORE_WRITE" if proposal["report"].get("WRITE_ATTEMPTED") == "NO" else "FAILED_IMPORT_OR_VERIFICATION"
            set_result(proposal["report"], result, "Specification import stopped.", exc)
            log.write("  STOPPED {0}: {1}".format(proposal["identifier"], error_text(exc)))
            log.write(traceback.format_exc())
            mark_remaining(
                writable, index + 1, "BATCH_STOPPED_AFTER_FAILURE",
                "A previous specification import failed or was not verified. No write was attempted.",
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
    log = Log(session)
    mode = configured_mode()
    input_path = configured_input_path()
    timestamp = stamp()
    dataset_type = configured_dataset_type()
    export_tool = configured_export_tool()
    log.write("=" * 72)
    log.write("J16 TEAMCENTER X SPECIFICATION DRAWING IMPORT")
    log.write("Build: {0} | Mode: {1}".format(BUILD, mode))
    log.write("Method: PDM FileManagement ImportFiles + re-export SHA-256")
    log.write("Dataset: {0}; relation resolved by exact baseline export".format(dataset_type))
    log.write("Master 3D action: NOT_TOUCHED")
    log.write("Input: {0}".format(input_path))
    log.write("=" * 72)
    report_path = ""
    fm = None
    try:
        if mode not in VALID_MODES:
            raise RuntimeError("USER_MODE/NX_J16_MODE must be DRY_RUN or APPLY_APPROVED.")
        if not os.path.isfile(input_path):
            raise RuntimeError("Import CSV not found: {0}".format(input_path))
        rows = read_csv(input_path)
        if not rows:
            raise RuntimeError("Import CSV contains no data rows.")
        pdm, fm = new_file_management(session)
        work_root = os.path.join(os.path.dirname(input_path), "J16_WORK_{0}".format(timestamp))
        os.makedirs(work_root)
        reports = execute(
            session, pdm, fm, rows, input_path, timestamp, mode, log,
            dataset_type, export_tool, work_root,
        )
        report_path = os.path.join(
            os.path.dirname(input_path), "J16_{0}_{1}.csv".format(mode, timestamp)
        )
        write_csv(report_path, reports)
        log.write("Report: {0}".format(report_path))
        log.write("Work evidence: {0}".format(work_root))
        for result, count in sorted(summary_counts(reports).items()):
            log.write("  {0}: {1}".format(result, count))
        if has_failure(reports, mode):
            log.write("FINAL STATUS: FAILED")
            log.write("Failures are recorded in the CSV report; handled row failures do not raise an NX prompt.")
        else:
            log.write("FINAL STATUS: SUCCESS")
    except Exception as exc:
        if "FINAL STATUS: FAILED" not in log.lines:
            log.write("FINAL STATUS: FAILED")
        log.write(error_text(exc))
        log.write(traceback.format_exc())
        raise
    finally:
        dispose(fm)
        try:
            log_dir = os.path.dirname(input_path) if input_path else io_root()
            os.makedirs(log_dir, exist_ok=True)
            write_log(
                os.path.join(log_dir, "J16_RUN_{0}_{1}.log".format(mode, timestamp)),
                log.lines,
            )
        except Exception:
            pass
    return report_path


if __name__ == "__main__":
    main()
