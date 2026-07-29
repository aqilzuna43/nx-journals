"""J16 V3 compatibility entry point for Teamcenter X drawing import.

This wrapper keeps the proven J16 V2 implementation in
_16_tc_offline_drawing_import_core_v2.py and adds two operational fixes:
- safely auto-resolve one unique Teamcenter AutoTranslate drawing filename when
  the CSV contains only the shortened dataset filename;
- do not raise an NX error prompt for row failures already written to the J16
  CSV report. Unexpected setup/runtime exceptions still raise normally.

NX X 2506 only.
"""

import importlib.util
import os
import traceback

import NXOpen
import NXOpen.UF


_CORE_FILE = "_16_tc_offline_drawing_import_core_v2.py"
_CORE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), _CORE_FILE)
if not os.path.isfile(_CORE_PATH):
    raise RuntimeError("J16 V2 core not found beside wrapper: {0}".format(_CORE_PATH))

_spec = importlib.util.spec_from_file_location("nx_j16_core_v2", _CORE_PATH)
_core = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_core)

# Re-export the V2 API so J17 can continue importing J16 from the standard path.
for _name in dir(_core):
    if not _name.startswith("__"):
        globals()[_name] = getattr(_core, _name)

BUILD = "J16-TCX-DRAWING-IMPORT-NX2506-V3"
_core.BUILD = BUILD
_original_local_preflight = _core.local_preflight


def _matching_native_drawings(folder, part_number, revision, drawing_index):
    """Return valid AutoTranslate drawing files in one explicit folder only."""
    if not folder or not os.path.isdir(folder):
        return []

    matches = []
    expected = expected_native(part_number, revision, drawing_index).lower()
    for name in os.listdir(folder):
        path = os.path.join(folder, name)
        if not os.path.isfile(path):
            continue
        if name.lower() == expected or valid_native(
            path, part_number, revision, drawing_index
        ):
            matches.append(os.path.abspath(path))

    # Windows filenames are case-insensitive. Deduplicate on normalized path.
    unique = {}
    for path in matches:
        unique[os.path.normcase(os.path.abspath(path))] = path
    return sorted(unique.values(), key=lambda value: value.lower())


def resolve_drawing_file(csv_path, supplied_value, part_number, revision, drawing_index):
    """Resolve an exact path or one unique AutoTranslate match beside the CSV.

    J16 never searches recursively and never guesses between multiple files.
    """
    requested = resolve_local_path(csv_path, supplied_value)
    if requested and os.path.isfile(requested):
        return requested, "EXACT", []
    if not clean(supplied_value):
        return requested, "NOT_FOUND", []

    folder = os.path.dirname(requested) if requested else os.path.dirname(csv_path)
    matches = _matching_native_drawings(
        folder, part_number, revision, drawing_index
    )
    if len(matches) == 1:
        return matches[0], "AUTO_RESOLVED", matches
    if len(matches) > 1:
        return requested, "MULTIPLE", matches
    return requested, "NOT_FOUND", []


def local_preflight(rows, csv_path, timestamp, mode):
    """Add safe filename discovery before the original V2 local preflight."""
    prepared_rows = []
    auto_resolved = {}
    multiple_matches = {}

    for source in rows:
        row = dict(source)
        row_number = row.get("_CSV_ROW", "")
        try:
            part_number, revision, drawing_index = parse_target(row)
            resolved, status, matches = resolve_drawing_file(
                csv_path,
                row.get("DRAWING_FILE"),
                part_number,
                revision,
                drawing_index,
            )
            if status == "AUTO_RESOLVED":
                requested = resolve_local_path(csv_path, row.get("DRAWING_FILE"))
                row["DRAWING_FILE"] = resolved
                auto_resolved[row_number] = (requested, resolved)
            elif status == "MULTIPLE":
                multiple_matches[row_number] = matches
        except Exception:
            # Preserve the original V2 input validation and error reporting.
            pass
        prepared_rows.append(row)

    reports, proposals = _original_local_preflight(
        prepared_rows, csv_path, timestamp, mode
    )

    for report in reports:
        row_number = report.get("CSV_ROW", "")
        if row_number in auto_resolved:
            requested, resolved = auto_resolved[row_number]
            prefix = (
                "Auto-resolved shortened DRAWING_FILE '{0}' to the unique "
                "Teamcenter AutoTranslate file '{1}'."
            ).format(requested, resolved)
            report["MESSAGE"] = (
                prefix + (" | " + report["MESSAGE"] if report.get("MESSAGE") else "")
            )
        if row_number in multiple_matches:
            matches = multiple_matches[row_number]
            report["RESULT"] = "ERROR_MULTIPLE_MATCHING_DRAWINGS"
            report["MESSAGE"] = (
                "More than one Teamcenter AutoTranslate drawing matched this target. "
                "Specify the exact DRAWING_FILE path: {0}"
            ).format(" | ".join(matches))

    if multiple_matches:
        proposals = [
            proposal
            for proposal in proposals
            if proposal["report"].get("CSV_ROW", "") not in multiple_matches
        ]

    return reports, proposals


# Functions defined in the V2 module resolve globals in that module. Patch only
# its local-preflight hook so execute() uses the V3 resolver without changing
# any UF Clone or overwrite safety behavior.
_core.local_preflight = local_preflight


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
    log.write("Filename resolution: exact path or one unique AutoTranslate match")
    log.write("Verification: completed clone-log evidence; fail closed if inconclusive")
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
            raise RuntimeError(
                "Import CSV contains no data rows: {0}".format(input_path)
            )

        api = resolve_clone_api(ufs, log)
        reports = _core.execute(
            api, rows, input_path, timestamp, current_mode, log
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
            log.write(
                "J16 completed with failed safety/import rows. The failure is "
                "already recorded in the CSV report; no NX exception prompt was raised."
            )
        else:
            log.write("FINAL STATUS: SUCCESS")

    except Exception as exc:
        # Unexpected setup/runtime faults are not normal row results and remain
        # visible to NX for troubleshooting.
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
