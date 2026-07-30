"""J19 - read-only Teamcenter drawing-import contract probe.

NX X 2506 only.

This journal proves the runtime contracts J16 depends on:
- the exact /specification/ identifier opens as that same managed drawing;
- PDMPart returns authoritative checkout state and owner;
- PDM FileManagement can export the exact UGPART specification for hashing;
- the proven J16 V1 writer can complete a UF Clone dry run.

J19 never checks out, saves, imports, or otherwise writes Teamcenter data.
It writes only local JSON, text, exported evidence, and clone dry-run files.
"""

import importlib.util
import json
import os
import traceback

import NXOpen
import NXOpen.UF


USER_PART_NUMBER = "264MN020818A01"
USER_REVISION = "A"
USER_DWG_INDEX = 1
BUILD = "J19-J16-TEAMCENTER-CONTRACT-NX2506-V1"


def load_j16():
    path = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "16_tc_offline_drawing_import.py",
    )
    spec = importlib.util.spec_from_file_location("nx_journal_16_probe", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


J16 = load_j16()


def configured_target():
    part_number = J16.env("NX_J19_PART_NUMBER") or USER_PART_NUMBER
    revision = J16.env("NX_J19_REVISION") or USER_REVISION
    raw_index = J16.env("NX_J19_DWG_INDEX") or str(USER_DWG_INDEX)
    try:
        drawing_index = int(raw_index)
    except Exception:
        raise RuntimeError("NX_J19_DWG_INDEX must be an integer.")
    if drawing_index < 1:
        raise RuntimeError("NX_J19_DWG_INDEX must be >= 1.")
    return part_number, revision, drawing_index


def configured_csv_path():
    value = J16.env("NX_TC_DRAWING_IMPORT_FILE")
    if value:
        return os.path.abspath(os.path.expanduser(value))
    return J16.configured_input_path()


def find_csv_evidence(csv_path, target):
    result = {
        "csv_path": csv_path,
        "drawing_file": "",
        "expected_sha256": "",
        "csv_row": "",
    }
    if not csv_path or not os.path.isfile(csv_path):
        return result
    part_number, revision, drawing_index = target
    for row in J16.read_csv(csv_path):
        try:
            candidate = J16.parse_target(row)
        except Exception:
            continue
        if (
            J16.upper(candidate[0]) == J16.upper(part_number)
            and J16.upper(candidate[1]) == J16.upper(revision)
            and candidate[2] == drawing_index
        ):
            result.update(
                drawing_file=J16.resolve_local_path(
                    csv_path, row.get("DRAWING_FILE")
                ),
                expected_sha256=J16.clean(row.get("EXPORT_SHA256")),
                csv_row=row.get("_CSV_ROW", ""),
            )
            break
    return result


def current_teamcenter_user(session):
    pdm = getattr(session, "PdmSession", None)
    method = getattr(pdm, "GetUserName", None)
    if not callable(method):
        return ""
    try:
        return J16.clean(method())
    except Exception:
        return ""


def write_json(path, value):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as handle:
        json.dump(value, handle, indent=2, sort_keys=True)
        handle.write("\n")


def main():
    session = NXOpen.Session.GetSession()
    ufs = NXOpen.UF.UFSession.GetUFSession()
    log = J16.Log(session)
    timestamp = J16.stamp()
    target = configured_target()
    part_number, revision, drawing_index = target
    identifier = J16.drawing_id(part_number, revision, drawing_index)
    csv_evidence = find_csv_evidence(configured_csv_path(), target)

    output_base = J16.env("NX_J19_OUTPUT_DIR")
    if not output_base:
        output_base = (
            os.path.dirname(csv_evidence["csv_path"])
            if csv_evidence["csv_path"]
            and os.path.isdir(os.path.dirname(csv_evidence["csv_path"]))
            else J16.io_root()
        )
    output_root = os.path.join(
        os.path.abspath(os.path.expanduser(output_base)),
        "J19_CONTRACT_{0}".format(timestamp),
    )
    os.makedirs(output_root, exist_ok=True)

    report = {
        "journal": BUILD,
        "timestamp": timestamp,
        "target": {
            "part_number": part_number,
            "revision": revision,
            "drawing_index": drawing_index,
            "identifier": identifier,
        },
        "current_teamcenter_user": current_teamcenter_user(session),
        "csv_evidence": csv_evidence,
        "checkout": {},
        "exact_dataset_export": {},
        "clone_dry_run": {},
        "teamcenter_write_attempted": False,
        "result": "PROBE_INCOMPLETE",
    }

    log.write("=" * 72)
    log.write("J19 READ-ONLY J16 TEAMCENTER CONTRACT PROBE")
    log.write("Build: {0}".format(BUILD))
    log.write("Target: {0}".format(identifier))
    log.write("Output: {0}".format(output_root))
    log.write("Teamcenter writes: FORBIDDEN")
    log.write("=" * 72)

    file_management = None
    try:
        checkout = J16.inspect_target_checkout(session, identifier, log)
        report["checkout"] = checkout
        log.write(
            "Checkout: state={0}; owner={1}; opened={2}; raw={3}".format(
                checkout.get("state", "UNKNOWN"),
                checkout.get("owner", "") or "<blank>",
                checkout.get("opened_identifier", "") or "<blank>",
                checkout.get("raw", "") or "<blank>",
            )
        )

        _, file_management = J16.new_file_management(session)
        probe_report = {
            "RELATION_TYPE": "",
            "BASELINE_EXPORT_PDI_CODE": "",
            "BASELINE_EXPORT_FILE": "",
        }
        proposal = {
            "report": probe_report,
            "part_number": part_number,
            "revision": revision,
            "drawing_index": drawing_index,
            "dataset_name": J16.dataset_name(
                part_number, revision, drawing_index
            ),
            "dataset_type": J16.configured_dataset_type(),
            "export_tool": J16.configured_export_tool(),
            "relation_type": "",
        }
        exported, exported_sha = J16.resolve_relation_and_export(
            file_management,
            proposal,
            os.path.join(output_root, "EXACT_DATASET_EXPORT"),
            "BASELINE",
            "BASELINE_EXPORT_PDI_CODE",
            "BASELINE_EXPORT_FILE",
        )
        expected_sha = csv_evidence["expected_sha256"]
        report["exact_dataset_export"] = {
            "status": "PASS",
            "relation_type": probe_report["RELATION_TYPE"],
            "pdi_code": probe_report["BASELINE_EXPORT_PDI_CODE"],
            "file": exported,
            "sha256": exported_sha,
            "expected_sha256": expected_sha,
            "matches_csv_baseline": (
                exported_sha.lower() == expected_sha.lower()
                if expected_sha
                else None
            ),
        }
        log.write(
            "Exact export: relation={0}; pdi={1}; sha256={2}; file={3}".format(
                probe_report["RELATION_TYPE"],
                probe_report["BASELINE_EXPORT_PDI_CODE"],
                exported_sha,
                exported,
            )
        )

        drawing = csv_evidence["drawing_file"]
        if drawing and os.path.isfile(drawing):
            clone_log = os.path.join(output_root, "J19_UFCLONE_DRY_RUN.clone")
            api = J16.resolve_clone_api(ufs, log)
            discovered = J16.import_one(
                api, drawing, clone_log, True, log
            )
            report["clone_dry_run"] = {
                "status": "PASS",
                "drawing_file": drawing,
                "drawing_sha256": J16.sha256(drawing),
                "clone_log": clone_log,
                "discovered_parts": list(discovered),
            }
            log.write(
                "UF Clone dry run: PASS; discovered={0}; log={1}".format(
                    len(discovered), clone_log
                )
            )
        else:
            report["clone_dry_run"] = {
                "status": "SKIPPED_NO_LOCAL_DRAWING",
                "drawing_file": drawing,
            }
            log.write(
                "UF Clone dry run: SKIPPED because the matching local "
                "DRAWING_FILE was not found."
            )

        exact_identity = (
            J16.upper(checkout.get("opened_identifier")).replace("\\", "/")
            == J16.upper(identifier).replace("\\", "/")
        )
        if (
            exact_identity
            and checkout.get("state") in ("CHECKED_IN", "CHECKED_OUT")
            and report["exact_dataset_export"].get("status") == "PASS"
        ):
            report["result"] = "PROBE_COMPLETE"
        log.write("FINAL STATUS: {0}".format(report["result"]))
    except Exception as exc:
        report["error"] = J16.error_text(exc)
        report["traceback"] = traceback.format_exc()
        log.write("FINAL STATUS: PROBE_INCOMPLETE")
        log.write(J16.error_text(exc))
        log.write(report["traceback"])
    finally:
        J16.dispose(file_management)
        json_path = os.path.join(
            output_root, "J19_CONTRACT_{0}.json".format(timestamp)
        )
        log_path = os.path.join(
            output_root, "J19_CONTRACT_{0}.log".format(timestamp)
        )
        write_json(json_path, report)
        J16.write_log(log_path, log.lines)

    return json_path


if __name__ == "__main__":
    main()
