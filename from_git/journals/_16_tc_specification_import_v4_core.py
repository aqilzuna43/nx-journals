"""J16 - Teamcenter X specification drawing replacement with direct verification.

NX X 2506 managed mode only.

Purpose:
- Replace the native .prt named reference of an existing UGPART drawing dataset
  attached by the Teamcenter specification relation.
- Preserve the master 3D Item/Revision completely. J16 does not clone, save,
  revise, overwrite, or otherwise modify the master 3D dataset.
- Block any target drawing that is checked out.
- Verify every write by exporting the exact managed specification afterward and
  comparing SHA-256 with the approved local source.

Run DRY_RUN first. APPLY_APPROVED writes only APPROVED=YES rows with ENGINEER.
"""

import csv
import datetime
import hashlib
import importlib.util
import os
import re
import shutil
import traceback
from collections import Counter

import NXOpen
import NXOpen.UF


USER_IMPORT_CSV = r""  # blank => <I/O root>\NX_TC_DRAWING_IMPORT.csv
USER_MODE = "DRY_RUN"  # DRY_RUN | APPLY_APPROVED

BUILD = "J16-TCX-SPECIFICATION-IMPORT-NX2506-V4"
DEFAULT_INPUT = "NX_TC_DRAWING_IMPORT.csv"
VALID_MODES = ("DRY_RUN", "APPLY_APPROVED")
DEFAULT_DATASET_TYPE = "UGPART"
DEFAULT_EXPORT_TOOL = "UGII V10-ALL"
RELATION_CANDIDATES = ("has specification", "specification", "IMAN_specification")

REQUIRED_COLUMNS = [
    "PART_NUMBER", "REVISION", "DWG_INDEX", "DRAWING_FILE", "APPROVED", "ENGINEER",
]

REPORT_COLUMNS = [
    "RUN_TIMESTAMP", "MODE", "CSV_ROW", "PART_NUMBER", "REVISION", "DWG_INDEX",
    "DRAWING_IDENTIFIER", "DATASET_NAME", "DATASET_TYPE", "RELATION_TYPE",
    "EXPORT_TOOL", "DRAWING_FILE", "SOURCE_SHA256", "CSV_EXPORT_SHA256",
    "TC_BASELINE_SHA256", "PREWRITE_TC_SHA256", "POST_IMPORT_TC_SHA256",
    "CHANGED_FROM_TC_BASELINE", "CHECKOUT_STATUS", "CHECKOUT_RECHECK",
    "IMPORT_METHOD", "MASTER_3D_ACTION", "STAGED_IMPORT_FILE", "STAGED_SHA256",
    "BASELINE_EXPORT_FILE", "PREWRITE_EXPORT_FILE", "POST_IMPORT_EXPORT_FILE",
    "BASELINE_EXPORT_PDI_CODE", "PREWRITE_EXPORT_PDI_CODE", "IMPORT_PDI_CODE",
    "POST_IMPORT_EXPORT_PDI_CODE", "WRITE_ATTEMPTED", "APPROVED", "ENGINEER",
    "RESULT", "MESSAGE",
]


def _load_clone_core():
    path = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "_16_tc_offline_drawing_import_core_v2.py",
    )
    if not os.path.isfile(path):
        return None
    spec = importlib.util.spec_from_file_location("nx_j16_clone_core_v2", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


_CLONE_CORE = _load_clone_core()
if _CLONE_CORE is not None:
    _J17_COMPAT_NAMES = (
        "Log", "clean", "upper", "env", "stamp", "io_root", "error_text",
        "dispose", "sha256", "resolve_clone_api", "terminate", "add_assembly",
        "naming_failures", "perform_clone", "set_action", "iterate_parts",
        "same_part", "read_text_file", "compact_line", "TARGET_BLOCK_TERMS",
        "GLOBAL_FAILURE_TERMS", "TARGET_SUCCESS_TERMS", "is_negated_failure",
    )
    for _compat_name in _J17_COMPAT_NAMES:
        if hasattr(_CLONE_CORE, _compat_name):
            globals()[_compat_name] = getattr(_CLONE_CORE, _compat_name)


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


def configured_dataset_type():
    return clean(env("NX_J16_DATASET_TYPE") or DEFAULT_DATASET_TYPE)


def configured_export_tool():
    return clean(env("NX_J16_EXPORT_TOOL") or DEFAULT_EXPORT_TOOL)


def configured_relation_candidates():
    override = clean(env("NX_J16_RELATION_TYPE"))
    return (override,) if override else RELATION_CANDIDATES


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


def normalized_name(value):
    return "".join(ch.lower() for ch in clean(value) if ch.isalnum())


def resolve_member(container, candidates, label):
    available = [name for name in dir(container) if not name.startswith("_")]
    normalized = {normalized_name(name): name for name in available}
    for candidate in candidates:
        if hasattr(container, candidate):
            return getattr(container, candidate)
    for candidate in candidates:
        actual = normalized.get(normalized_name(candidate))
        if actual and hasattr(container, actual):
            return getattr(container, actual)
    raise RuntimeError(
        "Unable to resolve {0}. Tried {1}; available: {2}".format(
            label, ", ".join(candidates), ", ".join(available) if available else "<none>"
        )
    )


def integer_list(value):
    if value is None or isinstance(value, bool):
        return []
    if isinstance(value, int):
        return [int(value)]
    if isinstance(value, (tuple, list)):
        return [int(item) for item in value if isinstance(item, int) and not isinstance(item, bool)]
    try:
        return [int(item) for item in value]
    except Exception:
        return []


def string_list(value):
    if isinstance(value, str):
        return [value] if value else []
    if isinstance(value, (tuple, list)):
        return [item for item in value if isinstance(item, str) and item]
    try:
        return [item for item in value if isinstance(item, str) and item]
    except Exception:
        return []


def parse_codes_and_strings(result):
    codes = []
    strings = []
    if isinstance(result, (tuple, list)):
        for item in result:
            candidate_codes = integer_list(item)
            candidate_strings = string_list(item)
            if candidate_codes and not codes:
                codes = candidate_codes
            strings.extend(candidate_strings)
    else:
        codes = integer_list(result)
        strings = string_list(result)
    return codes, strings


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
        if os.path.isfile(path) and valid_native(path, part_number, revision, drawing_index):
            matches.append(os.path.abspath(path))
    unique = {os.path.normcase(os.path.abspath(path)): path for path in matches}
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


def base_report(row, timestamp, mode, dataset_type, export_tool):
    return {
        "RUN_TIMESTAMP": timestamp, "MODE": mode, "CSV_ROW": row.get("_CSV_ROW", ""),
        "PART_NUMBER": row.get("PART_NUMBER", ""), "REVISION": row.get("REVISION", ""),
        "DWG_INDEX": row.get("DWG_INDEX", ""), "DRAWING_IDENTIFIER": "",
        "DATASET_NAME": "", "DATASET_TYPE": dataset_type, "RELATION_TYPE": "",
        "EXPORT_TOOL": export_tool, "DRAWING_FILE": row.get("DRAWING_FILE", ""),
        "SOURCE_SHA256": "", "CSV_EXPORT_SHA256": row.get("EXPORT_SHA256", ""),
        "TC_BASELINE_SHA256": "", "PREWRITE_TC_SHA256": "",
        "POST_IMPORT_TC_SHA256": "", "CHANGED_FROM_TC_BASELINE": "",
        "CHECKOUT_STATUS": "NOT_CHECKED", "CHECKOUT_RECHECK": "NOT_RUN",
        "IMPORT_METHOD": "NXOpen.PDM.FileManagement.ImportFiles",
        "MASTER_3D_ACTION": "NOT_TOUCHED", "STAGED_IMPORT_FILE": "",
        "STAGED_SHA256": "", "BASELINE_EXPORT_FILE": "", "PREWRITE_EXPORT_FILE": "",
        "POST_IMPORT_EXPORT_FILE": "", "BASELINE_EXPORT_PDI_CODE": "",
        "PREWRITE_EXPORT_PDI_CODE": "", "IMPORT_PDI_CODE": "",
        "POST_IMPORT_EXPORT_PDI_CODE": "", "WRITE_ATTEMPTED": "NO",
        "APPROVED": row.get("APPROVED", ""), "ENGINEER": row.get("ENGINEER", ""),
        "RESULT": "", "MESSAGE": "",
    }


def set_result(report, result, message, error=None):
    report["RESULT"] = result
    report["MESSAGE"] = message
    if error is not None:
        detail = error_text(error)
        if detail not in report["MESSAGE"]:
            report["MESSAGE"] += " | " + detail


def stage_source(source, stage_root, part_number, revision, drawing_index):
    folder = os.path.join(
        stage_root,
        re.sub(r"[^A-Za-z0-9_.-]", "_", "{0}_{1}_DWG{2}".format(
            part_number, revision, drawing_index
        )),
    )
    if os.path.isdir(folder):
        shutil.rmtree(folder)
    os.makedirs(folder)
    target = os.path.join(folder, expected_native(part_number, revision, drawing_index))
    shutil.copy2(source, target)
    return folder, target
