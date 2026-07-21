"""
Journal 08 - NX X / Teamcenter Drawing Identity Probe

Run this after manually opening one known dwg1/dwg2 specification and making
that drawing the active display part. The journal writes TXT and JSON reports
to Desktop\\NX_X_IDENTITY_PROBE.

The probe does not save or modify parts. It performs controlled reopen tests
using identifiers collected from the already-loaded display part. Any part
newly loaded by a test is restored and closed.

Target: NX 2312 embedded Python 3.10
"""

import datetime
import json
import os
import platform
import re
import sys
import traceback

import NXOpen
import NXOpen.UF


OUTPUT_FOLDER_NAME = "NX_X_IDENTITY_PROBE"
RUN_ACTIVE_IDENTIFIER_REOPEN_TESTS = True
RUN_GENERATED_AT_DB_TEST = True
MAX_LOADED_PARTS = 250
MAX_ATTRIBUTES = 500
MAX_REOPEN_CANDIDATES = 20

MEMBER_TOKENS = (
    "cloud", "database", "dataset", "db", "file", "identifier", "item",
    "journal", "name", "part", "path", "pdm", "revision", "teamcenter",
    "uid", "ugmgr",
)
ATTRIBUTE_TOKENS = ("DB_", "DATASET", "ITEM", "PART", "REV", "UID")
SAFE_ENV_VARS = (
    "UGII_BASE_DIR", "UGII_ROOT_DIR", "UGII_LANG", "UGII_TMP_DIR",
    "UGII_USER_DIR", "UGII_ENV_FILE", "UGII_CUSTOM_DIRECTORY_FILE",
    "UGII_UGMGR_COMMUNICATION", "UGII_UGMGR_HTTP_URL",
    "UGII_UGMGR_TC_ROOT", "UGII_TEAMCENTER_ROOT", "UGII_CLOUDDM_ROOT",
    "UGII_CLOUDDM_URL", "TEAMCENTER_ROOT", "TC_ROOT", "TC_DATA",
    "FMS_HOME", "FCC_HOME", "TCCS_CONFIG", "TEMP", "TMP",
    "LOCALAPPDATA", "APPDATA", "USERPROFILE",
)
RELATED_ENV_TOKENS = (
    "UGII", "UGMGR", "TEAMCENTER", "CLOUDDM", "FMS", "FCC", "TCCS",
    "TC_", "NX_",
)
SENSITIVE_TOKENS = (
    "AUTH", "COOKIE", "CREDENTIAL", "KEY", "PASSWORD", "SECRET", "TOKEN",
)
DRAWING_RE = re.compile(r"(?:^|[-_])DWG(\d+)(?:$|[^A-Z0-9])", re.I)


def text(value):
    if value is None:
        return ""
    try:
        return str(value).strip()
    except Exception:
        return ""


def short_repr(value, limit=1200):
    try:
        result = repr(value)
    except Exception as error:
        result = "<repr failed: {0}>".format(error)
    return result if len(result) <= limit else result[:limit] + "... <truncated>"


def type_name(value):
    if value is None:
        return "None"
    return "{0}.{1}".format(type(value).__module__, type(value).__name__)


def exception_info(error):
    info = {
        "type": type_name(error),
        "message": text(error),
        "repr": short_repr(error),
    }
    for name in ("ErrorCode", "error_code", "Code", "code", "Message", "message"):
        try:
            value = getattr(error, name)
            if not callable(value):
                info[name] = short_repr(value)
        except Exception:
            pass
    return info


def identity(nx_object):
    if nx_object is None:
        return ("NONE", "")
    try:
        return ("TAG", str(nx_object.Tag))
    except Exception:
        pass
    for name in ("FullPath", "Name", "Leaf"):
        try:
            value = text(getattr(nx_object, name))
            if value:
                return (name.upper(), value.upper())
        except Exception:
            pass
    return ("OBJECT", str(id(nx_object)))


def dispose(value):
    if value is not None:
        try:
            value.Dispose()
        except Exception:
            pass


def unwrap_open(value):
    if isinstance(value, (tuple, list)):
        part = value[0] if value else None
        status = value[1] if len(value) > 1 else None
        return part, status
    return value, None


def output_root():
    configured = text(os.environ.get("NX_JOURNALS_IO_DIR"))
    if configured:
        base = os.path.abspath(os.path.expanduser(configured))
    elif text(os.environ.get("USERPROFILE")):
        base = os.path.join(os.environ["USERPROFILE"], "Desktop")
    else:
        base = os.path.join(os.path.expanduser("~"), "Desktop")
    return os.path.join(base, OUTPUT_FOLDER_NAME)


class Logger:
    def __init__(self, session):
        self.lines = []
        self.window = None
        try:
            self.window = session.ListingWindow
            self.window.Open()
        except Exception:
            pass

    def write(self, message=""):
        value = str(message)
        self.lines.append(value)
        if self.window is not None:
            try:
                for line in value.splitlines() or [""]:
                    self.window.WriteFullline(line)
            except Exception:
                pass
        try:
            print(value)
        except Exception:
            pass

    def section(self, title):
        self.write("")
        self.write("=" * 78)
        self.write(title)
        self.write("=" * 78)


def safe_property(obj, name):
    record = {"name": name}
    try:
        value = getattr(obj, name)
        if callable(value):
            record.update({"status": "CALLABLE", "type": type_name(value)})
        else:
            record.update({
                "status": "VALUE",
                "type": type_name(value),
                "value": short_repr(value),
                "raw": value if isinstance(value, (str, int, float, bool)) else None,
            })
    except Exception as error:
        record.update({"status": "ERROR", "error": exception_info(error)})
    return record


def filtered_members(obj):
    try:
        names = dir(obj)
    except Exception:
        return []
    return sorted({
        name for name in names
        if any(token in name.lower() for token in MEMBER_TOKENS)
    })


def simple_member_values(obj, names):
    return [
        record for record in (safe_property(obj, name) for name in names)
        if record.get("status") == "VALUE"
    ]


def collection_count(value):
    try:
        return int(value.Count)
    except Exception:
        pass
    try:
        return len(list(value))
    except Exception:
        return None


def string_attribute(nx_object, name):
    if nx_object is None:
        return ""
    try:
        return text(nx_object.GetStringAttribute(name))
    except Exception:
        pass
    try:
        attr = nx_object.GetUserAttribute(
            name, NXOpen.NXObject.AttributeType.String, -1
        )
        return text(attr.StringValue)
    except Exception:
        return ""


def part_identity(part):
    number = (
        string_attribute(part, "DB_PART_NO")
        or string_attribute(part, "PART_NUMBER")
        or string_attribute(part, "ITEM_ID")
        or string_attribute(part, "ITEM_NUMBER")
    )
    revision = (
        string_attribute(part, "DB_PART_REV")
        or string_attribute(part, "REVISION")
        or string_attribute(part, "ITEM_REVISION")
        or string_attribute(part, "REVISION_ID")
    )
    return number, revision


def attribute_value(attr):
    values = {}
    for name in (
        "StringValue", "IntegerValue", "RealValue", "BooleanValue",
        "TimeValue", "ReferenceValue", "Value",
    ):
        try:
            value = getattr(attr, name)
            if not callable(value):
                values[name] = short_repr(value)
        except Exception:
            pass
    return values or {"repr": short_repr(attr)}


def user_attributes(part):
    result = {"status": "NOT_AVAILABLE", "attributes": []}
    if part is None:
        return result
    try:
        attrs = list(part.GetUserAttributes())
    except Exception as error:
        result.update({"status": "ERROR", "error": exception_info(error)})
        return result

    result["status"] = "SUCCESS"
    result["total_count"] = len(attrs)
    for attr in attrs[:MAX_ATTRIBUTES]:
        title = ""
        category = ""
        for name in ("Title", "Name"):
            try:
                title = text(getattr(attr, name))
                if title:
                    break
            except Exception:
                pass
        try:
            category = text(attr.Category)
        except Exception:
            pass

        joined = (title + " " + category).upper()
        if not any(token in joined for token in ATTRIBUTE_TOKENS):
            continue

        try:
            attr_type = short_repr(attr.Type)
        except Exception:
            attr_type = ""

        result["attributes"].append({
            "title": title,
            "category": category,
            "type": attr_type,
            "values": attribute_value(attr),
        })
    return result


def session_managed_mode(session):
    try:
        value = session.IsManagedMode
        return {
            "member_type": type_name(value),
            "callable": callable(value),
            "value": bool(value() if callable(value) else value),
        }
    except Exception as error:
        return {"error": exception_info(error)}


def environment_info():
    safe = {}
    for name in SAFE_ENV_VARS:
        if name not in os.environ:
            continue
        safe[name] = (
            "<redacted>"
            if any(token in name.upper() for token in SENSITIVE_TOKENS)
            else os.environ.get(name, "")
        )
    names = sorted({
        name for name in os.environ
        if any(token in name.upper() for token in RELATED_ENV_TOKENS)
    })
    return {"safe_values": dict(sorted(safe.items())), "related_names_only": names}


def nxopen_inventory():
    try:
        names = dir(NXOpen)
    except Exception:
        return []
    tokens = ("cloud", "collab", "database", "dm", "part", "pdm", "search", "team", "web")
    return sorted({
        name for name in names
        if any(token in name.lower() for token in tokens)
    })


def manager_info(session, name):
    result = {"property_name": name, "status": "NOT_AVAILABLE"}
    try:
        manager = getattr(session, name)
    except Exception as error:
        result.update({"status": "ERROR", "error": exception_info(error)})
        return result
    if manager is None:
        result["status"] = "NONE"
        return result

    members = filtered_members(manager)
    result.update({
        "status": "SUCCESS",
        "type": type_name(manager),
        "members": members,
        "simple_values": simple_member_values(manager, members),
    })
    return result


def uf_call(target, method_name, args):
    result = {"method": method_name, "args": [short_repr(arg) for arg in args]}
    try:
        method = getattr(target, method_name)
    except Exception as error:
        result.update({"status": "NOT_AVAILABLE", "error": exception_info(error)})
        return result
    if not callable(method):
        result.update({"status": "NOT_CALLABLE", "value": short_repr(method)})
        return result
    try:
        value = method(*args)
        result.update({
            "status": "SUCCESS",
            "return_type": type_name(value),
            "value": short_repr(value, 3000),
        })
        if isinstance(value, (str, int, float, bool)):
            result["raw_value"] = value
    except Exception as error:
        result.update({"status": "ERROR", "error": exception_info(error)})
    return result


def uf_info(part):
    result = {"status": "NOT_AVAILABLE", "part_calls": [], "ugmgr": {}}
    try:
        uf = NXOpen.UF.UFSession.GetUFSession()
    except Exception as error:
        result.update({"status": "ERROR", "error": exception_info(error)})
        return result

    result.update({
        "status": "SUCCESS",
        "type": type_name(uf),
        "members": [
            name for name in dir(uf)
            if any(token in name.lower() for token in ("part", "ugmgr", "pdm", "session"))
        ],
    })

    tag = None
    try:
        tag = part.Tag if part is not None else None
    except Exception:
        pass

    try:
        uf_part = uf.Part
        result["part_type"] = type_name(uf_part)
        result["part_members"] = filtered_members(uf_part)
        if tag is not None:
            result["part_calls"].append(uf_call(uf_part, "AskPartName", [tag]))
            result["part_calls"].append(uf_call(uf_part, "AskPartName2", [tag]))
        result["part_calls"].append(uf_call(uf_part, "AskDisplayPart", []))
        result["part_calls"].append(uf_call(uf_part, "AskWorkPart", []))
    except Exception as error:
        result["part_error"] = exception_info(error)

    ugmgr = None
    property_name = ""
    for name in ("Ugmgr", "UGMgr", "UgmgrSession"):
        try:
            candidate = getattr(uf, name)
            if candidate is not None:
                ugmgr = candidate
                property_name = name
                break
        except Exception:
            pass

    if ugmgr is None:
        result["ugmgr"] = {"status": "NOT_AVAILABLE"}
        return result

    calls = [uf_call(ugmgr, "IsActive", [])]
    if tag is not None:
        for name in (
            "AskDatabasePartName", "AskPartName", "AskPartNumber",
            "AskPartRevision", "AskPartRevisionId", "AskPartUid",
            "AskUidFromTag",
        ):
            calls.append(uf_call(ugmgr, name, [tag]))

    result["ugmgr"] = {
        "status": "SUCCESS",
        "property_name": property_name,
        "type": type_name(ugmgr),
        "members": filtered_members(ugmgr),
        "calls": calls,
    }
    return result


def part_record(part, display, work, include_uf=False):
    if part is None:
        return None
    properties = [
        safe_property(part, name)
        for name in (
            "Name", "Leaf", "FullPath", "PartName", "JournalIdentifier",
            "Tag", "IsModified", "IsFullyLoaded", "LoadState", "OwningPart",
        )
    ]
    number, revision = part_identity(part)
    members = filtered_members(part)
    result = {
        "type": type_name(part),
        "identity": list(identity(part)),
        "is_display_part": part is display,
        "is_work_part": part is work,
        "part_number": number,
        "revision": revision,
        "properties": properties,
        "filtered_members": members,
        "simple_member_values": (
            simple_member_values(part, members) if include_uf else []
        ),
        "user_attributes": (
            user_attributes(part)
            if include_uf
            else {"status": "SKIPPED_FOR_NON_ACTIVE_PART", "attributes": []}
        ),
    }
    try:
        result["drawing_sheet_count"] = collection_count(part.DrawingSheets)
    except Exception as error:
        result["drawing_sheet_error"] = exception_info(error)
    try:
        result["body_count"] = collection_count(part.Bodies)
    except Exception as error:
        result["body_count_error"] = exception_info(error)
    if include_uf:
        result["uf"] = uf_info(part)
    return result


def property_text(part, name):
    try:
        return text(getattr(part, name))
    except Exception:
        return ""


def candidate_identifiers(display_part, display_uf):
    candidates = []

    def add(source, value):
        value = text(value)
        if not value or any(item["value"] == value for item in candidates):
            return
        candidates.append({"source": source, "value": value})

    for name in ("FullPath", "PartName", "JournalIdentifier", "Leaf", "Name"):
        add("display_part." + name, property_text(display_part, name))

    for name in (
        "DB_FULL_NAME", "DB_PART_NAME", "DB_DATASET_NAME", "DB_MODEL_NAME",
        "DB_PART_NO", "PART_NUMBER", "ITEM_ID",
    ):
        add("attribute." + name, string_attribute(display_part, name))

    for call in display_uf.get("part_calls", []):
        raw_value = call.get("raw_value")
        if call.get("status") == "SUCCESS" and isinstance(raw_value, str):
            add("UF.Part." + call.get("method", ""), raw_value)

    for call in display_uf.get("ugmgr", {}).get("calls", []):
        raw_value = call.get("raw_value")
        if call.get("status") == "SUCCESS" and isinstance(raw_value, str):
            add("UF.Ugmgr." + call.get("method", ""), raw_value)
    return candidates


def drawing_index(part):
    for name in ("Name", "Leaf", "FullPath", "PartName", "JournalIdentifier"):
        match = DRAWING_RE.search(property_text(part, name).upper())
        if match:
            try:
                return int(match.group(1))
            except Exception:
                pass
    return None


def generated_at_db(display_part):
    number, revision = part_identity(display_part)
    index = drawing_index(display_part)
    if not number or not revision or index is None:
        return ""
    dataset = "{0}-{1}-dwg{2}".format(number, revision, index)
    return "@DB/{0}/{1}/UGPART/{2}".format(number, revision, dataset)


def restore_parts(session, display, work):
    try:
        if display is not None:
            result = session.Parts.SetDisplay(display, False, True)
            if isinstance(result, (tuple, list)) and len(result) > 1:
                dispose(result[1])
    except Exception:
        pass
    try:
        if work is not None:
            session.Parts.SetWork(work)
    except Exception:
        pass


def close_test_part(part, session, display, work, logger):
    restore_parts(session, display, work)
    try:
        part.Close(
            NXOpen.BasePart.CloseWholeTree.FalseValue,
            NXOpen.BasePart.CloseModified.CloseModified,
            None,
        )
        logger.write("    Closed part newly loaded by test.")
    except Exception as error:
        logger.write("    WARNING closing test-loaded part: {0}".format(exception_info(error)))


def reopen_test(
    session, source, candidate, preloaded, display, work, logger
):
    result = {"source": source, "candidate": candidate}
    logger.write("  {0}".format(source))
    logger.write("    Candidate: {0}".format(candidate))
    part = None
    status = None
    try:
        part, status = unwrap_open(session.Parts.OpenBase(candidate))
        result.update({
            "status": "SUCCESS",
            "returned_type": type_name(part),
            "returned_identity": list(identity(part)),
            "returned_name": property_text(part, "Name") if part else "",
            "returned_full_path": property_text(part, "FullPath") if part else "",
            "load_status_type": type_name(status),
        })
        was_preloaded = identity(part) in preloaded if part is not None else False
        result["was_preloaded"] = was_preloaded
        logger.write("    SUCCESS: {0}".format(result["returned_name"] or result["returned_identity"]))
        if part is not None and not was_preloaded:
            result["newly_loaded_by_test"] = True
            close_test_part(part, session, display, work, logger)
        else:
            result["newly_loaded_by_test"] = False
    except Exception as error:
        result.update({"status": "ERROR", "error": exception_info(error)})
        logger.write("    ERROR: {0}".format(result["error"]))
    finally:
        dispose(status)
    return result


def controlled_reopen_tests(session, display_part, display_uf, logger):
    result = {"tests": []}
    if display_part is None:
        result["status"] = "SKIPPED_NO_DISPLAY_PART"
        return result

    original_display = session.Parts.Display
    original_work = session.Parts.Work
    try:
        loaded = list(session.Parts)
    except Exception:
        loaded = []
    preloaded = {identity(part) for part in loaded}

    candidates = []
    if RUN_ACTIVE_IDENTIFIER_REOPEN_TESTS:
        candidates.extend(candidate_identifiers(display_part, display_uf))
    if RUN_GENERATED_AT_DB_TEST:
        value = generated_at_db(display_part)
        if value and not any(item["value"] == value for item in candidates):
            candidates.append({"source": "generated.classic_at_db", "value": value})

    if not candidates:
        result["status"] = "SKIPPED_NO_IDENTIFIERS"
        return result

    result["status"] = "EXECUTED"
    for item in candidates[:MAX_REOPEN_CANDIDATES]:
        result["tests"].append(reopen_test(
            session, item["source"], item["value"], preloaded,
            original_display, original_work, logger,
        ))
    restore_parts(session, original_display, original_work)
    return result


def write_txt(path, lines):
    with open(path, "w", encoding="utf-8-sig", newline="") as handle:
        handle.write("\n".join(str(line) for line in lines))
        handle.write("\n")


def write_json(path, payload):
    with open(path, "w", encoding="utf-8", newline="") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2, default=str)


def log_part(logger, title, record):
    logger.section(title)
    if record is None:
        logger.write("<none>")
        return
    logger.write("Type: {0}".format(record.get("type", "")))
    logger.write("Identity: {0}".format(record.get("identity", "")))
    logger.write("Part number: {0}".format(record.get("part_number", "")))
    logger.write("Revision: {0}".format(record.get("revision", "")))
    logger.write("Drawing sheets: {0}".format(record.get("drawing_sheet_count", "")))
    logger.write("Bodies: {0}".format(record.get("body_count", "")))
    for prop in record.get("properties", []):
        if prop.get("status") == "VALUE":
            logger.write("{0}: {1}".format(prop["name"], prop.get("value", "")))
    for attr in record.get("user_attributes", {}).get("attributes", []):
        logger.write(
            "ATTR {0} [{1}] = {2}".format(
                attr.get("title", ""), attr.get("category", ""), attr.get("values", {})
            )
        )
    uf = record.get("uf", {})
    logger.write("UF status: {0}".format(uf.get("status", "")))
    for call in uf.get("part_calls", []):
        logger.write(
            "UF.Part.{0}: {1} {2}".format(
                call.get("method", ""), call.get("status", ""),
                call.get("value", call.get("error", "")),
            )
        )
    ugmgr = uf.get("ugmgr", {})
    logger.write("UF UGMGR status: {0}".format(ugmgr.get("status", "")))
    for call in ugmgr.get("calls", []):
        logger.write(
            "UF.Ugmgr.{0}: {1} {2}".format(
                call.get("method", ""), call.get("status", ""),
                call.get("value", call.get("error", "")),
            )
        )


def main():
    session = NXOpen.Session.GetSession()
    logger = Logger(session)
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    folder = output_root()
    os.makedirs(folder, exist_ok=True)
    txt_path = os.path.join(folder, "NX_X_IDENTITY_PROBE_{0}.txt".format(timestamp))
    json_path = os.path.join(folder, "NX_X_IDENTITY_PROBE_{0}.json".format(timestamp))
    payload = {
        "probe": {
            "name": "Journal 08 - NX X / Teamcenter Drawing Identity Probe",
            "timestamp": timestamp,
            "txt_path": txt_path,
            "json_path": json_path,
        }
    }

    try:
        logger.section("NX X / TEAMCENTER DRAWING IDENTITY PROBE")
        logger.write("Run while a manually opened drawing is the active display part.")
        logger.write("TXT: {0}".format(txt_path))
        logger.write("JSON: {0}".format(json_path))

        payload["runtime"] = {
            "python_version": sys.version,
            "python_executable": sys.executable,
            "platform": platform.platform(),
            "machine": platform.machine(),
            "process_id": os.getpid(),
            "nxopen_module": short_repr(getattr(NXOpen, "__file__", "")),
            "current_directory": os.getcwd(),
        }
        logger.section("RUNTIME")
        for name, value in payload["runtime"].items():
            logger.write("{0}: {1}".format(name, value))

        session_members = filtered_members(session)
        payload["session"] = {
            "type": type_name(session),
            "managed_mode": session_managed_mode(session),
            "filtered_members": session_members,
            "simple_values": simple_member_values(session, session_members),
        }
        logger.section("SESSION")
        logger.write("Session.IsManagedMode: {0}".format(payload["session"]["managed_mode"]))

        payload["environment"] = environment_info()
        logger.section("ENVIRONMENT")
        for name, value in payload["environment"]["safe_values"].items():
            logger.write("{0}={1}".format(name, value))
        logger.write(
            "Other related names only: {0}".format(
                ", ".join(payload["environment"]["related_names_only"])
            )
        )

        payload["nxopen_inventory"] = nxopen_inventory()
        logger.section("NXOPEN PDM / CLOUD INVENTORY")
        logger.write(", ".join(payload["nxopen_inventory"]))

        manager_names = (
            "PdmSession", "PdmManager", "PdmSearchManager", "WebAppSession",
            "CollaborativeContentManager", "CollaborationApplicationSession",
            "CloudDmSession", "CloudDMSession",
        )
        payload["session_managers"] = [manager_info(session, name) for name in manager_names]
        for manager in payload["session_managers"]:
            logger.write(
                "{0}: {1} {2}".format(
                    manager["property_name"], manager["status"], manager.get("type", "")
                )
            )
            if manager.get("members"):
                logger.write("  Members: {0}".format(", ".join(manager["members"])))

        display = session.Parts.Display
        work = session.Parts.Work
        display_uf = uf_info(display)
        display_record = part_record(display, display, work, include_uf=True)
        work_record = part_record(work, display, work, include_uf=True)
        payload["active_parts"] = {"display": display_record, "work": work_record}
        log_part(logger, "DISPLAY PART", display_record)
        log_part(logger, "WORK PART", work_record)

        if display_record is None:
            logger.write("WARNING: No display part. Open a drawing and rerun.")
        elif not display_record.get("drawing_sheet_count"):
            logger.write("WARNING: Display part has no drawing sheets. Open a known drawing and rerun.")

        logger.section("ALL LOADED PARTS")
        try:
            loaded_parts = list(session.Parts)
            payload["loaded_parts"] = {
                "status": "SUCCESS",
                "total_count": len(loaded_parts),
                "parts": [
                    part_record(part, display, work)
                    for part in loaded_parts[:MAX_LOADED_PARTS]
                ],
            }
            logger.write("Loaded part count: {0}".format(len(loaded_parts)))
            for index, record in enumerate(payload["loaded_parts"]["parts"], start=1):
                name = next(
                    (
                        prop.get("value", "")
                        for prop in record.get("properties", [])
                        if prop.get("name") == "Name" and prop.get("status") == "VALUE"
                    ),
                    "",
                )
                logger.write(
                    "[{0}] {1}{2}{3}".format(
                        index, name,
                        " [DISPLAY]" if record.get("is_display_part") else "",
                        " [WORK]" if record.get("is_work_part") else "",
                    )
                )
        except Exception as error:
            payload["loaded_parts"] = {
                "status": "ERROR",
                "error": exception_info(error),
                "parts": [],
            }
            logger.write("ERROR enumerating loaded parts: {0}".format(exception_info(error)))

        logger.section("CONTROLLED REOPEN TESTS")
        payload["reopen_tests"] = controlled_reopen_tests(
            session, display, display_uf, logger
        )

        logger.section("NEXT ACTION")
        logger.write("1. Upload the TXT and JSON probe outputs.")
        logger.write(
            "2. Record one successful manual open using NX > Tools > Journal > Record "
            "and upload the recorded journal."
        )
        logger.write("3. Journal 07 remains unchanged until the canonical open method is proven.")
        payload["probe"]["status"] = "SUCCESS"

    except Exception as error:
        payload["probe"]["status"] = "FAILED"
        payload["probe"]["error"] = exception_info(error)
        payload["probe"]["traceback"] = traceback.format_exc()
        logger.section("UNHANDLED PROBE ERROR")
        logger.write(payload["probe"]["error"])
        logger.write(payload["probe"]["traceback"])

    finally:
        logger.write("")
        logger.write("Probe complete.")
        logger.write("TXT: {0}".format(txt_path))
        logger.write("JSON: {0}".format(json_path))
        try:
            write_json(json_path, payload)
        except Exception as error:
            logger.write("ERROR writing JSON: {0}".format(exception_info(error)))
        try:
            write_txt(txt_path, logger.lines)
        except Exception:
            pass


if __name__ == "__main__":
    main()
