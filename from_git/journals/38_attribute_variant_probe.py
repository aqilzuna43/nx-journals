"""Journal 38 - Attribute Variant Probe (NX 2506, read-only)

Answers one question about the native NX part file BEFORE it is exported and
pushed to a customer Teamcenter: does the file already carry BOTH name
variants of a business attribute - the stored title (e.g. "Commodity_Code")
AND the attribute-template alias (e.g. "Commodity Code") - or only the
canonical title?

Background: the deployed NX part attribute template (NXPartAttribute_FZ.xml)
defines every WAE business attribute with a display alias plus an internal
title.  Journal 04/05 read and write exactly one canonical title per field
under category WAEItem, so the journals themselves never create a second
variant.  If the customer's Teamcenter shows duplicate attribute entries with
identical values, one of two things is true:

  A. The exported native file already contained both variants (the probe
     reports NX_FILE_CARRIES_DUPLICATES).  Inspect the per-attribute flags in
     the JSON: PdmBased=True or OwnedBySystem=True identifies attributes that
     were mirrored into the part by a Teamcenter round trip rather than
     written locally.
  B. The file contains only canonical titles (the probe reports
     NX_FILE_CLEAN).  The duplication is then introduced on the customer's
     side - their TC conversion/import mapping reads the alias and the title
     (or maps one NX attribute into two TC properties), and the exact
     category/title contract in this JSON is the evidence to hand them.

Scope: active work part plus every unique assembly prototype (suppressed
occurrences skipped).  The journal is strictly read-only: no attribute
writes, no saves, no checkout, no Teamcenter calls.

Output: Listing Window summary plus, under
NX_ATTRIBUTE_VARIANT_PROBE on the desktop (or NX_JOURNALS_IO_DIR):
  J38_VARIANT_PROBE_<root>_<timestamp>.json   - full evidence for analysis
  J38_VARIANT_PROBE_<root>_<timestamp>.csv    - flat attribute list

Run via: NX > Tools > Journal > Play
"""

import csv
import datetime
import json
import os
import re
import traceback

import NXOpen


BUILD = "J38-NX2506-ATTRIBUTE-VARIANT-PROBE-V1"
OUTPUT_FOLDER = "NX_ATTRIBUTE_VARIANT_PROBE"
EXPECTED_CATEGORY = "WAEItem"

# (canonical_title, template_alias) pairs exactly as deployed in
# NXPartAttribute_FZ.xml.  The alias is display-only; the title is the
# attribute name stored in the part file and the only name Journal 04/05
# use.
CANONICAL_FIELDS = [
    ("Unit_Of_Measure", "UOM"),
    ("MFG", "Mfr. Name"),
    ("MPN", "Mfr. Part Number"),
    ("Stocking_Type", "Stocking Type"),
    ("WAE_VERSION", "WAE Version"),
    ("WAE_Hazardous", "WAE Hazardous"),
    ("COMMODITYTYPE", "Commodity Type"),
    ("Commodity_Code", "Commodity Code"),
    ("Export_Control_Number", "Export Control Number"),
    ("Country_of_Origin", "Country of Origin"),
    ("Temperature_Sensitive", "Temperature Sensitive"),
    ("LIFED", "Shelf Life Limited"),
    ("Serviceable_item_flag", "Serviceable item flag"),
    ("SERIAL_NUMBERED_PART", "Traceability"),
    ("COMPONENT_CLASS", "Part Classification"),
    ("NX_FINISH", "FINISH"),
]

_INVALID_FILENAME_CHARS = '<>:"/\\|?*'


def _text(value):
    return "" if value is None else str(value)


def clean(value):
    return _text(value).strip()


def error_text(error):
    code = clean(getattr(error, "ErrorCode", ""))
    suffix = " [{0}]".format(code) if code else ""
    return "{0}{1}".format(clean(error) or type(error).__name__, suffix)


def dispose(value):
    if value is None:
        return
    for method_name in ("Dispose", "FreeResource", "Destroy"):
        method = getattr(value, method_name, None)
        if callable(method):
            try:
                method()
            except Exception:
                pass
            return


def log_line(session, message):
    text = str(message)
    try:
        window = session.ListingWindow
        window.Open()
        for line in text.splitlines() or [""]:
            window.WriteFullline(line)
    except Exception:
        pass


def desktop_folder():
    for env_name in ("USERPROFILE", "HOME"):
        base = clean(os.environ.get(env_name))
        if base:
            desktop = os.path.join(base, "Desktop")
            if os.path.isdir(desktop):
                return desktop
            return base
    return os.path.abspath(os.curdir)


def io_root():
    configured = clean(os.environ.get("NX_JOURNALS_IO_DIR"))
    return os.path.abspath(
        os.path.expanduser(configured or desktop_folder())
    )


def clean_filename_token(value, fallback="UNKNOWN"):
    text = clean(value)
    if not text:
        return fallback
    result = "".join(
        "_" if char in _INVALID_FILENAME_CHARS or ord(char) < 32 else char
        for char in text
    ).strip(" .")
    return result or fallback


def enum_name(value):
    if value is None:
        return ""
    name = getattr(value, "name", None)
    return _text(name if name is not None else value).split(".")[-1]


def attribute_value(info):
    kind = enum_name(getattr(info, "Type", ""))
    numeric_kind = getattr(info, "Type", None)
    if kind in ("String", "5") or numeric_kind == 5:
        return getattr(info, "StringValue", ""), "String"
    if kind in ("Real", "Number", "4") or numeric_kind == 4:
        return getattr(info, "RealValue", None), "Number"
    if kind in ("Integer", "3") or numeric_kind == 3:
        return getattr(info, "IntegerValue", None), "Integer"
    if kind in ("Boolean", "1") or numeric_kind == 1:
        return getattr(info, "BooleanValue", None), "Boolean"
    if kind in ("Time", "6") or numeric_kind == 6:
        return getattr(info, "TimeValue", None), "Time"
    if kind in ("Reference", "7") or numeric_kind == 7:
        return getattr(info, "ReferenceValue", None), "Reference"
    return getattr(info, "StringValue", ""), kind or "String"


def get_string_attribute(nx_object, name):
    try:
        return clean(nx_object.GetStringAttribute(name))
    except Exception:
        pass
    try:
        attribute = nx_object.GetUserAttribute(
            name, NXOpen.NXObject.AttributeType.String, -1
        )
        return clean(attribute.StringValue)
    except Exception:
        return ""


def safe_part_name(part):
    for property_name in ("Name", "Leaf", "FullPath"):
        try:
            value = clean(getattr(part, property_name))
            if value:
                return value
        except Exception:
            pass
    return "UNKNOWN"


def part_identity(part):
    return {
        "number": (
            get_string_attribute(part, "DB_PART_NO")
            or get_string_attribute(part, "PART_NUMBER")
            or get_string_attribute(part, "ITEM_ID")
        ),
        "revision": (
            get_string_attribute(part, "DB_PART_REV")
            or get_string_attribute(part, "REVISION")
            or get_string_attribute(part, "ITEM_REVISION")
        ),
        "name": safe_part_name(part),
        "path": clean(getattr(part, "FullPath", "")),
        "pdm_part": "@DB" in clean(getattr(part, "FullPath", "")),
    }


def object_key(nx_object):
    for name in ("Tag", "JournalIdentifier", "FullPath", "Leaf"):
        try:
            value = getattr(nx_object, name)
        except Exception:
            continue
        if value:
            return "{0}:{1}".format(name, value)
    return "ID:{0}".format(id(nx_object))


def dump_attributes(nx_object):
    """Return the full attribute inventory with all read-only flags."""
    iterator = None
    result = []
    try:
        if callable(getattr(nx_object, "CreateAttributeIterator", None)):
            iterator = nx_object.CreateAttributeIterator()
            iterator.SetIncludeAlsoUnset(True)
            infos = list(nx_object.GetUserAttributes(iterator))
        else:
            infos = list(nx_object.GetUserAttributes())
        for info in infos:
            value, kind = attribute_value(info)
            result.append(
                {
                    "category": clean(getattr(info, "Category", "")),
                    "title": clean(getattr(info, "Title", "")),
                    "type": kind,
                    "value": value,
                    "unset": bool(getattr(info, "Unset", False)),
                    "locked": bool(getattr(info, "Locked", False)),
                    "owned_by_system": bool(
                        getattr(info, "OwnedBySystem", False)
                    ),
                    "pdm_based": bool(getattr(info, "PdmBased", False)),
                    "required": bool(getattr(info, "Required", False)),
                    "not_saved": bool(getattr(info, "NotSaved", False)),
                }
            )
    except Exception as error:
        result.append(
            {
                "category": "",
                "title": "ATTRIBUTE_ENUMERATION_FAILED",
                "type": "",
                "value": error_text(error),
                "unset": False,
                "locked": False,
                "owned_by_system": False,
                "pdm_based": False,
                "required": False,
                "not_saved": False,
            }
        )
    finally:
        dispose(iterator)
    result.sort(key=lambda item: (item["category"].lower(), item["title"].lower()))
    return result


def unique_prototypes(work_part):
    """Work part first, then every unique prototype in occurrence order."""
    ordered = []
    seen = set()

    def add(part):
        key = object_key(part)
        if key in seen:
            return
        seen.add(key)
        ordered.append(part)

    def walk(component):
        try:
            children = list(component.GetChildren())
        except Exception:
            return
        for child in children:
            try:
                if bool(child.IsSuppressed):
                    continue
            except Exception:
                pass
            try:
                prototype = child.Prototype
            except Exception:
                continue
            if prototype is None:
                continue
            add(prototype)
            walk(child)

    add(work_part)
    try:
        root = work_part.ComponentAssembly.RootComponent
    except Exception:
        root = None
    if root is not None:
        walk(root)
    return ordered


def normalize_title_key(value):
    """Collapse case, spaces, underscores, dots, hyphens: 'Commodity_Code'
    and 'Commodity Code' share one key; 'MFG' and 'Mfr. Name' do not."""
    return re.sub(r"[^0-9a-z]+", "", clean(value).casefold())


def normalize_value(value):
    text = clean(value)
    try:
        return format(float(text), ".12g")
    except (TypeError, ValueError):
        pass
    return " ".join(text.split()).casefold()


def values_equal(left, right):
    if normalize_value(left) == normalize_value(right):
        return True
    try:
        number_left = float(clean(left))
        number_right = float(clean(right))
    except (TypeError, ValueError):
        return False
    scale = max(abs(number_left), abs(number_right), 1e-30)
    return abs(number_left - number_right) <= 1e-12 * scale


def value_text(value):
    if value is None:
        return ""
    if isinstance(value, datetime.datetime):
        return value.isoformat()
    return clean(value)


def analyze_known_fields(attributes):
    """Census of the 16 canonical WAE fields: which variants are present."""
    by_title = {}
    for info in attributes:
        by_title.setdefault(clean(info["title"]).casefold(), []).append(info)

    fields = []
    for title, alias in CANONICAL_FIELDS:
        title_key = title.casefold()
        alias_key = alias.casefold()
        title_hits = list(by_title.get(title_key, []))
        alias_hits = (
            [] if alias_key == title_key else list(by_title.get(alias_key, []))
        )
        title_value = value_text(title_hits[0]["value"]) if title_hits else ""
        alias_value = value_text(alias_hits[0]["value"]) if alias_hits else ""
        if title_hits and alias_hits:
            status = "BOTH_VARIANTS"
        elif title_hits:
            status = "TITLE_ONLY"
        elif alias_hits:
            status = "ALIAS_ONLY"
        else:
            status = "NOT_PRESENT"
        if title_hits and alias_hits:
            relation = (
                "IDENTICAL_DUPLICATE"
                if values_equal(title_value, alias_value)
                else "CONFLICTING_DUPLICATE"
            )
        else:
            relation = ""
        fields.append(
            {
                "canonical_title": title,
                "template_alias": alias,
                "expected_category": EXPECTED_CATEGORY,
                "title_present": bool(title_hits),
                "alias_present": bool(alias_hits),
                "title_category": title_hits[0]["category"] if title_hits else "",
                "alias_category": alias_hits[0]["category"] if alias_hits else "",
                "title_value": title_value,
                "alias_value": alias_value,
                "title_pdm_based": bool(title_hits[0]["pdm_based"]) if title_hits else False,
                "alias_pdm_based": bool(alias_hits[0]["pdm_based"]) if alias_hits else False,
                "status": status,
                "relation": relation,
            }
        )
    return fields


def analyze_generic_groups(attributes, claimed_titles):
    """Normalized-title collisions among attributes NOT claimed by the
    canonical census - discovers unknown variants (e.g. lowercase names)."""
    groups = {}
    for info in attributes:
        title = clean(info["title"])
        if not title or title == "ATTRIBUTE_ENUMERATION_FAILED":
            continue
        if title.casefold() in claimed_titles:
            continue
        groups.setdefault(normalize_title_key(title), []).append(info)

    findings = []
    for key, members in sorted(groups.items()):
        distinct_titles = sorted(
            {clean(item["title"]) for item in members},
            key=lambda item: item.casefold(),
        )
        if len(distinct_titles) < 2:
            continue
        relations = set()
        for index in range(1, len(members)):
            relations.add(
                "IDENTICAL_DUPLICATE"
                if values_equal(members[0]["value"], members[index]["value"])
                else "CONFLICTING_DUPLICATE"
            )
        findings.append(
            {
                "normalized_key": key,
                "titles": distinct_titles,
                "relation": (
                    "IDENTICAL_DUPLICATE"
                    if relations == {"IDENTICAL_DUPLICATE"}
                    else "CONFLICTING_DUPLICATE"
                ),
                "members": [
                    {
                        "category": item["category"],
                        "title": item["title"],
                        "value": value_text(item["value"]),
                        "pdm_based": item["pdm_based"],
                        "owned_by_system": item["owned_by_system"],
                    }
                    for item in members
                ],
            }
        )
    return findings


def analyze_part(part):
    identity = part_identity(part)
    attributes = dump_attributes(part)
    fields = analyze_known_fields(attributes)

    claimed = set()
    for field in fields:
        claimed.add(field["canonical_title"].casefold())
        claimed.add(field["template_alias"].casefold())
    variant_groups = analyze_generic_groups(attributes, claimed)

    duplicated_fields = [
        field for field in fields if field["status"] in ("BOTH_VARIANTS", "ALIAS_ONLY")
    ]
    identical = [
        field for field in duplicated_fields if field["relation"] == "IDENTICAL_DUPLICATE"
    ]
    conflicting = [field for field in duplicated_fields if field["relation"] == "CONFLICTING_DUPLICATE"]
    verdict = "NX_FILE_CARRIES_DUPLICATES" if duplicated_fields else "NX_FILE_CLEAN"

    return {
        "part": identity,
        "tag": _text(getattr(part, "Tag", "")),
        "attribute_count": len(attributes),
        "attributes": attributes,
        "fields": fields,
        "variant_groups": variant_groups,
        "duplicate_field_count": len(duplicated_fields),
        "identical_duplicate_fields": [field["canonical_title"] for field in identical],
        "conflicting_duplicate_fields": [field["canonical_title"] for field in conflicting],
        "verdict": verdict,
    }


def build_report(session, work_part, run_datetime=None):
    now = run_datetime or datetime.datetime.now()
    managed_mode = getattr(session, "IsManagedMode", None)
    if callable(managed_mode):
        try:
            managed_mode = managed_mode()
        except Exception:
            managed_mode = None

    parts = [analyze_part(part) for part in unique_prototypes(work_part)]
    carrying = [entry for entry in parts if entry["verdict"] == "NX_FILE_CARRIES_DUPLICATES"]

    return {
        "build": BUILD,
        "generated_at": now.strftime("%Y-%m-%d %H:%M:%S"),
        "hostname": clean(os.environ.get("COMPUTERNAME", "")),
        "session_is_managed_mode": _text(managed_mode),
        "work_part": part_identity(work_part),
        "parts_probed": len(parts),
        "parts_clean": len(parts) - len(carrying),
        "parts_carrying_duplicates": len(carrying),
        "parts": parts,
        "interpretation": {
            "NX_FILE_CLEAN": (
                "The native file carries only canonical attribute titles. "
                "The duplicates seen in the customer Teamcenter are created "
                "by their conversion/import mapping, not by this file. "
                "Hand them the canonical_title list with expected_category "
                "WAEItem from this JSON as the mapping contract."
            ),
            "NX_FILE_CARRIES_DUPLICATES": (
                "The native file already carries both name variants. Check "
                "the flags in parts[].attributes: pdm_based or "
                "owned_by_system attributes were mirrored in by a "
                "Teamcenter round trip; plain title+alias pairs were likely "
                "written by an older template or manual entry. The export "
                "then pushes both variants."
            ),
        },
    }


def write_outputs(report, io_dir=None, run_datetime=None):
    now = run_datetime or datetime.datetime.now()
    file_timestamp = now.strftime("%Y%m%d_%H%M%S")
    token = clean_filename_token(
        report["work_part"]["number"] or report["work_part"]["name"]
    )
    folder = os.path.join(io_dir or io_root(), OUTPUT_FOLDER)
    if not os.path.isdir(folder):
        os.makedirs(folder)
    json_path = os.path.join(
        folder, "J38_VARIANT_PROBE_{0}_{1}.json".format(token, file_timestamp)
    )
    csv_path = os.path.join(
        folder, "J38_VARIANT_PROBE_{0}_{1}.csv".format(token, file_timestamp)
    )

    with open(json_path, "w", encoding="utf-8") as handle:
        json.dump(report, handle, indent=2, ensure_ascii=False)

    headers = [
        "PART_NUMBER",
        "PART_REV",
        "PART_NAME",
        "CATEGORY",
        "TITLE",
        "TYPE",
        "VALUE",
        "UNSET",
        "LOCKED",
        "OWNED_BY_SYSTEM",
        "PDM_BASED",
        "REQUIRED",
        "NOT_SAVED",
        "KNOWN_FIELD",
        "VARIANT_ROLE",
        "PART_VERDICT",
    ]
    with open(csv_path, "w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(headers)
        for part in report["parts"]:
            known_roles = {}
            for field in part["fields"]:
                if field["title_present"]:
                    known_roles[field["canonical_title"].casefold()] = field["canonical_title"]
                if field["alias_present"]:
                    known_roles[field["template_alias"].casefold()] = field["canonical_title"]
            for info in part["attributes"]:
                role_key = clean(info["title"]).casefold()
                writer.writerow(
                    [
                        part["part"]["number"],
                        part["part"]["revision"],
                        part["part"]["name"],
                        info["category"],
                        info["title"],
                        info["type"],
                        value_text(info["value"]),
                        "Y" if info["unset"] else "N",
                        "Y" if info["locked"] else "N",
                        "Y" if info["owned_by_system"] else "N",
                        "Y" if info["pdm_based"] else "N",
                        "Y" if info["required"] else "N",
                        "Y" if info["not_saved"] else "N",
                        known_roles.get(role_key, ""),
                        "TITLE" if role_key in {
                            field["canonical_title"].casefold() for field in part["fields"]
                        } else ("ALIAS" if role_key in {
                            field["template_alias"].casefold() for field in part["fields"]
                        } else ""),
                        part["verdict"],
                    ]
                )
    return csv_path, json_path


def run(session, io_dir=None, run_datetime=None):
    work_part = session.Parts.Work
    if work_part is None:
        raise RuntimeError("No active work part. Open the part to probe first.")
    report = build_report(session, work_part, run_datetime=run_datetime)
    csv_path, json_path = write_outputs(
        report, io_dir=io_dir, run_datetime=run_datetime
    )
    return csv_path, json_path, report


def log_summary(session, report):
    log_line(session, "-" * 72)
    log_line(
        session,
        "Parts probed: {0}  clean: {1}  carrying duplicates: {2}".format(
            report["parts_probed"],
            report["parts_clean"],
            report["parts_carrying_duplicates"],
        ),
    )
    for part in report["parts"]:
        log_line(
            session,
            "[{0}] {1} {2} - attributes: {3} - verdict: {4}".format(
                part["part"]["number"] or part["part"]["name"],
                part["part"]["revision"],
                part["part"]["name"],
                part["attribute_count"],
                part["verdict"],
            ),
        )
        for field in part["fields"]:
            if field["status"] in ("BOTH_VARIANTS", "ALIAS_ONLY"):
                log_line(
                    session,
                    "    {0}: title '{1}'={2!r} / alias '{3}'={4!r} -> {5} {6}".format(
                        field["canonical_title"],
                        field["canonical_title"],
                        field["title_value"],
                        field["template_alias"],
                        field["alias_value"],
                        field["status"],
                        field["relation"],
                    ),
                )
        for group in part["variant_groups"]:
            log_line(
                session,
                "    unexpected variant group {0}: {1} -> {2}".format(
                    group["normalized_key"],
                    " / ".join(group["titles"]),
                    group["relation"],
                ),
            )
    log_line(session, "-" * 72)
    log_line(session, "Interpretation:")
    if report["parts_carrying_duplicates"]:
        log_line(session, "  " + report["interpretation"]["NX_FILE_CARRIES_DUPLICATES"])
    else:
        log_line(session, "  " + report["interpretation"]["NX_FILE_CLEAN"])
    log_line(
        session,
        "Run the probe on the exact part(s) you export; if the customer "
        "returns files, probe those too and compare.",
    )


def main():
    session = NXOpen.Session.GetSession()
    log_line(session, "=" * 72)
    log_line(session, "J38 ATTRIBUTE VARIANT PROBE (read-only)")
    log_line(session, "Build: " + BUILD)
    log_line(
        session,
        "Question: does the native file carry BOTH attribute name variants "
        "(title 'Commodity_Code' AND alias 'Commodity Code')?",
    )
    log_line(session, "=" * 72)
    try:
        csv_path, json_path, report = run(session)
        log_summary(session, report)
        log_line(session, "CSV: " + csv_path)
        log_line(session, "JSON: " + json_path)
        log_line(
            session,
            "Send the JSON back for analysis; it proves whether the duplicate "
            "variants are already in the pushed file or created by the "
            "customer TC conversion.",
        )
    except Exception as error:
        log_line(session, "J38 FAILED: " + error_text(error))
        log_line(session, traceback.format_exc())


if __name__ == "__main__":
    main()
