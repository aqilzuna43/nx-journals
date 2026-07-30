"""
Journal 18 - Active Work Part Surface Area

Measures every face of each direct traditional solid body owned by the active
NX work part. Hidden solid bodies are included. Sheet bodies and convergent
bodies are ignored and counted in the result.

The journal is intentionally read-only. It does not traverse assemblies, open
Teamcenter objects, create measurement features or expressions, write
attributes, or save parts.

Target: NX 2312 and NX X 2506 embedded Python
Run via: NX > Tools > Journal > Play
"""

import csv
import datetime
import os
import re
import traceback

import NXOpen


BUILD = "J18-NX2506-WORK-PART-SURFACE-AREA-V1"
OUTPUT_FOLDER = "NX_SURFACE_AREA"
MEASUREMENT_ACCURACY = 0.99
AREA_DECIMAL_PLACES = 4

RESULT_COLUMNS = (
    "RUN_TIMESTAMP",
    "JOURNAL_BUILD",
    "STATUS",
    "ROW_TYPE",
    "DB_PART_NO",
    "DB_PART_REV",
    "PART_NAME",
    "BODY_INDEX",
    "BODY_NAME",
    "BODY_TAG",
    "FACE_COUNT",
    "SURFACE_AREA_M2",
    "MEASUREMENT_ACCURACY",
    "INCLUDED_SOLID_BODY_COUNT",
    "SKIPPED_SHEET_BODY_COUNT",
    "SKIPPED_CONVERGENT_BODY_COUNT",
    "MESSAGE",
)

_INVALID_FILENAME_CHARS = '<>:"/\\|?*'


def clean(value):
    return "" if value is None else str(value).strip()


def error_text(error):
    code = clean(getattr(error, "ErrorCode", ""))
    suffix = " [{0}]".format(code) if code else ""
    return "{0}{1}".format(clean(error) or type(error).__name__, suffix)


def dispose(value):
    if value is None:
        return
    for method_name in ("Dispose", "FreeResource"):
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

    try:
        print(text)
    except Exception:
        pass


def desktop_folder():
    profile = clean(os.environ.get("USERPROFILE"))
    if profile:
        return os.path.join(profile, "Desktop")

    fallback = os.path.expanduser("~")
    if fallback and fallback != "~":
        return os.path.join(fallback, "Desktop")
    return os.getcwd()


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


def get_string_attribute(nx_object, name):
    try:
        return clean(nx_object.GetStringAttribute(name))
    except Exception:
        pass

    try:
        attribute = nx_object.GetUserAttribute(
            name,
            NXOpen.NXObject.AttributeType.String,
            -1,
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
    number = (
        get_string_attribute(part, "DB_PART_NO")
        or get_string_attribute(part, "PART_NUMBER")
        or get_string_attribute(part, "ITEM_ID")
    )
    revision = (
        get_string_attribute(part, "DB_PART_REV")
        or get_string_attribute(part, "REVISION")
        or get_string_attribute(part, "ITEM_REVISION")
    )
    return {
        "number": number,
        "revision": revision,
        "name": safe_part_name(part),
    }


def body_name(body, index):
    try:
        value = clean(body.Name)
        if value:
            return value
    except Exception:
        pass
    return "BODY_{0}".format(index)


def body_tag(body):
    try:
        return clean(body.Tag)
    except Exception:
        return ""


def body_flag(body, property_name):
    value = getattr(body, property_name)
    return bool(value() if callable(value) else value)


def classify_direct_bodies(work_part):
    included = []
    skipped_sheet = 0
    skipped_convergent = 0

    for body in list(work_part.Bodies):
        if body_flag(body, "IsConvergentBody"):
            skipped_convergent += 1
        elif body_flag(body, "IsSolidBody"):
            included.append(body)
        elif body_flag(body, "IsSheetBody"):
            skipped_sheet += 1

    return {
        "included": included,
        "skipped_sheet": skipped_sheet,
        "skipped_convergent": skipped_convergent,
    }


def normalized_unit_token(value):
    text = clean(value).upper()
    text = text.replace("²", "2").replace("^", "")
    return re.sub(r"[^A-Z0-9]", "", text)


def unit_tokens(unit):
    values = []
    for property_name in (
        "Name",
        "Symbol",
        "Abbreviation",
        "TypeName",
    ):
        try:
            values.append(
                normalized_unit_token(getattr(unit, property_name))
            )
        except Exception:
            pass
    return {value for value in values if value}


def unit_matches(unit, wanted_tokens):
    return bool(unit_tokens(unit).intersection(wanted_tokens))


def resolve_measure_unit(unit_collection, measure_name, object_names, tokens):
    for object_name in object_names:
        try:
            unit = unit_collection.FindObject(object_name)
            if unit is not None:
                return unit
        except Exception:
            pass

    try:
        candidates = list(unit_collection.GetMeasureTypes(measure_name))
    except Exception as error:
        raise RuntimeError(
            "NX could not enumerate {0} units: {1}".format(
                measure_name,
                error_text(error),
            )
        )

    for unit in candidates:
        if unit_matches(unit, tokens):
            return unit

    available = sorted(
        {
            token
            for unit in candidates
            for token in unit_tokens(unit)
        }
    )
    raise RuntimeError(
        "NX square-metre measurement units are unavailable for {0}. "
        "Available unit tokens: {1}".format(
            measure_name,
            ", ".join(available) or "<none>",
        )
    )


def resolve_square_metre_units(work_part):
    units = work_part.UnitCollection
    area_unit = resolve_measure_unit(
        units,
        "Area",
        ("SquareMeter", "SquareMetre"),
        {"SQUAREMETER", "SQUAREMETRE", "M2"},
    )
    length_unit = resolve_measure_unit(
        units,
        "Length",
        ("Meter", "Metre"),
        {"METER", "METRE", "M"},
    )
    return area_unit, length_unit


def measure_body_area_m2(
    work_part,
    body,
    area_unit,
    length_unit,
):
    faces = list(body.GetFaces())
    if not faces:
        raise RuntimeError("Solid body contains no measurable faces.")

    measurement = None
    try:
        measurement = work_part.MeasureManager.NewFaceProperties(
            area_unit,
            length_unit,
            MEASUREMENT_ACCURACY,
            faces,
        )
        area = float(measurement.Area)
        if area < 0.0:
            raise RuntimeError(
                "NX returned a negative surface area: {0}".format(area)
            )
        return area, len(faces)
    finally:
        dispose(measurement)


def area_text(value):
    if value is None:
        return ""
    return ("{0:." + str(AREA_DECIMAL_PLACES) + "f}").format(value)


def base_row(timestamp, identity, counts):
    return {
        "RUN_TIMESTAMP": timestamp,
        "JOURNAL_BUILD": BUILD,
        "STATUS": "",
        "ROW_TYPE": "",
        "DB_PART_NO": identity.get("number", ""),
        "DB_PART_REV": identity.get("revision", ""),
        "PART_NAME": identity.get("name", ""),
        "BODY_INDEX": "",
        "BODY_NAME": "",
        "BODY_TAG": "",
        "FACE_COUNT": "",
        "SURFACE_AREA_M2": "",
        "MEASUREMENT_ACCURACY": "{0:.2f}".format(
            MEASUREMENT_ACCURACY
        ),
        "INCLUDED_SOLID_BODY_COUNT": counts.get("included", 0),
        "SKIPPED_SHEET_BODY_COUNT": counts.get("skipped_sheet", 0),
        "SKIPPED_CONVERGENT_BODY_COUNT": counts.get(
            "skipped_convergent",
            0,
        ),
        "MESSAGE": "",
    }


def failure_total_row(timestamp, identity, counts, status, message):
    row = base_row(timestamp, identity, counts)
    row["STATUS"] = status
    row["ROW_TYPE"] = "TOTAL"
    row["MESSAGE"] = message
    return row


def calculate_surface_rows(work_part, timestamp):
    identity = part_identity(work_part)
    classified = classify_direct_bodies(work_part)
    bodies = classified["included"]
    counts = {
        "included": len(bodies),
        "skipped_sheet": classified["skipped_sheet"],
        "skipped_convergent": classified["skipped_convergent"],
    }

    if not bodies:
        return [
            failure_total_row(
                timestamp,
                identity,
                counts,
                "FAILED_NO_SOLID_BODIES",
                (
                    "The active work part contains no direct traditional "
                    "solid bodies."
                ),
            )
        ]

    try:
        area_unit, length_unit = resolve_square_metre_units(work_part)
    except Exception as error:
        return [
            failure_total_row(
                timestamp,
                identity,
                counts,
                "FAILED_UNIT_RESOLUTION",
                error_text(error),
            )
        ]
    rows = []
    raw_total = 0.0
    failures = []

    for index, body in enumerate(bodies, start=1):
        row = base_row(timestamp, identity, counts)
        row["ROW_TYPE"] = "BODY"
        row["BODY_INDEX"] = index
        row["BODY_NAME"] = body_name(body, index)
        row["BODY_TAG"] = body_tag(body)
        try:
            raw_area, face_count = measure_body_area_m2(
                work_part,
                body,
                area_unit,
                length_unit,
            )
            raw_total += raw_area
            row["STATUS"] = "SUCCESS"
            row["FACE_COUNT"] = face_count
            row["SURFACE_AREA_M2"] = area_text(raw_area)
        except Exception as error:
            message = error_text(error)
            failures.append(
                "{0}: {1}".format(row["BODY_NAME"], message)
            )
            row["STATUS"] = "FAILED_MEASUREMENT"
            row["MESSAGE"] = message
        rows.append(row)

    total = base_row(timestamp, identity, counts)
    total["ROW_TYPE"] = "TOTAL"
    if failures:
        total["STATUS"] = "FAILED_BODY_MEASUREMENT"
        total["MESSAGE"] = (
            "At least one included solid body could not be measured; "
            "the fail-closed total is blank. {0}"
        ).format(" | ".join(failures))
    else:
        total["STATUS"] = "SUCCESS"
        total["SURFACE_AREA_M2"] = area_text(raw_total)
        total["MESSAGE"] = (
            "Surface area includes every face of every direct traditional "
            "solid body in the active work part."
        )
    rows.append(total)
    return rows


def write_csv(path, rows):
    with open(path, "w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=RESULT_COLUMNS,
            extrasaction="ignore",
        )
        writer.writeheader()
        writer.writerows(rows)


def output_path(identity, file_timestamp):
    folder = os.path.join(io_root(), OUTPUT_FOLDER)
    os.makedirs(folder, exist_ok=True)
    token = (
        identity.get("number")
        or identity.get("name")
        or "UNKNOWN"
    )
    filename = "J18_SURFACE_AREA_{0}_{1}.csv".format(
        clean_filename_token(token),
        file_timestamp,
    )
    return os.path.join(folder, filename)


def run(session, run_datetime=None):
    now = run_datetime or datetime.datetime.now().astimezone()
    timestamp = now.isoformat(timespec="seconds")
    file_timestamp = now.strftime("%Y%m%d_%H%M%S")
    try:
        work_part = session.Parts.Work
    except Exception:
        work_part = None

    if work_part is None:
        identity = {
            "number": "",
            "revision": "",
            "name": "UNKNOWN",
        }
        rows = [
            failure_total_row(
                timestamp,
                identity,
                {
                    "included": 0,
                    "skipped_sheet": 0,
                    "skipped_convergent": 0,
                },
                "FAILED_NO_WORK_PART",
                "No active NX work part is available.",
            )
        ]
    else:
        identity = part_identity(work_part)
        try:
            rows = calculate_surface_rows(work_part, timestamp)
        except Exception as error:
            rows = [
                failure_total_row(
                    timestamp,
                    identity,
                    {
                        "included": 0,
                        "skipped_sheet": 0,
                        "skipped_convergent": 0,
                    },
                    "FAILED_CALCULATION",
                    error_text(error),
                )
            ]

    path = output_path(identity, file_timestamp)
    write_csv(path, rows)
    return path, rows


def main():
    session = NXOpen.Session.GetSession()
    log_line(session, "=" * 72)
    log_line(session, "J18 ACTIVE WORK PART SURFACE AREA")
    log_line(session, "Build: " + BUILD)
    log_line(
        session,
        "Scope: direct traditional solid bodies; output unit: m^2",
    )
    log_line(session, "Measurement accuracy: {0:.2f}".format(
        MEASUREMENT_ACCURACY
    ))
    log_line(session, "=" * 72)

    try:
        path, rows = run(session)
        total = rows[-1]
        for row in rows:
            if row["ROW_TYPE"] == "BODY":
                log_line(
                    session,
                    "Body {0} ({1}): {2} m^2 [{3}]".format(
                        row["BODY_INDEX"],
                        row["BODY_NAME"],
                        row["SURFACE_AREA_M2"] or "<blank>",
                        row["STATUS"],
                    ),
                )

        log_line(
            session,
            "Included solids: {0}; skipped sheets: {1}; "
            "skipped convergent: {2}".format(
                total["INCLUDED_SOLID_BODY_COUNT"],
                total["SKIPPED_SHEET_BODY_COUNT"],
                total["SKIPPED_CONVERGENT_BODY_COUNT"],
            ),
        )
        log_line(
            session,
            "TOTAL SURFACE AREA: {0} m^2 [{1}]".format(
                total["SURFACE_AREA_M2"] or "<blank>",
                total["STATUS"],
            ),
        )
        if total["MESSAGE"]:
            log_line(session, total["MESSAGE"])
        log_line(session, "CSV: " + path)
    except Exception as error:
        log_line(session, "J18 FAILED: " + error_text(error))
        log_line(session, traceback.format_exc())
        raise


if __name__ == "__main__":
    main()
