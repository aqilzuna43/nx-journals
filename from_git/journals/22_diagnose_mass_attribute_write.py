"""Journal 22 - Mass Attribute Write Diagnostic (NX 2506)

Answers, on the ACTIVE WORK PART ONLY (fast), which mechanisms actually
write the standard NX roll-up attributes on this build:

  A. classic compute - MeasureManager.NewMassProperties / NewFaceProperties
     return mass / area for the part's solid bodies?
  B. native builder - does PropertiesManager.CreateMassPropertiesBuilder +
     UpdateOnSave + UpdateNow + Commit write NX_MassPropRollup* ?
  C. direct write - which category accepts AttributePropertiesBuilder
     writes for NX_MassPropRollupArea ("Rolled-Up Mass Properties" vs
     "Materials")?
  D. full before/after attribute dump (category/title/type/value) so the
     real attribute landscape is visible.

The journal writes controlled test values on the work part (only) and saves
it - run it on a disposable part.  Output: Listing Window plus
NX_MASS_SURFACE_UPDATE\\J22_DIAGNOSTIC_<root>_<timestamp>.csv and .json.

Target: NX X 2506 embedded Python
Run via: NX > Tools > Journal > Play
"""

import csv
import datetime
import json
import os
import traceback

import NXOpen


BUILD = "J22-NX2506-MASS-ATTRIBUTE-WRITE-DIAGNOSTIC-V1"
OUTPUT_FOLDER = "NX_MASS_SURFACE_UPDATE"
MEASUREMENT_ACCURACY = 0.99
ROLLUP_MASS_ATTRIBUTE = "NX_MassPropRollupMass"
ROLLUP_AREA_ATTRIBUTE = "NX_MassPropRollupArea"
ATTRIBUTE_CATEGORIES = ("Rolled-Up Mass Properties", "Materials")

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
    }


def body_flag(body, property_name):
    value = getattr(body, property_name)
    return bool(value() if callable(value) else value)


def solid_bodies(part):
    return [
        body
        for body in list(getattr(part, "Bodies", []))
        if body_flag(body, "IsSolidBody")
        and not body_flag(body, "IsConvergentBody")
    ]


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
    return getattr(info, "StringValue", ""), kind or "String"


def dump_attributes(part):
    """Return sorted (category, title, type, value, unset) tuples."""
    iterator = None
    result = []
    try:
        iterator = part.CreateAttributeIterator()
        for info in part.GetUserAttributes(iterator):
            value, kind = attribute_value(info)
            result.append(
                {
                    "category": clean(getattr(info, "Category", "")),
                    "title": clean(getattr(info, "Title", "")),
                    "type": kind,
                    "value": value,
                    "unset": bool(getattr(info, "Unset", False)),
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
            }
        )
    finally:
        dispose(iterator)
    result.sort(
        key=lambda item: (
            item["category"].lower(),
            item["title"].lower(),
        )
    )
    return result


def summarize_attributes(attributes):
    relevant = [
        item
        for item in attributes
        if (
            "MASS" in item["title"].upper()
            or "AREA" in item["title"].upper()
            or "VOLUME" in item["title"].upper()
            or "ROLLUP" in item["title"].upper()
        )
    ]
    return relevant


def resolve_area_units(work_part):
    units = work_part.UnitCollection
    try:
        area_unit = units.FindObject("SquareMeter") or units.FindObject(
            "SquareMetre"
        )
    except Exception:
        area_unit = None
    try:
        length_unit = units.FindObject("Meter") or units.FindObject("Metre")
    except Exception:
        length_unit = None
    try:
        mass_unit = units.FindObject("Kilogram")
    except Exception:
        mass_unit = None
    return area_unit, length_unit, mass_unit


def test_classic_compute(work_part, bodies):
    """Test A: NewMassProperties + NewFaceProperties on own solid bodies."""
    findings = []
    if not bodies:
        findings.append(
            {"test": "A_classic_compute", "status": "NO_SOLID_BODIES"}
        )
        return findings

    area_unit, length_unit, mass_unit = resolve_area_units(work_part)
    measure = work_part.MeasureManager

    mass_measurement = None
    try:
        mass_measurement = measure.NewMassProperties(
            [mass_unit], MEASUREMENT_ACCURACY, bodies
        )
        row = {
            "test": "A_classic_compute",
            "status": "OK",
            "mass_kg": getattr(mass_measurement, "Mass", None),
            "area_mm2": None,
            "volume_mm3": getattr(mass_measurement, "Volume", None),
        }
        area = getattr(mass_measurement, "Area", None)
        if area is not None:
            row["area_mm2"] = area
        findings.append(row)
    except Exception as error:
        findings.append(
            {
                "test": "A_classic_compute_NewMassProperties",
                "status": "FAILED",
                "message": error_text(error),
            }
        )
    finally:
        dispose(mass_measurement)

    face_area_total = 0.0
    face_failures = []
    for body in bodies:
        faces = list(body.GetFaces())
        if not faces:
            continue
        measurement = None
        try:
            measurement = measure.NewFaceProperties(
                area_unit, length_unit, MEASUREMENT_ACCURACY, faces
            )
            face_area_total += float(measurement.Area)
        except Exception as error:
            face_failures.append(
                "{0}: {1}".format(body.Name, error_text(error))
            )
        finally:
            dispose(measurement)
    findings.append(
        {
            "test": "A_classic_compute_NewFaceProperties",
            "status": "FAILED" if face_failures else "OK",
            "area_m2": face_area_total if not face_failures else None,
            "message": " | ".join(face_failures),
        }
    )
    return findings


def test_native_builder(work_part):
    """Test B: CreateMassPropertiesBuilder + UpdateOnSave + UpdateNow + Commit."""
    builder = None
    findings = []
    try:
        properties_manager = getattr(work_part, "PropertiesManager", None)
        if properties_manager is None:
            findings.append(
                {"test": "B_native_builder", "status": "NO_PROPERTIES_MANAGER"}
            )
            return findings
        builder = properties_manager.CreateMassPropertiesBuilder([work_part])
        builder.Accuracy = MEASUREMENT_ACCURACY
        options = getattr(builder, "UpdateOptions", None)
        yes = getattr(options, "Yes", None) if options is not None else None
        update_on_save_ok = False
        if getattr(builder, "UpdateOnSave", None) is not None and yes is not None:
            builder.UpdateOnSave = yes
            update_on_save_ok = True
        update_now = getattr(builder, "UpdateNow", None)
        commit = getattr(builder, "Commit", None)
        builder_ok = True
        messages = []
        if update_now is None:
            builder_ok = False
            messages.append("no UpdateNow")
        else:
            update_now()
        if commit is None:
            builder_ok = False
            messages.append("no Commit")
        else:
            commit()
        findings.append(
            {
                "test": "B_native_builder",
                "status": "OK" if builder_ok else "PARTIAL",
                "update_on_save_set": update_on_save_ok,
                "message": " | ".join(messages),
            }
        )
    except Exception as error:
        findings.append(
            {
                "test": "B_native_builder",
                "status": "FAILED",
                "message": error_text(error),
            }
        )
    finally:
        dispose(builder)
    return findings


def test_direct_write(session, part, title, value):
    """Test C: AttributePropertiesBuilder write per category."""
    findings = []
    for category in ATTRIBUTE_CATEGORIES:
        builder = None
        try:
            builder = session.AttributeManager.CreateAttributePropertiesBuilder(
                part,
                [part],
                NXOpen.AttributePropertiesBuilder.OperationType.Save,
            )
            builder.Category = category
            builder.Title = title
            builder.DataType = (
                NXOpen.AttributePropertiesBaseBuilder.DataTypeOptions.Number
            )
            builder.NumberValue = float(value)
            builder.Commit()
            findings.append(
                {
                    "test": "C_direct_write",
                    "category": category,
                    "status": "OK",
                }
            )
        except Exception as error:
            findings.append(
                {
                    "test": "C_direct_write",
                    "category": category,
                    "status": "FAILED",
                    "message": error_text(error),
                }
            )
        finally:
            dispose(builder)
    return findings


def build_report(
    session, work_part, timestamp, bodies, before, after, findings
):
    return {
        "build": BUILD,
        "timestamp": timestamp,
        "work_part": part_identity(work_part),
        "solid_body_count": len(bodies),
        "before_attributes": before,
        "after_attributes": after,
        "findings": findings,
    }


def run(session, run_datetime=None):
    now = run_datetime or datetime.datetime.now().astimezone()
    timestamp = now.isoformat(timespec="seconds")
    file_timestamp = now.strftime("%Y%m%d_%H%M%S")

    try:
        work_part = session.Parts.Work
    except Exception:
        work_part = None
    if work_part is None:
        raise RuntimeError("Open an NX 3D master part or assembly first.")

    bodies = solid_bodies(work_part)
    before = dump_attributes(work_part)

    findings = []
    findings.extend(test_classic_compute(work_part, bodies))
    findings.extend(test_native_builder(work_part))
    findings.extend(
        test_direct_write(session, work_part, ROLLUP_AREA_ATTRIBUTE, 123456.0)
    )
    findings.extend(
        test_direct_write(session, work_part, ROLLUP_MASS_ATTRIBUTE, 0.12345)
    )

    try:
        work_part.Save(
            NXOpen.BasePart.SaveComponents.FalseValue,
            NXOpen.BasePart.CloseAfterSave.FalseValue,
        )
        findings.append({"test": "save", "status": "OK"})
    except Exception as error:
        findings.append(
            {"test": "save", "status": "FAILED", "message": error_text(error)}
        )

    after = dump_attributes(work_part)
    report = build_report(
        session, work_part, timestamp, bodies, before, after, findings
    )

    folder = os.path.join(io_root(), OUTPUT_FOLDER)
    os.makedirs(folder, exist_ok=True)
    token = clean_filename_token(
        part_identity(work_part)["number"]
        or part_identity(work_part)["name"]
    )
    csv_path = os.path.join(
        folder,
        "J22_DIAGNOSTIC_{0}_{1}.csv".format(token, file_timestamp),
    )
    json_path = os.path.join(
        folder,
        "J22_DIAGNOSTIC_{0}_{1}.json".format(token, file_timestamp),
    )

    with open(csv_path, "w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(
            ["TEST", "CATEGORY", "STATUS", "VALUE", "MESSAGE"]
        )
        for item in findings:
            writer.writerow(
                [
                    item.get("test", ""),
                    item.get("category", ""),
                    item.get("status", ""),
                    item.get(
                        "mass_kg",
                        item.get("area_m2", item.get("area_mm2", "")),
                    ),
                    item.get("message", ""),
                ]
            )
    with open(json_path, "w", encoding="utf-8") as handle:
        json.dump(report, handle, indent=2, ensure_ascii=False)

    return csv_path, json_path, report


def main():
    session = NXOpen.Session.GetSession()
    log_line(session, "=" * 72)
    log_line(session, "J22 MASS ATTRIBUTE WRITE DIAGNOSTIC")
    log_line(session, "Build: " + BUILD)
    log_line(session, "Scope: active work part only; writes test values.")
    log_line(session, "=" * 72)
    try:
        csv_path, json_path, report = run(session)
        log_line(
            session,
            "Work part: {0} (solid bodies: {1})".format(
                report["work_part"]["number"] or report["work_part"]["name"],
                report["solid_body_count"],
            ),
        )
        for item in report["findings"]:
            log_line(
                session,
                "[{0}] {1} -> {2}{3}".format(
                    item.get("test", ""),
                    item.get("category", ""),
                    item.get("status", ""),
                    " | " + item["message"]
                    if item.get("message")
                    else "",
                ),
            )
        log_line(session, "CSV: " + csv_path)
        log_line(session, "JSON: " + json_path)
        log_line(
            session,
            "Send the JSON (or Listing output) to confirm which write "
            "mechanism and category work on this NX build.",
        )
    except Exception as error:
        log_line(session, "J22 FAILED: " + error_text(error))
        log_line(session, traceback.format_exc())
        raise


if __name__ == "__main__":
    main()
