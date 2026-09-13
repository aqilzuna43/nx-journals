"""Journal 18 - CSV-driven powder-coat surface area estimator.

Reads NX_POWDER_COAT_SCOPE.csv, opens one exact Teamcenter master part at a
time, measures all solid bodies, closes journal-opened parts, and writes detail
and powder-summary CSV reports. NX data is read-only.

Target: NX 2312 and NX X 2506 embedded Python.
"""

import csv
import datetime
import math
import os
import traceback
from collections import OrderedDict

import NXOpen
import NXOpen.UF

BUILD = "J18-NX2506-POWDER-COAT-AREA-V1"
INPUT_NAME = "NX_POWDER_COAT_SCOPE.csv"
OUTPUT_NAME = "NX_POWDER_COAT"
ACCURACY = 0.99

DEFAULTS = {
    "AREA_FACTOR": 1.0,
    "COATS": 1,
    "DFT_UM": 70.0,
    "SPECIFIC_GRAVITY": 1.50,
    "UTILISATION": 0.85,
    "CONTINGENCY": 0.10,
    "PACK_SIZE_KG": 20.0,
}

ALIASES = {
    "PN": ("DB_PART_NO", "PART_NUMBER", "PART NUMBER", "ITEM NUMBER"),
    "REV": ("DB_PART_REV", "REVISION", "ITEM REV"),
    "QTY": ("QUANTITY", "QTY"),
    "INCLUDE": ("INCLUDE", "POWDER_COAT", "POWDER COAT"),
    "DESC": ("PART_DESCRIPTION", "DB_PART_DESC", "DB_PART_NAME", "COMPONENT NAME"),
    "POWDER": ("POWDER_CODE", "POWDER CODE", "COLOUR", "COLOR", "FINISH"),
    "AREA_FACTOR": ("COATED_AREA_FACTOR", "AREA_FACTOR", "AREA FACTOR"),
    "COATS": ("COATS", "NUMBER_OF_COATS"),
    "DFT_UM": ("DFT_UM", "DFT", "FILM_THICKNESS_UM"),
    "SPECIFIC_GRAVITY": ("SPECIFIC_GRAVITY", "SG"),
    "UTILISATION": ("UTILISATION", "UTILIZATION"),
    "CONTINGENCY": ("CONTINGENCY", "RESERVE"),
    "PACK_SIZE_KG": ("PACK_SIZE_KG", "BAG_SIZE_KG"),
}

DETAIL_COLUMNS = (
    "DB_PART_NO", "DB_PART_REV", "PART_DESCRIPTION", "POWDER_CODE",
    "QUANTITY", "SOLID_BODY_COUNT", "SHEET_BODY_COUNT",
    "RAW_AREA_M2_PER_PART", "COATED_AREA_FACTOR", "COATED_AREA_M2_PER_PART",
    "COATS", "TOTAL_COATED_AREA_M2", "DFT_UM", "SPECIFIC_GRAVITY",
    "CURED_FILM_VOLUME_L", "THEORETICAL_POWDER_KG", "UTILISATION",
    "CONTINGENCY", "REQUIRED_POWDER_KG", "PACK_SIZE_KG", "OPEN_SOURCE",
    "RESULT", "MESSAGE", "DURATION_SECONDS",
)

SUMMARY_COLUMNS = (
    "POWDER_CODE", "DFT_UM", "SPECIFIC_GRAVITY", "UTILISATION",
    "CONTINGENCY", "PACK_SIZE_KG", "UNIQUE_PARTS", "TOTAL_QUANTITY",
    "TOTAL_COATED_AREA_M2", "CURED_FILM_VOLUME_L",
    "THEORETICAL_POWDER_KG", "REQUIRED_POWDER_KG", "BAGS_REQUIRED",
    "PURCHASE_QUANTITY_KG", "ESTIMATED_SPARE_KG",
)


def text(value):
    return "" if value is None else str(value).strip()


def norm(value):
    return " ".join(text(value).lstrip("\ufeff").split()).upper()


def io_root():
    configured = text(os.environ.get("NX_JOURNALS_IO_DIR"))
    if configured:
        return os.path.abspath(os.path.expanduser(configured))
    profile = text(os.environ.get("USERPROFILE"))
    return os.path.join(profile or os.path.expanduser("~"), "Desktop")


def source_csv():
    configured = text(os.environ.get("NX_POWDER_COAT_INPUT"))
    return os.path.abspath(os.path.expanduser(configured)) if configured else os.path.join(io_root(), INPUT_NAME)


def dispose(value):
    try:
        if value is not None:
            value.Dispose()
    except Exception:
        pass


def get_attr(obj, name):
    try:
        return text(obj.GetStringAttribute(name))
    except Exception:
        pass
    try:
        item = obj.GetUserAttribute(name, NXOpen.NXObject.AttributeType.String, -1)
        return text(item.StringValue)
    except Exception:
        return ""


def part_identity(part):
    pn = get_attr(part, "DB_PART_NO") or get_attr(part, "PART_NUMBER")
    rev = get_attr(part, "DB_PART_REV") or get_attr(part, "REVISION")
    return pn, rev


def part_name(part):
    for name in ("Name", "Leaf", "FullPath", "JournalIdentifier"):
        try:
            value = text(getattr(part, name))
            if value:
                return value
        except Exception:
            pass
    return "<unknown>"


def object_key(obj):
    try:
        return ("TAG", str(obj.Tag))
    except Exception:
        return ("OBJ", str(id(obj)))


def identity_matches(part, pn, rev):
    loaded_pn, loaded_rev = part_identity(part)
    if loaded_pn and loaded_rev:
        return loaded_pn.upper() == pn.upper() and loaded_rev.upper() == rev.upper()
    expected = "@DB/{0}/{1}".format(pn, rev).upper()
    for name in ("JournalIdentifier", "FullPath", "Name"):
        try:
            if text(getattr(part, name)).upper().startswith(expected):
                return True
        except Exception:
            pass
    return False


def is_drawing_nonmaster(part, pn, rev):
    token = "{0}-{1}-DWG".format(pn, rev).upper()
    for name in ("JournalIdentifier", "FullPath", "Name"):
        try:
            if token in text(getattr(part, name)).upper():
                return True
        except Exception:
            pass
    return False


def session_parts(session):
    try:
        return list(session.Parts)
    except Exception:
        return []


def unwrap_open(value):
    if isinstance(value, (tuple, list)):
        return value[0] if value else None, value[1] if len(value) > 1 else None
    return value, None


def close_part(part, logger):
    try:
        part.Close(
            NXOpen.BasePart.CloseWholeTree.FalseValue,
            NXOpen.BasePart.CloseModified.CloseModified,
            None,
        )
    except Exception as exc:
        logger.write("  WARNING close failed for {0}: {1}".format(part_name(part), exc))


def open_master(session, pn, rev, logger):
    for part in session_parts(session):
        if identity_matches(part, pn, rev) and not is_drawing_nonmaster(part, pn, rev):
            return part, False, "loaded session"

    before = {object_key(part) for part in session_parts(session)}
    attempts = ("@DB/{0}/{1}".format(pn, rev), "@DB/{0}/{1}/master".format(pn, rev))
    for spec in attempts:
        part = status = None
        logger.write("  OpenBase: {0}".format(spec))
        try:
            part, status = unwrap_open(session.Parts.OpenBase(spec))
        except Exception as exc:
            logger.write("    Not opened: {0}".format(exc))
        finally:
            dispose(status)
        if part is None:
            continue
        opened_here = object_key(part) not in before
        if identity_matches(part, pn, rev) and not is_drawing_nonmaster(part, pn, rev):
            return part, opened_here, spec
        if opened_here:
            close_part(part, logger)
    raise RuntimeError("Exact Teamcenter master could not be opened: {0}/{1}".format(pn, rev))


def is_solid(body):
    try:
        value = body.IsSolidBody
        return bool(value() if callable(value) else value)
    except Exception:
        return False


def ask_area_m2(uf, tags, accuracy):
    props = [0.0] * 47
    stats = [0.0] * 13
    uf.Modl.AskMassProps3d(
        tags, len(tags), 1, 4, 1.0, 1,
        [accuracy] + [0.0] * 10, props, stats,
    )
    return float(props[0])


def measure_part(part, uf, accuracy):
    bodies = list(part.Bodies)
    solids = [body for body in bodies if is_solid(body)]
    if not solids:
        raise RuntimeError("No solid bodies. Assembly-only and sheet-body parts are skipped.")
    try:
        area = ask_area_m2(uf, [body.Tag for body in solids], accuracy)
        return area, len(solids), len(bodies) - len(solids), ""
    except Exception as combined_error:
        total = 0.0
        failures = []
        for index, body in enumerate(solids, start=1):
            try:
                total += ask_area_m2(uf, [body.Tag], accuracy)
            except Exception as exc:
                failures.append("body {0}: {1}".format(index, exc))
        if failures and len(failures) == len(solids):
            raise RuntimeError("Mass-property measurement failed: {0}".format(combined_error))
        message = "Individual-body fallback used. " + " | ".join(failures)
        return total, len(solids), len(bodies) - len(solids), message


class Logger:
    def __init__(self, session):
        self.lines = []
        self.window = session.ListingWindow
        try:
            self.window.Open()
        except Exception:
            self.window = None

    def write(self, message=""):
        value = str(message)
        self.lines.append(value)
        if self.window is not None:
            try:
                self.window.WriteFullline(value)
            except Exception:
                pass
        try:
            print(value)
        except Exception:
            pass


def resolve_headers(fieldnames):
    available = {norm(name): name for name in (fieldnames or [])}
    resolved = {}
    for logical, aliases in ALIASES.items():
        for alias in aliases:
            if norm(alias) in available:
                resolved[logical] = available[norm(alias)]
                break
    missing = [name for name in ("PN", "REV", "QTY") if name not in resolved]
    if missing:
        raise RuntimeError("Missing required CSV columns: {0}".format(", ".join(missing)))
    return resolved


def value(row, headers, logical):
    return text(row.get(headers.get(logical, ""), ""))


def number(raw, default, label):
    if not text(raw):
        if default is None:
            raise RuntimeError("{0} is required.".format(label))
        return float(default)
    try:
        result = float(text(raw))
    except ValueError:
        raise RuntimeError("{0} must be numeric.".format(label))
    if result <= 0:
        raise RuntimeError("{0} must be greater than zero.".format(label))
    return result


def fraction(raw, default, label, allow_zero=False):
    data = text(raw)
    if not data:
        result = float(default)
    else:
        percent = data.endswith("%")
        try:
            result = float(data[:-1] if percent else data)
        except ValueError:
            raise RuntimeError("{0} must be numeric or a percentage.".format(label))
        if percent or result > 1:
            result /= 100.0
    if result > 1 or result < 0 or (result == 0 and not allow_zero):
        raise RuntimeError("{0} must be between {1} and 1.".format(label, "0" if allow_zero else ">0"))
    return result


def include_row(raw, has_column):
    if not has_column:
        return True
    return norm(raw) in ("YES", "Y", "TRUE", "1", "X", "INCLUDE")


def read_scope(path):
    grouped = OrderedDict()
    invalid = []
    with open(path, "r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        headers = resolve_headers(reader.fieldnames)
        for row_no, row in enumerate(reader, start=2):
            if not any(text(item) for item in row.values()):
                continue
            try:
                if not include_row(value(row, headers, "INCLUDE"), "INCLUDE" in headers):
                    continue
                item = {
                    "PN": value(row, headers, "PN"),
                    "REV": value(row, headers, "REV"),
                    "QTY": number(value(row, headers, "QTY"), None, "Quantity"),
                    "DESC": value(row, headers, "DESC"),
                    "POWDER": value(row, headers, "POWDER") or "UNSPECIFIED",
                    "AREA_FACTOR": fraction(value(row, headers, "AREA_FACTOR"), DEFAULTS["AREA_FACTOR"], "Area factor"),
                    "COATS": int(number(value(row, headers, "COATS"), DEFAULTS["COATS"], "Coats")),
                    "DFT_UM": number(value(row, headers, "DFT_UM"), DEFAULTS["DFT_UM"], "DFT"),
                    "SPECIFIC_GRAVITY": number(value(row, headers, "SPECIFIC_GRAVITY"), DEFAULTS["SPECIFIC_GRAVITY"], "Specific gravity"),
                    "UTILISATION": fraction(value(row, headers, "UTILISATION"), DEFAULTS["UTILISATION"], "Utilisation"),
                    "CONTINGENCY": fraction(value(row, headers, "CONTINGENCY"), DEFAULTS["CONTINGENCY"], "Contingency", True),
                    "PACK_SIZE_KG": number(value(row, headers, "PACK_SIZE_KG"), DEFAULTS["PACK_SIZE_KG"], "Pack size"),
                }
                if not item["PN"] or not item["REV"]:
                    raise RuntimeError("Part number and revision are required.")
                key = (
                    item["PN"].upper(), item["REV"].upper(), item["POWDER"].upper(),
                    item["AREA_FACTOR"], item["COATS"], item["DFT_UM"],
                    item["SPECIFIC_GRAVITY"], item["UTILISATION"],
                    item["CONTINGENCY"], item["PACK_SIZE_KG"],
                )
                if key in grouped:
                    grouped[key]["QTY"] += item["QTY"]
                else:
                    grouped[key] = item
            except Exception as exc:
                invalid.append((row_no, value(row, headers, "PN"), value(row, headers, "REV"), text(exc)))
    return list(grouped.values()), invalid


def detail_result(item, measured):
    raw_area, solid_count, sheet_count, open_source, measurement_message = measured
    coated_per_part = raw_area * item["AREA_FACTOR"]
    total_area = coated_per_part * item["QTY"] * item["COATS"]
    film_l = total_area * item["DFT_UM"] / 1000.0
    theoretical_kg = film_l * item["SPECIFIC_GRAVITY"]
    required_kg = theoretical_kg / item["UTILISATION"] * (1.0 + item["CONTINGENCY"])
    return {
        "DB_PART_NO": item["PN"], "DB_PART_REV": item["REV"],
        "PART_DESCRIPTION": item["DESC"], "POWDER_CODE": item["POWDER"],
        "QUANTITY": item["QTY"], "SOLID_BODY_COUNT": solid_count,
        "SHEET_BODY_COUNT": sheet_count, "RAW_AREA_M2_PER_PART": round(raw_area, 6),
        "COATED_AREA_FACTOR": item["AREA_FACTOR"],
        "COATED_AREA_M2_PER_PART": round(coated_per_part, 6), "COATS": item["COATS"],
        "TOTAL_COATED_AREA_M2": round(total_area, 6), "DFT_UM": item["DFT_UM"],
        "SPECIFIC_GRAVITY": item["SPECIFIC_GRAVITY"],
        "CURED_FILM_VOLUME_L": round(film_l, 6),
        "THEORETICAL_POWDER_KG": round(theoretical_kg, 6),
        "UTILISATION": item["UTILISATION"], "CONTINGENCY": item["CONTINGENCY"],
        "REQUIRED_POWDER_KG": round(required_kg, 6), "PACK_SIZE_KG": item["PACK_SIZE_KG"],
        "OPEN_SOURCE": open_source, "RESULT": "PARTIAL" if measurement_message else "SUCCESS",
        "MESSAGE": measurement_message,
    }


def summarize(rows):
    groups = OrderedDict()
    for row in rows:
        if row.get("RESULT") not in ("SUCCESS", "PARTIAL"):
            continue
        key = tuple(row[name] for name in ("POWDER_CODE", "DFT_UM", "SPECIFIC_GRAVITY", "UTILISATION", "CONTINGENCY", "PACK_SIZE_KG"))
        if key not in groups:
            groups[key] = {name: row[name] for name in ("POWDER_CODE", "DFT_UM", "SPECIFIC_GRAVITY", "UTILISATION", "CONTINGENCY", "PACK_SIZE_KG")}
            groups[key].update({"_parts": set(), "TOTAL_QUANTITY": 0.0, "TOTAL_COATED_AREA_M2": 0.0, "CURED_FILM_VOLUME_L": 0.0, "THEORETICAL_POWDER_KG": 0.0, "REQUIRED_POWDER_KG": 0.0})
        group = groups[key]
        group["_parts"].add((row["DB_PART_NO"], row["DB_PART_REV"]))
        group["TOTAL_QUANTITY"] += float(row["QUANTITY"])
        for name in ("TOTAL_COATED_AREA_M2", "CURED_FILM_VOLUME_L", "THEORETICAL_POWDER_KG", "REQUIRED_POWDER_KG"):
            group[name] += float(row[name])
    result = []
    for group in groups.values():
        group["UNIQUE_PARTS"] = len(group.pop("_parts"))
        bags = int(math.ceil(group["REQUIRED_POWDER_KG"] / float(group["PACK_SIZE_KG"])))
        purchase = bags * float(group["PACK_SIZE_KG"])
        group["BAGS_REQUIRED"] = bags
        group["PURCHASE_QUANTITY_KG"] = round(purchase, 6)
        group["ESTIMATED_SPARE_KG"] = round(purchase - group["REQUIRED_POWDER_KG"], 6)
        for name in ("TOTAL_QUANTITY", "TOTAL_COATED_AREA_M2", "CURED_FILM_VOLUME_L", "THEORETICAL_POWDER_KG", "REQUIRED_POWDER_KG"):
            group[name] = round(group[name], 6)
        result.append(group)
    return result


def write_csv(path, columns, rows):
    with open(path, "w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=columns)
        writer.writeheader()
        for row in rows:
            writer.writerow({name: row.get(name, "") for name in columns})


def main():
    session = NXOpen.Session.GetSession()
    uf = NXOpen.UF.UFSession.GetUFSession()
    logger = Logger(session)
    original_display, original_work = session.Parts.Display, session.Parts.Work
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    run_root = os.path.join(io_root(), OUTPUT_NAME, timestamp)
    reports = os.path.join(run_root, "REPORTS")
    logs = os.path.join(run_root, "LOGS")
    os.makedirs(reports)
    os.makedirs(logs)
    detail_path = os.path.join(reports, "POWDER_COAT_DETAIL_{0}.csv".format(timestamp))
    summary_path = os.path.join(reports, "POWDER_COAT_SUMMARY_{0}.csv".format(timestamp))
    log_path = os.path.join(logs, "POWDER_COAT_LOG_{0}.txt".format(timestamp))
    accuracy = float(text(os.environ.get("NX_POWDER_COAT_ACCURACY")) or ACCURACY)
    if not 0 < accuracy < 1:
        accuracy = ACCURACY
    rows = []
    cache = {}
    logger.write("{0} | READ ONLY | accuracy={1}".format(BUILD, accuracy))
    logger.write("Input: {0}".format(source_csv()))
    try:
        scope, invalid = read_scope(source_csv())
        for row_no, pn, rev, message in invalid:
            rows.append({"DB_PART_NO": pn, "DB_PART_REV": rev, "RESULT": "INVALID_INPUT", "MESSAGE": "Row {0}: {1}".format(row_no, message)})
        for index, item in enumerate(scope, start=1):
            started = datetime.datetime.now()
            logger.write("[{0}/{1}] {2}/{3} x {4}".format(index, len(scope), item["PN"], item["REV"], item["QTY"]))
            row = {"DB_PART_NO": item["PN"], "DB_PART_REV": item["REV"], "PART_DESCRIPTION": item["DESC"], "POWDER_CODE": item["POWDER"], "QUANTITY": item["QTY"], "RESULT": "FAILED"}
            try:
                key = (item["PN"].upper(), item["REV"].upper())
                measured = cache.get(key)
                if measured is None:
                    part, opened_here, source = open_master(session, item["PN"], item["REV"], logger)
                    try:
                        area, solid_count, sheet_count, message = measure_part(part, uf, accuracy)
                        measured = (area, solid_count, sheet_count, source, message)
                        cache[key] = measured
                    finally:
                        if opened_here:
                            close_part(part, logger)
                else:
                    logger.write("  Cached area reused.")
                row = detail_result(item, measured)
                logger.write("  {0:.6f} m2/part -> {1:.6f} kg".format(row["RAW_AREA_M2_PER_PART"], row["REQUIRED_POWDER_KG"]))
            except Exception as exc:
                row["MESSAGE"] = text(exc)
                logger.write("  FAILED: {0}".format(exc))
                logger.write(traceback.format_exc())
            row["DURATION_SECONDS"] = "{0:.3f}".format((datetime.datetime.now() - started).total_seconds())
            rows.append(row)
        write_csv(detail_path, DETAIL_COLUMNS, rows)
        write_csv(summary_path, SUMMARY_COLUMNS, summarize(rows))
        logger.write("Detail: {0}".format(detail_path))
        logger.write("Summary: {0}".format(summary_path))
        logger.write("CAUTION: area is full solid-body geometry; validate masks/internal faces through COATED_AREA_FACTOR.")
    finally:
        try:
            if original_display is not None:
                result = session.Parts.SetDisplay(original_display, False, True)
                if isinstance(result, (tuple, list)) and len(result) > 1:
                    dispose(result[1])
            if original_work is not None:
                session.Parts.SetWork(original_work)
        except Exception as exc:
            logger.write("WARNING restore failed: {0}".format(exc))
        with open(log_path, "w", encoding="utf-8-sig") as handle:
            handle.write("\n".join(logger.lines) + "\n")


if __name__ == "__main__":
    main()
