"""Journal 07 - CSV-driven Teamcenter PDF + STEP export.

NX 2312 embedded Python 3.10. No third-party modules or direct Teamcenter API.
PDF drawings use the resolver proven by Journal 09:
@DB/<part>/<rev>/specification/<part>-<rev>-dwg<n>
"""
import csv
import datetime
import os
import re
import traceback
import NXOpen

INPUT_FILENAME = "NX_EXPORT_SCOPE.csv"
OUTPUT_ROOT_FOLDER = "NX_BULK_EXPORT"
MAX_DWG = 9
VERIFY_FILES = True
CLOSE_OPENED_PARTS = True
TRUE = {"YES", "Y", "TRUE", "1", "X"}
FALSE = {"", "NO", "N", "FALSE", "0"}
BAD_FILENAME = '<>:"/\\|?*'
DWG_RE = re.compile(r"(?:^|[-_])DWG(\d+)(?:$|[^A-Z0-9])", re.I)
ALIASES = {
    "pn": ("DB_PART_NO", "Item Number", "PART_NUMBER", "Part Number"),
    "rev": ("DB_PART_REV", "Item Rev", "REVISION", "Revision"),
    "pdf": ("PDF", "Export_PDF", "EXPORT_PDF"),
    "step": ("STEP", "Export_STEP", "EXPORT_STEP"),
    "desc": ("PART_DESCRIPTION", "Part Description"),
    "module": ("PRIMARY_MODULE", "Primary Module"),
    "status": ("DATA_PACK_STATUS", "Status"),
    "owner": ("OWNER", "Owner"),
}
RESULT_COLUMNS = (
    "RUN_TIMESTAMP", "SOURCE_ROW_COUNT", "MERGED_ROW_COUNT", "DB_PART_NO",
    "DB_PART_REV", "PART_DESCRIPTION", "PRIMARY_MODULE", "DATA_PACK_STATUS",
    "OWNER", "PDF_REQUESTED", "PDF_RESULT", "PDF_FILE_COUNT", "PDF_FILES",
    "STEP_REQUESTED", "STEP_RESULT", "STEP_FILE", "LOADED_REVISION",
    "OVERALL_RESULT", "MESSAGE", "DURATION_SECONDS",
)


def text(value):
    return "" if value is None else str(value).strip()


def log(session, message, lines):
    value = str(message)
    lines.append(value)
    try:
        window = session.ListingWindow
        window.Open()
        for item in value.splitlines() or [""]:
            window.WriteFullline(item)
    except Exception:
        pass
    try:
        print(value)
    except Exception:
        pass


def err(error):
    values = [text(error) or type(error).__name__]
    for name in ("ErrorCode", "error_code", "Code", "code"):
        try:
            value = getattr(error, name)
            if not callable(value) and text(value):
                values.append("{0}={1}".format(name, text(value)))
        except Exception:
            pass
    return "; ".join(values)


def dispose(value):
    if value is not None:
        try:
            value.Dispose()
        except Exception:
            pass


def unwrap(value):
    if isinstance(value, (tuple, list)):
        return (value[0] if value else None, value[1] if len(value) > 1 else None)
    return value, None


def parts(session):
    try:
        return list(session.Parts)
    except Exception:
        return []


def jid(part):
    try:
        return text(part.JournalIdentifier)
    except Exception:
        return ""


def part_name(part):
    for name in ("Name", "Leaf", "FullPath", "JournalIdentifier"):
        try:
            value = text(getattr(part, name))
            if value:
                return value
        except Exception:
            pass
    return "part"


def key(part):
    try:
        return "TAG:" + str(part.Tag)
    except Exception:
        return "JID:" + jid(part).upper()


def attr(part, name):
    try:
        return text(part.GetStringAttribute(name))
    except Exception:
        pass
    try:
        value = part.GetUserAttribute(name, NXOpen.NXObject.AttributeType.String, -1)
        return text(value.StringValue)
    except Exception:
        return ""


def identity(part):
    return (
        attr(part, "DB_PART_NO") or attr(part, "PART_NUMBER") or attr(part, "ITEM_ID"),
        attr(part, "DB_PART_REV") or attr(part, "REVISION") or attr(part, "ITEM_REVISION"),
    )


def count(collection):
    try:
        return len(list(collection))
    except Exception:
        try:
            return int(collection.Count)
        except Exception:
            return 0


def sheet_count(part):
    try:
        return count(part.DrawingSheets)
    except Exception:
        return 0


def body_count(part):
    try:
        return count(part.Bodies)
    except Exception:
        return 0


def dwg_index(part):
    for value in (jid(part), part_name(part)):
        match = DWG_RE.search(value.upper())
        if match:
            try:
                return int(match.group(1))
            except Exception:
                pass
    return None


def set_display(session, part):
    result = session.Parts.SetDisplay(part, False, True)
    if isinstance(result, (tuple, list)) and len(result) > 1:
        dispose(result[1])


def restore(session, display, work, lines):
    if display is not None:
        try:
            set_display(session, display)
        except Exception as error:
            log(session, "ERROR restoring display: " + err(error), lines)
    if work is not None:
        try:
            session.Parts.SetWork(work)
        except Exception as error:
            log(session, "ERROR restoring work part: " + err(error), lines)


def close_opened(part, session, display, work, lines):
    if not CLOSE_OPENED_PARTS or part is None:
        return
    if display is None:
        log(session, "  WARNING: opened part left open; no original display part.", lines)
        return
    restore(session, display, work, lines)
    try:
        part.Close(NXOpen.BasePart.CloseWholeTree.FalseValue,
                   NXOpen.BasePart.CloseModified.CloseModified, None)
        log(session, "  Closed journal-opened part.", lines)
    except Exception as error:
        log(session, "  WARNING: close failed: " + err(error), lines)


def io_root():
    configured = text(os.environ.get("NX_JOURNALS_IO_DIR"))
    if configured:
        return os.path.abspath(os.path.expanduser(configured))
    profile = text(os.environ.get("USERPROFILE"))
    return os.path.join(profile or os.path.expanduser("~"), "Desktop")


def run_folders(root, stamp):
    base = os.path.join(root, OUTPUT_ROOT_FOLDER, stamp)
    path, index = base, 1
    while os.path.exists(path):
        path = base + "_{0:02d}".format(index)
        index += 1
    os.makedirs(path)
    result = {"run": path}
    for name in ("PDF", "STEP", "REPORTS", "LOGS"):
        folder = os.path.join(path, name)
        os.makedirs(folder)
        result[name.lower()] = folder
    return result


def safe_filename(value, fallback="part"):
    value = text(value)
    if not value:
        return fallback
    value = "".join("_" if c in BAD_FILENAME or ord(c) < 32 else c for c in value)
    return value.strip(" .") or fallback


def headers(fieldnames):
    if not fieldnames:
        raise ValueError("CSV has no header row")
    normalized = [(name, text(name).lstrip("\ufeff").upper()) for name in fieldnames]
    result = {}
    for logical, names in ALIASES.items():
        wanted = {name.upper() for name in names}
        for original, normalized_name in normalized:
            if normalized_name in wanted:
                result[logical] = original
                break
    missing = [name for name in ("pn", "rev", "pdf", "step") if name not in result]
    if missing:
        raise ValueError("Missing CSV column(s): " + ", ".join(missing))
    return result


def enabled(value, label, row_number, warnings):
    value = text(value).upper()
    if value in TRUE:
        return True
    if value in FALSE:
        return False
    warnings.append("Row {0}: unknown {1} value '{2}'".format(row_number, label, value))
    return False


def read_scope(path):
    merged, invalid = {}, []
    input_count = ignored = 0
    with open(path, "r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        names = headers(reader.fieldnames)
        for row_number, row in enumerate(reader, start=2):
            if not any(text(value) for value in row.values() if not isinstance(value, list)):
                continue
            input_count += 1
            warnings = []
            pdf = enabled(row.get(names["pdf"]), "PDF", row_number, warnings)
            step = enabled(row.get(names["step"]), "STEP", row_number, warnings)
            if not pdf and not step and not warnings:
                ignored += 1
                continue
            item = {
                "pn": text(row.get(names["pn"])), "rev": text(row.get(names["rev"])),
                "pdf": pdf, "step": step, "warnings": warnings,
                "rows": [row_number], "source_count": 1, "merged_count": 0,
                "desc": text(row.get(names.get("desc", ""), "")),
                "module": text(row.get(names.get("module", ""), "")),
                "status": text(row.get(names.get("status", ""), "")),
                "owner": text(row.get(names.get("owner", ""), "")),
            }
            problems = []
            if not item["pn"]:
                problems.append("part number is blank")
            if not item["rev"]:
                problems.append("revision is blank")
            if not pdf and not step:
                problems.append("no valid export request")
            if problems:
                item["warnings"].append("Row {0}: {1}".format(row_number, "; ".join(problems)))
                invalid.append(item)
                continue
            item_key = (item["pn"].upper(), item["rev"].upper())
            if item_key not in merged:
                merged[item_key] = item
                continue
            existing = merged[item_key]
            existing["rows"].append(row_number)
            existing["source_count"] += 1
            existing["merged_count"] += 1
            existing["pdf"] = existing["pdf"] or pdf
            existing["step"] = existing["step"] or step
            for name in ("desc", "module", "status", "owner"):
                if not existing[name] and item[name]:
                    existing[name] = item[name]
            for warning in warnings:
                if warning not in existing["warnings"]:
                    existing["warnings"].append(warning)
    return list(merged.values()), invalid, input_count, ignored


def result_row(stamp, item, invalid=False):
    pdf = "INVALID_INPUT" if invalid and item["pdf"] else ("PENDING" if item["pdf"] else "NOT_REQUESTED")
    step = "INVALID_INPUT" if invalid and item["step"] else ("PENDING" if item["step"] else "NOT_REQUESTED")
    return {
        "RUN_TIMESTAMP": stamp, "SOURCE_ROW_COUNT": item["source_count"],
        "MERGED_ROW_COUNT": item["merged_count"], "DB_PART_NO": item["pn"],
        "DB_PART_REV": item["rev"], "PART_DESCRIPTION": item["desc"],
        "PRIMARY_MODULE": item["module"], "DATA_PACK_STATUS": item["status"],
        "OWNER": item["owner"], "PDF_REQUESTED": "YES" if item["pdf"] else "NO",
        "PDF_RESULT": pdf, "PDF_FILE_COUNT": 0, "PDF_FILES": "",
        "STEP_REQUESTED": "YES" if item["step"] else "NO", "STEP_RESULT": step,
        "STEP_FILE": "", "LOADED_REVISION": item["rev"],
        "OVERALL_RESULT": "INVALID_INPUT" if invalid else "PENDING",
        "MESSAGE": " | ".join(item["warnings"]),
        "DURATION_SECONDS": "0.000" if invalid else "",
    }


def canonical_dwg(pn, rev, index):
    return "@DB/{0}/{1}/specification/{0}-{1}-dwg{2}".format(pn, rev, index)


def loaded_drawings(session, pn, rev):
    result = {}
    prefix = "@DB/{0}/{1}/SPECIFICATION/".format(pn, rev).upper()
    for part in parts(session):
        index = dwg_index(part)
        part_pn, part_rev = identity(part)
        if index and (jid(part).upper().startswith(prefix) or
                      (part_pn.upper() == pn.upper() and part_rev.upper() == rev.upper() and sheet_count(part))):
            result.setdefault(index, part)
    return result


def open_drawing(session, specification, preloaded, lines):
    log(session, "  Attempt drawing open: " + specification, lines)
    status = None
    try:
        part, status = unwrap(session.Parts.OpenDisplay(specification))
    except Exception as error:
        log(session, "    Not opened: " + err(error), lines)
        return None, err(error)
    finally:
        dispose(status)
    if part is None:
        return None, "OpenDisplay returned no part"
    opened = key(part) not in preloaded
    log(session, "    Opened: " + part_name(part), lines)
    log(session, "    JournalIdentifier: " + jid(part), lines)
    return (part, opened), ""


def pdf_filename(part, pn, rev, index, sheet, sheet_index, total):
    base = attr(part, "DRAWING_NUMBER") or pn
    name = "{0}_REV{1}_DWG{2}".format(safe_filename(base), safe_filename(rev, ""), index)
    if total > 1:
        name += "_SHEET{0:02d}".format(sheet_index)
        try:
            sheet_name = text(sheet.Name)
        except Exception:
            sheet_name = ""
        if sheet_name:
            name += "_" + safe_filename(sheet_name, "")
    return name + ".pdf"


def export_pdf_part(session, part, pn, rev, index, folder, lines):
    set_display(session, part)
    sheets = list(part.DrawingSheets)
    if not sheets:
        return [], ["DWG{0}: no drawing sheets".format(index)]
    log(session, "  Exporting DWG{0}: {1} sheet(s)".format(index, len(sheets)), lines)
    paths_out, failures = [], []
    for sheet_index, sheet in enumerate(sheets, start=1):
        path = os.path.join(folder, pdf_filename(part, pn, rev, index, sheet, sheet_index, len(sheets)))
        try:
            if os.path.exists(path):
                raise RuntimeError("output exists: " + path)
            sheet.Open()
            builder = part.PlotManager.CreatePrintPdfbuilder()
            try:
                builder.Action = NXOpen.PrintPDFBuilder.ActionOption.Native
                builder.Filename = path
                builder.Append = False
                builder.SourceBuilder.SetSheets([sheet])
                builder.Commit()
            finally:
                builder.Destroy()
            if VERIFY_FILES and not os.path.isfile(path):
                failures.append("DWG{0} sheet {1}: no PDF created".format(index, sheet_index))
            else:
                paths_out.append(path)
                log(session, "    PDF created: " + path, lines)
        except Exception as error:
            failures.append("DWG{0} sheet {1}: {2}".format(index, sheet_index, err(error)))
            log(session, traceback.format_exc(), lines)
    return paths_out, failures


def export_pdfs(session, item, folder, display, work, lines):
    pn, rev = item["pn"], item["rev"]
    preloaded = {key(part) for part in parts(session)}
    loaded = loaded_drawings(session, pn, rev)
    paths_out, failures, last_error = [], [], ""
    found = 0
    for index in range(1, MAX_DWG + 1):
        opened = False
        part = loaded.get(index)
        if part is not None:
            log(session, "  Drawing already loaded for DWG{0}: {1}".format(index, part_name(part)), lines)
        else:
            resolved, last_error = open_drawing(session, canonical_dwg(pn, rev, index), preloaded, lines)
            if resolved is None:
                log(session, "  DWG{0} not opened; stopping sequential search.".format(index), lines)
                break
            part, opened = resolved
        found += 1
        try:
            created, errors = export_pdf_part(session, part, pn, rev, index, folder, lines)
            paths_out.extend(created)
            failures.extend(errors)
        except Exception as error:
            failures.append("DWG{0}: {1}".format(index, err(error)))
            log(session, traceback.format_exc(), lines)
        finally:
            restore(session, display, work, lines)
            if opened:
                close_opened(part, session, display, work, lines)
    if paths_out and not failures:
        status = "SUCCESS"
    elif paths_out:
        status = "PARTIAL_SUCCESS"
    elif failures:
        status = "FAILED"
    else:
        status = "SKIPPED_NO_DRAWING"
    message = " | ".join(failures)
    if not found:
        message = "No drawing opened with canonical /specification/ identifier"
        if last_error:
            message += ": " + last_error
    return status, paths_out, message


def loaded_master(session, pn, rev):
    candidates = []
    for part in parts(session):
        part_pn, part_rev = identity(part)
        if part_pn.upper() != pn.upper() or part_rev.upper() != rev.upper():
            continue
        if "/SPECIFICATION/" in jid(part).upper() or (dwg_index(part) and sheet_count(part)):
            continue
        candidates.append(part)
    candidates.sort(key=lambda part: (body_count(part) <= 0, part_name(part).upper()))
    return candidates[0] if candidates else None


def open_master(session, pn, rev, lines):
    current = loaded_master(session, pn, rev)
    if current is not None:
        log(session, "  Master already loaded: " + part_name(current), lines)
        return current, False
    preloaded = {key(part) for part in parts(session)}
    for specification in ("@DB/{0}/{1}".format(pn, rev), "@DB/{0}/{1}/master".format(pn, rev)):
        log(session, "  Attempt master open: " + specification, lines)
        status = None
        try:
            part, status = unwrap(session.Parts.OpenBase(specification))
        except Exception as error:
            log(session, "    Not opened: " + err(error), lines)
            continue
        finally:
            dispose(status)
        if part is not None:
            log(session, "    Opened: " + part_name(part), lines)
            return part, key(part) not in preloaded
    return None, False


def export_step(session, item, folder, display, work, lines):
    part, opened = open_master(session, item["pn"], item["rev"], lines)
    if part is None:
        return "NOT_FOUND", "", "Master part could not be loaded"
    path = os.path.join(folder, "{0}_REV{1}.stp".format(safe_filename(item["pn"]), safe_filename(item["rev"], "")))
    try:
        set_display(session, part)
        session.Parts.SetWork(part)
        creator = session.DexManager.CreateStepCreator()
        try:
            try:
                creator.InputFile = part.FullPath
            except Exception:
                pass
            creator.OutputFile = path
            creator.ObjectTypes.Solids = True
            creator.ObjectTypes.Surfaces = True
            creator.ObjectTypes.Curves = True
            creator.ExportAs = NXOpen.StepCreator.ExportAsOption.Ap214
            creator.ProcessHoldFlag = True
            creator.Commit()
        finally:
            creator.Destroy()
        if VERIFY_FILES and not os.path.isfile(path):
            return "FAILED_NO_OUTPUT_FILE", "", "STEP builder created no file"
        log(session, "    STEP created: " + path, lines)
        return "SUCCESS", path, ""
    finally:
        restore(session, display, work, lines)
        if opened:
            close_opened(part, session, display, work, lines)


def overall(item, pdf_result, step_result):
    values = []
    if item["pdf"]:
        values.append(pdf_result)
    if item["step"]:
        values.append(step_result)
    if values and all(value == "SUCCESS" for value in values):
        return "SUCCESS"
    if any(value in ("SUCCESS", "PARTIAL_SUCCESS") for value in values):
        return "PARTIAL_SUCCESS"
    if item["pdf"] and not item["step"] and pdf_result == "SKIPPED_NO_DRAWING":
        return "SKIPPED_NO_DRAWING"
    if item["step"] and not item["pdf"] and step_result == "NOT_FOUND":
        return "NOT_FOUND"
    return "FAILED"


def process(session, item, folders, stamp, display, work, lines):
    start = datetime.datetime.now()
    row = result_row(stamp, item)
    messages = list(item["warnings"])
    if item["pdf"]:
        try:
            status, output, message = export_pdfs(session, item, folders["pdf"], display, work, lines)
            row["PDF_RESULT"], row["PDF_FILE_COUNT"], row["PDF_FILES"] = status, len(output), ";".join(output)
            if message:
                messages.append(message)
        except Exception as error:
            row["PDF_RESULT"] = "FAILED"
            messages.append("PDF export failed: " + err(error))
            log(session, traceback.format_exc(), lines)
    if item["step"]:
        try:
            status, path, message = export_step(session, item, folders["step"], display, work, lines)
            row["STEP_RESULT"], row["STEP_FILE"] = status, path
            if message:
                messages.append(message)
        except Exception as error:
            row["STEP_RESULT"] = "FAILED"
            messages.append("STEP export failed: " + err(error))
            log(session, traceback.format_exc(), lines)
    row["OVERALL_RESULT"] = overall(item, row["PDF_RESULT"], row["STEP_RESULT"])
    row["MESSAGE"] = " | ".join(dict.fromkeys(message for message in messages if message))
    row["DURATION_SECONDS"] = "{0:.3f}".format((datetime.datetime.now() - start).total_seconds())
    restore(session, display, work, lines)
    return row


def write_csv(path, rows):
    with open(path, "w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=RESULT_COLUMNS)
        writer.writeheader()
        for row in rows:
            writer.writerow({name: row.get(name, "") for name in RESULT_COLUMNS})


def write_log(path, lines):
    with open(path, "w", encoding="utf-8-sig", newline="") as handle:
        handle.write("\n".join(lines) + "\n")


def main():
    session = NXOpen.Session.GetSession()
    lines, rows, folders = [], [], None
    report = ""
    try:
        try:
            display = session.Parts.Display
        except Exception:
            display = None
        try:
            work = session.Parts.Work
        except Exception:
            work = None
        root = io_root()
        source = os.path.join(root, INPUT_FILENAME)
        log(session, "Journal 07 - CSV-driven PDF + STEP export", lines)
        log(session, "Input CSV: " + source, lines)
        log(session, "Drawing resolver: @DB/<part>/<rev>/specification/<part>-<rev>-dwg<n>", lines)
        log(session, "Drawing open method: session.Parts.OpenDisplay", lines)
        if not os.path.isfile(source):
            raise FileNotFoundError("Input CSV not found: " + source)
        items, invalid, input_count, ignored = read_scope(source)
        stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        folders = run_folders(root, stamp)
        report = os.path.join(folders["reports"], "EXPORT_RESULT_{0}.csv".format(stamp))
        logfile = os.path.join(folders["logs"], "EXPORT_LOG_{0}.txt".format(stamp))
        rows.extend(result_row(stamp, item, True) for item in invalid)
        log(session, "Input rows: {0}; unique requests: {1}; ignored: {2}; invalid: {3}".format(input_count, len(items), ignored, len(invalid)), lines)
        for index, item in enumerate(sorted(items, key=lambda value: (value["pn"], value["rev"])), start=1):
            log(session, "[{0}/{1}] {2} / {3}".format(index, len(items), item["pn"], item["rev"]), lines)
            row = process(session, item, folders, stamp, display, work, lines)
            rows.append(row)
            log(session, "  PDF: {0} ({1} file(s))".format(row["PDF_RESULT"], row["PDF_FILE_COUNT"]), lines)
            log(session, "  STEP: " + row["STEP_RESULT"], lines)
        restore(session, display, work, lines)
        write_csv(report, rows)
        counts = {}
        for row in rows:
            counts[row["OVERALL_RESULT"]] = counts.get(row["OVERALL_RESULT"], 0) + 1
        log(session, "Export complete", lines)
        for name in ("SUCCESS", "PARTIAL_SUCCESS", "SKIPPED_NO_DRAWING", "NOT_FOUND", "FAILED"):
            log(session, "{0}: {1}".format(name, counts.get(name, 0)), lines)
        log(session, "PDF files: {0}".format(sum(int(row["PDF_FILE_COUNT"]) for row in rows)), lines)
        log(session, "STEP files: {0}".format(sum(1 for row in rows if row["STEP_FILE"])), lines)
        log(session, "Result report: " + report, lines)
        write_log(logfile, lines)
    except Exception:
        log(session, "ERROR: Unhandled journal exception.", lines)
        log(session, traceback.format_exc(), lines)
    finally:
        try:
            restore(session, display, work, lines)
        except Exception:
            pass
        if folders:
            try:
                if report:
                    write_csv(report, rows)
                write_log(os.path.join(folders["logs"], "EXPORT_LOG_{0}.txt".format(os.path.basename(folders["run"]))), lines)
            except Exception:
                pass


if __name__ == "__main__":
    main()
