"""J15 - Teamcenter X offline drawing workflow.

EXPORT: export managed model + target drawing to native NX, preserve Teamcenter
AutoTranslate filenames, mark all non-drawing .prt files read-only, and write a
manifest with a SHA-256 baseline.

IMPORT_DRY_RUN / IMPORT_APPLY: import the native drawing package with every
object defaulted to UseExisting and only the exact target drawing set to
Overwrite. APPLY additionally requires APPROVED=YES and ENGINEER in manifest.

Target: NX 2312 / NX X 2506 embedded Python in a managed Teamcenter session.
"""

import csv
import datetime
import hashlib
import os
import stat
import traceback

import NXOpen
import NXOpen.UF

USER_MODE = "EXPORT"  # EXPORT | IMPORT_DRY_RUN | IMPORT_APPLY
USER_SCOPE_CSV = r""   # blank => <I/O root>\NX_TC_OFFLINE_SCOPE.csv
USER_MANIFEST_CSV = r""

BUILD = "J15-TCX-OFFLINE-DRAWING-V1"
OUT_DIR = "NX_TC_OFFLINE_DRAWINGS"
DEFAULT_SCOPE = "NX_TC_OFFLINE_SCOPE.csv"
MODES = ("EXPORT", "IMPORT_DRY_RUN", "IMPORT_APPLY")
MANIFEST_FIELDS = [
    "RUN_ID", "PART_NUMBER", "REVISION", "DWG_INDEX", "MODEL_IDENTIFIER",
    "DRAWING_IDENTIFIER", "PACKAGE_DIR", "DRAWING_FILE", "EXPORT_LOG",
    "EXPORT_SHA256", "EXPORTED_AT", "REFERENCE_PRT_COUNT", "APPROVED",
    "ENGINEER", "IMPORT_STATUS", "NOTES",
]
REPORT_FIELDS = [
    "RUN_TIMESTAMP", "MODE", "PART_NUMBER", "REVISION", "DWG_INDEX",
    "DRAWING_IDENTIFIER", "DRAWING_FILE", "EXPORTED_SHA256", "CURRENT_SHA256",
    "CHANGED", "APPROVED", "ENGINEER", "DEFAULT_IMPORT_ACTION",
    "DRAWING_IMPORT_ACTION", "DRY_RUN", "RESULT", "MESSAGE", "CLONE_LOG",
]


def text(v): return "" if v is None else str(v)
def clean(v): return text(v).strip()
def upper(v): return clean(v).upper()
def env(name): return clean(os.environ.get(name))
def stamp(): return datetime.datetime.now().strftime("%Y%m%d_%H%M%S")


def io_root():
    p = env("NX_JOURNALS_IO_DIR")
    if p:
        return os.path.abspath(os.path.expanduser(p))
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    return desktop if os.path.isdir(desktop) else os.getcwd()


def mode(): return upper(env("NX_TC_OFFLINE_MODE") or USER_MODE or "EXPORT")


def scope_path():
    p = env("NX_TC_OFFLINE_SCOPE_FILE") or clean(USER_SCOPE_CSV)
    return os.path.abspath(os.path.expanduser(p)) if p else os.path.join(io_root(), DEFAULT_SCOPE)


def manifest_path():
    p = env("NX_TC_OFFLINE_MANIFEST_FILE") or clean(USER_MANIFEST_CSV)
    return os.path.abspath(os.path.expanduser(p)) if p else ""


def err(e):
    code = clean(getattr(e, "ErrorCode", ""))
    return "{0}{1} - {2}".format(type(e).__name__, (":" + code) if code else "", text(e))


def dispose(v):
    if v is not None:
        try: v.Dispose()
        except Exception: pass


class Log:
    def __init__(self, session):
        self.lines = []
        try:
            self.lw = session.ListingWindow
            self.lw.Open()
        except Exception:
            self.lw = None

    def write(self, s=""):
        s = text(s)
        self.lines.append(s)
        if self.lw is not None:
            try: self.lw.WriteFullline(s)
            except Exception:
                try: self.lw.WriteLine(s)
                except Exception: pass
        try: print(s)
        except Exception: pass


def read_csv(path, required):
    for encoding in ("utf-8-sig", "utf-8", "cp1252"):
        try:
            with open(path, "r", encoding=encoding, newline="") as f:
                r = csv.DictReader(f)
                headers = [clean(x) for x in (r.fieldnames or [])]
                missing = [x for x in required if x not in headers]
                if missing:
                    raise RuntimeError("Missing CSV column(s): " + ", ".join(missing))
                rows = []
                for n, src in enumerate(r, 2):
                    row = {clean(k): clean(v) for k, v in src.items() if k is not None}
                    row["_CSV_ROW"] = n
                    rows.append(row)
                return rows
        except UnicodeDecodeError:
            pass
    raise RuntimeError("Unable to decode CSV: " + path)


def write_csv(path, fields, rows):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for row in rows:
            w.writerow({k: row.get(k, "") for k in fields})


def write_log(path, lines):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8-sig") as f:
        f.write("\n".join(lines) + "\n")


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        while True:
            block = f.read(1024 * 1024)
            if not block: break
            h.update(block)
    return h.hexdigest()


def model_id(pn, rev): return "@DB/{0}/{1}".format(pn, rev)
def dataset_name(pn, rev, idx): return "{0}-{1}-dwg{2}".format(pn, rev, idx)
def drawing_id(pn, rev, idx): return "@DB/{0}/{1}/specification/{2}".format(pn, rev, dataset_name(pn, rev, idx))
def expected_native(pn, rev, idx): return "{0}_{1}_s_{2}.prt".format(pn, rev, dataset_name(pn, rev, idx))


def valid_native(path, pn, rev, idx):
    name = os.path.basename(path).lower()
    if name == expected_native(pn, rev, idx).lower(): return True
    return "_s_" in name and name.endswith("-{0}-dwg{1}.prt".format(rev, idx).lower())


def find_drawing(folder, pn, rev, idx):
    exact = os.path.join(folder, expected_native(pn, rev, idx))
    if os.path.isfile(exact): return exact
    matches = [os.path.join(folder, n) for n in os.listdir(folder)
               if os.path.isfile(os.path.join(folder, n)) and valid_native(os.path.join(folder, n), pn, rev, idx)]
    if len(matches) == 1: return matches[0]
    if not matches: return ""
    raise RuntimeError("Multiple native drawing files matched: " + ", ".join(matches))


def protect_refs(folder, drawing):
    target = os.path.normcase(os.path.abspath(drawing))
    count = 0
    for name in os.listdir(folder):
        if not name.lower().endswith(".prt"): continue
        path = os.path.join(folder, name)
        if not os.path.isfile(path): continue
        m = os.stat(path).st_mode
        if os.path.normcase(os.path.abspath(path)) == target:
            os.chmod(path, m | stat.S_IWRITE)
        else:
            os.chmod(path, m & ~stat.S_IWRITE)
            count += 1
    return count


def terminate(clone):
    try: clone.Terminate()
    except Exception: pass


def add_assembly(clone, name):
    try: result = clone.AddAssembly(name)
    except TypeError: result = clone.AddAssembly(name, None)
    if isinstance(result, (tuple, list)):
        for v in result:
            if hasattr(v, "Dispose"): return v
    return None


def naming_failures(clone):
    try:
        r = clone.InitNamingFailures()
        return r[-1] if isinstance(r, (tuple, list)) and r else r
    except Exception:
        try: return NXOpen.UF.UFClone.NamingFailures()
        except Exception: return None


def perform(clone, failures):
    try: return clone.PerformClone(failures)
    except TypeError: return clone.PerformClone(None)


def iterate_parts(clone):
    out = []
    try: clone.StartIteration()
    except Exception: return out
    while True:
        try: r = clone.Iterate()
        except TypeError:
            try: r = clone.Iterate(None)
            except Exception: break
        except Exception: break
        if isinstance(r, (tuple, list)):
            name = ""
            for v in r:
                if isinstance(v, str): name = v
        else:
            name = clean(r)
        if not name: break
        out.append(name)
    return out


def setup_export(clone, folder, logfile):
    clone.Initialise(NXOpen.UF.UFClone.OperationClass.ExportOperation)
    clone.SetFamilyTreatment(NXOpen.UF.UFClone.FamilyTreatment.TreatAsLost)
    clone.SetDefNaming(NXOpen.UF.UFClone.NamingTechnique.Autotranslate)
    clone.SetDefItemType("")
    clone.SetDefDirectory(folder)
    try: clone.SetAssocFileRootDir(folder)
    except Exception: pass
    clone.SetDefAction(NXOpen.UF.UFClone.Action.Overwrite)
    clone.SetDefAssocFileCopy(True)
    clone.SetLogfile(logfile)
    try: clone.SetCloneRelatedDwgs(False)
    except Exception: pass


def export_package(ufs, folder, model, drawing, logfile, log):
    c = ufs.Clone
    load = None
    try:
        terminate(c)
        setup_export(c, folder, logfile)
        log.write("  Add assembly: " + model)
        load = add_assembly(c, model)
        log.write("  Add drawing:  " + drawing)
        c.AddPart(drawing)
        nf = naming_failures(c)
        c.SetDryrun(False)
        try: c.GenerateReport()
        except Exception: pass
        perform(c, nf)
    finally:
        dispose(load)
        terminate(c)


def parse_scope(row):
    pn, rev = clean(row.get("PART_NUMBER")), clean(row.get("REVISION"))
    if not pn or not rev: raise RuntimeError("PART_NUMBER and REVISION are required")
    try: idx = int(clean(row.get("DWG_INDEX")))
    except Exception: raise RuntimeError("DWG_INDEX must be an integer")
    if idx < 1: raise RuntimeError("DWG_INDEX must be >= 1")
    return pn, rev, idx


def do_export(ufs, log):
    path = scope_path()
    if not os.path.isfile(path): raise RuntimeError("Scope CSV not found: " + path)
    rows = read_csv(path, ["PART_NUMBER", "REVISION", "DWG_INDEX"])
    run = stamp()
    root = os.path.join(io_root(), OUT_DIR, run)
    os.makedirs(root, exist_ok=True)
    manifest = os.path.join(root, "TCX_OFFLINE_MANIFEST_{0}.csv".format(run))
    result = []
    for src in rows:
        rec = {k: "" for k in MANIFEST_FIELDS}
        rec.update({"RUN_ID": run, "IMPORT_STATUS": "NOT_IMPORTED"})
        try:
            pn, rev, idx = parse_scope(src)
            model, drawing = model_id(pn, rev), drawing_id(pn, rev, idx)
            folder = os.path.join(root, "{0}_{1}_DWG{2}".format(pn, rev, idx))
            os.makedirs(folder, exist_ok=True)
            elog = os.path.join(folder, "EXPORT_{0}_{1}_DWG{2}.clone".format(pn, rev, idx))
            rec.update({"PART_NUMBER": pn, "REVISION": rev, "DWG_INDEX": str(idx),
                        "MODEL_IDENTIFIER": model, "DRAWING_IDENTIFIER": drawing,
                        "PACKAGE_DIR": folder, "EXPORT_LOG": elog,
                        "EXPORTED_AT": datetime.datetime.now().isoformat(timespec="seconds")})
            log.write("EXPORT {0}/{1}/dwg{2}".format(pn, rev, idx))
            export_package(ufs, folder, model, drawing, elog, log)
            native = find_drawing(folder, pn, rev, idx)
            if not native: raise RuntimeError("Expected native drawing not found: " + expected_native(pn, rev, idx))
            rec["DRAWING_FILE"] = native
            rec["EXPORT_SHA256"] = sha256(native)
            rec["REFERENCE_PRT_COUNT"] = str(protect_refs(folder, native))
            rec["NOTES"] = "Export OK; all non-drawing .prt files set read-only"
            log.write("  Drawing: " + os.path.basename(native))
        except Exception as e:
            rec["IMPORT_STATUS"] = "EXPORT_FAILED"
            rec["NOTES"] = err(e)
            log.write("  FAILED: " + err(e))
            log.write(traceback.format_exc())
        result.append(rec)
        write_csv(manifest, MANIFEST_FIELDS, result)
    log.write("Manifest: " + manifest)
    return manifest


def validate_target(row):
    pn, rev = clean(row.get("PART_NUMBER")), clean(row.get("REVISION"))
    try: idx = int(clean(row.get("DWG_INDEX")))
    except Exception: raise RuntimeError("Invalid DWG_INDEX in manifest")
    path, did = clean(row.get("DRAWING_FILE")), clean(row.get("DRAWING_IDENTIFIER"))
    expected = drawing_id(pn, rev, idx)
    if did.upper() != expected.upper() or "/SPECIFICATION/" not in did.upper():
        raise RuntimeError("Manifest drawing identity is not the expected /specification/ target")
    if not os.path.isfile(path): raise RuntimeError("DRAWING_FILE not found: " + path)
    if not valid_native(path, pn, rev, idx):
        raise RuntimeError("Native drawing was renamed or does not match Teamcenter AutoTranslate naming")
    return pn, rev, idx, path, did


def same_part(candidate, target):
    if not clean(candidate): return False
    try:
        if os.path.normcase(os.path.abspath(candidate)) == os.path.normcase(os.path.abspath(target)): return True
    except Exception: pass
    return os.path.basename(candidate).lower() == os.path.basename(target).lower()


def import_one(ufs, drawing, folder, logfile, dry_run, log):
    c = ufs.Clone
    load = None
    try:
        terminate(c)
        c.Initialise(NXOpen.UF.UFClone.OperationClass.ImportOperation)
        c.SetFamilyTreatment(NXOpen.UF.UFClone.FamilyTreatment.TreatAsLost)
        c.SetDefNaming(NXOpen.UF.UFClone.NamingTechnique.Autotranslate)
        c.SetDefItemType("")
        c.SetDefDirectory(folder)
        try: c.SetAssocFileRootDir(folder)
        except Exception: pass
        c.SetDefAction(NXOpen.UF.UFClone.Action.UseExisting)  # CRITICAL SAFETY DEFAULT
        c.SetDefAssocFileCopy(True)
        c.SetLogfile(logfile)
        try: c.SetPropagateActions(False)
        except Exception: pass
        load = add_assembly(c, drawing)
        parts = iterate_parts(c)
        changed = False
        for p in parts:
            if same_part(p, drawing):
                c.SetAction(p, NXOpen.UF.UFClone.Action.Overwrite, "")
                changed = True
            else:
                try: c.SetAction(p, NXOpen.UF.UFClone.Action.UseExisting, "")
                except Exception: pass
        if not changed:
            c.SetAction(drawing, NXOpen.UF.UFClone.Action.Overwrite, "")
        nf = naming_failures(c)
        c.SetDryrun(bool(dry_run))
        try: c.GenerateReport()
        except Exception: pass
        perform(c, nf)
        log.write("  Default=UseExisting; drawing=Overwrite; dry_run={0}".format(dry_run))
    finally:
        dispose(load)
        terminate(c)


def report_row(row, m, ts):
    return {"RUN_TIMESTAMP": ts, "MODE": m, "PART_NUMBER": row.get("PART_NUMBER", ""),
            "REVISION": row.get("REVISION", ""), "DWG_INDEX": row.get("DWG_INDEX", ""),
            "DRAWING_IDENTIFIER": row.get("DRAWING_IDENTIFIER", ""), "DRAWING_FILE": row.get("DRAWING_FILE", ""),
            "EXPORTED_SHA256": row.get("EXPORT_SHA256", ""), "CURRENT_SHA256": "", "CHANGED": "",
            "APPROVED": row.get("APPROVED", ""), "ENGINEER": row.get("ENGINEER", ""),
            "DEFAULT_IMPORT_ACTION": "UseExisting", "DRAWING_IMPORT_ACTION": "Overwrite",
            "DRY_RUN": "YES" if m == "IMPORT_DRY_RUN" else "NO", "RESULT": "", "MESSAGE": "", "CLONE_LOG": ""}


def do_import(ufs, log, m):
    path = manifest_path()
    if not path or not os.path.isfile(path): raise RuntimeError("Set USER_MANIFEST_CSV/NX_TC_OFFLINE_MANIFEST_FILE to a valid manifest")
    rows = read_csv(path, ["PART_NUMBER", "REVISION", "DWG_INDEX", "DRAWING_IDENTIFIER", "DRAWING_FILE", "EXPORT_SHA256", "APPROVED", "ENGINEER"])
    ts = stamp()
    report = os.path.join(os.path.dirname(path), "TCX_OFFLINE_{0}_{1}.csv".format(m, ts))
    results = []
    for row in rows:
        r = report_row(row, m, ts)
        results.append(r)
        try:
            pn, rev, idx, drawing, did = validate_target(row)
            baseline = clean(row.get("EXPORT_SHA256"))
            if not baseline: raise RuntimeError("EXPORT_SHA256 is blank")
            current = sha256(drawing)
            r["CURRENT_SHA256"] = current
            changed = current.lower() != baseline.lower()
            r["CHANGED"] = "YES" if changed else "NO"
            log.write("IMPORT {0}/{1}/dwg{2} changed={3}".format(pn, rev, idx, changed))
            if not changed:
                r["RESULT"] = "SKIPPED_UNCHANGED"
                r["MESSAGE"] = "SHA-256 matches export snapshot"
                write_csv(report, REPORT_FIELDS, results)
                continue
            if m == "IMPORT_APPLY":
                if upper(row.get("APPROVED")) != "YES":
                    r["RESULT"], r["MESSAGE"] = "BLOCKED_NOT_APPROVED", "IMPORT_APPLY requires APPROVED=YES"
                    write_csv(report, REPORT_FIELDS, results)
                    continue
                if not clean(row.get("ENGINEER")):
                    r["RESULT"], r["MESSAGE"] = "BLOCKED_ENGINEER_REQUIRED", "IMPORT_APPLY requires ENGINEER"
                    write_csv(report, REPORT_FIELDS, results)
                    continue
            ilog = os.path.join(os.path.dirname(drawing), "IMPORT_{0}_{1}_{2}_DWG{3}.clone".format(m, pn, rev, idx))
            r["CLONE_LOG"] = ilog
            import_one(ufs, drawing, os.path.dirname(drawing), ilog, m == "IMPORT_DRY_RUN", log)
            r["RESULT"] = "DRY_RUN_OK" if m == "IMPORT_DRY_RUN" else "IMPORT_APPLIED"
            r["MESSAGE"] = "UFClone completed: default UseExisting, exact drawing Overwrite"
        except Exception as e:
            r["RESULT"], r["MESSAGE"] = "FAILED", err(e)
            log.write("  FAILED: " + err(e))
            log.write(traceback.format_exc())
            write_csv(report, REPORT_FIELDS, results)
            if m == "IMPORT_APPLY": break
        write_csv(report, REPORT_FIELDS, results)
    log.write("Import report: " + report)
    return report


def main():
    session = NXOpen.Session.GetSession()
    ufs = NXOpen.UF.UFSession.GetUFSession()
    log = Log(session)
    m = mode()
    log.write("=" * 72)
    log.write("J15 TEAMCENTER X OFFLINE DRAWING WORKFLOW")
    log.write("Build: {0} | Mode: {1} | I/O: {2}".format(BUILD, m, io_root()))
    log.write("=" * 72)
    try:
        if m not in MODES: raise RuntimeError("Invalid USER_MODE: " + m)
        output = do_export(ufs, log) if m == "EXPORT" else do_import(ufs, log, m)
        log.write("FINAL STATUS: SUCCESS")
        log.write("Primary output: " + output)
    except Exception as e:
        log.write("FINAL STATUS: FAILED")
        log.write(err(e))
        log.write(traceback.format_exc())
        raise
    finally:
        try:
            root = os.path.join(io_root(), OUT_DIR)
            os.makedirs(root, exist_ok=True)
            write_log(os.path.join(root, "J15_{0}_{1}.txt".format(m, stamp())), log.lines)
        except Exception: pass


if __name__ == "__main__":
    main()


def GetUnloadOption(dummy):
    return NXOpen.Session.LibraryUnloadOption.Immediately
