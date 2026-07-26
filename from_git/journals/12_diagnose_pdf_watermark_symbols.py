"""
Journal 12 - PDF Watermark and NX Catalog Symbol Diagnostic

Purpose:
    Compare NX PrintPDFBuilder settings for drawings where standard catalog
    symbols (for example omega or pi) are visible in NX but missing from the
    exported PDF.

Run 1:
    Display the affected drawing and run this journal. The canonical drawing
    identity is stored in NX_PDF_DIAGNOSTIC/LAST_TARGET.json.

Run 2:
    Close the drawing completely and run this journal again. The stored
    canonical identity is opened with session.Parts.OpenDisplay.

The journal creates five multipage PDFs. It does not update, modify, or save
the drawing, and it never changes layer visibility.

Target: NX 2312 and NX X 2506 embedded Python
"""

import datetime
import json
import os
import traceback

import NXOpen


JOURNAL_BUILD_ID = "J12-NX2506-PDF-SYMBOL-MATRIX-V1"
OUTPUT_FOLDER_NAME = "NX_PDF_DIAGNOSTIC"
TARGET_FILENAME = "LAST_TARGET.json"
DIAGNOSTIC_WATERMARK = "J12_WATERMARK_TEST"
VERIFY_OUTPUT_FILES = True


def clean(value):
    if value is None:
        return ""
    try:
        return str(value).strip()
    except Exception:
        return ""


def runtime_source_path():
    try:
        return os.path.abspath(__file__)
    except Exception:
        return "<unknown>"


def enum_text(value):
    text = clean(value)
    if text:
        return text
    try:
        return clean(value.name)
    except Exception:
        return "<unavailable>"


def output_root():
    configured = clean(os.environ.get("NX_JOURNALS_IO_DIR"))
    if configured:
        base = os.path.abspath(os.path.expanduser(configured))
    elif clean(os.environ.get("USERPROFILE")):
        base = os.path.join(os.environ["USERPROFILE"], "Desktop")
    else:
        base = os.path.join(os.path.expanduser("~"), "Desktop")
    return os.path.join(base, OUTPUT_FOLDER_NAME)


def variant_definitions():
    return [
        {
            "token": "00_WATERMARK_DISABLED_BASELINE",
            "add_watermark": False,
            "output_text": None,
            "custom_symbols_in_foreground": None,
        },
        {
            "token": "01_TEXT_WATERMARK",
            "add_watermark": True,
            "output_text": "TEXT",
            "custom_symbols_in_foreground": False,
        },
        {
            "token": "02_TEXT_FOREGROUND_SYMBOLS",
            "add_watermark": True,
            "output_text": "TEXT",
            "custom_symbols_in_foreground": True,
        },
        {
            "token": "03_POLYLINE_WATERMARK",
            "add_watermark": True,
            "output_text": "POLYLINES",
            "custom_symbols_in_foreground": False,
        },
        {
            "token": "04_POLYLINE_FOREGROUND_SYMBOLS",
            "add_watermark": True,
            "output_text": "POLYLINES",
            "custom_symbols_in_foreground": True,
        },
    ]


def drawing_sheet_count(part):
    if part is None:
        return 0
    try:
        return int(part.DrawingSheets.Count)
    except Exception:
        pass
    try:
        return len(list(part.DrawingSheets))
    except Exception:
        return 0


def part_name(part):
    if part is None:
        return ""
    for property_name in ("Name", "Leaf", "PartName"):
        try:
            value = clean(getattr(part, property_name))
            if value:
                return value
        except Exception:
            pass
    return "<unnamed>"


def journal_identifier(part):
    if part is None:
        return ""
    try:
        return clean(part.JournalIdentifier)
    except Exception:
        return ""


def object_identity(part):
    if part is None:
        return ("NONE", "")
    try:
        return ("TAG", str(part.Tag))
    except Exception:
        pass
    identifier = journal_identifier(part)
    if identifier:
        return ("JOURNAL_IDENTIFIER", identifier.upper())
    return ("NAME", part_name(part).upper())


def session_parts(session):
    try:
        return list(session.Parts)
    except Exception:
        return []


def is_canonical_drawing_identifier(identifier):
    normalized = clean(identifier).replace("\\", "/")
    return (
        normalized.upper().startswith("@DB/")
        and "/SPECIFICATION/" in normalized.upper()
    )


def identifier_key(identifier):
    return clean(identifier).replace("\\", "/").upper()


def target_record(part):
    identifier = journal_identifier(part)
    if not is_canonical_drawing_identifier(identifier):
        raise RuntimeError(
            "Displayed drawing has no canonical Teamcenter "
            "/specification/ JournalIdentifier."
        )
    if drawing_sheet_count(part) < 1:
        raise RuntimeError("Displayed part contains no drawing sheets.")
    return {
        "journal_identifier": identifier,
        "part_name": part_name(part),
        "sheet_count": drawing_sheet_count(part),
        "captured_at": datetime.datetime.now().isoformat(),
    }


def write_target(path, record):
    with open(path, "w", encoding="utf-8", newline="") as handle:
        json.dump(record, handle, indent=2, sort_keys=True)
        handle.write("\n")


def read_target(path):
    if not os.path.isfile(path):
        raise RuntimeError(
            "No saved diagnostic target exists. Display the affected "
            "drawing and run Journal 12 once before the closed-drawing run."
        )
    with open(path, "r", encoding="utf-8-sig") as handle:
        record = json.load(handle)
    identifier = clean(record.get("journal_identifier"))
    if not is_canonical_drawing_identifier(identifier):
        raise RuntimeError(
            "Saved diagnostic target has no valid canonical "
            "/specification/ identifier."
        )
    return record


def dispose(value):
    if value is not None:
        try:
            value.Dispose()
        except Exception:
            pass


def unwrap_open_result(value):
    if isinstance(value, (tuple, list)):
        part = value[0] if value else None
        load_status = value[1] if len(value) > 1 else None
        return part, load_status
    return value, None


def resolve_diagnostic_drawing(session, target_path, logger):
    display_part = session.Parts.Display
    if drawing_sheet_count(display_part) > 0:
        record = target_record(display_part)
        write_target(target_path, record)
        logger.write("Captured target: {0}".format(record["journal_identifier"]))
        return {
            "part": display_part,
            "mode": "PRELOADED",
            "opened_by_journal": False,
            "record": record,
        }

    record = read_target(target_path)
    identifier = record["journal_identifier"]
    preloaded = {object_identity(part) for part in session_parts(session)}
    logger.write("Calling session.Parts.OpenDisplay with saved target:")
    logger.write("  {0}".format(identifier))

    load_status = None
    try:
        part, load_status = unwrap_open_result(
            session.Parts.OpenDisplay(identifier)
        )
    finally:
        dispose(load_status)

    if part is None:
        raise RuntimeError("OpenDisplay returned no drawing part.")

    opened_by_journal = object_identity(part) not in preloaded
    return {
        "part": part,
        "mode": (
            "CLOSED_AUTO"
            if opened_by_journal
            else "PRELOADED_TARGET"
        ),
        "opened_by_journal": opened_by_journal,
        "record": record,
    }


def safe_property(value, name):
    try:
        return enum_text(getattr(value, name))
    except Exception as error:
        return "<unavailable: {0}>".format(clean(error))


def builder_snapshot(builder):
    names = (
        "AddWatermark",
        "Watermark",
        "OutputText",
        "CustomSymbolsInForeground",
        "Colors",
        "RasterImages",
        "ShadedGeometry",
        "ImageResolution",
    )
    return {name: safe_property(builder, name) for name in names}


def output_text_option(token):
    if token == "TEXT":
        return NXOpen.PrintPDFBuilder.OutputTextOption.Text
    if token == "POLYLINES":
        return NXOpen.PrintPDFBuilder.OutputTextOption.Polylines
    raise RuntimeError("Unknown PDF output-text option: {0}".format(token))


def apply_variant(builder, variant, watermark):
    builder.AddWatermark = bool(variant["add_watermark"])
    builder.Watermark = watermark

    if variant["output_text"] is not None:
        builder.OutputText = output_text_option(variant["output_text"])

    if variant["custom_symbols_in_foreground"] is not None:
        builder.CustomSymbolsInForeground = bool(
            variant["custom_symbols_in_foreground"]
        )


def export_pdf_variant(
    drawing_part,
    sheets,
    output_path,
    variant,
    watermark,
):
    if not sheets:
        raise RuntimeError("Drawing contains no sheets to export.")

    sheets[0].Open()
    builder = drawing_part.PlotManager.CreatePrintPdfbuilder()
    defaults = builder_snapshot(builder)
    applied = {}
    try:
        builder.Action = NXOpen.PrintPDFBuilder.ActionOption.Native
        builder.Filename = output_path
        builder.Append = False
        apply_variant(builder, variant, watermark)
        applied = builder_snapshot(builder)
        builder.SourceBuilder.SetSheets(sheets)
        builder.Commit()
    finally:
        builder.Destroy()

    if VERIFY_OUTPUT_FILES and not os.path.isfile(output_path):
        raise RuntimeError(
            "PDF builder committed but no output file was created: {0}".format(
                output_path
            )
        )
    return defaults, applied


def run_variant_matrix(
    drawing_part,
    sheets,
    run_folder,
    logger,
):
    outputs = []
    failures = []

    for variant in variant_definitions():
        token = variant["token"]
        output_path = os.path.join(run_folder, token + ".pdf")
        logger.write("")
        logger.write("Variant: {0}".format(token))
        logger.write("  Output: {0}".format(output_path))
        logger.write(
            "  Requested settings: AddWatermark={0}, "
            "OutputText={1}, CustomSymbolsInForeground={2}".format(
                variant["add_watermark"],
                variant["output_text"],
                variant["custom_symbols_in_foreground"],
            )
        )
        try:
            defaults, applied = export_pdf_variant(
                drawing_part,
                sheets,
                output_path,
                variant,
                DIAGNOSTIC_WATERMARK,
            )
            logger.write("  Builder defaults: {0}".format(defaults))
            logger.write("  Applied settings: {0}".format(applied))
            logger.write("  PDF created.")
            outputs.append(output_path)
        except Exception as error:
            failures.append(
                {
                    "token": token,
                    "message": clean(error),
                    "traceback": traceback.format_exc(),
                }
            )
            logger.write("  FAILED: {0}".format(error))
            logger.write(traceback.format_exc())

    return outputs, failures


def type_name(value):
    try:
        return type(value).__name__
    except Exception:
        return "<unknown>"


def log_sheet_inventory(part, sheets, logger):
    logger.write("Drawing sheets: {0}".format(len(sheets)))
    for index, sheet in enumerate(sheets, start=1):
        logger.write(
            "  Sheet {0}: name={1}, out_of_date={2}".format(
                index,
                clean(getattr(sheet, "Name", "<unnamed>")),
                safe_property(sheet, "IsOutOfDate"),
            )
        )
        try:
            views = list(sheet.GetDraftingViews())
            logger.write("    Drafting views: {0}".format(len(views)))
            for view in views:
                logger.write(
                    "      {0}: name={1}, out_of_date={2}".format(
                        type_name(view),
                        clean(getattr(view, "Name", "<unnamed>")),
                        safe_property(view, "IsOutOfDate"),
                    )
                )
        except Exception as error:
            logger.write(
                "    Drafting-view inventory unavailable: {0}".format(error)
            )

    try:
        logger.write("Work layer: {0}".format(part.Layers.WorkLayer))
    except Exception as error:
        logger.write("Work layer unavailable: {0}".format(error))

    logger.write("Non-empty object layers:")
    nonempty_layer_count = 0
    for layer in range(1, 257):
        try:
            objects = list(part.Layers.GetAllObjectsOnLayer(layer))
        except Exception:
            continue
        if not objects:
            continue

        nonempty_layer_count += 1
        type_counts = {}
        for obj in objects:
            name = type_name(obj)
            type_counts[name] = type_counts.get(name, 0) + 1
        try:
            state = enum_text(part.Layers.GetState(layer))
        except Exception as error:
            state = "<unavailable: {0}>".format(clean(error))
        summary = ", ".join(
            "{0}={1}".format(name, type_counts[name])
            for name in sorted(type_counts)
        )
        logger.write(
            "  Layer {0}: state={1}, objects={2} [{3}]".format(
                layer,
                state,
                len(objects),
                summary,
            )
        )
    if nonempty_layer_count == 0:
        logger.write("  <none reported>")


def set_display(session, part):
    result = session.Parts.SetDisplay(part, False, True)
    if isinstance(result, (tuple, list)) and len(result) > 1:
        dispose(result[1])


def restore_original_state(
    session,
    original_display,
    original_work,
    logger,
):
    logger.write("Restoring original NX display/work parts...")
    if original_display is not None:
        try:
            set_display(session, original_display)
            logger.write(
                "  Display restored: {0}".format(part_name(original_display))
            )
        except Exception as error:
            logger.write("  WARNING: display restore failed: {0}".format(error))
    if original_work is not None:
        try:
            session.Parts.SetWork(original_work)
            logger.write(
                "  Work restored: {0}".format(part_name(original_work))
            )
        except Exception as error:
            logger.write("  WARNING: work restore failed: {0}".format(error))


def close_diagnostic_part(part, logger):
    try:
        part.Close(
            NXOpen.BasePart.CloseWholeTree.FalseValue,
            NXOpen.BasePart.CloseModified.CloseModified,
            None,
        )
        logger.write("Diagnostic-opened drawing closed.")
    except Exception as error:
        logger.write(
            "WARNING: diagnostic-opened drawing close failed: {0}".format(
                error
            )
        )


def cleanup_diagnostic(
    session,
    drawing_part,
    opened_by_journal,
    original_display,
    original_work,
    original_sheet,
    logger,
):
    restore_original_state(
        session,
        original_display,
        original_work,
        logger,
    )

    if (
        not opened_by_journal
        and drawing_part is original_display
        and original_sheet is not None
    ):
        try:
            original_sheet.Open()
            logger.write("  Original drawing sheet restored.")
        except Exception as error:
            logger.write(
                "  WARNING: original sheet restore failed: {0}".format(error)
            )

    if opened_by_journal and drawing_part is not None:
        close_diagnostic_part(drawing_part, logger)


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


def write_log(path, lines):
    with open(path, "w", encoding="utf-8-sig", newline="") as handle:
        handle.write("\n".join(lines))
        handle.write("\n")


def main():
    session = NXOpen.Session.GetSession()
    logger = Logger(session)
    root = output_root()
    os.makedirs(root, exist_ok=True)
    target_path = os.path.join(root, TARGET_FILENAME)
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S_%f")

    original_display = session.Parts.Display
    original_work = session.Parts.Work
    original_sheet = None
    if drawing_sheet_count(original_display) > 0:
        try:
            original_sheet = original_display.DrawingSheets.CurrentDrawingSheet
        except Exception:
            pass

    drawing_part = None
    opened_by_journal = False
    run_folder = None
    final_status = "NOT_STARTED"
    outputs = []

    logger.write("=" * 78)
    logger.write("JOURNAL 12 - PDF WATERMARK AND CATALOG SYMBOL DIAGNOSTIC")
    logger.write("=" * 78)
    logger.write("Journal build: {0}".format(JOURNAL_BUILD_ID))
    logger.write("Journal source: {0}".format(runtime_source_path()))
    logger.write("Target file: {0}".format(target_path))
    logger.write(
        "Original display: {0}".format(part_name(original_display))
    )
    logger.write("Original work: {0}".format(part_name(original_work)))

    try:
        resolution = resolve_diagnostic_drawing(
            session,
            target_path,
            logger,
        )
        drawing_part = resolution["part"]
        opened_by_journal = resolution["opened_by_journal"]
        mode = resolution["mode"]
        expected_identifier = resolution["record"]["journal_identifier"]
        returned_identifier = journal_identifier(drawing_part)

        if identifier_key(returned_identifier) != identifier_key(
            expected_identifier
        ):
            raise RuntimeError(
                "Resolved drawing identifier does not match the saved "
                "diagnostic target. Expected {0}; returned {1}.".format(
                    expected_identifier,
                    returned_identifier,
                )
            )
        if drawing_sheet_count(drawing_part) < 1:
            raise RuntimeError("Resolved target contains no drawing sheets.")

        run_folder = os.path.join(
            root,
            "{0}_{1}".format(timestamp, mode),
        )
        os.makedirs(run_folder, exist_ok=False)

        logger.write("Mode: {0}".format(mode))
        logger.write("Run folder: {0}".format(run_folder))
        logger.write("Drawing name: {0}".format(part_name(drawing_part)))
        logger.write(
            "Drawing identifier: {0}".format(
                returned_identifier
            )
        )
        logger.write(
            "Opened by Journal 12: {0}".format(opened_by_journal)
        )

        set_display(session, drawing_part)
        sheets = list(drawing_part.DrawingSheets)
        log_sheet_inventory(drawing_part, sheets, logger)

        outputs, variant_failures = run_variant_matrix(
            drawing_part,
            sheets,
            run_folder,
            logger,
        )
        if len(outputs) == len(variant_definitions()):
            final_status = "SUCCESS"
        elif outputs:
            final_status = "PARTIAL_SUCCESS"
        else:
            final_status = "FAILED"
        logger.write(
            "Variant failures: {0}".format(len(variant_failures))
        )

    except Exception as error:
        final_status = "FAILED"
        logger.write("FAILED: {0}".format(error))
        logger.write(traceback.format_exc())

    finally:
        cleanup_diagnostic(
            session,
            drawing_part,
            opened_by_journal,
            original_display,
            original_work,
            original_sheet,
            logger,
        )
        logger.write("")
        logger.write("Generated PDF count: {0}".format(len(outputs)))
        logger.write("FINAL STATUS: {0}".format(final_status))

        log_folder = run_folder or root
        log_path = os.path.join(
            log_folder,
            "PDF_SYMBOL_DIAGNOSTIC_{0}.txt".format(timestamp),
        )
        logger.write("Log: {0}".format(log_path))
        try:
            write_log(log_path, logger.lines)
        except Exception as error:
            logger.write("WARNING: could not write log: {0}".format(error))


if __name__ == "__main__":
    main()
