"""
Journal 09 - Safe Teamcenter Specification Open Test

Purpose:
    Prove that NX 2312 can open one closed Teamcenter drawing specification
    directly from its canonical managed-mode JournalIdentifier:

        @DB/<part>/<revision>/specification/<part>-<revision>-dwg<index>

This journal does not export, save, or modify the drawing. It restores the
original display/work parts and closes only the drawing opened by this test.

Set the constants below or override them with environment variables:
    NX_TEST_PART_NO
    NX_TEST_PART_REV
    NX_TEST_DWG_INDEX
    NX_TEST_KEEP_OPEN=1

Target: NX 2312 embedded Python 3.10
"""

import datetime
import os
import traceback

import NXOpen


TEST_PART_NUMBER = "264MN024619A01"
TEST_REVISION = "A"
TEST_DRAWING_INDEX = 1
KEEP_OPEN_AFTER_SUCCESS = False

OUTPUT_FOLDER_NAME = "NX_X_SPEC_OPEN_TEST"


def clean(value):
    if value is None:
        return ""
    try:
        return str(value).strip()
    except Exception:
        return ""


def env_bool(name, default=False):
    value = clean(os.environ.get(name))
    if not value:
        return default
    return value.upper() in ("1", "TRUE", "YES", "Y", "ON")


def exception_details(error):
    details = {
        "type": "{0}.{1}".format(type(error).__module__, type(error).__name__),
        "message": clean(error),
    }
    for name in ("ErrorCode", "error_code", "Code", "code"):
        try:
            value = getattr(error, name)
            if not callable(value):
                details[name] = clean(value)
        except Exception:
            pass
    return details


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


def output_folder():
    configured = clean(os.environ.get("NX_JOURNALS_IO_DIR"))
    if configured:
        base = os.path.abspath(os.path.expanduser(configured))
    elif clean(os.environ.get("USERPROFILE")):
        base = os.path.join(os.environ["USERPROFILE"], "Desktop")
    else:
        base = os.path.join(os.path.expanduser("~"), "Desktop")
    return os.path.join(base, OUTPUT_FOLDER_NAME)


def journal_identifier(part):
    if part is None:
        return ""
    try:
        return clean(part.JournalIdentifier)
    except Exception:
        return ""


def part_name(part):
    if part is None:
        return ""
    try:
        return clean(part.Name)
    except Exception:
        return ""


def part_key(part):
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


def get_loaded_parts(session):
    try:
        return list(session.Parts)
    except Exception:
        return []


def find_loaded_by_identifier(session, expected_identifier):
    expected = expected_identifier.upper()
    for part in get_loaded_parts(session):
        if journal_identifier(part).upper() == expected:
            return part
    return None


def count_drawing_sheets(part):
    if part is None:
        return 0
    try:
        return len(list(part.DrawingSheets))
    except Exception:
        try:
            return int(part.DrawingSheets.Count)
        except Exception:
            return 0


def set_display(session, part):
    if part is None:
        return
    result = session.Parts.SetDisplay(part, False, True)
    if isinstance(result, (tuple, list)) and len(result) > 1:
        dispose(result[1])


def restore_original_state(session, original_display, original_work, logger):
    logger.write("Restoring original NX state...")
    if original_display is not None:
        try:
            set_display(session, original_display)
            logger.write("  Display restored: {0}".format(part_name(original_display)))
        except Exception as error:
            logger.write("  WARNING: display restore failed: {0}".format(
                exception_details(error)
            ))
    if original_work is not None:
        try:
            session.Parts.SetWork(original_work)
            logger.write("  Work part restored: {0}".format(part_name(original_work)))
        except Exception as error:
            logger.write("  WARNING: work-part restore failed: {0}".format(
                exception_details(error)
            ))


def close_test_part(part, session, original_display, original_work, logger):
    if part is None:
        return
    if original_display is None:
        logger.write(
            "Drawing left open because there was no original display part to restore."
        )
        return

    restore_original_state(session, original_display, original_work, logger)
    try:
        part.Close(
            NXOpen.BasePart.CloseWholeTree.FalseValue,
            NXOpen.BasePart.CloseModified.CloseModified,
            None,
        )
        logger.write("Test-opened drawing closed.")
    except Exception as error:
        logger.write("WARNING: drawing close failed: {0}".format(
            exception_details(error)
        ))


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


def resolve_test_scope():
    part_number = clean(os.environ.get("NX_TEST_PART_NO")) or TEST_PART_NUMBER
    revision = clean(os.environ.get("NX_TEST_PART_REV")) or TEST_REVISION
    index_text = clean(os.environ.get("NX_TEST_DWG_INDEX"))
    try:
        drawing_index = int(index_text) if index_text else int(TEST_DRAWING_INDEX)
    except Exception:
        drawing_index = int(TEST_DRAWING_INDEX)
    if drawing_index < 1:
        drawing_index = 1
    keep_open = env_bool("NX_TEST_KEEP_OPEN", KEEP_OPEN_AFTER_SUCCESS)
    return part_number, revision, drawing_index, keep_open


def main():
    session = NXOpen.Session.GetSession()
    logger = Logger(session)

    folder = output_folder()
    os.makedirs(folder, exist_ok=True)
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = os.path.join(
        folder, "SPECIFICATION_OPEN_TEST_{0}.txt".format(timestamp)
    )

    part_number, revision, drawing_index, keep_open = resolve_test_scope()
    drawing_name = "{0}-{1}-dwg{2}".format(
        part_number, revision, drawing_index
    )
    expected_identifier = (
        "@DB/{0}/{1}/specification/{2}".format(
            part_number, revision, drawing_name
        )
    )

    original_display = session.Parts.Display
    original_work = session.Parts.Work
    preloaded_keys = {part_key(part) for part in get_loaded_parts(session)}

    opened_part = None
    load_status = None
    opened_by_test = False
    final_status = "NOT_STARTED"

    logger.write("=" * 78)
    logger.write("JOURNAL 09 - SAFE TEAMCENTER SPECIFICATION OPEN TEST")
    logger.write("=" * 78)
    logger.write("Log: {0}".format(log_path))
    logger.write("Part number: {0}".format(part_number))
    logger.write("Revision: {0}".format(revision))
    logger.write("Drawing index: {0}".format(drawing_index))
    logger.write("Expected identifier: {0}".format(expected_identifier))
    logger.write("Original display: {0}".format(part_name(original_display)))
    logger.write("Original work: {0}".format(part_name(original_work)))
    logger.write("Keep open after success: {0}".format(keep_open))
    logger.write("")

    try:
        existing = find_loaded_by_identifier(session, expected_identifier)
        if existing is not None:
            final_status = "PRECONDITION_FAILED_ALREADY_LOADED"
            logger.write(
                "PRECONDITION FAILED: the exact drawing is already loaded."
            )
            logger.write("Loaded name: {0}".format(part_name(existing)))
            logger.write(
                "Close the drawing completely, then rerun Journal 09. "
                "A loaded drawing cannot prove that Teamcenter opening works."
            )
            return

        logger.write("Calling session.Parts.OpenDisplay(...) once with:")
        logger.write("  {0}".format(expected_identifier))

        opened_part, load_status = unwrap_open_result(
            session.Parts.OpenDisplay(expected_identifier)
        )

        if opened_part is None:
            final_status = "FAILED_NO_PART_RETURNED"
            logger.write("FAILED: OpenDisplay returned no part.")
            return

        opened_by_test = part_key(opened_part) not in preloaded_keys
        returned_identifier = journal_identifier(opened_part)
        sheet_count = count_drawing_sheets(opened_part)

        logger.write("OpenDisplay returned successfully.")
        logger.write("Returned name: {0}".format(part_name(opened_part)))
        logger.write("Returned identifier: {0}".format(returned_identifier))
        logger.write("Drawing sheet count: {0}".format(sheet_count))
        logger.write("Newly loaded by this test: {0}".format(opened_by_test))

        identifier_matches = (
            returned_identifier.upper() == expected_identifier.upper()
        )
        logger.write("Identifier matches request: {0}".format(identifier_matches))

        if sheet_count < 1:
            final_status = "FAILED_OPENED_WITHOUT_DRAWING_SHEETS"
            logger.write(
                "FAILED: the opened part has no drawing sheets. "
                "The identifier resolved, but not to a usable drawing."
            )
        elif not identifier_matches:
            final_status = "WARNING_IDENTIFIER_DIFFERENT"
            logger.write(
                "WARNING: the drawing opened and has sheets, but NX returned a "
                "different canonical JournalIdentifier."
            )
        else:
            final_status = "SUCCESS"
            logger.write(
                "SUCCESS: the canonical /specification/ identifier opened a "
                "usable drawing."
            )

    except Exception as error:
        final_status = "FAILED_EXCEPTION"
        logger.write("FAILED with exception:")
        logger.write("  {0}".format(exception_details(error)))
        logger.write(traceback.format_exc())

    finally:
        dispose(load_status)

        if opened_part is not None and opened_by_test and not keep_open:
            close_test_part(
                opened_part, session, original_display, original_work, logger
            )
        elif opened_part is not None and keep_open:
            logger.write("Drawing intentionally left open for inspection.")
        else:
            restore_original_state(
                session, original_display, original_work, logger
            )

        logger.write("")
        logger.write("FINAL STATUS: {0}".format(final_status))
        logger.write("Log: {0}".format(log_path))
        try:
            write_log(log_path, logger.lines)
        except Exception as error:
            logger.write("WARNING: could not write log: {0}".format(
                exception_details(error)
            ))


if __name__ == "__main__":
    main()
