# NXOpen Python journal
# Purpose: Test changing the Teamcenter Item Name through the legacy
#          UF_UGMGR SetPartNameDesc API.
# Scope: Dummy Part Name test only. NO revision change. NO CSV/bulk processing.
#
# Why this test exists:
# - DatabaseAttributeManager.SetAttribute("DB_PART_NAME", ...) returned NX 940049
#   ("The database attribute is not writable") on Teamcenter X.
# - UF_UGMGR_set_part_name_desc is a separate managed-mode database operation
#   intended to reset the name/description of an already-saved database part.
#
# IMPORTANT:
# - Run ONLY on a disposable Teamcenter item.
# - The default guard allows only Item ID AQIL-TEST.
# - This changes Item Name only, NOT Item ID / Part Number and NOT Revision.
# - No checkout, geometry save, revision, delete, or rollback is attempted.
# - If the server rejects the operation, capture the Listing Window output.

import NXOpen
import NXOpen.PDM
import NXOpen.UF


EXPECTED_ITEM_ID = "AQIL-TEST"
NEW_PART_NAME = "DUMMY-PART-NAME-UFUN-TEST"
DB_PART_NO = "DB_PART_NO"
DB_PART_NAME = "DB_PART_NAME"


def log(listing_window, text=""):
    listing_window.WriteLine(str(text))


def safe_dispose(obj):
    if obj is None:
        return
    try:
        obj.Dispose()
    except Exception:
        pass


def read_database_attribute(work_part, attribute_name):
    """Fresh read of a Teamcenter database attribute through PDMPart."""
    manager = None
    try:
        pdm_part = work_part.PDMPart
        if pdm_part is None:
            return None
        manager = pdm_part.NewDatabaseAttributeManager()
        manager.LoadAttributes(True)
        return manager.GetAttribute(attribute_name)
    except Exception:
        return None
    finally:
        safe_dispose(manager)


def get_item_id(work_part):
    """Prefer DB_PART_NO; fall back to parsing the managed-mode work-part name."""
    item_id = read_database_attribute(work_part, DB_PART_NO)
    if item_id:
        return str(item_id).strip()

    leaf = str(work_part.Leaf).strip()
    if leaf.startswith("@DB/"):
        leaf = leaf[4:]
    if "/" in leaf:
        leaf = leaf.split("/", 1)[0]
    return leaf.strip()


def ask_part_tag(ugmgr, item_id):
    """Return the Teamcenter database-part tag from the Python UF wrapper."""
    result = ugmgr.AskPartTag(item_id)
    if isinstance(result, tuple):
        return result[-1]
    return result


def is_null_tag(tag_value):
    """NXOpen Python represents UF tags as integer-like values; NULL_TAG is 0."""
    if tag_value is None:
        return True
    try:
        return int(tag_value) == 0
    except (TypeError, ValueError):
        return False


def ask_part_name_desc(ugmgr, database_part_tag):
    """Return (part_name, part_description) from UF_UGMGR."""
    result = ugmgr.AskPartNameDesc(database_part_tag)
    if isinstance(result, tuple):
        if len(result) >= 2:
            return result[0], result[1]
    # Defensive fallback in case a future wrapper returns an object-like result.
    return str(result), ""


def main():
    session = NXOpen.Session.GetSession()
    uf_session = NXOpen.UF.UFSession.GetUFSession()
    ugmgr = uf_session.Ugmgr

    listing = session.ListingWindow
    listing.Open()

    work_part = session.Parts.Work

    log(listing, "=" * 76)
    log(listing, "TEAMCENTER PART NAME CHANGE - TEST 2 - UF_UGMGR")
    log(listing, "=" * 76)

    if work_part is None:
        log(listing, "FAIL: No work part is open.")
        return

    log(listing, "Work part: {}".format(work_part.Leaf))
    try:
        log(listing, "Full path: {}".format(work_part.FullPath))
    except Exception:
        pass

    try:
        log(listing, "NX write access: {}".format(work_part.HasWriteAccess))
    except Exception:
        pass

    item_id = get_item_id(work_part)
    new_name = NEW_PART_NAME.strip()

    log(listing, "Detected Item ID: {}".format(item_id))
    log(listing, "Safety Item ID  : {}".format(EXPECTED_ITEM_ID))

    if not item_id:
        log(listing, "FAIL: Could not determine the Teamcenter Item ID.")
        return

    if item_id.upper() != EXPECTED_ITEM_ID.strip().upper():
        log(listing, "")
        log(listing, "SAFETY STOP: Current Item ID does not match EXPECTED_ITEM_ID.")
        log(listing, "Nothing was changed.")
        return

    if not new_name:
        log(listing, "FAIL: NEW_PART_NAME is blank. Nothing was changed.")
        return

    try:
        database_part_tag = ask_part_tag(ugmgr, item_id)
    except Exception as ex:
        log(listing, "")
        log(listing, "FAIL: UF_UGMGR could not resolve the Teamcenter Item.")
        log(listing, "Exception type: {}".format(type(ex).__name__))
        log(listing, "Message: {}".format(ex))
        error_code = getattr(ex, "ErrorCode", None)
        if error_code is not None:
            log(listing, "NX error code: {}".format(error_code))
        return

    log(listing, "Database part tag raw: {!r}".format(database_part_tag))
    log(listing, "Database part tag type: {}".format(type(database_part_tag).__name__))

    if is_null_tag(database_part_tag):
        log(listing, "FAIL: AskPartTag returned a null database part tag (0/None).")
        return

    log(listing, "Database part tag: {}".format(database_part_tag))

    try:
        old_name, old_desc = ask_part_name_desc(ugmgr, database_part_tag)
    except Exception as ex:
        log(listing, "")
        log(listing, "FAIL: UF_UGMGR AskPartNameDesc failed before any write attempt.")
        log(listing, "Exception type: {}".format(type(ex).__name__))
        log(listing, "Message: {}".format(ex))
        error_code = getattr(ex, "ErrorCode", None)
        if error_code is not None:
            log(listing, "NX error code: {}".format(error_code))
        return

    log(listing, "")
    log(listing, "UF_UGMGR current name : {}".format(old_name))
    log(listing, "UF_UGMGR description  : {}".format(old_desc))
    log(listing, "Requested new name    : {}".format(new_name))

    if str(old_name) == new_name:
        log(listing, "NO CHANGE: UF_UGMGR already reports the requested Part Name.")
        return

    log(listing, "")
    log(listing, "Calling UF_UGMGR SetPartNameDesc() ...")
    log(listing, "Description argument is blank so the existing description is not changed.")

    try:
        # UF_UGMGR documentation states that an empty part_desc leaves the
        # description unchanged. This test changes the Item Name only.
        ugmgr.SetPartNameDesc(database_part_tag, new_name, "")
        log(listing, "SetPartNameDesc() returned without an NXOpen exception.")
    except Exception as ex:
        log(listing, "")
        log(listing, "FAIL: Teamcenter rejected the UF_UGMGR Part Name operation.")
        log(listing, "Exception type: {}".format(type(ex).__name__))
        log(listing, "Message: {}".format(ex))
        error_code = getattr(ex, "ErrorCode", None)
        if error_code is not None:
            log(listing, "NX error code: {}".format(error_code))
        log(listing, "")
        log(listing, "No DB_PART_NAME attribute write was attempted.")
        log(listing, "No revision operation was attempted.")
        log(listing, "No retry or bypass was attempted.")
        return

    # Verification 1: query the Teamcenter database again through UF_UGMGR.
    try:
        verified_name, verified_desc = ask_part_name_desc(ugmgr, database_part_tag)
        log(listing, "")
        log(listing, "VERIFY 1 - UF_UGMGR DATABASE READ-BACK")
        log(listing, "Part Name   : {}".format(verified_name))
        log(listing, "Description : {}".format(verified_desc))
    except Exception as ex:
        log(listing, "")
        log(listing, "WARNING: SetPartNameDesc returned, but UF_UGMGR read-back failed.")
        log(listing, "Message: {}".format(ex))
        verified_name = None

    # Verification 2: refresh/read the familiar NX database attribute mapping.
    mapped_name = read_database_attribute(work_part, DB_PART_NAME)
    log(listing, "")
    log(listing, "VERIFY 2 - NX DB_PART_NAME READ-BACK")
    log(listing, "DB_PART_NAME: {}".format(mapped_name))

    log(listing, "")
    if str(verified_name) == new_name:
        log(listing, "PASS: UF_UGMGR reports the Teamcenter Item Name was changed.")
        if mapped_name is not None and str(mapped_name) != new_name:
            log(listing, "NOTE: NX DB_PART_NAME has not refreshed to the new value yet.")
            log(listing, "Reopen/refresh the part in NX and verify the Teamcenter UI.")
    else:
        log(listing, "NOT VERIFIED: The UF_UGMGR database read-back did not match.")
        log(listing, "Expected: {}".format(new_name))
        log(listing, "Actual  : {}".format(verified_name))

    log(listing, "")
    log(listing, "Original name: {}".format(old_name))
    log(listing, "Requested name: {}".format(new_name))
    log(listing, "")
    log(listing, "NOTE: This journal does NOT change Item ID or Revision.")
    log(listing, "=" * 76)


if __name__ == "__main__":
    main()
