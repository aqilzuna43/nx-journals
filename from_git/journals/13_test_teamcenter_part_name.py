# NXOpen Python journal
# Purpose: Test changing Teamcenter Item Name (DB_PART_NAME) for the CURRENT work part only.
# Scope: Part Name only. No revision change. No CSV/bulk processing.
#
# IMPORTANT:
# - Run only on a disposable / dummy Teamcenter item.
# - This changes the Teamcenter Item Name, not the Item ID / Part Number.
# - The journal deliberately does NOT auto-checkout or save geometry.
# - Edit NEW_PART_NAME below before running.

import NXOpen
import NXOpen.PDM

NEW_PART_NAME = "DUMMY-PART-NAME-CHANGE-TEST"
ATTRIBUTE_NAME = "DB_PART_NAME"


def log(listing_window, text=""):
    listing_window.WriteLine(str(text))


def safe_dispose(obj):
    if obj is None:
        return
    try:
        obj.Dispose()
    except Exception:
        pass


def read_part_name(pdm_part):
    manager = None
    try:
        manager = pdm_part.NewDatabaseAttributeManager()
        manager.LoadAttributes(True)
        return manager.GetAttribute(ATTRIBUTE_NAME)
    finally:
        safe_dispose(manager)


def main():
    session = NXOpen.Session.GetSession()
    listing = session.ListingWindow
    listing.Open()

    work_part = session.Parts.Work

    log(listing, "=" * 72)
    log(listing, "TEAMCENTER PART NAME CHANGE - DUMMY TEST")
    log(listing, "=" * 72)

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

    new_name = NEW_PART_NAME.strip()

    if not new_name:
        log(listing, "FAIL: NEW_PART_NAME is blank. Nothing changed.")
        return

    try:
        pdm_part = work_part.PDMPart
    except Exception as ex:
        log(listing, "FAIL: Could not access work_part.PDMPart.")
        log(listing, "This journal must be run in Teamcenter managed mode.")
        log(listing, "Exception: {}".format(ex))
        return

    if pdm_part is None:
        log(listing, "FAIL: PDMPart is unavailable. Is this NX session Teamcenter-managed?")
        return

    manager = None
    old_name = None

    try:
        manager = pdm_part.NewDatabaseAttributeManager()
        manager.LoadAttributes(True)

        old_name = manager.GetAttribute(ATTRIBUTE_NAME)

        log(listing, "")
        log(listing, "Current DB_PART_NAME : {}".format(old_name))
        log(listing, "Requested new name  : {}".format(new_name))
        log(listing, "")

        if old_name == new_name:
            log(listing, "NO CHANGE: Teamcenter Part Name already matches NEW_PART_NAME.")
            return

        manager.SetAttribute(ATTRIBUTE_NAME, new_name)
        log(listing, "Staged DB_PART_NAME update.")
        log(listing, "Calling StoreAttributes() ...")

        manager.StoreAttributes()

        log(listing, "StoreAttributes() completed without an NXOpen exception.")

    except Exception as ex:
        log(listing, "")
        log(listing, "FAIL: Teamcenter rejected or could not store the Part Name change.")
        log(listing, "Exception type: {}".format(type(ex).__name__))
        log(listing, "Message: {}".format(ex))
        error_code = getattr(ex, "ErrorCode", None)
        if error_code is not None:
            log(listing, "NX error code: {}".format(error_code))
        log(listing, "")
        log(listing, "No revision operation was attempted.")
        log(listing, "If this is a permissions/checkout error, capture this Listing Window output.")
        return

    finally:
        safe_dispose(manager)

    try:
        verified_name = read_part_name(pdm_part)

        log(listing, "")
        log(listing, "VERIFY FROM TEAMCENTER")
        log(listing, "DB_PART_NAME read-back: {}".format(verified_name))

        if verified_name == new_name:
            log(listing, "")
            log(listing, "PASS: Teamcenter Part Name was changed successfully.")
            log(listing, "Old name: {}".format(old_name))
            log(listing, "New name: {}".format(verified_name))
        else:
            log(listing, "")
            log(listing, "WARNING: StoreAttributes() returned successfully,")
            log(listing, "but Teamcenter read-back does not match the requested name.")
            log(listing, "Expected: {}".format(new_name))
            log(listing, "Actual  : {}".format(verified_name))
            log(listing, "Treat this as NOT VERIFIED.")

    except Exception as ex:
        log(listing, "")
        log(listing, "WARNING: Change was submitted, but fresh read-back failed.")
        log(listing, "Message: {}".format(ex))
        log(listing, "Check the Item Name directly in Teamcenter before rerunning.")

    log(listing, "")
    log(listing, "NOTE: This journal does not change DB_PART_REV and performs no revision action.")
    log(listing, "=" * 72)


if __name__ == "__main__":
    main()
