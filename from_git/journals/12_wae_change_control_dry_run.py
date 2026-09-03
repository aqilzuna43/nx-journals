# NX Journal 12 - WAE Change Control (Option B) DRY RUN
#
# Purpose
# -------
# Read-only preflight for a Teamcenter-controlled freeze/unfreeze process where:
# - DB_PART_REV remains the formal TCX revision.
# - WAE_VERSION is the working iteration within that TCX revision.
# - FREEZE means "baseline/check-in intent".
# - UNFREEZE means "checkout intent + WAE_VERSION increment intent".
#
# This DRY_RUN journal DOES NOT:
# - check out or check in any Teamcenter object;
# - write WAE_VERSION or any other attribute;
# - save any NX part;
# - create a TCX revision;
# - modify J04/J05 behavior.
#
# Run inside NX X / NX 2506 with a managed work part loaded.

import datetime
import json
import os

import NXOpen


BUILD = "J12-WAE-CHANGE-CONTROL-OPTION-B-DRY-RUN-V1"
MODE = "DRY_RUN"

# Set to "FREEZE" or "UNFREEZE" for the scenario you want to simulate.
USER_ACTION = os.environ.get("NX_J12_ACTION", "UNFREEZE").strip().upper()

ATTRIBUTE_CATEGORY = "WAEItem"
WAE_VERSION_TITLE = "WAE_VERSION"
DB_PART_NO_TITLE = "DB_PART_NO"
DB_PART_REV_TITLE = "DB_PART_REV"
CHECKED_OUT_USER_TITLE = "CHECKED_OUT_USER"
CHECKED_OUT_TITLE = "CHECKED_OUT"


def _listing_window(session):
    lw = session.ListingWindow
    try:
        lw.Open()
    except Exception:
        pass
    return lw


def _write(lw, text=""):
    try:
        lw.WriteLine(str(text))
    except Exception:
        pass


def _get_attrs(part):
    attrs = []
    try:
        infos = part.GetUserAttributes()
    except Exception:
        infos = []
    for info in infos:
        try:
            value = info.StringValue
        except Exception:
            try:
                value = info.RealValue
            except Exception:
                try:
                    value = info.IntegerValue
                except Exception:
                    value = ""
        attrs.append(
            {
                "category": str(getattr(info, "Category", "") or ""),
                "title": str(getattr(info, "Title", "") or ""),
                "type": str(getattr(info, "Type", "") or ""),
                "value": value,
            }
        )
    return attrs


def _find_attr(attrs, title, category=None):
    exact = []
    for a in attrs:
        if a["title"] != title:
            continue
        if category is not None and a["category"] != category:
            continue
        exact.append(a)
    if exact:
        return exact[0]

    # Fallback by title only for NX/TCX system attributes whose category can vary.
    if category is not None:
        for a in attrs:
            if a["title"] == title:
                return a
    return None


def _part_name(part):
    for attr in ("FullPath", "Name", "Leaf"):
        try:
            value = getattr(part, attr)
            if value:
                return str(value)
        except Exception:
            pass
    return "<unknown>"


def _is_managed(part):
    name = _part_name(part)
    if name.startswith("@DB/"):
        return True
    try:
        return bool(NXOpen.Session.GetSession().IsManagedMode)
    except Exception:
        return False


def _normalize_text(value):
    if value is None:
        return ""
    return str(value).strip()


def _parse_version(value):
    raw = _normalize_text(value)
    if not raw:
        return None, "WAE_VERSION is blank."
    try:
        parsed = int(raw)
    except Exception:
        return None, "WAE_VERSION must be an integer for controlled increment; found {!r}.".format(raw)
    if parsed < 1:
        return None, "WAE_VERSION must be >= 1; found {}.".format(parsed)
    return parsed, ""


def _checkout_snapshot(attrs):
    checked_out_user = _find_attr(attrs, CHECKED_OUT_USER_TITLE)
    checked_out = _find_attr(attrs, CHECKED_OUT_TITLE)

    user_value = _normalize_text(checked_out_user["value"]) if checked_out_user else ""
    checked_value = _normalize_text(checked_out["value"]) if checked_out else ""

    if user_value:
        state = "CHECKED_OUT"
    elif checked_value and checked_value not in ("0", "FALSE", "N", "NO"):
        state = "CHECKED_OUT_OR_FLAGGED"
    else:
        state = "NOT_POSITIVELY_CHECKED_OUT"

    return {
        "state": state,
        "checked_out_user": user_value,
        "checked_out_raw": checked_value,
    }


def _decision(action, managed, revision, version, version_error, checkout):
    reasons = []

    if action not in ("FREEZE", "UNFREEZE"):
        return {
            "result": "BLOCKED",
            "reasons": ["NX_J12_ACTION must be FREEZE or UNFREEZE."],
            "simulated_next_wae_version": None,
            "would_write": False,
            "would_save": False,
            "would_checkout": False,
            "would_checkin": False,
        }

    if not managed:
        reasons.append("Active work part is not positively identified as Teamcenter-managed.")
    if not revision:
        reasons.append("DB_PART_REV is blank or unavailable.")
    if version_error:
        reasons.append(version_error)

    next_version = None
    would_checkout = False
    would_checkin = False

    if action == "UNFREEZE" and version is not None:
        next_version = version + 1
        would_checkout = True
    elif action == "FREEZE":
        next_version = version
        would_checkin = True

    # DRY_RUN never performs writes/saves even when the simulated production path would.
    return {
        "result": "READY_FOR_PRODUCTION_IMPLEMENTATION" if not reasons else "BLOCKED",
        "reasons": reasons,
        "simulated_next_wae_version": next_version,
        "current_checkout_state": checkout["state"],
        "would_write": action == "UNFREEZE" and not reasons,
        "would_save": not reasons,
        "would_checkout": would_checkout and not reasons,
        "would_checkin": would_checkin and not reasons,
    }


def main():
    session = NXOpen.Session.GetSession()
    lw = _listing_window(session)

    _write(lw, "=" * 72)
    _write(lw, BUILD)
    _write(lw, "Mode: {}".format(MODE))
    _write(lw, "Requested scenario: {}".format(USER_ACTION))
    _write(lw, "=" * 72)

    part = session.Parts.Work
    if part is None:
        _write(lw, "BLOCKED: No active work part.")
        return

    attrs = _get_attrs(part)
    managed = _is_managed(part)

    part_no_attr = _find_attr(attrs, DB_PART_NO_TITLE)
    rev_attr = _find_attr(attrs, DB_PART_REV_TITLE)
    wae_attr = _find_attr(attrs, WAE_VERSION_TITLE, ATTRIBUTE_CATEGORY)

    part_no = _normalize_text(part_no_attr["value"]) if part_no_attr else ""
    revision = _normalize_text(rev_attr["value"]) if rev_attr else ""
    wae_raw = _normalize_text(wae_attr["value"]) if wae_attr else ""
    wae_version, wae_error = _parse_version(wae_raw)
    checkout = _checkout_snapshot(attrs)

    decision = _decision(
        USER_ACTION,
        managed,
        revision,
        wae_version,
        wae_error,
        checkout,
    )

    result = {
        "build": BUILD,
        "timestamp": datetime.datetime.now().astimezone().isoformat(timespec="seconds"),
        "mode": MODE,
        "requested_action": USER_ACTION,
        "work_part": {
            "name": _part_name(part),
            "managed": managed,
            "db_part_no": part_no,
            "db_part_rev": revision,
            "wae_version_raw": wae_raw,
            "wae_version": wae_version,
        },
        "checkout": checkout,
        "decision": decision,
        "safety": {
            "writes_performed": False,
            "save_performed": False,
            "checkout_performed": False,
            "checkin_performed": False,
            "revision_created": False,
        },
    }

    _write(lw, "Part: {}".format(part_no or _part_name(part)))
    _write(lw, "TCX managed: {}".format(managed))
    _write(lw, "DB_PART_REV: {}".format(revision or "<blank>"))
    _write(lw, "WAE_VERSION: {}".format(wae_raw or "<blank>"))
    _write(lw, "Checkout state: {}".format(checkout["state"]))
    if checkout["checked_out_user"]:
        _write(lw, "Checked-out user token: {}".format(checkout["checked_out_user"]))

    _write(lw, "-")
    _write(lw, "DRY-RUN DECISION: {}".format(decision["result"]))
    for reason in decision["reasons"]:
        _write(lw, "  BLOCKER: {}".format(reason))

    if USER_ACTION == "UNFREEZE" and wae_version is not None:
        _write(
            lw,
            "  SIMULATION: WAE_VERSION {} -> {}".format(
                wae_version, decision["simulated_next_wae_version"]
            ),
        )
        _write(lw, "  Production intent: explicit TCX checkout -> increment -> verify -> save")
    elif USER_ACTION == "FREEZE":
        _write(lw, "  SIMULATION: WAE_VERSION remains {}".format(wae_raw or "<blank>"))
        _write(lw, "  Production intent: validate/save -> TCX check-in / freeze baseline")

    _write(lw, "-")
    _write(lw, "SAFETY: no checkout, no check-in, no attribute write, no save, no revision creation.")

    # Write a local audit JSON only. This is intentionally outside Teamcenter/NX data.
    try:
        root = os.environ.get("TEMP") or os.environ.get("TMP") or os.getcwd()
        stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        path = os.path.join(root, "J12_WAE_CHANGE_CONTROL_DRY_RUN_{}.json".format(stamp))
        with open(path, "w", encoding="utf-8") as fh:
            json.dump(result, fh, indent=2, sort_keys=True)
        _write(lw, "Audit JSON: {}".format(path))
    except Exception as exc:
        _write(lw, "WARNING: Could not write local audit JSON: {}".format(exc))

    _write(lw, "=" * 72)


if __name__ == "__main__":
    main()
