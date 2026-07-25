"""Journal 11 - guarded Teamcenter checkout and attribute persistence test.

PROBE is the default and never changes NX/Teamcenter data.

FULL_REVERSIBLE requires all of:
  NX_J11_ALLOW_MUTATION=YES
  NX_J11_EXPECTED_PART_NUMBER=<disposable item>
  NX_J11_EXPECTED_REVISION=<exact revision>
  NX_J11_ATTRIBUTE=<business attribute title>
  NX_J11_TEST_VALUE=<temporary non-blank value>

The full test checks out the active 3D master, writes and saves one temporary
business value, reopens and verifies it, restores the original value, saves and
reopens again, and leaves the part checked out for review. It never checks in.
"""

import json
import os
import sys
import traceback
from datetime import datetime

import NXOpen

try:
    import NXOpen.UF
except Exception:
    NXOpen.UF = None


VALID_MODES = ("PROBE", "FULL_REVERSIBLE")


def _text(value):
    return "" if value is None else str(value)


def _clean(value):
    return _text(value).strip()


def _normalized(value):
    return " ".join(_clean(value).split()).upper()


def _dispose(value):
    if value is None:
        return
    for name in ("Dispose", "FreeResource"):
        method = getattr(value, name, None)
        if callable(method):
            try:
                method()
            except Exception:
                pass
            return


def _exception_record(error):
    return {
        "type": type(error).__name__,
        "error_code": _text(getattr(error, "ErrorCode", "")),
        "message": _text(error),
    }


def _runtime_root():
    script_parent = os.path.dirname(os.path.abspath(__file__))
    configured = _clean(os.environ.get("NX_JOURNALS_ROOT"))
    candidates = [
        configured,
        os.path.join(configured, "from_git") if configured else "",
        os.path.dirname(script_parent),
        os.getcwd(),
        os.path.join(os.getcwd(), "from_git"),
    ]
    for candidate in candidates:
        if candidate and os.path.isfile(
            os.path.join(
                os.path.abspath(candidate),
                "config",
                "attribute_reconciliation.json",
            )
        ):
            return os.path.abspath(candidate)
    raise RuntimeError(
        "Configuration not found. Deploy config beside journals or set "
        "NX_JOURNALS_ROOT."
    )


def _load_config():
    path = os.path.join(
        _runtime_root(), "config", "attribute_reconciliation.json"
    )
    with open(path, "r", encoding="utf-8-sig") as handle:
        return json.load(handle)


def _io_root():
    configured = _clean(os.environ.get("NX_JOURNALS_IO_DIR"))
    if configured:
        return os.path.abspath(configured)
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    return desktop if os.path.isdir(desktop) else os.getcwd()


def _rule_map(config):
    rules = {
        rule["logical_name"]: rule for rule in config.get("attributes", [])
    }
    result = {}
    for mapping in config["update_workflow"]["business_columns"]:
        rule = rules[mapping["logical_name"]]
        result[rule["attribute"]] = rule
    return result


def _enum_name(value):
    name = getattr(value, "name", None)
    return _text(name if name is not None else value).split(".")[-1]


def _attribute_value(info):
    kind = _enum_name(getattr(info, "Type", ""))
    numeric_kind = getattr(info, "Type", None)
    if kind in ("String", "5") or numeric_kind == 5:
        return getattr(info, "StringValue", "")
    if kind in ("Real", "Number", "4") or numeric_kind == 4:
        return getattr(info, "RealValue", None)
    if kind in ("Integer", "3") or numeric_kind == 3:
        return getattr(info, "IntegerValue", None)
    if kind in ("Boolean", "1") or numeric_kind == 1:
        return getattr(info, "BooleanValue", None)
    return getattr(info, "StringValue", "")


def _read_attribute(part, rule):
    iterator = None
    try:
        iterator = part.CreateAttributeIterator()
        iterator.SetIncludeOnlyCategory(rule["category"])
        iterator.SetIncludeOnlyTitle(rule["attribute"])
        iterator.SetIncludeAlsoUnset(True)
        matches = [
            info
            for info in part.GetUserAttributes(iterator)
            if _clean(getattr(info, "Category", "")) == rule["category"]
            and _clean(getattr(info, "Title", "")) == rule["attribute"]
        ]
        if not matches:
            return {
                "status": "MISSING",
                "raw": "",
                "flags": {},
            }
        if len(matches) != 1:
            return {
                "status": "AMBIGUOUS",
                "raw": "",
                "flags": {},
            }
        info = matches[0]
        raw = _attribute_value(info)
        return {
            "status": (
                "UNSET"
                if bool(getattr(info, "Unset", False))
                else ("BLANK" if _clean(raw) == "" else "POPULATED")
            ),
            "raw": raw,
            "flags": {
                "locked": bool(getattr(info, "Locked", False)),
                "owned_by_system": bool(
                    getattr(info, "OwnedBySystem", False)
                ),
                "pdm_based": bool(getattr(info, "PdmBased", False)),
            },
        }
    finally:
        _dispose(iterator)


def _identity(part, config):
    rules = {
        rule["logical_name"]: rule for rule in config["attributes"]
    }

    def read(rule):
        try:
            return _clean(part.GetStringAttribute(rule["attribute"]))
        except AttributeError:
            return _clean(_read_attribute(part, rule)["raw"])
        except Exception:
            return _clean(_read_attribute(part, rule)["raw"])

    return (
        read(rules["part_number"]),
        read(rules["revision"]),
    )


def _read_only(part):
    value = getattr(part, "IsReadOnly", None)
    try:
        value = value() if callable(value) else value
    except Exception:
        return None
    return None if value is None else bool(value)


def _managed_mode(session):
    value = getattr(session, "IsManagedMode", False)
    try:
        return bool(value() if callable(value) else value)
    except Exception:
        return False


def _pdm_part(part):
    value = getattr(part, "PDMPart", None)
    return value() if callable(value) else value


def _checkout_status(work_part, target):
    assembly = getattr(work_part, "ComponentAssembly", None)
    method = getattr(assembly, "GetCheckedoutStatusOfObjects", None)
    if not callable(method):
        return {
            "state": "API_UNAVAILABLE",
            "detail": "",
        }
    try:
        result = method()
        if not isinstance(result, tuple) or len(result) < 2:
            return {
                "state": "UNRECOGNIZED_API_RESULT",
                "detail": repr(result),
            }
        checked = {
            _text(getattr(item, "Tag", id(item))) for item in result[0]
        }
        unchecked = {
            _text(getattr(item, "Tag", id(item))) for item in result[1]
        }
        key = _text(getattr(target, "Tag", id(target)))
        if key in checked:
            state = "CHECKED_OUT"
        elif key in unchecked:
            state = "NOT_CHECKED_OUT"
        else:
            state = "NOT_RETURNED"
        return {"state": state, "detail": ""}
    except Exception as exc:
        return {
            "state": "STATUS_ERROR",
            "detail": _exception_record(exc),
        }


def _autolock_probe():
    result = {"status": "UNAVAILABLE", "value": None}
    try:
        uf_module = getattr(NXOpen, "UF", None)
        uf_session_type = getattr(uf_module, "UFSession", None)
        get_session = getattr(uf_session_type, "GetUFSession", None)
        if not callable(get_session):
            return result
        uf_session = get_session()
        manager = getattr(uf_session, "Ugmgr", None)
        for name in ("AskAutolockStatus", "AskAutoLockStatus"):
            method = getattr(manager, name, None)
            if callable(method):
                result.update(status="AVAILABLE", value=bool(method()))
                return result
    except Exception as exc:
        result.update(status="ERROR", error=_exception_record(exc))
    return result


def _simple_members(value):
    names = []
    try:
        candidates = dir(value)
    except Exception:
        return names
    for name in candidates:
        lower = name.lower()
        if any(
            token in lower
            for token in (
                "check",
                "pdm",
                "read",
                "reserve",
                "save",
                "lock",
                "reopen",
            )
        ):
            names.append(name)
    return sorted(names)


def _enum_members(container, name):
    try:
        return [
            member
            for member in dir(getattr(container, name))
            if not member.startswith("_")
        ]
    except Exception:
        return []


def probe(session, part, config):
    part_number, revision = _identity(part, config)
    pdm_part = _pdm_part(part)
    return {
        "python_version": sys.version,
        "nxopen_module": _text(getattr(NXOpen, "__file__", "")),
        "ugii_version": _clean(os.environ.get("UGII_VERSION")),
        "ugii_full_version": _clean(
            os.environ.get("UGII_FULL_VERSION")
        ),
        "managed_mode": _managed_mode(session),
        "part_number": part_number,
        "revision": revision,
        "part_identifier": _clean(
            getattr(part, "JournalIdentifier", "")
        )
        or _clean(getattr(part, "Name", "")),
        "read_only": _read_only(part),
        "checkout_status": _checkout_status(part, part),
        "autolock": _autolock_probe(),
        "part_members": _simple_members(part),
        "pdm_part_type": (
            type(pdm_part).__name__ if pdm_part is not None else ""
        ),
        "pdm_part_members": _simple_members(pdm_part),
        "session_members": _simple_members(session),
        "close_whole_tree_enum": _enum_members(
            NXOpen.BasePart, "CloseWholeTree"
        ),
        "close_modified_enum": _enum_members(
            NXOpen.BasePart, "CloseModified"
        ),
    }


def _builder_data_type(rule):
    enum = NXOpen.AttributePropertiesBaseBuilder.DataTypeOptions
    kind = _normalized(rule["type"])
    if kind == "BOOLEAN":
        return enum.Boolean
    if kind == "INTEGER":
        return enum.Integer
    if kind in ("NUMBER", "REAL"):
        return enum.Number
    return enum.String


def _set_builder_value(builder, rule, value):
    kind = _normalized(rule["type"])
    if kind == "BOOLEAN":
        builder.BooleanValue = (
            NXOpen.AttributePropertiesBaseBuilder.BooleanValueOptions.TrueValue
            if _normalized(value) in ("Y", "YES", "TRUE", "1")
            else NXOpen.AttributePropertiesBaseBuilder.BooleanValueOptions.FalseValue
        )
    elif kind == "INTEGER":
        builder.IntegerValue = int(value)
    elif kind in ("NUMBER", "REAL"):
        builder.NumberValue = float(value)
    else:
        builder.StringValue = _text(value)


def _write_attribute(session, part, rule, value):
    builder = None
    try:
        builder = session.AttributeManager.CreateAttributePropertiesBuilder(
            part,
            [part],
            NXOpen.AttributePropertiesBuilder.OperationType.Save,
        )
        builder.Category = rule["category"]
        builder.Title = rule["attribute"]
        builder.DataType = _builder_data_type(rule)
        _set_builder_value(builder, rule, value)
        builder.Commit()
    finally:
        _dispose(builder)


def _save(part):
    status = None
    try:
        status = part.Save(
            NXOpen.BasePart.SaveComponents.FalseValue,
            NXOpen.BasePart.CloseAfterSave.FalseValue,
        )
        if int(getattr(status, "NumberUnsavedParts", 0)) or int(
            getattr(status, "NumberUnsavedObjects", 0)
        ):
            raise RuntimeError("NX reported unsaved objects after save.")
    finally:
        _dispose(status)


def _enum_value(container, enum_name, member_names):
    enum = getattr(container, enum_name)
    for name in member_names:
        value = getattr(enum, name, None)
        if value is not None:
            return value
    raise RuntimeError(
        "{0} has none of: {1}".format(enum_name, ", ".join(member_names))
    )


def _unwrap(value):
    if isinstance(value, tuple):
        primary = value[0] if value else None
        for transient in value[1:]:
            _dispose(transient)
        return primary
    return value


def _reopen(part):
    whole_tree = _enum_value(
        NXOpen.BasePart,
        "CloseWholeTree",
        ("FalseValue", "CloseWholeTreeFalse"),
    )
    close_modified = _enum_value(
        NXOpen.BasePart,
        "CloseModified",
        ("CloseModified", "CloseModifiedCloseModified"),
    )
    return _unwrap(part.Reopen(whole_tree, close_modified, None))


def _checkout(part):
    pdm_part = _pdm_part(part)
    method = getattr(pdm_part, "Checkout", None)
    if not callable(method):
        raise RuntimeError("PDMPart.Checkout is unavailable.")
    method()
    if _read_only(part) is True:
        raise RuntimeError("Part remained read-only after checkout.")


def _assert_value(part, rule, expected, stage):
    actual = _read_attribute(part, rule)
    if _text(actual.get("raw", "")) != _text(expected):
        raise RuntimeError(
            "{0} verification failed. Expected {1!r}, observed {2!r}.".format(
                stage, expected, actual.get("raw", "")
            )
        )
    return actual


def _validate_test_value(value, rule):
    if _clean(value) == "":
        raise RuntimeError("Temporary test value must not be blank.")
    kind = _normalized(rule.get("type"))
    if kind in ("NUMBER", "REAL"):
        try:
            float(value)
        except ValueError:
            raise RuntimeError("Temporary test value is not a number.")
    elif kind == "INTEGER":
        try:
            int(value)
        except ValueError:
            raise RuntimeError("Temporary test value is not an integer.")
    allowed = rule.get("allowed_values")
    if allowed and _normalized(value) not in [
        _normalized(item) for item in allowed
    ]:
        raise RuntimeError(
            "Temporary test value is outside the controlled value set."
        )


def full_reversible_test(session, part, config, payload):
    if _normalized(os.environ.get("NX_J11_ALLOW_MUTATION")) != "YES":
        raise RuntimeError(
            "FULL_REVERSIBLE requires NX_J11_ALLOW_MUTATION=YES."
        )
    expected_part = _clean(
        os.environ.get("NX_J11_EXPECTED_PART_NUMBER")
    )
    expected_revision = _clean(
        os.environ.get("NX_J11_EXPECTED_REVISION")
    )
    attribute_title = _clean(os.environ.get("NX_J11_ATTRIBUTE"))
    test_value = _clean(os.environ.get("NX_J11_TEST_VALUE"))
    if not all(
        (expected_part, expected_revision, attribute_title, test_value)
    ):
        raise RuntimeError(
            "Expected part, revision, business attribute, and test value are "
            "all required."
        )
    actual_identity = _identity(part, config)
    if (
        _normalized(actual_identity[0]) != _normalized(expected_part)
        or _normalized(actual_identity[1]) != _normalized(
            expected_revision
        )
    ):
        raise RuntimeError(
            "Active part identity does not match the explicit disposable-item "
            "guard."
        )
    rules = _rule_map(config)
    rule = rules.get(attribute_title)
    if rule is None:
        raise RuntimeError(
            "NX_J11_ATTRIBUTE is not in the business allowlist."
        )
    _validate_test_value(test_value, rule)
    original = _read_attribute(part, rule)
    if original["status"] in ("AMBIGUOUS",):
        raise RuntimeError("Selected attribute is ambiguous.")
    flags = original.get("flags", {})
    if (
        flags.get("locked")
        or flags.get("owned_by_system")
        or flags.get("pdm_based")
    ):
        raise RuntimeError(
            "Selected attribute has runtime non-writable flags."
        )
    if _text(original.get("raw", "")) == test_value:
        raise RuntimeError("Temporary value must differ from the original.")

    original_value = original.get("raw", "")
    payload["mutation"] = {
        "attribute": attribute_title,
        "category": rule["category"],
        "original_value": original_value,
        "test_value": test_value,
        "checkout_before": _checkout_status(part, part),
        "read_only_before": _read_only(part),
        "steps": [],
        "restoration_required": False,
    }
    mutation = payload["mutation"]
    current_part = part
    temporary_saved = False
    restoration_confirmed = False
    try:
        _checkout(current_part)
        mutation["steps"].append("CHECKOUT_PASS")
        mutation["checkout_after"] = _checkout_status(
            current_part, current_part
        )
        mutation["read_only_after"] = _read_only(current_part)

        _write_attribute(
            session, current_part, rule, test_value
        )
        _assert_value(current_part, rule, test_value, "Immediate temporary")
        mutation["steps"].append("TEMPORARY_REREAD_PASS")
        _save(current_part)
        temporary_saved = True
        mutation["steps"].append("TEMPORARY_SAVE_PASS")

        current_part = _reopen(current_part)
        _assert_value(current_part, rule, test_value, "Reopen temporary")
        mutation["steps"].append("TEMPORARY_REOPEN_PASS")

        _write_attribute(
            session, current_part, rule, original_value
        )
        _assert_value(
            current_part, rule, original_value, "Immediate restoration"
        )
        mutation["steps"].append("RESTORE_REREAD_PASS")
        _save(current_part)
        mutation["steps"].append("RESTORE_SAVE_PASS")

        current_part = _reopen(current_part)
        _assert_value(
            current_part, rule, original_value, "Reopen restoration"
        )
        restoration_confirmed = True
        mutation["steps"].append("RESTORE_REOPEN_PASS")
        mutation["result"] = "PASS"
    except Exception as exc:
        mutation["result"] = "FAIL"
        mutation["error"] = _exception_record(exc)
        if temporary_saved and not restoration_confirmed:
            mutation["restoration_required"] = True
            try:
                _write_attribute(
                    session, current_part, rule, original_value
                )
                _save(current_part)
                _assert_value(
                    current_part,
                    rule,
                    original_value,
                    "Emergency restoration",
                )
                mutation["steps"].append("EMERGENCY_RESTORE_PASS")
                mutation["restoration_required"] = False
            except Exception as restore_exc:
                mutation["restoration_error"] = _exception_record(
                    restore_exc
                )
        raise
    finally:
        mutation["final_checkout"] = _checkout_status(
            current_part, current_part
        )
        mutation["final_read_only"] = _read_only(current_part)
        mutation["automatic_checkin"] = "NEVER"
    return current_part


def _listing(session):
    listing = getattr(session, "ListingWindow", None)
    if listing is not None:
        try:
            listing.Open()
        except Exception:
            pass
    return listing


def _log(listing, message):
    if listing is not None:
        try:
            listing.WriteLine(_text(message))
            return
        except Exception:
            pass
    print(_text(message))


def _write_payload(payload):
    output_root = _io_root()
    os.makedirs(output_root, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = os.path.join(
        output_root, "J11_CHECKOUT_ACCEPTANCE_{0}.json".format(stamp)
    )
    with open(path, "w", encoding="utf-8") as handle:
        json.dump(payload, handle, indent=2, ensure_ascii=False)
    return path


def main(session):
    listing = _listing(session)
    config = _load_config()
    part = getattr(getattr(session, "Parts", None), "Work", None)
    if part is None:
        raise RuntimeError("Open the disposable 3D master part first.")
    mode = _normalized(os.environ.get("NX_J11_MODE") or "PROBE")
    if mode not in VALID_MODES:
        raise RuntimeError(
            "NX_J11_MODE must be PROBE or FULL_REVERSIBLE."
        )
    payload = {
        "journal": "11_test_teamcenter_attribute_checkout.py",
        "timestamp": datetime.now().isoformat(),
        "mode": mode,
        "probe": probe(session, part, config),
        "result": "PROBE_COMPLETE" if mode == "PROBE" else "RUNNING",
    }
    error = None
    try:
        if mode == "FULL_REVERSIBLE":
            full_reversible_test(
                session, part, config, payload
            )
            payload["result"] = "PASS"
    except Exception as exc:
        error = exc
        payload["result"] = (
            "RESTORATION_REQUIRED"
            if payload.get("mutation", {}).get("restoration_required")
            else "FAIL"
        )
        payload["error"] = _exception_record(exc)
        payload["traceback"] = traceback.format_exc()
    path = _write_payload(payload)
    _log(listing, "Journal 11 result: " + payload["result"])
    _log(listing, "Journal 11 evidence: " + path)
    if payload["result"] == "RESTORATION_REQUIRED":
        _log(
            listing,
            "CRITICAL: RESTORATION_REQUIRED. Leave the item checked out and "
            "restore the original value before further work.",
        )
    _log(listing, "Journal 11 never checks the part in.")
    if error is not None:
        raise error
    return path


if __name__ == "__main__":
    main(NXOpen.Session.GetSession())
