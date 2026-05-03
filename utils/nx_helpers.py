"""NX 2312 helper utilities shared across journals."""

import os
import traceback

import NXOpen
import NXOpen.UF


def get_session():
    """Return the active NX session."""
    return NXOpen.Session.GetSession()


def get_work_part(session=None):
    """Return the current work part."""
    session = session or get_session()
    return session.Parts.Work


def get_listing_window(session=None):
    """Open and return the NX Listing Window."""
    session = session or get_session()
    lw = session.ListingWindow
    lw.Open()
    return lw


def log_lines(session, lines):
    """Write lines to stdout and the NX Listing Window."""
    lw = get_listing_window(session)
    for line in lines:
        for text in str(line).splitlines() or [""]:
            print(text)
            lw.WriteFullline(text)


def log_info(session, message):
    log_lines(session, [message])


def log_error(session, message):
    log_lines(session, [f"ERROR: {message}"])


def run_journal(main_func):
    """Run a journal entry point and surface uncaught errors in NX."""
    session = None
    try:
        session = get_session()
        main_func(session)
    except Exception:
        lines = ["ERROR: Unhandled journal exception.", traceback.format_exc()]
        if session is not None:
            log_lines(session, lines)
        else:
            for line in lines:
                print(line)


def require_work_part(session):
    """Return the work part, or report a clear error and return None."""
    part = get_work_part(session)
    if part is None:
        log_error(session, "No work part loaded.")
        return None
    return part


def prompt_folder(title):
    """
    Open an NX folder selection dialog with the given title.
    Returns the selected folder path string, or None if the user cancels.
    """
    ui = NXOpen.UI.GetUI()
    dialog = ui.CreateFolderSelectionDialog()
    dialog.SetTitle(title)
    try:
        response = dialog.Show()
        if response == NXOpen.SelectionDialog.DialogResponse.Ok:
            return dialog.Path
        return None
    finally:
        dialog.Destroy()


def get_root_component(part):
    """Return the root component for an assembly work part, or None."""
    try:
        return part.ComponentAssembly.RootComponent
    except Exception:
        return None


def traverse_assembly(component):
    """Yield every child occurrence below an assembly component."""
    if component is None:
        return
    for child in component.GetChildren():
        yield child
        yield from traverse_assembly(child)


def iter_occurrences(part):
    """Yield child occurrence components for a part's assembly tree."""
    root = get_root_component(part)
    if root is None:
        return
    yield from traverse_assembly(root)


def unique_prototype_parts(part, include_work_part=True):
    """Return unique prototype parts from the work part and assembly children."""
    unique = []
    seen = set()

    def add_part(candidate):
        if candidate is None:
            return
        key = getattr(candidate, "Tag", None) or getattr(candidate, "FullPath", None) or id(candidate)
        if key in seen:
            return
        seen.add(key)
        unique.append(candidate)

    if include_work_part:
        add_part(part)

    for component in iter_occurrences(part):
        add_part(component.Prototype)

    return unique


def safe_part_name(part, fallback="part"):
    """Return a filesystem-friendly part name."""
    try:
        full_path = part.FullPath
        if full_path:
            return os.path.splitext(os.path.basename(full_path))[0] or fallback
    except Exception:
        pass
    try:
        return part.Leaf or fallback
    except Exception:
        return fallback


def get_string_attr(nxobj, attr_name, fallback=""):
    """Read a string user attribute; return fallback if absent or unreadable."""
    if nxobj is None or not attr_name:
        return fallback
    try:
        attr = nxobj.GetUserAttribute(attr_name, NXOpen.NXObject.AttributeType.String, -1)
        return attr.StringValue.strip()
    except Exception:
        return fallback


def set_string_attr(nxobj, attr_name, value):
    """Write a string user attribute using the conservative NX2312 update path."""
    nxobj.SetUserAttribute(attr_name, -1, value, NXOpen.Update.Option.Now)
