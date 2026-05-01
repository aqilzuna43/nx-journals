"""Common NX Open helper utilities shared across journals."""

import NXOpen
import NXOpen.UF


def get_session():
    """Return the active NX session."""
    return NXOpen.Session.GetSession()


def get_work_part():
    """Return the current work part."""
    session = get_session()
    return session.Parts.Work


def prompt_folder(title):
    """
    Open a NX folder selection dialog with the given title.
    Returns the selected folder path string, or None if the user cancels.
    """
    ui = NXOpen.UI.GetUI()
    dialog = ui.CreateFolderSelectionDialog()
    dialog.SetTitle(title)
    response = dialog.Show()
    if response == NXOpen.SelectionDialog.DialogResponse.Ok:
        path = dialog.Path
        dialog.Destroy()
        return path
    dialog.Destroy()
    return None


def traverse_assembly(component, visited=None):
    """
    Recursive depth-first generator that yields every NXOpen.Assemblies.Component
    in the assembly tree rooted at *component*.

    Skips already-visited prototype tags to handle reused components without
    infinite loops, but still yields every occurrence (visiting logic only
    guards recursion, not yielding).

    Usage:
        root = part.ComponentAssembly.RootComponent
        for comp in traverse_assembly(root):
            ...
    """
    if visited is None:
        visited = set()

    if component is None:
        return

    for child in component.GetChildren():
        yield child

        proto_tag = child.Prototype.Tag if child.Prototype is not None else None
        if proto_tag is not None and proto_tag not in visited:
            visited.add(proto_tag)
            yield from traverse_assembly(child, visited)
