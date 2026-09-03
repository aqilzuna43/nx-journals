"""Journal 30 - freeze one selected CAD component at its current WAE version.

Preselect exactly one component row in Assembly Navigator, then run this
journal from its NX UI button.  APPLY saves only the selected component's
loaded prototype, checks in only that prototype, and verifies that the formal
Teamcenter revision and WAE_VERSION did not change.  A component that is
already checked in is verified without mutation.

This journal never writes WAE_VERSION and never creates or revises a
Teamcenter Item Revision.  Target: NX X 2506 embedded Python.
"""

import importlib.util
import os


BUILD = "J30-NX2506-CAD-FREEZE-V1"
USER_MODE = "APPLY"  # APPLY or DRY_RUN


def _load_common():
    path = os.path.abspath(
        os.path.join(os.path.dirname(__file__), "..", "utils", "wae_change_control.py")
    )
    if not os.path.isfile(path):
        raise RuntimeError("WAE change-control helper not found: " + path)
    spec = importlib.util.spec_from_file_location("nx_wae_change_control_j30", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def main():
    common = _load_common()
    common.run_ui("FREEZE", BUILD, USER_MODE, "NX_J30_MODE")


if __name__ == "__main__":
    main()
