"""Journal 31 - unfreeze one selected CAD component and advance WAE_VERSION.

Preselect exactly one component row in Assembly Navigator, then run this
journal from its NX UI button.  APPLY requires the selected component's
loaded prototype to be checked in, explicitly checks out only that prototype,
increments WAEItem/WAE_VERSION by exactly one, rereads it, saves it, and
leaves it checked out for CAD editing.

Rerunning against a checked-out component is blocked, preventing an accidental
second increment.  This journal never changes DB_PART_REV or creates a formal
Teamcenter revision.  Target: NX X 2506 embedded Python.
"""

import importlib.util
import os


BUILD = "J31-NX2506-CAD-UNFREEZE-V1"
USER_MODE = "APPLY"  # APPLY or DRY_RUN


def _load_common():
    path = os.path.abspath(
        os.path.join(os.path.dirname(__file__), "..", "utils", "wae_change_control.py")
    )
    if not os.path.isfile(path):
        raise RuntimeError("WAE change-control helper not found: " + path)
    spec = importlib.util.spec_from_file_location("nx_wae_change_control_j31", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def main():
    common = _load_common()
    common.run_ui("UNFREEZE", BUILD, USER_MODE, "NX_J31_MODE")


if __name__ == "__main__":
    main()
