"""Journal 30 - freeze active CAD or selected assembly components.

Preselect one or more component rows in Assembly Navigator and run this
journal to freeze their unique loaded prototypes. Geometry is normalized via
its owning component. With no resolved component selection, the active work
part is the sole target. APPLY validates and processes each exact Teamcenter
part/revision independently, saves and checks in operator-owned writable
targets, applies Teamcenter's configured Part_Freeze_Process one part at a
time, and verifies a frozen/read-only status without changing the formal
revision or WAE_VERSION. Invalid targets are skipped and reported while safe
targets continue. Both positive numeric WAE values and alphabetic WAE values
matching DB_PART_REV are valid freeze baselines.

This journal never writes WAE_VERSION and never creates or revises a
Teamcenter Item Revision.  Target: NX X 2506 embedded Python.
"""

import importlib.util
import os


BUILD = "J30-NX2506-CAD-FREEZE-V5"
EXPECTED_COMMON_BUILD = "WAE-CHANGE-CONTROL-V5"
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
    actual_build = getattr(module, "COMMON_BUILD", "<missing>")
    if actual_build != EXPECTED_COMMON_BUILD:
        raise RuntimeError(
            "J30 helper version mismatch: expected {0}, loaded {1} from {2}".format(
                EXPECTED_COMMON_BUILD, actual_build, path
            )
        )
    return module


def main():
    common = _load_common()
    common.run_ui("FREEZE", BUILD, USER_MODE, "NX_J30_MODE")


if __name__ == "__main__":
    main()
