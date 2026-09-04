"""Journal 34 - validate the fixed administrative CAD freeze CSV."""

import importlib.util
import os


BUILD = "J34-NX2506-VALIDATE-ADMIN-FREEZE-V1"
EXPECTED_COMMON_BUILD = "NX-ADMIN-FREEZE-V1"


def _load_common():
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "admin_freeze_common.py")
    spec = importlib.util.spec_from_file_location("nx_admin_freeze_j34", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    actual = getattr(module, "COMMON_BUILD", "<missing>")
    if actual != EXPECTED_COMMON_BUILD:
        raise RuntimeError(
            "J34 helper version mismatch: expected {0}, loaded {1}.".format(
                EXPECTED_COMMON_BUILD, actual
            )
        )
    return module


def main():
    _load_common().run_ui("VALIDATE", BUILD)


if __name__ == "__main__":
    main()
