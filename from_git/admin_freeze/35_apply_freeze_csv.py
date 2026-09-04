"""Journal 35 - apply the previously validated administrative CAD freeze CSV."""

import importlib.util
import os


BUILD = "J35-NX2506-APPLY-ADMIN-FREEZE-V1"
EXPECTED_COMMON_BUILD = "NX-ADMIN-FREEZE-V1"


def _load_common():
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "admin_freeze_common.py")
    spec = importlib.util.spec_from_file_location("nx_admin_freeze_j35", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    actual = getattr(module, "COMMON_BUILD", "<missing>")
    if actual != EXPECTED_COMMON_BUILD:
        raise RuntimeError(
            "J35 helper version mismatch: expected {0}, loaded {1}.".format(
                EXPECTED_COMMON_BUILD, actual
            )
        )
    return module


def main():
    _load_common().run_ui("APPLY", BUILD)


if __name__ == "__main__":
    main()
