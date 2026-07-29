"""J16 - Teamcenter X specification drawing import (NX X 2506)."""

import os
import sys

_JOURNAL_DIR = os.path.dirname(os.path.abspath(__file__))
if _JOURNAL_DIR not in sys.path:
    sys.path.insert(0, _JOURNAL_DIR)

from _16_tc_specification_import_v4_workflow import *


if __name__ == "__main__":
    main()
