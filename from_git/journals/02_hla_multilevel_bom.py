"""Journal 02 - NX-authoritative draft multilevel BOM.

Exports the active NX assembly in the FZ NX-import schema.  This journal is a
draft snapshot only; Journal 04 is the certification gate.
"""

import os
import sys
from datetime import datetime

import NXOpen
import NXOpen.UF

_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _REPO_ROOT not in sys.path:
    sys.path.insert(0, _REPO_ROOT)

from utils.attribute_reconciliation import (  # noqa: E402
    FZ_BOM_COLUMNS,
    bom_export_rows,
    collect_bom_nodes,
    load_config,
    write_csv,
)
from utils.nx_helpers import (  # noqa: E402
    get_output_folder,
    log_info,
    require_work_part,
    run_journal,
    safe_part_name,
)


def main(session):
    part = require_work_part(session)
    if part is None:
        return

    config = load_config(_REPO_ROOT)
    nodes, findings = collect_bom_nodes(part, config)
    rows = bom_export_rows(nodes)
    root_part_number = rows[0]["DB_PART_NO"] if rows else safe_part_name(part)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = os.path.join(
        get_output_folder(), "BOM_DRAFT_{0}_{1}.csv".format(root_part_number, timestamp)
    )
    write_csv(output_path, FZ_BOM_COLUMNS, rows)

    messages = [
        "Draft NX BOM exported.",
        "  Certification : DRAFT (run Journal 04 to certify)",
        "  BOM rows       : {0}".format(len(rows)),
        "  Findings       : {0}".format(len(findings)),
        "  Report         : {0}".format(output_path),
    ]
    for finding in findings:
        messages.append("  {0}: {1}".format(finding.get("code"), finding.get("message")))
    log_info(session, "\n".join(messages))


if __name__ == "__main__":
    run_journal(main)
