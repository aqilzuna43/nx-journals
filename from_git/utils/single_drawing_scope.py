"""Build a J25 single-drawing scope CSV from a J07 EXPORT_RESULT report.

Host-side utility. Pure Python 3.8+; no NXOpen, no Teamcenter, no NX runtime.

Why this exists
---------------
J07 exports one PDF per drawing specification it finds beneath a 3D master.
A master that carries more than one spec drawing (dwg1, dwg2, ...) produces
multiple PDFs, which violates the one-spec-drawing-per-master rule. J25
(25_tc_single_drawing_cleanup.py) removes the approved extra specifications,
but it needs an operator CSV with one row per offender.

This tool turns any J07 EXPORT_RESULT_<timestamp>.csv into that J25 scope
CSV. The report's PDF_FILES column already names every exported drawing as
<...>_DWG<n>.pdf, so the live dwg indices per part can be read back exactly.

Default policy (operator-overridable in the produced CSV):
- KEEP_DWG_INDEX defaults to 1 when dwg1 is live, otherwise the lowest live
  index.
- EXPECTED_REMOVE_DWG_INDICES is every other live index, pipe-separated.
- APPROVED stays NO and ENGINEER/CONFIRMATION stay blank: J25 APPLY_APPROVED
  still requires the human to review (DRY_RUN), sign the row, and flip
  USER_MODE. This tool never authorizes a deletion.

Fail-closed rules (any violation => no output file is written):
- a row reports PDF_FILE_COUNT > 1 but its PDF_FILES do not parse to the same
  count of _DWG<n> names;
- PDF_RESULT is not SUCCESS for a multi-drawing row;
- the same PART_NUMBER/REVISION appears more than once.

Usage:
    python from_git/utils/single_drawing_scope.py \
        --input  REPORTS/EXPORT_RESULT_20260815_093237.csv \
        --output NX_TC_SINGLE_DRAWING_SCOPE.csv

Output columns match J25's required input columns exactly:
PART_NUMBER, REVISION, KEEP_DWG_INDEX, EXPECTED_REMOVE_DWG_INDICES,
APPROVED, ENGINEER, CONFIRMATION.
"""

import argparse
import csv
import os
import re
import sys

SCOPE_COLUMNS = (
    "PART_NUMBER",
    "REVISION",
    "KEEP_DWG_INDEX",
    "EXPECTED_REMOVE_DWG_INDICES",
    "APPROVED",
    "ENGINEER",
    "CONFIRMATION",
)

DWG_SUFFIX = re.compile(r"_dwg(?P<index>[1-9][0-9]*)\.pdf$", re.IGNORECASE)


def basename_only(value):
    return os.path.basename(str(value).replace("\\", "/")).strip()


def dwg_index_from_file(value):
    """Return the dwg index encoded in one J07 PDF filename, or None."""
    match = DWG_SUFFIX.search(basename_only(value))
    if match is None:
        return None
    return int(match.group("index"))


def split_files(value):
    text = str(value or "").strip()
    if not text:
        return []
    return [item.strip() for item in text.split(";") if item.strip()]


def parse_row_indices(row):
    """Parse live dwg indices of one export-result row.

    Returns (indices, error). indices is a sorted unique int list.
    error is a message when the row is a multi-drawing row that cannot be
    trusted; otherwise None.
    """
    count_text = str(row.get("PDF_FILE_COUNT") or "").strip()
    try:
        count = int(count_text) if count_text else 0
    except ValueError:
        return [], "PDF_FILE_COUNT is not an integer: {0}".format(count_text)
    if count <= 1:
        return [], None
    result_text = str(row.get("PDF_RESULT") or "").strip().upper()
    if result_text != "SUCCESS":
        return [], (
            "multi-drawing row has PDF_RESULT={0}; only SUCCESS rows "
            "become cleanup candidates.".format(result_text or "<blank>")
        )
    indices = []
    for name in split_files(row.get("PDF_FILES")):
        index = dwg_index_from_file(name)
        if index is None:
            return [], (
                "PDF file does not carry a _DWG<n> suffix: {0}".format(name)
            )
        indices.append(index)
    unique = sorted(set(indices))
    if len(unique) != count:
        return [], (
            "PDF_FILE_COUNT={0} but PDF_FILES encodes {1} drawing(s): {2}".format(
                count, len(unique), str(row.get("PDF_FILES") or "")
            )
        )
    return unique, None


def scope_rows(export_rows):
    """Map J07 EXPORT_RESULT dict rows to J25 scope dict rows.

    Returns (rows, skipped, errors).
    rows      : list of dicts over SCOPE_COLUMNS, sorted by part/revision.
    skipped   : list of (part, revision, reason) for benign exclusions.
    errors    : list of (part, revision, reason); any error means the caller
                should refuse to write output (fail closed).
    """
    rows = []
    skipped = []
    errors = []
    seen = {}
    for source in export_rows:
        part = str(source.get("DB_PART_NO") or "").strip()
        revision = str(source.get("DB_PART_REV") or "").strip()
        key = (part, revision)
        if not part:
            continue
        if key in seen:
            errors.append((part, revision, "duplicate export row"))
            # Drop the earlier occurrence too: a duplicated part cannot be
            # turned into a trusted single-row scope.
            rows = [row for row in rows
                    if (row["PART_NUMBER"], row["REVISION"]) != key]
            skipped = [item for item in skipped if item[:2] != key]
            continue
        seen[key] = True
        indices, error = parse_row_indices(source)
        if error is not None:
            errors.append((part, revision, error))
            continue
        if not indices:
            skipped.append((part, revision, "single or no drawing exported"))
            continue
        keep = 1 if 1 in indices else indices[0]
        expected = [index for index in indices if index != keep]
        rows.append(
            {
                "PART_NUMBER": part,
                "REVISION": revision,
                "KEEP_DWG_INDEX": str(keep),
                "EXPECTED_REMOVE_DWG_INDICES": "|".join(
                    str(index) for index in expected
                ),
                "APPROVED": "NO",
                "ENGINEER": "",
                "CONFIRMATION": "",
            }
        )
    rows.sort(
        key=lambda item: (item["PART_NUMBER"], item["REVISION"])
    )
    return rows, skipped, errors


def write_scope(rows, output_path):
    """Write J25-compatible scope rows to output_path (UTF-8, no BOM)."""
    with open(output_path, "w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=SCOPE_COLUMNS)
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


def read_export_result(path):
    last_error = None
    for encoding in ("utf-8-sig", "utf-8", "cp1252"):
        try:
            with open(path, "r", encoding=encoding, newline="") as handle:
                return list(csv.DictReader(handle))
        except UnicodeDecodeError as error:
            last_error = error
    raise ValueError("Unable to decode {0}: {1}".format(path, last_error))


def build_scope_file(input_path, output_path):
    """Generate output_path from an EXPORT_RESULT CSV; returns (rows, skipped,
    errors). Raises ValueError when input is unreadable; refuses to write the
    output when any fail-closed error exists."""
    export_rows = read_export_result(input_path)
    rows, skipped, errors = scope_rows(export_rows)
    if errors:
        return rows, skipped, errors
    write_scope(rows, output_path)
    return rows, skipped, errors


def main(argv):
    parser = argparse.ArgumentParser(
        description=(
            "Generate a J25 NX_TC_SINGLE_DRAWING_SCOPE.csv from a J07 "
            "EXPORT_RESULT_<timestamp>.csv."
        )
    )
    parser.add_argument(
        "--input", required=True, help="Path to the J07 EXPORT_RESULT CSV."
    )
    parser.add_argument(
        "--output",
        default="NX_TC_SINGLE_DRAWING_SCOPE.csv",
        help="Output scope path (default: NX_TC_SINGLE_DRAWING_SCOPE.csv).",
    )
    args = parser.parse_args(argv)
    if not os.path.isfile(args.input):
        print("ERROR: input file not found: {0}".format(args.input))
        return 1
    try:
        rows, skipped, errors = build_scope_file(args.input, args.output)
    except ValueError as error:
        print("ERROR: {0}".format(error))
        return 2
    for part, revision, reason in skipped:
        print(
            "SKIP {0}/{1}: {2}".format(part, revision, reason)
        )
    if errors:
        print("ERROR: {0} row(s) cannot be trusted; no output written.".format(
            len(errors)
        ))
        for part, revision, reason in errors:
            print("  ERROR {0}/{1}: {2}".format(part, revision, reason))
        return 2
    print(
        "WROTE {0} scope row(s) to {1} (skipped {2}).".format(
            len(rows), args.output, len(skipped)
        ))
    print(
        "Review in J25 DRY_RUN, then set APPROVED=YES, ENGINEER, and "
        "CONFIRMATION=REMOVE_EXTRA_DRAWINGS per row before APPLY_APPROVED."
    )
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
