"""CSV input/output helpers for NX journals."""

import csv
import os


def write_csv(path, headers, rows):
    """Write an Excel-friendly UTF-8 CSV report."""
    with open(path, "w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.writer(fh)
        writer.writerow(headers)
        for row in rows:
            writer.writerow(["" if value is None else value for value in row])
    return path


def read_csv_rows(path):
    """Read a CSV file using common Teamcenter/Excel encodings."""
    encodings = ("utf-8-sig", "utf-8", "cp1252")
    last_error = None
    for encoding in encodings:
        try:
            with open(path, "r", encoding=encoding, newline="") as fh:
                return [row for row in csv.reader(fh)]
        except UnicodeDecodeError as exc:
            last_error = exc
    raise RuntimeError(
        f"Unable to read CSV with supported encodings (utf-8-sig, utf-8, cp1252): {path}"
    ) from last_error


def find_newest_csv(folder, prefix=None):
    """Return the newest CSV path in a folder, optionally preferring a prefix."""
    csvs = [name for name in os.listdir(folder) if name.lower().endswith(".csv")]
    if prefix:
        preferred = [name for name in csvs if name.lower().startswith(prefix.lower())]
        csvs = preferred or csvs
    if not csvs:
        return None
    return max((os.path.join(folder, name) for name in csvs), key=os.path.getmtime)
