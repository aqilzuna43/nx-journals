"""Tests for from_git/utils/single_drawing_scope.py (host-side, no NX)."""

import csv
import importlib.util
import io
import os
import tempfile
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
UTIL = ROOT / "from_git" / "utils" / "single_drawing_scope.py"
REPORTS_ZIP = (
    ROOT / "from_git" / "templates" / "LOGS" / "REPORTS.zip"
)

REQUIRED_COLUMNS = (
    "PART_NUMBER",
    "REVISION",
    "KEEP_DWG_INDEX",
    "EXPECTED_REMOVE_DWG_INDICES",
    "APPROVED",
    "ENGINEER",
    "CONFIRMATION",
)


def load_util():
    spec = importlib.util.spec_from_file_location(
        "single_drawing_scope", UTIL
    )
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def export_row(part, revision, files, count=None, result="SUCCESS"):
    return {
        "DB_PART_NO": part,
        "DB_PART_REV": revision,
        "PDF_FILE_COUNT": str(len(files)) if count is None else str(count),
        "PDF_RESULT": result,
        "PDF_FILES": ";".join(files),
    }


def pdf(part, index):
    return "C:\\out\\{0}_REVA_DWG{1}.pdf".format(part, index)


class ScopeUtilTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.u = load_util()

    def test_two_drawings_defaults_keep_1(self):
        rows, skipped, errors = self.u.scope_rows(
            [export_row("264MN021218A01", "A", [pdf("264MN021218A01", 1), pdf("264MN021218A01", 2)])]
        )
        self.assertEqual(errors, [])
        self.assertEqual(skipped, [])
        self.assertEqual(len(rows), 1)
        row = rows[0]
        self.assertEqual(row["PART_NUMBER"], "264MN021218A01")
        self.assertEqual(row["REVISION"], "A")
        self.assertEqual(row["KEEP_DWG_INDEX"], "1")
        self.assertEqual(row["EXPECTED_REMOVE_DWG_INDICES"], "2")
        self.assertEqual(row["APPROVED"], "NO")
        self.assertEqual(row["ENGINEER"], "")
        self.assertEqual(row["CONFIRMATION"], "")

    def test_three_drawings_expected_pipe_list(self):
        rows, skipped, errors = self.u.scope_rows(
            [export_row("264MN021262A01", "A",
                        [pdf("264MN021262A01", 1), pdf("264MN021262A01", 2), pdf("264MN021262A01", 3)])]
        )
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["EXPECTED_REMOVE_DWG_INDICES"], "2|3")

    def test_files_out_of_order_indices_sorted(self):
        rows, _, errors = self.u.scope_rows(
            [export_row("P1", "A", [pdf("P1", 2), pdf("P1", 1)])]
        )
        self.assertEqual(errors, [])
        self.assertEqual(rows[0]["EXPECTED_REMOVE_DWG_INDICES"], "2")

    def test_keep_lowest_live_when_dwg1_missing(self):
        rows, _, errors = self.u.scope_rows(
            [export_row("P2", "A", [pdf("P2", 2), pdf("P2", 3)])]
        )
        self.assertEqual(errors, [])
        self.assertEqual(rows[0]["KEEP_DWG_INDEX"], "2")
        self.assertEqual(rows[0]["EXPECTED_REMOVE_DWG_INDICES"], "3")

    def test_single_drawing_rows_are_skipped(self):
        rows, skipped, errors = self.u.scope_rows(
            [
                export_row("SINGLE", "A", [pdf("SINGLE", 1)]),
                export_row("NONE", "A", [], count=0, result="SKIPPED_NO_DRAWING"),
            ]
        )
        self.assertEqual(rows, [])
        self.assertEqual(errors, [])
        self.assertEqual(len(skipped), 2)

    def test_count_mismatch_fails_closed(self):
        rows, _, errors = self.u.scope_rows(
            [export_row("BAD1", "A", [pdf("BAD1", 1), pdf("BAD1", 2)], count=3)]
        )
        self.assertEqual(rows, [])
        self.assertEqual(len(errors), 1)
        self.assertIn("PDF_FILE_COUNT=3", errors[0][2])

    def test_unsuffixed_file_fails_closed(self):
        rows, _, errors = self.u.scope_rows(
            [
                export_row(
                    "BAD2", "A",
                    ["C:\\out\\BAD2_REVA.pdf", pdf("BAD2", 2)],
                )
            ]
        )
        self.assertEqual(rows, [])
        self.assertEqual(len(errors), 1)
        self.assertIn("_DWG<n>", errors[0][2])

    def test_non_success_multi_row_fails_closed(self):
        rows, _, errors = self.u.scope_rows(
            [export_row("BAD3", "A", [pdf("BAD3", 1), pdf("BAD3", 2)],
                        result="FAILED")]
        )
        self.assertEqual(rows, [])
        self.assertEqual(len(errors), 1)

    def test_duplicate_part_fails_closed(self):
        rows, _, errors = self.u.scope_rows(
            [
                export_row("DUP", "A", [pdf("DUP", 1), pdf("DUP", 2)]),
                export_row("DUP", "A", [pdf("DUP", 1), pdf("DUP", 2)]),
            ]
        )
        self.assertEqual(rows, [])
        self.assertEqual(len(errors), 1)
        self.assertIn("duplicate", errors[0][2])

    def test_rows_sorted_and_scope_columns_exact(self):
        rows, _, _ = self.u.scope_rows(
            [
                export_row("B", "A", [pdf("B", 1), pdf("B", 2)]),
                export_row("A", "A", [pdf("A", 1), pdf("A", 2)]),
            ]
        )
        self.assertEqual([r["PART_NUMBER"] for r in rows], ["A", "B"])
        self.assertEqual(tuple(rows[0].keys()), REQUIRED_COLUMNS)

    def test_write_scope_round_trip(self):
        rows, _, _ = self.u.scope_rows(
            [export_row("264MN021262A01", "A",
                        [pdf("264MN021262A01", 1), pdf("264MN021262A01", 2), pdf("264MN021262A01", 3)])]
        )
        with tempfile.TemporaryDirectory() as folder:
            output = os.path.join(folder, "scope.csv")
            self.u.write_scope(rows, output)
            with open(output, "r", encoding="utf-8-sig", newline="") as handle:
                reader = csv.DictReader(handle)
                self.assertEqual(tuple(reader.fieldnames), REQUIRED_COLUMNS)
                written = list(reader)
        self.assertEqual(len(written), 1)
        self.assertEqual(written[0]["EXPECTED_REMOVE_DWG_INDICES"], "2|3")
        self.assertEqual(written[0]["APPROVED"], "NO")

    def test_build_scope_file_errors_write_nothing(self):
        export_rows = [
            export_row("BAD", "A", [pdf("BAD", 1)], count=2),
            export_row("GOOD", "A", [pdf("GOOD", 1), pdf("GOOD", 2)]),
        ]
        path = os.path.join(tempfile.mkdtemp(), "export.csv")
        with open(path, "w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=list(export_rows[0].keys()))
            writer.writeheader()
            for item in export_rows:
                writer.writerow(item)
        output = os.path.join(tempfile.mkdtemp(), "scope.csv")
        rows, skipped, errors = self.u.build_scope_file(path, output)
        self.assertTrue(errors)
        self.assertFalse(os.path.exists(output))

    def test_read_export_result_bom_and_plain(self):
        for encoding in ("utf-8-sig", "utf-8"):
            with tempfile.NamedTemporaryFile(
                "w", encoding=encoding, suffix=".csv", delete=False
            ) as handle:
                handle.write(
                    "DB_PART_NO,DB_PART_REV,PDF_FILE_COUNT,PDF_RESULT,PDF_FILES\n"
                    "264MN021218A01,A,2,SUCCESS,C:\\x_REVA_DWG1.pdf;C:\\x_REVA_DWG2.pdf\n"
                )
                name = handle.name
            try:
                export_rows = self.u.read_export_result(name)
            finally:
                os.unlink(name)
            self.assertEqual(len(export_rows), 1)
            self.assertEqual(export_rows[0]["DB_PART_NO"], "264MN021218A01")


class CommittedReportEndToEndTests(unittest.TestCase):
    """Pin the generator against the committed J07 run (REPORTS.zip)."""

    @classmethod
    def setUpClass(cls):
        cls.u = load_util()
        if not REPORTS_ZIP.is_file():
            raise unittest.SkipTest("REPORTS.zip is not present in the repo")
        with zipfile.ZipFile(str(REPORTS_ZIP)) as archive:
            name = "REPORTS/EXPORT_RESULT_20260815_093237.csv"
            with archive.open(name) as handle:
                text = io.TextIOWrapper(handle, encoding="utf-8-sig")
                cls.export_rows = list(csv.DictReader(text))

    def test_known_sixteen_offenders(self):
        rows, skipped, errors = self.u.scope_rows(self.export_rows)
        self.assertEqual(errors, [])
        expected_parts = {
            "264MN021218A01", "264MN021262A01", "264MN021286A01",
            "264MN024185A01", "264MN024854A01", "264MN025545A01",
            "264MN026174A01", "264MN026184A01", "264MN029982A01",
            "264MN029984A01", "264MN030060A01", "264MN030070A01",
            "264MN030114A01", "264MN032670A01", "264MN032671A01",
            "264MN032706A01",
        }
        self.assertEqual(len(rows), len(expected_parts))
        self.assertEqual(
            {row["PART_NUMBER"] for row in rows}, expected_parts
        )
        for row in rows:
            self.assertEqual(row["REVISION"], "A")
            self.assertEqual(row["KEEP_DWG_INDEX"], "1")
            self.assertEqual(row["APPROVED"], "NO")
            self.assertEqual(row["ENGINEER"], "")
            self.assertEqual(row["CONFIRMATION"], "")
        by_part = {row["PART_NUMBER"]: row for row in rows}
        self.assertEqual(
            by_part["264MN021262A01"]["EXPECTED_REMOVE_DWG_INDICES"], "2|3"
        )
        for part in sorted(expected_parts - {"264MN021262A01"}):
            self.assertEqual(
                by_part[part]["EXPECTED_REMOVE_DWG_INDICES"], "2"
            )

    def test_single_drawing_rows_skipped(self):
        rows, skipped, errors = self.u.scope_rows(self.export_rows)
        single = [
            item for item in skipped
            if item[2] == "single or no drawing exported"
        ]
        self.assertEqual(len(single), 205 - 16)


if __name__ == "__main__":
    unittest.main()
