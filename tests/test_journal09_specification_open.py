import importlib.util
import os
import sys
import types
import unittest
from pathlib import Path
from unittest import mock


ROOT = Path(__file__).resolve().parents[1]
JOURNAL = (
    ROOT
    / "from_git"
    / "journals"
    / "09_test_teamcenter_specification_open.py"
)


def load_journal():
    sys.modules.setdefault("NXOpen", types.ModuleType("NXOpen"))
    spec = importlib.util.spec_from_file_location("journal09", JOURNAL)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class SpecificationOpenAcceptanceTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.journal = load_journal()

    def test_default_scope_matches_nx2506_acceptance_part(self):
        with mock.patch.dict(os.environ, {}, clear=True):
            self.assertEqual(
                self.journal.resolve_test_scope(),
                ("264MN028607A01", "A", 1, 3, False),
            )

    def test_scope_can_be_overridden_for_another_drawing(self):
        environment = {
            "NX_TEST_PART_NO": "TEST-PART",
            "NX_TEST_PART_REV": "B",
            "NX_TEST_DWG_INDEX": "2",
            "NX_TEST_EXPECTED_SHEET_COUNT": "4",
            "NX_TEST_KEEP_OPEN": "1",
        }
        with mock.patch.dict(os.environ, environment, clear=True):
            self.assertEqual(
                self.journal.resolve_test_scope(),
                ("TEST-PART", "B", 2, 4, True),
            )

    def test_exact_identifier_and_sheet_count_pass(self):
        identifier = (
            "@DB/264MN028607A01/A/specification/"
            "264MN028607A01-A-dwg1"
        )
        status, message = self.journal.evaluate_opened_drawing(
            identifier,
            identifier,
            3,
            3,
        )
        self.assertEqual(status, "SUCCESS")
        self.assertIn("expected sheet count", message)

    def test_wrong_sheet_count_fails(self):
        status, message = self.journal.evaluate_opened_drawing(
            "expected",
            "expected",
            2,
            3,
        )
        self.assertEqual(status, "FAILED_UNEXPECTED_SHEET_COUNT")
        self.assertIn("returned 2", message)

    def test_different_identifier_is_not_accepted(self):
        status, _message = self.journal.evaluate_opened_drawing(
            "expected",
            "different",
            3,
            3,
        )
        self.assertEqual(status, "WARNING_IDENTIFIER_DIFFERENT")

    def test_runtime_identity_marks_closed_spec_test(self):
        self.assertEqual(
            self.journal.JOURNAL_BUILD_ID,
            "J09-NX2506-CLOSED-SPEC-OPEN-V1",
        )
        self.assertTrue(
            self.journal.runtime_source_path().endswith(
                "09_test_teamcenter_specification_open.py"
            )
        )


if __name__ == "__main__":
    unittest.main()
