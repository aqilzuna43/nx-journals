import re
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
JOURNAL = ROOT / "Assembly" / "Diagnostic" / "NX_Assembly_Load_Diagnostic.vb"
README = ROOT / "Assembly" / "Diagnostic" / "README.md"
EXAMPLE = ROOT / "Assembly" / "Diagnostic" / "Example_Report.txt"


class AssemblyLoadDiagnosticContractTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.source = JOURNAL.read_text(encoding="utf-8")
        cls.readme = README.read_text(encoding="utf-8")
        cls.example = EXAMPLE.read_text(encoding="utf-8")

    def test_required_deliverables_exist(self):
        self.assertTrue(JOURNAL.is_file())
        self.assertTrue(README.is_file())
        self.assertTrue(EXAMPLE.is_file())

    def test_all_required_statuses_are_implemented(self):
        for status in (
            "OK",
            "MISSING_FILE",
            "PROTOTYPE_UNAVAILABLE",
            "UNLOADED",
            "INVALID_OBJECT",
            "ERROR",
        ):
            self.assertIn(f'"{status}"', self.source)

    def test_recursive_scan_uses_explicit_stack_and_continues_per_component(self):
        self.assertIn("Stack(Of ScanNode)", self.source)
        self.assertIn("While stack.Count > 0", self.source)
        self.assertIn("node.Component.GetChildren()", self.source)
        self.assertIn("record.AddIssue(\"Component.GetChildren\", ex)", self.source)
        self.assertIn("MaxComponentOccurrences", self.source)

    def test_missing_file_has_unloaded_prototype_fallback(self):
        self.assertIn("UFSession.GetUFSession()", self.source)
        self.assertIn("AskInstOfPartOcc", self.source)
        self.assertIn("AskPartNameOfChild", self.source)
        self.assertRegex(
            self.source,
            re.compile(
                r"If prototype Is Nothing Then.*?MISSING_FILE.*?"
                r"PROTOTYPE_UNAVAILABLE",
                re.DOTALL,
            ),
        )

    def test_invalid_om_signature_and_failed_operation_are_reported(self):
        self.assertIn("IM0541", self.source)
        self.assertIn("invalid or unsuitable OM object", self.source)
        self.assertIn('WriteField(writer, "Failed operation"', self.source)
        self.assertIn('WriteField(writer, "Exception"', self.source)

    def test_report_contains_hierarchy_health_and_teamcenter_fields(self):
        for label in (
            "Parent assembly",
            "Assembly path",
            "Level",
            "Part number / Item ID",
            "Revision",
            "Dataset / prototype",
            "Managed status",
            "Load state",
            "Reference set",
            "Recommended corrective action",
        ):
            self.assertIn(f'"{label}"', self.source)

    def test_output_contract_and_listing_window_progress(self):
        self.assertIn("NX_Assembly_Load_Diagnostic_Report.txt", self.source)
        self.assertIn("NX_JOURNALS_IO_DIR", self.source)
        for message in (
            "NX Assembly Diagnostic Started...",
            "Scanning assembly:",
            "Components found:",
            "Errors found:",
            "Report generated successfully.",
        ):
            self.assertIn(message, self.source)

    def test_journal_remains_diagnostic_only(self):
        forbidden_calls = (
            ".Save(",
            ".SaveAs(",
            ".Suppress(",
            ".Unsuppress(",
            ".AddComponent(",
            ".ReplaceComponent(",
            "EnsurePartsLoadedFully(",
            "OpenBase(",
        )
        for forbidden in forbidden_calls:
            self.assertNotIn(forbidden, self.source)

    def test_operator_docs_explain_run_and_statuses(self):
        self.assertIn("Tools > Journal > Play", self.readme)
        self.assertIn("read-only", self.readme)
        for status in (
            "MISSING_FILE",
            "PROTOTYPE_UNAVAILABLE",
            "UNLOADED",
            "INVALID_OBJECT",
        ):
            self.assertIn(status, self.readme)
        self.assertIn("Assembly path:", self.example)
        self.assertIn("Failed operation:", self.example)


if __name__ == "__main__":
    unittest.main()
