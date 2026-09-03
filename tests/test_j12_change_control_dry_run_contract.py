import pathlib
import unittest


ROOT = pathlib.Path(__file__).resolve().parents[1]
J12 = ROOT / "from_git" / "journals" / "12_wae_change_control_dry_run.py"


class TestJ12DryRunContract(unittest.TestCase):
    def test_journal_exists(self):
        self.assertTrue(J12.exists())

    def test_dry_run_has_no_mutation_calls(self):
        text = J12.read_text(encoding="utf-8")
        forbidden = [
            ".SetUserAttribute(",
            ".Save(",
            "PdmSession.Checkout",
            "PdmSession.Checkin",
            "CreateNewRevision",
        ]
        for token in forbidden:
            self.assertNotIn(token, text, msg=token)

    def test_expected_contract_is_present(self):
        text = J12.read_text(encoding="utf-8")
        for token in [
            'MODE = "DRY_RUN"',
            '"FREEZE"',
            '"UNFREEZE"',
            'WAE_VERSION_TITLE = "WAE_VERSION"',
            'DB_PART_REV_TITLE = "DB_PART_REV"',
            '"writes_performed": False',
            '"revision_created": False',
        ]:
            self.assertIn(token, text)


if __name__ == "__main__":
    unittest.main()
