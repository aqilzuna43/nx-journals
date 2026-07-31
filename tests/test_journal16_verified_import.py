import importlib.util
import os
import sys
import tempfile
import types
import unittest
from pathlib import Path
from unittest import mock


ROOT = Path(__file__).resolve().parents[1]
J16_PATH = (
    ROOT / "from_git" / "journals" / "16_tc_offline_drawing_import.py"
)
J19_PATH = (
    ROOT
    / "from_git"
    / "journals"
    / "19_test_teamcenter_drawing_import_contract.py"
)


def load_j16():
    nxopen = types.ModuleType("NXOpen")
    nxuf = types.ModuleType("NXOpen.UF")
    nxopen.UF = nxuf
    nxopen.BasePart = types.SimpleNamespace(
        CloseWholeTree=types.SimpleNamespace(FalseValue="FalseValue"),
        CloseModified=types.SimpleNamespace(CloseModified="CloseModified"),
    )
    sys.modules["NXOpen"] = nxopen
    sys.modules["NXOpen.UF"] = nxuf
    spec = importlib.util.spec_from_file_location("journal16_verified", J16_PATH)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class FakeLog:
    def __init__(self):
        self.lines = []

    def write(self, message=""):
        self.lines.append(str(message))


class FakeLoadStatus:
    def __init__(self):
        self.disposed = False

    def Dispose(self):
        self.disposed = True


class FakePdmFile:
    def __init__(self, name):
        self.name = name
        self.released = False

    def GetFileName(self):
        return self.name

    def FreeResource(self):
        self.released = True


class VerifiedImportTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.journal = load_j16()

    def make_row(
        self,
        drawing,
        part_number="TEST100",
        revision="A",
        drawing_index=1,
        baseline="",
    ):
        return {
            "_CSV_ROW": 2,
            "PART_NUMBER": part_number,
            "REVISION": revision,
            "DWG_INDEX": str(drawing_index),
            "DRAWING_FILE": drawing,
            "DRAWING_IDENTIFIER": self.journal.drawing_id(
                part_number, revision, drawing_index
            ),
            "EXPORT_SHA256": baseline,
            "APPROVED": "YES",
            "ENGINEER": "TESTER",
        }

    def make_drawing(self, folder, part_number="TEST100", revision="A", index=1):
        path = os.path.join(
            folder,
            self.journal.expected_native(part_number, revision, index),
        )
        with open(path, "wb") as handle:
            handle.write(
                ("drawing-{0}-{1}-{2}".format(part_number, revision, index)).encode(
                    "ascii"
                )
            )
        return path

    def checked_in(self, identifier=""):
        return {
            "state": "CHECKED_IN",
            "owner": "",
            "raw": "(False, '')",
            "opened_identifier": identifier,
        }

    def checked_out(self, owner="other.user", identifier=""):
        return {
            "state": "CHECKED_OUT",
            "owner": owner,
            "raw": "(True, '{0}')".format(owner),
            "opened_identifier": identifier,
        }

    def test_manual_settings_are_fail_safe(self):
        self.assertEqual("", self.journal.USER_IMPORT_CSV)
        self.assertEqual("DRY_RUN", self.journal.USER_MODE)
        with mock.patch.dict(os.environ, {}, clear=True):
            self.assertEqual("DRY_RUN", self.journal.configured_mode())

    def test_checkout_result_decodes_state_and_owner(self):
        checked = self.journal.decode_checkout_result((True, "other.user"))
        clear = self.journal.decode_checkout_result((False, ""))
        unknown = self.journal.decode_checkout_result(None)
        self.assertEqual("CHECKED_OUT", checked["state"])
        self.assertEqual("other.user", checked["owner"])
        self.assertEqual("CHECKED_IN", clear["state"])
        self.assertEqual("UNKNOWN", unknown["state"])

    def test_exact_open_returns_authoritative_checkout_owner(self):
        identifier = self.journal.drawing_id("TEST100", "A", 1)
        status = FakeLoadStatus()
        part = types.SimpleNamespace(
            JournalIdentifier=identifier,
            PDMPart=types.SimpleNamespace(
                GetCheckedoutStatusAndUser=mock.Mock(
                    return_value=(True, "other.user")
                )
            ),
            Close=mock.Mock(),
        )
        parts = mock.MagicMock()
        parts.__iter__.return_value = iter([])
        parts.FindObject.side_effect = RuntimeError("not loaded")
        parts.OpenBase.return_value = (part, status)
        session = types.SimpleNamespace(Parts=parts)

        result = self.journal.inspect_target_checkout(
            session, identifier, FakeLog()
        )

        self.assertEqual("CHECKED_OUT", result["state"])
        self.assertEqual("other.user", result["owner"])
        self.assertEqual(identifier, result["opened_identifier"])
        self.assertTrue(status.disposed)
        part.Close.assert_called_once()

    def test_checked_out_row_is_blocked_without_export(self):
        report = {"RESULT": "LOCAL_PREFLIGHT_OK"}
        proposal = {
            "report": report,
            "identifier": "exact",
            "part_number": "TEST100",
            "revision": "A",
            "drawing_index": 1,
        }
        checkout = self.checked_out("other.user", "exact")
        log = FakeLog()
        with mock.patch.object(
            self.journal,
            "inspect_target_checkout",
            return_value=checkout,
        ), mock.patch.object(
            self.journal, "retrieve_exact_associated_drawing"
        ) as export:
            self.journal.run_managed_preflight(
                object(), object(), [proposal], "evidence", log
            )

        self.assertEqual("BLOCKED_CHECKED_OUT", report["RESULT"])
        self.assertEqual("other.user", report["CHECKOUT_OWNER"])
        self.assertIn("other.user", report["MESSAGE"])
        self.assertTrue(
            any(
                "state=CHECKED_OUT" in line and "owner=other.user" in line
                for line in log.lines
            )
        )
        export.assert_not_called()

    def test_unknown_checkout_is_blocked(self):
        report = {"RESULT": "LOCAL_PREFLIGHT_OK"}
        proposal = {
            "report": report,
            "identifier": "exact",
            "part_number": "TEST100",
            "revision": "A",
            "drawing_index": 1,
        }
        checkout = {
            "state": "UNKNOWN",
            "owner": "",
            "raw": "unrecognized",
            "opened_identifier": "exact",
        }
        with mock.patch.object(
            self.journal,
            "inspect_target_checkout",
            return_value=checkout,
        ):
            self.journal.run_managed_preflight(
                object(), object(), [proposal], "evidence", FakeLog()
            )
        self.assertEqual("BLOCKED_CHECKOUT_UNKNOWN", report["RESULT"])
        self.assertIn("unrecognized", report["MESSAGE"])

    def test_blocked_row_does_not_prevent_clear_row_preflight(self):
        blocked_report = {"RESULT": "LOCAL_PREFLIGHT_OK"}
        clear_report = {
            "RESULT": "LOCAL_PREFLIGHT_OK",
            "BASELINE_SHA256": "",
        }
        proposals = [
            {
                "report": blocked_report,
                "identifier": "blocked",
                "part_number": "BLOCKED",
                "revision": "A",
                "drawing_index": 1,
            },
            {
                "report": clear_report,
                "identifier": "clear",
                "part_number": "CLEAR",
                "revision": "A",
                "drawing_index": 1,
                "preflight_sha": "source",
            },
        ]
        with mock.patch.object(
            self.journal,
            "inspect_target_checkout",
            side_effect=[
                self.checked_out("other.user", "blocked"),
                self.checked_in("clear"),
            ],
        ), mock.patch.object(
            self.journal,
            "retrieve_exact_associated_drawing",
            return_value=("baseline.prt", "teamcenter"),
        ) as export:
            self.journal.run_managed_preflight(
                object(), object(), proposals, "evidence", FakeLog()
            )
        self.assertEqual("BLOCKED_CHECKED_OUT", blocked_report["RESULT"])
        self.assertEqual("MANAGED_PREFLIGHT_OK", clear_report["RESULT"])
        export.assert_called_once()

    def test_csv_baseline_mismatch_blocks_stale_target(self):
        report = {
            "RESULT": "LOCAL_PREFLIGHT_OK",
            "BASELINE_SHA256": "expected",
        }
        proposal = {
            "report": report,
            "identifier": "exact",
            "part_number": "TEST100",
            "revision": "A",
            "drawing_index": 1,
            "preflight_sha": "source",
        }
        with mock.patch.object(
            self.journal,
            "inspect_target_checkout",
            return_value=self.checked_in("exact"),
        ), mock.patch.object(
            self.journal,
            "retrieve_exact_associated_drawing",
            return_value=("baseline.prt", "different"),
        ):
            self.journal.run_managed_preflight(
                object(), object(), [proposal], "evidence", FakeLog()
            )
        self.assertEqual("BLOCKED_STALE_TARGET", report["RESULT"])

    def test_verified_apply_requires_matching_post_export_hash(self):
        with tempfile.TemporaryDirectory() as folder:
            drawing = self.make_drawing(folder)
            source_sha = self.journal.sha256(drawing)
            row = self.make_row(drawing)
            identifier = row["DRAWING_IDENTIFIER"]

            def exported(_session, _fm, _proposal, _root, phase, *_fields):
                digest = source_sha if phase == "POSTIMPORT" else "tc-before"
                return (phase + ".prt", digest)

            with mock.patch.object(
                self.journal,
                "inspect_target_checkout",
                return_value=self.checked_in(identifier),
            ), mock.patch.object(
                self.journal,
                "retrieve_exact_associated_drawing",
                side_effect=exported,
            ), mock.patch.object(
                self.journal, "import_one"
            ) as import_one:
                reports = self.journal.execute(
                    object(),
                    object(),
                    object(),
                    [row],
                    os.path.join(folder, "input.csv"),
                    "stamp",
                    "APPLY_APPROVED",
                    FakeLog(),
                    os.path.join(folder, "evidence"),
                )

        self.assertEqual("IMPORT_VERIFIED", reports[0]["RESULT"])
        self.assertEqual("VERIFIED_SHA256", reports[0]["POST_IMPORT_VERIFICATION"])
        self.assertEqual("YES", reports[0]["WRITE_ATTEMPTED"])
        self.assertEqual(2, import_one.call_count)

    def test_post_export_mismatch_stops_later_writes(self):
        with tempfile.TemporaryDirectory() as folder:
            first = self.make_drawing(folder, "FIRST")
            second = self.make_drawing(folder, "SECOND")
            rows = [
                self.make_row(first, "FIRST"),
                self.make_row(second, "SECOND"),
            ]
            identifiers = [row["DRAWING_IDENTIFIER"] for row in rows]

            def exported(_session, _fm, proposal, _root, phase, *_fields):
                if phase == "POSTIMPORT":
                    return ("post.prt", "not-the-source")
                return ("before.prt", "before-" + proposal["part_number"])

            with mock.patch.object(
                self.journal,
                "inspect_target_checkout",
                side_effect=[
                    self.checked_in(identifiers[0]),
                    self.checked_in(identifiers[1]),
                    self.checked_in(identifiers[0]),
                ],
            ), mock.patch.object(
                self.journal,
                "retrieve_exact_associated_drawing",
                side_effect=exported,
            ), mock.patch.object(
                self.journal, "import_one"
            ) as import_one:
                reports = self.journal.execute(
                    object(),
                    object(),
                    object(),
                    rows,
                    os.path.join(folder, "input.csv"),
                    "stamp",
                    "APPLY_APPROVED",
                    FakeLog(),
                    os.path.join(folder, "evidence"),
                )

        self.assertEqual("FAILED_IMPORT_UNVERIFIED", reports[0]["RESULT"])
        self.assertEqual(
            "BATCH_STOPPED_AFTER_UNVERIFIED_WRITE", reports[1]["RESULT"]
        )
        self.assertEqual("YES", reports[0]["WRITE_ATTEMPTED"])
        self.assertEqual("NO", reports[1]["WRITE_ATTEMPTED"])
        self.assertEqual(3, import_one.call_count)

    def test_local_file_race_blocks_apply_before_write(self):
        with tempfile.TemporaryDirectory() as folder:
            drawing = self.make_drawing(folder)
            row = self.make_row(drawing)
            identifier = row["DRAWING_IDENTIFIER"]
            with mock.patch.object(
                self.journal,
                "sha256",
                side_effect=["source-before", "source-after"],
            ), mock.patch.object(
                self.journal,
                "inspect_target_checkout",
                return_value=self.checked_in(identifier),
            ), mock.patch.object(
                self.journal,
                "retrieve_exact_associated_drawing",
                return_value=("baseline.prt", "tc-before"),
            ), mock.patch.object(
                self.journal, "import_one"
            ) as import_one:
                reports = self.journal.execute(
                    object(),
                    object(),
                    object(),
                    [row],
                    os.path.join(folder, "input.csv"),
                    "stamp",
                    "APPLY_APPROVED",
                    FakeLog(),
                    os.path.join(folder, "evidence"),
                )
        self.assertEqual(
            "ERROR_FILE_CHANGED_AFTER_PREFLIGHT", reports[0]["RESULT"]
        )
        self.assertEqual("NO", reports[0]["WRITE_ATTEMPTED"])
        self.assertEqual(1, import_one.call_count)

    def test_checkout_recheck_blocks_before_apply(self):
        with tempfile.TemporaryDirectory() as folder:
            drawing = self.make_drawing(folder)
            row = self.make_row(drawing)
            identifier = row["DRAWING_IDENTIFIER"]
            with mock.patch.object(
                self.journal,
                "inspect_target_checkout",
                side_effect=[
                    self.checked_in(identifier),
                    self.checked_out("current.user", identifier),
                ],
            ), mock.patch.object(
                self.journal,
                "retrieve_exact_associated_drawing",
                return_value=("baseline.prt", "tc-before"),
            ), mock.patch.object(
                self.journal, "import_one"
            ) as import_one:
                reports = self.journal.execute(
                    object(),
                    object(),
                    object(),
                    [row],
                    os.path.join(folder, "input.csv"),
                    "stamp",
                    "APPLY_APPROVED",
                    FakeLog(),
                    os.path.join(folder, "evidence"),
                )
        self.assertEqual("BLOCKED_CHECKED_OUT", reports[0]["RESULT"])
        self.assertEqual("current.user", reports[0]["CHECKOUT_RECHECK_OWNER"])
        self.assertEqual("NO", reports[0]["WRITE_ATTEMPTED"])
        self.assertEqual(1, import_one.call_count)

    def test_exact_retrieval_failure_never_reaches_clone(self):
        with tempfile.TemporaryDirectory() as folder:
            drawing = self.make_drawing(folder)
            row = self.make_row(drawing)
            identifier = row["DRAWING_IDENTIFIER"]
            with mock.patch.object(
                self.journal,
                "inspect_target_checkout",
                return_value=self.checked_in(identifier),
            ), mock.patch.object(
                self.journal,
                "retrieve_exact_associated_drawing",
                side_effect=RuntimeError("PDI code 17"),
            ), mock.patch.object(
                self.journal, "import_one"
            ) as import_one:
                reports = self.journal.execute(
                    object(),
                    object(),
                    object(),
                    [row],
                    os.path.join(folder, "input.csv"),
                    "stamp",
                    "APPLY_APPROVED",
                    FakeLog(),
                    os.path.join(folder, "evidence"),
                )
        self.assertEqual(
            "FAILED_TARGET_BASELINE_RETRIEVAL", reports[0]["RESULT"]
        )
        self.assertIn("PDI code 17", reports[0]["MESSAGE"])
        import_one.assert_not_called()

    def test_blocked_row_is_skipped_while_clear_row_imports(self):
        with tempfile.TemporaryDirectory() as folder:
            blocked = self.make_drawing(folder, "BLOCKED")
            clear = self.make_drawing(folder, "CLEAR")
            rows = [
                self.make_row(blocked, "BLOCKED"),
                self.make_row(clear, "CLEAR"),
            ]
            identifiers = [row["DRAWING_IDENTIFIER"] for row in rows]

            def exported(_session, _fm, proposal, _root, phase, *_fields):
                digest = (
                    self.journal.sha256(proposal["drawing"])
                    if phase == "POSTIMPORT"
                    else "tc-before"
                )
                return (phase + ".prt", digest)

            with mock.patch.object(
                self.journal,
                "inspect_target_checkout",
                side_effect=[
                    self.checked_out("other.user", identifiers[0]),
                    self.checked_in(identifiers[1]),
                    self.checked_in(identifiers[1]),
                ],
            ), mock.patch.object(
                self.journal,
                "retrieve_exact_associated_drawing",
                side_effect=exported,
            ), mock.patch.object(
                self.journal, "import_one"
            ) as import_one:
                reports = self.journal.execute(
                    object(),
                    object(),
                    object(),
                    rows,
                    os.path.join(folder, "input.csv"),
                    "stamp",
                    "APPLY_APPROVED",
                    FakeLog(),
                    os.path.join(folder, "evidence"),
                )
        self.assertEqual("BLOCKED_CHECKED_OUT", reports[0]["RESULT"])
        self.assertEqual("IMPORT_VERIFIED", reports[1]["RESULT"])
        self.assertEqual(2, import_one.call_count)

    def test_supplied_identifier_mismatch_is_rejected_locally(self):
        with tempfile.TemporaryDirectory() as folder:
            drawing = self.make_drawing(folder)
            row = self.make_row(drawing)
            row["DRAWING_IDENTIFIER"] = "@DB/WRONG/A/specification/WRONG-A-dwg1"
            reports, proposals = self.journal.local_preflight(
                [row],
                os.path.join(folder, "input.csv"),
                "stamp",
                "APPLY_APPROVED",
            )
        self.assertEqual("ERROR_IDENTITY_MISMATCH", reports[0]["RESULT"])
        self.assertEqual([], proposals)

    def test_blocked_result_fails_the_run(self):
        report = {"RESULT": "BLOCKED_CHECKED_OUT", "APPROVED": "YES"}
        self.assertTrue(
            self.journal.has_failure([report], "APPLY_APPROVED")
        )

    def make_retrieval_context(self, folder, names):
        identifier = self.journal.drawing_id("TEST100", "A", 1)
        status = FakeLoadStatus()
        part = types.SimpleNamespace(
            JournalIdentifier=identifier,
            Close=mock.Mock(),
        )
        parts = mock.MagicMock()
        parts.__iter__.return_value = iter([])
        parts.FindObject.side_effect = RuntimeError("not loaded")
        parts.OpenBase.return_value = (part, status)
        session = types.SimpleNamespace(Parts=parts)
        pdm_files = [FakePdmFile(name) for name in names]
        download_folder = os.path.join(folder, "managed_download")

        def download(_parts, files):
            self.assertEqual(pdm_files, files)
            os.makedirs(download_folder, exist_ok=True)
            for value in files:
                with open(os.path.join(download_folder, value.name), "wb") as handle:
                    handle.write(("payload-" + value.name).encode("ascii"))
            os.chdir(download_folder)
            return None

        file_management = types.SimpleNamespace(
            GetAssociatedFiles=mock.Mock(return_value=(pdm_files,)),
            DownloadAssociatedFiles=mock.Mock(side_effect=download),
        )
        report = self.journal.base_report({}, "stamp", "DRY_RUN")
        proposal = {
            "report": report,
            "identifier": identifier,
            "part_number": "TEST100",
            "revision": "A",
            "drawing_index": 1,
            "log": FakeLog(),
        }
        return session, file_management, proposal, pdm_files, status

    def test_associated_retrieval_selects_exact_native_and_restores_cwd(self):
        with tempfile.TemporaryDirectory() as folder:
            names = [
                "dwg_SHEET-1.qaf",
                "qafmetadata.qaf",
                "TEST100_A_dwg1.prt",
            ]
            session, fm, proposal, files, status = self.make_retrieval_context(
                folder, names
            )
            original_cwd = os.getcwd()
            evidence, digest = self.journal.retrieve_exact_associated_drawing(
                session,
                fm,
                proposal,
                folder,
                "BASELINE",
                "BASELINE_EXPORT_PDI_CODE",
                "BASELINE_EXPORT_FILE",
            )

            self.assertEqual(original_cwd, os.getcwd())
            self.assertTrue(os.path.isfile(evidence))
            self.assertEqual(self.journal.sha256(evidence), digest)
            self.assertTrue(evidence.endswith("TEST100_A_dwg1.prt"))
            self.assertEqual(
                "N/A_ASSOCIATED_FILES",
                proposal["report"]["BASELINE_EXPORT_PDI_CODE"],
            )
            self.assertIn("dwg_SHEET-1.qaf", proposal["report"]["BASELINE_ASSOCIATED_FILES"])
            self.assertTrue(all(value.released for value in files))
            self.assertTrue(status.disposed)

    def test_associated_retrieval_blocks_missing_exact_native(self):
        with tempfile.TemporaryDirectory() as folder:
            session, fm, proposal, files, _ = self.make_retrieval_context(
                folder, ["dwg_SHEET-1.qaf", "qafmetadata.qaf"]
            )
            original_cwd = os.getcwd()
            with self.assertRaisesRegex(RuntimeError, "exactly one"):
                self.journal.retrieve_exact_associated_drawing(
                    session,
                    fm,
                    proposal,
                    folder,
                    "BASELINE",
                    "BASELINE_EXPORT_PDI_CODE",
                    "BASELINE_EXPORT_FILE",
                )
            self.assertEqual(original_cwd, os.getcwd())
            fm.DownloadAssociatedFiles.assert_not_called()
            self.assertTrue(all(value.released for value in files))

    def test_associated_retrieval_blocks_duplicate_exact_native(self):
        with tempfile.TemporaryDirectory() as folder:
            session, fm, proposal, _, _ = self.make_retrieval_context(
                folder,
                ["TEST100_A_dwg1.prt", "TEST100_A_dwg1.prt"],
            )
            with self.assertRaisesRegex(RuntimeError, "found 2"):
                self.journal.retrieve_exact_associated_drawing(
                    session,
                    fm,
                    proposal,
                    folder,
                    "BASELINE",
                    "BASELINE_EXPORT_PDI_CODE",
                    "BASELINE_EXPORT_FILE",
                )
            fm.DownloadAssociatedFiles.assert_not_called()

    def test_trial_mode_only_proposes_confirmed_target(self):
        with tempfile.TemporaryDirectory() as folder:
            target = self.make_drawing(
                folder,
                self.journal.TRIAL_PART_NUMBER,
                self.journal.TRIAL_REVISION,
                self.journal.TRIAL_DRAWING_INDEX,
            )
            other = self.make_drawing(folder, "OTHER")
            rows = [
                self.make_row(
                    target,
                    self.journal.TRIAL_PART_NUMBER,
                    self.journal.TRIAL_REVISION,
                    self.journal.TRIAL_DRAWING_INDEX,
                ),
                self.make_row(other, "OTHER"),
            ]
            reports, proposals = self.journal.local_preflight(
                rows,
                os.path.join(folder, "input.csv"),
                "stamp",
                "TRIAL_APPLY",
            )
        self.assertEqual(1, len(proposals))
        self.assertEqual(self.journal.TRIAL_PART_NUMBER, proposals[0]["part_number"])
        self.assertEqual("TRIAL_SCOPE_SKIPPED", reports[1]["RESULT"])

    def test_trial_requires_one_target_and_fresh_session(self):
        self.journal.require_one_trial_row(
            [
                {
                    "PART_NUMBER": self.journal.TRIAL_PART_NUMBER,
                    "REVISION": self.journal.TRIAL_REVISION,
                    "DWG_INDEX": str(self.journal.TRIAL_DRAWING_INDEX),
                }
            ]
        )
        with self.assertRaisesRegex(RuntimeError, "exactly one CSV row"):
            self.journal.require_one_trial_row([])
        loaded = types.SimpleNamespace(JournalIdentifier="@DB/LOADED/A")
        with self.assertRaisesRegex(RuntimeError, "fresh NX managed session"):
            self.journal.require_fresh_trial_session(
                types.SimpleNamespace(Parts=[loaded])
            )

    def test_trial_exact_target_reaches_one_verified_write(self):
        with tempfile.TemporaryDirectory() as folder:
            drawing = self.make_drawing(
                folder,
                self.journal.TRIAL_PART_NUMBER,
                self.journal.TRIAL_REVISION,
                self.journal.TRIAL_DRAWING_INDEX,
            )
            row = self.make_row(
                drawing,
                self.journal.TRIAL_PART_NUMBER,
                self.journal.TRIAL_REVISION,
                self.journal.TRIAL_DRAWING_INDEX,
            )
            identifier = row["DRAWING_IDENTIFIER"]
            source_sha = self.journal.sha256(drawing)
            session = types.SimpleNamespace(Parts=[])

            def retrieved(_session, _fm, _proposal, _root, phase, *_fields):
                digest = source_sha if phase == "POSTIMPORT" else "tc-before"
                return (phase + ".prt", digest)

            with mock.patch.object(
                self.journal,
                "inspect_target_checkout",
                return_value=self.checked_in(identifier),
            ), mock.patch.object(
                self.journal,
                "retrieve_exact_associated_drawing",
                side_effect=retrieved,
            ), mock.patch.object(
                self.journal, "import_one"
            ) as import_one:
                reports = self.journal.execute(
                    session,
                    object(),
                    object(),
                    [row],
                    os.path.join(folder, "input.csv"),
                    "stamp",
                    "TRIAL_APPLY",
                    FakeLog(),
                    os.path.join(folder, "evidence"),
                )

        self.assertEqual("IMPORT_VERIFIED", reports[0]["RESULT"])
        self.assertEqual("YES", reports[0]["WRITE_ATTEMPTED"])
        self.assertEqual(2, import_one.call_count)
        self.assertTrue(import_one.call_args_list[0].args[3])
        self.assertFalse(import_one.call_args_list[1].args[3])

    def test_j16_uses_associated_files_not_legacy_export(self):
        source = J16_PATH.read_text(encoding="utf-8")
        self.assertIn("GetAssociatedFiles", source)
        self.assertIn("DownloadAssociatedFiles", source)
        self.assertNotIn("file_management.ExportFiles", source)
        self.assertNotIn("resolve_relation_and_export", source)
        self.assertEqual("DRY_RUN", self.journal.USER_MODE)
        self.assertFalse(self.journal.BATCH_APPLY_ENABLED)

    def test_j19_source_has_no_teamcenter_mutation_calls(self):
        source = J19_PATH.read_text(encoding="utf-8")
        forbidden = (
            ".Checkout(",
            ".Save(",
            "ImportFiles(",
            "SetDryrun(False",
            "J16.import_one(",
            "NXOpen.UF",
        )
        for token in forbidden:
            self.assertNotIn(token, source)
        self.assertIn("teamcenter_write_attempted", source)
        self.assertIn("DownloadAssociatedFiles", source)
        self.assertIn("ExportNamedReferences", source)


if __name__ == "__main__":
    unittest.main()
