import importlib.util
import os
import sys
import tempfile
import types
import unittest
import zipfile
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
    MANAGED_SHA = "a" * 64
    MODEL_SHA = "c" * 64

    @classmethod
    def setUpClass(cls):
        cls.journal = load_j16()
        cls.real_retrieve_exact_master_native = (
            cls.journal.retrieve_exact_master_native
        )

    def setUp(self):
        self.master_checkout_patcher = mock.patch.object(
            self.journal,
            "inspect_master_checkout",
            side_effect=lambda _session, proposal, _log: self.checked_in(
                proposal.get("model_identifier", "")
            ),
        )
        self.master_retrieval_patcher = mock.patch.object(
            self.journal,
            "retrieve_exact_master_native",
            return_value=("master.prt", self.MODEL_SHA),
        )
        self.master_checkout_patcher.start()
        self.master_retrieval_patcher.start()
        self.addCleanup(self.master_checkout_patcher.stop)
        self.addCleanup(self.master_retrieval_patcher.stop)

    def make_row(
        self,
        drawing,
        part_number="TEST100",
        revision="A",
        drawing_index=1,
        baseline="",
    ):
        row = {
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
        if os.path.isfile(drawing):
            row["APPROVED_LOCAL_SHA256"] = self.journal.sha256(drawing)
        row["APPROVED_TC_BASELINE_SHA256"] = self.MANAGED_SHA
        return row

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
            "drawing_sheet_count": 1,
        }

    def checked_out(self, owner="other.user", identifier=""):
        return {
            "state": "CHECKED_OUT",
            "owner": owner,
            "raw": "(True, '{0}')".format(owner),
            "opened_identifier": identifier,
            "drawing_sheet_count": 1,
        }

    def fresh_session(self):
        return types.SimpleNamespace(Parts=[])

    def make_clone_proposal(self, drawing, part_number="TEST100"):
        report = self.journal.base_report({}, "stamp", "APPLY_APPROVED")
        return {
            "report": report,
            "drawing": drawing,
            "part_number": part_number,
            "revision": "A",
            "drawing_index": 1,
            "identifier": self.journal.drawing_id(part_number, "A", 1),
            "model_identifier": self.journal.model_id(part_number, "A"),
            "model_native_name": self.journal.managed_model_native_name(
                part_number, "A"
            ),
        }

    def clone_api(self, clone):
        return {
            "clone": clone,
            "import_operation": "IMPORT",
            "treat_as_lost": "TREAT_AS_LOST",
            "autotranslate": "AUTOTRANSLATE",
            "use_existing": "USE_EXISTING",
            "overwrite": "OVERWRITE",
        }

    def test_manual_settings_default_to_approved_one_run(self):
        self.assertEqual("", self.journal.USER_IMPORT_CSV)
        self.assertEqual("APPLY_APPROVED", self.journal.USER_MODE)
        with mock.patch.dict(os.environ, {}, clear=True):
            self.assertEqual("APPLY_APPROVED", self.journal.configured_mode())

    def test_controlled_bulk_write_limit_configuration(self):
        with mock.patch.dict(os.environ, {}, clear=True):
            self.assertEqual(
                self.journal.DEFAULT_MAX_APPROVED_WRITES,
                self.journal.configured_max_approved_writes(),
            )
        with mock.patch.dict(
            os.environ, {"NX_J16_MAX_APPROVED_WRITES": "7"}, clear=True
        ):
            self.assertEqual(7, self.journal.configured_max_approved_writes())
        for invalid in ("0", "101", "not-a-number"):
            with mock.patch.dict(
                os.environ,
                {"NX_J16_MAX_APPROVED_WRITES": invalid},
                clear=True,
            ):
                with self.assertRaises(RuntimeError):
                    self.journal.configured_max_approved_writes()

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

    def test_approved_tc_baseline_mismatch_blocks_stale_target(self):
        report = {
            "RESULT": "LOCAL_PREFLIGHT_OK",
            "MODE": "APPLY_ONE_APPROVED",
            "APPROVED_TC_BASELINE_SHA256": "a" * 64,
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
            return_value=("baseline.prt", "b" * 64),
        ):
            self.journal.run_managed_preflight(
                object(), object(), [proposal], "evidence", FakeLog()
            )
        self.assertEqual("BLOCKED_STALE_TARGET", report["RESULT"])

    def test_one_run_captures_blank_local_approval_hash(self):
        with tempfile.TemporaryDirectory() as folder:
            drawing = self.make_drawing(folder, "GENERIC100")
            source_sha = self.journal.sha256(drawing)
            row = self.make_row(drawing, "GENERIC100")
            row.pop("APPROVED_LOCAL_SHA256")
            reports, proposals = self.journal.local_preflight(
                [row],
                os.path.join(folder, "input.csv"),
                "stamp",
                "APPLY_ONE_APPROVED",
            )
        self.assertEqual("LOCAL_PREFLIGHT_OK", reports[0]["RESULT"])
        self.assertEqual(source_sha, reports[0]["APPROVED_LOCAL_SHA256"])
        self.assertEqual(1, len(proposals))

    def test_apply_one_blocks_local_file_changed_after_dry_run(self):
        with tempfile.TemporaryDirectory() as folder:
            drawing = self.make_drawing(folder, "GENERIC100")
            row = self.make_row(drawing, "GENERIC100")
            row["APPROVED_LOCAL_SHA256"] = "a" * 64
            row["APPROVED_TC_BASELINE_SHA256"] = "b" * 64
            reports, proposals = self.journal.local_preflight(
                [row],
                os.path.join(folder, "input.csv"),
                "stamp",
                "APPLY_ONE_APPROVED",
            )
        self.assertEqual(
            "BLOCKED_LOCAL_CHANGED_AFTER_APPROVAL", reports[0]["RESULT"]
        )
        self.assertEqual([], proposals)

    def test_dry_run_populates_both_approval_hashes(self):
        with tempfile.TemporaryDirectory() as folder:
            drawing = self.make_drawing(folder, "GENERIC100")
            source_sha = self.journal.sha256(drawing)
            row = self.make_row(drawing, "GENERIC100")
            reports, proposals = self.journal.local_preflight(
                [row],
                os.path.join(folder, "input.csv"),
                "stamp",
                "DRY_RUN",
            )
            managed_sha = "b" * 64
            with mock.patch.object(
                self.journal,
                "inspect_target_checkout",
                return_value=self.checked_in(row["DRAWING_IDENTIFIER"]),
            ), mock.patch.object(
                self.journal,
                "retrieve_exact_associated_drawing",
                return_value=("baseline.prt", managed_sha),
            ):
                self.journal.run_managed_preflight(
                    object(), object(), proposals, folder, FakeLog()
                )

        self.assertEqual(
            source_sha,
            reports[0]["APPROVED_LOCAL_SHA256"],
        )
        self.assertEqual(
            managed_sha, reports[0]["APPROVED_TC_BASELINE_SHA256"]
        )
        self.assertEqual("NO", reports[0]["APPROVED"])
        self.assertEqual("", reports[0]["ENGINEER"])
        self.assertEqual("MANAGED_PREFLIGHT_OK", reports[0]["RESULT"])

    def test_one_run_captures_blank_tc_approval_hash(self):
        with tempfile.TemporaryDirectory() as folder:
            drawing = self.make_drawing(folder, "GENERIC100")
            row = self.make_row(drawing, "GENERIC100")
            row["APPROVED_LOCAL_SHA256"] = self.journal.sha256(drawing)
            row.pop("APPROVED_TC_BASELINE_SHA256")
            reports, proposals = self.journal.local_preflight(
                [row],
                os.path.join(folder, "input.csv"),
                "stamp",
                "APPLY_ONE_APPROVED",
            )
            with mock.patch.object(
                self.journal,
                "inspect_target_checkout",
                return_value=self.checked_in(row["DRAWING_IDENTIFIER"]),
            ), mock.patch.object(
                self.journal,
                "retrieve_exact_associated_drawing",
                return_value=("baseline.prt", "b" * 64),
            ):
                self.journal.run_managed_preflight(
                    object(), object(), proposals, folder, FakeLog()
                )

        self.assertEqual("MANAGED_PREFLIGHT_OK", reports[0]["RESULT"])
        self.assertEqual("b" * 64, reports[0]["APPROVED_TC_BASELINE_SHA256"])

    def test_verified_apply_requires_matching_post_payload_hash(self):
        with tempfile.TemporaryDirectory() as folder:
            drawing = self.make_drawing(folder)
            source_sha = self.journal.sha256(drawing)
            row = self.make_row(drawing)
            identifier = row["DRAWING_IDENTIFIER"]

            def exported(_session, _fm, _proposal, _root, phase, *_fields):
                digest = (
                    source_sha if phase == "POSTIMPORT" else self.MANAGED_SHA
                )
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
                    self.fresh_session(),
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
        self.assertEqual("CHECKED_IN", reports[0]["POST_IMPORT_CHECKOUT_STATE"])
        self.assertEqual(identifier, reports[0]["POST_IMPORT_OPENED_IDENTIFIER"])
        self.assertEqual(1, reports[0]["POST_IMPORT_DRAWING_SHEET_COUNT"])
        self.assertEqual("YES", reports[0]["PRESERVE_3D_UNCHANGED"])
        self.assertEqual(self.MODEL_SHA, reports[0]["PRESERVE_3D_POST_SHA256"])
        self.assertEqual(
            "CHECKED_IN", reports[0]["PRESERVE_3D_POST_CHECKOUT_STATE"]
        )
        self.assertEqual("YES", reports[0]["WRITE_ATTEMPTED"])
        self.assertEqual(2, import_one.call_count)

    def test_transformed_managed_payload_continues_controlled_bulk(self):
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
                    return ("post.prt", "post-" + proposal["part_number"])
                return ("before.prt", self.MANAGED_SHA)

            with mock.patch.object(
                self.journal,
                "inspect_target_checkout",
                side_effect=[
                    self.checked_in(identifiers[0]),
                    self.checked_in(identifiers[1]),
                    self.checked_in(identifiers[0]),
                    self.checked_in(identifiers[0]),
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
                    self.fresh_session(),
                    object(),
                    object(),
                    rows,
                    os.path.join(folder, "input.csv"),
                    "stamp",
                    "APPLY_APPROVED",
                    FakeLog(),
                    os.path.join(folder, "evidence"),
                )

        self.assertEqual(
            "IMPORT_VERIFIED_MANAGED_TRANSFORM",
            reports[0]["RESULT"],
        )
        self.assertEqual(
            "IMPORT_VERIFIED_MANAGED_TRANSFORM",
            reports[1]["RESULT"],
        )
        self.assertEqual(
            "VERIFIED_MANAGED_TRANSFORM",
            reports[0]["POST_IMPORT_VERIFICATION"],
        )
        self.assertEqual(
            "CHECKED_IN", reports[0]["POST_IMPORT_CHECKOUT_STATE"]
        )
        self.assertFalse(self.journal.has_review_required(reports))
        self.assertFalse(self.journal.has_failure(reports, "APPLY_APPROVED"))
        self.assertEqual("YES", reports[0]["WRITE_ATTEMPTED"])
        self.assertEqual("YES", reports[1]["WRITE_ATTEMPTED"])
        self.assertEqual(4, import_one.call_count)

    def test_local_file_race_quarantines_before_write(self):
        with tempfile.TemporaryDirectory() as folder:
            drawing = self.make_drawing(folder)
            row = self.make_row(drawing)
            source_before = "b" * 64
            source_after = "c" * 64
            row["APPROVED_LOCAL_SHA256"] = source_before
            identifier = row["DRAWING_IDENTIFIER"]
            with mock.patch.object(
                self.journal,
                "sha256",
                side_effect=[source_before, source_after],
            ), mock.patch.object(
                self.journal,
                "inspect_target_checkout",
                return_value=self.checked_in(identifier),
            ), mock.patch.object(
                self.journal,
                "retrieve_exact_associated_drawing",
                return_value=("baseline.prt", self.MANAGED_SHA),
            ), mock.patch.object(
                self.journal, "import_one"
            ) as import_one:
                reports = self.journal.execute(
                    self.fresh_session(),
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
            "QUARANTINED_PREFLIGHT", reports[0]["RESULT"]
        )
        self.assertIn("ERROR_FILE_CHANGED_AFTER_PREFLIGHT", reports[0]["QUARANTINE_REASON"])
        self.assertEqual("NO", reports[0]["WRITE_ATTEMPTED"])
        self.assertEqual(1, import_one.call_count)

    def test_checkout_recheck_quarantines_before_apply(self):
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
                return_value=("baseline.prt", self.MANAGED_SHA),
            ), mock.patch.object(
                self.journal, "import_one"
            ) as import_one:
                reports = self.journal.execute(
                    self.fresh_session(),
                    object(),
                    object(),
                    [row],
                    os.path.join(folder, "input.csv"),
                    "stamp",
                    "APPLY_APPROVED",
                    FakeLog(),
                    os.path.join(folder, "evidence"),
                )
        self.assertEqual("QUARANTINED_CHECKED_OUT", reports[0]["RESULT"])
        self.assertEqual("current.user", reports[0]["CHECKOUT_RECHECK_OWNER"])
        self.assertEqual("NO", reports[0]["WRITE_ATTEMPTED"])
        self.assertEqual(1, import_one.call_count)

    def test_exact_retrieval_failure_is_quarantined_before_clone(self):
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
                    self.fresh_session(),
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
            "QUARANTINED_PREFLIGHT", reports[0]["RESULT"]
        )
        self.assertIn("PDI code 17", reports[0]["MESSAGE"])
        import_one.assert_not_called()

    def test_approved_preflight_blocker_is_quarantined_while_clear_row_imports(self):
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
                    else self.MANAGED_SHA
                )
                return (phase + ".prt", digest)

            with mock.patch.object(
                self.journal,
                "inspect_target_checkout",
                side_effect=[
                    self.checked_out("other.user", identifiers[0]),
                    self.checked_in(identifiers[1]),
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
                    self.fresh_session(),
                    object(),
                    object(),
                    rows,
                    os.path.join(folder, "input.csv"),
                    "stamp",
                    "APPLY_APPROVED",
                    FakeLog(),
                    os.path.join(folder, "evidence"),
                )
        self.assertEqual("QUARANTINED_CHECKED_OUT", reports[0]["RESULT"])
        self.assertEqual("IMPORT_VERIFIED", reports[1]["RESULT"])
        self.assertEqual("YES", reports[1]["WRITE_ATTEMPTED"])
        self.assertEqual(2, import_one.call_count)
        self.assertEqual(
            "COMPLETED_WITH_QUARANTINE",
            self.journal.final_run_status(reports, "APPLY_APPROVED"),
        )

    def test_foreign_payload_row_is_quarantined_while_valid_row_imports(self):
        with tempfile.TemporaryDirectory() as folder:
            bad = self.make_drawing(folder, "264MN032670A01")
            good = self.make_drawing(folder, "264MN026142A01")
            rows = [
                self.make_row(bad, "264MN032670A01"),
                self.make_row(good, "264MN026142A01"),
            ]

            def retrieved(_session, _fm, proposal, _root, phase, *_fields):
                if proposal["part_number"] == "264MN032670A01":
                    report = proposal["report"]
                    report["EXPECTED_MANAGED_NATIVE"] = "264MN032670A01_A_dwg1.prt"
                    report["OBSERVED_NATIVE_FILES"] = "264MN026184A01_A_dwg1.prt"
                    report["ANOMALY_EVIDENCE_FILE"] = "foreign.prt"
                    report["ANOMALY_SHA256"] = "f" * 64
                    raise RuntimeError("foreign native payload")
                digest = (
                    self.journal.sha256(proposal["drawing"])
                    if phase == "POSTIMPORT"
                    else self.MANAGED_SHA
                )
                return (phase + ".prt", digest)

            with mock.patch.object(
                self.journal,
                "inspect_target_checkout",
                side_effect=lambda _session, identifier, _log: self.checked_in(identifier),
            ), mock.patch.object(
                self.journal,
                "retrieve_exact_associated_drawing",
                side_effect=retrieved,
            ), mock.patch.object(
                self.journal, "import_one"
            ) as import_one:
                reports = self.journal.execute(
                    self.fresh_session(),
                    object(),
                    object(),
                    rows,
                    os.path.join(folder, "input.csv"),
                    "stamp",
                    "APPLY_APPROVED",
                    FakeLog(),
                    os.path.join(folder, "evidence"),
                )

        self.assertEqual(
            "QUARANTINED_TARGET_PAYLOAD_IDENTITY", reports[0]["RESULT"]
        )
        self.assertEqual("NO", reports[0]["WRITE_ATTEMPTED"])
        self.assertEqual("IMPORT_VERIFIED", reports[1]["RESULT"])
        self.assertEqual("YES", reports[1]["WRITE_ATTEMPTED"])
        self.assertEqual(2, import_one.call_count)

    def test_process_state_failure_aborts_managed_preflight(self):
        report = {
            "RESULT": "LOCAL_PREFLIGHT_OK",
            "MODE": "APPLY_APPROVED",
        }
        proposal = {
            "report": report,
            "identifier": "exact",
            "part_number": "TEST100",
            "revision": "A",
            "drawing_index": 1,
            "preflight_sha": "b" * 64,
        }
        with mock.patch.object(
            self.journal,
            "inspect_target_checkout",
            return_value=self.checked_in("exact"),
        ), mock.patch.object(
            self.journal,
            "retrieve_exact_associated_drawing",
            side_effect=self.journal.ProcessStateError("cwd restore failed"),
        ):
            with self.assertRaisesRegex(
                self.journal.ProcessStateError, "cwd restore failed"
            ):
                self.journal.run_managed_preflight(
                    object(), object(), [proposal], "evidence", FakeLog()
                )

    def test_execute_reports_process_wide_failure_without_losing_rows(self):
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
                side_effect=self.journal.ProcessStateError("cwd restore failed"),
            ), mock.patch.object(
                self.journal, "import_one"
            ) as import_one:
                reports = self.journal.execute(
                    self.fresh_session(),
                    object(),
                    object(),
                    [row],
                    os.path.join(folder, "input.csv"),
                    "stamp",
                    "APPLY_APPROVED",
                    FakeLog(),
                    os.path.join(folder, "evidence"),
                )

        self.assertEqual("FAILED_PROCESS_STATE", reports[0]["RESULT"])
        self.assertEqual("ABORTED", reports[0]["DISPOSITION"])
        self.assertEqual("FAILED", self.journal.final_run_status(reports, "APPLY_APPROVED"))
        import_one.assert_not_called()

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

    def test_unchanged_post_payload_is_failed_unverified(self):
        report = self.journal.base_report({}, "stamp", "APPLY_ONE_APPROVED")
        proposal = {
            "report": report,
            "identifier": "exact",
            "preflight_sha": "source",
        }
        review = self.journal.classify_post_import(
            proposal,
            "before",
            "before",
            self.checked_in("exact"),
            FakeLog(),
        )
        self.assertFalse(review)
        self.assertEqual("FAILED_IMPORT_UNVERIFIED", report["RESULT"])
        self.assertEqual(
            "FAILED_UNCHANGED_FROM_PREWRITE",
            report["POST_IMPORT_VERIFICATION"],
        )

    def test_accepted_trial_hashes_classify_as_managed_transform_verified(self):
        report = self.journal.base_report(
            {}, "20260731_210943", "APPLY_ONE_APPROVED"
        )
        proposal = {
            "report": report,
            "identifier": self.journal.drawing_id("264MN021218A01", "A", 1),
            "preflight_sha": (
                "43145a5d62429b0f020175da5ff8d8d5f1fac696161e6b23f9d09badb7045ebb"
            ),
        }
        review = self.journal.classify_post_import(
            proposal,
            "c80fcfda652ede4e25ef6995de64dbf7a587902ebd619cb8b40abf50a4739cf7",
            "ceca56a880a80e1b28d6f1eee6f5f25f56264f82653e8456083bf4b6218c654a",
            self.checked_in(proposal["identifier"]),
            FakeLog(),
        )
        self.assertFalse(review)
        self.assertEqual(
            "IMPORT_VERIFIED_MANAGED_TRANSFORM",
            report["RESULT"],
        )
        self.assertEqual(
            "VERIFIED_MANAGED_TRANSFORM",
            report["POST_IMPORT_VERIFICATION"],
        )
        self.assertEqual("CHECKED_IN", report["POST_IMPORT_CHECKOUT_STATE"])

    def test_changed_post_payload_still_checked_out_requires_manual_checkin(self):
        report = self.journal.base_report({}, "stamp", "APPLY_ONE_APPROVED")
        proposal = {
            "report": report,
            "identifier": "exact",
            "preflight_sha": "source",
        }
        review = self.journal.classify_post_import(
            proposal,
            "before",
            "managed-transform",
            self.checked_out("current.user", "exact"),
            FakeLog(),
        )
        self.assertTrue(review)
        self.assertEqual("MANUAL_CHECKIN_REQUIRED", report["RESULT"])
        self.assertEqual("current.user", report["POST_IMPORT_CHECKOUT_OWNER"])
        self.assertIn("J16 will not call check-in", report["MESSAGE"])
        self.assertTrue(self.journal.has_review_required([report]))

    def test_changed_post_payload_with_unknown_checkout_fails_closed(self):
        report = self.journal.base_report({}, "stamp", "APPLY_ONE_APPROVED")
        proposal = {
            "report": report,
            "identifier": "exact",
            "preflight_sha": "source",
        }
        checkout = {
            "state": "UNKNOWN",
            "owner": "",
            "raw": "unrecognized",
            "opened_identifier": "exact",
        }
        review = self.journal.classify_post_import(
            proposal,
            "before",
            "managed-transform",
            checkout,
            FakeLog(),
        )
        self.assertFalse(review)
        self.assertEqual("FAILED_IMPORT_UNVERIFIED", report["RESULT"])
        self.assertEqual(
            "FAILED_POST_CHECKOUT_UNKNOWN",
            report["POST_IMPORT_VERIFICATION"],
        )

    def test_post_import_retrieval_failure_remains_unverified(self):
        with tempfile.TemporaryDirectory() as folder:
            drawing = self.make_drawing(folder)
            row = self.make_row(drawing)
            identifier = row["DRAWING_IDENTIFIER"]

            def retrieved(_session, _fm, _proposal, _root, phase, *_fields):
                if phase == "POSTIMPORT":
                    raise RuntimeError("post download failed")
                return (phase + ".prt", self.MANAGED_SHA)

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
                    self.fresh_session(),
                    object(),
                    object(),
                    [row],
                    os.path.join(folder, "input.csv"),
                    "stamp",
                    "APPLY_APPROVED",
                    FakeLog(),
                    os.path.join(folder, "evidence"),
                )

        self.assertEqual("FAILED_IMPORT_UNVERIFIED", reports[0]["RESULT"])
        self.assertIn("post download failed", reports[0]["MESSAGE"])
        self.assertEqual("YES", reports[0]["WRITE_ATTEMPTED"])
        self.assertEqual(2, import_one.call_count)

    def test_unverified_bulk_write_stops_every_later_write(self):
        with tempfile.TemporaryDirectory() as folder:
            first = self.make_drawing(folder, "FIRST")
            second = self.make_drawing(folder, "SECOND")
            rows = [
                self.make_row(first, "FIRST"),
                self.make_row(second, "SECOND"),
            ]
            identifiers = [row["DRAWING_IDENTIFIER"] for row in rows]

            def retrieved(_session, _fm, _proposal, _root, phase, *_fields):
                if phase == "POSTIMPORT":
                    raise RuntimeError("post-import evidence unavailable")
                return (phase + ".prt", self.MANAGED_SHA)

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
                side_effect=retrieved,
            ), mock.patch.object(
                self.journal, "import_one"
            ) as import_one:
                reports = self.journal.execute(
                    self.fresh_session(),
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
            fm.DownloadAssociatedFiles.assert_called_once()

    def test_foreign_associated_native_is_captured_as_read_only_evidence(self):
        with tempfile.TemporaryDirectory() as folder:
            foreign = "264MN026184A01_A_dwg1.prt"
            session, fm, proposal, files, _ = self.make_retrieval_context(
                folder,
                ["dwg_SHEET-1.qaf", "qafmetadata.qaf", foreign],
            )
            original_cwd = os.getcwd()
            with self.assertRaisesRegex(RuntimeError, "found 0"):
                self.journal.retrieve_exact_associated_drawing(
                    session,
                    fm,
                    proposal,
                    folder,
                    "BASELINE",
                    "BASELINE_EXPORT_PDI_CODE",
                    "BASELINE_EXPORT_FILE",
                )

            report = proposal["report"]
            self.assertEqual(original_cwd, os.getcwd())
            self.assertEqual("TEST100_A_dwg1.prt", report["EXPECTED_MANAGED_NATIVE"])
            self.assertEqual(foreign, report["OBSERVED_NATIVE_FILES"])
            self.assertTrue(os.path.isfile(report["ANOMALY_EVIDENCE_FILE"]))
            self.assertTrue(self.journal.valid_sha256(report["ANOMALY_SHA256"]))
            self.assertTrue(all(value.released for value in files))

            report["RESULT"] = "FAILED_TARGET_BASELINE_RETRIEVAL"
            report["MESSAGE"] = "foreign native payload"
            report["APPROVED"] = "YES"
            self.journal.quarantine_report(report)
            self.assertEqual(
                "QUARANTINED_TARGET_PAYLOAD_IDENTITY", report["RESULT"]
            )

    def test_verified_receipt_skips_exact_rerun(self):
        with tempfile.TemporaryDirectory() as folder:
            drawing = self.make_drawing(folder)
            row = self.make_row(drawing)
            row.pop("APPROVED_LOCAL_SHA256")
            row["APPROVED_TC_BASELINE_SHA256"] = "d" * 64
            source_sha = self.journal.sha256(drawing)
            identifier = row["DRAWING_IDENTIFIER"]
            receipts = {
                self.journal.receipt_key(identifier, source_sha): {
                    "TARGET_IDENTIFIER": identifier,
                    "SOURCE_SHA256": source_sha,
                    "VERIFIED_POST_TC_SHA256": self.MANAGED_SHA,
                }
            }
            with mock.patch.object(
                self.journal,
                "inspect_target_checkout",
                return_value=self.checked_in(identifier),
            ), mock.patch.object(
                self.journal,
                "retrieve_exact_associated_drawing",
                return_value=("baseline.prt", self.MANAGED_SHA),
            ), mock.patch.object(
                self.journal, "import_one"
            ) as import_one:
                reports = self.journal.execute(
                    self.fresh_session(),
                    object(),
                    object(),
                    [row],
                    os.path.join(folder, "input.csv"),
                    "stamp",
                    "APPLY_APPROVED",
                    FakeLog(),
                    os.path.join(folder, "evidence"),
                    os.path.join(folder, self.journal.DEFAULT_RECEIPTS),
                    receipts,
                )

        self.assertEqual("SKIPPED_ALREADY_IMPORTED", reports[0]["RESULT"])
        self.assertEqual("MATCHED_VERIFIED_POST", reports[0]["RECEIPT_STATUS"])
        import_one.assert_not_called()

    def test_stale_receipt_falls_back_to_full_preflight(self):
        with tempfile.TemporaryDirectory() as folder:
            drawing = self.make_drawing(folder)
            row = self.make_row(drawing)
            source_sha = self.journal.sha256(drawing)
            identifier = row["DRAWING_IDENTIFIER"]
            reports, proposals = self.journal.local_preflight(
                [row],
                os.path.join(folder, "input.csv"),
                "stamp",
                "APPLY_APPROVED",
            )
            receipts = {
                self.journal.receipt_key(identifier, source_sha): {
                    "VERIFIED_POST_TC_SHA256": "c" * 64,
                }
            }
            with mock.patch.object(
                self.journal,
                "inspect_target_checkout",
                return_value=self.checked_in(identifier),
            ), mock.patch.object(
                self.journal,
                "retrieve_exact_associated_drawing",
                return_value=("baseline.prt", self.MANAGED_SHA),
            ):
                self.journal.run_managed_preflight(
                    object(), object(), proposals, folder, FakeLog(), receipts
                )

        self.assertEqual("MANAGED_PREFLIGHT_OK", reports[0]["RESULT"])
        self.assertEqual("STALE_CURRENT_TC_DIFFERS", reports[0]["RECEIPT_STATUS"])

    def test_one_run_verified_import_records_atomic_receipt(self):
        with tempfile.TemporaryDirectory() as folder:
            drawing = self.make_drawing(folder)
            row = self.make_row(drawing)
            row.pop("APPROVED_LOCAL_SHA256")
            row.pop("APPROVED_TC_BASELINE_SHA256")
            identifier = row["DRAWING_IDENTIFIER"]
            source_sha = self.journal.sha256(drawing)
            receipt_path = os.path.join(folder, self.journal.DEFAULT_RECEIPTS)

            def retrieved(_session, _fm, _proposal, _root, phase, *_fields):
                digest = source_sha if phase == "POSTIMPORT" else self.MANAGED_SHA
                return (phase + ".prt", digest)

            with mock.patch.object(
                self.journal,
                "inspect_target_checkout",
                return_value=self.checked_in(identifier),
            ), mock.patch.object(
                self.journal,
                "retrieve_exact_associated_drawing",
                side_effect=retrieved,
            ), mock.patch.object(self.journal, "import_one"):
                reports = self.journal.execute(
                    self.fresh_session(),
                    object(),
                    object(),
                    [row],
                    os.path.join(folder, "input.csv"),
                    "stamp",
                    "APPLY_APPROVED",
                    FakeLog(),
                    os.path.join(folder, "evidence"),
                    receipt_path,
                    {},
                )

            receipts, warning = self.journal.load_receipts(receipt_path)
            self.assertEqual("", warning)
            receipt = receipts[self.journal.receipt_key(identifier, source_sha)]
            self.assertEqual(source_sha, receipt["VERIFIED_POST_TC_SHA256"])
            self.assertEqual("RECORDED", reports[0]["RECEIPT_STATUS"])

    def test_receipt_write_failure_does_not_invalidate_verified_import(self):
        with tempfile.TemporaryDirectory() as folder:
            drawing = self.make_drawing(folder)
            row = self.make_row(drawing)
            identifier = row["DRAWING_IDENTIFIER"]
            source_sha = self.journal.sha256(drawing)

            def retrieved(_session, _fm, _proposal, _root, phase, *_fields):
                digest = source_sha if phase == "POSTIMPORT" else self.MANAGED_SHA
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
            ), mock.patch.object(
                self.journal,
                "write_receipts_atomic",
                side_effect=OSError("disk full"),
            ):
                reports = self.journal.execute(
                    self.fresh_session(),
                    object(),
                    object(),
                    [row],
                    os.path.join(folder, "input.csv"),
                    "stamp",
                    "APPLY_APPROVED",
                    FakeLog(),
                    os.path.join(folder, "evidence"),
                    os.path.join(folder, self.journal.DEFAULT_RECEIPTS),
                    {},
                )

        self.assertEqual("IMPORT_VERIFIED", reports[0]["RESULT"])
        self.assertIn("WRITE_FAILED", reports[0]["RECEIPT_STATUS"])
        self.assertEqual(
            "SUCCESS", self.journal.final_run_status(reports, "APPLY_APPROVED")
        )

    def test_malformed_receipt_is_nonfatal_and_ignored(self):
        with tempfile.TemporaryDirectory() as folder:
            path = os.path.join(folder, self.journal.DEFAULT_RECEIPTS)
            with open(path, "w", encoding="utf-8") as handle:
                handle.write("bad,columns\n1,2\n")
            receipts, warning = self.journal.load_receipts(path)
        self.assertEqual({}, receipts)
        self.assertIn("ignored", warning)

    def test_only_quarantined_approved_rows_fail_the_run(self):
        report = {
            "APPROVED": "YES",
            "RESULT": "QUARANTINED_PREFLIGHT",
            "DISPOSITION": "QUARANTINED",
        }
        self.assertEqual(
            "FAILED",
            self.journal.final_run_status([report], "APPLY_APPROVED"),
        )

    def test_apply_one_only_proposes_the_approved_generic_target(self):
        with tempfile.TemporaryDirectory() as folder:
            target = self.make_drawing(folder, "GENERIC100")
            other = self.make_drawing(folder, "OTHER")
            rows = [
                self.make_row(target, "GENERIC100"),
                self.make_row(other, "OTHER"),
            ]
            rows[0]["APPROVED_LOCAL_SHA256"] = self.journal.sha256(target)
            rows[0]["APPROVED_TC_BASELINE_SHA256"] = "a" * 64
            rows[1]["APPROVED"] = "NO"
            reports, proposals = self.journal.local_preflight(
                rows,
                os.path.join(folder, "input.csv"),
                "stamp",
                "APPLY_ONE_APPROVED",
            )
        self.assertEqual(1, len(proposals))
        self.assertEqual("GENERIC100", proposals[0]["part_number"])
        self.assertEqual("NOT_APPROVED", reports[1]["RESULT"])

    def test_controlled_modes_validate_approvals_and_fresh_session(self):
        approved = {
            "PART_NUMBER": "GENERIC100",
            "REVISION": "A",
            "DWG_INDEX": "1",
            "APPROVED": "YES",
        }
        self.assertEqual(
            1,
            self.journal.require_approved_rows(
                [
                    approved,
                    {"APPROVED": "NO"},
                ],
                "APPLY_ONE_APPROVED",
                1,
            ),
        )
        with self.assertRaisesRegex(RuntimeError, "exactly one APPROVED=YES"):
            self.journal.require_approved_rows([], "APPLY_ONE_APPROVED", 1)
        with self.assertRaisesRegex(RuntimeError, "found 2"):
            self.journal.require_approved_rows(
                [approved, dict(approved)], "APPLY_ONE_APPROVED", 2
            )
        with self.assertRaisesRegex(RuntimeError, "invalid row"):
            self.journal.require_approved_rows(
                [approved, {"APPROVED": "MAYBE"}],
                "APPLY_APPROVED",
                25,
            )
        with self.assertRaisesRegex(RuntimeError, "at least one"):
            self.journal.require_approved_rows(
                [{"APPROVED": "NO"}], "APPLY_APPROVED", 25
            )
        self.assertEqual(
            2,
            self.journal.require_approved_rows(
                [approved, dict(approved)], "APPLY_APPROVED", 2
            ),
        )
        with self.assertRaisesRegex(RuntimeError, "exceeds"):
            self.journal.require_approved_rows(
                [approved, dict(approved)], "APPLY_APPROVED", 1
            )
        loaded = types.SimpleNamespace(JournalIdentifier="@DB/LOADED/A")
        with self.assertRaisesRegex(RuntimeError, "fresh NX managed session"):
            self.journal.require_fresh_apply_session(
                types.SimpleNamespace(Parts=[loaded])
            )

    def test_apply_one_generic_target_reaches_one_verified_write(self):
        with tempfile.TemporaryDirectory() as folder:
            drawing = self.make_drawing(folder, "GENERIC100")
            row = self.make_row(drawing, "GENERIC100")
            identifier = row["DRAWING_IDENTIFIER"]
            source_sha = self.journal.sha256(drawing)
            approved_tc_sha = "a" * 64
            row["APPROVED_LOCAL_SHA256"] = source_sha
            row["APPROVED_TC_BASELINE_SHA256"] = approved_tc_sha
            session = types.SimpleNamespace(Parts=[])

            def retrieved(_session, _fm, _proposal, _root, phase, *_fields):
                digest = source_sha if phase == "POSTIMPORT" else approved_tc_sha
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
                    "APPLY_ONE_APPROVED",
                    FakeLog(),
                    os.path.join(folder, "evidence"),
                )

        self.assertEqual("IMPORT_VERIFIED", reports[0]["RESULT"])
        self.assertEqual("YES", reports[0]["WRITE_ATTEMPTED"])
        self.assertEqual(2, import_one.call_count)
        self.assertTrue(import_one.call_args_list[0].args[3])
        self.assertFalse(import_one.call_args_list[1].args[3])

    def test_clone_discovery_fails_closed_before_perform(self):
        with tempfile.TemporaryDirectory() as folder:
            drawing = self.make_drawing(folder)
            proposal = self.make_clone_proposal(drawing)
            model = os.path.join(folder, "TEST100_A.prt")
            cases = (
                ([], "discovered no parts"),
                ([model], "exact local drawing"),
                ([drawing], "preserved 3D master"),
            )
            for discovered, message in cases:
                with self.subTest(discovered=discovered), mock.patch.object(
                    self.journal, "iterate_parts", return_value=discovered
                ), mock.patch.object(
                    self.journal, "perform_clone"
                ) as perform:
                    with self.assertRaisesRegex(RuntimeError, message):
                        self.journal.import_one(
                            self.clone_api(mock.MagicMock()),
                            proposal,
                            os.path.join(folder, "clone.log"),
                            True,
                            FakeLog(),
                        )
                    perform.assert_not_called()

    def test_clone_assigns_only_drawing_overwrite_and_master_use_existing(self):
        with tempfile.TemporaryDirectory() as folder:
            drawing = self.make_drawing(folder)
            proposal = self.make_clone_proposal(drawing)
            model = os.path.join(folder, "TEST100_A.prt")
            reference = os.path.join(folder, "REFERENCE_A.prt")
            clone = mock.MagicMock()
            with mock.patch.object(
                self.journal,
                "iterate_parts",
                return_value=[model, drawing, reference],
            ), mock.patch.object(
                self.journal, "naming_failures", return_value=None
            ), mock.patch.object(
                self.journal, "perform_clone"
            ) as perform:
                discovered = self.journal.import_one(
                    self.clone_api(clone),
                    proposal,
                    os.path.join(folder, "clone.log"),
                    True,
                    FakeLog(),
                )

        self.assertEqual([model, drawing, reference], discovered)
        clone.SetDefAction.assert_called_once_with("USE_EXISTING")
        self.assertIn(
            mock.call(drawing, "OVERWRITE", ""), clone.SetAction.call_args_list
        )
        self.assertIn(
            mock.call(model, "USE_EXISTING", ""), clone.SetAction.call_args_list
        )
        self.assertEqual("YES", proposal["report"]["DRAWING_DISCOVERED"])
        self.assertEqual("YES", proposal["report"]["PRESERVE_3D_DISCOVERED"])
        perform.assert_called_once()

    def test_checked_out_master_blocks_before_clone_preflight(self):
        report = self.journal.base_report({}, "stamp", "APPLY_APPROVED")
        report["RESULT"] = "LOCAL_PREFLIGHT_OK"
        report["MODE"] = "APPLY_APPROVED"
        proposal = {
            "report": report,
            "identifier": self.journal.drawing_id("TEST100", "A", 1),
            "part_number": "TEST100",
            "revision": "A",
            "drawing_index": 1,
            "preflight_sha": "source",
        }
        with mock.patch.object(
            self.journal,
            "inspect_target_checkout",
            return_value=self.checked_in(proposal["identifier"]),
        ), mock.patch.object(
            self.journal,
            "retrieve_exact_associated_drawing",
            return_value=("drawing.prt", self.MANAGED_SHA),
        ), mock.patch.object(
            self.journal,
            "inspect_master_checkout",
            return_value=self.checked_out("other.user", "@DB/TEST100/A"),
        ), mock.patch.object(
            self.journal, "retrieve_exact_master_native"
        ) as retrieve_master:
            self.journal.run_managed_preflight(
                object(), object(), [proposal], "evidence", FakeLog()
            )

        self.assertEqual("BLOCKED_MASTER_CHECKED_OUT", report["RESULT"])
        self.assertIn("other.user", report["MESSAGE"])
        retrieve_master.assert_not_called()

    def test_unknown_master_checkout_blocks_fail_closed(self):
        report = {}
        blocked = self.journal.block_for_master_checkout(
            report,
            {
                "state": "UNKNOWN",
                "owner": "",
                "raw": "unrecognized NX output",
            },
            "clone preflight",
        )
        self.assertTrue(blocked)
        self.assertEqual("BLOCKED_MASTER_CHECKOUT_UNKNOWN", report["RESULT"])
        self.assertIn("unrecognized NX output", report["MESSAGE"])

    def test_master_change_before_apply_quarantines_without_write(self):
        with tempfile.TemporaryDirectory() as folder:
            drawing = self.make_drawing(folder)
            row = self.make_row(drawing)
            identifier = row["DRAWING_IDENTIFIER"]

            def target_retrieval(_s, _f, _p, _r, phase, *_fields):
                return (phase + ".prt", self.MANAGED_SHA)

            with mock.patch.object(
                self.journal,
                "inspect_target_checkout",
                return_value=self.checked_in(identifier),
            ), mock.patch.object(
                self.journal,
                "retrieve_exact_associated_drawing",
                side_effect=target_retrieval,
            ), mock.patch.object(
                self.journal,
                "retrieve_exact_master_native",
                side_effect=[
                    ("master-baseline.prt", self.MODEL_SHA),
                    ("master-prewrite.prt", "d" * 64),
                ],
            ), mock.patch.object(
                self.journal, "import_one"
            ) as import_one:
                reports = self.journal.execute(
                    self.fresh_session(),
                    object(),
                    object(),
                    [row],
                    os.path.join(folder, "input.csv"),
                    "stamp",
                    "APPLY_APPROVED",
                    FakeLog(),
                    os.path.join(folder, "evidence"),
                )

        self.assertEqual("QUARANTINED_PREFLIGHT", reports[0]["RESULT"])
        self.assertEqual("NO", reports[0]["WRITE_ATTEMPTED"])
        self.assertEqual(1, import_one.call_count)

    def test_changed_3d_after_write_fails_and_stops_later_rows(self):
        with tempfile.TemporaryDirectory() as folder:
            first = self.make_drawing(folder, "FIRST")
            second = self.make_drawing(folder, "SECOND")
            rows = [self.make_row(first, "FIRST"), self.make_row(second, "SECOND")]

            def target_retrieval(_s, _f, proposal, _r, phase, *_fields):
                if phase == "POSTIMPORT":
                    return (phase + ".prt", "post-" + proposal["part_number"])
                return (phase + ".prt", self.MANAGED_SHA)

            master_hashes = [
                ("first-base.prt", self.MODEL_SHA),
                ("second-base.prt", self.MODEL_SHA),
                ("first-pre.prt", self.MODEL_SHA),
                ("first-post.prt", "d" * 64),
            ]
            with mock.patch.object(
                self.journal,
                "inspect_target_checkout",
                return_value=self.checked_in(),
            ), mock.patch.object(
                self.journal,
                "retrieve_exact_associated_drawing",
                side_effect=target_retrieval,
            ), mock.patch.object(
                self.journal,
                "retrieve_exact_master_native",
                side_effect=master_hashes,
            ), mock.patch.object(
                self.journal, "import_one"
            ) as import_one:
                reports = self.journal.execute(
                    self.fresh_session(),
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
        self.assertEqual("NO", reports[0]["PRESERVE_3D_UNCHANGED"])
        self.assertEqual(
            "BATCH_STOPPED_AFTER_UNVERIFIED_WRITE", reports[1]["RESULT"]
        )
        self.assertEqual(3, import_one.call_count)

    def test_zero_sheet_post_import_is_failed_unverified(self):
        proposal = {
            "report": self.journal.base_report({}, "stamp", "APPLY_APPROVED"),
            "identifier": self.journal.drawing_id("TEST100", "A", 1),
            "preflight_sha": "source",
        }
        checkout = self.checked_in(proposal["identifier"])
        checkout["drawing_sheet_count"] = 0
        review = self.journal.classify_post_import(
            proposal,
            "before",
            "after",
            checkout,
            FakeLog(),
        )
        self.assertFalse(review)
        self.assertEqual(
            "FAILED_IMPORT_UNVERIFIED", proposal["report"]["RESULT"]
        )
        self.assertEqual(
            "FAILED_DRAWING_SHEETS_UNPROVEN",
            proposal["report"]["POST_IMPORT_VERIFICATION"],
        )

    def test_exact_master_retrieval_restores_cwd_and_hashes_native(self):
        with tempfile.TemporaryDirectory() as folder:
            identifier = self.journal.model_id("TEST100", "A")
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
            pdm_files = [
                FakePdmFile("qafmetadata.qaf"),
                FakePdmFile("TEST100_A.prt"),
            ]
            download_folder = os.path.join(folder, "managed_download")

            def download(_parts, files):
                os.makedirs(download_folder, exist_ok=True)
                for value in files:
                    with open(
                        os.path.join(download_folder, value.name), "wb"
                    ) as handle:
                        handle.write(("payload-" + value.name).encode("ascii"))
                os.chdir(download_folder)

            fm = types.SimpleNamespace(
                GetAssociatedFiles=mock.Mock(return_value=(pdm_files,)),
                DownloadAssociatedFiles=mock.Mock(side_effect=download),
            )
            proposal = self.make_clone_proposal(
                os.path.join(folder, "drawing.prt")
            )
            proposal["log"] = FakeLog()
            original_cwd = os.getcwd()
            evidence, digest = type(self).real_retrieve_exact_master_native(
                session, fm, proposal, folder, "BASELINE"
            )
            evidence_sha = self.journal.sha256(evidence)

        self.assertEqual(original_cwd, os.getcwd())
        self.assertTrue(evidence.endswith("TEST100_A.prt"))
        self.assertEqual(evidence_sha, digest)
        self.assertTrue(all(value.released for value in pdm_files))
        self.assertTrue(status.disposed)

    def test_j16_uses_associated_files_not_legacy_export(self):
        source = J16_PATH.read_text(encoding="utf-8")
        self.assertIn("GetAssociatedFiles", source)
        self.assertIn("DownloadAssociatedFiles", source)
        self.assertNotIn("file_management.ExportFiles", source)
        self.assertNotIn("resolve_relation_and_export", source)
        self.assertNotIn(".Checkin(", source)
        self.assertNotIn(".CheckIn(", source)
        self.assertNotIn(".Checkout(", source)
        self.assertNotIn(".CheckOut(", source)
        self.assertNotIn(".Save(", source)
        self.assertNotIn(".SaveAs(", source)
        self.assertNotIn(".Delete(", source)
        self.assertNotIn(".Revise(", source)
        self.assertIn("POST_IMPORT_CHECKOUT_STATE", source)
        self.assertIn("IMPORT_VERIFIED_MANAGED_TRANSFORM", source)
        self.assertIn("APPLY_ONE_APPROVED", source)
        self.assertNotIn("TRIAL_APPLY", source)
        self.assertNotIn("TRIAL_PART_NUMBER", source)
        self.assertEqual("APPLY_APPROVED", self.journal.USER_MODE)
        self.assertEqual("", self.journal.USER_IMPORT_CSV)
        self.assertTrue(self.journal.BATCH_APPLY_ENABLED)
        self.assertIn("DEFAULT_MAX_APPROVED_WRITES", source)

    def test_evidence_zip_includes_reports_logs_and_managed_evidence(self):
        with tempfile.TemporaryDirectory() as folder:
            evidence = os.path.join(folder, "J16_EVIDENCE_stamp")
            os.makedirs(os.path.join(evidence, "TARGET"))
            managed = os.path.join(evidence, "TARGET", "managed.prt")
            report = os.path.join(folder, "J16_APPLY_ONE_APPROVED_stamp.csv")
            run_log = os.path.join(folder, "J16_RUN_APPLY_ONE_APPROVED_stamp.log")
            clone_log = os.path.join(folder, "J16_APPLY.clone")
            source = os.path.join(folder, "offline_source.prt")
            for path, payload in (
                (managed, b"managed"),
                (report, b"report"),
                (run_log, b"log"),
                (clone_log, b"clone"),
                (source, b"source-must-not-be-packaged"),
            ):
                with open(path, "wb") as handle:
                    handle.write(payload)
            zip_path = evidence + ".zip"
            self.journal.zip_artifacts(
                zip_path, evidence, [report, run_log, clone_log]
            )
            with zipfile.ZipFile(zip_path) as archive:
                names = set(archive.namelist())

        self.assertIn("J16_EVIDENCE_stamp/TARGET/managed.prt", names)
        self.assertIn(os.path.basename(report), names)
        self.assertIn(os.path.basename(run_log), names)
        self.assertIn(os.path.basename(clone_log), names)
        self.assertNotIn(os.path.basename(source), names)

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
