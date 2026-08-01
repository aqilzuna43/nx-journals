import importlib.util
import os
import sys
import tempfile
import types
import unittest
from pathlib import Path
from unittest import mock


ROOT = Path(__file__).resolve().parents[1]
J17_PATH = (
    ROOT / "from_git" / "journals" / "17_tc_master_drawing_import.py"
)


def load_j17():
    nxopen = types.ModuleType("NXOpen")
    nxuf = types.ModuleType("NXOpen.UF")
    nxopen.UF = nxuf
    nxopen.BasePart = types.SimpleNamespace(
        CloseWholeTree=types.SimpleNamespace(FalseValue="FalseValue"),
        CloseModified=types.SimpleNamespace(CloseModified="CloseModified"),
    )
    sys.modules["NXOpen"] = nxopen
    sys.modules["NXOpen.UF"] = nxuf
    spec = importlib.util.spec_from_file_location("journal17_create", J17_PATH)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class FakeLog:
    def __init__(self):
        self.lines = []

    def write(self, message=""):
        self.lines.append(str(message))


class FakePdmFile:
    def __init__(self, name):
        self.name = name
        self.released = False

    def GetFileName(self):
        return self.name

    def FreeResource(self):
        self.released = True


class FakeLoadStatus:
    def __init__(self):
        self.disposed = False

    def Dispose(self):
        self.disposed = True


class MasterDrawingCreateTests(unittest.TestCase):
    MODEL_SHA = "a" * 64

    @classmethod
    def setUpClass(cls):
        cls.journal = load_j17()

    def make_source(self, folder, name="offline_master.prt"):
        path = os.path.join(folder, name)
        with open(path, "wb") as handle:
            handle.write(b"local-drawing-master-content")
        return path

    def make_row(self, source, part_number="MODEL100", drawing_index=3):
        return {
            "_CSV_ROW": 2,
            "PART_NUMBER": part_number,
            "REVISION": "A",
            "DWG_INDEX": str(drawing_index),
            "SOURCE_DRAWING_FILE": source,
            "APPROVED": "YES",
            "ENGINEER": "TESTER",
        }

    def exact_part(self, identifier, checkout="CHECKED_IN", owner="", sheets=0):
        return {
            "state": "EXISTS",
            "opened_identifier": identifier,
            "checkout_state": checkout,
            "checkout_owner": owner,
            "checkout_raw": repr((checkout == "CHECKED_OUT", owner)),
            "drawing_sheet_count": sheets,
            "detail": "exact",
            "error_code": "",
        }

    def not_openable(self):
        return {
            "state": "NOT_OPENABLE",
            "opened_identifier": "",
            "checkout_state": "UNKNOWN",
            "checkout_owner": "",
            "checkout_raw": "",
            "drawing_sheet_count": -1,
            "detail": "NXException - part not found",
            "error_code": "",
        }

    def retrieval(self, digest, native="managed.prt"):
        return {
            "associated_files": native,
            "native_name": native,
            "evidence_file": native,
            "sha256": digest,
            "download_cwd": "download",
        }

    def run_execute(
        self,
        folder,
        rows,
        inspect_side_effect,
        retrieve_side_effect,
        import_side_effect=None,
    ):
        stage = os.path.join(folder, "stage")
        evidence = os.path.join(folder, "evidence")
        os.makedirs(stage)
        os.makedirs(evidence)
        import_mock = mock.Mock(side_effect=import_side_effect)
        with mock.patch.object(
            self.journal, "inspect_exact_part", side_effect=inspect_side_effect
        ), mock.patch.object(
            self.journal,
            "retrieve_single_native",
            side_effect=retrieve_side_effect,
        ), mock.patch.object(
            self.journal, "import_one", import_mock
        ):
            reports = self.journal.execute(
                types.SimpleNamespace(Parts=[]),
                object(),
                object(),
                rows,
                os.path.join(folder, "input.csv"),
                stage,
                evidence,
                "stamp",
                "APPLY_APPROVED",
                FakeLog(),
            )
        return reports, import_mock

    def successful_inspector(self, target_id, model_id, post_checkout="CHECKED_IN"):
        target_calls = {"count": 0}

        def inspect(_session, identifier, _log):
            if identifier == model_id:
                return self.exact_part(model_id)
            self.assertEqual(target_id, identifier)
            target_calls["count"] += 1
            if target_calls["count"] <= 2:
                return self.not_openable()
            return self.exact_part(
                target_id,
                checkout=post_checkout,
                owner="tester" if post_checkout == "CHECKED_OUT" else "",
                sheets=2,
            )

        return inspect

    def successful_retriever(self, target_id, model_id, source_sha, post_sha=None):
        post_sha = post_sha or source_sha
        model_calls = {"count": 0}

        def retrieve(_session, _fm, identifier, _root, _log, *_rest):
            if identifier == model_id:
                model_calls["count"] += 1
                return self.retrieval(self.MODEL_SHA, "MODEL100_A.prt")
            self.assertEqual(target_id, identifier)
            return self.retrieval(post_sha, "MODEL100_A_dwg3.prt")

        return retrieve

    def test_defaults_to_single_run_production_contract(self):
        self.assertEqual("", self.journal.USER_IMPORT_CSV)
        self.assertEqual("APPLY_APPROVED", self.journal.USER_MODE)
        with mock.patch.dict(os.environ, {}, clear=True):
            self.assertEqual("APPLY_APPROVED", self.journal.configured_mode())
            self.assertEqual(
                25, self.journal.configured_max_approved_writes()
            )

    def test_current_j16_dependency_loads_without_legacy_log_symbols(self):
        self.assertTrue(callable(self.journal.J16.retrieve_exact_associated_drawing))
        source = J17_PATH.read_text(encoding="utf-8")
        self.assertNotIn("TARGET_BLOCK_TERMS", source)
        self.assertNotIn("GLOBAL_FAILURE_TERMS", source)
        self.assertNotIn("TARGET_SUCCESS_TERMS", source)
        self.assertNotIn("classify_log", source)

    def test_local_preflight_stages_proven_specification_name(self):
        with tempfile.TemporaryDirectory() as folder:
            source = self.make_source(folder)
            stage = os.path.join(folder, "stage")
            os.makedirs(stage)
            reports, proposals = self.journal.local_preflight(
                [self.make_row(source)],
                os.path.join(folder, "input.csv"),
                stage,
                "stamp",
                "APPLY_APPROVED",
            )
        self.assertEqual("LOCAL_PREFLIGHT_OK", reports[0]["RESULT"])
        self.assertEqual(
            "MODEL100_A_s_MODEL100-A-dwg3.prt",
            os.path.basename(proposals[0]["staged"]),
        )
        self.assertEqual(
            reports[0]["SOURCE_SHA256"], reports[0]["STAGED_SHA256"]
        )
        self.assertEqual(
            "@DB/MODEL100/A/specification/MODEL100-A-dwg3",
            proposals[0]["identifier"],
        )
        self.assertEqual("@DB/MODEL100/A", proposals[0]["model_identifier"])

    def test_invalid_drawing_index_is_rejected(self):
        row = {
            "PART_NUMBER": "MODEL100",
            "REVISION": "A",
            "DWG_INDEX": "zero",
        }
        with self.assertRaisesRegex(RuntimeError, "must be an integer"):
            self.journal.parse_row(row)

    def test_existing_destination_is_quarantined_with_zero_clone_calls(self):
        with tempfile.TemporaryDirectory() as folder:
            source = self.make_source(folder)
            row = self.make_row(source)
            target_id = self.journal.drawing_id("MODEL100", "A", 3)
            model_id = self.journal.master_id("MODEL100", "A")

            def inspect(_session, identifier, _log):
                if identifier == model_id:
                    return self.exact_part(model_id)
                return self.exact_part(target_id, sheets=2)

            reports, import_mock = self.run_execute(
                folder,
                [row],
                inspect,
                mock.Mock(),
            )
        self.assertEqual(
            "QUARANTINED_TARGET_ALREADY_EXISTS", reports[0]["RESULT"]
        )
        self.assertEqual("NO", reports[0]["WRITE_ATTEMPTED"])
        import_mock.assert_not_called()

    def test_ambiguous_destination_identity_is_quarantined(self):
        with tempfile.TemporaryDirectory() as folder:
            source = self.make_source(folder)
            row = self.make_row(source)
            model_id = self.journal.master_id("MODEL100", "A")

            def inspect(_session, identifier, _log):
                if identifier == model_id:
                    return self.exact_part(model_id)
                result = self.not_openable()
                result.update(state="IDENTITY_MISMATCH", detail="wrong object")
                return result

            reports, import_mock = self.run_execute(
                folder, [row], inspect, mock.Mock()
            )
        self.assertEqual(
            "QUARANTINED_TARGET_ALREADY_EXISTS", reports[0]["RESULT"]
        )
        import_mock.assert_not_called()

    def test_checked_out_preserved_3d_blocks_with_owner_and_zero_write(self):
        with tempfile.TemporaryDirectory() as folder:
            source = self.make_source(folder)
            row = self.make_row(source)
            model_id = self.journal.master_id("MODEL100", "A")

            def inspect(_session, identifier, _log):
                self.assertEqual(model_id, identifier)
                return self.exact_part(
                    model_id, checkout="CHECKED_OUT", owner="other.user"
                )

            reports, import_mock = self.run_execute(
                folder, [row], inspect, mock.Mock()
            )
        self.assertEqual("QUARANTINED_PREFLIGHT", reports[0]["RESULT"])
        self.assertIn("other.user", reports[0]["MESSAGE"])
        self.assertEqual("NO", reports[0]["WRITE_ATTEMPTED"])
        import_mock.assert_not_called()

    def test_exact_source_creation_is_verified_and_3d_is_unchanged(self):
        with tempfile.TemporaryDirectory() as folder:
            source = self.make_source(folder)
            row = self.make_row(source)
            source_sha = self.journal.J16.sha256(source)
            target_id = self.journal.drawing_id("MODEL100", "A", 3)
            model_id = self.journal.master_id("MODEL100", "A")
            reports, import_mock = self.run_execute(
                folder,
                [row],
                self.successful_inspector(target_id, model_id),
                self.successful_retriever(
                    target_id, model_id, source_sha
                ),
            )
        report = reports[0]
        self.assertEqual("SPECIFICATION_CREATED_VERIFIED", report["RESULT"])
        self.assertEqual("YES", report["WRITE_ATTEMPTED"])
        self.assertEqual("YES", report["PRESERVE_3D_UNCHANGED"])
        self.assertEqual("2", str(report["POST_IMPORT_DRAWING_SHEET_COUNT"]))
        self.assertEqual(2, import_mock.call_count)
        self.assertTrue(import_mock.call_args_list[0].args[3])
        self.assertFalse(import_mock.call_args_list[1].args[3])

    def test_managed_rewrite_is_verified_after_exact_create(self):
        with tempfile.TemporaryDirectory() as folder:
            source = self.make_source(folder)
            row = self.make_row(source)
            source_sha = self.journal.J16.sha256(source)
            target_id = self.journal.drawing_id("MODEL100", "A", 3)
            model_id = self.journal.master_id("MODEL100", "A")
            reports, _ = self.run_execute(
                folder,
                [row],
                self.successful_inspector(target_id, model_id),
                self.successful_retriever(
                    target_id, model_id, source_sha, post_sha="c" * 64
                ),
            )
        self.assertEqual(
            "SPECIFICATION_CREATED_VERIFIED_MANAGED_TRANSFORM",
            reports[0]["RESULT"],
        )
        self.assertEqual(
            "VERIFIED_MANAGED_TRANSFORM",
            reports[0]["POST_IMPORT_VERIFICATION"],
        )

    def test_local_file_race_quarantines_before_apply(self):
        with tempfile.TemporaryDirectory() as folder:
            source = self.make_source(folder)
            row = self.make_row(source)
            source_sha = self.journal.J16.sha256(source)
            target_id = self.journal.drawing_id("MODEL100", "A", 3)
            model_id = self.journal.master_id("MODEL100", "A")

            def mutate_after_dry(_api, _proposal, _logfile, dry_run, _log):
                if dry_run:
                    with open(source, "ab") as handle:
                        handle.write(b"changed")

            reports, import_mock = self.run_execute(
                folder,
                [row],
                self.successful_inspector(target_id, model_id),
                self.successful_retriever(target_id, model_id, source_sha),
                mutate_after_dry,
            )
        self.assertEqual("QUARANTINED_PREWRITE", reports[0]["RESULT"])
        self.assertEqual("NO", reports[0]["WRITE_ATTEMPTED"])
        self.assertEqual(1, import_mock.call_count)

    def test_destination_created_during_preflight_is_not_overwritten(self):
        with tempfile.TemporaryDirectory() as folder:
            source = self.make_source(folder)
            row = self.make_row(source)
            target_id = self.journal.drawing_id("MODEL100", "A", 3)
            model_id = self.journal.master_id("MODEL100", "A")
            target_calls = {"count": 0}

            def inspect(_session, identifier, _log):
                if identifier == model_id:
                    return self.exact_part(model_id)
                target_calls["count"] += 1
                if target_calls["count"] == 1:
                    return self.not_openable()
                return self.exact_part(target_id, sheets=2)

            reports, import_mock = self.run_execute(
                folder,
                [row],
                inspect,
                lambda *_args: self.retrieval(self.MODEL_SHA, "MODEL100_A.prt"),
            )
        self.assertEqual("QUARANTINED_PREWRITE", reports[0]["RESULT"])
        self.assertEqual("NO", reports[0]["WRITE_ATTEMPTED"])
        self.assertEqual(1, import_mock.call_count)

    def test_changed_preserved_3d_after_write_fails_verification(self):
        with tempfile.TemporaryDirectory() as folder:
            source = self.make_source(folder)
            row = self.make_row(source)
            source_sha = self.journal.J16.sha256(source)
            target_id = self.journal.drawing_id("MODEL100", "A", 3)
            model_id = self.journal.master_id("MODEL100", "A")
            model_calls = {"count": 0}

            def retrieve(_session, _fm, identifier, _root, _log, *_rest):
                if identifier == model_id:
                    model_calls["count"] += 1
                    digest = self.MODEL_SHA if model_calls["count"] < 3 else "b" * 64
                    return self.retrieval(digest, "MODEL100_A.prt")
                return self.retrieval(source_sha, "MODEL100_A_dwg3.prt")

            reports, _ = self.run_execute(
                folder,
                [row],
                self.successful_inspector(target_id, model_id),
                retrieve,
            )
        self.assertEqual("FAILED_IMPORT_UNVERIFIED", reports[0]["RESULT"])
        self.assertEqual("NO", reports[0]["PRESERVE_3D_UNCHANGED"])
        self.assertEqual("YES", reports[0]["WRITE_ATTEMPTED"])

    def test_post_import_checkout_requires_manual_checkin_and_stops_later(self):
        with tempfile.TemporaryDirectory() as folder:
            source1 = self.make_source(folder, "one.prt")
            source2 = self.make_source(folder, "two.prt")
            rows = [
                self.make_row(source1, "MODEL100", 3),
                self.make_row(source2, "MODEL200", 4),
            ]
            target_ids = {
                self.journal.drawing_id("MODEL100", "A", 3),
                self.journal.drawing_id("MODEL200", "A", 4),
            }
            model_ids = {
                self.journal.master_id("MODEL100", "A"),
                self.journal.master_id("MODEL200", "A"),
            }
            target_counts = {identifier: 0 for identifier in target_ids}

            def inspect(_session, identifier, _log):
                if identifier in model_ids:
                    return self.exact_part(identifier)
                target_counts[identifier] += 1
                if target_counts[identifier] <= 2:
                    return self.not_openable()
                return self.exact_part(
                    identifier,
                    checkout="CHECKED_OUT",
                    owner="tester",
                    sheets=2,
                )

            def retrieve(_session, _fm, identifier, _root, _log, *_rest):
                if identifier in model_ids:
                    return self.retrieval(
                        self.MODEL_SHA,
                        identifier.split("/")[2] + "_A.prt",
                    )
                suffix = "dwg3" if "MODEL100" in identifier else "dwg4"
                model = "MODEL100" if "MODEL100" in identifier else "MODEL200"
                return self.retrieval(
                    "c" * 64, "{0}_A_{1}.prt".format(model, suffix)
                )

            reports, import_mock = self.run_execute(
                folder, rows, inspect, retrieve
            )
        self.assertEqual("MANUAL_CHECKIN_REQUIRED", reports[0]["RESULT"])
        self.assertEqual(
            "REVIEW_NOT_ATTEMPTED_AFTER_PRIOR_WRITE", reports[1]["RESULT"]
        )
        self.assertEqual("NO", reports[1]["WRITE_ATTEMPTED"])
        self.assertEqual(3, import_mock.call_count)

    def test_clone_defaults_references_to_use_existing(self):
        staged = os.path.abspath("MODEL100_A_s_MODEL100-A-dwg3.prt")
        proposal = {
            "staged": staged,
            "model_part_number": "MODEL100",
            "model_revision": "A",
            "model_identifier": "@DB/MODEL100/A",
            "report": self.journal.base_report({}, "stamp", "DRY_RUN"),
        }
        clone = mock.MagicMock()
        api = {
            "clone": clone,
            "import_operation": "IMPORT",
            "treat_as_lost": "LOST",
            "autotranslate": "AUTO",
            "use_existing": "USE_EXISTING",
            "overwrite": "OVERWRITE",
        }
        parts = [staged, "@DB/MODEL100/A", "@DB/OTHER100/A"]
        with mock.patch.object(
            self.journal.J16, "add_assembly", return_value=None
        ), mock.patch.object(
            self.journal.J16, "iterate_parts", return_value=parts
        ), mock.patch.object(
            self.journal.J16, "set_action"
        ) as set_action, mock.patch.object(
            self.journal.J16, "perform_clone"
        ):
            self.journal.import_one(api, proposal, "clone.log", True, FakeLog())

        clone.SetDefAction.assert_called_once_with("USE_EXISTING")
        self.assertIn(
            mock.call(clone, staged, "OVERWRITE"), set_action.call_args_list
        )
        self.assertIn(
            mock.call(clone, "@DB/MODEL100/A", "USE_EXISTING"),
            set_action.call_args_list,
        )
        self.assertNotIn(
            mock.call(clone, "@DB/MODEL100/A", "OVERWRITE"),
            set_action.call_args_list,
        )

    def test_old_drawing_identity_cannot_satisfy_exact_3d_discovery(self):
        staged = os.path.abspath("MODEL100_A_s_MODEL100-A-dwg3.prt")
        parts = [staged, "MODEL100_A_dwg1.prt"]
        self.assertEqual(
            [],
            self.journal.find_model_references(
                parts, staged, "MODEL100", "A"
            ),
        )
        parts.append(os.path.abspath("MODEL100_A.prt"))
        self.assertEqual(
            [os.path.abspath("MODEL100_A.prt")],
            self.journal.find_model_references(
                parts, staged, "MODEL100", "A"
            ),
        )

    def test_single_native_retrieval_uses_exact_open_and_restores_cwd(self):
        with tempfile.TemporaryDirectory() as folder:
            identifier = "@DB/MODEL100/A"
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
            pdm_files = [FakePdmFile("meta.qaf"), FakePdmFile("MODEL100_A.prt")]
            download_dir = os.path.join(folder, "download")

            def download(_parts, files):
                self.assertEqual(pdm_files, files)
                os.makedirs(download_dir)
                with open(os.path.join(download_dir, "MODEL100_A.prt"), "wb") as handle:
                    handle.write(b"managed-model")
                os.chdir(download_dir)

            fm = types.SimpleNamespace(
                GetAssociatedFiles=mock.Mock(return_value=(pdm_files,)),
                DownloadAssociatedFiles=mock.Mock(side_effect=download),
            )
            original_cwd = os.getcwd()
            result = self.journal.retrieve_single_native(
                session,
                fm,
                identifier,
                os.path.join(folder, "evidence"),
                FakeLog(),
                "MODEL100_A.prt",
            )

        self.assertEqual(original_cwd, os.getcwd())
        self.assertEqual("MODEL100_A.prt", result["native_name"])
        self.assertEqual(64, len(result["sha256"]))
        self.assertTrue(all(value.released for value in pdm_files))
        self.assertTrue(status.disposed)

    def test_retrieval_identity_mismatch_closes_wrongly_opened_part(self):
        identifier = "@DB/MODEL100/A"
        status = FakeLoadStatus()
        part = types.SimpleNamespace(
            JournalIdentifier="@DB/WRONG100/A",
            Close=mock.Mock(),
        )
        parts = mock.MagicMock()
        parts.__iter__.return_value = iter([])
        parts.FindObject.side_effect = RuntimeError("not loaded")
        parts.OpenBase.return_value = (part, status)
        session = types.SimpleNamespace(Parts=parts)

        with self.assertRaisesRegex(RuntimeError, "does not match"):
            self.journal.open_exact_for_retrieval(
                session, identifier, FakeLog()
            )

        self.assertTrue(status.disposed)
        part.Close.assert_called_once()

    def test_retrieval_no_part_disposes_load_status(self):
        status = FakeLoadStatus()
        parts = mock.MagicMock()
        parts.__iter__.return_value = iter([])
        parts.FindObject.side_effect = RuntimeError("not loaded")
        parts.OpenBase.return_value = (None, status)

        with self.assertRaisesRegex(RuntimeError, "returned no part"):
            self.journal.open_exact_for_retrieval(
                types.SimpleNamespace(Parts=parts),
                "@DB/MODEL100/A",
                FakeLog(),
            )

        self.assertTrue(status.disposed)

    def test_process_state_failure_aborts_prior_clear_preflight_row(self):
        first = self.journal.base_report(
            {"APPROVED": "YES"}, "stamp", "APPLY_APPROVED"
        )
        second = self.journal.base_report(
            {"APPROVED": "YES"}, "stamp", "APPLY_APPROVED"
        )
        first["RESULT"] = "LOCAL_PREFLIGHT_OK"
        second["RESULT"] = "LOCAL_PREFLIGHT_OK"
        proposals = [{"report": first}, {"report": second}]

        def managed(*_args):
            first.update(
                RESULT="CLONE_PREFLIGHT_OK", DISPOSITION="PREFLIGHT_CLEAR"
            )
            raise self.journal.J16.ProcessStateError("cwd restore failed")

        with mock.patch.object(
            self.journal,
            "local_preflight",
            return_value=([first, second], proposals),
        ), mock.patch.object(
            self.journal, "managed_preflight", side_effect=managed
        ):
            reports = self.journal.execute(
                types.SimpleNamespace(Parts=[]),
                object(),
                object(),
                [{"APPROVED": "YES"}],
                "input.csv",
                "stage",
                "evidence",
                "stamp",
                "APPLY_APPROVED",
                FakeLog(),
            )

        self.assertEqual("FAILED_PROCESS_STATE", reports[0]["RESULT"])
        self.assertEqual("FAILED_PROCESS_STATE", reports[1]["RESULT"])

    def test_source_contains_no_direct_teamcenter_mutation_apis(self):
        source = J17_PATH.read_text(encoding="utf-8")
        forbidden = (
            ".Checkin(",
            ".CheckIn(",
            ".Checkout(",
            ".CheckOut(",
            ".Save(",
            ".SaveAs(",
            ".Delete(",
            ".Revise(",
            "ImportFiles(",
        )
        for token in forbidden:
            self.assertNotIn(token, source)
        self.assertIn("clone.SetDefAction(api[\"use_existing\"])", source)
        self.assertIn("api[\"overwrite\"]", source)
        self.assertIn("PRESERVE_3D_UNCHANGED", source)
        self.assertIn("QUARANTINED_TARGET_ALREADY_EXISTS", source)
        self.assertIn("_s_{2}.prt", source)
        self.assertNotIn("_m.prt", source)


if __name__ == "__main__":
    unittest.main()
