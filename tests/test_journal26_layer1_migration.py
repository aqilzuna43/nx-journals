import csv
import datetime
import importlib.util
import json
import os
import sys
import tempfile
import types
import unittest
from pathlib import Path
from unittest import mock


ROOT = Path(__file__).resolve().parents[1]
JOURNAL = (
    ROOT / "from_git" / "journals" / "26_move_solid_bodies_to_layer_1.py"
)


def load_journal():
    nxopen = types.ModuleType("NXOpen")
    nxopen.NXObject = types.SimpleNamespace(
        AttributeType=types.SimpleNamespace(String="String")
    )
    nxopen.Session = types.SimpleNamespace(
        MarkVisibility=types.SimpleNamespace(Visible="Visible"),
        LibraryUnloadOption=types.SimpleNamespace(Immediately="Immediately"),
    )
    prior = sys.modules.get("NXOpen")
    sys.modules["NXOpen"] = nxopen
    try:
        spec = importlib.util.spec_from_file_location("journal26", JOURNAL)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        return module
    finally:
        if prior is None:
            sys.modules.pop("NXOpen", None)
        else:
            sys.modules["NXOpen"] = prior


class FakeBody:
    def __init__(
        self,
        name,
        tag,
        layer,
        solid=True,
        sheet=False,
        convergent=False,
        blanked=False,
    ):
        self.Name = name
        self.Tag = tag
        self.Layer = layer
        self.IsSolidBody = solid
        self.IsSheetBody = sheet
        self.IsConvergentBody = convergent
        self.IsBlanked = blanked
        self.OwningPart = None


class FakePdmPart:
    def __init__(self, checked=True, owner="aqil"):
        self.checked = checked
        self.owner = owner
        self.calls = 0

    def GetCheckedoutStatusAndUser(self):
        self.calls += 1
        if self.checked == "ERROR":
            raise RuntimeError("checkout lookup failed")
        if self.checked == "UNKNOWN":
            return ("unknown-status", self.owner)
        return self.checked, self.owner


class FakePdmSession:
    def __init__(self, user="aqil"):
        self.user = user

    def GetUserName(self):
        return self.user


class FakeLayers:
    def __init__(self, part, work_layer=7):
        self.part = part
        self.WorkLayer = work_layer
        self.states = {index: "STATE_{0}".format(index) for index in range(1, 257)}
        self.calls = []
        self.move_error = None
        self.raise_after_move = False
        self.ignore_move = False
        self.mutate_blanking = False
        self.mutate_layer_state = False
        self.change_membership = False

    def GetState(self, layer):
        return self.states[layer]

    def MoveDisplayableObjects(self, layer, objects):
        self.calls.append((layer, list(objects)))
        if self.move_error and not self.raise_after_move:
            raise RuntimeError(self.move_error)
        if not self.ignore_move:
            for body in objects:
                body.Layer = layer
        if self.mutate_blanking and objects:
            objects[0].IsBlanked = not objects[0].IsBlanked
        if self.mutate_layer_state:
            self.states[2] = "MUTATED"
        if self.change_membership:
            self.part.Bodies = list(self.part.Bodies[:-1])
        if self.move_error and self.raise_after_move:
            raise RuntimeError(self.move_error)


class FakePart:
    _next_tag = 10000

    def __init__(
        self,
        bodies,
        managed=False,
        read_only=False,
        checked=True,
        owner="aqil",
    ):
        FakePart._next_tag += 1
        self.Tag = FakePart._next_tag
        self.Name = "TEST_PART"
        self.Leaf = "TEST_PART"
        self.FullPath = "@DB/TEST_PART/A" if managed else r"C:\temp\TEST_PART.prt"
        self.JournalIdentifier = self.FullPath
        self.IsReadOnly = read_only
        self.Bodies = list(bodies)
        self.Layers = FakeLayers(self)
        self.PDMPart = FakePdmPart(checked, owner) if managed else None
        self.attributes = {
            "DB_PART_NO": "TEST_PART",
            "DB_PART_REV": "A",
        }
        for body in self.Bodies:
            body.OwningPart = self

    def GetStringAttribute(self, name):
        return self.attributes.get(name, "")


class FakeListingWindow:
    def __init__(self):
        self.lines = []
        self.opened = False

    def Open(self):
        self.opened = True

    def WriteFullline(self, value):
        self.lines.append(str(value))


class FakeSession:
    def __init__(
        self,
        part,
        managed=False,
        user="aqil",
        undo_error=False,
        mark_error=False,
    ):
        self.Parts = types.SimpleNamespace(Work=part)
        self.IsManagedMode = managed
        self.PdmSession = FakePdmSession(user)
        self.ListingWindow = FakeListingWindow()
        self.undo_error = undo_error
        self.mark_error = mark_error
        self.set_mark_calls = []
        self.undo_calls = []
        self.delete_calls = []
        self._baseline = None

    def SetUndoMark(self, visibility, name):
        self.set_mark_calls.append((visibility, name))
        if self.mark_error:
            raise RuntimeError("undo mark unavailable")
        part = self.Parts.Work
        self._baseline = {
            "bodies": list(part.Bodies),
            "body_state": {
                body.Tag: (body.Layer, body.IsBlanked) for body in part.Bodies
            },
            "work_layer": part.Layers.WorkLayer,
            "states": dict(part.Layers.states),
        }
        return "MARK-1"

    def UndoToMark(self, mark, name):
        self.undo_calls.append((mark, name))
        if self.undo_error:
            raise RuntimeError("undo failed")
        part = self.Parts.Work
        part.Bodies = list(self._baseline["bodies"])
        for body in part.Bodies:
            body.Layer, body.IsBlanked = self._baseline["body_state"][body.Tag]
        part.Layers.WorkLayer = self._baseline["work_layer"]
        part.Layers.states = dict(self._baseline["states"])

    def DeleteUndoMark(self, mark, name):
        self.delete_calls.append((mark, name))


class LayerOneMigrationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.journal = load_journal()

    def run_in_temp(self, session, mode="DRY_RUN", **changes):
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        now = changes.pop(
            "run_datetime",
            datetime.datetime(
                2026,
                8,
                28,
                12,
                30,
                tzinfo=datetime.timezone(datetime.timedelta(hours=8)),
            ),
        )
        with mock.patch.object(self.journal, "io_root", return_value=folder.name):
            result = self.journal.run(
                session,
                run_datetime=now,
                mode=mode,
                **changes
            )
        return result

    def test_configured_mode_defaults_and_rejects_invalid_value(self):
        with mock.patch.dict(os.environ, {}, clear=True), mock.patch.object(
            self.journal, "USER_MODE", "DRY_RUN"
        ):
            self.assertEqual("DRY_RUN", self.journal.configured_mode())
        with mock.patch.dict(os.environ, {"NX_J26_MODE": "apply"}, clear=True):
            self.assertEqual("APPLY", self.journal.configured_mode())
        with mock.patch.dict(os.environ, {"NX_J26_MODE": "MOVE"}, clear=True):
            with self.assertRaisesRegex(RuntimeError, "DRY_RUN or APPLY"):
                self.journal.configured_mode()

    def test_capture_includes_blanked_solids_and_classifies_exclusions(self):
        solid = FakeBody("HIDDEN_SOLID", 1, 12, blanked=True)
        sheet = FakeBody("SHEET", 2, 20, solid=False, sheet=True)
        convergent = FakeBody("CONVERGENT", 3, 30, convergent=True)
        other = FakeBody("OTHER", 4, 40, solid=False)
        part = FakePart([solid, sheet, convergent, other])

        snapshot = self.journal.capture_snapshot(part)
        counts = self.journal.snapshot_counts(snapshot)

        self.assertEqual(
            [item["body_type"] for item in snapshot["bodies"]],
            ["TRADITIONAL_SOLID", "SHEET", "CONVERGENT", "OTHER"],
        )
        self.assertTrue(snapshot["bodies"][0]["blanked"])
        self.assertEqual(counts["eligible_solid_count"], 1)
        self.assertEqual(counts["skipped_sheet_count"], 1)
        self.assertEqual(counts["skipped_convergent_count"], 1)
        self.assertEqual(counts["skipped_other_count"], 1)

    def test_capture_fails_closed_for_foreign_owned_body(self):
        body = FakeBody("SOLID", 1, 2)
        part = FakePart([body])
        body.OwningPart = FakePart([])

        with self.assertRaisesRegex(RuntimeError, "not owned"):
            self.journal.capture_snapshot(part)

    def test_dry_run_writes_paired_evidence_without_mutation(self):
        move = FakeBody("MOVE", 11, 8, blanked=True)
        existing = FakeBody("EXISTING", 12, 1)
        sheet = FakeBody("SHEET", 13, 9, solid=False, sheet=True)
        part = FakePart([move, existing, sheet])
        session = FakeSession(part)
        original_states = dict(part.Layers.states)

        csv_path, json_path, report = self.run_in_temp(session)

        self.assertEqual(report["verdict"]["status"], "DRY_RUN_READY")
        self.assertEqual(move.Layer, 8)
        self.assertTrue(move.IsBlanked)
        self.assertEqual(part.Layers.calls, [])
        self.assertEqual(session.set_mark_calls, [])
        self.assertEqual(part.Layers.states, original_states)
        self.assertEqual(Path(csv_path).parent, Path(json_path).parent)
        self.assertEqual(Path(csv_path).parent.parent.name, "NX_LAYER_1_MIGRATION")
        self.assertTrue(Path(csv_path).read_bytes().startswith(b"\xef\xbb\xbf"))
        with open(csv_path, "r", encoding="utf-8-sig", newline="") as handle:
            rows = list(csv.DictReader(handle))
        self.assertEqual(rows[0]["ROW_TYPE"], "SUMMARY")
        move_row = next(row for row in rows if row["BODY_TAG"] == "11")
        self.assertEqual(move_row["ACTION"], "WOULD_MOVE")
        payload = json.loads(Path(json_path).read_text(encoding="utf-8"))
        self.assertEqual(payload["schema_version"], 1)
        self.assertEqual(len(payload["before"]["layer_states"]), 256)

    def test_already_compliant_is_idempotent_without_undo_mark(self):
        body = FakeBody("SOLID", 1, 1)
        part = FakePart([body])
        session = FakeSession(part)

        _, _, report = self.run_in_temp(session, mode="APPLY")

        self.assertEqual(report["verdict"]["status"], "ALREADY_COMPLIANT")
        self.assertEqual(part.Layers.calls, [])
        self.assertEqual(session.set_mark_calls, [])

    def test_no_eligible_solids_is_reported_as_noop(self):
        sheet = FakeBody("SHEET", 2, 25, solid=False, sheet=True)
        part = FakePart([sheet])
        session = FakeSession(part)

        csv_path, _, report = self.run_in_temp(session, mode="APPLY")

        self.assertEqual(report["verdict"]["status"], "NO_ELIGIBLE_SOLIDS")
        self.assertTrue(Path(csv_path).is_file())
        self.assertEqual(session.set_mark_calls, [])

    def test_no_work_part_still_writes_failure_evidence(self):
        session = FakeSession(None)

        csv_path, json_path, report = self.run_in_temp(session)

        self.assertEqual(report["verdict"]["status"], "FAILED_NO_WORK_PART")
        self.assertTrue(Path(csv_path).is_file())
        self.assertTrue(Path(json_path).is_file())

    def test_invalid_explicit_mode_fails_before_outputs(self):
        part = FakePart([FakeBody("SOLID", 1, 2)])
        with self.assertRaisesRegex(RuntimeError, "DRY_RUN or APPLY"):
            self.run_in_temp(FakeSession(part), mode="MOVE")

    def test_native_read_only_part_is_blocked(self):
        body = FakeBody("SOLID", 1, 2)
        part = FakePart([body], read_only=True)
        session = FakeSession(part)

        _, _, report = self.run_in_temp(session, mode="APPLY")

        self.assertEqual(report["verdict"]["status"], "BLOCKED_WRITE_ACCESS")
        self.assertEqual(body.Layer, 2)
        self.assertEqual(session.set_mark_calls, [])

    def test_managed_apply_requires_known_current_user_checkout(self):
        cases = (
            (False, "aqil", "aqil", "CHECKED_IN"),
            ("UNKNOWN", "", "aqil", "UNKNOWN"),
            (True, "other.user", "aqil", "CHECKED_OUT"),
            (True, "aqil", "", "CHECKED_OUT"),
        )
        for checked, owner, user, expected_state in cases:
            with self.subTest(checked=checked, owner=owner, user=user):
                body = FakeBody("SOLID", 1, 2)
                part = FakePart(
                    [body], managed=True, checked=checked, owner=owner
                )
                session = FakeSession(part, managed=True, user=user)

                _, _, report = self.run_in_temp(session, mode="APPLY")

                self.assertEqual(
                    report["verdict"]["status"], "BLOCKED_WRITE_ACCESS"
                )
                self.assertEqual(report["access"]["checkout_state"], expected_state)
                self.assertEqual(body.Layer, 2)

    def test_managed_apply_blocks_unknown_read_only_state(self):
        body = FakeBody("SOLID", 1, 2)
        part = FakePart(
            [body], managed=True, checked=True, owner="aqil"
        )
        del part.IsReadOnly
        session = FakeSession(part, managed=True, user="aqil")

        _, _, report = self.run_in_temp(session, mode="APPLY")

        self.assertEqual(report["verdict"]["status"], "BLOCKED_WRITE_ACCESS")
        self.assertIsNone(report["access"]["read_only"])
        self.assertIn("unavailable", report["access"]["message"])
        self.assertEqual(body.Layer, 2)

    def test_managed_current_user_apply_moves_exact_candidates(self):
        move_a = FakeBody("MOVE_A", 1, 2, blanked=True)
        move_b = FakeBody("MOVE_B", 2, 200)
        existing = FakeBody("EXISTING", 3, 1)
        sheet = FakeBody("SHEET", 4, 5, solid=False, sheet=True)
        part = FakePart(
            [move_a, move_b, existing, sheet],
            managed=True,
            checked=True,
            owner="aqil",
        )
        session = FakeSession(part, managed=True, user="AQIL")
        original_states = dict(part.Layers.states)
        original_work_layer = part.Layers.WorkLayer

        _, _, report = self.run_in_temp(session, mode="APPLY")

        self.assertEqual(report["verdict"]["status"], "APPLIED_VERIFIED")
        self.assertEqual(len(part.Layers.calls), 1)
        target, objects = part.Layers.calls[0]
        self.assertEqual(target, 1)
        self.assertEqual(objects, [move_a, move_b])
        self.assertEqual([move_a.Layer, move_b.Layer, existing.Layer], [1, 1, 1])
        self.assertEqual(sheet.Layer, 5)
        self.assertTrue(move_a.IsBlanked)
        self.assertEqual(part.Layers.states, original_states)
        self.assertEqual(part.Layers.WorkLayer, original_work_layer)
        self.assertEqual(len(session.set_mark_calls), 1)
        self.assertEqual(session.undo_calls, [])
        self.assertEqual(session.delete_calls, [])
        self.assertTrue(report["action"]["successful_change_left_undoable"])
        self.assertEqual(report["access"]["checkout_owner"], "aqil")

    def test_move_exception_rolls_back_and_verifies_baseline(self):
        body = FakeBody("SOLID", 1, 2)
        part = FakePart([body])
        part.Layers.move_error = "NX move failed"
        session = FakeSession(part)

        _, _, report = self.run_in_temp(session, mode="APPLY")

        self.assertEqual(report["verdict"]["status"], "ROLLED_BACK")
        self.assertEqual(report["rollback"]["status"], "ROLLED_BACK")
        self.assertEqual(body.Layer, 2)
        self.assertEqual(len(session.undo_calls), 1)
        self.assertEqual(len(session.delete_calls), 1)

    def test_post_move_layer_mismatch_rolls_back(self):
        body = FakeBody("SOLID", 1, 2)
        part = FakePart([body])
        part.Layers.ignore_move = True
        session = FakeSession(part)

        _, _, report = self.run_in_temp(session, mode="APPLY")

        self.assertEqual(report["verdict"]["status"], "ROLLED_BACK")
        self.assertTrue(report["action"]["verification_errors"])
        self.assertEqual(body.Layer, 2)

    def test_unexpected_blanking_layer_state_or_membership_change_rolls_back(self):
        flags = ("mutate_blanking", "mutate_layer_state", "change_membership")
        for flag in flags:
            with self.subTest(flag=flag):
                move = FakeBody("MOVE", 1, 2)
                retained = FakeBody("RETAINED", 2, 1)
                part = FakePart([move, retained])
                setattr(part.Layers, flag, True)
                session = FakeSession(part)

                _, _, report = self.run_in_temp(session, mode="APPLY")

                self.assertEqual(report["verdict"]["status"], "ROLLED_BACK")
                self.assertEqual([item.Tag for item in part.Bodies], [1, 2])
                self.assertEqual(move.Layer, 2)
                self.assertFalse(move.IsBlanked)
                self.assertEqual(part.Layers.states[2], "STATE_2")

    def test_partial_move_with_undo_failure_is_prominent(self):
        body = FakeBody("SOLID", 1, 2)
        part = FakePart([body])
        part.Layers.move_error = "partial NX failure"
        part.Layers.raise_after_move = True
        session = FakeSession(part, undo_error=True)

        _, _, report = self.run_in_temp(session, mode="APPLY")

        self.assertEqual(report["verdict"]["status"], "ROLLBACK_FAILED")
        self.assertEqual(report["rollback"]["status"], "ROLLBACK_FAILED")
        self.assertEqual(body.Layer, 1)
        self.assertIn("UndoToMark failed", report["rollback"]["error"])

    def test_evidence_failure_after_success_rolls_back_then_records_result(self):
        body = FakeBody("SOLID", 1, 2)
        part = FakePart([body])
        session = FakeSession(part)
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        original_write = self.journal.write_outputs
        calls = []

        def flaky_write(report, output_folder, stem):
            calls.append(report["verdict"]["status"])
            if len(calls) == 1:
                raise OSError("disk interrupted")
            return original_write(report, output_folder, stem)

        with mock.patch.object(
            self.journal, "io_root", return_value=folder.name
        ), mock.patch.object(
            self.journal, "write_outputs", side_effect=flaky_write
        ):
            csv_path, json_path, report = self.journal.run(
                session,
                run_datetime=datetime.datetime(2026, 8, 28, 12, 30),
                mode="APPLY",
            )

        self.assertEqual(calls, ["APPLIED_VERIFIED", "ROLLED_BACK"])
        self.assertEqual(report["verdict"]["status"], "ROLLED_BACK")
        self.assertEqual(body.Layer, 2)
        self.assertTrue(Path(csv_path).is_file())
        self.assertTrue(Path(json_path).is_file())
        self.assertIn("Evidence write failed", report["errors"][0])

    def test_undo_guard_failure_never_calls_move(self):
        body = FakeBody("SOLID", 1, 2)
        part = FakePart([body])
        session = FakeSession(part, mark_error=True)

        _, _, report = self.run_in_temp(session, mode="APPLY")

        self.assertEqual(report["verdict"]["status"], "ROLLED_BACK")
        self.assertFalse(report["action"]["attempted"])
        self.assertEqual(part.Layers.calls, [])
        self.assertEqual(body.Layer, 2)

    def test_log_line_uses_listing_window(self):
        session = FakeSession(None)

        self.journal.log_line(session, "first\nsecond")

        self.assertTrue(session.ListingWindow.opened)
        self.assertEqual(session.ListingWindow.lines, ["first", "second"])

    def test_journal_has_no_runtime_save_checkout_or_checkin_calls(self):
        source = JOURNAL.read_text(encoding="utf-8")

        self.assertNotIn(".Save(", source)
        self.assertNotIn(".Checkout(", source)
        self.assertNotIn(".Checkin(", source)
        self.assertNotIn("ComponentAssembly", source)


if __name__ == "__main__":
    unittest.main()
