"""저장한 업무일지 날짜를 바꾸면 내용이 새 날짜로 옮겨지는지."""
from __future__ import annotations

import os
import tempfile
import unittest
from datetime import date
from unittest.mock import patch

from openpyxl import Workbook, load_workbook


class _FakeSS(dict):
    def pop(self, key, default=None):
        return dict.pop(self, key, default)


def _write_template(path: str) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws["C5"] = ""
    ws["C8"] = ""
    ws["G8"] = ""
    wb.save(path)
    wb.close()


class WorklogDateChangeTest(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.cache = os.path.join(self._tmp.name, "worklog")
        self.arch = os.path.join(self._tmp.name, "Desktop", "업무", "일지")
        os.makedirs(self.cache, exist_ok=True)
        os.makedirs(self.arch, exist_ok=True)
        self.tpl = os.path.join(self.cache, "template.xlsx")
        _write_template(self.tpl)
        self.ss = _FakeSS()
        import worklog_tab as wt

        self.wt = wt
        self._patches = [
            patch.object(wt, "WORKLOG_DIR", self.cache),
            patch.object(wt, "WORKLOG_TEMPLATE", self.tpl),
            patch.object(wt, "resolve_worklog_archive_root", return_value=self.arch),
            patch.object(wt.st, "session_state", self.ss),
            patch.object(wt, "_schedule_worklog_remote_delete"),
        ]
        for p in self._patches:
            p.start()

    def tearDown(self):
        for p in reversed(self._patches):
            p.stop()
        self._tmp.cleanup()

    def _cells(self, d: date, client: str, content: str) -> dict:
        cells = self.wt._empty_cells(d)
        cells["C8"] = client
        cells["G8"] = content
        return cells

    def test_on_change_queues_pending_date_move(self):
        self.ss["worklog_selected"] = date(2026, 9, 7)
        self.ss["wl_date_pick"] = date(2026, 9, 8)
        self.wt._on_wl_date_pick_change()
        self.assertEqual(self.ss["wl_pending_date_change"], ("2026-09-07", "2026-09-08"))

    def test_apply_after_save_moves_file_archive_and_session(self):
        old, new = date(2026, 9, 7), date(2026, 9, 8)
        cells = self._cells(old, "거래처A", "저장한 내용")
        self.wt.save_worklog_cells(old, cells, force=True, allow_overwrite=True)
        self.assertTrue(os.path.isfile(self.wt.worklog_path(old)))
        self.assertTrue(self.wt.worklog_date_exists_in_archive(old))

        with patch.object(self.wt, "_cells_from_widgets", side_effect=lambda d: self.wt.read_worklog_cells(d)):
            err = self.wt.apply_worklog_date_change(old, new)
        self.assertEqual(err, "")
        self.assertFalse(os.path.isfile(self.wt.worklog_path(old)))
        self.assertTrue(os.path.isfile(self.wt.worklog_path(new)))
        moved = self.wt.read_worklog_cells(new)
        self.assertEqual(moved.get("C8"), "거래처A")
        self.assertEqual(moved.get("G8"), "저장한 내용")
        self.assertEqual(self.ss["worklog_selected"], new)
        self.assertEqual(self.ss["wl_date_pick"], new)
        self.assertEqual(self.ss["wl_date_sync"], new.isoformat())
        self.assertFalse(self.wt.worklog_date_exists_in_archive(old))
        self.assertTrue(self.wt.worklog_date_exists_in_archive(new))
        month_path = os.path.join(self.arch, "2026", "9월.xlsx")
        wb = load_workbook(month_path, read_only=True)
        try:
            self.assertNotIn("7", wb.sheetnames)
            self.assertIn("8", wb.sheetnames)
        finally:
            wb.close()

    def test_saved_7th_moves_to_3rd(self):
        old, new = date(2026, 9, 7), date(2026, 9, 3)
        self.wt.save_worklog_cells(old, self._cells(old, "거래처7", "7일 내용"), force=True, allow_overwrite=True)
        with patch.object(self.wt, "_cells_from_widgets", side_effect=lambda d: self.wt.read_worklog_cells(d)):
            err = self.wt.apply_worklog_date_change(old, new)
        self.assertEqual(err, "")
        self.assertFalse(os.path.isfile(self.wt.worklog_path(old)))
        self.assertTrue(os.path.isfile(self.wt.worklog_path(new)))
        moved = self.wt.read_worklog_cells(new)
        self.assertEqual(moved.get("C8"), "거래처7")
        self.assertEqual(moved.get("G8"), "7일 내용")
        self.assertEqual(self.ss["worklog_selected"], new)
        self.assertEqual(self.ss["wl_date_pick"], new)

    def test_saved_7th_overwrites_existing_3rd(self):
        old, new = date(2026, 9, 7), date(2026, 9, 3)
        self.wt.save_worklog_cells(old, self._cells(old, "A", "7일 내용"), force=True, allow_overwrite=True)
        self.wt.save_worklog_cells(new, self._cells(new, "B", "3일 옛내용"), force=True, allow_overwrite=True)
        self.ss["worklog_selected"] = old
        with patch.object(self.wt, "_cells_from_widgets", side_effect=lambda d: self.wt.read_worklog_cells(d)):
            err = self.wt.apply_worklog_date_change(old, new)
        self.assertEqual(err, "")
        self.assertFalse(os.path.isfile(self.wt.worklog_path(old)))
        self.assertEqual(self.wt.read_worklog_cells(new).get("G8"), "7일 내용")
        self.assertFalse(self.wt.worklog_date_exists_in_archive(old))
        self.assertTrue(self.wt.worklog_date_exists_in_archive(new))

    def test_empty_day_only_retargets_selected(self):
        old, new = date(2026, 9, 7), date(2026, 9, 8)
        self.ss["worklog_selected"] = old
        with patch.object(self.wt, "_cells_from_widgets", side_effect=lambda d: self.wt._empty_cells(d)):
            err = self.wt.apply_worklog_date_change(old, new)
        self.assertEqual(err, "")
        self.assertFalse(os.path.isfile(self.wt.worklog_path(old)))
        self.assertFalse(os.path.isfile(self.wt.worklog_path(new)))
        self.assertEqual(self.ss["worklog_selected"], new)
        self.assertEqual(self.ss["wl_date_pick"], new)

    def test_pending_runner_moves_saved_7th_to_3rd(self):
        old, new = date(2026, 9, 7), date(2026, 9, 3)
        self.wt.save_worklog_cells(old, self._cells(old, "A", "7일 내용"), force=True, allow_overwrite=True)
        self.ss["worklog_selected"] = old
        self.ss["wl_pending_date_change"] = (old.isoformat(), new.isoformat())
        with patch.object(self.wt, "_cells_from_widgets", side_effect=lambda d: self.wt.read_worklog_cells(d)):
            self.assertTrue(self.wt._run_pending_worklog_date_change())
        self.assertEqual(self.ss["worklog_selected"], new)
        self.assertEqual(self.ss["wl_date_pick"], new)
        self.assertFalse(self.ss.get("wl_date_err"))
        self.assertFalse(os.path.isfile(self.wt.worklog_path(old)))
        self.assertEqual(self.wt.read_worklog_cells(new).get("G8"), "7일 내용")

    def test_queue_save_uses_picked_date_and_pending_move(self):
        old, new = date(2026, 9, 7), date(2026, 9, 8)
        self.ss["worklog_selected"] = old
        self.ss["wl_date_pick"] = new
        self.wt._queue_worklog_save(old.isoformat())
        self.assertEqual(self.ss["wl_pending_date_change"], (old.isoformat(), new.isoformat()))
        self.assertTrue(self.ss[f"wl_do_save_{new.isoformat()}"])
        self.assertFalse(self.ss.get(f"wl_do_save_{old.isoformat()}"))

    def test_pending_move_keeps_save_flag_on_new_date(self):
        old, new = date(2026, 9, 7), date(2026, 9, 8)
        cells = self._cells(old, "거래처A", "저장한 내용")
        self.wt.save_worklog_cells(old, cells, force=True, allow_overwrite=True)
        self.ss["worklog_selected"] = old
        self.ss["wl_pending_date_change"] = (old.isoformat(), new.isoformat())
        self.ss[f"wl_do_save_{old.isoformat()}"] = True
        with patch.object(self.wt, "_cells_from_widgets", side_effect=lambda d: self.wt.read_worklog_cells(d)):
            self.assertTrue(self.wt._run_pending_worklog_date_change())
        self.assertFalse(self.ss.get(f"wl_do_save_{old.isoformat()}"))
        self.assertTrue(self.ss.get(f"wl_do_save_{new.isoformat()}"))
        self.assertEqual(self.ss["worklog_selected"], new)
        self.assertTrue(self.ss.get(f"wl_saved_ok_{new.isoformat()}"))
        self.assertTrue((self.ss.get(f"wl_open_ctx_{new.isoformat()}") or {}).get("had_local"))

    def test_save_after_move_overwrites_new_date(self):
        old, new = date(2026, 9, 7), date(2026, 9, 8)
        cells = self._cells(old, "거래처A", "저장한 내용")
        self.wt.save_worklog_cells(old, cells, force=True, allow_overwrite=True)
        with patch.object(self.wt, "_cells_from_widgets", side_effect=lambda d: self.wt.read_worklog_cells(d)):
            self.assertEqual(self.wt.apply_worklog_date_change(old, new), "")
        updated = self.wt.read_worklog_cells(new)
        updated["G8"] = "이동 후 수정"
        path = self.wt.save_worklog_cells(
            new, updated, force=True,
            allow_overwrite=bool(self.ss.get(f"wl_saved_ok_{new.isoformat()}")),
        )
        self.assertTrue(path)
        self.assertEqual(self.wt.read_worklog_cells(new).get("G8"), "이동 후 수정")
        self.assertFalse(os.path.isfile(self.wt.worklog_path(old)))


if __name__ == "__main__":
    unittest.main()
