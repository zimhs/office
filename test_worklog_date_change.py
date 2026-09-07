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

    def test_set_date_pick_queues_when_widget_already_exists(self):
        real_setitem = _FakeSS.__setitem__

        def _setitem(d, k, v):
            if k == "wl_date_pick":
                raise self.wt.StreamlitAPIException("cannot be modified after the widget is instantiated")
            return real_setitem(d, k, v)

        with patch.object(_FakeSS, "__setitem__", _setitem):
            self.wt._set_wl_date_pick(date(2026, 9, 3))
        self.assertEqual(self.ss.get("_wl_date_pick_next"), date(2026, 9, 3))
        self.assertEqual(self.ss.get("wl_date_sync"), "2026-09-03")
        self.wt._flush_queued_date_pick()
        self.assertEqual(self.ss.get("wl_date_pick"), date(2026, 9, 3))
        self.assertIsNone(self.ss.get("_wl_date_pick_next"))

    def test_set_date_pick_queues_when_widget_live(self):
        self.ss["_wl_date_pick_live"] = True
        self.ss["wl_date_pick"] = date(2026, 9, 7)
        self.wt._set_wl_date_pick(date(2026, 9, 3))
        self.assertEqual(self.ss.get("_wl_date_pick_next"), date(2026, 9, 3))
        self.assertEqual(self.ss.get("wl_date_pick"), date(2026, 9, 7))
        self.assertEqual(self.ss.get("wl_date_sync"), "2026-09-03")

    def test_on_change_retargets_without_writing_files(self):
        old, new = date(2026, 9, 7), date(2026, 9, 3)
        self.wt.save_worklog_cells(old, self._cells(old, "거래처7", "7일 내용"), force=True, allow_overwrite=True)
        self.ss["worklog_selected"] = old
        self.ss["wl_date_pick"] = new
        with patch.object(self.wt, "_cells_from_widgets", side_effect=lambda d: self.wt.read_worklog_cells(d)):
            self.wt._on_wl_date_pick_change()
        self.assertEqual(self.ss["worklog_selected"], new)
        self.assertEqual(self.ss.get("wl_date_retarget_from"), old.isoformat())
        self.assertFalse(self.ss.get("wl_date_err"))
        self.assertTrue(os.path.isfile(self.wt.worklog_path(old)))
        self.assertFalse(os.path.isfile(self.wt.worklog_path(new)))
        self.assertTrue(self.ss.get("wl_need_app_rerun"))

    def test_left_date_pick_consumes_retarget_before_widget(self):
        old, new = date(2026, 9, 7), date(2026, 9, 3)
        self.wt.save_worklog_cells(old, self._cells(old, "거래처7", "7일 내용"), force=True, allow_overwrite=True)
        self.ss["worklog_selected"] = old
        self.ss["wl_date_pick"] = new
        with patch.object(self.wt, "_cells_from_widgets", side_effect=lambda d: self.wt.read_worklog_cells(d)):
            got, moved = self.wt.consume_left_date_pick_move(old)
        self.assertTrue(moved)
        self.assertEqual(got, new)
        self.assertEqual(self.ss["worklog_selected"], new)
        self.assertFalse(self.ss.get("wl_date_err"))
        self.assertTrue(os.path.isfile(self.wt.worklog_path(old)))
        self.assertTrue(self.ss.get("wl_need_app_rerun"))

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

    def test_archive_only_7th_moves_to_3rd_updates_calendar_dots(self):
        old, new = date(2026, 9, 7), date(2026, 9, 3)
        self.wt.save_worklog_cells(old, self._cells(old, "거래처7", "7일 내용"), force=True, allow_overwrite=True)
        os.remove(self.wt.worklog_path(old))
        self.wt._invalidate_saved_dates_cache()
        self.ss.pop(f"wl_arch_exists_{old.isoformat()}", None)
        self.ss.pop(f"wl_presence_{old.isoformat()}_fast", None)
        self.ss.pop(f"wl_presence_{old.isoformat()}_all", None)
        self.assertFalse(os.path.isfile(self.wt.worklog_path(old)))
        self.assertTrue(self.wt.worklog_date_exists_in_archive(old))
        with patch.object(self.wt, "_cells_from_widgets", side_effect=lambda d: self.wt._empty_cells(d)):
            err = self.wt.apply_worklog_date_change(old, new)
        self.assertEqual(err, "")
        self.assertTrue(os.path.isfile(self.wt.worklog_path(new)))
        self.assertEqual(self.wt.read_worklog_cells(new).get("G8"), "7일 내용")
        saved = self.wt.list_saved_worklog_dates()
        self.assertIn(new.isoformat(), saved)
        self.assertNotIn(old.isoformat(), saved)
        cal = self.wt._saved_dates_for_calendar()
        self.assertIn(new.isoformat(), cal)
        self.assertNotIn(old.isoformat(), cal)

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

    def test_pending_runner_retargets_without_error(self):
        old, new = date(2026, 9, 7), date(2026, 9, 3)
        self.wt.save_worklog_cells(old, self._cells(old, "A", "7일 내용"), force=True, allow_overwrite=True)
        self.ss["worklog_selected"] = old
        self.ss["wl_pending_date_change"] = (old.isoformat(), new.isoformat())
        with patch.object(self.wt, "_cells_from_widgets", side_effect=lambda d: self.wt.read_worklog_cells(d)):
            self.assertTrue(self.wt._run_pending_worklog_date_change())
        self.assertEqual(self.ss["worklog_selected"], new)
        self.assertFalse(self.ss.get("wl_date_err"))
        self.assertTrue(os.path.isfile(self.wt.worklog_path(old)))

    def test_queue_save_uses_picked_date(self):
        old, new = date(2026, 9, 7), date(2026, 9, 8)
        self.ss["worklog_selected"] = old
        self.ss["wl_date_pick"] = new
        self.wt._queue_worklog_save(old.isoformat())
        self.assertEqual(self.ss["wl_date_retarget_from"], old.isoformat())
        self.assertTrue(self.ss[f"wl_do_save_{new.isoformat()}"])
        self.assertTrue(self.ss[f"wl_do_save_{old.isoformat()}"])

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

    def test_commit_save_overwrites_existing_3rd(self):
        """날짜를 3일로 바꾸고 저장하면 이미 있습니다 없이 3일에 덮어 쓴다."""
        old, new = date(2026, 9, 7), date(2026, 9, 3)
        self.wt.save_worklog_cells(old, self._cells(old, "거래처7", "7일 내용"), force=True, allow_overwrite=True)
        self.wt.save_worklog_cells(new, self._cells(new, "옛3", "3일 옛내용"), force=True, allow_overwrite=True)
        cells = self._cells(new, "거래처7", "7일 내용")
        path = self.wt.commit_worklog_date_save(old, new, cells)
        self.assertTrue(path)
        self.assertFalse(self.ss.get("wl_date_err"))
        self.assertFalse(os.path.isfile(self.wt.worklog_path(old)))
        self.assertEqual(self.wt.read_worklog_cells(new).get("G8"), "7일 내용")
        self.assertEqual(self.ss["worklog_selected"], new)

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

    def _make_archive_only(self, d: date, cells: dict) -> None:
        self.wt.save_worklog_cells(d, cells, force=True, allow_overwrite=True)
        os.remove(self.wt.worklog_path(d))
        self.wt._invalidate_saved_dates_cache()
        self.wt._invalidate_worklog_presence_cache(d)

    def test_empty_archive_sheet_does_not_block_save_or_move(self):
        """달력 • 없는 빈 3일 시트가 있어도 7일→3일 이동·저장이 된다."""
        empty3, dest = date(2026, 9, 3), date(2026, 9, 3)
        src = date(2026, 9, 7)
        self._make_archive_only(empty3, self.wt._empty_cells(empty3))
        self.assertTrue(self.wt.worklog_date_exists_in_archive(dest))
        self.assertFalse(self.wt.worklog_archive_has_saved_content(dest))
        ok, msg = self.wt.check_worklog_save_allowed(dest, had_local_at_open=False)
        self.assertTrue(ok, msg)
        self.assertEqual(msg, "")
        self.wt.save_worklog_cells(src, self._cells(src, "거래처7", "7일 내용"), force=True, allow_overwrite=True)
        with patch.object(self.wt, "_cells_from_widgets", side_effect=lambda d: self.wt.read_worklog_cells(d)):
            err = self.wt.apply_worklog_date_change(src, dest)
        self.assertEqual(err, "")
        self.assertEqual(self.wt.read_worklog_cells(dest).get("G8"), "7일 내용")
        self.assertFalse(os.path.isfile(self.wt.worklog_path(src)))

    def test_leftover_empty_day_file_does_not_block(self):
        dest = date(2026, 9, 3)
        leftover = os.path.join(self.arch, "2026", f"{dest.isoformat()}.xlsx")
        os.makedirs(os.path.dirname(leftover), exist_ok=True)
        _write_template(leftover)
        self.assertTrue(self.wt.worklog_date_exists_in_archive(dest))
        self.assertNotIn(dest.isoformat(), self.wt._list_archive_saved_dates())
        ok, msg = self.wt.check_worklog_save_allowed(dest, had_local_at_open=False)
        self.assertTrue(ok, msg)
        path = self.wt.save_worklog_cells(
            dest, self._cells(dest, "신규", "3일 신규"), force=True, allow_overwrite=False,
        )
        self.assertTrue(path)
        self.assertEqual(self.wt.read_worklog_cells(dest).get("G8"), "3일 신규")

    def test_overwrite_flag_allows_archive_only_dest(self):
        dest = date(2026, 9, 3)
        self._make_archive_only(dest, self._cells(dest, "옛3", "3일 옛내용"))
        self.assertTrue(self.wt.worklog_archive_has_saved_content(dest))
        ok, _ = self.wt.check_worklog_save_allowed(dest, had_local_at_open=False)
        self.assertFalse(ok)
        ok, msg = self.wt.check_worklog_save_allowed(dest, had_local_at_open=True)
        self.assertTrue(ok, msg)
        src = date(2026, 9, 7)
        self.wt.save_worklog_cells(src, self._cells(src, "거래처7", "7일 내용"), force=True, allow_overwrite=True)
        with patch.object(self.wt, "_cells_from_widgets", side_effect=lambda d: self.wt.read_worklog_cells(d)):
            err = self.wt.apply_worklog_date_change(src, dest)
        self.assertEqual(err, "")
        self.assertEqual(self.wt.read_worklog_cells(dest).get("G8"), "7일 내용")

    def test_real_archive_content_still_blocks_new_save(self):
        dest = date(2026, 9, 3)
        self._make_archive_only(dest, self._cells(dest, "기존", "있는 내용"))
        ok, msg = self.wt.check_worklog_save_allowed(dest, had_local_at_open=False)
        self.assertFalse(ok)
        self.assertIn("이미 있습니다", msg)
        with self.assertRaises(self.wt.WorklogSaveBlockedError):
            self.wt.save_worklog_cells(
                dest, self._cells(dest, "후입력", "막혀야 함"), force=True, allow_overwrite=False,
            )

    def test_no_calendar_dot_never_blocks(self):
        """달력에 3일 •가 없으면, 아카이브에 파일이 있어도 저장·이동을 막지 않는다."""
        dest = date(2026, 9, 3)
        leftover = os.path.join(self.arch, "2026", f"{dest.isoformat()}.xlsx")
        os.makedirs(os.path.dirname(leftover), exist_ok=True)
        _write_template(leftover)
        wb = load_workbook(leftover)
        try:
            wb.active["C8"] = "잔여"
            wb.active["G8"] = "달력에 안 보이는 잔여"
            wb.save(leftover)
        finally:
            wb.close()
        self.assertTrue(self.wt.worklog_date_exists_in_archive(dest))
        self.assertTrue(self.wt.worklog_archive_has_saved_content(dest))
        self.assertNotIn(dest.isoformat(), self.wt._saved_dates_for_calendar())
        ok, msg = self.wt.check_worklog_save_allowed(dest, had_local_at_open=False)
        self.assertTrue(ok, msg)
        src = date(2026, 9, 7)
        self.wt.save_worklog_cells(src, self._cells(src, "거래처7", "7일 내용"), force=True, allow_overwrite=True)
        with patch.object(self.wt, "_cells_from_widgets", side_effect=lambda d: self.wt.read_worklog_cells(d)):
            err = self.wt.apply_worklog_date_change(src, dest)
        self.assertEqual(err, "")
        self.assertEqual(self.wt.read_worklog_cells(dest).get("G8"), "7일 내용")

    def test_stale_saved_cache_without_dot_allows_save(self):
        dest = date(2026, 9, 3)
        self._make_archive_only(dest, self._cells(dest, "기존", "있는 내용"))
        self.ss["wl_saved_dates_cache"] = set()
        self.assertNotIn(dest.isoformat(), self.wt._saved_dates_for_calendar())
        ok, msg = self.wt.check_worklog_save_allowed(dest, had_local_at_open=False)
        self.assertTrue(ok, msg)


if __name__ == "__main__":
    unittest.main()
