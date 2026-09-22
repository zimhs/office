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

    def test_commit_save_adds_calendar_dot(self):
        """저장하면 달력 • 캐시에 그 날이 바로 들어간다."""
        d = date(2026, 9, 17)
        self.ss["wl_saved_dates_cache"] = {"2026-09-15"}
        self.wt.commit_worklog_date_save(d, d, self._cells(d, "거래처", "내용"))
        self.assertIn(d.isoformat(), self.wt._saved_dates_for_calendar())
        self.assertTrue(os.path.isfile(self.wt.worklog_path(d)))

    def test_collapse_empty_leading_page(self):
        first = self.wt._empty_cells(date(2026, 9, 17))
        second = self._cells(date(2026, 9, 17), "이엔에이치", "HE(10.2t)")
        packed = self.wt._attach_extra_pages(first, [second])
        collapsed = self.wt._collapse_empty_worklog_pages(packed)
        body, extras = self.wt._detach_extra_pages(collapsed)
        self.assertEqual(body.get("G8"), "HE(10.2t)")
        self.assertEqual(body.get("C8"), "이엔에이치")
        self.assertFalse(extras)

    def test_ui_iso_keeps_shared_editor_keys(self):
        """달력 이동 때도 입력칸 키는 같다. 위젯을 다시 붙이지 않아 로딩이 짧다."""
        self.assertEqual(self.wt._ui_iso("2026-09-07"), "ui")
        self.assertEqual(self.wt._ui_iso("2026-09-17"), "ui")
        self.assertEqual(
            self.wt._entry_lines_live_key("2026-09-07", 0),
            self.wt._entry_lines_live_key("2026-09-17", 0),
        )
        self.assertEqual(
            self.wt._entry_clients_live_key("2026-09-07", 0),
            self.wt._entry_clients_live_key("2026-09-17", 0),
        )

    def test_calendar_day_keeps_old_editor_and_loads_dest_editor(self):
        """달력으로 다른 날을 열면 전날 초안은 보관하고, 칸에는 이동일 저장본만 넣는다."""
        old, new = date(2026, 9, 17), date(2026, 9, 7)
        self.wt.save_worklog_cells(old, self._cells(old, "이엔에이치", "17일 HE"), force=True, allow_overwrite=True)
        self.wt.save_worklog_cells(new, self._cells(new, "가스코아", "7일 내근"), force=True, allow_overwrite=True)
        self.ss["worklog_selected"] = old
        self.wt._apply_entry_lines(old.isoformat(), 0, ["17일 HE"], remount_comp=True)
        self.wt._seed_entry_clients(old.isoformat(), 0, ["이엔에이치"])
        inst_before = int(self.ss.get(self.wt._entry_lines_inst_key(old.isoformat(), 0), 0) or 0)
        self.ss[self.wt._boot_key(old)] = True
        self.ss[self.wt._entries_key(old)] = [{"client": "이엔에이치", "content": "17일 HE", "lines": ["17일 HE"]}]
        self.wt._on_wl_cal_day(new.isoformat())
        self.assertEqual(self.ss["worklog_selected"], new)
        parked = self.ss.get(self.wt._parked_cells_key(old)) or {}
        self.assertIn("17일 HE", str(parked.get("G8") or ""))
        dest_live = self.ss.get(self.wt._entry_lines_live_key(new.isoformat(), 0)) or []
        self.assertIn("7일 내근", "\n".join(str(x) for x in dest_live))
        dest_cl = self.ss.get(self.wt._entry_clients_live_key(new.isoformat(), 0)) or []
        self.assertIn("가스코아", "\n".join(str(x) for x in dest_cl))
        self.assertEqual((self.ss.get(self.wt._view_cells_key(new)) or {}).get("G8"), "7일 내근")
        inst_after = int(self.ss.get(self.wt._entry_lines_inst_key(new.isoformat(), 0), 0) or 0)
        self.assertEqual(inst_after, inst_before)
        prog, replace = self.wt._program_editor_lines(new.isoformat(), 0, "ln", dest_live)
        self.assertTrue(replace)
        self.assertIn("7일 내근", "\n".join(prog or []))
        old_cs = ["채권관리", "경민산업"] + [""] * 30
        self.ss[self.wt._entry_lines_comp_key(new.isoformat(), 0)] = {"lines": old_cs}
        prog2, replace2 = self.wt._program_editor_lines(new.isoformat(), 0, "ln", dest_live)
        self.assertTrue(replace2)
        self.assertIn("7일 내근", "\n".join(prog2 or []))
        self.assertNotIn("채권관리", "\n".join(prog2 or []))
        self.ss.pop(self.wt._forced_editor_key(new.isoformat(), 0, "ln"), None)
        self.ss[self.wt._bound_editor_key(new.isoformat(), 0, "ln")] = True
        prog3, replace3 = self.wt._program_editor_lines(new.isoformat(), 0, "ln", dest_live)
        self.assertFalse(replace3)
        self.assertIsNone(prog3)

    def test_draft_date_change_opens_dest_without_copying(self):
        """저장 전 초안은 원래 날짜에 남고, 이동일은 그 날 값만 연다."""
        old, new = date(2026, 9, 7), date(2026, 9, 3)
        self.ss["worklog_selected"] = old
        self.ss[self.wt._boot_key(old)] = True
        self.ss[self.wt._entries_key(old)] = [{"client": "거래처A", "content": "초안", "lines": ["초안"]}]
        self.ss["wl_date_pick"] = new
        with patch.object(self.wt, "_cells_from_widgets", side_effect=lambda d: self._cells(d, "거래처A", "초안") if d == old else self.wt._empty_cells(d)):
            self.wt._on_wl_date_pick_change()
        self.assertEqual(self.ss["worklog_selected"], new)
        self.assertFalse(self.ss.get("wl_date_retarget_from"))
        self.assertNotIn(old.isoformat(), self.ss.get("wl_purge_dates") or [])
        self.assertFalse(os.path.isfile(self.wt.worklog_path(old)))
        self.assertEqual((self.ss.get(self.wt._entries_key(old)) or [{}])[0].get("content"), "초안")
        dest = (self.ss.get(self.wt._entries_key(new)) or [{}])[0]
        self.assertFalse(str(dest.get("content") or "").strip())

    def test_date_move_mode_retargets_draft_and_keeps_until_save(self):
        """날짜변경을 켠 뒤 날짜를 고르면 편집중 내용이 그 날로 옮겨진다. 저장 전 예전 파일은 남는다."""
        old, new = date(2026, 9, 21), date(2026, 9, 18)
        self.wt.save_worklog_cells(old, self._cells(old, "원에이엔지", "21일 초안"), force=True, allow_overwrite=True)
        self.ss["worklog_selected"] = old
        self.ss["wl_date_move_mode"] = True
        self.ss["wl_date_pick"] = new
        with patch.object(self.wt, "_cells_from_widgets", side_effect=lambda d: self._cells(d, "원에이엔지", "21일 초안")):
            self.wt._on_wl_date_pick_change()
        self.assertEqual(self.ss["worklog_selected"], new)
        self.assertFalse(self.ss.get("wl_date_move_mode"))
        self.assertEqual(self.ss.get("wl_date_retarget_from"), old.isoformat())
        self.assertIn(old.isoformat(), self.ss.get("wl_purge_dates") or [])
        self.assertTrue(os.path.isfile(self.wt.worklog_path(old)))
        dest = (self.ss.get(self.wt._entries_key(new)) or [{}])[0]
        self.assertEqual(dest.get("content"), "21일 초안")
        self.assertEqual((self.ss.get(self.wt._view_cells_key(new)) or {}).get("G8"), "21일 초안")

    def test_date_move_mode_blocks_saved_dest(self):
        old, new = date(2026, 9, 21), date(2026, 9, 18)
        self.wt.save_worklog_cells(old, self._cells(old, "A", "21일"), force=True, allow_overwrite=True)
        self.wt.save_worklog_cells(new, self._cells(new, "B", "18일 저장본"), force=True, allow_overwrite=True)
        self.ss["worklog_selected"] = old
        self.ss["wl_date_move_mode"] = True
        self.ss["wl_date_pick"] = new
        self.wt._on_wl_date_pick_change()
        self.assertEqual(self.ss["worklog_selected"], old)
        self.assertEqual(self.ss.get("wl_date_err"), self.wt._WL_MOVE_BLOCK_MSG)
        self.assertTrue(os.path.isfile(self.wt.worklog_path(old)))
        self.assertEqual(self.wt.read_worklog_cells(new).get("G8"), "18일 저장본")

    def test_save_after_saved_date_change_keeps_old_until_commit(self):
        """날짜칸만 바꾸면 예전 파일은 그대로다. 저장 커밋을 해야 이동된다."""
        old, new = date(2026, 9, 7), date(2026, 9, 3)
        self.wt.save_worklog_cells(old, self._cells(old, "거래처7", "7일 내용"), force=True, allow_overwrite=True)
        self.ss["worklog_selected"] = old
        self.ss[f"wl_saved_ok_{old.isoformat()}"] = True
        self.ss["wl_date_pick"] = new
        with patch.object(self.wt, "_cells_from_widgets", side_effect=lambda d: self.wt.read_worklog_cells(d)):
            self.wt._on_wl_date_pick_change()
        self.assertEqual(self.ss["worklog_selected"], new)
        self.assertFalse(str(((self.ss.get(self.wt._entries_key(new)) or [{}])[0]).get("content") or "").strip())
        self.assertTrue(os.path.isfile(self.wt.worklog_path(old)))
        self.wt.commit_worklog_date_save(old, new, self._cells(new, "거래처7", "7일 내용"))
        self.assertFalse(os.path.isfile(self.wt.worklog_path(old)))
        self.assertTrue(os.path.isfile(self.wt.worklog_path(new)))
        self.assertEqual(self.wt.read_worklog_cells(new).get("G8"), "7일 내용")

    def test_on_change_after_save_loads_dest_not_current(self):
        """저장한 뒤 빈 날짜로 옮기면 그 날(빈 값)을 연다. 예전 날은 파일에 남는다."""
        old, new = date(2026, 9, 7), date(2026, 9, 3)
        self.wt.save_worklog_cells(old, self._cells(old, "거래처7", "7일 내용"), force=True, allow_overwrite=True)
        self.ss["worklog_selected"] = old
        self.ss[f"wl_saved_ok_{old.isoformat()}"] = True
        self.ss["wl_date_pick"] = new
        with patch.object(self.wt, "_cells_from_widgets", side_effect=lambda d: self.wt.read_worklog_cells(d)):
            self.wt._on_wl_date_pick_change()
        self.assertEqual(self.ss["worklog_selected"], new)
        self.assertFalse(self.ss.get("wl_date_retarget_from"))
        self.assertNotIn(old.isoformat(), self.ss.get("wl_purge_dates") or [])
        self.assertFalse(self.ss.get("wl_date_err"))
        self.assertFalse(str(((self.ss.get(self.wt._entries_key(new)) or [{}])[0]).get("content") or "").strip())
        self.assertTrue(os.path.isfile(self.wt.worklog_path(old)))
        self.assertEqual(self.wt.read_worklog_cells(old).get("G8"), "7일 내용")
        self.assertTrue(self.wt.worklog_date_exists_in_archive(old))
        self.assertFalse(os.path.isfile(self.wt.worklog_path(new)))
        cal = self.wt._saved_dates_for_calendar()
        self.assertIn(old.isoformat(), cal)
        self.assertNotIn(new.isoformat(), cal)
        self.assertTrue(self.ss.get("wl_skip_sync_once"))

    def test_stored_cells_uses_archive_when_local_file_empty(self):
        dest = date(2026, 9, 3)
        self.wt.save_worklog_cells(dest, self._cells(dest, "B", "3일 옛내용"), force=True, allow_overwrite=True)
        path = self.wt.worklog_path(dest)
        self.wt.write_cells_to_path(path, dest, self.wt._empty_cells(dest), force_template=True)
        packed = self.wt._stored_cells_for_date(dest)
        self.assertEqual(packed.get("G8"), "3일 옛내용")

    def test_open_reloads_file_even_if_dest_session_has_other_day(self):
        """대상 날짜 세션에 다른 날 내용이 남아 있어도 저장본을 다시 심는다."""
        old, new = date(2026, 9, 7), date(2026, 9, 3)
        self.wt.save_worklog_cells(old, self._cells(old, "A", "7일 내용"), force=True, allow_overwrite=True)
        self.wt.save_worklog_cells(new, self._cells(new, "B", "3일 옛내용"), force=True, allow_overwrite=True)
        self.ss["worklog_selected"] = old
        self.ss[self.wt._boot_key(new)] = True
        self.ss[self.wt._entries_key(new)] = [{"client": "A", "content": "7일 내용", "lines": ["7일 내용"]}]
        self.ss[self.wt._view_cells_key(new)] = self._cells(new, "A", "7일 내용")
        self.ss["wl_date_pick"] = new
        self.wt._on_wl_date_pick_change()
        self.assertEqual(self.ss["worklog_selected"], new)
        self.assertEqual((self.ss.get(self.wt._entries_key(new)) or [{}])[0].get("content"), "3일 옛내용")
        self.assertEqual((self.ss.get(self.wt._view_cells_key(new)) or {}).get("G8"), "3일 옛내용")

    def test_date_change_onto_saved_3rd_loads_that_day(self):
        """저장된 3일로 날짜를 바꾸면 3일 저장본을 업무내용·요약에 불러온다."""
        old, new = date(2026, 9, 7), date(2026, 9, 3)
        self.wt.save_worklog_cells(old, self._cells(old, "A", "7일 내용"), force=True, allow_overwrite=True)
        self.wt.save_worklog_cells(new, self._cells(new, "B", "3일 옛내용"), force=True, allow_overwrite=True)
        self.ss["worklog_selected"] = old
        self.ss[f"wl_saved_ok_{old.isoformat()}"] = True
        self.ss["wl_date_pick"] = new
        with patch.object(self.wt, "_cells_from_widgets", side_effect=lambda d: self.wt.read_worklog_cells(d)):
            self.wt._on_wl_date_pick_change()
        self.assertEqual(self.ss["worklog_selected"], new)
        self.assertFalse(self.ss.get("wl_date_err"))
        self.assertEqual((self.ss.get(self.wt._entries_key(new)) or [{}])[0].get("content"), "3일 옛내용")
        self.assertEqual((self.ss.get(self.wt._view_cells_key(new)) or {}).get("G8"), "3일 옛내용")
        self.assertTrue(os.path.isfile(self.wt.worklog_path(old)))
        self.assertEqual(self.wt.read_worklog_cells(old).get("G8"), "7일 내용")
        self.assertEqual(self.wt.read_worklog_cells(new).get("G8"), "3일 옛내용")
        self.assertTrue(self.ss.get("wl_skip_sync_once"))
        self.assertFalse(self.ss.get("wl_need_app_rerun"))

    def test_left_date_pick_opens_dest_before_widget(self):
        old, new = date(2026, 9, 7), date(2026, 9, 3)
        self.wt.save_worklog_cells(old, self._cells(old, "거래처7", "7일 내용"), force=True, allow_overwrite=True)
        self.wt.save_worklog_cells(new, self._cells(new, "거래처3", "3일 내용"), force=True, allow_overwrite=True)
        self.ss["worklog_selected"] = old
        self.ss["wl_date_pick"] = new
        with patch.object(self.wt, "_cells_from_widgets", side_effect=lambda d: self.wt.read_worklog_cells(d)):
            got, moved = self.wt.consume_left_date_pick_move(old)
        self.assertTrue(moved)
        self.assertEqual(got, new)
        self.assertEqual(self.ss["worklog_selected"], new)
        self.assertFalse(self.ss.get("wl_date_err"))
        self.assertEqual((self.ss.get(self.wt._entries_key(new)) or [{}])[0].get("content"), "3일 내용")
        self.assertEqual((self.ss.get(self.wt._view_cells_key(new)) or {}).get("G8"), "3일 내용")
        self.assertTrue(os.path.isfile(self.wt.worklog_path(old)))
        self.assertEqual(self.wt.read_worklog_cells(old).get("G8"), "7일 내용")
        self.assertNotIn(old.isoformat(), self.ss.get("wl_purge_dates") or [])
        self.assertTrue(self.ss.get("wl_skip_sync_once"))
        self.assertFalse(self.ss.get("wl_need_app_rerun"))

    def test_skip_sync_once_does_not_call_drive(self):
        self.ss["wl_skip_sync_once"] = True
        self.ss["_wl_dash_filter_sig"] = self.wt._dashboard_top_filter_sig()
        with patch.object(self.wt, "_wl_is_streamlit_cloud", return_value=False):
            self.wt._maybe_sync_worklog_remote()
        self.assertFalse(self.ss.get("wl_skip_sync_once"))
        self.assertFalse(self.ss.get("_wl_drive_sync_ts"))

    def test_prepare_day_state_can_skip_remote_pull(self):
        dest = date(2026, 9, 3)
        with patch.object(self.wt, "_try_pull_remote_worklog_day") as pull:
            self.wt._prepare_worklog_day_state(dest, skip_remote_pull=True)
        pull.assert_not_called()

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
        self.assertEqual(err, self.wt._WL_MOVE_BLOCK_MSG)
        self.assertTrue(os.path.isfile(self.wt.worklog_path(old)))
        self.assertEqual(self.wt.read_worklog_cells(new).get("G8"), "3일 옛내용")

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

    def test_pending_runner_opens_dest_without_error(self):
        old, new = date(2026, 9, 7), date(2026, 9, 3)
        self.wt.save_worklog_cells(old, self._cells(old, "A", "7일 내용"), force=True, allow_overwrite=True)
        self.wt.save_worklog_cells(new, self._cells(new, "B", "3일 내용"), force=True, allow_overwrite=True)
        self.ss["worklog_selected"] = old
        self.ss["wl_pending_date_change"] = (old.isoformat(), new.isoformat())
        with patch.object(self.wt, "_cells_from_widgets", side_effect=lambda d: self.wt.read_worklog_cells(d)):
            self.assertTrue(self.wt._run_pending_worklog_date_change())
        self.assertEqual(self.ss["worklog_selected"], new)
        self.assertFalse(self.ss.get("wl_date_err"))
        self.assertEqual((self.ss.get(self.wt._entries_key(new)) or [{}])[0].get("content"), "3일 내용")
        self.assertTrue(os.path.isfile(self.wt.worklog_path(old)))

    def test_queue_save_flags_open_editor_date(self):
        old, new = date(2026, 9, 7), date(2026, 9, 8)
        self.ss["worklog_selected"] = old
        self.ss["wl_date_pick"] = new
        self.wt._queue_worklog_save(old.isoformat())
        self.assertTrue(self.ss[f"wl_do_save_{old.isoformat()}"])
        self.assertFalse(self.ss.get(f"wl_do_save_{new.isoformat()}"))

    def test_pending_move_does_not_auto_save_new_date(self):
        """날짜만 바꾼 뒤에는 저장 버튼을 다시 눌러야 한다."""
        old, new = date(2026, 9, 7), date(2026, 9, 8)
        cells = self._cells(old, "거래처A", "저장한 내용")
        self.wt.save_worklog_cells(old, cells, force=True, allow_overwrite=True)
        self.ss["worklog_selected"] = old
        self.ss["wl_pending_date_change"] = (old.isoformat(), new.isoformat())
        self.ss[f"wl_do_save_{old.isoformat()}"] = True
        with patch.object(self.wt, "_cells_from_widgets", side_effect=lambda d: self.wt.read_worklog_cells(d)):
            self.assertTrue(self.wt._run_pending_worklog_date_change())
        self.assertFalse(self.ss.get(f"wl_do_save_{old.isoformat()}"))
        self.assertFalse(self.ss.get(f"wl_do_save_{new.isoformat()}"))
        self.assertEqual(self.ss["worklog_selected"], new)
        self.assertFalse(self.ss.get(f"wl_saved_ok_{new.isoformat()}"))
        self.assertTrue(os.path.isfile(self.wt.worklog_path(old)))
        self.assertFalse(os.path.isfile(self.wt.worklog_path(new)))

    def test_commit_save_overwrites_existing_3rd(self):
        """날짜만 바꿔 저장하면 대상 날짜를 덮어쓰고 예전 날짜는 지운다."""
        old, new = date(2026, 9, 7), date(2026, 9, 3)
        self.wt.save_worklog_cells(old, self._cells(old, "거래처7", "7일 내용"), force=True, allow_overwrite=True)
        self.wt.save_worklog_cells(new, self._cells(new, "옛3", "3일 옛내용"), force=True, allow_overwrite=True)
        cells = self._cells(new, "거래처7", "7일 내용")
        path = self.wt.commit_worklog_date_save(old, new, cells)
        self.assertTrue(path)
        self.assertFalse(self.ss.get("wl_date_err"))
        self.assertFalse(os.path.isfile(self.wt.worklog_path(old)))
        self.assertEqual(self.wt.read_worklog_cells(new).get("G8"), "7일 내용")

    def test_date_change_before_save_does_not_copy_or_purge(self):
        """저장 전 날짜 변경은 초안을 복사하지 않고, 예전 날짜도 지우지 않는다."""
        old, new = date(2026, 9, 7), date(2026, 9, 3)
        self.ss["worklog_selected"] = old
        self.ss[self.wt._boot_key(old)] = True
        self.ss[self.wt._entries_key(old)] = [{"client": "거래처7", "content": "7일 내용", "lines": ["7일 내용"]}]
        self.ss["wl_date_pick"] = new
        with patch.object(self.wt, "_cells_from_widgets", side_effect=lambda d: self._cells(d, "거래처7", "7일 내용") if d == old else self.wt._empty_cells(d)):
            self.wt._on_wl_date_pick_change()
        self.assertEqual(self.ss["worklog_selected"], new)
        self.assertFalse(self.ss.get("wl_date_retarget_from"))
        self.assertFalse(os.path.isfile(self.wt.worklog_path(old)))
        self.assertFalse(os.path.isfile(self.wt.worklog_path(new)))
        self.assertEqual((self.ss.get(self.wt._entries_key(old)) or [{}])[0].get("content"), "7일 내용")

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

    def test_leftover_empty_does_not_block_move(self):
        src, dest = date(2026, 9, 7), date(2026, 9, 3)
        leftover = os.path.join(self.arch, "2026", f"{dest.isoformat()}.xlsx")
        os.makedirs(os.path.dirname(leftover), exist_ok=True)
        _write_template(leftover)
        self.wt.save_worklog_cells(src, self._cells(src, "거래처7", "7일 내용"), force=True, allow_overwrite=True)
        with patch.object(self.wt, "_cells_from_widgets", side_effect=lambda d: self.wt.read_worklog_cells(d)):
            err = self.wt.apply_worklog_date_change(src, dest)
        self.assertEqual(err, "")
        self.assertEqual(self.wt.read_worklog_cells(dest).get("G8"), "7일 내용")

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
        self.assertEqual(err, self.wt._WL_MOVE_BLOCK_MSG)
        self.assertEqual(self.wt.read_worklog_cells(dest).get("G8"), "3일 옛내용")

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
        """달력에 3일 •가 없어도, 저장된 내용이 있으면 자료를 옮길 수 없다."""
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
        src = date(2026, 9, 7)
        self.wt.save_worklog_cells(src, self._cells(src, "거래처7", "7일 내용"), force=True, allow_overwrite=True)
        with patch.object(self.wt, "_cells_from_widgets", side_effect=lambda d: self.wt.read_worklog_cells(d)):
            err = self.wt.apply_worklog_date_change(src, dest)
        self.assertEqual(err, self.wt._WL_MOVE_BLOCK_MSG)
        self.assertTrue(os.path.isfile(self.wt.worklog_path(src)))
        self.assertTrue(self.wt.worklog_archive_has_saved_content(dest))

    def test_stale_saved_cache_without_dot_allows_save(self):
        dest = date(2026, 9, 3)
        self._make_archive_only(dest, self._cells(dest, "기존", "있는 내용"))
        self.ss["wl_saved_dates_cache"] = set()
        self.assertNotIn(dest.isoformat(), self.wt._saved_dates_for_calendar())
        ok, msg = self.wt.check_worklog_save_allowed(dest, had_local_at_open=False)
        self.assertTrue(ok, msg)

    def test_save_purges_old_even_if_selected_already_new(self):
        """화면 날짜가 이미 3일이어도 purge 목록의 7일은 저장 때 삭제된다."""
        old, new = date(2026, 9, 7), date(2026, 9, 3)
        self.wt.save_worklog_cells(old, self._cells(old, "거래처7", "7일 내용"), force=True, allow_overwrite=True)
        self.ss["worklog_selected"] = new
        self.ss["wl_purge_dates"] = [old.isoformat()]
        path = self.wt.commit_worklog_date_save(new, new, self._cells(new, "거래처7", "7일 내용"))
        self.assertTrue(path)
        self.assertFalse(os.path.isfile(self.wt.worklog_path(old)))
        self.assertFalse(self.wt.worklog_date_exists_in_archive(old))
        self.assertTrue(os.path.isfile(self.wt.worklog_path(new)))
        self.assertEqual(self.wt.read_worklog_cells(new).get("G8"), "7일 내용")

    def test_calendar_opens_saved_day_without_copying(self):
        """달력에서 다른 날을 누르면 그 날 저장본을 불러오고, 지금 입력은 복사·저장하지 않는다."""
        old, new = date(2026, 9, 7), date(2026, 9, 3)
        self.wt.save_worklog_cells(old, self._cells(old, "거래처7", "7일 내용"), force=True, allow_overwrite=True)
        self.wt.save_worklog_cells(new, self._cells(new, "거래처3", "3일 내용"), force=True, allow_overwrite=True)
        self.ss["worklog_selected"] = new
        self.ss[self.wt._boot_key(new)] = True
        self.ss[self.wt._entries_key(new)] = [{"client": "거래처3", "content": "3일 내용", "lines": ["3일 내용"]}]
        self.ss[f"wl_do_save_{new.isoformat()}"] = True
        with patch.object(self.wt, "_cells_from_widgets", side_effect=lambda d: self.wt.read_worklog_cells(d)):
            self.wt._on_wl_cal_day(old.isoformat())
        self.assertEqual(self.ss["worklog_selected"], old)
        self.assertTrue(os.path.isfile(self.wt.worklog_path(new)))
        self.assertEqual(self.wt.read_worklog_cells(new).get("G8"), "3일 내용")
        self.assertEqual((self.ss.get(self.wt._entries_key(new)) or [{}])[0].get("content"), "3일 내용")
        self.assertEqual((self.ss.get(self.wt._entries_key(old)) or [{}])[0].get("content"), "7일 내용")
        self.assertEqual((self.ss.get(self.wt._view_cells_key(old)) or {}).get("G8"), "7일 내용")
        self.assertFalse(self.ss.get("wl_date_retarget_from"))
        self.assertFalse(self.ss.get(f"wl_do_save_{old.isoformat()}"))
        self.assertFalse(self.ss.get(f"wl_do_save_{new.isoformat()}"))

    def test_calendar_empty_day_opens_without_saving(self):
        """달력의 빈 날짜는 열기만 한다. 저장 파일이 생기지 않는다."""
        old, new = date(2026, 9, 7), date(2026, 9, 16)
        self.wt.save_worklog_cells(old, self._cells(old, "거래처7", "7일 내용"), force=True, allow_overwrite=True)
        self.ss["worklog_selected"] = old
        self.ss[self.wt._boot_key(old)] = True
        self.ss[self.wt._entries_key(old)] = [{"client": "거래처7", "content": "7일 내용", "lines": ["7일 내용"]}]
        self.ss[f"wl_do_save_{old.isoformat()}"] = True
        with patch.object(self.wt, "_cells_from_widgets", side_effect=lambda d: self.wt.read_worklog_cells(d)):
            self.wt._on_wl_cal_day(new.isoformat())
        self.assertEqual(self.ss["worklog_selected"], new)
        self.assertTrue(os.path.isfile(self.wt.worklog_path(old)))
        self.assertEqual(self.wt.read_worklog_cells(old).get("G8"), "7일 내용")
        self.assertEqual((self.ss.get(self.wt._entries_key(old)) or [{}])[0].get("content"), "7일 내용")
        self.assertFalse(self.ss.get("wl_date_retarget_from"))
        self.assertNotIn(old.isoformat(), self.ss.get("wl_purge_dates") or [])
        self.assertFalse(os.path.isfile(self.wt.worklog_path(new)))
        self.assertFalse(self.ss.get(f"wl_do_save_{new.isoformat()}"))

    def test_confirm_delete_closes_popover_on_flush(self):
        """확정 클릭 후 위젯을 그리기 전에 팝업 인스턴스를 올려 닫는다."""
        d = date(2026, 9, 3)
        self.ss["worklog_selected"] = d
        self.ss["wl_del_day_inst"] = 0
        self.ss["wl_del_day_open_0"] = True
        self.wt._on_confirm_delete_day()
        self.assertEqual(self.ss.get("wl_do_delete_day"), d.isoformat())
        self.assertGreaterEqual(int(self.ss.get("wl_scroll_pin") or 0), 1)
        self.assertTrue(self.ss.get("wl_del_day_force_close"))
        self.assertEqual(self.wt._worklog_delete_popover_key(), "wl_del_day_open_0")
        self.wt._flush_worklog_delete_popover()
        self.assertFalse(self.ss.get("wl_del_day_force_close"))
        self.assertEqual(self.ss.get("wl_del_day_inst"), 1)
        self.assertEqual(self.wt._worklog_delete_popover_key(), "wl_del_day_open_1")
        self.assertFalse(self.ss.get("wl_del_day_open_1"))
        self.assertIs(self.ss.get("wl_del_day_open_0"), False)

    def test_delete_old_date_does_not_wipe_moved_date(self):
        """예전 날짜 삭제 버튼은 그 날만 지우고, 옮긴 날짜 데이터는 남긴다."""
        old, new = date(2026, 9, 7), date(2026, 9, 3)
        self.wt.save_worklog_cells(old, self._cells(old, "잔여7", "남은 7일"), force=True, allow_overwrite=True)
        self.wt.save_worklog_cells(new, self._cells(new, "거래처3", "옮긴 내용"), force=True, allow_overwrite=True)
        self.ss["worklog_selected"] = new
        self.ss[self.wt._boot_key(new)] = True
        self.ss[self.wt._entries_key(new)] = [{"client": "거래처3", "content": "옮긴 내용", "lines": ["옮긴 내용"]}]
        with patch.object(self.wt, "_cells_from_widgets", side_effect=lambda d: self.wt.read_worklog_cells(d)):
            self.wt._on_wl_cal_day(old.isoformat())
        self.ss["wl_do_delete_day"] = old.isoformat()
        self.assertTrue(self.wt._run_pending_worklog_day_delete())
        self.assertFalse(os.path.isfile(self.wt.worklog_path(old)))
        self.assertFalse(self.wt.worklog_date_exists_in_archive(old))
        self.assertTrue(os.path.isfile(self.wt.worklog_path(new)))
        self.assertTrue(self.wt.worklog_date_exists_in_archive(new))
        self.assertEqual(self.wt.read_worklog_cells(new).get("G8"), "옮긴 내용")

    def test_drop_saved_date_from_cache_keeps_others(self):
        """확정 삭제 후 달력 캐시는 그 날짜만 빼고 월별 파일을 다시 읽지 않는다."""
        self.ss["wl_saved_dates_cache"] = {"2026-09-03", "2026-09-07"}
        self.wt._drop_saved_date_from_cache("2026-09-07")
        self.assertEqual(self.ss["wl_saved_dates_cache"], {"2026-09-03"})

    def test_local_purge_does_not_push_month_archive_to_drive(self):
        """로컬 확정 삭제는 Drive 월별 복사를 UI 스레드에서 하지 않는다."""
        d = date(2026, 9, 3)
        self.wt.save_worklog_cells(d, self._cells(d, "거래처3", "지울 내용"), force=True, allow_overwrite=True)
        self.ss["wl_saved_dates_cache"] = {d.isoformat(), "2026-09-07"}
        with patch("drive_autoload.push_worklog_month_archive_to_drive") as push:
            removed = self.wt.purge_worklog_day_files(d, remote=False)
        push.assert_not_called()
        self.assertTrue(removed)
        self.assertFalse(os.path.isfile(self.wt.worklog_path(d)))
        self.assertEqual(self.ss.get("wl_saved_dates_cache"), {"2026-09-07"})


if __name__ == "__main__":
    unittest.main()
