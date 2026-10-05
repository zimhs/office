"""월별 일지(…/일지/2026/9월.xlsx) 시트 저장·삭제·연월 자동생성."""
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


def _write_day_xlsx(path: str, label: str) -> None:
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    wb = Workbook()
    ws = wb.active
    # 본문 칸(C8/G8) — 달력 • / draft 판정이 이 행을 본다
    ws["C8"] = label
    ws["G8"] = label
    wb.save(path)
    wb.close()


class WorklogMonthArchiveTest(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.root = os.path.join(self._tmp.name, "Desktop", "업무", "일지")
        os.makedirs(self.root, exist_ok=True)
        self.ss = _FakeSS()
        import worklog_tab as wt

        self.wt = wt
        self._p_root = patch.object(wt, "resolve_worklog_archive_root", return_value=self.root)
        self._p_ss = patch.object(wt.st, "session_state", self.ss)
        self._p_root.start()
        self._p_ss.start()

    def tearDown(self):
        self._p_ss.stop()
        self._p_root.stop()
        self._tmp.cleanup()

    def test_save_creates_year_and_month_file_with_day_sheet(self):
        d = date(2026, 9, 3)
        day_xlsx = os.path.join(self._tmp.name, "2026-09-03.xlsx")
        _write_day_xlsx(day_xlsx, "day-3")

        year_dir = os.path.join(self.root, "2026")
        month_path = os.path.join(year_dir, "9월.xlsx")
        self.assertFalse(os.path.isdir(year_dir))
        self.assertFalse(os.path.exists(month_path))

        got = self.wt.upsert_worklog_archive_sheet(d, day_xlsx)
        self.assertEqual(got, month_path)
        self.assertTrue(os.path.isdir(year_dir))
        self.assertTrue(os.path.isfile(month_path))

        wb = load_workbook(month_path, read_only=True)
        try:
            self.assertIn("3", wb.sheetnames)
            self.assertEqual(wb["3"]["C8"].value, "day-3")
        finally:
            wb.close()

    def test_second_day_adds_another_sheet_same_month_file(self):
        d3 = date(2026, 9, 3)
        d5 = date(2026, 9, 5)
        p3 = os.path.join(self._tmp.name, "2026-09-03.xlsx")
        p5 = os.path.join(self._tmp.name, "2026-09-05.xlsx")
        _write_day_xlsx(p3, "day-3")
        _write_day_xlsx(p5, "day-5")
        self.wt.upsert_worklog_archive_sheet(d3, p3)
        self.wt.upsert_worklog_archive_sheet(d5, p5)

        month_path = os.path.join(self.root, "2026", "9월.xlsx")
        wb = load_workbook(month_path, read_only=True)
        try:
            self.assertEqual(sorted(wb.sheetnames, key=lambda n: int(n) if n.isdigit() else n), ["3", "5"])
        finally:
            wb.close()

    def test_delete_removes_sheet_keeps_other_days(self):
        d3 = date(2026, 9, 3)
        d5 = date(2026, 9, 5)
        p3 = os.path.join(self._tmp.name, "2026-09-03.xlsx")
        p5 = os.path.join(self._tmp.name, "2026-09-05.xlsx")
        _write_day_xlsx(p3, "day-3")
        _write_day_xlsx(p5, "day-5")
        self.wt.upsert_worklog_archive_sheet(d3, p3)
        self.wt.upsert_worklog_archive_sheet(d5, p5)

        removed = self.wt.delete_worklog_archive_sheet(d3)
        month_path = os.path.join(self.root, "2026", "9월.xlsx")
        self.assertEqual(removed, month_path)
        self.assertTrue(os.path.isfile(month_path))
        wb = load_workbook(month_path, read_only=True)
        try:
            self.assertNotIn("3", wb.sheetnames)
            self.assertIn("5", wb.sheetnames)
        finally:
            wb.close()

    def test_delete_last_sheet_removes_month_file(self):
        d = date(2026, 9, 3)
        p = os.path.join(self._tmp.name, "2026-09-03.xlsx")
        _write_day_xlsx(p, "day-3")
        self.wt.upsert_worklog_archive_sheet(d, p)
        month_path = os.path.join(self.root, "2026", "9월.xlsx")
        self.assertTrue(os.path.isfile(month_path))
        self.wt.delete_worklog_archive_sheet(d)
        self.assertFalse(os.path.isfile(month_path))

    def test_new_month_creates_separate_file(self):
        d9 = date(2026, 9, 30)
        d10 = date(2026, 10, 1)
        p9 = os.path.join(self._tmp.name, "2026-09-30.xlsx")
        p10 = os.path.join(self._tmp.name, "2026-10-01.xlsx")
        _write_day_xlsx(p9, "sep")
        _write_day_xlsx(p10, "oct")
        self.wt.upsert_worklog_archive_sheet(d9, p9)
        self.wt.upsert_worklog_archive_sheet(d10, p10)
        self.assertTrue(os.path.isfile(os.path.join(self.root, "2026", "9월.xlsx")))
        self.assertTrue(os.path.isfile(os.path.join(self.root, "2026", "10월.xlsx")))

    def test_exists_lookup_does_not_create_year_folder(self):
        d = date(2027, 1, 1)
        self.assertFalse(self.wt.worklog_date_exists_in_archive(d))
        self.assertFalse(os.path.isdir(os.path.join(self.root, "2027")))

    def test_create_worklog_day_local_makes_sheet_without_overwrite(self):
        d = date(2026, 9, 4)
        tpl = os.path.join(self._tmp.name, "template.xlsx")
        cache = os.path.join(self._tmp.name, "cache")
        os.makedirs(cache, exist_ok=True)
        # 빈 템플릿 — 본문 칸이 없어야 달력 • 가 안 뜬다
        wb_tpl = Workbook()
        wb_tpl.save(tpl)
        wb_tpl.close()
        with patch.object(self.wt, "WORKLOG_TEMPLATE", tpl), patch.object(self.wt, "WORKLOG_DIR", cache):
            info = self.wt.create_worklog_day_local(d)
            self.assertTrue(info["created"])
            month_path = os.path.join(self.root, "2026", "9월.xlsx")
            self.assertTrue(os.path.isfile(month_path))
            self.assertTrue(os.path.isfile(os.path.join(cache, "2026-09-04.xlsx")))
            again = self.wt.create_worklog_day_local(d)
            self.assertFalse(again["created"])
            self.ss.pop("wl_saved_dates_cache", None)
            dates = self.wt.list_saved_worklog_dates()
            self.assertNotIn("2026-09-04", dates)
            # 내용 저장 후에만 •
            day_path = os.path.join(cache, "2026-09-04.xlsx")
            _write_day_xlsx(day_path, "saved-4")
            self.wt.upsert_worklog_archive_sheet(d, day_path, allow_overwrite=True)
            self.ss.pop("wl_saved_dates_cache", None)
            dates2 = self.wt.list_saved_worklog_dates()
            self.assertIn("2026-09-04", dates2)
        wb = load_workbook(os.path.join(self.root, "2026", "9월.xlsx"), read_only=True)
        try:
            self.assertIn("4", wb.sheetnames)
        finally:
            wb.close()

    def test_padded_day_sheets_normalize_and_show_on_calendar(self):
        """옛 03·04 시트 → 정규 3·4 로 합치고 내용 있는 날만 •."""
        month_dir = os.path.join(self.root, "2026")
        os.makedirs(month_dir, exist_ok=True)
        month_path = os.path.join(month_dir, "10월.xlsx")
        wb = Workbook()
        ws1 = wb.active
        ws1.title = "1"
        ws1["C8"] = "oct1-new"
        ws1["G8"] = "oct1-new"
        for day, label in (("03", "old-3"), ("04", "old-4"), ("05", "old-5"), ("07", "old-7"), ("10", "old-10")):
            ws = wb.create_sheet(day)
            ws["C8"] = label
            ws["G8"] = label
        # 빈 유령 시트는 • 제외
        wb.create_sheet("06")
        wb.save(month_path)
        wb.close()

        self.ss.pop("wl_saved_dates_cache", None)
        dates = self.wt.list_saved_worklog_dates()
        self.assertIn("2026-10-01", dates)
        self.assertIn("2026-10-03", dates)
        self.assertIn("2026-10-04", dates)
        self.assertIn("2026-10-05", dates)
        self.assertIn("2026-10-07", dates)
        self.assertIn("2026-10-10", dates)
        self.assertNotIn("2026-10-06", dates)

        wb2 = load_workbook(month_path)
        try:
            self.assertIn("3", wb2.sheetnames)
            self.assertNotIn("03", wb2.sheetnames)
            self.assertIn("4", wb2.sheetnames)
            self.assertNotIn("04", wb2.sheetnames)
            cells = self.wt.read_worklog_cells_from_archive(date(2026, 10, 3))
            self.assertIsNotNone(cells)
            self.assertIn("old-3", str(cells.get("C8") or cells.get("G8") or ""))
        finally:
            wb2.close()

    def test_try_pull_prefers_archive_over_drive_day(self):
        """로컬 일자 파일이 없을 때 월별 시트를 Drive 옛 파일보다 먼저 복원한다."""
        d = date(2026, 10, 1)
        cache = os.path.join(self._tmp.name, "cache2")
        os.makedirs(cache, exist_ok=True)
        day_src = os.path.join(self._tmp.name, "seed-10-01.xlsx")
        _write_day_xlsx(day_src, "from-archive")
        self.wt.upsert_worklog_archive_sheet(d, day_src, allow_overwrite=True)
        local_day = os.path.join(cache, "2026-10-01.xlsx")
        self.assertFalse(os.path.isfile(local_day))

        with patch.object(self.wt, "WORKLOG_DIR", cache), patch.object(
            self.wt, "_wl_is_streamlit_cloud", return_value=False
        ), patch.object(self.wt, "_worklog_day_marked_deleted", return_value=False), patch(
            "drive_autoload.sync_worklog_bidirectional"
        ) as sync_mock:
            ok = self.wt._try_pull_remote_worklog_day(d)
            self.assertTrue(ok)
            sync_mock.assert_not_called()
        self.assertTrue(os.path.isfile(local_day))
        cells = self.wt.read_worklog_cells(d)
        self.assertTrue(self.wt._worklog_cells_have_draft(cells))
        self.assertIn("from-archive", str(cells.get("C8") or "") + str(cells.get("G8") or ""))

    def test_wrong_year_iso_sheet_does_not_mark_calendar_or_load(self):
        """2026/2월.xlsx 안의 2024-02-07 시트는 2026-02-07로 취급하지 않는다."""
        year_dir = os.path.join(self.root, "2026")
        os.makedirs(year_dir, exist_ok=True)
        month_path = os.path.join(year_dir, "2월.xlsx")
        wb = Workbook()
        ws = wb.active
        ws.title = "2024-02-07"
        ws["C5"] = "2024-02-07 (금)"
        ws["C8"] = "from-2024"
        ws["G8"] = "from-2024"
        wb.save(month_path)
        wb.close()

        self.ss.pop("wl_saved_dates_cache", None)
        dates = self.wt.list_saved_worklog_dates()
        self.assertNotIn("2026-02-07", dates)
        self.assertNotIn("2024-02-07", dates)

        self.assertIsNone(self.wt.read_worklog_cells_from_archive(date(2026, 2, 7)))
        # 정규화해도 다른 연 ISO는 일 시트로 합치지 않는다
        wb2 = load_workbook(month_path)
        try:
            self.assertIn("2024-02-07", wb2.sheetnames)
            self.assertNotIn("7", wb2.sheetnames)
        finally:
            wb2.close()

    def test_sheet_with_mismatched_c5_date_skipped_for_year(self):
        """시트명 7 이어도 C5가 2024면 2026-02-07 달력·불러오기에서 제외."""
        year_dir = os.path.join(self.root, "2026")
        os.makedirs(year_dir, exist_ok=True)
        month_path = os.path.join(year_dir, "2월.xlsx")
        wb = Workbook()
        ws = wb.active
        ws.title = "7"
        ws["C5"] = "2024-02-07 (수)"
        ws["C8"] = "ghost-2024"
        ws["G8"] = "ghost-2024"
        wb.save(month_path)
        wb.close()

        self.ss.pop("wl_saved_dates_cache", None)
        dates = self.wt.list_saved_worklog_dates()
        self.assertNotIn("2026-02-07", dates)
        self.assertIsNone(self.wt.read_worklog_cells_from_archive(date(2026, 2, 7)))

    def test_save_writes_selected_year_month_sheet(self):
        """저장은 선택한 날짜의 연도 폴더·월 파일·일 시트에 쓴다."""
        d = date(2026, 2, 7)
        day_xlsx = os.path.join(self._tmp.name, "day.xlsx")
        _write_day_xlsx(day_xlsx, "feb7-2026")
        # C5도 지정일에 맞춤
        wb = load_workbook(day_xlsx)
        wb.active["C5"] = self.wt.format_worklog_date(d)
        wb.save(day_xlsx)
        wb.close()

        got = self.wt.upsert_worklog_archive_sheet(d, day_xlsx, allow_overwrite=True)
        expect = os.path.join(self.root, "2026", "2월.xlsx")
        self.assertEqual(got, expect)
        self.assertTrue(os.path.isfile(expect))
        self.assertFalse(os.path.exists(os.path.join(self.root, "2024", "2월.xlsx")))

        cells = self.wt.read_worklog_cells_from_archive(d)
        self.assertIsNotNone(cells)
        self.assertIn("feb7-2026", str(cells.get("C8") or ""))
        self.assertTrue(str(cells.get("date") or "").startswith("2026-02-07"))

        self.ss.pop("wl_saved_dates_cache", None)
        self.assertIn("2026-02-07", self.wt.list_saved_worklog_dates())

    def test_openpyxl_duplicate_sheet_101_normalized_on_load(self):
        """openpyxl 충돌명 101 + 낡은 10 → 정규 10으로 합치고 새 값을 불러온다."""
        month_dir = os.path.join(self.root, "2026")
        os.makedirs(month_dir, exist_ok=True)
        month_path = os.path.join(month_dir, "10월.xlsx")
        wb = Workbook()
        ws_old = wb.active
        ws_old.title = "10"
        ws_old["C5"] = self.wt.format_worklog_date(date(2026, 10, 10))
        ws_old["C8"] = "STALE-OLD"
        ws_old["G8"] = "STALE-OLD"
        ws_new = wb.create_sheet("101")
        ws_new["C5"] = self.wt.format_worklog_date(date(2026, 10, 10))
        ws_new["C8"] = "CORRECT-NEW"
        ws_new["G8"] = "CORRECT-NEW"
        for day, lab in ((3, "d3"), (4, "d4"), (5, "d5"), (7, "d7")):
            ws = wb.create_sheet(str(day))
            ws["C8"] = lab
            ws["G8"] = lab
        wb.save(month_path)
        wb.close()

        cells = self.wt.read_worklog_cells_from_archive(date(2026, 10, 10))
        self.assertIsNotNone(cells)
        self.assertIn("CORRECT-NEW", str(cells.get("C8") or "") + str(cells.get("G8") or ""))
        self.assertNotIn("STALE-OLD", str(cells.get("C8") or "") + str(cells.get("G8") or ""))

        wb2 = load_workbook(month_path)
        try:
            self.assertIn("10", wb2.sheetnames)
            self.assertNotIn("101", wb2.sheetnames)
            self.assertEqual(wb2["10"]["C8"].value, "CORRECT-NEW")
            self.assertEqual(
                sorted((n for n in wb2.sheetnames if n.isdigit()), key=int),
                ["3", "4", "5", "7", "10"],
            )
        finally:
            wb2.close()

    def test_first_month_upsert_drops_leftover_day_sheet_not_101(self):
        """일자 파일에 Sheet+잔여 '10'이 있어도 월 파일은 '10'만 만들고 활성 내용을 담는다."""
        day_xlsx = os.path.join(self._tmp.name, "day-10-dup.xlsx")
        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet"
        ws["C5"] = self.wt.format_worklog_date(date(2026, 10, 10))
        ws["C8"] = "CORRECT-NEW"
        ws["G8"] = "CORRECT-NEW"
        leftover = wb.create_sheet("10")
        leftover["C5"] = self.wt.format_worklog_date(date(2026, 10, 10))
        leftover["C8"] = "STALE-OLD"
        leftover["G8"] = "STALE-OLD"
        wb.save(day_xlsx)
        wb.close()

        got = self.wt.upsert_worklog_archive_sheet(date(2026, 10, 10), day_xlsx)
        month_path = os.path.join(self.root, "2026", "10월.xlsx")
        self.assertEqual(got, month_path)

        mwb = load_workbook(month_path)
        try:
            self.assertEqual(mwb.sheetnames, ["10"])
            self.assertEqual(mwb["10"]["C8"].value, "CORRECT-NEW")
            self.assertNotIn("101", mwb.sheetnames)
        finally:
            mwb.close()

        cells = self.wt.read_worklog_cells_from_archive(date(2026, 10, 10))
        self.assertIsNotNone(cells)
        self.assertIn("CORRECT-NEW", str(cells.get("C8") or ""))

    def test_openpyxl_duplicate_base_day_helper(self):
        self.assertEqual(self.wt._openpyxl_duplicate_base_day("101"), 10)
        self.assertEqual(self.wt._openpyxl_duplicate_base_day("102"), 10)
        self.assertEqual(self.wt._openpyxl_duplicate_base_day("32"), 3)
        self.assertIsNone(self.wt._openpyxl_duplicate_base_day("10"))
        self.assertIsNone(self.wt._openpyxl_duplicate_base_day("11"))
        self.assertIsNone(self.wt._openpyxl_duplicate_base_day("10p2"))

    def test_archive_root_prefers_local_desktop_over_other_computers(self):
        home = os.path.join(self._tmp.name, "home")
        local = os.path.join(home, "Desktop", "업무", "일지")
        other = os.path.join(
            home,
            "Library",
            "CloudStorage",
            "GoogleDrive-x",
            "다른 컴퓨터",
            "내 컴퓨터 (1)",
            "Desktop",
            "업무",
            "일지",
        )
        os.makedirs(local, exist_ok=True)
        os.makedirs(other, exist_ok=True)
        os.makedirs(os.path.join(other, "2026"), exist_ok=True)
        with open(os.path.join(other, "2026", "3월.xlsx"), "wb") as f:
            f.write(b"PK")
        # setUp 의 resolve mock 을 잠시 끄고 실제 우선순위를 검사한다
        self._p_root.stop()
        try:
            with patch.object(self.wt.os.path, "expanduser", side_effect=lambda p: home if p == "~" else p), patch.object(
                self.wt,
                "_iter_google_drive_roots",
                return_value=[os.path.join(home, "Library", "CloudStorage", "GoogleDrive-x")],
            ):
                self.ss.pop("_wl_archive_root_cache", None)
                got = self.wt.resolve_worklog_archive_root()
            self.assertEqual(got, local)
        finally:
            self._p_root.start()


if __name__ == "__main__":
    unittest.main()
