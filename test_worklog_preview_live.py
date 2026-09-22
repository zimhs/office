"""왼쪽 업무일지 미리보기 — 입력 중 iframe 재부착 없이 칸만 패치."""
from __future__ import annotations

import os
import re
import tempfile
import unittest
from datetime import date

from openpyxl import Workbook


class WorklogPreviewLiveTest(unittest.TestCase):
    def setUp(self):
        import worklog_tab as wt

        self.wt = wt

    def test_preview_cell_html_matches_excel_escape(self):
        self.assertEqual(self.wt._wl_preview_cell_html("G8", "상세하러간다"), "상세하러간다")
        self.assertEqual(self.wt._wl_preview_cell_html("G8", "A B"), "A&nbsp;B")
        self.assertEqual(self.wt._wl_preview_cell_html("G8", "A\nB"), "A<br>B")
        self.assertEqual(self.wt._wl_preview_cell_html("G8", "<x>"), "&lt;x&gt;")
        self.assertEqual(self.wt._wl_preview_cell_html("D40", ""), "&nbsp;")
        self.assertEqual(self.wt._wl_preview_cell_html("D40", self.wt._WL_SOFT_BLANK), "&nbsp;")
        self.assertEqual(self.wt._wl_preview_cell_html("G8", "   "), "")

    def test_preview_patches_cover_date_and_body_cells(self):
        d = date(2026, 9, 4)
        cells = self.wt._empty_cells(d)
        cells["G8"] = "상세하러간다"
        cells["C8"] = "거래처"
        cells["Y8"] = "비고"
        cells["D40"] = "내일"
        patches = self.wt._wl_preview_patches(cells)
        self.assertEqual(patches["G8"], "상세하러간다")
        self.assertEqual(patches["C8"], "거래처")
        self.assertEqual(patches["Y8"], "비고")
        self.assertEqual(patches["D40"], "내일")
        self.assertIn("date", patches)
        self.assertIn("G39", patches)
        self.assertIn("C39", patches)
        self.assertIn("Y39", patches)
        self.assertEqual(patches["D41"], "&nbsp;")

    def test_workbook_html_marks_live_cells(self):
        tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        tmp.close()
        try:
            wb = Workbook()
            ws = wb.active
            ws["C5"] = "2026-09-04 (금)"
            ws["C8"] = "거래처"
            ws["G8"] = "상세하러간다"
            ws["Y8"] = "비고"
            ws["D40"] = "내일"
            wb.save(tmp.name)
            wb.close()
            html = self.wt.workbook_to_html(tmp.name, include_logo=False)
            self.assertIn('data-wl="date"', html)
            self.assertIn('data-wl-page="1"', html)
            self.assertIn('data-wl="C8"', html)
            self.assertIn('data-wl="G8"', html)
            self.assertIn('data-wl="Y8"', html)
            self.assertIn('data-wl="D40"', html)
            self.assertIn("비고", html)
            self.assertIn("상세하러간다", html)
            self.assertIn("word-break:break-all", html)
            self.assertIn("white-space:pre-wrap", html)
            self.assertIn("white-space:nowrap", html)
        finally:
            os.unlink(tmp.name)

    def test_workbook_html_marks_extra_page(self):
        tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        tmp.close()
        try:
            wb = Workbook()
            wb.active["G8"] = "1페이지"
            ws2 = wb.create_sheet("p2")
            ws2["G8"] = "2페이지"
            wb.save(tmp.name)
            wb.close()
            html = self.wt.workbook_to_html(tmp.name, include_logo=False)
            self.assertIn('data-wl-page="1"', html)
            self.assertIn('data-wl-page="2"', html)
            self.assertIn('data-wl="p2-G8"', html)
            self.assertIn("2페이지", html)
        finally:
            os.unlink(tmp.name)

    def test_excel_host_html_is_fragment_not_document(self):
        tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        tmp.close()
        try:
            wb = Workbook()
            wb.save(tmp.name)
            wb.close()
            html = self.wt._excel_preview_host_html(tmp.name, scale=0.65)
            self.assertNotIn("<!DOCTYPE html>", html)
            self.assertIn("wl-sheet", html)
            self.assertIn("sheet-scale", html)
            self.assertIn("data-wl=", html)
            self.assertIn("wl-preview-page", html)
            self.assertIn("word-break:break-all", html)
            self.assertIn("white-space:pre-wrap", html)
            self.assertIn("white-space:nowrap", html)
            self.assertIn("td[data-wl^='Y']", html)
        finally:
            os.unlink(tmp.name)


    def test_missing_remark_merge_still_spans_y_ab(self):
        tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        tmp.close()
        try:
            wb = Workbook()
            ws = wb.active
            ws.merge_cells("Y8:AB8")
            ws["Y12"] = "카타파타"
            wb.save(tmp.name)
            wb.close()
            html = self.wt.workbook_to_html(tmp.name, include_logo=False)
            self.assertIn('data-wl="Y12"', html)
            self.assertIn("카타파타", html)
            m = re.search(r"<td([^>]*data-wl=\"Y12\"[^>]*)>", html)
            self.assertIsNotNone(m)
            attrs = m.group(1)
            self.assertIn('colspan="4"', attrs)
            self.assertIn("white-space:nowrap", attrs)
            self.assertNotIn("word-break:break-all", attrs)
            host = self.wt._excel_preview_host_html(tmp.name, scale=0.65)
            y_css = host.split("td[data-wl^='Y']", 1)[-1][:400]
            self.assertIn("white-space:nowrap", y_css)
            self.assertNotIn("word-break:break-all", y_css.split("}" , 1)[0])
        finally:
            os.unlink(tmp.name)


if __name__ == "__main__":
    unittest.main()
