"""업무입력 상단 엑셀 서식(채우기·글자색·정렬)."""
from __future__ import annotations

import os
import tempfile
import unittest
from unittest.mock import patch


class WorklogFormatBarTest(unittest.TestCase):
    def setUp(self):
        import worklog_tab as wt

        self.wt = wt
        self.td = tempfile.TemporaryDirectory()
        self.folder = os.path.join(self.td.name, "worklog")
        os.makedirs(self.folder, exist_ok=True)
        self.p_dir = patch.object(wt, "WORKLOG_DIR", self.folder)
        self.p_dir.start()

    def tearDown(self):
        self.p_dir.stop()
        self.td.cleanup()

    def test_set_fmt_roundtrip_and_payload(self):
        iso = "2026-11-06"
        ss = {}
        with patch.object(self.wt.st, "session_state", ss):
            self.wt._wl_set_fmt(iso, 0, "ln", 3, fill=self.wt._WL_FMT_FILL_YELLOW, color=self.wt._WL_FMT_COLOR_RED, align="center")
            payload = self.wt._wl_fmt_styles_payload(iso, 0, "ln", 5)
            self.assertEqual(payload[3]["fill"], self.wt._WL_FMT_FILL_YELLOW)
            self.assertEqual(payload[3]["color"], self.wt._WL_FMT_COLOR_RED)
            self.assertEqual(payload[3]["align"], "center")
            self.assertEqual(payload[0], {})
            self.wt._wl_set_fmt(iso, 0, "ln", 3, fill="", color="")
            payload = self.wt._wl_fmt_styles_payload(iso, 0, "ln", 5)
            self.assertNotIn("fill", payload[3])
            self.assertNotIn("color", payload[3])
            self.assertEqual(payload[3].get("align"), "center")
        rows = self.wt._wl_read_fmt(iso)
        self.assertEqual(rows["0:ln:3"]["fill"], "none")
        self.assertEqual(rows["0:ln:3"]["align"], "center")

    def test_focus_parse_and_toolbar_in_source(self):
        ss = {"wl_active_cell_key": "wl_ent_ln_2026-11-06_0_4"}
        with patch.object(self.wt.st, "session_state", ss):
            kind, page, j = self.wt._wl_fmt_focus_cell("2026-11-06", 0)
        self.assertEqual((kind, page, j), ("ln", 0, 4))
        with open(self.wt.__file__, encoding="utf-8") as f:
            src = f.read()
        self.assertIn("_render_worklog_format_bar", src)
        self.assertIn("노랑", src)
        self.assertIn("빨강글씨", src)


if __name__ == "__main__":
    unittest.main(verbosity=2)
