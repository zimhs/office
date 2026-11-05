"""업무일지 캡처 붙여넣기·크기 조절."""
from __future__ import annotations

import base64
import os
import tempfile
import unittest
from unittest.mock import patch

_PNG_1 = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
)


class WorklogCaptureTest(unittest.TestCase):
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

    def test_clamp_and_iso_from_path(self):
        self.assertEqual(self.wt._wl_clamp_cap_width(5), 20)
        self.assertEqual(self.wt._wl_clamp_cap_width(140), 100)
        self.assertEqual(self.wt._wl_clamp_cap_width("70"), 70)
        self.assertEqual(
            self.wt._wl_iso_from_worklog_path("/x/_preview_2026-11-06.xlsx"),
            "2026-11-06",
        )

    def test_png_size_and_roundtrip(self):
        w, h = self.wt._wl_image_wh(_PNG_1)
        self.assertEqual((w, h), (1, 1))
        png = self.wt._wl_normalize_capture_png(_PNG_1)
        self.assertTrue(png and png[:8] == b"\x89PNG\r\n\x1a\n")
        iso = "2026-11-06"
        self.wt._wl_write_captures(
            iso,
            [{"id": "ab12", "file": "ab12.png", "w_pct": 55, "page": 0, "w0": 1, "h0": 1}],
        )
        folder = self.wt._wl_capture_dir(iso)
        os.makedirs(folder, exist_ok=True)
        with open(os.path.join(folder, "ab12.png"), "wb") as f:
            f.write(_PNG_1)
        rows = self.wt._wl_read_captures(iso)
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["w_pct"], 55)
        html = self.wt._wl_capture_html_block(
            os.path.join(self.folder, "2026-11-06.xlsx"), 1, 1.0
        )
        self.assertIn("data:image/png;base64,", html)
        self.assertIn("renderCaps", self.wt._WL_LINES_JS)
        self.assertIn("captures", self.wt._WL_LINES_JS)
        self.wt._wl_purge_captures(iso)
        self.assertFalse(os.path.isdir(self.wt._wl_capture_dir(iso)))

    def test_payload_and_cap_state(self):
        iso = "2026-11-06"
        folder = self.wt._wl_capture_dir(iso)
        os.makedirs(folder, exist_ok=True)
        with open(os.path.join(folder, "ab12.png"), "wb") as f:
            f.write(_PNG_1)
        rows = [{"id": "ab12", "file": "ab12.png", "w_pct": 40, "page": 0, "w0": 1, "h0": 1}]
        self.wt._wl_write_captures(iso, rows)
        ss = {self.wt._wl_caps_session_key(iso): rows}
        with patch.object(self.wt.st, "session_state", ss):
            payload = self.wt._wl_capture_payload(iso, 0)
            self.assertEqual(len(payload), 1)
            self.assertEqual(payload[0]["id"], "ab12")
            self.assertEqual(payload[0]["w_pct"], 40)
            self.assertTrue(str(payload[0]["src"]).startswith("data:image/png;base64,"))
            self.assertEqual(self.wt._wl_capture_payload(iso, 1), [])
            ss["ck"] = {"cap_w": '{"id":"ab12","w_pct":88,"t":1}'}
            self.wt._wl_apply_cap_state(iso, ss["ck"], kind="cap_w")
            self.assertEqual(ss[self.wt._wl_caps_session_key(iso)][0]["w_pct"], 88)
            ss["ck"] = {"cap_del": '{"id":"ab12","t":2}'}
            self.wt._wl_apply_cap_state(iso, ss["ck"], kind="cap_del")
            self.assertEqual(ss[self.wt._wl_caps_session_key(iso)], [])

    def test_ingest_pending_file_before_editor(self):
        iso = "2026-11-06"

        class _F:
            name = "a.png"

            def getvalue(self):
                return _PNG_1

        ss = {f"wl_cap_file_{iso}_0": _F()}
        with patch.object(self.wt.st, "session_state", ss):
            added = self.wt._wl_ingest_pending_captures(iso, 0)
            self.assertTrue(added)
            self.assertEqual(ss.get(f"wl_cap_file_n_{iso}"), 1)
            self.assertEqual(len(ss[self.wt._wl_caps_session_key(iso)]), 1)
            self.assertFalse(self.wt._wl_ingest_pending_captures(iso, 0))

    def test_add_capture_bytes_assigns_id(self):
        iso = "2026-11-06"
        ss = {}
        with patch.object(self.wt.st, "session_state", ss):
            cid = self.wt._wl_add_capture_bytes(iso, _PNG_1, page=0)
        self.assertTrue(cid)
        self.assertEqual(len(cid), 8)
        rows = self.wt._wl_read_captures(iso)
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["id"], cid)

    def test_decode_data_url(self):
        url = "data:image/png;base64," + base64.b64encode(_PNG_1).decode("ascii")
        blob = self.wt._wl_decode_data_url(url)
        self.assertEqual(blob[:8], b"\x89PNG\r\n\x1a\n")
        self.assertIsNone(self.wt._wl_decode_data_url("hello"))
