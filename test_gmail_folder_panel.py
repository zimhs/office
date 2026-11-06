"""메일 탭 하단 신일가스 폴더 보기."""
from __future__ import annotations

import base64
import unittest

import gmail_folder_panel as gf
import google_calendar_tab as gc


def _b64(text: str) -> str:
    return base64.urlsafe_b64encode(text.encode("utf-8")).decode("ascii").rstrip("=")


class GmailFolderPanelTest(unittest.TestCase):
    def test_plain_body_prefers_text(self):
        payload = {
            "mimeType": "multipart/alternative",
            "parts": [
                {
                    "mimeType": "text/plain",
                    "body": {"data": _b64("안녕 본문")},
                },
                {
                    "mimeType": "text/html",
                    "body": {"data": _b64("<b>안녕</b> 본문")},
                },
            ],
        }
        self.assertEqual(gf.gmail_plain_body(payload), "안녕 본문")

    def test_html_only_strips_tags(self):
        payload = {
            "mimeType": "text/html",
            "body": {"data": _b64("<p>견적 <b>요청</b></p>")},
        }
        text = gf.gmail_plain_body(payload)
        self.assertIn("견적", text)
        self.assertIn("요청", text)
        self.assertNotIn("<p>", text)

    def test_panel_wired_under_mail_tab_only(self):
        with open("gmail_folder_panel.py", encoding="utf-8") as f:
            src = f.read()
        self.assertIn("render_shinil_gmail_folder", src)
        self.assertIn("신일가스 메일 폴더", src)
        self.assertIn("gf_shinil_sel", src)
        self.assertNotIn("MutationObserver", src)
        self.assertNotIn("@st.dialog", src)
        self.assertIn("GC_GMAIL_LABEL", src)
        with open("price_increase_tab.py", encoding="utf-8") as f:
            pi = f.read()
        self.assertIn("render_shinil_gmail_folder", pi)
        self.assertEqual(gc.GC_GMAIL_LABEL, "신일가스")


if __name__ == "__main__":
    unittest.main()
