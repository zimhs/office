"""구글 캘린더 탭 — 내장 URL·탭 연결."""
from __future__ import annotations

import unittest
from urllib.parse import parse_qs, urlparse

import google_calendar_tab as gc


class GoogleCalendarTabTest(unittest.TestCase):
    def test_embed_url_month_korea(self):
        url = gc.calendar_embed_url(
            ["user@gmail.com", gc.GC_KR_HOLIDAY], height=780
        )
        p = urlparse(url)
        self.assertEqual(p.netloc, "calendar.google.com")
        self.assertTrue(p.path.endswith("/embed"))
        q = parse_qs(p.query)
        self.assertEqual(q.get("hl"), ["ko"])
        self.assertEqual(q.get("ctz"), ["Asia/Seoul"])
        self.assertEqual(q.get("mode"), ["MONTH"])
        self.assertEqual(q.get("wkst"), ["1"])
        self.assertIn("user@gmail.com", q.get("src", []))
        self.assertIn(gc.GC_KR_HOLIDAY, q.get("src", []))

    def test_month_weeks_start_sunday(self):
        weeks = gc.month_weeks(2026, 10)
        self.assertEqual(weeks[0][0].weekday(), 6)
        self.assertEqual(gc.iso_week_label(weeks[0]), 40)

    def test_auth_url_asks_calendar_and_tasks(self):
        url = gc.auth_url("cid.apps.googleusercontent.com", "state-1")
        self.assertIn("accounts.google.com", url)
        self.assertIn("calendar", url)
        self.assertIn("tasks", url)

    def test_ui_wired_as_new_tab_only(self):
        with open("google_calendar_tab.py", encoding="utf-8") as f:
            src = f.read()
        self.assertNotIn("tasks.google.com/embed", src)
        self.assertIn("할 일 목록", src)
        self.assertIn("구글 계정 연결", src)
        self.assertIn("render_google_calendar_tab", src)
        with open("app.py", encoding="utf-8") as f:
            app = f.read()
        self.assertIn("📆 구글캘린더", app)
        self.assertIn("with tab13:", app)
        self.assertNotIn("📅 방문·할일", app)
        tab13 = app[app.index("with tab13:") : app.index("구글캘린더 탭 오류")]
        self.assertIn("render_google_calendar_tab", tab13)
        self.assertIn("min_tabs=13", app)
        self.assertIn("render_visit_delivery_header", src)
        self.assertIn("@st.fragment", src)
        self.assertIn("_gc_pending_new", src)
        self.assertIn('st.rerun(scope="fragment")', src)


if __name__ == "__main__":
    unittest.main()
