"""구글 캘린더 탭 — 내장 URL·탭 연결."""
from __future__ import annotations

import json
import os
import tempfile
import unittest
from unittest.mock import patch
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
        self.assertIn("redirect_uri=", url)

    def test_redirect_candidates_include_local_and_cloud(self):
        items = gc._redirect_candidates("http://127.0.0.1:8501/")
        self.assertIn("http://127.0.0.1:8501/", items)
        self.assertIn("http://localhost:8501/", items)
        self.assertIn(gc._CLOUD_REDIRECT, items)

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
        self.assertIn("gc_new_note", src)
        self.assertIn('body["notes"]', src)
        self.assertIn("_gc_note_edit", src)
        self.assertIn("gc_ne_", src)
        self.assertNotIn('with st.expander("메모"', src)
        self.assertIn("GC_TOKEN_HOME", src)
        self.assertIn("refresh_token", src)
        self.assertIn("_gc_pending_month", src)
        self.assertIn('key="gc_next"', src)
        self.assertIn("st-key-gc_ne_", src)
        self.assertIn('type="tertiary"', src)
        self.assertNotIn('help="메모"', src)
        self.assertNotIn("st.popover", src)
        self.assertIn("@st.dialog", src)
        self.assertIn("delete_event", src)
        self.assertIn("gc-x", src)
        self.assertIn("_gc_pending_del_event", src)
        self.assertIn("일정 저장", src)
        self.assertIn("gc_ev_note", src)
        self.assertIn('body["description"]', src)
        self.assertIn("gc_ovc=", src)
        self.assertIn("_gc_oauth_code_used", src)
        self.assertIn("_gc_oauth_redirect", src)
        self.assertIn("redirect_uri_mismatch", src)
        self.assertIn("_gc_refresh_tasks", src)
        self.assertIn("gc_task_import", src)
        self.assertIn("gc_task_export", src)
        self.assertIn("export_tasks_to_google", src)
        self.assertIn("불러오기", src)
        self.assertIn("내보내기", src)
        self.assertNotIn("gc_task_sync", src)


    def test_month_html_x_only_on_writable(self):
        from datetime import date as d
        html = gc._month_html(
            2026,
            10,
            {
                "2026-10-05": [
                    {
                        "title": "미팅",
                        "color": "#1a73e8",
                        "hhmm": "09:00",
                        "id": "evt1",
                        "cal_id": "user@gmail.com",
                        "writable": True,
                    }
                ]
            },
            d(2026, 10, 5),
        )
        self.assertIn("gc-x", html)
        self.assertIn("gc_xc=", html)
        self.assertIn("×", html)
        html_ro = gc._month_html(
            2026,
            10,
            {
                "2026-10-05": [
                    {
                        "title": "한글날",
                        "color": "#d93025",
                        "hhmm": "",
                        "id": "hol1",
                        "cal_id": gc.GC_KR_HOLIDAY,
                        "writable": False,
                    }
                ]
            },
            d(2026, 10, 5),
        )
        self.assertNotIn('class="gc-x"', html_ro)
        self.assertIn("gc_ovc=", html)


class GoogleCalendarTokenPersistTest(unittest.TestCase):
    def test_load_token_from_home_when_cache_missing(self):
        with tempfile.TemporaryDirectory() as td:
            home = os.path.join(td, "home.json")
            cache = os.path.join(td, "cache.json")
            payload = {
                "refresh_token": "rt-1",
                "access_token": "at-1",
                "expires_in": 3599,
            }
            with open(home, "w", encoding="utf-8") as f:
                json.dump(payload, f)
            with patch.object(gc, "GC_TOKEN_HOME", home), patch.object(
                gc, "GC_TOKEN_ST", os.path.join(td, "st.json")
            ), patch.object(gc, "GC_TOKEN", cache), patch.object(
                gc.st, "session_state", {}
            ):
                tok = gc.load_token()
            self.assertEqual(tok.get("refresh_token"), "rt-1")
            self.assertTrue(os.path.exists(cache))

    def test_save_keeps_refresh_token(self):
        with tempfile.TemporaryDirectory() as td:
            home = os.path.join(td, "home.json")
            cache = os.path.join(td, "cache.json")
            stpath = os.path.join(td, "st.json")
            with patch.object(gc, "GC_TOKEN_HOME", home), patch.object(
                gc, "GC_TOKEN_ST", stpath
            ), patch.object(gc, "GC_TOKEN", cache), patch.object(
                gc.st,
                "session_state",
                {"_gc_token": {"refresh_token": "keep-me", "access_token": "old"}},
            ):
                gc.save_token({"access_token": "new"})
            with open(home, encoding="utf-8") as f:
                saved = json.load(f)
            self.assertEqual(saved.get("refresh_token"), "keep-me")
            self.assertEqual(saved.get("access_token"), "new")

    def test_usable_without_access_token(self):
        self.assertTrue(gc._usable_token({"refresh_token": "r"}))
        self.assertFalse(gc._usable_token({}))

    def test_local_tasks_kept_until_import(self):
        with tempfile.TemporaryDirectory() as td:
            home = os.path.join(td, "tasks.json")
            with patch.object(gc, "GC_TASKS_HOME", home), patch.object(
                gc, "GC_TASKS_ST", os.path.join(td, "st.json")
            ), patch.object(gc, "GC_TASKS", os.path.join(td, "cache.json")), patch.object(
                gc.st, "session_state", {}
            ), patch.object(gc, "list_tasks", return_value=[{"id": "g1", "title": "구글만"}]):
                gc.save_local_tasks(
                    "list-1",
                    [{"id": "local-a", "title": "대시보드", "notes": "남김", "local_only": True}],
                )
                shown = gc._gc_cached_tasks("list-1")
                self.assertEqual(shown[0]["title"], "대시보드")
                imported = gc._gc_cached_tasks("list-1", force=True)
                self.assertEqual(imported[0]["title"], "구글만")


if __name__ == "__main__":
    unittest.main()
