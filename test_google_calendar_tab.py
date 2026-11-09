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

    def test_mini_month_weeks_start_monday(self):
        weeks = gc.mini_month_weeks(2026, 11)
        self.assertEqual(weeks[0][0].isoformat(), "2026-10-26")
        self.assertEqual(weeks[0][0].weekday(), 0)
        self.assertEqual(weeks[0][6].day, 1)

    def test_auth_url_asks_calendar_and_tasks(self):
        url = gc.auth_url("cid.apps.googleusercontent.com", "state-1")
        self.assertIn("accounts.google.com", url)
        self.assertIn("calendar", url)
        self.assertIn("tasks", url)
        self.assertIn("gmail.modify", url)
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
        self.assertIn("_gc_cal_panel", src)
        self.assertIn("_gc_cache_add_event", src)
        self.assertIn("ThreadPoolExecutor", src)
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
        self.assertNotIn("@st.dialog", src)
        self.assertNotIn("gc_ovc=", src)
        self.assertIn("gc_memo_pop", src)
        self.assertIn("_gc_place_memo_pop", src)
        self.assertIn("_gc_sel_chip", src)
        self.assertIn("_render_event_memo_pop", src)
        self.assertIn("font-size: 11px", src)
        self.assertIn("st-key-gc_td_", src)
        self.assertIn("#1a73e8", src)
        self.assertIn("#e8f0fe", src)
        self.assertIn("patch_event", src)
        self.assertIn("_gc_memo_open", src)
        self.assertIn("gc_ev_note", src)
        self.assertIn("_writable_calendars", src)
        self.assertIn('body["description"]', src)
        self.assertIn("연동실행", src)
        self.assertIn('key="gc_sync_run"', src)
        self.assertIn("on_click=_gc_run_sync_click", src)
        self.assertIn("def _gc_run_sync_click", src)
        sync_fn = src.split("def sync_shinil_gmail_to_calendar", 1)[1].split("\ndef ", 1)[0]
        self.assertNotIn("_EV_MEM.pop", sync_fn)
        self.assertNotIn("_retitle_mail_calendar_events", sync_fn)
        click = src.split("def _gc_run_sync_click", 1)[1].split("\ndef ", 1)[0]
        self.assertNotIn("st.rerun()", click)
        self.assertIn("_SYNC_BG.submit", click)
        self.assertIn("def _gc_sync_job", src)
        self.assertIn("def _gc_take_sync_result", src)
        self.assertNotIn("_gc_sync_google(y, m)", click)
        self.assertIn("GC_MAIL_SYNC_BATCH = 20", src)
        self.assertIn("gc_head_acts", src)
        self.assertIn("gc_cal_top", src)
        self.assertIn("gc_ym", src)
        self.assertIn("font-size: 18px", src)
        self.assertIn("_gc_sync_google", src)
        self.assertIn("_render_month_grid", src)
        self.assertIn("_render_mini_month", src)
        self.assertIn("gc_mini", src)
        self.assertIn("gc_mdp_", src)
        self.assertIn("_gc_cached_holiday_keys", src)
        self.assertIn("gc_md_", src)
        self.assertIn("mini_month_weeks", src)
        self.assertIn("[0.82, 2.70, 1.05]", src)
        self.assertIn("_gc_mini_panel", src)
        self.assertIn("_gc_mini_set_month", src)
        self.assertIn("_gc_set_view_month", src)
        self.assertIn("gc_mini_year", src)
        self.assertNotIn("_gc_sync_view_month", src)
        self.assertIn("gc_xd_", src)
        self.assertIn("st-key-gc_xd_", src)
        self.assertIn("translateY(-50%)", src)
        self.assertIn("text-overflow: ellipsis", src)
        self.assertIn("padding: 0 18px 0 4px", src)
        self.assertIn("_gc_pending_del_event", src)
        self.assertIn("gc_evlist_", src)
        self.assertIn("gc_evpick_", src)
        self.assertIn("_gc_shrink_event_titles", src)
        self.assertIn('st-key-gc_evpick_"] button p', src)
        self.assertIn("gc-evpick-fs", src)
        self.assertIn("__gcEvpickObs.disconnect", src)
        self.assertNotIn("MutationObserver", src)
        self.assertIn("_gc_schedule_prefetch", src)
        self.assertIn("_gc_pf_ym", src)
        self.assertIn("__gcDayPick", src)
        self.assertIn("dropChipMove", src)
        self.assertNotIn("bindChipMove", src)
        self.assertNotIn("gc_move_payload", src)
        self.assertNotIn("_gc_apply_day_move", src)
        self.assertNotIn("cursor: grab", src)
        self.assertIn("gc-js-sel", src)
        self.assertIn("st-key-gc_evpick_", src)
        shrink = src.split("def _gc_shrink_event_titles", 1)[1].split("def ", 1)[0]
        self.assertNotIn("st.html", shrink)
        self.assertNotIn("unsafe_allow_javascript", shrink)
        self.assertIn("font-size: 13px", src)
        self.assertIn("overflow-y: auto", src)
        self.assertIn("_gc_set_sel_event", src)
        self.assertIn("gc_day_", src)
        self.assertIn("_gc_sel_day", src)
        self.assertIn("_gc_set_sel_day", src)
        self.assertIn('st.session_state["gc_ev_day"] = day', src)
        self.assertIn("gc-sel", src)
        self.assertIn("gc_cellpink", src)
        self.assertIn("gc_celltoday", src)
        self.assertIn("#fdecea", src)
        self.assertIn("#e8f4ff", src)
        self.assertIn("border-radius: 10px", src)
        self.assertIn("flex-wrap: nowrap", src)
        self.assertIn("gc-dow-lab", src)
        self.assertNotIn("gc_ovc=", src)
        self.assertIn("일정 저장", src)
        self.assertNotIn('with st.expander("일정 추가"', src)
        self.assertIn("_gc_add_pop", src)
        self.assertIn("_gc_open_add_pop", src)
        self.assertIn('key="gc_add_open"', src)
        self.assertIn("gc_add_pop", src)
        self.assertIn("_render_add_event_pop", src)
        self.assertIn('st.rerun(scope="fragment")', src)
        self.assertNotIn('st.session_state["_vc_month"] = date(pending_m', src)
        self.assertIn("_gc_oauth_code_used", src)
        self.assertIn("_gc_oauth_redirect", src)
        self.assertIn("redirect_uri_mismatch", src)
        self.assertIn("_gc_refresh_tasks", src)
        self.assertIn("merge_tasks_with_google", src)
        self.assertIn("uncomplete_task", src)
        self.assertIn("include_completed", src)
        self.assertIn("완료된 할 일", src)
        self.assertIn("_gc_pending_restore", src)
        self.assertIn("gc_tr_", src)
        self.assertIn("_merge_event_days", src)
        self.assertNotIn("gc_task_import", src)
        self.assertNotIn("gc_task_export", src)
        self.assertNotIn("gc_task_sync", src)
        self.assertIn("sync_shinil_gmail_to_calendar", src)
        self.assertIn("GC_GMAIL_LABEL", src)
        self.assertIn("label_name: str = GC_GMAIL_LABEL", src)
        self.assertIn("GC_GMAIL_INBOX", src)
        self.assertIn("calendar_chip_label", src)
        self.assertIn("font-size:10px!important", src)
        self.assertNotIn("GC_BODY_FS", src)
        self.assertIn("_gc_load_disk_events", src)
        self.assertIn("GC_EV_CACHE", src)
        self.assertIn("gmail.modify", src)
        self.assertIn("개인업무", src)
        self.assertIn("GC_GMAIL_PERSONAL", src)
        self.assertIn('key="gc_mail_move_personal"', src)
        self.assertIn('key="gc_mail_move_shinil"', src)
        self.assertIn('key="gc_mail_del"', src)
        self.assertIn("gmail_move_inbox_to_label", src)
        self.assertIn("gmail_trash_mail", src)
        self.assertIn("_gc_pending_mail_del", src)
        self.assertIn("Gmail 권한 다시 연결", src)
        self.assertIn("token_has_gmail_modify", src)
        self.assertIn("연결 해제 후 처음부터", src)
        consume = src.split("def _consume_oauth_redirect", 1)[1].split("def ", 1)[0]
        self.assertNotIn("_usable_token(load_token())", consume)
        self.assertIn("구글 캘린더·할 일·Gmail이 연결되었습니다", src)
        self.assertNotIn("gmail.readonly", src)
        self.assertIn("insert_mail_day_event", src)
        self.assertIn("메일 열기:", src)
        self.assertIn("Gmail API", src)
        self.assertIn("_auth_error_message", src)
        self.assertIn("_handle_refresh_failure", src)
        self.assertIn("구글 인증이 만료되었습니다", src)


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
        self.assertIn("미팅", html)
        self.assertIn("gc-chip", html)
        self.assertIn("gc-table", html)
        self.assertNotIn("gc_ovc=", html)
        self.assertNotIn("<a ", html)
        self.assertIn("class=\"gc-x\"", html)
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
        self.assertNotIn("gc_ovc=", html_ro)
        self.assertNotIn("<a ", html_ro)
        self.assertNotIn("class=\"gc-x\"", html_ro)
        self.assertIn("한글날", html_ro)
        html_sel = gc._month_html(
            2026,
            10,
            {},
            d(2026, 10, 5),
            selected=d(2026, 10, 12),
        )
        self.assertIn("gc-sel", html_sel)
        self.assertIn("gc-today", html)
        html_sun = gc._month_html(2026, 10, {}, d(2026, 10, 5))
        self.assertIn("gc-sun", html_sun)
        self.assertIn("#fdecea", html_sun)
        html_hol = gc._month_html(
            2026,
            10,
            {
                "2026-10-09": [
                    {
                        "title": "한글날",
                        "cal_id": gc.GC_KR_HOLIDAY,
                        "writable": False,
                    }
                ]
            },
            d(2026, 10, 5),
        )
        self.assertIn("gc-hol", html_hol)
        self.assertTrue(gc._gc_is_holiday_day([{"cal_id": gc.GC_KR_HOLIDAY}]))
        self.assertFalse(gc._gc_is_holiday_day([{"cal_id": "user@gmail.com"}]))

    def test_month_cache_skips_google_fetch(self):
        cached = {"2026-10-05": [{"title": "미팅", "id": "e1"}]}
        box = {"ev:2026-10": cached}
        with patch.object(gc, "_gc_api_box", return_value=box), patch.object(
            gc, "_events_by_day"
        ) as fetch:
            out = gc._gc_cached_events([{"id": "c1"}], 2026, 10)
        self.assertIs(out, cached)
        fetch.assert_not_called()

    def test_mem_cache_skips_google_on_month_move(self):
        gc._EV_MEM["ev:2026-11"] = {"2026-11-01": [{"title": "미리", "id": "p1"}]}
        box: dict = {}
        try:
            with patch.object(gc, "_gc_api_box", return_value=box), patch.object(
                gc, "_events_by_day"
            ) as fetch:
                out = gc._gc_cached_events([{"id": "c1"}], 2026, 11)
            self.assertEqual(out["2026-11-01"][0]["title"], "미리")
            fetch.assert_not_called()
        finally:
            gc._EV_MEM.pop("ev:2026-11", None)

    def test_shift_month_and_prefetch_skips_cached(self):
        self.assertEqual(gc._gc_shift_month(2026, 1, -1), (2025, 12))
        self.assertEqual(gc._gc_shift_month(2026, 12, 1), (2027, 1))
        gc._EV_MEM["ev:2026-09"] = {}
        gc._EV_MEM["ev:2026-11"] = {}
        try:
            with patch.object(gc, "_events_by_day") as fetch:
                gc._gc_schedule_prefetch([{"id": "c1"}], 2026, 10)
            fetch.assert_not_called()
        finally:
            gc._EV_MEM.pop("ev:2026-09", None)
            gc._EV_MEM.pop("ev:2026-11", None)


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

    def test_writable_calendars_skip_holiday(self):
        items = gc._writable_calendars(
            [
                {
                    "id": gc.GC_KR_HOLIDAY,
                    "summary": "대한민국 공휴일",
                    "writable": False,
                },
                {
                    "id": "user@gmail.com",
                    "summary": "업무",
                    "writable": True,
                    "primary": True,
                },
            ]
        )
        self.assertEqual([c["summary"] for c in items], ["업무"])
        with self.assertRaises(RuntimeError):
            gc.insert_event(
                gc.GC_KR_HOLIDAY, "제목", __import__("datetime").date(2026, 11, 11), True
            )

    def test_refresh_keeps_rotated_refresh_token(self):
        class _Resp:
            ok = True
            content = b"{}"
            status_code = 200

            def json(self):
                return {"access_token": "at-new", "refresh_token": "rt-rotated", "expires_in": 3600}

        with patch.object(gc.requests, "post", return_value=_Resp()):
            out = gc.refresh_token("cid", "csec", {"refresh_token": "rt-old"})
        self.assertEqual(out.get("refresh_token"), "rt-rotated")
        self.assertEqual(out.get("access_token"), "at-new")
        self.assertTrue(out.get("_obtained"))

    def test_refresh_400_is_korean_not_token_url(self):
        class _Resp:
            ok = False
            status_code = 400
            content = b"{}"

            def json(self):
                return {
                    "error": "invalid_grant",
                    "error_description": "Token has been expired or revoked.",
                }

        with patch.object(gc.requests, "post", return_value=_Resp()):
            with self.assertRaises(RuntimeError) as ctx:
                gc.refresh_token("cid", "csec", {"refresh_token": "rt-dead"})
        msg = str(ctx.exception)
        self.assertIn("다시 연결", msg)
        self.assertNotIn("oauth2.googleapis.com", msg)

    def test_expired_without_obtained_tries_refresh(self):
        self.assertTrue(gc._expired({"access_token": "at", "refresh_token": "rt"}))
        self.assertTrue(gc._expired({"refresh_token": "rt"}))
        self.assertFalse(
            gc._expired(
                {
                    "access_token": "at",
                    "expires_in": 3600,
                    "_obtained": __import__("datetime")
                    .datetime.now(__import__("datetime").timezone.utc)
                    .isoformat(),
                }
            )
        )

    def test_usable_without_access_token(self):
        self.assertTrue(gc._usable_token({"refresh_token": "r"}))
        self.assertFalse(gc._usable_token({}))

    def test_local_tasks_merge_keeps_dashboard_and_google(self):
        with tempfile.TemporaryDirectory() as td:
            home = os.path.join(td, "tasks.json")
            with patch.object(gc, "GC_TASKS_HOME", home), patch.object(
                gc, "GC_TASKS_ST", os.path.join(td, "st.json")
            ), patch.object(gc, "GC_TASKS", os.path.join(td, "cache.json")), patch.object(
                gc.st, "session_state", {}
            ), patch.object(
                gc, "list_tasks", return_value=[{"id": "g1", "title": "구글만"}]
            ), patch.object(
                gc, "add_task", side_effect=RuntimeError("offline")
            ):
                gc.save_local_tasks(
                    "list-1",
                    [{"id": "local-a", "title": "대시보드", "notes": "남김", "local_only": True}],
                )
                merged = gc.merge_tasks_with_google("list-1")
                titles = {t["title"] for t in merged}
                self.assertIn("대시보드", titles)
                self.assertIn("구글만", titles)


class GoogleCalendarGmailSyncTest(unittest.TestCase):
    def test_thread_url_and_received_date(self):
        from datetime import date as d
        from datetime import datetime, timezone

        url = gc.gmail_thread_url("18abcDEF")
        self.assertTrue(url.startswith("https://mail.google.com/mail/u/0/#all/"))
        self.assertIn("18abcDEF", url)
        ms = int(datetime(2026, 11, 6, 8, 0, tzinfo=timezone.utc).timestamp() * 1000)
        self.assertEqual(gc.gmail_received_date(ms), d(2026, 11, 6))
        self.assertIsNone(gc.gmail_received_date("bad"))

    def test_mail_event_notes_include_link(self):
        from datetime import date as d

        notes = gc.mail_event_notes(
            subject="견적 요청서",
            sender="A <a@b.com>",
            received=d(2026, 11, 6),
            snippet="견적 문의",
            thread_id="tid1",
        )
        self.assertIn("제목: 견적 요청서", notes)
        self.assertIn("보낸사람: A <a@b.com>", notes)
        self.assertIn("받은날: 2026-11-06", notes)
        self.assertIn("메일 열기:", notes)
        self.assertIn("tid1", notes)
        self.assertEqual(gc.mail_event_title("납품"), "메일")
        self.assertEqual(
            gc.calendar_chip_label({"title": "[메일] Fwd: 구매발주", "description": "메일 열기: x"}),
            "메일",
        )
        self.assertEqual(gc.calendar_chip_label({"title": "미팅", "hhmm": "09:00"}), "09:00 미팅")

    def test_mail_move_and_trash_helpers(self):
        ids = gc.gmail_ids_from_event(
            {
                "description": "제목: a\n메일 열기: https://mail.google.com/mail/u/0/#all/18abcDEF",
                "gmail_msg": "mid1",
            }
        )
        self.assertEqual(ids["thread"], "18abcDEF")
        self.assertEqual(ids["msg"], "mid1")
        body = gc.gmail_move_body("Label_9")
        self.assertEqual(body["addLabelIds"], ["Label_9"])
        self.assertEqual(body["removeLabelIds"], ["INBOX"])
        self.assertIn("threads/tid9/trash", gc.gmail_trash_path(thread_id="tid9"))
        self.assertEqual(gc.GC_GMAIL_PERSONAL, "개인업무")
        self.assertEqual(gc.GC_GMAIL_LABEL, "신일가스")

    def test_match_gmail_label_leaf_and_exact(self):
        labels = [
            {"id": "1", "name": "INBOX"},
            {"id": "2", "name": "거래처/신일가스"},
            {"id": "3", "name": "신일가스"},
        ]
        hit = gc.match_gmail_label(labels, "신일가스")
        self.assertEqual(hit.get("id"), "3")
        hit2 = gc.match_gmail_label(
            [{"id": "2", "name": "거래처/신일가스"}], "신일가스"
        )
        self.assertEqual(hit2.get("id"), "2")
        self.assertEqual(gc.match_gmail_label(labels, "없는폴더"), {})

    def test_sync_skips_already_done_and_inserts_new(self):
        from datetime import date as d

        with tempfile.TemporaryDirectory() as td:
            home = os.path.join(td, "sync.json")
            cache = os.path.join(td, "cache.json")
            created = []

            def _insert(cal_id, **kwargs):
                kwargs["cal_id"] = cal_id
                created.append(kwargs)
                return {
                    "id": "ev-" + kwargs["msg_id"],
                    "summary": kwargs["title"],
                    "start": {"date": kwargs["day"].isoformat()},
                    "end": {"date": kwargs["day"].isoformat()},
                    "description": kwargs["notes"],
                }

            listed = [
                {"id": "old1", "threadId": "t-old"},
                {"id": "new1", "threadId": "t-new"},
            ]
            with patch.object(gc, "GC_GMAIL_SYNC_HOME", home), patch.object(
                gc, "GC_GMAIL_SYNC", cache
            ), patch.object(gc.st, "session_state", {"_gc_api": {}}), patch.object(
                gc,
                "list_gmail_labels",
                return_value=[{"id": "Label_shinil", "name": "신일가스"}],
            ), patch.object(
                gc, "list_gmail_message_ids", return_value=listed
            ), patch.object(
                gc, "_list_synced_mail_ids_from_calendar", return_value=set()
            ), patch.object(
                gc, "_primary_writable_cal_id", return_value="user@gmail.com"
            ), patch.object(
                gc,
                "gmail_message_meta",
                return_value={
                    "id": "new1",
                    "threadId": "t-new",
                    "subject": "발주",
                    "sender": "a@b.com",
                    "snippet": "본문",
                    "day": d(2026, 11, 6),
                },
            ), patch.object(
                gc, "insert_mail_day_event", side_effect=_insert
            ), patch.object(
                gc, "_gc_cache_add_event"
            ), patch.object(
                gc, "_retitle_mail_calendar_events", return_value=0
            ):
                gc.save_gmail_sync_state({"ids": {"old1": "2026-01-01"}})
                out = gc.sync_shinil_gmail_to_calendar(since=d(2026, 11, 6))
            self.assertEqual(out["added"], 1)
            self.assertEqual(out["skipped"], 1)
            self.assertEqual(created[0]["msg_id"], "new1")
            self.assertEqual(created[0]["day"], d(2026, 11, 6))
            self.assertIn("제목: 발주", created[0]["notes"])
            self.assertIn("메일 열기:", created[0]["notes"])
            self.assertEqual(out["label"], "신일가스")

    def test_set_sel_event_opens_memo(self):
        from datetime import date as d

        state = {}
        with patch.object(gc.st, "session_state", state):
            gc._gc_set_sel_event(d(2026, 11, 9), "e1", "20261109_0")
        self.assertTrue(state["_gc_memo_open"])
        self.assertEqual(state["_gc_sel_event"], "e1")
        self.assertEqual(state["_gc_sel_chip"], "20261109_0")

    def test_find_event_by_id(self):
        ev = {"id": "e1", "title": "회의"}
        found = gc._gc_find_event({"2026-11-09": [ev]}, "e1")
        self.assertEqual(found["title"], "회의")


if __name__ == "__main__":
    unittest.main()
