"""방문·할일 캘린더 — 저장·이력·탭 연결."""
from __future__ import annotations

import tempfile
import unittest
from datetime import date
from pathlib import Path
from unittest.mock import patch

import pandas as pd


class VisitCalendarTest(unittest.TestCase):
    def setUp(self):
        import visit_calendar_tab as vc

        self.vc = vc
        self.tmp = tempfile.TemporaryDirectory()
        self.store = str(Path(self.tmp.name) / "store.json")
        self.dir_patch = patch.object(vc, "VC_DIR", self.tmp.name)
        self.path_patch = patch.object(vc, "VC_STORE", self.store)
        self.dir_patch.start()
        self.path_patch.start()

    def tearDown(self):
        self.path_patch.stop()
        self.dir_patch.stop()
        self.tmp.cleanup()

    def test_todo_roundtrip(self):
        item = self.vc.add_todo(
            {"업체명": "지엠레이저", "세부사항": "지엠철강분 외상대", "due": date(2026, 9, 16)}
        )
        self.assertEqual(item["title"], "지엠레이저")
        self.assertEqual(item["note"], "지엠철강분 외상대")
        self.assertEqual(item["client"], "지엠레이저")
        self.assertEqual(item["due"], "2026-09-16")
        self.assertFalse(item["done"])
        self.assertFalse(item["starred"])
        self.assertTrue(self.vc.toggle_todo(item["id"], done=True))
        self.assertTrue(self.vc.load_store()["todos"][0]["done"])
        self.assertTrue(self.vc.delete_todo(item["id"]))
        self.assertEqual(self.vc.load_store()["todos"], [])

    def test_todo_star_toggle(self):
        item = self.vc.add_todo({"업체명": "이엔에이치", "세부사항": "단가인상"})
        self.assertTrue(self.vc.toggle_todo_star(item["id"]))
        self.assertTrue(self.vc.load_store()["todos"][0]["starred"])
        self.assertTrue(self.vc.toggle_todo_star(item["id"]))
        self.assertFalse(self.vc.load_store()["todos"][0]["starred"])

    def test_todo_update_keeps_id(self):
        item = self.vc.add_todo(
            {"업체명": "예제", "세부사항": "첫내용", "due": date(2026, 9, 2)}
        )
        updated = self.vc.update_todo(
            item["id"],
            {"업체명": "예제수정", "세부사항": "바꿈", "due": date(2026, 9, 10)},
        )
        self.assertEqual(updated["id"], item["id"])
        self.assertEqual(updated["title"], "예제수정")
        self.assertEqual(updated["note"], "바꿈")
        self.assertEqual(updated["due"], "2026-09-10")
        self.assertEqual(self.vc.load_store()["todos"][0]["title"], "예제수정")
        with self.assertRaises(ValueError):
            self.vc.update_todo("missing", {"업체명": "없음"})

    def test_direct_visit_toggle_marks_day(self):
        client = "라콜 주식회사(구.신정우)"
        self.assertIsNone(self.vc._direct_visit_on(self.vc.load_store(), date(2026, 8, 7), client))
        self.vc.add_visit({"date": date(2026, 8, 7), "staff": "김혁수", "client": client})
        hit = self.vc._direct_visit_on(self.vc.load_store(), date(2026, 8, 7), client)
        self.assertIsNotNone(hit)
        chips = self.vc.calendar_chips(self.vc.load_store(), [], date(2026, 8, 1), [])
        kinds = {c["kind"] for c in chips["2026-08-07"]}
        self.assertIn("visit", kinds)
        self.assertTrue(self.vc.delete_visit(hit["id"]))
        self.assertIsNone(self.vc._direct_visit_on(self.vc.load_store(), date(2026, 8, 7), client))

    def test_planned_visit_shows_prefixed_name(self):
        client = "라콜 주식회사(구.신정우)"
        item = self.vc.add_visit(
            {"date": date(2026, 8, 8), "staff": "김혁수", "client": client, "status": "planned"}
        )
        self.assertEqual(item["status"], "planned")
        chips = self.vc.calendar_chips(self.vc.load_store(), [], date(2026, 8, 1), [])
        kinds = {c["kind"] for c in chips["2026-08-08"]}
        self.assertIn("planned", kinds)
        self.assertNotIn("visit", kinds)
        nm, mark = self.vc._strip_visit_mark(chips["2026-08-08"])
        self.assertEqual(mark, "planned")
        self.assertTrue(nm.startswith("예·"))
        self.assertIn("라콜", nm)

    def test_month_visit_rows_splits_done_and_planned(self):
        self.vc.add_visit(
            {
                "date": date(2026, 9, 7),
                "staff": "김혁수",
                "client": "한신테크",
                "status": "done",
            }
        )
        self.vc.add_visit(
            {
                "date": date(2026, 9, 23),
                "staff": "김혁수",
                "client": "라콜 주식회사(구.신정우)",
                "status": "planned",
            }
        )
        self.vc.add_visit(
            {
                "date": date(2026, 8, 8),
                "staff": "김혁수",
                "client": "다른달",
                "status": "planned",
            }
        )
        rows = self.vc.month_visit_rows(self.vc.load_store(), date(2026, 9, 1), "김혁수")
        self.assertEqual(len(rows), 2)
        by_st = {r["status"]: r["client"] for r in rows}
        self.assertEqual(by_st["done"], "한신테크")
        self.assertEqual(by_st["planned"], "라콜 주식회사(구.신정우)")
        html = self.vc._month_schedule_items_html(
            [r for r in rows if r["status"] == "planned"],
            date(2026, 9, 23),
            "라콜 주식회사(구.신정우)",
            "planned",
        )
        self.assertIn("방문예정", html)
        self.assertIn("라콜 주식회사", html)
        self.assertIn("23일", html)

    def test_clear_day_visit_removes_calendar_and_month(self):
        client = "라콜 주식회사(구.신정우)"
        item = self.vc.add_visit(
            {"date": date(2026, 9, 23), "staff": "김혁수", "client": client, "status": "planned"}
        )
        self.assertTrue(self.vc.clear_day_visit(date(2026, 9, 23), client))
        store = self.vc.load_store()
        self.assertEqual(store["visits"], [])
        self.assertFalse(self.vc.delete_visit(item["id"]))
        chips = self.vc.calendar_chips(store, [], date(2026, 9, 1), [])
        self.assertNotIn("2026-09-23", chips)
        self.assertEqual(self.vc.month_visit_rows(store, date(2026, 9, 1), "김혁수"), [])
        nm, mark = self.vc._strip_visit_mark(chips.get("2026-09-23") or [])
        self.assertEqual(nm, "")
        self.assertEqual(mark, "")
        self.assertIsNone(self.vc._direct_visit_on(store, date(2026, 9, 23), client))

    def test_iso_drops_timestamp_time(self):
        self.assertEqual(self.vc._iso(pd.Timestamp("2026-09-04 00:00:00")), "2026-09-04")
        self.assertEqual(self.vc._iso(date(2026, 9, 4)), "2026-09-04")

    def test_delivery_rows_on_selected_day(self):
        df = pd.DataFrame(
            {
                "담당자": ["김혁수", "김혁수", "김혁수"],
                "거래처": ["라콜 주식회사(구.신정우)", "라콜 주식회사(구.신정우)", "다른곳"],
                "매출일_dt": [date(2026, 8, 11), date(2026, 8, 11), date(2026, 8, 11)],
                "품목명": ["N2 (kg, Bulk)", "CO2 (kg, Bulk)", "AR"],
                "출고량": [10, 5, 1],
                "매출액": [100000, 50000, 9],
            }
        )
        rows = self.vc._sales_delivery_rows(df, "김혁수", "라콜 주식회사(구.신정우)")
        self.assertEqual(len(rows), 2)
        day = self.vc.delivery_rows_on(rows, date(2026, 8, 11))
        self.assertEqual(len(day), 2)
        self.assertEqual({r["item"] for r in day}, {"N2 (kg, Bulk)", "CO2 (kg, Bulk)"})
        self.assertEqual(self.vc.delivery_rows_on(rows, date(2026, 8, 12)), [])
        other = self.vc._sales_delivery_rows(df, "김혁수", "다른곳")
        self.assertEqual(len(other), 1)
        self.assertEqual(other[0]["item"], "AR")
        idx = self.vc._sales_by_client(df, "김혁수")
        self.assertEqual(len(idx[self.vc._company_key("라콜 주식회사(구.신정우)")]), 2)

    def test_visit_and_history_merge(self):
        self.vc.add_visit(
            {"date": date(2026, 9, 10), "staff": "김혁수", "client": "한신테크", "note": "미팅"}
        )
        df = pd.DataFrame(
            {
                "담당자": ["김혁수", "김혁수"],
                "거래처": ["한신테크", "다른곳"],
                "매출일_dt": [date(2026, 8, 1), date(2026, 7, 1)],
            }
        )
        hist = self.vc.merge_visit_history(df, "김혁수", "한신테크")
        sources = {h["source"] for h in hist}
        self.assertIn("직접입력", sources)
        self.assertIn("납품", sources)
        days = [h["date"] for h in hist]
        self.assertIn("2026-09-10", days)
        self.assertIn("2026-08-01", days)
        self.assertNotIn("2026-07-01", days)

    def test_marked_dates_include_todos(self):
        store = {
            "todos": [{"due": "2026-09-16", "done": False}],
            "visits": [{"date": "2026-09-03"}],
        }
        marks = self.vc.marked_dates(store, [{"date": "2026-09-16", "source": "업무일지"}], date(2026, 9, 1))
        self.assertIn("todo", marks["2026-09-16"])
        self.assertIn("worklog", marks["2026-09-16"])
        self.assertIn("visit", marks["2026-09-03"])

    def test_calendar_chips_bulk_and_other(self):
        store = {
            "todos": [{"due": "2026-09-16", "done": False, "title": "견적"}],
            "visits": [{"date": "2026-09-10", "client": "한신테크"}],
        }
        deliveries = [
            {"date": "2026-09-10", "item": "N2 (kg, Bulk)", "qty": 1200, "bulk": True},
            {"date": "2026-09-10", "item": "CO2 (kg, Bulk)", "qty": 800, "bulk": True},
            {"date": "2026-09-10", "item": "아세틸렌", "qty": 2, "bulk": False},
        ]
        chips = self.vc.calendar_chips(store, [], date(2026, 9, 1), deliveries)
        labels = {c["label"] for c in chips["2026-09-10"]}
        self.assertIn("한신테크", labels)
        self.assertIn("N2 1,200", labels)
        self.assertIn("CO2 800", labels)
        self.assertIn("실린더", labels)
        self.assertNotIn("벌크외", labels)
        self.assertEqual(chips["2026-09-16"][0]["kind"], "todo")

    def test_delivery_chips_only_bulk_and_cylinder(self):
        store = {
            "todos": [{"due": "2026-09-16", "done": False, "title": "견적"}],
            "visits": [{"date": "2026-09-10", "client": "한신테크"}],
        }
        deliveries = [
            {"date": "2026-09-10", "item": "N2 (kg, Bulk)", "qty": 1200, "bulk": True},
            {"date": "2026-09-10", "item": "아세틸렌", "qty": 2, "bulk": False},
        ]
        visit = self.vc.calendar_chips(store, [], date(2026, 9, 1), deliveries)
        only = self.vc.delivery_chips(date(2026, 9, 1), deliveries)
        self.assertIn("todo", {c["kind"] for c in visit["2026-09-16"]})
        self.assertNotIn("2026-09-16", only)
        kinds = {c["kind"] for c in only["2026-09-10"]}
        self.assertEqual(kinds, {"bulk", "other"})
        self.assertNotIn("visit", kinds)

    def test_t2d_delivery_rows_skip_full_staff_map(self):
        df = pd.DataFrame(
            {
                "담당자": ["김혁수", "김혁수", "김혁수"],
                "거래처": ["한국메티슨특수가스", "한국메티슨특수가스", "한국메티슨특수가스"],
                "매출일_dt": [
                    pd.Timestamp("2026-09-10"),
                    pd.Timestamp("2026-09-04"),
                    pd.Timestamp("2026-08-20"),
                ],
                "품목명": ["N2 (kg, Bulk)", "아세틸렌", "CO2 (kg, Bulk)"],
                "출고량": [1200.0, 2.0, 800.0],
                "매출액": [1000.0, 50.0, 400.0],
            }
        )
        rows = self.vc._t2d_delivery_rows(df, "김혁수", "한국메티슨특수가스")
        self.assertEqual(len(rows), 3)
        self.assertTrue(any(r.get("bulk") for r in rows))
        self.assertFalse(any(r["date"] == "2026-09-16" for r in rows))
        sep = self.vc._t2d_delivery_rows(
            df, "김혁수", "한국메티슨특수가스", month=date(2026, 9, 1)
        )
        self.assertEqual(len(sep), 2)
        self.assertTrue(all(str(r["date"]).startswith("2026-09-") for r in sep))
        aug = self.vc._t2d_delivery_rows(
            df, "김혁수", "한국메티슨특수가스", month=date(2026, 8, 1)
        )
        self.assertEqual(len(aug), 1)
        self.assertEqual(aug[0]["date"], "2026-08-20")

    def test_tab2_delivery_status_uses_main_filters(self):
        with open(self.vc.__file__, encoding="utf-8") as f:
            src = f.read()
        fn = src[src.index("def render_tab2_delivery_status") :]
        strip = src[src.index("def _render_t2d_day_strip") : src.index("def render_tab2_delivery_status")]
        self.assertIn("납품현황", fn)
        self.assertIn("delivery_chips", fn)
        self.assertIn('key="t2d_strip_host"', strip)
        self.assertIn('"hideNames": True', strip)
        self.assertIn('heading="납품 내역"', fn)
        self.assertIn("_t2d_delivery_rows", fn)
        self.assertIn("month=month", fn)
        self.assertNotIn("_sales_delivery_rows(df, staff, client)", fn)
        self.assertNotIn('selectbox("담당자"', fn)
        self.assertNotIn('selectbox("거래처"', fn)
        self.assertNotIn('key="vc_staff"', fn)
        self.assertNotIn('key="vc_client"', fn)
        self.assertNotIn("_render_todo_panel", fn)
        self.assertNotIn("_render_month_schedule", fn)
        with open("app.py", encoding="utf-8") as f:
            app = f.read()
        tab2 = app.split("# Tab 2:", 1)[1].split("with tab3:", 1)[0]
        self.assertIn("render_tab2_delivery_status", tab2)
        self.assertIn("df_client_filtered", tab2[tab2.find("render_tab2_delivery_status") - 180 : tab2.find("render_tab2_delivery_status")])
        self.assertLess(tab2.find("render_tab2_delivery_status"), tab2.find("tab2_action_btns"))
        self.assertNotIn("_t2_df = full_df", tab2)

    def test_is_bulk_item(self):
        self.assertTrue(self.vc._is_bulk_item("N2 (kg, Bulk)"))
        self.assertFalse(self.vc._is_bulk_item("아세틸렌"))

    def test_strip_tag_bulk_cylinder_visit_todo(self):
        self.assertEqual(self.vc._strip_tag([{"kind": "bulk"}]), "벌크")
        self.assertEqual(self.vc._strip_tag([{"kind": "other"}]), "실린더")
        self.assertEqual(self.vc._strip_tag([{"kind": "bulk"}, {"kind": "other"}]), "벌·실")
        self.assertEqual(self.vc._strip_tag([{"kind": "visit"}]), "")
        self.assertEqual(self.vc._strip_tag([{"kind": "todo"}]), "할일")
        self.assertEqual(self.vc._strip_tag([{"kind": "worklog"}]), "일지")
        self.assertEqual(self.vc._strip_tag([{"kind": "todo"}, {"kind": "bulk"}]), "벌크")

    def test_strip_days_payload_marks_selected_and_visit_name(self):
        chips = {
            "2026-09-18": [
                {"kind": "visit", "label": "한국메디손톡"},
                {"kind": "bulk", "label": "N2 10"},
            ]
        }
        days = self.vc._strip_days_payload(
            date(2026, 9, 1), chips, date(2026, 9, 18), date(2026, 9, 17)
        )
        cell = days[17]
        self.assertEqual(cell["iso"], "2026-09-18")
        self.assertTrue(cell["sel"])
        self.assertEqual(cell["kind"], "bulk")
        self.assertEqual(cell["name"], "한국메디손톡")
        self.assertTrue(days[16]["today"])
        self.assertEqual(self.vc._strip_tag([]), "")
        self.assertEqual(self.vc._client_short("라콜 주식회사(구.신정우)"), "라콜")
        self.assertEqual(
            self.vc._strip_visit_name([{"kind": "visit", "label": "라콜 주식회사(구.신정우)"}]),
            "라콜",
        )
        self.assertEqual(self.vc._strip_visit_name([{"kind": "bulk", "label": "N2"}]), "")

    def test_weekday_name_and_weekend_color(self):
        wed = date(2026, 9, 16)
        sat = date(2026, 9, 19)
        sun = date(2026, 9, 20)
        self.assertEqual(self.vc._weekday_name(wed), "수")
        self.assertEqual(self.vc._weekday_name(sat), "토")
        self.assertEqual(self.vc._weekday_name(sun), "일")
        self.assertEqual(self.vc._weekday_color(sat), self.vc._VC_SAT_FG)
        self.assertEqual(self.vc._weekday_color(sun), self.vc._VC_SUN_FG)
        self.assertEqual(self.vc._weekday_color(wed), "#6b7280")

    def test_purge_drops_day_button_keys(self):
        class _SS(dict):
            pass

        ss = _SS()
        ss["vc_day_2026-09-17"] = True
        ss["_vc_selected"] = date(2026, 9, 17)
        ss["_dash_bak_visit"] = {"vc_day_2026-09-17": True, "_vc_selected": date(2026, 9, 17)}
        with patch.object(self.vc.st, "session_state", ss):
            self.vc._purge_vc_button_keys()
        self.assertNotIn("vc_day_2026-09-17", ss)
        self.assertNotIn("vc_day_2026-09-17", ss["_dash_bak_visit"])
        self.assertEqual(ss["_vc_selected"], date(2026, 9, 17))

    def test_cached_history_skips_second_merge(self):
        df = pd.DataFrame(
            {
                "담당자": ["김혁수"],
                "거래처": ["한신테크"],
                "매출일_dt": [date(2026, 8, 1)],
            }
        )
        with patch.object(self.vc, "merge_visit_history", wraps=self.vc.merge_visit_history) as merged:
            a = self.vc._cached_history(df, "김혁수", "한신테크")
            b = self.vc._cached_history(df, "김혁수", "한신테크")
        self.assertEqual(a, b)
        self.assertEqual(merged.call_count, 1)

    def test_company_key_matches_suffix(self):
        self.assertEqual(self.vc._company_key("㈜한신테크"), self.vc._company_key("한신테크"))

    def test_ui_and_app_tab_wired(self):
        with open(self.vc.__file__, encoding="utf-8") as f:
            src = f.read()
        self.assertIn("할일 목록", src)
        self.assertIn('placeholder="거래처명 입력"', src)
        self.assertIn("accept_new_options=True", src)
        self.assertNotIn("거래처를 선택하면 방문일이 표시됩니다", src)
        self.assertIn("업체명", src)
        self.assertIn("세부사항", src)
        self.assertIn("방문 내역", src)
        self.assertIn("기방문", src)
        self.assertIn("def month_visit_rows", src)
        self.assertIn("방문예정", src)
        self.assertIn("vc_plan_chk_", src)
        self.assertIn("vc_visit_del_", src)
        self.assertIn("def clear_day_visit", src)
        self.assertIn("숨긴 할일 보기", src)
        self.assertIn("vc_todo_show_hidden", src)
        self.assertIn("vc_todo_chk_", src)
        self.assertIn("vc_todo_pick_", src)
        self.assertIn("할 일 수정", src)
        self.assertIn("def update_todo", src)
        self.assertIn("vc_visit_chk_", src)
        self.assertIn("실린더", src)
        self.assertIn("vc_strip_", src)
        self.assertIn("_VC_SAT_BG", src)
        self.assertIn("_VC_SUN_BG", src)
        self.assertIn("def _strip_tag", src)
        self.assertIn("def _render_day_strip", src)
        self.assertIn("visit_day_strip_v1", src)
        self.assertIn('setStateValue("iso"', src)
        self.assertIn("on_iso_change=_on_strip_iso_change", src)
        self.assertIn("def _visit_day_block", src)
        self.assertIn("def _weekday_name", src)
        self.assertIn("vc-wd", src)
        self.assertNotIn("def _render_calendar", src)
        self.assertNotIn('key=f"vc_day_{iso}"', src)
        self.assertNotIn("def _render_visit_history", src)
        self.assertNotIn('key="vc_todo_radio"', src)
        self.assertIn("on_click=_on_pick_day", src)
        self.assertIn("@st.fragment", src)
        self.assertIn("scope=\"fragment\"", src)
        self.assertIn("def _purge_vc_button_keys", src)
        self.assertIn("st.columns([1, 1]", src)
        self.assertIn("calendar_chips", src)
        self.assertNotIn('st.session_state["vc_day_', src)
        with open("app.py", encoding="utf-8") as f:
            app = f.read()
        self.assertIn("📅 방문·할일", app)
        self.assertIn("visit_calendar_tab", app)
        self.assertIn("tab13", app)
        self.assertIn("min_tabs=13", app)
        self.assertIn('_DASH_VC_STATE_PREFIXES = ("_vc_",)', app)
        tab13 = app[app.index("with tab13:") : app.index("방문·할일 탭 오류")]
        self.assertIn("render_visit_calendar_tab", tab13)
        self.assertNotIn("_dash_defer_heavy_stub", tab13)
        self.assertIn("시작부터 펼침", tab13)
        mount = app[app.index("def _dash_should_defer_heavy_tab") : app.index("def _dash_defer_heavy_stub")]
        self.assertNotIn("_DASH_TAB_VISIT", mount)
        body = src[src.index("def _render_visit_body") : src.index("def _render_day_strip")]
        self.assertNotIn("_ensure_worklog_index()", body)
        self.assertNotIn("_sales_by_client(df, staff)", body)
        self.assertIn("_staff_from_store", body)
        self.assertIn("_staff_clients", body)
        self.assertIn("달력·방문·할일에 연동", src)

    def test_staff_lists_do_not_scan_all_staff_from_sales(self):
        df = pd.DataFrame({"담당자": ["홍길동"] * 200, "거래처": ["한신테크"] * 200})
        self.vc.add_visit(
            {"date": date(2026, 9, 17), "staff": "김혁수", "client": "이엔에이치", "status": "done"}
        )
        names = self.vc._staff_names(df)
        self.assertIn("김혁수", names)
        self.assertNotIn("홍길동", names)
        clients = self.vc._staff_clients(df, "김혁수")
        self.assertIn("이엔에이치", clients)
        self.assertNotIn("한신테크", clients)

    def test_staff_clients_include_all_sales_for_staff(self):
        df = pd.DataFrame(
            {
                "담당자": ["김혁수", "김혁수", "홍길동"],
                "거래처": ["한국메티슨특수가스", "한신테크", "다른곳"],
                "거래처_원본": ["한국메티슨특수가스", "한신테크", "다른곳"],
            }
        )
        self.vc.add_visit(
            {"date": date(2026, 9, 17), "staff": "김혁수", "client": "이엔에이치", "status": "done"}
        )
        clients = self.vc._staff_clients(df, "김혁수")
        self.assertIn("한국메티슨특수가스", clients)
        self.assertIn("한신테크", clients)
        self.assertIn("이엔에이치", clients)
        self.assertNotIn("다른곳", clients)


if __name__ == "__main__":
    unittest.main()
