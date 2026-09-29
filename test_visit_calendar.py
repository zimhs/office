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

    def test_month_schedule_cells_mark_visit_plan_and_prior(self):
        self.vc.add_visit(
            {"date": date(2026, 9, 7), "staff": "김혁수", "client": "한신테크", "status": "done"}
        )
        self.vc.add_visit(
            {
                "date": date(2026, 9, 23),
                "staff": "김혁수",
                "client": "라콜 주식회사(구.신정우)",
                "status": "planned",
            }
        )
        weeks = self.vc.month_cal_weeks(date(2026, 9, 1))
        self.assertEqual(len(weeks[0]), 7)
        self.assertEqual(weeks[0][0].weekday(), 6)
        self.assertEqual(weeks[0][6].weekday(), 5)
        days = [d for w in weeks for d in w]
        cells = self.vc.schedule_cells(
            self.vc.load_store(),
            days,
            "김혁수",
            [{"date": "2026-09-10", "source": "업무일지", "client": "한신테크"}],
            "한신테크",
        )
        kinds7 = {x["kind"] for x in cells["2026-09-07"]}
        kinds23 = {x["kind"] for x in cells["2026-09-23"]}
        kinds10 = {x["kind"] for x in cells["2026-09-10"]}
        self.assertIn("visit", kinds7)
        self.assertTrue(any(x.get("mine") for x in cells["2026-09-07"]))
        self.assertIn("planned", kinds23)
        self.assertNotIn("prior", kinds10)
        self.assertNotIn("기방문", str(cells))
        html = self.vc._mcal_chips_html(cells["2026-09-23"])
        self.assertIn("planned", html)
        self.assertIn("라콜", html)
        self.assertEqual(html.count("<span class='vc-mcal-slot"), 5)
        empty = self.vc._mcal_chips_html([])
        self.assertEqual(empty.count("vc-mcal-slot empty"), 5)
        six = self.vc._mcal_chips_html(
            [{"kind": "visit", "label": f"거래처{i}"} for i in range(6)]
        )
        self.assertEqual(six.count("<span class='vc-mcal-slot"), 5)
        self.assertIn("+1", six)
        grid = self.vc._mcal_head_html() + self.vc._mcal_month_html(
            date(2026, 9, 1),
            date(2026, 9, 23),
            date(2026, 9, 17),
            weeks,
            cells,
        )
        self.assertIn("vc-mcal-head", grid)
        self.assertIn("vc-mcal-table", grid)
        self.assertIn(">일</div>", grid)
        self.assertIn(">토</div>", grid)
        self.assertIn("vc-mcal-th sun", grid)
        self.assertIn("<td ", grid)
        lab = self.vc._mcal_button_label(date(2026, 9, 23), cells["2026-09-23"])
        self.assertTrue(lab.startswith("23\n"))
        self.assertEqual(lab.count("\n"), 5)
        self.assertIn("라콜", lab)

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
        days = self.vc._strip_days_payload(
            date(2026, 9, 1), only, date(2026, 9, 10), date(2026, 9, 17)
        )
        cell = next(x for x in days if x["iso"] == "2026-09-10")
        self.assertEqual(cell["name"], "")
        self.assertEqual(cell["mark"], "")
        self.assertIn(cell["tag"], ("벌·실", "벌크", "실린더"))

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
                "단가": [0.83, 25.0, 0.5],
            }
        )
        rows = self.vc._t2d_delivery_rows(df, "김혁수", "한국메티슨특수가스")
        again = self.vc._t2d_delivery_rows(df, "김혁수", "한국메티슨특수가스")
        self.assertEqual(len(rows), 3)
        self.assertEqual([r["date"] for r in rows], [r["date"] for r in again])
        self.assertTrue(any(r.get("bulk") for r in rows))
        by_item = {r["item"]: r for r in rows}
        self.assertAlmostEqual(by_item["아세틸렌"]["unit_price"], 25.0)
        self.assertAlmostEqual(by_item["N2 (kg, Bulk)"]["unit_price"], 0.83)
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
        self.assertIn("_vc_use_day_strip_component()", strip)
        self.assertIn("_render_strip_buttons(days, \"t2d_strip_\"", strip)
        self.assertNotIn("_strip_fallback_html(", strip)
        self.assertNotIn("if _VC_STRIP is not None:", strip)
        self.assertIn('"hideNames": True', strip)
        self.assertIn('heading="납품 내역"', fn)
        self.assertIn("show_unit_price=True", fn)
        self.assertIn("_t2d_delivery_rows", fn)
        self.assertIn("month=month", fn)
        self.assertIn("_apply_pending_t2d_strip_pick()", fn)
        self.assertIn("_apply_pending_t2d_month_nav()", fn)
        self.assertIn('if k == "t2d_strip_host"', src)
        pick = src[src.index("def _on_t2d_pick_day") : src.index("def _on_t2d_shift_month")]
        self.assertNotIn("_t2d_month", pick)
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
        ss["vc_mcal_2026-09-17"] = True
        ss["vc_mcal_date"] = date(2026, 9, 17)
        ss["vc_mcal_box"] = object()
        ss["vc_strip_host"] = {"iso": "2026-09-17"}
        ss["vc_mcal_host"] = {"iso": "2026-09-17"}
        ss["_vc_selected"] = date(2026, 9, 17)
        ss["_dash_bak_visit"] = {
            "vc_day_2026-09-17": True,
            "vc_mcal_date": date(2026, 9, 17),
            "vc_strip_host": {"iso": "2026-09-17"},
            "vc_mcal_host": {"iso": "2026-09-17"},
            "_vc_selected": date(2026, 9, 17),
        }
        with patch.object(self.vc.st, "session_state", ss):
            self.vc._purge_vc_button_keys()
        self.assertNotIn("vc_day_2026-09-17", ss)
        self.assertNotIn("vc_mcal_2026-09-17", ss)
        self.assertNotIn("vc_day_2026-09-17", ss["_dash_bak_visit"])
        self.assertIn("vc_mcal_date", ss)
        self.assertIn("vc_mcal_box", ss)
        self.assertIn("vc_mcal_date", ss["_dash_bak_visit"])
        self.assertIn("vc_strip_host", ss)
        self.assertIn("vc_mcal_host", ss)
        self.assertIn("vc_strip_host", ss["_dash_bak_visit"])
        self.assertIn("vc_mcal_host", ss["_dash_bak_visit"])
        self.assertEqual(ss["_vc_selected"], date(2026, 9, 17))
        self.assertFalse(self.vc._is_vc_purge_key("vc_mcal_date"))
        self.assertFalse(self.vc._is_vc_purge_key("vc_mcal_box"))
        self.assertFalse(self.vc._is_vc_purge_key("vc_mcal_iso"))
        self.assertFalse(self.vc._is_vc_purge_key("vc_mcal_host"))
        self.assertFalse(self.vc._is_vc_purge_key("vc_touch_slide"))
        self.assertTrue(self.vc._is_vc_purge_key("vc_mcal_2026-09-17"))

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
        strip_fn = src[src.index("def _render_day_strip") : src.index("def _month_row_html")]
        self.assertIn("_render_strip_buttons(days, \"vc_strip_\"", strip_fn)
        self.assertIn("_strip_button_theme_css", strip_fn)
        self.assertNotIn("pick_href=True", strip_fn)
        self.assertIn("_strip_fallback_html(days, selected.isoformat())", strip_fn)
        self.assertIn("_vc_is_touch_ui()", strip_fn)
        self.assertNotIn("_render_cloud_day_pick", strip_fn)
        self.assertNotIn("_render_week_pick_buttons", strip_fn)
        self.assertIn("def _vc_use_day_strip_component", src)
        self.assertIn("_vc_use_day_strip_component()", src)
        self.assertIn("def _vc_is_mac_local", src)
        mac_fn = src[src.index("def _vc_is_mac_local") : src.index("def _vc_is_streamlit_cloud")]
        self.assertIn("return _vc_is_darwin_local() and not _vc_is_touch_ui()", mac_fn)
        self.assertNotIn("return True", mac_fn)
        self.assertNotIn("sys.platform != \"darwin\"", src[src.index("def _vc_is_touch_ui") : src.index("def _vc_is_streamlit_cloud")])
        self.assertIn("def _vc_is_darwin_local", src)
        self.assertIn("st.fragment(_visit_day_block)()", src)
        self.assertNotIn("def _on_cloud_day_change", src)
        self.assertIn("def _vc_ensure_strip", src)
        self.assertIn("def _vc_clear_strip_hosts_if_unused", src)
        self.assertIn("streamlit.app", src)
        self.assertIn("visit_day_strip_v1", src)
        ensure = src[src.index("def _vc_ensure_strip") : src.index("def _on_strip_iso_change")]
        self.assertIn("_vc_is_darwin_local()", ensure)
        self.assertIn('setStateValue("iso"', src)
        self.assertIn("_apply_v2_iso_host", src)
        self.assertIn("def _on_strip_iso_change", src)
        self.assertNotIn("on_iso_change=_on_strip_iso_change", src[src.index("def _render_day_strip") : src.index("def _month_row_html")])
        self.assertNotIn("default={\"iso\": selected.isoformat()}", src[src.index("def _render_day_strip") : src.index("def _month_row_html")])
        self.assertIn("def _visit_day_block", src)
        self.assertIn("def month_cal_weeks", src)
        self.assertIn("_CAL_HEADERS", src)
        self.assertEqual(self.vc._CAL_HEADERS, ("일", "월", "화", "수", "목", "금", "토"))
        self.assertEqual(self.vc._CAL_FIRST, 6)
        self.assertIn('key="vc_mcal_box"', src)
        self.assertIn("st-key-vc_mcal_box", src)
        self.assertIn("border: 1.5px solid #9aa8bc", src)
        self.assertIn("vc-grid-stColumn-v3", src)
        self.assertIn("def _mcal_grid_css", src)
        self.assertIn("def _apply_pending_mcal_pick", src)
        self.assertIn("def _apply_pending_month_nav", src)
        self.assertIn("def _vc_day_payload", src)
        payload = src[src.index("def _vc_day_payload") : src.index("def _worklog_visit_dates")]
        self.assertIn("ck = (staff, client)", payload)
        self.assertNotIn("month.year", payload)
        mount_fn = src[src.index("def render_visit_calendar_tab") : src.index("def _render_visit_body")]
        self.assertNotIn("def _visit_body", mount_fn)
        self.assertIn("_render_visit_body(df, latest_update_str)", mount_fn)
        self.assertNotIn("_vc_inject_visit_tab_hold_script", src)
        self.assertNotIn("st.components.v1.html", src)
        self.assertIn("_apply_pending_mcal_pick()", mount_fn)
        self.assertIn("_apply_pending_strip_pick()", mount_fn)
        self.assertLess(
            mount_fn.index("_apply_pending_strip_pick()"),
            mount_fn.index("_purge_vc_button_keys()"),
        )
        self.assertLess(
            mount_fn.index("_apply_pending_mcal_pick()"),
            mount_fn.index("_purge_vc_button_keys()"),
        )
        self.assertIn("def _mcal_button_label", src)
        self.assertNotIn("disabled=out", src)
        pick = src[src.index("def _on_pick_day") : src.index("def _on_shift_month")]
        self.assertNotIn("_vc_month", pick)
        self.assertIn("#e8eaed", src)
        self.assertIn(".stColumn", src)
        self.assertIn("border-radius: 10px", src)
        self.assertIn("#eceff3", src)
        self.assertIn("min-height: 1.22rem", src)
        self.assertIn('gap="small"', src)
        self.assertIn("_MCAL_SLOTS = 5", src)
        self.assertIn("vc-mcal-slot", src)
        self.assertIn("text-align: center", src)
        self.assertNotIn('.stColumn:has(div[class*="st-key-vc_mcal_{iso}"]){{background:#1a73e8', src)
        self.assertIn("#fff3f1", src)
        self.assertIn("#eef4fc", src)
        self.assertIn("def schedule_cells", src)
        self.assertIn("def _render_month_cal", src)
        self.assertIn("vc_mcal_", src)
        cal = src[src.index("def _render_month_cal") : src.index("def _render_month_schedule")]
        self.assertIn("_render_touch_mcal_html(", cal)
        self.assertNotIn('key="vc_mcal_date"', cal)
        self.assertNotIn("pick_href=True", cal)
        self.assertNotIn("스케줄 날짜", cal)
        self.assertIn("_vc_use_mcal_component()", cal)
        self.assertIn("_vc_is_touch_ui()", cal)
        self.assertIn("_render_mcal_day_buttons(", cal)
        touch_fn = src[src.index("def _render_touch_mcal_html") : src.index("def _on_mcal_date_change")]
        self.assertIn("_mcal_month_html(", touch_fn)
        self.assertIn("pick_href=False", touch_fn)
        self.assertIn("st.slider(", touch_fn)
        self.assertNotIn("st.dataframe(", touch_fn)
        self.assertNotIn("st.selectbox(", touch_fn)
        self.assertIn("visit_month_cal_v1", src)
        self.assertIn("vc_mcal_host", src)
        btns = src[src.index("def _render_mcal_day_buttons") : src.index("def _mcal_head_html")]
        self.assertIn('key=f"vc_mcal_{iso}"', btns)
        self.assertNotIn("opacity:.35", btns)
        self.assertNotIn("d.month != month.month", btns)
        self.assertNotIn("_vc_use_mcal_date_input()", cal)
        self.assertNotIn("st.date_input(", cal)
        self.assertNotIn("on_change=_on_mcal_date_change", cal)
        self.assertNotIn("st.pills(", cal)
        self.assertNotIn('key="vc_mcal_iso"', cal)
        self.assertNotIn('key="vc_mcal_day"', cal)
        self.assertIn("def _apply_pending_strip_pick", src)
        self.assertIn("_apply_pending_strip_pick()", src)
        self.assertNotIn("def _render_week_pick_buttons", src)
        self.assertNotIn("def _sunday_week", src)
        self.assertIn("_vc_pick_href(iso)", src[src.index("def _strip_fallback_html") : src.index("def _render_month_cal")])
        t2d_html = src[src.index("def _render_t2d_day_strip") : src.index("def render_tab2_delivery_status")]
        self.assertIn("_render_strip_wd_row(days)", t2d_html)
        self.assertNotIn("pick_href=True", t2d_html)
        self.assertIn("로컬과 같이 fragment", src)
        self.assertIn("def _on_mcal_date_change", src)
        self.assertIn("def _sync_selected_from_mcal_widget", src)
        self.assertIn("def _vc_date_field", src)
        self.assertIn("_apply_v2_iso_host(\"vc_strip_host\")", src)
        self.assertIn("_apply_v2_iso_host(\"vc_mcal_host\")", src)
        self.assertIn("vc_client_q", src)
        self.assertIn("vc_light_todo_sel", src)
        self.assertIn("_month_schedule_items_html(done", src)
        self.assertIn("def _strip_fallback_html", src)
        self.assertIn("def _apply_query_day_pick", src)
        self.assertIn("vc-mcal-hit", src)
        self.assertNotIn("def week_days", src)
        self.assertNotIn("def week_schedule_cells", src)
        self.assertIn("def _weekday_name", src)
        self.assertIn("vc-wd", src)
        self.assertNotIn("def _render_calendar", src)
        self.assertNotIn('key=f"vc_day_{iso}"', src)
        self.assertNotIn("def _render_visit_history", src)
        self.assertNotIn('key="vc_todo_radio"', src)
        self.assertIn("on_click=on_pick", src)
        self.assertIn("_on_pick_day", src)
        self.assertIn("st.fragment(_visit_day_block)()", src)
        self.assertIn("st.fragment(_t2d_body)()", src)
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
        self.assertIn("탭 클릭 때 stub remount", tab13)
        self.assertIn("if not is_touch_ui() and not _is_streamlit_cloud()", tab13)
        self.assertIn("v2 재등록으로 방문탭이 안 끝난다", tab13)
        self.assertIn("heavy_indices=(9, 10, 11)", app)
        self.assertNotIn("heavy_indices=(9, 10, 11, 12)", app)
        self.assertIn("v2 재등록으로 방문탭이 안 끝난다", tab13)
        mount = app[app.index("def _dash_should_defer_heavy_tab") : app.index("def _dash_defer_heavy_stub")]
        self.assertNotIn("_DASH_TAB_VISIT", mount)
        body = src[src.index("def _render_visit_body") : src.index("def _render_day_strip")]
        self.assertNotIn("_ensure_worklog_index()", body)
        self.assertNotIn("_sales_by_client(df, staff)", body)
        self.assertIn("_staff_from_store", body)
        self.assertIn("_staff_clients", body)
        self.assertIn("달력·방문·할일에 연동", src)
        pre = src[src.index("def _render_visit_body") : src.index("def _visit_day_block")]
        self.assertNotIn('key="vc_prev_month"', pre)
        self.assertNotIn('key="vc_next_month"', pre)
        self.assertNotIn('key="vc_jump_today"', pre)
        self.assertNotIn('key="vc_staff"', pre)
        self.assertNotIn('key="vc_client"', pre)
        self.assertIn('setdefault("vc_strip_host"', pre)
        self.assertIn('setdefault("vc_mcal_host"', pre)
        block = src[src.index("def _visit_day_block") : src.index("def _client_short")]
        self.assertIn("_apply_pending_month_nav()", block)
        self.assertIn("_apply_pending_mcal_pick()", block)
        self.assertIn("_apply_v2_calendar_picks()", block)
        self.assertNotIn('_apply_v2_iso_host("vc_strip_host")', block)
        self.assertNotIn('_apply_v2_iso_host("vc_mcal_host")', block)
        self.assertNotIn("_pick_touch_visible_day", src)
        self.assertNotIn('key="vc_touch_day"', src)
        self.assertNotIn('key="vc_touch_mcal"', src)
        self.assertNotIn("on_select=_on_touch_mcal_select", src)
        self.assertNotIn("selection_mode=\"single-cell\"", src)
        self.assertIn('key="vc_touch_slide"', src)
        self.assertIn("st.slider(", src)
        touch = src[src.index("def _render_touch_mcal_html") : src.index("def _on_mcal_date_change")]
        self.assertNotIn("st.dataframe(", touch)
        self.assertNotIn("st.selectbox(", touch)
        self.assertIn("st-key-vc_visit_frag", src)
        self.assertIn('[data-testid="stElementContainer"]', src)
        self.assertIn(':has(> [data-testid="stLayoutWrapper"] > [class*="st-key-vc_visit_frag"])', src)
        self.assertIn('key="vc_next_month"', block)
        self.assertIn('key="vc_jump_today"', block)
        self.assertNotIn("on_click=_on_shift_month", block)
        self.assertNotIn("on_click=_on_jump_today", block)
        self.assertIn('key="vc_staff"', block)
        self.assertIn('key="vc_client"', block)
        self.assertIn("_vc_day_payload", block)
        self.assertIn("_render_day_strip", block)
        self.assertIn("delivery_chips(mon, month_deliveries)", block)
        self.assertNotIn("calendar_chips(store, history, mon, month_deliveries)", block)
        self.assertIn("_render_delivery_list", block)
        self.assertNotIn("show_unit_price=True", block)
        self.assertLess(block.index("left, right = st.columns"), block.index("_render_month_cal"))
        self.assertLess(block.index("_vc_delivery_col()"), block.index("_render_todo_panel"))
        self.assertLess(block.index("_render_todo_panel"), block.index("with right:"))
        self.assertLess(block.index("_render_month_cal"), block.index("_render_day_agenda"))
        self.assertLess(block.index("_render_day_agenda"), block.index("_render_month_schedule"))
        self.assertGreater(block.index("_render_day_agenda"), block.index("with right:"))
        self.assertIn("def _vc_is_touch_ui", src)
        self.assertIn("_vc_use_ipad_layout()", block)
        self.assertIn("pointer: coarse", src)
        self.assertIn("orientation: portrait", src)
        self.assertIn("orientation: landscape", src)
        self.assertIn("max-width: 850px", src)
        self.assertIn("min-width: 851px", src)
        self.assertIn("월간 달력을 전폭으로", block)

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

    def test_touch_slide_index_maps_visible_days(self):
        month = date(2026, 9, 1)
        weeks = self.vc.month_cal_weeks(month)
        vis = self.vc._touch_visible_days(month)
        self.assertEqual(len(vis), sum(len(w) for w in weeks))
        self.assertEqual(self.vc._touch_slide_index(month, date(2026, 9, 17)), vis.index(date(2026, 9, 17)))
        self.assertEqual(vis[0].month, 8)
        self.assertEqual(vis[-1].month, 10)
        self.assertEqual(self.vc._touch_slide_index(month, vis[-1]), len(vis) - 1)

    def test_v2_calendar_picks_clicked_host_not_stale_other(self):
        class _SS(dict):
            pass

        def _run(ss):
            with patch.object(self.vc.st, "session_state", ss), patch.object(
                self.vc, "_vc_is_darwin_local", return_value=True
            ), patch.object(self.vc, "_vc_is_touch_ui", return_value=False):
                self.vc._apply_v2_calendar_picks()

        strip_click = _SS()
        strip_click["_vc_selected"] = date(2026, 9, 10)
        strip_click["vc_strip_host"] = {"iso": "2026-09-17"}
        strip_click["vc_mcal_host"] = {"iso": "2026-09-10"}
        strip_click["_vc_v2_cal_ready"] = True
        strip_click["_vc_v2_strip_seen"] = "2026-09-10"
        strip_click["_vc_v2_mcal_seen"] = "2026-09-10"
        _run(strip_click)
        self.assertEqual(strip_click["_vc_selected"], date(2026, 9, 17))
        _run(strip_click)
        self.assertEqual(strip_click["_vc_selected"], date(2026, 9, 17))

        mcal_click = _SS()
        mcal_click["_vc_selected"] = date(2026, 9, 10)
        mcal_click["vc_strip_host"] = {"iso": "2026-09-10"}
        mcal_click["vc_mcal_host"] = {"iso": "2026-09-23"}
        mcal_click["_vc_v2_cal_ready"] = True
        mcal_click["_vc_v2_strip_seen"] = "2026-09-10"
        mcal_click["_vc_v2_mcal_seen"] = "2026-09-10"
        _run(mcal_click)
        self.assertEqual(mcal_click["_vc_selected"], date(2026, 9, 23))
        _run(mcal_click)
        self.assertEqual(mcal_click["_vc_selected"], date(2026, 9, 23))

        leftover = _SS()
        leftover["_vc_selected"] = date(2026, 9, 28)
        leftover["vc_strip_host"] = {"iso": "2026-09-10"}
        leftover["vc_mcal_host"] = {"iso": "2026-09-10"}
        _run(leftover)
        self.assertEqual(leftover["_vc_selected"], date(2026, 9, 28))

    def test_day_payload_skips_worklog_index(self):
        with patch.object(self.vc.st, "session_state", {}), patch.object(
            self.vc, "_build_worklog_client_index", side_effect=AssertionError("slow worklog scan")
        ):
            store, deliveries, history = self.vc._vc_day_payload(
                None, "김혁수", "대영가스상사", date(2026, 9, 1)
            )
        self.assertEqual(history, [])
        self.assertIsInstance(store, dict)

    def test_save_store_mirrors_to_drive_copy(self):
        drive = tempfile.TemporaryDirectory()
        self.addCleanup(drive.cleanup)
        with patch.object(self.vc, "_vc_drive_store_path", return_value=""):
            self.vc.add_visit(
                {"date": date(2026, 9, 29), "staff": "김혁수", "client": "에스엔케이", "status": "planned"}
            )
        self.assertEqual(self.vc._vc_drive_store_path(), "")
        dst = Path(drive.name) / "방문할일.json"
        with patch.object(self.vc, "_vc_drive_store_path", return_value=str(dst)):
            self.vc.add_visit(
                {"date": date(2026, 9, 30), "staff": "김혁수", "client": "엠케이러스", "status": "done"}
            )
        self.assertTrue(dst.is_file())
        self.assertEqual(dst.read_text(encoding="utf-8"), Path(self.store).read_text(encoding="utf-8"))
        self.assertIn("엠케이러스", dst.read_text(encoding="utf-8"))

    def test_visit_tab_open_on_cloud_and_ipad(self):
        for cloud, touch in ((True, False), (False, True), (True, True), (False, False)):
            with patch.object(self.vc, "_vc_is_streamlit_cloud", return_value=cloud), patch.object(
                self.vc, "_vc_is_touch_ui", return_value=touch
            ):
                self.assertFalse(self.vc._vc_tab_paused())
        css = self.vc._mcal_button_theme_css(
            date(2026, 9, 1), date(2026, 9, 29), date(2026, 9, 29), self.vc.month_cal_weeks(date(2026, 9, 1))
        )
        self.assertIn('st-key-vc_mcal_2"] button{align-items:flex-start', css)
        self.assertIn("white-space:pre-line", css)
        self.assertIn("text-align:center", css)
        src = Path(self.vc.__file__).read_text(encoding="utf-8")
        mount = src[src.index("def render_visit_calendar_tab") : src.index("def _render_visit_body")]
        self.assertLess(mount.index("_vc_tab_paused()"), mount.index("_render_visit_body("))
        self.assertIn("expanded=False", mount)

    def test_t2d_strip_buttons_on_cloud_mac_unchanged(self):
        src = Path(self.vc.__file__).read_text(encoding="utf-8")
        strip = src[src.index("def _render_t2d_day_strip") : src.index("def render_tab2_delivery_status")]
        self.assertLess(strip.index("_vc_use_day_strip_component()"), strip.index("_vc_is_darwin_local()"))
        self.assertNotIn("_strip_fallback_html", strip)
        self.assertIn('replace("st-key-vc_strip_", "st-key-t2d_strip_")', strip)
        body = src[src.index("def render_tab2_delivery_status") :]
        self.assertIn("if _vc_is_mac_local() or not _vc_is_darwin_local():", body)
        darwin_ret = strip.index("return", strip.index("if _vc_is_darwin_local():"))
        for rule in ("gap:3px", "min-width:0", ".vc-strip-html .vc-strip-wd{display:flex"):
            self.assertGreater(strip.index(rule), darwin_ret)

    def test_touch_buttons_flag_only_swaps_ipad_calendar(self):
        src = Path(self.vc.__file__).read_text(encoding="utf-8")
        strip = src[src.index("def _render_day_strip") : src.index("def _month_row_html")]
        self.assertIn("_vc_is_touch_ui() and not _vc_touch_buttons()", strip)
        cal = src[src.index("def _render_month_cal") : src.index("def _render_month_schedule")]
        self.assertIn("elif _vc_is_touch_ui() and not _vc_touch_buttons():", cal)
        self.assertIn("_render_mcal_day_buttons(", cal)
        self.assertLess(cal.index("_vc_use_mcal_component()"), cal.index("_vc_touch_buttons()"))
        self.assertIn('query_params.get("vc_btn"', src)

        class _SS(dict):
            pass

        class _QP(dict):
            pass

        with patch.object(self.vc.st, "session_state", _SS()), patch.object(self.vc.st, "query_params", _QP()):
            self.assertTrue(self.vc._vc_touch_buttons())
        with patch.object(self.vc.st, "session_state", _SS()), patch.object(
            self.vc.st, "query_params", _QP(vc_btn="0")
        ):
            self.assertFalse(self.vc._vc_touch_buttons())

    def test_mcal_slot_click_picks_client_with_shade(self):
        marks = [
            {"kind": "visit", "label": "에스엔케이", "mine": False},
            {"kind": "planned", "label": "엠케이러스", "mine": True},
            {"kind": "prior", "label": "기방문", "mine": False},
        ]
        slots = self.vc._mcal_slot_view(marks)
        self.assertEqual(slots[0]["client"], "에스엔케이")
        self.assertEqual(slots[1]["client"], "엠케이러스")
        self.assertTrue(slots[1]["pick"])
        self.assertFalse(slots[0]["pick"])
        self.assertEqual(slots[2]["client"], "")
        self.assertEqual(slots[3]["kind"], "empty")
        html = self.vc._mcal_chips_html(marks)
        self.assertIn("planned mine pick", html)
        src = Path(self.vc.__file__).read_text(encoding="utf-8")
        self.assertIn('setStateValue("client"', src)
        self.assertIn("data-client", src)
        self.assertIn(".vc-mcal-slot.pick", src)

        class _SS(dict):
            pass

        ss = _SS()
        ss["_vc_selected"] = date(2026, 9, 29)
        ss["vc_strip_host"] = {"iso": "2026-09-29"}
        ss["vc_mcal_host"] = {"iso": "2026-09-29", "client": "엠케이러스"}
        ss["_vc_v2_cal_ready"] = True
        ss["_vc_v2_strip_seen"] = "2026-09-29"
        ss["_vc_v2_mcal_seen"] = "2026-09-29"
        ss["_vc_v2_mcal_client_seen"] = ""
        with patch.object(self.vc.st, "session_state", ss), patch.object(
            self.vc, "_vc_is_darwin_local", return_value=True
        ), patch.object(self.vc, "_vc_is_touch_ui", return_value=False):
            self.vc._apply_v2_calendar_picks()
            self.assertEqual(ss["vc_client"], "엠케이러스")
            ss["vc_client"] = "다른곳"
            self.vc._apply_v2_calendar_picks()
            self.assertEqual(ss["vc_client"], "다른곳")

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
