"""분기 사업분석 — 라벨 계산·엑셀 반영(레이아웃·수식 유지)."""
from __future__ import annotations

import os
import shutil
import tempfile
import unittest
from datetime import date

import pandas as pd

import quarterly_analysis_tab as qa


class QuarterlyAnalysisTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="qa_test_")
        self.addCleanup(shutil.rmtree, self.tmp, ignore_errors=True)
        self.template = os.path.join(self.tmp, "template.xlsx")
        self.work = os.path.join(self.tmp, "work.xlsx")
        src = qa.QA_TEMPLATE
        if not os.path.isfile(src):
            self.skipTest(f"template missing: {src}")
        shutil.copy2(src, self.template)
        shutil.copy2(src, self.work)
        self._orig = (qa.QA_DIR, qa.QA_TEMPLATE, qa.QA_WORK)
        qa.QA_DIR = self.tmp
        qa.QA_TEMPLATE = self.template
        qa.QA_WORK = self.work

        def _restore():
            qa.QA_DIR, qa.QA_TEMPLATE, qa.QA_WORK = self._orig

        self.addCleanup(_restore)

    def _sample_df(self) -> pd.DataFrame:
        rows = []
        # 2026-01 sales ~ 100백만, bulk O2 10톤, cylinder 5병
        rows.append(
            {
                "매출일_dt": date(2026, 1, 10),
                "품목명": "O2 (kg, Bulk)",
                "출고량": 10000,
                "매출액": 50_000_000,
            }
        )
        rows.append(
            {
                "매출일_dt": date(2026, 1, 11),
                "품목명": "CO2 (kg, Bulk)",
                "출고량": 5000,
                "매출액": 50_000_000,
            }
        )
        rows.append(
            {
                "매출일_dt": date(2026, 1, 12),
                "품목명": "O2 (공업용, 40L)",
                "출고량": 5,
                "매출액": 100_000,
            }
        )
        # prior year
        rows.append(
            {
                "매출일_dt": date(2025, 1, 10),
                "품목명": "O2 (kg, Bulk)",
                "출고량": 8000,
                "매출액": 80_000_000,
            }
        )
        return pd.DataFrame(rows)

    def test_compute_sales_and_qty(self):
        goals = {"E6": 1500, "G6": 1400, "I6": 1500}
        vals = qa.compute_dashboard_values(
            self._sample_df(), year=2026, months=(1, 2, 3), goals=goals
        )
        self.assertEqual(vals["E6"], 1500)  # 목표 직접입력
        self.assertEqual(vals["E7"], 100)  # 실적 1월 백만
        self.assertEqual(vals["E9"], 80)  # 전년 1월
        self.assertEqual(vals["D21"], 10)  # L-O2 당해 1월 TON
        self.assertEqual(vals["E21"], 8)  # L-O2 전년 1월
        self.assertEqual(vals["D27"], 5)  # L-O2~LPG 외 → 실린더
        self.assertTrue(str(vals["E10"]).startswith("▲") or str(vals["E10"]).startswith("▼"))

    def test_non_bulk_goes_to_cylinder(self):
        df = pd.DataFrame(
            [
                {
                    "매출일_dt": date(2026, 1, 1),
                    "품목명": "O2 (kg, Bulk)",
                    "출고량": 1000,
                    "매출액": 1,
                },
                {
                    "매출일_dt": date(2026, 1, 1),
                    "품목명": "N2 (LGC, 175L)",
                    "출고량": 12,
                    "매출액": 1,
                },
                {
                    "매출일_dt": date(2026, 1, 1),
                    "품목명": "He (UHP, 47L)",
                    "출고량": 7,
                    "매출액": 1,
                },
            ]
        )
        vals = qa.compute_dashboard_values(
            df, year=2026, months=(1, 2, 3), goals={"E6": 1, "G6": 1, "I6": 1}
        )
        self.assertEqual(vals["D21"], 1)  # bulk O2 only
        # LGC·He 등 5종 외는 전부 실린더
        self.assertEqual(vals["D27"], 19)

    def test_apply_keeps_formulas(self):
        before = qa._qa_read_cells(self.work, ["K6", "D26", "E7"])
        self.assertTrue(str(before["K6"]).startswith("="))
        self.assertTrue(str(before["D26"]).startswith("="))
        dash = qa.compute_dashboard_values(
            self._sample_df(),
            year=2026,
            months=(1, 2, 3),
            goals={"E6": 1538, "G6": 1371, "I6": 1525},
        )
        n = qa._qa_write_cells(self.work, dash)
        self.assertGreater(n, 0)
        after = qa._qa_read_cells(self.work, ["K6", "D26", "E7", "E6"])
        self.assertEqual(after["K6"], before["K6"])
        self.assertEqual(after["D26"], before["D26"])
        self.assertEqual(after["E7"], 100)
        self.assertEqual(after["E6"], 1538)

    def test_quarter_scroll_options(self):
        opts = qa._qa_quarter_options(start_year=2025, end_year=2026)
        self.assertIn("2025년 1분기", opts)
        self.assertIn("2026년 4분기", opts)
        self.assertEqual(qa._qa_parse_quarter_label("2026년 2분기"), (2026, 2))
        self.assertEqual(qa._QA_QUARTER_MONTHS[2], (4, 5, 6))

    def test_goal_and_perf_table_html(self):
        goal_html = qa._qa_goal_table_html(
            2026,
            (1, 2, 3),
            (1538, 1371, 1525),
            (1429, 1135, 1584),
            (1152, 1261, 1343),
            ("▼7%", "▼17%", "▲4%", "▼6%"),
            ("▲24%", "▼9%", "▲17%", "▲10%"),
        )
        self.assertIn("■ 목표 / 달성", goal_html)
        self.assertIn("1,538", goal_html)
        self.assertIn("▼7%", goal_html)
        self.assertIn("전년 매출", goal_html)
        dash = qa.compute_dashboard_values(
            self._sample_df(),
            year=2026,
            months=(1, 2, 3),
            goals={"E6": 1500, "G6": 1400, "I6": 1500},
        )
        rows = qa._qa_build_perf_rows(dash)
        labels = [r["label"] for r in rows]
        self.assertEqual(labels[:5], ["L-O2", "L-N2", "L-AR", "L-CO2", "LPG"])
        self.assertIn("Bulk 합계", labels)
        self.assertIn("Gas Cylinder", labels)
        perf_html = qa._qa_perf_table_html(2026, (1, 2, 3), rows)
        self.assertIn("■ 실적 / 분석", perf_html)
        self.assertIn("L-O2", perf_html)
        self.assertIn("qa-bulk", perf_html)

    def test_ui_has_quarter_buttons(self):
        with open(qa.__file__, encoding="utf-8") as f:
            src = f.read()
        self.assertIn('f"{i}분기"', src)
        self.assertIn("qa_qbtn_", src)
        self.assertIn("_qa_goal_table_html", src)
        self.assertIn("_qa_perf_table_html", src)
        self.assertIn("목표 입력", src)

    def test_app_wires_tab(self):
        with open("app.py", encoding="utf-8") as f:
            app = f.read()
        self.assertIn("📊 분기 사업분석", app)
        self.assertIn("tab14", app)
        self.assertIn("render_quarterly_analysis_tab", app)
        self.assertIn("quarterly_analysis_tab", app)
        self.assertIn("_DASH_TAB_QUARTERLY", app)


if __name__ == "__main__":
    unittest.main()
