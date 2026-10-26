"""엑셀: 사이드바는 통합 다운로드, 표 라벨 옆에서 그라디언트 그대로 받기."""
from __future__ import annotations

import io
import unittest

import pandas as pd


class DashExcelExportSourceTest(unittest.TestCase):
    def _src(self) -> str:
        with open("app.py", encoding="utf-8") as f:
            return f.read()

    def test_sidebar_restored_without_tab_label_pickers(self):
        src = self._src()
        excel = src[src.index('st.sidebar.subheader("📥 엑셀 내보내기")') : src.index('st.sidebar.subheader("📽️ PPT 내보내기")')]
        self.assertIn("전체 분석 시트별 엑셀 다운로드", excel)
        self.assertNotIn("탭 선택", excel)
        self.assertNotIn("라벨 선택", excel)
        self.assertNotIn("dash_excel_tabs", excel)

    def test_label_titles_have_small_excel_button(self):
        src = self._src()
        self.assertIn("def _dash_label_with_xlsx", src)
        self.assertIn("def _dash_xlsx_from_styler", src)
        self.assertIn('key=f"dash_xlsx_{key}"', src)
        self.assertIn("_dash_label_with_xlsx(", src)
        self.assertIn("xlsx_name", src)


class DashExcelConvertTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        import app as dash_app

        cls.app = dash_app

    def test_styler_excel_keeps_heatmap(self):
        from openpyxl import load_workbook

        df = pd.DataFrame(
            {"2025": [10.0, 50.0, 200.0], "2026": [1.0, 80.0, 300.0]},
            index=["1월", "2월", "3월"],
        )
        disp = self.app.get_display_df_with_sum(df, "연간 합계")
        styled = self.app.style_with_sum(disp, "{:,.0f}", "Blues", axis=None)
        raw = self.app._dash_xlsx_from_styler(styled, sheet="연도별월매출")
        wb = load_workbook(io.BytesIO(raw))
        self.assertEqual(wb.sheetnames, ["연도별월매출"])
        ws = wb["연도별월매출"]
        self.assertEqual(ws.cell(1, 1).fill.fgColor.rgb[-6:].upper(), "1E3A8A")
        fills = {
            (ws.cell(r, 2).fill.fgColor.rgb or "")[-6:].upper()
            for r in range(2, 5)
        }
        self.assertGreaterEqual(len(fills), 2)

    def test_convert_sheets_still_writes_sum_gradient(self):
        from openpyxl import load_workbook

        df = pd.DataFrame({"2025": [10.0, 200.0]}, index=["1월", "2월"])
        raw = self.app.convert_dash_excel_sheets(
            [
                {
                    "sheet": "T1_연도별월매출",
                    "df": df,
                    "use_index": True,
                    "add_sum": True,
                    "sum_label": "연간 합계",
                    "cmap": "Blues",
                    "num_fmt": "#,##0",
                }
            ]
        )
        wb = load_workbook(io.BytesIO(raw))
        ws = wb["T1_연도별월매출"]
        self.assertEqual(ws.cell(4, 1).value, "연간 합계")
        self.assertEqual(ws.cell(4, 2).value, 210)
