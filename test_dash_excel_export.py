"""사이드바 엑셀: 탭·라벨 선택, 대시보드 그라디언트 색 유지."""
from __future__ import annotations

import io
import unittest

import pandas as pd


class DashExcelExportSourceTest(unittest.TestCase):
    def _excel_block(self) -> str:
        with open("app.py", encoding="utf-8") as f:
            src = f.read()
        start = src.index('st.sidebar.subheader("📥 엑셀 내보내기")')
        end = src.index('st.sidebar.subheader("📽️ PPT 내보내기")')
        return src[start:end]

    def test_sidebar_picks_tab_and_label(self):
        excel = self._excel_block()
        self.assertIn("탭 선택", excel)
        self.assertIn("라벨 선택", excel)
        self.assertIn("dash_excel_tabs", excel)
        self.assertIn("dash_excel_labels", excel)
        self.assertIn("convert_dash_excel_sheets", excel)
        self.assertIn("그라디언트", excel)
        self.assertNotIn("전체 분석 시트별 엑셀 다운로드", excel)

    def test_catalog_covers_existing_sheets(self):
        excel = self._excel_block()
        for label in (
            "연도별 월 매출 (만원)",
            "거래처 품목별 사용량",
            "품목별 매출액 (만원)",
            "품목별 출고량",
            "품목별 적용단가",
            "담당자별 매출 (만원)",
            "상세 거래 내역",
            "채권 관리 현황",
        ):
            self.assertIn(label, excel)


class DashExcelConvertTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        import app as dash_app

        cls.app = dash_app

    def test_hex_and_safe_sheet(self):
        self.assertEqual(self.app._dash_excel_hex("#6baed6"), "6BAED6")
        self.assertEqual(self.app._dash_excel_hex("rgb(107, 174, 214)"), "6BAED6")
        used = set()
        a = self.app._dash_excel_safe_sheet("T1_연도별월매출", used)
        b = self.app._dash_excel_safe_sheet("T1_연도별월매출", used)
        self.assertEqual(a, "T1_연도별월매출")
        self.assertNotEqual(a, b)
        self.assertTrue(b.startswith("T1_연도별월매출"))

    def test_selected_sheet_keeps_heatmap_and_sum(self):
        from openpyxl import load_workbook

        df = pd.DataFrame(
            {"2025": [10.0, 50.0, 200.0], "2026": [1.0, 80.0, 300.0]},
            index=["1월", "2월", "3월"],
        )
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
        self.assertEqual(wb.sheetnames, ["T1_연도별월매출"])
        ws = wb["T1_연도별월매출"]
        self.assertEqual(ws.cell(1, 1).value, "구분")
        self.assertEqual(ws.cell(1, 1).fill.fgColor.rgb[-6:].upper(), "1E3A8A")
        self.assertEqual(ws.cell(5, 1).value, "연간 합계")
        self.assertEqual(ws.cell(5, 2).value, 260)
        self.assertEqual(ws.cell(5, 2).fill.fgColor.rgb[-6:].upper(), "E2E8F0")
        fills = {
            (ws.cell(r, 2).fill.fgColor.rgb or "")[-6:].upper()
            for r in (2, 3, 4)
        }
        self.assertGreaterEqual(len(fills), 2)
        self.assertTrue(all(len(c) == 6 and c != "000000" for c in fills))

    def test_only_picked_labels_are_written(self):
        from openpyxl import load_workbook

        a = pd.DataFrame({"x": [1]}, index=["1월"])
        b = pd.DataFrame({"y": [2]}, index=["1월"])
        raw = self.app.convert_dash_excel_sheets(
            [
                {"sheet": "T1_연도별월매출", "df": a, "use_index": True, "cmap": "Blues"},
                {"sheet": "T3_품목매출", "df": b, "use_index": True, "cmap": "Blues"},
            ]
        )
        wb = load_workbook(io.BytesIO(raw))
        self.assertEqual(wb.sheetnames, ["T1_연도별월매출", "T3_품목매출"])
