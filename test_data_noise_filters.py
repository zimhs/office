"""채권/매출 CSV 노이즈 행 필터 검증."""
import io
import os
import unittest

import pandas as pd

# Streamlit 캐시/페이지 설정 없이 유틸 함수만 로드
with open(os.path.join(os.path.dirname(__file__), "app.py"), encoding="utf-8") as f:
    _src = f.read().split("# 5. 메인 실행 흐름")[0]
exec(_src, globals())


class DebtNoiseFilterTest(unittest.TestCase):
    def test_real_debt_file_has_no_export_footer_client(self):
        debt_path = os.path.join(os.path.dirname(__file__), "uploaded_cache", "debt.csv")
        if not os.path.exists(debt_path):
            self.skipTest("debt.csv 없음")

        with open(debt_path, "rb") as f:
            debt_df = load_debt_file(f.read())

        clients = debt_df["거래처"].astype(str)
        self.assertFalse(clients.str.match(r"^\d{4}-\d{2}-\d{2}\s", na=False).any())
        self.assertNotIn("2026-08-05 오후 5:49:25", set(clients))

    def test_synthetic_debt_footer_is_removed(self):
        csv_text = """거래처,구분,1월,2월
가스코아산(대창),이월,100,0
가스코아산(대창),잔액,100,0
2026-08-05 오후 5:49:25,총매출,,,
,총수금,,,
"""
        debt_df = load_debt_file(csv_text.encode("utf-8-sig"))
        self.assertEqual(debt_df["거래처"].nunique(), 1)
        self.assertEqual(debt_df.iloc[0]["거래처"], "가스코아산(대창)")


class SalesNoiseFilterTest(unittest.TestCase):
    def test_carryover_and_timestamp_rows_are_removed(self):
        csv_text = """,,,,,,
영업담당자,거래처,매출일,품목명,출고량,단가,매출액
담당자없음,가스코아산.,08/04,"CO2 (kg, Bulk)",100,220,22000
담당자없음,가스코아산(대창),,[이월미수액],,,330000
2026-08-06 오후 12:09:58,,,,,,
"""
        sales_df = load_uploaded_files_from_bytes([("20260805.csv", csv_text.encode("utf-8-sig"))])
        self.assertEqual(len(sales_df), 1)
        self.assertEqual(sales_df.iloc[0]["품목명"], "CO2 (kg, Bulk)")


if __name__ == "__main__":
    unittest.main()
