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

    def test_gascocoasan_bigo_splits_to_subclients(self):
        csv_text = """영업담당자,거래처,매출일,품목명,출고량,단가,매출액,비고
담당자없음,가스코아산.,08/04,"CO2 (kg, Bulk)",100,220,22000,동희
담당자없음,가스코아산.,08/05,"N2 (kg, Bulk)",50,200,10000,성우하이텍
담당자없음,가스코아산.,08/06,"AR (kg, Bulk)",10,300,3000,
담당자없음,가스코아산.,08/07,"O2 (kg, Bulk)",1,100,100,미검공병출고
"""
        sales_df = load_uploaded_files_from_bytes([("202608.csv", csv_text.encode("utf-8-sig"))])
        clients = set(sales_df["거래처"].astype(str))
        self.assertIn("가스코아산(동희)", clients)
        self.assertIn("가스코아산(성우하이텍)", clients)
        self.assertIn("가스코아산.", clients)  # 비고 없음·업무메모는 부모 유지
        self.assertNotIn("가스코아산(미검공병출고)", clients)
        sub = sales_df[sales_df["거래처"] == "가스코아산(동희)"]
        self.assertEqual(len(sub), 1)
        self.assertEqual(sub.iloc[0]["담당자"], "가스코아산")
        self.assertEqual(float(sub.iloc[0]["출고량"]), 100.0)

    def test_other_clients_bigo_also_splits(self):
        csv_text = """영업담당자,거래처,매출일,품목명,출고량,단가,매출액,비고
김담당,태광산업가스,08/04,"CO2 (kg, Bulk)",80,220,17600,영신쿼츠
김담당,태광산업가스,08/05,"N2 (kg, Bulk)",20,200,4000,3,000kg
김담당,디에스텍,08/06,"AR (kg, Bulk)",15,300,4500,델리치푸드
김담당,두산판금,08/07,"O2 (kg, Bulk)",5,100,500,11월 거래분
"""
        sales_df = load_uploaded_files_from_bytes([("202608.csv", csv_text.encode("utf-8-sig"))])
        clients = set(sales_df["거래처"].astype(str))
        self.assertIn("태광산업가스(영신쿼츠)", clients)
        self.assertIn("디에스텍(델리치푸드)", clients)
        # kg·거래분 메모는 분리하지 않음
        self.assertIn("태광산업가스", clients)
        self.assertIn("두산판금", clients)
        self.assertNotIn("태광산업가스(3,000kg)", clients)
        self.assertNotIn("두산판금(11월 거래분)", clients)
        sub = sales_df[sales_df["거래처"] == "태광산업가스(영신쿼츠)"]
        self.assertEqual(len(sub), 1)
        self.assertEqual(float(sub.iloc[0]["출고량"]), 80.0)

    def test_parent_selection_aggregates_subclients(self):
        csv_text = """영업담당자,거래처,매출일,품목명,출고량,단가,매출액,비고
김담당,가스코아산.,08/04,"CO2 (kg, Bulk)",100,220,22000,동희
김담당,가스코아산.,08/05,"N2 (kg, Bulk)",50,200,10000,성우하이텍
김담당,가스코아산.,08/06,"AR (kg, Bulk)",10,300,3000,
"""
        sales_df = load_uploaded_files_from_bytes([("202608.csv", csv_text.encode("utf-8-sig"))])
        self.assertIn("거래처_원본", sales_df.columns)
        parent = filter_df_by_selected_client(sales_df, "가스코아산.")
        child = filter_df_by_selected_client(sales_df, "가스코아산(동희)")
        self.assertEqual(len(parent), 3)  # 종속 2 + 부모잔여 1
        self.assertEqual(float(parent["출고량"].sum()), 160.0)
        self.assertEqual(len(child), 1)
        self.assertEqual(float(child["출고량"].sum()), 100.0)
        opts = list_filter_client_options(sales_df)
        self.assertIn("가스코아산.", opts)
        self.assertIn("가스코아산(동희)", opts)
        self.assertIn("가스코아산(성우하이텍)", opts)


class IndustryStaffMapTest(unittest.TestCase):
    def test_poem_maps_to_kim_from_industry_csv(self):
        """업체대분류 영업담당자: 포엠주식회사 → 김혁수 (월별 매출에 담당자 열 없어도)."""
        ind_path = os.path.join(os.path.dirname(__file__), "uploaded_cache", "industry.csv")
        if not os.path.exists(ind_path):
            ind_path = os.path.join(os.path.dirname(__file__), "uploaded_cache", "industry.csv")
        if not os.path.exists(ind_path):
            self.skipTest("업체대분류.csv 없음")
        with open(ind_path, "rb") as f:
            staff_map = load_industry_staff_map(f.read())
        self.assertEqual(staff_map.get("포엠주식회사"), "김혁수")
        df = pd.DataFrame(
            {
                "거래처": ["포엠주식회사", "포엠주식회사"],
                "담당자": ["미지정", "미지정"],
            }
        )
        out = _apply_industry_staff_mapping(df, staff_map)
        self.assertTrue((out["담당자"] == "김혁수").all())
        # 이미 지정된 담당자는 덮지 않음
        kept = _apply_industry_staff_mapping(
            pd.DataFrame({"거래처": ["포엠주식회사"], "담당자": ["박철수"]}),
            staff_map,
        )
        self.assertEqual(kept.iloc[0]["담당자"], "박철수")


if __name__ == "__main__":
    unittest.main()
