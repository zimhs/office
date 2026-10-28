"""채권관리 부가세 별도/포함 토글 · 다른 탭 표기는 부가세 별도."""
from __future__ import annotations

import unittest

import pandas as pd


class DebtVatModeTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        import app as dash_app

        cls.app = dash_app

    def test_inclusive_keeps_raw_exclusive_divides(self):
        df = pd.DataFrame(
            {"거래처": ["A"], "구분": ["잔액"], "10월": [11000.0], "11월": [22000.0]}
        )
        inc = self.app._debt_apply_vat_mode(df, "inclusive")
        self.assertEqual(inc.loc[0, "10월"], 11000.0)
        exc = self.app._debt_apply_vat_mode(df, "exclusive")
        self.assertAlmostEqual(float(exc.loc[0, "10월"]), 10000.0)
        self.assertAlmostEqual(float(exc.loc[0, "11월"]), 20000.0)

    def test_october_matches_erp_sales_and_total(self):
        # 202610.csv 총합계: 매출액 1,802,398,597 · 합계액 1,982,638,456
        inc_oct = 1_982_638_456.0
        df = pd.DataFrame({"거래처": ["합"], "구분": ["매출"], "10월": [inc_oct]})
        inc = self.app._debt_apply_vat_mode(df, "inclusive")
        self.assertAlmostEqual(float(inc.loc[0, "10월"]), inc_oct)
        exc = self.app._debt_apply_vat_mode(df, "exclusive")
        self.assertAlmostEqual(float(exc.loc[0, "10월"]), 1_802_398_596.3636, places=2)

    def test_tab_labels_say_vat_exclusive(self):
        with open("app.py", encoding="utf-8") as f:
            src = f.read()
        self.assertIn('key="t5_vat_ex"', src)
        self.assertIn('key="t5_vat_in"', src)
        self.assertIn('key="t5_vat_row"', src)
        self.assertIn('key="t5_vat_col"', src)
        self.assertIn('key="t5_vat_upd"', src)
        self.assertIn("#E11D48", src)
        self.assertIn("justify-content: flex-end", src)
        self.assertIn("min-height: 30px", src)
        _btn = src[src.index('key="t5_vat_ex"') : src.index('key="t5_vat_in"') + 120]
        self.assertIn('width="content"', _btn)
        self.assertIn("부가세 별도", src)
        self.assertIn("부가세 포함", src)
        self.assertNotIn("상위 30위 거래처 월별 실적 (VAT포함, 만원)", src)
        self.assertNotIn("단위: 만원, VAT 포함", src)
        self.assertNotIn("매출액 (VAT 포함, 만원)", src)
        self.assertNotIn("만원·VAT포함", src)
        self.assertIn("매출액 (부가세 별도, 만원)", src)
        self.assertIn("품목별 매출(만원·부가세 별도)", src)
        vat_div = [ln for ln in src.splitlines() if "/ 1.1" in ln]
        self.assertEqual(len(vat_div), 1)
        self.assertIn("pd.to_numeric(out[c], errors=\"coerce\") / 1.1", src)
        self.assertNotIn("pd.to_numeric(out[c], errors=\"coerce\") * 1.1", src)

    def test_yearly_monthly_pivot_is_exclusive_manwon(self):
        df = pd.DataFrame(
            {"월": ["1월"], "연도": ["2026"], "매출액": [20000.0]}
        )
        pvt = self.app._yearly_monthly_pivot(df, ["1월"], ["2026"])
        self.assertAlmostEqual(float(pvt.loc["1월", "2026"]), 2.0)

    def test_tab2_ytd_uses_current_year_like_tab1(self):
        with open("app.py", encoding="utf-8") as f:
            src = f.read()
        t2 = src[src.index("# 당해년도 누적 매출 · 부가세 별도 (영업종합요약과 동일 기준"):]
        self.assertIn('df_client_filtered["연도"].astype(str) == _cur_year_c', t2)
        self.assertIn("_tot_label_c", t2)


if __name__ == "__main__":
    unittest.main()
