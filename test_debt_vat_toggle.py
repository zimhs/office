"""채권관리 부가세 별도/포함 토글 · 다른 탭 표기는 부가세 별도."""
from __future__ import annotations

import unittest

import pandas as pd


class DebtVatModeTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        import app as dash_app

        cls.app = dash_app

    def test_exclusive_keeps_raw_inclusive_multiplies(self):
        df = pd.DataFrame(
            {"거래처": ["A"], "구분": ["잔액"], "10월": [11000.0], "11월": [22000.0]}
        )
        raw = self.app._debt_apply_vat_mode(df, "exclusive")
        self.assertEqual(raw.loc[0, "10월"], 11000.0)
        inc = self.app._debt_apply_vat_mode(df, "inclusive")
        self.assertAlmostEqual(float(inc.loc[0, "10월"]), 12100.0)
        self.assertAlmostEqual(float(inc.loc[0, "11월"]), 24200.0)

    def test_tab_labels_say_vat_exclusive(self):
        with open("app.py", encoding="utf-8") as f:
            src = f.read()
        self.assertIn('key="t5_vat_ex"', src)
        self.assertIn('key="t5_vat_in"', src)
        self.assertIn("부가세 별도", src)
        self.assertIn("부가세 포함", src)
        self.assertNotIn("상위 30위 거래처 월별 실적 (VAT포함, 만원)", src)
        self.assertNotIn("단위: 만원, VAT 포함", src)
        self.assertNotIn("매출액 (VAT 포함, 만원)", src)
        self.assertNotIn("만원·VAT포함", src)
        self.assertIn("매출액 (부가세 별도, 만원)", src)
        self.assertIn("품목별 매출(만원·부가세 별도)", src)
        vat_mult = [ln for ln in src.splitlines() if "* 1.1" in ln]
        self.assertEqual(len(vat_mult), 1)
        self.assertIn("pd.to_numeric(out[c], errors=\"coerce\") * 1.1", src)

    def test_yearly_monthly_pivot_is_exclusive_manwon(self):
        df = pd.DataFrame(
            {"월": ["1월"], "연도": ["2026"], "매출액": [20000.0]}
        )
        pvt = self.app._yearly_monthly_pivot(df, ["1월"], ["2026"])
        self.assertAlmostEqual(float(pvt.loc["1월", "2026"]), 2.0)


if __name__ == "__main__":
    unittest.main()
