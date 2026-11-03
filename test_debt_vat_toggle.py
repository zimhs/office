"""채권관리 부가세 별도/포함 토글 · 매출 탭은 별도 원금액에 포함 시 ×1.1."""
from __future__ import annotations

import unittest
from unittest.mock import patch

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
        inc_oct = 1_982_638_456.0
        df = pd.DataFrame({"거래처": ["합"], "구분": ["매출"], "10월": [inc_oct]})
        inc = self.app._debt_apply_vat_mode(df, "inclusive")
        self.assertAlmostEqual(float(inc.loc[0, "10월"]), inc_oct)
        exc = self.app._debt_apply_vat_mode(df, "exclusive")
        self.assertAlmostEqual(float(exc.loc[0, "10월"]), 1_802_398_596.3636, places=2)

    def test_set_vat_mode_callback(self):
        with patch.object(self.app.st, "session_state", {}):
            self.app._set_dash_vat_mode("inclusive")
            self.assertEqual(self.app.st.session_state["dash_vat_mode"], "inclusive")
            self.app._set_dash_vat_mode("exclusive")
            self.assertEqual(self.app.st.session_state["dash_vat_mode"], "exclusive")

    def test_sales_scale_only_when_inclusive(self):
        with patch.object(self.app.st, "session_state", {"dash_vat_mode": "exclusive"}):
            self.assertEqual(self.app._sales_vat_mult(), 1.0)
            self.assertAlmostEqual(float(self.app._scale_sales_money(180239.8597)), 180239.8597)
        with patch.object(self.app.st, "session_state", {"dash_vat_mode": "inclusive"}):
            self.assertEqual(self.app._sales_vat_mult(), 1.1)
            self.assertAlmostEqual(float(self.app._scale_sales_money(180239.8597)), 198263.84567)

    def test_vat_chips_on_sales_tabs(self):
        with open("app.py", encoding="utf-8") as f:
            src = f.read()
        for prefix in ("t1", "t2", "t3", "t4", "t5"):
            self.assertIn(f'_render_vat_chips(latest_update_str, "{prefix}")' if prefix != "t5" else '_render_vat_chips(debt_update_str, "t5")', src)
        self.assertIn("def _tab1_vat_main", src)
        self.assertIn("on_click=_set_dash_vat_mode", src)
        self.assertIn('args=("exclusive",)', src)
        self.assertIn('args=("inclusive",)', src)
        self.assertIn("font-size: 13px !important", src[src.index("def _dash_vat_chip_css"): src.index("def _render_vat_chips")])
        self.assertIn("padding: 6px 12px !important", src[src.index("def _dash_vat_chip_css"): src.index("def _render_vat_chips")])
        self.assertIn("min-height: 32px !important", src[src.index("def _dash_vat_chip_css"): src.index("def _render_vat_chips")])
        self.assertIn("dash-vat-upd-badge", src)
        self.assertIn('key=f"{prefix}_vat_ex"', src)
        self.assertNotIn("pd.to_numeric(out[c], errors=\"coerce\") * 1.1", src)
        self.assertIn("pd.to_numeric(out[c], errors=\"coerce\") / 1.1", src)

    def test_yearly_monthly_pivot_is_exclusive_manwon(self):
        df = pd.DataFrame(
            {"월": ["1월"], "연도": ["2026"], "매출액": [20000.0]}
        )
        pvt = self.app._yearly_monthly_pivot(df, ["1월"], ["2026"])
        self.assertAlmostEqual(float(pvt.loc["1월", "2026"]), 2.0)

    def test_tab2_ytd_uses_current_year_like_tab1(self):
        with open("app.py", encoding="utf-8") as f:
            src = f.read()
        t2 = src[src.index("def _tab2_vat_main"):]
        self.assertIn('df_client_filtered["연도"].astype(str) == _cur_year_c', t2)
        self.assertIn("_tot_label_c", t2)

    def test_scale_skips_seq_and_scales_sales(self):
        df = pd.DataFrame({"순번": [1, 2], "1월": [100.0, 200.0]})
        with patch.object(self.app.st, "session_state", {"dash_vat_mode": "inclusive"}):
            out = self.app._scale_sales_money(df)
        self.assertEqual(int(out.loc[0, "순번"]), 1)
        self.assertAlmostEqual(float(out.loc[0, "1월"]), 110.0)

    def test_sales_money_labels_follow_vat_mode(self):
        with open("app.py", encoding="utf-8") as f:
            src = f.read()
        self.assertNotIn("1️⃣ 매출액 (부가세 별도, 만원)", src)
        self.assertIn("_scale_sales_money(pivot_m_client)", src)
        self.assertIn("_scale_sales_money(ranking_pivot)", src)
        self.assertIn("if metric == \"sales\":\n                pvt = _scale_sales_money(pvt)", src)
        self.assertIn("_scale_sales_money(triple[\"sales\"])", src)
        self.assertIn('view_detail_df["매출액"] = _scale_sales_money(view_detail_df["매출액"])', src)
        self.assertIn('unassigned_summary["총매출액_만원"] = _scale_sales_money(unassigned_summary["총매출액_만원"])', src)
        self.assertIn("def _dash_vat_lbl", src)


if __name__ == "__main__":
    unittest.main()
