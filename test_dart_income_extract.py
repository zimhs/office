"""DART 손익계산서에서 매출액·영업이익을 표 기준으로 읽는다."""
from __future__ import annotations

import ast
import unittest

import pandas as pd


_HELPERS = (
    "_fmt_dart_amount",
    "_fmt_dart_amount_with_year",
    "_parse_dart_money_cell",
    "_is_revenue_label",
    "_is_operating_profit_label",
    "_pick_current_amount",
    "_detect_amount_unit_mult",
    "_extract_income_amounts",
    "_year_from_audit_item",
    "_pick_fin_account",
    "_parse_dart_viewer_section",
)


def _load_helpers():
    with open("app.py", encoding="utf-8") as f:
        tree = ast.parse(f.read())
    nodes = [
        n
        for n in tree.body
        if isinstance(n, ast.FunctionDef) and n.name in _HELPERS
    ]
    names = {n.name for n in nodes}
    missing = set(_HELPERS) - names
    if missing:
        raise AssertionError(f"missing helpers: {missing}")
    ns: dict = {
        "re": __import__("re"),
        "pd": pd,
        "BeautifulSoup": __import__("bs4", fromlist=["BeautifulSoup"]).BeautifulSoup,
    }
    exec(compile(ast.Module(body=nodes, type_ignores=[]), "app.py", "exec"), ns)
    return ns


SAMPLE_TABLE = """
<p>(단위: 원)</p>
<table>
  <tr><th>과목</th><th>주석</th><th>제 27(당) 기</th><th>제 26(전) 기</th></tr>
  <tr><td>Ⅰ. 매출액</td><td>3</td><td>12,345,678,901</td><td>11,000,000,000</td></tr>
  <tr><td>Ⅱ. 매출원가</td><td>4</td><td>10,000,000,000</td><td>9,000,000,000</td></tr>
  <tr><td>Ⅴ. 영업이익</td><td></td><td>234,567,890</td><td>200,000,000</td></tr>
</table>
"""

NOTE_FIRST_TEXT = "매출액 3 12,345,678,901 11,000,000,000 영업이익 234,567,890"


class DartIncomeExtractTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.h = _load_helpers()

    def test_table_reads_current_year_not_note(self):
        found = self.h["_extract_income_amounts"](SAMPLE_TABLE)
        self.assertEqual(found["revenue"], "12,345,678,901 원")
        self.assertEqual(found["profit"], "234,567,890 원")

    def test_skips_cogs_and_note_number(self):
        self.assertFalse(self.h["_is_revenue_label"]("매출원가"))
        self.assertTrue(self.h["_is_revenue_label"]("Ⅰ. 매출액"))
        self.assertEqual(
            self.h["_pick_current_amount"](["3", "12,345,678,901", "11,000,000,000"]),
            12345678901,
        )

    def test_text_fallback_skips_note(self):
        found = self.h["_extract_income_amounts"](NOTE_FIRST_TEXT)
        self.assertEqual(found["revenue"], "12,345,678,901 원")
        self.assertEqual(found["profit"], "234,567,890 원")

    def test_cheonwon_unit(self):
        html = """
        <p>(단위: 천원)</p>
        <table>
          <tr><td>매출액</td><td>1,234,567</td></tr>
          <tr><td>영업이익(손실)</td><td>(12,345)</td></tr>
        </table>
        """
        found = self.h["_extract_income_amounts"](html)
        self.assertEqual(found["revenue"], "1,234,567,000 원")
        self.assertEqual(found["profit"], "-12,345,000 원")

    def test_year_from_audit_name(self):
        self.assertEqual(
            self.h["_year_from_audit_item"]({"name": "감사보고서 (2025.12)"}),
            "2025",
        )

    def test_pick_fin_account_prefers_ofs_is(self):
        df = pd.DataFrame(
            [
                {"account_nm": "매출액", "sj_div": "BS", "fs_div": "CFS", "thstrm_amount": "1"},
                {"account_nm": "매출액", "sj_div": "IS", "fs_div": "OFS", "thstrm_amount": "99"},
                {"account_nm": "매출원가", "sj_div": "IS", "fs_div": "OFS", "thstrm_amount": "2"},
            ]
        )
        row = self.h["_pick_fin_account"](df, ["매출액"], ["매출액"])
        self.assertEqual(str(row["thstrm_amount"]), "99")

    def test_revenue_label_with_note(self):
        html = """
        <table>
          <tr><td>매출액(주석2)</td><td>31,232,147,191</td><td>23,663,099,944</td></tr>
          <tr><td>영업이익</td><td>5,026,397,254</td><td>3,136,447,933</td></tr>
        </table>
        """
        found = self.h["_extract_income_amounts"](html)
        self.assertEqual(found["revenue"], "31,232,147,191 원")
        self.assertEqual(found["profit"], "5,026,397,254 원")

    def test_parse_viewer_section(self):
        js = """
        node2['text'] = "손 익 계 산 서";
        node2['dcmNo'] = "11222790";
        node2['eleId'] = "5";
        node2['offset'] = "58497";
        node2['length'] = "30978";
        node2['dtd'] = "dart4.xsd";
        """
        sec = self.h["_parse_dart_viewer_section"](js, "손익계산서")
        self.assertEqual(sec["dcmNo"], "11222790")
        self.assertEqual(sec["eleId"], "5")
        self.assertEqual(sec["offset"], "58497")

    def test_app_uses_audit_income_fallback(self):
        with open("app.py", encoding="utf-8") as f:
            src = f.read()
        self.assertIn("extract_dart_audit_income", src)
        self.assertIn("fnlttSinglAcntAll.json", src)
        self.assertIn("DART 감사보고서 손익계산서", src)


if __name__ == "__main__":
    unittest.main()
