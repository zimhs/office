"""거래처 분석 — 기업정보 보기/닫기 문구와 펼침 카드."""
from __future__ import annotations

import ast
import unittest


def _app_source() -> str:
    with open("app.py", encoding="utf-8") as f:
        return f.read()


class Tab2CorpToggleTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.src = _app_source()

    def test_closed_label_is_보기(self):
        self.assertIn("🏢 기업정보 보기", self.src)
        self.assertNotIn("🏢 기업 기본/재무정보 보기", self.src)

    def test_open_label_is_닫기(self):
        self.assertIn("🏢 기업정보 닫기", self.src)

    def test_label_follows_show_corp_info(self):
        self.assertIn(
            'btn_label = "🏢 기업정보 닫기" if _corp_open else "🏢 기업정보 보기"',
            self.src,
        )

    def test_click_reruns_so_label_matches_panel(self):
        self.assertIn("st.session_state.show_corp_info = not _corp_open", self.src)
        idx = self.src.find("st.session_state.show_corp_info = not _corp_open")
        window = self.src[idx : idx + 80]
        self.assertIn("st.rerun()", window)

    def test_corp_card_css_exists(self):
        self.assertIn(".tab2-corp-card", self.src)
        self.assertIn(".tab2-corp-title", self.src)
        self.assertIn(".tab2-corp-grid .v", self.src)
        self.assertIn("grid-template-columns: repeat(3, minmax(0, 1fr))", self.src)
        self.assertIn("wide_keys", self.src)

    def test_no_audit_dict_empty_crash(self):
        self.assertNotIn("_latest_audit.empty", self.src)

    def test_card_html_not_markdown_indented(self):
        self.assertIn('class="tab2-corp-card"', self.src)
        self.assertNotRegex(
            self.src,
            r'\n[ ]{4,}<div class="tab2-corp-card">',
        )
        tree = ast.parse(self.src)
        self.assertTrue(tree.body)


if __name__ == "__main__":
    unittest.main()
