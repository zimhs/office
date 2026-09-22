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

    def test_click_toggles_without_full_rerun(self):
        self.assertIn("def _tab2_toggle_corp_info", self.src)
        self.assertIn("on_click=_tab2_toggle_corp_info", self.src)
        self.assertIn("@st.fragment\n        def _tab2_corp_actions", self.src)
        tab2 = self.src[self.src.index("with tab2:") : self.src.index("with tab3:")]
        dart_idx = tab2.find("key=\"btn_dart_info\"")
        window = tab2[max(0, dart_idx - 80) : dart_idx + 220]
        self.assertNotIn("st.rerun()", window)
        self.assertNotIn("st.rerun()", tab2[tab2.find("def _tab2_corp_actions") : tab2.find("_tab2_corp_actions()")])

    def test_corp_card_css_exists(self):
        self.assertIn(".tab2-corp-card", self.src)
        self.assertIn(".tab2-corp-title", self.src)
        self.assertIn(".tab2-corp-grid .v", self.src)

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
