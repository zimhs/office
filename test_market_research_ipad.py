"""시장조사 탭 — 아이패드 검색 선택 적용. 맥 경로는 그대로."""
from __future__ import annotations

import unittest


class MarketResearchIpadSelectTest(unittest.TestCase):
    def setUp(self):
        import market_research_tab as mr

        self.mr = mr

    def test_touch_helpers_exist(self):
        self.assertTrue(callable(self.mr._mr_is_touch_ui))
        self.assertTrue(callable(self.mr._mr_touch_sync_apply))
        self.assertEqual(self.mr._mr_touch_kwargs(), {})

    def test_mac_keeps_fragment_and_apply_caption(self):
        with open(self.mr.__file__, encoding="utf-8") as f:
            src = f.read()
        self.assertIn("@st.fragment\ndef _mr_filter_and_results(", src)
        self.assertIn("지역 → 산업단지 → 공급사 순으로 고른 뒤 **적용**을 누르세요.", src)
        self.assertIn("if _mr_is_touch_ui():", src)
        self.assertIn("_mr_filter_and_results_body", src)
        self.assertIn("html.dashboard-touch-mode [role=\"option\"]", src)
        self.assertIn("검색을 고르면 바로 적용됩니다.", src)

    def test_mac_widgets_do_not_force_on_change(self):
        with open(self.mr.__file__, encoding="utf-8") as f:
            src = f.read()
        self.assertIn("**touch_kw", src)
        self.assertIn('key="mr_w_region"', src)
        self.assertIn('key="mr_apply_btn"', src)


if __name__ == "__main__":
    unittest.main()
