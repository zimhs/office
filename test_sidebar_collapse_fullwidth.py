"""사이드바 접힘 시 본문 전체폭 CSS·고정바 스크립트 회귀."""
from __future__ import annotations

import os
import unittest


class SidebarCollapseFullwidthTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        path = os.path.join(os.path.dirname(__file__), "app.py")
        with open(path, encoding="utf-8") as f:
            cls.src = f.read()

    def test_css_collapses_sidebar_width(self):
        self.assertIn("dashboard-sidebar-collapsed", self.src)
        self.assertIn("max-width: 0 !important", self.src)
        self.assertIn("forceLocalSidebarLayout", self.src)

    def test_main_fullwidth_when_sidebar_collapsed(self):
        self.assertIn("stExpandSidebarButton", self.src)
        self.assertIn("flex: 1 1 auto", self.src)

    def test_sticky_script_uses_fullwidth_when_closed(self):
        self.assertIn("sidebarLooksOpen()", self.src)
        self.assertIn("sidebarExpandBtnVisible", self.src)
        self.assertIn("STICKY_SCRIPT_VER_MAC = 44", self.src)
        self.assertIn("__dashboardStickyGeoFrozen = null", self.src)

    def test_expand_button_not_killed_by_false_collapse(self):
        """열린 사이드바를 접힌 것으로 오인하면 >> 가 DOM에서 사라진다."""
        self.assertIn("collapsedSure", self.src)
        self.assertIn("openSure", self.src)
        self.assertIn("clearForcedSidebarStyles", self.src)
        # 사이드바 자식을 display:none 하면 열기 버튼/내부 컨트롤이 깨질 수 있음
        self.assertNotIn(
            'html.dashboard-sidebar-collapsed:not(.dashboard-touch-mode) [data-testid="stSidebar"] > div',
            self.src,
        )
        self.assertIn(
            '[data-testid="stExpandSidebarButton"] {\n'
            "                visibility: visible !important;",
            self.src,
        )


if __name__ == "__main__":
    unittest.main(verbosity=2)
