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
        self.assertIn("grid-template-columns: 0 minmax(0, 1fr)", self.src)

    def test_main_fullwidth_when_sidebar_collapsed(self):
        self.assertIn("stExpandSidebarButton", self.src)
        self.assertIn("flex: 1 1 auto", self.src)

    def test_sticky_script_uses_fullwidth_when_closed(self):
        self.assertIn("sidebarLooksOpen()", self.src)
        self.assertIn("sidebarExpandBtnVisible", self.src)
        self.assertIn("sidebarExpandBtnPresent", self.src)
        self.assertIn("fullWidthBarRect", self.src)
        self.assertIn("enforceCollapsedStickyBar", self.src)
        self.assertIn("STICKY_SCRIPT_VER_MAC = 47", self.src)
        self.assertIn("data-dash-sticky-ver", self.src)
        self.assertIn("__dashboardStickyGeoFrozen = null", self.src)
        self.assertIn("접힘: 본문 rect/동결/Cloud 보정과 무관하게 고정바를 화면 전체폭으로", self.src)

    def test_local_desktop_not_treated_as_ipad(self):
        """localhost Mac이 touch로 오인되면 Mac 접힘 보정이 스킵된다."""
        self.assertIn("로컬 Desktop(localhost): Mac 트랙패드/터치스크린을 iPad로 오인하면", self.src)
        self.assertIn("if (isLocalDesktopHost())", self.src)
        self.assertIn("__dashboardCollapseWatch", self.src)

    def test_expand_button_kept_and_width_preserved(self):
        """>> 는 유지하고, 열린 폭은 removeProperty로 지우지 않고 저장/복구."""
        self.assertIn("saveSidebarInlineSize", self.src)
        self.assertIn("restoreSidebarInlineSize", self.src)
        self.assertIn("data-dash-sb-saved", self.src)
        self.assertIn("collapsedSure", self.src)
        self.assertIn("openSure", self.src)
        self.assertIn("if (sidebarExpandBtnPresent()) return false;", self.src)
        self.assertIn(
            '[data-testid="stExpandSidebarButton"] {\n'
            "                visibility: visible !important;",
            self.src,
        )
        self.assertIn(
            'html.dashboard-sidebar-collapsed:not(.dashboard-touch-mode) [data-testid="stSidebar"] > div',
            self.src,
        )


if __name__ == "__main__":
    unittest.main(verbosity=2)
