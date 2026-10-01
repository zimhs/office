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
        self.assertIn("STICKY_SCRIPT_VER_MAC = 46", self.src)
        self.assertIn("__dashboardStickyGeoFrozen = null", self.src)
        # 접힘 시 고정바 좌표를 동결/본문 rect보다 나중에 전체폭으로 덮어씀
        self.assertIn("접힘: 본문 rect/동결/Cloud 보정과 무관하게 고정바를 화면 전체폭으로", self.src)

    def test_expand_button_kept_and_width_preserved(self):
        """>> 는 유지하고, 열린 폭은 removeProperty로 지우지 않고 저장/복구."""
        self.assertIn("saveSidebarInlineSize", self.src)
        self.assertIn("restoreSidebarInlineSize", self.src)
        self.assertIn("data-dash-sb-saved", self.src)
        self.assertIn("collapsedSure", self.src)
        self.assertIn("openSure", self.src)
        # >> DOM 존재만으로 접힘 (크기/opacity 검사 실패로 고정바가 남는 것 방지)
        self.assertIn("if (sidebarExpandBtnPresent()) return false;", self.src)
        self.assertIn(
            '[data-testid="stExpandSidebarButton"] {\n'
            "                visibility: visible !important;",
            self.src,
        )
        # 접힘 시 사이드바 자식 숨김(>> 는 헤더에 있어 안전) + grid 잔여폭 제거
        self.assertIn(
            'html.dashboard-sidebar-collapsed:not(.dashboard-touch-mode) [data-testid="stSidebar"] > div',
            self.src,
        )


if __name__ == "__main__":
    unittest.main(verbosity=2)
