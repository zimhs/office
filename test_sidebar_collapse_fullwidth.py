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

    def test_sticky_script_stable_collapse(self):
        self.assertIn("sidebarLooksOpen()", self.src)
        self.assertIn("sidebarIsCollapsedSure", self.src)
        self.assertIn("STICKY_SCRIPT_VER_MAC = 48", self.src)
        self.assertIn("data-dash-sticky-ver", self.src)
        self.assertIn("keepExpandBtnAlive", self.src)
        # 700ms 강제 폴링은 진동 원인 — 제거·무력화되어야 함
        self.assertIn("v47 접힘 감시 인터벌 제거", self.src)
        self.assertIn("__dashboardCollapseWatch = null", self.src)

    def test_aria_true_never_forced_collapsed(self):
        """aria=true 이면 접힘 강제 금지 — >> 실종 방지."""
        self.assertIn("if (aria === 'true') return false;", self.src)
        self.assertIn("sidebarIsCollapsedSure", self.src)
        # JS로 사이드바 자식 display:none 강제하지 않음
        self.assertNotIn("node.style.setProperty('display', 'none', 'important');", self.src)

    def test_local_desktop_not_treated_as_ipad(self):
        self.assertIn("로컬 Desktop(localhost): Mac 트랙패드/터치스크린을 iPad로 오인하면", self.src)


if __name__ == "__main__":
    unittest.main(verbosity=2)
