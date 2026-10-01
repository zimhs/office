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
        self.assertIn("STICKY_SCRIPT_VER_MAC = 43", self.src)
        self.assertIn("__dashboardStickyGeoFrozen = null", self.src)


if __name__ == "__main__":
    unittest.main(verbosity=2)
