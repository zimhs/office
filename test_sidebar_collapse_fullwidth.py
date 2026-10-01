"""사이드바 접힘 — 분기사업분석 안정 시점 회귀(강제 JS 없음)."""
from __future__ import annotations

import os
import unittest


class SidebarCollapseFullwidthTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        path = os.path.join(os.path.dirname(__file__), "app.py")
        with open(path, encoding="utf-8") as f:
            cls.src = f.read()

    def test_streamlit_default_collapse_restored(self):
        self.assertIn("STICKY_SCRIPT_VER_MAC = 49", self.src)
        self.assertIn("_sticky_py_ver = 106", self.src)
        self.assertIn("분기사업분석 안정 시점", self.src)
        # 진동 원인이던 JS 강제 접힘 제거
        self.assertNotIn("function forceLocalSidebarLayout", self.src)
        self.assertNotIn("function enforceCollapsedStickyBar", self.src)
        self.assertNotIn("__dashboardCollapseWatch = setInterval", self.src)

    def test_css_only_when_aria_false(self):
        self.assertIn(
            '[data-testid="stSidebar"][aria-expanded="false"]',
            self.src,
        )
        self.assertIn("grid-template-columns: 0 minmax(0, 1fr)", self.src)
        self.assertIn("stExpandSidebarButton", self.src)

    def test_sticky_reads_collapsed_without_mutating_sidebar(self):
        self.assertIn("sidebarLooksOpen()", self.src)
        self.assertIn("keepExpandBtnAlive", self.src)
        self.assertIn("사이드바 DOM/style 은 절대 수정하지 않음", self.src)

    def test_local_desktop_not_treated_as_ipad(self):
        self.assertIn("로컬 Desktop(localhost): Mac 트랙패드/터치스크린을 iPad로 오인하면", self.src)


if __name__ == "__main__":
    unittest.main(verbosity=2)
