"""모든 탭은 시작부터 펼침 — 화면 불러오기 stub 로 접히지 않아야 한다."""
from __future__ import annotations

import ast
import unittest


_DEFER_FNS = (
    "_dash_should_defer_light_tab",
    "_dash_should_defer_heavy_tab",
)


class DashTabExpandTest(unittest.TestCase):
    def test_defer_helpers_always_return_false(self):
        with open("app.py", encoding="utf-8") as f:
            tree = ast.parse(f.read())
        found: set[str] = set()
        for node in tree.body:
            if not isinstance(node, ast.FunctionDef) or node.name not in _DEFER_FNS:
                continue
            found.add(node.name)
            returns = [n for n in ast.walk(node) if isinstance(n, ast.Return)]
            self.assertTrue(returns, msg=f"{node.name} has no return")
            for ret in returns:
                self.assertIsInstance(ret.value, ast.Constant, msg=f"{node.name} return is not constant")
                self.assertIs(ret.value.value, False, msg=f"{node.name} must always return False")
        self.assertEqual(found, set(_DEFER_FNS))


class CloudClipFixIsolationTest(unittest.TestCase):
    def test_cloud_clip_css_not_inside_shared_inject(self):
        with open("app.py", encoding="utf-8") as f:
            src = f.read()
        shared = src.split("def inject_custom_css():", 1)[1].split("def parse_date_series_robust", 1)[0]
        self.assertNotIn("dashboard-cloud-tab2-head-gap", shared)
        self.assertIn("def inject_cloud_clip_fix_css", src)
        self.assertIn("if _is_streamlit_cloud():\n    inject_cloud_clip_fix_css()", src)
        cloud_css = src.split("def inject_cloud_clip_fix_css():", 1)[1].split(
            "# 5. 메인 실행 흐름", 1
        )[0]
        self.assertNotIn("#dashboard-sticky-spacer", cloud_css)
        warm = src.split("def _dash_map_autowarm_fragment", 1)[1].split(
            "if not st.session_state.get(\"_dash_map_autowarm_done\")", 1
        )[0]
        self.assertIn("if _is_streamlit_cloud():", warm)
        self.assertIn("function pushCloudContentBelowBar", src)
        self.assertIn("need = Math.round(bottom - mainTop);", src)
        self.assertIn("streamlit.app", src)
        self.assertIn("function detectCloudHost", src)
        self.assertIn("function isLocalDesktopHost", src)
        self.assertIn("if (isLocalDesktopHost()) return false;", src)
        self.assertIn("function firstCloudContentEl", src)
        self.assertIn("dashboard-cloud-content-pad", src)
        self.assertIn("os.path.isdir(\"/mount/src\")", src)
        self.assertIn("/home/adminuser", src)
        self.assertIn("if (!cloudMode) return h;", src)
        self.assertIn("padding-top: 15px !important;", cloud_css)
        self.assertIn("z-index: 100000 !important;", cloud_css)
        self.assertNotIn("z-index: 100000 !important;", shared)
        self.assertIn("cloudMode ? '100' : '990'", src)
        self.assertIn("function cloudAvoidSidebarOverlap", src)
        self.assertIn("function syncIpadTopShield", src)
        self.assertIn("var topPx = 44", src)
        self.assertIn("최상단 Streamlit 줄", src)
        self.assertNotIn(
            "html.dashboard-touch-mode [data-testid=\"stToolbar\"],\n"
            "            html.dashboard-touch-mode [data-testid=\"stDecoration\"] {\n"
            "                display: none !important;",
            src,
        )
        self.assertNotIn("Math.max(Math.round(topPx) || 0, Math.round(barBottom)", src)
        self.assertIn("html.dashboard-touch-mode #dashboard-top-shield", src)
        self.assertIn("Safari 주소창 접힘", src)
        self.assertIn("if (prev && Math.abs(w - prev) < 24) return", src)
        self.assertIn("parentWin.__dashboardIpadRaf = null", src)
        self.assertIn("if (isTouchPadEarly()) return h", src)
        self.assertIn("fromNav(parentWin.navigator)", src)
        self.assertIn("dashboard-cloud-clipfix.dashboard-touch-mode", src)
        self.assertIn("if (!cloudMode) return h;", src)
        tab2 = src.split("# Tab 2:", 1)[1].split("with tab3:", 1)[0]
        self.assertIn("dashboard-cloud-tab2-head-gap", tab2)
        self.assertIn("if _is_streamlit_cloud():", tab2)
        self.assertIn("영업 실적 및 요약", tab2)
        self.assertIn("연도별 월 매출 추이", tab2)
        self.assertLess(
            tab2.find("영업 실적 및 요약"),
            tab2.find("연도별 월 매출 추이"),
        )
        self.assertLess(
            tab2.find("품목별 상세 분석"),
            tab2.find("매출 비교 ("),
        )
        _summary_idx = tab2.find("영업 실적 및 요약")
        _btn_idx = tab2.find("tab2_action_btns")
        self.assertGreater(_summary_idx, _btn_idx)

    def test_cloud_command_opens_one_tab(self):
        with open("dashboard_Cloud.command", encoding="utf-8") as f:
            src = f.read()
        self.assertNotIn('open -a "Google Chrome" "$CLOUD"', src)
        self.assertIn('open -a "Google Chrome"', src)
        self.assertIn("make new tab at end of tabs of keepWin", src)
        self.assertIn('if seenCloud then', src)

    def test_filter_clear_uses_per_field_buttons(self):
        with open("app.py", encoding="utf-8") as f:
            src = f.read()
        self.assertIn("def _dash_on_filter_clear_client", src)
        self.assertIn('key="_dash_filter_clear_client_btn"', src)
        self.assertIn("_dash_filter_clear_client_btn", src)
        self.assertIn("_DASH_FILTER_BIND_VER = 9", src)
        self.assertIn("clickClearRerunBtn(fieldKey)", src)


if __name__ == "__main__":
    unittest.main()
