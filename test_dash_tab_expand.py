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
        self.assertIn("Math.round(bot)", src)
        self.assertNotIn("return Math.max(got, 64)", src)
        self.assertIn("stSidebarCollapsedControl", src)
        self.assertIn(
            "html:not(.dashboard-touch-mode) [data-testid=\"stHeader\"]",
            src,
        )
        self.assertIn("맥 로컬·Cloud: 좌 >>", src)
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
        self.assertIn("dashboard-ipad-chrome-css", src)
        self.assertIn("function ensureIpadChromeVisible", src)
        self.assertIn("z-index: 999980 !important", src)
        self.assertNotIn("z-index: 999999 !important", src)
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
        self.assertIn("_DASH_FILTER_BIND_VER = 12", src)
        self.assertIn("function bindDirectInputs", src)
        self.assertIn("function clearForNewInput", src)
        self.assertIn("function holdCleared", src)
        self.assertIn("DATE_KEYS = ['dash_filter_start', 'dash_filter_end']", src)
        self.assertIn("clickClearRerunBtn(fieldKey)", src)


class WorklogDraftSurviveFilterTest(unittest.TestCase):
    """상단 검색/필터 후 일일업무일지로 돌아와도 미저장 초안이 남아야 한다."""

    def test_worklog_restore_does_not_pop_drafts(self):
        with open("app.py", encoding="utf-8") as f:
            src = f.read()
        restore = src.split("def _dash_restore_session_keys", 1)[1].split(
            "def _dash_should_defer_light_tab", 1
        )[0]
        self.assertIn('store_key != "_dash_bak_worklog"', restore)
        self.assertIn("미저장 초안", restore)

    def test_worklog_draft_keys_are_state_not_widgets(self):
        with open("app.py", encoding="utf-8") as f:
            src = f.read()
        prefixes = src.split("_DASH_WL_STATE_PREFIXES = (", 1)[1].split(")", 1)[0]
        for p in (
            "wl_entries_",
            "wl_next_",
            "wl_notes_",
            "wl_lines_live_",
            "wl_clients_live_",
            "worklog_booted_",
        ):
            self.assertIn(f'"{p}"', prefixes)
        classifier = src.split("def _dash_is_wl_widget_key", 1)[1].split(
            "_DASH_FILTER_ALL_STAFF", 1
        )[0]
        state_idx = classifier.find("_DASH_WL_STATE_PREFIXES")
        wl_idx = classifier.find('key.startswith("wl_")')
        self.assertGreater(state_idx, -1)
        self.assertGreater(wl_idx, -1)
        self.assertLess(state_idx, wl_idx)

    def _load_restore_helpers(self):
        import types

        with open("app.py", encoding="utf-8") as f:
            src = f.read()
        prefix_src = src[
            src.find("# 업무일지 — 편집 상태") : src.find("_DASH_FILTER_ALL_STAFF")
        ]
        restore_src = src[
            src.find("def _dash_restore_session_keys") : src.find(
                "def _dash_should_defer_light_tab"
            )
        ]
        ns = {
            "st": types.SimpleNamespace(session_state={}),
            "_dash_is_mr_widget_key": lambda key: False,
            "_dash_is_pi_widget_key": lambda key: False,
        }
        exec(prefix_src + restore_src, ns)
        return ns

    def test_restore_keeps_unsaved_worklog_drafts(self):
        ns = self._load_restore_helpers()
        st = ns["st"]
        iso = "2026-09-09"
        st.session_state.update(
            {
                "_dash_bak_worklog": {"worklog_selected": iso},
                f"wl_entries_{iso}": [{"client": "테스트거래처", "content": "초안내용"}],
                f"wl_next_area_{iso}": "익일 초안",
                f"wl_notes_area_{iso}": "특이 초안",
                f"wl_lines_live_{iso}_0": ["줄1", ""],
                f"worklog_booted_{iso}": True,
                "wl_date_pick": iso,
            }
        )
        ns["_dash_restore_session_keys"]("_dash_bak_worklog")
        self.assertEqual(st.session_state[f"wl_next_area_{iso}"], "익일 초안")
        self.assertEqual(st.session_state[f"wl_notes_area_{iso}"], "특이 초안")
        self.assertEqual(
            st.session_state[f"wl_entries_{iso}"][0]["client"], "테스트거래처"
        )
        self.assertTrue(st.session_state[f"worklog_booted_{iso}"])
        self.assertEqual(st.session_state[f"wl_lines_live_{iso}_0"], ["줄1", ""])
        self.assertEqual(st.session_state["wl_date_pick"], iso)
        self.assertFalse(ns["_dash_is_wl_widget_key"](f"wl_entries_{iso}"))
        self.assertFalse(ns["_dash_is_wl_widget_key"](f"wl_next_area_{iso}"))
        self.assertTrue(ns["_dash_is_wl_widget_key"]("wl_date_pick"))


if __name__ == "__main__":
    unittest.main()
