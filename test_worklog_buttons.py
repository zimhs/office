"""업무일지 저장·추가·삭제 버튼이 CCv2 rerun 폭주로 죽지 않는지."""
from __future__ import annotations

import unittest
from unittest.mock import patch


class _FakeSS(dict):
    def pop(self, key, default=None):
        return dict.pop(self, key, default)


class WorklogButtonRestoreTest(unittest.TestCase):
    def setUp(self):
        import worklog_tab as wt

        self.wt = wt
        self.ss = _FakeSS()
        self._p_ss = patch.object(wt.st, "session_state", self.ss)
        self._p_ss.start()

    def tearDown(self):
        self._p_ss.stop()

    def test_queue_save_add_delete_flags(self):
        iso = "2026-09-06"
        self.wt._queue_worklog_save(iso)
        self.wt._queue_worklog_add(iso)
        self.wt._queue_worklog_del_entry(iso, 2)
        self.assertTrue(self.ss[f"wl_do_save_{iso}"])
        self.assertTrue(self.ss[f"wl_do_add_{iso}"])
        self.assertEqual(self.ss[f"wl_do_del_{iso}"], 2)

    def test_lines_editor_does_not_emit_while_typing(self):
        js = self.wt._WL_LINES_JS
        self.assertIn("lastEmitted", js)
        self.assertNotIn("function softEmit", js)
        self.assertIn("else localOnly(cur);", js)
        self.assertIn("mode === \"blur\"", js)

    def test_component_names_bumped_for_cache(self):
        with open(self.wt.__file__, encoding="utf-8") as f:
            src = f.read()
        self.assertIn('"worklog_entry_lines_v20"', src)
        self.assertIn('"worklog_cell_nav_hook_v24"', src)
        self.assertNotIn('"worklog_entry_lines_v19"', src)
        self.assertNotIn('"worklog_entry_lines_v18"', src)
        self.assertNotIn('"worklog_cell_nav_hook_v23"', src)

    def test_enter_hook_is_singleton_and_silent_on_click(self):
        js = self.wt._WL_ENTER_HOOK_JS
        self.assertIn("__wlEnterHookOff", js)
        self.assertNotIn('document.addEventListener("click", onSel, true)', js)
        self.assertNotIn('document.addEventListener("focusin", onFocusIn, true)', js)
        self.assertIn('document.addEventListener("keydown", onKey, true)', js)

    def test_action_buttons_use_on_click_not_extra_rerun(self):
        with open(self.wt.__file__, encoding="utf-8") as f:
            src = f.read()
        self.assertIn("on_click=_queue_worklog_save", src)
        self.assertIn("on_click=_queue_worklog_add", src)
        self.assertIn("on_click=_queue_worklog_del_entry", src)
        # 저장/추가/삭제 클릭 직후 fragment를 한 번 더 rerun 하면 ERROR·먹통
        self.assertNotIn('st.session_state[f"wl_do_save_{iso2}"] = True\n                    _wl_rerun()', src)
        self.assertNotIn('st.session_state[f"wl_do_add_{iso2}"] = True\n                    _wl_rerun()', src)

    def test_enter_hook_host_cannot_cover_buttons(self):
        with open(self.wt.__file__, encoding="utf-8") as f:
            src = f.read()
        self.assertIn("st-key-wl_enter_hook_", src)
        self.assertIn("pointer-events: none !important", src)
        self.assertIn("height=1", src)

    def test_comp_state_write_remounts_after_api_error(self):
        iso = "2026-09-06"
        self.ss["wl_lines_inst_2026-09-06_0"] = 0
        real_setitem = _FakeSS.__setitem__

        def _setitem(d, k, v):
            if k == "wl_lines_comp_2026-09-06_0_i0":
                raise self.wt.StreamlitAPIException("cannot be modified after instantiate")
            return real_setitem(d, k, v)

        with patch.object(_FakeSS, "__setitem__", _setitem):
            self.wt._set_comp_lines_state(iso, 0, ["새줄", ""], focus_j=0)
        self.assertEqual(self.ss.get("wl_lines_inst_2026-09-06_0"), 1)
        self.assertEqual(self.ss.get("wl_lines_comp_2026-09-06_0_i1")["lines"], ["새줄", ""])

    def test_date_input_moves_saved_day_instead_of_switching(self):
        with open(self.wt.__file__, encoding="utf-8") as f:
            src = f.read()
        self.assertIn("on_change=_on_wl_date_pick_change", src)
        self.assertIn("def apply_worklog_date_change", src)
        self.assertIn("def _set_wl_date_pick", src)
        self.assertIn("저장됨 · 날짜를 3일처럼 바꾸면 이 일지가 그 날짜로 이동합니다.", src)
        self.assertNotIn("if os.path.exists(worklog_path(picked)):", src)
        self.assertNotIn('st.session_state["wl_date_pick"] = selected', src)


if __name__ == "__main__":
    unittest.main()
