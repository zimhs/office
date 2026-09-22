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

    def test_pin_worklog_scroll_clears_focus(self):
        iso = "2026-09-16"
        self.ss[f"wl_focus_ln_{iso}"] = f"wl_ent_ln_{iso}_0_3_g0"
        self.ss[f"wl_focus_caret_{iso}"] = 2
        self.ss["wl_active_cell_key"] = f"wl_ent_ln_{iso}_0_3_g0"
        self.wt._pin_worklog_scroll()
        self.assertEqual(self.ss.get("wl_scroll_pin"), 1)
        self.assertNotIn(f"wl_focus_ln_{iso}", self.ss)
        self.assertNotIn("wl_active_cell_key", self.ss)

    def test_queue_page_clears_stale_focus(self):
        iso = "2026-09-16"
        self.ss[f"wl_focus_ln_{iso}"] = f"wl_ent_ln_{iso}_0_3_g0"
        self.ss[f"wl_focus_caret_{iso}"] = 2
        self.ss["wl_active_cell_key"] = f"wl_ent_ln_{iso}_0_3_g0"
        self.wt._queue_worklog_page(iso, 1)
        self.assertEqual(self.ss[self.wt._page_idx_key(iso)], 1)
        self.assertNotIn(f"wl_focus_ln_{iso}", self.ss)
        self.assertNotIn(f"wl_focus_caret_{iso}", self.ss)
        self.assertNotIn("wl_active_cell_key", self.ss)
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
        self.assertNotIn('textContent = "×"', js)

    def test_component_names_bumped_for_cache(self):
        with open(self.wt.__file__, encoding="utf-8") as f:
            src = f.read()
        self.assertIn('"worklog_entry_lines_v53"', src)
        self.assertNotIn('"worklog_entry_lines_v52"', src)
        self.assertNotIn('"worklog_entry_lines_v51"', src)
        self.assertNotIn('"worklog_entry_lines_v50"', src)
        self.assertNotIn('"worklog_entry_lines_v49"', src)
        self.assertNotIn('"worklog_entry_lines_v43"', src)
        self.assertNotIn('"worklog_entry_lines_v42"', src)
        self.assertNotIn('"worklog_entry_lines_v41"', src)
        self.assertNotIn('"worklog_entry_lines_v40"', src)
        self.assertNotIn('"worklog_entry_lines_v39"', src)
        self.assertNotIn('"worklog_entry_lines_v38"', src)
        self.assertNotIn('"worklog_entry_lines_v37"', src)
        self.assertNotIn('"worklog_entry_lines_v36"', src)
        self.assertNotIn('"worklog_entry_lines_v34"', src)
        self.assertNotIn('"worklog_entry_lines_v33"', src)
        self.assertNotIn('"worklog_entry_lines_v32"', src)
        self.assertNotIn('"worklog_entry_lines_v31"', src)
        self.assertNotIn('"worklog_entry_lines_v30"', src)
        self.assertIn('"worklog_cell_nav_hook_v32"', src)
        self.assertNotIn('"worklog_cell_nav_hook_v31"', src)
        self.assertNotIn('"worklog_cell_nav_hook_v30"', src)
        self.assertNotIn('"worklog_entry_lines_v29"', src)
        self.assertNotIn('"worklog_entry_lines_v28"', src)
        self.assertNotIn('"worklog_entry_lines_v26"', src)
        self.assertNotIn('"worklog_entry_lines_v25"', src)
        self.assertNotIn('"worklog_entry_lines_v24"', src)
        self.assertNotIn('"worklog_entry_lines_v23"', src)
        self.assertNotIn('"worklog_entry_lines_v22"', src)
        self.assertNotIn('"worklog_entry_lines_v20"', src)
        self.assertNotIn('"worklog_entry_lines_v19"', src)
        self.assertNotIn('"worklog_cell_nav_hook_v29"', src)
        self.assertNotIn('"worklog_cell_nav_hook_v28"', src)
        self.assertNotIn('"worklog_cell_nav_hook_v27"', src)
        self.assertNotIn('"worklog_cell_nav_hook_v26"', src)
        self.assertNotIn('"worklog_cell_nav_hook_v25"', src)
        self.assertNotIn('"worklog_cell_nav_hook_v24"', src)
        self.assertIn('"worklog_live_preview_v2"', src)
        self.assertIn("scrollToPreviewPage", src)
        self.assertIn("page=page_n", src)
        self.assertIn('"worklog_print_launch_v2"', src)
        self.assertNotIn('"worklog_print_launch_v1"', src)
        self.assertNotIn('window.open("", "_blank")', src)
        self.assertIn('wl-print-frame', src)

    def test_enter_hook_is_singleton_and_silent_on_click(self):
        js = self.wt._WL_ENTER_HOOK_JS
        self.assertIn("__wlEnterHookOff", js)
        self.assertNotIn('document.addEventListener("click", onSel, true)', js)
        self.assertNotIn('document.addEventListener("focusin", onFocusIn, true)', js)
        self.assertIn('document.addEventListener("keydown", onKey, true)', js)
        self.assertIn("composedPath", js)
        self.assertIn("queryDeep", js)
        self.assertIn('e.key === "ArrowUp"', js)
        self.assertIn('e.key === "ArrowLeft"', js)
        self.assertIn('e.key === "ArrowRight"', js)
        self.assertIn("scrollCellIntoTabView", js)
        self.assertIn("preventScroll: true", js)
        self.assertIn("clearFocusTimers", js)
        self.assertIn("findCell(kind, lj, ei)", js)
        self.assertIn("function slotOk(part)", js)
        self.assertIn("wlNavWorklogArrow", js)
        self.assertIn("\\d{4}-\\d{2}-\\d{2}|ui", js)
        self.assertNotIn("setTimeout(go, 600)", js)
        self.assertIn('tag === "INPUT" || tag === "TEXTAREA"', js)
        self.assertIn("return false;", js)

    def test_lines_editor_does_not_steal_clicked_cell(self):
        js = self.wt._WL_LINES_JS
        self.assertIn("function shouldApplyDataFocus()", js)
        self.assertIn("function mergeIncoming(next)", js)
        self.assertIn("function fillEmptyFrom(src, dest)", js)
        self.assertIn("const isoChanged", js)
        self.assertIn("normalize(incoming)", js)
        self.assertIn("__wlLinesMem", js)
        self.assertIn('ch === " " || ch === "\\t"', js)
        self.assertIn("return 0.58;", js)
        self.assertIn("function lineOver", js)
        self.assertIn("const side8", js)
        self.assertIn("side8 && displayUnits(s) <= maxU", js)
        self.assertIn("if (!side8) return displayUnits(s) > maxU;", js)
        self.assertIn("if (!side8) return fitByUnits(s, maxU);", js)
        self.assertIn("function measureOrigPx", js)
        self.assertIn("function fitLine", js)
        self.assertIn("function fitByOrigPx", js)
        self.assertIn("function applyFillScale", js)
        self.assertIn("function inputInnerW", js)
        self.assertIn("function measureShowPx", js)
        self.assertIn("function fitByShowPx", js)
        self.assertIn("function bindDragSelect", js)
        self.assertIn("function selectedText", js)
        self.assertIn("function paintSel", js)
        self.assertIn("function deleteSelected", js)
        self.assertIn("addEventListener(\"cut\"", js)
        self.assertIn("Backspace", js)
        self.assertIn("--wl-show-pt", js)
        self.assertIn("--wl-look-pt", js)
        self.assertIn("look_scale", js)
        self.assertIn("fontPt * scale", js)
        self.assertNotIn("function leadLen", js)
        self.assertNotIn("function bodyOf", js)
        self.assertNotIn("function innerW", js)
        self.assertNotIn("view_scale", js)
        with open(self.wt.__file__, encoding="utf-8") as f:
            src = f.read()
        self.assertIn('"replace": 1 if replace else 0', src)
        self.assertIn("def _comp_send_focus", src)
        self.assertIn("def _maybe_remember_comp_focus", src)
        self.assertIn("def _coalesce_editor_lines", src)
        self.assertIn('_comp_send_focus(iso, entry_i, "ln")', src)
        self.assertIn('"slot": str(entry_i)', src)
        self.assertIn("def _collapse_empty_worklog_pages", src)

    def test_empty_incoming_does_not_wipe_typed_lines(self):
        self.assertEqual(
            self.wt._coalesce_editor_lines(["", "", ""], ["첫번째", "두번째"], None),
            ["첫번째", "두번째"],
        )
        self.assertEqual(self.wt._coalesce_editor_lines(None, None), [""])

    def test_blur_does_not_restore_previous_content_cell(self):
        iso = "2026-09-16"
        self.ss[self.wt._comp_focus_seen_key(iso, 0, "ln")] = 0
        self.wt._maybe_remember_comp_focus(
            iso, 0, "ln", {"lines": ["첫번째", "두번째"], "focus": 0}, self.wt._entry_line_key
        )
        self.assertNotIn(f"wl_focus_ln_{iso}", self.ss)

    def test_first_focus_sighting_does_not_restore(self):
        iso = "2026-09-16"
        self.wt._maybe_remember_comp_focus(
            iso, 0, "ln", {"lines": ["첫번째"], "focus": 0}, self.wt._entry_line_key
        )
        self.assertNotIn(f"wl_focus_ln_{iso}", self.ss)
        self.assertEqual(self.ss[self.wt._comp_focus_seen_key(iso, 0, "ln")], 0)

    def test_wrap_focus_move_is_remembered(self):
        iso = "2026-09-16"
        self.ss[self.wt._comp_focus_seen_key(iso, 0, "ln")] = 0
        self.wt._maybe_remember_comp_focus(
            iso, 0, "ln", {"lines": ["첫번째", "두번째"], "focus": 1}, self.wt._entry_line_key
        )
        self.assertEqual(self.ss.get(f"wl_focus_ln_{iso}"), self.wt._entry_line_key(iso, 0, 1))

    def test_comp_send_focus_is_one_shot_from_pending_key(self):
        iso = "2026-09-16"
        self.assertEqual(self.wt._comp_send_focus(iso, 0, "ln"), -1)
        self.ss[f"wl_focus_ln_{iso}"] = self.wt._entry_line_key(iso, 0, 4)
        self.assertEqual(self.wt._comp_send_focus(iso, 0, "ln"), 4)
        self.assertEqual(self.wt._comp_send_focus(iso, 1, "ln"), -1)

    def test_action_buttons_use_on_click_not_extra_rerun(self):
        with open(self.wt.__file__, encoding="utf-8") as f:
            src = f.read()
        self.assertIn("on_click=_queue_worklog_save", src)
        self.assertNotIn("＋ 항목 추가", src)
        self.assertNotIn("이 항목 삭제", src)
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
        inst_k = self.wt._entry_lines_inst_key(iso, 0)
        self.ss[inst_k] = 0
        blocked = self.wt._entry_lines_comp_key(iso, 0)
        real_setitem = _FakeSS.__setitem__

        def _setitem(d, k, v):
            if k == blocked:
                raise self.wt.StreamlitAPIException("cannot be modified after instantiate")
            return real_setitem(d, k, v)

        with patch.object(_FakeSS, "__setitem__", _setitem):
            self.wt._set_comp_lines_state(iso, 0, ["새줄", ""], focus_j=0)
        self.assertEqual(self.ss.get(inst_k), 1)
        self.assertEqual(self.ss.get(self.wt._entry_lines_comp_key(iso, 0))["lines"], ["새줄", ""])

    def test_date_input_opens_saved_day_instead_of_copying(self):
        with open(self.wt.__file__, encoding="utf-8") as f:
            src = f.read()
        self.assertIn("on_change=_on_wl_date_pick_change", src)
        self.assertIn("def apply_worklog_date_change", src)
        self.assertIn("def _set_wl_date_pick", src)
        self.assertIn("def _render_worklog_date_toolbar", src)
        self.assertIn("consume_left_date_pick_move", src)
        self.assertIn("이미 저장된 데이터가 있으면 자료를 옮길 수 없습니다.", src)
        self.assertIn("def try_retarget_worklog_editor_date", src)
        self.assertNotIn("날짜를 바꾸면 • 도 그 날로 이동합니다", src)
        self.assertIn("def commit_worklog_date_save", src)
        self.assertIn("allow_overwrite=True", src)
        self.assertIn("2026-09-22 · 입력칸 여백활용 · 내용줄=미리보기", src)
        self.assertIn("flex: 1 1 auto !important", src)
        self.assertIn("_WL_INPUT_LOOK_SCALE", src)
        self.assertIn("_WL_INPUT_SIDE_WIDEN_PX", src)
        self.assertIn("def _input_look_scale", src)
        self.assertIn("_side = max(int(_cw), int(_rw))", src)
        self.assertIn("st.columns([_side, _gw, _side]", src)
        self.assertNotIn("flex: 1.15 1 0 !important", src)
        self.assertIn("def _clear_wl_cal_nav_keys", src)
        self.assertIn('_clear_wl_cal_nav_keys()', src)
        self.assertIn("word-break:break-all !important", src)
        self.assertIn("def _inject_body_col_merges", src)
        self.assertIn("def _ensure_worklog_body_merges", src)
        self.assertIn("wl_left_excel_html_v27_", src)
        self.assertIn("_fit_preview_col_widths", src)
        self.assertIn("wl_entry_sheet", src)
        self.assertIn("wl_fit_root", src)
        self.assertIn("def _wl_ipad_preview_h", src)
        self.assertIn("orientation: portrait", src)
        self.assertIn("orientation: landscape", src)
        self.assertIn("100dvh", src)
        self.assertIn("_input_cell_px", src)
        self.assertIn("look_scale", src)
        self.assertIn("def _move_worklog_editor_to_date", src)
        self.assertIn("wl_date_move_mode", src)
        self.assertIn("날짜변경", src)
        self.assertIn("여기서 가로채 다시 심으면 이전 칸 값이 빠진다", src)
        self.assertIn("def _ui_iso", src)
        self.assertIn("날짜가 바뀌어도 같은 위젯을 유지한다", src)
        self.assertIn("def _park_shared_editor_to_date", src)
        self.assertIn("remount_comp=False", src)
        self.assertIn("purge_worklog_day_files", src)
        self.assertIn("def _open_worklog_saved_date", src)
        self.assertIn("def _reload_worklog_date_from_storage", src)
        self.assertIn("def _flush_worklog_delete_popover", src)
        self.assertIn("wl_del_day_force_close", src)
        self.assertNotIn('on_change="rerun"', src)
        pick = src[src.index("def _on_wl_date_pick_change") : src.index("def _on_wl_cal_day")]
        self.assertNotIn("try_retarget_worklog_editor_date", pick)
        self.assertIn("_move_worklog_editor_to_date", pick)
        self.assertIn("_open_worklog_saved_date", pick)
        cal = src[src.index("def _on_wl_cal_day") : src.index("def _run_pending_worklog_date_change")]
        self.assertNotIn("try_retarget_worklog_editor_date", cal)
        self.assertIn("_move_worklog_editor_to_date", cal)
        self.assertIn("_open_worklog_saved_date", cal)
        self.assertIn("날짜변경」을 켠 뒤 날짜를 고르면", src)
        self.assertNotIn("날짜만 바꾸면 입력 중인 내용은 그대로입니다", src)
        self.assertNotIn('value=selected,', src)
        self.assertNotIn("if os.path.exists(worklog_path(picked)):", src)
        self.assertNotIn('st.session_state["wl_date_pick"] = selected', src)
        self.assertIn("skip_remote_pull=True", src)
        self.assertNotIn('["wl_need_app_rerun"] = True', src)
        self.assertNotIn('["wl_need_app_rerun"] = 1', src)
        self.assertNotIn("_wl_rerun(full=True)", src)
        self.assertIn("def _drop_saved_date_from_cache", src)
        self.assertIn("def _remember_calendar_saved_date", src)
        self.assertIn("def _pin_worklog_scroll", src)
        self.assertIn('"worklog_scroll_lock_v3"', src)
        self.assertIn("keep_editor", src)
        self.assertIn("st-key-wl_del_", src)
        self.assertIn("__wlScrollSnap", src)
        self.assertIn("wl_enter_hook_nav", src)
        self.assertIn("wl_day_{d.isoformat()}", src)
        self.assertIn("'s' if has else 'n'", src)
        panel = src[src.index("def _render_worklog_input_panel") : src.index("def _render_worklog_sync_ui")]
        self.assertNotIn("_run_pending_worklog_day_delete", panel)
        frag = src.split("def _worklog_body", 1)[1].split("_worklog_body()", 1)[0]
        self.assertIn("_run_pending_worklog_day_delete()", frag)
        self.assertIn("_prepare_worklog_day_state", frag)


if __name__ == "__main__":
    unittest.main()
