"""상단 담당자·거래처·품목 필터: 보고 있는 탭만 다시 그린다."""

from __future__ import annotations

import unittest
from unittest.mock import patch

import pandas as pd


class _FakeSS(dict):
    def pop(self, key, default=None):
        return dict.pop(self, key, default)

    def setdefault(self, key, default=None):
        if key not in self:
            self[key] = default
        return self[key]


class DashNameFilterReplayTest(unittest.TestCase):
    def setUp(self):
        import app as dash_app

        self.app = dash_app
        self.ss = _FakeSS()
        self.click = patch.object(dash_app, "_dash_click_repaint_tabs")
        self.active = patch.object(dash_app, "_dash_active_tab_idx", return_value=1)
        self._patches = [
            patch.object(dash_app.st, "session_state", self.ss),
            patch.object(dash_app, "_dash_mark_tabs_skipped"),
            self.click,
            self.active,
        ]
        for p in self._patches:
            p.start()
        self.click_mock = self.click.get_mock() if hasattr(self.click, "get_mock") else None

    def tearDown(self):
        for p in reversed(self._patches):
            p.stop()

    def test_client_change_repaints_only_active_tab(self):
        self.ss["_dash_name_replay_prev"] = ((), "전체 거래처", ())
        with patch.object(self.app, "_dash_click_repaint_tabs") as click:
            self.app._dash_replay_open_tab([], "라쿨 주식회사(구.신정우)", [])
            self.assertEqual(self.ss.get("_dash_repaint_queue"), [1])
            click.assert_called_once_with([1])
        self.assertTrue(self.ss.get("_dash_filter_changed_flag"))

    def test_summary_tab_skips_chain_repaint(self):
        self.ss["_dash_name_replay_prev"] = ((), "전체 거래처", ())
        with patch.object(self.app, "_dash_active_tab_idx", return_value=0), patch.object(
            self.app, "_dash_click_repaint_tabs"
        ) as click:
            self.app._dash_replay_open_tab([], "라쿨 주식회사(구.신정우)", [])
            self.assertNotIn("_dash_repaint_queue", self.ss)
            click.assert_not_called()

    def test_many_clients_filter_is_vectorized_union(self):
        df = pd.DataFrame(
            {
                "거래처": ["A", "A(공장)", "B", "C"],
                "거래처_원본": ["A", "A", "B", "C"],
                "매출액": [1, 2, 3, 4],
            }
        )
        out = self.app._filter_df_by_many_clients(df, ["A", "B"])
        self.assertEqual(sorted(out["거래처"].tolist()), ["A", "A(공장)", "B"])


if __name__ == "__main__":
    unittest.main()
