"""거래처분석 추가 사업장 주소 — 카카오맵에만 같이 쓴다."""
from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch


class Tab2ExtraSitesTest(unittest.TestCase):
    def setUp(self):
        import app as dash_app

        self.app = dash_app
        self.tmp = tempfile.TemporaryDirectory()
        self.path = str(Path(self.tmp.name) / "client_extra_sites.json")
        self.patcher = patch.object(dash_app, "_TAB2_EXTRA_SITES_FILE", self.path)
        self.patcher.start()

    def tearDown(self):
        self.patcher.stop()
        self.tmp.cleanup()

    def test_add_named_site_under_parent(self):
        err = self.app._tab2_add_extra_site("라쿨", "라쿨1사업장", "경기도 화성시 팔탄면")
        self.assertEqual(err, "")
        rows = self.app._tab2_extra_sites_for("라쿨")
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["name"], "라쿨1사업장")
        self.assertEqual(rows[0]["addr"], "경기도 화성시 팔탄면")
        data = json.loads(Path(self.path).read_text(encoding="utf-8"))
        self.assertIn("라쿨", data)

    def test_same_name_updates_address(self):
        self.app._tab2_add_extra_site("라쿨", "라쿨1사업장", "옛주소")
        self.app._tab2_add_extra_site("라쿨", "라쿨1사업장", "새주소")
        rows = self.app._tab2_extra_sites_for("라쿨")
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["addr"], "새주소")

    def test_delete_site(self):
        self.app._tab2_add_extra_site("라쿨", "라쿨1사업장", "경기도")
        self.app._tab2_delete_extra_site("라쿨", "라쿨1사업장")
        self.assertEqual(self.app._tab2_extra_sites_for("라쿨"), [])

    def test_map_rows_follow_parent_or_site_name(self):
        self.app._tab2_add_extra_site("라쿨", "라쿨1사업장", "경기도 화성")
        staff = {"라쿨": "김혁수"}
        hit = self.app._tab6_extra_map_rows(staff, ["라쿨"])
        self.assertEqual(len(hit), 1)
        self.assertEqual(hit[0]["거래처"], "라쿨1사업장")
        self.assertEqual(hit[0]["담당자"], "김혁수")
        by_name = self.app._tab6_extra_map_rows(staff, ["라쿨1사업장"])
        self.assertEqual(len(by_name), 1)
        hidden = self.app._tab6_extra_map_rows(staff, None, visible_parents=set())
        self.assertEqual(hidden, [])

    def test_ui_only_on_tab2_and_tab6(self):
        with open(self.app.__file__, encoding="utf-8") as f:
            src = f.read()
        tab2 = src[src.index("with tab2:") : src.index("with tab3:")]
        tab6 = src[src.index("with tab6:") : src.index("with tab7:")]
        tab3 = src[src.index("with tab3:") : src.index("with tab4:")]
        self.assertIn("_render_tab2_extra_addr_box", tab2)
        self.assertIn("@st.fragment\ndef _render_tab2_extra_addr_box", src)
        self.assertIn("@st.fragment\ndef _render_tab6_addr_book", src)
        self.assertIn("tab2_xapply_", src)
        self.assertIn("추가 사업장 이름", src)
        self.assertIn("_tab2_extra_site_row", src)
        self.assertIn('"×"', src[src.index("_tab2_extra_site_row") : src.index("_tab6_extra_site_options")])
        self.assertIn("disabled=False", src)
        self.assertIn('key_prefix="tab2_xdel"', src)
        self.assertIn("tab6_xdel", src)
        self.assertIn("_render_tab6_addr_book", tab6)
        self.assertIn("_tab6_extra_map_rows", tab6)
        apply_src = src[src.index("def _tab2_apply_extra_site") : src.index("def _tab2_on_delete_extra_site")]
        del_src = src[src.index("def _tab2_on_delete_extra_site") : src.index("def _tab2_extra_widget_key")]
        self.assertNotIn("map_force_rebuild", apply_src)
        self.assertNotIn("map_force_rebuild", del_src)
        self.assertNotIn("_render_tab2_extra_addr_box", tab3)
        self.assertNotIn("tab2_xapply_", tab3)
        self.assertNotIn("with tab1:", tab2)
        self.assertNotIn("with tab3:", tab2)
        self.assertNotIn("with tab5:", tab6)
        self.assertNotIn("with tab7:", tab6)


if __name__ == "__main__":
    unittest.main()
