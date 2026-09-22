"""카카오맵 탭 — 드래그 네모 안 업체(담당자·거래처·주소)."""
from __future__ import annotations

import unittest


class Tab6BoxSelectTest(unittest.TestCase):
    def setUp(self):
        import app as dash_app

        self.app = dash_app

    def test_clients_in_box_include_staff_name(self):
        rows = [
            {"lat": 37.10, "lon": 127.10, "거래처": "한신테크", "담당자": "김혁수", "주소": "화성시 동탄"},
            {"lat": 38.00, "lon": 128.00, "거래처": "멀리있는곳", "담당자": "홍길동", "주소": "강릉"},
            {"lat": 37.11, "lon": 127.12, "name": "이엔에이치", "staff": "김혁수", "addr": "오산"},
        ]
        hit = self.app._tab6_clients_in_box(rows, 37.05, 127.05, 37.20, 127.20)
        names = {r["거래처"] for r in hit}
        self.assertEqual(names, {"한신테크", "이엔에이치"})
        by_name = {r["거래처"]: r for r in hit}
        self.assertEqual(by_name["한신테크"]["담당자"], "김혁수")
        self.assertEqual(by_name["한신테크"]["주소"], "화성시 동탄")
        self.assertEqual(by_name["이엔에이치"]["담당자"], "김혁수")
        self.assertNotIn("멀리있는곳", names)

    def test_box_flips_inverted_bounds(self):
        rows = [{"lat": 37.0, "lon": 127.0, "거래처": "A", "담당자": "김혁수", "주소": "수원"}]
        hit = self.app._tab6_clients_in_box(rows, 37.5, 127.5, 36.5, 126.5)
        self.assertEqual(len(hit), 1)
        self.assertEqual(hit[0]["담당자"], "김혁수")

    def test_tab6_has_drag_button_and_list_columns(self):
        with open(self.app.__file__, encoding="utf-8") as f:
            src = f.read()
        tab6 = src[src.index("with tab6:") : src.index("with tab7:")]
        self.assertIn("⬚ 드래그 선택", tab6)
        self.assertIn('key="tab6_box_btn"', tab6)
        self.assertIn("def _tab6_toggle_box_mode", src)
        self.assertIn("tab6_box_mode", tab6)
        self.assertIn("영역 안 업체", src)
        self.assertIn("<th>담당자</th><th>거래처</th><th>주소</th>", src)
        self.assertIn("def _tab6_box_select_html", src)
        self.assertIn("_tab6_box_select_html(", tab6)
        self.assertIn("components.html(", tab6)
        self.assertIn("max-width: none !important", src)
        self.assertNotIn("tab6_box_map_v2", src)
        self.assertNotIn("setStateValue", src)
        self.assertIn("render_plotly_chart(fig_map", tab6)
        self.assertNotIn("with tab5:", tab6)
        self.assertNotIn("with tab7:", tab6)

    def test_box_select_html_is_self_contained_iframe(self):
        html = self.app._tab6_box_select_html(
            37.4,
            127.1,
            8,
            'L.tileLayer("https://example/tile/{z}/{x}/{y}.png", {maxZoom:19}).addTo(map);',
            [{"lat": 37.4, "lon": 127.1, "name": "한신테크", "staff": "김혁수", "addr": "화성", "color": "#1a73e8"}],
        )
        self.assertIn("<!DOCTYPE html>", html)
        self.assertIn("leaflet.css", html)
        self.assertIn("leaflet.js", html)
        self.assertIn("L.rectangle", html)
        self.assertIn("pointerdown", html)
        self.assertIn("touchstart", html)
        self.assertIn("setPointerCapture", html)
        self.assertIn("touch-action: none", html)
        self.assertIn("애플펜슬", html)
        self.assertIn("t6-nav", html)
        self.assertIn("data-act", html)
        self.assertIn("map.zoomIn", html)
        self.assertIn("map.panBy", html)
        self.assertIn('data-act="n"', html)
        self.assertIn('data-act="w"', html)
        self.assertIn('data-act="e"', html)
        self.assertIn('data-act="s"', html)
        self.assertIn("pointerup", html)
        self.assertIn("touchend", html)
        self.assertIn("담당자", html)
        self.assertIn("한신테크", html)
        self.assertIn("김혁수", html)


if __name__ == "__main__":
    unittest.main()
