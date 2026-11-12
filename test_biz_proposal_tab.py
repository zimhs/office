"""사업계획서: 16페이지를 나눠 저장하고, 2025년 실적은 원본과 맞춘다."""
from __future__ import annotations

import zipfile
from pathlib import Path

import biz_proposal_tab as bp

PPT = {
    2025: {
        "화성": {1: 849, 2: 900, 3: 959, 4: 1194, 5: 1022, 6: 1137, 7: 1519, 8: 1393, 9: 1387, 10: 1107, 11: 1056, 12: 1177},
        "아산": {1: 303, 2: 360, 3: 383, 4: 393, 5: 354, 6: 375, 7: 360, 8: 364, 9: 412, 10: 348, 11: 386, 12: 373},
    }
}


def _template() -> Path:
    path = bp.find_template()
    assert path is not None and path.is_file()
    return path


def test_sixteen_pages_each_have_a_save_button():
    src = Path("biz_proposal_tab.py").read_text(encoding="utf-8")
    assert "PAGES = tuple(range(1, 17))" in src
    assert 'f"{page}페이지 저장"' in src
    app = Path("app.py").read_text(encoding="utf-8")
    assert "📑 사업계획서" in app
    assert "render_biz_proposal_tab" in app
    assert "min_tabs=13" in app
    assert "tab14" not in app


def test_2025_plant_months_match_the_plan():
    won = bp.plant_won(2025, 12)
    for plant, months in PPT[2025].items():
        got = bp.millions(won[plant], 12)
        for month, expected in months.items():
            assert abs(got[month - 1] - expected) <= 1, (plant, month, got[month - 1], expected)


def test_2025_bulk_amounts_match_known_rows():
    totals = bp.item_totals(2025, 12)
    assert abs(int(round(totals["B:CO2"]["금액"])) - 9146015313) <= 100
    assert int(round(totals["B:CO2"]["수량"])) == 40436132
    assert abs(int(round(totals["B:O2"]["금액"])) - 1155094256) <= 100
    assert int(round(totals["B:LPG"]["수량"])) == 196480
    assert bp.client_kg(2025, 12, "라쿨", "LPG") == 192742


def test_page_save_keeps_theme_and_writes_the_chart(tmp_path: Path | None = None):
    template = _template()
    root = tmp_path or Path("/tmp/biz_proposal_test")
    if root.exists():
        import shutil
        shutil.rmtree(root)
    root.mkdir(parents=True)
    dest = root / "2026년 사업계획.pptx"
    bp.save_page(5, 2026, 12, path=dest, template=template)
    with zipfile.ZipFile(template) as old, zipfile.ZipFile(dest) as new:
        assert old.read("ppt/theme/theme1.xml") == new.read("ppt/theme/theme1.xml")
        assert old.read("ppt/slides/slide1.xml") == new.read("ppt/slides/slide1.xml")
        slide5 = new.read("ppt/slides/slide5.xml").decode("utf-8")
    assert "1,152" in slide5
    bp.save_page(4, 2026, 12, plan={"아산": 6000, "화성": 15764, "전체": 21764}, path=dest, template=template)
    with zipfile.ZipFile(dest) as book:
        chart = book.read("ppt/charts/chart1.xml").decode("utf-8")
        assert book.read("ppt/slides/slide5.xml").decode("utf-8") == slide5
        assert book.read("ppt/theme/theme1.xml") == zipfile.ZipFile(template).read("ppt/theme/theme1.xml")
    assert "<c:v>6000</c:v>" in chart
    assert "<c:v>4417</c:v>" in chart
    import shutil
    shutil.rmtree(root)


def test_plan_months_match_the_template_sum():
    assert sum(bp.PLAN_ASAN) == 5695
    assert sum(bp.PLAN_HWA) == 15765


if __name__ == "__main__":
    test_sixteen_pages_each_have_a_save_button()
    test_plan_months_match_the_template_sum()
    test_2025_plant_months_match_the_plan()
    test_2025_bulk_amounts_match_known_rows()
    test_page_save_keeps_theme_and_writes_the_chart()
    print("ok")
