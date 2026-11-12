"""사업계획서: 16페이지를 나눠 저장하고, 2025년 실적은 원본과 맞춘다."""
from __future__ import annotations

import re
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
    if path is None:
        for year in (2026, 2027, 2025):
            candidate = bp.proposal_path(year)
            if candidate is not None and candidate.is_file():
                path = candidate
                break
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


def test_save_name_is_two_digit_and_updates_the_same_file(tmp_path: Path | None = None):
    assert bp.proposal_name(2026) == "26년 사업계획서.pptx"
    template = _template()
    root = tmp_path or Path("/tmp/biz_proposal_name")
    if root.exists():
        import shutil
        shutil.rmtree(root)
    root.mkdir(parents=True)
    existing = root / bp.proposal_name(2026)
    existing.write_bytes(b"placeholder")
    assert bp.resolve_proposal_path(root, 2026) == existing
    fresh = bp.resolve_proposal_path(root, 2027)
    assert fresh == root / "2027" / "27년 사업계획서.pptx"
    dest = root / "2026" / bp.proposal_name(2026)
    bp.save_page(1, 2027, 12, path=dest, template=template)
    with zipfile.ZipFile(dest) as book:
        slide1 = book.read("ppt/slides/slide1.xml")
    bp.save_page(5, 2026, 12, path=dest, template=template)
    with zipfile.ZipFile(dest) as book:
        assert book.read("ppt/slides/slide1.xml") == slide1
        assert "1,152" in book.read("ppt/slides/slide5.xml").decode("utf-8")
    import shutil
    shutil.rmtree(root)


def test_existing_file_saves_without_desktop_template(tmp_path: Path | None = None):
    template = _template()
    root = tmp_path or Path("/tmp/biz_proposal_notemplate")
    if root.exists():
        import shutil
        shutil.rmtree(root)
    dest = root / "2027" / "27년 사업계획서.pptx"
    dest.parent.mkdir(parents=True)
    import shutil
    shutil.copy(template, dest)
    saved = bp.save_page(
        4,
        2027,
        12,
        plan={"아산": 1111, "화성": 2222, "전체": 3333},
        path=dest,
        template=Path("/tmp/missing-business-plan-template.pptx"),
    )
    assert saved == dest
    chart = zipfile.ZipFile(dest).read("ppt/charts/chart1.xml").decode("utf-8")
    assert "<c:v>1111</c:v>" in chart
    shutil.rmtree(root)


def test_page2_strategy_lines_add_and_remove(tmp_path: Path | None = None):
    src = Path("biz_proposal_tab.py").read_text(encoding="utf-8")
    assert "항목 추가" in src and "삭제" in src
    template = _template()
    root = tmp_path or Path("/tmp/biz_proposal_strategy")
    if root.exists():
        import shutil
        shutil.rmtree(root)
    dest = root / "2027" / "27년 사업계획서.pptx"
    dest.parent.mkdir(parents=True)
    import shutil
    shutil.copy(template, dest)
    bp.save_page(2, 2027, 12, strategy_lines=["가나다", "둘"], path=dest, template=template)
    with zipfile.ZipFile(template) as old, zipfile.ZipFile(dest) as new:
        assert old.read("ppt/theme/theme1.xml") == new.read("ppt/theme/theme1.xml")
        assert old.read("ppt/slides/slide1.xml") == new.read("ppt/slides/slide1.xml")
        slide2 = new.read("ppt/slides/slide2.xml").decode("utf-8")
    assert bp.strategy_lines_in(slide2) == ["가나다", "둘"]
    shape = next(block for block in re.findall(r"<p:sp\b.*?</p:sp>", slide2, flags=re.S) if "년도 세부 영업전략" in block)
    assert ">27</a:t>" in shape
    assert "총 판매현황" in slide2
    assert "영업방침 및 영업목표" in slide2
    bp.save_page(2, 2027, 12, strategy_lines=["가", "나", "다", "라"], path=dest, template=template)
    slide2 = zipfile.ZipFile(dest).read("ppt/slides/slide2.xml").decode("utf-8")
    assert bp.strategy_lines_in(slide2) == ["가", "나", "다", "라"]
    bp.save_page(2, 2027, 12, strategy_lines=[], path=dest, template=template)
    slide2 = zipfile.ZipFile(dest).read("ppt/slides/slide2.xml").decode("utf-8")
    assert bp.strategy_lines_in(slide2) == []
    assert "년도 세부 영업전략" in slide2
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
    test_save_name_is_two_digit_and_updates_the_same_file()
    test_existing_file_saves_without_desktop_template()
    test_page2_strategy_lines_add_and_remove()
    print("ok")
