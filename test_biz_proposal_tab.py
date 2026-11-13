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
    assert "edit_col, view_col = st.columns([1, 1.15])" in src
    assert "st.image" in src
    assert "_show_page_preview" in src


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


def test_page2_policy_and_review_are_editable(tmp_path: Path | None = None):
    src = Path("biz_proposal_tab.py").read_text(encoding="utf-8")
    assert "1. {POLICY_TITLE}" in src
    assert "년 사업계획 대비" in src
    template = _template()
    root = tmp_path or Path("/tmp/biz_proposal_toc")
    if root.exists():
        import shutil
        shutil.rmtree(root)
    dest = root / "27년 사업계획서.pptx"
    dest.parent.mkdir(parents=True)
    import shutil
    shutil.copy(template, dest)
    bp.save_page(
        2,
        2027,
        12,
        policy_lines=["방침가", "방침나"],
        review_lines=["판매하나", "담당둘"],
        strategy_lines=["전략셋"],
        path=dest,
        template=template,
    )
    slide2 = zipfile.ZipFile(dest).read("ppt/slides/slide2.xml").decode("utf-8")
    assert bp._shape_lines(slide2, "policy") == ["방침가", "방침나"]
    assert bp._shape_lines(slide2, "review") == ["판매하나", "담당둘"]
    assert bp._shape_lines(slide2, "strategy") == ["전략셋"]
    review = next(block for block in re.findall(r"<p:sp\b.*?</p:sp>", slide2, flags=re.S) if "사업계획 대비" in block)
    assert ">27</a:t>" in review and ">26</a:t>" in review
    policy = next(block for block in re.findall(r"<p:sp\b.*?</p:sp>", slide2, flags=re.S) if "방침가" in block)
    assert "영업방침 및 영업목표" in policy
    with zipfile.ZipFile(template) as old, zipfile.ZipFile(dest) as new:
        assert old.read("ppt/theme/theme1.xml") == new.read("ppt/theme/theme1.xml")
        assert old.read("ppt/slides/slide1.xml") == new.read("ppt/slides/slide1.xml")
    shutil.rmtree(root)


def test_page3_three_sections_are_editable(tmp_path: Path | None = None):
    import shutil

    src = Path("biz_proposal_tab.py").read_text(encoding="utf-8")
    assert "1. {DIRECTIVE_LABEL}" in src
    assert "2. {FOCUS_LABEL}" in src
    assert "3. {TASK_LABEL}" in src
    template = _template()
    root = tmp_path or Path("/tmp/biz_proposal_goals")
    if root.exists():
        shutil.rmtree(root)
    root.mkdir(parents=True)
    dest = root / "27년 사업계획서.pptx"
    real = bp.proposal_path(2026)
    before = real.read_bytes() if real is not None and real.is_file() else None
    bp.save_page(
        3,
        2027,
        12,
        directive_lines=["방침하나", "방침둘"],
        focus_lines=["목표하나"],
        task_lines=["과제하나", "과제둘", "과제셋"],
        path=dest,
        template=template,
    )
    slide3 = zipfile.ZipFile(dest).read("ppt/slides/slide3.xml").decode("utf-8")
    assert bp._goal_lines(slide3, "directive") == ["방침하나", "방침둘"]
    assert bp._goal_lines(slide3, "focus") == ["목표하나"]
    assert bp._goal_lines(slide3, "task") == ["과제하나", "과제둘", "과제셋"]
    directive = next(block for block in re.findall(r"<p:sp\b.*?</p:sp>", slide3, flags=re.S) if "방침하나" in block)
    assert "buAutoNum" in directive
    assert "영업방침" in bp._compact_title("".join(bp._para_texts(re.findall(r"<a:p\b.*?</a:p>", directive, flags=re.S)[0])))
    bp.save_page(3, 2027, 12, task_lines=[], path=dest, template=template)
    slide3 = zipfile.ZipFile(dest).read("ppt/slides/slide3.xml").decode("utf-8")
    assert bp._goal_lines(slide3, "directive") == ["방침하나", "방침둘"]
    assert bp._goal_lines(slide3, "focus") == ["목표하나"]
    assert bp._goal_lines(slide3, "task") == []
    assert "추진과제" in slide3.replace(" ", "")
    with zipfile.ZipFile(template) as old, zipfile.ZipFile(dest) as new:
        assert old.read("ppt/theme/theme1.xml") == new.read("ppt/theme/theme1.xml")
        assert old.read("ppt/slides/slide2.xml") == new.read("ppt/slides/slide2.xml")
    if before is not None:
        assert real.read_bytes() == before
    shutil.rmtree(root)


def test_page7_bulk_lines_follow_the_total_growth(tmp_path: Path | None = None):
    import shutil

    src = Path("biz_proposal_tab.py").read_text(encoding="utf-8")
    assert "벌크운영 계획" in src
    assert "증가율" in src
    template = _template()
    root = tmp_path or Path("/tmp/biz_proposal_bulk")
    if root.exists():
        shutil.rmtree(root)
    root.mkdir(parents=True)
    dest = root / "27년 사업계획서.pptx"
    real = bp.proposal_path(2026)
    before = real.read_bytes() if real is not None and real.is_file() else None
    pct = bp.bulk_growth_percent(2026, 12, bp.PLAN_TONS)
    prev = bp.tons_of(2025, 12)["합계"]
    assert pct == int(round((sum(bp.PLAN_TONS.values()) - prev) / prev * 100))
    bp.save_page(
        7,
        2026,
        12,
        bulk_lines=["- 26년 벌크 전년대비 목표 17% 증대", "- 26년 신규 항목"],
        path=dest,
        template=template,
    )
    slide7 = zipfile.ZipFile(dest).read("ppt/slides/slide7.xml").decode("utf-8")
    assert bp._bulk_lines_in(slide7) == [f"- 26년 벌크 전년대비 목표 {pct}% 증대", "- 26년 신규 항목"]
    assert f">{pct}%</a:t>" in slide7
    bp.save_page(7, 2027, 12, bulk_lines=["- 26년 벌크 전년대비 목표 1% 증대"], path=dest, template=template)
    slide7 = zipfile.ZipFile(dest).read("ppt/slides/slide7.xml").decode("utf-8")
    tuned = bp._bulk_lines_in(slide7)
    assert tuned is not None and tuned[0].startswith("- 27년 벌크 전년대비 목표 ")
    assert "신규 항목" not in slide7
    assert "벌크운영 계획" in slide7
    with zipfile.ZipFile(template) as old, zipfile.ZipFile(dest) as new:
        assert old.read("ppt/theme/theme1.xml") == new.read("ppt/theme/theme1.xml")
        assert old.read("ppt/slides/slide1.xml") == new.read("ppt/slides/slide1.xml")
    if before is not None:
        assert real.read_bytes() == before
    shutil.rmtree(root)


def test_page8_clients_aggregate_without_changing_the_format(tmp_path: Path | None = None):
    import shutil

    src = Path("biz_proposal_tab.py").read_text(encoding="utf-8")
    assert "거래처 추가" in src
    assert "page8_figures" in src
    template = _template()
    root = tmp_path or Path("/tmp/biz_proposal_page8")
    if root.exists():
        shutil.rmtree(root)
    root.mkdir(parents=True)
    dest = root / "26년 사업계획서.pptx"
    real = bp.proposal_path(2026)
    before = real.read_bytes() if real is not None and real.is_file() else None
    original = zipfile.ZipFile(template).read("ppt/slides/slide8.xml").decode("utf-8")
    rows = [
        {"product": "LO2", "client": "가스코아산", "budget": "", "note": ""},
        {"product": "LO2", "client": "추가거래처", "budget": "1000", "note": "메모"},
        {"product": "LPG", "client": "신규 거래선", "budget": "5000", "note": ""},
    ]
    bp.save_page(8, 2026, 12, page8_rows=rows, path=dest, template=template)
    slide8 = zipfile.ZipFile(dest).read("ppt/slides/slide8.xml").decode("utf-8")
    parsed = bp._page8_rows_in(slide8)
    assert [(row["product"], row["client"]) for row in parsed] == [
        ("LO2", "가스코아산"),
        ("LO2", "추가거래처"),
        ("LPG", "신규 거래선"),
    ]
    assert 'rowSpan="2"' in slide8
    assert bp.client_kg(2025, 12, "태성 EPS", "LCO2") == bp.client_kg(2025, 12, "태성EPS", "LCO2")
    actual = bp.client_kg(2025, 12, "가스코아산", "LO2")
    assert actual is not None
    assert bp._fmt_int(actual) in slide8
    assert bp._fmt_int(-round(actual)) in slide8
    gas_row = next(row for row in re.findall(r"<a:tr\b.*?</a:tr>", slide8, flags=re.S) if "가스코아산" in row)
    assert "FF0000" in gas_row
    added = next(row for row in re.findall(r"<a:tr\b.*?</a:tr>", slide8, flags=re.S) if "추가거래처" in row)
    assert "메모" in added
    added_delta = re.findall(r"<a:tc\b.*?</a:tc>", added, flags=re.S)[4]
    assert "FF0000" not in added_delta
    assert ">1,000</a:t>" in added
    new_row = next(row for row in re.findall(r"<a:tr\b.*?</a:tr>", slide8, flags=re.S) if "신규 거래선" in row)
    assert ">-" in new_row or ">-</a:t>" in new_row or ">-</a:t>" in new_row
    assert ">5,000</a:t>" in new_row
    assert re.search(r"<a:tblGrid>.*?</a:tblGrid>", original, flags=re.S).group(0) == re.search(
        r"<a:tblGrid>.*?</a:tblGrid>", slide8, flags=re.S
    ).group(0)
    assert 'h="417189"' in slide8
    assert "HY견고딕" in added
    with zipfile.ZipFile(template) as old, zipfile.ZipFile(dest) as new:
        assert old.read("ppt/theme/theme1.xml") == new.read("ppt/theme/theme1.xml")
        assert old.read("ppt/slides/slide1.xml") == new.read("ppt/slides/slide1.xml")
        assert old.read("ppt/slides/slide7.xml") == new.read("ppt/slides/slide7.xml")
    if before is not None:
        assert real.read_bytes() == before
    shutil.rmtree(root)


def test_preview_deck_updates_a_copy_and_leaves_the_saved_file():
    import shutil

    src_template = _template()
    template_before = src_template.read_bytes()
    real = bp.proposal_path(2026)
    before = real.read_bytes() if real is not None and real.is_file() else None
    root = Path("/tmp/biz_proposal_preview_test")
    if root.exists():
        shutil.rmtree(root)
    root.mkdir(parents=True)
    dest = root / "preview.pptx"
    specs = [
        {"page": 1, "year": 2026, "month": 12, "plan": None, "tons": None, "asan": None, "hwa": None, "policy": None, "review": None, "strategy": None},
        {"page": 2, "year": 2027, "month": 12, "plan": None, "tons": None, "asan": None, "hwa": None, "policy": ["방침가"], "review": ["판매하나"], "strategy": ["전략셋"]},
    ]
    original = bp._preview_base
    bp._preview_base = lambda: src_template
    try:
        bp._build_preview_deck(specs, dest)
    finally:
        bp._preview_base = original
    slide2 = zipfile.ZipFile(dest).read("ppt/slides/slide2.xml").decode("utf-8")
    assert "방침가" in slide2
    assert src_template.read_bytes() == template_before
    if before is not None:
        assert real.read_bytes() == before
    shutil.rmtree(root)


def test_client_lookup_uses_one_index_and_page8_preview_is_the_only_change():
    assert bp.client_kg(2025, 12, "태성 EPS", "LCO2") == bp.client_kg(2025, 12, "태성EPS", "LCO2")
    assert bp.client_kg(2025, 12, "제원산업", "LCO2") > 0
    assert bp.client_kg(2025, 12, "제원사업", "LCO2") == 0.0
    before = [
        {"page": 1, "year": 2026, "page8": None},
        {"page": 8, "year": 2026, "page8": [{"client": "가스텍"}]},
    ]
    after = [
        {"page": 1, "year": 2026, "page8": None},
        {"page": 8, "year": 2026, "page8": [{"client": "제원산업"}]},
    ]
    assert bp._changed_preview_pages(before, after) == [8]
    moved = [
        {"page": 1, "year": 2027, "page8": None},
        {"page": 8, "year": 2026, "page8": [{"client": "제원산업"}]},
    ]
    assert bp._changed_preview_pages(before, moved) == [1, 8]


def test_page10_new_clients_fill_gas_and_growth_is_calculated():
    import shutil

    src = Path("biz_proposal_tab.py").read_text(encoding="utf-8")
    assert "전체 적용" in src
    assert "staff_new_clients" in src
    news = bp.staff_new_clients(2026, 12, "박준우")
    assert news
    assert all(row["gas"] for row in news)
    gong = next(row for row in news if "공성기계" in row["client"])
    assert "LGC" in gong["gas"]
    assert bp.volume_change_text(2026, 12, "두산판금") == "▲38% 증대"
    assert bp.staff_top_clients(2026, 12, "박준우")[0]["client"] == "두산판금"
    template = _template()
    root = Path("/tmp/biz_proposal_page10")
    if root.exists():
        shutil.rmtree(root)
    root.mkdir(parents=True)
    dest = root / "26년 사업계획서.pptx"
    real = bp.proposal_path(2026)
    before = real.read_bytes() if real is not None and real.is_file() else None
    original = zipfile.ZipFile(template).read("ppt/slides/slide10.xml").decode("utf-8")
    groups = [
        {
            "staff": "박준우 대리(25년 3월 입사)",
            "news": [{"client": gong["client"], "gas": gong["gas"]}, {"client": "추가신규", "gas": "LO2(BULK)"}],
            "growth": [{"client": "두산판금"}],
        }
    ]
    bp.save_page(10, 2026, 12, page10_groups=groups, path=dest, template=template)
    slide10 = zipfile.ZipFile(dest).read("ppt/slides/slide10.xml").decode("utf-8")
    parsed = bp._page10_groups_in(slide10)
    assert parsed[0]["staff"] == "박준우 대리(25년 3월 입사)"
    assert parsed[0]["news"][0]["gas"] == gong["gas"]
    assert parsed[0]["growth"][0]["client"] == "두산판금"
    row = next(item for item in re.findall(r"<a:tr\b.*?</a:tr>", slide10, flags=re.S) if "두산판금" in item)
    assert "▲38% 증대" in row
    assert "93백" not in row
    assert 'rowSpan="2"' in slide10
    headed = bp._apply_page10_heading(slide10, 2027)
    assert re.search(r">\s*26<", headed)
    assert "25년 3월" in headed
    assert re.search(r"<a:tblGrid>.*?</a:tblGrid>", original, flags=re.S).group(0) == re.search(
        r"<a:tblGrid>.*?</a:tblGrid>", slide10, flags=re.S
    ).group(0)
    with zipfile.ZipFile(template) as old, zipfile.ZipFile(dest) as new:
        assert old.read("ppt/theme/theme1.xml") == new.read("ppt/theme/theme1.xml")
        assert old.read("ppt/slides/slide1.xml") == new.read("ppt/slides/slide1.xml")
        assert old.read("ppt/slides/slide8.xml") == new.read("ppt/slides/slide8.xml")
    if before is not None:
        assert real.read_bytes() == before
    shutil.rmtree(root)


def test_page11_uses_the_same_staff_editor_and_growth_text():
    import shutil

    src = Path("biz_proposal_tab.py").read_text(encoding="utf-8")
    assert "page11_groups" in src
    assert "read_page10_groups(year, page)" in src
    news = bp.staff_new_clients(2026, 12, "김효동")
    top = bp.staff_top_clients(2026, 12, "김효동")
    assert news and news[0]["gas"]
    assert top and top[0]["change"] == bp.volume_change_text(2026, 12, top[0]["client"])
    template = _template()
    original = zipfile.ZipFile(template).read("ppt/slides/slide11.xml").decode("utf-8")
    assert bp._page10_groups_in(original)[0]["staff"].startswith("김효동")
    root = Path("/tmp/biz_proposal_page11")
    if root.exists():
        shutil.rmtree(root)
    root.mkdir(parents=True)
    dest = root / "26년 사업계획서.pptx"
    real = bp.proposal_path(2026)
    before = real.read_bytes() if real is not None and real.is_file() else None
    groups = [
        {
            "staff": "김효동 과장(25년 6월 입사)",
            "news": [{"client": news[0]["client"], "gas": news[0]["gas"]}, {"client": news[1]["client"], "gas": news[1]["gas"]}],
            "growth": [{"client": top[0]["client"]}],
        }
    ]
    bp.save_page(11, 2026, 12, page11_groups=groups, path=dest, template=template)
    slide11 = zipfile.ZipFile(dest).read("ppt/slides/slide11.xml").decode("utf-8")
    parsed = bp._page10_groups_in(slide11)
    assert parsed[0]["news"][0]["gas"] == news[0]["gas"]
    assert parsed[0]["growth"][0]["client"] == top[0]["client"]
    assert top[0]["change"] in slide11
    assert "신규계약" not in slide11
    assert 'rowSpan="2"' in slide11
    assert "_2" in slide11
    assert "25년 6월" in slide11
    assert re.search(r"<a:tblGrid>.*?</a:tblGrid>", original, flags=re.S).group(0) == re.search(
        r"<a:tblGrid>.*?</a:tblGrid>", slide11, flags=re.S
    ).group(0)
    with zipfile.ZipFile(template) as old, zipfile.ZipFile(dest) as new:
        assert old.read("ppt/theme/theme1.xml") == new.read("ppt/theme/theme1.xml")
        assert old.read("ppt/slides/slide1.xml") == new.read("ppt/slides/slide1.xml")
        assert old.read("ppt/slides/slide10.xml") == new.read("ppt/slides/slide10.xml")
    if before is not None:
        assert real.read_bytes() == before
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
    test_page2_policy_and_review_are_editable()
    test_page3_three_sections_are_editable()
    test_page7_bulk_lines_follow_the_total_growth()
    test_page8_clients_aggregate_without_changing_the_format()
    test_client_lookup_uses_one_index_and_page8_preview_is_the_only_change()
    test_page10_new_clients_fill_gas_and_growth_is_calculated()
    test_page11_uses_the_same_staff_editor_and_growth_text()
    test_preview_deck_updates_a_copy_and_leaves_the_saved_file()
    print("ok")
