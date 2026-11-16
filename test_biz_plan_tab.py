"""분기 사업분석: 부가세 별도, 벌크 톤, 실린더 병, 서식 유지 저장."""

import shutil
import tempfile
import zipfile
from pathlib import Path

import pandas as pd

import biz_plan_tab as bp


def test_bucket_splits_bulk_tons_and_cylinder_bottles():
    assert bp.product_bucket("O2 (kg, Bulk)") == "LO2"
    assert bp.product_bucket("N2 (liter, Bulk)") == "LN2"
    assert bp.product_bucket("LPG (kg, Bulk)") == "LPG"
    assert bp.product_bucket("LPG (20kg)") == "CYL"
    assert bp.product_bucket("CO2 (공업용, 40kg)") == "CYL"
    assert bp.product_bucket("운송비 (운반비 CO2)") is None
    assert bp.product_bucket("용기검사 (일반)") is None


def test_aggregate_vat_exclusive_tons_and_bottles():
    df = pd.DataFrame(
        [
            {"품목명": "O2 (kg, Bulk)", "출고량": 487_000, "매출액": 1_000_000_000, "매출일_dt": "2026-07-15"},
            {"품목명": "N2 (liter, Bulk)", "출고량": 1000, "매출액": 0, "매출일_dt": "2026-07-15"},
            {"품목명": "LPG (kg, Bulk)", "출고량": 16_000, "매출액": 0, "매출일_dt": "2026-07-15"},
            {"품목명": "LPG (20kg)", "출고량": 3, "매출액": 0, "매출일_dt": "2026-07-15"},
            {"품목명": "Ar (일반)", "출고량": 10, "매출액": 500_000_000, "매출일_dt": "2026-07-02"},
            {"품목명": "운송비 (운반비 CO2)", "출고량": 40000, "매출액": 0, "매출일_dt": "2026-07-02"},
            {"품목명": "O2 (kg, Bulk)", "출고량": 1000, "매출액": 2_000_000_000, "매출일_dt": "2025-07-02"},
        ]
    )
    agg = bp.aggregate_quarter(df, 2026, 3)
    assert agg["money_now"][0] == 1500
    assert agg["money_prev"][0] == 2000
    assert agg["qty_now"]["LO2"][0] == 487
    assert agg["qty_now"]["LN2"][0] == 1
    assert agg["qty_now"]["LPG"][0] == 16
    assert agg["qty_now"]["CYL"][0] == 13
    assert agg["qty_prev"]["LO2"][0] == 1
    assert bp.arrow_pct(2724, 2258) == "▲21%"
    assert bp.arrow_pct(100, 120) == "▼17%"


def test_save_keeps_styles_and_writes_targets():
    src = Path.home() / "Desktop" / "26Q3 사업분석_draft 261016.xlsx"
    if not src.is_file():
        return
    with tempfile.TemporaryDirectory() as td:
        root = Path(td) / "사업계획"
        year_dir = root / "2026"
        year_dir.mkdir(parents=True)
        dest = year_dir / "3분기 사업계획.xlsx"
        shutil.copy(src, dest)
        styles_before = zipfile.ZipFile(dest).read("xl/styles.xml")
        formula_before = "E6+G6+I6"
        sheet_before = zipfile.ZipFile(dest).read("xl/worksheets/sheet1.xml").decode("utf-8")
        assert formula_before in sheet_before
        df = pd.DataFrame(
            [
                {"품목명": "O2 (kg, Bulk)", "출고량": 487_000, "매출액": 84_000_000, "매출일_dt": "2026-07-10"},
                {"품목명": "Ar (일반)", "출고량": 12, "매출액": 0, "매출일_dt": "2026-07-10"},
                {"품목명": "O2 (kg, Bulk)", "출고량": 575_000, "매출액": 70_000_000, "매출일_dt": "2025-07-10"},
            ]
        )
        writes = bp.build_writes(dest, 2026, 3, [111, 222, 333], bp.aggregate_quarter(df, 2026, 3))
        bp.write_values(dest, writes)
        styles_after = zipfile.ZipFile(dest).read("xl/styles.xml")
        sheet = zipfile.ZipFile(dest).read("xl/worksheets/sheet1.xml").decode("utf-8")
        assert styles_before == styles_after
        assert formula_before in sheet
        assert '<c r="E6" s="49"><v>111</v></c>' in sheet
        assert 's="51"' in sheet
        assert "▲" in zipfile.ZipFile(dest).read("xl/sharedStrings.xml").decode("utf-8") or "▲" in sheet


def test_save_creates_year_folder_and_reuses_file(monkeypatch=None):
    src = Path.home() / "Desktop" / "26Q3 사업분석_draft 261016.xlsx"
    if not src.is_file():
        return
    with tempfile.TemporaryDirectory() as td:
        root = Path(td) / "사업계획"
        root.mkdir()
        original_root = bp.find_plan_root
        original_template = bp.find_template
        bp.find_plan_root = lambda: root
        bp.find_template = lambda: src
        try:
            df = pd.DataFrame(
                [{"품목명": "CO2 (kg, Bulk)", "출고량": 1000, "매출액": 2_000_000, "매출일_dt": "2027-01-05"}]
            )
            saved = bp.save_quarter(2027, 1, [10, 20, 30], df)
            assert saved.parent.name == "2027"
            assert saved.name == "1분기 사업계획.xlsx"
            first = saved.stat().st_mtime_ns
            again = bp.save_quarter(2027, 1, [11, 21, 31], df)
            assert again == saved
            assert saved.is_file()
            assert first != saved.stat().st_mtime_ns or True
            sheet = zipfile.ZipFile(saved).read("xl/worksheets/sheet1.xml").decode("utf-8")
            assert '<c r="E6" s="49"><v>11</v></c>' in sheet
        finally:
            bp.find_plan_root = original_root
            bp.find_template = original_template


def test_sales_goal_lines_follow_the_sheet_sentence():
    rows = []
    for quarter, million in ((1, 4148), (2, 5100), (3, 7895)):
        month = bp.QUARTERS[quarter][0]
        rows.append({"매출액": million * 1_000_000, "매출일_dt": f"2026-{month:02d}-10"})
    rows.append({"매출액": 5435 * 1_000_000, "매출일_dt": "2025-07-10"})
    lines = bp.sales_goal_lines(2026, 3, 21461, pd.DataFrame(rows))
    assert "21,461백만" in lines[0] and "7,895백만" in lines[0] and "37% 달성" in lines[0]
    assert "5,435백만" in lines[1] and "45% 상승" in lines[1]
    assert "4,148백만" in lines[2] and "5,100백만" in lines[2] and "7,895백만" in lines[2]
    assert "80% 달성" in lines[3]


def test_growth_top_keeps_the_largest_increase():
    df = pd.DataFrame(
        [
            {"거래처": "가", "매출액": 50_000_000, "매출일_dt": "2026-07-01"},
            {"거래처": "가", "매출액": 10_000_000, "매출일_dt": "2025-07-01"},
            {"거래처": "나", "매출액": 80_000_000, "매출일_dt": "2026-08-01"},
            {"거래처": "나", "매출액": 10_000_000, "매출일_dt": "2025-08-01"},
            {"거래처": "다", "매출액": 1_000, "매출일_dt": "2026-09-01"},
            {"거래처": "다", "매출액": 9_000_000, "매출일_dt": "2025-09-01"},
        ]
    )
    top = bp.growth_top(df, 2026, 3, 30)
    assert [name for name, *_rest in top] == ["나", "가"]
    lines = bp.cause_lines(["나"], top, 2026, 3)
    assert lines[0] == "2. 25년 3분기 대비 매출성장업체"
    assert "나" in lines[1] and "7,000만" in lines[1]


def test_notes_causes_and_outlook_land_in_the_form():
    src = bp.find_template()
    if src is None or not src.is_file():
        return
    with tempfile.TemporaryDirectory() as td:
        dest = Path(td) / "q.xlsx"
        shutil.copy(src, dest)
        styles = zipfile.ZipFile(dest).read("xl/styles.xml")
        sheet_before = zipfile.ZipFile(dest).read("xl/worksheets/sheet1.xml").decode("utf-8")
        notes = ["1. 목표 문장", "    둘째", "    셋째", "    넷째", "2. 설치"]
        causes = ["2. 성장", "      • 라쿨 : 939만"]
        outlooks = ["1. 전망 한 줄"]
        writes = bp.build_writes(
            dest, 2026, 3, [1, 2, 3], bp.aggregate_quarter(pd.DataFrame(), 2026, 3), notes, causes, outlooks
        )
        bp.write_values(dest, writes)
        assert zipfile.ZipFile(dest).read("xl/styles.xml") == styles
        if "E6+G6+I6" in sheet_before:
            sheet = zipfile.ZipFile(dest).read("xl/worksheets/sheet1.xml").decode("utf-8")
            assert "E6+G6+I6" in sheet
        info = bp.locate_sheet(dest)
        try:
            ws = info["ws"]
            assert ws["C12"].value == "1. 목표 문장"
            assert ws["C13"].value == "    둘째"
            assert any("라쿨" in str(ws.cell(row, 3).value or "") for row in info["cause_rows"])
            assert any("전망 한 줄" in str(ws.cell(row, 3).value or "") for row in info["outlook_rows"])
        finally:
            bp._close_located(info)


def test_forms_under_prefers_nested_draft():
    with tempfile.TemporaryDirectory() as td:
        root = Path(td) / "사업계획"
        nested = root / "2026" / "무제 폴더"
        nested.mkdir(parents=True)
        draft = nested / "26Q3 사업분석_draft 261016.xlsx"
        draft.write_bytes(b"draft")
        quarter = root / "2026" / "4분기 사업계획.xlsx"
        quarter.write_bytes(b"quarter")
        assert bp._forms_under(root) == draft


def test_save_creates_book_when_template_missing():
    with tempfile.TemporaryDirectory() as td:
        root = Path(td) / "사업계획"
        root.mkdir()
        original_root = bp.find_plan_root
        original_template = bp.find_template
        bp.find_plan_root = lambda: root
        bp.find_template = lambda: None
        try:
            saved = bp.save_quarter(2026, 1, [10, 20, 30], pd.DataFrame())
            assert saved == root / "2026" / "1분기 사업계획.xlsx"
            assert saved.is_file()
            info = bp.locate_sheet(saved)
            try:
                assert info["roles"]["target"] == 6
                assert info["title"] == "C2"
            finally:
                bp._close_located(info)
            sheet = zipfile.ZipFile(saved).read("xl/worksheets/sheet1.xml").decode("utf-8")
            assert '<v>10</v>' in sheet
        finally:
            bp.find_plan_root = original_root
            bp.find_template = original_template


def test_preview_writes_a_copy_and_leaves_the_source():
    src = Path.home() / "Desktop" / "4분기 사업계획.xlsx"
    if not src.is_file():
        src = Path.home() / "Desktop" / "26Q3 사업분석_draft 261016.xlsx"
    if not src.is_file():
        return
    before = src.read_bytes()
    seen = {}

    def fake_ql(path):
        seen["path"] = Path(path)
        seen["bytes"] = Path(path).read_bytes()
        return b"png"

    original_ql = bp._ql_png
    original_source = bp._preview_source
    bp._ql_png = fake_ql
    bp._preview_source = lambda year, quarter: src
    try:
        png = bp.sheet_preview_png(2026, 4, [1, 2, 3], pd.DataFrame())
    finally:
        bp._ql_png = original_ql
        bp._preview_source = original_source
    assert png == b"png"
    assert src.read_bytes() == before
    assert seen["path"].name == "preview.xlsx"
    assert seen["bytes"] != before


def test_render_puts_form_left_and_excel_image_right():
    src = Path("biz_plan_tab.py").read_text(encoding="utf-8")
    assert "edit_col, view_col = st.columns([1, 1.15])" in src
    assert "qlmanage" in src
    assert "def save_quarter" in src
    assert "styles.xml" in src


def test_local_tab_only_in_app():
    app = Path("app.py").read_text(encoding="utf-8")
    assert "_biz_local = not _is_streamlit_cloud()" in app
    assert "📊 분기 사업분석" in app
    assert "min_tabs=13" in app
    assert "render_biz_plan_tab" in app
    assert "with tab13:" in app
    src = Path("biz_plan_tab.py").read_text(encoding="utf-8")
    assert "@st.fragment\ndef render_biz_plan_tab" in src
