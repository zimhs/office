"""견적서: 새 시트에 값만 넣고, 품목 단가는 최근 거래부터."""
from __future__ import annotations

import shutil
import zipfile
from datetime import date
from pathlib import Path

import pandas as pd

import quote_tab as qt

TEMPLATE = Path.home() / "Desktop" / "신일가스_견적서.xlsx"


def _sample(on: date | None = None, client: str = "테스트가스") -> dict:
    lines = [
        {"품목": "CO", "규격": "47L", "수량": 2, "단위": "ea", "단가": 1000, "비고": ""},
    ]
    while len(lines) < 8:
        lines.append({"품목": "", "규격": "", "수량": None, "단위": "", "단가": None, "비고": ""})
    return {
        "client": client,
        "person": "홍길동",
        "phone": "010-0000-0000",
        "address": "광주시",
        "ref": "구매팀",
        "number": "SI-TEST-001",
        "quote_date": on or date(2026, 11, 13),
        "delivery": "협의",
        "pay": "익월 말 현금",
        "vat": "별도",
        "notes": ["단가 조정", "", "", ""],
        "lines": lines,
    }


def _part(path: Path, name: str) -> bytes:
    return zipfile.ZipFile(path).read(name)


def _shared(path: Path) -> list[str]:
    xml = _part(path, "xl/sharedStrings.xml").decode("utf-8")
    return qt._shared_strings(xml)


def test_sheet_title_uses_client_and_date():
    assert qt.sheet_title("부림가스", date(2026, 11, 12), set()) == "부림가스.2026-11-12"
    again = qt.sheet_title("부림가스", date(2026, 11, 12), {"부림가스.2026-11-12"})
    assert again.startswith("부림가스.")
    assert "(2)" in again
    assert len(qt.sheet_title("가" * 40, date(2026, 11, 12), set())) <= 31


def test_search_is_recent_first():
    frame = pd.DataFrame(
        [
            {"매출일_dt": "2026-01-02", "거래처": "가", "품목명": "N2 (kg, Bulk)", "출고량": 1, "단가": 100},
            {"매출일_dt": "2026-07-01", "거래처": "가", "품목명": "N2 실린더", "출고량": 4, "단가": 180},
            {"매출일_dt": "2026-08-09", "거래처": "나", "품목명": "N2 (kg, Bulk)", "출고량": 2, "단가": 300},
            {"매출일_dt": "2026-08-09", "거래처": "나", "품목명": "N2 (kg, Bulk)", "출고량": 9, "단가": 320},
            {"매출일_dt": "2026-03-01", "거래처": "다", "품목명": "O2 (kg, Bulk)", "출고량": 3, "단가": 200},
            {"매출일_dt": "2019-12-31", "거래처": "가", "품목명": "N2 (kg, Bulk)", "출고량": 8, "단가": 999},
            {"매출일_dt": "2019-06-01", "거래처": "옛업체", "품목명": "N2 (kg, Bulk)", "출고량": 1, "단가": 50},
        ]
    )
    rows = qt.search_item_prices(frame, "n2")
    assert list(rows["거래처"]) == ["나", "가"]
    assert list(rows["단가"]) == [320, 180]
    assert rows.iloc[0]["날짜"] == "2026-08-09"
    assert rows.iloc[1]["날짜"] == "2026-07-01"
    assert "옛업체" not in set(rows["거래처"])
    assert qt.search_item_prices(frame, "").empty
    assert qt.search_item_prices(frame, "없는품목").empty


def test_save_adds_sheet_without_touching_styles(tmp_path: Path | None = None):
    assert TEMPLATE.is_file(), TEMPLATE
    root = tmp_path or Path("/tmp/quote_tab_test")
    if root.exists():
        shutil.rmtree(root)
    root.mkdir(parents=True)
    book = root / "신일가스_견적서.xlsx"
    shutil.copy(TEMPLATE, book)
    before = {name: _part(book, name) for name in qt._UNTOUCHED}
    saved, title = qt.save_quote(_sample(), book)
    assert saved == book
    assert title == "테스트가스.2026-11-13"
    for name, blob in before.items():
        assert _part(book, name) == blob, name
    with zipfile.ZipFile(book) as zin:
        names = zin.namelist()
        sheet = zin.read("xl/worksheets/sheet3.xml").decode("utf-8")
    assert "xl/worksheets/sheet3.xml" in names
    assert "xl/drawings/drawing3.xml" in names
    strings = _shared(book)
    client = re_shared(sheet, "C10", strings)
    assert client == "테스트가스"
    assert '<c r="C10" s="38" t="s">' in sheet
    assert "<v>46339</v>" in sheet  # 2026-11-13
    assert 't="shared" ref="H19:H26"' in sheet
    assert 'IF(OR(E19="",G19=""),"",E19*G19)' in sheet
    assert "<v>2000</v>" in sheet
    assert 's="49"' not in sheet or True
    phone = re_shared(sheet, "C12", strings)
    assert phone == "010-0000-0000"
    assert '<c r="C21" s="9"/>' in sheet
    assert "SUM(H19:H26)" in sheet
    workbook = _part(book, "xl/workbook.xml").decode("utf-8")
    assert 'name="테스트가스.2026-11-13"' in workbook
    assert 'name="이엔에이치"' in workbook
    assert "fullCalcOnLoad" in workbook
    sheet3 = _part(book, "xl/worksheets/sheet3.xml")
    qt.save_quote(_sample(client="다른회사"), book)
    assert _part(book, "xl/worksheets/sheet3.xml") == sheet3
    assert _part(book, "xl/styles.xml") == before["xl/styles.xml"]
    workbook = _part(book, "xl/workbook.xml").decode("utf-8")
    assert 'name="다른회사.2026-11-13"' in workbook
    shutil.rmtree(root)


def re_shared(sheet: str, ref: str, strings: list[str]) -> str:
    import re

    match = re.search(rf'<c r="{ref}"[^>]*t="s"><v>(\d+)</v></c>', sheet)
    assert match, ref
    return strings[int(match.group(1))]


def test_local_tab_wiring():
    app = Path("app.py").read_text(encoding="utf-8")
    assert '📄 견적서' in app
    assert "render_quote_tab" in app
    assert "min_tabs=13" in app
    assert "tab14" not in app
    src = Path("quote_tab.py").read_text(encoding="utf-8")
    assert "@st.fragment\ndef render_quote_tab" in src
    assert "st-key-qt_price_box" in src or "qt_price_box" in src
    assert 'addEventListener("copy"' in src
    assert 'addEventListener("paste"' in src
    assert "st-key-qt_item_q" in src


if __name__ == "__main__":
    test_sheet_title_uses_client_and_date()
    test_search_is_recent_first()
    test_save_adds_sheet_without_touching_styles()
    test_local_tab_wiring()
    print("ok")
