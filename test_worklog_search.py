"""일일업무일지 통합검색 — 거래처·내용 단어."""
from __future__ import annotations

import tempfile
from datetime import date
from pathlib import Path

from openpyxl import Workbook

import worklog_tab as wl


def test_words_match_client_or_content():
    rows = [
        {"date": "2026-07-03", "page": 1, "client": "삼성물산", "content": "밸브 교체"},
        {"date": "2026-07-04", "page": 1, "client": "한빛", "content": "입금 확인"},
    ]
    assert wl.worklog_search_words("  삼성,  밸브 ") == ["삼성", "밸브"]
    assert wl.worklog_search_words("   ") == []
    hit = wl.search_worklog_rows(rows, "밸브")
    assert [r["date"] for r in hit] == ["2026-07-03"]
    assert wl.search_worklog_rows(rows, "삼성 밸브")[0]["client"] == "삼성물산"
    assert wl.search_worklog_rows(rows, "삼성 입금") == []
    assert wl.search_worklog_rows(rows, "") == []


def test_month_file_finds_client_content_and_extra_page():
    with tempfile.TemporaryDirectory() as td:
        path = str(Path(td) / "7월.xlsx")
        wb = Workbook()
        ws = wb.active
        ws.title = "3"
        ws["C5"] = "2026-07-03 (금)"
        ws["C8"] = "예성이엔지"
        ws["G8"] = "탱크 점검"
        ws["G9"] = "패킹 교체"
        extra = wb.create_sheet("3p2")
        extra["C8"] = "예성이엔지"
        extra["G8"] = "야간 출동"
        other = wb.create_sheet("4")
        other["C8"] = "다른업체"
        other["G8"] = "견적"
        wb.save(path)
        wb.close()
        rows = wl.worklog_search_rows_from_month(path, 2026, 7)
        by_content = {r["content"]: r for r in rows}
        assert "탱크 점검\n패킹 교체" in by_content
        assert by_content["탱크 점검\n패킹 교체"]["client"] == "예성이엔지"
        assert by_content["탱크 점검\n패킹 교체"]["date"] == "2026-07-03"
        night = wl.search_worklog_rows(rows, "야간")
        assert len(night) == 1 and night[0]["page"] == 2
        assert wl.search_worklog_rows(rows, "예성 패킹")[0]["date"] == "2026-07-03"
        assert wl.search_worklog_rows(rows, "견적")[0]["date"] == "2026-07-04"


def test_search_tab_is_inside_worklog():
    src = Path("worklog_tab.py").read_text(encoding="utf-8")
    assert '["작성", "통합검색"]' in src
    assert 'if panel == "통합검색"' in src
    assert "def _render_worklog_search" in src
    assert "거래처와 내용에서 단어를 찾습니다" in src
