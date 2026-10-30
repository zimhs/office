"""구글 캘린더·할 일 — 대시보드 새 탭에 구글 월간 달력과 할 일 목록을 내장한다."""
from __future__ import annotations

import json
import os
from urllib.parse import quote, urlencode

import streamlit as st
import streamlit.components.v1 as components

GC_DIR = os.path.join("uploaded_cache", "google_calendar")
GC_STORE = os.path.join(GC_DIR, "calendars.json")
GC_KR_HOLIDAY = "ko.south_korea#holiday@group.v.calendar.google.com"
GC_CAL_OPEN = "https://calendar.google.com/calendar/u/0/r"
GC_TASKS_EMBED = (
    "https://tasks.google.com/embed/?origin=https://calendar.google.com&fullWidth=1"
)
_GC_DEFAULT_SRCS = (GC_KR_HOLIDAY,)


def _gc_load_srcs() -> list[str]:
    if os.path.exists(GC_STORE):
        try:
            with open(GC_STORE, "r", encoding="utf-8") as f:
                data = json.load(f)
            srcs = data.get("srcs") if isinstance(data, dict) else data
            if isinstance(srcs, list):
                out = [str(x).strip() for x in srcs if str(x).strip()]
                if out:
                    return out
        except Exception:
            pass
    return list(_GC_DEFAULT_SRCS)


def _gc_save_srcs(srcs: list[str]) -> None:
    os.makedirs(GC_DIR, exist_ok=True)
    clean = [s.strip() for s in srcs if s and str(s).strip()]
    with open(GC_STORE, "w", encoding="utf-8") as f:
        json.dump({"srcs": clean or list(_GC_DEFAULT_SRCS)}, f, ensure_ascii=False, indent=2)


def calendar_embed_url(srcs: list[str] | None = None, *, height: int = 780) -> str:
    """구글 캘린더 월간 내장 URL. src는 지메일 또는 캘린더 ID."""
    items = [s.strip() for s in (srcs or _gc_load_srcs()) if s and str(s).strip()]
    if not items:
        items = list(_GC_DEFAULT_SRCS)
    q = {
        "height": str(int(height)),
        "wkst": "1",
        "ctz": "Asia/Seoul",
        "hl": "ko",
        "mode": "MONTH",
        "showTitle": "0",
        "showNav": "1",
        "showDate": "1",
        "showPrint": "0",
        "showTabs": "1",
        "showCalendars": "1",
        "showTz": "0",
    }
    base = "https://calendar.google.com/calendar/embed?" + urlencode(q)
    src_q = "&".join("src=" + quote(s, safe="") for s in items)
    return base + "&" + src_q


def render_google_calendar_tab() -> None:
    """구글 캘린더(월) + 할 일 목록. 기존 탭과 상태를 공유하지 않는다."""
    srcs = _gc_load_srcs()
    with st.expander("캘린더 연결 (지메일·캘린더 ID)", expanded=False):
        raw = st.text_area(
            "한 줄에 하나. 구글 캘린더 설정 → 일정의 캘린더 통합 → 캘린더 ID.",
            value="\n".join(srcs),
            height=110,
            key="gc_srcs_area",
            help="예: 내gmail@gmail.com  /  ko.south_korea#holiday@group.v.calendar.google.com",
        )
        c_a, c_b = st.columns([1, 3])
        with c_a:
            if st.button("연결 저장", type="primary", key="gc_save_srcs"):
                lines = [ln.strip() for ln in (raw or "").splitlines() if ln.strip()]
                _gc_save_srcs(lines or list(_GC_DEFAULT_SRCS))
                st.rerun()
        with c_b:
            st.caption(
                "할 일과 내 일정은 이 브라우저에서 구글 계정으로 로그인되어 있어야 보입니다. "
                "구글이 전체 캘린더 화면 iframe을 막아서, 월간 달력 내장 + 할 일 패널로 같은 구성을 넣었습니다."
            )
    st.link_button("구글 캘린더 전체 화면", GC_CAL_OPEN, width="stretch")
    cal_h = 780
    cal_url = calendar_embed_url(_gc_load_srcs(), height=cal_h)
    left, right = st.columns([3.15, 1.12], gap="small")
    with left:
        components.iframe(cal_url, height=cal_h, scrolling=True)
    with right:
        components.iframe(GC_TASKS_EMBED, height=cal_h, scrolling=True)
