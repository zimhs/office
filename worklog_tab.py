"""일일업무일지 탭 (tab10 전용).

다른 탭·공유 헬퍼와 상태를 섞지 않는다.
iPad(Safari)에서도 입력·저장·조회가 깨지지 않도록 단순 UI + 로컬 JSON 저장.
"""

from __future__ import annotations

import json
import os
import uuid
from datetime import date, datetime
from typing import Any

import pandas as pd
import streamlit as st

CACHE_DIR = "./uploaded_cache"
WORKLOG_FILE = os.path.join(CACHE_DIR, "worklog_entries.json")
WORKLOG_COLS = [
    "id",
    "work_date",
    "staff",
    "client",
    "category",
    "title",
    "content",
    "status",
    "updated_at",
]
CATEGORIES = ["방문", "전화", "견적", "수금", "클레임", "내부", "기타"]
STATUSES = ["진행", "완료", "보류"]


def _ensure_cache_dir() -> None:
    try:
        os.makedirs(CACHE_DIR, exist_ok=True)
    except Exception:
        pass


def _empty_df() -> pd.DataFrame:
    return pd.DataFrame(columns=WORKLOG_COLS)


def _normalize_df(df: pd.DataFrame | None) -> pd.DataFrame:
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return _empty_df()
    out = df.copy()
    for col in WORKLOG_COLS:
        if col not in out.columns:
            out[col] = ""
    out = out[WORKLOG_COLS]
    for col in WORKLOG_COLS:
        out[col] = out[col].fillna("").astype(str)
    return out.reset_index(drop=True)


def load_worklog() -> pd.DataFrame:
    _ensure_cache_dir()
    if not os.path.exists(WORKLOG_FILE):
        return _empty_df()
    try:
        with open(WORKLOG_FILE, "r", encoding="utf-8") as f:
            raw = json.load(f)
        if isinstance(raw, list):
            return _normalize_df(pd.DataFrame(raw))
        if isinstance(raw, dict) and isinstance(raw.get("entries"), list):
            return _normalize_df(pd.DataFrame(raw["entries"]))
    except Exception:
        pass
    return _empty_df()


def save_worklog(df: pd.DataFrame) -> bool:
    _ensure_cache_dir()
    try:
        clean = _normalize_df(df)
        payload = {"version": 1, "entries": clean.to_dict(orient="records")}
        tmp = WORKLOG_FILE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
        os.replace(tmp, WORKLOG_FILE)
        return True
    except Exception:
        return False


def _is_touch_ui() -> bool:
    """app.py is_touch_ui와 동일 분기. 모듈 단독 실행·iPad Safari 대응."""
    try:
        if st.session_state.get("force_touch_ui") is True:
            return True
        v = st.query_params.get("touch_ui", "")
        if isinstance(v, list):
            v = v[0] if v else ""
        if str(v) == "1":
            st.session_state["force_touch_ui"] = True
            return True
    except Exception:
        pass
    try:
        cookies = getattr(st.context, "cookies", None)
        if cookies is not None and str(cookies.get("dashboard_touch", "")) == "1":
            st.session_state["force_touch_ui"] = True
            return True
    except Exception:
        pass
    try:
        import re

        headers = getattr(st.context, "headers", None)
        ua = ""
        if headers is not None:
            ua = str(headers.get("User-Agent") or headers.get("user-agent") or "")
        if ua and (
            re.search(r"iPad|iPhone|iPod", ua)
            or ("Macintosh" in ua and "Mobile" in ua)
        ):
            st.session_state["force_touch_ui"] = True
            return True
    except Exception:
        pass
    return bool(st.session_state.get("force_touch_ui"))


def _inject_worklog_touch_css() -> None:
    """일일업무일지 탭만 — 터치 타깃·입력 폭 안정화 (다른 탭 CSS 미변경)."""
    st.markdown(
        """
        <style>
        .worklog-touch-scope textarea,
        .worklog-touch-scope input,
        .worklog-touch-scope [data-baseweb="input"] input,
        .worklog-touch-scope [data-baseweb="textarea"] textarea,
        .worklog-touch-scope [data-baseweb="select"] {
            font-size: 16px !important; /* iOS 자동 줌 방지 */
            min-height: 44px !important;
        }
        .worklog-touch-scope div[data-testid="stButton"] > button,
        .worklog-touch-scope div[data-testid="stDownloadButton"] > button {
            min-height: 48px !important;
            width: 100% !important;
            font-size: 15px !important;
            font-weight: 600 !important;
        }
        .worklog-touch-scope [data-testid="stDataFrame"] {
            -webkit-overflow-scrolling: touch !important;
            max-width: 100% !important;
        }
        .worklog-card {
            border: 1px solid #E2E8F0;
            border-radius: 10px;
            background: #FFFFFF;
            padding: 12px 14px;
            margin: 0 0 10px 0;
        }
        .worklog-card .meta {
            color: #64748B;
            font-size: 12px;
            margin-bottom: 4px;
        }
        .worklog-card .title {
            color: #0F172A;
            font-size: 15px;
            font-weight: 700;
            margin-bottom: 4px;
        }
        .worklog-card .body {
            color: #334155;
            font-size: 14px;
            white-space: pre-wrap;
            word-break: break-word;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def _new_entry(
    *,
    work_date: date,
    staff: str,
    client: str,
    category: str,
    title: str,
    content: str,
    status: str,
) -> dict[str, str]:
    return {
        "id": uuid.uuid4().hex[:12],
        "work_date": work_date.isoformat(),
        "staff": (staff or "").strip() or "미지정",
        "client": (client or "").strip(),
        "category": category or "기타",
        "title": (title or "").strip() or "(제목 없음)",
        "content": (content or "").strip(),
        "status": status or "진행",
        "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }


def _filter_df(
    df: pd.DataFrame,
    *,
    start: date | None,
    end: date | None,
    staff: str,
    client_q: str,
    status: str,
) -> pd.DataFrame:
    if df.empty:
        return df
    out = df.copy()
    try:
        wd = pd.to_datetime(out["work_date"], errors="coerce")
        if start is not None:
            out = out.loc[wd >= pd.Timestamp(start)]
            wd = pd.to_datetime(out["work_date"], errors="coerce")
        if end is not None:
            out = out.loc[wd <= pd.Timestamp(end)]
    except Exception:
        pass
    if staff and staff != "전체":
        out = out.loc[out["staff"].astype(str) == staff]
    if client_q:
        q = client_q.strip().lower()
        mask = (
            out["client"].astype(str).str.lower().str.contains(q, na=False)
            | out["title"].astype(str).str.lower().str.contains(q, na=False)
            | out["content"].astype(str).str.lower().str.contains(q, na=False)
        )
        out = out.loc[mask]
    if status and status != "전체":
        out = out.loc[out["status"].astype(str) == status]
    if out.empty:
        return _empty_df()
    return out.sort_values(by=["work_date", "updated_at"], ascending=[False, False]).reset_index(
        drop=True
    )


def _render_entry_cards(df: pd.DataFrame) -> None:
    if df.empty:
        st.info("조건에 맞는 업무일지가 없습니다.")
        return
    for _, row in df.head(80).iterrows():
        meta = (
            f"{row.get('work_date', '')} · {row.get('staff', '')} · "
            f"{row.get('category', '')} · {row.get('status', '')}"
        )
        client = str(row.get("client") or "").strip()
        title = str(row.get("title") or "")
        body = str(row.get("content") or "")
        head = f"{title}" + (f" — {client}" if client else "")
        st.markdown(
            f"<div class='worklog-card'>"
            f"<div class='meta'>{_esc(meta)}</div>"
            f"<div class='title'>{_esc(head)}</div>"
            f"<div class='body'>{_esc(body)}</div>"
            f"</div>",
            unsafe_allow_html=True,
        )


def _esc(s: Any) -> str:
    import html as _html

    return _html.escape("" if s is None else str(s)).replace("\n", "<br>")


def render_worklog_tab(latest_update_str: str | None = None) -> None:
    """app.py tab10 진입점. 예외는 탭 안에서만 표시하고 앱 전체를 죽이지 않음."""
    touch = _is_touch_ui()
    try:
        _ensure_cache_dir()
        if touch:
            _inject_worklog_touch_css()
            st.markdown("<div class='worklog-touch-scope'>", unsafe_allow_html=True)

        head_l, head_r = st.columns([4, 1])
        with head_l:
            st.markdown(
                "<div class='sub-header dashboard-tab-panel-head'>📝 일일업무일지</div>",
                unsafe_allow_html=True,
            )
        with head_r:
            badge = latest_update_str or datetime.now().strftime("%Y-%m-%d")
            st.markdown(
                f"<div style='text-align:right;margin-top:0;'>"
                f"<span style='background:#FFFFFF;color:#475569;padding:6px 12px;"
                f"border-radius:8px;font-size:13px;font-weight:600;"
                f"border:1px solid #CBD5E1;'>⏱️ {badge}</span></div>",
                unsafe_allow_html=True,
            )

        df = load_worklog()

        st.markdown("#### 새 일지 작성")
        today = date.today()
        if touch:
            d_col = st.container()
            with d_col:
                work_date = st.date_input("업무일", value=today, key="wl_work_date")
                staff = st.text_input("담당자", value="", key="wl_staff", placeholder="이름")
                client = st.text_input("거래처", value="", key="wl_client", placeholder="거래처명")
                category = st.selectbox("구분", CATEGORIES, index=0, key="wl_category")
                status = st.selectbox("상태", STATUSES, index=0, key="wl_status")
                title = st.text_input("제목", value="", key="wl_title", placeholder="요약 제목")
                content = st.text_area(
                    "내용",
                    value="",
                    key="wl_content",
                    height=160,
                    placeholder="방문 내용, 요청사항, 후속 조치 등",
                )
        else:
            r1 = st.columns([1, 1, 1, 1, 1])
            with r1[0]:
                work_date = st.date_input("업무일", value=today, key="wl_work_date")
            with r1[1]:
                staff = st.text_input("담당자", value="", key="wl_staff", placeholder="이름")
            with r1[2]:
                client = st.text_input("거래처", value="", key="wl_client", placeholder="거래처명")
            with r1[3]:
                category = st.selectbox("구분", CATEGORIES, index=0, key="wl_category")
            with r1[4]:
                status = st.selectbox("상태", STATUSES, index=0, key="wl_status")
            title = st.text_input("제목", value="", key="wl_title", placeholder="요약 제목")
            content = st.text_area(
                "내용",
                value="",
                key="wl_content",
                height=120,
                placeholder="방문 내용, 요청사항, 후속 조치 등",
            )

        if st.button("💾 일지 저장", key="wl_save_btn", type="primary", use_container_width=True):
            entry = _new_entry(
                work_date=work_date if isinstance(work_date, date) else today,
                staff=staff,
                client=client,
                category=category,
                title=title,
                content=content,
                status=status,
            )
            next_df = pd.concat([df, pd.DataFrame([entry])], ignore_index=True)
            if save_worklog(next_df):
                st.session_state["wl_flash_ok"] = "저장되었습니다."
                st.rerun()
            else:
                st.error("저장에 실패했습니다. 디스크 권한·용량을 확인해 주세요.")

        if st.session_state.pop("wl_flash_ok", None):
            st.success("저장되었습니다.")

        st.markdown("---")
        st.markdown("#### 조회 · 관리")

        staff_opts = ["전체"] + sorted(
            {s for s in df["staff"].astype(str).tolist() if s and s != "nan"}
        )
        if touch:
            f_start = st.date_input("시작일", value=today.replace(day=1), key="wl_f_start")
            f_end = st.date_input("종료일", value=today, key="wl_f_end")
            f_staff = st.selectbox("담당자 필터", staff_opts, index=0, key="wl_f_staff")
            f_status = st.selectbox("상태 필터", ["전체"] + STATUSES, index=0, key="wl_f_status")
            f_q = st.text_input("검색(거래처·제목·내용)", value="", key="wl_f_q")
        else:
            f1, f2, f3, f4 = st.columns([1, 1, 1, 2])
            with f1:
                f_start = st.date_input("시작일", value=today.replace(day=1), key="wl_f_start")
            with f2:
                f_end = st.date_input("종료일", value=today, key="wl_f_end")
            with f3:
                f_staff = st.selectbox("담당자 필터", staff_opts, index=0, key="wl_f_staff")
            with f4:
                f_status = st.selectbox("상태 필터", ["전체"] + STATUSES, index=0, key="wl_f_status")
            f_q = st.text_input("검색(거래처·제목·내용)", value="", key="wl_f_q")

        view = _filter_df(
            df,
            start=f_start if isinstance(f_start, date) else None,
            end=f_end if isinstance(f_end, date) else None,
            staff=f_staff,
            client_q=f_q,
            status=f_status,
        )

        st.caption(f"표시 {len(view)}건 / 전체 {len(df)}건")

        if touch:
            _render_entry_cards(view)
        else:
            show = view.rename(
                columns={
                    "work_date": "업무일",
                    "staff": "담당자",
                    "client": "거래처",
                    "category": "구분",
                    "title": "제목",
                    "content": "내용",
                    "status": "상태",
                    "updated_at": "수정시각",
                }
            )
            cols = [c for c in ["업무일", "담당자", "거래처", "구분", "제목", "내용", "상태", "수정시각"] if c in show.columns]
            st.dataframe(show[cols], use_container_width=True, hide_index=True, height=min(480, 48 + 36 * max(1, len(show))))

        # 삭제 (선택)
        if not view.empty:
            labels = {
                str(r["id"]): f"{r['work_date']} | {r['staff']} | {r['title']}"
                for _, r in view.iterrows()
            }
            del_id = st.selectbox(
                "삭제할 일지",
                options=["(선택 없음)"] + list(labels.keys()),
                format_func=lambda x: labels.get(x, x),
                key="wl_del_id",
            )
            if st.button("🗑️ 선택 일지 삭제", key="wl_del_btn", use_container_width=True):
                if del_id and del_id != "(선택 없음)":
                    kept = df.loc[df["id"].astype(str) != str(del_id)].copy()
                    if save_worklog(kept):
                        st.session_state["wl_flash_ok"] = "삭제되었습니다."
                        st.rerun()
                    else:
                        st.error("삭제 저장에 실패했습니다.")

        # 내보내기
        if not view.empty:
            csv_bytes = view.to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                "⬇️ 조회결과 CSV",
                data=csv_bytes,
                file_name=f"일일업무일지_{today.strftime('%Y%m%d')}.csv",
                mime="text/csv",
                use_container_width=True,
                key="wl_csv_dl",
            )

        if touch:
            st.markdown("</div>", unsafe_allow_html=True)

    except Exception as exc:
        st.error(f"일일업무일지 탭 오류: {exc}")
        st.caption("다른 탭은 그대로 사용 가능합니다. 페이지를 새로고침해 주세요.")
