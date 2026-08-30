"""업무일지 + 공문 전용 Streamlit 대시보드.

메인 영업 대시보드(app.py)와 별도 포트·브라우저에서 실행합니다.
모듈·uploaded_cache 는 상위(dashboard) 폴더를 공유합니다.
"""
from __future__ import annotations

import os
import sys

# 상위 dashboard 루트를 cwd·import 경로로 (uploaded_cache, worklog_tab 등 공유)
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)
if _THIS := os.path.dirname(os.path.abspath(__file__)):
    if _THIS not in sys.path:
        sys.path.insert(0, _THIS)
os.chdir(_ROOT)

import streamlit as st

from dev_mode import apply_dev_ui_gate, dev_caption
from sales_loader import load_sales_for_letter_tab

try:
    import worklog_tab as _worklog_tab_mod

    _WL_BUILD = str(getattr(_worklog_tab_mod, "_WL_UI_BUILD", "") or "").strip()
except Exception:
    _WL_BUILD = ""
_APP_BUILD = f"2026-08-30 · 업무일지공문2탭" + (f" · {_WL_BUILD}" if _WL_BUILD else "")

apply_dev_ui_gate()

st.set_page_config(
    page_title="업무일지·공문",
    layout="wide",
    page_icon="📝",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
    .block-container { padding-top: 0.75rem; }
    div[data-testid="stTabs"] {
        position: sticky;
        top: 0;
        z-index: 200;
        background: #fff;
        padding-top: 0.25rem;
        padding-bottom: 0.35rem;
        border-bottom: 1px solid #E2E8F0;
        margin-bottom: 0.5rem;
    }
    div[data-testid="stTabs"] button[data-baseweb="tab"] {
        font-size: 1rem !important;
        font-weight: 600 !important;
        padding: 0.45rem 1.1rem !important;
    }
    #wl-dash-head {
        margin: 0 0 0.35rem 0;
        font-size: 1.05rem;
        font-weight: 700;
        color: #0F172A;
    }
    #wl-dash-sub {
        margin: 0 0 0.5rem 0;
        font-size: 0.82rem;
        color: #64748B;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    f"<p id='wl-dash-head'>업무일지 · 공문 전용 대시보드</p>"
    f"<p id='wl-dash-sub'>아래 탭에서 <b>📝 일일업무일지</b> 와 <b>📨 공문</b> 을 선택 · 빌드 {_APP_BUILD}</p>",
    unsafe_allow_html=True,
)
dev_caption(f"업무일지·공문 전용 · 포트 8502 · {_APP_BUILD}")

tab_wl, tab_letter = st.tabs(["📝 일일업무일지", "📨 공문"])

with tab_wl:
    try:
        import worklog_tab as _worklog_tab

        _worklog_tab.render_worklog_tab()
    except ModuleNotFoundError:
        st.error("worklog_tab.py 를 찾을 수 없습니다. dashboard 루트에 있는지 확인하세요.")
    except Exception as _wl_err:
        st.error(f"일일업무일지 탭 오류: {_wl_err}")

with tab_letter:
    try:
        import price_increase_tab as _pi_tab

        _sales_df, _latest = load_sales_for_letter_tab()
        if _sales_df.empty:
            st.info(
                "매출 CSV가 없습니다. 메인 대시보드와 같이 "
                "`uploaded_cache/sales/` 에 2026.csv 등을 두거나 "
                "Drive 동기화 후 새로고침하세요."
            )
        _pi_tab.render_price_increase_tab(_sales_df, latest_update_str=_latest)
    except ModuleNotFoundError:
        st.error("price_increase_tab.py 를 찾을 수 없습니다.")
    except Exception as _pi_err:
        st.error(f"공문 탭 오류: {_pi_err}")
