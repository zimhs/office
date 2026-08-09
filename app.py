import io
import os
import re
import sys
import html
import json
import subprocess
import shutil
import urllib.parse
import requests
from bs4 import BeautifulSoup
import numpy as np
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
import plotly.graph_objects as go
import plotly.express as px
import datetime
from matplotlib.colors import LinearSegmentedColormap
# Tab3 히트맵 — 파스텔만 사용 (진한 네이비/브라운 제외)
_TAB3_CMAP_BLUE = LinearSegmentedColormap.from_list(
    "tab3_blue", ["#F8FAFC", "#DBEAFE", "#93C5FD"]
)
_TAB3_CMAP_GREEN = LinearSegmentedColormap.from_list(
    "tab3_green", ["#F8FAFC", "#DCFCE7", "#86EFAC"]
)
_TAB3_CMAP_ORANGE = LinearSegmentedColormap.from_list(
    "tab3_orange", ["#FFF7ED", "#FFEDD5", "#FDBA74"]
)
# OpenDartReader 임포트 (설치되지 않았을 경우를 대비한 예외 처리)
try:
    import OpenDartReader
except ImportError:
    OpenDartReader = None
# 페이지 및 Styler 가동 한도 설정 (과부하 방지용)
pd.set_option("styler.render.max_elements", 2000000)
st.set_page_config(page_title="통합 영업 분석 대시보드", layout="wide", initial_sidebar_state="expanded")
# ==========================================
# 0. 로컬 파일 자동 저장을 위한 디렉토리 설정
# ==========================================
CACHE_DIR = "./uploaded_cache"
os.makedirs(CACHE_DIR, exist_ok=True)
# ==========================================
# 1. 상단 공백 최소화 및 사이드바 무손실 복구 CSS
# ==========================================
def inject_custom_css():
    st.markdown(
        """
        <script>
            document.documentElement.lang = 'ko';
            document.documentElement.classList.add('notranslate');
        </script>
        <meta name="google" content="notranslate" />
        <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
        <style>
            /* 기본 불필요 UI 숨김 */
            div[data-baseweb="select"] + div:has(span) { display: none !important; }
            div[data-testid="stMultiSelect"] [data-testid="stWidgetInstructions"] { display: none !important; }
            small[data-testid="stCaptionContainer"] { display: none !important; }
            
            html, body, .stApp {
                background-color: #F8FAFC !important;
                color: #1E293B !important;
                font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
                -webkit-tap-highlight-color: transparent;
            }
            
            /* ★★★ 상단 여백 최소화 (Streamlit 헤더는 그대로) ★★★ */
            section.main .block-container { 
                padding-top: 0.25rem !important; 
                padding-bottom: 0.75rem !important;
                padding-left: 1rem !important;
                padding-right: 1rem !important;
                max-width: 99% !important; 
            }
            section.main [data-testid="stVerticalBlock"]:first-child,
            section.main [data-testid="stVerticalBlockBorderWrapper"]:first-child {
                margin-top: 0 !important;
                padding-top: 0 !important;
            }
            /* 사이드바 — Streamlit 기본 동작 유지, 색상만 보조 */
            [data-testid="stSidebar"] {
                background-color: #F1F5F9 !important;
                border-right: 1px solid #E2E8F0;
            }
            [data-testid="stSidebar"] .block-container { 
                padding-top: 1.5rem !important; 
            }
            [data-testid="stSidebar"] p, [data-testid="stSidebar"] span, 
            [data-testid="stSidebar"] label, [data-testid="stSidebar"] .stMarkdown {
                color: #334155 !important;
            }
            div[data-testid="column"] { align-self: flex-start; }
            /* 채권관리 월/연체 토글 — 해당 컨테이너에만 적용 (다른 탭 버튼에 영향 금지) */
            [class*="st-key-debt_compact_btns"] div[data-testid="stButton"] > button {
                min-height: 26px !important;
                height: 26px !important;
                padding: 0 6px !important;
                font-size: 11px !important;
                font-weight: 600 !important;
                line-height: 1.1 !important;
                border-radius: 6px !important;
                max-width: 88px !important;
                margin: 0 auto !important;
            }
            [class*="st-key-debt_compact_btns"] div[data-testid="column"] {
                padding-left: 1px !important;
                padding-right: 1px !important;
            }
            /* Tab2 액션 버튼 — 가로 넓게, 글씨 한 줄로 박스 안에 */
            .st-key-tab2_action_btns div[data-testid="stButton"],
            .st-key-tab2_action_btns div[data-testid="stLinkButton"] {
                width: 100% !important;
            }
            .st-key-tab2_action_btns div[data-testid="stButton"] > button,
            .st-key-tab2_action_btns a[data-testid="stLinkButton"] {
                width: 100% !important;
                max-width: none !important;
                min-height: 48px !important;
                height: 48px !important;
                padding: 0 16px !important;
                font-size: 13px !important;
                font-weight: 600 !important;
                line-height: 1.2 !important;
                border-radius: 8px !important;
                text-align: center !important;
                white-space: nowrap !important;
                overflow: hidden !important;
                text-overflow: ellipsis !important;
                box-sizing: border-box !important;
            }
            .metric-box {
                background: #FFFFFF;
                padding: 16px 20px;
                border-radius: 10px;
                border: 1px solid #E2E8F0;
                box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
                margin-bottom: 8px;
            }
            .metric-label { color: #64748B; font-size: 13px; font-weight: 500; margin-bottom: 6px; }
            .metric-value { color: #0F172A; font-size: 20px; font-weight: 700; }
            
            .sub-header {
                color: #1E3A8A;
                font-size: 17px;
                font-weight: 700;
                margin-top: 0 !important;
                margin-bottom: 12px;
                border-left: 4px solid #1E3A8A;
                padding-left: 10px;
            }
            .dashboard-tabs-host-compact [role="tabpanel"] {
                padding-top: 8px !important;
            }
            .dashboard-tab-title-row {
                min-height: 44px !important;
                align-items: flex-start !important;
            }
            
            div[role="radiogroup"] {
                padding: 10px;
                background: white;
                border-radius: 8px;
                border: 1px solid #E2E8F0;
            }
            /* ===== 상단 통합 고정바 (필터 + 탭) ===== */
            #dashboard-sticky-spacer {
                width: 100% !important;
                height: var(--dashboard-fixed-bar-height, 108px) !important;
                margin: 0 !important;
                padding: 0 !important;
                pointer-events: none !important;
                flex-shrink: 0 !important;
            }
            #dashboard-top-shield {
                pointer-events: none !important;
            }
            section.main {
                overflow: visible !important;
            }
            .dashboard-tabs-host-compact {
                margin-top: 0 !important;
                padding-top: 0 !important;
                height: auto !important;
                max-height: none !important;
                overflow: visible !important;
            }
            /* tablist만 숨김 — tabpanel 직접 자식은 반드시 유지 */
            .dashboard-tabs-host-compact > div:has(> [role="tablist"]),
            .dashboard-tabs-host-compact > div.dashboard-tabs-list-shell:empty {
                display: none !important;
                height: 0 !important;
                min-height: 0 !important;
                margin: 0 !important;
                padding: 0 !important;
                overflow: hidden !important;
            }
            .dashboard-tabs-host-compact > div[role="tabpanel"],
            .dashboard-tabs-host-compact [role="tabpanel"] {
                display: block !important;
                height: auto !important;
                max-height: none !important;
                min-height: 0 !important;
                overflow: visible !important;
            }
            .dashboard-filter-sticky {
                position: fixed !important;
                top: var(--dashboard-bar-top, 2.75rem) !important;
                left: var(--dashboard-bar-left, 0px) !important;
                width: var(--dashboard-bar-width, 100%) !important;
                max-width: var(--dashboard-bar-width, 100%) !important;
                z-index: 990 !important;
                background-color: #FFFFFF !important;
                border: 1px solid #BFDBFE !important;
                border-radius: 0 0 6px 6px !important;
                padding: 1px 6px 0 6px !important;
                box-shadow: 0 2px 8px -2px rgba(15, 23, 42, 0.06) !important;
                box-sizing: border-box !important;
                overflow-x: hidden !important;
                overflow-y: visible !important;
                margin: 0 !important;
            }
            .dashboard-filter-sticky-touch {
                z-index: 999999 !important;
                top: 3.65rem !important;
                border: 2px solid #2563EB !important;
                border-top: 3px solid #2563EB !important;
                border-radius: 0 0 8px 8px !important;
                padding: 8px 10px 0 10px !important;
                box-shadow: 0 8px 14px -4px rgba(37, 99, 235, 0.18) !important;
                margin-top: 0 !important;
            }
            html.dashboard-touch-mode [data-testid="stHeader"],
            html.dashboard-touch-mode [data-testid="stToolbar"],
            html.dashboard-touch-mode [data-testid="stDecoration"] {
                z-index: 1000001 !important;
                position: relative !important;
            }
            html.dashboard-touch-mode [data-testid="stSidebar"],
            html.dashboard-touch-mode [data-testid="stSidebarBackdrop"],
            html.dashboard-touch-mode section[data-testid="stSidebar"] {
                z-index: 1000005 !important;
            }
            html.dashboard-touch-mode [data-testid="stAppViewContainer"],
            html.dashboard-touch-mode section.main,
            html.dashboard-touch-mode section.main .block-container {
                overflow: visible !important;
            }
            /* 고정바 내부 공백 최소화 */
            .dashboard-filter-sticky [data-testid="stVerticalBlock"],
            .dashboard-filter-sticky [data-testid="stHorizontalBlock"] {
                gap: 0.1rem !important;
            }
            .dashboard-filter-sticky [data-testid="column"] {
                padding-top: 0 !important;
                padding-bottom: 0 !important;
            }
            .dashboard-filter-sticky [data-testid="stWidgetLabel"],
            .dashboard-filter-sticky label {
                margin-bottom: 0 !important;
                padding-bottom: 0 !important;
                font-size: 11px !important;
                line-height: 1.15 !important;
                min-height: auto !important;
            }
            .dashboard-filter-sticky [data-testid="stTextInput"],
            .dashboard-filter-sticky [data-testid="stSelectbox"],
            .dashboard-filter-sticky [data-testid="stMultiSelect"] {
                margin-bottom: 0 !important;
                padding-bottom: 0 !important;
            }
            .dashboard-filter-sticky [data-testid="stTextInput"] > div > div,
            .dashboard-filter-sticky [data-baseweb="select"] > div {
                min-height: 1.85rem !important;
            }
            .dashboard-filter-sticky [data-testid="stMarkdownContainer"] {
                margin: 0 !important;
                padding: 0 !important;
            }
            div[data-testid="stVerticalBlockBorderWrapper"].dashboard-filter-sticky {
                padding-top: 2px !important;
                padding-bottom: 0 !important;
            }
            /* 고정바 하단 구분선 — 가늘고 연하게 */
            .dashboard-filter-sticky::after {
                content: '' !important;
                display: block !important;
                height: 1px !important;
                margin: 2px -6px 0 -6px !important;
                background: #CBD5E1 !important;
                box-shadow: none !important;
                border-radius: 0 !important;
            }
            /* 탭 패널 높이 — 대시보드 메인 탭만 */
            .dashboard-tabs-host-compact,
            .dashboard-tabs-host-compact > div,
            .dashboard-tabs-host-compact [role="tabpanel"] {
                overflow: visible !important;
                max-height: none !important;
                height: auto !important;
            }
            div[data-testid="stDataFrame"] > div {
                -webkit-overflow-scrolling: touch;
            }
            /* 탭: 필터 고정바 내부 하단, 세로모드 가로 스크롤 */
            .dashboard-filter-sticky [role="tablist"],
            .dashboard-tabs-in-filter {
                display: flex !important;
                flex-wrap: nowrap !important;
                overflow-x: auto !important;
                overflow-y: visible !important;
                -webkit-overflow-scrolling: touch !important;
                width: 100% !important;
                max-width: 100% !important;
                margin-top: 1px !important;
                margin-bottom: 0 !important;
                padding: 0 0 2px 0 !important;
                border-top: 1px solid #E2E8F0 !important;
                background-color: #FFFFFF !important;
                box-sizing: border-box !important;
                scrollbar-width: none;
                touch-action: pan-x !important;
            }
            .dashboard-filter-sticky [role="tab"] {
                flex: 0 0 auto !important;
                white-space: nowrap !important;
                min-width: -webkit-max-content !important;
                min-width: max-content !important;
                padding: 2px 8px 3px 8px !important;
                font-size: 12px !important;
            }
            .dashboard-filter-sticky [role="tablist"]::-webkit-scrollbar,
            .dashboard-tabs-in-filter::-webkit-scrollbar {
                display: none;
            }
            .dashboard-tabs-host-compact [role="tabpanel"] {
                padding-top: 4px !important;
            }
            .dashboard-tab-panel-head {
                padding-top: 2px !important;
                margin-bottom: 6px !important;
            }
            .dashboard-tabs-host-compact [role="tabpanel"]:not([hidden]) {
                padding-bottom: 48px !important;
            }
            /* Streamlit 기본 스크롤 복구 */
            [data-testid="stAppViewContainer"] {
                overflow-y: auto !important;
                overflow-x: hidden !important;
                -webkit-overflow-scrolling: touch !important;
            }
            .dashboard-filter-sticky [role="tab"] p,
            .dashboard-filter-sticky [role="tab"] span,
            .dashboard-filter-sticky [role="tab"] label {
                white-space: nowrap !important;
            }
            @media (max-width: 1024px) {
                .dashboard-filter-sticky {
                    padding: 2px 6px 0 6px !important;
                }
                .dashboard-filter-sticky [role="tab"] {
                    font-size: 10px !important;
                    padding: 2px 6px 3px 6px !important;
                }
            }
            @supports (top: env(safe-area-inset-top)) {
                .dashboard-filter-sticky {
                    padding-left: max(10px, env(safe-area-inset-left)) !important;
                    padding-right: max(10px, env(safe-area-inset-right)) !important;
                }
            }
            /* iPad only: Top30 표·도넛 가로 100% — 직계 칼럼만 (중첩 월버튼 칼럼 제외) */
            html.dashboard-touch-mode .top30-touch-scope,
            html.dashboard-touch-mode .top30-touch-row {
                width: 100% !important;
                max-width: 100% !important;
                box-sizing: border-box !important;
            }
            html.dashboard-touch-mode .top30-touch-row {
                display: flex !important;
                flex-direction: column !important;
                flex-wrap: nowrap !important;
                align-items: stretch !important;
                gap: 0.75rem !important;
            }
            html.dashboard-touch-mode .top30-touch-row > [data-testid="column"],
            html.dashboard-touch-mode .top30-touch-row > [data-testid="stColumn"],
            html.dashboard-touch-mode .top30-touch-row > div > [data-testid="column"],
            html.dashboard-touch-mode .top30-touch-row > div > [data-testid="stColumn"] {
                width: 100% !important;
                min-width: 100% !important;
                max-width: 100% !important;
                flex: 1 1 100% !important;
                box-sizing: border-box !important;
            }
            html.dashboard-touch-mode .top30-touch-row iframe,
            html.dashboard-touch-mode .top30-touch-row [data-testid="stPlotlyChart"],
            html.dashboard-touch-mode .top30-touch-row [data-testid="stPlotlyChart"] > div {
                width: 100% !important;
                max-width: 100% !important;
                min-width: 0 !important;
            }
            html.dashboard-touch-mode .top30-touch-row [data-testid="stPlotlyChart"],
            html.dashboard-touch-mode .top30-touch-row [data-testid="stPlotlyChart"] iframe {
                min-height: 420px !important;
                height: 420px !important;
            }
        </style>
        """,
        unsafe_allow_html=True,
    )
# ==========================================
# 2. 날짜 파싱 및 데이터 정규화 유틸리티
# ==========================================
def parse_date_series_robust(series, default_year="2026"):
    if series.empty:
        return pd.Series(pd.NaT, index=series.index)
    s_str = series.astype(str).str.strip()
    parsed = pd.Series(pd.NaT, index=series.index)
    digits = s_str.str.replace(r"\D", "", regex=True)
    cond_8 = (digits.str.len() == 8) & parsed.isna()
    if cond_8.any():
        parsed[cond_8] = pd.to_datetime(
            digits[cond_8], format="%Y%m%d", errors="coerce"
        )
    cond_6 = (digits.str.len() == 6) & parsed.isna()
    if cond_6.any():
        parsed[cond_6] = pd.to_datetime(
            "20" + digits[cond_6], format="%Y%m%d", errors="coerce"
        )
    remaining_mask = (
        parsed.isna() & (s_str != "") & (s_str != "nan") & (s_str != "None")
    )
    if remaining_mask.any():
        rem_series = s_str[remaining_mask]
        parts_df = rem_series.str.split(r"[-/.\s]+", expand=True)
        parts_df = parts_df.apply(lambda col: col.str.strip()).replace("", None)
        num_parts = parts_df.notna().sum(axis=1)
        cond_3 = num_parts >= 3
        if cond_3.any():
            sub = parts_df[cond_3]
            y = sub[0].astype(str).str.strip()
            y = np.where(y.str.len() == 2, "20" + y, y)
            m = sub[1].astype(str).str.strip().str.zfill(2)
            d = sub[2].astype(str).str.strip().str.zfill(2)
            dt_str = y + "-" + m + "-" + d
            parsed.loc[sub.index] = pd.to_datetime(
                dt_str, format="%Y-%m-%d", errors="coerce"
            )
        cond_2 = num_parts == 2
        if cond_2.any():
            sub = parts_df[cond_2]
            m = sub[0].astype(str).str.strip().str.zfill(2)
            d = sub[1].astype(str).str.strip().str.zfill(2)
            dt_str = str(default_year) + "-" + m + "-" + d
            parsed.loc[sub.index] = pd.to_datetime(
                dt_str, format="%Y-%m-%d", errors="coerce"
            )
    valid_range = (parsed >= pd.Timestamp("2000-01-01")) & (
        parsed <= pd.Timestamp("2099-12-31")
    )
    parsed[~valid_range] = pd.NaT
    return parsed
def normalize_items_vectorized(df):
    if "품목명" not in df.columns or df.empty:
        return df
    p_str = df["품목명"].astype(str)
    p_upper = p_str.str.upper().str.replace(" ", "")
    is_bulk = p_upper.str.contains("BULK", na=False) | p_str.str.contains("벌크", na=False)
    
    is_ar = is_bulk & (p_upper.str.contains("AR", na=False) | p_str.str.contains("아르곤|아르", na=False))
    is_co2 = is_bulk & (p_upper.str.contains("CO2", na=False) | p_str.str.contains("탄산", na=False))
    is_o2 = is_bulk & ~is_co2 & (p_upper.str.contains("O2", na=False) | p_str.str.contains("산소", na=False))
    is_n2 = is_bulk & (p_upper.str.contains("N2", na=False) | p_str.str.contains("질소", na=False))
    is_n2_liter = is_n2 & (p_upper.str.contains(r"\bL\b|LITER", regex=True, na=False) | p_str.str.contains("리터", na=False))
    df.loc[is_ar, "품목명"] = "AR (kg, Bulk)"
    df.loc[is_co2, "품목명"] = "CO2 (kg, Bulk)"
    df.loc[is_o2, "품목명"] = "O2 (kg, Bulk)"
    df.loc[is_n2, "품목명"] = "N2 (kg, Bulk)"
    return df
def get_exact_original_price(series):
    s = series[series > 0]
    if s.empty:
        return 0
    return s.mode().iloc[0]
def apply_forward_unit_price(unit_price_df, qty_df, years, all_months):
    """당월 출고 없음(0)이어도 직전 최종 변경 단가를 이월 표시."""
    if unit_price_df.empty:
        return unit_price_df
    filled = unit_price_df.copy()
    chron_cols = [
        f"{yr[2:]}년 {m}"
        for yr in sorted(years)
        for m in all_months
        if f"{yr[2:]}년 {m}" in filled.columns
    ]
    for idx in filled.index:
        last_price = 0.0
        for col in chron_cols:
            price = float(filled.at[idx, col]) if pd.notna(filled.at[idx, col]) else 0.0
            qty = 0.0
            if idx in qty_df.index and col in qty_df.columns:
                qty = float(qty_df.at[idx, col]) if pd.notna(qty_df.at[idx, col]) else 0.0
            if qty > 0 and price > 0:
                last_price = price
            elif last_price > 0 and (qty == 0 or price == 0):
                filled.at[idx, col] = last_price
    return filled
@st.cache_data
def convert_dfs_to_excel(dfs_dict):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        sheets_written = False
        for sheet_name, (df, use_index) in dfs_dict.items():
            if not df.empty:
                if isinstance(df.columns, pd.MultiIndex):
                    df_to_save = df.copy()
                    df_to_save.columns = [
                        "_".join([str(c) for c in col if c])
                        for col in df_to_save.columns
                    ]
                    df_to_save.to_excel(
                        writer, sheet_name=sheet_name, index=use_index
                    )
                else:
                    df.to_excel(writer, sheet_name=sheet_name, index=use_index)
                sheets_written = True
        
        # [핵심] 저장할 데이터가 전혀 없어 시트가 0개일 때 발생하는 크래시 방지
        if not sheets_written:
            pd.DataFrame({"알림": ["저장할 데이터가 없습니다."]}).to_excel(writer, sheet_name="No Data", index=False)
            
    return output.getvalue()
def _ppt_format_value(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    if isinstance(val, (pd.Timestamp, datetime.datetime, datetime.date)):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, (int, np.integer)):
        return f"{val:,}"
    if isinstance(val, (float, np.floating)):
        if abs(val) >= 100:
            return f"{val:,.0f}"
        return f"{val:,.2f}"
    text = str(val).replace("\n", " ")
    return text[:40] + ("…" if len(text) > 40 else "")
def _ppt_add_title_bar(slide, title, subtitle=""):
    from pptx.util import Inches, Pt
    box = slide.shapes.add_textbox(Inches(0.4), Inches(0.25), Inches(9.2), Inches(0.7))
    tf = box.text_frame
    tf.text = title
    p = tf.paragraphs[0]
    p.font.size = Pt(24)
    p.font.bold = True
    if subtitle:
        sub = slide.shapes.add_textbox(Inches(0.4), Inches(0.85), Inches(9.2), Inches(0.45))
        stf = sub.text_frame
        stf.text = subtitle
        stf.paragraphs[0].font.size = Pt(11)
def _ppt_add_bullets(slide, lines, top=1.35):
    from pptx.util import Inches, Pt
    box = slide.shapes.add_textbox(Inches(0.55), Inches(top), Inches(9.0), Inches(5.5))
    tf = box.text_frame
    tf.word_wrap = True
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = line
        p.font.size = Pt(14)
        p.level = 0
def _ppt_add_dataframe(slide, df, max_rows=14, max_cols=8, top=1.25):
    from pptx.util import Inches
    if df is None or df.empty:
        _ppt_add_bullets(slide, ["표시할 데이터가 없습니다."], top=top)
        return
    view = df.copy()
    if isinstance(view.columns, pd.MultiIndex):
        view.columns = ["_".join(str(c) for c in col if c) for col in view.columns]
    view = view.iloc[:max_rows, :max_cols]
    rows = len(view) + 1
    cols = len(view.columns) + (1 if view.index.name or not isinstance(view.index, pd.RangeIndex) else 0)
    has_index = not isinstance(view.index, pd.RangeIndex) or view.index.name
    if not has_index:
        cols = len(view.columns)
    if cols == 0:
        _ppt_add_bullets(slide, ["표시할 데이터가 없습니다."], top=top)
        return
    left, width, height = Inches(0.35), Inches(9.3), Inches(0.28)
    table_shape = slide.shapes.add_table(rows, cols, left, Inches(top), width, height * rows)
    table = table_shape.table
    col_offset = 0
    if has_index:
        table.cell(0, 0).text = str(view.index.name or "")
        col_offset = 1
    for c, col_name in enumerate(view.columns):
        table.cell(0, c + col_offset).text = _ppt_format_value(col_name)
    for r in range(len(view)):
        row = view.iloc[r]
        if has_index:
            table.cell(r + 1, 0).text = _ppt_format_value(view.index[r])
        for c, col_name in enumerate(view.columns):
            table.cell(r + 1, c + col_offset).text = _ppt_format_value(row[col_name])
def _ppt_try_add_chart(slide, fig, top=1.2):
    from pptx.util import Inches
    if fig is None:
        return False
    try:
        img_bytes = fig.to_image(format="png", width=960, height=540, scale=1)
        slide.shapes.add_picture(io.BytesIO(img_bytes), Inches(0.45), Inches(top), width=Inches(9.1))
        return True
    except Exception:
        return False
def _ppt_new_content_slide(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])
def _ppt_tab_slide(prs, tab_title, section_title, df, chart_fig=None, bullets=None):
    slide = _ppt_new_content_slide(prs)
    _ppt_add_title_bar(slide, tab_title, section_title)
    row_top = 1.25
    if bullets:
        _ppt_add_bullets(slide, bullets, top=row_top)
        row_top = 2.15
    if chart_fig is not None and _ppt_try_add_chart(slide, chart_fig, top=row_top):
        return
    _ppt_add_dataframe(slide, df, top=row_top)
def _ppt_cover_subtitle(latest_update_str, selected_client, selected_staff_tuple):
    today_str = datetime.date.today().strftime("%Y-%m-%d")
    staff_label = ", ".join(selected_staff_tuple) if selected_staff_tuple else "전체"
    return (
        f"생성일: {today_str}\n"
        f"데이터 기준: {latest_update_str}\n"
        f"조회 거래처: {selected_client}\n"
        f"담당자: {staff_label}"
    )
def _ppt_save_presentation(prs):
    out = io.BytesIO()
    prs.save(out)
    return out.getvalue()
@st.cache_data(show_spinner="PPT 파일 생성 중...")
def convert_dashboard_to_ppt(
    latest_update_str,
    selected_client,
    selected_staff_tuple,
    latest_month_str_total,
    tot_sales_val,
    cur_sales_val,
    mom_rate_total,
    avg_rate_total,
    pivot_m_total,
    pivot_m_client,
    client_item_qty_pivot,
    sales_p,
    qty_p,
    staff_pivot,
    df_detail,
    filtered_debt_df,
    df_base,
    df_tank,
    df_vaporizer,
    df_integrated,
    all_months_tuple,
    years_tuple,
    target_items_tuple,
):
    try:
        from pptx import Presentation
    except ImportError as exc:
        raise ImportError("python-pptx 패키지가 필요합니다. pip install python-pptx") from exc
    all_months = list(all_months_tuple)
    years = list(years_tuple)
    target_items = list(target_items_tuple)
    prs = Presentation()
    cover = prs.slides.add_slide(prs.slide_layouts[0])
    cover.shapes.title.text = "통합 영업 분석 대시보드"
    cover.placeholders[1].text = _ppt_cover_subtitle(latest_update_str, selected_client, selected_staff_tuple)
    # Tab1
    _ppt_tab_slide(
        prs, "📌 Tab1 · 영업 종합 요약", "주요 실적 지표", pivot_m_total,
        bullets=[
            f"총 누적 매출(VAT포함): {tot_sales_val:,.0f} 만원",
            f"최근 월 매출 ({latest_month_str_total}): {cur_sales_val:,.0f} 만원",
            f"전월 대비(MoM): {mom_rate_total:+.0f}%",
            f"월평균 대비 증감: {avg_rate_total:+.0f}%",
        ],
    )
    _ppt_tab_slide(
        prs, "📌 Tab1 · 영업 종합 요약", "연도별 월 매출 추이 (차트)", pivot_m_total,
        chart_fig=create_stacked_bar_chart(pivot_m_total, title_text=""),
    )
    _ppt_tab_slide(prs, "📌 Tab1 · 영업 종합 요약", "연도별 월 매출 (표)", pivot_m_total)
    if not df_base.empty and target_items:
        item_pivot = cached_get_item_pivot(
            df_base, target_items[0], "매출액 (만원)", all_months, years
        )
        _ppt_tab_slide(prs, "📌 Tab1 · 영업 종합 요약", f"주요 품목 분석 — {target_items[0]}", item_pivot)
    # Tab2
    tab2_title = f"🏢 Tab2 · 거래처 분석 [{selected_client}]"
    _ppt_tab_slide(
        prs, tab2_title, "연도별 월 매출 추이 (차트)", pivot_m_client,
        chart_fig=create_stacked_bar_chart(pivot_m_client, title_text="") if not pivot_m_client.empty else None,
    )
    _ppt_tab_slide(prs, tab2_title, "연도별 월 매출 (표)", pivot_m_client)
    _ppt_tab_slide(prs, tab2_title, "거래처별 품목 사용량", client_item_qty_pivot)
    # Tab3
    sales_view = sales_p * 1.1 / 10000 if not sales_p.empty else sales_p
    _ppt_tab_slide(prs, "📦 Tab3 · 품목 및 단가 분석", "품목별 매출액 (만원, VAT포함)", sales_view)
    _ppt_tab_slide(prs, "📦 Tab3 · 품목 및 단가 분석", "품목별 출고량", qty_p)
    # Tab4
    _ppt_tab_slide(prs, "👤 Tab4 · 담당자 & 상세내역", "담당자별 월 매출 (만원)", staff_pivot)
    detail_view = df_detail.sort_values(by="매출일_dt", ascending=False).head(20) if not df_detail.empty else df_detail
    _ppt_tab_slide(prs, "👤 Tab4 · 담당자 & 상세내역", "거래 상세 내역 (최신 20건)", detail_view)
    # Tab5
    _ppt_tab_slide(prs, "📌 Tab5 · 채권 관리", "채권 현황", filtered_debt_df)
    # Tab6
    if not df_base.empty:
        map_summary = (
            df_base.groupby("담당자")["거래처"].nunique().reset_index(name="거래처 수").sort_values("거래처 수", ascending=False)
        )
        addr_count = df_base["거래처"].nunique()
        _ppt_tab_slide(prs, "📍 Tab6 · 지도 분포", f"담당자별 거래처 수 (전체 {addr_count}곳)", map_summary)
    else:
        _ppt_tab_slide(prs, "📍 Tab6 · 지도 분포", "거래처 분포", pd.DataFrame())
    # Tab7
    _ppt_tab_slide(prs, "🏭 Tab7 · 설비 재고", "탱크 재고 현황", df_tank)
    _ppt_tab_slide(prs, "🏭 Tab7 · 설비 재고", "기화기 재고 현황", df_vaporizer)
    # Tab8
    _ppt_tab_slide(prs, "🛢️ Tab8 · 통합 탱크 재고", "통합 탱크 재고 현황", df_integrated)
    return _ppt_save_presentation(prs)
# ==========================================
# ★ 네이버 크롤링 + DART API 하이브리드 유틸리티 ★
# ==========================================
@st.cache_data(show_spinner=False, max_entries=100)
def get_company_info_hybrid(company_name, dart_api_key=None):
    clean_name = re.sub(r'\(.*?\)|\[.*?\]|주식회사|㈜|\(주\)|주\)', '', company_name).strip()
    if not clean_name:
        clean_name = company_name 
        
    info = {
        "ceo": "정보 없음",
        "industry": "정보 없음",
        "revenue": "정보 없음",
        "profit": "정보 없음",
        "clean_name": clean_name,
        "source": "정보 없음",
        "dart_error": "" 
    }
    dart_success = False
    if dart_api_key and OpenDartReader is not None:
        try:
            dart = OpenDartReader(dart_api_key)
            corp_info = dart.company(clean_name)
            
            if corp_info and 'ceo_nm' in corp_info:
                info['ceo'] = corp_info['ceo_nm']
                info['industry'] = corp_info.get('induty_nm', '정보 없음')
                
            fin_state = dart.finstate(clean_name, 2023)
            if fin_state is not None and not fin_state.empty:
                sales_row = fin_state[fin_state['account_nm'].str.contains('매출', na=False)]
                profit_row = fin_state[fin_state['account_nm'].str.contains('영업이익', na=False)]
                
                if not sales_row.empty:
                    s_val = str(sales_row['thstrm_amount'].values[0]).replace(',', '')
                    info['revenue'] = f"{int(s_val):,} 원"
                
                if not profit_row.empty:
                    p_val = str(profit_row['thstrm_amount'].values[0]).replace(',', '')
                    info['profit'] = f"{int(p_val):,} 원"
                    
                info["source"] = "금융감독원 DART (2023)"
                dart_success = True
        except Exception as e:
            info["dart_error"] = str(e)
    if not dart_success or info['revenue'] == "정보 없음":
        try:
            search_query = urllib.parse.quote(clean_name + " 기업정보")
            url = f"https://search.naver.com/search.naver?query={search_query}"
            
            headers = {
                "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
            }
            
            response = requests.get(url, headers=headers, timeout=3)
            
            if response.status_code == 200:
                soup = BeautifulSoup(response.text, 'html.parser')
                texts = list(soup.stripped_strings)
                
                for i, text in enumerate(texts):
                    if text in ["대표자", "대표자명"]:
                        if info["ceo"] == "정보 없음" and i + 1 < len(texts):
                            info["ceo"] = texts[i + 1]
                    elif text in ["업종", "산업(업종)"]:
                        if info["industry"] == "정보 없음" and i + 1 < len(texts):
                            info["industry"] = texts[i + 1]
                    elif text in ["매출액"]:
                        if info["revenue"] == "정보 없음" and i + 1 < len(texts):
                            info["revenue"] = texts[i + 1]
                    elif text in ["영업이익"]:
                        if info["profit"] == "정보 없음" and i + 1 < len(texts):
                            info["profit"] = texts[i + 1]
                
                info["source"] = "네이버 기업정보 요약"
        except Exception:
            pass
            
    return info
# ==========================================
# ★ 카카오 지오코딩 API ★
# ==========================================
@st.cache_data(show_spinner=False, max_entries=5000)
def get_lat_lon_kakao(company_name, address, rest_api_key):
    headers = {"Authorization": f"KakaoAK {rest_api_key}"}
    
    clean_addr = ""
    if address and address != "등록된 주소 정보가 없습니다.":
        clean_addr = re.sub(r'\(.*?\)|\[.*?\]', '', str(address))
        clean_addr = clean_addr.split(',')[0].strip()
        
    temp_name = re.sub(r'\(주\)|\(유\)|\(합\)|주식회사|㈜', '', str(company_name))
    match = re.search(r'\((.*?)\)', temp_name)
    
    clean_name = ""
    if match:
        clean_name = match.group(1).strip()
    else:
        clean_name = re.sub(r'^[zZ]', '', temp_name).strip()
    if clean_addr:
        try:
            res = requests.get("https://dapi.kakao.com/v2/local/search/address.json", headers=headers, params={"query": clean_addr}, timeout=3)
            if res.status_code == 200 and res.json().get('documents'):
                return float(res.json()['documents'][0]['y']), float(res.json()['documents'][0]['x'])
        except: pass
        
        try:
            res = requests.get("https://dapi.kakao.com/v2/local/search/keyword.json", headers=headers, params={"query": clean_addr}, timeout=3)
            if res.status_code == 200 and res.json().get('documents'):
                return float(res.json()['documents'][0]['y']), float(res.json()['documents'][0]['x'])
        except: pass
        
    if clean_name:
        try:
            res = requests.get("https://dapi.kakao.com/v2/local/search/keyword.json", headers=headers, params={"query": clean_name}, timeout=3)
            if res.status_code == 200 and res.json().get('documents'):
                return float(res.json()['documents'][0]['y']), float(res.json()['documents'][0]['x'])
        except: pass
        
    if company_name:
        try:
            res = requests.get("https://dapi.kakao.com/v2/local/search/keyword.json", headers=headers, params={"query": str(company_name)}, timeout=3)
            if res.status_code == 200 and res.json().get('documents'):
                return float(res.json()['documents'][0]['y']), float(res.json()['documents'][0]['x'])
        except: pass
    return None, None
KAKAO_REST_API_KEY = "21a8c4d7312051598c2e05dba0b9c0c7"
@st.cache_data(show_spinner=False, max_entries=2000)
def kakao_place_search(query, rest_api_key=None, size=15):
    """주소·상호·명칭 통합검색 → 후보 목록.
    반환: [{lat, lon, label, place_name, address}, ...]
    """
    key = rest_api_key or KAKAO_REST_API_KEY
    q = str(query or "").strip()
    if not q:
        return []
    headers = {"Authorization": f"KakaoAK {key}"}
    cands = []
    seen = set()

    def _add(lat, lon, label, place_name="", address=""):
        try:
            lat_f, lon_f = float(lat), float(lon)
        except Exception:
            return
        sig = (round(lat_f, 5), round(lon_f, 5), str(label).strip())
        if sig in seen:
            return
        seen.add(sig)
        cands.append(
            {
                "lat": lat_f,
                "lon": lon_f,
                "label": str(label).strip() or q,
                "place_name": str(place_name or "").strip(),
                "address": str(address or "").strip(),
            }
        )

    # 1) 키워드(장소명·상호·지점)
    try:
        res = requests.get(
            "https://dapi.kakao.com/v2/local/search/keyword.json",
            headers=headers,
            params={"query": q, "size": min(int(size or 15), 15)},
            timeout=5,
        )
        if res.status_code == 200:
            for doc in res.json().get("documents") or []:
                place = str(doc.get("place_name") or "").strip()
                road = str(doc.get("road_address_name") or "").strip()
                jibun = str(doc.get("address_name") or "").strip()
                addr = road or jibun
                label = f"{place} · {addr}" if place and addr else (place or addr or q)
                _add(doc.get("y"), doc.get("x"), label, place, addr)
    except Exception:
        pass
    # 2) 주소 검색 (키워드에 없을 때 보강)
    try:
        res = requests.get(
            "https://dapi.kakao.com/v2/local/search/address.json",
            headers=headers,
            params={"query": q},
            timeout=5,
        )
        if res.status_code == 200:
            for doc in res.json().get("documents") or []:
                addr = doc.get("address_name") or q
                if doc.get("road_address"):
                    addr = doc["road_address"].get("address_name") or addr
                elif doc.get("address"):
                    addr = doc["address"].get("address_name") or addr
                _add(doc.get("y"), doc.get("x"), addr, "", addr)
    except Exception:
        pass
    return cands
def geocode_address_kakao(address, rest_api_key=None):
    """주소·상호·명칭 통합검색 → (lat, lon, 표기주소). 후보 1순위 사용."""
    cands = kakao_place_search(address, rest_api_key=rest_api_key, size=10)
    if not cands:
        return None, None, ""
    c0 = cands[0]
    return c0["lat"], c0["lon"], c0["label"]
def haversine_km(lat1, lon1, lat2, lon2):
    """두 좌표 직선거리(km) — Haversine."""
    r = 6371.0
    p1, p2 = np.radians(lat1), np.radians(lat2)
    dphi = np.radians(lat2 - lat1)
    dlmb = np.radians(lon2 - lon1)
    a = np.sin(dphi / 2) ** 2 + np.cos(p1) * np.cos(p2) * np.sin(dlmb / 2) ** 2
    return float(2 * r * np.arcsin(np.sqrt(a)))
# 한국도로공사 차종: 20톤 벌크로리 ≈ 5종(특수). 길찾기 API 통행료는 보통 1종(승용).
# 실무 근사 배수(1종 대비): 3종 1.5 / 4종 2.0 / 5종 2.5
TOLL_CLASS_5_MULT = 2.5
def _toll_for_bulk20t(toll_class1_won):
    """1종 통행료 → 20톤 벌크(5종) 편도 통행료(원, 10원 단위)."""
    raw = float(toll_class1_won or 0) * TOLL_CLASS_5_MULT
    return float(int(round(raw / 10.0) * 10))
@st.cache_data(show_spinner=False, max_entries=2000)
def kakao_route_from_coords(o_lat, o_lon, d_lat, d_lon, o_label, d_label, rest_api_key=None, _route_v=4):
    """좌표 기준 자동차 거리(km) + 통행료(원, 편도)."""
    _ = _route_v
    key = rest_api_key or KAKAO_REST_API_KEY
    o_label = str(o_label or "출발")
    d_label = str(d_label or "도착")
    try:
        o_lat, o_lon = float(o_lat), float(o_lon)
        d_lat, d_lon = float(d_lat), float(d_lon)
    except Exception:
        return {
            "ok": False, "km": 0.0, "toll": 0.0, "toll_class1": 0.0, "method": "",
            "origin_label": o_label, "dest_label": d_label,
            "message": "좌표가 올바르지 않습니다.",
        }
    try:
        headers = {"Authorization": f"KakaoAK {key}", "Content-Type": "application/json"}
        res = requests.get(
            "https://apis-navi.kakaomobility.com/v1/directions",
            headers=headers,
            params={
                "origin": f"{o_lon},{o_lat}",
                "destination": f"{d_lon},{d_lat}",
                "priority": "RECOMMEND",
                "car_hipass": "true",
            },
            timeout=8,
        )
        if res.status_code == 200:
            routes = res.json().get("routes") or []
            if routes:
                summary = routes[0].get("summary") or {}
                dist_m = float(summary.get("distance") or 0)
                fare = summary.get("fare") or {}
                toll_c1 = float(fare.get("toll") or 0)
                toll_5 = _toll_for_bulk20t(toll_c1)
                if dist_m > 0:
                    msg = (
                        f"카카오 길찾기 {dist_m/1000.0:.1f}km  ·  "
                        f"통행료 1종 {toll_c1:,.0f}원 → 20톤벌크(5종×{TOLL_CLASS_5_MULT:g}) {toll_5:,.0f}원/편도"
                    )
                    return {
                        "ok": True,
                        "km": dist_m / 1000.0,
                        "toll": toll_5,
                        "toll_class1": toll_c1,
                        "method": "kakao_mobility",
                        "origin_label": o_label,
                        "dest_label": d_label,
                        "message": msg,
                    }
    except Exception:
        pass
    straight = haversine_km(o_lat, o_lon, d_lat, d_lon)
    road_km = straight * 1.3
    return {
        "ok": True,
        "km": road_km,
        "toll": 0.0,
        "toll_class1": 0.0,
        "method": "haversine_x1.3",
        "origin_label": o_label,
        "dest_label": d_label,
        "message": (
            f"길찾기 API 미사용/실패 → 직선 {straight:.1f}km × 1.3 = 도로근사 {road_km:.1f}km  ·  "
            f"통행료는 수동 입력"
        ),
    }
@st.cache_data(show_spinner=False, max_entries=2000)
def kakao_route_distance_km(origin_addr, dest_addr, rest_api_key=None, _route_v=4):
    """출발·도착(주소/상호/명칭) → 자동차 거리(km) + 통행료(원, 편도)."""
    _ = _route_v
    key = rest_api_key or KAKAO_REST_API_KEY
    o_lat, o_lon, o_label = geocode_address_kakao(origin_addr, key)
    d_lat, d_lon, d_label = geocode_address_kakao(dest_addr, key)
    if o_lat is None or d_lat is None:
        return {
            "ok": False,
            "km": 0.0,
            "toll": 0.0,
            "toll_class1": 0.0,
            "method": "",
            "origin_label": o_label or str(origin_addr),
            "dest_label": d_label or str(dest_addr),
            "message": "출발/도착을 찾지 못했습니다. 주소·상호·명칭을 확인해 주세요.",
        }
    return kakao_route_from_coords(
        o_lat, o_lon, d_lat, d_lon, o_label, d_label, rest_api_key=key, _route_v=_route_v
    )
def _parse_ton_number(token):
    """톤수 숫자 토큰 → float. '4.9'/'4,9' 지원."""
    t = str(token or "").strip().replace(",", ".")
    t = re.sub(r"[^0-9.]", "", t)
    if not t or t == ".":
        return 0.0
    try:
        return float(t)
    except Exception:
        return 0.0
# 액화가스 내용적(L) → kg 환산 밀도 (대략값, 액화 기준 kg/L)
GAS_DENSITY_KG_PER_L = {
    "질소": 0.808,    # LN2
    "알곤": 1.395,    # LAr
    "산소": 1.141,    # LOX
    "탄산": 1.101,    # LCO2
    "수소": 0.0708,   # LH2
    "헬륨": 0.125,    # LHe
}
# 액화 kg → 기체 Nm³ 환산 (0℃, 1atm 대략값)
GAS_NM3_PER_KG = {
    "질소": 0.7996,
    "알곤": 0.5605,
    "산소": 0.6998,
    "탄산": 0.5090,
    "수소": 11.126,
    "헬륨": 5.596,
}
GAS_OPTIONS = list(GAS_DENSITY_KG_PER_L.keys())
def liters_to_tank_kg(liters, gas_name):
    """내용적(L) × 가스밀도 → 탱크 용량(kg)."""
    dens = float(GAS_DENSITY_KG_PER_L.get(str(gas_name), 0.808) or 0.808)
    return float(liters or 0) * dens
def tank_kg_to_nm3(tank_kg, gas_name):
    """탱크 용량(kg) → 기체 환산(Nm³)."""
    factor = float(GAS_NM3_PER_KG.get(str(gas_name), 0.0) or 0.0)
    return float(tank_kg or 0) * factor
def tank_liters_to_nm3(liters, gas_name):
    """내용적(L) → 기체 환산(Nm³) = L × 밀도 × Nm³/kg."""
    return tank_kg_to_nm3(liters_to_tank_kg(liters, gas_name), gas_name)
def nm3_per_h_to_kg_per_h(nm3_h, gas_name):
    """시간당 루베(Nm³/h) → kg/h."""
    factor = float(GAS_NM3_PER_KG.get(str(gas_name), 0.0) or 0.0)
    if factor <= 0:
        return 0.0
    return float(nm3_h or 0) / factor
def kg_per_h_to_nm3_per_h(kg_h, gas_name):
    """시간당 kg/h → 루베(Nm³/h)."""
    factor = float(GAS_NM3_PER_KG.get(str(gas_name), 0.0) or 0.0)
    return float(kg_h or 0) * factor
def gas_conversion_rows(liters, tank_kg, selected_gas):
    """선택 가스·전 가스 기체환산 표용 rows."""
    rows = []
    for g in GAS_OPTIONS:
        dens = GAS_DENSITY_KG_PER_L[g]
        nm3_per_kg = GAS_NM3_PER_KG[g]
        # 동일 내용적(L) 기준 비교 + 현재 탱크 kg 기준(선택 가스)도 표기
        kg_from_l = float(liters or 0) * dens
        nm3_from_l = kg_from_l * nm3_per_kg
        nm3_from_kg = float(tank_kg or 0) * nm3_per_kg if g == selected_gas else None
        rows.append(
            {
                "가스": g + (" ★" if g == selected_gas else ""),
                "밀도(kg/L)": dens,
                "Nm³/kg": nm3_per_kg,
                "내용적 기준 kg": round(kg_from_l, 1),
                "내용적 기준 Nm³": round(nm3_from_l, 1),
                "현재탱크 Nm³": round(nm3_from_kg, 1) if nm3_from_kg is not None else "—",
            }
        )
    return rows
def compute_tank_usage_cycle(
    tank_kg, hourly_usage_kg, operating_hours, fill_ratio=0.8, days_per_month=30.0
):
    """시간당 사용량·가동시간 → 탱크 사용주기.
    충전기준 = 탱크kg × 80% 소모 시 충전.
    사용주기(일) = 충전기준 ÷ (시간당사용량 × 일가동시간)
    월 사용량 = 일사용량 × 월가동일수
    """
    tank = float(tank_kg or 0)
    hourly = float(hourly_usage_kg or 0)
    hours = float(operating_hours or 0)
    days_m = max(0.0, float(days_per_month or 0))
    ratio = float(fill_ratio or 0) or 0.8
    charge_kg = tank * ratio
    daily_kg = hourly * hours
    monthly_kg = daily_kg * days_m
    if hourly <= 0 or charge_kg <= 0:
        return {
            "ok": False,
            "charge_kg": charge_kg,
            "daily_kg": daily_kg,
            "monthly_kg": monthly_kg,
            "days_per_month": days_m,
            "cycle_days": 0.0,
            "cycle_hours": 0.0,
            "fills_per_month": 0.0,
            "message": "시간당 사용량과 탱크 용량을 입력하세요.",
        }
    cycle_hours = charge_kg / hourly  # 가동시간 누적 기준
    cycle_days = (charge_kg / daily_kg) if daily_kg > 0 else 0.0
    fills_per_month = (days_m / cycle_days) if cycle_days > 0 else 0.0
    return {
        "ok": True,
        "charge_kg": charge_kg,
        "daily_kg": daily_kg,
        "monthly_kg": monthly_kg,
        "days_per_month": days_m,
        "cycle_days": cycle_days,
        "cycle_hours": cycle_hours,
        "fills_per_month": fills_per_month,
        "message": (
            f"충전기준 {charge_kg:,.0f}kg(탱크×{ratio*100:.0f}%) ÷ "
            f"일사용 {daily_kg:,.0f}kg = 사용주기 {cycle_days:,.0f}일"
        ),
    }
def parse_tank_capacity_kg(tank_value):
    """TANK 용량 → kg.
    - 숫자만 입력: kg로 그대로 사용 (예: 4900)
    - 구형식(4.9t, 35T*2 등): 톤→kg 환산 (하위호환)
    """
    if isinstance(tank_value, (int, float)) and not isinstance(tank_value, bool):
        return float(tank_value) if float(tank_value) > 0 else 0.0
    s = str(tank_value or "").strip()
    if not s:
        return 0.0
    for ch in ("，", "·", "．", "･", "。"):
        s = s.replace(ch, ".")
    s_norm = s.replace(",", ".")
    # 숫자만 → kg
    if re.fullmatch(r"\d+(?:\.\d+)?", s_norm):
        return float(s_norm)
    # 구형식: 톤 표기 → kg
    unit = r"(?:TON|톤|T|t)"
    pairs = re.findall(
        rf"(?<![0-9.])(\d+(?:[.,]\d+)?)\s*{unit}\s*[*xX×]\s*(\d+(?:[.,]\d+)?)",
        s,
    )
    if pairs:
        tons = sum(_parse_ton_number(a) * _parse_ton_number(b) for a, b in pairs)
        return float(tons * 1000.0)
    singles = re.findall(rf"(?<![0-9.])(\d+(?:[.,]\d+)?)\s*{unit}\b", s)
    if singles:
        tons = sum(_parse_ton_number(a) for a in singles)
        return float(tons * 1000.0)
    return 0.0
def compute_roundtrips_from_tank(monthly_usage_kg, tank_value, fill_ratio=0.8):
    """왕복횟수 = ceil(월평균공급량 ÷ (탱크용량kg × 0.8)).
    탱크(kg)의 80%가 소모되었을 때 1회 충전(왕복).
    """
    import math
    tank_kg = parse_tank_capacity_kg(tank_value)
    ratio = float(fill_ratio or 0) or 0.8
    charge_kg = tank_kg * ratio
    monthly = float(monthly_usage_kg or 0)
    if charge_kg <= 0:
        return {
            "ok": False,
            "roundtrips": 0.0,
            "roundtrips_exact": 0.0,
            "tank_tons": 0.0,
            "tank_kg": tank_kg,
            "usable_kg": 0.0,
            "fill_ratio": ratio,
            "message": "TANK 용량(kg)을 숫자로 입력하세요. 예: 4900",
        }
    exact = (monthly / charge_kg) if monthly > 0 else 0.0
    trips = int(math.ceil(exact - 1e-12)) if exact > 0 else 0
    return {
        "ok": True,
        "roundtrips": float(trips),
        "roundtrips_exact": exact,
        "tank_tons": tank_kg / 1000.0,
        "tank_kg": tank_kg,
        "usable_kg": charge_kg,
        "fill_ratio": ratio,
        "message": (
            f"탱크 {tank_kg:,.0f}kg의 {ratio*100:.0f}% 소모 시 충전 "
            f"= {charge_kg:,.0f}kg/회 → 월공급 {monthly:,.0f}kg ÷ 충전량 = {trips}회 "
            f"(정확 {exact:.2f}, 올림)"
        ),
    }
def profit_int_comma_input(label, *, key, value=0, help=None):
    """정수 입력 + 천단위 콤마 표시 (소수점 없음). Tab9 투자비 전용."""
    txt_key = f"{key}__comma"
    try:
        init_n = int(round(float(value or 0)))
    except Exception:
        init_n = 0
    if key not in st.session_state:
        st.session_state[key] = init_n
    if txt_key not in st.session_state:
        st.session_state[txt_key] = f"{int(st.session_state[key]):,}"

    def _sync():
        raw = str(st.session_state.get(txt_key, "") or "")
        digits = re.sub(r"[^\d]", "", raw)
        n = int(digits) if digits else 0
        st.session_state[key] = n
        st.session_state[txt_key] = f"{n:,}"

    st.text_input(label, key=txt_key, on_change=_sync, help=help)
    # 화면에 남아 있는 미커밋 값도 숫자로 해석
    raw_now = str(st.session_state.get(txt_key, "") or "")
    digits_now = re.sub(r"[^\d]", "", raw_now)
    if digits_now:
        st.session_state[key] = int(digits_now)
    return float(st.session_state.get(key, 0) or 0)
def compute_logistics_unit_cost(
    distance_km,
    fuel_price_per_l,
    fuel_efficiency_km_per_l,
    toll_fee=0.0,
    round_trips=1.0,
    load_kg=1.0,
):
    """표준 탱크로리 물류비(원/kg) 산식.
    연료비(편도) = 거리(km) × 유류비(원/L) ÷ 연비(km/L)
    편도비용 = 연료비 + 통행료
    왕복총비용 = 편도비용 × 왕복횟수
    물류비(원/kg) = 왕복총비용 ÷ 공급량(kg)
    (엑셀 P18과 동일: (KM×유류비/연비+통행료)×왕복÷KG)
    """
    eff = float(fuel_efficiency_km_per_l or 0) or 1.0
    kg = float(load_kg or 0) or 1.0
    fuel_one_way = float(distance_km or 0) * float(fuel_price_per_l or 0) / eff
    one_way = fuel_one_way + float(toll_fee or 0)
    total = one_way * float(round_trips or 0)
    per_kg = total / kg
    return {
        "fuel_one_way": fuel_one_way,
        "one_way_cost": one_way,
        "round_trip_cost": total,
        "per_kg": per_kg,
    }
DIESEL_PRICE_FILE = os.path.join(CACHE_DIR, "diesel_market_price.json")
def fetch_diesel_market_price():
    """전국 주유소 경유 평균가(원/L).
    소스: 네이버 금융 OIL_LO(경유) = 한국석유공사 Opinet 전국평균.
    20톤 벌크로리(경유) 물류비 산정용 시장 정가 기준.
    """
    url = "https://finance.naver.com/marketindex/oilDetail.naver?marketindexCd=OIL_LO"
    headers = {"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36"}
    try:
        res = requests.get(url, headers=headers, timeout=10)
        res.encoding = res.apparent_encoding or "euc-kr"
        soup = BeautifulSoup(res.text, "html.parser")
        nt = soup.select_one("p.no_today")
        price = None
        if nt:
            for b in nt.select(".blind"):
                b.decompose()
            spaced = re.sub(r"\s+", "", nt.get_text(" ", strip=True))
            num = re.search(r"([0-9]+(?:,[0-9]+)*(?:\.[0-9]+)?)", spaced)
            if num:
                price = float(num.group(1).replace(",", ""))
        if price is None or price < 500 or price > 5000:
            return {"ok": False, "price": 0.0, "asof": "", "source": "", "message": "경유 시세를 파싱하지 못했습니다."}
        asof = datetime.date.today().strftime("%Y-%m-%d")
        dm = re.search(r"(20\d{2}\.\d{2}\.\d{2})", res.text)
        if dm:
            asof = dm.group(1).replace(".", "-")
        return {
            "ok": True,
            "price": round(price, 2),
            "asof": asof,
            "source": "네이버금융·오피넷 전국 경유 평균",
            "message": f"경유 {price:,.2f} 원/L ({asof})",
        }
    except Exception as e:
        return {"ok": False, "price": 0.0, "asof": "", "source": "", "message": f"경유 시세 조회 실패: {e}"}
def load_diesel_price_cache():
    if os.path.exists(DIESEL_PRICE_FILE):
        try:
            with open(DIESEL_PRICE_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict):
                return data
        except Exception:
            pass
    return {}
def save_diesel_price_cache(data):
    try:
        with open(DIESEL_PRICE_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass
def get_diesel_price_monthly(force_refresh=False):
    """매월 시장가 자동 갱신(캐시). force_refresh면 즉시 재조회."""
    month = datetime.date.today().strftime("%Y-%m")
    cached = load_diesel_price_cache()
    if (
        not force_refresh
        and cached.get("ok")
        and cached.get("month") == month
        and float(cached.get("price") or 0) > 0
    ):
        return cached
    fetched = fetch_diesel_market_price()
    if fetched.get("ok"):
        fetched["month"] = month
        fetched["vehicle"] = "20톤 벌크로리 · 경유"
        save_diesel_price_cache(fetched)
        return fetched
    # 조회 실패 시 이전 캐시라도 사용
    if cached.get("ok") and float(cached.get("price") or 0) > 0:
        cached = dict(cached)
        cached["message"] = f"최신 조회 실패 → 캐시 사용 ({cached.get('asof','')})"
        return cached
    fetched["month"] = month
    return fetched
# ==========================================
# ★ 메모 생성 AppleScript ★ 
# ==========================================
def open_macos_notes_folder(client_name, dart_api_key, df_integrated=None):
    safe_client_name = client_name.replace('"', '\\"')
    info = get_company_info_hybrid(client_name, dart_api_key)
    encoded_name = urllib.parse.quote(info['clean_name'])
    
    # 통합 탱크 재고 연동 HTML 생성
    inventory_html = ""
    if df_integrated is not None and not df_integrated.empty and '거래처(사용처)/보관장소' in df_integrated.columns:
        client_inv = df_integrated[df_integrated['거래처(사용처)/보관장소'].astype(str).str.contains(client_name, regex=False, na=False)]
        if not client_inv.empty:
            inventory_html = "<h3>🛢️ 설치/보관 장비 현황 (통합 탱크 재고)</h3><ul>"
            for _, row in client_inv.iterrows():
                item = row.get('품목', '미상')
                status = row.get('사용구분', '')
                serial = row.get('일련(제조)번호', 'S/N 없음')
                vol = row.get('저장부피(L)', '')
                weight = row.get('저장무게(kg)', '')
                
                vol_str = f"{vol}L" if pd.notna(vol) and str(vol).strip() != "" else ""
                weight_str = f"{weight}kg" if pd.notna(weight) and str(weight).strip() != "" else ""
                cap_str = f" / 용량: {vol_str} {weight_str}".strip() if vol_str or weight_str else ""
                
                inventory_html += f"<li><b>[{item}]</b> {status} (S/N: {serial}{cap_str})</li>"
            inventory_html += "</ul><br>"
    note_content = f"""
    <h1>{safe_client_name}</h1>
    <br>
    <h3>📌 요약 기업 정보 (데이터 출처: {info['source']})</h3>
    <ul>
        <li><b>대표자:</b> {info['ceo']}</li>
        <li><b>업종:</b> {info['industry']}</li>
        <li><b>매출액:</b> {info['revenue']}</li>
        <li><b>영업이익:</b> {info['profit']}</li>
    </ul>
    <br>
    {inventory_html}
    <h3>🔗 상세 정보 원클릭 검색</h3>
    <ul>
        <li><a href="https://search.naver.com/search.naver?query={encoded_name} 기업정보">네이버에서 '{info['clean_name']}' 재무정보 보기</a></li>
        <li><a href="https://www.saramin.co.kr/zf_user/search/company?searchword={encoded_name}">사람인에서 '{info['clean_name']}' 기업/채용 검색</a></li>
        <li><a href="https://www.jobkorea.co.kr/Search/?stext={encoded_name}&tabType=corp">잡코리아에서 '{info['clean_name']}' 기업 검색</a></li>
    </ul>
    <br>
    <h3>📝 영업 및 특이사항</h3>
    <p></p>
    """
    safe_note_content = note_content.replace('"', '\\"')
    script = f"""
    tell application "Notes"
        activate
        
        set targetFolderName to "거래처"
        set noteName to "{safe_client_name}"
        set noteFound to false
        set targetAcc to missing value
        
        repeat with acc in accounts
            try
                if exists folder targetFolderName of acc then
                    set parentFolder to folder targetFolderName of acc
                    set foundNotes to (notes of parentFolder whose name is noteName)
                    if (count of foundNotes) > 0 then
                        show (item 1 of foundNotes)
                        set noteFound to true
                        exit repeat
                    end if
                end if
            end try
        end repeat
        
        if not noteFound then
            repeat with acc in accounts
                if name of acc is "iCloud" then
                    set targetAcc to acc
                    exit repeat
                end if
            end repeat
            
            if targetAcc is missing value then
                set targetAcc to first account
            end if
            
            if not (exists folder targetFolderName of targetAcc) then
                make new folder at targetAcc with properties {{name:targetFolderName}}
            end if
            
            set parentFolder to folder targetFolderName of targetAcc
            
            set newNote to make new note at parentFolder with properties {{body:"{safe_note_content}"}}
            
            show newNote
        end if
    end tell
    """
    try:
        process = subprocess.Popen(['osascript', '-'], stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        stdout, stderr = process.communicate(script.encode('utf-8'))
        
        if process.returncode == 0:
            return True
        else:
            return False
    except Exception as e:
        return False
def create_stacked_bar_chart(pivot_df, title_text="", y_suffix="", y_format=",.0f"):
    fig = go.Figure()
    sorted_years = sorted(pivot_df.columns, key=lambda x: str(x))
    
    color_map = {
        '2020': '#0052CC',
        '2021': '#4C9AFF',
        '2022': '#FF2B2B',
        '2023': '#FF9999',
        '2024': '#00B894',
        '2025': '#55E6A5',
        '2026': '#FF9F1A'
    }
    for yr in sorted_years:
        col_name = str(yr)
        color = color_map.get(col_name, None)
        fig.add_trace(
            go.Bar(
                x=pivot_df.index,
                y=pivot_df[yr],
                name=col_name,
                marker_color=color,
                hovertemplate=f"%{{x}} ({col_name}): %{{y:{y_format}}}{y_suffix}<extra></extra>"
            )
        )
    layout_args = dict(
        barmode='stack',
        xaxis=dict(title=None, tickangle=0),
        yaxis=dict(title=None, gridcolor='#E2E8F0'),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.25,
            xanchor="center",
            x=0.5
        ),
        margin=dict(l=10, r=10, t=10, b=40), 
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        height=420
    )
    
    if title_text:
        layout_args['title'] = dict(text=title_text, font=dict(size=14, color="#334155"))
        layout_args['margin']['t'] = 40
        
    fig.update_layout(**layout_args)
    return fig
# ==========================================
# 3. 데이터 로딩 & 메모리 캐싱 (최적화) - Error 무시(on_bad_lines) 적용
# ==========================================
@st.cache_data(show_spinner="주소록을 읽어오는 중입니다...")
def load_address_file(address_bytes):
    if not address_bytes:
        return {}
    try:
        for enc in ["utf-8-sig", "cp949", "euc-kr", "utf-8"]:
            try:
                temp_addr = pd.read_csv(io.BytesIO(address_bytes), encoding=enc, on_bad_lines='skip', engine='python')
                if len(temp_addr.columns) >= 2:
                    k_col = temp_addr.columns[0]
                    v_col = temp_addr.columns[1]
                    temp_addr = temp_addr.dropna(subset=[k_col])
                    return (
                        temp_addr.astype(str)
                        .set_index(k_col)[v_col]
                        .to_dict()
                    )
                break
            except UnicodeDecodeError:
                continue
    except Exception:
        pass
    return {}
def _drop_debt_noise_rows(df):
    """ERP 내보내기 푸터·합계 행을 제거 (거래처/채권 데이터 오염 방지)."""
    if df.empty or "거래처" not in df.columns:
        return df
    client_raw = df["거래처"].astype(str).str.strip()
    gubun_raw = (
        df["구분"].astype(str).str.strip()
        if "구분" in df.columns
        else pd.Series("", index=df.index)
    )
    noise = (
        client_raw.str.match(r"^\d{4}-\d{2}-\d{2}\s", na=False)
        | gubun_raw.str.contains(r"총매출|총수금|합계|소계", case=False, na=False)
        | client_raw.isin(["", "nan", "None", "NaN"])
    )
    return df.loc[~noise].copy()
def _drop_sales_noise_rows(df):
    """매출 CSV 하단 타임스탬프·이월미수 행 제거."""
    if df.empty:
        return df
    noise = pd.Series(False, index=df.index)
    if "거래처" in df.columns:
        noise |= df["거래처"].astype(str).str.match(r"^\d{4}-\d{2}-\d{2}\s", na=False)
    if "품목명" in df.columns:
        noise |= df["품목명"].astype(str).str.contains(r"이월\s*미수|\[이월", na=False)
    return df.loc[~noise].copy()
@st.cache_data(show_spinner="업종 분류 데이터를 읽어오는 중입니다...")
def load_industry_file(industry_bytes):
    if not industry_bytes:
        return {}
    try:
        for enc in ["utf-8-sig", "cp949", "euc-kr", "utf-8"]:
            try:
                temp_ind = pd.read_csv(io.BytesIO(industry_bytes), encoding=enc, on_bad_lines='skip', engine='python')
                temp_ind.columns = temp_ind.columns.astype(str).str.strip()
                
                c_client = next((c for c in temp_ind.columns if "거래처" in c or "상호" in c), None)
                c_ind = next((c for c in temp_ind.columns if "분류" in c or "업종" in c), None)
                
                if c_client and c_ind:
                    temp_ind = temp_ind.dropna(subset=[c_client, c_ind])
                    return temp_ind.astype(str).set_index(c_client)[c_ind].to_dict()
                break
            except UnicodeDecodeError:
                continue
    except Exception:
        pass
    return {}
@st.cache_data(show_spinner="채권 데이터를 읽어오는 중입니다...")
def load_debt_file(debt_bytes):
    if not debt_bytes:
        return pd.DataFrame()
    try:
        for enc in ["utf-8-sig", "cp949", "euc-kr", "utf-8"]:
            try:
                df_direct = pd.read_csv(io.BytesIO(debt_bytes), encoding=enc, on_bad_lines='skip', engine='python')
                df_direct.columns = df_direct.columns.astype(str).str.strip()
                
                # ★ 채권 CSV 파일에 잘못 생성된 빈 껍데기 열(Unnamed) 무조건 제거
                df_direct = df_direct.loc[:, ~df_direct.columns.str.contains('^Unnamed')]
                
                if "거래처" in df_direct.columns and "구분" in df_direct.columns:
                    df_direct = _drop_debt_noise_rows(df_direct)
                    # 💡 거래처 이름 공백 제거로 누락 데이터 완벽 방지
                    df_direct["거래처"] = df_direct["거래처"].replace("", np.nan).ffill().astype(str).str.strip()
                    
                    def map_gubun(val):
                        val_clean = re.sub(r'[^가-힣a-zA-Z0-9]', '', str(val))
                        
                        if not val_clean or val_clean == 'nan': 
                            return "매출"
                            
                        if any(k in val_clean for k in ["이월", "전월", "기초"]): return "이월"
                        if any(k in val_clean for k in ["익월", "다음", "차월"]): return "익월"
                        if any(k in val_clean for k in ["잔액", "미수", "현재", "기말", "잔금"]): return "잔액"
                        if any(k in val_clean for k in ["수금", "입금", "결제", "회수", "대변", "감소"]): return "수금"
                        
                        return "매출"
                    df_direct["구분"] = df_direct["구분"].apply(map_gubun)
                    valid_types = ["익월", "이월", "매출", "수금", "잔액"]
                    df_filtered = df_direct[df_direct["구분"].isin(valid_types)].copy()
                    
                    if not df_filtered.empty:
                        month_cols = [c for c in df_filtered.columns if c not in ["거래처", "구분"]]
                        for m_col in month_cols:
                            df_filtered[m_col] = (
                                pd.to_numeric(
                                    df_filtered[m_col].astype(str).str.replace(r"[^\d.-]", "", regex=True),
                                    errors="coerce"
                                ).fillna(0)
                            )
                        return df_filtered
                break
            except UnicodeDecodeError:
                continue
    except Exception:
        pass
    return pd.DataFrame()
@st.cache_data(show_spinner="데이터 파싱 중입니다...")
def load_uploaded_files_from_bytes(file_tuples):
    if not file_tuples:
        return pd.DataFrame()
    df_list = []
    for file_name, content in file_tuples:
        try:
            decoded_text = None
            for enc in ["utf-8-sig", "cp949", "euc-kr", "utf-8"]:
                try:
                    decoded_text = content.decode(enc)
                    break
                except UnicodeDecodeError:
                    continue
            if decoded_text is None:
                decoded_text = content.decode("utf-8", errors="replace")
            lines = [line for line in decoded_text.splitlines() if line.strip()]
            if not lines:
                continue
            header_idx = 0
            max_matches = 0
            for i, line in enumerate(lines[:30]):
                cells = re.split(r',|\t', line)
                matches = sum(1 for cell in cells if any(kw in cell for kw in ["거래처", "품목", "매출", "단가", "수량", "담당", "일자"]))
                
                if matches > max_matches:
                    max_matches = matches
                    header_idx = i
                
                if max_matches >= 3:
                    break
            df = pd.read_csv(io.StringIO("\n".join(lines[header_idx:])), on_bad_lines='skip', engine='python')
            df.columns = df.columns.astype(str).str.strip()
            cols = list(df.columns)
            def find_col(priority_keywords, exclude_keywords=[]):
                clean_cols = [c.replace(" ", "") for c in cols]
                
                for kw in priority_keywords:
                    kw_clean = kw.replace(" ", "")
                    for orig_c, clean_c in zip(cols, clean_cols):
                        if any(ex in clean_c for ex in exclude_keywords):
                            continue
                        if kw_clean == clean_c:
                            return orig_c
                            
                for kw in priority_keywords:
                    kw_clean = kw.replace(" ", "")
                    for orig_c, clean_c in zip(cols, clean_cols):
                        if any(ex in clean_c for ex in exclude_keywords):
                            continue
                        if kw_clean in clean_c:
                            return orig_c
                return None
            c_staff = find_col(["담당자", "담당자명", "영업담당", "영업사원", "담당"], ["코드", "ID", "번호"])
            c_client = find_col(["거래처", "거래처명", "상호명", "고객명", "회사명", "상호", "고객"], ["코드", "ID", "번호", "담당", "영업"])
            c_item = find_col(["품목명", "제품명", "상품명", "품목", "제품"], ["코드", "ID", "번호", "규격"])
            c_sales = find_col(["매출액", "금액", "매출"], ["일", "자", "수량", "량", "단가"])
            c_qty = find_col(["출고량", "수량", "출고"], ["액", "금액", "단가"])
            c_price = find_col(["단가", "단 가", "판매단가", "공급단가"], ["액", "금액", "수량", "량"])
            c_date = find_col(["매출일자", "매출일", "일자", "날짜", "출고일"])
            rename_dict = {}
            if c_client: rename_dict[c_client] = "거래처"
            if c_item: rename_dict[c_item] = "품목명"
            if c_staff: rename_dict[c_staff] = "담당자"
            if c_sales: rename_dict[c_sales] = "매출액"
            if c_qty: rename_dict[c_qty] = "출고량"
            if c_price: rename_dict[c_price] = "단가"
            if c_date: rename_dict[c_date] = "매출일자_raw"
            df = df.rename(columns=rename_dict)
            df = _drop_sales_noise_rows(df)
            for req in ["거래처", "품목명", "담당자"]:
                if req not in df.columns:
                    df[req] = "미지정"
            file_year = next(
                (y for y in ["2020", "2021", "2022", "2023", "2024", "2025", "2026"] if y in file_name),
                "2026"
            )
            date_col = "매출일자_raw" if "매출일자_raw" in df.columns else df.columns[0]
            df["매출일_dt"] = parse_date_series_robust(df[date_col], default_year=file_year)
            df["매출액"] = pd.to_numeric(
                df["매출액"].astype(str).str.replace(r"[^\d.-]", "", regex=True), errors="coerce"
            ).fillna(0) if "매출액" in df.columns else 0
            df["출고량"] = pd.to_numeric(
                df["출고량"].astype(str).str.replace(r"[^\d.-]", "", regex=True), errors="coerce"
            ).fillna(0) if "출고량" in df.columns else 0
            df["단가"] = pd.to_numeric(
                df["단가"].astype(str).str.replace(r"[^\d.-]", "", regex=True), errors="coerce"
            ).fillna(0) if "단가" in df.columns else 0
            df["거래처"] = df["거래처"].fillna("미지정").astype(str).str.strip()
            df["담당자"] = df["담당자"].fillna("미지정").astype(str).str.strip()
            df = normalize_items_vectorized(df)
            df = df.dropna(subset=["매출일_dt"])
            df["연도"] = df["매출일_dt"].dt.year.astype(str)
            df["월"] = df["매출일_dt"].dt.strftime("%m월")
            df["연도월_정렬"] = df["연도"].astype(str).str[2:] + "년 " + df["월"].astype(str)
            if not df.empty:
                df_list.append(df)
        except Exception as e:
            st.sidebar.error(f"파일 읽기 오류 ({file_name}): {e}")
    result_df = pd.concat(df_list, ignore_index=True) if df_list else pd.DataFrame()
    
    # ★ '미지정' 및 '담당자없음' 데이터 통합 및 최신 담당자 자동 추론 로직
    if not result_df.empty and "거래처" in result_df.columns and "담당자" in result_df.columns:
        result_df["담당자"] = result_df["담당자"].fillna("미지정").astype(str).str.strip()
        
        # '담당자없음'을 포함하여 비어있거나 누락된 표현들을 모두 '미지정'으로 통일
        invalid_staff_markers = ["", "nan", "None", "NAT", "NaN", "담당자없음", "지정안함", "없음"]
        result_df["담당자"] = result_df["담당자"].replace(invalid_staff_markers, "미지정")
        
        temp_df = result_df.dropna(subset=["매출일_dt"]).sort_values("매출일_dt")
        valid_staff_map = temp_df[~temp_df["담당자"].isin(["미지정"])].groupby("거래처")["담당자"].last().to_dict()
        
        mask_unassigned = result_df["담당자"] == "미지정"
        result_df.loc[mask_unassigned, "담당자"] = (
            result_df.loc[mask_unassigned, "거래처"]
            .map(valid_staff_map)
            .fillna("미지정")
        )
        result_df.loc[result_df["담당자"].isin(invalid_staff_markers), "담당자"] = "미지정"
    # 수동 지정 매핑 적용 (무손실 보존)
    manual_map_path = os.path.join(CACHE_DIR, "manual_staff_mapping.csv")
    if os.path.exists(manual_map_path) and not result_df.empty:
        try:
            manual_df = pd.read_csv(manual_map_path)
            manual_dict = manual_df.set_index("거래처")["담당자"].to_dict()
            mask_manual = result_df["거래처"].isin(manual_dict.keys())
            result_df.loc[mask_manual, "담당자"] = result_df.loc[mask_manual, "거래처"].map(manual_dict)
        except Exception:
            pass
    return result_df
# ==========================================
# 4. 피벗 및 지표 계산 연산 캐싱 (최적화)
# ==========================================
@st.cache_data
def cached_get_yearly_monthly_pivot(data_df, all_months, years):
    if data_df.empty:
        return pd.DataFrame(0, index=all_months, columns=[str(y) for y in years])
    
    pvt = data_df.pivot_table(
        index="월", columns="연도", values="매출액", aggfunc="sum"
    ).fillna(0) * 1.1 / 10000
    
    pvt = pvt.reindex(index=all_months, fill_value=0)
    all_yrs = [str(y) for y in years]
    pvt = pvt.reindex(columns=all_yrs, fill_value=0)
    return pvt
@st.cache_data
def cached_get_item_pivot(data_df, item_name, metric, all_months, years):
    if data_df.empty:
        return pd.DataFrame(0, index=all_months, columns=[str(y) for y in years])
        
    df_item = data_df[data_df["품목명"] == item_name]
    if df_item.empty:
        return pd.DataFrame(0, index=all_months, columns=[str(y) for y in years])
        
    if metric == "매출액 (만원)":
        pvt = df_item.pivot_table(index="월", columns="연도", values="매출액", aggfunc="sum").fillna(0) * 1.1 / 10000
    elif "출고량" in metric:
        pvt = df_item.pivot_table(index="월", columns="연도", values="출고량", aggfunc="sum").fillna(0)
        target_bulks = ["CO2 (kg, Bulk)", "N2 (kg, Bulk)", "O2 (kg, Bulk)", "AR (kg, Bulk)"]
        if item_name in target_bulks:
            pvt = pvt / 1000
    elif metric == "총매출 대비 비중 (%)":
        pvt_item = df_item.pivot_table(index="월", columns="연도", values="매출액", aggfunc="sum").fillna(0)
        pvt_total = data_df.pivot_table(index="월", columns="연도", values="매출액", aggfunc="sum").fillna(0)
        pvt = (pvt_item / pvt_total.replace(0, np.nan) * 100).fillna(0)
        
    pvt = pvt.reindex(index=all_months, fill_value=0)
    all_yrs = [str(y) for y in years]
    pvt = pvt.reindex(columns=all_yrs, fill_value=0)
    return pvt
@st.cache_data
def cached_get_industry_pivot(data_df, industry_name, metric, all_months, years):
    if data_df.empty:
        return pd.DataFrame(0, index=all_months, columns=[str(y) for y in years])
        
    df_ind = data_df[data_df["업종"] == industry_name]
    if df_ind.empty:
        return pd.DataFrame(0, index=all_months, columns=[str(y) for y in years])
        
    if metric == "매출액 (만원)":
        pvt = df_ind.pivot_table(index="월", columns="연도", values="매출액", aggfunc="sum").fillna(0) * 1.1 / 10000
    elif metric == "출고량":
        pvt = df_ind.pivot_table(index="월", columns="연도", values="출고량", aggfunc="sum").fillna(0)
    elif metric == "총매출 대비 비중 (%)":
        pvt_ind = df_ind.pivot_table(index="월", columns="연도", values="매출액", aggfunc="sum").fillna(0)
        pvt_total = data_df.pivot_table(index="월", columns="연도", values="매출액", aggfunc="sum").fillna(0)
        pvt = (pvt_ind / pvt_total.replace(0, np.nan) * 100).fillna(0)
        
    pvt = pvt.reindex(index=all_months, fill_value=0)
    all_yrs = [str(y) for y in years]
    pvt = pvt.reindex(columns=all_yrs, fill_value=0)
    return pvt
@st.cache_data
def cached_tab3_pivots(target_tab3_df, years, all_months):
    sales_p = pd.DataFrame()
    qty_p = pd.DataFrame()
    unit_price_p = pd.DataFrame()
    
    if target_tab3_df.empty:
        return sales_p, qty_p, unit_price_p
    sales_raw_p = target_tab3_df.pivot_table(index="품목명", columns="연도월_정렬", values="매출액", aggfunc="sum").fillna(0)
    qty_raw_p = target_tab3_df.pivot_table(index="품목명", columns="연도월_정렬", values="출고량", aggfunc="sum").fillna(0)
    sales_expanded_data = {}
    qty_expanded_data = {}
    for yr in years:
        yr_short = yr[2:]
        yr_sales_sum = 0
        yr_qty_sum = 0
        for m in all_months:
            col_key = f"{yr_short}년 {m}"
            s_val = sales_raw_p[col_key] if col_key in sales_raw_p.columns else 0
            q_val = qty_raw_p[col_key] if col_key in qty_raw_p.columns else 0
            sales_expanded_data[col_key] = s_val
            qty_expanded_data[col_key] = q_val
            yr_sales_sum += s_val
            yr_qty_sum += q_val
        sales_expanded_data[f"{yr_short}년 연간총합"] = yr_sales_sum
        qty_expanded_data[f"{yr_short}년 연간총합"] = yr_qty_sum
    sales_p = pd.DataFrame(sales_expanded_data, index=sales_raw_p.index)
    qty_p = pd.DataFrame(qty_expanded_data, index=qty_raw_p.index)
    latest_col = None
    for yr in years:
        for m in reversed(all_months):
            chk_c = f"{yr[2:]}년 {m}"
            if chk_c in qty_p.columns and qty_p[chk_c].sum() > 0:
                latest_col = chk_c
                break
        if latest_col: break
    if not latest_col and len(qty_p.columns) > 0:
        latest_col = qty_p.columns[0]
    if latest_col and latest_col in qty_p.columns:
        qty_p = qty_p.sort_values(by=latest_col, ascending=False)
        sales_p = sales_p.reindex(qty_p.index)
    df_for_price = target_tab3_df.copy()
    target_clients = ["두산판금", "드림맥", "모베이스전전", "지엔티테크", "태광기업", "경민산업", "동주산업"]
    mask_n2_special = (df_for_price["거래처"].isin(target_clients)) & (df_for_price["품목명"] == "N2 (kg, Bulk)")
    
    df_for_price.loc[mask_n2_special, "단가"] = df_for_price.loc[mask_n2_special, "단가"] * 1.238
    raw_up = df_for_price.pivot_table(index="품목명", columns="연도월_정렬", values="단가", aggfunc=get_exact_original_price)
    
    if not raw_up.empty:
        unit_price_p = raw_up.fillna(0)
        
        desired_price_cols = [f"{yr[2:]}년 {m}" for yr in years for m in reversed(all_months)]
        existing_cols = [c for c in desired_price_cols if c in unit_price_p.columns]
        unit_price_p = unit_price_p[existing_cols]
        unit_price_p = apply_forward_unit_price(unit_price_p, qty_p, years, all_months)
        
    return sales_p, qty_p, unit_price_p
@st.cache_data
def cached_filter_tab3_year_columns(sales_p, qty_p, selected_detail_years, avail_years_short, all_months):
    """Tab3 연도별 컬럼 필터 (캐시로 탭 전환·재렌더 가속)."""
    def _filter_year_columns(df):
        if df.empty:
            return df
        cols = []
        for y in avail_years_short:
            if y in selected_detail_years:
                for m in reversed(all_months):
                    c = f"{y}년 {m}"
                    if c in df.columns:
                        cols.append(c)
            tot = f"{y}년 연간총합"
            if tot in df.columns:
                cols.append(tot)
        return df[[c for c in cols if c in df.columns]]
    sales_vat = sales_p * 1.1 / 10000
    return _filter_year_columns(sales_vat), _filter_year_columns(qty_p)
@st.cache_data
def cached_tab3_item_client_detail(
    df, item_name, metric, years, all_months, selected_years_short
):
    """품목 1개 → 거래처×월 상세 피벗 (sales=만원 VAT포함 / qty=출고량)."""
    if df is None or df.empty or not item_name:
        return pd.DataFrame()
    d = df[df["품목명"] == item_name]
    if d.empty:
        return pd.DataFrame()
    if metric == "sales":
        raw = (
            d.pivot_table(index="거래처", columns="연도월_정렬", values="매출액", aggfunc="sum")
            .fillna(0)
            * 1.1
            / 10000
        )
    else:
        raw = d.pivot_table(
            index="거래처", columns="연도월_정렬", values="출고량", aggfunc="sum"
        ).fillna(0)
        bulk = ["CO2 (kg, Bulk)", "N2 (kg, Bulk)", "O2 (kg, Bulk)", "AR (kg, Bulk)"]
        if item_name in bulk:
            raw = raw / 1000
    if raw.empty:
        return pd.DataFrame()
    expanded = {}
    for yr in years:
        yr_short = yr[2:]
        yr_sum = 0
        for m in all_months:
            col_key = f"{yr_short}년 {m}"
            val = raw[col_key] if col_key in raw.columns else 0
            expanded[col_key] = val
            yr_sum = yr_sum + val
        expanded[f"{yr_short}년 연간총합"] = yr_sum
    pvt = pd.DataFrame(expanded, index=raw.index)
    avail_years_short = [y[2:] for y in years]
    cols = []
    for y in avail_years_short:
        if y in selected_years_short:
            for m in reversed(all_months):
                c = f"{y}년 {m}"
                if c in pvt.columns:
                    cols.append(c)
        tot = f"{y}년 연간총합"
        if tot in pvt.columns:
            cols.append(tot)
    pvt = pvt[[c for c in cols if c in pvt.columns]]
    if pvt.empty:
        return pvt
    sort_col = None
    for c in pvt.columns:
        if "연간총합" in str(c):
            sort_col = c
            break
    if sort_col is None and len(pvt.columns):
        sort_col = pvt.columns[0]
    if sort_col:
        pvt = pvt.sort_values(by=sort_col, ascending=False)
    return pvt
def render_tab3_item_client_expanders(df_src, items, metric, years, all_months, selected_years_short):
    """상단에서 선택한 품목의 거래처별 상세 — 클릭 펼침 (매출/출고 각각)."""
    if not items:
        return
    label_metric = "매출" if metric == "sales" else "출고량"
    cmap = "Blues" if metric == "sales" else "Greens"
    bulk = ("CO2 (kg, Bulk)", "N2 (kg, Bulk)", "O2 (kg, Bulk)", "AR (kg, Bulk)")
    for item in items:
        with st.expander(
            f"📂 [{item}] 거래처별 {label_metric} 상세 데이터 파보기 (클릭하여 펼치기)",
            expanded=False,
        ):
            pvt = cached_tab3_item_client_detail(
                df_src,
                item,
                metric,
                tuple(years),
                tuple(all_months),
                tuple(sorted(selected_years_short)),
            )
            if pvt is None or pvt.empty:
                st.info(f"선택한 조건에서 [{item}] 거래처 상세 데이터가 없습니다.")
                continue
            pvt_disp = get_display_df_with_sum(pvt, "합계")
            fmt = "{:,.1f}" if metric == "qty" and item in bulk else "{:,.0f}"
            st.dataframe(
                style_with_sum(pvt_disp, fmt, cmap, axis=None),
                use_container_width=True,
                height=min(420, 80 + 28 * min(len(pvt_disp), 12)),
            )
# ==========================================
# ★ 누락되었던 필수 캐시 함수 복구 완료 ★
# ==========================================
@st.cache_data
def cached_client_item_qty_pivot(df_client_filtered, years, all_months):
    if df_client_filtered.empty:
        return pd.DataFrame()
    raw_ci_qty = df_client_filtered.pivot_table(
        index="품목명", columns="연도월_정렬", values="출고량", aggfunc="sum"
    ).fillna(0)
    ci_expanded_data = {}
    for yr in years:
        yr_short = yr[2:]
        for m in all_months:
            col_key = f"{yr_short}년 {m}"
            q_val = raw_ci_qty[col_key] if col_key in raw_ci_qty.columns else 0
            ci_expanded_data[col_key] = q_val
    return pd.DataFrame(ci_expanded_data, index=raw_ci_qty.index)
# ==========================================
# ★ 초고속 빈 껍데기 필터링 로직 ★
# ==========================================
def prepare_active_df_fast(df, target_col):
    if df is None or df.empty:
        return None, [], None
    
    df_active = df.loc[(df != 0).any(axis=1), (df != 0).any(axis=0)].copy()
    
    if df_active.empty:
        return None, [], None
    
    df_active.index.name = None
    df_active = df_active.reset_index().rename(columns={"index": "품목명"})
    df_active.columns.name = None 
    numeric_cols = [c for c in df_active.columns if c != "품목명"]
    highlight_col_name = None
    if target_col and target_col in df_active.columns:
        highlight_col_name = target_col
        
    return df_active, numeric_cols, highlight_col_name
@st.cache_data
def cached_prepare_active_df(df, target_col):
    """Tab3 표 렌더용 활성 행·열 필터 (캐시)."""
    return prepare_active_df_fast(df, target_col)
DEBT_PINK_SOFT = "#FFE4E6"
DEBT_CLIENT_STRIPE_A = "#FFFFFF"
DEBT_CLIENT_STRIPE_B = "#E2E8F0"
def payment_term_credit_months(term):
    """결제조건 → 정상채권으로 인정하는 최대 경과개월(age).
    age0=당월. 익월말 → age<=1(당월·전월) 정상, age>=2 연체.
    """
    if not term or not str(term).strip():
        return 1
    t = str(term).replace(" ", "").replace("　", "")
    if "선매출" in t:
        return 99
    if "당월" in t or "발행후" in t:
        return 0
    m = re.search(r"(\d+)일", t)
    if m:
        days = int(m.group(1))
        # 40일 ≈ 다다음달 10일 결제 → 2개월 유예 / 60일+ → 2개월 유지
        if days >= 40:
            return 2
        return 1
    if "2개월" in t or "이개월" in t:
        return 2
    if "익월말60" in t:
        return 2
    if "bulk-2" in t.lower() or "bulk-2개월" in t or "/bulk" in t.lower():
        return 2
    if "익월" in t:
        return 1
    return 1
def analyze_client_ar(sales_vals, bal_vals, credit_months, current_idx):
    """당월 잔액 배분 → (정상액, 연체액, 연체 매출열 set, 연체개월수).
    연체개월수 = 결제조건 초과 경과개월의 최대값. 잔액0 → 0.
    """
    try:
        cur_bal = float(bal_vals[current_idx]) if pd.notna(bal_vals[current_idx]) else 0.0
    except (TypeError, ValueError, IndexError):
        cur_bal = 0.0
    if cur_bal <= 0:
        return 0.0, 0.0, set(), 0
    remaining = cur_bal
    normal = 0.0
    overdue = 0.0
    pink_cols = set()
    overdue_months = 0
    for c in range(current_idx, -1, -1):
        if remaining <= 0:
            break
        try:
            sal = float(sales_vals[c]) if pd.notna(sales_vals[c]) else 0.0
        except (TypeError, ValueError):
            sal = 0.0
        sal = max(sal, 0.0)
        portion = min(sal, remaining) if sal > 0 else 0.0
        if portion > 0:
            age = current_idx - c
            if age <= credit_months:
                normal += portion
            else:
                overdue += portion
                pink_cols.add(c)
                overdue_months = max(overdue_months, age - credit_months)
            remaining -= portion
    if remaining > 0:
        overdue += remaining
        overdue_months = max(overdue_months, max(1, current_idx - credit_months))
    return normal, overdue, pink_cols, overdue_months
def format_debt_status_label(overdue_amt, overdue_months):
    """정상 / 연체 1~3개월 / 악성(연체 4개월+)."""
    if overdue_amt <= 0 or overdue_months <= 0:
        return "정상"
    if overdue_months >= 4:
        return "악성"
    return f"연체 {int(overdue_months)}개월"
def apply_debt_style_fast(df, highlight_debt=True, payment_terms_map=None):
    """분홍 = 연체분만. 익월말이면 전월 매출은 정상(분홍 X). 잔액0이면 분홍 없음."""
    styles = np.full(df.shape, '', dtype=object)
    terms_map = payment_terms_map or {}
    clients = df.index.get_level_values('거래처')
    gubuns = df.index.get_level_values('구분')
    u_clients_fast = clients.unique()
    color_map_fast = {
        client: DEBT_CLIENT_STRIPE_A if i % 2 == 0 else DEBT_CLIENT_STRIPE_B
        for i, client in enumerate(u_clients_fast)
    }
    for r in range(df.shape[0]):
        client = clients[r]
        if client == "📌 [전체 합계]":
            row_style = 'background-color: #E2E8F0; font-weight: 700;'
        else:
            row_style = f'background-color: {color_map_fast.get(client, DEBT_CLIENT_STRIPE_A)};'
        styles[r, :] = row_style
    def apply_pink_cell(old_style, pink_bg=DEBT_PINK_SOFT):
        s = re.sub(r'background-color:\s*#[0-9a-fA-F]+;', '', old_style)
        s = re.sub(r'color:\s*#[0-9a-fA-F]+;', '', s)
        s = re.sub(r'font-weight:\s*\d+;', '', s)
        return s + f' background-color: {pink_bg};'
    if df.shape[1] == 0:
        return pd.DataFrame(styles, index=df.index, columns=df.columns)
    current_month_idx = df.shape[1] - 1
    for client in u_clients_fast:
        if client == "📌 [전체 합계]":
            continue
        client_rows = np.where(clients == client)[0]
        i_sal = -1
        i_bal = -1
        for r in client_rows:
            if gubuns[r] == '매출':
                i_sal = r
            elif gubuns[r] == '잔액':
                i_bal = r
        if i_sal == -1 or i_bal == -1:
            continue
        term = resolve_payment_term(client, terms_map)
        credit = payment_term_credit_months(term)
        sales_vals = [df.iat[i_sal, c] for c in range(df.shape[1])]
        bal_vals = [df.iat[i_bal, c] for c in range(df.shape[1])]
        _normal, overdue, pink_cols, _od_m = analyze_client_ar(
            sales_vals, bal_vals, credit, current_month_idx
        )
        if overdue > 0:
            styles[i_bal, current_month_idx] = apply_pink_cell(
                styles[i_bal, current_month_idx], DEBT_PINK_SOFT
            )
            for c in pink_cols:
                styles[i_sal, c] = apply_pink_cell(styles[i_sal, c], DEBT_PINK_SOFT)
    return pd.DataFrame(styles, index=df.index, columns=df.columns)
def compute_debt_status_by_client(disp_debt, payment_terms_map=None):
    """거래처 → 채권구분 라벨 (정상 / 연체 1~3개월 / 악성)."""
    terms_map = payment_terms_map or {}
    if disp_debt is None or disp_debt.empty:
        return {}
    clients = disp_debt.index.get_level_values("거래처")
    gubuns = disp_debt.index.get_level_values("구분")
    current_idx = disp_debt.shape[1] - 1
    out = {}
    for client in clients.unique():
        if client == "📌 [전체 합계]":
            continue
        rows = np.where(clients == client)[0]
        i_sal = i_bal = -1
        for r in rows:
            if gubuns[r] == "매출":
                i_sal = r
            elif gubuns[r] == "잔액":
                i_bal = r
        if i_sal < 0 or i_bal < 0 or current_idx < 0:
            out[client] = "정상"
            continue
        term = resolve_payment_term(client, terms_map)
        credit = payment_term_credit_months(term)
        sales_vals = [disp_debt.iat[i_sal, c] for c in range(disp_debt.shape[1])]
        bal_vals = [disp_debt.iat[i_bal, c] for c in range(disp_debt.shape[1])]
        _n, overdue, _p, od_m = analyze_client_ar(sales_vals, bal_vals, credit, current_idx)
        out[client] = format_debt_status_label(overdue, od_m)
    return out
def _debt_label_cell_style(client, gubun, color_map, compact=False):
    """거래처분석 탭과 동일 계열: 13px / weight 400 · 중간정렬."""
    pad = "4px 4px" if compact else "6px 8px"
    base = (
        f"padding:{pad};border-bottom:1px solid #E2E8F0;white-space:nowrap;"
        "font-size:13px;font-weight:400;line-height:1.4;vertical-align:middle;"
        "font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;"
        "overflow:hidden;text-overflow:ellipsis;color:#31333F;"
    )
    if client == "📌 [전체 합계]":
        return base + "background-color:#E2E8F0;font-weight:600;text-align:center;"
    bg = color_map.get(client, "#FFFFFF")
    return base + f"background-color:{bg};text-align:center;"
def _debt_num_cell_font(compact=False):
    """거래처분석 탭 dataframe과 비슷한 숫자 글씨."""
    pad = "4px 4px" if compact else "6px 8px"
    return (
        f"padding:{pad};border-bottom:1px solid #E2E8F0;white-space:nowrap;"
        "font-size:13px;font-weight:400;line-height:1.4;vertical-align:middle;text-align:center;"
        "font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;"
        "color:#31333F;"
    )
PAYMENT_TERMS_PATH = os.path.join(CACHE_DIR, "payment_terms.csv")
PAYMENT_TERMS_FALLBACK = os.path.join(os.path.dirname(os.path.abspath(__file__)), "payment_terms.csv")
@st.cache_data(show_spinner=False)
def load_payment_terms_map():
    """거래처별 결제조건 로드 (입금기준표 반영)."""
    path = PAYMENT_TERMS_PATH if os.path.exists(PAYMENT_TERMS_PATH) else PAYMENT_TERMS_FALLBACK
    if not os.path.exists(path):
        return {}
    try:
        df = pd.read_csv(path, encoding="utf-8-sig")
    except Exception:
        try:
            df = pd.read_csv(path, encoding="cp949")
        except Exception:
            return {}
    if "거래처" not in df.columns or "결제조건" not in df.columns:
        return {}
    out = {}
    for _, row in df.iterrows():
        name = str(row["거래처"]).strip()
        term = "" if pd.isna(row["결제조건"]) else str(row["결제조건"]).strip()
        if name:
            out[name] = term
    return out
def resolve_payment_term(client_name, terms_map):
    """정확 매칭 → strip 매칭 → 접두(가스코아산/세진가스텍) 규칙."""
    if not client_name or client_name == "📌 [전체 합계]":
        return ""
    name = str(client_name)
    if name in terms_map:
        return terms_map[name]
    stripped = name.strip()
    if stripped in terms_map:
        return terms_map[stripped]
    for k, v in terms_map.items():
        if k.strip() == stripped:
            return v
    if name.startswith("가스코아산") or name.startswith("Z가스코아산"):
        return terms_map.get("가스코아산.", "익월말현금")
    if name.startswith("세진가스텍"):
        return "2개월"
    return ""
def render_interactive_html_table(
    headers,
    body_html,
    height=420,
    show_sum_popup=False,
    toolbar_hint=None,
    freeze_left_cols=0,
    freeze_left_widths=None,
    freeze_right_header=False,
    freeze_right_widths=None,
):
    hint = toolbar_hint or "셀 클릭 · ⌘/Ctrl+클릭 또는 ⊕다중선택 · 복사 버튼/⌘C"
    sum_controls = ""
    if show_sum_popup:
        sum_controls = """
            <div class="dash-sum-inline" id="dashSumInline">
                <span class="dash-sum-label">선택 합계</span>
                <span id="dashSumValue">0</span>
                <span class="dash-sum-meta" id="dashSumCount">0개</span>
            </div>
        """
    freeze_css = ""
    parts = [
        """
        thead th {
            position: sticky;
            top: 0;
            z-index: 5;
            background: #F0F2F6 !important;
            box-shadow: 0 1px 0 #CBD5E1;
        }
        """
    ]
    if freeze_left_cols > 0:
        widths = freeze_left_widths or ([150, 64] + [100] * max(0, freeze_left_cols - 2))
        left = 0
        for i in range(min(freeze_left_cols, len(widths))):
            w = widths[i]
            parts.append(
                f"""
        th.dash-freeze-{i}, td.dash-freeze-{i} {{
            position: sticky;
            left: {left}px;
            min-width: {w}px;
            max-width: {w}px;
            width: {w}px;
            box-shadow: 2px 0 0 #CBD5E1;
            overflow: hidden;
            text-overflow: ellipsis;
        }}
        th.dash-freeze-{i} {{
            top: 0;
            z-index: 8;
            background: #F0F2F6 !important;
        }}
        td.dash-freeze-{i} {{
            z-index: 4;
        }}
                """
            )
            left += w
    if freeze_right_header:
        # 우측 2열: r0=연체(right:0), r1=결제(right:r0폭)
        rw = freeze_right_widths or [64, 72]
        r0_w = int(rw[0]) if len(rw) > 0 else 64
        r1_w = int(rw[1]) if len(rw) > 1 else 72
        parts.append(
            f"""
        th.dash-freeze-r0, td.dash-freeze-r0 {{
            position: sticky;
            right: 0;
            z-index: 6;
            min-width: {r0_w}px;
            max-width: {r0_w}px;
            width: {r0_w}px;
            box-shadow: -1px 0 0 #E2E8F0;
            background-clip: padding-box;
            word-break: keep-all;
        }}
        th.dash-freeze-r1, td.dash-freeze-r1 {{
            position: sticky;
            right: {r0_w}px;
            z-index: 6;
            min-width: {r1_w}px;
            max-width: {r1_w}px;
            width: {r1_w}px;
            box-shadow: -1px 0 0 #E2E8F0;
            background-clip: padding-box;
            word-break: keep-all;
            overflow: hidden;
        }}
        th.dash-freeze-r0, th.dash-freeze-r1 {{
            top: 0;
            z-index: 9;
            background: #F0F2F6 !important;
            color: #31333F;
            font-weight: 600;
        }}
        td.dash-freeze-r0, td.dash-freeze-r1 {{
            z-index: 5;
        }}
            """
        )
    freeze_css = "".join(parts)
    header_cells = []
    n_h = len(headers)
    for i, h in enumerate(headers):
        classes = []
        if i < freeze_left_cols:
            classes.append(f"dash-freeze-{i}")
        if freeze_right_header and n_h >= 2:
            if i == n_h - 1:
                classes.append("dash-freeze-r0")  # 연체개월수
            elif i == n_h - 2:
                classes.append("dash-freeze-r1")  # 결제조건
        freeze_cls = f' class="{" ".join(classes)}"' if classes else ""
        header_cells.append(
            f'<th{freeze_cls} style="padding:6px 8px;border-bottom:1px solid #E2E8F0;'
            f'background:#F0F2F6;text-align:center;vertical-align:middle;'
            f'font-size:13px;font-weight:600;white-space:nowrap;color:#31333F;">'
            f'{html.escape(str(h))}</th>'
        )
    page_html = f"""
    <!DOCTYPE html>
    <html><head><meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <style>
        html, body {{
            margin: 0;
            height: 100%;
            overflow: hidden;
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
            font-size: 13px;
            font-weight: 400;
            color: #31333F;
        }}
        .dash-shell {{
            display: flex;
            flex-direction: column;
            height: 100%;
            min-height: 100%;
        }}
        .dash-table-toolbar {{
            padding: 6px 10px;
            font-size: 12px;
            color: #64748B;
            background: #F8FAFC;
            border: 1px solid #E2E8F0;
            border-bottom: none;
            border-radius: 4px 4px 0 0;
            display: flex;
            gap: 8px;
            align-items: center;
            flex-wrap: wrap;
            flex-shrink: 0;
        }}
        .dash-action-btn, .dash-multi-btn {{
            padding: 4px 10px;
            font-size: 12px;
            border: 1px solid #CBD5E1;
            border-radius: 4px;
            background: #fff;
            cursor: pointer;
            color: #334155;
        }}
        .dash-multi-btn.active {{ background: #DBEAFE; border-color: #2563EB; color: #1D4ED8; }}
        .dash-sum-inline {{
            display: none;
            margin-left: auto;
            align-items: center;
            gap: 8px;
            padding: 4px 10px;
            border-radius: 6px;
            background: #1E293B;
            color: #fff;
            font-size: 13px;
            white-space: nowrap;
        }}
        .dash-sum-inline.show {{ display: inline-flex; }}
        .dash-sum-label {{ opacity: 0.85; font-size: 12px; }}
        #dashSumValue {{ font-size: 15px; font-weight: 700; }}
        .dash-sum-meta {{ opacity: 0.75; font-size: 11px; }}
        .wrap {{
            flex: 1 1 auto;
            min-height: 0;
            overflow: auto;
            -webkit-overflow-scrolling: touch;
            overscroll-behavior: contain;
            touch-action: pan-x pan-y;
            border: 1px solid #E2E8F0;
            border-radius: 0 0 4px 4px;
            background: #fff;
        }}
        table {{ width: max-content; min-width: 100%; border-collapse: separate; border-spacing: 0; }}
        th, td {{ font-weight: 400; }}
        .dash-cell-selectable {{ cursor: cell; user-select: none; -webkit-user-select: none; }}
        .dash-cell-selectable.selected {{
            outline: 2px solid #2563EB;
            outline-offset: -2px;
            background-color: #DBEAFE !important;
        }}
        {freeze_css}
    </style></head><body>
    <div class="dash-shell">
    <div class="dash-table-toolbar">
        <span>{html.escape(hint)}</span>
        <button type="button" class="dash-multi-btn" id="dashMultiBtn">⊕ 다중선택</button>
        <button type="button" class="dash-action-btn" id="dashCopyBtn">📋 선택 복사</button>
        {sum_controls}
    </div>
    <div class="wrap">
        <table id="dashTable">
            <thead><tr>{''.join(header_cells)}</tr></thead>
            <tbody>{body_html}</tbody>
        </table>
    </div>
    </div>
    <script>
    (function() {{
        const showSum = {'true' if show_sum_popup else 'false'};
        const selected = new Set();
        let multiMode = false;
        const popup = document.getElementById('dashSumInline');
        const sumValue = document.getElementById('dashSumValue');
        const sumCount = document.getElementById('dashSumCount');
        function fmt(n) {{
            return Math.round(n).toLocaleString('ko-KR');
        }}
        function updateSumUI() {{
            if (!showSum) return;
            let total = 0, count = 0;
            selected.forEach(td => {{
                const raw = td.dataset.raw;
                if (raw === undefined || raw === '') return;
                const v = parseFloat(raw);
                if (!isNaN(v)) {{ total += v; count += 1; }}
            }});
            if (count > 0) {{
                if (sumValue) sumValue.textContent = fmt(total);
                if (sumCount) sumCount.textContent = count + '개 숫자 셀';
                if (popup) popup.classList.add('show');
            }} else {{
                if (popup) popup.classList.remove('show');
            }}
        }}
        document.querySelectorAll('.dash-cell-selectable').forEach(td => {{
            td.addEventListener('click', function(e) {{
                const additive = multiMode || e.ctrlKey || e.metaKey;
                if (!additive) {{
                    selected.forEach(el => el.classList.remove('selected'));
                    selected.clear();
                }}
                if (td.classList.contains('selected')) {{
                    td.classList.remove('selected');
                    selected.delete(td);
                }} else {{
                    td.classList.add('selected');
                    selected.add(td);
                }}
                updateSumUI();
            }});
        }});
        document.getElementById('dashMultiBtn').addEventListener('click', function() {{
            multiMode = !multiMode;
            this.classList.toggle('active', multiMode);
        }});
        function selectedText() {{
            const rows = new Map();
            selected.forEach(td => {{
                const tr = td.parentElement;
                if (!tr) return;
                const cells = Array.from(tr.querySelectorAll('td'));
                const idx = cells.indexOf(td);
                const key = Array.from(tr.parentElement.children).indexOf(tr);
                if (!rows.has(key)) rows.set(key, {{ order: key, parts: [] }});
                const label = (td.dataset.raw !== undefined && td.dataset.raw !== '')
                    ? td.dataset.raw
                    : (td.textContent || '').trim();
                rows.get(key).parts.push({{ idx: idx, text: label }});
            }});
            return Array.from(rows.values())
                .sort((a, b) => a.order - b.order)
                .map(r => r.parts.sort((a, b) => a.idx - b.idx).map(p => p.text).join('\\t'))
                .join('\\n');
        }}
        async function copySelected() {{
            const text = selectedText();
            const btn = document.getElementById('dashCopyBtn');
            if (!text) {{
                if (btn) {{
                    const prev = btn.textContent;
                    btn.textContent = '선택 없음';
                    setTimeout(() => {{ btn.textContent = prev; }}, 900);
                }}
                return;
            }}
            try {{
                if (navigator.clipboard && navigator.clipboard.writeText) {{
                    await navigator.clipboard.writeText(text);
                }} else {{
                    const ta = document.createElement('textarea');
                    ta.value = text;
                    ta.style.position = 'fixed';
                    ta.style.left = '-9999px';
                    document.body.appendChild(ta);
                    ta.select();
                    document.execCommand('copy');
                    document.body.removeChild(ta);
                }}
                if (btn) {{
                    const prev = btn.textContent;
                    btn.textContent = '✓ 복사됨';
                    setTimeout(() => {{ btn.textContent = prev; }}, 900);
                }}
            }} catch (err) {{
                if (btn) {{
                    const prev = btn.textContent;
                    btn.textContent = '복사 실패';
                    setTimeout(() => {{ btn.textContent = prev; }}, 900);
                }}
            }}
        }}
        const copyBtn = document.getElementById('dashCopyBtn');
        if (copyBtn) copyBtn.addEventListener('click', function(e) {{
            e.preventDefault();
            copySelected();
        }});
        document.addEventListener('keydown', function(e) {{
            if ((e.metaKey || e.ctrlKey) && (e.key === 'c' || e.key === 'C')) {{
                if (selected.size > 0) {{
                    e.preventDefault();
                    copySelected();
                }}
            }}
        }});
    }})();
    </script>
    </body></html>
    """
    components.html(page_html, height=height, scrolling=True)
def render_tab3_dataframe_table(
    df, fmt, target_col, key_prefix="tab3", table_kind="sales", sort_order=None
):
    """Tab3 표 — 연한 파스텔 그라디언트 + 선택 행 복사.
    sort_order: 단가표용 품목명 순서(매출액 내림차순 인덱스).
    """
    df_active, numeric_cols, highlight_col_name = cached_prepare_active_df(df, target_col)
    if df_active is None or df_active.empty:
        return
    annual_cols = [c for c in numeric_cols if "연간총합" in str(c)]
    month_cols = [c for c in numeric_cols if "연간총합" not in str(c)]
    # 정렬: 단가표는 매출액 순(sort_order), 그 외는 해당 표 고실적 순
    if table_kind == "price" and sort_order is not None:
        rank = {name: i for i, name in enumerate(sort_order)}
        df_active = df_active.assign(
            _rk=df_active["품목명"].map(lambda x: rank.get(x, 10**9))
        ).sort_values(by="_rk").drop(columns=["_rk"])
    else:
        sort_col = None
        if table_kind != "price":
            if annual_cols:
                sort_col = annual_cols[0]
            elif highlight_col_name in numeric_cols:
                sort_col = highlight_col_name
            elif month_cols:
                sort_col = month_cols[0]
        if sort_col and sort_col in df_active.columns:
            df_active = df_active.sort_values(by=sort_col, ascending=False, na_position="last")
        elif numeric_cols and table_kind != "price":
            df_active = df_active.assign(_rk=df_active[numeric_cols].sum(axis=1)).sort_values(
                by="_rk", ascending=False
            ).drop(columns=["_rk"])
    styled = df_active.style.format(fmt, subset=numeric_cols)
    # 파스텔 히트맵 — 매출/출고만 (적용단가는 그라디언트 없음)
    if table_kind == "price":
        if highlight_col_name and highlight_col_name in numeric_cols:
            styled = styled.apply(
                lambda s: ["color:#B91C1C;background-color:#FEE2E2;"] * len(s),
                subset=[highlight_col_name],
                axis=0,
            )
        styled = styled.apply(
            lambda col: (
                ["border-right:2px solid #CBD5E1;background-color:#FAFAFA;"] * len(col)
                if col.name == "품목명"
                else [""] * len(col)
            ),
            axis=0,
        )
    else:
        cmap_month = _TAB3_CMAP_BLUE if table_kind == "sales" else _TAB3_CMAP_GREEN
        if month_cols:
            styled = styled.background_gradient(
                cmap=cmap_month, subset=month_cols, axis=0
            )
            styled = styled.set_properties(subset=month_cols, **{"color": "#0F172A"})
        if annual_cols:
            styled = styled.background_gradient(
                cmap=_TAB3_CMAP_ORANGE, subset=annual_cols, axis=0
            )
            styled = styled.set_properties(
                subset=annual_cols,
                **{
                    "font-weight": "700",
                    "border-left": "2px solid #FDBA74",
                    "color": "#0F172A",
                },
            )
        if highlight_col_name and highlight_col_name in numeric_cols:
            styled = styled.apply(
                lambda s: [
                    "color:#9F1239;font-weight:700;background-color:#FFE4E6;"
                    "border-left:2px solid #FB7185;border-right:2px solid #FB7185;"
                ]
                * len(s),
                subset=[highlight_col_name],
                axis=0,
            )
        styled = styled.apply(
            lambda col: (
                [
                    "font-weight:700;background-color:#F8FAFC;color:#0F172A;"
                    "border-right:2px solid #94A3B8;"
                ]
                * len(col)
                if col.name == "품목명"
                else [""] * len(col)
            ),
            axis=0,
        )
    event = st.dataframe(
        styled,
        use_container_width=True,
        height=420,
        hide_index=True,
        key=f"{key_prefix}_grid",
        on_select="rerun",
        selection_mode="multi-row",
        column_config={
            "품목명": st.column_config.TextColumn("품목명", width="medium"),
        },
    )
    rows = []
    try:
        rows = list(event.selection.rows or [])
    except Exception:
        rows = []
    if rows:
        subset = df_active.iloc[rows].copy()
        for c in numeric_cols:
            if c in subset.columns:
                subset[c] = pd.to_numeric(subset[c], errors="coerce").map(
                    lambda x: "" if pd.isna(x) else format(x, ",.0f")
                )
        tsv = subset.to_csv(sep="\t", index=False)
        enc = urllib.parse.quote(tsv, safe="")
        components.html(
            f"""
            <div style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;">
              <button id="tab3CopyBtn" type="button"
                style="padding:6px 12px;border:1px solid #CBD5E1;border-radius:6px;
                       background:#fff;color:#334155;font-size:13px;font-weight:600;cursor:pointer;">
                📋 선택 행 복사 ({len(rows)}행)
              </button>
              <span id="tab3CopyMsg" style="margin-left:8px;font-size:12px;color:#64748B;"></span>
            </div>
            <script>
            (function() {{
              var btn = document.getElementById("tab3CopyBtn");
              var msg = document.getElementById("tab3CopyMsg");
              if (!btn) return;
              btn.addEventListener("click", async function() {{
                var text = decodeURIComponent("{enc}");
                try {{
                  if (navigator.clipboard && navigator.clipboard.writeText) {{
                    await navigator.clipboard.writeText(text);
                  }} else {{
                    var ta = document.createElement("textarea");
                    ta.value = text;
                    document.body.appendChild(ta);
                    ta.select();
                    document.execCommand("copy");
                    document.body.removeChild(ta);
                  }}
                  if (msg) msg.textContent = "✓ 복사됨 — 엑셀/메모장에 붙여넣기 하세요";
                  btn.textContent = "✓ 복사됨 ({len(rows)}행)";
                }} catch (e) {{
                  if (msg) msg.textContent = "복사 실패 — 아래 텍스트를 길게 눌러 복사하세요";
                }}
              }});
            }})();
            </script>
            """,
            height=46,
        )
    else:
        st.caption("행을 클릭해 선택한 뒤 「선택 행 복사」로 복사할 수 있습니다. (⌘/Ctrl 클릭으로 여러 행)")
def render_debt_interactive_table(disp_debt, highlight_debt, height=700, payment_terms_map=None):
    """채권관리 표 — 우측: 결제조건 + 연체개월수(정상/연체1~3개월/악성)."""
    terms_map = payment_terms_map or {}
    style_df = apply_debt_style_fast(
        disp_debt, highlight_debt=highlight_debt, payment_terms_map=terms_map
    )
    status_map = compute_debt_status_by_client(disp_debt, terms_map)
    numeric_cols = list(disp_debt.columns)
    clients = disp_debt.index.get_level_values("거래처")
    u_clients = clients.unique()
    color_map = {
        client: DEBT_CLIENT_STRIPE_A if i % 2 == 0 else DEBT_CLIENT_STRIPE_B
        for i, client in enumerate(u_clients)
    }
    long_cnt = sum(1 for v in status_map.values() if v == "악성")
    # 거래처분석 탭과 동일: 13px / 보통 굵기 · 중간정렬 (연체 분홍만 강조)
    compact = is_touch_ui()
    if compact:
        headers = ["거래처", "구분"] + numeric_cols + ["결제", "연체"]
        left_w, right_w = [108, 40], [52, 64]
        hint = "셀 클릭 · 분홍=연체 · 결제/연체 열 축소 · 연체: 정상/1~3M/악성(4M+)"
    else:
        headers = ["거래처", "구분"] + numeric_cols + ["결제조건", "연체개월수"]
        left_w, right_w = [160, 68], [110, 120]
        hint = "셀 클릭 선택 · 분홍=연체 · 연체개월수: 정상 / 연체 1~3개월 / 악성(4개월+)"
    cell_font = _debt_num_cell_font(compact=compact)
    body_rows = []
    for r, (idx, row) in enumerate(disp_debt.iterrows()):
        client, gubun = idx[0], idx[1]
        cells = [
            f'<td class="dash-cell-selectable dash-freeze-0" style="{_debt_label_cell_style(client, "", color_map, compact=compact)}" title="{html.escape(str(client))}">{html.escape(str(client))}</td>',
            f'<td class="dash-cell-selectable dash-freeze-1" style="{_debt_label_cell_style(client, gubun, color_map, compact=compact)}">{html.escape(str(gubun))}</td>',
        ]
        for col in numeric_cols:
            val = row[col]
            try:
                num = float(val) if pd.notna(val) else 0.0
            except (TypeError, ValueError):
                num = 0.0
            display = f"{num:,.0f}"
            extra_style = str(style_df.at[idx, col]) if col in style_df.columns else ""
            cells.append(
                f'<td class="dash-cell-selectable" style="{cell_font}{extra_style}" '
                f'data-raw="{num}">{display}</td>'
            )
        # 결제조건·연체는 잔액 행에만 1회 · 글자크기/굵기 동일 · 중간정렬
        if client == "📌 [전체 합계]":
            term_bg = "#E2E8F0"
            if gubun == "잔액":
                term = ""
                status = f"악성 {long_cnt}곳" if long_cnt else "—"
                if compact and long_cnt:
                    status = f"악성{long_cnt}"
            else:
                term, status = "", ""
        else:
            term_bg = color_map.get(client, DEBT_CLIENT_STRIPE_A)
            if gubun == "잔액":
                term = resolve_payment_term(client, terms_map)
                status = status_map.get(client, "정상")
            else:
                term, status = "", ""
        status_disp = status
        if compact and status.startswith("연체 ") and status.endswith("개월"):
            status_disp = status.replace("연체 ", "").replace("개월", "M")
        meta_style = f"{cell_font}background-color:{term_bg};white-space:normal;"
        term_show = term if gubun == "잔액" else ""
        term_fallback = "—" if (client != "📌 [전체 합계]" and gubun == "잔액" and not term) else ""
        cells.append(
            f'<td class="dash-freeze-r1" style="{meta_style}" title="{html.escape(term)}">'
            f"{html.escape(term_show or term_fallback)}</td>"
        )
        if status == "악성" or (isinstance(status, str) and status.startswith("악성")):
            st_color = "#BE123C"
        elif status == "정상":
            st_color = "#047857"
        elif "연체" in str(status) and "개월" in str(status):
            st_color = "#C2410C"
        else:
            st_color = "#31333F"
        cells.append(
            f'<td class="dash-freeze-r0" style="{meta_style}color:{st_color};" '
            f'title="{html.escape(status)}">{html.escape(status_disp)}</td>'
        )
        body_rows.append(f"<tr>{''.join(cells)}</tr>")
    render_interactive_html_table(
        headers,
        "".join(body_rows),
        height=height + 40,
        show_sum_popup=True,
        toolbar_hint=hint,
        freeze_left_cols=2,
        freeze_left_widths=left_w,
        freeze_right_header=True,
        freeze_right_widths=right_w,
    )
def build_debt_client_balance_matrix(debt_df, month_cols):
    """거래처 × 월 잔액 피벗 (잔액>0 업체만)."""
    if debt_df is None or debt_df.empty or not month_cols:
        return pd.DataFrame()
    bal = debt_df[debt_df["구분"] == "잔액"].copy()
    if bal.empty:
        return pd.DataFrame()
    cols = [c for c in month_cols if c in bal.columns]
    if not cols:
        return pd.DataFrame()
    mat = bal.groupby("거래처", as_index=True)[cols].sum()
    mat = mat[(mat.fillna(0) > 0).any(axis=1)]
    return mat
def compute_debt_status_from_raw(debt_df, month_cols, payment_terms_map=None):
    """원본 채권 df → 거래처별 연체 라벨 (정상 포함)."""
    meta = compute_debt_od_meta_from_raw(debt_df, month_cols, payment_terms_map)
    return {c: v["label"] for c, v in meta.items()}
def compute_debt_od_meta_from_raw(debt_df, month_cols, payment_terms_map=None):
    """거래처 → {label, months, overdue_amt, pink_months, od_month_amts, cur_bal}."""
    terms_map = payment_terms_map or {}
    if debt_df is None or debt_df.empty or not month_cols:
        return {}
    cols = [c for c in month_cols if c in debt_df.columns]
    if not cols:
        return {}
    out = {}
    for client, g in debt_df.groupby("거래처"):
        sal = g[g["구분"] == "매출"]
        bal = g[g["구분"] == "잔액"]
        if sal.empty or bal.empty:
            out[client] = {
                "label": "정상",
                "months": 0,
                "overdue_amt": 0.0,
                "pink_months": [],
                "od_month_amts": {},
                "cur_bal": 0.0,
            }
            continue
        sales_vals = [float(sal[c].sum()) if c in sal.columns else 0.0 for c in cols]
        bal_vals = [float(bal[c].sum()) if c in bal.columns else 0.0 for c in cols]
        credit = payment_term_credit_months(resolve_payment_term(client, terms_map))
        _n, overdue, pink_cols, od_m = analyze_client_ar(
            sales_vals, bal_vals, credit, len(cols) - 1
        )
        pink_months = [cols[i] for i in sorted(pink_cols)]
        od_month_amts = {
            cols[i]: float(sales_vals[i]) for i in pink_cols if sales_vals[i] > 0
        }
        try:
            cur_bal = float(bal_vals[-1]) if pd.notna(bal_vals[-1]) else 0.0
        except (TypeError, ValueError, IndexError):
            cur_bal = 0.0
        out[client] = {
            "label": format_debt_status_label(overdue, od_m),
            "months": int(od_m) if overdue > 0 else 0,
            "overdue_amt": float(overdue),
            "pink_months": pink_months,
            "od_month_amts": od_month_amts,
            "cur_bal": max(0.0, cur_bal),
        }
    return out
def _debt_od_filter_categories(status_map):
    """데이터에 존재하는 연체 필터 버튼 목록 (정상 제외)."""
    labels = {v for v in status_map.values() if v and v != "정상"}
    ordered = []
    for n in range(1, 12):
        lab = f"연체 {n}개월"
        if lab in labels:
            ordered.append(lab)
    if "악성" in labels:
        ordered.append("악성")
    for lab in sorted(labels):
        if lab not in ordered:
            ordered.append(lab)
    return ordered
def render_debt_month_rank_panel(
    debt_df, month_cols, payment_terms_map=None, height=480, status_month_cols=None
):
    """화면 절반: 연체업체만 · 연체개월수 다중필터 · 채권액 내림차순."""
    terms_map = payment_terms_map or {}
    mat = build_debt_client_balance_matrix(debt_df, month_cols)
    if mat.empty:
        st.info("표시할 잔액 채권 데이터가 없습니다.")
        return
    cols = list(mat.columns)
    sort_m = cols[-1]
    status_cols = status_month_cols or cols
    od_meta = compute_debt_od_meta_from_raw(debt_df, status_cols, terms_map)
    status_map = {c: v["label"] for c, v in od_meta.items()}
    # 정상채권 제외
    overdue_idx = [c for c in mat.index if status_map.get(c, "정상") != "정상"]
    mat = mat.loc[overdue_idx] if overdue_idx else mat.iloc[0:0]
    if mat.empty:
        st.info("연체 거래처가 없습니다. (정상채권은 이 표에서 제외됩니다)")
        return
    filter_cats = _debt_od_filter_categories(
        {c: status_map[c] for c in mat.index if c in status_map}
    )
    if not filter_cats:
        st.info("연체개월수 필터 대상이 없습니다.")
        return
    if "debt_od_filters" not in st.session_state:
        st.session_state["debt_od_filters"] = list(filter_cats)
    else:
        kept = [x for x in st.session_state["debt_od_filters"] if x in filter_cats]
        st.session_state["debt_od_filters"] = kept if kept else list(filter_cats)
    st.markdown(
        "<div style='font-size:14px;font-weight:600;color:#334155;margin:18px 0 8px;'>"
        "📊 연체개월수 기준 채권 현황 <span style='font-size:12px;font-weight:500;color:#64748B;'>"
        "(거래처 선택 무시·전체 · 담당자 필터 적용 · 정상 제외 · 연체개월 다중선택 · 화면 절반)</span></div>",
        unsafe_allow_html=True,
    )
    # 연체개월수 / 악성 다중 토글 (아담한 버튼 — 이 컨테이너에만 compact CSS)
    with st.container(key="debt_compact_btns_od"):
        bm = st.columns(len(filter_cats), gap="small")
        for i, cat in enumerate(filter_cats):
            with bm[i]:
                on = cat in st.session_state["debt_od_filters"]
                if st.button(
                    cat,
                    key=f"debt_od_filt_{cat}",
                    type="primary" if on else "secondary",
                    width="stretch",
                ):
                    cur = list(st.session_state["debt_od_filters"])
                    if on:
                        if len(cur) > 1:
                            cur = [x for x in cur if x != cat]
                    else:
                        cur.append(cat)
                        cur = [x for x in filter_cats if x in cur]
                    st.session_state["debt_od_filters"] = cur
                    st.rerun()
    selected = st.session_state["debt_od_filters"]
    mat_f = mat.loc[[c for c in mat.index if status_map.get(c) in selected]]
    if mat_f.empty:
        st.info("선택한 연체개월수에 해당하는 업체가 없습니다. 버튼을 하나 이상 켜 주세요.")
        return
    # 연체 있는 월만 열로 표시 (거래처별 해당 없으면 — / 상단 월토글과 무관)
    display_months = [
        m
        for m in status_cols
        if any(m in (od_meta.get(c, {}).get("pink_months") or []) for c in mat_f.index)
    ]
    # 연체합계 기준 내림차순 (동률이면 당월 잔액)
    clients_sorted = sorted(
        list(mat_f.index),
        key=lambda c: (
            float(od_meta.get(c, {}).get("overdue_amt") or 0.0),
            float(od_meta.get(c, {}).get("cur_bal") or 0.0),
        ),
        reverse=True,
    )
    left, right = st.columns([1, 1], gap="medium")
    with left:
        compact = is_touch_ui()
        bal_hdr = f"{sort_m}잔액"
        if compact:
            headers = ["거래처"] + display_months + ["연체합계", bal_hdr, "결제", "연체"]
            left_w, right_w = [100], [48, 60]
        else:
            headers = ["거래처"] + display_months + ["연체합계", bal_hdr, "결제조건", "연체개월수"]
            left_w, right_w = [140], [110, 120]
        cell = _debt_num_cell_font(compact=compact)
        body = []
        for i, client in enumerate(clients_sorted):
            bg = DEBT_CLIENT_STRIPE_A if i % 2 == 0 else DEBT_CLIENT_STRIPE_B
            meta = od_meta.get(client, {})
            od_amts = meta.get("od_month_amts") or {}
            pink_set = set(meta.get("pink_months") or [])
            tds = [
                f'<td class="dash-freeze-0" style="{cell}background:{bg};'
                f'overflow:hidden;text-overflow:ellipsis;" '
                f'title="{html.escape(str(client))}">{html.escape(str(client))}</td>'
            ]
            for m in display_months:
                if m in pink_set and float(od_amts.get(m, 0) or 0) > 0:
                    v = float(od_amts[m])
                    tds.append(
                        f'<td class="dash-cell-selectable" style="{cell}background:{DEBT_PINK_SOFT};" '
                        f'data-raw="{v}">{v:,.0f}</td>'
                    )
                else:
                    tds.append(
                        f'<td class="dash-cell-selectable" style="{cell}background:{bg};color:#94A3B8;" '
                        f'data-raw="">—</td>'
                    )
            od_total = float(meta.get("overdue_amt") or 0.0)
            cur_bal = float(meta.get("cur_bal") or 0.0)
            tds.append(
                f'<td class="dash-cell-selectable" style="{cell}background:{bg};font-weight:600;" '
                f'data-raw="{od_total}">{od_total:,.0f}</td>'
            )
            tds.append(
                f'<td class="dash-cell-selectable" style="{cell}background:{bg};" '
                f'data-raw="{cur_bal}">{cur_bal:,.0f}</td>'
            )
            term = resolve_payment_term(client, terms_map) or "—"
            status = status_map.get(client, "악성")
            status_disp = status
            if compact and status.startswith("연체 ") and status.endswith("개월"):
                status_disp = status.replace("연체 ", "").replace("개월", "M")
            if status == "악성":
                sc = "#BE123C"
            elif "연체" in str(status) and "개월" in str(status):
                sc = "#C2410C"
            else:
                sc = "#31333F"
            tds.append(
                f'<td class="dash-freeze-r1" style="{cell}background-color:{bg};white-space:normal;" '
                f'title="{html.escape(term)}">{html.escape(term)}</td>'
            )
            tds.append(
                f'<td class="dash-freeze-r0" style="{cell}background-color:{bg};color:{sc};" '
                f'title="{html.escape(status)}">{html.escape(status_disp)}</td>'
            )
            body.append(f"<tr>{''.join(tds)}</tr>")
        sel_txt = ", ".join(selected)
        od_m_txt = ", ".join(display_months) if display_months else "없음"
        render_interactive_html_table(
            headers,
            "".join(body),
            height=height,
            show_sum_popup=True,
            toolbar_hint=f"필터: {sel_txt} · 연체월만({od_m_txt}) · 연체합계↓ · 정상 제외",
            freeze_left_cols=1,
            freeze_left_widths=left_w,
            freeze_right_header=True,
            freeze_right_widths=right_w,
        )
    with right:
        top_series = pd.Series(
            {
                c: float(od_meta.get(c, {}).get("overdue_amt") or 0.0)
                for c in clients_sorted
            }
        )
        top_n = top_series[top_series > 0].head(15)
        if top_n.empty:
            st.info("선택한 연체 구간에 연체금액이 있는 거래처가 없습니다.")
        else:
            fig = go.Figure(
                go.Bar(
                    x=top_n.values,
                    y=[str(i) for i in top_n.index],
                    orientation="h",
                    marker_color="#FCA5A5",
                    text=[f"{v:,.0f}" for v in top_n.values],
                    textposition="outside",
                )
            )
            fig.update_layout(
                title=dict(
                    text=f"연체금액 Top {len(top_n)}",
                    font=dict(size=14, color="#334155"),
                ),
                yaxis=dict(autorange="reversed", tickfont=dict(size=11)),
                xaxis=dict(title="원", tickfont=dict(size=11)),
                margin=dict(l=10, r=40, t=40, b=30),
                height=height,
                plot_bgcolor="#FFFFFF",
                paper_bgcolor="#FFFFFF",
                showlegend=False,
            )
            chart_key = "debt_od_" + "_".join(selected)[:40]
            render_plotly_chart(fig, use_container_width=True, key=chart_key)
        # 선택 구간별 업체 수·연체합계
        rows_sum = []
        for cat in filter_cats:
            clients_c = [c for c in mat.index if status_map.get(c) == cat]
            if not clients_c:
                continue
            amt = float(sum(od_meta.get(c, {}).get("overdue_amt") or 0.0 for c in clients_c))
            on = "ON" if cat in selected else "OFF"
            rows_sum.append({"구분": cat, "업체수": len(clients_c), "연체합계": amt, "선택": on})
        if rows_sum:
            st.markdown(
                "<div style='font-size:12px;font-weight:600;color:#64748B;margin:8px 0 4px;'>"
                "연체개월수별 요약</div>",
                unsafe_allow_html=True,
            )
            sum_df = pd.DataFrame(rows_sum)
            st.dataframe(
                sum_df.style.format({"연체합계": "{:,.0f}"}),
                use_container_width=True,
                hide_index=True,
                height=min(280, 48 + 32 * len(sum_df)),
            )
@st.cache_data
def cached_staff_pivot(df_base, desired_order):
    if df_base.empty:
        return pd.DataFrame()
    staff_raw = (df_base.pivot_table(index="담당자", columns="연도월_정렬", values="매출액", aggfunc="sum").fillna(0) * 1.1 / 10000)
    staff_cols = [c for c in desired_order if c in staff_raw.columns]
    
    df_p = staff_raw.reindex(columns=staff_cols, fill_value=0)
    
    row_totals = df_p.sum(axis=1)
    total_all = row_totals.sum()
    
    if total_all > 0:
        prop = row_totals / total_all * 100
    else:
        prop = 0.0
        
    df_p.insert(0, "총 매출 합계 (만원)", row_totals)
    df_p.insert(0, "매출 비중 (%)", prop)
    
    df_p = df_p.sort_values(by="총 매출 합계 (만원)", ascending=False)
    return df_p
@st.cache_data
def cached_ranking_pivot(df_base, current_year, sel_staff, all_months):
    df_ranking = df_base[(df_base["담당자"] == sel_staff) & (df_base["연도"] == current_year)]
    if df_ranking.empty:
        return pd.DataFrame()
    
    ranking_pivot = (df_ranking.pivot_table(index="거래처", columns="월", values="매출액", aggfunc="sum").fillna(0) * 1.1 / 10000)
    ranking_pivot = ranking_pivot.reindex(columns=all_months, fill_value=0)
    ranking_pivot["당해 누적 (만원)"] = ranking_pivot.sum(axis=1)
    ranking_pivot = ranking_pivot.sort_values(by="당해 누적 (만원)", ascending=False)
    return ranking_pivot
# ----------------------------------------------------
# 탭 전체 적용 업데이트 뱃지 렌더링 유틸
# ----------------------------------------------------
def render_update_badge(date_str):
    return f"<div style='text-align: right; margin-top: 0;'><span style='background-color: #FFFFFF; color: #475569; padding: 6px 12px; border-radius: 8px; font-size: 13px; font-weight: 600; border: 1px solid #CBD5E1; box-shadow: 0 1px 2px rgba(0,0,0,0.05);'>⏱️ 데이터 업데이트: {date_str}</span></div>"
def inject_sticky_tabs_script():
    """필터+탭 상단 fixed.
    - 맥: 현재 안정 코드 무손실 유지
    - iPad: 0804-최종 DOM 해킹 (fixed + 3.65rem + spacer + RAF)
    """
    components.html(
        """
        <script>
        (function() {
            var parentDoc = window.parent.document;
            var parentWin = window.parent;
            var SPACER_ID = 'dashboard-sticky-spacer';
            var SHIELD_ID = 'dashboard-top-shield';
            var syncTimer = null;
            var lastH = 0;
            if (parentWin.__dashboardStickyObserver) {
                parentWin.__dashboardStickyObserver.disconnect();
            }
            if (parentWin.__dashboardStickyBootInterval) {
                clearInterval(parentWin.__dashboardStickyBootInterval);
            }
            if (parentWin.__dashboardStickyTouchInterval) {
                clearInterval(parentWin.__dashboardStickyTouchInterval);
            }
            parentWin.__dashboardStickyRafLoop = false;
            if (parentWin.__dashboardIpadRaf) {
                parentWin.cancelAnimationFrame(parentWin.__dashboardIpadRaf);
                parentWin.__dashboardIpadRaf = null;
            }
            function isTouchPad() {
                var ua = navigator.userAgent || '';
                var ios = /iPad|iPhone|iPod/.test(ua);
                var ipadOs = navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1;
                return ios || ipadOs;
            }
            var touchMode = isTouchPad();
            /* ===== Mac 전용 (무손실) ===== */
            function getTopOffsetMac() {
                var header = parentDoc.querySelector('[data-testid="stHeader"]');
                if (!header) return 46;
                var rect = header.getBoundingClientRect();
                if (rect.bottom > 0) {
                    return Math.round(rect.bottom);
                }
                if (header.offsetHeight > 0) {
                    return header.offsetHeight;
                }
                return 46;
            }
            function getMainRect() {
                var block = parentDoc.querySelector('section.main .block-container');
                return block ? block.getBoundingClientRect() : null;
            }
            function findMainTabList() {
                var lists = parentDoc.querySelectorAll('div[role="tablist"]');
                for (var i = 0; i < lists.length; i++) {
                    if (lists[i].textContent.indexOf('📌 영업 종합 요약') !== -1) {
                        return lists[i];
                    }
                }
                return null;
            }
            function findFilterBox() {
                var marker = parentDoc.getElementById('sticky-marker');
                if (!marker) return null;
                return marker.closest('div[data-testid="stVerticalBlockBorderWrapper"]') ||
                       marker.closest('div[data-testid="stVerticalBlock"]');
            }
            function findMainTabsHost() {
                var hosts = parentDoc.querySelectorAll('div[data-testid="stTabs"]');
                for (var i = 0; i < hosts.length; i++) {
                    if (hosts[i].querySelector('[role="tabpanel"]')) return hosts[i];
                }
                return null;
            }
            function containsTabPanel(el) {
                if (!el || el.nodeType !== 1) return false;
                if (el.getAttribute && el.getAttribute('role') === 'tabpanel') return true;
                return !!(el.querySelector && el.querySelector('[role="tabpanel"]'));
            }
            function ensureSpacer(filterBox) {
                var spacer = parentDoc.getElementById(SPACER_ID);
                if (!spacer) {
                    spacer = parentDoc.createElement('div');
                    spacer.id = SPACER_ID;
                    filterBox.parentNode.insertBefore(spacer, filterBox);
                } else if (spacer.nextElementSibling !== filterBox) {
                    filterBox.parentNode.insertBefore(spacer, filterBox);
                }
                return spacer;
            }
            function syncTopShield(top, rect) {
                var shield = parentDoc.getElementById(SHIELD_ID);
                if (!shield) {
                    shield = parentDoc.createElement('div');
                    shield.id = SHIELD_ID;
                    parentDoc.body.appendChild(shield);
                }
                shield.style.setProperty('display', 'block', 'important');
                shield.style.setProperty('position', 'fixed', 'important');
                shield.style.setProperty('top', '0', 'important');
                shield.style.setProperty('left', rect.left + 'px', 'important');
                shield.style.setProperty('width', rect.width + 'px', 'important');
                shield.style.setProperty('height', top + 'px', 'important');
                shield.style.setProperty('background', '#F8FAFC', 'important');
                shield.style.setProperty('z-index', '989', 'important');
                shield.style.setProperty('pointer-events', 'none', 'important');
            }
            function mountTabs(filterBox, tabList) {
                if (!filterBox || !tabList) return false;
                filterBox.classList.add('dashboard-filter-sticky');
                if (touchMode) {
                    filterBox.classList.add('dashboard-filter-sticky-touch');
                } else {
                    filterBox.classList.remove('dashboard-filter-sticky-touch');
                }
                if (!filterBox.contains(tabList)) {
                    filterBox.appendChild(tabList);
                }
                tabList.classList.add('dashboard-tabs-in-filter');
                var tabsHost = findMainTabsHost();
                if (tabsHost) {
                    tabsHost.classList.add('dashboard-tabs-host-compact');
                    Array.from(tabsHost.children).forEach(function(child) {
                        if (!containsTabPanel(child)) {
                            child.classList.add('dashboard-tabs-list-shell');
                            child.style.setProperty('display', 'none', 'important');
                            child.style.setProperty('height', '0', 'important');
                        }
                    });
                }
                return true;
            }
            /* ===== iPad: 0804-최종 해킹 ===== */
            function isSidebarOpen() {
                var sidebar = parentDoc.querySelector('[data-testid="stSidebar"]');
                if (!sidebar) return false;
                if (sidebar.getAttribute('aria-expanded') === 'true') return true;
                return sidebar.getBoundingClientRect().width > 100;
            }
            function cleanupIpadPortal() {
                var portal = parentDoc.getElementById('dashboard-ipad-portal');
                if (!portal) return;
                var filterBox = findFilterBox();
                if (filterBox && portal.contains(filterBox)) {
                    var spacer = parentDoc.getElementById(SPACER_ID);
                    if (spacer && spacer.parentNode) {
                        spacer.parentNode.insertBefore(filterBox, spacer.nextSibling);
                    } else {
                        parentDoc.body.appendChild(filterBox);
                    }
                }
                portal.remove();
            }
            function applyIpad0804Hack() {
                cleanupIpadPortal();
                var marker = parentDoc.getElementById('sticky-marker');
                if (!marker) return;
                var targetBox = marker.closest('div[data-testid="stVerticalBlockBorderWrapper"]') ||
                                marker.closest('div[data-testid="stVerticalBlock"]');
                if (!targetBox) return;
                var tabList = findMainTabList();
                var stTabs = findMainTabsHost();
                var tabHeader = tabList || (stTabs ? stTabs.querySelector('div:first-child') : null);
                if (!tabHeader) return;
                if (!targetBox.contains(tabHeader)) {
                    targetBox.appendChild(tabHeader);
                }
                targetBox.classList.add('dashboard-filter-sticky');
                targetBox.classList.add('dashboard-filter-sticky-touch');
                tabHeader.classList.add('dashboard-tabs-in-filter');
                if (stTabs) {
                    stTabs.classList.add('dashboard-tabs-host-compact');
                    Array.from(stTabs.children).forEach(function(child) {
                        if (!containsTabPanel(child)) {
                            child.classList.add('dashboard-tabs-list-shell');
                            child.style.setProperty('display', 'none', 'important');
                            child.style.setProperty('height', '0', 'important');
                        }
                    });
                }
                /* iPad: Share/메뉴·≫ 버튼 아래 + 상단 파란선이 보이도록 내림 */
                var open = isSidebarOpen();
                targetBox.style.setProperty('position', 'fixed', 'important');
                targetBox.style.setProperty('top', '3.65rem', 'important');
                targetBox.style.setProperty('z-index', open ? '1' : '999999', 'important');
                targetBox.style.setProperty('background-color', '#FFFFFF', 'important');
                targetBox.style.setProperty('border', '2px solid #2563EB', 'important');
                targetBox.style.setProperty('border-top', '3px solid #2563EB', 'important');
                targetBox.style.setProperty('border-radius', '0 0 8px 8px', 'important');
                targetBox.style.setProperty('padding', '8px 10px 0px 10px', 'important');
                targetBox.style.setProperty('margin-top', '0', 'important');
                targetBox.style.setProperty('box-shadow', '0 8px 14px -4px rgba(37, 99, 235, 0.18)', 'important');
                targetBox.style.setProperty('-webkit-transform', 'none', 'important');
                targetBox.style.setProperty('transform', 'none', 'important');
                tabHeader.style.setProperty('padding', '0 10px 0px 10px', 'important');
                tabHeader.style.setProperty('margin-top', '5px', 'important');
                tabHeader.style.setProperty('border-bottom', 'none', 'important');
                tabHeader.style.setProperty('background-color', 'transparent', 'important');
                var spacer = parentDoc.getElementById(SPACER_ID);
                if (!spacer) {
                    spacer = parentDoc.createElement('div');
                    spacer.id = SPACER_ID;
                }
                if (spacer.parentNode !== targetBox.parentNode || spacer.nextElementSibling !== targetBox) {
                    targetBox.parentNode.insertBefore(spacer, targetBox);
                }
                spacer.style.height = targetBox.offsetHeight + 'px';
                spacer.style.width = '100%';
                spacer.style.marginBottom = '12px';
                spacer.style.display = 'block';
                parentWin.__dashboardIpadTarget = targetBox;
                parentWin.__dashboardIpadSpacer = spacer;
            }
            function syncIpadWidthLoop() {
                var targetBox = parentWin.__dashboardIpadTarget;
                var spacer = parentWin.__dashboardIpadSpacer || parentDoc.getElementById(SPACER_ID);
                if (!targetBox || !parentDoc.body.contains(targetBox)) {
                    applyIpad0804Hack();
                    targetBox = parentWin.__dashboardIpadTarget;
                    spacer = parentWin.__dashboardIpadSpacer;
                }
                if (targetBox && spacer && parentDoc.body.contains(spacer)) {
                    var rect = spacer.getBoundingClientRect();
                    targetBox.style.setProperty('position', 'fixed', 'important');
                    targetBox.style.setProperty('top', '3.65rem', 'important');
                    targetBox.style.setProperty('width', rect.width + 'px', 'important');
                    targetBox.style.setProperty('left', rect.left + 'px', 'important');
                    targetBox.style.setProperty('z-index', isSidebarOpen() ? '1' : '999999', 'important');
                    spacer.style.height = targetBox.offsetHeight + 'px';
                }
                parentWin.__dashboardIpadRaf = parentWin.requestAnimationFrame(syncIpadWidthLoop);
            }
            function syncFixedBar() {
                if (touchMode) {
                    applyIpad0804Hack();
                    return;
                }
                /* ===== Mac 무손실 분기 (변경 금지) ===== */
                var filterBox = findFilterBox();
                var tabList = findMainTabList();
                if (!filterBox || !tabList) return;
                mountTabs(filterBox, tabList);
                var rectMac = getMainRect();
                if (!rectMac) return;
                var topMac = getTopOffsetMac();
                filterBox.style.setProperty('position', 'fixed', 'important');
                filterBox.style.setProperty('top', topMac + 'px', 'important');
                filterBox.style.setProperty('left', rectMac.left + 'px', 'important');
                filterBox.style.setProperty('width', rectMac.width + 'px', 'important');
                filterBox.style.setProperty('max-width', rectMac.width + 'px', 'important');
                filterBox.style.setProperty('z-index', '990', 'important');
                parentDoc.documentElement.style.setProperty('--dashboard-bar-top', topMac + 'px');
                parentDoc.documentElement.style.setProperty('--dashboard-bar-left', rectMac.left + 'px');
                parentDoc.documentElement.style.setProperty('--dashboard-bar-width', rectMac.width + 'px');
                syncTopShield(topMac, rectMac);
                var barHMac = filterBox.offsetHeight + 4;
                if (Math.abs(barHMac - lastH) > 1) {
                    var spacerMac = ensureSpacer(filterBox);
                    spacerMac.style.height = barHMac + 'px';
                    spacerMac.style.display = 'block';
                    parentDoc.documentElement.style.setProperty('--dashboard-fixed-bar-height', barHMac + 'px');
                    lastH = barHMac;
                }
            }
            function scheduleSync(ms) {
                if (syncTimer) clearTimeout(syncTimer);
                syncTimer = setTimeout(syncFixedBar, ms || 40);
            }
            if (touchMode) {
                parentDoc.documentElement.classList.add('dashboard-touch-mode');
                /* iPad UI 분기 플래그 — 맥 분기에는 진입하지 않음 */
                try {
                    if (!parentWin.__dashboardTouchUiSynced) {
                        parentWin.__dashboardTouchUiSynced = true;
                        var tu = new URL(parentWin.location.href);
                        if (tu.searchParams.get('touch_ui') !== '1') {
                            tu.searchParams.set('touch_ui', '1');
                            parentWin.location.replace(tu.toString());
                            return;
                        }
                    }
                } catch (eTu) {}
                var boot = 0;
                parentWin.__dashboardStickyBootInterval = setInterval(function() {
                    applyIpad0804Hack();
                    boot++;
                    if (boot > 40) {
                        clearInterval(parentWin.__dashboardStickyBootInterval);
                        parentWin.__dashboardStickyBootInterval = null;
                    }
                }, 200);
                parentWin.__dashboardStickyTouchInterval = setInterval(function() {
                    applyIpad0804Hack();
                }, 800);
                var observer = new MutationObserver(function() { scheduleSync(80); });
                observer.observe(parentDoc.body, { childList: true, subtree: true });
                parentWin.__dashboardStickyObserver = observer;
                parentDoc.addEventListener('click', function(e) {
                    if (e.target.closest('[data-testid="collapsedControl"]') ||
                        e.target.closest('[data-testid="stSidebar"]') ||
                        e.target.closest('[role="tab"]')) {
                        scheduleSync(30);
                        scheduleSync(300);
                    }
                }, true);
                parentWin.addEventListener('resize', function() { scheduleSync(80); }, { passive: true });
                parentWin.addEventListener('orientationchange', function() {
                    scheduleSync(120);
                    scheduleSync(450);
                }, { passive: true });
                applyIpad0804Hack();
                syncIpadWidthLoop();
            } else {
                var pollCount = 0;
                parentWin.__dashboardStickyBootInterval = setInterval(function() {
                    syncFixedBar();
                    pollCount++;
                    if (pollCount > 100) {
                        clearInterval(parentWin.__dashboardStickyBootInterval);
                        parentWin.__dashboardStickyBootInterval = null;
                    }
                }, 150);
                var observer = new MutationObserver(function() { scheduleSync(60); });
                observer.observe(parentDoc.body, {
                    childList: true,
                    subtree: true,
                    attributes: true
                });
                parentWin.__dashboardStickyObserver = observer;
                var scrollRoot = parentDoc.querySelector('[data-testid="stAppViewContainer"]') || parentDoc;
                scrollRoot.addEventListener('scroll', function() { scheduleSync(30); }, { passive: true, capture: true });
                parentDoc.addEventListener('scroll', function() { scheduleSync(30); }, { passive: true, capture: true });
                parentWin.addEventListener('scroll', function() { scheduleSync(30); }, { passive: true });
                parentWin.addEventListener('resize', function() { scheduleSync(80); }, { passive: true });
                parentWin.addEventListener('pageshow', function() { scheduleSync(80); }, { passive: true });
                if (parentWin.visualViewport) {
                    parentWin.visualViewport.addEventListener('resize', function() { scheduleSync(40); }, { passive: true });
                    parentWin.visualViewport.addEventListener('scroll', function() { scheduleSync(40); }, { passive: true });
                }
                parentDoc.addEventListener('click', function(e) {
                    if (e.target.closest('[data-testid="collapsedControl"]') ||
                        e.target.closest('[role="tab"]')) {
                        scheduleSync(50);
                        scheduleSync(300);
                    }
                }, true);
                var sidebar = parentDoc.querySelector('[data-testid="stSidebar"]');
                if (sidebar) {
                    sidebar.addEventListener('transitionend', function() { scheduleSync(40); });
                }
                parentWin.__dashboardStickyTouchInterval = setInterval(function() {
                    syncFixedBar();
                }, 3000);
                syncFixedBar();
            }
        })();
        </script>
        """,
        height=0,
        width=0,
    )
# ----------------------------------------------------
# 날짜 유지용 로컬 캐시 함수 (7, 8번 탭 전용)
# ----------------------------------------------------
TAB7_DATE_FILE = os.path.join(CACHE_DIR, "tab7_date.txt")
TAB8_DATE_FILE = os.path.join(CACHE_DIR, "tab8_date.txt")
PROFIT_CACHE_FILE = os.path.join(CACHE_DIR, "profitability.json")
# 수익성분석.xlsx 기본값 (한국메티슨 시트 함수 그대로 이식)
PROFIT_DEFAULTS = {
    "project_name": "한국메티슨",
    "tank_gas": "질소",
    "tank_capacity_mode": "liters",  # liters | kg
    "tank_liters": 123763.0,  # 내용적 L (100000kg ÷ 0.808 ≈)
    "tank_spec": 100000.0,  # kg (환산값 또는 직접입력)
    "hourly_usage_mode": "kg",  # nm3 | kg
    "hourly_usage_kg": 50.0,  # 시간당 사용량 kg/h
    "hourly_usage_nm3": 39.98,  # 시간당 루베 Nm³/h (질소 50kg×0.7996)
    "operating_hours": 8.0,   # 일 가동시간 h
    "operating_days_per_month": 22.0,  # 월 가동일수 (일사용×가동일 = 월사용)
    "auto_monthly_from_cycle": False,  # 사용주기→월평균 공급량 자동반영
    "tank_price": 350_000_000.0,
    "monthly_usage_kg": 400_000.0,
    "vaporizer_capacity": 1500.0,
    "vaporizer_qty_note": "* 2ea",
    "vaporizer_price": 31_500_000.0,
    "construction_cost": 190_000_000.0,
    "purchase_unit": 120.0,
    "logistics_unit": 20.0,
    "supply_unit": 210.0,
    "interest_rate": 0.05,
    "mgmt_rate": 0.145,
    "depreciation_months": 120.0,
    "equipment_rent": 0.0,
    "rent_count": 0.0,
    "logi_km": 70.0,
    "logi_fuel_price": 1400.0,
    "logi_efficiency": 2.5,
    "logi_toll": 0.0,
    "logi_roundtrips": 10.0,
    "logi_supply_kg": 20_000.0,
    "logi_origin": "경기도 화성시 마도면 쌍송리 신일가스",
    "logi_dest": "",
}
def load_profit_inputs():
    """수익성분석 입력값 로드 (없으면 엑셀 기본값)."""
    data = dict(PROFIT_DEFAULTS)
    if os.path.exists(PROFIT_CACHE_FILE):
        try:
            with open(PROFIT_CACHE_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
            if isinstance(saved, dict):
                data.update({k: saved[k] for k in PROFIT_DEFAULTS if k in saved})
        except Exception:
            pass
    return data
def save_profit_inputs(data):
    try:
        with open(PROFIT_CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return True
    except Exception:
        return False
def compute_profitability(p):
    """엑셀 수익성분석 함수 동일 적용.
    C10=C11*12, C20=C9+C16+C18, C27=C26-(C24+C25),
    C30=H36*0.145, C35=F35*H35, C36=F36/120, H36=C11*C26,
    C37=F37*H37/12, C38=F38-H38-C30, C39=C38*3+(F39*H39),
    P18=(J18*K18/L18+M18)*N18/O18
    """
    monthly_usage = float(p.get("monthly_usage_kg") or 0)
    tank_price = float(p.get("tank_price") or 0)
    vap_price = float(p.get("vaporizer_price") or 0)
    const_cost = float(p.get("construction_cost") or 0)
    purchase = float(p.get("purchase_unit") or 0)
    logistics = float(p.get("logistics_unit") or 0)
    supply = float(p.get("supply_unit") or 0)
    rate = float(p.get("interest_rate") or 0)
    mgmt_rate = float(p.get("mgmt_rate") or 0)
    dep_m = float(p.get("depreciation_months") or 120) or 120.0
    rent = float(p.get("equipment_rent") or 0)
    rent_n = float(p.get("rent_count") or 0)
    yearly_usage = monthly_usage * 12  # C10
    total_invest = tank_price + vap_price + const_cost  # C20
    margin_kg = supply - (purchase + logistics)  # C27
    monthly_sales = monthly_usage * supply  # H36
    monthly_gross = monthly_usage * margin_kg  # C35
    depreciation = total_invest / dep_m  # C36
    finance = total_invest * rate / 12  # C37
    mgmt = monthly_sales * mgmt_rate  # C30
    invest_cost = depreciation + finance  # H38
    monthly_profit = monthly_gross - invest_cost - mgmt  # C38
    three_month = monthly_profit * 3 + (rent * rent_n)  # C39
    eff = float(p.get("logi_efficiency") or 0) or 1.0
    supply_kg = float(p.get("logi_supply_kg") or 0) or 1.0
    logi_per_kg = (
        (float(p.get("logi_km") or 0) * float(p.get("logi_fuel_price") or 0) / eff
         + float(p.get("logi_toll") or 0))
        * float(p.get("logi_roundtrips") or 0)
        / supply_kg
    )  # P18
    return {
        "yearly_usage": yearly_usage,
        "total_invest": total_invest,
        "margin_kg": margin_kg,
        "monthly_sales": monthly_sales,
        "monthly_gross": monthly_gross,
        "depreciation": depreciation,
        "finance": finance,
        "mgmt": mgmt,
        "invest_cost": invest_cost,
        "monthly_profit": monthly_profit,
        "three_month": three_month,
        "logi_per_kg": logi_per_kg,
    }
@st.cache_data(show_spinner=False, max_entries=8)
def _cached_profitability_report_excel(p_json, r_json, route_json, diesel_json):
    """동일 입력 반복 시 엑셀 재생성 생략 (Tab9 입력 체감 속도)."""
    return build_profitability_report_excel(
        json.loads(p_json),
        json.loads(r_json),
        route_info=json.loads(route_json),
        diesel_info=json.loads(diesel_json),
    )
def build_profitability_report_excel(p, r, route_info=None, diesel_info=None):
    """스크린샷형 수익성 분석 보고서 엑셀(bytes). Tab9 전용."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "수익성분석"
    # 열 폭 (보고서 가독성)
    for col, w in enumerate([3, 38, 28, 28, 28, 3], start=1):
        ws.column_dimensions[get_column_letter(col)].width = w

    thin = Border(
        left=Side(style="thin", color="94A3B8"),
        right=Side(style="thin", color="94A3B8"),
        top=Side(style="thin", color="94A3B8"),
        bottom=Side(style="thin", color="94A3B8"),
    )
    dash = Border(
        left=Side(style="dashed", color="94A3B8"),
        right=Side(style="dashed", color="94A3B8"),
        top=Side(style="dashed", color="94A3B8"),
        bottom=Side(style="dashed", color="94A3B8"),
    )
    title_fill = PatternFill("solid", fgColor="F1F5F9")
    green_fill = PatternFill("solid", fgColor="DCFCE7")
    note_font = Font(name="맑은 고딕", size=8, color="64748B")
    body_font = Font(name="맑은 고딕", size=11, color="0F172A")
    bold_font = Font(name="맑은 고딕", size=11, bold=True, color="0F172A")
    head_font = Font(name="맑은 고딕", size=12, bold=True, color="0F172A")
    title_font = Font(name="맑은 고딕", size=16, bold=True, color="0F172A")
    green_font = Font(name="맑은 고딕", size=9, bold=True, color="166534")
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    def _won(v):
        try:
            return f"{float(v):,.0f}"
        except Exception:
            return "0"

    def _num(v, d=1):
        try:
            return f"{float(v):,.{d}f}"
        except Exception:
            return "0"

    def put(row, col, text, font=None, align=None, fill=None, border=None, merge_to=None):
        cell = ws.cell(row=row, column=col, value=text)
        cell.font = font or body_font
        cell.alignment = align or left_align
        if fill:
            cell.fill = fill
        if border:
            cell.border = border
        if merge_to:
            ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=merge_to)
            for c in range(col, merge_to + 1):
                ws.cell(row=row, column=c).border = border or Border()
                ws.cell(row=row, column=c).fill = fill or PatternFill()
                ws.cell(row=row, column=c).alignment = align or left_align
        return cell

    name = str(p.get("project_name") or "프로젝트")
    dep_m = float(p.get("depreciation_months") or 120) or 120.0
    dep_years = dep_m / 12.0
    rate = float(p.get("interest_rate") or 0)
    mgmt_rate = float(p.get("mgmt_rate") or 0)
    route_info = route_info or {}
    diesel_info = diesel_info or {}

    # 1) 제목
    put(2, 2, "수익성 분석", font=title_font, align=center, fill=title_fill, border=thin, merge_to=5)
    ws.row_dimensions[2].height = 32

    # 2) 보고 헤더
    put(4, 2, f"◆ [{name}] 투자대비 수익성 분석 보고", font=head_font, merge_to=5)
    put(
        5, 2,
        f"TANK: {p.get('tank_gas','')} {float(p.get('tank_liters') or 0):,.0f} L → "
        f"{parse_tank_capacity_kg(p.get('tank_spec')):,.0f} kg · 기화기 {float(p.get('vaporizer_capacity') or 0):,.0f} Nm3/hr "
        f"{p.get('vaporizer_qty_note','')} · 년 사용량(자동) {_won(r['yearly_usage'])} kg",
        font=note_font, merge_to=5,
    )

    # 3) 장비 투자비
    row = 7
    put(row, 2, "○ 장비 투자비", font=bold_font, merge_to=5)
    row = 8
    _xlsx_kg = parse_tank_capacity_kg(p.get("tank_spec"))
    _xlsx_gas = str(p.get("tank_gas") or "질소")
    _xlsx_nm3 = tank_kg_to_nm3(_xlsx_kg, _xlsx_gas)
    put(
        row, 2,
        f"1. TANK 구매 : {_xlsx_gas} {float(p.get('tank_liters') or 0):,.0f} L "
        f"→ {_xlsx_kg:,.0f} kg · 기체환산 {_xlsx_nm3:,.1f} Nm³",
        font=body_font, merge_to=5,
    )
    row = 9
    put(row, 2, f"   구입가: {_won(p.get('tank_price'))} 원", font=body_font)
    put(row, 3, f"년 사용량: {_won(r['yearly_usage'])} kg", font=body_font)
    put(row, 4, f"월평균 공급량: {_won(p.get('monthly_usage_kg'))} kg", font=body_font)
    row = 10
    put(row, 2, f"2. 기화기 구매 : {float(p.get('vaporizer_capacity') or 0):,.0f} Nm3/hr {p.get('vaporizer_qty_note','')}", font=body_font, merge_to=5)
    row = 11
    put(row, 2, f"   구입가: {_won(p.get('vaporizer_price'))} 원", font=body_font, merge_to=5)
    row = 12
    put(row, 2, f"3. 공사비용: {_won(p.get('construction_cost'))} 원", font=body_font, merge_to=5)
    row = 13
    put(row, 2, f"총 투자금액: {_won(r['total_invest'])} 원", font=bold_font, merge_to=5)
    row = 14
    put(row, 2, "합계 = TANK + 기화기 + 공사비용", font=note_font, merge_to=5)

    # 4) 단가
    row = 16
    put(row, 2, "○ 단가", font=bold_font, merge_to=5)
    row = 17
    put(row, 2, f"1. 매입단가: {_num(p.get('purchase_unit'), 1)} 원/kg", font=body_font, merge_to=3)
    put(row, 4, f"2. 물류비: {_num(p.get('logistics_unit'), 2)} 원/kg", font=body_font, merge_to=5)
    row = 18
    put(row, 2, f"3. 공급단가: {_num(p.get('supply_unit'), 1)} 원/kg", font=body_font, merge_to=3)
    put(row, 4, f"4. 매출이익: {_num(r['margin_kg'], 1)} 원/kg", font=body_font, merge_to=5)
    row = 19
    put(row, 2, "매출이익 = 공급단가 − (매입단가 + 물류비)", font=note_font, merge_to=5)

    # 5) 일반관리비
    row = 21
    put(row, 2, "○ 일반관리비", font=bold_font)
    put(row, 3, f"{_won(r['mgmt'])} 원/월", font=bold_font)
    put(row, 4, f"(월 매출 {mgmt_rate*100:.1f}%)", font=note_font, merge_to=5)
    row = 22
    put(row, 2, "일반관리비 = 월매출 × 관리비율  ·  월매출 = 월평균공급량 × 공급단가", font=note_font, merge_to=5)

    # 6) 사용량 대비 영업이익 표
    row = 24
    put(row, 2, "○ 사용량 대비 영업이익", font=bold_font, merge_to=5)
    table = [
        [
            f"월평균 매출이익: {_won(r['monthly_gross'])} 원/월",
            f"월 사용량(kg): {_won(p.get('monthly_usage_kg'))}",
            f"매출마진: {_num(r['margin_kg'], 1)}",
        ],
        [
            "월사용량 × 매출이익",
            "프로젝트 월 평균 공급량",
            "공급 − (매입 + 물류)",
        ],
        [
            f"장비 감가상각비({dep_years:.0f}년): {_won(r['depreciation'])} 원/월",
            f"투자총액: {_won(r['total_invest'])}",
            f"월 매출: {_won(r['monthly_sales'])}",
        ],
        [
            f"투자합계 ÷ {dep_m:.0f}개월",
            "TANK+기화기+공사",
            "월사용량 × 공급단가",
        ],
        [
            f"금융비: {_won(r['finance'])} 원/월",
            f"원금: {_won(r['total_invest'])}",
            f"이자율: {rate*100:.1f}%",
        ],
        [
            "원금 × 이자율 ÷ 12",
            "총 투자금액",
            "입력 이자율",
        ],
        [
            f"월평균 이익금: {_won(r['monthly_profit'])} 원/월",
            f"투자비용: {_won(r['invest_cost'])}",
            f"일반관리비: {_won(r['mgmt'])}",
        ],
        [
            "매출이익 − 투자비용 − 관리비",
            "감가상각 + 금융비",
            f"월매출 × {mgmt_rate*100:.1f}%",
        ],
        [
            f"최근 3개월 매출 실이익: {_won(r['three_month'])} 원",
            f"장비 임대료: {_won(p.get('equipment_rent'))}",
            f"부가횟수: {_num(p.get('rent_count'), 0)}",
        ],
        [
            "월이익×3 + 임대료×횟수",
            "입력 임대료",
            "입력 부가횟수",
        ],
    ]
    t_row = 25
    for i, cells in enumerate(table):
        is_note = (i % 2 == 1)
        for j, text in enumerate(cells):
            c = put(
                t_row, 2 + j, text,
                font=note_font if is_note else body_font,
                border=dash,
                align=left_align,
            )
            if not is_note:
                c.font = bold_font if j == 0 and ("이익금" in text or "실이익" in text or "매출이익" in text) else body_font
        ws.row_dimensions[t_row].height = 18 if is_note else 22
        t_row += 1

    # 7) 물류비 계산 설명 (스크린샷 작은글씨 전부)
    t_row += 1
    put(t_row, 2, "○ 물류비 계산 상세 (20톤 벌크로리 · 경유)", font=bold_font, merge_to=5)
    t_row += 1
    put(
        t_row, 2,
        f"거리 { _num(p.get('logi_km'), 1) } km · 유류비 {_num(p.get('logi_fuel_price'), 2)} 원/L · "
        f"연비 {_num(p.get('logi_efficiency'), 1)} km/L · 통행료(편도·5종) {_won(p.get('logi_toll'))} 원 · "
        f"왕복 {_num(p.get('logi_roundtrips'), 0)} 회 · 월평균공급량 {_won(p.get('logi_supply_kg'))} kg",
        font=note_font, merge_to=5,
    )
    t_row += 1
    put(
        t_row, 2,
        f"물류비(원/kg) = ((거리 × 유류비 ÷ 연비) + 통행료) × 왕복 ÷ 월평균공급량  =  {_num(r['logi_per_kg'], 2)} 원/kg",
        font=note_font, merge_to=5,
    )
    t_row += 1
    _rt_xlsx = compute_roundtrips_from_tank(
        p.get("monthly_usage_kg"), p.get("tank_spec"), fill_ratio=0.8
    )
    put(
        t_row, 2,
        _rt_xlsx.get("message")
        or "왕복횟수 = ceil(월평균공급량 ÷ (탱크총용량kg × 80%))",
        font=note_font, merge_to=5,
    )
    t_row += 1
    route_msg = str(route_info.get("message") or "")
    if route_msg:
        put(t_row, 2, route_msg, font=green_font, fill=green_fill, border=thin, merge_to=5)
        ws.row_dimensions[t_row].height = 22
        t_row += 1
    o_lab = route_info.get("origin_label") or p.get("logi_origin") or ""
    d_lab = route_info.get("dest_label") or p.get("logi_dest") or ""
    if o_lab or d_lab:
        put(
            t_row, 2,
            f"출발: {o_lab}  →  도착: {d_lab}"
            + (f"  ·  {route_msg}" if route_msg else ""),
            font=note_font, merge_to=5,
        )
        t_row += 1
    if diesel_info.get("ok"):
        put(
            t_row, 2,
            f"경유 시장가 {_num(diesel_info.get('price'), 2)} 원/L  ·  "
            f"{diesel_info.get('source','')}  ·  기준일 {diesel_info.get('asof','')}  ·  "
            f"{diesel_info.get('month', datetime.date.today().strftime('%Y-%m'))} 자동반영",
            font=note_font, merge_to=5,
        )
        t_row += 1
    elif diesel_info.get("message"):
        put(t_row, 2, str(diesel_info.get("message")), font=note_font, merge_to=5)
        t_row += 1

    t_row += 1
    put(
        t_row, 2,
        "적용 함수: 년사용량=월×12 · 합계=탱크+기화기+공사 · 매출이익=공급−(매입+물류) · "
        "관리비=월매출×비율 · 감가=투자÷개월 · 금융=원금×이자÷12 · "
        "월이익=매출이익−(감가+금융)−관리비 · 3개월=월이익×3+(임대×횟수) · "
        "통행료=카카오1종×2.5(20톤벌크 5종)",
        font=note_font, merge_to=5,
    )
    t_row += 2
    put(t_row, 2, "1/1페이지", font=note_font, align=center, fill=title_fill, border=thin, merge_to=5)

    ws.print_title_rows = "1:2"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()
def get_saved_date(file_path):
    if os.path.exists(file_path):
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                return datetime.datetime.strptime(f.read().strip(), "%Y-%m-%d").date()
        except:
            pass
    return datetime.date.today()
def set_saved_date(file_path, date_val):
    with open(file_path, "w", encoding="utf-8") as f:
        f.write(date_val.strftime("%Y-%m-%d"))
# ==========================================
# ★ 탭 전체 적용 표 합계행 렌더링 유틸리티 (무손실 보존) ★
# ==========================================
def get_display_df_with_sum(df, sum_label="연간 합계", text_cols=None):
    if df is None or df.empty: return df
    disp = df.copy()
    numeric_cols = disp.select_dtypes(include=[np.number]).columns
    sum_series = disp[numeric_cols].sum()
    disp.loc[sum_label] = sum_series
    if text_cols:
        for col in text_cols:
            if col in disp.columns:
                disp.at[sum_label, col] = "총 합계"
    return disp
def style_with_sum(disp_df, fmt_str, cmap=None, subset_cols=None, axis=None):
    if disp_df.empty:
        return disp_df.style
    if subset_cols is None:
        subset_cols = disp_df.select_dtypes(include=[np.number]).columns
    grad_subset = pd.IndexSlice[disp_df.index[:-1], subset_cols]
    styled = disp_df.style.format(fmt_str)
    if cmap:
        styled = styled.background_gradient(cmap=cmap, axis=axis, subset=grad_subset)
    styled = styled.apply(
        lambda s: ['font-weight: 800; background-color: #E2E8F0; color: #0F172A;'] * len(s),
        subset=pd.IndexSlice[disp_df.index[-1], :],
        axis=1,
    )
    return styled
def inject_top30_month_bridge():
    """iframe(월 헤더 클릭) → 부모 URL query 갱신 + iPad Top30 도넛 스택. sticky/맥 로직 분리."""
    components.html(
        """
        <script>
        (function () {
            var parentWin = window.parent;
            if (!parentWin) return;
            var parentDoc = parentWin.document;
            function isTouchPad() {
                try {
                    var ua = parentWin.navigator.userAgent || "";
                    var ios = /iPad|iPhone|iPod/.test(ua);
                    var ipadOs = parentWin.navigator.platform === "MacIntel" && parentWin.navigator.maxTouchPoints > 1;
                    return ios || ipadOs || parentDoc.documentElement.classList.contains("dashboard-touch-mode");
                } catch (e) { return false; }
            }
            /* iPad ↔ 맥 UI 분기용 query (한 번만). 맥에는 touch_ui 안 붙음 */
            try {
                if (!parentWin.__dashboardTouchUiSynced) {
                    parentWin.__dashboardTouchUiSynced = true;
                    var url = new URL(parentWin.location.href);
                    var cur = url.searchParams.get("touch_ui");
                    var touch = isTouchPad();
                    if (touch && cur !== "1") {
                        url.searchParams.set("touch_ui", "1");
                        parentWin.location.replace(url.toString());
                        return;
                    }
                    if (!touch && cur === "1") {
                        url.searchParams.delete("touch_ui");
                        parentWin.location.replace(url.toString());
                        return;
                    }
                }
            } catch (eSync) {}
            if (!parentWin.__dashboardTop30MonthBridge) {
                parentWin.__dashboardTop30MonthBridge = true;
                parentWin.addEventListener("message", function (e) {
                    try {
                        var d = e.data;
                        if (!d || d.type !== "dashboard-top30-month" || !d.month) return;
                        var url2 = new URL(parentWin.location.href);
                        url2.searchParams.set("top30_month", d.month);
                        parentWin.location.href = url2.toString();
                    } catch (err) {}
                });
            }
            /* iPad(touch-mode)만: Top30 표+도넛 행을 세로 100% 폭 — 맥은 즉시 return */
            function forceFullWidth(el) {
                if (!el || !el.style) return;
                el.style.setProperty("width", "100%", "important");
                el.style.setProperty("min-width", "100%", "important");
                el.style.setProperty("max-width", "100%", "important");
                el.style.setProperty("flex", "1 1 100%", "important");
                el.style.setProperty("box-sizing", "border-box", "important");
            }
            function applyTop30FullWidthRow(row) {
                row.classList.add("top30-touch-row");
                row.style.setProperty("display", "flex", "important");
                row.style.setProperty("flex-direction", "column", "important");
                row.style.setProperty("flex-wrap", "nowrap", "important");
                row.style.setProperty("align-items", "stretch", "important");
                row.style.setProperty("width", "100%", "important");
                row.style.setProperty("max-width", "100%", "important");
                Array.prototype.forEach.call(row.children, forceFullWidth);
                /* 직계(표|도넛) 칼럼만 — 중첩 칼럼에 min-width:100% 금지 */
                var topCols = row.querySelectorAll(':scope > [data-testid="column"], :scope > [data-testid="stColumn"], :scope > div > [data-testid="column"], :scope > div > [data-testid="stColumn"]');
                if (!topCols.length) {
                    topCols = [];
                    Array.prototype.forEach.call(row.children, function (ch) {
                        if (ch.matches && (ch.matches('[data-testid="column"]') || ch.matches('[data-testid="stColumn"]'))) {
                            topCols.push(ch);
                        } else if (ch.querySelector) {
                            var inner = ch.querySelector(':scope > [data-testid="column"], :scope > [data-testid="stColumn"]');
                            if (inner) topCols.push(inner);
                        }
                    });
                }
                Array.prototype.forEach.call(topCols, forceFullWidth);
                row.querySelectorAll("iframe").forEach(function (f) {
                    f.style.setProperty("width", "100%", "important");
                    f.style.setProperty("max-width", "100%", "important");
                    try { f.setAttribute("width", "100%"); } catch (e0) {}
                });
                var plots = row.querySelectorAll('[data-testid="stPlotlyChart"]');
                plots.forEach(function (p) {
                    p.style.setProperty("width", "100%", "important");
                    p.style.setProperty("min-height", "420px", "important");
                    p.style.setProperty("height", "420px", "important");
                });
                try { parentWin.dispatchEvent(new Event("resize")); } catch (e1) {}
                setTimeout(function () {
                    try { parentWin.dispatchEvent(new Event("resize")); } catch (e2) {}
                    plots.forEach(function (p) {
                        var ifr = p.querySelector("iframe");
                        if (!ifr) return;
                        try {
                            var pw = ifr.contentWindow;
                            if (pw && pw.Plotly) {
                                var gds = ifr.contentDocument.querySelectorAll(".js-plotly-plot");
                                for (var g = 0; g < gds.length; g++) {
                                    try { pw.Plotly.Plots.resize(gds[g]); } catch (e3) {}
                                }
                            }
                        } catch (e4) {}
                    });
                }, 250);
            }
            function stackTop30DonutRow() {
                try {
                    if (!isTouchPad()) return;
                    var flag = parentDoc.querySelector(".top30-section-flag");
                    if (!flag) return;
                    var scope = flag.closest('[data-testid="stVerticalBlock"]') || flag.parentElement;
                    if (!scope) return;
                    scope.classList.add("top30-touch-scope");
                    forceFullWidth(scope);
                    var rows = scope.querySelectorAll('[data-testid="stHorizontalBlock"]');
                    for (var i = 0; i < rows.length; i++) {
                        var row = rows[i];
                        var cols = row.querySelectorAll(':scope > [data-testid="column"], :scope > [data-testid="stColumn"], :scope > div > [data-testid="column"], :scope > div > [data-testid="stColumn"]');
                        if (!cols.length) {
                            cols = row.querySelectorAll('[data-testid="column"], [data-testid="stColumn"]');
                        }
                        /* 표|도넛(플롯 포함) 행만 풀폭 — 월 버튼 12칸 행은 제외 */
                        if (cols.length === 12) continue;
                        var hasPlot = !!row.querySelector('[data-testid="stPlotlyChart"]');
                        if (hasPlot) {
                            applyTop30FullWidthRow(row);
                        }
                    }
                } catch (err) {}
            }
            stackTop30DonutRow();
            if (!parentWin.__dashboardTop30StackTimer) {
                parentWin.__dashboardTop30StackTimer = parentWin.setInterval(stackTop30DonutRow, 600);
            }
        })();
        </script>
        """,
        height=0,
    )
def is_touch_ui():
    """iPad UI 분기. query touch_ui / cookie / session. 맥은 False."""
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
    return bool(st.session_state.get("force_touch_ui"))
def render_plotly_chart(fig, *, key=None, use_container_width=True, allow_drag=False, height=None, **kwargs):
    """맥: st.plotly_chart 무손실.
    iPad: components.html 직접 렌더 → 컨트롤바는 그래프 탭 시에만 표시.
    """
    if not is_touch_ui():
        st.plotly_chart(fig, use_container_width=use_container_width, key=key, **kwargs)
        return
    try:
        layout_h = fig.layout.height
    except Exception:
        layout_h = None
    try:
        h = int(height or layout_h or 450)
    except Exception:
        h = 450
    try:
        updates = {
            "clickmode": "none",
            "margin": dict(t=56, r=12, b=12, l=12),
            "modebar": dict(
                orientation="h",
                bgcolor="rgba(15, 23, 42, 0.92)",
                color="#E2E8F0",
                activecolor="#38BDF8",
            ),
        }
        if not allow_drag:
            updates["dragmode"] = False
        fig.update_layout(**updates)
    except Exception:
        pass
    config = {
        "displayModeBar": True,
        "displaylogo": False,
        "responsive": True,
        "scrollZoom": False,
        "doubleClick": "reset",
        "modeBarButtonsToRemove": ["lasso2d", "select2d"],
    }
    try:
        inner = fig.to_html(include_plotlyjs="cdn", full_html=False, config=config)
    except Exception:
        st.plotly_chart(fig, use_container_width=use_container_width, key=key, config=config, **kwargs)
        return
    page_html = f"""<!DOCTYPE html>
<html><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<style>
  html, body {{ margin:0; padding:0; background:#fff; overflow:hidden; }}
  .wrap {{ position:relative; width:100%; }}
  .modebar-container, .modebar {{
    opacity:0 !important;
    visibility:hidden !important;
    pointer-events:none !important;
    transition: opacity .15s ease;
    z-index:9999 !important;
  }}
  body.mb-show .modebar-container,
  body.mb-show .modebar {{
    opacity:1 !important;
    visibility:visible !important;
    display:flex !important;
    pointer-events:auto !important;
  }}
  .modebar-container {{
    top:8px !important; right:8px !important; left:auto !important;
  }}
  .modebar {{
    background:rgba(15,23,42,.94) !important;
    border-radius:8px !important; padding:4px !important;
  }}
  .modebar-btn {{ min-width:32px !important; min-height:32px !important; }}
</style></head>
<body>
<div class="wrap" id="plot-wrap">
  {inner}
</div>
<script>
(function() {{
  var body = document.body;
  function showBar() {{ body.classList.add("mb-show"); }}
  function onTap(e) {{
    if (e.target && e.target.closest && e.target.closest(".modebar")) return;
    showBar();
  }}
  document.addEventListener("click", onTap, true);
  document.addEventListener("touchend", onTap, true);
}})();
</script>
</body></html>"""
    components.html(page_html, height=h + 8, scrolling=False)
def inject_ipad_plotly_controls():
    """iPad: touch cookie/query 동기화 + 예전 원복 플로팅 버튼 제거. 맥 무손실."""
    components.html(
        """
        <script>
        (function () {
            var parentWin = window.parent;
            if (!parentWin) return;
            var parentDoc = parentWin.document;
            function isTouchPad() {
                try {
                    var ua = parentWin.navigator.userAgent || "";
                    var ios = /iPad|iPhone|iPod/.test(ua);
                    var ipadOs = parentWin.navigator.platform === "MacIntel" && parentWin.navigator.maxTouchPoints > 1;
                    return ios || ipadOs || parentDoc.documentElement.classList.contains("dashboard-touch-mode");
                } catch (e) { return false; }
            }
            try {
                var old = parentDoc.getElementById("dashboard-ipad-plotly-reset");
                if (old) old.remove();
            } catch (e0) {}
            if (!isTouchPad()) return;
            try {
                parentDoc.cookie = "dashboard_touch=1; path=/; max-age=31536000; SameSite=Lax";
            } catch (e1) {}
            try {
                if (!parentWin.__dashboardTouchUiSynced) {
                    parentWin.__dashboardTouchUiSynced = true;
                    var url = new URL(parentWin.location.href);
                    if (url.searchParams.get("touch_ui") !== "1") {
                        url.searchParams.set("touch_ui", "1");
                        parentWin.location.replace(url.toString());
                        return;
                    }
                }
            } catch (e2) {}
        })();
        </script>
        """,
        height=0,
    )
def render_frozen_styler_html(
    styled,
    height=450,
    freeze_left_n=2,
    freeze_widths=None,
    clickable_cols=None,
    query_param=None,
    active_col=None,
):
    """Styler HTML 렌더 + 헤더/좌측열 틀고정 (데이터·스타일 무손실)."""
    widths = freeze_widths or ([44, 160] if freeze_left_n >= 2 else [160])
    left_css = []
    left = 0
    for i in range(freeze_left_n):
        w = widths[i] if i < len(widths) else 100
        n = i + 1
        left_css.append(
            f"""
            table thead tr th:nth-child({n}),
            table tbody tr th:nth-child({n}),
            table tbody tr td:nth-child({n}) {{
                position: sticky;
                left: {left}px;
                z-index: 4;
                min-width: {w}px;
                max-width: {w}px;
                box-shadow: 2px 0 0 #CBD5E1;
            }}
            table thead tr th:nth-child({n}) {{
                z-index: 8;
                background: #F0F2F6 !important;
            }}
            table tbody tr th:nth-child({n}),
            table tbody tr td:nth-child({n}) {{
                background-color: #FFFFFF;
                z-index: 5;
            }}
            table tbody tr:last-child th:nth-child({n}),
            table tbody tr:last-child td:nth-child({n}) {{
                background-color: #E2E8F0 !important;
            }}
            """
        )
        left += w
    table_html = styled.to_html()
    click_js = ""
    if clickable_cols and query_param:
        def _link_thead(match):
            block = match.group(0)
            for col in clickable_cols:
                active = "color:#B91C1C;font-weight:700;" if active_col == col else "color:inherit;font-weight:600;"
                block = re.sub(
                    rf"(<th[^>]*>)\s*{re.escape(str(col))}\s*(</th>)",
                    rf'\1<a href="#" role="button" data-top30-month="{col}" '
                    rf'style="text-decoration:none;cursor:pointer;display:block;{active}" '
                    rf'title="클릭: {col} 매출 순위·그래프 적용">{col}</a>\2',
                    block,
                )
            return block
        table_html = re.sub(r"<thead>.*?</thead>", _link_thead, table_html, count=1, flags=re.DOTALL)
        click_js = """
        <script>
        (function () {
            function notify(month) {
                try {
                    window.parent.postMessage({ type: "dashboard-top30-month", month: month }, "*");
                } catch (err) {}
            }
            document.querySelectorAll("a[data-top30-month]").forEach(function (a) {
                a.addEventListener("click", function (e) {
                    e.preventDefault();
                    e.stopPropagation();
                    notify(a.getAttribute("data-top30-month"));
                });
            });
        })();
        </script>
        """
    page_html = f"""
    <!DOCTYPE html>
    <html><head><meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <style>
        html, body {{
            margin: 0; height: 100%; overflow: hidden;
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
            font-size: 13px; color: #31333F;
        }}
        .wrap {{
            height: 100%; overflow: auto;
            -webkit-overflow-scrolling: touch;
            border: 1px solid #E2E8F0; border-radius: 4px; background: #fff;
        }}
        table {{
            border-collapse: separate; border-spacing: 0;
            width: max-content; min-width: 100%;
        }}
        th, td {{
            padding: 6px 10px; white-space: nowrap;
            border-bottom: 1px solid #E2E8F0;
            font-weight: 400;
        }}
        thead th {{
            position: sticky; top: 0; z-index: 6;
            background: #F0F2F6 !important;
            box-shadow: 0 1px 0 #CBD5E1;
            font-weight: 600; text-align: right;
        }}
        thead th:nth-child(1), thead th:nth-child(2) {{ text-align: left; }}
        thead th a:hover {{ text-decoration: underline !important; }}
        {''.join(left_css)}
    </style></head>
    <body><div class="wrap">{table_html}</div>{click_js}</body></html>
    """
    components.html(page_html, height=height, scrolling=True)
# ==========================================
# 5. 메인 실행 흐름 및 영구 캐싱 관리
# ==========================================
inject_custom_css()
st.sidebar.header("📁 데이터 업로드 및 유지")
address_file_up = st.sidebar.file_uploader("거래처 주소록 (CSV)", type=["csv"])
industry_file_up = st.sidebar.file_uploader("🏢 거래처 업종 분류 (CSV)", type=["csv"])
debt_file_up = st.sidebar.file_uploader("채권 데이터 (채권.csv)", type=["csv"])
uploaded_files_up = st.sidebar.file_uploader("매출 데이터 (다중 업로드)", type=["csv"], accept_multiple_files=True)
st.sidebar.markdown("---")
st.sidebar.subheader("🏭 설비 재고 관리 (선택)")
tank_file_up = st.sidebar.file_uploader("탱크 재고현황 (CSV/Excel)", type=["csv", "xlsx"])
vaporizer_file_up = st.sidebar.file_uploader("기화기 재고현황 (CSV/Excel)", type=["csv", "xlsx"])
st.sidebar.markdown("---")
st.sidebar.subheader("🛢️ 통합 탱크 재고")
local_int_path = "통합탱크재고.csv"
int_bytes = None
int_name = ""
if os.path.exists(local_int_path):
    st.sidebar.success(f"✅ '{local_int_path}' 자동 로드됨")
    with open(local_int_path, "rb") as f:
        int_bytes = f.read()
    int_name = local_int_path
else:
    integrated_file_up = st.sidebar.file_uploader("통합 탱크 재고현황 (CSV)", type=["csv"])
    if integrated_file_up is not None:
        int_bytes = integrated_file_up.getvalue()
        int_name = integrated_file_up.name
@st.cache_data(show_spinner="설비 데이터를 읽어오는 중입니다...")
def load_equipment_file(file_bytes, file_name):
    if not file_bytes:
        return pd.DataFrame()
    try:
        if file_name.endswith('.csv'):
            decoded = None
            for enc in ['utf-8-sig', 'utf-8', 'cp949', 'euc-kr']:
                try:
                    decoded = file_bytes.decode(enc)
                    break
                except UnicodeDecodeError:
                    continue
            
            if decoded is None:
                decoded = file_bytes.decode('utf-8', errors='replace')
            
            lines = [line.strip().strip('"') for line in decoded.splitlines() if line.strip()]
            
            df = pd.read_csv(io.StringIO("\n".join(lines)), on_bad_lines='skip', engine='python')
            return df
        else:
            return pd.read_excel(io.BytesIO(file_bytes))
    except Exception as e:
        st.sidebar.error(f"파일 읽기 오류 ({file_name}): {e}")
        return pd.DataFrame()
st.sidebar.markdown("---")
st.sidebar.subheader("🔑 Open DART API 설정")
API_KEY_FILE = os.path.join(CACHE_DIR, "dart_api_key.txt")
saved_api_key = ""
if os.path.exists(API_KEY_FILE):
    with open(API_KEY_FILE, "r", encoding="utf-8") as f:
        saved_api_key = f.read().strip()
dart_api_key = st.sidebar.text_input(
    "DART API 키 (재무정보 연동용)", 
    value=saved_api_key, 
    type="password", 
    help="금융감독원 Open DART API 키를 입력하세요. 한번 입력하면 자동 저장되어 새로고침해도 유지됩니다."
)
if dart_api_key and dart_api_key != saved_api_key:
    with open(API_KEY_FILE, "w", encoding="utf-8") as f:
        f.write(dart_api_key)
addr_cache_path = os.path.join(CACHE_DIR, "address.csv")
industry_cache_path = os.path.join(CACHE_DIR, "industry.csv")
debt_cache_path = os.path.join(CACHE_DIR, "debt.csv")
tank_cache_path = os.path.join(CACHE_DIR, "tank_cache.dat")
vaporizer_cache_path = os.path.join(CACHE_DIR, "vaporizer_cache.dat")
integrated_cache_path = os.path.join(CACHE_DIR, "integrated_cache.dat")
sales_cache_dir = os.path.join(CACHE_DIR, "sales")
os.makedirs(sales_cache_dir, exist_ok=True)
if address_file_up is not None:
    addr_bytes = address_file_up.getvalue()
    with open(addr_cache_path, "wb") as f: f.write(addr_bytes)
elif os.path.exists(addr_cache_path):
    with open(addr_cache_path, "rb") as f: addr_bytes = f.read()
elif os.path.exists("주소.csv"):
    with open("주소.csv", "rb") as f: addr_bytes = f.read()
else:
    addr_bytes = None
if industry_file_up is not None:
    ind_bytes = industry_file_up.getvalue()
    with open(industry_cache_path, "wb") as f: f.write(ind_bytes)
elif os.path.exists(industry_cache_path):
    with open(industry_cache_path, "rb") as f: ind_bytes = f.read()
elif os.path.exists("업체대분류.csv"):
    with open("업체대분류.csv", "rb") as f: ind_bytes = f.read()
else:
    ind_bytes = None
if debt_file_up is not None:
    debt_bytes = debt_file_up.getvalue()
    with open(debt_cache_path, "wb") as f: f.write(debt_bytes)
elif os.path.exists(debt_cache_path):
    with open(debt_cache_path, "rb") as f: debt_bytes = f.read()
else:
    debt_bytes = None
    for f_name in os.listdir("."):
        if f_name.startswith("채권") and f_name.endswith(".csv"):
            with open(f_name, "rb") as f: debt_bytes = f.read()
            break
if tank_file_up is not None:
    tank_bytes = tank_file_up.getvalue()
    tank_name = tank_file_up.name
    with open(tank_cache_path, "wb") as f: f.write(tank_bytes)
    with open(tank_cache_path + "_name.txt", "w", encoding="utf-8") as f: f.write(tank_name)
elif os.path.exists(tank_cache_path) and os.path.exists(tank_cache_path + "_name.txt"):
    with open(tank_cache_path, "rb") as f: tank_bytes = f.read()
    with open(tank_cache_path + "_name.txt", "r", encoding="utf-8") as f: tank_name = f.read().strip()
elif os.path.exists("탱크.csv"):
    with open("탱크.csv", "rb") as f: tank_bytes = f.read()
    tank_name = "탱크.csv"
else:
    tank_bytes = None
    tank_name = ""
if vaporizer_file_up is not None:
    vaporizer_bytes = vaporizer_file_up.getvalue()
    vaporizer_name = vaporizer_file_up.name
    with open(vaporizer_cache_path, "wb") as f: f.write(vaporizer_bytes)
    with open(vaporizer_cache_path + "_name.txt", "w", encoding="utf-8") as f: f.write(vaporizer_name)
elif os.path.exists(vaporizer_cache_path) and os.path.exists(vaporizer_cache_path + "_name.txt"):
    with open(vaporizer_cache_path, "rb") as f: vaporizer_bytes = f.read()
    with open(vaporizer_cache_path + "_name.txt", "r", encoding="utf-8") as f: vaporizer_name = f.read().strip()
elif os.path.exists("기화기.csv"):
    with open("기화기.csv", "rb") as f: vaporizer_bytes = f.read()
    vaporizer_name = "기화기.csv"
else:
    vaporizer_bytes = None
    vaporizer_name = ""
if int_bytes is not None:
    with open(integrated_cache_path, "wb") as f: f.write(int_bytes)
    with open(integrated_cache_path + "_name.txt", "w", encoding="utf-8") as f: f.write(int_name)
elif os.path.exists(integrated_cache_path) and os.path.exists(integrated_cache_path + "_name.txt"):
    with open(integrated_cache_path, "rb") as f: int_bytes = f.read()
    with open(integrated_cache_path + "_name.txt", "r", encoding="utf-8") as f: int_name = f.read().strip()
else:
    int_bytes = None
    int_name = ""
sales_file_tuples = []
if uploaded_files_up and len(uploaded_files_up) > 0:
    for f_name in os.listdir(sales_cache_dir):
        os.remove(os.path.join(sales_cache_dir, f_name))
    for f in uploaded_files_up:
        f_bytes = f.getvalue()
        f_path = os.path.join(sales_cache_dir, f.name)
        with open(f_path, "wb") as sf: sf.write(f_bytes)
        sales_file_tuples.append((f.name, f_bytes))
else:
    if os.path.exists(sales_cache_dir):
        for f_name in os.listdir(sales_cache_dir):
            if f_name.endswith(".csv"):
                f_path = os.path.join(sales_cache_dir, f_name)
                with open(f_path, "rb") as sf: sales_file_tuples.append((f_name, sf.read()))
                
    if not sales_file_tuples:
        for f_name in os.listdir("."):
            if re.match(r"^20\d{2}.*\.csv$", f_name):
                with open(f_name, "rb") as sf:
                    sales_file_tuples.append((f_name, sf.read()))
if st.sidebar.button("🗑️ 저장된 캐시 데이터 초기화"):
    for p in [addr_cache_path, industry_cache_path, debt_cache_path, 
              tank_cache_path, tank_cache_path + "_name.txt", 
              vaporizer_cache_path, vaporizer_cache_path + "_name.txt",
              integrated_cache_path, integrated_cache_path + "_name.txt",
              TAB7_DATE_FILE, TAB8_DATE_FILE]:
        if os.path.exists(p): os.remove(p)
    if os.path.exists(API_KEY_FILE):
        os.remove(API_KEY_FILE)
    for f_name in os.listdir(sales_cache_dir):
        os.remove(os.path.join(sales_cache_dir, f_name))
    st.rerun()
addr_dict = load_address_file(addr_bytes) if addr_bytes else {}
industry_dict = load_industry_file(ind_bytes) if ind_bytes else {}
debt_df = load_debt_file(debt_bytes) if debt_bytes else pd.DataFrame()
df_tank = load_equipment_file(tank_bytes, tank_name) if tank_bytes else pd.DataFrame()
df_vaporizer = load_equipment_file(vaporizer_bytes, vaporizer_name) if vaporizer_bytes else pd.DataFrame()
df_integrated = load_equipment_file(int_bytes, int_name) if int_bytes else pd.DataFrame()
full_df = load_uploaded_files_from_bytes(sales_file_tuples) if sales_file_tuples else pd.DataFrame()
if not full_df.empty:
    is_deposit_row = full_df["품목명"].astype(str).str.contains("입금", na=False)
    full_df = full_df[~is_deposit_row].copy()
    full_df["업종"] = full_df["거래처"].map(industry_dict).fillna("미분류")
target_items = [
    "CO2 (kg, Bulk)",
    "N2 (kg, Bulk)",
    "O2 (kg, Bulk)",
    "AR (kg, Bulk)",
]
latest_update_str = "데이터 없음"
selected_staff = []
selected_client = "전체 거래처"
df_base = pd.DataFrame()
df_client_filtered = pd.DataFrame()
pivot_m_total = pd.DataFrame()
client_item_qty_pivot = pd.DataFrame()
sales_p = pd.DataFrame()
qty_p = pd.DataFrame()
unit_price_p = pd.DataFrame()
staff_pivot = pd.DataFrame()
df_detail = pd.DataFrame()
all_months = [f"{i:02d}월" for i in range(1, 13)]
years = ["2026"]
if not full_df.empty:
    latest_dt_overall = full_df["매출일_dt"].max()
    if pd.notnull(latest_dt_overall):
        latest_update_str = latest_dt_overall.strftime("%Y-%m-%d")
    # ==============================================================
    # 필터 영역 (상단 고정은 inject_sticky_tabs_script에서 처리)
    # ==============================================================
    st.markdown("<div id='dashboard-sticky-spacer'></div>", unsafe_allow_html=True)
    try:
        filter_container = st.container(border=True)
    except TypeError:
        filter_container = st.container()
        
    with filter_container:
        st.markdown("<div id='sticky-marker' style='display:none;'></div>", unsafe_allow_html=True)
        fc1, fc2, fc3, fc4, fc5 = st.columns([1, 1, 1, 1, 1])
        start_date = fc1.text_input("📅 조회 시작", "200101")
        end_date = fc2.text_input("📅 조회 종료", "261231")
        start_dt = pd.to_datetime(start_date, format="%y%m%d", errors="coerce")
        end_dt = pd.to_datetime(end_date, format="%y%m%d", errors="coerce")
        if pd.isna(start_dt): start_dt = pd.Timestamp("2000-01-01")
        if pd.isna(end_dt): end_dt = pd.Timestamp("2099-12-31")
        df_base = full_df[(full_df["매출일_dt"] >= start_dt) & (full_df["매출일_dt"] <= end_dt)].copy()
        selected_staff = fc3.multiselect("👤 담당자", sorted(df_base["담당자"].unique()) if not df_base.empty else [])
        df_staff_filtered = df_base[df_base["담당자"].isin(selected_staff)] if selected_staff else df_base.copy()
        all_clients = sorted(df_staff_filtered["거래처"].unique()) if not df_staff_filtered.empty else []
        client_options = ["전체 거래처"] + all_clients
        selected_client = fc4.selectbox("🏢 거래처", options=client_options, index=0)
        df_client_filtered = df_staff_filtered[df_staff_filtered["거래처"] == selected_client] if selected_client != "전체 거래처" else df_staff_filtered.copy()
        available_items = sorted(df_client_filtered["품목명"].unique()) if not df_client_filtered.empty else []
        selected_item = fc5.multiselect("📦 품목명", available_items)
        df_f = df_client_filtered[df_client_filtered["품목명"].isin(selected_item)] if selected_item else df_client_filtered.copy()
    raw_years = sorted(full_df["연도"].unique()) if "연도" in full_df.columns else ["2026"]
    years = sorted(raw_years, reverse=True)
    desired_order = [f"{y[2:]}년 {m}" for y in years for m in all_months]
    pivot_m_total = cached_get_yearly_monthly_pivot(df_base, all_months, years)
    client_item_qty_pivot = cached_client_item_qty_pivot(df_client_filtered, years, all_months)
    sales_p, qty_p, unit_price_p = cached_tab3_pivots(df_f, years, all_months)
    staff_pivot = cached_staff_pivot(df_base, desired_order)
    detail_cols = ["매출일_dt", "담당자", "거래처", "품목명", "출고량", "단가", "매출액"]
    df_detail = df_f[detail_cols].copy() if not df_f.empty else pd.DataFrame(columns=detail_cols)
    df_total_monthly = df_base.groupby(df_base["매출일_dt"].dt.to_period("M"))["매출액"].sum()
    if not df_total_monthly.empty:
        latest_period_total = df_total_monthly.index.max()
        cur_month_sales_total = df_total_monthly.loc[latest_period_total] * 1.1
        prev_period_total = latest_period_total - 1
        prev_month_sales_total = df_total_monthly.get(prev_period_total, 0.0) * 1.1
        mom_rate_total = ((cur_month_sales_total - prev_month_sales_total) / prev_month_sales_total * 100) if prev_month_sales_total > 0 else 0.0
        avg_monthly_sales_total = (df_total_monthly.mean() * 1.1)
        avg_rate_total = ((cur_month_sales_total - avg_monthly_sales_total) / avg_monthly_sales_total * 100) if avg_monthly_sales_total > 0 else 0.0
        latest_month_str_total = latest_period_total.strftime("%Y년 %m월")
    else:
        cur_month_sales_total = prev_month_sales_total = mom_rate_total = avg_monthly_sales_total = avg_rate_total = 0.0
        latest_month_str_total = "-"
    if not df_client_filtered.empty:
        df_client_monthly = df_client_filtered.groupby(df_client_filtered["매출일_dt"].dt.to_period("M"))["매출액"].sum() * 1.1
        latest_period_client = df_client_monthly.index.max()
        cur_month_sales_client = df_client_monthly.loc[latest_period_client]
        prev_period_client = latest_period_client - 1
        prev_month_sales_client = df_client_monthly.get(prev_period_client, 0.0)
        mom_rate_client = ((cur_month_sales_client - prev_month_sales_client) / prev_month_sales_client * 100) if prev_month_sales_client > 0 else 0.0
        avg_monthly_sales_client = df_client_monthly.mean()
        avg_rate_client = ((cur_month_sales_client - avg_monthly_sales_client) / avg_monthly_sales_client * 100) if avg_monthly_sales_client > 0 else 0.0
        latest_month_str_client = latest_period_client.strftime("%Y년 %m월")
    else:
        cur_month_sales_client = prev_month_sales_client = mom_rate_client = avg_monthly_sales_client = avg_rate_client = 0.0
        latest_month_str_client = "-"
else:
    cur_month_sales_total = prev_month_sales_total = mom_rate_total = avg_monthly_sales_total = avg_rate_total = 0.0
    latest_month_str_total = "-"
    cur_month_sales_client = prev_month_sales_client = mom_rate_client = avg_monthly_sales_client = avg_rate_client = 0.0
    latest_month_str_client = "-"
# 담당자만 반영한 채권(거래처 선택 무관) — 연체개월수 요약표용
staff_debt_df = pd.DataFrame()
filtered_debt_df = pd.DataFrame()
if not debt_df.empty:
    if selected_staff:
        valid_staff_clients = full_df[full_df["담당자"].isin(selected_staff)]["거래처"].unique()
        staff_debt_df = debt_df[debt_df["거래처"].isin(valid_staff_clients)].copy()
    else:
        staff_debt_df = debt_df.copy()
    filtered_debt_df = staff_debt_df.copy()
    if selected_client != "전체 거래처":
        filtered_debt_df = filtered_debt_df[filtered_debt_df["거래처"] == selected_client].copy()
client_addr_raw = addr_dict.get(selected_client, "등록된 주소 정보가 없습니다.")
if pd.isna(client_addr_raw) or str(client_addr_raw).strip().lower() == 'nan' or not str(client_addr_raw).strip():
    client_addr = "등록된 주소 정보가 없습니다."
else:
    client_addr = str(client_addr_raw)
st.sidebar.markdown("---")
st.sidebar.subheader("📥 엑셀 내보내기")
sheets_dict = {
    "연도별_월매출(만원)": (pivot_m_total, True),
    "거래처별_품목별사용량": (client_item_qty_pivot, True),
    "선택거래처_품목별_매출액(만원)": (sales_p * 1.1 / 10000, True),
    "선택거래처_품목별_출고량": (qty_p, True),
    "선택거래처_품목별_적용단가": (unit_price_p, True),
    "담당자별_매출(만원)": (staff_pivot, True),
    "상세거래내역": (df_detail, False),
}
if not filtered_debt_df.empty:
    sheets_dict["채권관리_현황"] = (filtered_debt_df, False)
excel_data = convert_dfs_to_excel(sheets_dict)
st.sidebar.download_button(
    label="📊 전체 분석 시트별 엑셀 다운로드",
    data=excel_data,
    file_name="통합영업분석_시트별보고서.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)
st.sidebar.subheader("📽️ PPT 내보내기")
_ppt_sig = (
    selected_client,
    tuple(sorted(selected_staff)),
    latest_update_str,
    latest_month_str_total,
)
if st.session_state.get("ppt_gen_sig") != _ppt_sig:
    st.session_state["ppt_ready_bytes"] = None
try:
    if st.sidebar.button("📽️ PPT 생성", key="gen_dashboard_ppt", use_container_width=True):
        with st.spinner("PPT 파일 생성 중..."):
            tot_sales_val_export = df_base["매출액"].sum() * 1.1 / 10000 if not df_base.empty else 0.0
            cur_sales_val_export = cur_month_sales_total / 10000
            pivot_m_client_export = cached_get_yearly_monthly_pivot(df_client_filtered, all_months, years)
            st.session_state["ppt_ready_bytes"] = convert_dashboard_to_ppt(
                latest_update_str,
                selected_client,
                tuple(selected_staff),
                latest_month_str_total,
                tot_sales_val_export,
                cur_sales_val_export,
                mom_rate_total,
                avg_rate_total,
                pivot_m_total,
                pivot_m_client_export,
                client_item_qty_pivot,
                sales_p,
                qty_p,
                staff_pivot,
                df_detail,
                filtered_debt_df,
                df_base,
                df_tank,
                df_vaporizer,
                df_integrated,
                tuple(all_months),
                tuple(years),
                tuple(target_items),
            )
            st.session_state["ppt_gen_sig"] = _ppt_sig
    if st.session_state.get("ppt_ready_bytes") and st.session_state.get("ppt_gen_sig") == _ppt_sig:
        st.sidebar.download_button(
            label="⬇️ PPT 다운로드",
            data=st.session_state["ppt_ready_bytes"],
            file_name=f"통합영업분석_전체탭_{datetime.date.today().strftime('%Y%m%d')}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
            key="dl_dashboard_ppt",
        )
    elif st.session_state.get("ppt_gen_sig") and st.session_state.get("ppt_gen_sig") != _ppt_sig:
        st.sidebar.caption("검색 조건이 변경되었습니다. PPT를 다시 생성해 주세요.")
except ImportError:
    st.sidebar.warning("PPT 내보내기: `pip install python-pptx kaleido` 설치 후 이용하세요.")
except Exception as exc:
    st.sidebar.error(f"PPT 생성 오류: {exc}")
tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, tab9 = st.tabs(
    [
        "📌 영업 종합 요약",
        "🏢 거래처 분석",
        "📦 품목 및 단가 분석",
        "👤 담당자 & 상세내역",
        "📌 채권 관리",
        "📍 카카오맵",
        "🏭 설비 재고 현황",
        "🛢️ 통합 탱크 재고",
        "📈 수익성 분석",
    ]
)
inject_sticky_tabs_script()
inject_ipad_plotly_controls()
# Tab 1: 📌 영업 종합 요약
with tab1:
    t1_c1, t1_c2 = st.columns([4, 1])
    t1_c1.markdown("<div class='sub-header dashboard-tab-panel-head'>📊 전체 영업 주요 실적 지표</div>", unsafe_allow_html=True)
    t1_c2.markdown(render_update_badge(latest_update_str), unsafe_allow_html=True)
    
    m1, m2, m3, m4 = st.columns(4)
    
    tot_sales_val = df_base["매출액"].sum() * 1.1 / 10000 if not df_base.empty else 0.0
    cur_sales_val = cur_month_sales_total / 10000
    m1.markdown(f"<div class='metric-box'><div class='metric-label'>총 누적 매출 (VAT포함)</div><div class='metric-value'>{tot_sales_val:,.0f} 만원</div></div>", unsafe_allow_html=True)
    m2.markdown(f"<div class='metric-box'><div class='metric-label'>최근 월 매출 ({latest_month_str_total})</div><div class='metric-value'>{cur_sales_val:,.0f} 만원</div></div>", unsafe_allow_html=True)
    m3.markdown(f"<div class='metric-box'><div class='metric-label'>전월 대비 (MoM)</div><div class='metric-value' style='color:{'#E11D48' if mom_rate_total < 0 else '#2563EB'};'>{mom_rate_total:+.0f}%</div></div>", unsafe_allow_html=True)
    m4.markdown(f"<div class='metric-box'><div class='metric-label'>월평균 대비 증감</div><div class='metric-value' style='color:{'#E11D48' if avg_rate_total < 0 else '#2563EB'};'>{avg_rate_total:+.0f}%</div></div>", unsafe_allow_html=True)
    st.markdown("<div class='sub-header dashboard-tab-panel-head'>📊 전체 영업 연도별 월 매출 추이</div>", unsafe_allow_html=True)
    col_left, col_right = st.columns([1, 1])
    
    with col_left:
        pivot_m_total_disp = get_display_df_with_sum(pivot_m_total, "연간 합계")
        st.dataframe(style_with_sum(pivot_m_total_disp, "{:,.0f}", "Blues", axis=None), use_container_width=True, height=460)
    with col_right:
        render_plotly_chart(
            create_stacked_bar_chart(pivot_m_total, title_text=""),
            use_container_width=True, key="tab1_total_chart"
        )
        
    st.markdown("---")
    st.markdown("<div class='sub-header dashboard-tab-panel-head'>📦 주요 4대 품목 상세 분석</div>", unsafe_allow_html=True)
    
    sel_col1, sel_col2 = st.columns([1, 1])
    with sel_col1:
        selected_target_item = st.radio("🔍 분석할 품목 선택", target_items, horizontal=True, key="overall_item_radio")
    with sel_col2:
        selected_metric = st.radio("📊 분석 지표 선택", ["매출액 (만원)", "출고량", "총매출 대비 비중 (%)"], horizontal=True, key="overall_metric_radio")
        
    item_pivot = cached_get_item_pivot(df_base, selected_target_item, selected_metric, all_months, years)
    
    i_col_left, i_col_right = st.columns([1, 1])
    with i_col_left:
        item_pivot_disp = get_display_df_with_sum(item_pivot, "연간 합계")
        if "비중" in selected_metric:
            st.dataframe(style_with_sum(item_pivot_disp, "{:,.1f}%", "Purples", axis=None), use_container_width=True, height=460)
            y_suf, y_fmt = "%", ",.1f"
        elif "출고량" in selected_metric:
            st.dataframe(style_with_sum(item_pivot_disp, "{:,.1f}", "Greens", axis=None), use_container_width=True, height=460)
            y_suf, y_fmt = " 천kg", ",.1f"
        else:
            st.dataframe(style_with_sum(item_pivot_disp, "{:,.0f}", "Blues", axis=None), use_container_width=True, height=460)
            y_suf, y_fmt = " 만원", ",.0f"
            
    with i_col_right:
        render_plotly_chart(
            create_stacked_bar_chart(
                item_pivot, 
                title_text="", 
                y_suffix=y_suf, 
                y_format=y_fmt
            ),
            use_container_width=True, key="tab1_item_chart"
        )
    st.markdown("---")
    st.markdown("<div class='sub-header dashboard-tab-panel-head'>🏭 업종별(분류별) 상세 분석</div>", unsafe_allow_html=True)
    
    if "업종" in df_base.columns:
        available_industries = sorted(list(df_base["업종"].unique()))
        
        if len(available_industries) == 1 and available_industries[0] == "미분류":
            st.info("💡 현재 모든 거래처가 '미분류' 상태입니다. 왼쪽 사이드바에서 '🏢 거래처 업종 분류 (CSV)' 파일을 업로드하시면 정확한 업종별 상세 분석이 가능합니다.")
            
        if available_industries:
            ind_col1, ind_col2 = st.columns([1, 1])
            with ind_col1:
                selected_industry = st.selectbox("🔍 분석할 업종(분류) 선택", available_industries, key="industry_selectbox")
            with ind_col2:
                selected_ind_metric = st.radio("📊 분석 지표 선택", ["매출액 (만원)", "출고량", "총매출 대비 비중 (%)"], horizontal=True, key="industry_metric_radio")
            
            ind_pivot = cached_get_industry_pivot(df_base, selected_industry, selected_ind_metric, all_months, years)
            
            i_col_left2, i_col_right2 = st.columns([1, 1])
            with i_col_left2:
                ind_pivot_disp = get_display_df_with_sum(ind_pivot, "연간 합계")
                if "비중" in selected_ind_metric:
                    st.dataframe(style_with_sum(ind_pivot_disp, "{:,.1f}%", "Purples", axis=None), use_container_width=True, height=460)
                    y_suf_i, y_fmt_i = "%", ",.1f"
                elif "출고량" in selected_ind_metric:
                    st.dataframe(style_with_sum(ind_pivot_disp, "{:,.0f}", "Greens", axis=None), use_container_width=True, height=460)
                    y_suf_i, y_fmt_i = "", ",.0f"
                else:
                    st.dataframe(style_with_sum(ind_pivot_disp, "{:,.0f}", "Blues", axis=None), use_container_width=True, height=460)
                    y_suf_i, y_fmt_i = " 만원", ",.0f"
            
            with i_col_right2:
                render_plotly_chart(
                    create_stacked_bar_chart(
                        ind_pivot, 
                        title_text="", 
                        y_suffix=y_suf_i, 
                        y_format=y_fmt_i
                    ),
                    use_container_width=True, key="tab1_industry_chart"
                )
            st.markdown("<br>", unsafe_allow_html=True)
            with st.expander(f"📂 [{selected_industry}] 소속 거래처 상세 데이터 파보기 (클릭하여 펼치기)", expanded=False):
                df_ind_detail = df_base[df_base["업종"] == selected_industry]
                
                if not df_ind_detail.empty:
                    ind_client_pivot = df_ind_detail.pivot_table(index="거래처", columns="연도", values="매출액", aggfunc="sum").fillna(0) * 1.1 / 10000
                    
                    avail_years_ind = sorted([y for y in ind_client_pivot.columns if str(y).isdigit()])
                    ind_client_pivot = ind_client_pivot.reindex(columns=avail_years_ind, fill_value=0)
                    ind_client_pivot["총 누적매출"] = ind_client_pivot.sum(axis=1)
                    ind_client_pivot = ind_client_pivot.sort_values(by="총 누적매출", ascending=False)
                    
                    ind_client_pivot_disp = get_display_df_with_sum(ind_client_pivot, "합계")
                    st.dataframe(style_with_sum(ind_client_pivot_disp, "{:,.0f}", "Blues", axis=None), use_container_width=True, height=280)
                    
                    st.markdown("<hr style='margin: 15px 0px; border-top: 1px dashed #E2E8F0;'>", unsafe_allow_html=True)
                    
                    ind_clients = sorted(df_ind_detail["거래처"].unique())
                    
                    c1, c2, c3 = st.columns([1, 1, 1])
                    with c1:
                        sel_ind_client = st.selectbox(f"🏢 [{selected_industry}] 거래처 선택", ind_clients, key="ind_client_sel")
                    with c2:
                        client_items = sorted(df_ind_detail[df_ind_detail["거래처"] == sel_ind_client]["품목명"].unique())
                        sel_ind_item = st.selectbox("📦 품목 선택", client_items if client_items else ["없음"], key="ind_item_sel")
                    with c3:
                        sel_ind_sub_metric = st.radio("📊 지표 선택", ["매출액 (만원)", "출고량", "총매출 대비 비중 (%)"], horizontal=True, key="ind_sub_metric")
                    
                    if client_items:
                        df_target_client = df_ind_detail[df_ind_detail["거래처"] == sel_ind_client]
                        sub_item_pivot = cached_get_item_pivot(df_target_client, sel_ind_item, sel_ind_sub_metric, all_months, years)
                        
                        sc1, sc2 = st.columns([1, 1])
                        with sc1:
                            sub_item_pivot_disp = get_display_df_with_sum(sub_item_pivot, "연간 합계")
                            if "비중" in sel_ind_sub_metric:
                                st.dataframe(style_with_sum(sub_item_pivot_disp, "{:,.1f}%", "Purples", axis=None), use_container_width=True, height=410)
                                y_suf_sub, y_fmt_sub = "%", ",.1f"
                            elif "출고량" in sel_ind_sub_metric:
                                if sel_ind_item in target_items:
                                    y_suf_sub, y_fmt_sub = " 천kg", ",.1f"
                                elif "LPG" in str(sel_ind_item).upper():
                                    y_suf_sub, y_fmt_sub = " kg", ",.0f"
                                else:
                                    y_suf_sub, y_fmt_sub = " 개(병)", ",.0f"
                                st.dataframe(style_with_sum(sub_item_pivot_disp, f"{{:{y_fmt_sub}}}", "Greens", axis=None), use_container_width=True, height=410)
                            else:
                                st.dataframe(style_with_sum(sub_item_pivot_disp, "{:,.0f}", "Blues", axis=None), use_container_width=True, height=410)
                                y_suf_sub, y_fmt_sub = " 만원", ",.0f"
                                
                        with sc2:
                            render_plotly_chart(
                                create_stacked_bar_chart(sub_item_pivot, title_text="", y_suffix=y_suf_sub, y_format=y_fmt_sub),
                                use_container_width=True, key="ind_client_sub_chart"
                            )
    st.markdown("---")
    st.markdown("<div class='sub-header dashboard-tab-panel-head'>🏆 당해년도 상위 30위 거래처 실적 (1월~12월) 및 업종 비중 · 월 헤더 클릭</div>", unsafe_allow_html=True)
    
    if not df_base.empty:
        current_year_str = str(df_base["연도"].max())
        df_curr_year = df_base[df_base["연도"] == current_year_str]
        if not df_curr_year.empty:
            latest_dt = df_curr_year["매출일_dt"].max()
            latest_m = latest_dt.strftime("%m월")
            # 월 선택 → 해당 월 순위·도넛만 변경 (집계 로직 무손실)
            inject_top30_month_bridge()
            if "top30_month" not in st.session_state:
                st.session_state["top30_month"] = latest_m
            try:
                qp_m = st.query_params.get("top30_month", None)
                if isinstance(qp_m, list):
                    qp_m = qp_m[0] if qp_m else None
                if qp_m is not None:
                    qp_m = urllib.parse.unquote(str(qp_m))
                    if qp_m in all_months:
                        st.session_state["top30_month"] = qp_m
            except Exception:
                pass
            if st.session_state.get("top30_month") not in all_months:
                st.session_state["top30_month"] = latest_m
            rank_m = st.session_state["top30_month"]
            rank_month_label = f"{current_year_str}년 {rank_m}"
            # top30-section-flag: iPad CSS 스택용 마커 (맥 레이아웃/데이터 무손실)
            with st.container():
                st.markdown("<div class='top30-section-flag' style='display:none'></div>", unsafe_allow_html=True)
                p_col1, p_col2 = st.columns([1.2, 0.8])
                with p_col1:
                    st.markdown(
                        f"<div style='font-size: 14px; font-weight: 600; color: #334155; margin-bottom: 10px;'>"
                        f"🥇 [{rank_month_label} 기준] 상위 30위 거래처 월별 실적 (VAT포함, 만원)"
                        f"<span style='font-size:12px;font-weight:500;color:#64748B;margin-left:8px;'>"
                        f"{'← 월 선택 또는 표 헤더 클릭' if is_touch_ui() else '← 월 버튼 또는 표 헤더 클릭'}"
                        f"</span></div>",
                        unsafe_allow_html=True,
                    )
                    # 맥: 이전 가로 월버튼 / iPad(touch_ui=1): 현재 selectbox
                    if is_touch_ui():
                        _m_idx = all_months.index(rank_m) if rank_m in all_months else 0
                        picked_m = st.selectbox(
                            "기준 월",
                            all_months,
                            index=_m_idx,
                            key="top30_month_select",
                            label_visibility="collapsed",
                        )
                        if picked_m != st.session_state.get("top30_month"):
                            st.session_state["top30_month"] = picked_m
                            try:
                                st.query_params["top30_month"] = picked_m
                            except Exception:
                                pass
                            st.rerun()
                    else:
                        m_cols = st.columns(12, gap="small")
                        _clicked_m = None
                        for _i, _m in enumerate(all_months):
                            with m_cols[_i]:
                                if st.button(
                                    _m,
                                    key=f"top30_month_btn_{_m}",
                                    type="primary" if _m == rank_m else "secondary",
                                    use_container_width=True,
                                ):
                                    _clicked_m = _m
                        if _clicked_m:
                            st.session_state["top30_month"] = _clicked_m
                            try:
                                st.query_params["top30_month"] = _clicked_m
                            except Exception:
                                pass
                            st.rerun()
                    rank_m = st.session_state["top30_month"]
                    rank_month_label = f"{current_year_str}년 {rank_m}"
                    
                    pvt_curr = df_curr_year.pivot_table(index="거래처", columns="월", values="매출액", aggfunc="sum").fillna(0) * 1.1 / 10000
                    pvt_curr = pvt_curr.reindex(columns=all_months, fill_value=0)
                    
                    if rank_m in pvt_curr.columns:
                        pvt_curr = pvt_curr.sort_values(by=rank_m, ascending=False)
                    top30_pvt = pvt_curr.head(30).reset_index()
                    
                    top30_pvt.index = range(1, len(top30_pvt) + 1)
                    
                    top30_pvt_disp = get_display_df_with_sum(top30_pvt, sum_label="합계", text_cols=["거래처"])
                    
                    fmt_dict = {m: "{:,.0f}" for m in all_months}
                    styled_top30 = style_with_sum(top30_pvt_disp, fmt_dict, "Blues", subset_cols=all_months, axis=0)
                    
                    if rank_m in all_months:
                        styled_top30 = styled_top30.apply(
                            lambda s: ['color: #B91C1C; font-weight: bold; background-color: #DBEAFE;'] * len(s),
                            subset=[rank_m],
                            axis=0
                        )
                    
                    render_frozen_styler_html(
                        styled_top30,
                        height=450,
                        freeze_left_n=2,
                        freeze_widths=[44, 160],
                        clickable_cols=all_months,
                        query_param="top30_month",
                        active_col=rank_m,
                    )
                with p_col2:
                    st.markdown(
                        f"<div class='top30-donut-title' style='font-size: 14px; font-weight: 600; color: #334155; margin-bottom: 10px;'>"
                        f"🍩 [{rank_month_label}] 업종별 매출 비중</div>",
                        unsafe_allow_html=True,
                    )
                    
                    df_rank_month = df_curr_year[df_curr_year["월"] == rank_m]
                    ind_sales = df_rank_month.groupby("업종")["매출액"].sum().reset_index()
                    ind_sales = ind_sales[ind_sales["매출액"] > 0]
                    
                    if ind_sales.empty:
                        st.info(f"{rank_month_label} 업종별 매출 데이터가 없습니다.")
                    else:
                        fig_donut = px.pie(
                            ind_sales, 
                            values='매출액', 
                            names='업종', 
                            hole=0.4,
                            color_discrete_sequence=px.colors.qualitative.Pastel
                        )
                        fig_donut.update_traces(textposition='inside', textinfo='percent+label')
                        fig_donut.update_layout(
                            showlegend=True,
                            legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5),
                            margin=dict(l=10, r=10, t=10, b=10),
                            height=450,
                            autosize=True,
                        )
                        render_plotly_chart(
                            fig_donut,
                            use_container_width=True,
                            key=f"top30_donut_{rank_m}",
                        )
        else:
            st.info("당해년도 매출 데이터가 없습니다.")
    st.markdown("---")
    if not df_base.empty and '매출일_dt' in df_base.columns:
        latest_period = df_base["매출일_dt"].dt.to_period("M").max()
        prev_period = latest_period - 1
        
        curr_m_label = latest_period.strftime("%m월")
        prev_m_label = prev_period.strftime("%m월")
        
        st.markdown(f"<div class='sub-header dashboard-tab-panel-head'>📊 전월 대비 실적 증감 및 신규/이탈 분석 ({prev_m_label} vs {curr_m_label})</div>", unsafe_allow_html=True)
        
        df_prev = df_base[df_base["매출일_dt"].dt.to_period("M") == prev_period].groupby("거래처")["매출액"].sum().reset_index().rename(columns={"매출액": f"{prev_m_label} 매출"})
        df_curr = df_base[df_base["매출일_dt"].dt.to_period("M") == latest_period].groupby("거래처")["매출액"].sum().reset_index().rename(columns={"매출액": f"{curr_m_label} 매출"})
        df_diff = pd.merge(df_prev, df_curr, on="거래처", how="outer").fillna(0)
        df_diff["매출 증감액"] = df_diff[f"{curr_m_label} 매출"] - df_diff[f"{prev_m_label} 매출"]
        
        df_diff[[f"{prev_m_label} 매출", f"{curr_m_label} 매출", "매출 증감액"]] = (df_diff[[f"{prev_m_label} 매출", f"{curr_m_label} 매출", "매출 증감액"]] * 1.1) / 10000
        top_gains = df_diff[(df_diff[f"{prev_m_label} 매출"] > 0) & (df_diff["매출 증감액"] > 0)].sort_values(by="매출 증감액", ascending=False).head(10)
        top_drops = df_diff[(df_diff[f"{prev_m_label} 매출"] > 0) & (df_diff[f"{curr_m_label} 매출"] > 0) & (df_diff["매출 증감액"] < 0)].sort_values(by="매출 증감액", ascending=True).head(10)
        new_clients = df_diff[(df_diff[f"{prev_m_label} 매출"] == 0) & (df_diff[f"{curr_m_label} 매출"] > 0)].sort_values(by=f"{curr_m_label} 매출", ascending=False)
        lost_clients = df_diff[(df_diff[f"{prev_m_label} 매출"] > 0) & (df_diff[f"{curr_m_label} 매출"] == 0)].sort_values(by=f"{prev_m_label} 매출", ascending=False)
        mom_view = st.radio(
            "전월 대비 분석 보기",
            ["🚀 상승 Top 10", "📉 하락 Top 10", "🎉 신규/재개 거래처", "⚠️ 미거래/이탈 의심"],
            horizontal=True,
            key="tab1_mom_view",
            label_visibility="collapsed",
        )
        if mom_view == "🚀 상승 Top 10":
            st.markdown(f"**🔥 기존 거래처 중 매출이 가장 많이 [상승]한 10곳 (단위: 만원, VAT 포함)**")
            if top_gains.empty:
                st.info("해당 조건에 맞는 상승 거래처가 없습니다.")
            else:
                st.dataframe(
                    top_gains.style.format({f"{prev_m_label} 매출": "{:,.0f}", f"{curr_m_label} 매출": "{:,.0f}", "매출 증감액": "{:,.0f}"})
                    .apply(lambda s: ['color: #2563EB; font-weight: bold;' if v > 0 else '' for v in s], subset=['매출 증감액']),
                    use_container_width=True, hide_index=True, height=min(420, 38 + len(top_gains) * 35)
                )
        elif mom_view == "📉 하락 Top 10":
            st.markdown(f"**📉 기존 거래처 중 매출이 가장 많이 [하락]한 10곳 (단위: 만원, VAT 포함)**")
            if top_drops.empty:
                st.info("해당 조건에 맞는 하락 거래처가 없습니다.")
            else:
                st.dataframe(
                    top_drops.style.format({f"{prev_m_label} 매출": "{:,.0f}", f"{curr_m_label} 매출": "{:,.0f}", "매출 증감액": "{:,.0f}"})
                    .apply(lambda s: ['color: #B91C1C; font-weight: bold;' if v < 0 else '' for v in s], subset=['매출 증감액']),
                    use_container_width=True, hide_index=True, height=min(420, 38 + len(top_drops) * 35)
                )
        elif mom_view == "🎉 신규/재개 거래처":
            st.markdown(f"**🎉 {prev_m_label}엔 거래가 없었으나 {curr_m_label}에 새롭게 매출이 발생한 곳 (총 {len(new_clients)}곳, 단위: 만원, VAT 포함)**")
            if new_clients.empty:
                st.info("신규/재개 거래처가 없습니다.")
            else:
                st.dataframe(
                    new_clients[["거래처", f"{curr_m_label} 매출"]].style.format({f"{curr_m_label} 매출": "{:,.0f}"}),
                    use_container_width=True, hide_index=True, height=min(480, 38 + len(new_clients) * 35)
                )
        else:
            st.markdown(f"**⚠️ {prev_m_label}엔 매출이 있었으나 {curr_m_label}엔 거래가 없는 곳 (총 {len(lost_clients)}곳, 단위: 만원, VAT 포함)**")
            if lost_clients.empty:
                st.info("미거래/이탈 의심 거래처가 없습니다.")
            else:
                st.dataframe(
                    lost_clients[["거래처", f"{prev_m_label} 매출"]].style.format({f"{prev_m_label} 매출": "{:,.0f}"}),
                    use_container_width=True, hide_index=True, height=min(480, 38 + len(lost_clients) * 35)
                )
# Tab 2: 🏢 거래처 분석
with tab2:
    t2_c1, t2_c2 = st.columns([4, 1])
    t2_c1.markdown(f"<div class='sub-header dashboard-tab-panel-head'>🏢 [{selected_client}] 영업 실적 및 요약</div>", unsafe_allow_html=True)
    t2_c2.markdown(render_update_badge(latest_update_str), unsafe_allow_html=True)
    
    if "show_corp_info" not in st.session_state:
        st.session_state.show_corp_info = False
    # 가로 넓은 직사각형 버튼 — 글씨 한 줄로 박스 안에
    with st.container(key="tab2_action_btns"):
        btn_c1, btn_c2, btn_c3 = st.columns([2.4, 2.0, 1.8], gap="medium")
        with btn_c1:
            if st.button(
                "📝 macOS 메모에서 노트 열기/생성",
                key="btn_notes",
                width="stretch",
            ):
                open_macos_notes_folder(selected_client, dart_api_key, df_integrated)
        with btn_c2:
            btn_label = "🏢 기업정보 닫기" if st.session_state.show_corp_info else "🏢 기업 기본/재무정보 보기"
            if st.button(btn_label, key="btn_dart_info", width="stretch"):
                st.session_state.show_corp_info = not st.session_state.show_corp_info
                st.rerun()
        with btn_c3:
            if client_addr != "등록된 주소 정보가 없습니다.":
                kakao_url = f"https://map.kakao.com/link/search/{urllib.parse.quote(client_addr)}"
                st.link_button("🗺️ 카카오맵에서 주소 보기", kakao_url, width="stretch")
            else:
                st.button(
                    "🗺️ 카카오맵에서 주소 보기",
                    disabled=True,
                    key="btn_kakao_disabled",
                    width="stretch",
                )
    if st.session_state.show_corp_info:
        with st.spinner("DART 및 네이버 기업 정보를 불러오는 중..."):
            c_info = get_company_info_hybrid(selected_client, dart_api_key)
            st.markdown(f"""
            - **출처:** {c_info['source']}
            - **대표자:** {c_info['ceo']}
            - **업종:** {c_info['industry']}
            - **매출액:** {c_info['revenue']}
            - **영업이익:** {c_info['profit']}
            """)
    m1, m2, m3, m4 = st.columns(4)
    tot_sales_c = df_client_filtered["매출액"].sum() * 1.1 / 10000 if not df_client_filtered.empty else 0.0
    
    cur_sales_c = cur_month_sales_client / 10000
    m1.markdown(f"<div class='metric-box'><div class='metric-label'>총 누적 매출 (VAT포함)</div><div class='metric-value'>{tot_sales_c:,.0f} 만원</div></div>", unsafe_allow_html=True)
    m2.markdown(f"<div class='metric-box'><div class='metric-label'>최근 월 매출 ({latest_month_str_client})</div><div class='metric-value'>{cur_sales_c:,.0f} 만원</div></div>", unsafe_allow_html=True)
    m3.markdown(f"<div class='metric-box'><div class='metric-label'>전월 대비 (MoM)</div><div class='metric-value' style='color:{'#E11D48' if mom_rate_client < 0 else '#2563EB'};'>{mom_rate_client:+.0f}%</div></div>", unsafe_allow_html=True)
    m4.markdown(f"<div class='metric-box'><div class='metric-label'>월평균 대비 증감</div><div class='metric-value' style='color:{'#E11D48' if avg_rate_client < 0 else '#2563EB'};'>{avg_rate_client:+.0f}%</div></div>", unsafe_allow_html=True)
    if not df_client_filtered.empty:
        pivot_m_client = cached_get_yearly_monthly_pivot(df_client_filtered, all_months, years)
        cl, cr = st.columns([1, 1])
        with cl:
            pivot_m_client_disp = get_display_df_with_sum(pivot_m_client, "연간 합계")
            st.dataframe(style_with_sum(pivot_m_client_disp, "{:,.0f}", "Blues", axis=None), use_container_width=True, height=460)
        with cr:
            render_plotly_chart(
                create_stacked_bar_chart(pivot_m_client, title_text=""),
                use_container_width=True, key="tab2_client_total_chart"
            )
        st.markdown("---")
        st.markdown(f"<div class='sub-header dashboard-tab-panel-head'>📦 [{selected_client}] 품목별 상세 분석</div>", unsafe_allow_html=True)
        
        client_available_items = sorted(df_client_filtered["품목명"].unique())
        if client_available_items:
            item_ratios = {}
            if not df_client_filtered.empty:
                latest_dt_c = df_client_filtered["매출일_dt"].max()
                if pd.notnull(latest_dt_c):
                    latest_ym = latest_dt_c.strftime("%Y-%m")
                    df_cm = df_client_filtered[df_client_filtered["매출일_dt"].dt.strftime("%Y-%m") == latest_ym]
                    tot_sales = df_cm["매출액"].sum()
                    if tot_sales > 0:
                        grp = df_cm.groupby("품목명")["매출액"].sum()
                        for item, val in grp.items():
                            item_ratios[item] = (val / tot_sales) * 100
            def format_item_with_ratio(item_name):
                pct = item_ratios.get(item_name, 0.0)
                return f"{item_name} (당월 {pct:.1f}%)"
            sel_col1_c, sel_col2_c = st.columns([1, 1])
            with sel_col1_c:
                selected_target_item_c = st.selectbox(
                    "🔍 분석할 품목 선택 (전체 거래 품목)", 
                    options=client_available_items, 
                    format_func=format_item_with_ratio,
                    key="client_item_selectbox"
                )
            with sel_col2_c:
                selected_metric_c = st.radio("📊 분석 지표 선택", ["매출액 (만원)", "출고량", "총매출 대비 비중 (%)"], horizontal=True, key="client_metric_radio")
                
            client_item_pivot = cached_get_item_pivot(df_client_filtered, selected_target_item_c, selected_metric_c, all_months, years)
            
            i_col_left_c, i_col_right_c = st.columns([1, 1])
            with i_col_left_c:
                client_item_pivot_disp = get_display_df_with_sum(client_item_pivot, "연간 합계")
                if "비중" in selected_metric_c:
                    st.dataframe(style_with_sum(client_item_pivot_disp, "{:,.1f}%", "Purples", axis=None), use_container_width=True, height=460)
                    y_suf_c, y_fmt_c = "%", ",.1f"
                elif "출고량" in selected_metric_c:
                    if selected_target_item_c in target_items:
                        y_suf_c, y_fmt_c = " 천kg", ",.1f"
                    elif "LPG" in str(selected_target_item_c).upper():
                        y_suf_c, y_fmt_c = " kg", ",.0f"
                    else:
                        y_suf_c, y_fmt_c = " 개(병)", ",.0f"
                    st.dataframe(style_with_sum(client_item_pivot_disp, f"{{:{y_fmt_c}}}", "Greens", axis=None), use_container_width=True, height=460)
                else:
                    st.dataframe(style_with_sum(client_item_pivot_disp, "{:,.0f}", "Blues", axis=None), use_container_width=True, height=460)
                    y_suf_c, y_fmt_c = " 만원", ",.0f"
                    
            with i_col_right_c:
                render_plotly_chart(
                    create_stacked_bar_chart(
                        client_item_pivot, 
                        title_text="", 
                        y_suffix=y_suf_c, 
                        y_format=y_fmt_c
                    ),
                    use_container_width=True, key="tab2_client_item_chart"
                )
# Tab 3: 📦 품목 및 단가 분석
with tab3:
    t3_c1, t3_c2 = st.columns([4, 1])
    t3_c1.markdown(f"<div class='sub-header dashboard-tab-panel-head'>📦 [{selected_client}] 품목별 실적 분석</div>", unsafe_allow_html=True)
    t3_c2.markdown(render_update_badge(latest_update_str), unsafe_allow_html=True)
    
    latest_dt_overall = df_base["매출일_dt"].max() if not df_base.empty else None
    target_month_col = latest_dt_overall.strftime("%y년 %m월") if pd.notnull(latest_dt_overall) else None
    avail_years_short = [y[2:] for y in years]
    current_year_short = str(df_base["연도"].max())[2:] if not df_base.empty else (avail_years_short[0] if avail_years_short else "26")
    
    selected_detail_years = st.multiselect(
        "📅 월별 상세 내역을 펼쳐볼 연도 선택 (단가표 제외)",
        options=avail_years_short,
        default=[current_year_short] if current_year_short in avail_years_short else avail_years_short[:1],
        format_func=lambda x: f"20{x}년",
        key="tab3_detail_years",
    )
        
    sales_p_filtered, qty_p_filtered = cached_filter_tab3_year_columns(
        sales_p,
        qty_p,
        tuple(sorted(selected_detail_years)),
        tuple(avail_years_short),
        tuple(all_months),
    )
    st.markdown("<div style='font-size: 14px; font-weight: 600; color: #334155; margin-bottom: 10px;'>1️⃣ 매출액 (VAT 포함, 만원)</div>", unsafe_allow_html=True)
    render_tab3_dataframe_table(
        sales_p_filtered, "{:,.0f}", target_month_col, key_prefix="tab3_sales", table_kind="sales"
    )
    # 상단 담당자·품목 선택 시 → 품목별 거래처 매출 상세 (클릭 펼침)
    if selected_staff and selected_item:
        render_tab3_item_client_expanders(
            df_f,
            list(selected_item),
            "sales",
            years,
            all_months,
            selected_detail_years,
        )
    elif selected_item and not selected_staff:
        st.caption("💡 거래처별 상세를 보려면 상단 고정바에서 담당자를 먼저 선택한 뒤 품목을 선택하세요.")
        
    st.markdown("<div style='font-size: 14px; font-weight: 600; color: #334155; margin-bottom: 10px;'>2️⃣ 출고량</div>", unsafe_allow_html=True)
    render_tab3_dataframe_table(
        qty_p_filtered, "{:,.0f}", target_month_col, key_prefix="tab3_qty", table_kind="qty"
    )
    if selected_staff and selected_item:
        render_tab3_item_client_expanders(
            df_f,
            list(selected_item),
            "qty",
            years,
            all_months,
            selected_detail_years,
        )
        
    st.markdown("<div style='font-size: 14px; font-weight: 600; color: #334155; margin-bottom: 10px;'>3️⃣ 적용 단가 (실제 원본 단가) - 전체 기간 월별 고정 표시</div>", unsafe_allow_html=True)
    _price_sort_order = list(sales_p.index) if not sales_p.empty else None
    render_tab3_dataframe_table(
        unit_price_p,
        "{:,.0f}",
        target_month_col,
        key_prefix="tab3_price",
        table_kind="price",
        sort_order=_price_sort_order,
    )
    
# Tab 4: 👤 담당자 & 상세내역
with tab4:
    t4_c1, t4_c2 = st.columns([4, 1])
    t4_c1.markdown("<div class='sub-header dashboard-tab-panel-head'>👤 담당자별 월 매출 실적 (만원)</div>", unsafe_allow_html=True)
    t4_c2.markdown(render_update_badge(latest_update_str), unsafe_allow_html=True)
    
    if not staff_pivot.empty:
        format_dict = {col: "{:,.0f}" for col in staff_pivot.columns if col != "매출 비중 (%)"}
        format_dict["매출 비중 (%)"] = "{:,.1f}%"
        
        monthly_cols = [c for c in staff_pivot.columns if c not in ["매출 비중 (%)", "총 매출 합계 (만원)"]]
        
        styled_staff = (
            staff_pivot.style.format(format_dict)
            .background_gradient(cmap="Purples", subset=["매출 비중 (%)"])
            .background_gradient(cmap="Oranges", subset=["총 매출 합계 (만원)"])
            .background_gradient(cmap="Blues", subset=monthly_cols)
        )
        st.dataframe(styled_staff, use_container_width=True, height=350)
    st.markdown("<div class='sub-header dashboard-tab-panel-head'>🏆 담당자별 거래처 매출 순위 (당해년도)</div>", unsafe_allow_html=True)
    if not df_base.empty:
        current_year = str(df_base["연도"].max())
        all_staffs = sorted(df_base["담당자"].unique())
        
        sel_staff = st.selectbox("👤 순위를 조회할 담당자 선택", all_staffs, key="ranking_staff_select")
        ranking_pivot = cached_ranking_pivot(df_base, current_year, sel_staff, all_months)
        
        if not ranking_pivot.empty:
            r_col1, r_col2 = st.columns([1.2, 1])
            
            with r_col1:
                st.dataframe(
                    ranking_pivot.style.format("{:,.0f}")
                    .background_gradient(cmap="Blues", subset=all_months)
                    .background_gradient(cmap="Oranges", subset=["당해 누적 (만원)"]),
                    use_container_width=True, height=380
                )
                
            with r_col2:
                top_clients = ranking_pivot.head(10).sort_values(by="당해 누적 (만원)", ascending=True)
                
                fig_ranking = px.bar(
                    top_clients.reset_index(),
                    x="당해 누적 (만원)",
                    y="거래처",
                    orientation='h',
                    text="당해 누적 (만원)",
                    color="당해 누적 (만원)",
                    color_continuous_scale="Blues"
                )
                fig_ranking.update_traces(texttemplate='%{text:,.0f}', textposition='outside')
                fig_ranking.update_layout(
                    xaxis_title=None,
                    yaxis_title=None,
                    margin=dict(l=10, r=40, t=10, b=10),
                    plot_bgcolor='rgba(0,0,0,0)',
                    paper_bgcolor='rgba(0,0,0,0)',
                    coloraxis_showscale=False,
                    height=380
                )
                render_plotly_chart(fig_ranking, use_container_width=True, key=f"ranking_chart_{sel_staff}")
        else:
            st.info(f"💡 {current_year}년에 선택한 담당자({sel_staff})의 거래처 매출 실적 데이터가 없습니다.")
    if not df_base.empty:
        with st.expander("⚠️ 담당자 미지정 신규/누락 거래처 (직접 지정) 열기/닫기", expanded=True):
            unassigned_df = df_base[df_base["담당자"] == "미지정"]
            
            custom_staffs = ["가스코아산"]
            existing_staffs = [s for s in full_df["담당자"].unique() if s != "미지정"]
            combined_staffs = sorted(list(set(existing_staffs + custom_staffs)))
            
            all_staff_options = ["미지정"] + combined_staffs
            
            if not unassigned_df.empty:
                unassigned_summary = unassigned_df.groupby("거래처").agg(
                    최근매출일=("매출일_dt", "max"),
                    총매출액_만원=("매출액", lambda x: sum(x) * 1.1 / 10000)
                ).reset_index()
                
                unassigned_summary = unassigned_summary.sort_values(by="총매출액_만원", ascending=False)
                unassigned_summary["최근매출일"] = unassigned_summary["최근매출일"].dt.strftime("%Y-%m-%d")
                unassigned_summary["담당자지정"] = "미지정"
                
                st.warning(f"💡 자동 추론으로도 담당자를 찾을 수 없는 거래처가 총 **{len(unassigned_summary)}곳** 있습니다. 표의 **'담당자지정'** 열을 클릭하여 담당자를 선택하고 아래 저장 버튼을 누르세요.")
                
                edited_unassigned = st.data_editor(
                    unassigned_summary,
                    column_config={
                        "거래처": st.column_config.TextColumn("거래처", disabled=True),
                        "최근매출일": st.column_config.TextColumn("최근매출일", disabled=True),
                        "총매출액_만원": st.column_config.NumberColumn("총매출액(만원)", disabled=True, format="%d"),
                        "담당자지정": st.column_config.SelectboxColumn(
                            "👤 담당자 지정 (클릭하여 변경)",
                            help="이 거래처의 담당자를 선택하세요.",
                            options=all_staff_options,
                            required=True
                        )
                    },
                    use_container_width=True,
                    hide_index=True,
                    key="unassigned_editor"
                )
                
                if st.button("💾 변경된 담당자 저장 및 전체 대시보드 적용", type="primary"):
                    changed_rows = edited_unassigned[edited_unassigned["담당자지정"] != "미지정"]
                    if not changed_rows.empty:
                        manual_map_path = os.path.join(CACHE_DIR, "manual_staff_mapping.csv")
                        if os.path.exists(manual_map_path):
                            try:
                                existing_map = pd.read_csv(manual_map_path).set_index("거래처")["담당자"].to_dict()
                            except:
                                existing_map = {}
                        else:
                            existing_map = {}
                            
                        for _, row in changed_rows.iterrows():
                            existing_map[row["거래처"]] = row["담당자지정"]
                            
                        save_df = pd.DataFrame(list(existing_map.items()), columns=["거래처", "담당자"])
                        save_df.to_csv(manual_map_path, index=False, encoding="utf-8-sig")
                        
                        st.success("✅ 담당자 지정이 완료되었습니다! 대시보드를 새로고침합니다.")
                        load_uploaded_files_from_bytes.clear() 
                        st.rerun()
            else:
                st.success("🎉 모든 거래처에 담당자가 완벽하게 지정되어 있습니다!")
                
        with st.expander("🔄 이미 지정된 기존 거래처 담당자 수정/강제 변경하기 열기/닫기"):
            assigned_df = df_base[df_base["담당자"] != "미지정"]
            if not assigned_df.empty:
                assigned_summary = assigned_df.groupby("거래처").agg(
                    현재담당자=("담당자", "first"),
                    최근매출일=("매출일_dt", "max")
                ).reset_index()
                
                assigned_summary["새담당자변경"] = assigned_summary["현재담당자"]
                assigned_summary["최근매출일"] = assigned_summary["최근매출일"].dt.strftime("%Y-%m-%d")
                
                all_staffs_for_edit_assign = combined_staffs + ["미지정"]
                
                st.info("💡 잘못 지정된 거래처나 인수인계된 거래처의 담당자를 새롭게 변경할 수 있습니다.")
                edited_assigned = st.data_editor(
                    assigned_summary,
                    column_config={
                        "거래처": st.column_config.TextColumn("거래처", disabled=True),
                        "현재담당자": st.column_config.TextColumn("현재 담당자", disabled=True),
                        "최근매출일": st.column_config.TextColumn("최근매출일", disabled=True),
                        "새담당자변경": st.column_config.SelectboxColumn(
                            "👤 새 담당자로 변경 (클릭)",
                            options=all_staffs_for_edit_assign,
                            required=True
                        )
                    },
                    use_container_width=True,
                    hide_index=True,
                    key="assigned_editor"
                )
                
                if st.button("💾 변경된 기존 거래처 담당자 저장", type="primary", key="save_assigned_btn"):
                    changed_assigned = edited_assigned[edited_assigned["새담당자변경"] != edited_assigned["현재담당자"]]
                    if not changed_assigned.empty:
                        manual_map_path = os.path.join(CACHE_DIR, "manual_staff_mapping.csv")
                        if os.path.exists(manual_map_path):
                            try:
                                existing_map = pd.read_csv(manual_map_path).set_index("거래처")["담당자"].to_dict()
                            except:
                                existing_map = {}
                        else:
                            existing_map = {}
                            
                        for _, row in changed_assigned.iterrows():
                            existing_map[row["거래처"]] = row["새담당자변경"]
                            
                        save_df = pd.DataFrame(list(existing_map.items()), columns=["거래처", "담당자"])
                        save_df.to_csv(manual_map_path, index=False, encoding="utf-8-sig")
                        
                        st.success("✅ 담당자 변경이 완료되었습니다! 대시보드를 새로고침합니다.")
                        load_uploaded_files_from_bytes.clear() 
                        st.rerun()
    st.markdown("<div class='sub-header dashboard-tab-panel-head'>📋 거래 상세 내역 (최신순 800건)</div>", unsafe_allow_html=True)
    if not df_detail.empty:
        view_detail_df = df_detail.sort_values(by="매출일_dt", ascending=False).head(800)
        
        styled_detail = (
            view_detail_df
            .style.format({
                "출고량": "{:,.0f}",
                "단가": "{:,.0f}",
                "매출액": "{:,.0f}",
                "매출일_dt": lambda t: t.strftime("%Y-%m-%d") if pd.notnull(t) else ""
            })
            .background_gradient(subset=["매출액"], cmap="Blues")
        )
        st.dataframe(styled_detail, use_container_width=True, height=600, hide_index=True)
# Tab 5: 📌 채권 관리
with tab5:
    latest_month = None
    if not filtered_debt_df.empty:
        # 데이터가 0(없는) 달 제거 로직 추가
        numeric_cols_temp = [c for c in filtered_debt_df.columns if c not in ["거래처", "구분"]]
        valid_numeric_cols = [c for c in numeric_cols_temp if filtered_debt_df[c].abs().sum() > 0]
        
        filtered_debt_df = filtered_debt_df[["거래처", "구분"] + valid_numeric_cols]
        numeric_cols_debt = valid_numeric_cols
        
        if numeric_cols_debt:
            latest_month = numeric_cols_debt[-1]
            
    debt_update_str = f"{latest_month} 기준" if latest_month else "데이터 없음"
    
    t5_c1, t5_c2 = st.columns([4, 1])
    t5_c1.markdown("<div class='sub-header dashboard-tab-panel-head'>💰 채권(외상대금) 관리 현황 및 연령 분석</div>", unsafe_allow_html=True)
    t5_c2.markdown(render_update_badge(debt_update_str), unsafe_allow_html=True)
    
    if not debt_df.empty:
        if not filtered_debt_df.empty:
            numeric_cols = [c for c in filtered_debt_df.columns if c not in ["거래처", "구분"]]
            
            total_outstanding = 0
            warning_count = 0
            
            if latest_month:
                u_clients = filtered_debt_df['거래처'].unique()
                for uc in u_clients:
                    c_mask = filtered_debt_df['거래처'] == uc
                    b_val = filtered_debt_df.loc[c_mask & (filtered_debt_df['구분'] == '잔액'), latest_month].sum()
                    s_val = filtered_debt_df.loc[c_mask & (filtered_debt_df['구분'] == '매출'), latest_month].sum()
                    
                    total_outstanding += max(0, b_val)
                    if b_val > 0 and b_val > s_val:
                        warning_count += 1
                        
            m1, m2 = st.columns(2)
            m1.markdown(f"<div class='metric-box'><div class='metric-label'>총 미수금 잔액 ({latest_month} 기준)</div><div class='metric-value'>{total_outstanding:,.0f} 원</div></div>", unsafe_allow_html=True)
            m2.markdown(f"<div class='metric-box'><div class='metric-label'>매출 초과 악성/지연 채권 업체 수</div><div class='metric-value' style='color:#E11D48;'>{warning_count} 곳</div></div>", unsafe_allow_html=True)
            
            st.markdown("<br>", unsafe_allow_html=True)
            summary_rows = []
            for gubun in ["이월", "익월", "매출", "수금", "잔액", "합계"]: 
                if gubun in filtered_debt_df["구분"].values:
                    sum_vals = filtered_debt_df[filtered_debt_df["구분"] == gubun][numeric_cols].sum()
                    row_data = {"거래처": "📌 [전체 합계]", "구분": gubun}
                    for col in numeric_cols:
                        row_data[col] = sum_vals[col]
                    summary_rows.append(row_data)
            
            if summary_rows:
                summary_df = pd.DataFrame(summary_rows)
                disp_debt = pd.concat([filtered_debt_df, summary_df], ignore_index=True)
            else:
                disp_debt = filtered_debt_df.copy()
            
            gubun_order = {"이월": 1, "매출": 2, "수금": 3, "잔액": 4, "합계": 5}
            disp_debt["구분순위"] = disp_debt["구분"].map(gubun_order).fillna(99)
            client_order = {c: i for i, c in enumerate(disp_debt["거래처"].unique())}
            disp_debt["거래처순위"] = disp_debt["거래처"].map(client_order)
            
            disp_debt = disp_debt.sort_values(by=["거래처순위", "구분순위"]).drop(columns=["거래처순위", "구분순위"])
            
            disp_debt = disp_debt.set_index(["거래처", "구분"])
            debt_highlight = selected_client != "전체 거래처"
            df_height = 500 if selected_client != "전체 거래처" else 700
            # 월 접기/펼치기: 버튼 토글 (거래처·구분 고정)
            if "debt_visible_months" not in st.session_state:
                st.session_state["debt_visible_months"] = list(numeric_cols)
            else:
                kept = [c for c in st.session_state["debt_visible_months"] if c in numeric_cols]
                st.session_state["debt_visible_months"] = kept if kept else list(numeric_cols)
            with st.container(key="debt_compact_btns_months"):
                m_cols = st.columns(len(numeric_cols), gap="small")
                for _i, _m in enumerate(numeric_cols):
                    with m_cols[_i]:
                        _on = _m in st.session_state["debt_visible_months"]
                        if st.button(
                            str(_m),
                            key=f"debt_month_tog_{_m}",
                            type="primary" if _on else "secondary",
                            width="stretch",
                        ):
                            cur = list(st.session_state["debt_visible_months"])
                            if _on:
                                if len(cur) > 1:
                                    cur = [c for c in cur if c != _m]
                            else:
                                cur.append(_m)
                                cur = [c for c in numeric_cols if c in cur]
                            st.session_state["debt_visible_months"] = cur
                            st.rerun()
            show_cols = [c for c in numeric_cols if c in st.session_state["debt_visible_months"]]
            if not show_cols:
                st.info("표시할 월을 하나 이상 선택하세요. (거래처·구분은 항상 표시)")
            else:
                # 입금기준표 결제조건 → 표 맨 우측 고정 표시
                if os.path.exists(PAYMENT_TERMS_FALLBACK) and not os.path.exists(PAYMENT_TERMS_PATH):
                    try:
                        shutil.copy2(PAYMENT_TERMS_FALLBACK, PAYMENT_TERMS_PATH)
                    except Exception:
                        pass
                payment_terms_map = load_payment_terms_map()
                render_debt_interactive_table(
                    disp_debt[show_cols],
                    debt_highlight,
                    height=df_height,
                    payment_terms_map=payment_terms_map,
                )
                # 상세표 아래: 연체개월수 요약 — 거래처(상위검색) 무시, 담당자 필터만 적용
                if not staff_debt_df.empty:
                    _staff_num = [c for c in staff_debt_df.columns if c not in ("거래처", "구분")]
                    _staff_months = [c for c in _staff_num if staff_debt_df[c].abs().sum() > 0]
                    _rank_months = [c for c in show_cols if c in _staff_months] or _staff_months
                    render_debt_month_rank_panel(
                        staff_debt_df,
                        _rank_months,
                        payment_terms_map=payment_terms_map,
                        height=480,
                        status_month_cols=_staff_months,
                    )
                else:
                    render_debt_month_rank_panel(
                        filtered_debt_df,
                        show_cols,
                        payment_terms_map=payment_terms_map,
                        height=480,
                        status_month_cols=numeric_cols,
                    )
# Tab 6: 📍 대한민국 V-World 고해상도 한글/위성 지도 적용
with tab6:
    t6_c1, t6_c2 = st.columns([4, 1])
    t6_c1.markdown("<div class='sub-header dashboard-tab-panel-head'>📍 담당자별 거래처 지도 분포 (대한민국 V-World 지도)</div>", unsafe_allow_html=True)
    t6_c2.markdown(render_update_badge(latest_update_str), unsafe_allow_html=True)
    
    rest_api_key = "21a8c4d7312051598c2e05dba0b9c0c7"
    
    map_col1, map_col2, map_col3 = st.columns([1, 1, 1])
    with map_col1:
        map_style_choice = st.radio(
            "🗺️ 지도 배경 스타일 선택",
            ["일반 지도 (V-World 한글 기본도)", "위성 지도 (V-World 고해상도 위성)"],
            horizontal=True,
            key="map_style_radio"
        )
    with map_col2:
        all_staff_list = sorted(df_base["담당자"].unique()) if not df_base.empty else []
        map_selected_staff = st.multiselect(
            "👤 지도 전용 담당자 선택", 
            options=all_staff_list, 
            default=all_staff_list,
            key="map_staff_multiselect"
        )
    with map_col3:
        all_map_clients = sorted(df_base["거래처"].unique()) if not df_base.empty else []
        map_selected_client = st.multiselect(
            "🏢 특정 거래처 위치 검색", 
            options=all_map_clients,
            placeholder="검색할 거래처명을 입력하세요...",
            key="map_client_multiselect"
        )
    
    if map_selected_client:
        addr_display_html = "<div style='background-color: #F1F5F9; padding: 8px 12px; border-radius: 6px; border: 1px solid #CBD5E1; margin-top: 5px; margin-bottom: 15px; font-size: 13px; color: #334155;'>"
        for sc in map_selected_client:
            raw_a = addr_dict.get(sc, "등록된 주소 정보가 없습니다.")
            clean_a = str(raw_a) if pd.notna(raw_a) and str(raw_a).strip().lower() != 'nan' and str(raw_a).strip() else "등록된 주소 정보가 없습니다."
            addr_display_html += f"<div>📍 <b>{sc}:</b> {clean_a}</div>"
        addr_display_html += "</div>"
        st.markdown(addr_display_html, unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)
    ctrl_space, ctrl_c1, ctrl_c2, ctrl_c3, ctrl_c4 = st.columns([5, 1.2, 1.2, 1.2, 1.2])
    
    with ctrl_space:
        st.empty()
    with ctrl_c1:
        btn_load_map = st.button("🗺️ 지도 새로고침/조회", type="primary", use_container_width=True)
    with ctrl_c2:
        btn_zoom_in = st.button("➕ 확대 (+)", use_container_width=True)
    with ctrl_c3:
        btn_zoom_out = st.button("➖ 축소 (-)", use_container_width=True)
    with ctrl_c4:
        btn_reset_map = st.button("🏠 기본 위치", use_container_width=True)
    if btn_load_map or btn_zoom_in or btn_zoom_out or btn_reset_map or "show_map" not in st.session_state:
        st.session_state.show_map = True
        
    if st.session_state.show_map:
        target_map_df = df_base.copy()
        
        if map_selected_client:
            target_map_df = target_map_df[target_map_df["거래처"].isin(map_selected_client)]
        elif map_selected_staff:
            target_map_df = target_map_df[target_map_df["담당자"].isin(map_selected_staff)]
        
        if not target_map_df.empty:
            unique_clients_df = target_map_df[['거래처', '담당자']].drop_duplicates(subset=['거래처'])
            
            map_data = []
            total_cnt = len(unique_clients_df)
            progress_text = "최초 1회 주소 좌표 변환 중입니다 (이후부터는 빠르게 로딩됩니다) 🚀"
            my_bar = st.progress(0, text=progress_text)
            
            invalid_clients = [] 
            
            for i, (_, row) in enumerate(unique_clients_df.iterrows()):
                c_name = row['거래처']
                c_staff = row['담당자']
                
                c_addr_raw = addr_dict.get(c_name, "")
                if pd.isna(c_addr_raw) or str(c_addr_raw).strip().lower() == 'nan' or not str(c_addr_raw).strip():
                    c_addr = "등록된 주소 정보가 없습니다."
                else:
                    c_addr = str(c_addr_raw)
                
                lat, lon = get_lat_lon_kakao(c_name, c_addr, rest_api_key)
                
                if lat is not None and lon is not None:
                    map_data.append({
                        "거래처": c_name,
                        "담당자": c_staff,
                        "주소": c_addr,
                        "lat": lat,
                        "lon": lon
                    })
                else:
                    invalid_clients.append(c_name)
                    
                my_bar.progress((i + 1) / total_cnt, text=f"{progress_text} ({i+1}/{total_cnt})")
            
            my_bar.empty()
            
            if map_data:
                map_df = pd.DataFrame(map_data)
                center_lat = map_df['lat'].mean()
                center_lon = map_df['lon'].mean()
                
                default_zoom = 13 if map_selected_client and len(map_selected_client) <= 3 else 8
                
                if "map_zoom" not in st.session_state or btn_reset_map or btn_load_map:
                    st.session_state.map_zoom = default_zoom
                
                if btn_zoom_in: 
                    st.session_state.map_zoom = min(st.session_state.map_zoom + 2, 20)
                elif btn_zoom_out: 
                    st.session_state.map_zoom = max(st.session_state.map_zoom - 2, 2) 
                
                fig_map = px.scatter_mapbox(
                    map_df,
                    lat="lat",
                    lon="lon",
                    color="담당자",
                    hover_name="거래처",
                    hover_data={"주소": True, "lat": False, "lon": False, "담당자": False},
                    zoom=st.session_state.map_zoom,
                    center={"lat": center_lat, "lon": center_lon},
                    height=600
                )
                fig_map.update_traces(marker=dict(size=14, opacity=0.9))
                
                vworld_base = "https://xdworld.vworld.kr/2d/Base/service/{z}/{x}/{y}.png"
                vworld_sat = "https://xdworld.vworld.kr/2d/Satellite/service/{z}/{x}/{y}.jpeg"
                vworld_hybrid = "https://xdworld.vworld.kr/2d/Hybrid/service/{z}/{x}/{y}.png"
                
                if "일반" in map_style_choice:
                    mapbox_layers = [
                        {"below": 'traces', "sourcetype": "raster", "source": [vworld_base]}
                    ]
                else:
                    mapbox_layers = [
                        {"below": 'traces', "sourcetype": "raster", "source": [vworld_sat]},
                        {"below": 'traces', "sourcetype": "raster", "source": [vworld_hybrid]}
                    ]
                
                fig_map.update_layout(
                    mapbox_style="white-bg",
                    mapbox_layers=mapbox_layers,
                    margin={"r": 0, "t": 10, "l": 0, "b": 0},
                    legend=dict(
                        orientation="h",
                        yanchor="bottom",
                        y=-0.15,
                        xanchor="center",
                        x=0.5
                    )
                )
                
                dynamic_key = f"map_chart_{hash(str(map_selected_staff))}_{hash(str(map_selected_client))}"
                render_plotly_chart(fig_map, use_container_width=True, key=dynamic_key, allow_drag=True)
                if invalid_clients:
                    with st.expander("⚠️ 지도에 표시되지 않은 거래처 (주소 정보 없음 또는 좌표 변환 실패)"):
                        st.write(", ".join(invalid_clients))
            else:
                st.info("조건에 맞는 거래처 데이터가 없습니다.")
# Tab 7: 🏭 설비 재고 현황
with tab7:
    t7_c1, t7_c2 = st.columns([4, 1])
    t7_c1.markdown("<div class='sub-header dashboard-tab-panel-head'>🏭 고압가스 탱크 및 기화기 재고 현황</div>", unsafe_allow_html=True)
    with t7_c2:
        st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)
        tab7_default = get_saved_date(TAB7_DATE_FILE)
        tab7_date = st.date_input("기준일", value=tab7_default, key="tab7_date", label_visibility="collapsed")
        if tab7_date != tab7_default:
            set_saved_date(TAB7_DATE_FILE, tab7_date)
    if not df_tank.empty or not df_vaporizer.empty:
        eq_col1, eq_col2, eq_col3, eq_col4 = st.columns(4)
        
        with eq_col1:
            available_branches = []
            if not df_tank.empty and '지사' in df_tank.columns:
                available_branches.extend(df_tank['지사'].dropna().unique())
            if not df_vaporizer.empty and '지사' in df_vaporizer.columns:
                available_branches.extend(df_vaporizer['지사'].dropna().unique())
            selected_branch = st.selectbox("📍 지사 선택 (전체 조회)", ["전체 지사"] + sorted(list(set(available_branches))))
        
        with eq_col2:
            selected_equip_type = st.selectbox("🛢️ 설비 종류 선택", ["전체 보기", "탱크 재고", "기화기 재고"])
        with eq_col3:
            selected_status = st.selectbox("📌 사용구분 필터", ["전체 상태", "유휴 장비", "거래처 사용중"])
        with eq_col4:
            eq_items = []
            if not df_tank.empty and '품목' in df_tank.columns:
                eq_items.extend(df_tank['품목'].dropna().astype(str).tolist())
            if not df_vaporizer.empty and '기화형식' in df_vaporizer.columns:
                eq_items.extend(df_vaporizer['기화형식'].dropna().astype(str).tolist())
            unique_eq_items = sorted(list(set([i.strip() for i in eq_items if i.strip() != ''])))
            selected_eq_item = st.selectbox("📦 품목/형식 선택", ["전체 품목/형식"] + unique_eq_items)
        st.markdown("---")
        if not df_tank.empty:
            if '사용구분' in df_tank.columns:
                mask_idle = df_tank['사용구분'].astype(str).str.contains('유휴')
                df_tank.loc[mask_idle, '사용구분'] = '🟢 ' + df_tank.loc[mask_idle, '사용구분'].astype(str).str.replace('🟢 ', '').str.replace('🏢 ', '')
                df_tank.loc[~mask_idle, '사용구분'] = '🏢 ' + df_tank.loc[~mask_idle, '사용구분'].astype(str).str.replace('🟢 ', '').str.replace('🏢 ', '')
        if not df_vaporizer.empty:
            if '사용구분' in df_vaporizer.columns:
                mask_idle_v = df_vaporizer['사용구분'].astype(str).str.contains('유휴')
                df_vaporizer.loc[mask_idle_v, '사용구분'] = '🟢 ' + df_vaporizer.loc[mask_idle_v, '사용구분'].astype(str).str.replace('🟢 ', '').str.replace('🏢 ', '')
                df_vaporizer.loc[~mask_idle_v, '사용구분'] = '🏢 ' + df_vaporizer.loc[~mask_idle_v, '사용구분'].astype(str).str.replace('🟢 ', '').str.replace('🏢 ', '')
        if selected_equip_type in ["전체 보기", "탱크 재고"]:
            st.markdown("<div style='font-size: 16px; font-weight: 700; color: #1E3A8A; margin-bottom: 10px;'>🛢️ 초저온 탱크 재고 현황</div>", unsafe_allow_html=True)
            if not df_tank.empty:
                filtered_tank = df_tank.copy()
                if selected_branch != "전체 지사" and '지사' in filtered_tank.columns:
                    filtered_tank = filtered_tank[filtered_tank['지사'].astype(str).str.contains(selected_branch)]
                if selected_status != "전체 상태" and '사용구분' in filtered_tank.columns:
                    filtered_tank = filtered_tank[filtered_tank['사용구분'].astype(str).str.contains(selected_status)]
                if selected_eq_item != "전체 품목/형식" and '품목' in filtered_tank.columns:
                    filtered_tank = filtered_tank[filtered_tank['품목'].astype(str).str.strip() == selected_eq_item]
                
                st.dataframe(filtered_tank, use_container_width=True, height=350, hide_index=True)
        if selected_equip_type in ["전체 보기", "기화기 재고"]:
            if selected_equip_type == "전체 보기":
                st.markdown("<br>", unsafe_allow_html=True)
            
            st.markdown(f"<div style='font-size: 16px; font-weight: 700; color: #1E3A8A; margin-bottom: 10px;'>♨️ 기화기 재고 현황</div>", unsafe_allow_html=True)
            if not df_vaporizer.empty:
                filtered_vap = df_vaporizer.copy()
                if selected_branch != "전체 지사" and '지사' in filtered_vap.columns:
                    filtered_vap = filtered_vap[filtered_vap['지사'].astype(str).str.contains(selected_branch)]
                if selected_status != "전체 상태" and '사용구분' in filtered_vap.columns:
                    filtered_vap = filtered_vap[filtered_vap['사용구분'].astype(str).str.contains(selected_status)]
                if selected_eq_item != "전체 품목/형식" and '기화형식' in filtered_vap.columns:
                    filtered_vap = filtered_vap[filtered_vap['기화형식'].astype(str).str.strip() == selected_eq_item]
                
                st.dataframe(filtered_vap, use_container_width=True, height=350, hide_index=True)
# Tab 8: 🛢️ 통합 탱크 재고
with tab8:
    t8_c1, t8_c2 = st.columns([4, 1])
    t8_c1.markdown("<div class='sub-header dashboard-tab-panel-head'>🛢️ 통합 고압가스 탱크 재고 현황</div>", unsafe_allow_html=True)
    with t8_c2:
        st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)
        tab8_default = get_saved_date(TAB8_DATE_FILE)
        tab8_date = st.date_input("기준일", value=tab8_default, key="int_date", label_visibility="collapsed")
        if tab8_date != tab8_default:
            set_saved_date(TAB8_DATE_FILE, tab8_date)
        
    if not df_integrated.empty:
        int_col1, int_col2 = st.columns(2)
        with int_col1:
            items = ["전체 품목"] + sorted([str(x) for x in df_integrated['품목'].dropna().unique() if str(x).strip()])
            sel_item = st.selectbox("📦 품목 선택", items, key="int_item")
        with int_col2:
            statuses = ["전체 상태"] + sorted([str(x) for x in df_integrated['사용구분'].dropna().unique() if str(x).strip()])
            sel_status = st.selectbox("📌 사용구분", statuses, key="int_status")
        st.markdown("---")
        df_int_filtered = df_integrated.copy()
        if sel_item != "전체 품목":
            df_int_filtered = df_int_filtered[df_int_filtered['품목'].astype(str) == sel_item]
        if sel_status != "전체 상태":
            df_int_filtered = df_int_filtered[df_int_filtered['사용구분'].astype(str) == sel_status]
        total_tanks = len(df_int_filtered)
        idle_tanks = len(df_int_filtered[df_int_filtered['사용구분'].astype(str).str.contains('유휴', na=False)])
        inuse_tanks = total_tanks - idle_tanks
        
        k1, k2, k3 = st.columns(3)
        k1.markdown(f"<div class='metric-box'><div class='metric-label'>총 탱크 수량</div><div class='metric-value'>{total_tanks:,} 기</div></div>", unsafe_allow_html=True)
        k2.markdown(f"<div class='metric-box'><div class='metric-label'>🟢 유휴 장비 (대기중)</div><div class='metric-value' style='color:#059669;'>{idle_tanks:,} 기</div></div>", unsafe_allow_html=True)
        k3.markdown(f"<div class='metric-box'><div class='metric-label'>🏢 사용/충전중</div><div class='metric-value' style='color:#2563EB;'>{inuse_tanks:,} 기</div></div>", unsafe_allow_html=True)
        st.markdown("<div style='font-size: 14px; font-weight: 600; color: #334155; margin-bottom: 10px; margin-top: 20px;'>📋 상세 재고 데이터</div>", unsafe_allow_html=True)
        
        df_display = df_int_filtered.copy()
        if '사용구분' in df_display.columns:
            mask_idle = df_display['사용구분'].astype(str).str.contains('유휴')
            df_display.loc[mask_idle, '사용구분'] = '🟢 ' + df_display.loc[mask_idle, '사용구분'].astype(str).str.replace('🟢 ', '').str.replace('🏢 ', '')
            df_display.loc[~mask_idle, '사용구분'] = '🏢 ' + df_display.loc[~mask_idle, '사용구분'].astype(str).str.replace('🟢 ', '').str.replace('🏢 ', '')
        st.dataframe(df_display, use_container_width=True, height=600, hide_index=True)
    else:
        st.warning("통합 탱크 재고 데이터가 없습니다. 폴더에 '통합탱크재고.csv'를 넣거나 왼쪽 사이드바에서 업로드해주세요.")
# Tab 9: 📈 수익성 분석 (엑셀 함수 동일 적용)
# 입력 변경 시 Tab9만 부분 재실행 (다른 탭·상단 로딩 생략)
@st.fragment
def _render_profitability_analysis_tab(latest_update_str):
    t9_c1, t9_c2 = st.columns([4, 1])
    t9_c1.markdown(
        "<div class='sub-header dashboard-tab-panel-head'>📈 투자대비 수익성 분석</div>",
        unsafe_allow_html=True,
    )
    t9_c2.markdown(render_update_badge(latest_update_str), unsafe_allow_html=True)
    st.caption("엑셀「수익성분석.xlsx」함수를 그대로 적용합니다. 입력값을 바꾸면 결과가 즉시 재계산됩니다.")
    if "profit_inputs" not in st.session_state:
        st.session_state["profit_inputs"] = load_profit_inputs()
    p0 = st.session_state["profit_inputs"]
    _pf_keys = [
        "pf_name", "pf_tank_gas", "pf_tank_cap_mode",
        "pf_tank_liters", "pf_tank_liters__comma",
        "pf_tank_kg", "pf_tank_kg__comma", "pf_tank_spec",
        "pf_hourly_mode", "pf_hourly_usage", "pf_hourly_nm3", "pf_operating_hours",
        "pf_operating_days", "pf_auto_monthly",
        "pf_tank_price", "pf_tank_price__comma", "pf_usage", "pf_usage__comma",
        "pf_const", "pf_const__comma", "pf_vap_cap", "pf_vap_cap__comma",
        "pf_vap_note", "pf_vap_price", "pf_vap_price__comma",
        "pf_buy", "pf_logi", "pf_supply",
        "pf_rate", "pf_mgmt", "pf_dep", "pf_rent", "pf_rent_n",
        "pf_origin", "pf_dest", "pf_origin_cands", "pf_dest_cands",
        "pf_origin_q", "pf_dest_q", "pf_origin_pick", "pf_dest_pick",
        "pf_lkm", "pf_lfuel", "pf_leff", "pf_ltoll", "pf_lrt", "pf_lkg",
    ]
    # —— 입력 (물류비 계산 → 단가 순, 글씨·위젯 스타일은 단가란과 동일) ——
    st.markdown("##### ◆ 프로젝트 / 장비 투자비")
    c_name, c_gas = st.columns([1, 1])
    project_name = c_name.text_input("거래처/프로젝트명", value=str(p0.get("project_name", "")), key="pf_name")
    _gas0 = str(p0.get("tank_gas") or PROFIT_DEFAULTS["tank_gas"])
    if _gas0 not in GAS_OPTIONS:
        _gas0 = GAS_OPTIONS[0]
    tank_gas = c_gas.selectbox(
        "탱크 가스 종류",
        GAS_OPTIONS,
        index=GAS_OPTIONS.index(_gas0),
        key="pf_tank_gas",
        help="질소·알곤·산소·탄산·수소·헬륨. 내용적(L) 입력 시 kg 환산에 사용됩니다.",
    )
    _dens = float(GAS_DENSITY_KG_PER_L.get(tank_gas, 0.808))
    _mode0 = str(p0.get("tank_capacity_mode") or "liters")
    if _mode0 not in ("liters", "kg"):
        _mode0 = "liters"
    tank_capacity_mode = st.radio(
        "탱크 용량 입력 방식",
        options=["liters", "kg"],
        index=0 if _mode0 == "liters" else 1,
        format_func=lambda m: "내용적(L) → kg 환산" if m == "liters" else "용량(kg) 직접 입력",
        horizontal=True,
        key="pf_tank_cap_mode",
        help="L로 넣거나, kg를 바로 넣을 수 있습니다.",
    )
    _kg0 = parse_tank_capacity_kg(p0.get("tank_spec", PROFIT_DEFAULTS["tank_spec"]))
    _liters0 = p0.get("tank_liters")
    if _liters0 is None or float(_liters0 or 0) <= 0:
        _liters0 = (_kg0 / _dens) if _dens > 0 and _kg0 > 0 else float(PROFIT_DEFAULTS["tank_liters"])
    t_l, t_k = st.columns([1, 1])
    if tank_capacity_mode == "liters":
        with t_l:
            tank_liters = profit_int_comma_input(
                "TANK 내용적 (L)",
                key="pf_tank_liters",
                value=_liters0,
                help=f"{tank_gas} 밀도 {_dens:g} kg/L × 내용적(L) = 용량(kg)",
            )
        tank_kg = round(liters_to_tank_kg(tank_liters, tank_gas))
        st.session_state["pf_tank_kg"] = int(tank_kg)
        st.session_state["pf_tank_kg__comma"] = f"{int(tank_kg):,}"
        t_k.markdown(
            f"<div style='padding-top:0.2rem;'>"
            f"<div style='font-size:0.875rem;color:#31333F;margin-bottom:0.25rem;'>TANK 용량 (kg) 환산</div>"
            f"<div style='font-size:1rem;font-weight:600;color:#0F172A;'>{tank_kg:,.0f}</div>"
            f"<div style='font-size:0.75rem;color:#64748B;'>"
            f"{tank_gas} {_dens:g} kg/L × {float(tank_liters):,.0f} L</div>"
            f"</div>",
            unsafe_allow_html=True,
        )
    else:
        with t_k:
            tank_kg = profit_int_comma_input(
                "TANK 용량 (kg) 직접입력",
                key="pf_tank_kg",
                value=_kg0 if _kg0 > 0 else 4900,
                help="용량을 kg로 바로 입력합니다. 왕복횟수 계산에 사용됩니다.",
            )
        tank_liters = round(float(tank_kg) / _dens) if _dens > 0 else 0.0
        st.session_state["pf_tank_liters"] = int(tank_liters)
        st.session_state["pf_tank_liters__comma"] = f"{int(tank_liters):,}"
        t_l.markdown(
            f"<div style='padding-top:0.2rem;'>"
            f"<div style='font-size:0.875rem;color:#31333F;margin-bottom:0.25rem;'>TANK 내용적 (L) 환산</div>"
            f"<div style='font-size:1rem;font-weight:600;color:#0F172A;'>{tank_liters:,.0f}</div>"
            f"<div style='font-size:0.75rem;color:#64748B;'>"
            f"{float(tank_kg):,.0f} kg ÷ {tank_gas} {_dens:g} kg/L</div>"
            f"</div>",
            unsafe_allow_html=True,
        )
    tank_spec = float(tank_kg)  # 저장·계산용 (단위: kg)
    _nm3 = tank_kg_to_nm3(tank_kg, tank_gas)
    _nm3_per_kg = float(GAS_NM3_PER_KG.get(tank_gas, 0))
    st.caption(
        f"기체환산({tank_gas}): {float(tank_kg):,.0f} kg × {_nm3_per_kg:g} Nm³/kg = "
        f"**{_nm3:,.1f} Nm³**  ·  내용적 {float(tank_liters):,.0f} L 기준 "
        f"{tank_liters_to_nm3(tank_liters, tank_gas):,.1f} Nm³"
    )
    with st.expander("각 가스별 밀도·기체환산 비교", expanded=False):
        st.caption("동일 내용적(L)으로 가스별 kg·Nm³를 비교합니다. ★ = 현재 선택 가스. (0℃·1atm 대략값)")
        _gas_df = pd.DataFrame(gas_conversion_rows(tank_liters, tank_kg, tank_gas))
        _gas_df["밀도(kg/L)"] = _gas_df["밀도(kg/L)"].map(lambda x: f"{float(x):.4g}")
        _gas_df["Nm³/kg"] = _gas_df["Nm³/kg"].map(lambda x: f"{float(x):.4g}")
        _gas_df["내용적 기준 kg"] = _gas_df["내용적 기준 kg"].map(lambda x: f"{float(x):,.1f}")
        _gas_df["내용적 기준 Nm³"] = _gas_df["내용적 기준 Nm³"].map(lambda x: f"{float(x):,.1f}")
        _gas_df["현재탱크 Nm³"] = _gas_df["현재탱크 Nm³"].map(
            lambda x: f"{float(x):,.1f}" if isinstance(x, (int, float)) else str(x)
        )
        st.dataframe(_gas_df, hide_index=True, width="stretch")
    # 별도 작은 타일: 탱크 사용주기 (화면 복잡도 ↓ — 접힌 expander)
    _nm3pkg = float(GAS_NM3_PER_KG.get(tank_gas, 0) or 0)
    with st.expander("⏱ 탱크 사용주기 (시간당 사용량 · 가동시간)", expanded=False):
        st.caption(
            f"충전기준 = 탱크×80% · 사용주기 = 충전기준 ÷ (시간당kg × 일가동) · "
            f"{tank_gas} 환산 {_nm3pkg:g} Nm³/kg"
        )
        # 기존 float 세션값 → 정수 (format=%d 호환)
        for _pk in ("pf_hourly_nm3", "pf_hourly_usage", "pf_operating_hours", "pf_operating_days"):
            if _pk in st.session_state:
                try:
                    st.session_state[_pk] = int(round(float(st.session_state[_pk])))
                except Exception:
                    pass
        _hmode0 = str(p0.get("hourly_usage_mode") or "kg")
        if _hmode0 not in ("nm3", "kg"):
            _hmode0 = "kg"
        hourly_usage_mode = st.radio(
            "시간당 사용량 입력",
            options=["nm3", "kg"],
            index=0 if _hmode0 == "nm3" else 1,
            format_func=lambda m: "루베(Nm³/h) → kg/h 환산" if m == "nm3" else "kg/h 직접 입력",
            horizontal=True,
            key="pf_hourly_mode",
        )
        _kg_h0 = float(p0.get("hourly_usage_kg", PROFIT_DEFAULTS["hourly_usage_kg"]) or 0)
        _nm3_h0 = p0.get("hourly_usage_nm3")
        if _nm3_h0 is None or float(_nm3_h0 or 0) <= 0:
            _nm3_h0 = kg_per_h_to_nm3_per_h(_kg_h0, tank_gas)
        u_a, u_b, u_o = st.columns(3)
        if hourly_usage_mode == "nm3":
            with u_a:
                hourly_usage_nm3 = st.number_input(
                    "시간당 사용량 (Nm³/h, 루베)",
                    min_value=0,
                    step=1,
                    value=int(round(float(_nm3_h0))),
                    format="%d",
                    key="pf_hourly_nm3",
                    help=f"{tank_gas}: kg/h = Nm³/h ÷ {_nm3pkg:g}",
                )
            hourly_usage_kg = nm3_per_h_to_kg_per_h(hourly_usage_nm3, tank_gas)
            st.session_state["pf_hourly_usage"] = int(round(hourly_usage_kg))
            u_b.markdown(
                f"<div style='padding-top:0.2rem;'>"
                f"<div style='font-size:0.875rem;color:#31333F;margin-bottom:0.25rem;'>시간당 사용량 (kg/h) 환산</div>"
                f"<div style='font-size:1rem;font-weight:600;color:#0F172A;'>{hourly_usage_kg:,.0f}</div>"
                f"<div style='font-size:0.75rem;color:#64748B;'>{tank_gas} {float(hourly_usage_nm3):,.0f} Nm³/h ÷ {_nm3pkg:g}</div>"
                f"</div>",
                unsafe_allow_html=True,
            )
        else:
            with u_b:
                hourly_usage_kg = st.number_input(
                    "시간당 사용량 (kg/h)",
                    min_value=0,
                    step=1,
                    value=int(round(float(_kg_h0))),
                    format="%d",
                    key="pf_hourly_usage",
                )
            hourly_usage_nm3 = kg_per_h_to_nm3_per_h(hourly_usage_kg, tank_gas)
            st.session_state["pf_hourly_nm3"] = int(round(hourly_usage_nm3))
            u_a.markdown(
                f"<div style='padding-top:0.2rem;'>"
                f"<div style='font-size:0.875rem;color:#31333F;margin-bottom:0.25rem;'>시간당 사용량 (Nm³/h) 환산</div>"
                f"<div style='font-size:1rem;font-weight:600;color:#0F172A;'>{hourly_usage_nm3:,.0f}</div>"
                f"<div style='font-size:0.75rem;color:#64748B;'>{tank_gas} {float(hourly_usage_kg):,.0f} kg/h × {_nm3pkg:g}</div>"
                f"</div>",
                unsafe_allow_html=True,
            )
        operating_hours = u_o.number_input(
            "일 가동시간 (h/일)",
            min_value=0,
            max_value=24,
            step=1,
            value=int(round(float(p0.get("operating_hours", PROFIT_DEFAULTS["operating_hours"])))),
            format="%d",
            key="pf_operating_hours",
        )
        d_days, d_auto = st.columns([1, 2])
        operating_days = d_days.number_input(
            "월 가동일수 (일/월)",
            min_value=1,
            max_value=31,
            step=1,
            value=int(round(float(p0.get("operating_days_per_month", PROFIT_DEFAULTS["operating_days_per_month"])))),
            format="%d",
            key="pf_operating_days",
            help="월 사용량 = 일 사용량(시간당×일가동) × 월 가동일수. 예: 주5일≈22일, 연중무휴≈30일",
        )
        _cycle = compute_tank_usage_cycle(
            tank_kg,
            hourly_usage_kg,
            operating_hours,
            fill_ratio=0.8,
            days_per_month=operating_days,
        )
        _monthly_est = float(_cycle.get("monthly_kg") or 0)
        auto_monthly_from_cycle = d_auto.checkbox(
            f"월 평균 공급량에 자동 반영 (일사용량 × {int(operating_days)}일)",
            value=bool(p0.get("auto_monthly_from_cycle", False)),
            key="pf_auto_monthly",
            help="체크 시 아래 「월 평균 공급량」= 일사용량 × 월가동일수. 해제하면 직접 입력합니다.",
        )
        if _cycle.get("ok"):
            c1, c2, c3, c4 = st.columns(4)
            c1.markdown(
                f"<div class='metric-box'><div class='metric-label'>일 사용량</div>"
                f"<div class='metric-value' style='font-size:18px;'>{_cycle['daily_kg']:,.0f} kg/일</div></div>",
                unsafe_allow_html=True,
            )
            c2.markdown(
                f"<div class='metric-box'><div class='metric-label'>사용주기</div>"
                f"<div class='metric-value' style='font-size:18px;color:#2563EB;'>{_cycle['cycle_days']:,.0f} 일</div></div>",
                unsafe_allow_html=True,
            )
            c3.markdown(
                f"<div class='metric-box'><div class='metric-label'>가동시간 기준</div>"
                f"<div class='metric-value' style='font-size:18px;'>{_cycle['cycle_hours']:,.0f} h</div></div>",
                unsafe_allow_html=True,
            )
            c4.markdown(
                f"<div class='metric-box'><div class='metric-label'>월 사용량({int(operating_days)}일)</div>"
                f"<div class='metric-value' style='font-size:18px;'>{_monthly_est:,.0f} kg</div></div>",
                unsafe_allow_html=True,
            )
            st.caption(
                f"탱크 {float(tank_kg):,.0f} kg · 충전기준 {_cycle['charge_kg']:,.0f} kg(80%) · "
                f"월가동 {int(operating_days)}일 · 월충전 약 {_cycle['fills_per_month']:,.0f}회 · {_cycle['message']}"
            )
            if auto_monthly_from_cycle:
                _mu = int(round(_monthly_est))
                st.session_state["pf_usage"] = _mu
                st.session_state["pf_usage__comma"] = f"{_mu:,}"
                st.caption(
                    f"→ 월 평균 공급량 {_mu:,} kg "
                    f"(일 {_cycle['daily_kg']:,.0f} × {int(operating_days)}일) 자동 반영 중"
                )
        else:
            st.caption(_cycle.get("message") or "입력값을 확인하세요.")
    i1, i2, i3 = st.columns(3)
    with i1:
        tank_price = profit_int_comma_input(
            "1. TANK 구입가 (원)", key="pf_tank_price", value=p0["tank_price"]
        )
    with i2:
        if auto_monthly_from_cycle and _cycle.get("ok"):
            monthly_usage = float(st.session_state.get("pf_usage", round(_monthly_est)))
            st.markdown(
                f"<div style='padding-top:0.2rem;'>"
                f"<div style='font-size:0.875rem;color:#31333F;margin-bottom:0.25rem;'>월 평균 공급량 (kg)</div>"
                f"<div style='font-size:1rem;font-weight:600;color:#0F172A;'>{monthly_usage:,.0f}</div>"
                f"<div style='font-size:0.75rem;color:#64748B;'>사용주기 자동반영 (직접입력은 위 체크 해제)</div>"
                f"</div>",
                unsafe_allow_html=True,
            )
        else:
            monthly_usage = profit_int_comma_input(
                "월 평균 공급량 (kg)",
                key="pf_usage",
                value=p0["monthly_usage_kg"],
                help="직접 입력. 사용주기에서 자동반영하려면 위 체크박스를 켜세요.",
            )
    with i3:
        construction = profit_int_comma_input(
            "3. 공사비용 (원)", key="pf_const", value=p0["construction_cost"]
        )
    v1, v2, v3 = st.columns(3)
    with v1:
        vap_cap = profit_int_comma_input(
            "2. 기화기 용량 (Nm3/hr)", key="pf_vap_cap", value=p0["vaporizer_capacity"]
        )
    vap_note = v2.text_input("기화기 수량 메모", value=str(p0.get("vaporizer_qty_note", "")), key="pf_vap_note")
    with v3:
        vap_price = profit_int_comma_input(
            "기화기 구입가 (원)", key="pf_vap_price", value=p0["vaporizer_price"]
        )
    # —— 물류비 계산 (단가 바로 위) ——
    st.markdown("##### ◎ 물류비 계산")
    st.caption(
        "20톤 벌크로리 · 경유 · 통행료 5종  ·  "
        "왕복=탱크용량×80%충전 기준 자동  ·  "
        "((거리km × 유류비/L ÷ 연비) + 통행료) × 왕복 ÷ 월평균공급량 → 「2. 물류비」반영"
    )
    # 경유 시장가(오피넷) — 매월 자동 반영
    force_diesel = st.session_state.pop("pf_diesel_force", False)
    diesel_info = get_diesel_price_monthly(force_refresh=force_diesel)
    month_now = datetime.date.today().strftime("%Y-%m")
    if diesel_info.get("ok"):
        if st.session_state.get("pf_diesel_applied_month") != month_now or force_diesel:
            st.session_state["pf_lfuel"] = float(diesel_info["price"])
            st.session_state["pf_diesel_applied_month"] = month_now
    a1, a2 = st.columns(2)
    logi_origin = a1.text_input(
        "출발지 (주소·상호·명칭)",
        value=str(p0.get("logi_origin", PROFIT_DEFAULTS["logi_origin"])),
        key="pf_origin",
        placeholder="예: 신일가스 화성공장 / 경기도 화성시 …",
        help="상호·지점명이 여러 개면 아래에서 선택할 수 있습니다.",
    )
    logi_dest = a2.text_input(
        "도착지 (주소·상호·명칭)",
        value=str(p0.get("logi_dest", "")),
        key="pf_dest",
        placeholder="예: 제이아이금속 / 경북 구미시 …",
        help="상호·지점명이 여러 개면 아래에서 선택할 수 있습니다.",
    )
    _oq = str(logi_origin or "").strip()
    _dq = str(logi_dest or "").strip()
    _o_cands = st.session_state.get("pf_origin_cands") if st.session_state.get("pf_origin_q") == _oq else []
    _d_cands = st.session_state.get("pf_dest_cands") if st.session_state.get("pf_dest_q") == _dq else []
    _o_cands = _o_cands or []
    _d_cands = _d_cands or []
    # 지점이 여러 개면 선택 UI
    if _o_cands or _d_cands:
        p1, p2 = st.columns(2)
        with p1:
            if len(_o_cands) > 1:
                st.selectbox(
                    f"출발 지점 선택 ({len(_o_cands)}곳)",
                    options=list(range(len(_o_cands))),
                    format_func=lambda i: _o_cands[i]["label"],
                    key="pf_origin_pick",
                )
            elif len(_o_cands) == 1:
                st.caption(f"출발: {_o_cands[0]['label']}")
        with p2:
            if len(_d_cands) > 1:
                st.selectbox(
                    f"도착 지점 선택 ({len(_d_cands)}곳)",
                    options=list(range(len(_d_cands))),
                    format_func=lambda i: _d_cands[i]["label"],
                    key="pf_dest_pick",
                )
            elif len(_d_cands) == 1:
                st.caption(f"도착: {_d_cands[0]['label']}")
    btn_r, btn_f = st.columns(2)
    if btn_r.button("📍 거리·통행료 조회", key="pf_route_btn", type="secondary", width="stretch"):
        with st.spinner("통합검색·거리·통행료 조회 중…"):
            need_search = (
                st.session_state.get("pf_origin_q") != _oq
                or st.session_state.get("pf_dest_q") != _dq
                or not st.session_state.get("pf_origin_cands")
                or not st.session_state.get("pf_dest_cands")
            )
            if need_search:
                o_cands = kakao_place_search(_oq)
                d_cands = kakao_place_search(_dq)
                st.session_state["pf_origin_cands"] = o_cands
                st.session_state["pf_dest_cands"] = d_cands
                st.session_state["pf_origin_q"] = _oq
                st.session_state["pf_dest_q"] = _dq
                st.session_state.pop("pf_origin_pick", None)
                st.session_state.pop("pf_dest_pick", None)
                if not o_cands or not d_cands:
                    st.session_state["pf_route_info"] = {
                        "ok": False,
                        "message": "출발/도착을 찾지 못했습니다. 주소·상호·명칭을 확인해 주세요.",
                    }
                    st.error(st.session_state["pf_route_info"]["message"])
                elif len(o_cands) > 1 or len(d_cands) > 1:
                    st.warning("지점이 여러 개입니다. 아래에서 지점을 고른 뒤 다시 「거리·통행료 조회」를 눌러주세요.")
                    st.rerun()
                else:
                    oi, di = 0, 0
                    route = kakao_route_from_coords(
                        o_cands[oi]["lat"], o_cands[oi]["lon"],
                        d_cands[di]["lat"], d_cands[di]["lon"],
                        o_cands[oi]["label"], d_cands[di]["label"],
                    )
                    st.session_state["pf_route_info"] = route
                    if route.get("ok"):
                        st.session_state["pf_lkm"] = round(float(route["km"]), 1)
                        st.session_state["pf_ltoll"] = float(route.get("toll") or 0)
                        st.success(route.get("message") or f"거리 {route['km']:.1f} km")
                    else:
                        st.error(route.get("message") or "거리 조회 실패")
            else:
                o_cands = st.session_state.get("pf_origin_cands") or []
                d_cands = st.session_state.get("pf_dest_cands") or []
                if not o_cands or not d_cands:
                    st.error("출발/도착을 찾지 못했습니다. 주소·상호·명칭을 확인해 주세요.")
                else:
                    oi = int(st.session_state.get("pf_origin_pick", 0) or 0) if len(o_cands) > 1 else 0
                    di = int(st.session_state.get("pf_dest_pick", 0) or 0) if len(d_cands) > 1 else 0
                    oi = max(0, min(oi, len(o_cands) - 1))
                    di = max(0, min(di, len(d_cands) - 1))
                    route = kakao_route_from_coords(
                        o_cands[oi]["lat"], o_cands[oi]["lon"],
                        d_cands[di]["lat"], d_cands[di]["lon"],
                        o_cands[oi]["label"], d_cands[di]["label"],
                    )
                    st.session_state["pf_route_info"] = route
                    if route.get("ok"):
                        st.session_state["pf_lkm"] = round(float(route["km"]), 1)
                        st.session_state["pf_ltoll"] = float(route.get("toll") or 0)
                        st.success(route.get("message") or f"거리 {route['km']:.1f} km")
                    else:
                        st.error(route.get("message") or "거리 조회 실패")
    if btn_f.button("⛽ 경유 시세 새로고침", key="pf_diesel_btn", width="stretch"):
        st.session_state["pf_diesel_force"] = True
        st.rerun()
    route_info = st.session_state.get("pf_route_info") or {}
    if route_info.get("ok"):
        st.caption(
            f"출발: {route_info.get('origin_label','')} → 도착: {route_info.get('dest_label','')}  ·  "
            f"{route_info.get('message','')}"
        )
    if diesel_info.get("ok"):
        st.caption(
            f"경유 시장가 {float(diesel_info['price']):,.2f} 원/L  ·  "
            f"{diesel_info.get('source','')}  ·  기준일 {diesel_info.get('asof','')}  ·  "
            f"{month_now} 자동반영"
        )
    else:
        st.caption(diesel_info.get("message") or "경유 시세 조회 실패 — 유류비를 수동 입력하세요.")
    if "pf_lkm" not in st.session_state:
        st.session_state["pf_lkm"] = float(p0.get("logi_km", 70.0))
    if "pf_lfuel" not in st.session_state:
        st.session_state["pf_lfuel"] = float(
            diesel_info["price"] if diesel_info.get("ok") else p0.get("logi_fuel_price", 1400.0)
        )
    l1, l2, l3 = st.columns(3)
    logi_km = l1.number_input("거리 (km)", min_value=0.0, step=0.1, key="pf_lkm")
    logi_fuel = l2.number_input(
        "유류비 (원/L, 경유)", min_value=0.0, step=1.0, key="pf_lfuel",
    )
    logi_eff = l3.number_input(
        "연비 (km/L, 20톤벌크)", min_value=0.1, step=0.1,
        value=float(p0.get("logi_efficiency", 2.5)), key="pf_leff",
    )
    l4, l5, l6 = st.columns(3)
    if "pf_ltoll" not in st.session_state:
        st.session_state["pf_ltoll"] = float(p0.get("logi_toll", 0.0))
    logi_toll = l4.number_input(
        "통행료 (원, 편도·20톤벌크)", min_value=0.0, step=100.0, key="pf_ltoll",
    )
    # 왕복횟수 = ceil(월평균공급량 ÷ (탱크총용량kg × 80%))
    _rt_info = compute_roundtrips_from_tank(monthly_usage, tank_spec, fill_ratio=0.8)
    logi_rt = float(_rt_info["roundtrips"]) if _rt_info.get("ok") else float(p0.get("logi_roundtrips", 0) or 0)
    st.session_state["pf_lrt"] = logi_rt
    l5.markdown(
        f"<div style='padding-top:0.2rem;'>"
        f"<div style='font-size:0.875rem;color:#31333F;margin-bottom:0.25rem;'>왕복 횟수 (회/월)</div>"
        f"<div style='font-size:1rem;font-weight:600;color:#0F172A;'>{logi_rt:,.0f}</div>"
        f"<div style='font-size:0.75rem;color:#64748B;'>탱크 80% 소모시 충전</div>"
        f"</div>",
        unsafe_allow_html=True,
    )
    # 물류비 공급량 = 프로젝트 「월 평균 공급량」 그대로 (별도 입력 없음)
    logi_kg = float(monthly_usage) if float(monthly_usage or 0) > 0 else 1.0
    st.session_state["pf_lkg"] = logi_kg
    l6.markdown(
        f"<div style='padding-top:0.2rem;'>"
        f"<div style='font-size:0.875rem;color:#31333F;margin-bottom:0.25rem;'>월평균 공급량 (kg)</div>"
        f"<div style='font-size:1rem;font-weight:600;color:#0F172A;'>{logi_kg:,.0f}</div>"
        f"<div style='font-size:0.75rem;color:#64748B;'>프로젝트 월 평균 공급량 연동</div>"
        f"</div>",
        unsafe_allow_html=True,
    )
    logi_calc = compute_logistics_unit_cost(
        logi_km, logi_fuel, logi_eff, logi_toll, logi_rt, logi_kg
    )
    logi_per = int(round(float(logi_calc["per_kg"])))
    # 계산값을 단가 「물류비」위젯에 즉시 반영 (단가 렌더 전에 세팅)
    st.session_state["pf_logi"] = int(logi_per)
    st.caption(_rt_info.get("message") or "")
    st.caption(
        f"계산 물류비 {logi_per:,} 원/kg  ·  "
        f"편도연료 {logi_calc['fuel_one_way']:,.0f}원  ·  "
        f"왕복총비용 {logi_calc['round_trip_cost']:,.0f}원"
    )
    st.markdown("##### ◎ 단가")
    for _uk in ("pf_buy", "pf_logi", "pf_supply", "pf_dep", "pf_rent", "pf_rent_n"):
        if _uk in st.session_state:
            try:
                st.session_state[_uk] = int(round(float(st.session_state[_uk])))
            except Exception:
                pass
    u1, u2, u3 = st.columns(3)
    with u1:
        purchase_unit = st.number_input(
            "1. 매입단가 (원/kg)",
            min_value=0,
            step=1,
            value=int(round(float(p0["purchase_unit"]))),
            format="%d",
            key="pf_buy",
        )
    with u2:
        logistics_unit = st.number_input(
            "2. 물류비 (원/kg)",
            min_value=0,
            step=1,
            format="%d",
            key="pf_logi",
        )
        st.caption(
            f"적용: ((거리 {float(logi_km):,.0f}km × 유류 {float(logi_fuel):,.0f}원/L ÷ 연비 {float(logi_eff):g}) "
            f"+ 통행료 {float(logi_toll):,.0f}원) × 왕복 {float(logi_rt):,.0f}회 "
            f"÷ 월공급 {float(logi_kg):,.0f}kg = **{logi_per:,} 원/kg** → 위 칸에 자동반영"
        )
    with u3:
        supply_unit = st.number_input(
            "3. 공급단가 (원/kg)",
            min_value=0,
            step=1,
            value=int(round(float(p0["supply_unit"]))),
            format="%d",
            key="pf_supply",
        )
    st.markdown("##### ◎ 금융 / 관리 / 임대")
    f1, f2, f3, f4 = st.columns(4)
    interest = f1.number_input(
        "이자율 (예: 0.05=5%)", min_value=0.0, max_value=1.0, step=0.005, format="%.3f",
        value=float(p0["interest_rate"]), key="pf_rate",
    )
    mgmt_rate = f2.number_input(
        "일반관리비 비율 (월매출)", min_value=0.0, max_value=1.0, step=0.005, format="%.3f",
        value=float(p0["mgmt_rate"]), key="pf_mgmt",
    )
    dep_months = f3.number_input(
        "감가상각 개월 (10년=120)",
        min_value=1,
        step=12,
        value=int(round(float(p0["depreciation_months"]))),
        format="%d",
        key="pf_dep",
    )
    rent = f4.number_input(
        "장비 임대료 (원)",
        min_value=0,
        step=10000,
        value=int(round(float(p0["equipment_rent"]))),
        format="%d",
        key="pf_rent",
    )
    rent_count = st.number_input(
        "부가횟수",
        min_value=0,
        step=1,
        value=int(round(float(p0["rent_count"]))),
        format="%d",
        key="pf_rent_n",
    )
    b_save, b_reset = st.columns(2)
    submitted = b_save.button("💾 계산 / 저장", type="primary", width="stretch", key="pf_save_btn")
    reset = b_reset.button("↺ 엑셀 기본값으로 초기화", width="stretch", key="pf_reset_btn")
    if reset:
        st.session_state["profit_inputs"] = dict(PROFIT_DEFAULTS)
        save_profit_inputs(st.session_state["profit_inputs"])
        for k in _pf_keys:
            st.session_state.pop(k, None)
        st.session_state.pop("pf_route_info", None)
        st.session_state.pop("pf_diesel_applied_month", None)
        st.session_state.pop("pf_diesel_force", None)
        st.rerun()
    if submitted:
        new_p = {
            "project_name": project_name,
            "tank_gas": tank_gas,
            "tank_capacity_mode": tank_capacity_mode,
            "tank_liters": float(tank_liters),
            "tank_spec": tank_spec,
            "hourly_usage_mode": str(hourly_usage_mode),
            "hourly_usage_kg": float(hourly_usage_kg),
            "hourly_usage_nm3": float(hourly_usage_nm3),
            "operating_hours": float(operating_hours),
            "operating_days_per_month": float(operating_days),
            "auto_monthly_from_cycle": bool(auto_monthly_from_cycle),
            "tank_price": tank_price,
            "monthly_usage_kg": monthly_usage,
            "vaporizer_capacity": vap_cap,
            "vaporizer_qty_note": vap_note,
            "vaporizer_price": vap_price,
            "construction_cost": construction,
            "purchase_unit": purchase_unit,
            "logistics_unit": float(logistics_unit),
            "supply_unit": supply_unit,
            "interest_rate": interest,
            "mgmt_rate": mgmt_rate,
            "depreciation_months": dep_months,
            "equipment_rent": rent,
            "rent_count": rent_count,
            "logi_km": float(logi_km),
            "logi_fuel_price": float(logi_fuel),
            "logi_efficiency": float(logi_eff),
            "logi_toll": float(logi_toll),
            "logi_roundtrips": float(logi_rt),
            "logi_supply_kg": float(logi_kg),
            "logi_origin": str(logi_origin),
            "logi_dest": str(logi_dest),
        }
        st.session_state["profit_inputs"] = new_p
        save_profit_inputs(new_p)
        st.success("저장되었습니다. 아래 결과가 갱신됩니다.")
    # 화면 결과도 현재 입력(자동 반영된 물류비 포함) 기준으로 표시
    p = {
        "project_name": project_name,
        "tank_gas": tank_gas,
        "tank_capacity_mode": tank_capacity_mode,
        "tank_liters": float(tank_liters),
        "tank_spec": tank_spec,
        "hourly_usage_mode": str(hourly_usage_mode),
        "hourly_usage_kg": float(hourly_usage_kg),
        "hourly_usage_nm3": float(hourly_usage_nm3),
        "operating_hours": float(operating_hours),
        "operating_days_per_month": float(operating_days),
        "auto_monthly_from_cycle": bool(auto_monthly_from_cycle),
        "tank_price": tank_price,
        "monthly_usage_kg": monthly_usage,
        "vaporizer_capacity": vap_cap,
        "vaporizer_qty_note": vap_note,
        "vaporizer_price": vap_price,
        "construction_cost": construction,
        "purchase_unit": purchase_unit,
        "logistics_unit": float(logistics_unit),
        "supply_unit": supply_unit,
        "interest_rate": interest,
        "mgmt_rate": mgmt_rate,
        "depreciation_months": dep_months,
        "equipment_rent": rent,
        "rent_count": rent_count,
        "logi_km": float(logi_km),
        "logi_fuel_price": float(logi_fuel),
        "logi_efficiency": float(logi_eff),
        "logi_toll": float(logi_toll),
        "logi_roundtrips": float(logi_rt),
        "logi_supply_kg": float(logi_kg),
        "logi_origin": str(logi_origin),
        "logi_dest": str(logi_dest),
    }
    st.session_state["profit_inputs"] = p
    r = compute_profitability(p)
    st.markdown("---")
    st.markdown(
        f"<div class='sub-header dashboard-tab-panel-head'>◆ [{html.escape(str(p.get('project_name') or '프로젝트'))}] 투자대비 수익성 분석 보고</div>",
        unsafe_allow_html=True,
    )
    st.caption(
        f"TANK: {p.get('tank_gas','')} {float(p.get('tank_liters') or 0):,.0f} L → "
        f"{parse_tank_capacity_kg(p.get('tank_spec')):,.0f} kg · "
        f"기화기 {p.get('vaporizer_capacity',0):,.0f} Nm3/hr {p.get('vaporizer_qty_note','')} · "
        f"년 사용량(자동) {r['yearly_usage']:,.0f} kg"
    )
    m1, m2, m3, m4 = st.columns(4)
    m1.markdown(
        f"<div class='metric-box'><div class='metric-label'>합계 투자액 (C9+C16+C18)</div>"
        f"<div class='metric-value'>{r['total_invest']:,.0f} 원</div></div>",
        unsafe_allow_html=True,
    )
    m2.markdown(
        f"<div class='metric-box'><div class='metric-label'>매출이익 (원/kg)</div>"
        f"<div class='metric-value'>{r['margin_kg']:,.1f} 원/kg</div></div>",
        unsafe_allow_html=True,
    )
    m3.markdown(
        f"<div class='metric-box'><div class='metric-label'>월평균 이익금</div>"
        f"<div class='metric-value' style='color:{'#059669' if r['monthly_profit']>=0 else '#E11D48'};'>"
        f"{r['monthly_profit']:,.0f} 원/월</div></div>",
        unsafe_allow_html=True,
    )
    m4.markdown(
        f"<div class='metric-box'><div class='metric-label'>최근 3개월 매출 실이익</div>"
        f"<div class='metric-value'>{r['three_month']:,.0f} 원</div></div>",
        unsafe_allow_html=True,
    )
    left, right = st.columns(2)
    def _pf_i(v):
        try:
            return f"{int(round(float(v or 0))):,}"
        except Exception:
            return "0"
    with left:
        st.markdown("**◎ 단가 · 투자 요약**")
        summary_df = pd.DataFrame(
            [
                {"항목": "TANK 구입가", "값": _pf_i(p["tank_price"]), "단위": "원"},
                {"항목": "기화기 구입가", "값": _pf_i(p["vaporizer_price"]), "단위": "원"},
                {"항목": "공사비용", "값": _pf_i(p["construction_cost"]), "단위": "원"},
                {"항목": "합계 금액", "값": _pf_i(r["total_invest"]), "단위": "원"},
                {"항목": "년 사용량", "값": _pf_i(r["yearly_usage"]), "단위": "kg"},
                {"항목": "월 평균 공급량", "값": _pf_i(p["monthly_usage_kg"]), "단위": "kg"},
                {"항목": "매입단가", "값": _pf_i(p["purchase_unit"]), "단위": "원/kg"},
                {"항목": "물류비", "값": _pf_i(p["logistics_unit"]), "단위": "원/kg"},
                {"항목": "공급단가", "값": _pf_i(p["supply_unit"]), "단위": "원/kg"},
                {"항목": "매출이익", "값": _pf_i(r["margin_kg"]), "단위": "원/kg"},
                {"항목": "물류비 계산(P18)", "값": _pf_i(r["logi_per_kg"]), "단위": "원/kg"},
            ]
        )
        st.dataframe(summary_df, hide_index=True, width="stretch", height=420)
    with right:
        st.markdown("**◎ 사용량 대비 영업이익 (월)**")
        result_df = pd.DataFrame(
            [
                {"항목": "월평균 매출이익", "계산": "월사용량 × 매출이익", "값": _pf_i(r["monthly_gross"]), "단위": "원/월"},
                {"항목": "장비 감가상각", "계산": f"투자합계 ÷ {p['depreciation_months']:.0f}", "값": _pf_i(r["depreciation"]), "단위": "원/월"},
                {"항목": "금융비", "계산": "원금 × 이자율 ÷ 12", "값": _pf_i(r["finance"]), "단위": "원/월"},
                {"항목": "월 매출", "계산": "월사용량 × 공급단가", "값": _pf_i(r["monthly_sales"]), "단위": "원/월"},
                {"항목": "일반관리비", "계산": f"월매출 × {p['mgmt_rate']*100:.1f}%", "값": _pf_i(r["mgmt"]), "단위": "원/월"},
                {"항목": "투자비용(상각+금융)", "계산": "감가상각 + 금융비", "값": _pf_i(r["invest_cost"]), "단위": "원/월"},
                {"항목": "월평균 이익금", "계산": "매출이익 − 투자비용 − 관리비", "값": _pf_i(r["monthly_profit"]), "단위": "원/월"},
                {"항목": "최근 3개월 실이익", "계산": "월이익×3 + 임대료×횟수", "값": _pf_i(r["three_month"]), "단위": "원"},
            ]
        )
        st.dataframe(result_df, hide_index=True, width="stretch", height=420)
    st.info(
        "적용 함수: `년사용량=월×12` · `합계=탱크+기화기+공사` · `매출이익=공급−(매입+물류)` · "
        "`관리비=월매출×14.5%` · `감가=투자÷120` · `금융=원금×이자÷12` · "
        "`월이익=매출이익−(감가+금융)−관리비` · `3개월=월이익×3+(임대×횟수)` · "
        "`물류원/kg=(KM×유류비/연비+통행료)×왕복÷KG`"
    )
    # Tab9 전용 — 스크린샷형 수익성 보고서 엑셀 (동일 입력이면 캐시 재사용)
    _route_for_xlsx = st.session_state.get("pf_route_info") or {}
    _diesel_for_xlsx = diesel_info if isinstance(diesel_info, dict) else {}
    try:
        _profit_xlsx = _cached_profitability_report_excel(
            json.dumps(p, ensure_ascii=False, sort_keys=True, default=str),
            json.dumps(r, ensure_ascii=False, sort_keys=True, default=str),
            json.dumps(_route_for_xlsx, ensure_ascii=False, sort_keys=True, default=str),
            json.dumps(_diesel_for_xlsx, ensure_ascii=False, sort_keys=True, default=str),
        )
        st.download_button(
            "📥 수익성 분석 보고서 엑셀 내보내기",
            data=_profit_xlsx,
            file_name=f"수익성분석_{str(p.get('project_name') or '보고서')}_{datetime.date.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
            key="pf_report_xlsx_btn",
            help="스크린샷과 같은 보고서 형태. 거리·통행료·경유시세 설명은 작은 글씨로 포함됩니다.",
        )
    except Exception as _xlsx_err:
        st.warning(f"보고서 엑셀 생성 실패: {_xlsx_err}")

with tab9:
    _render_profitability_analysis_tab(latest_update_str)
