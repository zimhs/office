import io
import re
import sys
import subprocess
import urllib.parse
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px

# 페이지 및 Styler 가동 한도 설정
pd.set_option("styler.render.max_elements", 2000000)
st.set_page_config(
    page_title="통합 영업 분석 대시보드",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ==========================================
# 1. 아이패드/모바일 최적화 CSS Injection
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
            /* multiselect 경고 문구 및 안내 숨기기 */
            div[data-baseweb="select"] + div:has(span) {
                display: none !important;
            }
            div[data-testid="stMultiSelect"] [data-testid="stWidgetInstructions"] {
                display: none !important;
            }
            small[data-testid="stCaptionContainer"] {
                display: none !important;
            }
            
            /* 아이패드 터치 스크롤 및 전체 레이아웃 최적화 */
            html, body, .stApp {
                background-color: #F8FAFC !important;
                color: #1E293B !important;
                font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
                -webkit-tap-highlight-color: transparent;
                -webkit-overflow-scrolling: touch;
            }
            [data-testid="stSidebar"] {
                background-color: #F1F5F9 !important;
                border-right: 1px solid #E2E8F0;
            }
            [data-testid="stSidebar"] .block-container { padding-top: 1.5rem; }
            [data-testid="stSidebar"] p, [data-testid="stSidebar"] span, 
            [data-testid="stSidebar"] label, [data-testid="stSidebar"] .stMarkdown {
                color: #334155 !important;
            }

            /* 상단 검색창/필터 높이 및 터치 타겟 반응형 정렬 */
            div[data-testid="column"] {
                align-self: flex-start;
            }
            div[data-testid="stTextInput"], div[data-testid="stMultiSelect"], div[data-testid="stSelectbox"] {
                min-height: 75px;
            }
            div[data-testid="stTextInput"] label, div[data-testid="stMultiSelect"] label, div[data-testid="stSelectbox"] label {
                font-size: 13px !important;
                font-weight: 600 !important;
                color: #475569 !important;
                white-space: nowrap !important;
                overflow: hidden;
                text-overflow: ellipsis;
                margin-bottom: 4px !important;
            }

            /* KPI 메트릭 카드 */
            .metric-box {
                background: #FFFFFF;
                padding: 16px 20px;
                border-radius: 12px;
                border: 1px solid #E2E8F0;
                box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.03), 0 2px 4px -1px rgba(0, 0, 0, 0.02);
                margin-bottom: 8px;
            }
            .metric-label {
                color: #64748B;
                font-size: 13px;
                font-weight: 600;
                margin-bottom: 6px;
            }
            .metric-value {
                color: #0F172A;
                font-size: 22px;
                font-weight: 700;
            }
            
            /* 대기업 스타일 카드 패널 */
            .ent-card {
                background-color: #FFFFFF;
                border: 1px solid #E2E8F0;
                border-radius: 12px;
                padding: 20px;
                box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.05);
                margin-bottom: 20px;
            }
            
            /* 서브 헤더 */
            .sub-header {
                color: #1E3A8A;
                font-size: 17px;
                font-weight: 700;
                margin-top: 15px;
                margin-bottom: 12px;
                border-left: 4px solid #2563EB;
                padding-left: 10px;
            }

            /* 아이패드 테이블 가로 스크롤 최적화 */
            div[data-testid="stDataFrame"] {
                -webkit-overflow-scrolling: touch;
                border-radius: 8px;
            }
        </style>
    """,
        unsafe_allow_html=True,
    )


# ==========================================
# 2. 연간총합 색상 강조 유틸리티 함수
# ==========================================
def style_table_with_totals(
    df, cmap="Blues", total_col="연간총합", axis_grad=1
):
    """
    DataFrame의 연간총합 컬럼에 황금색 하이라이트를 주고 나머지 숫자에 배경 그라데이션을 적용하는 함수
    """
    if df.empty:
        return df

    num_cols = df.select_dtypes(include=[np.number]).columns
    grad_cols = [c for c in num_cols if c != total_col]

    styler = df.style.format("{:,.0f}")

    if grad_cols:
        styler = styler.background_gradient(
            cmap=cmap, subset=grad_cols, axis=axis_grad
        )

    if total_col in df.columns:
        styler = styler.set_properties(
            subset=[total_col],
            **{
                "background-color": "#FEF3C7",
                "color": "#0F172A",
                "font-weight": "bold",
                "border-left": "2px solid #F59E0B",
            },
        )

    return styler


# ==========================================
# 3. 날짜 파싱 및 데이터 정규화 유틸리티
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
        parts_df = parts_df.apply(lambda col: col.str.strip()).replace(
            "", None
        )
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
    """
    품목명 정규화 및 N2(liter) -> kg 환산 처리 (1 Liter N2 = 0.808 kg N2)
    """
    if "품목명" not in df.columns or df.empty:
        return df

    p_str = df["품목명"].astype(str)
    p_upper = p_str.str.upper().str.replace(" ", "")

    is_n2 = p_upper.str.contains("N2", na=False) | p_str.str.contains(
        "질소", na=False
    )
    is_liter = p_upper.str.contains("LITER|L", na=False) | p_str.str.contains(
        "리터", na=False
    )
    is_n2_liter = is_n2 & is_liter & ~p_upper.str.contains("BULK", na=False)

    if "출고량" in df.columns and is_n2_liter.any():
        df.loc[is_n2_liter, "출고량"] = df.loc[is_n2_liter, "출고량"] * 0.808

    is_ar = p_upper.str.contains("AR", na=False) | p_str.str.contains(
        "아르곤|아르", na=False
    )
    is_co2 = p_upper.str.contains("CO2", na=False) | p_str.str.contains(
        "탄산|이산화탄소", na=False
    )
    is_o2 = (~is_co2) & (
        p_upper.str.contains("O2", na=False) | p_str.str.contains("산소", na=False)
    )

    df.loc[is_ar, "품목명"] = "AR (kg, Bulk)"
    df.loc[is_co2, "품목명"] = "CO2 (kg, Bulk)"
    df.loc[is_o2, "품목명"] = "O2 (kg, Bulk)"
    df.loc[is_n2, "품목명"] = "N2 (kg, Bulk)"

    return df


def open_macos_note(client_name):
    if sys.platform != "darwin":
        return False

    script = f"""
    tell application "Notes"
        activate
        try
            set targetNote to first note of folder "거래처" whose name contains "{client_name}"
            show targetNote
        on error
            tell folder "거래처"
                make new note with properties {{name:"{client_name}", body:"--- {client_name} 영업 및 특사항 메모 ---\\n\\n"}}
            end tell
            set targetNote to first note of folder "거래처" whose name contains "{client_name}"
            show targetNote
        end try
    end tell
    """
    try:
        subprocess.run(["osascript", "-e", script], check=True)
        return True
    except Exception:
        return False


@st.cache_data
def convert_dfs_to_excel(dfs_dict):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
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
    return output.getvalue()


# 인라인 품목 상세 분석 렌더링 함수 (Plotly 차트 적용)
def render_item_detail_section(
    item_name, df_item, view_type, unique_key_suffix
):
    unit_label = "만 원" if view_type == "매출액" else "kg / 개"
    val_series = (
        df_item["매출액"] / 10000 if view_type == "매출액" else df_item["출고량"]
    )
    total_sum = val_series.sum()

    st.markdown(
        f"""
        <div style="background-color: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 12px; padding: 20px; margin-top: 15px; margin-bottom: 25px; box-shadow: 0 1px 3px rgba(0,0,0,0.05);">
            <h4 style="color: #1E3A8A; margin-top: 0; margin-bottom: 12px; font-weight: 700;">📦 [{item_name}] 상세 분석 보고서 ({view_type})</h4>
            <div style="background-color: #F8FAFC; border-radius: 8px; padding: 12px 16px; margin-bottom: 16px; border: 1px solid #E2E8F0; display: flex; justify-content: space-between; align-items: center;">
                <span style="color: #475569; font-weight: 600; font-size: 14px;">총 {view_type} 누계</span>
                <span style="color: #2563EB; font-weight: 800; font-size: 20px;">{total_sum:,.0f} <span style="font-size: 14px; font-weight: 500; color: #64748B;">{unit_label}</span></span>
            </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        "<h5 style='color: #1E3A8A; font-weight: 700;'>📈 1. 연도별 요약</h5>",
        unsafe_allow_html=True,
    )
    col_c, col_t = st.columns([1, 1])
    yr_summary = (
        df_item.assign(표시값=val_series)
        .groupby("연도", as_index=False)["표시값"]
        .sum()
    )

    with col_c:
        st.markdown(
            "<p style='font-size: 13px; font-weight: 700; color: #334155;'>📊 연도별 추이</p>",
            unsafe_allow_html=True,
        )
        if not yr_summary.empty and total_sum > 0:
            fig_bar = px.bar(
                yr_summary,
                x="연도",
                y="표시값",
                text_auto=",.0f",
                color_discrete_sequence=["#2563EB"],
            )
            fig_bar.update_layout(
                height=200,
                margin=dict(l=10, r=10, t=10, b=10),
                xaxis_title=None,
                yaxis_title=None,
                xaxis=dict(type="category"),
            )
            st.plotly_chart(
                fig_bar,
                use_container_width=True,
                key=f"chart_{unique_key_suffix}",
            )
        else:
            st.info("데이터가 없습니다.")

    with col_t:
        st.markdown(
            f"<p style='font-size: 13px; font-weight: 700; color: #334155;'>📋 연도별 {view_type} 요약</p>",
            unsafe_allow_html=True,
        )
        if not yr_summary.empty:
            yr_summary["비중"] = (
                yr_summary["표시값"] / total_sum * 100
            ).map("{:.1f}%".format)
            yr_summary_disp = yr_summary.rename(
                columns={"표시값": f"{view_type}({unit_label})"}
            )
            st.dataframe(
                yr_summary_disp.style.format(
                    {f"{view_type}({unit_label})": "{:,.0f}"}
                ),
                hide_index=True,
                use_container_width=True,
                height=200,
                key=f"yr_df_{unique_key_suffix}",
            )
        else:
            st.info("데이터가 없습니다.")

    st.markdown("<hr style='margin: 20px 0;'>", unsafe_allow_html=True)
    st.markdown(
        "<h5 style='color: #1E3A8A; font-weight: 700;'>📅 2. 월별 상세 추이</h5>",
        unsafe_allow_html=True,
    )
    all_months = [f"{i:02d}월" for i in range(1, 13)]
    pivot_detail = (
        df_item.assign(표시값=val_series)
        .pivot_table(
            index="연도", columns="월", values="표시값", aggfunc="sum"
        )
        .fillna(0)
    )

    pivot_detail = pivot_detail.reindex(columns=all_months, fill_value=0)
    pivot_detail["연간총합"] = pivot_detail.sum(axis=1)

    if not pivot_detail.empty:
        st.dataframe(
            style_table_with_totals(pivot_detail, cmap="Blues"),
            use_container_width=True,
            height=200,
            key=f"m_df_{unique_key_suffix}",
        )
    else:
        st.info("상세 월별 데이터가 없습니다.")

    st.markdown("</div>", unsafe_allow_html=True)


# ==========================================
# 4. 데이터 로딩 & 메모리 캐싱
# ==========================================
@st.cache_data(show_spinner="주소록을 읽어오는 중입니다...")
def load_address_file(address_bytes):
    if not address_bytes:
        return {}
    try:
        for enc in ["utf-8-sig", "cp949", "euc-kr", "utf-8"]:
            try:
                temp_addr = pd.read_csv(io.BytesIO(address_bytes), encoding=enc)
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


@st.cache_data(show_spinner="데이터를 파싱 및 캐싱 중입니다...")
def load_uploaded_files(uploaded_files):
    if not uploaded_files:
        return pd.DataFrame()

    df_list = []
    for file in uploaded_files:
        try:
            content = file.getvalue()
            decoded_text = None
            for enc in ["cp949", "euc-kr", "utf-8-sig", "utf-8"]:
                try:
                    decoded_text = content.decode(enc)
                    break
                except UnicodeDecodeError:
                    continue
            if decoded_text is None:
                decoded_text = content.decode("utf-8", errors="ignore")

            lines = [
                line for line in decoded_text.splitlines() if line.strip()
            ]
            if not lines:
                continue

            header_idx = 0
            for i, line in enumerate(lines[:30]):
                if any(
                    k in line
                    for k in [
                        "거래처",
                        "상호",
                        "품목",
                        "제품",
                        "매출액",
                        "담당",
                        "일자",
                        "금액",
                        "단가",
                    ]
                ):
                    header_idx = i
                    break

            df = pd.read_csv(io.StringIO("\n".join(lines[header_idx:])))
            df.columns = df.columns.astype(str).str.strip()
            cols = list(df.columns)

            def find_col(priority_keywords, exclude_keywords=[]):
                for kw in priority_keywords:
                    for c in cols:
                        if any(ex in c for ex in exclude_keywords):
                            continue
                        if kw == c or kw in c:
                            return c
                return None

            c_staff = find_col(
                ["담당자명", "영업담당", "담당자", "영업사원", "담당"],
                ["코드", "ID", "번호"],
            )
            c_client = find_col(
                [
                    "거래처명",
                    "상호명",
                    "고객명",
                    "회사명",
                    "거래처",
                    "상호",
                    "고객",
                ],
                ["코드", "ID", "번호", "담당", "영업"],
            )
            c_item = find_col(
                ["품목명", "제품명", "상품명", "품목", "제품"],
                ["코드", "ID", "번호", "규격"],
            )
            c_sales = find_col(
                ["매출액", "금액", "매출"], ["일", "자", "수량", "량", "단가"]
            )
            c_qty = find_col(
                ["출고량", "수량", "출고"], ["액", "금액", "단가"]
            )
            c_price = find_col(
                ["단가", "단 가", "판매단가", "공급단가"],
                ["액", "금액", "수량", "량"],
            )
            c_date = find_col(["매출일자", "매출일", "일자", "날짜", "출고일"])

            rename_dict = {}
            if c_client:
                rename_dict[c_client] = "거래처"
            if c_item:
                rename_dict[c_item] = "품목명"
            if c_staff:
                rename_dict[c_staff] = "담당자"
            if c_sales:
                rename_dict[c_sales] = "매출액"
            if c_qty:
                rename_dict[c_qty] = "출고량"
            if c_price:
                rename_dict[c_price] = "단가"
            if c_date:
                rename_dict[c_date] = "매출일자_raw"

            df = df.rename(columns=rename_dict)

            for req in ["거래처", "품목명", "담당자"]:
                if req not in df.columns:
                    df[req] = "미지정"

            file_year = next(
                (
                    y
                    for y in [
                        "2020",
                        "2021",
                        "2022",
                        "2023",
                        "2024",
                        "2025",
                        "2026",
                    ]
                    if y in file.name
                ),
                "2026",
            )
            date_col = (
                "매출일자_raw" if "매출일자_raw" in df.columns else df.columns[0]
            )

            df["매출일_dt"] = parse_date_series_robust(
                df[date_col], default_year=file_year
            )

            df["매출액"] = (
                pd.to_numeric(
                    df["매출액"]
                    .astype(str)
                    .str.replace(r"[^\d.-]", "", regex=True),
                    errors="coerce",
                ).fillna(0)
                if "매출액" in df.columns
                else 0
            )
            df["출고량"] = (
                pd.to_numeric(
                    df["출고량"]
                    .astype(str)
                    .str.replace(r"[^\d.-]", "", regex=True),
                    errors="coerce",
                ).fillna(0)
                if "출고량" in df.columns
                else 0
            )
            df["단가"] = (
                pd.to_numeric(
                    df["단가"]
                    .astype(str)
                    .str.replace(r"[^\d.-]", "", regex=True),
                    errors="coerce",
                ).fillna(0)
                if "단가" in df.columns
                else 0
            )

            df["거래처"] = df["거래처"].fillna("미지정").astype(str).str.strip()
            df["담당자"] = df["담당자"].fillna("미지정").astype(str).str.strip()

            df = normalize_items_vectorized(df)
            df = df.dropna(subset=["매출일_dt"])

            if not df.empty:
                df_list.append(df)
        except Exception as e:
            st.sidebar.error(f"파일 읽기 오류 ({file.name}): {e}")

    result_df = (
        pd.concat(df_list, ignore_index=True) if df_list else pd.DataFrame()
    )

    if not result_df.empty and "거래처" in result_df.columns:
        result_df["거래처"] = result_df["거래처"].replace(
            {"Z바인컴퍼니": "아이스푸드앤바인(구.바인컴퍼니)"}
        )

    return result_df


# ==========================================
# 5. 메인 실행 흐름 (UI 구성)
# ==========================================
inject_custom_css()

st.title("📊 통합 영업 분석 대시보드")
st.markdown(
    "<p style='color: #64748B; margin-bottom: 15px;'>실시간 영업 데이터 모니터링 및 품목·거래처별 다차원 분석 시스템</p>",
    unsafe_allow_html=True,
)

st.sidebar.header("📁 데이터 업로드")
address_file = st.sidebar.file_uploader("거래처 주소록 (CSV)", type=["csv"])
uploaded_files = st.sidebar.file_uploader(
    "매출 데이터 (다중 업로드)", type=["csv"], accept_multiple_files=True
)

addr_dict = load_address_file(address_file.getvalue()) if address_file else {}
full_df = (
    load_uploaded_files(uploaded_files) if uploaded_files else pd.DataFrame()
)

if not full_df.empty:
    is_deposit_row = full_df["품목명"].astype(str).str.contains("입금", na=False)
    full_df = full_df[~is_deposit_row].copy()

# 주요 4대 품목 목록
target_items = [
    "AR (kg, Bulk)",
    "O2 (kg, Bulk)",
    "CO2 (kg, Bulk)",
    "N2 (kg, Bulk)",
]

if not full_df.empty:
    filter_container = st.container()
    with filter_container:
        fc1, fc2, fc3, fc4, fc5 = st.columns([1, 1, 1, 1, 1])

        start_date = fc1.text_input("📅 조회 시작", "200101")
        end_date = fc2.text_input("📅 조회 종료", "261231")

        start_dt = pd.to_datetime(start_date, format="%y%m%d", errors="coerce")
        end_dt = pd.to_datetime(end_date, format="%y%m%d", errors="coerce")

        if pd.isna(start_dt):
            start_dt = pd.Timestamp("2000-01-01")
        if pd.isna(end_dt):
            end_dt = pd.Timestamp("2099-12-31")

        df_base = full_df[
            (full_df["매출일_dt"] >= start_dt) & (full_df["매출일_dt"] <= end_dt)
        ].copy()

        selected_staff = fc3.multiselect(
            "👤 담당자",
            sorted(df_base["담당자"].unique()) if not df_base.empty else [],
        )
        df_staff_filtered = (
            df_base[df_base["담당자"].isin(selected_staff)]
            if selected_staff
            else df_base.copy()
        )

        all_clients = (
            sorted(df_staff_filtered["거래처"].unique())
            if not df_staff_filtered.empty
            else []
        )

        selected_client_list = fc4.multiselect(
            "🏢 거래처",
            options=all_clients,
            max_selections=1,
            placeholder="거래처 검색...",
        )
        selected_client = (
            selected_client_list[0] if selected_client_list else "전체 거래처"
        )

        df_client_filtered = (
            df_staff_filtered[df_staff_filtered["거래처"] == selected_client]
            if selected_client != "전체 거래처"
            else df_staff_filtered.copy()
        )

        available_items = (
            sorted(df_client_filtered["품목명"].unique())
            if not df_client_filtered.empty
            else []
        )
        selected_item = fc5.multiselect("📦 품목명", available_items)

    df_f = (
        df_client_filtered[df_client_filtered["품목명"].isin(selected_item)]
        if selected_item
        else df_client_filtered.copy()
    )

    if not df_f.empty:
        df_f["연도"] = df_f["매출일_dt"].dt.year.astype(str)
        df_f["월"] = df_f["매출일_dt"].dt.strftime("%m월")
        df_f["분기"] = df_f["매출일_dt"].dt.to_period("Q").astype(str).str[-2:]
        existing_years = sorted(df_f["연도"].unique())
        sorted_cols = [
            (y, m)
            for y in existing_years
            for m in [f"{i:02d}월" for i in range(1, 13)]
        ]
    else:
        existing_years = []
        sorted_cols = []

    all_months = [f"{i:02d}월" for i in range(1, 13)]

    # 1. 전체 연도별 월 매출 (연간총합 포함)
    pivot_m = pd.DataFrame()
    if not df_f.empty:
        pivot_m = (
            df_f.pivot_table(
                index="연도", columns="월", values="매출액", aggfunc="sum"
            ).fillna(0)
            * 1.1
            / 10000
        )
        pivot_m = pivot_m.reindex(columns=all_months, fill_value=0)
        pivot_m["연간총합"] = pivot_m.sum(axis=1)

    # 2. 거래처별 월별 매출 (연간총합 포함)
    client_pivot = pd.DataFrame()
    years = (
        sorted(full_df["연도"].unique())
        if not full_df.empty and "연도" in full_df.columns
        else (sorted(df_f["연도"].unique()) if not df_f.empty else ["2026"])
    )
    if not df_f.empty:
        df_f["연도월_정렬"] = (
            df_f["연도"].astype(str).str[2:] + "년 " + df_f["월"].astype(str)
        )
        desired_order = [f"{y[2:]}년 {m}" for m in all_months for y in years]

        client_pivot_raw = (
            df_f.pivot_table(
                index="거래처",
                columns="연도월_정렬",
                values="매출액",
                aggfunc="sum",
            ).fillna(0)
            / 10000
        )
        actual_cols = [
            c for c in desired_order if c in client_pivot_raw.columns
        ]
        client_pivot = client_pivot_raw.reindex(
            columns=actual_cols, fill_value=0
        )
        client_pivot["연간총합"] = client_pivot.sum(axis=1)

    # 3. 품목 및 단가 분석용 데이터
    sales_p = pd.DataFrame()
    qty_p = pd.DataFrame()
    unit_price_p = pd.DataFrame()
    valid_cols = [c for c in sorted_cols if not df_f.empty]

    if not df_f.empty:
        sales_raw_p = df_f.pivot_table(
            index="품목명",
            columns=["연도", "월"],
            values="매출액",
            aggfunc="sum",
        ).fillna(0)

        qty_raw_p = df_f.pivot_table(
            index="품목명",
            columns=["연도", "월"],
            values="출고량",
            aggfunc="sum",
        ).fillna(0)

        sales_expanded_data = {}
        qty_expanded_data = {}

        for yr in existing_years:
            for m in all_months:
                col_key = (yr, m)
                sales_expanded_data[col_key] = (
                    sales_raw_p[col_key] if col_key in sales_raw_p.columns else 0
                )
                qty_expanded_data[col_key] = (
                    qty_raw_p[col_key] if col_key in qty_raw_p.columns else 0
                )

            yr_sales_sum = sum(
                sales_raw_p[(yr, m)]
                for m in all_months
                if (yr, m) in sales_raw_p.columns
            )
            sales_expanded_data[(yr, "연간총합")] = yr_sales_sum

            yr_qty_sum = sum(
                qty_raw_p[(yr, m)]
                for m in all_months
                if (yr, m) in qty_raw_p.columns
            )
            qty_expanded_data[(yr, "연간총합")] = yr_qty_sum

        sales_p = pd.DataFrame(sales_expanded_data, index=sales_raw_p.index)
        qty_p = pd.DataFrame(qty_expanded_data, index=qty_raw_p.index)

        unit_price_p = df_f[df_f["단가"] > 0].pivot_table(
            index="품목명",
            columns=["연도", "월"],
            values="단가",
            aggfunc="first",
        )
        if valid_cols and not unit_price_p.empty:
            unit_price_p = unit_price_p.reindex(
                index=sales_raw_p.index, columns=valid_cols, fill_value=0
            )

    # 4. 담당자별 매출 (연간총합 포함)
    staff_pivot = pd.DataFrame()
    if not df_f.empty:
        staff_pivot = (
            df_f.pivot_table(
                index="담당자",
                columns=["연도", "월"],
                values="매출액",
                aggfunc="sum",
            ).fillna(0)
            / 10000
        )
        if valid_cols and not staff_pivot.empty:
            staff_pivot = staff_pivot.reindex(columns=valid_cols, fill_value=0)
            staff_pivot["연간총합"] = staff_pivot.sum(axis=1)

    # 5. 상세 거래 내역
    df_detail = pd.DataFrame()
    if not df_f.empty:
        detail_cols = [
            "매출일_dt",
            "담당자",
            "거래처",
            "품목명",
            "출고량",
            "단가",
            "매출액",
        ]
        df_detail = df_f[detail_cols].copy()

    st.sidebar.markdown("---")
    st.sidebar.subheader("📥 엑셀 내보내기")
    if not df_f.empty:
        sheets_dict = {
            "연도별_월매출(만원)": (pivot_m, True),
            "거래처별_월별매출(만원)": (client_pivot, True),
            "품목별_매출액(만원)": (sales_p / 10000, True),
            "품목별_출고량": (qty_p, True),
            "품목별_적용단가": (unit_price_p, True),
            "담당자별_매출(만원)": (staff_pivot, True),
            "상세거래내역": (df_detail, False),
        }
        excel_data = convert_dfs_to_excel(sheets_dict)
        st.sidebar.download_button(
            label="📊 전체 분석 시트별 엑셀 다운로드",
            data=excel_data,
            file_name="통합영업분석_시트별보고서.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    total_sales = df_f["매출액"].sum() if not df_f.empty else 0
    total_qty = df_f["출고량"].sum() if not df_f.empty else 0
    label_suffix = "(선택 품목)" if selected_item else "(전체 품목)"

    m1, m2 = st.columns(2)
    m1.markdown(
        f"""<div class="metric-box"><div class="metric-label">💰 총 매출 합계 {label_suffix}</div><div class="metric-value">{total_sales:,.0f} <span style="font-size: 15px; font-weight: normal; color: #64748B;">원</span></div></div>""",
        unsafe_allow_html=True,
    )
    m2.markdown(
        f"""<div class="metric-box"><div class="metric-label">📦 총 출고량 {label_suffix}</div><div class="metric-value">{total_qty:,.0f} <span style="font-size: 15px; font-weight: normal; color: #64748B;">개/kg</span></div></div>""",
        unsafe_allow_html=True,
    )

    st.markdown("<br>", unsafe_allow_html=True)

    tab1, tab2, tab3, tab4 = st.tabs(
        [
            "📌 영업 종합 요약",
            "🏢 거래처 분석",
            "📦 품목 및 단가 분석",
            "👤 담당자 & 상세내역",
        ]
    )

    # ------------------------------------
    # TAB 1: 영업 종합 요약
    # ------------------------------------
    with tab1:
        st.markdown(
            '<div class="sub-header">📈 전체 월별 매출 추이 및 연도별 비교 (만 원 단위)</div>',
            unsafe_allow_html=True,
        )
        col_table1, col_chart1 = st.columns([1, 1])

        with col_table1:
            st.markdown("**📋 연도별 전체 월 매출 데이터 (VAT 포함, 만 원)**")
            if not pivot_m.empty:
                st.dataframe(
                    style_table_with_totals(pivot_m, cmap="Blues"),
                    use_container_width=True,
                    height=280,
                )
            else:
                st.info("조회된 데이터가 없습니다.")

        with col_chart1:
            st.markdown("**📊 연도별 월 매출 추이 비교**")
            if not pivot_m.empty:
                # Plotly 데이터 꺾은선 차트 (아이패드 WebKit 렌더링 최적화)
                df_chart = pivot_m.drop(
                    columns=["연간총합"], errors="ignore"
                ).reset_index()
                df_chart_melted = df_chart.melt(
                    id_vars=["연도"], var_name="월", value_name="매출액(만원)"
                )
                fig_m = px.line(
                    df_chart_melted,
                    x="월",
                    y="매출액(만원)",
                    color="연도",
                    markers=True,
                )
                fig_m.update_layout(
                    height=280,
                    margin=dict(l=10, r=10, t=10, b=10),
                    hovermode="x unified",
                    xaxis_title=None,
                    yaxis_title=None,
                )
                st.plotly_chart(fig_m, use_container_width=True)
            else:
                st.info("차트를 표시할 데이터가 없습니다.")

        # 주요 4대 품목 상세 인라인 분석
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(
            '<div class="sub-header">🎯 주요 4대 품목 (AR, O2, CO2, N2) 상세 분석</div>',
            unsafe_allow_html=True,
        )

        for item in target_items:
            df_item = df_f[df_f["품목명"] == item]
            if not df_item.empty:
                render_item_detail_section(
                    item_name=item,
                    df_item=df_item,
                    view_type="출고량",
                    unique_key_suffix=f"tab1_{item.split()[0]}",
                )

    # ------------------------------------
    # TAB 2: 거래처 분석
    # ------------------------------------
    with tab2:
        st.markdown(
            '<div class="sub-header">🏢 거래처별 월별 매출 분석 (만 원)</div>',
            unsafe_allow_html=True,
        )
        if not client_pivot.empty:
            st.dataframe(
                style_table_with_totals(client_pivot, cmap="Greens"),
                use_container_width=True,
                height=350,
            )
        else:
            st.info("거래처별 데이터가 없습니다.")

        if selected_client != "전체 거래처":
            st.markdown("---")
            st.markdown(f"### 📍 **{selected_client}** 상세 내역 및 영업 메모")

            c_info1, c_info2 = st.columns([2, 1])
            with c_info1:
                client_addr = addr_dict.get(selected_client, "주소록 정보 없음")
                st.write(f"**🏢 거래처명:** {selected_client}")
                st.write(f"**📍 등록 주소:** {client_addr}")

            with c_info2:
                if sys.platform == "darwin":
                    if st.button(
                        f"📝 macOS 메모 연동 ({selected_client})",
                        use_container_width=True,
                    ):
                        if open_macos_note(selected_client):
                            st.success("macOS 메모 앱이 실행되었습니다.")
                        else:
                            st.error("메모 연동 실패")

            df_curr_client = df_f[df_f["거래처"] == selected_client]
            if not df_curr_client.empty:
                st.markdown("**📦 품목별 출고 내역**")
                client_item_summary = (
                    df_curr_client.groupby("품목명")[["출고량", "매출액"]]
                    .sum()
                    .reset_index()
                )
                client_item_summary["매출액(만원)"] = (
                    client_item_summary["매출액"] / 10000
                )
                st.dataframe(
                    client_item_summary[
                        ["품목명", "출고량", "매출액(만원)"]
                    ].style.format(
                        {"출고량": "{:,.0f}", "매출액(만원)": "{:,.1f}"}
                    ),
                    hide_index=True,
                    use_container_width=True,
                )

    # ------------------------------------
    # TAB 3: 품목 및 단가 분석
    # ------------------------------------
    with tab3:
        st.markdown(
            '<div class="sub-header">📦 품목별 매출 · 출고량 · 적용단가 종합 분석</div>',
            unsafe_allow_html=True,
        )
        p_tab1, p_tab2, p_tab3 = st.tabs(
            [
                "💰 품목별 매출액 (만 원)",
                "📦 품목별 출고량 (kg/개)",
                "🏷️ 품목별 적용단가 (원)",
            ]
        )

        with p_tab1:
            if not sales_p.empty:
                st.dataframe(
                    style_table_with_totals(sales_p / 10000, cmap="Oranges"),
                    use_container_width=True,
                    height=300,
                )
            else:
                st.info("매출액 데이터가 없습니다.")

        with p_tab2:
            if not qty_p.empty:
                st.dataframe(
                    style_table_with_totals(qty_p, cmap="Purples"),
                    use_container_width=True,
                    height=300,
                )
            else:
                st.info("출고량 데이터가 없습니다.")

        with p_tab3:
            if not unit_price_p.empty:
                st.dataframe(
                    unit_price_p.style.format("{:,.0f}"),
                    use_container_width=True,
                    height=300,
                )
            else:
                st.info("단가 데이터가 없습니다.")

    # ------------------------------------
    # TAB 4: 담당자 & 상세내역
    # ------------------------------------
    with tab4:
        st.markdown(
            '<div class="sub-header">👤 담당자별 영업 실적 (만 원)</div>',
            unsafe_allow_html=True,
        )
        if not staff_pivot.empty:
            st.dataframe(
                style_table_with_totals(staff_pivot, cmap="YlGnBu"),
                use_container_width=True,
                height=250,
            )
        else:
            st.info("담당자별 실적 데이터가 없습니다.")

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(
            '<div class="sub-header">📋 상세 거래 내역 데이터</div>',
            unsafe_allow_html=True,
        )
        if not df_detail.empty:
            df_detail_disp = df_detail.copy()
            df_detail_disp["매출일"] = df_detail_disp["매출일_dt"].dt.strftime(
                "%Y-%m-%d"
            )
            df_detail_disp = df_detail_disp[
                [
                    "매출일",
                    "담당자",
                    "거래처",
                    "품목명",
                    "출고량",
                    "단가",
                    "매출액",
                ]
            ]
            st.dataframe(
                df_detail_disp.style.format(
                    {"출고량": "{:,.0f}", "단가": "{:,.0f}", "매출액": "{:,.0f}"}
                ),
                hide_index=True,
                use_container_width=True,
                height=400,
            )
        else:
            st.info("상세 거래 내역이 없습니다.")
else:
    st.info("👈 왼쪽 사이드바에서 매출 데이터 CSV 파일을 업로드해 주세요.")
