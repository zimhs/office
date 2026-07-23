import io
import re
import subprocess
import urllib.parse
import numpy as np
import pandas as pd
import streamlit as st

# 페이지 및 Styler 가동 한도 설정
pd.set_option("styler.render.max_elements", 2000000)
st.set_page_config(page_title="통합 영업 분석 대시보드", layout="wide")


# ==========================================
# 1. 스타일 및 디자인 Injection (검색바 + 탭 완전 상단 고정 CSS)
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
            /* 🚫 multiselect 1개 제한 경고 문구 및 안내 숨기기 */
            div[data-baseweb="select"] + div:has(span) {
                display: none !important;
            }
            div[data-testid="stMultiSelect"] [data-testid="stWidgetInstructions"] {
                display: none !important;
            }
            small[data-testid="stCaptionContainer"] {
                display: none !important;
            }
            
            html, body, .stApp {
                background-color: #F8FAFC !important;
                color: #1E293B !important;
                font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
                -webkit-tap-highlight-color: transparent;
            }

            /* ⚠️ Sticky 핵심: 부모 컨테이너들의 overflow를 visible로 변경하여 sticky 작동 유도 */
            .main, .main .block-container, [data-testid="stMainBlockContainer"], [data-testid="stVerticalBlock"] {
                overflow: visible !important;
            }
            
            /* 📌 메인 영역 상단 여백 및 st.title (대제목) 슬림화 */
            .main .block-container {
                padding-top: 1.2rem !important;
                padding-bottom: 2rem !important;
            }
            h1 {
                font-size: 1.25rem !important;
                font-weight: 700 !important;
                margin-top: 0px !important;
                margin-bottom: 0.5rem !important;
                padding: 0px !important;
                color: #0F172A !important;
            }
            
            [data-testid="stSidebar"] {
                background-color: #F1F5F9 !important;
                border-right: 1px solid #E2E8F0;
            }
            [data-testid="stSidebar"] .block-container { padding-top: 1rem; }
            [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 {
                font-size: 1rem !important;
            }
            [data-testid="stSidebar"] p, [data-testid="stSidebar"] span, 
            [data-testid="stSidebar"] label, [data-testid="stSidebar"] .stMarkdown {
                color: #334155 !important;
            }

            /* 📌 1단계: 상단 고정 필터 컨테이너 (Sticky Filter Bar) */
            div[data-testid="stVerticalBlock"] > div:has(div.sticky-filter-marker) {
                position: -webkit-sticky !important;
                position: sticky !important;
                top: 2.5rem !important;
                z-index: 999 !important;
                background-color: #FFFFFF !important;
                padding: 6px 12px !important;
                border-radius: 8px;
                border: 1px solid #E2E8F0;
                box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05);
                margin-bottom: 8px;
            }
            
            /* 입력창 내부 라벨 및 여백 축소 */
            div[data-testid="stVerticalBlock"] > div:has(div.sticky-filter-marker) div[data-testid="stMarkdownContainer"] p,
            div[data-testid="stVerticalBlock"] > div:has(div.sticky-filter-marker) label {
                font-size: 12px !important;
                font-weight: 600 !important;
                margin-bottom: 2px !important;
            }
            
            div[data-testid="stVerticalBlock"] > div:has(div.sticky-filter-marker) [data-testid="stColumn"] {
                padding: 0 4px !important;
            }
            
            /* 📌 2단계: 탭 메뉴 전체 상단 고정 (Sticky Tabs Container & List) */
            div[data-testid="stTabs"] {
                position: relative !important;
                overflow: visible !important;
            }

            div[data-testid="stTabs"] > div[data-baseweb="tab-list"] {
                position: -webkit-sticky !important;
                position: sticky !important;
                top: 6.8rem !important; /* 검색 필터 바 밑에 부착 */
                z-index: 998 !important;
                background-color: #F8FAFC !important;
                padding-top: 6px;
                padding-bottom: 4px;
                gap: 4px;
                border-bottom: 2px solid #CBD5E1;
                overflow-x: auto;
                -webkit-overflow-scrolling: touch;
            }
            
            .stTabs [data-baseweb="tab"] {
                height: 36px;
                white-space: pre-wrap;
                background-color: #E2E8F0;
                border-radius: 6px 6px 0px 0px;
                color: #475569;
                font-size: 13px;
                font-weight: 600;
                padding: 0px 14px;
            }
            
            .stTabs [aria-selected="true"] {
                background-color: #FFFFFF !important;
                color: #2563EB !important;
                border-bottom: 2px solid #2563EB !important;
                border-top: 1px solid #CBD5E1;
                border-left: 1px solid #CBD5E1;
                border-right: 1px solid #CBD5E1;
            }
            
            /* KPI 메트릭 카드 */
            .metric-box {
                background: #FFFFFF;
                padding: 10px 14px;
                border-radius: 8px;
                border: 1px solid #E2E8F0;
                box-shadow: 0 2px 4px rgba(0, 0, 0, 0.03);
                margin-bottom: 8px;
            }
            .metric-label {
                color: #64748B;
                font-size: 11px;
                font-weight: 500;
                margin-bottom: 2px;
            }
            .metric-value {
                color: #0F172A;
                font-size: 18px;
                font-weight: 700;
            }
            
            /* 서브 헤더 */
            .sub-header {
                color: #2563EB;
                font-size: 14px;
                font-weight: 700;
                margin-top: 14px;
                margin-bottom: 8px;
                border-left: 3px solid #2563EB;
                padding-left: 6px;
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
    if "품목명" not in df.columns or df.empty:
        return df

    p_str = df["품목명"].astype(str)
    p_upper = p_str.str.upper().str.replace(" ", "")

    is_bulk = p_upper.str.contains("BULK", na=False) | p_str.str.contains(
        "벌크", na=False
    )
    is_ar = is_bulk & (
        p_upper.str.contains("AR", na=False)
        | p_str.str.contains("아르곤|아르", na=False)
    )
    is_co2 = is_bulk & (
        p_upper.str.contains("CO2", na=False)
        | p_str.str.contains("탄산", na=False)
    )
    is_o2 = (
        is_bulk
        & ~is_co2
        & (
            p_upper.str.contains("O2", na=False)
            | p_str.str.contains("산소", na=False)
        )
    )
    is_n2 = is_bulk & (
        p_upper.str.contains("N2", na=False)
        | p_str.str.contains("질소", na=False)
    )

    is_n2_liter = is_n2 & (
        p_upper.str.contains("L|LITER", na=False)
        | p_str.str.contains("리터", na=False)
    )

    if "출고량" in df.columns:
        df.loc[is_n2_liter, "출고량"] = df.loc[is_n2_liter, "출고량"] * 0.808

    df.loc[is_ar, "품목명"] = "AR (kg, Bulk)"
    df.loc[is_co2, "품목명"] = "CO2 (kg, Bulk)"
    df.loc[is_o2, "품목명"] = "O2 (kg, Bulk)"
    df.loc[is_n2, "품목명"] = "N2 (kg, Bulk)"

    return df


def open_macos_note(client_name):
    script = f"""
    tell application "Notes"
        activate
        try
            set targetNote to first note of folder "거래처" whose name contains "{client_name}"
            show targetNote
        on error
            tell folder "거래처"
                make new note with properties {{name:"{client_name}", body:"--- {client_name} 영업 및 특이사항 메모 ---\\n\\n"}}
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
                df.to_excel(writer, sheet_name=sheet_name, index=use_index)
    return output.getvalue()


# ==========================================
# 3. 데이터 로딩 & 메모리 캐싱 (@st.cache_data)
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


@st.cache_data(show_spinner="데이터를 최초 파싱 및 캐싱 중입니다...")
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
# 4. 메인 실행 흐름 (UI 구성)
# ==========================================
inject_custom_css()

# 슬림한 메인 대제목
st.title("📊 통합 영업 분석 대시보드")

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

target_items = [
    "CO2 (kg, Bulk)",
    "N2 (kg, Bulk)",
    "O2 (kg, Bulk)",
    "AR (kg, Bulk)",
]

if not full_df.empty:
    # 📌 1단계: 화면 상단 스크롤 고정 필터 영역 (Sticky Bar 1)
    with st.container():
        st.markdown('<div class="sticky-filter-marker"></div>', unsafe_allow_html=True)
        
        fc1, fc2, fc3, fc4, fc5 = st.columns([0.75, 0.75, 1.0, 1.5, 1.0])

        start_date = fc1.text_input("📅 시작", "200101")
        end_date = fc2.text_input("📅 종료", "261231")

        start_dt = pd.to_datetime(start_date, format="%y%m%d", errors="coerce")
        end_dt = pd.to_datetime(end_date, format="%y%m%d", errors="coerce")

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
            "🏢 거래처 선택",
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
        selected_item = fc5.multiselect("📦 품목 선택", available_items)

    df_f = (
        df_client_filtered[df_client_filtered["품목명"].isin(selected_item)]
        if selected_item
        else df_client_filtered.copy()
    )

    if not df_f.empty:
        df_f["연도"] = df_f["매출일_dt"].dt.year.astype(str)
        df_f["월"] = df_f["매출일_dt"].dt.strftime("%m월")
        df_f["분기"] = df_f["매출일_dt"].dt.to_period("Q").astype(str)
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

    # 1. 연도별 월 매출 (VAT 포함, 만원 단위)
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

    # 2. 거래처별 월별 매출
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
        actual_cols = [c for c in desired_order if c in client_pivot_raw.columns]
        client_pivot = client_pivot_raw.reindex(columns=actual_cols, fill_value=0)

    # 3. 품목 및 단가 분석용 데이터
    main_df = (
        df_f[df_f["품목명"].isin(target_items)].copy()
        if not df_f.empty
        else pd.DataFrame()
    )
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

        sales_expanded_cols = []
        sales_expanded_data = {}

        qty_expanded_cols = []
        qty_expanded_data = {}

        for yr in existing_years:
            for m in all_months:
                col_key = (yr, m)
                sales_expanded_data[col_key] = (
                    sales_raw_p[col_key] if col_key in sales_raw_p.columns else 0
                )
                sales_expanded_cols.append(col_key)

                qty_expanded_data[col_key] = (
                    qty_raw_p[col_key] if col_key in qty_raw_p.columns else 0
                )
                qty_expanded_cols.append(col_key)

            yr_sales_sum = sum(
                sales_raw_p[(yr, m)]
                for m in all_months
                if (yr, m) in sales_raw_p.columns
            )
            sales_expanded_data[(yr, "연간총합")] = yr_sales_sum
            sales_expanded_cols.append((yr, "연간총합"))

            yr_qty_sum = sum(
                qty_raw_p[(yr, m)]
                for m in all_months
                if (yr, m) in qty_raw_p.columns
            )
            qty_expanded_data[(yr, "연간총합")] = yr_qty_sum
            qty_expanded_cols.append((yr, "연간총합"))

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

    # 4. 담당자별 매출
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

    # 📌 2단계: 상단 고정 탭 바 (Sticky Tabs)
    tab1, tab2, tab3, tab4 = st.tabs(
        [
            "📌 영업 종합 요약",
            "🏢 거래처 분석",
            "📦 품목 및 단가 분석",
            "👤 담당자 & 상세내역",
        ]
    )

    # KPI 메트릭 정보
    total_sales = df_f["매출액"].sum() if not df_f.empty else 0
    total_qty = df_f["출고량"].sum() if not df_f.empty else 0
    label_suffix = "(선택 품목)" if selected_item else "(전체 품목)"

    # ------------------------------------
    # TAB 1: 영업 종합 요약 (수정 및 개선)
    # ------------------------------------
    with tab1:
        m1, m2 = st.columns(2)
        m1.markdown(
            f"""<div class="metric-box"><div class="metric-label">💰 총 매출 합계 {label_suffix}</div><div class="metric-value">{total_sales:,.0f} <span style="font-size: 13px; font-weight: normal; color: #64748B;">원</span></div></div>""",
            unsafe_allow_html=True,
        )
        m2.markdown(
            f"""<div class="metric-box"><div class="metric-label">📦 총 출고량 {label_suffix}</div><div class="metric-value">{total_qty:,.0f} <span style="font-size: 13px; font-weight: normal; color: #64748B;">개</span></div></div>""",
            unsafe_allow_html=True,
        )

        # 당해 연도 추출 (가장 최신 연도 기준)
        current_year = existing_years[-1] if existing_years else "2026"
        df_curr = df_f[df_f["연도"] == current_year] if not df_f.empty else pd.DataFrame()

        # ------------------------------------
        # 1. 당해 월별 합계 매출 (표 + 그래프)
        # ------------------------------------
        st.markdown(
            f'<div class="sub-header">📅 {current_year}년 당해 월별 합계 매출 (VAT 포함, 만 원 단위)</div>',
            unsafe_allow_html=True,
        )
        col_m_tbl, col_m_cht = st.columns([1, 1])

        curr_month_sales = pd.DataFrame(index=all_months, columns=["매출액(만원)"]).fillna(0.0)

        if not df_curr.empty:
            m_grouped = (
                df_curr.groupby("월")["매출액"].sum() * 1.1 / 10000
            )
            for m in all_months:
                if m in m_grouped.index:
                    curr_month_sales.loc[m, "매출액(만원)"] = m_grouped[m]

        with col_m_tbl:
            st.markdown(f"**📋 {current_year}년 월별 매출 표**")
            st.dataframe(
                curr_month_sales.style.format("{:,.0f}").background_gradient(
                    cmap="Blues", axis=0
                ),
                use_container_width=True,
                height=350,
            )

        with col_m_cht:
            st.markdown(f"**📊 {current_year}년 월별 매출 추이**")
            st.bar_chart(curr_month_sales["매출액(만원)"], use_container_width=True, height=350)

        # ------------------------------------
        # 2. 당해 분기별 매출 (표 + 그래프)
        # ------------------------------------
        st.markdown(
            f'<div class="sub-header">📊 {current_year}년 당해 분기별 매출 현황 (VAT 포함, 만 원 단위)</div>',
            unsafe_allow_html=True,
        )
        col_q_tbl, col_q_cht = st.columns([1, 1])

        q_labels = ["1분기 (1~3월)", "2분기 (4~6월)", "3분기 (7~9월)", "4분기 (10~12월)"]
        curr_q_sales = pd.DataFrame(index=q_labels, columns=["매출액(만원)"]).fillna(0.0)

        if not df_curr.empty:
            q_map = {
                "01월": "1분기 (1~3월)", "02월": "1분기 (1~3월)", "03월": "1분기 (1~3월)",
                "04월": "2분기 (4~6월)", "05월": "2분기 (4~6월)", "06월": "2분기 (4~6월)",
                "07월": "3분기 (7~9월)", "08월": "3분기 (7~9월)", "09월": "3분기 (7~9월)",
                "10월": "4분기 (10~12월)", "11월": "4분기 (10~12월)", "12월": "4분기 (10~12월)",
            }
            df_curr_copy = df_curr.copy()
            df_curr_copy["분기_구분"] = df_curr_copy["월"].map(q_map)
            q_grouped = df_curr_copy.groupby("분기_구분")["매출액"].sum() * 1.1 / 10000

            for q in q_labels:
                if q in q_grouped.index:
                    curr_q_sales.loc[q, "매출액(만원)"] = q_grouped[q]

        with col_q_tbl:
            st.markdown(f"**📋 {current_year}년 분기별 매출 표**")
            st.dataframe(
                curr_q_sales.style.format("{:,.0f}").background_gradient(
                    cmap="Blues", axis=0
                ),
                use_container_width=True,
                height=240,
            )

        with col_q_cht:
            st.markdown(f"**📊 {current_year}년 분기별 매출 비교**")
            st.bar_chart(curr_q_sales["매출액(만원)"], use_container_width=True, height=240)

        # ------------------------------------
        # 3. 전체 연도 월별 매출 비교 (종합 시야 제공)
        # ------------------------------------
        st.markdown(
            '<div class="sub-header">📈 전체 연도 동월 비교 종합 (VAT 포함, 만 원 단위)</div>',
            unsafe_allow_html=True,
        )
        col_table1, col_chart1 = st.columns([1.2, 1.8])

        with col_table1:
            st.markdown("**📋 연도별 전체 월 매출 데이터**")
            if not pivot_m.empty:
                styled_pivot_m = (
                    pivot_m.style.format("{:,.0f}").background_gradient(
                        cmap="Blues", axis=None
                    )
                )
                st.dataframe(
                    styled_pivot_m, use_container_width=True, height=300
                )
            else:
                st.info("데이터 없음")

        with col_chart1:
            st.markdown("**📊 연도별 월 매출 그래프 비교**")
            if not df_f.empty:
                chart_m = (
                    df_f.pivot_table(
                        index="월",
                        columns="연도",
                        values="매출액",
                        aggfunc="sum",
                    ).fillna(0)
                    * 1.1
                    / 10000
                )
                chart_m = chart_m.reindex(all_months, fill_value=0)
                st.bar_chart(chart_m, use_container_width=True, height=300)
            else:
                st.info("표시할 그래프 데이터가 없습니다.")

    # ------------------------------------
    # TAB 2: 거래처 분석
    # ------------------------------------
    with tab2:
        st.markdown(
            '<div class="sub-header">🏢 거래처별 월별 비교 현황 (만 원 단위)</div>',
            unsafe_allow_html=True,
        )
        if not client_pivot.empty:
            year_groups = {}
            for col in client_pivot.columns:
                yr_match = col.split("년")[0] + "년" if "년" in col else "기타"
                m_match = col.split("년")[-1].strip() if "년" in col else col
                year_groups.setdefault(yr_match, []).append((col, m_match))

            html_table = """
            <div style="max-height: 520px; overflow-y: auto; overflow-x: auto; -webkit-overflow-scrolling: touch; border: 1px solid #CBD5E1; border-radius: 8px; margin-bottom: 20px; background-color: #FFFFFF;">
                <table style="width: 100%; border-collapse: separate; border-spacing: 0; font-size: 13px; font-family: -apple-system, sans-serif;">
                    <thead>
                        <!-- 1단 연도 헤더 -->
                        <tr style="background-color: #E2E8F0; position: sticky; top: 0; z-index: 15;">
                            <th rowspan="2" style="padding: 12px; text-align: center; color: #0F172A !important; font-weight: 700; position: sticky; left: 0; background-color: #E2E8F0; z-index: 25; min-width: 180px; border-bottom: 2px solid #94A3B8; border-right: 2px solid #94A3B8;">거래처명</th>
            """
            
            for yr, cols in year_groups.items():
                span = len(cols)
                html_table += f'<th colspan="{span}" style="padding: 8px; text-align: center; color: #1E3A8A !important; font-weight: 700; background-color: #DBEAFE; border-bottom: 1px solid #94A3B8; border-right: 2px solid #2563EB;">20{yr if len(yr)==3 else yr}</th>'

            html_table += '</tr><tr style="background-color: #F8FAFC; position: sticky; top: 35px; z-index: 10;">'

            for yr, cols in year_groups.items():
                for idx, (col_full, month_str) in enumerate(cols):
                    is_last = (idx == len(cols) - 1)
                    border_right = "border-right: 2px solid #2563EB;" if is_last else "border-right: 1px solid #E2E8F0;"
                    html_table += f'<th style="padding: 8px 10px; text-align: right; color: #334155 !important; font-weight: 600; min-width: 80px; white-space: nowrap; border-bottom: 2px solid #94A3B8; {border_right}">{month_str}</th>'

            html_table += "</tr></thead><tbody>"

            for row_idx, (client_name, row) in enumerate(client_pivot.iterrows()):
                bg_color = "#FFFFFF" if row_idx % 2 == 0 else "#F8FAFC"
                html_table += f'<tr style="background-color: {bg_color};">'
                html_table += f'<td style="padding: 8px 12px; text-align: left; color: #0F172A !important; font-weight: 600; position: sticky; left: 0; background-color: {bg_color}; border-right: 2px solid #94A3B8; border-bottom: 1px solid #E2E8F0; white-space: nowrap;">{client_name}</td>'
                
                col_pointer = 0
                for yr, cols in year_groups.items():
                    for idx, (col_full, m_str) in enumerate(cols):
                        val = row.iloc[col_pointer]
                        col_pointer += 1
                        val_str = f"{val:,.0f}" if val != 0 else "-"
                        text_color = "#0F172A" if val != 0 else "#CBD5E1"
                        font_weight = "600" if val != 0 else "400"
                        
                        is_last = (idx == len(cols) - 1)
                        border_right = "border-right: 2px solid #2563EB;" if is_last else "border-right: 1px solid #F1F5F9;"
                        
                        html_table += f'<td style="padding: 8px 10px; text-align: right; color: {text_color} !important; font-weight: {font_weight}; border-bottom: 1px solid #E2E8F0; {border_right}">{val_str}</td>'
                
                html_table += "</tr>"
                
            html_table += "</tbody></table></div>"
            
            st.markdown(html_table, unsafe_allow_html=True)
            
        else:
            st.info("거래처별 데이터가 없습니다.")

        if selected_client and selected_client != "전체 거래처":
            st.markdown(
                f'<div class="sub-header">📝 거래처 메모 연동 ({selected_client})</div>',
                unsafe_allow_html=True,
            )
            col_memo1, col_memo2 = st.columns(2)

            with col_memo1:
                if st.button(
                    f"📝 '{selected_client}' 메모 열기/생성",
                    use_container_width=True,
                ):
                    success = open_macos_note(selected_client)
                    if success:
                        st.success(
                            f"메모 앱에서 '{selected_client}' 노트를 열거나 생성했습니다."
                        )
                    else:
                        st.info(
                            f"📱 아이패드에서는 [iOS 메모 앱]을 직접 열거나 아래 딥링크 버튼을 활용해주세요."
                        )

                notes_url = f"mobilenotes://"
                st.link_button("📱 아이패드 메모 앱 바로가기", notes_url, use_container_width=True)

            with col_memo2:
                st.info(
                    "💡 기기 제약 없는 웹 기반 노트 서비스 또는 대시보드 내 직접 입력 기능을 활용하실 수 있습니다."
                )

        if addr_dict and not df_f.empty:
            st.markdown(
                '<div class="sub-header">📍 선택 거래처 지도 연동</div>',
                unsafe_allow_html=True,
            )
            matched_clients = df_f["거래처"].unique()
            map_cols = st.columns(3)
            for idx, c in enumerate(matched_clients[:6]):
                addr = addr_dict.get(c)
                if addr and str(addr).strip() not in ["nan", "None", ""]:
                    with map_cols[idx % 3]:
                        st.markdown(f"**{c}**")
                        st.caption(str(addr))
                        safe_addr = urllib.parse.quote(str(addr).strip())
                        st.link_button(
                            "🗺️ Kakao 지도 보기",
                            f"https://map.kakao.com/?q={safe_addr}",
                            use_container_width=True,
                        )

    # ------------------------------------
    # TAB 3: 품목 및 단가 분석
    # ------------------------------------
    with tab3:
        st.markdown(
            '<div class="sub-header">📅 주요품목 분기별 매출 비교 (만 원 단위)</div>',
            unsafe_allow_html=True,
        )
        col_t3, col_c3 = st.columns([1.1, 1.9])

        with col_t3:
            if not main_df.empty:
                q_pivot = (
                    main_df.pivot_table(
                        index="분기",
                        columns="품목명",
                        values="매출액",
                        aggfunc="sum",
                    ).fillna(0)
                    / 10000
                )
                q_pivot = q_pivot.reindex(columns=target_items, fill_value=0)
                st.dataframe(
                    q_pivot.style.format("{:,.0f}"), use_container_width=True
                )
            else:
                st.info("주요품목 데이터가 없습니다.")

        with col_c3:
            if not main_df.empty:
                q_chart = (
                    main_df.pivot_table(
                        index="분기",
                        columns="품목명",
                        values="매출액",
                        aggfunc="sum",
                    ).fillna(0)
                    / 10000
                )
                q_chart = q_chart.reindex(columns=target_items, fill_value=0)
                st.bar_chart(q_chart, use_container_width=True)
            else:
                st.info("그래프 데이터가 없습니다.")

        st.markdown(
            '<div class="sub-header">📦 품목별 월별 매출액 / 출고량 / 실질 적용 단가 분석 (연도별 연간 합계 포함)</div>',
            unsafe_allow_html=True,
        )
        if not df_f.empty and not sales_p.empty:

            def highlight_annual_total(df):
                styles = pd.DataFrame("", index=df.index, columns=df.columns)
                for col in df.columns:
                    if (
                        isinstance(col, tuple)
                        and len(col) > 1
                        and col[1] == "연간총합"
                    ):
                        styles[col] = (
                            "background-color: #FEF08A; font-weight: bold;"
                            " color: #854D0E;"
                        )
                return styles

            @st.dialog("🎯 [매출액] 품목별 연도별 비교 (원형 그래프)")
            def show_sales_item_pie_chart(item_name):
                st.markdown(f"**📦 선택 품목: `{item_name}`**")
                item_df = df_f[df_f["품목명"] == item_name]
                if not item_df.empty:
                    yr_sales = (
                        item_df.groupby("연도")["매출액"].sum() / 10000
                    )

                    import altair as alt

                    chart_data = yr_sales.reset_index()
                    chart_data.columns = ["연도", "매출액"]

                    pie_chart = (
                        alt.Chart(chart_data)
                        .mark_arc(innerRadius=50)
                        .encode(
                            theta=alt.Theta(
                                field="매출액", type="quantitative"
                            ),
                            color=alt.Color(
                                field="연도",
                                type="nominal",
                                legend=alt.Legend(title="연도"),
                            ),
                            tooltip=[
                                "연도",
                                alt.Tooltip(
                                    "매출액:Q",
                                    format=",.1f",
                                    title="매출액(만원)",
                                ),
                            ],
                        )
                        .properties(width=400, height=300)
                    )

                    st.altair_chart(pie_chart, use_container_width=True)

                    st.markdown("---")
                    st.markdown("**📋 연도별 매출 상세 내역 (만 원)**")
                    st.dataframe(
                        yr_sales.to_frame(name="매출액(만원)").style.format(
                            "{:,.1f}"
                        ),
                        use_container_width=True,
                    )
                else:
                    st.warning("선택한 품목의 상세 데이터가 없습니다.")

            st.markdown(
                "**📋 [매출액] 품목별 월별 합계 및 연도별 총합 (만 원)**"
            )
            items_list = list(sales_p.index)
            col_sel1, col_btn1 = st.columns([2, 1])
            with col_sel1:
                selected_sales_popup_item = st.selectbox(
                    "🔍 연도별 매출 비교를 확인할 품목 선택",
                    items_list,
                    key="sales_popup_item_select",
                )
            with col_btn1:
                st.write("")
                st.write("")
                if st.button(
                    "📊 매출액 팝업 그래프 보기",
                    use_container_width=True,
                    key="btn_sales_popup",
                ):
                    show_sales_item_pie_chart(selected_sales_popup_item)

            styled_sales_p = (
                (sales_p / 10000)
                .style.apply(highlight_annual_total, axis=None)
                .format("{:,.0f}")
            )
            st.dataframe(styled_sales_p, use_container_width=True)

            @st.dialog("🎯 [출고량] 품목별 연도별 비교 (원형 그래프)")
            def show_qty_item_pie_chart(item_name):
                st.markdown(f"**📦 선택 품목: `{item_name}`**")
                item_df = df_f[df_f["품목명"] == item_name]
                if not item_df.empty:
                    yr_qty = item_df.groupby("연도")["출고량"].sum()

                    import altair as alt

                    chart_data = yr_qty.reset_index()
                    chart_data.columns = ["연도", "출고량"]

                    pie_chart = (
                        alt.Chart(chart_data)
                        .mark_arc(innerRadius=50)
                        .encode(
                            theta=alt.Theta(
                                field="출고량", type="quantitative"
                            ),
                            color=alt.Color(
                                field="연도",
                                type="nominal",
                                legend=alt.Legend(title="연도"),
                            ),
                            tooltip=[
                                "연도",
                                alt.Tooltip(
                                    "출고량:Q",
                                    format=",.0f",
                                    title="출고량",
                                ),
                            ],
                        )
                        .properties(width=400, height=300)
                    )

                    st.altair_chart(pie_chart, use_container_width=True)

                    st.markdown("---")
                    st.markdown("**📋 연도별 출고량 상세 내역**")
                    st.dataframe(
                        yr_qty.to_frame(name="출고량").style.format(
                            "{:,.0f}"
                        ),
                        use_container_width=True,
                    )
                else:
                    st.warning("선택한 품목의 상세 데이터가 없습니다.")

            st.markdown(
                "**📋 [출고량] 품목별 월별 합계 및 연도별 총합**"
            )
            col_sel2, col_btn2 = st.columns([2, 1])
            with col_sel2:
                selected_qty_popup_item = st.selectbox(
                    "🔍 연도별 출고량 비교를 확인할 품목 선택",
                    items_list,
                    key="qty_popup_item_select",
                )
            with col_btn2:
                st.write("")
                st.write("")
                if st.button(
                    "📊 출고량 팝업 그래프 보기",
                    use_container_width=True,
                    key="btn_qty_popup",
                ):
                    show_qty_item_pie_chart(selected_qty_popup_item)

            styled_qty_p = qty_p.style.apply(
                highlight_annual_total, axis=None
            ).format("{:,.0f}")
            st.dataframe(styled_qty_p, use_container_width=True)

            st.markdown(
                "**📋 [적용 단가] 품목별 월별 실질 적용 단가 (원)**"
            )
            if not unit_price_p.empty:
                st.dataframe(
                    unit_price_p.style.format("{:,.0f}"),
                    use_container_width=True,
                )
            else:
                st.info("단가 데이터가 없습니다.")
        else:
            st.info("품목 및 단가 분석 데이터가 없습니다.")

    # ------------------------------------
    # TAB 4: 담당자 & 상세내역
    # ------------------------------------
    with tab4:
        st.markdown(
            '<div class="sub-header">👤 담당자별 월별 매출 (만 원 단위)</div>',
            unsafe_allow_html=True,
        )
        if not staff_pivot.empty:
            st.dataframe(
                staff_pivot.style.format("{:,.0f}"),
                use_container_width=True,
            )
        else:
            st.info("담당자별 데이터가 없습니다.")

        st.markdown(
            '<div class="sub-header">📄 상세 거래 내역</div>',
            unsafe_allow_html=True,
        )
         if not df_detail.empty:
            styled_detail = df_detail.copy()
            styled_detail["매출일_dt"] = styled_detail[
                "매출일_dt"
            ].dt.strftime("%Y-%m-%d")
            st.dataframe(
                styled_detail.style.format(
                    {"출고량": "{:,.0f}", "단가": "{:,.0f}", "매출액": "{:,.0f}"}
                ),
                use_container_width=True,
            )
        else:
            st.info("상세 거래 내역이 없습니다.")
else:
    st.info("👈 왼쪽 사이드바에서 데이터 파일(CSV)을 업로드해주세요.")
