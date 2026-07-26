import io
import re
import sys
import subprocess
import numpy as np
import pandas as pd
import streamlit as st

# 페이지 및 Styler 가동 한도 설정
pd.set_option("styler.render.max_elements", 2000000)
st.set_page_config(page_title="통합 영업 분석 대시보드", layout="wide")


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
            div[data-baseweb="select"] + div:has(span) { display: none !important; }
            div[data-testid="stMultiSelect"] [data-testid="stWidgetInstructions"] { display: none !important; }
            small[data-testid="stCaptionContainer"] { display: none !important; }
            
            html, body, .stApp {
                background-color: #F8FAFC !important;
                color: #1E293B !important;
                font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
                -webkit-tap-highlight-color: transparent;
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

            div[data-testid="column"] { align-self: flex-start; }
            div[data-testid="stTextInput"], div[data-testid="stMultiSelect"] { min-height: 80px; }
            div[data-testid="stTextInput"] label, div[data-testid="stMultiSelect"] label {
                font-size: 13px !important;
                font-weight: 600 !important;
                white-space: nowrap !important;
                overflow: hidden;
                text-overflow: ellipsis;
                margin-bottom: 4px !important;
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
                color: #2563EB;
                font-size: 17px;
                font-weight: 700;
                margin-top: 25px;
                margin-bottom: 12px;
                border-left: 4px solid #2563EB;
                padding-left: 10px;
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


# ==========================================
# 3. 데이터 로딩 & 메모리 캐싱
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

            df["연도"] = df["매출일_dt"].dt.year.astype(str)
            df["월"] = df["매출일_dt"].dt.strftime("%m월")

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

target_items = [
    "CO2 (kg, Bulk)",
    "N2 (kg, Bulk)",
    "O2 (kg, Bulk)",
    "AR (kg, Bulk)",
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
        if "연도" not in df_f.columns:
            df_f["연도"] = df_f["매출일_dt"].dt.year.astype(str)
        if "월" not in df_f.columns:
            df_f["월"] = df_f["매출일_dt"].dt.strftime("%m월")
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

    # 1. 연도별 월 매출 (전체 데이터 기반 요약용)
    pivot_m = pd.DataFrame()
    if not full_df.empty:
        full_df_temp = full_df.copy()
        if "연도" not in full_df_temp.columns:
            full_df_temp["연도"] = full_df_temp["매출일_dt"].dt.year.astype(str)
        if "월" not in full_df_temp.columns:
            full_df_temp["월"] = full_df_temp["매출일_dt"].dt.strftime("%m월")

        pivot_m = (
            full_df_temp.pivot_table(
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
        actual_cols = [
            c for c in desired_order if c in client_pivot_raw.columns
        ]
        client_pivot = client_pivot_raw.reindex(
            columns=actual_cols, fill_value=0
        )

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

        # 🟢 단가 데이터프레임 구성 (멀티인덱스 컬럼을 평탄화하여 0 또는 빈값에 대한 스타일 오류 원천 차단)
        raw_up = df_f[df_f["단가"] > 0].pivot_table(
            index="품목명",
            columns=["연도", "월"],
            values="단가",
            aggfunc="first",
        )
        if not raw_up.empty:
            if isinstance(raw_up.columns, pd.MultiIndex):
                raw_up.columns = [
                    f"{y} {m}" for y, m in raw_up.columns
                ]
            unit_price_p = raw_up.fillna(0)
        else:
            unit_price_p = pd.DataFrame()

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
            "품목별_매출액(만원)": (sales_p * 1.1 / 10000, True),
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
        f"""<div class="metric-box"><div class="metric-label">📦 총 출고량 {label_suffix}</div><div class="metric-value">{total_qty:,.0f} <span style="font-size: 15px; font-weight: normal; color: #64748B;">수량</span></div></div>""",
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
        all_df_clean = full_df.copy()
        if not all_df_clean.empty and "매출일_dt" in all_df_clean.columns:
            all_df_clean["연도"] = all_df_clean["매출일_dt"].dt.year.astype(str)
            all_df_clean["월숫자"] = all_df_clean["매출일_dt"].dt.month
            max_yr = all_df_clean["연도"].max()
            df_max_yr = all_df_clean[all_df_clean["연도"] == max_yr]
            max_m = (
                df_max_yr["월숫자"].max() if not df_max_yr.empty else 1
            )
            prev_m = max_m - 1 if max_m > 1 else 12
            prev_m_yr = max_yr if max_m > 1 else str(int(max_yr) - 1)
            prev_yr = str(int(max_yr) - 1)

            cur_sales = (
                df_max_yr[df_max_yr["월숫자"] == max_m]["매출액"].sum() * 1.1
            )
            prev_month_sales = (
                all_df_clean[
                    (all_df_clean["연도"] == prev_m_yr)
                    & (all_df_clean["월숫자"] == prev_m)
                ]["매출액"].sum()
                * 1.1
            )
            prev_year_sales = (
                all_df_clean[
                    (all_df_clean["연도"] == prev_yr)
                    & (all_df_clean["월숫자"] == max_m)
                ]["매출액"].sum()
                * 1.1
            )

            mom_diff = (
                ((cur_sales - prev_month_sales) / prev_month_sales * 100)
                if prev_month_sales > 0
                else 0
            )
            yoy_diff = (
                ((cur_sales - prev_year_sales) / prev_year_sales * 100)
                if prev_year_sales > 0
                else 0
            )

            cur_qty = df_max_yr[df_max_yr["월숫자"] == max_m]["출고량"].sum()
            prev_month_qty = all_df_clean[
                (all_df_clean["연도"] == prev_m_yr)
                & (all_df_clean["월숫자"] == prev_m)
            ]["출고량"].sum()
            qty_mom_diff = (
                ((cur_qty - prev_month_qty) / prev_month_qty * 100)
                if prev_month_qty > 0
                else 0
            )
        else:
            max_yr, max_m, cur_sales, mom_diff, yoy_diff, cur_qty, qty_mom_diff = (
                "2026",
                12,
                0,
                0,
                0,
                0,
                0,
            )

        st.markdown(
            f'<div class="sub-header">📊 최근 전체 실적 요약 지표 ({max_yr}년 {max_m}월 기준)</div>',
            unsafe_allow_html=True,
        )
        um1, um2, um3, um4 = st.columns(4)
        um1.markdown(
            f"""<div class="metric-box"><div class="metric-label">💰 당월 매출액 (VAT포함)</div><div class="metric-value">{cur_sales:,.0f} <span style="font-size: 13px; color: #64748B;">원</span></div></div>""",
            unsafe_allow_html=True,
        )
        um2.markdown(
            f"""<div class="metric-box"><div class="metric-label">📈 전월 대비 매출증감</div><div class="metric-value" style="color: {'#EF4444' if mom_diff < 0 else '#10B981'};">{mom_diff:+.1f}%</div></div>""",
            unsafe_allow_html=True,
        )
        um3.markdown(
            f"""<div class="metric-box"><div class="metric-label">📊 전년 동월 대비 매출증감</div><div class="metric-value" style="color: {'#EF4444' if yoy_diff < 0 else '#10B981'};">{yoy_diff:+.1f}%</div></div>""",
            unsafe_allow_html=True,
        )
        um4.markdown(
            f"""<div class="metric-box"><div class="metric-label">📦 당월 출고량 / 전월비</div><div class="metric-value">{cur_qty:,.0f} <span style="font-size: 13px; color: #64748B;">수량</span> <span style="font-size: 14px; color: {'#EF4444' if qty_mom_diff < 0 else '#10B981'};">({qty_mom_diff:+.1f}%)</span></div></div>""",
            unsafe_allow_html=True,
        )

        st.markdown("<br>", unsafe_allow_html=True)

        st.markdown(
            '<div class="sub-header">📈 전체 월별 매출 추이 및 연도별 비교 (만 원 단위)</div>',
            unsafe_allow_html=True,
        )
        col_table1, col_chart1 = st.columns([1, 1])

        with col_table1:
            st.markdown(
                "**📋 연도별 전체 월 매출 데이터 (VAT 포함, 만 원)**"
            )
            if not pivot_m.empty:
                styled_pivot_m = pivot_m.style.format(
                    "{:,.0f}"
                ).background_gradient(cmap="Blues", axis=None)
                st.dataframe(
                    styled_pivot_m, use_container_width=True, height=360
                )
            else:
                st.info("데이터 없음")

        with col_chart1:
            st.markdown("**📊 연도 동월 비교 그래프 (VAT 포함)**")
            if not full_df.empty:
                chart_m = (
                    full_df.pivot_table(
                        index="월",
                        columns="연도",
                        values="매출액",
                        aggfunc="sum",
                    ).fillna(0)
                    * 1.1
                    / 10000
                )
                chart_m = chart_m.reindex(all_months, fill_value=0)
                st.bar_chart(chart_m, use_container_width=True, height=360)
            else:
                st.info("표시할 그래프 데이터가 없습니다.")

        st.markdown("---")

        st.markdown(
            '<div class="sub-header">🧪 주요 품목별 매출액 및 출고량 상세 분석 (연도 × 월)</div>',
            unsafe_allow_html=True,
        )

        if not full_df.empty:
            df_major = full_df[
                full_df["품목명"].isin(target_items)
            ].copy()

            if not df_major.empty:
                selected_analysis_item = st.selectbox(
                    "📦 개별 분석할 주요 품목 선택",
                    target_items,
                    key="major_item_selector",
                )
                selected_metric_type = st.radio(
                    "📊 분석 지표 선택",
                    ["매출액 (VAT 포함, 만 원)", "출고량 (천 단위)"],
                    horizontal=True,
                    key="major_metric_selector",
                )

                df_sub_item = df_major[
                    df_major["품목명"] == selected_analysis_item
                ]

                if not df_sub_item.empty:
                    if "매출액" in selected_metric_type:
                        item_pivot_table = (
                            df_sub_item.pivot_table(
                                index="연도",
                                columns="월",
                                values="매출액",
                                aggfunc="sum",
                            ).fillna(0)
                            * 1.1
                            / 10000
                        )
                        item_chart_data = (
                            df_sub_item.pivot_table(
                                index="월",
                                columns="연도",
                                values="매출액",
                                aggfunc="sum",
                            ).fillna(0)
                            * 1.1
                            / 10000
                        )
                    else:
                        item_pivot_table = (
                            df_sub_item.pivot_table(
                                index="연도",
                                columns="월",
                                values="출고량",
                                aggfunc="sum",
                            ).fillna(0)
                            / 1000
                        )
                        item_chart_data = (
                            df_sub_item.pivot_table(
                                index="월",
                                columns="연도",
                                values="출고량",
                                aggfunc="sum",
                            ).fillna(0)
                            / 1000
                        )

                    item_pivot_table = item_pivot_table.reindex(
                        columns=all_months, fill_value=0
                    )
                    item_chart_data = item_chart_data.reindex(
                        all_months, fill_value=0
                    )

                    col_itbl, col_ichart = st.columns([1, 1])
                    with col_itbl:
                        st.markdown(f"**📋 [{selected_analysis_item}] 월별 데이터표**")
                        styled_item_tbl = item_pivot_table.style.format(
                            "{:,.1f}"
                        ).background_gradient(cmap="Purples", axis=None)
                        st.dataframe(
                            styled_item_tbl,
                            use_container_width=True,
                            height=320,
                        )
                    with col_ichart:
                        st.markdown(
                            f"**📊 [{selected_analysis_item}] 연도별 월 비교 그래프**"
                        )
                        st.bar_chart(
                            item_chart_data, use_container_width=True, height=320
                        )

    # ------------------------------------
    # TAB 2: 거래처 분석
    # ------------------------------------
    with tab2:
        st.markdown(
            '<div class="sub-header">🏢 거래처별 월별 매출 현황 (만 원 단위, VAT 포함)</div>',
            unsafe_allow_html=True,
        )
        if not client_pivot.empty:
            client_display = client_pivot * 1.1
            styled_client = client_display.style.format(
                "{:,.0f}"
            ).background_gradient(cmap="Blues", axis=None)
            st.dataframe(styled_client, use_container_width=True, height=450)
        else:
            st.info("조건에 해당하는 거래처 데이터가 없습니다.")

    # ------------------------------------
    # TAB 3: 품목 및 단가 분석
    # ------------------------------------
    with tab3:
        st.markdown(
            '<div class="sub-header">🔍 선택된 거래처의 품목별 상세 데이터 및 월별 누적 비교</div>',
            unsafe_allow_html=True,
        )

        tab3_view_mode = st.radio(
            "조회 지표 선택",
            ["매출액 (만 원, VAT 포함)", "출고량"],
            horizontal=True,
            key="tab3_metric_radio",
        )

        if not df_client_filtered.empty:
            all_client_items = sorted(df_client_filtered["품목명"].unique())

            if all_client_items:
                selected_analysis_item = st.selectbox(
                    "📦 상세 내역 및 그래프를 조회할 품목 선택",
                    all_client_items,
                    key="tab3_major_item_selector",
                )

                df_sub_item = df_client_filtered[
                    df_client_filtered["품목명"] == selected_analysis_item
                ]

                if not df_sub_item.empty:
                    if "매출액" in tab3_view_mode:
                        item_pivot_table = (
                            df_sub_item.pivot_table(
                                index="연도",
                                columns="월",
                                values="매출액",
                                aggfunc="sum",
                            ).fillna(0)
                            * 1.1
                            / 10000
                        )
                        item_chart_data = (
                            df_sub_item.pivot_table(
                                index="월",
                                columns="연도",
                                values="매출액",
                                aggfunc="sum",
                            ).fillna(0)
                            * 1.1
                            / 10000
                        )
                        unit_label = "매출액 (만 원)"
                    else:
                        item_pivot_table = df_sub_item.pivot_table(
                            index="연도",
                            columns="월",
                            values="출고량",
                            aggfunc="sum",
                        ).fillna(0)
                        item_chart_data = df_sub_item.pivot_table(
                            index="월",
                            columns="연도",
                            values="출고량",
                            aggfunc="sum",
                        ).fillna(0)
                        unit_label = "출고량"

                    item_pivot_table = item_pivot_table.reindex(
                        columns=all_months, fill_value=0
                    )
                    item_chart_data = item_chart_data.reindex(
                        all_months, fill_value=0
                    )

                    col_itbl, col_ichart = st.columns([1, 1])
                    with col_itbl:
                        st.markdown(
                            f"**📋 [{selected_analysis_item}] 연도별·월별 {unit_label} 표**"
                        )
                        styled_item_tbl = item_pivot_table.style.format(
                            "{:,.0f}"
                        ).background_gradient(cmap="Blues", axis=None)
                        st.dataframe(
                            styled_item_tbl,
                            use_container_width=True,
                            height=360,
                        )
                    with col_ichart:
                        st.markdown(
                            f"**📊 [{selected_analysis_item}] 월별 연도 비교 그래프 (누적 바)**"
                        )
                        st.bar_chart(
                            item_chart_data, use_container_width=True, height=360
                        )
            else:
                st.info("선택된 거래처에 품목 데이터가 없습니다.")
        else:
            st.info("좌측 상단 필터에서 거래처를 선택해주세요.")

        st.markdown("---")

        st.markdown(
            '<div class="sub-header">🏷️ 품목별 적용 단가 추이</div>',
            unsafe_allow_html=True,
        )
        if not unit_price_p.empty:
            # 🟢 단가 데이터프레임의 0원/결측치 셀에 스타일이 잘못 적용되어 검은색 배경으로 깨지는 현상을 방지하는 커스텀 스타일 적용
            def color_non_zero_price(val):
                if pd.isna(val) or val == 0:
                    return "color: #94A3B8; background-color: #F8FAFC;"
                return "color: #0F172A; background-color: #ECFDF5; font-weight: 600;"

            styled_up = unit_price_p.style.format(
                lambda x: f"{x:,.0f}" if pd.notna(x) and x > 0 else "-"
            ).map(color_non_zero_price)

            st.dataframe(
                styled_up,
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
            '<div class="sub-header">👤 담당자별 월 매출 현황 (만 원, VAT 포함)</div>',
            unsafe_allow_html=True,
        )
        if not staff_pivot.empty:
            staff_display = staff_pivot * 1.1
            st.dataframe(
                staff_display.style.format("{:,.0f}").background_gradient(
                    cmap="PuBu", axis=None
                ),
                use_container_width=True,
                height=300,
            )
        else:
            st.info("담당자별 데이터가 없습니다.")

        st.markdown(
            '<div class="sub-header">📄 조건별 상세 거래 내역 원본</div>',
            unsafe_allow_html=True,
        )
        if not df_detail.empty:
            detail_view = df_detail.copy()
            detail_view["매출일_dt"] = detail_view["매출일_dt"].dt.strftime(
                "%Y-%m-%d"
            )
            st.dataframe(detail_view, use_container_width=True, height=400)
        else:
            st.info("상세 거래 내역이 없습니다.")
else:
    st.info(
        "👈 좌측 사이드바에서 **매출 데이터(CSV)** 파일을 업로드하여 대시보드를 활성화하세요."
    )
