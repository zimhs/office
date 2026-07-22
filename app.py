import io
import re
import urllib.parse
import altair as alt
import numpy as np
import pandas as pd
import streamlit as st

# 페이지 설정
pd.set_option("styler.render.max_elements", 2000000)
st.set_page_config(page_title="통합 영업 분석 대시보드", layout="wide")


# ==========================================
# 1. 커스텀 CSS (다크모드/라이트모드 시인성 보정)
# ==========================================
def inject_custom_css():
    st.markdown(
        """
        <script>
            document.documentElement.lang = 'ko';
            document.documentElement.classList.add('notranslate');
        </script>
        <meta name="google" content="notranslate" />
        <style>
            /* 전체 배경 및 기본 글자색 강제 고정 */
            html, body, .stApp {
                background-color: #FFFFFF !important;
                color: #0F172A !important;
                font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
            }
            [data-testid="stSidebar"] {
                background-color: #F1F5F9 !important;
                border-right: 1px solid #E2E8F0;
            }
            
            /* KPI 메트릭 카드 */
            .metric-box {
                background: #F8FAFC;
                padding: 16px 20px;
                border-radius: 10px;
                border: 1px solid #E2E8F0;
                box-shadow: 0 2px 4px rgba(0,0,0,0.02);
            }
            .metric-label {
                color: #475569;
                font-size: 13px;
                font-weight: 600;
                margin-bottom: 6px;
            }
            .metric-value {
                color: #0F172A;
                font-size: 22px;
                font-weight: 700;
            }
            
            /* 서브 헤더 */
            .sub-header {
                color: #2563EB;
                font-size: 17px;
                font-weight: 700;
                margin-top: 20px;
                margin-bottom: 12px;
                border-left: 4px solid #2563EB;
                padding-left: 10px;
            }
        </style>
    """,
        unsafe_allow_html=True,
    )


# ==========================================
# 2. 유틸리티 함수 (날짜 파싱, 품목 정규화)
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
                df.to_excel(writer, sheet_name=sheet_name, index=use_index)
    return output.getvalue()


# ==========================================
# 3. 데이터 로딩 & 캐싱
# ==========================================
@st.cache_data(show_spinner="주소록 읽는 중...")
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


@st.cache_data(show_spinner="데이터 파싱 및 캐싱 중...")
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
                (y for y in ["2020", "2021", "2022", "2023", "2024", "2025", "2026"] if y in file.name), "2026"
            )
            date_col = (
                "매출일자_raw" if "매출일자_raw" in df.columns else df.columns[0]
            )

            df["매출일_dt"] = parse_date_series_robust(
                df[date_col], default_year=file_year
            )

            df["매출액"] = (
                pd.to_numeric(
                    df["매출액"].astype(str).str.replace(r"[^\d.-]", "", regex=True),
                    errors="coerce",
                ).fillna(0)
                if "매출액" in df.columns
                else 0
            )
            df["출고량"] = (
                pd.to_numeric(
                    df["출고량"].astype(str).str.replace(r"[^\d.-]", "", regex=True),
                    errors="coerce",
                ).fillna(0)
                if "출고량" in df.columns
                else 0
            )
            df["단가"] = (
                pd.to_numeric(
                    df["단가"].astype(str).str.replace(r"[^\d.-]", "", regex=True),
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
# 4. 메인 대시보드
# ==========================================
inject_custom_css()

st.title("📊 통합 영업 분석 대시보드")

st.sidebar.header("📁 데이터 관리")
address_file = st.sidebar.file_uploader("거래처 주소록 (CSV)", type=["csv"])
uploaded_files = st.sidebar.file_uploader(
    "매출 데이터 (다중 업로드)", type=["csv"], accept_multiple_files=True
)

addr_dict = (
    load_address_file(address_file.getvalue()) if address_file else {}
)
full_df = load_uploaded_files(uploaded_files) if uploaded_files else pd.DataFrame()

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
    st.sidebar.write("---")
    st.sidebar.subheader("📅 기간 필터")
    start_date = st.sidebar.text_input("조회 시작 (YYMMDD)", "200101")
    end_date = st.sidebar.text_input("조회 종료 (YYMMDD)", "261231")

    start_dt = pd.to_datetime(start_date, format="%y%m%d", errors="coerce")
    end_dt = pd.to_datetime(end_date, format="%y%m%d", errors="coerce")

    df_base = full_df[
        (full_df["매출일_dt"] >= start_dt) & (full_df["매출일_dt"] <= end_dt)
    ].copy()

    st.markdown("### 🔎 필터링")
    c1, c2, c3 = st.columns(3)

    selected_staff = c1.multiselect(
        "👤 담당자 필터",
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
    
    selected_client_list = c2.multiselect(
        "🏢 거래처 검색 및 선택", 
        options=all_clients,
        max_selections=1,
        placeholder="거래처 검색..."
    )
    selected_client = selected_client_list[0] if selected_client_list else "전체 거래처"

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
    selected_item = c3.multiselect("📦 품목명 필터", available_items)

    df_f = (
        df_client_filtered[df_client_filtered["품목명"].isin(selected_item)]
        if selected_item
        else df_client_filtered.copy()
    )

    all_months = [f"{i:02d}월" for i in range(1, 13)]

    if not df_f.empty:
        df_f["연도"] = df_f["매출일_dt"].dt.year.astype(str)
        df_f["월"] = df_f["매출일_dt"].dt.strftime("%m월")
        df_f["분기"] = df_f["매출일_dt"].dt.to_period("Q").astype(str)
        existing_years = sorted(df_f["연도"].unique())
        sorted_cols = [
            (y, m)
            for y in existing_years
            for m in all_months
        ]
    else:
        existing_years = []
        sorted_cols = []

    # ------------------------------------
    # 데이터 피벗 연산
    # ------------------------------------
    # 1. 연도별 월 매출
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
    if not df_f.empty:
        df_f["연도월_정렬"] = (
            df_f["연도"].astype(str).str[2:] + "년 " + df_f["월"].astype(str)
        )
        years_for_pivot = existing_years if existing_years else ["2026"]
        desired_order = [f"{y[2:]}년 {m}" for y in years_for_pivot for m in all_months]
        
        client_pivot = (
            df_f.pivot_table(
                index="거래처",
                columns="연도월_정렬",
                values="매출액",
                aggfunc="sum",
            ).fillna(0)
            / 10000
        )
        client_pivot = client_pivot.reindex(columns=desired_order, fill_value=0)

    # 3. 품목 피벗
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
                sales_expanded_data[col_key] = sales_raw_p[col_key] if col_key in sales_raw_p.columns else 0
                qty_expanded_data[col_key] = qty_raw_p[col_key] if col_key in qty_raw_p.columns else 0
            
            yr_sales_sum = sum(sales_raw_p[(yr, m)] for m in all_months if (yr, m) in sales_raw_p.columns)
            sales_expanded_data[(yr, "연간총합")] = yr_sales_sum

            yr_qty_sum = sum(qty_raw_p[(yr, m)] for m in all_months if (yr, m) in qty_raw_p.columns)
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

    # 5. 상세 내역
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

    # 상단 요약 KPI
    total_sales = df_f["매출액"].sum() if not df_f.empty else 0
    total_qty = df_f["출고량"].sum() if not df_f.empty else 0

    m1, m2 = st.columns(2)
    m1.markdown(
        f"""<div class="metric-box"><div class="metric-label">💰 총 매출 합계</div><div class="metric-value">{total_sales:,.0f} 원</div></div>""",
        unsafe_allow_html=True,
    )
    m2.markdown(
        f"""<div class="metric-box"><div class="metric-label">📦 총 출고량</div><div class="metric-value">{total_qty:,.0f} 개</div></div>""",
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

    # TAB 1: 종합 요약
    with tab1:
        st.markdown(
            '<div class="sub-header">📈 전체 월별 매출 추이 (만 원 단위)</div>',
            unsafe_allow_html=True,
        )
        col_table1, col_chart1 = st.columns([1.2, 1.8])

        with col_table1:
            if not pivot_m.empty:
                st.dataframe(pivot_m.round(0), use_container_width=True, height=400)

        with col_chart1:
            if not df_f.empty:
                chart_m = (
                    df_f.pivot_table(
                        index="월", columns="연도", values="매출액", aggfunc="sum"
                    ).fillna(0)
                    * 1.1
                    / 10000
                )
                chart_m = chart_m.reindex(all_months, fill_value=0)
                st.bar_chart(chart_m, use_container_width=True, height=400)

    # TAB 2: 거래처 분석
    with tab2:
        st.markdown(
            '<div class="sub-header">🏢 거래처별 월별 비교 현황 (만 원 단위)</div>',
            unsafe_allow_html=True,
        )
        if not client_pivot.empty:
            # 안전한 표 표시 (스타일링 비활성화를 통해 다크모드 글자 안 보이는 현상 차단)
            st.dataframe(client_pivot.round(0), use_container_width=True, height=550)
        else:
            st.info("거래처별 데이터가 없습니다.")

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

    # TAB 3: 품목 분석
    with tab3:
        st.markdown(
            '<div class="sub-header">📦 품목별 월별 매출액 / 출고량 / 적용 단가</div>',
            unsafe_allow_html=True,
        )
        if not df_f.empty and not sales_p.empty:
            st.markdown("**📋 [매출액] 품목별 월별 합계 (만 원)**")
            st.dataframe((sales_p / 10000).round(0), use_container_width=True)

            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("**📋 [출고량] 품목별 월별 합계**")
            st.dataframe(qty_p.round(0), use_container_width=True)

            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("**📋 [적용 단가] 품목별 월별 단가 추이**")
            if not unit_price_p.empty:
                st.dataframe(unit_price_p.round(0), use_container_width=True)

    # TAB 4: 담당자
    with tab4:
        st.markdown(
            '<div class="sub-header">👤 담당자별 월별 매출 (만 원 단위)</div>',
            unsafe_allow_html=True,
        )
        if not staff_pivot.empty:
            st.dataframe(staff_pivot.round(0), use_container_width=True)

        st.markdown(
            '<div class="sub-header">📄 상세 거래 내역 데이터</div>',
            unsafe_allow_html=True,
        )
        if not df_detail.empty:
            df_detail_display = df_detail.copy()
            df_detail_display["매출일"] = df_detail_display["매출일_dt"].dt.strftime("%Y-%m-%d")
            df_detail_display = df_detail_display.drop(columns=["매출일_dt"])
            cols_order = ["매출일", "담당자", "거래처", "품목명", "출고량", "단가", "매출액"]
            df_detail_display = df_detail_display[cols_order].sort_values(by="매출일", ascending=False)

            st.dataframe(df_detail_display, use_container_width=True, height=450)
else:
    st.info("👈 왼쪽 사이드바에서 데이터 파일(CSV)을 업로드해주세요.")
