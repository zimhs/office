

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
            
            /* 성공 알림 메시지 영역 완전히 숨기기 */
            div[data-testid="stSuccess"] { display: none !important; }
            
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


@st.cache_data(show_spinner="채권 데이터를 읽어오는 중입니다...")
def load_debt_file(debt_bytes):
    if not debt_bytes:
        return pd.DataFrame()
    try:
        for enc in ["utf-8-sig", "cp949", "euc-kr", "utf-8"]:
            try:
                df_raw = pd.read_csv(
                    io.BytesIO(debt_bytes), encoding=enc, header=None
                )
                if df_raw.shape[0] > 6:
                    months_map = {
                        8: "1월",
                        11: "2월",
                        12: "3월",
                        16: "4월",
                        18: "5월",
                        19: "6월",
                        20: "7월",
                    }
                    parsed_rows = []
                    current_client = ""
                    for idx in range(6, df_raw.shape[0]):
                        c_val = df_raw.iloc[idx, 0]
                        if pd.notna(c_val) and str(c_val).strip() != "":
                            current_client = str(c_val).strip()

                        g_val = df_raw.iloc[idx, 7]
                        if (
                            pd.notna(g_val)
                            and str(g_val).strip() in ["이월", "매출", "수금", "잔액"]
                            and current_client
                        ):
                            row_data = {
                                "거래처": current_client,
                                "구분": str(g_val).strip(),
                            }
                            for col_idx, m_name in months_map.items():
                                if col_idx < df_raw.shape[1]:
                                    val_str = str(
                                        df_raw.iloc[idx, col_idx]
                                    ).replace(",", "")
                                    try:
                                        row_data[m_name] = float(val_str)
                                    except:
                                        row_data[m_name] = 0.0
                                else:
                                    row_data[m_name] = 0.0
                            parsed_rows.append(row_data)

                    if parsed_rows:
                        return pd.DataFrame(parsed_rows)
                break
            except UnicodeDecodeError:
                continue
    except Exception:
        pass
    return pd.DataFrame()


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


# 맥북 메모 앱 연동용 AppleScript 함수
def open_macos_notes_folder(client_name):
    script = f'''
    tell application "Notes"
        activate
        try
            set defaultFolder to folder "거래처" of default account
            set targetFolder to folder "{client_name}" of defaultFolder
            
            if (count of notes of targetFolder) > 0 then
                show note 1 of targetFolder
            else
                set newNote to make new note at targetFolder with properties {{name:"{client_name}", body:"{client_name} 관련 영업 및 특이사항 입력"}}
                show newNote
            end if
        on error
            try
                set defaultFolder to folder "거래처" of default account
                set targetFolder to make new folder at defaultFolder with properties {{name:"{client_name}"}}
                set newNote to make new note at targetFolder with properties {{name:"{client_name}", body:"{client_name} 관련 영업 및 특이사항 입력"}}
                show newNote
            on error
                set defaultFolder to make new folder at default account with properties {{name:"거래처"}}
                set targetFolder to make new folder at defaultFolder with properties {{name:"{client_name}"}}
                set newNote to make new note at targetFolder with properties {{name:"{client_name}", body:"{client_name} 관련 영업 및 특이사항 입력"}}
                show newNote
            end try
        end try
    end tell
    '''
    try:
        subprocess.run(["osascript", "-e", script], check=True)
        return True
    except Exception:
        return False


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
debt_file = st.sidebar.file_uploader("채권 데이터 (채권.csv)", type=["csv"])
uploaded_files = st.sidebar.file_uploader(
    "매출 데이터 (다중 업로드)", type=["csv"], accept_multiple_files=True
)

addr_dict = load_address_file(address_file.getvalue()) if address_file else {}
debt_df = load_debt_file(debt_file.getvalue()) if debt_file else pd.DataFrame()
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
    if "연도" not in full_df.columns:
        full_df["연도"] = full_df["매출일_dt"].dt.year.astype(str)
    if "월" not in full_df.columns:
        full_df["월"] = full_df["매출일_dt"].dt.strftime("%m월")
    
    full_df["연도월_정렬"] = (
        full_df["연도"].astype(str).str[2:] + "년 " + full_df["월"].astype(str)
    )

    all_months = [f"{i:02d}월" for i in range(1, 13)]
    raw_years = sorted(full_df["연도"].unique())
    years = sorted(raw_years, reverse=True)
    desired_order = [f"{y[2:]}년 {m}" for y in years for m in all_months]

    # 상단 공통 필터 영역
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

    df_filtered_all = (
        df_client_filtered[df_client_filtered["품목명"].isin(selected_item)]
        if selected_item
        else df_client_filtered.copy()
    )

    st.markdown("<br>", unsafe_allow_html=True)

    tab1, tab2, tab3, tab4, tab5 = st.tabs(
        [
            "📌 영업 종합 요약",
            "🏢 거래처 분석",
            "📦 품목 및 단가 분석",
            "👤 담당자 & 상세내역",
            "📌 채권 관리",
        ]
    )

    # ------------------------------------
    # TAB 1: 영업 종합 요약 (검색 조건 무관 전체 데이터 기반 요약 카드 배치)
    # ------------------------------------
    with tab1:
        global_df_all = full_df.copy()
        global_monthly = global_df_all.groupby(["연도", "월"])["매출액"].sum().reset_index()
        global_monthly["매출액_VAT"] = global_monthly["매출액"] * 1.1
        
        cur_global_sales = 0.0
        prev_global_sales = 0.0
        global_mom_rate = 0.0
        has_global_mom = False
        
        avg_monthly_change_rate = 0.0
        has_avg_change = False

        if not global_monthly.empty:
            global_monthly["월_num"] = global_monthly["월"].astype(str).str.replace("월", "", regex=False).astype(int)
            global_monthly = global_monthly.sort_values(["연도", "월_num"])
            
            global_latest = global_monthly.iloc[-1]
            c_y = int(global_latest["연도"])
            c_m = int(global_latest["월_num"])
            cur_global_sales = global_latest["매출액_VAT"]
            
            p_y = c_y - 1 if c_m == 1 else c_y
            p_m = 12 if c_m == 1 else c_m - 1
            
            global_prev_row = global_monthly[(global_monthly["연도"] == str(p_y)) & (global_monthly["월_num"] == p_m)]
            if not global_prev_row.empty:
                prev_global_sales = global_prev_row["매출액_VAT"].values[0]
                if prev_global_sales > 0:
                    global_mom_rate = ((cur_global_sales - prev_global_sales) / prev_global_sales) * 100
                    has_global_mom = True

            if len(global_monthly) >= 2:
                global_monthly["전월대비증감률"] = global_monthly["매출액_VAT"].pct_change() * 100
                avg_monthly_change_rate = global_monthly["전월대비증감률"].mean()
                has_avg_change = True

        cc1, cc2, cc3 = st.columns(3)
        
        cc1.markdown(f'<div class="metric-box"><div class="metric-label">📅 당월 총매출현황 (VAT 포함)</div><div class="metric-value">{cur_global_sales:,.0f} <span style="font-size: 13px; font-weight: normal; color: #64748B;">원</span></div></div>', unsafe_allow_html=True)
        
        gmom_color = "#EF4444" if global_mom_rate > 0 else ("#3B82F6" if global_mom_rate < 0 else "#64748B")
        gmom_sign = "+" if global_mom_rate > 0 else ""
        gmom_text = f'<strong style="color: {gmom_color};">{gmom_sign}{global_mom_rate:.1f}%</strong>' if has_global_mom else "비교 데이터 없음"
        cc2.markdown(f'<div class="metric-box"><div class="metric-label">📈 전월 매출증감현황</div><div class="metric-value" style="font-size: 18px;">{gmom_text} <span style="font-size: 13px; font-weight: normal; color: #64748B;">(전월 대비)</span></div></div>', unsafe_allow_html=True)
        
        gavg_color = "#EF4444" if avg_monthly_change_rate > 0 else ("#3B82F6" if avg_monthly_change_rate < 0 else "#64748B")
        gavg_sign = "+" if avg_monthly_change_rate > 0 else ""
        gavg_text = f'<strong style="color: {gavg_color};">{gavg_sign}{avg_monthly_change_rate:.1f}%</strong>' if has_avg_change else "데이터 부족"
        cc3.markdown(f'<div class="metric-box"><div class="metric-label">📊 월평균매출 증감현황</div><div class="metric-value" style="font-size: 18px;">{gavg_text} <span style="font-size: 13px; font-weight: normal; color: #64748B;">(평균 증감률)</span></div></div>', unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        st.markdown(
            '<div class="sub-header">📈 전체 월별 매출 추이 (만 원 단위)</div>',
            unsafe_allow_html=True,
        )
        col_table1, col_chart1 = st.columns([1, 1])

        pivot_m = (
            full_df.pivot_table(
                index="월", columns="연도", values="매출액", aggfunc="sum"
            ).fillna(0)
            * 1.1
            / 10000
        )
        pivot_m = pivot_m.reindex(index=all_months, fill_value=0)
        pivot_m = pivot_m.reindex(
            columns=sorted(pivot_m.columns, reverse=True), fill_value=0
        )

        with col_table1:
            st.markdown(
                "**📋 연도별 전체 월 매출 데이터 (VAT 포함, 만 원)**"
            )
            styled_pivot_m = pivot_m.style.format(
                "{:,.0f}"
            ).background_gradient(cmap="Blues", axis=None)
            st.dataframe(
                styled_pivot_m, use_container_width=True, height=360
            )

        with col_chart1:
            st.markdown("**📊 연도별 월 매출 비교 그래프 (VAT 포함)**")
            st.bar_chart(pivot_m, use_container_width=True, height=360)

        st.markdown("---")
        st.markdown(
            '<div class="sub-header">🧪 주요 품목별 매출액 및 출고량 상세 분석</div>',
            unsafe_allow_html=True,
        )

        df_major = full_df[full_df["품목명"].isin(target_items)].copy()
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
                val_col = "매출액" if "매출액" in selected_metric_type else "출고량"
                scale = (1.1 / 10000) if "매출액" in selected_metric_type else (1 / 1000)
                
                item_raw_p = (
                    df_sub_item.pivot_table(
                        index="월",
                        columns="연도",
                        values=val_col,
                        aggfunc="sum",
                    ).fillna(0)
                    * scale
                ).reindex(all_months, fill_value=0)

                c_t1, c_c1 = st.columns([1, 1])
                with c_t1:
                    styled_item_p = item_raw_p.style.format(
                        "{:,.0f}"
                    ).background_gradient(cmap="Purples", axis=None)
                    st.dataframe(
                        styled_item_p, use_container_width=True, height=360
                    )
                with c_c1:
                    st.bar_chart(
                        item_raw_p,
                        use_container_width=True,
                        height=360,
                    )

    # ------------------------------------
    # 공통 함수: 선택된 거래처 기준 요약 카드 렌더링
    # ------------------------------------
    def render_selected_client_metrics():
        if selected_client != "전체 거래처":
            target_client_df = full_df[full_df["거래처"] == selected_client].copy()
        else:
            target_client_df = full_df.copy()

        client_monthly = target_client_df.groupby(["연도", "월"])["매출액"].sum().reset_index()
        client_monthly["매출액_VAT"] = client_monthly["매출액"] * 1.1
        
        cur_c_sales = 0.0
        prev_c_sales = 0.0
        c_mom_rate = 0.0
        has_c_mom = False
        
        avg_c_change_rate = 0.0
        has_avg_c_change = False

        if not client_monthly.empty:
            client_monthly["월_num"] = client_monthly["월"].astype(str).str.replace("월", "", regex=False).astype(int)
            client_monthly = client_monthly.sort_values(["연도", "월_num"])
            
            c_latest = client_monthly.iloc[-1]
            c_y = int(c_latest["연도"])
            c_m = int(c_latest["월_num"])
            cur_c_sales = c_latest["매출액_VAT"]
            
            p_y = c_y - 1 if c_m == 1 else c_y
            p_m = 12 if c_m == 1 else c_m - 1
            
            c_prev_row = client_monthly[(client_monthly["연도"] == str(p_y)) & (client_monthly["월_num"] == p_m)]
            if not c_prev_row.empty:
                prev_c_sales = c_prev_row["매출액_VAT"].values[0]
                if prev_c_sales > 0:
                    c_mom_rate = ((cur_c_sales - prev_c_sales) / prev_c_sales) * 100
                    has_c_mom = True

            if len(client_monthly) >= 2:
                client_monthly["전월대비증감률"] = client_monthly["매출액_VAT"].pct_change() * 100
                avg_c_change_rate = client_monthly["전월대비증감률"].mean()
                has_avg_c_change = True

        st.markdown(f"### 🎯 [{selected_client}] 요약 지표")
        sc1, sc2, sc3 = st.columns(3)
        
        sc1.markdown(f'<div class="metric-box"><div class="metric-label">📅 당월 총매출액 (VAT 포함)</div><div class="metric-value">{cur_c_sales:,.0f} <span style="font-size: 13px; font-weight: normal; color: #64748B;">원</span></div></div>', unsafe_allow_html=True)
        
        cmom_color = "#EF4444" if c_mom_rate > 0 else ("#3B82F6" if c_mom_rate < 0 else "#64748B")
        cmom_sign = "+" if c_mom_rate > 0 else ""
        cmom_text = f'<strong style="color: {cmom_color};">{cmom_sign}{c_mom_rate:.1f}%</strong>' if has_c_mom else "비교 데이터 없음"
        sc2.markdown(f'<div class="metric-box"><div class="metric-label">📈 전월대비 증감현황</div><div class="metric-value" style="font-size: 18px;">{cmom_text}</div></div>', unsafe_allow_html=True)
        
        cavg_color = "#EF4444" if avg_c_change_rate > 0 else ("#3B82F6" if avg_c_change_rate < 0 else "#64748B")
        cavg_sign = "+" if avg_c_change_rate > 0 else ""
        cavg_text = f'<strong style="color: {cavg_color};">{cavg_sign}{avg_c_change_rate:.1f}%</strong>' if has_avg_c_change else "데이터 부족"
        sc3.markdown(f'<div class="metric-box"><div class="metric-label">📊 월평균매출 증감현황</div><div class="metric-value" style="font-size: 18px;">{cavg_text}</div></div>', unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)

    # ------------------------------------
    # TAB 2: 거래처 분석
    # ------------------------------------
    with tab2:
        render_selected_client_metrics()
        
        # 1. 주소 및 연동 버튼 영역 (맥북 메모앱/카카오맵)
        with st.container():
            if selected_client != "전체 거래처":
                target_addr = addr_dict.get(selected_client, "")
                if target_addr and target_addr != "nan":
                    st.markdown(f'<div style="font-size: 18px; font-weight: 700; color: #1E293B; margin-bottom: 12px;">📍 주소: {target_addr}</div>', unsafe_allow_html=True)
                else:
                    st.markdown('<div style="font-size: 16px; font-weight: 600; color: #94A3B8; margin-bottom: 12px;">📍 주소 정보가 없습니다.</div>', unsafe_allow_html=True)
            else:
                st.markdown('<div style="font-size: 16px; font-weight: 600; color: #94A3B8; margin-bottom: 12px;">📍 전체 거래처 조회 중 (거래처를 선택하세요)</div>', unsafe_allow_html=True)

            if selected_client != "전체 거래처":
                target_addr = addr_dict.get(selected_client, "")
                map_url = f"https://map.kakao.com/?q={target_addr}" if (target_addr and target_addr != "nan") else "#"
                map_disabled_style = "" if (target_addr and target_addr != "nan") else "pointer-events: none; opacity: 0.5;"
                
                col_b1, col_b2 = st.columns(2)
                with col_b1:
                    if st.button("📝 맥북 메모앱 연동", use_container_width=True, key="tab2_memo_btn"):
                        open_macos_notes_folder(selected_client)
                with col_b2:
                    st.markdown(f'''
                        <div style="display: flex; align-items: center; width: 100%;">
                            <a href="{map_url}" target="_blank" style="text-decoration: none; width: 100%; {map_disabled_style}">
                                <div style="background-color: #FFFFFF; border: 1px solid rgba(49, 51, 63, 0.2); padding: 0.5rem 1rem; border-radius: 0.5rem; text-align: center; color: rgb(49, 51, 63); font-weight: 400; font-size: 14px; box-shadow: 0 1px 2px rgba(0,0,0,0.05); height: 38px; display: flex; align-items: center; justify-content: center;">
                                    🗺️ 카카오맵 연동
                                </div>
                            </a>
                        </div>
                    ''', unsafe_allow_html=True)
            else:
                col_b1, col_b2 = st.columns(2)
                with col_b1:
                    st.button("📝 맥북 메모앱 연동", use_container_width=True, disabled=True, key="tab2_memo_btn_disabled")
                with col_b2:
                    st.markdown('''
                        <div style="display: flex; align-items: center; width: 100%;">
                            <div style="background-color: #F8FAFC; border: 1px solid rgba(49, 51, 63, 0.1); padding: 0.5rem 1rem; border-radius: 0.5rem; text-align: center; color: #94A3B8; font-weight: 400; font-size: 14px; width: 100%; height: 38px; display: flex; align-items: center; justify-content: center; cursor: not-allowed;">
                                🗺️ 카카오맵 연동
                            </div>
                        </div>
                    ''', unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # 2. 🏢 거래처별 월별 전체품목 실적 분석 (참조형식) 및 버튼·표/그래프 위쪽 배치
        st.markdown(
            '<div class="sub-header">🏢 거래처별 월별 전체품목 실적 분석 (참조형식)</div>',
            unsafe_allow_html=True,
        )

        # 매출현황 / 출고량 2개 선택 버튼 생성
        tab2_metric_choice = st.radio(
            "📊 분석 지표 선택",
            ["매출액 (VAT 포함, 만 원)", "출고량"],
            horizontal=True,
            key="tab2_metric_radio"
        )

        # 데이터 집계 및 피벗 구성
        if selected_client != "전체 거래처":
            t2_base_df = full_df[full_df["거래처"] == selected_client].copy()
        else:
            t2_base_df = full_df.copy()

        if not t2_base_df.empty:
            if "매출액(VAT포함)" not in t2_base_df.columns:
                t2_base_df["매출액(VAT포함)"] = t2_base_df["매출액"] * 1.1

            if "매출액" in tab2_metric_choice:
                t2_val_col = "매출액(VAT포함)"
                t2_scale = 1.0 / 10000  # 만 원 단위
                t2_cmap = "Blues"
            else:
                t2_val_col = "출고량"
                t2_scale = 1.0
                t2_cmap = "Greens"

            t2_pivot = (
                t2_base_df.pivot_table(
                    index="월", columns="연도", values=t2_val_col, aggfunc="sum"
                ).fillna(0)
                * t2_scale
            )
            t2_pivot = t2_pivot.reindex(index=all_months, fill_value=0)
            t2_pivot = t2_pivot.reindex(
                columns=sorted(t2_pivot.columns, reverse=True), fill_value=0
            )

            col_t2_tbl, col_t2_chart = st.columns([1, 1])

            with col_t2_tbl:
                unit_label = "만 원" if "매출액" in tab2_metric_choice else "수량"
                st.markdown(f"**📋 월별 상세 피벗 테이블 ({unit_label})**")
                styled_t2_p = t2_pivot.style.format("{:,.0f}").background_gradient(cmap=t2_cmap, axis=None)
                st.dataframe(styled_t2_p, use_container_width=True, height=360)

            with col_t2_chart:
                st.markdown(f"**📊 연도별 월 비교 그래프 ({tab2_metric_choice})**")
                st.bar_chart(t2_pivot, use_container_width=True, height=360)
        else:
            st.info("선택된 조건에 해당하는 데이터가 없습니다.")

        st.markdown("---")

        # 3. 🏢 거래처별 품목 선택 월별 출고량 분석 (요청하신 기능 추가)
        st.markdown(
            '<div class="sub-header">📦 거래처별 품목 선택 월별 출고량 분석</div>',
            unsafe_allow_html=True,
        )

        if not t2_base_df.empty:
            available_client_items = sorted(t2_base_df["품목명"].unique())
            selected_client_item = st.selectbox(
                "📦 분석할 품목 선택",
                available_client_items,
                key="tab2_item_selector"
            )

            df_client_item_sub = t2_base_df[t2_base_df["품목명"] == selected_client_item]

            if not df_client_item_sub.empty:
                item_pivot = (
                    df_client_item_sub.pivot_table(
                        index="월", columns="연도", values="출고량", aggfunc="sum"
                    ).fillna(0)
                )
                item_pivot = item_pivot.reindex(index=all_months, fill_value=0)
                item_pivot = item_pivot.reindex(
                    columns=sorted(item_pivot.columns, reverse=True), fill_value=0
                )

                col_i_tbl, col_i_chart = st.columns([1, 1])
                with col_i_tbl:
                    st.markdown(f"**📋 [{selected_client_item}] 월별 출고량 피벗 테이블**")
                    styled_item_p = item_pivot.style.format("{:,.0f}").background_gradient(cmap="Greens", axis=None)
                    st.dataframe(styled_item_p, use_container_width=True, height=360)

                with col_i_chart:
                    st.markdown(f"**📊 [{selected_client_item}] 연도별 월 출고량 비교 그래프**")
                    st.bar_chart(item_pivot, use_container_width=True, height=360)
            else:
                st.info("선택한 품목에 대한 출고량 데이터가 없습니다.")

        st.markdown("---")

        # 4. 🏢 거래처별 전체 품목 상세 거래 내역 (아래쪽 배치)
        st.markdown(
            '<div class="sub-header">🏢 거래처별 전체 품목 상세 거래 내역</div>',
            unsafe_allow_html=True,
        )

        df_tab2_detail = df_client_filtered.copy()
        if not df_tab2_detail.empty:
            df_tab2_detail["매출액(VAT포함)"] = df_tab2_detail["매출액"] * 1.1
            display_cols = [c for c in ["매출일_dt", "담당자", "거래처", "품목명", "출고량", "매출액", "매출액(VAT포함)", "단가"] if c in df_tab2_detail.columns]
            
            styled_tab2_df = df_tab2_detail[display_cols].sort_values(by="매출일_dt", ascending=False).style.format(
                {"출고량": "{:,.0f}", "매출액": "{:,.0f}", "매출액(VAT포함)": "{:,.0f}", "단가": "{:,.0f}"}
            )
            st.dataframe(styled_tab2_df, use_container_width=True, height=500)
        else:
            st.info("조회된 상세 거래 데이터가 없습니다.")

    # ------------------------------------
    # TAB 3: 품목 및 단가 분석
    # ------------------------------------
    with tab3:
        render_selected_client_metrics()
        
        st.markdown(
            f'<div class="sub-header">📦 [{selected_client}] 품목별 월별 출고량 분석</div>',
            unsafe_allow_html=True,
        )

        if selected_client != "전체 거래처":
            t3_target_df = full_df[full_df["거래처"] == selected_client].copy()
        else:
            t3_target_df = full_df.copy()

        if not t3_target_df.empty:
            # MultiIndex 컬럼(연도, 월) 형태로 품목별 월별 출고량 피벗 구성
            t3_pivot = t3_target_df.pivot_table(
                index="품목명",
                columns=["연도", "월"],
                values="출고량",
                aggfunc="sum"
            ).fillna(0)

            # 연도 목록 내림차순 정렬 및 월 정렬 보장
            existing_years = sorted(t3_target_df["연도"].unique(), reverse=True)
            multi_cols = [(y, m) for y in existing_years for m in all_months]
            
            # 존재하는 컬럼만 필터링 후 재인덱싱
            valid_multi_cols = [col for col in multi_cols if col in t3_pivot.columns]
            t3_pivot = t3_pivot.reindex(columns=valid_multi_cols, fill_value=0)

            # 연도별 합계 컬럼 추가
            for y in existing_years:
                y_cols = [c for c in t3_pivot.columns if c[0] == y]
                if y_cols:
                    t3_pivot[(y, "연도합계")] = t3_pivot[y_cols].sum(axis=1)

            # 최종 컬럼 정렬 (연도별 월들 뒤에 연도합계 배치)
            final_cols = []
            for y in existing_years:
                m_cols = [(y, m) for m in all_months if (y, m) in t3_pivot.columns]
                sum_col = [(y, "연도합계")] if (y, "연도합계") in t3_pivot.columns else []
                final_cols.extend(m_cols + sum_col)

            t3_pivot = t3_pivot.reindex(columns=final_cols, fill_value=0)

            def highlight_sum_columns(s):
                is_sum = [col[1] == "연도합계" for col in s.index]
                return ['background-color: #FEF3C7; font-weight: bold;' if v else '' for v in is_sum]

            styled_t3 = t3_pivot.style.format("{:,.0f}").background_gradient(cmap="Greens", axis=None)
            
            st.dataframe(styled_t3, use_container_width=True, height=450)
        else:
            st.info("출고량 분석을 위한 데이터가 없습니다.")

        st.markdown("---")

        # ── 요청하신 품목별 선택 월별 출고량 분석(표+그래프)을 품목별 월별 출고량 분석 아래로 이동 ──
        st.markdown(
            '<div class="sub-header">📦 선택 품목별 월별 상세 출고량 분석 (표 및 그래프)</div>',
            unsafe_allow_html=True,
        )

        if not t3_target_df.empty:
            available_t3_items = sorted(t3_target_df["품목명"].unique())
            selected_t3_item = st.selectbox(
                "📦 상세 분석할 품목 선택",
                available_t3_items,
                key="tab3_item_selector"
            )

            df_t3_item_sub = t3_target_df[t3_target_df["품목명"] == selected_t3_item]

            if not df_t3_item_sub.empty:
                t3_item_pivot = (
                    df_t3_item_sub.pivot_table(
                        index="월", columns="연도", values="출고량", aggfunc="sum"
                    ).fillna(0)
                )
                t3_item_pivot = t3_item_pivot.reindex(index=all_months, fill_value=0)
                t3_item_pivot = t3_item_pivot.reindex(
                    columns=sorted(t3_item_pivot.columns, reverse=True), fill_value=0
                )

                col_t3_tbl, col_t3_chart = st.columns([1, 1])
                with col_t3_tbl:
                    st.markdown(f"**📋 [{selected_t3_item}] 월별 출고량 피벗 테이블**")
                    styled_t3_item_p = t3_item_pivot.style.format("{:,.0f}").background_gradient(cmap="Greens", axis=None)
                    st.dataframe(styled_t3_item_p, use_container_width=True, height=360)

                with col_t3_chart:
                    st.markdown(f"**📊 [{selected_t3_item}] 연도별 월 출고량 비교 그래프**")
                    st.bar_chart(t3_item_pivot, use_container_width=True, height=360)
            else:
                st.info("선택한 품목에 대한 출고량 데이터가 없습니다.")

    # ------------------------------------
    # TAB 4: 담당자 & 상세내역
    # ------------------------------------
    with tab4:
        render_selected_client_metrics()
        
        st.markdown(
            '<div class="sub-header">👤 담당자별 실적 및 상세 거래 내역</div>',
            unsafe_allow_html=True,
        )

        if not df_filtered_all.empty:
            staff_summary = (
                df_filtered_all.groupby("담당자")
                .agg(매출액=("매출액", "sum"), 출고량=("출고량", "sum"))
                .reset_index()
            )
            staff_summary["매출액(VAT포함)"] = staff_summary["매출액"] * 1.1
            st.dataframe(
                staff_summary.style.format(
                    {"매출액": "{:,.0f}", "출고량": "{:,.0f}", "매출액(VAT포함)": "{:,.0f}"}
                ),
                use_container_width=True,
            )

            st.markdown("---")
            st.markdown("**🔍 조건별 상세 매출 데이터 원본**")
            st.dataframe(
                df_filtered_all[
                    [
                        "매출일_dt",
                        "담당자",
                        "거래처",
                        "품목명",
                        "출고량",
                        "매출액",
                        "단가",
                    ]
                ].sort_values(by="매출일_dt", ascending=False).style.format(
                    {"출고량": "{:,.0f}", "매출액": "{:,.0f}", "단가": "{:,.0f}"}
                ),
                use_container_width=True,
                height=400,
            )

    # ------------------------------------
    # TAB 5: 채권 관리
    # ------------------------------------
    with tab5:
        render_selected_client_metrics()
        
        st.markdown(
            '<div class="sub-header">📌 거래처별 채권 및 월별 잔액 현황</div>',
            unsafe_allow_html=True,
        )

        if not debt_df.empty:
            numeric_cols = debt_df.select_dtypes(include=[np.number]).columns
            debt_styled = debt_df.style.format({col: "{:,.0f}" for col in numeric_cols})
            st.dataframe(debt_styled, use_container_width=True, height=500)
        else:
            st.info(
                "📁 사이드바에서 올바른 형식의 채권 데이터(채권.csv)를 업로드해 주세요."
            )

else:
    st.info(
        "👈 왼쪽 사이드바에서 분석할 **매출 데이터(CSV)** 파일을 업로드해 주세요."
    )
