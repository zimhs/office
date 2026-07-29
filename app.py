import io
import os
import re
import sys
import subprocess
import urllib.parse
import numpy as np
import pandas as pd
import streamlit as st
import plotly.graph_objects as go

# 페이지 및 Styler 가동 한도 설정
pd.set_option("styler.render.max_elements", 2000000)
st.set_page_config(page_title="통합 영업 분석 대시보드", layout="wide")

# ==========================================
# 0. 로컬 파일 자동 저장을 위한 디렉토리 설정
# ==========================================
CACHE_DIR = "./uploaded_cache"
os.makedirs(CACHE_DIR, exist_ok=True)


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
                color: #1E3A8A;
                font-size: 17px;
                font-weight: 700;
                margin-top: 20px;
                margin-bottom: 12px;
                border-left: 4px solid #1E3A8A;
                padding-left: 10px;
            }
            
            /* Radio 버튼 그룹을 좀 더 돋보이게 */
            div[role="radiogroup"] {
                padding: 10px;
                background: white;
                border-radius: 8px;
                border: 1px solid #E2E8F0;
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


def open_macos_notes_folder(client_name):
    safe_client_name = client_name.replace('"', '\\"')

    script = f"""
    tell application "Notes"
        activate
        
        set targetFolderName to "거래처"
        set noteName to "{safe_client_name}"
        set noteFound to false
        
        try
            set targetAccount to account "iCloud"
            if exists folder targetFolderName of targetAccount then
                set parentFolder to folder targetFolderName of targetAccount
                set foundNotes to (notes of parentFolder whose name is noteName)
                if (count of foundNotes) > 0 then
                    show (item 1 of foundNotes)
                    set noteFound to true
                end if
            end if
        end try
        
        if not noteFound then
            repeat with acc in accounts
                if exists folder targetFolderName of acc then
                    set parentFolder to folder targetFolderName of acc
                    set foundNotes to (notes of parentFolder whose name is noteName)
                    if (count of foundNotes) > 0 then
                        show (item 1 of foundNotes)
                        set noteFound to true
                        exit repeat
                    end if
                end if
            end repeat
        end if
        
        if not noteFound then
            try
                set targetAccount to account "iCloud"
            on error
                set targetAccount to first account
            end try
            
            if not (exists folder targetFolderName of targetAccount) then
                make new folder at targetAccount with properties {{name:targetFolderName}}
            end if
            
            set parentFolder to folder targetFolderName of targetAccount
            set newNote to make new note at parentFolder with properties {{name:noteName, body:"<h1>" & noteName & "</h1><br><p>영업 및 특이사항:</p>"}}
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
            error_msg = stderr.decode('utf-8', errors='ignore')
            st.warning(f"AppleScript 오류: {error_msg}")
            return False
    except Exception as e:
        return False


def create_stacked_bar_chart(pivot_df, title_text, y_suffix="", y_format=",.0f"):
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

    fig.update_layout(
        title=dict(text=title_text, font=dict(size=15, color="#1E293B")),
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
        margin=dict(l=10, r=10, t=40, b=40),
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        height=420
    )
    return fig


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


@st.cache_data(show_spinner="채권 데이터를 읽어오는 중입니다...")
def load_debt_file(debt_bytes):
    if not debt_bytes:
        return pd.DataFrame()
    try:
        for enc in ["utf-8-sig", "cp949", "euc-kr", "utf-8"]:
            try:
                df_direct = pd.read_csv(io.BytesIO(debt_bytes), encoding=enc)
                df_direct.columns = df_direct.columns.astype(str).str.strip()
                
                if "거래처" in df_direct.columns and "구분" in df_direct.columns:
                    df_direct["거래처"] = df_direct["거래처"].replace("", np.nan).ffill()
                    df_direct["구분"] = df_direct["구분"].astype(str).str.strip()
                    
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


@st.cache_data(show_spinner="데이터를 파싱 및 캐싱 중입니다...")
def load_uploaded_files_from_bytes(file_tuples):
    if not file_tuples:
        return pd.DataFrame()

    df_list = []
    for file_name, content in file_tuples:
        try:
            decoded_text = None
            for enc in ["cp949", "euc-kr", "utf-8-sig", "utf-8"]:
                try:
                    decoded_text = content.decode(enc)
                    break
                except UnicodeDecodeError:
                    continue
            if decoded_text is None:
                decoded_text = content.decode("utf-8", errors="ignore")

            lines = [line for line in decoded_text.splitlines() if line.strip()]
            if not lines:
                continue

            header_idx = 0
            for i, line in enumerate(lines[:30]):
                if any(
                    k in line
                    for k in [
                        "거래처", "상호", "품목", "제품", "매출액", "담당", "일자", "금액", "단가"
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

            c_staff = find_col(["담당자명", "영업담당", "담당자", "영업사원", "담당"], ["코드", "ID", "번호"])
            c_client = find_col(["거래처명", "상호명", "고객명", "회사명", "거래처", "상호", "고객"], ["코드", "ID", "번호", "담당", "영업"])
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
    return result_df


# ==========================================
# 4. 메인 실행 흐름 및 영구 캐싱 관리
# ==========================================
inject_custom_css()

st.title("📊 통합 영업 분석 대시보드")
st.markdown("<p style='color: #64748B; margin-bottom: 15px;'>실시간 영업 데이터 모니터링 및 품목·거래처별 다차원 분석 시스템</p>", unsafe_allow_html=True)

st.sidebar.header("📁 데이터 업로드 및 유지")

# 1. 파일 업로더
address_file_up = st.sidebar.file_uploader("거래처 주소록 (CSV)", type=["csv"])
debt_file_up = st.sidebar.file_uploader("채권 데이터 (채권.csv)", type=["csv"])
uploaded_files_up = st.sidebar.file_uploader("매출 데이터 (다중 업로드)", type=["csv"], accept_multiple_files=True)

# 2. 로컬 캐시 경로 설정
addr_cache_path = os.path.join(CACHE_DIR, "address.csv")
debt_cache_path = os.path.join(CACHE_DIR, "debt.csv")
sales_cache_dir = os.path.join(CACHE_DIR, "sales")
os.makedirs(sales_cache_dir, exist_ok=True)

# --- 주소록 처리 ---
if address_file_up is not None:
    addr_bytes = address_file_up.getvalue()
    with open(addr_cache_path, "wb") as f:
        f.write(addr_bytes)
elif os.path.exists(addr_cache_path):
    with open(addr_cache_path, "rb") as f:
        addr_bytes = f.read()
else:
    addr_bytes = None

# --- 채권 파일 처리 ---
if debt_file_up is not None:
    debt_bytes = debt_file_up.getvalue()
    with open(debt_cache_path, "wb") as f:
        f.write(debt_bytes)
elif os.path.exists(debt_cache_path):
    with open(debt_cache_path, "rb") as f:
        debt_bytes = f.read()
else:
    debt_bytes = None

# --- 매출 데이터 파일들 처리 ---
if uploaded_files_up and len(uploaded_files_up) > 0:
    for f_name in os.listdir(sales_cache_dir):
        os.remove(os.path.join(sales_cache_dir, f_name))
    
    sales_file_tuples = []
    for f in uploaded_files_up:
        f_bytes = f.getvalue()
        f_path = os.path.join(sales_cache_dir, f.name)
        with open(f_path, "wb") as sf:
            sf.write(f_bytes)
        sales_file_tuples.append((f.name, f_bytes))
else:
    sales_file_tuples = []
    if os.path.exists(sales_cache_dir):
        for f_name in os.listdir(sales_cache_dir):
            if f_name.endswith(".csv"):
                f_path = os.path.join(sales_cache_dir, f_name)
                with open(f_path, "rb") as sf:
                    sales_file_tuples.append((f_name, sf.read()))

# 캐시 초기화 버튼
if st.sidebar.button("🗑️ 저장된 캐시 데이터 초기화"):
    for p in [addr_cache_path, debt_cache_path]:
        if os.path.exists(p): os.remove(p)
    for f_name in os.listdir(sales_cache_dir):
        os.remove(os.path.join(sales_cache_dir, f_name))
    st.rerun()

# 데이터 로딩 실행
addr_dict = load_address_file(addr_bytes) if addr_bytes else {}
debt_df = load_debt_file(debt_bytes) if debt_bytes else pd.DataFrame()
full_df = load_uploaded_files_from_bytes(sales_file_tuples) if sales_file_tuples else pd.DataFrame()

if not full_df.empty:
    is_deposit_row = full_df["품목명"].astype(str).str.contains("입금", na=False)
    full_df = full_df[~is_deposit_row].copy()

# 주요 품목 4가지 선언
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

        if pd.isna(start_dt): start_dt = pd.Timestamp("2000-01-01")
        if pd.isna(end_dt): end_dt = pd.Timestamp("2099-12-31")

        df_base = full_df[(full_df["매출일_dt"] >= start_dt) & (full_df["매출일_dt"] <= end_dt)].copy()

        selected_staff = fc3.multiselect("👤 담당자", sorted(df_base["담당자"].unique()) if not df_base.empty else [])
        df_staff_filtered = df_base[df_base["담당자"].isin(selected_staff)] if selected_staff else df_base.copy()

        all_clients = sorted(df_staff_filtered["거래처"].unique()) if not df_staff_filtered.empty else []
        selected_client_list = fc4.multiselect("🏢 거래처", options=all_clients, max_selections=1, placeholder="거래처 검색...")
        selected_client = selected_client_list[0] if selected_client_list else "전체 거래처"

        df_client_filtered = df_staff_filtered[df_staff_filtered["거래처"] == selected_client] if selected_client != "전체 거래처" else df_staff_filtered.copy()

        available_items = sorted(df_client_filtered["품목명"].unique()) if not df_client_filtered.empty else []
        selected_item = fc5.multiselect("📦 품목명", available_items)

        df_f = df_client_filtered[df_client_filtered["품목명"].isin(selected_item)] if selected_item else df_client_filtered.copy()

    all_months = [f"{i:02d}월" for i in range(1, 13)]
    raw_years = sorted(full_df["연도"].unique()) if "연도" in full_df.columns else ["2026"]
    years = sorted(raw_years, reverse=True)
    desired_order = [f"{y[2:]}년 {m}" for y in years for m in all_months]

    # --- Pivot 계산 유틸리티 ---
    def get_yearly_monthly_pivot(data_df):
        if data_df.empty:
            return pd.DataFrame(0, index=all_months, columns=[str(y) for y in years])
        
        pvt = data_df.pivot_table(
            index="월", columns="연도", values="매출액", aggfunc="sum"
        ).fillna(0) * 1.1 / 10000
        
        pvt = pvt.reindex(index=all_months, fill_value=0)
        all_yrs = [str(y) for y in years]
        pvt = pvt.reindex(columns=all_yrs, fill_value=0)
        return pvt

    # 특정 품목에 대한 지표별(매출액/출고량/비중) 피벗을 계산하는 함수
    def get_item_pivot(data_df, item_name, metric):
        if data_df.empty:
            return pd.DataFrame(0, index=all_months, columns=[str(y) for y in years])
            
        df_item = data_df[data_df["품목명"] == item_name]
        if df_item.empty:
            return pd.DataFrame(0, index=all_months, columns=[str(y) for y in years])
            
        if metric == "매출액 (만원)":
            pvt = df_item.pivot_table(index="월", columns="연도", values="매출액", aggfunc="sum").fillna(0) * 1.1 / 10000
        elif metric == "출고량 (kg)":
            pvt = df_item.pivot_table(index="월", columns="연도", values="출고량", aggfunc="sum").fillna(0)
        elif metric == "총매출 대비 비중 (%)":
            pvt_item = df_item.pivot_table(index="월", columns="연도", values="매출액", aggfunc="sum").fillna(0)
            pvt_total = data_df.pivot_table(index="월", columns="연도", values="매출액", aggfunc="sum").fillna(0)
            pvt = (pvt_item / pvt_total.replace(0, np.nan) * 100).fillna(0)
            
        pvt = pvt.reindex(index=all_months, fill_value=0)
        all_yrs = [str(y) for y in years]
        pvt = pvt.reindex(columns=all_yrs, fill_value=0)
        return pvt

    pivot_m_total = get_yearly_monthly_pivot(df_base)

    client_item_qty_pivot = pd.DataFrame()
    if not df_client_filtered.empty:
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

        client_item_qty_pivot = pd.DataFrame(ci_expanded_data, index=raw_ci_qty.index)

    sales_p = pd.DataFrame()
    qty_p = pd.DataFrame()
    unit_price_p = pd.DataFrame()

    target_tab3_df = df_f
    if not target_tab3_df.empty:
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

        raw_up = target_tab3_df[target_tab3_df["단가"] > 0].pivot_table(index="품목명", columns="연도월_정렬", values="단가", aggfunc="median")
        if not raw_up.empty:
            unit_price_p = raw_up.fillna(0)
            if latest_col in unit_price_p.columns:
                unit_price_p = unit_price_p.sort_values(by=latest_col, ascending=False)

    staff_pivot = pd.DataFrame()
    if not df_base.empty:
        staff_raw = (df_base.pivot_table(index="담당자", columns="연도월_정렬", values="매출액", aggfunc="sum").fillna(0) / 10000)
        staff_cols = [c for c in desired_order if c in staff_raw.columns]
        staff_pivot = staff_raw.reindex(columns=staff_cols, fill_value=0)

    detail_cols = ["매출일_dt", "담당자", "거래처", "품목명", "출고량", "단가", "매출액"]
    df_detail = df_f[detail_cols].copy() if not df_f.empty else pd.DataFrame(columns=detail_cols)

    if not df_base.empty:
        df_total_monthly = df_base.groupby(df_base["매출일_dt"].dt.to_period("M"))["매출액"].sum()
        latest_period_total = df_total_monthly.index.max()
        cur_month_sales_total = df_total_monthly.loc[latest_period_total]
        prev_period_total = latest_period_total - 1
        prev_month_sales_total = df_total_monthly.get(prev_period_total, 0.0)

        mom_rate_total = ((cur_month_sales_total - prev_month_sales_total) / prev_month_sales_total * 100) if prev_month_sales_total > 0 else 0.0
        avg_monthly_sales_total = df_total_monthly.mean()
        avg_rate_total = ((cur_month_sales_total - avg_monthly_sales_total) / avg_monthly_sales_total * 100) if avg_monthly_sales_total > 0 else 0.0
        latest_month_str_total = latest_period_total.strftime("%Y년 %m월")
    else:
        cur_month_sales_total = prev_month_sales_total = mom_rate_total = avg_monthly_sales_total = avg_rate_total = 0.0
        latest_month_str_total = "-"

    if not df_client_filtered.empty:
        df_client_monthly = df_client_filtered.groupby(df_client_filtered["매출일_dt"].dt.to_period("M"))["매출액"].sum()
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


    # ==========================================
    # 채권 데이터 (debt_df) 담당자 및 거래처 필터 종속 처리
    # ==========================================
    filtered_debt_df = pd.DataFrame()
    if not debt_df.empty:
        if selected_staff:
            # 선택된 담당자들의 거래처 목록 추출 (매출 데이터 기준)
            valid_staff_clients = full_df[full_df["담당자"].isin(selected_staff)]["거래처"].unique()
            filtered_debt_df = debt_df[debt_df["거래처"].isin(valid_staff_clients)].copy()
        else:
            filtered_debt_df = debt_df.copy()
            
        if selected_client != "전체 거래처":
            filtered_debt_df = filtered_debt_df[filtered_debt_df["거래처"] == selected_client].copy()

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
    
    # 엑셀 다운로드에도 담당자/거래처 필터가 종속 적용된 채권 데이터 사용
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

    tab1, tab2, tab3, tab4, tab5 = st.tabs(
        [
            "📌 영업 종합 요약",
            "🏢 거래처 분석",
            "📦 품목 및 단가 분석",
            "👤 담당자 & 상세내역",
            "📌 채권 관리",
        ]
    )

    # ==========================================
    # Tab 1: 📌 영업 종합 요약 (전체 기준)
    # ==========================================
    with tab1:
        st.markdown("<div class='sub-header'>📊 전체 영업 주요 실적 지표</div>", unsafe_allow_html=True)
        m1, m2, m3, m4 = st.columns(4)
        
        tot_sales_val = df_base["매출액"].sum() * 1.1 / 10000 if not df_base.empty else 0.0
        cur_sales_val = cur_month_sales_total * 1.1 / 10000

        m1.markdown(f"<div class='metric-box'><div class='metric-label'>총 누적 매출 (VAT포함)</div><div class='metric-value'>{tot_sales_val:,.0f} 만원</div></div>", unsafe_allow_html=True)
        m2.markdown(f"<div class='metric-box'><div class='metric-label'>최근 월 매출 ({latest_month_str_total})</div><div class='metric-value'>{cur_sales_val:,.0f} 만원</div></div>", unsafe_allow_html=True)
        m3.markdown(f"<div class='metric-box'><div class='metric-label'>전월 대비 (MoM)</div><div class='metric-value' style='color:{'#E11D48' if mom_rate_total < 0 else '#2563EB'};'>{mom_rate_total:+.0f}%</div></div>", unsafe_allow_html=True)
        m4.markdown(f"<div class='metric-box'><div class='metric-label'>월평균 대비 증감</div><div class='metric-value' style='color:{'#E11D48' if avg_rate_total < 0 else '#2563EB'};'>{avg_rate_total:+.0f}%</div></div>", unsafe_allow_html=True)

        # 전체 매출 표 & 차트 (히트맵 배경 적용)
        st.markdown("<div class='sub-header'>📊 전체 영업 연도별 월 매출 추이</div>", unsafe_allow_html=True)
        col_left, col_right = st.columns([1, 1])
        
        with col_left:
            st.markdown("##### 📋 연도별 월별 매출액 (만원) 데이터 (VAT 포함)")
            st.dataframe(pivot_m_total.style.format("{:,.0f}").background_gradient(cmap="Blues", axis=None))

        with col_right:
            st.plotly_chart(
                create_stacked_bar_chart(pivot_m_total, "전체 영업 연도별 월 매출", "만원"),
                use_container_width=True
            )
            
        st.markdown("---")
        
        # 주요 4대 품목 상세 분석 표 & 차트 연동 섹션 (히트맵 배경 적용)
        st.markdown("<div class='sub-header'>📦 주요 4대 품목 상세 분석</div>", unsafe_allow_html=True)
        
        sel_col1, sel_col2 = st.columns([1, 1])
        with sel_col1:
            selected_target_item = st.radio("🔍 분석할 품목 선택", target_items, horizontal=True, key="overall_item_radio")
        with sel_col2:
            selected_metric = st.radio("📊 분석 지표 선택", ["매출액 (만원)", "출고량 (kg)", "총매출 대비 비중 (%)"], horizontal=True, key="overall_metric_radio")
            
        item_pivot = get_item_pivot(df_base, selected_target_item, selected_metric)
        
        i_col_left, i_col_right = st.columns([1, 1])
        with i_col_left:
            st.markdown(f"##### 📋 {selected_target_item} {selected_metric.split(' ')[0]} 데이터")
            
            # 선택된 지표에 맞춰 포맷팅 및 그라데이션(컬러맵) 다르게 적용
            if "비중" in selected_metric:
                st.dataframe(item_pivot.style.format("{:,.1f}%").background_gradient(cmap="Purples", axis=None))
                y_suf = "%"
                y_fmt = ",.1f"
            elif "출고량" in selected_metric:
                st.dataframe(item_pivot.style.format("{:,.0f}").background_gradient(cmap="Greens", axis=None))
                y_suf = " kg"
                y_fmt = ",.0f"
            else:
                st.dataframe(item_pivot.style.format("{:,.0f}").background_gradient(cmap="Blues", axis=None))
                y_suf = " 만원"
                y_fmt = ",.0f"
                
        with i_col_right:
            st.plotly_chart(
                create_stacked_bar_chart(
                    item_pivot, 
                    f"{selected_target_item} 연도별 월별 {selected_metric.split(' ')[0]}", 
                    y_suffix=y_suf, 
                    y_format=y_fmt
                ),
                use_container_width=True
            )

    # ==========================================
    # Tab 2: 🏢 거래처 분석
    # ==========================================
    with tab2:
        st.markdown(f"<div class='sub-header'>🏢 [{selected_client}] 영업 실적 및 요약</div>", unsafe_allow_html=True)
        
        # 주소록 데이터 매칭
        client_addr = addr_dict.get(selected_client, "등록된 주소 정보가 없습니다.")
        st.info(f"📍 주소: {client_addr}")
        
        # 버튼 영역을 나누어 나란히 배치
        btn_c1, btn_c2 = st.columns([1, 1])
        
        with btn_c1:
            # macOS 메모 앱 연동
            if st.button("📝 macOS 메모 앱에서 거래처 노트 열기/생성", key="btn_notes"):
                if open_macos_notes_folder(selected_client):
                    st.success(f"'{selected_client}' 메모가 활성화되었습니다.")
                else:
                    st.error("메모 앱을 열 수 없습니다. (macOS 환경인지 확인해주세요)")
                    
        with btn_c2:
            # 카카오맵 주소 검색 연동
            if client_addr != "등록된 주소 정보가 없습니다.":
                kakao_url = f"https://map.kakao.com/link/search/{urllib.parse.quote(client_addr)}"
                st.link_button("🗺️ 카카오맵에서 주소 보기", kakao_url)
            else:
                st.button("🗺️ 카카오맵에서 주소 보기", disabled=True, key="btn_kakao_disabled")

        # 거래처 주요 지표
        m1, m2, m3, m4 = st.columns(4)
        tot_sales_c = df_client_filtered["매출액"].sum() * 1.1 / 10000 if not df_client_filtered.empty else 0.0
        cur_sales_c = cur_month_sales_client * 1.1 / 10000

        m1.markdown(f"<div class='metric-box'><div class='metric-label'>총 누적 매출 (VAT포함)</div><div class='metric-value'>{tot_sales_c:,.0f} 만원</div></div>", unsafe_allow_html=True)
        m2.markdown(f"<div class='metric-box'><div class='metric-label'>최근 월 매출 ({latest_month_str_client})</div><div class='metric-value'>{cur_sales_c:,.0f} 만원</div></div>", unsafe_allow_html=True)
        m3.markdown(f"<div class='metric-box'><div class='metric-label'>전월 대비 (MoM)</div><div class='metric-value' style='color:{'#E11D48' if mom_rate_client < 0 else '#2563EB'};'>{mom_rate_client:+.0f}%</div></div>", unsafe_allow_html=True)
        m4.markdown(f"<div class='metric-box'><div class='metric-label'>월평균 대비 증감</div><div class='metric-value' style='color:{'#E11D48' if avg_rate_client < 0 else '#2563EB'};'>{avg_rate_client:+.0f}%</div></div>", unsafe_allow_html=True)

        if not df_client_filtered.empty:
            pivot_m_client = get_yearly_monthly_pivot(df_client_filtered)
            cl, cr = st.columns([1, 1])
            with cl:
                st.markdown("##### 📋 거래처 월별 매출액 (만원)")
                st.dataframe(pivot_m_client.style.format("{:,.0f}").background_gradient(cmap="Blues", axis=None))
            with cr:
                st.plotly_chart(
                    create_stacked_bar_chart(pivot_m_client, f"[{selected_client}] 연도별 월 매출", "만원"),
                    use_container_width=True
                )
                
            st.markdown("---")
            
            # 선택 거래처의 '모든' 품목 상세 분석 표 & 차트 (히트맵 배경 적용)
            st.markdown(f"<div class='sub-header'>📦 [{selected_client}] 품목별 상세 분석</div>", unsafe_allow_html=True)
            
            client_available_items = sorted(df_client_filtered["품목명"].unique())
            
            if client_available_items:
                sel_col1_c, sel_col2_c = st.columns([1, 1])
                with sel_col1_c:
                    selected_target_item_c = st.selectbox("🔍 분석할 품목 선택 (전체 거래 품목)", client_available_items, key="client_item_selectbox")
                with sel_col2_c:
                    selected_metric_c = st.radio("📊 분석 지표 선택", ["매출액 (만원)", "출고량 (kg)", "총매출 대비 비중 (%)"], horizontal=True, key="client_metric_radio")
                    
                client_item_pivot = get_item_pivot(df_client_filtered, selected_target_item_c, selected_metric_c)
                
                i_col_left_c, i_col_right_c = st.columns([1, 1])
                with i_col_left_c:
                    st.markdown(f"##### 📋 {selected_target_item_c} {selected_metric_c.split(' ')[0]} 데이터")
                    
                    if "비중" in selected_metric_c:
                        st.dataframe(client_item_pivot.style.format("{:,.1f}%").background_gradient(cmap="Purples", axis=None))
                        y_suf_c = "%"
                        y_fmt_c = ",.1f"
                    elif "출고량" in selected_metric_c:
                        st.dataframe(client_item_pivot.style.format("{:,.0f}").background_gradient(cmap="Greens", axis=None))
                        y_suf_c = " kg"
                        y_fmt_c = ",.0f"
                    else:
                        st.dataframe(client_item_pivot.style.format("{:,.0f}").background_gradient(cmap="Blues", axis=None))
                        y_suf_c = " 만원"
                        y_fmt_c = ",.0f"
                        
                with i_col_right_c:
                    st.plotly_chart(
                        create_stacked_bar_chart(
                            client_item_pivot, 
                            f"[{selected_client}] {selected_target_item_c} 연도별 월별 {selected_metric_c.split(' ')[0]}", 
                            y_suffix=y_suf_c, 
                            y_format=y_fmt_c
                        ),
                        use_container_width=True
                    )
            else:
                st.info("해당 거래처에 대한 품목별 상세 데이터가 존재하지 않습니다.")
        else:
            st.warning("선택한 조건에 해당하는 거래처 데이터가 없습니다.")

    # ==========================================
    # Tab 3: 📦 품목 및 단가 분석
    # ==========================================
    with tab3:
        st.markdown(f"<div class='sub-header'>📦 [{selected_client}] 품목별 실적 분석</div>", unsafe_allow_html=True)
        
        st.markdown("##### 1️⃣ 매출액 (VAT 포함, 만원)")
        if not sales_p.empty:
            # ---------------------------------------------------------
            # 요청사항 반영: 소수점 1자리({:,.1f})를 소수점 제거({:,.0f})로 변경
            # ---------------------------------------------------------
            st.dataframe((sales_p * 1.1 / 10000).style.format("{:,.0f}").background_gradient(cmap="Blues", axis=None))
        else:
            st.info("데이터가 없습니다.")
            
        st.markdown("##### 2️⃣ 출고량")
        if not qty_p.empty:
            st.dataframe(qty_p.style.format("{:,.0f}").background_gradient(cmap="Greens", axis=None))
        else:
            st.info("데이터가 없습니다.")
            
        st.markdown("##### 3️⃣ 적용 단가 (중간값)")
        if not unit_price_p.empty:
            st.dataframe(unit_price_p.style.format("{:,.0f}").background_gradient(cmap="Oranges", axis=None))
        else:
            st.info("데이터가 없습니다.")

    # ==========================================
    # Tab 4: 👤 담당자 & 상세내역
    # ==========================================
    with tab4:
        st.markdown("<div class='sub-header'>👤 담당자별 월 매출 실적 (만원)</div>", unsafe_allow_html=True)
        if not staff_pivot.empty:
            st.dataframe(staff_pivot.style.format("{:,.0f}").background_gradient(cmap="Blues", axis=None))
        else:
            st.info("데이터가 없습니다.")

        st.markdown("<div class='sub-header'>📋 거래 상세 내역 (최신순)</div>", unsafe_allow_html=True)
        if not df_detail.empty:
            # 상세 내역은 텍스트 컬럼이 섞여 있으므로 특정 수치형 컬럼에만 그라데이션 적용
            st.dataframe(
                df_detail.sort_values(by="매출일_dt", ascending=False)
                .style.format({"출고량": "{:,.0f}", "단가": "{:,.0f}", "매출액": "{:,.0f}"})
                .background_gradient(subset=["매출액", "출고량"], cmap="Blues"),
                use_container_width=True
            )
        else:
            st.info("데이터가 없습니다.")

    # ==========================================
    # Tab 5: 📌 채권 관리
    # ==========================================
    with tab5:
        st.markdown("<div class='sub-header'>💰 채권(외상대금) 관리 현황</div>", unsafe_allow_html=True)
        
        if not debt_df.empty:
            if not filtered_debt_df.empty:
                numeric_cols = [c for c in filtered_debt_df.columns if c not in ["거래처", "구분"]]
                
                # --- 표 스타일링 및 그라디에이션/하이라이트 함수 ---
                def apply_debt_style(df_to_style):
                    # 1. 거래처별 구분을 확실히 하기 위해 회색 톤을 더 진하게(#E2E8F0) 변경
                    unique_clients = df_to_style['거래처'].unique()
                    color_map = {client: '#FFFFFF' if i % 2 == 0 else '#E2E8F0' for i, client in enumerate(unique_clients)}
                    
                    def row_style(row):
                        bg_color = color_map.get(row['거래처'], '#FFFFFF')
                        # 기본 배경색 설정
                        style_list = [f'background-color: {bg_color}'] * len(row)
                        
                        # '구분' 값에 따른 추가 텍스트 스타일링
                        idx_gubun = df_to_style.columns.get_loc('구분')
                        
                        if row['구분'] == '잔액':
                            # 잔액 행: 전체 굵게 + 파란색 폰트
                            style_list = [f'background-color: {bg_color}; font-weight: 700; color: #1E3A8A;' for _ in range(len(row))]
                        else:
                            if row['구분'] == '매출':
                                style_list[idx_gubun] += '; color: #E11D48; font-weight: 600;'
                            elif row['구분'] == '수금':
                                style_list[idx_gubun] += '; color: #059669; font-weight: 600;'
                                
                        return style_list

                    # 숫자 포맷팅 및 행별 스타일 적용
                    return (
                        df_to_style.style
                        .format(subset=numeric_cols, formatter="{:,.0f}")
                        .apply(row_style, axis=1)
                    )

                # --- 데이터 출력 ---
                # 선택된 거래처가 전체가 아닐경우 표 높이를 조절하여 보기 편하게 함
                df_height = 500 if selected_client != "전체 거래처" else 700
                st.dataframe(apply_debt_style(filtered_debt_df), use_container_width=True, height=df_height)
            else:
                st.warning("선택한 담당자나 거래처에 해당하는 채권 데이터가 없습니다.")
        else:
            st.warning("업로드된 채권 데이터가 없습니다. 사이드바에서 파일을 등록해주세요.")

else:
    st.info("👈 왼쪽 사이드바에서 매출 데이터를 업로드하면 분석 대시보드가 활성화됩니다.")
