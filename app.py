import io
import os
import re
import sys
import subprocess
import urllib.parse
import requests
from bs4 import BeautifulSoup
import numpy as np
import pandas as pd
import streamlit as st
import plotly.graph_objects as go
import plotly.express as px

# OpenDartReader 임포트 (설치되지 않았을 경우를 대비한 예외 처리)
try:
    import OpenDartReader
except ImportError:
    OpenDartReader = None

# 페이지 및 Styler 가동 한도 설정
pd.set_option("styler.render.max_elements", 2000000)
st.set_page_config(page_title="통합 영업 분석 대시보드", layout="wide")

# ==========================================
# 0. 로컬 파일 자동 저장을 위한 디렉토리 설정
# ==========================================
CACHE_DIR = "./uploaded_cache"
os.makedirs(CACHE_DIR, exist_ok=True)


# ==========================================
# 1. 아이패드/모바일 최적화 CSS Injection (터치 및 스크롤 개선)
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

            /* 아이패드/모바일 터치 스크롤 영역 부드럽게 처리 */
            div[data-testid="stDataFrame"] {
                -webkit-overflow-scrolling: touch;
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
            
            /* Radio 버튼 그룹 최적화 */
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

    is_bulk = p_upper.str.contains("BULK", na=False) | p_str.str.contains("벌크", na=False)
    
    is_ar = is_bulk & (p_upper.str.contains("AR", na=False) | p_str.str.contains("아르곤|아르", na=False))
    is_co2 = is_bulk & (p_upper.str.contains("CO2", na=False) | p_str.str.contains("탄산", na=False))
    is_o2 = is_bulk & ~is_co2 & (p_upper.str.contains("O2", na=False) | p_str.str.contains("산소", na=False))
    is_n2 = is_bulk & (p_upper.str.contains("N2", na=False) | p_str.str.contains("질소", na=False))

    is_n2_liter = is_n2 & (p_upper.str.contains("L|LITER", na=False) | p_str.str.contains("리터", na=False))

    if is_n2_liter.any():
        if "출고량" in df.columns:
            df.loc[is_n2_liter, "출고량"] = df.loc[is_n2_liter, "출고량"] * 0.808
        if "단가" in df.columns:
            df.loc[is_n2_liter, "단가"] = df.loc[is_n2_liter, "단가"] * 1.238
        if "매출액" in df.columns and "출고량" in df.columns and "단가" in df.columns:
            df.loc[is_n2_liter, "매출액"] = df.loc[is_n2_liter, "출고량"] * df.loc[is_n2_liter, "단가"]

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


# ==========================================
# ★ 메모 생성 AppleScript ★
# ==========================================
def open_macos_notes_folder(client_name, dart_api_key):
    safe_client_name = client_name.replace('"', '\\"')

    info = get_company_info_hybrid(client_name, dart_api_key)
    encoded_name = urllib.parse.quote(info['clean_name'])
    
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


@st.cache_data(show_spinner="데이터 파싱 중입니다...")
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
                decoded_text = content.decode("utf-8", errors="replace")

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

            df = pd.read_csv(io.StringIO("\n".join(lines[header_idx:])), on_bad_lines='skip', engine='python')
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

            # 품목명 정규화 (N2 liter -> N2 kg 자동 변환 및 보정)
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
# 4. 피벗 및 지표 계산 연산 캐싱 (속도 극대화 🚀)
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

    raw_up = target_tab3_df[target_tab3_df["단가"] > 0].pivot_table(index="품목명", columns="연도월_정렬", values="단가", aggfunc="median")
    if not raw_up.empty:
        unit_price_p = raw_up.fillna(0)
        
        desired_price_cols = [f"{yr[2:]}년 {m}" for yr in years for m in all_months]
        existing_cols = [c for c in desired_price_cols if c in unit_price_p.columns]
        unit_price_p = unit_price_p[existing_cols]
        
        if latest_col in unit_price_p.columns:
            unit_price_p = unit_price_p.sort_values(by=latest_col, ascending=False)
            
    return sales_p, qty_p, unit_price_p


# ==========================================
# ★ 초고속 빈 껍데기 필터링 로직 ★
# ==========================================
def prepare_active_df_fast(df, target_col):
    if df is None or df.empty:
        return None, [], None
    
    # [최적화 핵심] 데이터 누락 에러(NaN)를 없애고 안전하게 빈 칸만 삭제
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
def cached_staff_pivot(df_base, desired_order):
    if df_base.empty:
        return pd.DataFrame()
    staff_raw = (df_base.pivot_table(index="담당자", columns="연도월_정렬", values="매출액", aggfunc="sum").fillna(0) / 10000)
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
# 5. 메인 실행 흐름 및 영구 캐싱 관리
# ==========================================
inject_custom_css()

st.sidebar.header("📁 데이터 업로드 및 유지")

address_file_up = st.sidebar.file_uploader("거래처 주소록 (CSV)", type=["csv"])
industry_file_up = st.sidebar.file_uploader("🏢 거래처 업종 분류 (CSV)", type=["csv"])
debt_file_up = st.sidebar.file_uploader("채권 데이터 (채권.csv)", type=["csv"])
uploaded_files_up = st.sidebar.file_uploader("매출 데이터 (다중 업로드)", type=["csv"], accept_multiple_files=True)

# ==========================================
# ★ 설비(탱크/기화기) 재고 관리 파일 업로드 추가 ★
# ==========================================
st.sidebar.markdown("---")
st.sidebar.subheader("🏭 설비 재고 관리 (선택)")
tank_file_up = st.sidebar.file_uploader("탱크 재고현황 (CSV/Excel)", type=["csv", "xlsx"])
vaporizer_file_up = st.sidebar.file_uploader("기화기 재고현황 (CSV/Excel)", type=["csv", "xlsx"])

# CSV 큰따옴표 이중묶임 에러 및 열 개수 불일치 방지용 함수 (on_bad_lines 적용)
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
    help="금융감독원 Open DART API 키를 입력하세요. 한 번 입력하면 자동 저장되어 새로고침해도 유지됩니다."
)

if dart_api_key and dart_api_key != saved_api_key:
    with open(API_KEY_FILE, "w", encoding="utf-8") as f:
        f.write(dart_api_key)

# ----------------------------------------------------
# 캐시 파일 경로 설정 및 유지 기능 (탱크 및 기화기 추가)
# ----------------------------------------------------
addr_cache_path = os.path.join(CACHE_DIR, "address.csv")
industry_cache_path = os.path.join(CACHE_DIR, "industry.csv")
debt_cache_path = os.path.join(CACHE_DIR, "debt.csv")
tank_cache_path = os.path.join(CACHE_DIR, "tank_cache.dat")
vaporizer_cache_path = os.path.join(CACHE_DIR, "vaporizer_cache.dat")
sales_cache_dir = os.path.join(CACHE_DIR, "sales")
os.makedirs(sales_cache_dir, exist_ok=True)

# 주소록 로딩
if address_file_up is not None:
    addr_bytes = address_file_up.getvalue()
    with open(addr_cache_path, "wb") as f:
        f.write(addr_bytes)
elif os.path.exists(addr_cache_path):
    with open(addr_cache_path, "rb") as f:
        addr_bytes = f.read()
else:
    addr_bytes = None

# 업종 데이터 로딩
if industry_file_up is not None:
    ind_bytes = industry_file_up.getvalue()
    with open(industry_cache_path, "wb") as f:
        f.write(ind_bytes)
elif os.path.exists(industry_cache_path):
    with open(industry_cache_path, "rb") as f:
        ind_bytes = f.read()
else:
    ind_bytes = None

# 채권 데이터 로딩
if debt_file_up is not None:
    debt_bytes = debt_file_up.getvalue()
    with open(debt_cache_path, "wb") as f:
        f.write(debt_bytes)
elif os.path.exists(debt_cache_path):
    with open(debt_cache_path, "rb") as f:
        debt_bytes = f.read()
else:
    debt_bytes = None

# 탱크 데이터 캐싱
if tank_file_up is not None:
    tank_bytes = tank_file_up.getvalue()
    tank_name = tank_file_up.name
    with open(tank_cache_path, "wb") as f:
        f.write(tank_bytes)
    with open(tank_cache_path + "_name.txt", "w", encoding="utf-8") as f:
        f.write(tank_name)
elif os.path.exists(tank_cache_path) and os.path.exists(tank_cache_path + "_name.txt"):
    with open(tank_cache_path, "rb") as f:
        tank_bytes = f.read()
    with open(tank_cache_path + "_name.txt", "r", encoding="utf-8") as f:
        tank_name = f.read().strip()
else:
    tank_bytes = None
    tank_name = ""

# 기화기 데이터 캐싱
if vaporizer_file_up is not None:
    vaporizer_bytes = vaporizer_file_up.getvalue()
    vaporizer_name = vaporizer_file_up.name
    with open(vaporizer_cache_path, "wb") as f:
        f.write(vaporizer_bytes)
    with open(vaporizer_cache_path + "_name.txt", "w", encoding="utf-8") as f:
        f.write(vaporizer_name)
elif os.path.exists(vaporizer_cache_path) and os.path.exists(vaporizer_cache_path + "_name.txt"):
    with open(vaporizer_cache_path, "rb") as f:
        vaporizer_bytes = f.read()
    with open(vaporizer_cache_path + "_name.txt", "r", encoding="utf-8") as f:
        vaporizer_name = f.read().strip()
else:
    vaporizer_bytes = None
    vaporizer_name = ""

# 매출 데이터 로딩 (다중 파일)
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

# 캐시 초기화 시 탱크, 기화기 파일도 함께 제거되도록 수정
if st.sidebar.button("🗑️ 저장된 캐시 데이터 초기화"):
    for p in [addr_cache_path, industry_cache_path, debt_cache_path, 
              tank_cache_path, tank_cache_path + "_name.txt", 
              vaporizer_cache_path, vaporizer_cache_path + "_name.txt"]:
        if os.path.exists(p): os.remove(p)
    if os.path.exists(API_KEY_FILE):
        os.remove(API_KEY_FILE)
    for f_name in os.listdir(sales_cache_dir):
        os.remove(os.path.join(sales_cache_dir, f_name))
    st.rerun()

# 실제 데이터프레임 파싱 적용
addr_dict = load_address_file(addr_bytes) if addr_bytes else {}
industry_dict = load_industry_file(ind_bytes) if ind_bytes else {}
debt_df = load_debt_file(debt_bytes) if debt_bytes else pd.DataFrame()
df_tank = load_equipment_file(tank_bytes, tank_name) if tank_bytes else pd.DataFrame()
df_vaporizer = load_equipment_file(vaporizer_bytes, vaporizer_name) if vaporizer_bytes else pd.DataFrame()
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

    pivot_m_total = cached_get_yearly_monthly_pivot(df_base, all_months, years)
    client_item_qty_pivot = cached_client_item_qty_pivot(df_client_filtered, years, all_months)
    sales_p, qty_p, unit_price_p = cached_tab3_pivots(df_f, years, all_months)
    staff_pivot = cached_staff_pivot(df_base, desired_order)

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

    filtered_debt_df = pd.DataFrame()
    if not debt_df.empty:
        if selected_staff:
            valid_staff_clients = full_df[full_df["담당자"].isin(selected_staff)]["거래처"].unique()
            filtered_debt_df = debt_df[debt_df["거래처"].isin(valid_staff_clients)].copy()
        else:
            filtered_debt_df = debt_df.copy()
            
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

    # ==========================================
    # ★ 7번 탭 설비 재고 현황 추가 반영 ★
    # ==========================================
    tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs(
        [
            "📌 영업 종합 요약",
            "🏢 거래처 분석",
            "📦 품목 및 단가 분석",
            "👤 담당자 & 상세내역",
            "📌 채권 관리",
            "📍 카카오맵",
            "🏭 설비 재고 현황"
        ]
    )

    # Tab 1: 📌 영업 종합 요약
    with tab1:
        st.markdown("<div class='sub-header'>📊 전체 영업 주요 실적 지표</div>", unsafe_allow_html=True)
        m1, m2, m3, m4 = st.columns(4)
        
        tot_sales_val = df_base["매출액"].sum() * 1.1 / 10000 if not df_base.empty else 0.0
        cur_sales_val = cur_month_sales_total * 1.1 / 10000

        m1.markdown(f"<div class='metric-box'><div class='metric-label'>총 누적 매출 (VAT포함)</div><div class='metric-value'>{tot_sales_val:,.0f} 만원</div></div>", unsafe_allow_html=True)
        m2.markdown(f"<div class='metric-box'><div class='metric-label'>최근 월 매출 ({latest_month_str_total})</div><div class='metric-value'>{cur_sales_val:,.0f} 만원</div></div>", unsafe_allow_html=True)
        m3.markdown(f"<div class='metric-box'><div class='metric-label'>전월 대비 (MoM)</div><div class='metric-value' style='color:{'#E11D48' if mom_rate_total < 0 else '#2563EB'};'>{mom_rate_total:+.0f}%</div></div>", unsafe_allow_html=True)
        m4.markdown(f"<div class='metric-box'><div class='metric-label'>월평균 대비 증감</div><div class='metric-value' style='color:{'#E11D48' if avg_rate_total < 0 else '#2563EB'};'>{avg_rate_total:+.0f}%</div></div>", unsafe_allow_html=True)

        st.markdown("<div class='sub-header'>📊 전체 영업 연도별 월 매출 추이</div>", unsafe_allow_html=True)
        col_left, col_right = st.columns([1, 1])
        
        with col_left:
            st.markdown("<div style='font-size: 14px; font-weight: 600; color: #334155; margin-bottom: 10px;'>📋 연도별 월별 매출액 (만원) 데이터 (VAT 포함)</div>", unsafe_allow_html=True)
            st.dataframe(pivot_m_total.style.format("{:,.0f}").background_gradient(cmap="Blues", axis=None), use_container_width=True, height=420)

        with col_right:
            st.markdown("<div style='font-size: 14px; font-weight: 600; color: #334155; margin-bottom: 10px;'>📈 전체 영업 연도별 월 매출</div>", unsafe_allow_html=True)
            st.plotly_chart(
                create_stacked_bar_chart(pivot_m_total, title_text=""),
                use_container_width=True, key="tab1_total_chart"
            )
            
        st.markdown("---")
        st.markdown("<div class='sub-header'>📦 주요 4대 품목 상세 분석</div>", unsafe_allow_html=True)
        
        sel_col1, sel_col2 = st.columns([1, 1])
        with sel_col1:
            selected_target_item = st.radio("🔍 분석할 품목 선택", target_items, horizontal=True, key="overall_item_radio")
        with sel_col2:
            selected_metric = st.radio("📊 분석 지표 선택", ["매출액 (만원)", "출고량", "총매출 대비 비중 (%)"], horizontal=True, key="overall_metric_radio")
            
        item_pivot = cached_get_item_pivot(df_base, selected_target_item, selected_metric, all_months, years)
        
        i_col_left, i_col_right = st.columns([1, 1])
        with i_col_left:
            st.markdown(f"<div style='font-size: 14px; font-weight: 600; color: #334155; margin-bottom: 10px;'>📋 {selected_target_item} {selected_metric.split(' ')[0]} 데이터</div>", unsafe_allow_html=True)
            
            if "비중" in selected_metric:
                st.dataframe(item_pivot.style.format("{:,.1f}%").background_gradient(cmap="Purples", axis=None), use_container_width=True, height=420)
                y_suf = "%"
                y_fmt = ",.1f"
            elif "출고량" in selected_metric:
                st.dataframe(item_pivot.style.format("{:,.1f}").background_gradient(cmap="Greens", axis=None), use_container_width=True, height=420)
                y_suf = " 천kg"
                y_fmt = ",.1f"
            else:
                st.dataframe(item_pivot.style.format("{:,.0f}").background_gradient(cmap="Blues", axis=None), use_container_width=True, height=420)
                y_suf = " 만원"
                y_fmt = ",.0f"
                
        with i_col_right:
            st.markdown(f"<div style='font-size: 14px; font-weight: 600; color: #334155; margin-bottom: 10px;'>📈 {selected_target_item} 연도별 월별 {selected_metric.split(' ')[0]}</div>", unsafe_allow_html=True)
            st.plotly_chart(
                create_stacked_bar_chart(
                    item_pivot, 
                    title_text="", 
                    y_suffix=y_suf, 
                    y_format=y_fmt
                ),
                use_container_width=True, key="tab1_item_chart"
            )

        st.markdown("---")
        st.markdown("<div class='sub-header'>🏭 업종별(분류별) 상세 분석</div>", unsafe_allow_html=True)
        
        if "업종" in df_base.columns:
            available_industries = sorted([ind for ind in df_base["업종"].unique() if ind != "미분류"])
            if available_industries:
                ind_col1, ind_col2 = st.columns([1, 1])
                with ind_col1:
                    selected_industry = st.selectbox("🔍 분석할 업종(분류) 선택", available_industries, key="industry_selectbox")
                with ind_col2:
                    selected_ind_metric = st.radio("📊 분석 지표 선택", ["매출액 (만원)", "출고량", "총매출 대비 비중 (%)"], horizontal=True, key="industry_metric_radio")
                
                ind_pivot = cached_get_industry_pivot(df_base, selected_industry, selected_ind_metric, all_months, years)
                
                i_col_left2, i_col_right2 = st.columns([1, 1])
                with i_col_left2:
                    st.markdown(f"<div style='font-size: 14px; font-weight: 600; color: #334155; margin-bottom: 10px;'>📋 {selected_industry} {selected_ind_metric.split(' ')[0]} 데이터</div>", unsafe_allow_html=True)
                    
                    if "비중" in selected_ind_metric:
                        st.dataframe(ind_pivot.style.format("{:,.1f}%").background_gradient(cmap="Purples", axis=None), use_container_width=True, height=420)
                        y_suf_i = "%"
                        y_fmt_i = ",.1f"
                    elif "출고량" in selected_ind_metric:
                        st.dataframe(ind_pivot.style.format("{:,.0f}").background_gradient(cmap="Greens", axis=None), use_container_width=True, height=420)
                        y_suf_i = ""
                        y_fmt_i = ",.0f"
                    else:
                        st.dataframe(ind_pivot.style.format("{:,.0f}").background_gradient(cmap="Blues", axis=None), use_container_width=True, height=420)
                        y_suf_i = " 만원"
                        y_fmt_i = ",.0f"
                
                with i_col_right2:
                    st.markdown(f"<div style='font-size: 14px; font-weight: 600; color: #334155; margin-bottom: 10px;'>📈 [{selected_industry}] 연도별 월별 {selected_ind_metric.split(' ')[0]}</div>", unsafe_allow_html=True)
                    st.plotly_chart(
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
                        st.markdown(f"<div style='font-size: 14px; font-weight: 600; color: #1E3A8A; margin-bottom: 10px;'>🏆 [{selected_industry}] 소속 전체 거래처 연도별 매출 요약 (단위: 만원)</div>", unsafe_allow_html=True)
                        
                        ind_client_pivot = df_ind_detail.pivot_table(index="거래처", columns="연도", values="매출액", aggfunc="sum").fillna(0) * 1.1 / 10000
                        
                        avail_years_ind = sorted([y for y in ind_client_pivot.columns if str(y).isdigit()])
                        ind_client_pivot = ind_client_pivot.reindex(columns=avail_years_ind, fill_value=0)
                        ind_client_pivot["총 누적매출"] = ind_client_pivot.sum(axis=1)
                        ind_client_pivot = ind_client_pivot.sort_values(by="총 누적매출", ascending=False)
                        
                        st.dataframe(ind_client_pivot.style.format("{:,.0f}").background_gradient(cmap="Blues", axis=None), use_container_width=True, height=250)
                        
                        st.markdown("<hr style='margin: 15px 0px; border-top: 1px dashed #E2E8F0;'>", unsafe_allow_html=True)
                        
                        st.markdown(f"<div style='font-size: 14px; font-weight: 600; color: #1E3A8A; margin-bottom: 10px;'>🔍 개별 거래처 / 품목별 월 상세 분석</div>", unsafe_allow_html=True)
                        
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
                                if "비중" in sel_ind_sub_metric:
                                    st.dataframe(sub_item_pivot.style.format("{:,.1f}%").background_gradient(cmap="Purples", axis=None), use_container_width=True, height=380)
                                    y_suf_sub = "%"
                                    y_fmt_sub = ",.1f"
                                elif "출고량" in sel_ind_sub_metric:
                                    if sel_ind_item in target_items:
                                        y_suf_sub = " 천kg"
                                        y_fmt_sub = ",.1f"
                                    elif "LPG" in str(sel_ind_item).upper():
                                        y_suf_sub = " kg"
                                        y_fmt_sub = ",.0f"
                                    else:
                                        y_suf_sub = " 개(병)"
                                        y_fmt_sub = ",.0f"
                                    st.dataframe(sub_item_pivot.style.format(f"{{:{y_fmt_sub}}}").background_gradient(cmap="Greens", axis=None), use_container_width=True, height=380)
                                else:
                                    st.dataframe(sub_item_pivot.style.format("{:,.0f}").background_gradient(cmap="Blues", axis=None), use_container_width=True, height=380)
                                    y_suf_sub = " 만원"
                                    y_fmt_sub = ",.0f"
                                    
                            with sc2:
                                st.plotly_chart(
                                    create_stacked_bar_chart(sub_item_pivot, title_text="", y_suffix=y_suf_sub, y_format=y_fmt_sub),
                                    use_container_width=True, key="ind_client_sub_chart"
                                )
                        else:
                            st.info("해당 업종에 소속된 거래처 데이터가 없습니다.")
            else:
                st.info("💡 사이드바에 [🏢 거래처 업종 분류 (CSV)] 파일을 업로드하시면 업종별 실적 분석이 활성화됩니다.")
        else:
            st.info("💡 사이드바에 [🏢 거래처 업종 분류 (CSV)] 파일을 업로드하시면 업종별 실적 분석이 활성화됩니다.")

    # Tab 2: 🏢 거래처 분석
    with tab2:
        st.markdown(f"<div class='sub-header'>🏢 [{selected_client}] 영업 실적 및 요약</div>", unsafe_allow_html=True)
        st.info(f"📍 주소: {client_addr}")
        
        btn_c1, btn_c2, btn_c3 = st.columns([1.5, 1, 1])
        
        with btn_c1:
            if st.button("📝 macOS 메모 앱에서 거래처 노트 열기/생성", key="btn_notes"):
                open_macos_notes_folder(selected_client, dart_api_key)
                
        with btn_c2:
            if "show_corp_info" not in st.session_state:
                st.session_state.show_corp_info = False
                
            btn_label = "🏢 기업정보 닫기" if st.session_state.show_corp_info else "🏢 기업 기본/재무정보 보기"
            
            if st.button(btn_label, key="btn_dart_info"):
                st.session_state.show_corp_info = not st.session_state.show_corp_info
                st.rerun()
                
            if st.session_state.show_corp_info:
                with st.spinner("DART 및 네이버 기업 정보를 불러오는 중..."):
                    c_info = get_company_info_hybrid(selected_client, dart_api_key)
                    st.success("정보 로딩 완료!")
                    
                    if c_info["source"] != "금융감독원 DART (2023)" and dart_api_key:
                        if c_info.get("dart_error"):
                            st.warning(f"⚠️ DART 연동 실패로 네이버 정보를 가져왔습니다. (원인: {c_info['dart_error']})")
                        else:
                            st.warning("⚠️ 해당 기업은 DART에 등록되지 않았거나 재무정보가 없어 네이버 정보를 가져왔습니다.")
                            
                    st.markdown(f"""
                    - **출처:** {c_info['source']}
                    - **대표자:** {c_info['ceo']}
                    - **업종:** {c_info['industry']}
                    - **매출액:** {c_info['revenue']}
                    - **영업이익:** {c_info['profit']}
                    """)
                    
        with btn_c3:
            if client_addr != "등록된 주소 정보가 없습니다.":
                kakao_url = f"https://map.kakao.com/link/search/{urllib.parse.quote(client_addr)}"
                st.link_button("🗺️ 카카오맵에서 주소 보기", kakao_url)
            else:
                st.button("🗺️ 카카오맵에서 주소 보기", disabled=True, key="btn_kakao_disabled")

        m1, m2, m3, m4 = st.columns(4)
        tot_sales_c = df_client_filtered["매출액"].sum() * 1.1 / 10000 if not df_client_filtered.empty else 0.0
        cur_sales_c = cur_month_sales_client * 1.1 / 10000

        m1.markdown(f"<div class='metric-box'><div class='metric-label'>총 누적 매출 (VAT포함)</div><div class='metric-value'>{tot_sales_c:,.0f} 만원</div></div>", unsafe_allow_html=True)
        m2.markdown(f"<div class='metric-box'><div class='metric-label'>최근 월 매출 ({latest_month_str_client})</div><div class='metric-value'>{cur_sales_c:,.0f} 만원</div></div>", unsafe_allow_html=True)
        m3.markdown(f"<div class='metric-box'><div class='metric-label'>전월 대비 (MoM)</div><div class='metric-value' style='color:{'#E11D48' if mom_rate_client < 0 else '#2563EB'};'>{mom_rate_client:+.0f}%</div></div>", unsafe_allow_html=True)
        m4.markdown(f"<div class='metric-box'><div class='metric-label'>월평균 대비 증감</div><div class='metric-value' style='color:{'#E11D48' if avg_rate_client < 0 else '#2563EB'};'>{avg_rate_client:+.0f}%</div></div>", unsafe_allow_html=True)

        if not df_client_filtered.empty:
            pivot_m_client = cached_get_yearly_monthly_pivot(df_client_filtered, all_months, years)
            cl, cr = st.columns([1, 1])
            with cl:
                st.markdown("<div style='font-size: 14px; font-weight: 600; color: #334155; margin-bottom: 10px;'>📋 거래처 월별 매출액 (만원)</div>", unsafe_allow_html=True)
                st.dataframe(pivot_m_client.style.format("{:,.0f}").background_gradient(cmap="Blues", axis=None), use_container_width=True, height=420)
            with cr:
                st.markdown(f"<div style='font-size: 14px; font-weight: 600; color: #334155; margin-bottom: 10px;'>📈 [{selected_client}] 연도별 월 매출</div>", unsafe_allow_html=True)
                st.plotly_chart(
                    create_stacked_bar_chart(pivot_m_client, title_text=""),
                    use_container_width=True, key="tab2_client_total_chart"
                )
                
            st.markdown("---")
            st.markdown(f"<div class='sub-header'>📦 [{selected_client}] 품목별 상세 분석</div>", unsafe_allow_html=True)
            
            client_available_items = sorted(df_client_filtered["품목명"].unique())
            if client_available_items:
                sel_col1_c, sel_col2_c = st.columns([1, 1])
                with sel_col1_c:
                    selected_target_item_c = st.selectbox("🔍 분석할 품목 선택 (전체 거래 품목)", client_available_items, key="client_item_selectbox")
                with sel_col2_c:
                    selected_metric_c = st.radio("📊 분석 지표 선택", ["매출액 (만원)", "출고량", "총매출 대비 비중 (%)"], horizontal=True, key="client_metric_radio")
                    
                client_item_pivot = cached_get_item_pivot(df_client_filtered, selected_target_item_c, selected_metric_c, all_months, years)
                
                i_col_left_c, i_col_right_c = st.columns([1, 1])
                with i_col_left_c:
                    st.markdown(f"<div style='font-size: 14px; font-weight: 600; color: #334155; margin-bottom: 10px;'>📋 {selected_target_item_c} {selected_metric_c.split(' ')[0]} 데이터</div>", unsafe_allow_html=True)
                    
                    if "비중" in selected_metric_c:
                        st.dataframe(client_item_pivot.style.format("{:,.1f}%").background_gradient(cmap="Purples", axis=None), use_container_width=True, height=420)
                        y_suf_c = "%"
                        y_fmt_c = ",.1f"
                    elif "출고량" in selected_metric_c:
                        if selected_target_item_c in target_items:
                            y_suf_c = " 천kg"
                            y_fmt_c = ",.1f"
                        elif "LPG" in str(selected_target_item_c).upper():
                            y_suf_c = " kg"
                            y_fmt_c = ",.0f"
                        else:
                            y_suf_c = " 개(병)"
                            y_fmt_c = ",.0f"
                        st.dataframe(client_item_pivot.style.format(f"{{:{y_fmt_c}}}").background_gradient(cmap="Greens", axis=None), use_container_width=True, height=420)
                    else:
                        st.dataframe(client_item_pivot.style.format("{:,.0f}").background_gradient(cmap="Blues", axis=None), use_container_width=True, height=420)
                        y_suf_c = " 만원"
                        y_fmt_c = ",.0f"
                        
                with i_col_right_c:
                    st.markdown(f"<div style='font-size: 14px; font-weight: 600; color: #334155; margin-bottom: 10px;'>📈 [{selected_client}] {selected_target_item_c} 연도별 월별 {selected_metric_c.split(' ')[0]}</div>", unsafe_allow_html=True)
                    st.plotly_chart(
                        create_stacked_bar_chart(
                            client_item_pivot, 
                            title_text="", 
                            y_suffix=y_suf_c, 
                            y_format=y_fmt_c
                        ),
                        use_container_width=True, key="tab2_client_item_chart"
                    )
            else:
                st.info("해당 거래처에 대한 품목별 상세 데이터가 존재하지 않습니다.")
        else:
            st.warning("선택한 조건에 해당하는 거래처 데이터가 없습니다.")

    # Tab 3: 📦 품목 및 단가 분석 (초고속 빈 열/행 필터링 무손실 렌더링)
    with tab3:
        st.markdown(f"<div class='sub-header'>📦 [{selected_client}] 품목별 실적 분석</div>", unsafe_allow_html=True)
        
        latest_dt_overall = df_base["매출일_dt"].max() if not df_base.empty else None
        target_month_col = latest_dt_overall.strftime("%y년 %m월") if pd.notnull(latest_dt_overall) else None
        
        def render_native_dataframe(df, cmap, fmt, target_col):
            df_active, numeric_cols, highlight_col_name = prepare_active_df_fast(df, target_col)
            
            if df_active is None or df_active.empty:
                st.info("실적 데이터가 없습니다.")
                return

            styled = df_active.style.format(fmt, subset=numeric_cols).background_gradient(cmap=cmap, subset=numeric_cols, axis=None)

            if highlight_col_name:
                styled = styled.apply(lambda s: ['color: #B91C1C; font-weight: bold;'] * len(s), subset=[highlight_col_name], axis=0)

            st.dataframe(
                styled, 
                use_container_width=True, 
                height=400,
                hide_index=True,  
                column_config={
                    "품목명": st.column_config.TextColumn("품목명", width="medium")
                }
            )

        st.markdown("<div style='font-size: 14px; font-weight: 600; color: #334155; margin-bottom: 10px;'>1️⃣ 매출액 (VAT 포함, 만원)</div>", unsafe_allow_html=True)
        render_native_dataframe((sales_p * 1.1 / 10000), "Blues", "{:,.0f}", target_month_col)
            
        st.markdown("<div style='font-size: 14px; font-weight: 600; color: #334155; margin-bottom: 10px;'>2️⃣ 출고량</div>", unsafe_allow_html=True)
        render_native_dataframe(qty_p, "Greens", "{:,.0f}", target_month_col)
            
        st.markdown("<div style='font-size: 14px; font-weight: 600; color: #334155; margin-bottom: 10px;'>3️⃣ 적용 단가 (중간값)</div>", unsafe_allow_html=True)
        render_native_dataframe(unit_price_p, "Oranges", "{:,.0f}", target_month_col)

    # Tab 4: 👤 담당자 & 상세내역
    with tab4:
        st.markdown("<div class='sub-header'>👤 담당자별 월 매출 실적 (만원)</div>", unsafe_allow_html=True)
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
        else:
            st.info("데이터가 없습니다.")

        st.markdown("<div class='sub-header'>🏆 담당자별 거래처 매출 순위 (당해년도)</div>", unsafe_allow_html=True)
        if not df_base.empty:
            current_year = str(df_base["연도"].max())
            all_staffs = sorted(df_base["담당자"].unique())
            
            sel_staff = st.selectbox("👤 순위를 조회할 담당자 선택", all_staffs, key="ranking_staff_select")
            ranking_pivot = cached_ranking_pivot(df_base, current_year, sel_staff, all_months)
            
            if not ranking_pivot.empty:
                r_col1, r_col2 = st.columns([1.2, 1])
                
                with r_col1:
                    st.markdown(f"<div style='font-size: 14px; font-weight: 600; color: #334155; margin-bottom: 10px;'>📋 [{sel_staff}] {current_year}년 거래처 월별 매출 순위</div>", unsafe_allow_html=True)
                    st.dataframe(
                        ranking_pivot.style.format("{:,.0f}")
                        .background_gradient(cmap="Blues", subset=all_months)
                        .background_gradient(cmap="Oranges", subset=["당해 누적 (만원)"]),
                        use_container_width=True, height=380
                    )
                    
                with r_col2:
                    st.markdown(f"<div style='font-size: 14px; font-weight: 600; color: #334155; margin-bottom: 10px;'>📈 상위 10개 거래처 누적 매출액 비교</div>", unsafe_allow_html=True)
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
                    st.plotly_chart(fig_ranking, use_container_width=True, key=f"ranking_chart_{sel_staff}")
            else:
                st.info(f"{current_year}년에 해당 담당자의 거래 내역이 없습니다.")
        else:
            st.info("데이터가 없습니다.")

        st.markdown("<div class='sub-header'>📋 거래 상세 내역 (최신순 800건)</div>", unsafe_allow_html=True)
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
        else:
            st.info("데이터가 없습니다.")

    # Tab 5: 📌 채권 관리
    with tab5:
        st.markdown("<div class='sub-header'>💰 채권(외상대금) 관리 현황 및 연령 분석</div>", unsafe_allow_html=True)
        st.info("💡 **스마트 채권 분석 알림:** 마지막 잔액이 당월 매출보다 클 경우, **어느 달부터 미수금이 밀려 있는지 역추적**하여 매출 칸에 🔴빨간색으로 경고합니다. 정상적으로 수금된 잔액(당월 매출 이하)은 🔵파란색으로 표시됩니다.")
        
        if not debt_df.empty:
            if not filtered_debt_df.empty:
                numeric_cols = [c for c in filtered_debt_df.columns if c not in ["거래처", "구분"]]
                latest_month = numeric_cols[-1] if numeric_cols else None
                
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
                
                def apply_debt_style_fast(df):
                    styles = np.full(df.shape, '', dtype=object)
                    
                    cols = list(df.columns)
                    idx_gubun = cols.index('구분')
                    idx_client = cols.index('거래처')
                    
                    month_col_indices = [cols.index(c) for c in numeric_cols]
                    
                    u_clients_fast = df['거래처'].unique()
                    color_map_fast = {client: '#FFFFFF' if i % 2 == 0 else '#E2E8F0' for i, client in enumerate(u_clients_fast)}
                    
                    for r in range(df.shape[0]):
                        client = df.iat[r, idx_client]
                        bg_color = color_map_fast.get(client, '#FFFFFF')
                        for c in range(df.shape[1]):
                            styles[r, c] = f'background-color: {bg_color};'
                            
                    client_rows = df.groupby('거래처').indices 
                    
                    for client, rows in client_rows.items():
                        i_sal = -1
                        i_bal = -1
                        i_sugum = -1
                        
                        for r in rows:
                            gubun = df.iat[r, idx_gubun]
                            if gubun == '매출': i_sal = r
                            elif gubun == '잔액': i_bal = r
                            elif gubun == '수금': i_sugum = r
                            
                        if i_sal != -1:
                            styles[i_sal, idx_gubun] += ' color: #E11D48; font-weight: 600;'
                        if i_sugum != -1:
                            styles[i_sugum, idx_gubun] += ' color: #059669; font-weight: 600;'
                        if i_bal != -1:
                            for c in range(df.shape[1]):
                                styles[i_bal, c] += ' font-weight: 700; color: #334155;'
                                
                        if i_sal != -1 and i_bal != -1 and month_col_indices:
                            l_month_idx = -1
                            for c in reversed(month_col_indices):
                                val = df.iat[i_bal, c]
                                if pd.notna(val) and val > 0:
                                    l_month_idx = c
                                    break
                                    
                            if l_month_idx != -1:
                                final_bal = df.iat[i_bal, l_month_idx]
                                sal_latest = df.iat[i_sal, l_month_idx]
                                
                                final_bal = float(final_bal) if pd.notna(final_bal) else 0.0
                                sal_latest = float(sal_latest) if pd.notna(sal_latest) else 0.0
                                
                                if final_bal > 0:
                                    if final_bal > sal_latest:
                                        styles[i_bal, l_month_idx] += ' background-color: #FEE2E2; color: #B91C1C; font-weight: 800;'
                                        accumulated = 0.0
                                        for c in reversed(month_col_indices):
                                            if c > l_month_idx: continue
                                            val = df.iat[i_sal, c]
                                            val = float(val) if pd.notna(val) else 0.0
                                            
                                            if accumulated < final_bal:
                                                styles[i_sal, c] += ' background-color: #FEE2E2; color: #B91C1C; font-weight: 800;'
                                                accumulated += val
                                            else:
                                                break
                                    else:
                                        styles[i_bal, l_month_idx] += ' background-color: #EFF6FF; color: #1D4ED8; font-weight: 800;'
                                        
                    return pd.DataFrame(styles, index=df.index, columns=df.columns)

                df_height = 500 if selected_client != "전체 거래처" else 700
                st.dataframe(
                    filtered_debt_df.style
                    .format(subset=numeric_cols, formatter="{:,.0f}")
                    .apply(apply_debt_style_fast, axis=None),
                    use_container_width=True, height=df_height
                )
            else:
                st.warning("선택한 담당자나 거래처에 해당하는 채권 데이터가 없습니다.")
        else:
            st.warning("업로드된 채권 데이터가 없습니다. 사이드바에서 파일을 등록해주세요.")

    # Tab 6: 📍 대한민국 V-World 고해상도 한글/위성 지도 적용
    with tab6:
        st.markdown("<div class='sub-header'>📍 담당자별 거래처 지도 분포 (대한민국 V-World 지도)</div>", unsafe_allow_html=True)
        
        rest_api_key = "21a8c4d7312051598c2e05dba0b9c0c7"
        
        map_col1, map_col2 = st.columns([1, 1])
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
        
        if "show_map" not in st.session_state:
            st.session_state.show_map = False
            
        if st.button("🗺️ 담당자별 거래처 지도 생성하기", type="primary"):
            st.session_state.show_map = True
                
        if st.session_state.show_map:
            if map_selected_staff:
                target_map_df = df_base[df_base["담당자"].isin(map_selected_staff)]
            else:
                target_map_df = df_base.copy()
            
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
                    
                    fig_map = px.scatter_mapbox(
                        map_df,
                        lat="lat",
                        lon="lon",
                        color="담당자",
                        hover_name="거래처",
                        hover_data={"주소": True, "lat": False, "lon": False, "담당자": False},
                        zoom=8,
                        center={"lat": center_lat, "lon": center_lon},
                        height=650
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
                    
                    dynamic_key = f"map_chart_{hash(str(map_selected_staff))}"
                    st.plotly_chart(fig_map, use_container_width=True, key=dynamic_key)
                    
                    if invalid_clients:
                        with st.expander("⚠️ 지도에 표시되지 않은 거래처 (주소 정보 없음 또는 좌표 변환 실패)"):
                            st.write(", ".join(invalid_clients))
                else:
                    st.warning("선택한 조건에 해당하는 거래처 중 유효한 주소가 등록된 곳이 없어 지도를 그릴 수 없습니다.")
            else:
                st.warning("선택한 담당자의 거래처 데이터가 없습니다.")
                
    # Tab 7: 🏭 설비 재고 현황
    with tab7:
        st.markdown("<div class='sub-header'>🏭 고압가스 탱크 및 기화기 재고 현황</div>", unsafe_allow_html=True)
        st.info("💡 **Tip:** CSV나 엑셀 형태로 자산 데이터를 업로드하면 지사별/상태별 유휴 장비를 즉각적으로 조회할 수 있습니다.")

        if not df_tank.empty or not df_vaporizer.empty:
            # --- 4분할 상단 필터 레이아웃 ---
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
                # 동적으로 탱크는 '품목', 기화기는 '기화형식' 추출
                eq_items = []
                if not df_tank.empty and '품목' in df_tank.columns:
                    eq_items.extend(df_tank['품목'].dropna().astype(str).tolist())
                if not df_vaporizer.empty and '기화형식' in df_vaporizer.columns:
                    eq_items.extend(df_vaporizer['기화형식'].dropna().astype(str).tolist())
                unique_eq_items = sorted(list(set([i.strip() for i in eq_items if i.strip() != ''])))
                selected_eq_item = st.selectbox("📦 품목/형식 선택", ["전체 품목/형식"] + unique_eq_items)

            # --- 데이터 업데이트 기준일 (원하시는 위치) ---
            st.markdown("<br>", unsafe_allow_html=True)
            date_col1, date_col2 = st.columns([1.5, 4.5])
            with date_col1:
                st.date_input("📅 데이터 업데이트 기준일")

            st.markdown("---")

            def style_status(val):
                color = '#059669' if '유휴' in str(val) else '#334155'
                weight = 'bold' if '유휴' in str(val) else 'normal'
                return f'color: {color}; font-weight: {weight}'

            # --- 탱크 재고 렌더링 (전체보기 또는 탱크선택 시) ---
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
                    
                    styled_tank = filtered_tank.style.map(style_status, subset=['사용구분'] if '사용구분' in filtered_tank.columns else [])
                    st.dataframe(styled_tank, use_container_width=True, height=350, hide_index=True)
                else:
                    st.warning("업로드된 탱크 재고 데이터가 없습니다.")

            # --- 기화기 재고 렌더링 (전체보기 또는 기화기선택 시) ---
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
                    
                    styled_vap = filtered_vap.style.map(style_status, subset=['사용구분'] if '사용구분' in filtered_vap.columns else [])
                    st.dataframe(styled_vap, use_container_width=True, height=350, hide_index=True)
                else:
                    st.warning("업로드된 기화기 재고 데이터가 없습니다.")

        else:
            st.info("좌측 사이드바에서 데이터 파일을 업로드해주세요. 임시로 이미지를 확인하시려면 아래 버튼을 눌러주세요.")
            if st.button("📷 첨부한 이미지로 임시 확인하기"):
                img_col1, img_col2 = st.columns(2)
                with img_col1:
                    st.image("스크린샷 2026-07-31 오전 8.25.11.png", caption="탱크 현황 (표 변환 권장)", use_column_width=True)
                with img_col2:
                    st.image("스크린샷 2026-07-31 오전 8.25.20.png", caption="기화기 현황 (표 변환 권장)", use_column_width=True)

else:
    st.info("👈 왼쪽 사이드바에서 매출 데이터를 업로드하면 분석 대시보드가 활성화됩니다.")
