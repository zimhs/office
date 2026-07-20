import streamlit as st
import pandas as pd
import io
import urllib.parse
import re

# 페이지 설정
st.set_page_config(page_title="통합 영업 분석 대시보드", layout="wide")

# CSS: 스타일링
st.markdown("""
    <style>
        .stApp { background-color: #0E1117; color: #E0E0E0; }
        [data-testid="stSidebar"] { font-size: 12px; }
        .metric-box { background-color: #161B22; padding: 15px; border-radius: 8px; text-align: center; border: 1px solid #30363d; }
        .metric-label { color: #8B949E; font-size: 12px; margin-bottom: 5px; }
        .metric-value { color: #FFFFFF; font-size: 18px; font-weight: bold; }
        .sub-header { color: #58A6FF; font-size: 16px; font-weight: bold; margin-top: 30px; margin-bottom: 10px; border-left: 4px solid #58A6FF; padding-left: 8px; }
    </style>
""", unsafe_allow_html=True)

st.title("📊 통합 영업 분석 대시보드")

# 1. 데이터 로드 및 처리
df_list = []
addr_dict = {}

st.sidebar.header("📁 데이터 관리")
address_file = st.sidebar.file_uploader("거래처 주소록", type=["csv"])
uploaded_files = st.sidebar.file_uploader("매출 데이터 (다중)", type=["csv"], accept_multiple_files=True)

if uploaded_files:
    for file in uploaded_files:
        try:
            file.seek(0)
            lines = file.getvalue().decode('utf-8', errors='ignore').splitlines()
            header_idx = next((i for i, line in enumerate(lines) if '매출' in line), 0)
            df = pd.read_csv(io.StringIO("\n".join(lines[header_idx:])))
            col_map = {'거래처': ['거래처', '회사명'], '품목명': ['품목', '제품'], '담당자': ['담당자', '영업']}
            new_cols = {c: next((k for k, keywords in col_map.items() if any(kw in c for kw in keywords)), c) for c in df.columns}
            df = df.rename(columns=new_cols)
            year_val = next((y for y in ['2024', '2025', '2026'] if y in file.name), '2026')
            date_col = next((c for c in df.columns if '매출' in c), df.columns[0])
            df['매출일_dt'] = pd.to_datetime(df[date_col].astype(str) + f"/{year_val}", format='%m/%d/%Y', errors='coerce')
            df['매출액'] = pd.to_numeric(df['매출액'].astype(str).str.replace(r'[^\d.]', '', regex=True), errors='coerce').fillna(0)
            df['출고량'] = pd.to_numeric(df['출고량'], errors='coerce').fillna(0)
            
            # 검색 누락 방지: 거래처명 공백 제거 및 품목명 처리
            processed_rows = []
            for idx, row in df.iterrows():
                if '거래처' in row:
                    row['거래처'] = str(row['거래처']).strip()
                
                p_name_raw = str(row['품목명']).strip()
                p_upper = p_name_raw.upper().replace(" ", "")
                
                matched_item = None
                if "(KG,BULK)" in p_upper or "BULK" in p_upper:
                    if "AR" in p_upper or "아르곤" in p_name_raw:
                        matched_item = 'AR (kg, Bulk)'
                    elif "CO2" in p_upper or "탄산" in p_name_raw:
                        matched_item = 'CO2 (kg, Bulk)'
                    elif "O2" in p_upper or "산소" in p_name_raw:
                        matched_item = 'O2 (kg, Bulk)'
                    elif "N2" in p_upper or "질소" in p_name_raw:
                        if "L" in p_upper or "LITER" in p_upper or "리터" in p_name_raw:
                            row['출고량'] = row['출고량'] * 0.808
                        matched_item = 'N2 (kg, Bulk)'
                
                if matched_item:
                    row['품목명'] = matched_item
                
                processed_rows.append(row)
            
            if processed_rows:
                df_filtered_temp = pd.DataFrame(processed_rows)
                df_list.append(df_filtered_temp.dropna(subset=['매출일_dt']))
        except: continue

    if address_file:
        try:
            temp_addr = pd.read_csv(address_file)
            addr_dict = temp_addr.set_index(temp_addr.columns[0])[temp_addr.columns[1]].to_dict()
        except: pass

    if df_list:
        full_df = pd.concat(df_list, ignore_index=True)
    else:
        full_df = pd.DataFrame(columns=['매출일_dt', '담당자', '거래처', '품목명', '출고량', '매출액'])
    
    # 4대 주요 품목 정의
    target_items = [
        'CO2 (kg, Bulk)', 
        'N2 (kg, Bulk)', 
        'O2 (kg, Bulk)', 
        'AR (kg, Bulk)'
    ]

    # 사이드바 설정
    st.sidebar.write("---")
    start_date = st.sidebar.text_input("조회 시작 (YYMMDD)", '240101')
    end_date = st.sidebar.text_input("조회 종료 (YYMMDD)", '261231')

    # 기본 기간 필터 적용 데이터프레임
    df_base = full_df[(full_df['매출일_dt'] >= pd.to_datetime(start_date, format='%y%m%d', errors='coerce')) & 
                      (full_df['매출일_dt'] <= pd.to_datetime(end_date, format='%y%m%d', errors='coerce'))].copy() if not full_df.empty else full_df.copy()

    # 필터 레이아웃
    c1, c2, c3 = st.columns(3)
    
    selected_staff = c1.multiselect("담당자", sorted(df_base['담당자'].dropna().unique()) if not df_base.empty else [])
    df_staff_filtered = df_base[df_base['담당자'].isin(selected_staff)] if selected_staff and not df_base.empty else df_base.copy()

    # 거래처 검색 및 선택
    available_clients = sorted(df_staff_filtered['거래처'].dropna().unique()) if not df_staff_filtered.empty else []
    selected_client = c2.multiselect("거래처", available_clients)
    df_client_filtered = df_staff_filtered[df_staff_filtered['거래처'].isin(selected_client)] if selected_client and not df_staff_filtered.empty else df_staff_filtered.copy()

    # 품목 선택 필터
    available_items = sorted(df_client_filtered['품목명'].dropna().unique()) if not df_client_filtered.empty else []
    selected_item = c3.multiselect("품목명", available_items)
    
    # 품목 검색 조건에 따른 데이터 적용
    if not df_client_filtered.empty:
        if selected_item:
            df_f = df_client_filtered[df_client_filtered['품목명'].isin(selected_item)].copy()
        else:
            df_f = df_client_filtered.copy()
    else:
        df_f = pd.DataFrame(columns=['매출일_dt', '담당자', '거래처', '품목명', '출고량', '매출액', '연도', '월', '년월', '분기'])

    # 상단 요약 지표
    m1, m2 = st.columns(2)
    total_sales = df_f['매출액'].sum() if not df_f.empty else 0
    total_qty = df_f['출고량'].sum() if not df_f.empty else 0
    
    label_suffix = "(선택 품목)" if selected_item else "(검색된 전체 품목)"
    m1.markdown(f"""<div class="metric-box"><div class="metric-label">총 매출합계 {label_suffix}</div><div class="metric-value">{total_sales:,.0f} 원</div></div>""", unsafe_allow_html=True)
    m2.markdown(f"""<div class="metric-box"><div class="metric-label">총 출고량 {label_suffix}</div><div class="metric-value">{total_qty:,.0f} 개</div></div>""", unsafe_allow_html=True)

    if not df_f.empty:
        df_f['연도'] = df_f['매출일_dt'].dt.year.astype(str)
        df_f['월'] = df_f['매출일_dt'].dt.strftime('%m월')
        df_f['년월'] = df_f['매출일_dt'].dt.strftime('%y-%m')
        df_f['분기'] = df_f['매출일_dt'].dt.to_period('Q').astype(str)

    # 1. 전체 월별 매출 추이 비교 (종속 전체 품목 대상)
    st.markdown('<div class="sub-header">📈 전체 월별 매출 추이 비교 (천 원 단위)</div>', unsafe_allow_html=True)
    col_table1, col_chart1 = st.columns([1.1, 1.9])
    
    with col_table1:
        st.markdown("**📋 연도별 월 매출 데이터**")
        if not df_f.empty:
            yearly_monthly_pivot = df_f.pivot_table(index='연도', columns='월', values='매출액', aggfunc='sum').fillna(0) / 1000
            all_m_cols = [f"{i:02d}월" for i in range(1, 13)]
            yearly_monthly_pivot = yearly_monthly_pivot[[m for m in all_m_cols if m in yearly_monthly_pivot.columns]]
            st.dataframe(yearly_monthly_pivot.style.format("{:,.0f}"), use_container_width=True)
        else:
            st.warning("데이터 없음")
        
    with col_chart1:
        st.markdown("**📊 연도 동월 비교 막대그래프**")
        if not df_f.empty:
            total_compare_chart = df_f.pivot_table(index='월', columns='연도', values='매출액', aggfunc='sum').fillna(0) / 1000
            total_compare_chart = total_compare_chart.reindex([m for m in [f"{i:02d}월" for i in range(1, 13)] if m in total_compare_chart.index])
            st.bar_chart(total_compare_chart, use_container_width=True)
        else:
            st.info("표시할 그래프 데이터가 없습니다.")

    # 2. 거래처별 연도별 월별 매출 현황
    st.markdown('<div class="sub-header">🏢 거래처별 연도별 월별 매출 현황 (천 원 단위)</div>', unsafe_allow_html=True)
    if not df_f.empty:
        client_year_month_pivot = df_f.pivot_table(index='거래처', columns=['연도', '월'], values='매출액', aggfunc='sum').fillna(0) / 1000
        existing_years = sorted(df_f['연도'].unique())
        sorted_cols = [(y, m) for y in existing_years for m in [f"{i:02d}월" for i in range(1, 13)]]
        valid_cols = [col for col in sorted_cols if col in client_year_month_pivot.columns]
        if valid_cols:
            client_year_month_pivot = client_year_month_pivot[valid_cols]
        st.dataframe(client_year_month_pivot.style.format("{:,.0f}"), use_container_width=True)
    else:
        st.info("거래처별 데이터를 표시할 내용이 없습니다.")

    # 3. 담당자별 연도별 월별 매출 현황
    st.markdown('<div class="sub-header">👤 담당자별 연도별 월별 매출 현황 (천 원 단위)</div>', unsafe_allow_html=True)
    if not df_f.empty:
        staff_year_month_pivot = df_f.pivot_table(index='담당자', columns=['연도', '월'], values='매출액', aggfunc='sum').fillna(0) / 1000
        if valid_cols:
            staff_year_month_pivot = staff_year_month_pivot.reindex(columns=valid_cols, fill_value=0)
        st.dataframe(staff_year_month_pivot.style.format("{:,.0f}"), use_container_width=True)
    else:
        st.info("담당자별 데이터를 표시할 내용이 없습니다.")

    # 주요 품목 전용 데이터 분리 (주요 4품목만 집중 비교)
    main_df = df_f[df_f['품목명'].isin(target_items)].copy() if not df_f.empty else df_f.copy()

    # 4. 주요품목 분기별 비교 현황
    st.markdown(f'<div class="sub-header">📅 주요품목 ({", ".join(target_items)}) 분기별 비교 현황 (천 원 단위)</div>', unsafe_allow_html=True)
    col_table3, col_chart3 = st.columns([1.1, 1.9])
    
    with col_table3:
        st.markdown("**📋 주요품목 분기별 매출 데이터**")
        if not main_df.empty:
            quarter_pivot = main_df.pivot_table(index='분기', columns='품목명', values='매출액', aggfunc='sum').fillna(0) / 1000
            quarter_pivot = quarter_pivot.reindex(columns=target_items, fill_value=0)
            st.dataframe(quarter_pivot.style.format("{:,.0f}"), use_container_width=True)
        else:
            st.info("분기별 데이터를 표시할 주요품목 내용이 없습니다.")
            
    with col_chart3:
        st.markdown("**📊 분기별 주요품목 비교 그래프**")
        if not main_df.empty:
            quarter_chart_df = main_df.pivot_table(index='분기', columns='품목명', values='매출액', aggfunc='sum').fillna(0) / 1000
            quarter_chart_df = quarter_chart_df.reindex(columns=target_items, fill_value=0)
            st.bar_chart(quarter_chart_df, use_container_width=True)
        else:
            st.info("표시할 분기별 그래프 데이터가 없습니다.")

    # 5. 품목별 연도별·월별 매출 및 출고량 종합 요약표
    st.markdown('<div class="sub-header">📦 품목별 연도별·월별 매출 및 출고량 종합 요약표</div>', unsafe_allow_html=True)
    if not df_f.empty:
        sales_pivot = df_f.pivot_table(index='품목명', columns=['연도', '월'], values='매출액', aggfunc='sum').fillna(0) / 1000
        if valid_cols:
            sales_pivot = sales_pivot[valid_cols]
        
        st.markdown("**📋 [매출액] 품목별 연도/월별 합계 (천 원)**")
        st.dataframe(sales_pivot.style.format("{:,.0f}"), use_container_width=True)

        qty_pivot = df_f.pivot_table(index='품목명', columns=['연도', '월'], values='출고량', aggfunc='sum').fillna(0)
        if valid_cols:
            qty_pivot = qty_pivot[valid_cols]
            
        st.markdown("**📋 [출고량] 품목별 연도/월별 합계 (소수점 제외)**")
        st.dataframe(qty_pivot.style.format("{:,.0f}"), use_container_width=True)
    else:
        st.info("요약표를 표시할 데이터가 없습니다.")

    # 상세 내역 아코디언
    with st.expander("📝 전체 상세 거래 내역 보기"):
        if not df_f.empty:
            st.dataframe(df_f[['매출일_dt', '담당자', '거래처', '품목명', '출고량', '매출액']].style.format({
                '출고량': '{:,.2f}',
                '매출액': '{:,.0f}'
            }), use_container_width=True)
        else:
            st.info("상세 내역이 없습니다.")

    # 사이드바 지도 검색
    if selected_client and addr_dict:
        st.sidebar.write("---")
        st.sidebar.subheader("📍 거래처 지도 검색")
        for c in selected_client:
            addr = addr_dict.get(c)
            if addr:
                st.sidebar.markdown(f"**{c}**")
                st.sidebar.caption(f"{addr}")
                st.sidebar.link_button(f"🗺️ 지도 보기", f"https://map.kakao.com/?q={urllib.parse.quote(addr)}", use_container_width=True)
else:
    st.info("좌측 사이드바에서 CSV 파일을 업로드해 주세요.")