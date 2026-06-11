import streamlit as st
import pandas as pd
import numpy as np
import io
import re

# ==========================================
# 1. 페이지 설정 및 CSS (스타일 유지)
# ==========================================
st.set_page_config(page_title="월간 매출 보고서", layout="wide")
st.markdown("""
    <style>
    .block-container { padding: 2rem 3rem; }
    .report-table { border-collapse: collapse !important; font-family: 'Malgun Gothic', sans-serif; font-size: 12px; width: 100%; background-color: white; }
    .table-container { overflow-x: auto; border: 2px solid #002060; box-shadow: 2px 2px 10px rgba(0,0,0,0.1); margin-bottom: 1rem !important; padding: 0px !important; display: inline-block; width: auto; min-width: 100%; box-sizing: border-box; }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 2. 필수 함수 (NameError 방지: 최상단 정의)
# ==========================================
def format_k_val(val):
    if pd.isna(val) or isinstance(val, str) or val == '': return val
    v = val / 1_000.0
    rounded_int = int(round(v, 0))
    if rounded_int == 0: return str(round(v, 2)) if round(v, 2) != 0 else "0"
    return f"{rounded_int:,}"

def format_percentage_html(val):
    if pd.isna(val) or isinstance(val, str) or val == '': return val
    pct_str = f"{val:.0%}"
    if val >= 1.0: return f'<span style="color: #00b050; font-weight: bold;">{pct_str} ▲</span>'
    elif val > 0: return f'<span style="color: #c00000; font-weight: bold;">{pct_str} ▼</span>'
    return pct_str

def get_numeric_cols(df):
    return [col for col in df.columns if any(x in str(col) for x in ['FC3', 'FC1', 'ACT', 'ACHI'])]

def load_and_preprocess(file):
    xl = pd.ExcelFile(file); sheets = xl.sheet_names
    df = pd.read_excel(xl, sheet_name=sheets[0], header=4).iloc[:, :26]
    df.columns = ['Year', 'Month', 'Desc.', 'Date', 'STP', 'Customer', 'LK No.', "Q'ty", 'Rev. ($)', 'Rev. (€)', 'Rev. ₩', 'BIZ Type', 'Group 1', 'Group 2', 'Project', 'PF', 'Item', 'Source', 'KOx', 'Memo', 'CPS', 'EUR:USD', 'EUR:KRW', 'Business Type', 'Curr.', 'Con.']
    df['BIZ Type'] = df['BIZ Type'].replace(['COMM', 'comm', 'COMMERCIAL', 'commercial'], 'COMM').fillna('Unknown')
    df['Year'] = pd.to_numeric(df['Year'], errors='coerce').fillna(0).astype(int)
    df['Month'] = pd.to_numeric(df['Month'], errors='coerce').fillna(0).astype(int)
    df['Rev. (€)'] = pd.to_numeric(df['Rev. (€)'], errors='coerce').fillna(0)
    return df

def render_html_table(df):
    # 헤더 ACT 노란색 변환 및 스타일링
    df_d = df.replace(0, '')
    new_cols = []
    for col in df_d.columns:
        c_name = str(col[1] if isinstance(col, tuple) else col)
        new_col = (col[0], col[1].replace('ACT', '<span style="color: #FFD700;">ACT</span>')) if isinstance(col, tuple) else f'<span style="color: #FFD700;">{col}</span>' if 'ACT' in c_name else col
        new_cols.append(new_col)
    df_d.columns = pd.MultiIndex.from_tuples(new_cols) if isinstance(df.columns, pd.MultiIndex) else new_cols
    
    format_dict = {col: format_percentage_html if 'ACHI' in str(col) else format_k_val for col in df_d.columns}
    styler = df_d.style.format(format_dict, na_rep='').set_table_attributes('class="report-table"')
    # 합계 행 노란색 강조
    styler.apply(lambda row: ['background-color: #ffffe0 !important; color: #002060 !important; font-weight: bold !important; border-top: 2px solid #8ea9db !important; border-bottom: 2px solid #8ea9db !important;'] * len(row) 
                 if any(k in str(row.name) for k in ['TTL', 'Total', 'Subtotal', '소계']) else [''] * len(row), axis=1)
    return f'<div class="table-container">{styler.to_html(escape=False)}</div>'

# ==========================================
# 3. 메인 실행부
# ==========================================
st.title("📊 통합 월간 매출 보고서 (FC vs ACT 자동 집계)")
uploaded_file = st.sidebar.file_uploader("SAP/엑셀 데이터를 업로드하세요.", type=['xlsx', 'xls'])

if uploaded_file:
    raw_df = load_and_preprocess(uploaded_file)
    selected_year = st.sidebar.selectbox("연도", sorted(raw_df['Year'].unique()))
    selected_month = st.sidebar.selectbox("월", sorted(raw_df['Month'].unique()))

    # 1. CPS 보고서 (ACT 내림차순 정렬 & 소계 유지)
    df_cps, p_col, c_col = build_summary_report(raw_df, ['CPS'], selected_year, selected_month, 'TTL (K.€)')
    if not df_cps.empty:
        # 소계 행 분리 후 데이터 정렬
        total_row = df_cps.loc[[df_cps.index[-1]]] if 'TTL' in str(df_cps.index[-1]) else pd.DataFrame()
        data_rows = df_cps.iloc[:-1]
        df_cps = pd.concat([data_rows.sort_values(by=(c_col, 'ACT'), ascending=False), total_row])
        st.subheader("📌 1. 매출 요약 (CPS 기준)")
        st.markdown(render_html_table(df_cps), unsafe_allow_html=True)

    # 2. Item 보고서 (ACT 내림차순 정렬 & 소계 유지)
    df_item, p_col, c_col = build_summary_report(raw_df[raw_df['Item'].isin(['ICCU1', 'ICCU2', 'VCMS'])], ['Item'], selected_year, selected_month, 'TTL (K.€)')
    if not df_item.empty:
        total_row = df_item.loc[[df_item.index[-1]]]
        data_rows = df_item.iloc[:-1]
        df_item = pd.concat([data_rows.sort_values(by=(c_col, 'ACT'), ascending=False), total_row])
        st.subheader("📌 2. 매출 요약 (Item 기준)")
        st.markdown(render_html_table(df_item), unsafe_allow_html=True)

    # 3. Biz Type 보고서 (함수 내부에서 KOx별 정렬 처리됨)
    df_biz = get_biz_type_detailed_report(raw_df, selected_year, selected_month)
    if not df_biz.empty:
        st.subheader("📌 3. 비즈니스 타입별 매출 요약")
        st.markdown(render_html_table(df_biz), unsafe_allow_html=True)

    # 4 & 5. PE/Core 보고서 (기존 함수 그대로)
    st.subheader("📌 4. Power Electronics 비즈니스")
    st.markdown(render_html_table(get_biz_report(raw_df, "Power", selected_year, selected_month)[0]), unsafe_allow_html=True)
    
    st.subheader("📌 5. Core 비즈니스")
    st.markdown(render_html_table(get_biz_report(raw_df, "Core", selected_year, selected_month)[0]), unsafe_allow_html=True)

    # 6. HKMC 요약 (추가)
    st.subheader("📌 6. HKMC 매출 요약 (KEM-KR/CN)")
    mask = (raw_df['Group 1'] == 'HKMC') & (raw_df['KOx'].isin(['KEM-KR', 'KEM-CN']))
    df_hkmc, p_col, c_col = build_summary_report(raw_df[mask], ['KOx'], selected_year, selected_month, 'TTL (K.€)')
    if not df_hkmc.empty:
        total_row = df_hkmc.loc[[df_hkmc.index[-1]]]
        data_rows = df_hkmc.iloc[:-1]
        df_hkmc = pd.concat([data_rows.sort_values(by=(c_col, 'ACT'), ascending=False), total_row])
        st.markdown(render_html_table(df_hkmc), unsafe_allow_html=True)

else:
    st.info("👈 좌측 사이드바에서 파일을 업로드하세요.")
