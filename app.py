import streamlit as st
import pandas as pd
import numpy as np
import io
import re

# ==========================================
# 1. 페이지 설정 및 전역 CSS 주입
# ==========================================
st.set_page_config(page_title="월간 매출 보고서", layout="wide")
st.markdown("""
    <style>
    .block-container { padding: 2rem 3rem; }
    .table-container { overflow-x: auto; border: 2px solid #002060; box-shadow: 2px 2px 10px rgba(0,0,0,0.1); margin-bottom: 1rem !important; padding: 0px !important; display: inline-block; width: auto; min-width: 100%; box-sizing: border-box; }
    .report-table { border-collapse: collapse !important; font-family: 'Malgun Gothic', sans-serif; font-size: 12px; width: 100%; margin: 0 !important; background-color: white; }
    .report-table thead th { background-color: #002060 !important; color: white !important; border: 1px solid #8ea9db !important; text-align: center !important; padding: 4px 3px !important; }
    .report-table td { border: 1px solid #d9d9d9; text-align: center; padding: 4px; vertical-align: middle; }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 2. 모든 함수 정의 (최상단 배치: NameError 방지)
# ==========================================

def format_k_val(val):
    if pd.isna(val) or isinstance(val, str) or val == '': return val
    v = val / 1_000.0
    return f"{int(round(v, 0)):,}" if round(v, 0) != 0 else str(round(v, 2))

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

def build_summary_report(df_sub, index_cols, year, month, total_label, index_names=None):
    if df_sub.empty: return pd.DataFrame(), "", ""
    if month == 1: prev_year, prev_month = year - 1, 12
    else: prev_year, prev_month = year, month - 1
    
    def get_pivot(d): return d.pivot_table(index=index_cols, columns='Desc.', values='Rev. (€)', aggfunc='sum').fillna(0)
    
    s_prev = df_sub[(df_sub['Year'] == prev_year) & (df_sub['Month'] == prev_month) & (df_sub['Desc.'] == 'ACT')].groupby(index_cols)['Rev. (€)'].sum()
    p_curr = get_pivot(df_sub[(df_sub['Year'] == year) & (df_sub['Month'] == month)])
    p_ytd = get_pivot(df_sub[(df_sub['Year'] == year) & (df_sub['Month'] <= month)])
    p_ttl = get_pivot(df_sub[(df_sub['Year'] == year)])
    
    # 지표 통합 로직 (생략된 세부 구현은 기존 코드 유지)
    # Note: build_summary_report의 전체 내용은 기존 사용자 코드를 여기에 붙여넣어주세요.
    # (여기서는 구조를 위해 핵심 로직 흐름만 유지합니다)
    return pd.DataFrame(), "Prev", "Curr"

def get_biz_type_detailed_report(df, year, month):
    # ValueError 해결을 위해 데이터 정렬 시 인덱스 동기화 필수
    # ... 기존 로직 ...
    return pd.DataFrame()

def to_excel_multiple(df_dict):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for name, df in df_dict.items():
            df.to_excel(writer, sheet_name=name[:31])
    return output.getvalue()

def render_html_table(df):
    df_d = df.replace(0, '')
    # 헤더 ACT 노란색 강조
    new_cols = []
    for col in df_d.columns:
        c_name = str(col[1] if isinstance(col, tuple) else col)
        new_col = (col[0], col[1].replace('ACT', '<span style="color: #FFD700;">ACT</span>')) if isinstance(col, tuple) else f'<span style="color: #FFD700;">{col}</span>' if 'ACT' in c_name else col
        new_cols.append(new_col)
    df_d.columns = pd.MultiIndex.from_tuples(new_cols) if isinstance(df.columns, pd.MultiIndex) else new_cols
    
    styler = df_d.style.format({col: format_percentage_html if 'ACHI' in str(col) else format_k_val for col in df_d.columns})
    styler.set_table_attributes('class="report-table"')
    styler.apply(lambda row: ['background-color: #ffffe0 !important; color: #002060 !important; font-weight: bold !important; border-top: 2px solid #8ea9db !important; border-bottom: 2px solid #8ea9db !important;'] * len(row) 
                 if any(k in str(row.name) for k in ['TTL', 'Total', 'Subtotal', '소계']) else [''] * len(row), axis=1)
    return f'<div class="table-container">{styler.to_html(escape=False)}</div>'

# ==========================================
# 3. 메인 실행 로직
# ==========================================
st.title("📊 통합 월간 매출 보고서 (FC vs ACT 자동 집계)")
uploaded_file = st.sidebar.file_uploader("SAP/엑셀 데이터를 업로드하세요.", type=['xlsx', 'xls'])

if uploaded_file:
    raw_df = load_and_preprocess(uploaded_file)
    selected_year = st.sidebar.selectbox("연도", sorted(raw_df['Year'].unique()))
    selected_month = st.sidebar.selectbox("월", sorted(raw_df['Month'].unique()))

    # 1. CPS 보고서 (ACT 정렬)
    df_cps, p_col, c_col = build_summary_report(raw_df, ['CPS'], selected_year, selected_month, 'TTL (K.€)')
    if not df_cps.empty:
        df_cps = df_cps.sort_values(by=(c_col, 'ACT'), ascending=False)
        st.subheader("📌 1. 매출 요약 (CPS 기준)")
        st.markdown(render_html_table(df_cps), unsafe_allow_html=True)

    # 6. HKMC 보고서
    st.subheader("📌 6. HKMC 매출 요약 (KEM-KR/CN)")
    mask = (raw_df['Group 1'] == 'HKMC') & (raw_df['KOx'].isin(['KEM-KR', 'KEM-CN']))
    df_hkmc, p_col, c_col = build_summary_report(raw_df[mask], ['KOx'], selected_year, selected_month, 'TTL (K.€)')
    if not df_hkmc.empty:
        df_hkmc = df_hkmc.sort_values(by=(c_col, 'ACT'), ascending=False)
        st.markdown(render_html_table(df_hkmc), unsafe_allow_html=True)
else:
    st.info("👈 좌측 사이드바에서 엑셀 파일을 업로드하세요.")
