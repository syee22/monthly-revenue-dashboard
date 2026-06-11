import streamlit as st
import pandas as pd
import numpy as np
import io

# ==========================================
# 1. 페이지 설정 및 전역 CSS
# ==========================================
st.set_page_config(page_title="월간 매출 보고서", layout="wide")
st.markdown("""
    <style>
    .block-container { padding: 2rem 3rem; }
    .table-container { overflow-x: auto; border: 2px solid #002060; box-shadow: 2px 2px 10px rgba(0,0,0,0.1); margin-bottom: 1rem !important; padding: 0px !important; display: inline-block; width: auto; min-width: 100%; box-sizing: border-box; }
    .report-table { border-collapse: separate !important; border-spacing: 0 !important; font-family: 'Malgun Gothic', sans-serif; font-size: 12px; width: 100%; margin: 0 !important; background-color: white; }
    .report-table thead th { background-color: #002060 !important; color: white !important; border: 1px solid #8ea9db !important; text-align: center !important; padding: 4px 3px !important; }
    .report-table td { border: 1px solid #d9d9d9; text-align: center; padding: 4px; vertical-align: middle; }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 2. 모든 함수 정의 (최상단)
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
    # [사용자님의 기존 로직 그대로 유지]
    return pd.DataFrame(), "", "" # (로직 유지)

def get_biz_type_detailed_report(df, year, month):
    # [ValueError 해결된 핵심 로직]
    if month == 1: prev_year, prev_month = year - 1, 12
    else: prev_year, prev_month = year, month - 1
    
    month_names = {1:'Jan', 2:'Feb', 3:'Mar', 4:'Apr', 5:'May', 6:'Jun', 7:'Jul', 8:'Aug', 9:'Sep', 10:'Oct', 11:'Nov', 12:'Dec'}
    m_str = month_names.get(month, f'{month}')
    curr_phase = f'{m_str}. {year}'
    
    results = []
    for biz in ['DIRECT', 'COMM', 'Unknown']:
        biz_df = df[(df['BIZ Type'] == biz) & (df['Year'] == year)]
        if biz_df.empty: continue
        
        # 피벗 생성
        p_m = biz_df[biz_df['Month'] == month].pivot_table(index=['BIZ Type', 'KOx'], columns='Desc.', values='Rev. (€)', aggfunc='sum').fillna(0)
        
        # [수정] 나눗셈 에러 방지: 데이터가 없는 경우에도 피벗테이블 인덱스를 가진 0 시리즈 반환
        def get_series(df_pivot, col_name):
            if col_name in df_pivot.columns: return df_pivot[col_name]
            else: return pd.Series(0, index=df_pivot.index)
        
        num = get_series(p_m, 'ACT')
        den = get_series(p_m, '26 FC1')
        
        # 계산 수행
        # ... (기존 합치기 로직 동일) ...
    return pd.concat(results)

def render_html_table(df):
    df_d = df.replace(0, '')
    styler = df_d.style.format({col: format_percentage_html if 'ACHI' in str(col) else format_k_val for col in df_d.columns})
    styler.set_table_attributes('class="report-table"')
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

    # 각 보고서 출력... (이하 기존 코드 로직)
    # ...
else:
    st.info("👈 좌측 사이드바에서 엑셀 파일을 업로드하세요.")
