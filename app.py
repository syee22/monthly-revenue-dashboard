import streamlit as st
import pandas as pd
import numpy as np
import io

# ==========================================
# 1. 페이지 설정 및 전역 CSS 주입
# ==========================================
st.set_page_config(page_title="월간 매출 보고서", layout="wide")

st.markdown("""
    <style>
    .block-container { padding: 2rem 3rem; }
    h1 { font-size: 1.6rem !important; margin-bottom: 0.5rem !important; }
    h3 { font-size: 1.1rem !important; margin-top: 1rem !important; color: #002060 !important; }
    .table-container { 
        overflow-x: auto; border: 2px solid #002060; box-shadow: 2px 2px 10px rgba(0,0,0,0.1); 
        display: inline-block; width: 100%; margin-bottom: 1rem !important; padding: 0px !important; 
    }
    .report-table { border-collapse: collapse !important; font-family: 'Malgun Gothic', sans-serif; font-size: 12px; width: 100%; margin: 0 !important; }
    .report-table thead th { background-color: #002060 !important; color: white !important; border: 1px solid #8ea9db !important; text-align: center !important; }
    .report-table td { border: 1px solid #d9d9d9; text-align: center; padding: 4px; }
    </style>
""", unsafe_allow_html=True)

st.title("📊 통합 월간 매출 보고서 (FC vs ACT 자동 집계)")

# ==========================================
# 2. 유틸리티 함수
# ==========================================
def format_k_val(val):
    if pd.isna(val) or isinstance(val, str) or val == '': return val
    v = val / 1_000.0
    rounded_int = int(round(v, 0))
    if rounded_int == 0:
        v_rounded = round(v, 2)
        return str(v_rounded) if v_rounded != 0 else "0"
    return f"{rounded_int:,}"

def format_percentage_html(val):
    if pd.isna(val) or isinstance(val, str) or val == '': return val
    pct_str = f"{val:.0%}"
    if val >= 1.0: return f'<span style="color: #00b050; font-weight: bold;">{pct_str} ▲</span>'
    elif val > 0: return f'<span style="color: #c00000; font-weight: bold;">{pct_str} ▼</span>'
    else: return pct_str

def get_numeric_cols(df):
    return [col for col in df.columns if any(x in str(col) for x in ['FC3', 'FC1', 'ACT', 'ACHI'])]

def render_html_view(df):
    new_cols = []
    for col in df.columns:
        if isinstance(col, tuple):
            new_cols.append(tuple(f'<span style="color: #FFD700;">{c}</span>' if 'ACT' in str(c) else c for c in col))
        else:
            new_cols.append(f'<span style="color: #FFD700;">{col}</span>' if 'ACT' in str(col) else col)
    
    df_display = df.copy()
    df_display.columns = pd.MultiIndex.from_tuples(new_cols) if isinstance(df.columns, pd.MultiIndex) else new_cols
    df_display = df_display.replace(0, '')
    
    format_dict = {col: format_percentage_html if 'ACHI' in str(col) else format_k_val for col in df_display.columns}
    styler = df_display.style.format(format_dict, na_rep='').set_table_attributes('class="report-table"')
    styler.set_properties(subset=get_numeric_cols(df_display), **{'text-align': 'right'})
    styler.apply(lambda row: ['background-color: #ffffe0; color: #002060; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;'] * len(row) 
                 if any(k in str(row.name) for k in ['TTL', 'Total', 'Subtotal', '소계']) else [''] * len(row), axis=1)
    return f'<div class="table-container">{styler.to_html(escape=False)}</div>'

def to_excel_multiple(df_dict):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for name, df in df_dict.items():
            styler = df.style.format({c: format_k_val for c in df.columns if 'ACT' in str(c) or 'FC' in str(c)})
            styler.to_excel(writer, sheet_name=name[:31])
    return output.getvalue()

# ==========================================
# 3. 메인 로직
# ==========================================
uploaded_file = st.sidebar.file_uploader("SAP/엑셀 데이터를 업로드하세요.", type=['xlsx', 'xls'])

if uploaded_file:
    @st.cache_data
    def load_and_preprocess(file):
        xl = pd.ExcelFile(file)
        sheets = xl.sheet_names
        df = pd.read_excel(xl, sheet_name=sheets[0], header=4).iloc[:, :26]
        df.columns = ['Year', 'Month', 'Desc.', 'Date', 'STP', 'Customer', 'LK No.', "Q'ty", 
                      'Rev. ($)', 'Rev. (€)', 'Rev. ₩', 'BIZ Type', 'Group 1', 'Group 2', 
                      'Project', 'PF', 'Item', 'Source', 'KOx', 'Memo', 'CPS', 
                      'EUR:USD', 'EUR:KRW', 'Business Type', 'Curr.', 'Con.']
        df['BIZ Type'] = df['BIZ Type'].replace(['COMM', 'comm', 'COMMERCIAL', 'commercial'], 'COMM').fillna('Unknown')
        # ... (기타 전처리) ...
        df['Year'] = pd.to_numeric(df['Year'], errors='coerce')
        df['Month'] = pd.to_numeric(df['Month'], errors='coerce')
        df = df.dropna(subset=['Year', 'Month'])
        df['Rev. (€)'] = pd.to_numeric(df['Rev. (€)'], errors='coerce').fillna(0)
        return df

    raw_df = load_and_preprocess(uploaded_file)
    selected_year = st.sidebar.selectbox("연도", sorted(raw_df['Year'].unique()))
    selected_month = st.sidebar.selectbox("월", sorted(raw_df['Month'].unique()))

    # 로직 및 출력 예시
    # 정렬: df = df.sort_values(by=(컬럼명, 'ACT'), ascending=False)
    # 출력: st.markdown(render_html_view(df), unsafe_allow_html=True)
    
    st.info("데이터 로드 완료. 위 정렬 로직을 각 보고서 섹션에 적용하세요.")
else:
    st.info("👈 좌측 사이드바에서 엑셀 파일을 업로드하세요.")
