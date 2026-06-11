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
    .table-container { overflow-x: auto; border: 2px solid #002060; box-shadow: 2px 2px 10px rgba(0,0,0,0.1); margin-bottom: 1rem !important; padding: 0px !important; display: inline-block; width: auto; min-width: 100%; box-sizing: border-box; }
    .report-table { border-collapse: separate !important; border-spacing: 0 !important; font-family: 'Malgun Gothic', sans-serif; font-size: 12px; width: 100%; margin: 0 !important; background-color: white; }
    .report-table thead th { background-color: #002060 !important; color: white !important; border: 1px solid #8ea9db !important; text-align: center !important; padding: 4px 3px !important; }
    .report-table td { border: 1px solid #d9d9d9; text-align: center; padding: 4px; vertical-align: middle; }
    </style>
""", unsafe_allow_html=True)

st.title("📊 통합 월간 매출 보고서 (FC vs ACT 자동 집계)")

# ==========================================
# 2. 필수 함수 (최상위 정의)
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

# [기존 load_and_preprocess, build_summary_report, get_biz_type_detailed_report, get_biz_report, to_excel_multiple 함수들을 여기에 그대로 두시면 됩니다]
# (이하 생략된 기존 함수들의 로직이 그대로 들어가야 합니다)

def render_html_view(df):
    # 컬럼명 헤더에서 ACT만 노란색 처리
    new_cols = []
    for col in df.columns:
        c_name = str(col[1] if isinstance(col, tuple) else col)
        if 'ACT' in c_name:
            new_col = (col[0], col[1].replace('ACT', '<span style="color: #FFD700;">ACT</span>')) if isinstance(col, tuple) else f'<span style="color: #FFD700;">{col}</span>'
            new_cols.append(new_col)
        else: new_cols.append(col)
    
    df_d = df.copy()
    df_d.columns = pd.MultiIndex.from_tuples(new_cols) if isinstance(df.columns, pd.MultiIndex) else new_cols
    df_d = df_d.replace(0, '')
    
    format_dict = {col: format_percentage_html if 'ACHI' in str(col) else format_k_val for col in df_d.columns}
    styler = df_d.style.format(format_dict, na_rep='').set_table_attributes('class="report-table"')
    styler.apply(lambda row: ['background-color: #ffffe0 !important; color: #002060 !important; font-weight: bold !important; border-top: 2px solid #8ea9db !important; border-bottom: 2px solid #8ea9db !important;'] * len(row) 
                 if any(k in str(row.name) for k in ['TTL', 'Total', 'Subtotal', '소계']) else [''] * len(row), axis=1)
    return f'<div class="table-container">{styler.to_html(escape=False)}</div>'

# ==========================================
# 3. 메인 실행 로직
# ==========================================
uploaded_file = st.sidebar.file_uploader("SAP/엑셀 데이터를 업로드하세요.", type=['xlsx', 'xls'])

if uploaded_file:
    raw_df = load_and_preprocess(uploaded_file)
    selected_year = st.sidebar.selectbox("연도", sorted(raw_df['Year'].unique()))
    selected_month = st.sidebar.selectbox("월", sorted(raw_df['Month'].unique()))

    # 1. 매출 요약 (CPS)
    df_cps, p_col, c_col = build_summary_report(raw_df, ['CPS'], selected_year, selected_month, 'TTL (K.€)')
    if not df_cps.empty:
        df_cps = df_cps.sort_values(by=(c_col, 'ACT'), ascending=False)
        st.subheader("📌 1. 매출 요약 (CPS 기준)")
        st.markdown(render_html_view(df_cps), unsafe_allow_html=True)

    # 2. 매출 요약 (Item)
    df_item_raw = raw_df[raw_df['Item'].isin(['ICCU1', 'ICCU2', 'VCMS'])]
    df_item, p_col, c_col = build_summary_report(df_item_raw, ['Item'], selected_year, selected_month, 'TTL (K.€)')
    if not df_item.empty:
        df_item = df_item.sort_values(by=(c_col, 'ACT'), ascending=False)
        st.subheader("📌 2. 매출 요약 (Item 기준)")
        st.markdown(render_html_view(df_item), unsafe_allow_html=True)

    # 3. 비즈니스 타입 요약
    df_biz = get_biz_type_detailed_report(raw_df, selected_year, selected_month)
    if not df_biz.empty:
        st.subheader("📌 3. 비즈니스 타입별 매출 요약")
        st.markdown(render_html_view(df_biz), unsafe_allow_html=True)

    # 6. HKMC 신규 추가
    st.subheader("📌 6. HKMC 매출 요약 (KEM-KR/CN)")
    mask = (raw_df['Group 1'] == 'HKMC') & (raw_df['KOx'].isin(['KEM-KR', 'KEM-CN']))
    df_hkmc, p_col, c_col = build_summary_report(raw_df[mask], ['KOx'], selected_year, selected_month, 'TTL (K.€)')
    if not df_hkmc.empty:
        df_hkmc = df_hkmc.sort_values(by=(c_col, 'ACT'), ascending=False)
        st.markdown(render_html_view(df_hkmc), unsafe_allow_html=True)

else:
    st.info("파일을 업로드하세요.")
