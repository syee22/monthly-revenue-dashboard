import streamlit as st
import pandas as pd
import numpy as np
import io
import re

# ==========================================
# 1. 페이지 설정 및 CSS
# ==========================================
st.set_page_config(page_title="월간 매출 보고서", layout="wide")
st.markdown("""
    <style>
    .block-container { padding: 2rem 3rem; }
    .table-container { overflow-x: auto; border: 2px solid #002060; box-shadow: 2px 2px 10px rgba(0,0,0,0.1); margin-bottom: 1rem !important; padding: 0px !important; display: inline-block; width: auto; min-width: 100%; }
    .report-table { border-collapse: separate !important; border-spacing: 0 !important; font-family: 'Malgun Gothic', sans-serif; font-size: 12px; width: 100%; margin: 0 !important; }
    .report-table thead th { background-color: #002060 !important; color: white !important; border: 1px solid #8ea9db !important; text-align: center !important; padding: 4px 3px !important; }
    .report-table td { border: 1px solid #d9d9d9; text-align: center; padding: 4px; vertical-align: middle; }
    </style>
""", unsafe_allow_html=True)

st.title("📊 통합 월간 매출 보고서 (FC vs ACT 자동 집계)")

# ==========================================
# 2. 필수 함수들 (최상단 정의)
# ==========================================
def format_k_val(val):
    if pd.isna(val) or isinstance(val, str) or val == '': return val
    v = val / 1_000.0
    rounded_int = int(round(v, 0))
    if rounded_int == 0: return str(round(v, 2)) if round(v, 2) != 0 else "0"
    return f"{rounded_int:,}"

def get_numeric_cols(df):
    return [col for col in df.columns if any(x in str(col) for x in ['FC3', 'FC1', 'ACT', 'ACHI'])]

def render_html_view(df):
    # 컬럼 헤더 ACT 노란색 적용
    new_cols = []
    for col in df.columns:
        if isinstance(col, tuple):
            new_cols.append(tuple(f'<span style="color: #FFD700;">{c}</span>' if 'ACT' in str(c) else c for c in col))
        else:
            new_cols.append(f'<span style="color: #FFD700;">{col}</span>' if 'ACT' in str(col) else col)
    
    df_display = df.copy()
    df_display.columns = pd.MultiIndex.from_tuples(new_cols) if isinstance(df.columns, pd.MultiIndex) else new_cols
    df_display = df_display.replace(0, '')
    
    styler = df_display.style.format(lambda x: format_percentage_html(x) if 'ACHI' in str(x) else format_k_val(x) if isinstance(x, (int, float)) else x)
    styler.set_table_attributes('class="report-table"')
    styler.set_properties(subset=get_numeric_cols(df_display), **{'text-align': 'right'})
    
    # 소계/합계 행 강제 노란색 스타일링
    styler.apply(lambda row: [
        'background-color: #ffffe0 !important; color: #002060 !important; font-weight: bold !important; border-top: 2px solid #8ea9db !important; border-bottom: 2px solid #8ea9db !important;' 
        if any(k in str(row.name) for k in ['TTL', 'Total', 'Subtotal', '소계']) else '' for _ in row
    ], axis=1)
    
    return f'<div class="table-container">{styler.to_html(escape=False)}</div>'

# [여기에 build_summary_report, get_biz_type_detailed_report, get_biz_report 함수를 그대로 붙여넣으세요]
# (이 함수들이 여기에 있어야 NameError가 발생하지 않습니다)

# ==========================================
# 3. 메인 로직
# ==========================================
uploaded_file = st.sidebar.file_uploader("SAP/엑셀 데이터를 업로드하세요.", type=['xlsx', 'xls'])

if uploaded_file:
    raw_df = load_and_preprocess(uploaded_file)
    selected_year = st.sidebar.selectbox("연도", sorted(raw_df['Year'].unique()))
    selected_month = st.sidebar.selectbox("월", sorted(raw_df['Month'].unique()))

    # 1. 매출 요약 (CPS) -> ACT 정렬
    df_cps, p_col, c_col = build_summary_report(raw_df, ['CPS'], selected_year, selected_month, 'TTL (K.€)')
    if not df_cps.empty:
        df_cps = df_cps.sort_values(by=(c_col, 'ACT'), ascending=False)
        st.subheader("📌 1. 매출 요약 (CPS 기준)")
        st.markdown(render_html_view(df_cps), unsafe_allow_html=True)

    # 2. 매출 요약 (Item) -> ACT 정렬
    df_item_raw = raw_df[raw_df['Item'].isin(['ICCU1', 'ICCU2', 'VCMS'])]
    df_item, p_col, c_col = build_summary_report(df_item_raw, ['Item'], selected_year, selected_month, 'TTL (K.€)')
    if not df_item.empty:
        df_item = df_item.sort_values(by=(c_col, 'ACT'), ascending=False)
        st.subheader("📌 2. 매출 요약 (Item 기준)")
        st.markdown(render_html_view(df_item), unsafe_allow_html=True)

    # 3. 비즈니스 타입 요약 -> ACT 정렬
    df_biz = get_biz_type_detailed_report(raw_df, selected_year, selected_month)
    if not df_biz.empty:
        # 여기는 Biz Type, KOx 구조에 맞춰 정렬 (상위 레벨 기준으로 ACT 정렬)
        df_biz = df_biz.sort_values(by=[(f'{selected_month}. {selected_year}', 'ACT')], ascending=False)
        st.subheader("📌 3. 비즈니스 타입별 매출 요약")
        st.markdown(render_html_view(df_biz), unsafe_allow_html=True)

else:
    st.info("파일을 업로드하세요.")
