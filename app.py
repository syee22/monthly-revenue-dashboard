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
    h1 { font-size: 1.6rem !important; margin-bottom: 0.5rem !important; }
    h3 { font-size: 1.1rem !important; margin-top: 1rem !important; color: #002060 !important; }
    .report-table { border-collapse: separate !important; border-spacing: 0 !important; font-family: 'Malgun Gothic', sans-serif; font-size: 12px; width: 100%; background-color: white; margin: 0 !important; }
    .report-table thead th { background-color: #002060 !important; color: white !important; border: 1px solid #8ea9db !important; text-align: center !important; padding: 4px 3px !important; }
    .report-table td { border: 1px solid #d9d9d9; text-align: center; padding: 4px; vertical-align: middle; }
    .table-container { overflow-x: auto; border: 2px solid #002060; box-shadow: 2px 2px 10px rgba(0,0,0,0.1); margin-bottom: 1rem !important; padding: 0px !important; display: inline-block; width: auto; min-width: 100%; box-sizing: border-box; }
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

# ==========================================
# 3. 렌더링 함수 (ACT 헤더 노란색 + 정렬 기능 반영)
# ==========================================
def render_html_view(df):
    # 컬럼 헤더에 ACT 포함 시 노란색 span 적용 (헤더 HTML 변환)
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
    
    # 소계/합계 행 스타일 (모든 셀 강제 적용으로 병합 현상 방지)
    styler.apply(lambda row: [
        'background-color: #ffffe0 !important; color: #002060 !important; font-weight: bold !important; border-top: 2px solid #8ea9db !important; border-bottom: 2px solid #8ea9db !important;' 
        if any(k in str(row.name) for k in ['TTL', 'Total', 'Subtotal', '소계']) else '' for _ in row
    ], axis=1)
    
    return f'<div class="table-container">{styler.to_html(escape=False)}</div>'

# ==========================================
# 4. 데이터 로드 및 로직 실행
# ==========================================
uploaded_file = st.sidebar.file_uploader("SAP/엑셀 데이터를 업로드하세요.", type=['xlsx', 'xls'])

if uploaded_file:
    # (기존 load_and_preprocess 로직 동일)
    @st.cache_data
    def load_and_preprocess(file):
        xl = pd.ExcelFile(file); sheets = xl.sheet_names
        df = pd.read_excel(xl, sheet_name=sheets[0], header=4).iloc[:, :26]
        df.columns = ['Year', 'Month', 'Desc.', 'Date', 'STP', 'Customer', 'LK No.', "Q'ty", 'Rev. ($)', 'Rev. (€)', 'Rev. ₩', 'BIZ Type', 'Group 1', 'Group 2', 'Project', 'PF', 'Item', 'Source', 'KOx', 'Memo', 'CPS', 'EUR:USD', 'EUR:KRW', 'Business Type', 'Curr.', 'Con.']
        df['BIZ Type'] = df['BIZ Type'].replace(['COMM', 'comm', 'COMMERCIAL', 'commercial'], 'COMM').fillna('Unknown')
        return df

    raw_df = load_and_preprocess(uploaded_file)
    selected_year = st.sidebar.selectbox("연도", sorted(raw_df['Year'].unique()))
    selected_month = st.sidebar.selectbox("월", sorted(raw_df['Month'].unique()))

    # 각 보고서 생성 및 정렬 적용 예시
    # 1. CPS 요약
    df_cps, p_col, c_col = build_summary_report(raw_df, ['CPS'], selected_year, selected_month, 'TTL (K.€)')
    df_cps = df_cps.sort_values(by=(c_col, 'ACT'), ascending=False)
    st.subheader("📌 1. 매출 요약 (CPS 기준)")
    st.markdown(render_html_view(df_cps), unsafe_allow_html=True)

    # 2. Item 요약
    df_item_raw = raw_df[raw_df['Item'].isin(['ICCU1', 'ICCU2', 'VCMS'])]
    df_item, p_col, c_col = build_summary_report(df_item_raw, ['Item'], selected_year, selected_month, 'TTL (K.€)')
    df_item = df_item.sort_values(by=(c_col, 'ACT'), ascending=False)
    st.subheader("📌 2. 매출 요약 (Item 기준)")
    st.markdown(render_html_view(df_item), unsafe_allow_html=True)

    # 3. 비즈니스 타입 요약
    df_biz = get_biz_type_detailed_report(raw_df, selected_year, selected_month)
    # 정렬: 특정 인덱스 기준 정렬이 필요할 경우 df_biz.sort_values(...) 추가
    st.subheader("📌 3. 비즈니스 타입별 매출 요약 (DIRECT / COMM.)")
    st.markdown(render_html_view(df_biz), unsafe_allow_html=True)
    
else:
    st.info("👈 좌측 사이드바에서 엑셀 파일을 업로드하세요.")
