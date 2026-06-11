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
    .table-container { overflow-x: auto; border: 2px solid #002060; box-shadow: 2px 2px 10px rgba(0,0,0,0.1); margin-bottom: 1rem !important; padding: 0px !important; display: inline-block; width: auto; min-width: 100%; box-sizing: border-box; }
    .report-table { border-collapse: separate !important; border-spacing: 0 !important; font-family: 'Malgun Gothic', sans-serif; font-size: 12px; width: 100%; margin: 0 !important; background-color: white; }
    .report-table thead th { background-color: #002060 !important; color: white !important; border: 1px solid #8ea9db !important; text-align: center !important; padding: 4px 3px !important; }
    .report-table td { border: 1px solid #d9d9d9; text-align: center; padding: 4px; vertical-align: middle; }
    </style>
""", unsafe_allow_html=True)

st.title("📊 통합 월간 매출 보고서 (FC vs ACT 자동 집계)")

# ==========================================
# 2. 유틸리티 함수 (전역 위치로 이동)
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

# 렌더링 엔진: 헤더 노란색 + 합계 행 음영 처리
def render_html_table(df):
    df_display = df.copy()
    
    # 1. 컬럼 헤더에 ACT 포함 시 노란색 처리
    new_cols = []
    for col in df_display.columns:
        col_str = str(col[1]) if isinstance(col, tuple) else str(col)
        if 'ACT' in col_str:
            new_col = col.replace('ACT', '<span style="color: #FFD700;">ACT</span>') if isinstance(col, str) else tuple(f'<span style="color: #FFD700;">{c}</span>' if 'ACT' in str(c) else c for c in col)
            new_cols.append(new_col)
        else:
            new_cols.append(col)
    df_display.columns = pd.MultiIndex.from_tuples(new_cols) if isinstance(df.columns, pd.MultiIndex) else new_cols
    df_display = df_display.replace(0, '')

    # 2. 스타일링
    format_dict = {col: format_percentage_html if 'ACHI' in str(col) else format_k_val for col in df_display.columns}
    styler = df_display.style.format(format_dict, na_rep='').set_table_attributes('class="report-table"')
    styler.set_properties(subset=get_numeric_cols(df_display), **{'text-align': 'right'})
    
    # 합계 행 음영 강제 적용
    styler.apply(lambda row: [
        'background-color: #ffffe0 !important; color: #002060 !important; font-weight: bold !important; border-top: 2px solid #8ea9db !important; border-bottom: 2px solid #8ea9db !important;' 
        if any(k in str(row.name) for k in ['TTL', 'Total', 'Subtotal', '소계']) else '' for _ in row
    ], axis=1)
    
    return f'<div class="table-container">{styler.to_html(escape=False)}</div>'

# [여기에 build_summary_report, get_biz_type_detailed_report, get_biz_report, to_excel_multiple, load_and_preprocess 함수들을 모두 넣으세요]

# ==========================================
# 3. 메인 로직
# ==========================================
uploaded_file = st.sidebar.file_uploader("SAP/엑셀 데이터를 업로드하세요.", type=['xlsx', 'xls'])

if uploaded_file:
    raw_df = load_and_preprocess(uploaded_file)
    selected_year = st.sidebar.selectbox("연도", sorted(raw_df['Year'].unique()))
    selected_month = st.sidebar.selectbox("월", sorted(raw_df['Month'].unique()))

    # 1. 매출 요약 (CPS) -> ACT 기준 내림차순 정렬
    df_cps, p_col, c_col = build_summary_report(raw_df, ['CPS'], selected_year, selected_month, 'TTL (K.€)')
    if not df_cps.empty:
        df_cps = df_cps.sort_values(by=(c_col, 'ACT'), ascending=False)
        st.subheader("📌 1. 매출 요약 (CPS 기준)")
        st.markdown(render_html_table(df_cps), unsafe_allow_html=True)

    # 2. 매출 요약 (Item) -> ACT 기준 내림차순 정렬
    df_item_raw = raw_df[raw_df['Item'].isin(['ICCU1', 'ICCU2', 'VCMS'])]
    df_item, p_col, c_col = build_summary_report(df_item_raw, ['Item'], selected_year, selected_month, 'TTL (K.€)')
    if not df_item.empty:
        df_item = df_item.sort_values(by=(c_col, 'ACT'), ascending=False)
        st.subheader("📌 2. 매출 요약 (Item 기준)")
        st.markdown(render_html_table(df_item), unsafe_allow_html=True)

    # 3. 비즈니스 타입별 매출 요약 -> 정렬
    df_biz = get_biz_type_detailed_report(raw_df, selected_year, selected_month)
    if not df_biz.empty:
        # 데이터 정렬 (필요시 기준 컬럼 수정)
        st.subheader("📌 3. 비즈니스 타입별 매출 요약")
        st.markdown(render_html_table(df_biz), unsafe_allow_html=True)

else:
    st.info("👈 좌측 사이드바에서 엑셀 파일을 업로드하세요.")
