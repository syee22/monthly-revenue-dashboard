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
    .report-table { border-collapse: collapse !important; font-family: 'Malgun Gothic', sans-serif; font-size: 12px; width: 100%; background-color: white; }
    .report-table thead th { background-color: #002060 !important; color: white !important; }
    .report-table td { border: 1px solid #d9d9d9; text-align: center; padding: 4px; }
    .table-container { overflow-x: auto; border: 2px solid #002060; margin-bottom: 1rem; }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 2. 모든 함수 정의 (최상단 배치)
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

def to_excel_multiple(df_dict):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet_name, df in df_dict.items():
            styler = df.style.format(lambda x: format_k_val(x) if isinstance(x, (int, float)) else x)
            styler.apply(lambda row: ['background-color: #ffffe0; color: #002060; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;'] * len(row) if any(k in str(row.name) for k in ['TTL', 'Total', 'Subtotal', '소계']) else [''] * len(row), axis=1)
            styler.to_excel(writer, sheet_name=sheet_name[:31])
    return output.getvalue()

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
    # 사용자님의 원본 build_summary_report 로직을 그대로 사용 (복사)
    # 생략된 부분은 기존 코드의 해당 함수 내용을 여기에 붙여넣으세요.
    return pd.DataFrame(), "", ""

def get_biz_type_detailed_report(df, year, month):
    # 사용자님의 원본 get_biz_type_detailed_report 로직 그대로 사용 (복사)
    return pd.DataFrame()

def get_biz_report(df, biz_type, year, month):
    # 사용자님의 원본 get_biz_report 로직 그대로 사용 (복사)
    return pd.DataFrame(), []

def render_html_view(df):
    styler = df.replace(0, '').style.format({col: format_percentage_html if 'ACHI' in str(col) else format_k_val for col in df.columns}).set_table_attributes('class="report-table"')
    styler.apply(lambda row: ['background-color: #ffffe0; color: #002060; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;'] * len(row) if any(k in str(row.name) for k in ['TTL', 'Total', 'Subtotal', '소계']) else [''] * len(row), axis=1)
    return f'<div class="table-container">{styler.to_html()}</div>'

def render_biz_html_table(df):
    styler = df.replace(0, '').style.format({col: format_percentage_html if 'ACHI' in str(col) else format_k_val for col in df.columns}).set_table_attributes('class="report-table"')
    styler.apply(lambda row: ['background-color: #ffffe0; color: #002060; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;' if '소계' in str(row.name) or 'Total' in str(row.name) else '' for _ in row], axis=1)
    return f'<div class="table-container">{styler.to_html()}</div>'

# ==========================================
# 4. 메인 실행부
# ==========================================
st.title("📊 통합 월간 매출 보고서 (FC vs ACT 자동 집계)")
uploaded_file = st.sidebar.file_uploader("SAP/엑셀 데이터를 업로드하세요.", type=['xlsx', 'xls'])

if uploaded_file:
    raw_df = load_and_preprocess(uploaded_file)
    selected_year = st.sidebar.selectbox("연도", sorted(raw_df['Year'].unique()))
    selected_month = st.sidebar.selectbox("월", sorted(raw_df['Month'].unique()))
    reports_to_download = {}

    # 1. 매출 요약 (CPS)
    st.subheader("📌 1. 매출 요약 (CPS 기준)")
    df_cps, p_col, c_col = build_summary_report(raw_df, ['CPS'], selected_year, selected_month, 'TTL (K.€)')
    if not df_cps.empty: st.markdown(render_html_view(df_cps), unsafe_allow_html=True); reports_to_download["CPS_Summary"] = df_cps

    # 3. 비즈니스 타입별 매출 요약 (순서 변경)
    st.subheader("📌 3. 비즈니스 타입별 매출 요약 (DIRECT / COMM.)")
    df_biz = get_biz_type_detailed_report(raw_df, selected_year, selected_month)
    if not df_biz.empty: st.markdown(render_html_view(df_biz), unsafe_allow_html=True); reports_to_download["Biz_Type_Summary"] = df_biz

    # 5. Core 비즈니스 (순서 변경)
    st.subheader("📌 5. Core 비즈니스")
    df_core, phase_names_core = get_biz_report(raw_df, "Core", selected_year, selected_month)
    if not df_core.empty: st.markdown(render_biz_html_table(df_core), unsafe_allow_html=True); reports_to_download["Core_Biz"] = df_core

    # 4. Power Electronics 비즈니스 (순서 변경)
    st.subheader("📌 4. Power Electronics 비즈니스")
    df_pe, phase_names_pe = get_biz_report(raw_df, "Power", selected_year, selected_month)
    if not df_pe.empty: st.markdown(render_biz_html_table(df_pe), unsafe_allow_html=True); reports_to_download["PE_Biz"] = df_pe

    # 2. 매출 요약 (Item) (순서 변경)
    st.subheader("📌 2. 매출 요약 (Item 기준)")
    df_item_raw = raw_df[raw_df['Item'].isin(['ICCU1', 'ICCU2', 'VCMS'])]
    df_item, p_col, c_col = build_summary_report(df_item_raw, ['Item'], selected_year, selected_month, 'TTL (K.€)')
    if not df_item.empty: st.markdown(render_html_view(df_item), unsafe_allow_html=True); reports_to_download["Item_Summary"] = df_item

    if reports_to_download:
        st.write("---")
        st.download_button("📥 전체 요약 리포트 다운로드", data=to_excel_multiple(reports_to_download), file_name="Report.xlsx")
