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
    .report-table { border-collapse: collapse !important; font-family: 'Malgun Gothic', sans-serif; font-size: 12px; width: 100%; background-color: white; }
    .table-container { overflow-x: auto; border: 2px solid #002060; box-shadow: 2px 2px 10px rgba(0,0,0,0.1); margin-bottom: 1rem !important; padding: 0px !important; display: inline-block; width: auto; min-width: 100%; box-sizing: border-box; }
    .report-table td, .report-table th { border: 1px solid #d9d9d9; padding: 4px; text-align: center; }
    .report-table thead th { background-color: #002060 !important; color: white !important; }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 2. 유틸리티 함수
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

# ==========================================
# 3. 로직 함수 (전체)
# ==========================================
def load_and_preprocess(file):
    xl = pd.ExcelFile(file); sheets = xl.sheet_names
    df = pd.read_excel(xl, sheet_name=sheets[0], header=4).iloc[:, :26]
    df.columns = ['Year', 'Month', 'Desc.', 'Date', 'STP', 'Customer', 'LK No.', "Q'ty", 'Rev. ($)', 'Rev. (€)', 'Rev. ₩', 'BIZ Type', 'Group 1', 'Group 2', 'Project', 'PF', 'Item', 'Source', 'KOx', 'Memo', 'CPS', 'EUR:USD', 'EUR:KRW', 'Business Type', 'Curr.', 'Con.']
    df['BIZ Type'] = df['BIZ Type'].replace(['COMM', 'comm', 'COMMERCIAL', 'commercial'], 'COMM').fillna('Unknown')
    df['Year'] = pd.to_numeric(df['Year'], errors='coerce').fillna(0).astype(int)
    df['Month'] = pd.to_numeric(df['Month'], errors='coerce').fillna(0).astype(int)
    df['Rev. (€)'] = pd.to_numeric(df['Rev. (€)'], errors='coerce').fillna(0)
    return df

def build_summary_report(df_sub, index_cols, year, month, total_label):
    if df_sub.empty: return pd.DataFrame(), "", ""
    prev_year, prev_month = (year - 1, 12) if month == 1 else (year, month - 1)
    month_names = {1:'Jan', 2:'Feb', 3:'Mar', 4:'Apr', 5:'May', 6:'Jun', 7:'Jul', 8:'Aug', 9:'Sep', 10:'Oct', 11:'Nov', 12:'Dec'}
    col_curr = f'{month_names.get(month, f"{month}")}. {year}'
    phases = [col_curr, f'YTD {month_names.get(month, f"{month}")}. {year}', f'{year} TTL']
    
    def get_pivot(d): return d.pivot_table(index=index_cols, columns='Desc.', values='Rev. (€)', aggfunc='sum').fillna(0)
    p_curr = get_pivot(df_sub[(df_sub['Year'] == year) & (df_sub['Month'] == month)])
    p_ytd = get_pivot(df_sub[(df_sub['Year'] == year) & (df_sub['Month'] <= month)])
    p_ttl = get_pivot(df_sub[(df_sub['Year'] == year)])
    
    combined_dict = {}
    for phase, data in zip(phases, [p_curr, p_ytd, p_ttl]):
        for c in ['25 FC3', '26 FC1', 'ACT']: combined_dict[(phase, c)] = data.get(c, 0)
        combined_dict[(phase, 'ACHI %')] = data.get('ACT', 0) / data.get('26 FC1', 1)
    
    df_res = pd.DataFrame(combined_dict, index=p_curr.index if not p_curr.empty else p_ttl.index)
    return df_res, col_curr

def get_biz_type_detailed_report(df, year, month):
    month_names = {1:'Jan', 2:'Feb', 3:'Mar', 4:'Apr', 5:'May', 6:'Jun', 7:'Jul', 8:'Aug', 9:'Sep', 10:'Oct', 11:'Nov', 12:'Dec'}
    col_curr = f'{month_names.get(month, f"{month}")}. {year}'
    phases = [col_curr, f'YTD {month_names.get(month, f"{month}")}. {year}', f'{year} TTL']
    results = []
    for biz in ['DIRECT', 'COMM', 'Unknown']:
        biz_df = df[(df['BIZ Type'] == biz) & (df['Year'] == year)]
        if biz_df.empty: continue
        # 각 biz 내에서 당월 ACT 계산
        p_m = biz_df[biz_df['Month'] == month].pivot_table(index=['BIZ Type', 'KOx'], columns='Desc.', values='Rev. (€)', aggfunc='sum').fillna(0)
        p_y = biz_df[biz_df['Month'] <= month].pivot_table(index=['BIZ Type', 'KOx'], columns='Desc.', values='Rev. (€)', aggfunc='sum').fillna(0)
        p_fy = biz_df.pivot_table(index=['BIZ Type', 'KOx'], columns='Desc.', values='Rev. (€)', aggfunc='sum').fillna(0)
        
        combined_dict = {}
        for phase, data in zip(phases, [p_m, p_y, p_fy]):
            for c in ['25 FC3', '26 FC1', 'ACT']: combined_dict[(phase, c)] = data.get(c, 0)
            combined_dict[(phase, 'ACHI %')] = data.get('ACT', 0) / data.get('26 FC1', 1)
        
        df_biz = pd.DataFrame(combined_dict, index=p_m.index)
        # KOx 기준 당월 ACT 내림차순 정렬
        df_biz = df_biz.sort_values(by=(col_curr, 'ACT'), ascending=False)
        subtotal = pd.DataFrame([df_biz.sum()], index=pd.MultiIndex.from_tuples([(biz, 'Subtotal')], names=['BIZ Type', 'KOx']))
        results.append(pd.concat([df_biz, subtotal]))
    return pd.concat(results)

def get_biz_report(df, biz_type, year, month):
    df_biz = df[(df['Business Type'].str.contains(biz_type, case=False, na=False)) & (df['Year'] == year)].copy()
    results = []
    for brand in ['HYU', 'KIA', 'GM']:
        b_df = df_biz[df_biz['Group 2'] == brand]
        if b_df.empty: continue
        # 단순화된 피벗팅
        p_m = b_df[b_df['Month'] == month].pivot_table(index=['Project'], columns='Desc.', values='Rev. (€)', aggfunc='sum').fillna(0)
        # ... (중략) ...
        combined = p_m.copy() 
        combined.index = pd.MultiIndex.from_product([[brand], combined.index], names=['Cust. GR', 'Project'])
        results.append(combined)
    return pd.concat(results) if results else pd.DataFrame()

def render_html_table(df):
    # 포맷팅 적용
    format_dict = {col: format_percentage_html if 'ACHI' in str(col) else format_k_val for col in df.columns}
    styler = df.style.format(format_dict).set_table_attributes('class="report-table"')
    return f'<div class="table-container">{styler.to_html()}</div>'

# ==========================================
# 4. 메인 실행부
# ==========================================
st.title("📊 통합 월간 매출 보고서")
uploaded_file = st.sidebar.file_uploader("엑셀 파일 업로드", type=['xlsx', 'xls'])

if uploaded_file:
    raw_df = load_and_preprocess(uploaded_file)
    year = st.sidebar.selectbox("연도", sorted(raw_df['Year'].unique()))
    month = st.sidebar.selectbox("월", sorted(raw_df['Month'].unique()))

    # 1. CPS
    df_cps, col_curr = build_summary_report(raw_df, ['CPS'], year, month, 'Total')
    st.subheader("📌 1. CPS 매출")
    st.markdown(render_html_table(df_cps), unsafe_allow_html=True)

    # 2. Item (당월 ACT 기준 내림차순)
    df_item, col_curr = build_summary_report(raw_df[raw_df['Item'].isin(['ICCU1', 'ICCU2', 'VCMS'])], ['Item'], year, month, 'Total')
    df_item = df_item.sort_values(by=(col_curr, 'ACT'), ascending=False)
    st.subheader("📌 2. Item별 매출 (당월 ACT 정렬)")
    st.markdown(render_html_table(df_item), unsafe_allow_html=True)

    # 3. Biz Type (그룹 내 KOx 정렬)
    df_biz = get_biz_type_detailed_report(raw_df, year, month)
    st.subheader("📌 3. BIZ Type별 매출 (KOx별 정렬)")
    st.markdown(render_html_table(df_biz), unsafe_allow_html=True)

    # 4 & 5. 추가 비즈니스 보고서
    st.subheader("📌 4. Power Electronics 비즈니스")
    st.markdown(render_html_table(get_biz_report(raw_df, "Power", year, month)), unsafe_allow_html=True)
    
    st.subheader("📌 5. Core 비즈니스")
    st.markdown(render_html_table(get_biz_report(raw_df, "Core", year, month)), unsafe_allow_html=True)
