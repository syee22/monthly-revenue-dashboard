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
    .table-container { overflow-x: auto; border: 2px solid #002060; box-shadow: 2px 2px 10px rgba(0,0,0,0.1); margin-bottom: 1rem !important; padding: 0px !important; display: inline-block; width: auto; min-width: 100%; box-sizing: border-box; }
    .report-table { border-collapse: collapse !important; font-family: 'Malgun Gothic', sans-serif; font-size: 12px; width: 100%; margin: 0 !important; background-color: white; }
    .report-table thead th { background-color: #002060 !important; color: white !important; border: 1px solid #8ea9db !important; text-align: center !important; padding: 4px 3px !important; }
    .report-table td { border: 1px solid #d9d9d9; text-align: center; padding: 4px; vertical-align: middle; }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 2. 필수 함수 정의 (모두 최상단 배치)
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

def build_summary_report(df_sub, index_cols, year, month, total_label, index_names=None):
    if df_sub.empty: return pd.DataFrame(), "", ""
    if month == 1: prev_year, prev_month = year - 1, 12
    else: prev_year, prev_month = year, month - 1
    month_names = {1:'Jan', 2:'Feb', 3:'Mar', 4:'Apr', 5:'May', 6:'Jun', 7:'Jul', 8:'Aug', 9:'Sep', 10:'Oct', 11:'Nov', 12:'Dec'}
    m_str, pm_str = month_names.get(month, f'{month}'), month_names.get(prev_month, f'{prev_month}')
    col_prev, phase_curr, phase_ytd, phase_ttl = f'{pm_str}. {year if month != 1 else prev_year}', f'{m_str}. {year}', f'YTD {m_str}. {year}', f'{year} TTL'
    phases = [phase_curr, phase_ytd, phase_ttl]
    
    def get_pivot(d): return d.pivot_table(index=index_cols, columns='Desc.', values='Rev. (€)', aggfunc='sum').fillna(0) if not d.empty else pd.DataFrame()
    s_prev = df_sub[(df_sub['Year'] == prev_year) & (df_sub['Month'] == prev_month) & (df_sub['Desc.'] == 'ACT')].groupby(index_cols)['Rev. (€)'].sum()
    p_curr, p_ytd, p_ttl = get_pivot(df_sub[(df_sub['Year'] == year) & (df_sub['Month'] == month)]), get_pivot(df_sub[(df_sub['Year'] == year) & (df_sub['Month'] <= month)]), get_pivot(df_sub[(df_sub['Year'] == year)])
    
    all_indices = set()
    for p in [s_prev, p_curr, p_ytd, p_ttl]:
        if not p.empty: all_indices.update(p.index.tolist() if isinstance(p.index, pd.MultiIndex) else [(x,) for x in p.index.tolist()])
    if not all_indices: return pd.DataFrame(), col_prev, phase_curr
    all_indices = sorted(list(all_indices), key=lambda x: tuple(str(i) for i in x))
    current_index_names = index_names if index_names else (['CPS'] if index_cols == ['CPS'] else index_cols)
    idx = pd.MultiIndex.from_tuples(all_indices, names=current_index_names) if len(index_cols) > 1 else pd.Index([x[0] for x in all_indices], name=current_index_names[0])
    combined_dict, col_tuples = {}, [('', col_prev)]
    for p in phases:
        for c in ['25 FC3', '26 FC1', 'ACT', 'ACHI %']: col_tuples.append((p, c))
    combined_dict[('', col_prev)] = s_prev.reindex(idx).fillna(0) if not s_prev.empty else pd.Series(0, index=idx)
    for phase_name, data in zip(phases, [p_curr, p_ytd, p_ttl]):
        for c in ['25 FC3', '26 FC1', 'ACT']: combined_dict[(phase_name, c)] = data[c].reindex(idx).fillna(0) if not data.empty and c in data.columns else pd.Series(0, index=idx)
        num = pd.Series(combined_dict[(phase_name, 'ACT')]); den = pd.Series(combined_dict[(phase_name, '26 FC1')])
        combined_dict[(phase_name, 'ACHI %')] = num.div(den).replace([np.inf, -np.inf], 0).fillna(0)
    final_df = pd.DataFrame(combined_dict); final_df.columns = pd.MultiIndex.from_tuples(col_tuples); final_df.index.names = current_index_names
    total_row = final_df.sum(numeric_only=True)
    for phase_name in phases:
        num = total_row.get((phase_name, 'ACT'), 0); den = total_row.get((phase_name, '26 FC1'), 0)
        total_row[(phase_name, 'ACHI %')] = num / den if den != 0 else 0
    t_label = "TTL (K.€)"
    t_index = tuple([''] * (len(final_df.index.names)-1) + [t_label]) if isinstance(final_df.index, pd.MultiIndex) else t_label
    t_df = pd.DataFrame([total_row], index=[t_index] if not isinstance(final_df.index, pd.MultiIndex) else pd.MultiIndex.from_tuples([t_index], names=final_df.index.names))
    return pd.concat([final_df, t_df]), col_prev, phase_curr

def get_biz_type_detailed_report(df, year, month):
    if month == 1: prev_year, prev_month = year - 1, 12
    else: prev_year, prev_month = year, month - 1
    month_names = {1:'Jan', 2:'Feb', 3:'Mar', 4:'Apr', 5:'May', 6:'Jun', 7:'Jul', 8:'Aug', 9:'Sep', 10:'Oct', 11:'Nov', 12:'Dec'}
    m_str, pm_str = month_names.get(month, f'{month}'), month_names.get(prev_month, f'{prev_month}')
    curr_phase = f'{m_str}. {year}'
    phase_names = [curr_phase, f'YTD {m_str}. {year}', f'{year} TTL']
    prev_phase_name = f'{pm_str}. {prev_year}'
    results = []
    biz_categories = ['DIRECT', 'COMM', 'Unknown']
    for biz in biz_categories:
        biz_df = df[(df['BIZ Type'] == biz) & (df['Year'] == year)]
        if biz_df.empty: continue
        p_m = biz_df[biz_df['Month'] == month].pivot_table(index=['BIZ Type', 'KOx'], columns='Desc.', values='Rev. (€)', aggfunc='sum').fillna(0)
        p_y = biz_df[biz_df['Month'] <= month].pivot_table(index=['BIZ Type', 'KOx'], columns='Desc.', values='Rev. (€)', aggfunc='sum').fillna(0)
        p_fy = biz_df.pivot_table(index=['BIZ Type', 'KOx'], columns='Desc.', values='Rev. (€)', aggfunc='sum').fillna(0)
        p_prev = df[(df['BIZ Type'] == biz) & (df['Year'] == prev_year) & (df['Month'] == prev_month)].pivot_table(index=['BIZ Type', 'KOx'], columns='Desc.', values='Rev. (€)', aggfunc='sum').fillna(0)
        combined_dict = {(prev_phase_name, 'ACT'): p_prev.get('ACT', 0)}
        for phase_name, data in [(phase_names[0], p_m), (phase_names[1], p_y), (phase_names[2], p_fy)]:
            for c in ['25 FC3', '26 FC1', 'ACT']: combined_dict[(phase_name, c)] = data.get(c, 0)
            num = pd.Series(data.get('ACT', 0)); den = pd.Series(data.get('26 FC1', 0))
            combined_dict[(phase_name, 'ACHI %')] = num.div(den).replace([np.inf, -np.inf], 0).fillna(0)
        combined = pd.DataFrame(combined_dict, index=p_m.index)
        # 소계 행 분리 후 정렬
        data_rows = combined.iloc[:-1] if len(combined) > 1 else combined
        sorted_data = data_rows.sort_values(by=(curr_phase, 'ACT'), ascending=False)
        subtotal = combined.sum(numeric_only=True)
        for p_name in phase_names:
            num = subtotal.get((p_name, 'ACT'), 0); den = subtotal.get((p_name, '26 FC1'), 0)
            subtotal[(p_name, 'ACHI %')] = num / den if den != 0 else 0
        results.append(sorted_data)
        results.append(pd.DataFrame([subtotal], index=pd.MultiIndex.from_tuples([(biz, 'Subtotal')], names=['BIZ Type', 'KOx'])))
    return pd.concat(results)

def get_biz_report(df, biz_type, year, month):
    # (사용자님의 기존 get_biz_report 로직을 여기에 그대로 붙여넣으세요)
    # 로직이 생략되었으나 전체 코드 구동을 위해 반드시 포함 필요
    return pd.DataFrame(), []

def to_excel_multiple(df_dict):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet_name, df in df_dict.items():
            df.to_excel(writer, sheet_name=sheet_name[:31])
    return output.getvalue()

def render_html_table(df):
    df_d = df.replace(0, '')
    # 헤더 ACT 노란색 변환
    new_cols = []
    for col in df_d.columns:
        c_name = str(col[1] if isinstance(col, tuple) else col)
        new_col = (col[0], col[1].replace('ACT', '<span style="color: #FFD700;">ACT</span>')) if isinstance(col, tuple) else f'<span style="color: #FFD700;">{col}</span>' if 'ACT' in c_name else col
        new_cols.append(new_col)
    df_d.columns = pd.MultiIndex.from_tuples(new_cols) if isinstance(df.columns, pd.MultiIndex) else new_cols
    
    format_dict = {col: format_percentage_html if 'ACHI' in str(col) else format_k_val for col in df_d.columns}
    styler = df_d.style.format(format_dict, na_rep='').set_table_attributes('class="report-table"')
    styler.apply(lambda row: ['background-color: #ffffe0 !important; color: #002060 !important; font-weight: bold !important; border-top: 2px solid #8ea9db !important; border-bottom: 2px solid #8ea9db !important;'] * len(row) 
                 if any(k in str(row.name) for k in ['TTL', 'Total', 'Subtotal', '소계']) else [''] * len(row), axis=1)
    return f'<div class="table-container">{styler.to_html(escape=False)}</div>'

# ==========================================
# 4. 메인 실행 로직
# ==========================================
st.title("📊 통합 월간 매출 보고서 (FC vs ACT 자동 집계)")
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
        st.markdown(render_html_table(df_cps), unsafe_allow_html=True)

    # 2. 매출 요약 (Item)
    df_item, p_col, c_col = build_summary_report(raw_df[raw_df['Item'].isin(['ICCU1', 'ICCU2', 'VCMS'])], ['Item'], selected_year, selected_month, 'TTL (K.€)')
    if not df_item.empty:
        df_item = df_item.sort_values(by=(c_col, 'ACT'), ascending=False)
        st.subheader("📌 2. 매출 요약 (Item 기준)")
        st.markdown(render_html_table(df_item), unsafe_allow_html=True)

    # 3. 비즈니스 타입 요약
    df_biz = get_biz_type_detailed_report(raw_df, selected_year, selected_month)
    if not df_biz.empty:
        st.subheader("📌 3. 비즈니스 타입별 매출 요약")
        st.markdown(render_html_table(df_biz), unsafe_allow_html=True)

    # 6. HKMC 요약
    st.subheader("📌 6. HKMC 매출 요약 (KEM-KR/CN)")
    mask = (raw_df['Group 1'] == 'HKMC') & (raw_df['KOx'].isin(['KEM-KR', 'KEM-CN']))
    df_hkmc, p_col, c_col = build_summary_report(raw_df[mask], ['KOx'], selected_year, selected_month, 'TTL (K.€)')
    if not df_hkmc.empty:
        df_hkmc = df_hkmc.sort_values(by=(c_col, 'ACT'), ascending=False)
        st.markdown(render_html_table(df_hkmc), unsafe_allow_html=True)
else:
    st.info("👈 좌측 사이드바에서 엑셀 파일을 업로드하세요.")
