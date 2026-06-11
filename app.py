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
    .report-table { border-collapse: collapse; font-family: 'Malgun Gothic', sans-serif; font-size: 13px; width: 100%; background-color: white; }
    .report-table th { background-color: #002060 !important; color: white !important; border: 1px solid #8ea9db !important; text-align: center !important; padding: 6px 4px !important; font-weight: 600 !important; }
    .report-table td { border: 1px solid #d9d9d9; text-align: right; padding: 5px; vertical-align: middle; }
    .report-table th.row_heading, .report-table td.row_heading { text-align: center !important; border-right: 1px solid #8ea9db; }
    .table-container { overflow-x: auto; max-height: 750px; border: 2px solid #002060; box-shadow: 2px 2px 10px rgba(0,0,0,0.1); }
    .report-table thead th { position: sticky; top: 0; z-index: 10; }
    </style>
""", unsafe_allow_html=True)

st.title("📊 월간 매출 보고서 (Core & Power Electronics)")

# ==========================================
# 2. 포맷팅 및 다운로드 함수
# ==========================================
def format_k_val(val):
    if pd.isna(val) or isinstance(val, str) or val == '': return val
    v = val / 1_000.0
    rounded_int = int(round(v, 0))
    if rounded_int == 0:
        v_rounded = round(v, 2)
        if v_rounded == 0: return "0"
        return f"{v_rounded:,.1f}" if round(v_rounded, 1) == v_rounded else f"{v_rounded:,.2f}"
    return f"{rounded_int:,}"

def format_percentage_html(val):
    if pd.isna(val) or isinstance(val, str) or val == '': return val
    pct_str = f"{val:.0%}"
    if val >= 1.0: return f'<span style="color: #00b050; font-weight: bold;">{pct_str} ▲</span>'
    elif val > 0: return f'<span style="color: #c00000; font-weight: bold;">{pct_str} ▼</span>'
    else: return pct_str

def to_excel(df):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df_formatted = df.copy().fillna(0)
    for col in df_formatted.columns:
        if 'ACHI' not in col[1]:
            df_formatted[col] = df_formatted[col].apply(lambda x: x/1000 if isinstance(x, (int, float, np.number)) else x)
    df_formatted.to_excel(writer, index=True, sheet_name='Sheet1')
    writer.close()
    return output.getvalue()

# ==========================================
# 3. 데이터 전처리
# ==========================================
uploaded_file = st.sidebar.file_uploader("엑셀 파일 업로드", type=['xlsx', 'xls'])

if uploaded_file:
    @st.cache_data
    def load_and_preprocess(file):
        df = pd.read_excel(file, header=4).iloc[:, :26]
        df.columns = ['Year', 'Month', 'Desc.', 'Date', 'STP', 'Customer', 'LK No.', "Q'ty", 'Rev. ($)', 'Rev. (€)', 'Rev. (₩)', 'BIZ Type', 'Group 1', 'Group 2', 'Project', 'PF', 'Item', 'Source', 'KOx', 'Memo', 'CPS', 'EUR:USD', 'EUR:KRW', 'Business Type', 'Curr.', 'Con.']
        df['Year'] = pd.to_numeric(df['Year'], errors='coerce')
        df['Month'] = pd.to_numeric(df['Month'], errors='coerce')
        df = df.dropna(subset=['Year', 'Month'])
        df['Rev. (€)'] = pd.to_numeric(df['Rev. (€)'], errors='coerce').fillna(0)
        df.loc[(df['Item'] == 'VCMS') & (df['Source'] == 'KEM-KR'), 'Business Type'] = 'Power electronics'
        df.loc[(df['Item'] == 'VCMS') & (df['Source'] == 'KOASIA'), 'Business Type'] = 'Core Business'
        df.loc[df['Group 1'] == 'GM', 'Group 2'] = 'GM'
        return df.astype({'Year': int, 'Month': int})

    raw_df = load_and_preprocess(uploaded_file)
    selected_year = st.sidebar.selectbox("연도", sorted(raw_df['Year'].unique()))
    selected_month = st.sidebar.selectbox("월", sorted(raw_df['Month'].unique()))

    # ==========================================
    # 4. 비즈니스 로직 및 렌더링
    # ==========================================
    def get_biz_report(df, biz_type, year, month):
        prev_year, prev_month = (year - 1, 12) if month == 1 else (year, month - 1)
        df_biz = df[(df['Business Type'].str.contains(biz_type, case=False, na=False)) & (df['Year'] == year)].copy()
        
        m_names = {1:'Jan', 2:'Feb', 3:'Mar', 4:'Apr', 5:'May', 6:'Jun', 7:'Jul', 8:'Aug', 9:'Sep', 10:'Oct', 11:'Nov', 12:'Dec'}
        p_names = [f'{m_names[month]}. {year}', f'YTD {m_names[month]}. {year}', f'{year} TTL']
        
        results = []
        for brand in ['HYU', 'KIA', 'GM']:
            brand_df = df_biz[df_biz['Group 2'] == brand].copy()
            p_m = brand_df[brand_df['Month'] == month].pivot_table(index=['Project', 'Con.'], columns='Desc.', values='Rev. (€)', aggfunc='sum').fillna(0)
            p_y = brand_df[brand_df['Month'] <= month].pivot_table(index=['Project', 'Con.'], columns='Desc.', values='Rev. (€)', aggfunc='sum').fillna(0)
            p_fy = brand_df.pivot_table(index=['Project', 'Con.'], columns='Desc.', values='Rev. (€)', aggfunc='sum').fillna(0)
            
            p_prev = df[(df['Year']==prev_year)&(df['Month']==prev_month)&(df['Group 2']==brand)].pivot_table(index=['Project', 'Con.'], columns='Desc.', values='Rev. (€)', aggfunc='sum').fillna(0)

            if biz_type == "Core" and brand in ['HYU', 'KIA']:
                top = p_m.nlargest(10, 'ACT' if 'ACT' in p_m.columns else p_m.columns[0]).index
                def group_others(p):
                    main = p.loc[p.index.isin(top)]
                    oth = p.loc[~p.index.isin(top)].sum().to_frame().T
                    oth.index = pd.MultiIndex.from_tuples([('Others', '')], names=['Project', 'Con.'])
                    return pd.concat([main, oth])
                p_m, p_y, p_fy, p_prev = group_others(p_m), group_others(p_y), group_others(p_fy), group_others(p_prev)

            # 데이터 통합 및 합계 로직
            combined = pd.DataFrame({(f'{m_names[prev_month]}. {prev_year}', 'ACT'): p_prev.get('ACT', 0)})
            for p_name, data in [(p_names[0], p_m), (p_names[1], p_y), (p_names[2], p_fy)]:
                for c in ['25 FC3', '26 FC1', 'ACT']: combined[(p_name, c)] = data.get(c, 0)
                combined[(p_name, 'ACHI %')] = np.where(combined[(p_name, '26 FC1')] != 0, combined[(p_name, 'ACT')] / combined[(p_name, '26 FC1')], 0)
            
            combined = combined.sort_values(by=(p_names[0], 'ACT'), ascending=False)
            subtotal = combined.sum(numeric_only=True)
            for p in p_names: subtotal[(p, 'ACHI %')] = subtotal[(p, 'ACT')] / subtotal[(p, '26 FC1')] if subtotal[(p, '26 FC1')] != 0 else 0
            
            combined.index = pd.MultiIndex.from_tuples([(brand, p, c) for p, c in combined.index], names=['Cust. GR', 'Project', 'Con.'])
            results.append(combined)
            if brand != 'GM': results.append(pd.DataFrame([subtotal], index=pd.MultiIndex.from_tuples([(brand, '소계', '')], names=['Cust. GR', 'Project', 'Con.'])))
            
        final_df = pd.concat(results)
        grand_total = final_df[final_df.index.get_level_values(1) != '소계'].sum(numeric_only=True)
        final_df = pd.concat([final_df, pd.DataFrame([grand_total], index=pd.MultiIndex.from_tuples([('', f"{biz_type} Business Sales Revenue (K.€)", '')], names=['Cust. GR', 'Project', 'Con.']))])
        
        # HTML 변환
        styler = final_df.style.format({col: format_percentage_html if 'ACHI' in col[1] else format_k_val for col in final_df.columns})
        styler.set_table_attributes('class="report-table"')
        styler.applymap_index(lambda v: 'color: #002060; font-weight: bold;', axis=0)
        styler.applymap_index(lambda v: 'color: #0070c0; font-weight: bold;', axis=0, level=1)
        
        html = styler.to_html()
        html = re.sub(r'<th[^>]*>.*?Sales Revenue.*?</th>', lambda m: m.group(0).replace('<th', '<th colspan="3"'), html)
        return f'<div class="table-container">{html}</div>'

    st.markdown(get_biz_report(raw_df, "Core", selected_year, selected_month), unsafe_allow_html=True)
    st.markdown(get_biz_report(raw_df, "Power", selected_year, selected_month), unsafe_allow_html=True)
