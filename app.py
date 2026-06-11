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
    
    h1 { font-size: 1.6rem !important; margin-bottom: 0.5rem !important; padding-bottom: 0 !important; }
    h3 { font-size: 1.1rem !important; margin-top: 1rem !important; margin-bottom: 0.5rem !important; color: #002060 !important; }
    
    .report-table {
        border-collapse: collapse;
        font-family: 'Malgun Gothic', sans-serif;
        font-size: 12px;
        width: 100%;
        background-color: white;
    }
    
    .report-table th {
        background-color: #002060 !important;
        color: white !important;
        border: 1px solid #8ea9db !important;
        text-align: center !important;
        padding: 4px 3px !important;
        font-weight: 600 !important;
        font-size: 11.5px !important;
    }
    
    .report-table td {
        border: 1px solid #d9d9d9;
        text-align: right;
        padding: 4px;
        vertical-align: middle;
    }
    
    .report-table th.row_heading, .report-table td.row_heading {
        text-align: center !important;
        border-right: 1px solid #8ea9db;
    }

    .report-table th.row_heading.level0 { color: #002060 !important; font-weight: bold !important; }
    .report-table th.row_heading.level1 { color: #0070c0 !important; font-weight: bold !important; }
    
    .table-container {
        overflow-x: auto;
        overflow-y: auto;
        max-height: 750px;
        border: 2px solid #002060;
        box-shadow: 2px 2px 10px rgba(0,0,0,0.1);
    }
    
    .report-table thead th {
        position: sticky;
        top: 0;
        z-index: 10;
    }
    </style>
""", unsafe_allow_html=True)

st.title("📊 월간 매출 보고서 (Core & Power Electronics)")

# ==========================================
# 2. 포맷팅 함수 (화면용 - K단위 유지)
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
    if val >= 1.0: 
        return f'<span style="color: #00b050; font-weight: bold;">{pct_str} ▲</span>'
    elif val > 0:  
        return f'<span style="color: #c00000; font-weight: bold;">{pct_str} ▼</span>'
    else:
        return pct_str

# ==========================================
# 3. 엑셀 통합 다운로드 함수 (다중 시트 지원)
# ==========================================
def to_excel_multiple(df_dict):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    
    workbook = writer.book
    num_fmt = workbook.add_format({'num_format': '#,##0'})
    pct_fmt = workbook.add_format({'num_format': '0%'})
    
    # 딕셔너리로 받은 데이터프레임들을 각각의 시트에 저장
    for sheet_name, df in df_dict.items():
        df_formatted = df.copy().fillna(0)
        df_formatted.to_excel(writer, index=True, sheet_name=sheet_name)
        worksheet = writer.sheets[sheet_name]
        
        start_col = 3
        for i, col in enumerate(df_formatted.columns):
            fmt = pct_fmt if 'ACHI' in col[1] else num_fmt
            worksheet.set_column(start_col + i, start_col + i, 12, fmt)
            
    writer.close()
    return output.getvalue()

# ==========================================
# 4. 데이터 로드 및 전처리
# ==========================================
uploaded_file = st.sidebar.file_uploader("매출 데이터 엑셀 파일을 업로드하세요.", type=['xlsx', 'xls'])

if uploaded_file:
    @st.cache_data
    def load_and_preprocess(file):
        df = pd.read_excel(file, header=4).iloc[:, :26]
        df.columns = ['Year', 'Month', 'Desc.', 'Date', 'STP', 'Customer', 'LK No.', "Q'ty", 
                      'Rev. ($)', 'Rev. (€)', 'Rev. (₩)', 'BIZ Type', 'Group 1', 'Group 2', 
                      'Project', 'PF', 'Item', 'Source', 'KOx', 'Memo', 'CPS', 
                      'EUR:USD', 'EUR:KRW', 'Business Type', 'Curr.', 'Con.']
        
        df['Year'] = pd.to_numeric(df['Year'], errors='coerce')
        df['Month'] = pd.to_numeric(df['Month'], errors='coerce')
        df = df.dropna(subset=['Year', 'Month'])
        
        df['Rev. (€)'] = pd.to_numeric(df['Rev. (€)'], errors='coerce').fillna(0)
        df.loc[(df['Item'] == 'VCMS') & (df['Source'] == 'KEM-KR'), 'Business Type'] = 'Power electronics'
        df.loc[(df['Item'] == 'VCMS') & (df['Source'] == 'KOASIA'), 'Business Type'] = 'Core Business'
        df.loc[df['Group 1'] == 'GM', 'Group 2'] = 'GM'
        df = df.replace([np.inf, -np.inf], 0)
        
        df['Year'] = df['Year'].astype(int)
        df['Month'] = df['Month'].astype(int)
        return df

    raw_df = load_and_preprocess(uploaded_file)
    years = sorted(raw_df['Year'].unique())
    selected_year = st.sidebar.selectbox("연도", years, index=len(years)-1 if years else 0)
    selected_month = st.sidebar.selectbox("월", sorted(raw_df['Month'].unique()))

    # ==========================================
    # 5. 핵심 비즈니스 로직
    # ==========================================
    def get_biz_report(df, biz_type, year, month):
        if month == 1:
            prev_year, prev_month = year - 1, 12
        else:
            prev_year, prev_month = year, month - 1

        df_biz = df[(df['Business Type'].str.contains(biz_type, case=False, na=False)) & (df['Year'] == year)].copy()
        
        month_names = {1:'Jan', 2:'Feb', 3:'Mar', 4:'Apr', 5:'May', 6:'Jun', 7:'Jul', 8:'Aug', 9:'Sep', 10:'Oct', 11:'Nov', 12:'Dec'}
        m_str, pm_str = month_names.get(month, f'{month}'), month_names.get(prev_month, f'{prev_month}')
        
        phase_names = [f'{m_str}. {year}', f'YTD {m_str}. {year}', f'{year} TTL']
        prev_phase_name = f'{pm_str}. {prev_year}'
        
        results = []
        for brand in ['HYU', 'KIA', 'GM']:
            brand_df = df_biz[df_biz['Group 2'] == brand].copy()
            if brand_df.empty: continue

            p_m = brand_df[brand_df['Month'] == month].pivot_table(index=['Project', 'Con.'], columns='Desc.', values='Rev. (€)', aggfunc='sum').fillna(0)
            p_y = brand_df[brand_df['Month'] <= month].pivot_table(index=['Project', 'Con.'], columns='Desc.', values='Rev. (€)', aggfunc='sum').fillna(0)
            p_fy = brand_df.pivot_table(index=['Project', 'Con.'], columns='Desc.', values='Rev. (€)', aggfunc='sum').fillna(0)

            brand_df_prev = df[(df['Business Type'].str.contains(biz_type, case=False, na=False)) & 
                               (df['Year'] == prev_year) & (df['Month'] == prev_month) & (df['Group 2'] == brand)].copy()
            p_prev = brand_df_prev.pivot_table(index=['Project', 'Con.'], columns='Desc.', values='Rev. (€)', aggfunc='sum').fillna(0)

            if biz_type == "Core" and brand in ['HYU', 'KIA']:
                if not p_m.empty and 'ACT' in p_m.columns:
                    top = p_m[p_m['ACT'] >= 10000].index
                else:
                    top = p_m.index
                
                def group_others(p):
                    if p.empty: return pd.DataFrame(columns=['25 FC3', '26 FC1', 'ACT']).reindex(pd.MultiIndex.from_tuples([], names=['Project', 'Con.']))
                    main = p.loc[p.index.isin(top)]
                    oth = p.loc[~p.index.isin(top)].sum().to_frame().T
                    oth.index = pd.MultiIndex.from_tuples([('Others', '')], names=['Project', 'Con.'])
                    return pd.concat([main, oth])
                
                p_m, p_y, p_fy, p_prev = group_others(p_m), group_others(p_y), group_others(p_fy), group_others(p_prev)

            p_prev = p_prev.reindex(p_m.index, fill_value=0)
            p_y = p_y.reindex(p_m.index, fill_value=0)
            p_fy = p_fy.reindex(p_m.index, fill_value=0)

            combined_dict = {}
            combined_dict[(prev_phase_name, 'ACT')] = p_prev.get('ACT', 0)
            
            phases = [(phase_names[0], p_m), (phase_names[1], p_y), (phase_names[2], p_fy)]
            for phase_name, data in phases:
                for c in ['25 FC3', '26 FC1', 'ACT']:
                    combined_dict[(phase_name, c)] = data.get(c, 0)
                
                fc1_vals = data.get('26 FC1', 0)
                act_vals = data.get('ACT', 0)
                achi_vals = np.where(fc1_vals != 0, act_vals / fc1_vals, 0)
                combined_dict[(phase_name, 'ACHI %')] = pd.Series(achi_vals, index=data.index)

            combined = pd.DataFrame(combined_dict, index=p_m.index)
            
            if ('Others', '') in combined.index:
                oth_row = combined.loc[[('Others', '')]]
                main_rows = combined.drop(index=('Others', ''))
                main_rows = main_rows.sort_values(by=(phase_names[0], 'ACT'), ascending=False)
                combined = pd.concat([main_rows, oth_row])
            else:
                combined = combined.sort_values(by=(phase_names[0], 'ACT'), ascending=False)
            
            subtotal = combined.sum(numeric_only=True)
            for p_name in phase_names:
                sum_fc1 = subtotal.get((p_name, '26 FC1'), 0)
                sum_act = subtotal.get((p_name, 'ACT'), 0)
                subtotal[(p_name, 'ACHI %')] = sum_act / sum_fc1 if sum_fc1 != 0 else 0
            
            combined.index = pd.MultiIndex.from_tuples([(brand, p, c) for p, c in combined.index], names=['Cust. GR', 'Project', 'Con.'])
            results.append(combined)
            
            if brand != 'GM':
                sub_row = pd.DataFrame([subtotal], index=pd.MultiIndex.from_tuples([(brand, '소계', '')], names=['Cust. GR', 'Project', 'Con.']))
                results.append(sub_row)
            
        if not results: return pd.DataFrame(), []
        
        final_df = pd.concat(results)
        
        grand_total = final_df[final_df.index.get_level_values(1) != '소계'].sum(numeric_only=True)
        for p_name in phase_names:
            total_fc1 = grand_total.get((p_name, '26 FC1'), 0)
            total_act = grand_total.get((p_name, 'ACT'), 0)
            grand_total[(p_name, 'ACHI %')] = total_act / total_fc1 if total_fc1 != 0 else 0
            
        if "Core" in biz_type:
            grand_total_label = 'Core Business Sales Revenue (K.€)'
        else:
            grand_total_label = 'Power Electronics Business Sales Revenue (K.€)'
            
        grand_row = pd.DataFrame([grand_total], index=pd.MultiIndex.from_tuples([('', grand_total_label, '')], names=['Cust. GR', 'Project', 'Con.']))
        return pd.concat([final_df, grand_row]), phase_names

    # ==========================================
    # 6. 스타일링 및 순수 HTML 렌더링
    # ==========================================
    def render_html_table(df, phase_names):
        df = df.replace(0, '') 
        format_dict = {col: format_percentage_html if 'ACHI' in col[1] else format_k_val for col in df.columns}
        styler = df.style.format(format_dict, na_rep='')
        styler.set_table_attributes('class="report-table"')
        
        def highlight_specifics(row):
            styles = [''] * len(row)
            if row.name[1] == '소계':
                styles = ['background-color: #f2f2f2; color: #333; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;'] * len(row)
            elif 'Sales Revenue' in str(row.name[1]):
                styles = ['background-color: #e2efda; color: black; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;'] * len(row)
            return styles
            
        styler.apply(highlight_specifics, axis=1)
        
        current_month_label = phase_names[0]
        border_styles = []
        border_styles.append({'selector': 'th', 'props': [('vertical-align', 'middle')]})
        border_styles.append({'selector': 'td', 'props': [('vertical-align', 'middle')]})
        
        for i, col in enumerate(df.columns):
            if col[0] == current_month_label and col[1] == '25 FC3':
                border_styles.append({'selector': f'.col{i}', 'props': [('border-left', '2px solid red')]})
            elif col[0] == current_month_label and 'ACHI' in col[1]:
                border_styles.append({'selector': f'.col{i}', 'props': [('border-right', '2px solid red')]})
            elif 'ACHI' in col[1]:
                border_styles.append({'selector': f'.col{i}', 'props': [('border-right', '1px solid #8ea9db')]})

        styler.set_table_styles(border_styles, overwrite=False)
        html = styler.to_html()
        
        pattern = r'<th[^>]*level0[^>]*>.*?</th>\s*<th[^>]*level1[^>]*>(.*?)</th>\s*<th[^>]*level2[^>]*>.*?</th>'
        def merge_colspan(match):
            text = match.group(1)
            if "Sales Revenue" in text:
                return f'<th colspan="3" class="row_heading" style="text-align: center !important; background-color: #e2efda !important; color: black !important; font-weight: bold !important; border-top: 2px solid #8ea9db !important; border-bottom: 2px solid #8ea9db !important; border-right: 1px solid #8ea9db !important; z-index: 1;">{text}</th>'
            return match.group(0)
            
        html = re.sub(pattern, merge_colspan, html)
        return f'<div class="table-container">{html}</div>'

    # ==========================================
    # 7. 화면 출력 (HTML) 및 통합 다운로드
    # ==========================================
    
    # 엑셀로 다운로드할 데이터프레임들을 모아둘 딕셔너리
    reports_to_download = {}
    
    for b_type in ["Core", "Power"]:
        st.subheader(f"📊 {b_type} Business")
        report, phase_names = get_biz_report(raw_df, b_type, selected_year, selected_month)
        
        if not report.empty:
            html_table = render_html_table(report, phase_names)
            st.markdown(html_table, unsafe_allow_html=True)
            # 다운로드 목록에 추가 (시트 이름으로 사용)
            reports_to_download[b_type] = report
        else:
            st.write("조회된 데이터가 없습니다.")

    # 두 표의 렌더링이 모두 끝나면 화면 하단에 통합 다운로드 버튼 1개만 생성
    if reports_to_download:
        st.write("---") # 구분선 추가
        st.download_button(
            label="📥 통합 리포트 엑셀 다운로드 (Core & Power 시트 분리)", 
            data=to_excel_multiple(reports_to_download), 
            file_name=f"Integrated_Report_{selected_year}_{selected_month}.xlsx"
        )
else:
    st.info("파일을 업로드하면 보고서가 생성됩니다.")
