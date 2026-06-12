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
        border-collapse: collapse !important;
        font-family: 'Malgun Gothic', sans-serif;
        font-size: 12px;
        width: 100%;
        background-color: white;
    }
    .report-table tr { border-bottom: none !important; }
    .report-table td, .report-table th { border-bottom: none !important; border-top: none !important; }
    .report-table th, .report-table td {
        max-width: 250px;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }
    .report-table thead th {
        background-color: #002060 !important;
        color: white !important;
        border: 1px solid #8ea9db !important;
        text-align: center !important;
        padding: 4px 3px !important;
        font-weight: 600 !important;
        font-size: 11.5px !important;
        position: sticky;
        top: 0;
        z-index: 10;
    }
    .report-table td {
        border: 1px solid #d9d9d9;
        text-align: center;
        padding: 4px;
        vertical-align: middle;
    }
    .report-table .row_heading {
        background-color: #f8f9fa; 
        color: #333 !important;
        text-align: left !important;
        padding-left: 10px !important;
        border: 1px solid #d9d9d9 !important;
        vertical-align: middle !important;
        font-weight: bold !important;
    }
    .table-container {
        overflow-x: auto;
        border: 2px solid #002060;
        box-shadow: 2px 2px 10px rgba(0,0,0,0.1);
        margin-bottom: 1rem !important; 
        padding: 0px !important;
        display: inline-block; 
        width: auto;
        min-width: 100%;
        box-sizing: border-box;
    }
    .report-table {
        margin: 0 !important;
        border-collapse: collapse !important;
    }
    </style>
""", unsafe_allow_html=True)

st.title("📊 통합 월간 매출 보고서 (FC vs ACT 자동 집계)")

# ==========================================
# 2. 포맷팅 및 공통 스타일 함수
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

def color_index_cells(v):
    if str(v) == 'HYU': return 'background-color: #e6f2ff;'  # 하늘색
    if str(v) == 'KIA': return 'background-color: #ffe6e6;'  # 분홍색
    return ''

def apply_common_styles(styler, apply_hkmc_color=False):
    # 총계/소계 강조
    styler.apply(lambda row: [
        'background-color: #ffffe0; color: #002060; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;' 
        if any(keyword in str(row.name) for keyword in ['TTL', 'Total', 'Subtotal', '소계']) 
        else '' 
        for _ in row
    ], axis=1)
    
    # 4번, 5번, 6번 테이블에만 HYU/KIA 인덱스 셀 색상 적용
    if apply_hkmc_color:
        if hasattr(styler, 'map_index'):
            styler.map_index(color_index_cells, axis=0, level=0)
        elif hasattr(styler, 'applymap_index'):
            styler.applymap_index(color_index_cells, axis=0, level=0)
            
    return styler

def to_excel_multiple(df_dict):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet_name, df in df_dict.items():
            styler = df.style.format(lambda x: format_k_val(x) if isinstance(x, (int, float)) else x)
            
            # 4번, 5번 및 6번 시트에만 색상 적용
            apply_color = sheet_name in ["PE_HKMC_Summary", "PE_Biz_Detailed", "Core_Biz"]
            styler = apply_common_styles(styler, apply_hkmc_color=apply_color)
            
            styler.to_excel(writer, sheet_name=sheet_name[:31])
            
            worksheet = writer.sheets[sheet_name[:31]]
            for i, col in enumerate(df.columns):
                worksheet.set_column(i+1, i+1, 15)
    return output.getvalue()

# ==========================================
# 3. 데이터 로드 및 전처리
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
        
        if 'BIZ Type' in df.columns:
            df['BIZ Type'] = df['BIZ Type'].replace(['COMM', 'comm', 'COMMERCIAL', 'commercial'], 'COMM')
            df['BIZ Type'] = df['BIZ Type'].fillna('Unknown')
            
        sop_dict = {}
        if len(sheets) > 1:
            df_sop = pd.read_excel(xl, sheet_name=sheets[1])
            sop_dict = dict(zip(df_sop.iloc[:, 0], df_sop.iloc[:, 3]))
            
        df['SOP'] = df['Project'].map(sop_dict)
        df['SOP'] = pd.to_datetime(df['SOP'], errors='coerce').dt.strftime('%Y.%m').fillna(df['SOP'].astype(str))
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
        df['Date'] = df['Date'].astype(str).str.replace('00:00:00', '').str.strip()
        
        return df

    raw_df = load_and_preprocess(uploaded_file)
    years = sorted(raw_df['Year'].unique())
    selected_year = st.sidebar.selectbox("연도", years, index=len(years)-1 if years else 0)
    selected_month = st.sidebar.selectbox("월", sorted(raw_df['Month'].unique()))

    # ==========================================
    # 4. 핵심 비즈니스 로직
    # ==========================================
    def get_numeric_cols(df):
        return [col for col in df.columns if any(x in str(col) for x in ['FC3', 'FC1', 'ACT', 'ACHI'])]

    def build_summary_report(df_sub, index_cols, year, month, total_label="TTL (K.€)", index_names=None, sort_by_current_act=False):
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
            num = pd.Series(combined_dict[(phase_name, 'ACT')])
            den = pd.Series(combined_dict[(phase_name, '26 FC1')])
            combined_dict[(phase_name, 'ACHI %')] = num.div(den).replace([np.inf, -np.inf], 0).fillna(0)
            
        final_df = pd.DataFrame(combined_dict)
        final_df.columns = pd.MultiIndex.from_tuples(col_tuples)
        final_df.index.names = current_index_names
        final_df = final_df.loc[(final_df.filter(like='ACT').sum(axis=1) != 0) | (final_df.filter(like='FC1').sum(axis=1) != 0)]
        
        if sort_by_current_act and (phase_curr, 'ACT') in final_df.columns:
            final_df = final_df.sort_values(by=(phase_curr, 'ACT'), ascending=False)
            
        if 'BIZ Type' in final_df.index.names:
            cats = pd.CategoricalDtype(categories=['DIRECT', 'COMM', 'Unknown'], ordered=True)
            try:
                final_df.index = final_df.index.set_levels(final_df.index.levels[0].astype(cats), level=0)
                final_df = final_df.sort_index(level=0)
            except: pass
            
        total_row = final_df.sum(numeric_only=True)
        for phase_name in phases:
            num = total_row.get((phase_name, 'ACT'), 0)
            den = total_row.get((phase_name, '26 FC1'), 0)
            total_row[(phase_name, 'ACHI %')] = num / den if den != 0 else 0
            
        t_label = total_label
        t_index = tuple([t_label] + [''] * (len(final_df.index.names)-1)) if isinstance(final_df.index, pd.MultiIndex) else t_label
        t_df = pd.DataFrame([total_row], index=[t_index] if not isinstance(final_df.index, pd.MultiIndex) else pd.MultiIndex.from_tuples([t_index], names=final_df.index.names))
        
        return pd.concat([final_df, t_df]), col_prev, phase_curr

    def get_biz_type_detailed_report(df, year, month):
        if month == 1: prev_year, prev_month = year - 1, 12
        else: prev_year, prev_month = year, month - 1
        
        month_names = {1:'Jan', 2:'Feb', 3:'Mar', 4:'Apr', 5:'May', 6:'Jun', 7:'Jul', 8:'Aug', 9:'Sep', 10:'Oct', 11:'Nov', 12:'Dec'}
        m_str, pm_str = month_names.get(month, f'{month}'), month_names.get(prev_month, f'{prev_month}')
        phase_names = [f'{m_str}. {year}', f'YTD {m_str}. {year}', f'{year} TTL']
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
                num = pd.Series(data.get('ACT', 0))
                den = pd.Series(data.get('26 FC1', 0))
                combined_dict[(phase_name, 'ACHI %')] = num.div(den).replace([np.inf, -np.inf], 0).fillna(0)
                
            combined = pd.DataFrame(combined_dict, index=p_m.index)
            
            if (phase_names[0], 'ACT') in combined.columns:
                combined = combined.sort_values(by=(phase_names[0], 'ACT'), ascending=False)
                
            subtotal = combined.sum(numeric_only=True)
            for p_name in phase_names:
                num = subtotal.get((p_name, 'ACT'), 0)
                den = subtotal.get((p_name, '26 FC1'), 0)
                subtotal[(p_name, 'ACHI %')] = num / den if den != 0 else 0
                
            results.append(combined)
            results.append(pd.DataFrame([subtotal], index=pd.MultiIndex.from_tuples([(biz, 'Subtotal')], names=['BIZ Type', 'KOx'])))
            
        return pd.concat(results)

    def get_biz_report(df, biz_type, year, month):
        if month == 1: prev_year, prev_month = year - 1, 12
        else: prev_year, prev_month = year, month - 1
        
        df_biz = df[(df['Business Type'].str.contains(biz_type, case=False, na=False)) & (df['Year'] == year)].copy()
        month_names = {1:'Jan', 2:'Feb', 3:'Mar', 4:'Apr', 5:'May', 6:'Jun', 7:'Jul', 8:'Aug', 9:'Sep', 10:'Oct', 11:'Nov', 12:'Dec'}
        m_str, pm_str = month_names.get(month, f'{month}'), month_names.get(prev_month, f'{prev_month}')
        phase_names = [f'{m_str}. {year}', f'YTD {m_str}. {year}', f'{year} TTL']
        prev_phase_name = f'{pm_str}. {prev_year}'
        
        results = []
        for brand in ['HYU', 'KIA', 'GM']:
            brand_df = df_biz[df_biz['Group 2'] == brand].copy()
            if brand_df.empty: continue
            
            p_m = brand_df[brand_df['Month'] == month].pivot_table(index=['Project', 'Con.', 'SOP'], columns='Desc.', values='Rev. (€)', aggfunc='sum').fillna(0)
            p_y = brand_df[brand_df['Month'] <= month].pivot_table(index=['Project', 'Con.', 'SOP'], columns='Desc.', values='Rev. (€)', aggfunc='sum').fillna(0)
            p_fy = brand_df.pivot_table(index=['Project', 'Con.', 'SOP'], columns='Desc.', values='Rev. (€)', aggfunc='sum').fillna(0)
            p_prev = df[(df['Business Type'].str.contains(biz_type, case=False, na=False)) & (df['Year'] == prev_year) & (df['Month'] == prev_month) & (df['Group 2'] == brand)].pivot_table(index=['Project', 'Con.', 'SOP'], columns='Desc.', values='Rev. (€)', aggfunc='sum').fillna(0)
            
            if "Core" in biz_type and brand in ['HYU', 'KIA']:
                top = p_m[p_m['ACT'] >= 10000].index if not p_m.empty and 'ACT' in p_m.columns else p_m.index
                def group_others(p):
                    if p.empty: return pd.DataFrame(columns=['25 FC3', '26 FC1', 'ACT']).reindex(pd.MultiIndex.from_tuples([], names=['Project', 'Con.', 'SOP']))
                    main = p.loc[p.index.isin(top)]; oth = p.loc[~p.index.isin(top)].sum().to_frame().T; oth.index = pd.MultiIndex.from_tuples([('Others', '', '')], names=['Project', 'Con.', 'SOP'])
                    return pd.concat([main, oth])
                p_m, p_y, p_fy, p_prev = group_others(p_m), group_others(p_y), group_others(p_fy), group_others(p_prev)
                
            p_prev, p_y, p_fy = p_prev.reindex(p_m.index, fill_value=0), p_y.reindex(p_m.index, fill_value=0), p_fy.reindex(p_m.index, fill_value=0)
            combined_dict = {(prev_phase_name, 'ACT'): p_prev.get('ACT', 0)}
            
            for phase_name, data in [(phase_names[0], p_m), (phase_names[1], p_y), (phase_names[2], p_fy)]:
                for c in ['25 FC3', '26 FC1', 'ACT']: combined_dict[(phase_name, c)] = data.get(c, 0)
                num = pd.Series(data.get('ACT', 0))
                den = pd.Series(data.get('26 FC1', 0))
                combined_dict[(phase_name, 'ACHI %')] = num.div(den).replace([np.inf, -np.inf], 0).fillna(0)
                
            combined = pd.DataFrame(combined_dict, index=p_m.index)
            if ('Others', '', '') in combined.index: 
                combined = pd.concat([combined.drop(index=('Others', '', '')).sort_values(by=(phase_names[0], 'ACT'), ascending=False), combined.loc[[('Others', '', '')]]])
            else: 
                combined = combined.sort_values(by=(phase_names[0], 'ACT'), ascending=False)
                
            subtotal = combined.sum(numeric_only=True)
            for p_name in phase_names:
                num = subtotal.get((p_name, 'ACT'), 0)
                den = subtotal.get((p_name, '26 FC1'), 0)
                subtotal[(p_name, 'ACHI %')] = num / den if den != 0 else 0
                
            combined.index = pd.MultiIndex.from_tuples([(brand, p, c, s) for p, c, s in combined.index], names=['Cust. GR', 'Project', 'Con.', 'SOP'])
            results.append(combined)
            if brand != 'GM': results.append(pd.DataFrame([subtotal], index=pd.MultiIndex.from_tuples([(brand, '소계', '', '')], names=['Cust. GR', 'Project', 'Con.', 'SOP'])))
            
        final_df = pd.concat(results)
        grand_total = final_df[final_df.index.get_level_values(1) != '소계'].sum(numeric_only=True)
        for p_name in phase_names:
            num = grand_total.get((p_name, 'ACT'), 0)
            den = grand_total.get((p_name, '26 FC1'), 0)
            grand_total[(p_name, 'ACHI %')] = num / den if den != 0 else 0
            
        grand_row = pd.DataFrame([grand_total], index=pd.MultiIndex.from_tuples([('', f'{biz_type} Total', '', '')], names=['Cust. GR', 'Project', 'Con.', 'SOP']))
        return pd.concat([final_df, grand_row]), phase_names

    def build_trend_report(df, end_year, end_month):
        months = []
        curr_y, curr_m = end_year, end_month
        for _ in range(12):
            months.append((curr_y, curr_m))
            curr_m -= 1
            if curr_m == 0:
                curr_m = 12
                curr_y -= 1
        months.reverse() 
        
        month_names = {1:'Jan', 2:'Feb', 3:'Mar', 4:'Apr', 5:'May', 6:'Jun', 7:'Jul', 8:'Aug', 9:'Sep', 10:'Oct', 11:'Nov', 12:'Dec'}
        df_act = df[df['Desc.'] == 'ACT']
        
        pivot_data = {}
        for y, m in months:
            col_name = f"{month_names[m]}.{str(y)[-2:]}"
            temp_df = df_act[(df_act['Year'] == y) & (df_act['Month'] == m)]
            pivot_data[col_name] = temp_df.groupby('Group 2')['Rev. (€)'].sum()
            
        trend_df = pd.DataFrame(pivot_data)
        row_order = ['HYU', 'KIA', 'GM']
        trend_df = trend_df.reindex(row_order).fillna(0)
        
        trend_df.loc['TTL (K.€)'] = trend_df.sum(numeric_only=True)
        trend_df.index.name = ''
        return trend_df

    # ==========================================
    # 5. 스타일링 및 렌더링
    # ==========================================
    def render_html_view(df, phase_curr, apply_color=False):
        df_display = df.replace(0, '')
        format_dict = {col: format_percentage_html if 'ACHI' in col[1] else format_k_val for col in df.columns}
        styler = df_display.style.format(format_dict, na_rep='').set_table_attributes('class="report-table"')
        styler.set_table_styles([
            {'selector': 'th, td', 'props': [('border-collapse', 'separate')]},
            {'selector': 'tr', 'props': [('display', 'table-row')]}
        ])
        
        numeric_cols = get_numeric_cols(df)
        styler.set_properties(subset=numeric_cols, **{'text-align': 'right'})
        styler = apply_common_styles(styler, apply_hkmc_color=apply_color)
        return f'<div class="table-container">{styler.to_html()}</div>'

    def render_biz_html_table(df, apply_color=False):
        df_display = df.replace(0, '')
        format_dict = {col: format_percentage_html if 'ACHI' in col[1] else format_k_val for col in df.columns}
        styler = df_display.style.format(format_dict, na_rep='').set_table_attributes('class="report-table"')
        
        numeric_cols = get_numeric_cols(df)
        styler.set_properties(subset=numeric_cols, **{'text-align': 'right'})
        styler = apply_common_styles(styler, apply_hkmc_color=apply_color)
        return f'<div class="table-container">{styler.to_html()}</div>'

    def render_trend_html_table(df, apply_color=False):
        df_display = df.replace(0, '')
        format_dict = {col: format_k_val for col in df.columns}
        styler = df_display.style.format(format_dict, na_rep='').set_table_attributes('class="report-table"')
        
        styler.set_properties(**{'text-align': 'right'})
        styler = apply_common_styles(styler, apply_hkmc_color=apply_color)
        return f'<div class="table-container">{styler.to_html()}</div>'

    # ==========================================
    # 6. 화면 출력
    # ==========================================
    reports_to_download = {}

    st.subheader("📌 1. 매출 요약 (CPS 기준)")
    df_cps, p_col, c_col = build_summary_report(raw_df, ['CPS'], selected_year, selected_month, 'TTL (K.€)')
    if not df_cps.empty: 
        st.markdown(render_html_view(df_cps, c_col, apply_color=False), unsafe_allow_html=True)
        reports_to_download["CPS_Summary"] = df_cps

    st.subheader("📌 2. 매출 요약 (Item 기준)")
    df_item_raw = raw_df[raw_df['Item'].isin(['ICCU1', 'ICCU2', 'VCMS'])]
    df_item, p_col, c_col = build_summary_report(df_item_raw, ['Item'], selected_year, selected_month, 'TTL (K.€)', sort_by_current_act=True)
    if not df_item.empty: 
        st.markdown(render_html_view(df_item, c_col, apply_color=False), unsafe_allow_html=True)
        reports_to_download["Item_Summary"] = df_item

    st.subheader("📌 3. 비즈니스 타입별 매출 요약 (DIRECT / COMM.)")
    df_biz = get_biz_type_detailed_report(raw_df, selected_year, selected_month)
    if not df_biz.empty: 
        st.markdown(render_html_view(df_biz, "", apply_color=False), unsafe_allow_html=True)
        reports_to_download["Biz_Type_Summary"] = df_biz

    st.subheader("📌 4. Power Electronics 매출 요약 (HKMC 기준)")
    df_pe_raw = raw_df[raw_df['Business Type'].str.contains("Power", case=False, na=False)].copy()
    if not df_pe_raw.empty:
        df_pe_raw['Cust. GR'] = df_pe_raw['Group 2'].replace({'HYU': 'HKMC', 'KIA': 'HKMC'})
        df_pe_hkmc = df_pe_raw[df_pe_raw['Cust. GR'] == 'HKMC']
        
        df_pe_summary, p_col, c_col = build_summary_report(
            df_pe_hkmc, 
            ['Cust. GR', 'KOx'], 
            selected_year, 
            selected_month, 
            total_label='PE Biz Rev. TTL (K.€)', 
            sort_by_current_act=True
        )
        if not df_pe_summary.empty:
            # 4번 테이블에 적용
            st.markdown(render_html_view(df_pe_summary, c_col, apply_color=True), unsafe_allow_html=True)
            reports_to_download["PE_HKMC_Summary"] = df_pe_summary

    st.subheader("📌 5. Power Electronics 비즈니스 (상세)")
    df_pe, phase_names_pe = get_biz_report(raw_df, "Power", selected_year, selected_month)
    if not df_pe.empty: 
        # 5번 테이블에 적용
        st.markdown(render_biz_html_table(df_pe, apply_color=True), unsafe_allow_html=True)
        reports_to_download["PE_Biz_Detailed"] = df_pe

    st.subheader("📌 6. Core 비즈니스")
    df_core, phase_names_core = get_biz_report(raw_df, "Core", selected_year, selected_month)
    if not df_core.empty: 
        # 6번 테이블에 적용 (수정됨)
        st.markdown(render_biz_html_table(df_core, apply_color=True), unsafe_allow_html=True)
        reports_to_download["Core_Biz"] = df_core

    st.subheader("📌 7. 고객사별 월별 매출 트렌드 (최근 12개월)")
    df_trend = build_trend_report(raw_df, selected_year, selected_month)
    if not df_trend.empty:
        st.markdown(render_trend_html_table(df_trend, apply_color=False), unsafe_allow_html=True)
        reports_to_download["12M_Trend_Report"] = df_trend

    if reports_to_download:
        st.write("---")
        st.download_button("📥 전체 요약 리포트 엑셀 다운로드 (시트별 분리)", data=to_excel_multiple(reports_to_download), file_name=f"Monthly_Closing_Report_{selected_year}_{selected_month:02d}.xlsx", use_container_width=True)

else:
    st.info("👈 좌측 사이드바에서 엑셀 파일을 업로드하시면 요약 리포트가 자동 생성됩니다.")
