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
        background-color: #f8f9fa !important;
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
        
        /* 여기서 여백을 조절합니다 */
        margin-bottom: 1rem !important; 
        padding: 0px !important;
        
        /* 테이블 크기에 딱 맞게 설정 */
        display: inline-block; 
        width: auto;
        min-width: 100%; /* 너비는 최소 100%를 유지하되 */
        box-sizing: border-box;
    }
    
    /* 테이블의 불필요한 기본 margin 제거 */
    .report-table {
        margin: 0 !important;
        border-collapse: collapse !important;
    }
    </style>
""", unsafe_allow_html=True)

st.title("📊 통합 월간 매출 보고서 (FC vs ACT 자동 집계)")

# ==========================================
# 2. 포맷팅 및 고품질 엑셀 다운로드 함수
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

def to_excel_multiple(df_dict):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet_name, df in df_dict.items():
            # 1. 스타일링을 위한 Styler 생성
            styler = df.style.format(lambda x: format_k_val(x) if isinstance(x, (int, float)) else x)
            
            # 2. 배경색/테두리 조건부 서식 적용 (화면과 동일하게)
            def apply_row_style(row):
                if 'TTL (K.€)' in str(row.name) or 'Total' in str(row.name) or 'Subtotal' in str(row.name):
                    return ['background-color: #ffffe0; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;'] * len(row)
                return [''] * len(row)
            
            styler.apply(apply_row_style, axis=1)
            
            # 3. 엑셀로 내보내기
            styler.to_excel(writer, sheet_name=sheet_name[:31])
            
            # 4. 열 너비 자동 조정
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
            df['BIZ Type'] = df['BIZ Type'].replace(['COMM', 'comm'], 'COMMERCIAL')
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
        
        # 인덱스 이름 설정 (첫 번째 컬럼을 CPS로 명시)
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
        
        total_row = final_df.sum(numeric_only=True)
        for phase_name in phases:
            num = total_row.get((phase_name, 'ACT'), 0)
            den = total_row.get((phase_name, '26 FC1'), 0)
            total_row[(phase_name, 'ACHI %')] = num / den if den != 0 else 0
        
        # 합계 행 라벨 변경
        t_label = "TTL (K.€)"
        t_index = tuple([''] * (len(final_df.index.names)-1) + [t_label]) if isinstance(final_df.index, pd.MultiIndex) else t_label
        t_df = pd.DataFrame([total_row], index=[t_index] if not isinstance(final_df.index, pd.MultiIndex) else pd.MultiIndex.from_tuples([t_index], names=final_df.index.names))
        return pd.concat([final_df, t_df]), col_prev, phase_curr

    # (이하 get_biz_type_detailed_report, get_biz_report 함수는 기존과 동일)
    def get_biz_type_detailed_report(df, year, month):
        if month == 1: prev_year, prev_month = year - 1, 12
        else: prev_year, prev_month = year, month - 1
        month_names = {1:'Jan', 2:'Feb', 3:'Mar', 4:'Apr', 5:'May', 6:'Jun', 7:'Jul', 8:'Aug', 9:'Sep', 10:'Oct', 11:'Nov', 12:'Dec'}
        m_str, pm_str = month_names.get(month, f'{month}'), month_names.get(prev_month, f'{prev_month}')
        phase_names = [f'{m_str}. {year}', f'YTD {m_str}. {year}', f'{year} TTL']
        prev_phase_name = f'{pm_str}. {prev_year}'
        results = []
        biz_categories = ['DIRECT', 'COMMERCIAL', 'Unknown']
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
            if ('Others', '', '') in combined.index: combined = pd.concat([combined.drop(index=('Others', '', '')).sort_values(by=(phase_names[0], 'ACT'), ascending=False), combined.loc[[('Others', '', '')]]])
            else: combined = combined.sort_values(by=(phase_names[0], 'ACT'), ascending=False)
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

    # ==========================================
    # 5. 스타일링 및 렌더링
    # ==========================================
    def render_html_view(df, phase_curr):
        df_display = df.replace(0, '')
        format_dict = {col: format_percentage_html if 'ACHI' in col[1] else format_k_val for col in df.columns}
        styler = df_display.style.format(format_dict, na_rep='').set_table_attributes('class="report-table"')
        numeric_cols = get_numeric_cols(df)
        styler.set_properties(subset=numeric_cols, **{'text-align': 'right'})
        # 'TTL (K.€)' 또는 'Total'이 들어간 행에 동일 음영 적용
        styler.apply(lambda row: ['background-color: #ffffe0; color: #002060; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;'] * len(row) if 'TTL (K.€)' in str(row.name) or 'Total' in str(row.name) or 'Subtotal' in str(row.name) else [''] * len(row), axis=1)
        return f'<div class="table-container">{styler.to_html()}</div>'

    def render_biz_html_table(df):
        df_display = df.replace(0, '')
        format_dict = {col: format_percentage_html if 'ACHI' in col[1] else format_k_val for col in df.columns}
        styler = df_display.style.format(format_dict, na_rep='').set_table_attributes('class="report-table"')
        numeric_cols = get_numeric_cols(df)
        styler.set_properties(subset=numeric_cols, **{'text-align': 'right'})
        styler.apply(lambda row: ['background-color: #ffffe0; color: #002060; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;' if '소계' in str(row.name) or 'Total' in str(row.name) else '' for _ in row], axis=1)
        return f'<div class="table-container">{styler.to_html()}</div>'

    # ==========================================
    # 6. 화면 출력
    # ==========================================
    reports_to_download = {}
    
    st.subheader("📌 1. 매출 요약 (CPS 기준)")
    df_cps, p_col, c_col = build_summary_report(raw_df, ['CPS'], selected_year, selected_month, 'TTL (K.€)')
    if not df_cps.empty: st.markdown(render_html_view(df_cps, c_col), unsafe_allow_html=True); reports_to_download["CPS_Summary"] = df_cps

    st.subheader("📌 2. 매출 요약 (Item 기준)")
    df_item_raw = raw_df[raw_df['Item'].isin(['ICCU1', 'ICCU2', 'VCMS'])]
    df_item, p_col, c_col = build_summary_report(df_item_raw, ['Item'], selected_year, selected_month, 'TTL (K.€)')
    if not df_item.empty: st.markdown(render_html_view(df_item, c_col), unsafe_allow_html=True); reports_to_download["Item_Summary"] = df_item

    st.subheader("📌 3. 비즈니스 타입별 매출 요약 (DIRECT / COMM.)")
    df_biz = get_biz_type_detailed_report(raw_df, selected_year, selected_month)
    if not df_biz.empty: st.markdown(render_html_view(df_biz, ""), unsafe_allow_html=True); reports_to_download["Biz_Type_Summary"] = df_biz

    st.subheader("📌 4. Power Electronics 비즈니스")
    df_pe, phase_names_pe = get_biz_report(raw_df, "Power", selected_year, selected_month)
    if not df_pe.empty: st.markdown(render_biz_html_table(df_pe), unsafe_allow_html=True); reports_to_download["PE_Biz"] = df_pe

    st.subheader("📌 5. Core 비즈니스")
    df_core, phase_names_core = get_biz_report(raw_df, "Core", selected_year, selected_month)
    if not df_core.empty: st.markdown(render_biz_html_table(df_core), unsafe_allow_html=True); reports_to_download["Core_Biz"] = df_core

    if reports_to_download:
        st.write("---")
        st.download_button("📥 전체 5개 요약 리포트 엑셀 다운로드 (시트별 분리)", data=to_excel_multiple(reports_to_download), file_name=f"Monthly_Closing_Report_{selected_year}_{selected_month:02d}.xlsx", use_container_width=True)
else:
    st.info("👈 좌측 사이드바에서 엑셀 파일을 업로드하시면 5가지 요약 리포트가 자동 생성됩니다.")
