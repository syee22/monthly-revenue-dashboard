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
    
    /* 상단 컬럼 헤더 스타일 */
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
    
    /* 일반 데이터 셀 스타일 */
    .report-table td {
        border: 1px solid #d9d9d9;
        text-align: right;
        padding: 4px;
        vertical-align: middle;
    }
    
    /* 좌측 인덱스(Row Heading) 스타일 */
    .report-table th.row_heading, .report-table td.row_heading {
        background-color: white !important;
        color: #333 !important;
        text-align: center !important;
        border: 1px solid #d9d9d9 !important;
        border-right: 1px solid #8ea9db;
        vertical-align: middle !important;
    }

    .report-table th.row_heading.level0 { color: #002060 !important; font-weight: bold !important; }
    .report-table th.row_heading.level1 { color: #0070c0 !important; font-weight: bold !important; }
    
    .table-container {
        overflow-x: auto;
        overflow-y: auto;
        max-height: 750px;
        border: 2px solid #002060;
        box-shadow: 2px 2px 10px rgba(0,0,0,0.1);
        margin-bottom: 2rem;
    }
    </style>
""", unsafe_allow_html=True)

st.title("📊 통합 월간 매출 보고서 (FC vs ACT 자동 집계)")

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
    
    for sheet_name, df in df_dict.items():
        if not df.empty:
            df_formatted = df.copy().fillna(0)
            df_formatted.to_excel(writer, index=True, sheet_name=sheet_name[:31])
            worksheet = writer.sheets[sheet_name[:31]]
            
            start_col = len(df.index.names)
            for i, col in enumerate(df_formatted.columns):
                # 컬럼명에 ACHI가 포함되어 있으면 퍼센트 포맷, 아니면 숫자 포맷 적용
                fmt = pct_fmt if 'ACHI' in str(col) else num_fmt
                worksheet.set_column(start_col + i, start_col + i, 12, fmt)
            
    writer.close()
    return output.getvalue()

# ==========================================
# 4. 데이터 로드 및 전처리
# ==========================================
uploaded_file = st.sidebar.file_uploader("SAP/엑셀 데이터를 업로드하세요.", type=['xlsx', 'xls'])

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
        
        # SOP(Date) 포맷 정리
        df['Date'] = df['Date'].astype(str).str.replace('00:00:00', '').str.strip()
        
        return df

    raw_df = load_and_preprocess(uploaded_file)
    years = sorted(raw_df['Year'].unique())
    selected_year = st.sidebar.selectbox("연도", years, index=len(years)-1 if years else 0)
    selected_month = st.sidebar.selectbox("월", sorted(raw_df['Month'].unique()))

    # ==========================================
    # 5-1. 핵심 비즈니스 로직 (1~3번 범용 리포트용)
    # ==========================================
    def build_summary_report(df_sub, index_cols, year, month, total_label, index_names=None):
        if df_sub.empty:
            return pd.DataFrame(), "", ""

        if month == 1:
            prev_year, prev_month = year - 1, 12
        else:
            prev_year, prev_month = year, month - 1

        month_names = {1:'Jan', 2:'Feb', 3:'Mar', 4:'Apr', 5:'May', 6:'Jun', 7:'Jul', 8:'Aug', 9:'Sep', 10:'Oct', 11:'Nov', 12:'Dec'}
        m_str = month_names.get(month, f'{month}')
        pm_str = month_names.get(prev_month, f'{prev_month}')

        col_prev = f'{pm_str}. {year if month != 1 else prev_year}'
        phase_curr = f'{m_str}. {year}'
        phase_ytd = f'YTD {m_str}. {year}'
        phase_ttl = f'{year} TTL'
        phases = [phase_curr, phase_ytd, phase_ttl]

        def get_pivot(d):
            if d.empty: return pd.DataFrame()
            return d.pivot_table(index=index_cols, columns='Desc.', values='Rev. (€)', aggfunc='sum').fillna(0)

        df_prev = df_sub[(df_sub['Year'] == prev_year) & (df_sub['Month'] == prev_month) & (df_sub['Desc.'] == 'ACT')]
        s_prev = df_prev.groupby(index_cols)['Rev. (€)'].sum() if not df_prev.empty else pd.Series(dtype=float)

        df_curr = df_sub[(df_sub['Year'] == year) & (df_sub['Month'] == month)]
        df_ytd = df_sub[(df_sub['Year'] == year) & (df_sub['Month'] <= month)]
        df_ttl = df_sub[(df_sub['Year'] == year)]

        p_curr = get_pivot(df_curr)
        p_ytd = get_pivot(df_ytd)
        p_ttl = get_pivot(df_ttl)

        all_indices = set()
        for p in [s_prev, p_curr, p_ytd, p_ttl]:
            if not p.empty:
                if isinstance(p.index, pd.MultiIndex):
                    all_indices.update(p.index.tolist())
                else:
                    all_indices.update([(x,) for x in p.index.tolist()])

        if not all_indices:
            return pd.DataFrame(), col_prev, phase_curr

        all_indices = sorted(list(all_indices), key=lambda x: tuple(str(i) for i in x))
        
        if len(index_cols) > 1:
            idx = pd.MultiIndex.from_tuples(all_indices, names=index_names or index_cols)
        else:
            idx = pd.Index([x[0] for x in all_indices], name=(index_names[0] if index_names else index_cols[0]))

        combined_dict = {}
        col_tuples = [('', col_prev)]
        for p in phases:
            for c in ['25 FC3', '26 FC1', 'ACT', 'ACHI %']:
                col_tuples.append((p, c))

        if not s_prev.empty:
            s_prev.index = s_prev.index if len(index_cols) == 1 else pd.MultiIndex.from_tuples(s_prev.index)
            combined_dict[('', col_prev)] = s_prev.reindex(idx).fillna(0)
        else:
            combined_dict[('', col_prev)] = pd.Series(0, index=idx)

        for phase_name, data in zip(phases, [p_curr, p_ytd, p_ttl]):
            for c in ['25 FC3', '26 FC1', 'ACT']:
                if not data.empty and c in data.columns:
                    data.index = data.index if len(index_cols) == 1 else pd.MultiIndex.from_tuples(data.index)
                    combined_dict[(phase_name, c)] = data[c].reindex(idx).fillna(0)
                else:
                    combined_dict[(phase_name, c)] = pd.Series(0, index=idx)

            fc1 = combined_dict[(phase_name, '26 FC1')]
            act = combined_dict[(phase_name, 'ACT')]
            combined_dict[(phase_name, 'ACHI %')] = np.where(fc1 != 0, act / fc1, 0)

        final_df = pd.DataFrame(combined_dict)
        final_df.columns = pd.MultiIndex.from_tuples(col_tuples)

        check_cols = [c for c in final_df.columns if c[1] in ['25 FC3', '26 FC1', 'ACT']]
        if check_cols:
            final_df = final_df.loc[(final_df[check_cols] != 0).any(axis=1)]

        if (phase_curr, 'ACT') in final_df.columns:
            if len(index_cols) > 1:
                temp_level0 = final_df.index.get_level_values(0)
                final_df[('__temp__', 'level0')] = temp_level0
                final_df = final_df.sort_values(by=[('__temp__', 'level0'), (phase_curr, 'ACT')], ascending=[True, False])
                final_df = final_df.drop(columns=[('__temp__', 'level0')])
            else:
                final_df = final_df.sort_values(by=(phase_curr, 'ACT'), ascending=False)

        total_row = final_df.sum(numeric_only=True)
        for phase_name in phases:
            t_fc1 = total_row[(phase_name, '26 FC1')]
            t_act = total_row[(phase_name, 'ACT')]
            total_row[(phase_name, 'ACHI %')] = t_act / t_fc1 if t_fc1 != 0 else 0

        if len(index_cols) > 1:
            t_idx = ('', total_label) if len(index_cols) == 2 else tuple([''] * (len(index_cols)-1) + [total_label])
            t_df = pd.DataFrame([total_row], index=pd.MultiIndex.from_tuples([t_idx], names=index_names or index_cols))
        else:
            t_df = pd.DataFrame([total_row], index=pd.Index([total_label], name=(index_names[0] if index_names else index_cols[0])))

        final_df = pd.concat([final_df, t_df])
        return final_df, col_prev, phase_curr

    # ==========================================
    # 5-2. 핵심 비즈니스 로직 (4~5번 Core/Power 전용 그룹핑 로직)
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

            # Core 혹은 Power 비즈니스의 주요 프로젝트 필터링(10K 기준)
            if brand in ['HYU', 'KIA']:
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
    # 6-1. 스타일링 및 렌더링 (1~3번 범용 리포트용)
    # ==========================================
    def render_html_view(df, phase_curr):
        df = df.replace(0, '') 
        format_dict = {col: format_percentage_html if 'ACHI' in col[1] else format_k_val for col in df.columns}
        
        styler = df.style.format(format_dict, na_rep='')
        styler.set_table_attributes('class="report-table"')
        
        def highlight_totals(row):
            row_name = str(row.name[-1]) if isinstance(row.name, tuple) else str(row.name)
            if 'TTL' in row_name:
                return ['background-color: #ffffe0; color: #002060; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;'] * len(row)
            return [''] * len(row)
            
        styler.apply(highlight_totals, axis=1)
        
        border_styles = [
            {'selector': 'th', 'props': [('vertical-align', 'middle')]},
            {'selector': 'td', 'props': [('vertical-align', 'middle')]}
        ]
        
        for i, col in enumerate(df.columns):
            if col[0] == phase_curr and col[1] == '25 FC3':
                border_styles.append({'selector': f'.col{i}', 'props': [('border-left', '2px solid #c00000')]})
            elif col[0] == phase_curr and 'ACHI' in col[1]:
                border_styles.append({'selector': f'.col{i}', 'props': [('border-right', '2px solid #c00000')]})
            elif 'ACHI' in col[1]:
                border_styles.append({'selector': f'.col{i}', 'props': [('border-right', '1px solid #8ea9db')]})

        styler.set_table_styles(border_styles, overwrite=False)
        return f'<div class="table-container">{styler.to_html()}</div>'

    # ==========================================
    # 6-2. 스타일링 및 렌더링 (4~5번 Core/Power 전용 리포트용)
    # ==========================================
    def render_biz_html_table(df):
        df = df.replace(0, '') 
        format_dict = {col: format_percentage_html if 'ACHI' in col[1] else format_k_val for col in df.columns}
        styler = df.style.format(format_dict, na_rep='')
        styler.set_table_attributes('class="report-table"')
        
        styler.apply(lambda row: [
            'background-color: #f2f2f2; color: #333; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;' if row.name[1] == '소계' 
            else 'background-color: #e2efda; color: black; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;' if 'Sales Revenue' in str(row.name[1]) 
            else '' for _ in row
        ], axis=1)
        
        # 합계 행의 컬럼 인덱스를 병합(colspan)하여 깔끔하게 보여주기 위한 정규식 처리
        html = re.sub(r'<th[^>]*level0[^>]*>.*?</th>\s*<th[^>]*level1[^>]*>(.*?)</th>\s*<th[^>]*level2[^>]*>.*?</th>', 
                      lambda m: m.group(0).replace('<th', '<th colspan="3"') if "Sales Revenue" in m.group(1) else m.group(0), 
                      styler.to_html())
        return f'<div class="table-container">{html}</div>'

    # ==========================================
    # 7. 화면 출력 (5가지 View) 및 통합 다운로드
    # ==========================================
    reports_to_download = {}

    st.subheader("📌 1. 매출 요약 (CPS 기준)")
    df_cps, p_col, c_col = build_summary_report(raw_df, ['CPS'], selected_year, selected_month, 'TTL (K.€)')
    if not df_cps.empty:
        st.markdown(render_html_view(df_cps, c_col), unsafe_allow_html=True)
        reports_to_download["CPS_Summary"] = df_cps

    st.subheader("📌 2. 매출 요약 (Item 기준)")
    df_item_raw = raw_df[raw_df['Item'].isin(['ICCU1', 'ICCU2', 'VCMS'])]
    df_item, p_col, c_col = build_summary_report(df_item_raw, ['Item'], selected_year, selected_month, 'TTL (K.€)')
    if not df_item.empty:
        st.markdown(render_html_view(df_item, c_col), unsafe_allow_html=True)
        reports_to_download["Item_Summary"] = df_item

    st.subheader("📌 3. 비즈니스 타입별 매출 요약 (DIRECT / COMM.)")
    df_biz, p_col, c_col = build_summary_report(raw_df, ['BIZ Type', 'KOx'], selected_year, selected_month, 'Sales Rev. TTL (K.€)', index_names=['Biz Type', 'KOx'])
    if not df_biz.empty:
        st.markdown(render_html_view(df_biz, c_col), unsafe_allow_html=True)
        reports_to_download["Biz_Type_Summary"] = df_biz

    # --- 여기서부터 기존 로직 적용 (Core / Power) ---
    st.subheader("📌 4. Power Electronics 비즈니스 (고객사별)")
    df_pe, phase_names_pe = get_biz_report(raw_df, "Power", selected_year, selected_month)
    if not df_pe.empty:
        st.markdown(render_biz_html_table(df_pe), unsafe_allow_html=True)
        reports_to_download["PE_Biz"] = df_pe

    st.subheader("📌 5. Core 비즈니스 (고객사별)")
    df_core, phase_names_core = get_biz_report(raw_df, "Core", selected_year, selected_month)
    if not df_core.empty:
        st.markdown(render_biz_html_table(df_core), unsafe_allow_html=True)
        reports_to_download["Core_Biz"] = df_core

    # 통합 다운로드 버튼
    if reports_to_download:
        st.write("---")
        st.download_button(
            label="📥 전체 5개 요약 리포트 엑셀 다운로드 (시트별 분리)", 
            data=to_excel_multiple(reports_to_download), 
            file_name=f"Monthly_Closing_Report_{selected_year}_{selected_month:02d}.xlsx",
            use_container_width=True
        )
else:
    st.info("👈 좌측 사이드바에서 엑셀 파일을 업로드하시면 5가지 요약 리포트가 자동 생성됩니다.")
