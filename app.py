import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import uuid
import plotly.express as px
import plotly.graph_objects as go

# ==========================================
# 1. 페이지 설정 및 전역 변수 설정
# ==========================================
st.set_page_config(page_title="Sales Revenue & Price Report", layout="wide")

MONTH_NAMES = {1:'Jan', 2:'Feb', 3:'Mar', 4:'Apr', 5:'May', 6:'Jun', 7:'Jul', 8:'Aug', 9:'Sep', 10:'Oct', 11:'Nov', 12:'Dec'}
BIZ_CONFIG = {"Power": "PE Biz", "Core": "Core Biz"}

# ==========================================
# 전역 CSS 주입
# ==========================================
st.markdown("""<style>
.block-container { padding: 2rem 3rem; }
h1 { font-size: 1.6rem !important; margin-bottom: 0.5rem !important; padding-bottom: 0 !important; }
h3 { font-size: 1.1rem !important; margin-top: 1rem !important; margin-bottom: 0.5rem !important; color: #002060 !important; }
.table-container { overflow-x: auto; box-shadow: 2px 2px 10px rgba(0,0,0,0.1); margin-bottom: 1rem !important; padding: 2px !important; display: inline-block; width: auto; min-width: 100%; box-sizing: border-box; background-color: white; }
.report-table { border-collapse: collapse !important; font-family: 'Arial', sans-serif; font-size: 12px; width: 100%; background-color: white; margin: 0 !important; border: 2px solid #002060 !important; }
.report-table tr { border-bottom: none !important; }
.report-table td, .report-table th { border-bottom: none !important; border-top: none !important; }
.report-table th, .report-table td { max-width: 250px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.report-table thead th { background-color: #002060 !important; color: white !important; border: 1px solid #8ea9db !important; text-align: center !important; padding: 4px 3px !important; font-weight: 600 !important; font-size: 11.5px !important; position: sticky; top: 0; z-index: 10; }
.report-table td { border: 1px solid #d9d9d9; text-align: center; padding: 4px; vertical-align: middle; }
.report-table .row_heading { color: #002060 !important; text-align: left !important; padding-left: 10px !important; border: 1px solid #d9d9d9 !important; vertical-align: middle !important; font-weight: bold !important; }
.report-table tr.total-row th, .report-table tr.total-row td { background-color: #ffffe0 !important; color: #002060 !important; font-weight: bold !important; border-top: 2px solid #8ea9db !important; border-bottom: 2px solid #8ea9db !important; }
.report-table tr.total-row-hyu th, .report-table tr.total-row-hyu td { background-color: #e6f2ff !important; color: #002060 !important; font-weight: bold !important; border-top: 2px solid #8ea9db !important; border-bottom: 2px solid #8ea9db !important; }
.report-table tr.total-row-kia th, .report-table tr.total-row-kia td { background-color: #ffe6e6 !important; color: #002060 !important; font-weight: bold !important; border-top: 2px solid #8ea9db !important; border-bottom: 2px solid #8ea9db !important; }
.report-table tr.total-row-gm th, .report-table tr.total-row-gm td { background-color: #e6e6e6 !important; color: #002060 !important; font-weight: bold !important; border-top: 2px solid #8ea9db !important; border-bottom: 2px solid #8ea9db !important; }
.report-table tr.total-row-direct th, .report-table tr.total-row-direct td { background-color: #99caff !important; color: #002060 !important; font-weight: bold !important; border-top: 2px solid #8ea9db !important; border-bottom: 2px solid #8ea9db !important; }
.report-table tr.total-row-comm th, .report-table tr.total-row-comm td { background-color: #d0d0d0 !important; color: #002060 !important; font-weight: bold !important; border-top: 2px solid #8ea9db !important; border-bottom: 2px solid #8ea9db !important; }
.report-table tr.total-row-exrate th, .report-table tr.total-row-exrate td { background-color: #e2efda !important; color: #002060 !important; font-weight: bold !important; border-top: 2px solid #8ea9db !important; border-bottom: 2px solid #8ea9db !important; }

.report-table.biz-table .row_heading { text-align: center !important; padding-left: 4px !important; padding-right: 4px !important; }
.report-table th.level2, .report-table th.level3, .report-table.biz-table th.level2, .report-table.biz-table th.level3 { 
    width: 50px !important; min-width: 50px !important; max-width: 50px !important; padding-left: 2px !important; padding-right: 2px !important; text-align: center !important; font-size: 11px !important; white-space: normal !important; word-break: break-all !important; 
}
.report-table tbody tr { height: 15px !important; max-height: 15px !important; }
.report-table tbody td, .report-table tbody th { height: 15px !important; max-height: 15px !important; padding-top: 0px !important; padding-bottom: 0px !important; line-height: 15px !important; }
</style>""", unsafe_allow_html=True)

# ==========================================
# 2. 동적 테두리 및 포맷터 생성
# ==========================================
def get_trend_highlight_css(table_id):
    return f"<style>#{table_id} thead tr:nth-child(1) th:last-child {{ border-top: 4px solid #c00000 !important; border-left: 4px solid #c00000 !important; border-right: 4px solid #c00000 !important; }} #{table_id} tbody td:last-child {{ border-left: 4px solid #c00000 !important; border-right: 4px solid #c00000 !important; }} #{table_id} tbody tr:last-child td:last-child {{ border-bottom: 4px solid #c00000 !important; }}</style>"

def get_dynamic_highlight_css(table_id, df, highlight_phase):
    if not highlight_phase: return ""
    cols = list(df.columns)
    start_col, end_col = -1, -1
    level0_cols = []
    for i, col in enumerate(cols):
        c0 = col[0] if isinstance(col, tuple) else str(col)
        if not level0_cols or level0_cols[-1] != c0: level0_cols.append(c0)
        if c0 == highlight_phase:
            if start_col == -1: start_col = i
            end_col = i
    if start_col == -1: return ""
    num_indices = df.index.nlevels
    target_th_row0 = num_indices + level0_cols.index(highlight_phase) + 1
    target_th_row1_start = start_col + 1
    target_th_row1_end = end_col + 1
    td_start = start_col + 1
    td_end = end_col + 1
    return f"<style>#{table_id} thead tr:nth-child(1) th:nth-child({target_th_row0}) {{ border-top: 5px solid #c00000 !important; border-left: 5px solid #c00000 !important; border-right: 5px solid #c00000 !important; }} #{table_id} thead tr:nth-child(2) th:nth-child({target_th_row1_start}) {{ border-left: 5px solid #c00000 !important; }} #{table_id} thead tr:nth-child(2) th:nth-child({target_th_row1_end}) {{ border-right: 5px solid #c00000 !important; }} #{table_id} tbody td:nth-of-type({td_start}) {{ border-left: 5px solid #c00000 !important; }} #{table_id} tbody td:nth-of-type({td_end}) {{ border-right: 5px solid #c00000 !important; }} #{table_id} tbody tr:last-child td:nth-of-type(n+{td_start}):nth-of-type(-n+{td_end}) {{ border-bottom: 5px solid #c00000 !important; }}</style>"

def get_numeric_cols(df): 
    return [col for col in df.columns if any(x in str(col) for x in ['FC3', 'FC1', 'ACT', 'ACHI'])]

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
    shadow = "text-shadow: 1px 1px 1px rgba(0,0,0,0.3);"
    if 0.95 <= val <= 1.0:
        bar_html = '<span style="display:inline-block; width:9.5px; height:2px; background-color:#404040; vertical-align:middle; margin-bottom:1px; margin-left:3px;"></span>'
        return f'<span style="color: #404040; font-style: italic;">{pct_str} {bar_html}</span>'
    elif val > 1.0: return f'<span style="color: #145A32; font-style: italic;">{pct_str} <span style="{shadow}">▲</span></span>'
    elif val > 0: return f'<span style="color: #B03A2E; font-style: italic;">{pct_str} <span style="{shadow}">▼</span></span>'
    return f'<span style="font-style: italic;">{pct_str}</span>'

def format_percentage_html_no_trend(val):
    if pd.isna(val) or isinstance(val, str) or val == '': return val
    return f'<span style="color: #000000; font-weight: bold; font-style: italic;">{val:.0%}</span>'

def color_index_cells(v):
    val_str = str(v)
    if val_str == 'HYU': return 'background-color: #e6f2ff;'  
    if val_str == 'KIA': return 'background-color: #ffe6e6;'  
    if val_str == 'GM': return 'background-color: #e6e6e6;'
    if val_str == 'DIRECT': return 'background-color: #e6f2ff;' 
    if val_str == 'COMM': return 'background-color: #f2f2f2;' 
    return ''

def apply_common_styles(styler, apply_hkmc_color=False, is_export=False):
    imp = "" if is_export else " !important"
    def style_row(row):
        row_str = str(row.name)
        base = ''
        if 'HYU_소계' in row_str: base = f'background-color: #e6f2ff{imp}; color: #002060; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;'
        elif 'KIA_소계' in row_str: base = f'background-color: #ffe6e6{imp}; color: #002060; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;'
        elif 'GM_소계' in row_str: base = f'background-color: #e6e6e6{imp}; color: #002060; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;'
        elif 'DIRECT_Subtotal_숨김' in row_str: base = f'background-color: #e6f2ff{imp}; color: #002060; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;'
        elif 'COMM_Subtotal_숨김' in row_str: base = f'background-color: #f2f2f2{imp}; color: #002060; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;'
        elif 'Unknown_Subtotal_숨김' in row_str: base = f'background-color: #ffffe0{imp}; color: #002060; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;'
        elif 'FC1 EX-RATE' in row_str: base = f'background-color: #e2efda{imp}; color: #002060; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;'
        elif 'GRAND_TOTAL_MERGE' in row_str or any(k in row_str for k in ['TTL', 'Total', '소계']): base = f'background-color: #ffffe0{imp}; color: #002060; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;'
        return [base] * len(row)
    
    styler.apply(style_row, axis=1)
    
    def highlight_total_index(val):
        l = str(val)
        if 'FC1 EX-RATE' in l: return f'background-color: #e2efda{imp}; color: #002060; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;'
        elif any(k in l for k in ['HYU_소계', 'KIA_소계', 'GM_소계', 'DIRECT_Subtotal_숨김', 'COMM_Subtotal_숨김', 'Unknown_Subtotal_숨김']): return f'background-color: #f0f0f0{imp}; color: #002060; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;'
        elif 'GRAND_TOTAL_MERGE' in l or any(k in l for k in ['TTL', 'Total', '소계']): return f'background-color: #ffffe0{imp}; color: #002060; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;'
        return ''
        
    for i in range(styler.index.nlevels): styler.map_index(highlight_total_index, axis=0, level=i)
    
    if apply_hkmc_color:
        if hasattr(styler, 'map_index'): styler.map_index(color_index_cells, axis=0, level=0)
        elif hasattr(styler, 'applymap_index'): styler.applymap_index(color_index_cells, axis=0, level=0)
        
    return styler

def optimize_html_headers(html_str, df):
    try:
        thead_start, thead_end = html_str.find('<thead>'), html_str.find('</thead>')
        if thead_start == -1 or thead_end == -1: return html_str
        thead_html = html_str[thead_start:thead_end+8]
        tr_matches = list(re.finditer(r'<tr[^>]*>(.*?)</tr>', thead_html, re.IGNORECASE | re.DOTALL))
        if len(tr_matches) < 2: return html_str
        
        row0_inner, row1_inner = tr_matches[0].group(1), tr_matches[1].group(1)
        ths0 = re.findall(r'<th[^>]*>.*?</th>', row0_inner, re.IGNORECASE | re.DOTALL)
        ths1 = re.findall(r'<th[^>]*>.*?</th>', row1_inner, re.IGNORECASE | re.DOTALL)
        
        index_names = list(df.index.names)
        for i in range(df.index.nlevels):
            if i < len(ths0) and i < len(ths1):
                name = str(index_names[i]) if index_names[i] is not None else ""
                width_style = "width: 50px !important; min-width: 50px !important; max-width: 50px !important; padding-left: 2px !important; padding-right: 2px !important; white-space: normal !important;" if name in ['Con.', 'SOP'] else "min-width: 80px;"
                ths0[i] = f'<th rowspan="2" style="vertical-align: middle !important; text-align: center !important; background-color: #002060 !important; color: white !important; border: 1px solid #8ea9db !important; {width_style}">{name}</th>'
                ths1[i] = ''
        new_thead = f"<thead>\n<tr>{''.join(ths0)}</tr>\n<tr>{''.join(ths1)}</tr>\n</thead>"
        return html_str[:thead_start] + new_thead + html_str[thead_end+8:]
    except Exception: return html_str

def post_process_html_styles(html_str):
    if '<tbody>' not in html_str: return html_str
    def process_row(match):
        row = match.group(0)
        if 'HYU_소계' in row: row = re.sub(r'^<tr', r'<tr class="total-row-hyu"', row.replace('HYU_소계', ''))
        elif 'KIA_소계' in row: row = re.sub(r'^<tr', r'<tr class="total-row-kia"', row.replace('KIA_소계', ''))
        elif 'GM_소계' in row: row = re.sub(r'^<tr', r'<tr class="total-row-gm"', row.replace('GM_소계', ''))
        elif 'DIRECT_Subtotal_숨김' in row: row = re.sub(r'^<tr', r'<tr class="total-row-direct"', row.replace('DIRECT_Subtotal_숨김', '')) 
        elif 'COMM_Subtotal_숨김' in row: row = re.sub(r'^<tr', r'<tr class="total-row-comm"', row.replace('COMM_Subtotal_숨김', '')) 
        elif 'Unknown_Subtotal_숨김' in row: row = re.sub(r'^<tr', r'<tr class="total-row"', row.replace('Unknown_Subtotal_숨김', '')) 
        elif 'FC1 EX-RATE' in row: row = re.sub(r'^<tr', r'<tr class="total-row-exrate"', row)
        elif 'GRAND_TOTAL_MERGE_START' in row:
            row = re.sub(r'^<tr', r'<tr class="total-row"', row)
            label_match = re.search(r'GRAND_TOTAL_MERGE_START(.*?)</th', row)
            if label_match:
                row = re.sub(r'<th[^>]*>GRAND_TOTAL_MERGE_START.*?</th>\s*<th[^>]*>GRAND_TOTAL_MERGE_DEL</th>\s*<th[^>]*>GRAND_TOTAL_MERGE_DEL</th>\s*<th[^>]*>GRAND_TOTAL_MERGE_DEL</th>',
                             f'<th colspan="4" style="text-align: left !important; padding-left: 15px !important; background-color: #ffffe0 !important; color: #002060 !important; font-weight: bold !important; border-top: 2px solid #8ea9db !important; border-bottom: 2px solid #8ea9db !important;">{label_match.group(1).strip()}</th>', row, flags=re.DOTALL)
        elif any(k in row for k in ['TTL', 'Total', '소계']): row = re.sub(r'^<tr', r'<tr class="total-row"', row)
        return row
    parts = html_str.split('<tbody>', 1)
    return parts[0] + '<tbody>' + re.sub(r'<tr[^>]*>.*?</tr>', process_row, parts[1], flags=re.DOTALL)

def to_excel_multiple(df_dict):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet_name, original_df in df_dict.items():
            df = original_df.copy()
            if isinstance(df.index, pd.MultiIndex):
                new_tuples = []
                for t in df.index:
                    new_t = list(t)
                    if isinstance(new_t[0], str) and 'GRAND_TOTAL_MERGE_START' in new_t[0]:
                        new_t[0] = new_t[0].replace('GRAND_TOTAL_MERGE_START', '')
                        for i in range(1, len(new_t)):
                            if new_t[i] == 'GRAND_TOTAL_MERGE_DEL': new_t[i] = ''
                    new_tuples.append(tuple(new_t))
                df.index = pd.MultiIndex.from_tuples(new_tuples, names=df.index.names)
            
            styler = df.style.format(lambda x: format_k_val(x) if isinstance(x, (int, float)) else x)
            apply_color = sheet_name in ["PE_HKMC_Summary", "PE_Biz_Detailed", "Core_Biz", "Biz_Type_Summary"]
            styler = apply_common_styles(styler, apply_hkmc_color=apply_color, is_export=True)
            
            styler.to_excel(writer, sheet_name=sheet_name[:31])
            worksheet = writer.sheets[sheet_name[:31]]
            for i in range(len(df.columns)): worksheet.set_column(i+1, i+1, 15)
    return output.getvalue()


# ==========================================
# 3. 데이터 로딩 및 집계 함수
# ==========================================
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
        df['BIZ Type'] = df['BIZ Type'].replace(['COMM', 'comm', 'COMMERCIAL', 'commercial'], 'COMM').fillna('Unknown')
        
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
    
    df = df.replace([np.inf, -np.inf], 0)
    df['Year'] = df['Year'].astype(int)
    df['Month'] = df['Month'].astype(int)
    df['Date'] = df['Date'].astype(str).str.replace('00:00:00', '').str.strip()
    return df

def build_summary_report(df_sub, index_cols, year, month, total_label="TTL (K.€)", index_names=None, sort_by_current_act=False, add_ex_rate=False):
    if df_sub.empty: return pd.DataFrame(), "", ""
    prev_year, prev_month = (year - 1, 12) if month == 1 else (year, month - 1)
    m_str, pm_str = MONTH_NAMES.get(month, f'{month}'), MONTH_NAMES.get(prev_month, f'{prev_month}')
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
    
    combined_dict, col_tuples = {}, [(col_prev, 'ACT')]
    for p in phases:
        for c in ['25 FC3', '26 FC1', 'ACT', 'ACHI %']: col_tuples.append((p, c))
        
    combined_dict[(col_prev, 'ACT')] = s_prev.reindex(idx).fillna(0) if not s_prev.empty else pd.Series(0, index=idx)
    for phase_name, data in zip(phases, [p_curr, p_ytd, p_ttl]):
        for c in ['25 FC3', '26 FC1', 'ACT']: combined_dict[(phase_name, c)] = data[c].reindex(idx).fillna(0) if not data.empty and c in data.columns else pd.Series(0, index=idx)
        num = pd.Series(combined_dict[(phase_name, 'ACT')])
        den = pd.Series(combined_dict[(phase_name, '26 FC1')])
        combined_dict[(phase_name, 'ACHI %')] = num.div(den).replace([np.inf, -np.inf], 0).fillna(0)
        
    final_df = pd.DataFrame(combined_dict)
    final_df.columns = pd.MultiIndex.from_tuples(col_tuples)
    final_df.index.names = current_index_names
    final_df = final_df.loc[(final_df.filter(like='ACT').sum(axis=1) != 0) | (final_df.filter(like='FC1').sum(axis=1) != 0)]
    
    if sort_by_current_act and (phase_curr, 'ACT') in final_df.columns: final_df = final_df.sort_values(by=(phase_curr, 'ACT'), ascending=False)
        
    if 'BIZ Type' in final_df.index.names:
        cats = pd.CategoricalDtype(categories=['DIRECT', 'COMM', 'Unknown'], ordered=True)
        try:
            final_df.index = final_df.index.set_levels(final_df.index.levels[0].astype(cats), level=0)
            final_df = final_df.sort_index(level=0)
        except: pass
        
    total_row = final_df.sum(numeric_only=True)
    for phase_name in phases:
        num, den = total_row.get((phase_name, 'ACT'), 0), total_row.get((phase_name, '26 FC1'), 0)
        total_row[(phase_name, 'ACHI %')] = num / den if den != 0 else 0
        
    if isinstance(final_df.index, pd.MultiIndex):
        total_idx = tuple([total_label] + [''] * (len(final_df.index.names)-1))
        t_df = pd.DataFrame([total_row], index=pd.MultiIndex.from_tuples([total_idx], names=final_df.index.names))
    else:
        t_df = pd.DataFrame([total_row], index=pd.Index([total_label], name=final_df.index.name))
        
    dfs_to_concat = [final_df, t_df]

    # --- [MODIFIED] PE Biz 전용 FC1 EX-RATE 로직 ---
    if add_ex_rate:
        def calc_ex_rate_act(df_target, target_year, target_month_list):
            total_val = 0
            for kox in df_target['KOx'].unique():
                kox_df = df_target[(df_target['KOx'] == kox) & (df_target['Year'] == target_year)]
                fc1_df = kox_df[kox_df['Desc.'] == '26 FC1']
                rate_col = 'EUR:KRW' if kox in ['KOKOR', 'KEM-KR'] else 'EUR:USD'
                
                fc1_rates = pd.to_numeric(fc1_df[rate_col], errors='coerce').replace(0, np.nan).dropna()
                fc1_rate = fc1_rates.iloc[0] if not fc1_rates.empty else np.nan
                
                for m in target_month_list:
                    m_act_df = kox_df[(kox_df['Desc.'] == 'ACT') & (kox_df['Month'] == m)]
                    act_sum = m_act_df['Rev. (€)'].sum()
                    if act_sum == 0: continue
                    
                    act_rates = pd.to_numeric(m_act_df[rate_col], errors='coerce').replace(0, np.nan).dropna()
                    act_rate = act_rates.iloc[0] if not act_rates.empty else np.nan
                    
                    if pd.notna(fc1_rate) and pd.notna(act_rate) and act_rate != 0:
                        total_val += act_sum * (fc1_rate / act_rate)
                    else:
                        total_val += act_sum
            return total_val
            
        ex_rate_row = pd.Series(0.0, index=total_row.index)
        ex_rate_row[(col_prev, 'ACT')] = calc_ex_rate_act(df_sub, prev_year, [prev_month])
        
        month_lists = {
            phase_curr: [month],
            phase_ytd: list(range(1, month + 1)),
            phase_ttl: list(range(1, 13))
        }
        
        for p_name in phases:
            ex_rate_row[(p_name, 'ACT')] = calc_ex_rate_act(df_sub, year, month_lists[p_name])
            ex_rate_row[(p_name, '26 FC1')] = total_row.get((p_name, '26 FC1'), 0)
            ex_rate_row[(p_name, '25 FC3')] = total_row.get((p_name, '25 FC3'), 0)
            den = total_row.get((p_name, '26 FC1'), 0)
            if den != 0:
                ex_rate_row[(p_name, 'ACHI %')] = ex_rate_row[(p_name, 'ACT')] / den
            else:
                ex_rate_row[(p_name, 'ACHI %')] = 0
                
        if isinstance(final_df.index, pd.MultiIndex):
            ex_idx = tuple(['FC1 EX-RATE'] + [''] * (len(final_df.index.names)-1))
            ex_df = pd.DataFrame([ex_rate_row], index=pd.MultiIndex.from_tuples([ex_idx], names=final_df.index.names))
        else:
            ex_df = pd.DataFrame([ex_rate_row], index=pd.Index(['FC1 EX-RATE'], name=final_df.index.name))
            
        dfs_to_concat.append(ex_df)
        
    return pd.concat(dfs_to_concat), col_prev, phase_curr

def get_biz_type_detailed_report(df, year, month):
    prev_year, prev_month = (year - 1, 12) if month == 1 else (year, month - 1)
    m_str, pm_str = MONTH_NAMES.get(month, f'{month}'), MONTH_NAMES.get(prev_month, f'{prev_month}')
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
            combined_dict[(phase_name, 'ACHI %')] = pd.Series(data.get('ACT', 0)).div(pd.Series(data.get('26 FC1', 0))).replace([np.inf, -np.inf], 0).fillna(0)
            
        combined = pd.DataFrame(combined_dict, index=p_m.index)
        if (phase_names[0], 'ACT') in combined.columns: combined = combined.sort_values(by=(phase_names[0], 'ACT'), ascending=False)
            
        subtotal = combined.sum(numeric_only=True)
        for p_name in phase_names:
            num, den = subtotal.get((p_name, 'ACT'), 0), subtotal.get((p_name, '26 FC1'), 0)
            subtotal[(p_name, 'ACHI %')] = num / den if den != 0 else 0
            
        results.append(combined)
        results.append(pd.DataFrame([subtotal], index=pd.MultiIndex.from_tuples([(biz, f'{biz}_Subtotal_숨김')], names=['BIZ Type', 'KOx'])))
        
    if not results: return pd.DataFrame(), phase_names[0]
    
    final_df = pd.concat(results)
    grand_total = final_df[~final_df.index.get_level_values(1).str.contains('Subtotal_숨김', na=False)].sum(numeric_only=True)
    for p_name in phase_names:
        num, den = grand_total.get((p_name, 'ACT'), 0), grand_total.get((p_name, '26 FC1'), 0)
        grand_total[(p_name, 'ACHI %')] = num / den if den != 0 else 0
        
    grand_row = pd.DataFrame([grand_total], index=pd.MultiIndex.from_tuples([('TTL (K.€)', ' ')], names=['BIZ Type', 'KOx']))
    return pd.concat([final_df, grand_row]), phase_names[0]

def get_biz_report(df, biz_type, year, month):
    prev_year, prev_month = (year - 1, 12) if month == 1 else (year, month - 1)
    df_biz = df[(df['Business Type'].str.contains(biz_type, case=False, na=False))].copy()
    m_str, pm_str = MONTH_NAMES.get(month, f'{month}'), MONTH_NAMES.get(prev_month, f'{prev_month}')
    p_curr, p_ytd, p_ttl, p_prev = f'{m_str}. {year}', f'YTD {m_str}. {year}', f'{year} TTL', f'{pm_str}. {prev_year}'
    
    brands = ['HYU', 'KIA'] if biz_type == 'Power' else ['HYU', 'KIA', 'GM']
    results = []
    
    for brand in brands:
        if brand == 'GM':
            brand_df = df_biz[df_biz['Project'] == 'GM'].copy()
            prev_mask = (df['Business Type'].str.contains(biz_type, case=False, na=False)) & (df['Year'] == prev_year) & (df['Month'] == prev_month) & (df['Project'] == 'GM')
            subtotal_dict = {(p_prev, 'ACT'): df[prev_mask & (df['Desc.'] == 'ACT')]['Rev. (€)'].sum() if not df[prev_mask].empty else 0.0}
            
            for i, d_df in enumerate([brand_df[brand_df['Month'] == month], brand_df[brand_df['Month'] <= month], brand_df]):
                phase_name = [p_curr, p_ytd, p_ttl][i]
                for c in ['25 FC3', '26 FC1', 'ACT']: 
                    subtotal_dict[(phase_name, c)] = d_df[d_df['Desc.'] == c]['Rev. (€)'].sum()
                subtotal_dict[(phase_name, 'ACHI %')] = subtotal_dict[(phase_name, 'ACT')] / subtotal_dict[(phase_name, '26 FC1')] if subtotal_dict[(phase_name, '26 FC1')] != 0 else 0.0
            
            res_df = pd.DataFrame([pd.Series(subtotal_dict)], index=pd.MultiIndex.from_tuples([(brand, f'{brand}_소계', '', '')], names=['Cust. GR', 'Project', 'Con.', 'SOP']))
            results.append(res_df)
            continue
        else:
            brand_df = df_biz[df_biz['Group 2'] == brand].copy()
            prev_mask = (df['Business Type'].str.contains(biz_type, case=False, na=False)) & (df['Year'] == prev_year) & (df['Month'] == prev_month) & (df['Group 2'] == brand)
            if brand_df.empty and df[prev_mask].empty: continue
            
            p_m = brand_df[brand_df['Month'] == month].pivot_table(index=['Project', 'Con.', 'SOP'], columns='Desc.', values='Rev. (€)', aggfunc='sum').fillna(0)
            p_y = brand_df[brand_df['Month'] <= month].pivot_table(index=['Project', 'Con.', 'SOP'], columns='Desc.', values='Rev. (€)', aggfunc='sum').fillna(0)
            p_fy = brand_df.pivot_table(index=['Project', 'Con.', 'SOP'], columns='Desc.', values='Rev. (€)', aggfunc='sum').fillna(0)
            p_prev_df = df[prev_mask].pivot_table(index=['Project', 'Con.', 'SOP'], columns='Desc.', values='Rev. (€)', aggfunc='sum').fillna(0)
            
            all_idx = set()
            for p in [p_prev_df, p_m, p_y, p_fy]:
                if not p.empty: all_idx.update(p.index.tolist())
            if not all_idx: continue
            
            idx = pd.MultiIndex.from_tuples(sorted(list(all_idx)), names=['Project', 'Con.', 'SOP'])
            p_prev_df = p_prev_df.reindex(idx, fill_value=0)
            p_m = p_m.reindex(idx, fill_value=0)
            p_y = p_y.reindex(idx, fill_value=0)
            p_fy = p_fy.reindex(idx, fill_value=0)
            
            # Others 그룹화 (Core Biz의 HYU, KIA만)
            if "Core" in biz_type and brand in ['HYU', 'KIA']:
                act_col = p_m['ACT'] if 'ACT' in p_m.columns else pd.Series(0, index=idx)
                top = act_col[act_col >= 10000].index
                def group_others(p):
                    if p.empty: return pd.DataFrame(columns=['25 FC3', '26 FC1', 'ACT']).reindex(pd.MultiIndex.from_tuples([], names=['Project', 'Con.', 'SOP']))
                    main = p.loc[p.index.isin(top)]
                    oth = p.loc[~p.index.isin(top)].sum().to_frame().T
                    oth.index = pd.MultiIndex.from_tuples([('Others', '', '')], names=['Project', 'Con.', 'SOP'])
                    return pd.concat([main, oth])
                p_m, p_y, p_fy, p_prev_df = group_others(p_m), group_others(p_y), group_others(p_fy), group_others(p_prev_df)
                idx = p_m.index
                
            combined_dict = {(p_prev, 'ACT'): p_prev_df['ACT'] if 'ACT' in p_prev_df.columns else pd.Series(0, index=idx)}
            for phase_name, data in [(p_curr, p_m), (p_ytd, p_y), (p_ttl, p_fy)]:
                for c in ['25 FC3', '26 FC1', 'ACT']:
                    combined_dict[(phase_name, c)] = data[c] if c in data.columns else pd.Series(0, index=idx)
                num = pd.Series(combined_dict[(phase_name, 'ACT')])
                den = pd.Series(combined_dict[(phase_name, '26 FC1')])
                combined_dict[(phase_name, 'ACHI %')] = num.div(den).replace([np.inf, -np.inf], 0).fillna(0)
            
            combined = pd.DataFrame(combined_dict, index=idx)
            
            if ('Others', '', '') in combined.index: 
                combined = pd.concat([combined.drop(index=('Others', '', '')).sort_values(by=(p_curr, 'ACT'), ascending=False), combined.loc[[('Others', '', '')]]])
            elif not combined.empty and (p_curr, 'ACT') in combined.columns: 
                combined = combined.sort_values(by=(p_curr, 'ACT'), ascending=False)
            
            subtotal = combined.sum(numeric_only=True) if not combined.empty else pd.Series(0, index=combined.columns)
            for p_name in [p_curr, p_ytd, p_ttl]:
                num, den = subtotal.get((p_name, 'ACT'), 0), subtotal.get((p_name, '26 FC1'), 0)
                subtotal[(p_name, 'ACHI %')] = num / den if den != 0 else 0
            
            combined.index = pd.MultiIndex.from_tuples([(brand, p, c, s) for p, c, s in combined.index], names=['Cust. GR', 'Project', 'Con.', 'SOP'])
            results.append(combined)
            results.append(pd.DataFrame([subtotal], index=pd.MultiIndex.from_tuples([(brand, f'{brand}_소계', '', '')], names=['Cust. GR', 'Project', 'Con.', 'SOP'])))
            
    if not results: return pd.DataFrame(), [p_curr, p_ytd, p_ttl]
    
    final_df = pd.concat(results)
    
    grand_total = final_df[final_df.index.get_level_values(1).str.contains('_소계', na=False)].sum(numeric_only=True)
    for p_name in [p_curr, p_ytd, p_ttl]:
        num, den = grand_total.get((p_name, 'ACT'), 0), grand_total.get((p_name, '26 FC1'), 0)
        grand_total[(p_name, 'ACHI %')] = num / den if den != 0 else 0
        
    grand_label = f'{BIZ_CONFIG.get(biz_type, biz_type)} Rev. TTL (K.€)'
    grand_row = pd.DataFrame([grand_total], index=pd.MultiIndex.from_tuples(
        [(f'GRAND_TOTAL_MERGE_START{grand_label}', 'GRAND_TOTAL_MERGE_DEL', 'GRAND_TOTAL_MERGE_DEL', 'GRAND_TOTAL_MERGE_DEL')], 
        names=['Cust. GR', 'Project', 'Con.', 'SOP']
    ))
    
    return pd.concat([final_df, grand_row]), [p_curr, p_ytd, p_ttl]

def build_trend_report(df, end_year, end_month):
    months, curr_y, curr_m = [], end_year, end_month
    for _ in range(12):
        months.append((curr_y, curr_m))
        curr_m -= 1
        if curr_m == 0: curr_m, curr_y = 12, curr_y - 1
    months.reverse() 
    
    df_act = df[df['Desc.'] == 'ACT']
    pivot_data = {}
    for y, m in months:
        col_name = f"{MONTH_NAMES[m]}.{str(y)[-2:]}"
        temp_df = df_act[(df_act['Year'] == y) & (df_act['Month'] == m)]
        pivot_data[col_name] = {'HYU': temp_df[temp_df['Group 2'] == 'HYU']['Rev. (€)'].sum(), 'KIA': temp_df[temp_df['Group 2'] == 'KIA']['Rev. (€)'].sum(), 'GM': temp_df[temp_df['Project'] == 'GM']['Rev. (€)'].sum()}
        
    trend_df = pd.DataFrame(pivot_data).reindex(['HYU', 'KIA', 'GM']).fillna(0)
    trend_df.loc['TTL (K.€)'] = trend_df.sum(numeric_only=True)
    trend_df.index.name = trend_df.columns.name = None 
    trend_df.columns = [col.strip() for col in trend_df.columns.values]
    return trend_df

def render_html_view(df, phase_curr, apply_color=False, is_biz=False):
    table_id = f"table_{uuid.uuid4().hex[:8]}"
    df_display = df.replace(0, '')
    format_dict = {col: (format_percentage_html_no_trend if 'TTL' in str(col[0]) else format_percentage_html) if 'ACHI' in str(col[1]) else format_k_val for col in df.columns}
            
    table_class = 'class="report-table biz-table"' if is_biz else 'class="report-table"'
    styler = df_display.style.format(format_dict, na_rep='').set_table_attributes(table_class)
    if not is_biz: styler.set_table_styles([{'selector': 'th, td', 'props': [('border-collapse', 'separate')]}, {'selector': 'tr', 'props': [('display', 'table-row')]}])
    
    styler.set_properties(subset=get_numeric_cols(df), **{'text-align': 'right'})
    styler = apply_common_styles(styler, apply_hkmc_color=apply_color)
    
    html_str = post_process_html_styles(optimize_html_headers(styler.to_html(), df))
    return f'{get_dynamic_highlight_css(table_id, df, phase_curr)}<div id="{table_id}" class="table-container">{html_str}</div>'

def render_biz_html_table(df, phase_curr, apply_color=False):
    table_id = f"table_{uuid.uuid4().hex[:8]}"
    df_display = df.replace(0, '')
    format_dict = {col: (format_percentage_html_no_trend if 'TTL' in str(col[0]) else format_percentage_html) if 'ACHI' in str(col[1]) else format_k_val for col in df.columns}
    styler = df_display.style.format(format_dict, na_rep='').set_table_attributes('class="report-table biz-table"')
    styler.set_properties(subset=get_numeric_cols(df), **{'text-align': 'right'})
    styler = apply_common_styles(styler, apply_hkmc_color=apply_color)
    html_str = post_process_html_styles(optimize_html_headers(styler.to_html(), df))
    return f'{get_dynamic_highlight_css(table_id, df, phase_curr)}<div id="{table_id}" class="table-container">{html_str}</div>'

def render_trend_html_table(df, apply_color=False):
    table_id = f"table_{uuid.uuid4().hex[:8]}"
    styler = df.replace(0, '').style.format({col: format_k_val for col in df.columns}, na_rep='').set_table_attributes('class="report-table trend-table"')
    styler.set_properties(**{'text-align': 'right'})
    html_str = post_process_html_styles(apply_common_styles(styler, apply_hkmc_color=apply_color).to_html())
    return f'{get_trend_highlight_css(table_id)}<div id="{table_id}" class="table-container">{html_str}</div>'

# ==========================================
# 4. 사이드바 및 메인 로직
# ==========================================
st.sidebar.title("📌 메뉴 설정")
selected_menu = st.sidebar.radio("원하시는 작업을 선택하세요.", ["매출 보고서", "판매가 조회"])
st.sidebar.divider()

if selected_menu == "매출 보고서":
    
    st.title("📊 매출 보고서 (Monthly Report)")
    uploaded_file = st.sidebar.file_uploader("월간 회의용 엑셀 데이터를 업로드하세요.", type=['xlsx', 'xls'], key="sales_uploader")

    if uploaded_file:
        raw_df = load_and_preprocess(uploaded_file)
        years = sorted(raw_df['Year'].unique())
        selected_year = st.sidebar.selectbox("연도", years, index=len(years)-1 if years else 0)
        selected_month = st.sidebar.selectbox("월", sorted(raw_df['Month'].unique()))
        
        reports_to_download = {}
        
        # ==========================================
        # 시각화 대시보드 (Plotly)
        # ==========================================
        st.markdown("### 📈 Visual Dashboard")
        col1, col2 = st.columns(2)
        
        with col1:
            df_trend_data = build_trend_report(raw_df, selected_year, selected_month)
            if not df_trend_data.empty:
                plot_df = df_trend_data.drop('TTL (K.€)').reset_index().melt(id_vars='index', var_name='Month', value_name='Rev')
                plot_df.rename(columns={'index': 'Brand'}, inplace=True)
                plot_df['Rev'] = plot_df['Rev'] / 1000.0  # K.€ 단위 변환
                
                fig1 = px.bar(plot_df, x='Month', y='Rev', color='Brand', 
                              title='12 Months Revenue Trend (K.€)',
                              color_discrete_map={'HYU': '#002060', 'KIA': '#8ea9db', 'GM': '#d9d9d9'})
                fig1.update_layout(plot_bgcolor='rgba(0,0,0,0)', yaxis=(dict(showgrid=True, gridcolor='#e6e6e6')),
                                   margin=dict(l=20, r=20, t=40, b=20), legend_title_text='')
                st.plotly_chart(fig1, use_container_width=True)

        with col2:
            df_chart2 = raw_df[(raw_df['Year'] == selected_year) & (raw_df['Month'] == selected_month) & (raw_df['Desc.'].isin(['26 FC1', 'ACT']))]
            if not df_chart2.empty:
                chart2_pivot = df_chart2.pivot_table(index='CPS', columns='Desc.', values='Rev. (€)', aggfunc='sum').fillna(0) / 1000.0
                chart2_pivot.reset_index(inplace=True)
                if '26 FC1' not in chart2_pivot.columns: chart2_pivot['26 FC1'] = 0
                if 'ACT' not in chart2_pivot.columns: chart2_pivot['ACT'] = 0
                
                proj_pivot = df_chart2.pivot_table(index=['CPS', 'Project'], columns='Desc.', values='Rev. (€)', aggfunc='sum').fillna(0).reset_index()
                if '26 FC1' not in proj_pivot.columns: proj_pivot['26 FC1'] = 0
                if 'ACT' not in proj_pivot.columns: proj_pivot['ACT'] = 0
                
                fc1_hover_texts = []
                act_hover_texts = []
                
                for cps in chart2_pivot['CPS']:
                    cps_data = proj_pivot[proj_pivot['CPS'] == cps]
                    
                    cps_fc1 = cps_data.sort_values(by='26 FC1', ascending=False).head(5)
                    if not cps_fc1.empty and cps_fc1['26 FC1'].sum() != 0:
                        lines = ["<br><b>[Top 5 Projects (FC1 Target)]</b>"]
                        for rank, (_, row) in enumerate(cps_fc1.iterrows(), 1):
                            lines.append(f"{rank}. {row['Project']} : {row['26 FC1'] / 1000.0:,.0f} K.€")
                        fc1_hover_texts.append("<br>".join(lines))
                    else:
                        fc1_hover_texts.append("<br><b>[No FC1 Data]</b>")
                        
                    cps_act = cps_data.sort_values(by='ACT', ascending=False).head(5)
                    if not cps_act.empty and cps_act['ACT'].sum() != 0:
                        lines = ["<br><b>[Top 5 Projects (ACT Actual)]</b>"]
                        for rank, (_, row) in enumerate(cps_act.iterrows(), 1):
                            act_val = row['ACT'] / 1000.0
                            fc1_val = row['26 FC1'] / 1000.0
                            diff = act_val - fc1_val
                            
                            pct_str = "N/A"
                            if fc1_val != 0:
                                pct_str = f"{(diff / fc1_val) * 100:+.1f}%"
                            elif act_val > 0:
                                pct_str = "+100.0%"
                                
                            if diff < 0:
                                diff_html = f"<span style='color: darkred;'>{diff:,.0f} K.€ ({pct_str})</span>"
                            elif diff > 0:
                                diff_html = f"<span style='color: darkblue;'>+{diff:,.0f} K.€ ({pct_str})</span>"
                            else:
                                diff_html = "0 K.€ (0%)"
                                
                            lines.append(f"{rank}. {row['Project']} : {act_val:,.0f} K.€ | {diff_html}")
                        act_hover_texts.append("<br>".join(lines))
                    else:
                        act_hover_texts.append("<br><b>[No ACT Data]</b>")
                
                chart2_pivot['fc1_hover'] = fc1_hover_texts
                chart2_pivot['act_hover'] = act_hover_texts
                
                fig2 = go.Figure(data=[
                    go.Bar(name='FC1', 
                           x=chart2_pivot['CPS'], 
                           y=chart2_pivot['26 FC1'], 
                           marker_color='#c7c7c7',
                           text=chart2_pivot['26 FC1'].apply(lambda x: f'{x:,.0f}'),
                           textposition='inside', textfont=dict(color='white', size=12, weight='bold'),
                           customdata=chart2_pivot['fc1_hover'],
                           hovertemplate="<b>CPS: %{x}</b><br>FC1: %{y:,.0f} K.€%{customdata}<extra></extra>"),
                    go.Bar(name='ACT', 
                           x=chart2_pivot['CPS'], 
                           y=chart2_pivot['ACT'], 
                           marker_color='#1f77b4',
                           text=chart2_pivot['ACT'].apply(lambda x: f'{x:,.0f}'),
                           textposition='inside', textfont=dict(color='white', size=12, weight='bold'),
                           customdata=chart2_pivot['act_hover'],
                           hovertemplate="<b>CPS: %{x}</b><br>ACT: %{y:,.0f} K.€%{customdata}<extra></extra>")
                ])
                fig2.update_layout(barmode='group', title=f'[{MONTH_NAMES.get(selected_month)}] FC1 vs ACT by CPS (K.€)',
                                   plot_bgcolor='rgba(0,0,0,0)', yaxis=(dict(showgrid=True, gridcolor='#e6e6e6', visible=False)),
                                   margin=dict(l=20, r=20, t=50, b=20),
                                   legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
                st.plotly_chart(fig2, use_container_width=True)
                
        st.markdown("---")

        # ==========================================
        # 테이블 영역
        # ==========================================
        st.subheader("📌 Sales Revenue Trend")
        if not df_trend_data.empty:
            st.markdown(render_trend_html_table(df_trend_data, apply_color=False), unsafe_allow_html=True)
            reports_to_download["12M_Trend_Report"] = df_trend_data
            
        st.subheader("📌 CPS별 매출액 요약")
        df_cps, p_col, c_col = build_summary_report(raw_df, ['CPS'], selected_year, selected_month, 'TTL (K.€)')
        if not df_cps.empty: 
            st.markdown(render_html_view(df_cps, c_col, apply_color=False), unsafe_allow_html=True)
            reports_to_download["CPS_Summary"] = df_cps

        st.subheader("📌 PE Item 매출액 요약")
        df_item_raw = raw_df[raw_df['Item'].isin(['ICCU1', 'ICCU2', 'VCMS'])]
        df_item, p_col, c_col = build_summary_report(df_item_raw, ['Item'], selected_year, selected_month, 'TTL (K.€)', index_names=['CPS'], sort_by_current_act=True)
        if not df_item.empty: 
            st.markdown(render_html_view(df_item, c_col, apply_color=False), unsafe_allow_html=True)
            reports_to_download["Item_Summary"] = df_item

        st.subheader("📌 DIRECT & COMMISSION 매출액 요약")
        df_biz_type, c_col = get_biz_type_detailed_report(raw_df, selected_year, selected_month)
        if not df_biz_type.empty: 
            st.markdown(render_html_view(df_biz_type, c_col, apply_color=True), unsafe_allow_html=True)
            reports_to_download["Biz_Type_Summary"] = df_biz_type

        # --- PE Biz 전용 FC1 EX-RATE 로직 적용 ---
        st.subheader("📌 Sales Revenue: Power Electronics")
        df_pe_raw = raw_df[raw_df['Business Type'].str.contains("Power", case=False, na=False)].copy()
        if not df_pe_raw.empty:
            df_pe_raw['Cust. GR'] = df_pe_raw['Group 2'].replace({'HYU': 'HKMC', 'KIA': 'HKMC'})
            df_pe_summary, p_col, c_col = build_summary_report(
                df_pe_raw[df_pe_raw['Cust. GR'] == 'HKMC'], 
                ['Cust. GR', 'KOx'], 
                selected_year, selected_month, 
                total_label='PE Biz Rev. TTL (K.€)', 
                sort_by_current_act=True,
                add_ex_rate=True) 
            if not df_pe_summary.empty:
                st.markdown(render_html_view(df_pe_summary, c_col, apply_color=True), unsafe_allow_html=True)
                reports_to_download["PE_HKMC_Summary"] = df_pe_summary

        for filter_key, display_name in BIZ_CONFIG.items():
            st.subheader(f"📌 Sales Revenue: {display_name}")
            df_biz, phase_names = get_biz_report(raw_df, filter_key, selected_year, selected_month)
            if not df_biz.empty:
                st.markdown(render_biz_html_table(df_biz, phase_names[0], apply_color=True), unsafe_allow_html=True)
                reports_to_download[f"{display_name}_Detailed"] = df_biz

        if reports_to_download:
            st.write("---")
            st.download_button("📥 월간회의 자료용 엑셀 다운로드", data=to_excel_multiple(reports_to_download), file_name=f"Monthly_Closing_Report_{selected_year}_{selected_month:02d}.xlsx", use_container_width=True)
    else:
        st.info("👈 좌측 메뉴에서 '월간 회의용 엑셀 파일'을 업로드하시면 요약 리포트가 생성됩니다.")

# ==========================================
# 5. 판매가 조회 메뉴 로직 
# ==========================================
elif selected_menu == "판매가 조회":
    st.title("💰 판매가 조회 (Price Lookup)")
    
    uploaded_txt_files = st.sidebar.file_uploader("판매가 TXT 파일들을 업로드하세요.", type=['txt'], accept_multiple_files=True, key="price_uploader")
    
    if uploaded_txt_files:
        st.success(f"📂 총 {len(uploaded_txt_files)}개의 파일을 처리 중입니다.")
        parsed_data = []
        for txt_file in uploaded_txt_files:
            try: content = txt_file.getvalue().decode('utf-8')
            except UnicodeDecodeError: content = txt_file.getvalue().decode('cp949')
            
            sales_org, distr_channel, current_customer = "", "", ""
            for line in content.split('\n'):
                line = line.strip('\r')
                if line.startswith("Sales Org."):
                    if m := re.search(r'Sales Org\.\s+(\d+)', line): sales_org = m.group(1)
                elif line.startswith("Distr. Channel"):
                    if m := re.search(r'Distr\. Channel\s+(\d+)', line): distr_channel = m.group(1)
                elif line.startswith("\t"):
                    parts = [p.strip() for p in line.split('\t')]
                    if len(parts) > 1:
                        if parts[1].isdigit() and parts[2] == '': current_customer = parts[1]
                        elif len(parts) >= 16 and parts[2] in ['YPR0', 'ZADD']:
                            try: amt, per = float(parts[10].replace(',', '')), float(parts[12].replace(',', ''))
                            except: amt, per = 0, 0
                            price = int(amt / per) if (amt / per).is_integer() else round(amt / per, 2) if per != 0 else ""
                            parsed_data.append({
                                "Sales Org.": sales_org, "Distr. Channel": distr_channel, "Customer": current_customer,
                                "CnTy": parts[1], "Condition Type": parts[2], "Material": parts[5], "Material Description": parts[6],
                                "Amount": amt, "Unit": parts[11], "Unit Size": per, "UoM": parts[13],
                                "Valid From": parts[14], "Valid to": parts[15], "Price": price
                            })
        
        if parsed_data:
            df_final = pd.DataFrame(parsed_data)
            st.subheader("📋 정제된 단가 데이터")
            st.dataframe(df_final, use_container_width=True)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, sheet_name="Sheet2", index=False, startrow=2)
                ws = writer.sheets["Sheet2"]
                ws.set_column('A:B', 12); ws.set_column('C:C', 10); ws.set_column('G:G', 40)
                
            st.download_button("📥 통합 결과 엑셀 다운로드", data=output.getvalue(), file_name="결과.xlsx", use_container_width=True)
        else:
            st.warning("분석할 수 있는 데이터가 없습니다. txt 파일 형식을 확인해주세요.")
