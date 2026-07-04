import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import uuid
import datetime
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
h1 { font-size: 1rem !important; margin-bottom: 0.3rem !important; padding-bottom: 0 !important; }
h3 { font-size: 1.1rem !important; margin-top: 0.4rem !important; margin-bottom: 0.3rem !important; color: #002060 !important; }
hr { margin-top: 0.3rem !important; margin-bottom: 0.3rem !important; border: none !important; border-top: 1px solid #d9d9d9 !important; }
div[data-testid="stForm"] { margin-top: 0rem !important; margin-bottom: 0.4rem !important; padding: 1rem !important; }
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
# 2. 모든 함수 정의
# ==========================================
def get_trend_highlight_css(table_id):
    return f"<style>#{table_id} thead tr:nth-child(1) th:last-child {{ border-top: 4px solid #c00000 !important; border-left: 4px solid #c00000 !important; border-right: 4px solid #c00000 !important; }} #{table_id} tbody td:last-child {{ border-left: 4px solid #c00000 !important; border-right: 4px solid #c00000 !important; }} #{table_id} tbody tr:last-child td:last-child {{ border-bottom: 4px solid #c00000 !important; }}</style>"

def get_dynamic_highlight_css(table_id, df, highlight_phase):
    if not highlight_phase: return ""
    cols = list(df.columns)
    start_col, end_col = -1, -1
    act_col_idx = -1
    level0_cols = []
    for i, col in enumerate(cols):
        c0 = col[0] if isinstance(col, tuple) else str(col)
        if not level0_cols or level0_cols[-1] != c0: level0_cols.append(c0)
        if c0 == highlight_phase:
            if start_col == -1: start_col = i
            end_col = i
            if isinstance(col, tuple) and len(col) > 1 and col[1] == 'ACT': act_col_idx = i
            elif str(col) == 'ACT': act_col_idx = i
    if start_col == -1: return ""
    
    num_indices = df.index.nlevels
    target_th_row0 = num_indices + level0_cols.index(highlight_phase) + 1
    css = f"<style>\n#{table_id} thead tr:nth-child(1) th:nth-child({target_th_row0}) {{ border-top: 5px solid #c00000 !important; border-left: 5px solid #c00000 !important; border-right: 5px solid #c00000 !important; }}\n"
    css += f"#{table_id} tbody td:nth-of-type({start_col + 1}) {{ border-left: 5px solid #c00000 !important; }}\n"
    css += f"#{table_id} tbody td:nth-of-type({end_col + 1}) {{ border-right: 5px solid #c00000 !important; }}\n</style>"
    return css

def get_numeric_cols(df): return [col for col in df.columns if any(x in str(col) for x in ['FC3', 'FC1', 'ACT', 'ACHI'])]

def format_k_val(val):
    if pd.isna(val) or isinstance(val, str) or val == '': return val
    v = val / 1_000.0
    rounded_int = int(round(v, 0))
    if rounded_int == 0: return str(round(v, 2)) if round(v, 2) != 0 else "0"
    return f"{rounded_int:,}"

def format_percentage_html(val):
    if pd.isna(val) or isinstance(val, str) or val == '': return val
    pct_str = f"{val:.0%}"
    if 0.95 <= val <= 1.0: return f'<span style="color: #404040; font-style: italic;">{pct_str}</span>'
    elif val > 1.0: return f'<span style="color: #145A32; font-style: italic;">{pct_str} ▲</span>'
    elif val > 0: return f'<span style="color: #B03A2E; font-style: italic;">{pct_str} ▼</span>'
    return f'<span style="font-style: italic;">{pct_str}</span>'

def format_percentage_html_no_trend(val):
    return f'<span style="color: #000000; font-weight: bold; font-style: italic;">{val:.0%}</span>' if pd.notna(val) else val

def apply_common_styles(styler, apply_hkmc_color=False, is_export=False):
    imp = "" if is_export else " !important"
    def style_row(row):
        row_str = str(row.name)
        base = ''
        if 'HYU_소계' in row_str: base = f'background-color: #e6f2ff{imp}; color: #002060; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;'
        elif 'KIA_소계' in row_str: base = f'background-color: #ffe6e6{imp}; color: #002060; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;'
        elif 'GM_소계' in row_str: base = f'background-color: #e6e6e6{imp}; color: #002060; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;'
        elif 'GRAND_TOTAL_MERGE' in row_str or any(k in row_str for k in ['TTL', 'Total', '소계']): base = f'background-color: #ffffe0{imp}; color: #002060; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;'
        return [base] * len(row)
    styler.apply(style_row, axis=1)
    if hasattr(styler, 'apply_index'):
        def style_row_index(idx):
            res = []
            for label in idx:
                l = str(label)
                if 'HYU_소계' in l: res.append(f'background-color: #e6f2ff{imp}; color: #e6f2ff; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;')
                elif 'KIA_소계' in l: res.append(f'background-color: #ffe6e6{imp}; color: #ffe6e6; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;')
                elif 'GM_소계' in l: res.append(f'background-color: #e6e6e6{imp}; color: #e6e6e6; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;')
                elif 'GRAND_TOTAL_MERGE' in l or any(k in l for k in ['TTL', 'Total', '소계']): res.append(f'background-color: #ffffe0{imp}; color: #002060; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;')
                else: res.append('')
            return res
        styler.apply_index(style_row_index, axis=0)
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
                ths0[i] = f'<th rowspan="2" style="vertical-align: middle !important; text-align: center !important; background-color: #002060 !important; color: white !important; border: 1px solid #8ea9db !important;">{name}</th>'
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
        elif any(k in row for k in ['TTL', 'Total', '소계']): row = re.sub(r'^<tr', r'<tr class="total-row"', row)
        return row
    parts = html_str.split('<tbody>', 1)
    return parts[0] + '<tbody>' + re.sub(r'<tr[^>]*>.*?</tr>', process_row, parts[1], flags=re.DOTALL)

def to_excel_multiple(df_dict):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet_name, original_df in df_dict.items():
            df = original_df.copy()
            styler = df.style.format(lambda x: format_k_val(x) if isinstance(x, (int, float)) else x)
            styler = apply_common_styles(styler, is_export=True)
            styler.to_excel(writer, sheet_name=sheet_name[:31])
            worksheet = writer.sheets[sheet_name[:31]]
            for i in range(len(df.columns)): worksheet.set_column(i+1, i+1, 15)
    return output.getvalue()

def parse_sop_date(val):
    if pd.isna(val) or str(val).strip() == '' or str(val).strip().lower() == 'nan': return ""
    if isinstance(val, (datetime.datetime, datetime.date, pd.Timestamp)): return val.strftime("%Y.%m.01")
    val_str = str(val).strip()
    m = re.search(r'\b(\d{1,2})\.(\d{4})\b', val_str)
    if m: return f"{m.group(2)}.{m.group(1).zfill(2)}.01"
    try: return pd.to_datetime(val_str).strftime("%Y.%m.01")
    except: return val_str

@st.cache_data
def load_and_preprocess(file):
    xl = pd.ExcelFile(file)
    sheets = xl.sheet_names
    df = pd.read_excel(xl, sheet_name=sheets[0], header=4).iloc[:, :26]
    df.columns = ['Year', 'Month', 'Desc.', 'Date', 'STP', 'Customer', 'LK No.', "Q'ty", 
                  'Rev. ($)', 'Rev. (€)', 'Rev. ₩', 'BIZ Type', 'Group 1', 'Group 2', 
                  'Project', 'PF', 'Item', 'Source', 'KOx', 'Memo', 'CPS', 
                  'EUR:USD', 'EUR:KRW', 'Business Type', 'Curr.', 'Con.']
    if 'BIZ Type' in df.columns: df['BIZ Type'] = df['BIZ Type'].replace(['COMM', 'comm', 'COMMERCIAL', 'commercial'], 'COMM').fillna('Unknown')
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

    if add_ex_rate:
        def calc_ex_rate_act(df_target, target_year, target_month_list):
            total_val = 0
            for kox in df_target['KOx'].unique():
                if pd.isna(kox): continue
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
                    if pd.notna(fc1_rate) and pd.notna(act_rate) and act_rate != 0: total_val += act_sum * (fc1_rate / act_rate)
                    else: total_val += act_sum
            return total_val
            
        ex_rate_row = pd.Series(np.nan, index=total_row.index)
        month_lists = { phase_curr: [month], phase_ytd: list(range(1, month + 1)), phase_ttl: list(range(1, 13)) }
        for p_name in phases:
            act_val = calc_ex_rate_act(df_sub, year, month_lists[p_name])
            ex_rate_row[(p_name, 'ACT')] = act_val
            den = total_row.get((p_name, '26 FC1'), 0)
            ex_rate_row[(p_name, 'ACHI %')] = act_val / den if den != 0 else 0
                
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
        if (phase_names[0], '26 FC1') in combined.columns and (phase_names[0], 'ACT') in combined.columns:
            fc1_curr = combined[(phase_names[0], '26 FC1')].abs()
            act_curr = combined[(phase_names[0], 'ACT')].abs()
            combined = combined[(fc1_curr >= 0.01) | (act_curr >= 0.01)]
            
        if combined.empty: continue
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

    def calc_ex_rate_act(df_target, target_year, target_month_list):
        total_val = 0
        for kox in df_target['KOx'].unique():
            if pd.isna(kox): continue
            kox_df = df_target[(df_target['KOx'] == kox) & (df_target['Year'] == target_year)]
            fc1_df = kox_df[kox_df['Desc.'] == '26 FC1']
            rate_col = 'EUR:KRW' if kox in ['KOKOR', 'KEM-KR'] else 'EUR:USD'
            fc1_rates = pd.to_numeric(fc1_df[rate_col], errors='coerce').replace(0, np.nan).dropna()
            fc1_rate = fc1_rates.iloc[0] if not fc1_rates.empty else np.nan
            for m_idx in target_month_list:
                m_act_df = kox_df[(kox_df['Desc.'] == 'ACT') & (kox_df['Month'] == m_idx)]
                act_sum = m_act_df['Rev. (€)'].sum()
                if act_sum == 0: continue
                act_rates = pd.to_numeric(m_act_df[rate_col], errors='coerce').replace(0, np.nan).dropna()
                act_rate = act_rates.iloc[0] if not act_rates.empty else np.nan
                if pd.notna(fc1_rate) and pd.notna(act_rate) and act_rate != 0: total_val += act_sum * (fc1_rate / act_rate)
                else: total_val += act_sum
        return total_val

    ex_rate_row = pd.Series(np.nan, index=grand_total.index)
    valid_biz_df = df[df['BIZ Type'].isin(biz_categories)]
    month_lists = { phase_names[0]: [month], phase_names[1]: list(range(1, month + 1)), phase_names[2]: list(range(1, 13)) }
    for p_name in phase_names:
        act_val = calc_ex_rate_act(valid_biz_df, year, month_lists[p_name])
        ex_rate_row[(p_name, 'ACT')] = act_val
        den = grand_total.get((p_name, '26 FC1'), 0)
        ex_rate_row[(p_name, 'ACHI %')] = act_val / den if den != 0 else 0
        
    ex_df = pd.DataFrame([ex_rate_row], index=pd.MultiIndex.from_tuples([('FC1 EX-RATE', ' ')], names=['BIZ Type', 'KOx']))
    return pd.concat([final_df, grand_row, ex_df]), phase_names[0]

def get_core_biz_summary_report(df, year, month):
    prev_year, prev_month = (year - 1, 12) if month == 1 else (year, month - 1)
    m_str, pm_str = MONTH_NAMES.get(month, f'{month}'), MONTH_NAMES.get(prev_month, f'{prev_month}')
    phase_names = [f'{m_str}. {year}', f'YTD {m_str}. {year}', f'{year} TTL']
    prev_phase_name = f'{pm_str}. {prev_year}'

    df_core = df[df['Business Type'].str.contains("Core", case=False, na=False)].copy()
    if df_core.empty: return pd.DataFrame(), phase_names[0]

    df_core['Cust. GR'] = df_core['Group 1'].replace({'HYU': 'HKMC', 'KIA': 'HKMC', 'GM': 'GM'})

    results = []
    unique_grs = []
    for g in ['HKMC', 'GM']:
        if g in df_core['Cust. GR'].values: unique_grs.append(g)
    for g in df_core['Cust. GR'].dropna().unique():
        if g not in ['HKMC', 'GM']: unique_grs.append(g)

    for gr in unique_grs:
        gr_df = df_core[df_core['Cust. GR'] == gr]
        p_m = gr_df[(gr_df['Year'] == year) & (gr_df['Month'] == month)].pivot_table(index=['Cust. GR', 'KOx'], columns='Desc.', values='Rev. (€)', aggfunc='sum').fillna(0)
        p_y = gr_df[(gr_df['Year'] == year) & (gr_df['Month'] <= month)].pivot_table(index=['Cust. GR', 'KOx'], columns='Desc.', values='Rev. (€)', aggfunc='sum').fillna(0)
        p_fy = gr_df[gr_df['Year'] == year].pivot_table(index=['Cust. GR', 'KOx'], columns='Desc.', values='Rev. (€)', aggfunc='sum').fillna(0)
        p_prev = gr_df[(gr_df['Year'] == prev_year) & (gr_df['Month'] == prev_month)].pivot_table(index=['Cust. GR', 'KOx'], columns='Desc.', values='Rev. (€)', aggfunc='sum').fillna(0)

        all_idx = set(p_m.index.tolist() + p_y.index.tolist() + p_fy.index.tolist() + p_prev.index.tolist())
        if not all_idx: continue

        order_map = {'KOASIA': 1, 'KOKOR': 2, 'KOIN': 3, 'KOA': 4}
        idx_list = sorted(list(all_idx), key=lambda x: order_map.get(x[1], 99))
        idx = pd.MultiIndex.from_tuples(idx_list, names=['Cust. GR', 'KOx'])

        combined_dict = {(prev_phase_name, 'ACT'): p_prev.reindex(idx, fill_value=0).get('ACT', pd.Series(0, index=idx))}
        for phase_name, data in [(phase_names[0], p_m), (phase_names[1], p_y), (phase_names[2], p_fy)]:
            data = data.reindex(idx, fill_value=0)
            for c in ['25 FC3', '26 FC1', 'ACT']: combined_dict[(phase_name, c)] = data.get(c, pd.Series(0, index=idx))
            num = pd.Series(combined_dict[(phase_name, 'ACT')])
            den = pd.Series(combined_dict[(phase_name, '26 FC1')])
            combined_dict[(phase_name, 'ACHI %')] = num.div(den).replace([np.inf, -np.inf], 0).fillna(0)

        combined = pd.DataFrame(combined_dict, index=idx)
        if (phase_names[0], '26 FC1') in combined.columns and (phase_names[0], 'ACT') in combined.columns:
            fc1_curr = combined[(phase_names[0], '26 FC1')].abs()
            act_curr = combined[(phase_names[0], 'ACT')].abs()
            combined = combined[(fc1_curr >= 0.01) | (act_curr >= 0.01)]
            
        if combined.empty: continue

        subtotal = combined.sum(numeric_only=True)
        for p_name in phase_names:
            den = subtotal.get((p_name, '26 FC1'), 0)
            subtotal[(p_name, 'ACHI %')] = subtotal.get((p_name, 'ACT'), 0) / den if den != 0 else 0

        results.append(combined)
        subtotal_idx_name = f'HYU_소계' if gr == 'HKMC' else f'GM_소계'
        results.append(pd.DataFrame([subtotal], index=pd.MultiIndex.from_tuples([('', subtotal_idx_name)], names=['Cust. GR', 'KOx'])))

    if not results: return pd.DataFrame(), phase_names[0]

    final_df = pd.concat(results)
    grand_total = final_df[final_df.index.get_level_values(1).str.contains('소계', na=False)].sum(numeric_only=True)
    for p_name in phase_names:
        den = grand_total.get((p_name, '26 FC1'), 0)
        grand_total[(p_name, 'ACHI %')] = grand_total.get((p_name, 'ACT'), 0) / den if den != 0 else 0

    grand_row = pd.DataFrame([grand_total], index=pd.MultiIndex.from_tuples([('Core Biz Rev. TTL (K.€)', ' ')], names=['Cust. GR', 'KOx']))

    def calc_ex_rate_act(df_target, target_year, target_month_list):
        total_val = 0
        for kox in df_target['KOx'].unique():
            if pd.isna(kox): continue
            kox_df = df_target[(df_target['KOx'] == kox) & (df_target['Year'] == target_year)]
            fc1_df = kox_df[kox_df['Desc.'] == '26 FC1']
            rate_col = 'EUR:KRW' if kox in ['KOKOR', 'KEM-KR'] else 'EUR:USD'
            fc1_rates = pd.to_numeric(fc1_df[rate_col], errors='coerce').replace(0, np.nan).dropna()
            fc1_rate = fc1_rates.iloc[0] if not fc1_rates.empty else np.nan
            for m_idx in target_month_list:
                m_act_df = kox_df[(kox_df['Desc.'] == 'ACT') & (kox_df['Month'] == m_idx)]
                act_sum = m_act_df['Rev. (€)'].sum()
                if act_sum == 0: continue
                act_rates = pd.to_numeric(m_act_df[rate_col], errors='coerce').replace(0, np.nan).dropna()
                act_rate = act_rates.iloc[0] if not act_rates.empty else np.nan
                if pd.notna(fc1_rate) and pd.notna(act_rate) and act_rate != 0: total_val += act_sum * (fc1_rate / act_rate)
                else: total_val += act_sum
        return total_val

    ex_rate_row = pd.Series(np.nan, index=grand_total.index)
    month_lists = { phase_names[0]: [month], phase_names[1]: list(range(1, month + 1)), phase_names[2]: list(range(1, 13)) }
    for p_name in phase_names:
        act_val = calc_ex_rate_act(df_core, year, month_lists[p_name])
        ex_rate_row[(p_name, 'ACT')] = act_val
        den = grand_total.get((p_name, '26 FC1'), 0)
        ex_rate_row[(p_name, 'ACHI %')] = act_val / den if den != 0 else 0

    ex_df = pd.DataFrame([ex_rate_row], index=pd.MultiIndex.from_tuples([('FC1 EX-RATE', ' ')], names=['Cust. GR', 'KOx']))
    return pd.concat([final_df, grand_row, ex_df]), phase_names[0]

def get_biz_report(df, biz_type, year, month):
    if month == 1: prev_year, prev_month = year - 1, 12
    else: prev_year, prev_month = year, month - 1
    
    df_biz = df[(df['Business Type'].str.contains(biz_type, case=False, na=False)) & (df['Year'] == year)].copy()
    m_str, pm_str = MONTH_NAMES.get(month, f'{month}'), MONTH_NAMES.get(prev_month, f'{prev_month}')
    phase_names = [f'{m_str}. {year}', f'YTD {m_str}. {year}', f'{year} TTL']
    prev_phase_name = f'{pm_str}. {prev_year}'
    
    brands = ['HYU', 'KIA', 'GM']
    if biz_type == 'Power': brands = ['HYU', 'KIA']
    
    results = []
    for brand in brands:
        if brand == 'GM':
            brand_df = df_biz[df_biz['Project'] == 'GM'].copy()
            prev_mask = (df['Business Type'].str.contains(biz_type, case=False, na=False)) & (df['Year'] == prev_year) & (df['Month'] == prev_month) & (df['Project'] == 'GM')
            subtotal_dict = {}
            prev_act = df[prev_mask & (df['Desc.'] == 'ACT')]['Rev. (€)'].sum() if not df[prev_mask].empty else 0.0
            subtotal_dict[(prev_phase_name, 'ACT')] = prev_act
            
            df_m = brand_df[brand_df['Month'] == month]
            for c in ['25 FC3', '26 FC1', 'ACT']: subtotal_dict[(phase_names[0], c)] = df_m[df_m['Desc.'] == c]['Rev. (€)'].sum()
            subtotal_dict[(phase_names[0], 'ACHI %')] = subtotal_dict[(phase_names[0], 'ACT')] / subtotal_dict[(phase_names[0], '26 FC1')] if subtotal_dict[(phase_names[0], '26 FC1')] != 0 else 0.0
            
            df_y = brand_df[brand_df['Month'] <= month]
            for c in ['25 FC3', '26 FC1', 'ACT']: subtotal_dict[(phase_names[1], c)] = df_y[df_y['Desc.'] == c]['Rev. (€)'].sum()
            subtotal_dict[(phase_names[1], 'ACHI %')] = subtotal_dict[(phase_names[1], 'ACT')] / subtotal_dict[(phase_names[1], '26 FC1')] if subtotal_dict[(phase_names[1], '26 FC1')] != 0 else 0.0
            
            for c in ['25 FC3', '26 FC1', 'ACT']: subtotal_dict[(phase_names[2], c)] = brand_df[brand_df['Desc.'] == c]['Rev. (€)'].sum()
            subtotal_dict[(phase_names[2], 'ACHI %')] = subtotal_dict[(phase_names[2], 'ACT')] / subtotal_dict[(phase_names[2], '26 FC1')] if subtotal_dict[(phase_names[2], '26 FC1')] != 0 else 0.0
            
            subtotal = pd.Series(subtotal_dict)
            results.append(pd.DataFrame([subtotal], index=pd.MultiIndex.from_tuples([(brand, f'{brand}_소계', '', '')], names=['Cust. GR', 'Project', 'Con.', 'SOP'])))
            continue
        else:
            brand_df = df_biz[df_biz['Group 2'] == brand].copy()
            prev_mask = (df['Business Type'].str.contains(biz_type, case=False, na=False)) & (df['Year'] == prev_year) & (df['Month'] == prev_month) & (df['Group 2'] == brand)
            if brand_df.empty and df[prev_mask].empty: continue
            
            p_m = brand_df[brand_df['Month'] == month].pivot_table(index=['Project', 'Con.', 'SOP'], columns='Desc.', values='Rev. (€)', aggfunc='sum').fillna(0)
            p_y = brand_df[brand_df['Month'] <= month].pivot_table(index=['Project', 'Con.', 'SOP'], columns='Desc.', values='Rev. (€)', aggfunc='sum').fillna(0)
            p_fy = brand_df.pivot_table(index=['Project', 'Con.', 'SOP'], columns='Desc.', values='Rev. (€)', aggfunc='sum').fillna(0)
            p_prev = df[prev_mask].pivot_table(index=['Project', 'Con.', 'SOP'], columns='Desc.', values='Rev. (€)', aggfunc='sum').fillna(0)
            
            all_idx = set()
            for p in [p_prev, p_m, p_y, p_fy]:
                if not p.empty: all_idx.update(p.index.tolist())
            if not all_idx: continue
            
            idx = pd.MultiIndex.from_tuples(sorted(list(all_idx)), names=['Project', 'Con.', 'SOP'])
            p_prev = p_prev.reindex(idx, fill_value=0)
            p_m = p_m.reindex(idx, fill_value=0)
            p_y = p_y.reindex(idx, fill_value=0)
            p_fy = p_fy.reindex(idx, fill_value=0)
            
            if "Core" in biz_type and brand in ['HYU', 'KIA']:
                act_col = p_m['ACT'] if 'ACT' in p_m.columns else pd.Series(0, index=idx)
                top = act_col[act_col >= 10000].index
                def group_others(p):
                    if p.empty: return pd.DataFrame(columns=['25 FC3', '26 FC1', 'ACT']).reindex(pd.MultiIndex.from_tuples([], names=['Project', 'Con.', 'SOP']))
                    main = p.loc[p.index.isin(top)]
                    oth = p.loc[~p.index.isin(top)].sum().to_frame().T
                    oth.index = pd.MultiIndex.from_tuples([('Others', '', '')], names=['Project', 'Con.', 'SOP'])
                    return pd.concat([main, oth])
                p_m, p_y, p_fy, p_prev = group_others(p_m), group_others(p_y), group_others(p_fy), group_others(p_prev)
                idx = p_m.index
                
            combined_dict = {}
            combined_dict[(prev_phase_name, 'ACT')] = p_prev['ACT'] if 'ACT' in p_prev.columns else pd.Series(0, index=idx)
            for phase_name, data in [(phase_names[0], p_m), (phase_names[1], p_y), (phase_names[2], p_fy)]:
                for c in ['25 FC3', '26 FC1', 'ACT']:
                    combined_dict[(phase_name, c)] = data[c] if c in data.columns else pd.Series(0, index=idx)
                num = pd.Series(combined_dict[(phase_name, 'ACT')])
                den = pd.Series(combined_dict[(phase_name, '26 FC1')])
                combined_dict[(phase_name, 'ACHI %')] = num.div(den).replace([np.inf, -np.inf], 0).fillna(0)
            
            combined = pd.DataFrame(combined_dict, index=idx)
            if (phase_names[0], '26 FC1') in combined.columns and (phase_names[0], 'ACT') in combined.columns:
                fc1_curr = combined[(phase_names[0], '26 FC1')].abs()
                act_curr = combined[(phase_names[0], 'ACT')].abs()
                combined = combined[(fc1_curr >= 0.01) | (act_curr >= 0.01)]
                
            if combined.empty: continue
            
            if ('Others', '', '') in combined.index: 
                combined = pd.concat([combined.drop(index=('Others', '', '')).sort_values(by=(phase_names[0], 'ACT'), ascending=False), combined.loc[[('Others', '', '')]]])
            else: 
                if not combined.empty and (phase_names[0], 'ACT') in combined.columns:
                    combined = combined.sort_values(by=(phase_names[0], 'ACT'), ascending=False)
            
            subtotal = combined.sum(numeric_only=True) if not combined.empty else pd.Series(0, index=combined.columns)
            for p_name in phase_names:
                num = subtotal.get((p_name, 'ACT'), 0)
                den = subtotal.get((p_name, '26 FC1'), 0)
                subtotal[(p_name, 'ACHI %')] = num / den if den != 0 else 0
            
            combined.index = pd.MultiIndex.from_tuples([(brand, p, c, s) for p, c, s in combined.index], names=['Cust. GR', 'Project', 'Con.', 'SOP'])
            results.append(combined)
            results.append(pd.DataFrame([subtotal], index=pd.MultiIndex.from_tuples([(brand, f'{brand}_소계', '', '')], names=['Cust. GR', 'Project', 'Con.', 'SOP'])))
            
    if not results: return pd.DataFrame(), phase_names
    final_df = pd.concat(results)
    
    grand_total = final_df[final_df.index.get_level_values(1).str.contains('소계', na=False)].sum(numeric_only=True)
    for p_name in phase_names:
        num = grand_total.get((p_name, 'ACT'), 0)
        den = grand_total.get((p_name, '26 FC1'), 0)
        grand_total[(p_name, 'ACHI %')] = num / den if den != 0 else 0
        
    grand_label = f'{BIZ_CONFIG.get(biz_type, biz_type)} Rev. TTL (K.€)'
    grand_row = pd.DataFrame([grand_total], index=pd.MultiIndex.from_tuples([(f'GRAND_TOTAL_MERGE_START{grand_label}', 'GRAND_TOTAL_MERGE_DEL', 'GRAND_TOTAL_MERGE_DEL', 'GRAND_TOTAL_MERGE_DEL')], names=['Cust. GR', 'Project', 'Con.', 'SOP']))
    return pd.concat([final_df, grand_row]), phase_names

def build_trend_report(df, end_year, end_month):
    months, curr_y, curr_m = [], end_year, end_month
    for _ in range(12):
        months.append((curr_y, curr_m))
        curr_m -= 1
        if curr_m == 0: curr_m, curr_y = 12, curr_y - 1
    months.reverse() 
    
    df_act = df[df['Desc.'] == 'ACT']
    cps_list = sorted(df_act['CPS'].dropna().unique().tolist())
    
    pivot_data = {}
    for y, m in months:
        col_name = f"{MONTH_NAMES[m]}.{str(y)[-2:]}"
        temp_df = df_act[(df_act['Year'] == y) & (df_act['Month'] == m)]
        pivot_data[col_name] = temp_df.groupby('CPS')['Rev. (€)'].sum()
        
    trend_df = pd.DataFrame(pivot_data).reindex(cps_list).fillna(0)
    trend_df = trend_df.loc[(trend_df.sum(axis=1) != 0)]
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
selected_menu = st.sidebar.radio("원하시는 작업을 선택하세요.", ["매출 보고서", "판매가 조회", "AQL status 정리", "생산 실적 분석"])
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
        
        col1, col2 = st.columns(2)
        with col1:
            df_trend_data = build_trend_report(raw_df, selected_year, selected_month)
            if not df_trend_data.empty:
                plot_df = df_trend_data.drop('TTL (K.€)').reset_index().melt(id_vars='index', var_name='Month', value_name='Rev')
                plot_df.rename(columns={'index': 'CPS'}, inplace=True)
                plot_df['Rev'] = plot_df['Rev'] / 1000.0
                fig1 = px.bar(plot_df, x='Month', y='Rev', color='CPS', title='12 Months Revenue Trend (K.€)', color_discrete_map={'PE': '#002060', 'DC': '#8ea9db', 'CC': '#355E3B', 'CE':'#CD853F' })
                fig1.update_layout(height=300, plot_bgcolor='rgba(0,0,0,0)', yaxis=(dict(showgrid=True, gridcolor='#e6e6e6')), margin=dict(l=20, r=20, t=40, b=20), legend_title_text='')
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
                
                fc1_hover_texts, act_hover_texts = [], []
                for cps in chart2_pivot['CPS']:
                    cps_data = proj_pivot[proj_pivot['CPS'] == cps]
                    cps_fc1 = cps_data.sort_values(by='26 FC1', ascending=False).head(5)
                    if not cps_fc1.empty and cps_fc1['26 FC1'].sum() != 0:
                        lines = ["<br><b>[Top 5 Projects (FC1 Target)]</b>"]
                        for rank, (_, row) in enumerate(cps_fc1.iterrows(), 1): lines.append(f"{rank}. {row['Project']} : {row['26 FC1'] / 1000.0:,.0f} K.€")
                        fc1_hover_texts.append("<br>".join(lines))
                    else: fc1_hover_texts.append("<br><b>[No FC1 Data]</b>")
                        
                    cps_act = cps_data.sort_values(by='ACT', ascending=False).head(5)
                    if not cps_act.empty and cps_act['ACT'].sum() != 0:
                        lines = ["<br><b>[Top 5 Projects (ACT Actual)]</b>"]
                        for rank, (_, row) in enumerate(cps_act.iterrows(), 1):
                            act_val, fc1_val = row['ACT'] / 1000.0, row['26 FC1'] / 1000.0
                            diff = act_val - fc1_val
                            pct_str = f"{(diff / fc1_val) * 100:+.1f}%" if fc1_val != 0 else "+100.0%" if act_val > 0 else "N/A"
                            diff_html = f"<span style='color: darkred;'>{diff:,.0f} K.€ ({pct_str})</span>" if diff < 0 else f"<span style='color: darkblue;'>+{diff:,.0f} K.€ ({pct_str})</span>" if diff > 0 else "0 K.€ (0%)"
                            lines.append(f"{rank}. {row['Project']} : {act_val:,.0f} K.€ | {diff_html}")
                        act_hover_texts.append("<br>".join(lines))
                    else: act_hover_texts.append("<br><b>[No ACT Data]</b>")
                
                chart2_pivot['fc1_hover'], chart2_pivot['act_hover'] = fc1_hover_texts, act_hover_texts
                
                fig2 = go.Figure(data=[
                    go.Bar(name='FC1', x=chart2_pivot['CPS'], y=chart2_pivot['26 FC1'], marker_color='#c7c7c7', text=chart2_pivot['26 FC1'].apply(lambda x: f'{x:,.0f}'), textposition='inside', textfont=dict(color='white', size=12, weight='bold'), customdata=chart2_pivot['fc1_hover'], hovertemplate="<b>CPS: %{x}</b><br>FC1: %{y:,.0f} K.€%{customdata}<extra></extra>"),
                    go.Bar(name='ACT', x=chart2_pivot['CPS'], y=chart2_pivot['ACT'], marker_color='#1f77b4', text=chart2_pivot['ACT'].apply(lambda x: f'{x:,.0f}'), textposition='inside', textfont=dict(color='white', size=12, weight='bold'), customdata=chart2_pivot['act_hover'], hovertemplate="<b>CPS: %{x}</b><br>ACT: %{y:,.0f} K.€%{customdata}<extra></extra>")
                ])
                fig2.update_layout(height=300,barmode='group', title=f'[{MONTH_NAMES.get(selected_month)}] FC1 vs ACT by CPS (K.€)', plot_bgcolor='rgba(0,0,0,0)', yaxis=(dict(showgrid=True, gridcolor='#e6e6e6', visible=False)), margin=dict(l=20, r=20, t=50, b=20), legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
                st.plotly_chart(fig2, use_container_width=True)
                
        st.markdown("---")
        st.subheader("📌 Sales Revenue Trend")
        if not df_trend_data.empty:
            st.markdown(render_trend_html_table(df_trend_data, apply_color=False), unsafe_allow_html=True)
            reports_to_download["12M_Trend_Report"] = df_trend_data
            
        st.subheader("📌 CPS별 매출액 요약")
        df_cps, p_col, c_col = build_summary_report(raw_df, ['CPS'], selected_year, selected_month, 'TTL (K.€)')
        if not df_cps.empty: 
            st.markdown(render_html_view(df_cps, c_col, apply_color=False), unsafe_allow_html=True)
            reports_to_download["CPS_Summary"] = df_cps

        st.subheader("📌 DIRECT & COMMISSION 매출액 요약")
        df_biz_type, c_col = get_biz_type_detailed_report(raw_df, selected_year, selected_month)
        if not df_biz_type.empty: 
            st.markdown(render_html_view(df_biz_type, c_col, apply_color=True), unsafe_allow_html=True)
            reports_to_download["Biz_Type_Summary"] = df_biz_type

        st.subheader("📌 Sales Revenue: Power Electronics")
        df_pe_raw = raw_df[raw_df['Business Type'].str.contains("Power", case=False, na=False)].copy()
        if not df_pe_raw.empty:
            df_pe_raw['Cust. GR'] = df_pe_raw['Group 2'].replace({'HYU': 'HKMC', 'KIA': 'HKMC'})
            df_pe_summary, p_col, c_col = build_summary_report(df_pe_raw[df_pe_raw['Cust. GR'] == 'HKMC'], ['Cust. GR', 'KOx'], selected_year, selected_month, total_label='PE Biz Rev. TTL (K.€)', sort_by_current_act=True, add_ex_rate=True) 
            if not df_pe_summary.empty:
                st.markdown(render_html_view(df_pe_summary, c_col, apply_color=True), unsafe_allow_html=True)
                reports_to_download["PE_HKMC_Summary"] = df_pe_summary
                
        st.subheader("📌 PE Item 매출액 요약")
        df_item_raw = raw_df[raw_df['Item'].isin(['ICCU1', 'ICCU2', 'VCMS'])]
        df_item, p_col, c_col = build_summary_report(df_item_raw, ['Item'], selected_year, selected_month, 'TTL (K.€)', index_names=['CPS'], sort_by_current_act=True)
        if not df_item.empty: 
            st.markdown(render_html_view(df_item, c_col, apply_color=False), unsafe_allow_html=True)
            reports_to_download["Item_Summary"] = df_item
            
        st.subheader("📌 Core Biz Summary")
        df_core_grp, c_col = get_core_biz_summary_report(raw_df, selected_year, selected_month)
        if not df_core_grp.empty:
            st.markdown(render_html_view(df_core_grp, c_col, apply_color=True), unsafe_allow_html=True)
            reports_to_download["Core_Biz_Grp1_Summary"] = df_core_grp

        for filter_key, display_name in BIZ_CONFIG.items():
            st.subheader(f"📌 Sales Revenue: {display_name}")
            df_biz, phase_names = get_biz_report(raw_df, filter_key, selected_year, selected_month)
            if not df_biz.empty:
                st.markdown(render_biz_html_table(df_biz, phase_names[0], apply_color=True), unsafe_allow_html=True)
                reports_to_download[f"{display_name}_Detailed"] = df_biz

        if reports_to_download:
            st.write("---")
            st.download_button("📥 월간회의 자료용 엑셀 다운로드", data=to_excel_multiple(reports_to_download), file_name=f"Monthly_Closing_Report_{selected_year}_{selected_month:02d}.xlsx", use_container_width=True)
    else: st.info("👈 좌측 메뉴에서 '월간 회의용 엑셀 파일'을 업로드하시면 요약 리포트가 생성됩니다.")

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
                    if len(parts) > 2:
                        if parts[1].isdigit() and parts[2] == '': current_customer = parts[1]
                        elif len(parts) >= 16 and parts[1] in ['YPR0', 'ZADD', 'YSPR']:
                            try: 
                                amt = float(parts[10].replace(',', ''))
                                per = float(parts[12].replace(',', ''))
                                price = int(amt / per) if (amt / per).is_integer() else round(amt / per, 2) if per != 0 else ""
                            except: price = ""
                            
                            def format_date(d_str):
                                d_str = str(d_str).strip()
                                if not d_str: return d_str
                                pts = re.split(r'[\.\-\/]', d_str)
                                if len(pts) == 3:
                                    if len(pts[0]) == 4: return f"{pts[0]}-{pts[1].zfill(2)}-{pts[2].zfill(2)}"
                                    elif len(pts[2]) == 4: return f"{pts[2]}-{pts[1].zfill(2)}-{pts[0].zfill(2)}"
                                try: return pd.to_datetime(d_str, dayfirst=True).strftime('%Y-%m-%d')
                                except Exception: return d_str
                                
                            v_from = format_date(parts[14])
                            v_to = format_date(parts[15])
                            parsed_data.append({"Sales Org.": sales_org, "Distr. Channel": distr_channel, "Customer": current_customer, "CnTy": parts[1], "Condition Type": parts[2], "Material": parts[5], "Material Description": parts[6], "From": v_from, "To": v_to, "Price": price, "Curr.": parts[11]})
        
        if parsed_data:
            df_final = pd.DataFrame(parsed_data)
            st.markdown("---")
            st.subheader("🔍 특정 일자/조건 기준 단가 합산 시뮬레이터")
            st.info("입력하신 조건과 조회 기준일(Target Date)에 유효한(From~To 사이) 단가를 필터링하여 합산합니다.")
            
            with st.form("price_simulator_form"):
                c1, c2, c3, c4 = st.columns(4)
                with c1: sim_org = st.text_input("Sales Org. (입력 시 필터)")
                with c2: sim_distr = st.text_input("Distr. Channel (입력 시 필터)")
                with c3: sim_cust = st.text_input("Customer (입력 시 필터)")
                with c4: sim_date = st.date_input("조회 기준일 (Target Date)", value=datetime.date.today())
                sim_mats = st.text_area("조회할 Material 리스트 (엔터 또는 쉼표(,)로 구분하여 여러 개 입력)")
                submitted = st.form_submit_button("단가 합산 조회하기")
                
            if submitted:
                cond = pd.Series(True, index=df_final.index)
                if sim_org.strip(): cond &= df_final['Sales Org.'] == sim_org.strip()
                if sim_distr.strip(): cond &= df_final['Distr. Channel'] == sim_distr.strip()
                if sim_cust.strip(): cond &= df_final['Customer'] == sim_cust.strip()
                if sim_mats.strip():
                    mat_list = [m.strip() for m in re.split(r'[\n,]', sim_mats) if m.strip()]
                    cond &= df_final['Material'].astype(str).isin(mat_list)
                    
                target_date_str = sim_date.strftime("%Y-%m-%d")
                cond &= (df_final['From'] <= target_date_str) & (df_final['To'] >= target_date_str)
                df_sim = df_final[cond].copy()
                
                if not df_sim.empty:
                    df_sim['Price_Num'] = pd.to_numeric(df_sim['Price'], errors='coerce').fillna(0)
                    base_info = df_sim.groupby(['Sales Org.', 'Distr. Channel', 'Customer', 'Material'], as_index=False).agg({'Material Description': 'first', 'Curr.': 'first', 'CnTy': lambda x: ' + '.join(x.dropna().astype(str).unique())}).rename(columns={'CnTy': 'Memo'})
                    pivot_prices = df_sim.pivot_table(index=['Sales Org.', 'Distr. Channel', 'Customer', 'Material'], columns='CnTy', values='Price_Num', aggfunc='sum').fillna(0).reset_index()
                    df_grouped = pd.merge(base_info, pivot_prices, on=['Sales Org.', 'Distr. Channel', 'Customer', 'Material'])
                    cnty_cols = [c for c in pivot_prices.columns if c not in ['Sales Org.', 'Distr. Channel', 'Customer', 'Material']]
                    standard_sum_cols = [c for c in cnty_cols if c != 'YSPR']
                    df_grouped['Sales price'] = df_grouped[standard_sum_cols].sum(axis=1)
                    cols_order = ['Sales Org.', 'Distr. Channel', 'Customer', 'Material', 'Material Description'] + cnty_cols + ['Sales price', 'Memo', 'Curr.']
                    df_grouped = df_grouped[cols_order]
                    total_sales_price = df_grouped['Sales price'].sum()
                    
                    st.success(f"### 🎉 전체 합산 단가 (Total Sales Price): {total_sales_price:,.2f} (조회된 자재: {len(df_grouped)}건)")
                    st.dataframe(df_grouped, use_container_width=True)
                    
                    out_sim = io.BytesIO()
                    with pd.ExcelWriter(out_sim, engine='xlsxwriter') as writer:
                        df_grouped.to_excel(writer, sheet_name="Simulation_Result", index=False)
                        ws = writer.sheets["Simulation_Result"]
                        ws.set_column('A:E', 15)
                        for i, _ in enumerate(cnty_cols): ws.set_column(5+i, 5+i, 12)
                        sales_price_idx = 5 + len(cnty_cols)
                        ws.set_column(sales_price_idx, sales_price_idx, 15)
                        ws.set_column(sales_price_idx+1, sales_price_idx+1, 20)
                        ws.set_column(sales_price_idx+2, sales_price_idx+2, 10)
                    st.download_button("📥 시뮬레이션 결과 엑셀 다운로드", data=out_sim.getvalue(), file_name="Simulation_Result.xlsx", use_container_width=True)
                else: st.warning("조건에 일치하며 해당 일자에 유효한 단가 데이터가 없습니다.")

            st.markdown("---")
            st.subheader("📋 정제된 단가 데이터 전체 목록")
            st.dataframe(df_final, use_container_width=True)
            
            df_export = df_final.copy()
            df_export.columns = ["Sales Org.", "Distr. Channel", "Customer", "CnTy", "Condition Type", "Material", "Material Description", "From", "To", "Price", "Curr."]
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_export.to_excel(writer, sheet_name="Sheet1", index=False, startrow=3, header=False)
                workbook = writer.book
                ws = writer.sheets["Sheet1"]
                header_format = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#F0F0F0', 'align': 'center'})
                headers = ["Sales Org.", "Distr. Channel", "Customer", "CnTy", "Condition Type", "Material", "Material Description", "From", "To", "Price", "Curr."]
                for col_num, value in enumerate(headers): ws.write(2, col_num, value, header_format)
                ws.set_column('A:E', 12); ws.set_column('F:F', 12); ws.set_column('G:G', 40)
                ws.set_column('H:I', 12); ws.set_column('J:K', 10)
            st.download_button("📥 통합 결과 엑셀 전체 다운로드", data=output.getvalue(), file_name="결과.xlsx", use_container_width=True)
        else: st.warning("분석할 수 있는 데이터가 없습니다. txt 파일 형식을 확인해주세요.")

elif selected_menu == "AQL status 정리":
    st.title("📑 AQL Status 정리")
    st.info("엑셀 파일을 업로드하면 PRJT 기준으로 Status에 따라 PF Desc.를 그룹화하고, 동일한 KOKOR SOP 날짜를 정리해 드립니다.")
    uploaded_aql_file = st.sidebar.file_uploader("AQL 엑셀 데이터를 업로드하세요.", type=['xlsx', 'xls'], key="aql_uploader")
    if uploaded_aql_file:
        try:
            df_temp = pd.read_excel(uploaded_aql_file, header=None)
            header_row_idx = -1
            required_cols = ['PRJT', 'Status', 'PF Desc.', 'KOKOR SOP']
            
            for idx, row in df_temp.iterrows():
                row_vals = [str(x).strip() for x in row.values if pd.notna(x)]
                if all(c in row_vals for c in required_cols):
                    header_row_idx = idx
                    break
            if header_row_idx == -1: st.error(f"엑셀 파일 내에서 필수 컬럼({', '.join(required_cols)})을 찾을 수 없습니다.")
            else:
                df_aql = pd.read_excel(uploaded_aql_file, header=header_row_idx)
                st.success("📂 데이터를 성공적으로 불러왔습니다. 정리를 완료했습니다!")
                df_aql['PRJT'] = df_aql['PRJT'].fillna('Unknown')
                
                res_rows = []
                valid_statuses = ['awarded to kostal', 'acq. start / rfq rec.', 'in planning']
                
                for prjt, group in df_aql.groupby('PRJT'):
                    if prjt == 'Unknown': continue
                    a_list = []
                    t_list = []
                    sop_set = set()
                    
                    for _, row in group.iterrows():
                        status = str(row['Status']).strip().lower()
                        pf_desc = str(row['PF Desc.']).strip()
                        if status in valid_statuses:
                            formatted_sop = parse_sop_date(row.get('KOKOR SOP'))
                            if formatted_sop: sop_set.add(formatted_sop)
                            if pf_desc != 'nan' and pf_desc:
                                if status == 'awarded to kostal':
                                    if pf_desc not in a_list: a_list.append(pf_desc)
                                else:
                                    if pf_desc not in t_list: t_list.append(pf_desc)
                            
                    item_str_parts = []
                    if a_list: item_str_parts.append("[A] " + ", ".join(a_list))
                    if t_list: item_str_parts.append("[T] " + ", ".join(t_list))
                    item_result = " ".join(item_str_parts)
                    
                    if len(sop_set) == 1: sop_result = list(sop_set)[0]
                    elif len(sop_set) > 1: sop_result = "SOP to be checked"
                    else: sop_result = ""
                    
                    if item_result or sop_result: res_rows.append({'PRJT': prjt, 'ITEM': item_result, 'SOP': sop_result})
                        
                if res_rows:
                    df_result = pd.DataFrame(res_rows)
                    st.markdown("---")
                    st.subheader("📋 정리된 AQL Status 목록")
                    st.dataframe(df_result, use_container_width=True)
                    
                    output_aql = io.BytesIO()
                    with pd.ExcelWriter(output_aql, engine='xlsxwriter') as writer:
                        df_result.to_excel(writer, sheet_name="AQL_Status", index=False)
                        worksheet = writer.sheets["AQL_Status"]
                        worksheet.set_column('A:A', 20) 
                        worksheet.set_column('B:B', 70) 
                        worksheet.set_column('C:C', 20) 
                    st.download_button(label="📥 AQL 정리 결과 엑셀 다운로드", data=output_aql.getvalue(), file_name=f"AQL_Status_Summary_{datetime.date.today().strftime('%Y%m%d')}.xlsx", use_container_width=True)
                else: st.warning("조건에 해당하는 유효한 PRJT 및 데이터가 없습니다.")
        except Exception as e: st.error(f"파일을 처리하는 중 오류가 발생했습니다: {e}")
    else: st.info("👈 좌측 메뉴에서 'AQL 엑셀 데이터'를 업로드하시면 요약 리포트가 생성됩니다.")

elif selected_menu == "생산 실적 분석":
    st.title("📈 생산 실적 분석 (YoY Performance)")
    st.info("엑셀 파일을 업로드하면 날짜형식으로 된 컬럼들을 자동으로 인식하여 지정한 월의 전년 동월 대비 증감을 분석합니다.")
    
    uploaded_prod_file = st.sidebar.file_uploader("생산 실적 데이터를 업로드하세요.", type=['xlsx', 'xls'], key="prod_uploader")
    
    if uploaded_prod_file:
        try:
            df_prod = pd.read_excel(uploaded_prod_file)
            st.success("📂 데이터를 성공적으로 불러왔습니다.")
            
            hk_col = 'H/K'
            cn_col = 'CN'
            car_col = 'Car code master'
            
            missing = [c for c in [hk_col, cn_col] if c not in df_prod.columns]
            if missing:
                st.error(f"엑셀 파일에 다음 필수 컬럼이 없습니다: {', '.join(missing)}\n해당 데이터 포맷이 맞는지 확인해주세요.")
            else:
                date_cols = []
                date_mapping = {}
                for col in df_prod.columns:
                    try:
                        dt = pd.to_datetime(str(col))
                        date_cols.append(col)
                        date_mapping[col] = dt
                    except: pass
                
                if not date_cols:
                    st.error("엑셀 파일의 컬럼명에서 날짜(연/월) 정보를 인식할 수 없습니다. (예: 2026-01, 2026.01 등의 컬럼 필요)")
                else:
                    id_vars = [c for c in [hk_col, cn_col, car_col] if c in df_prod.columns]
                    df_melted = df_prod.melt(id_vars=id_vars, value_vars=date_cols, var_name='RawDate', value_name='Qty')
                    df_melted['Date'] = df_melted['RawDate'].map(date_mapping)
                    df_melted['Year'] = df_melted['Date'].dt.year
                    df_melted['Month'] = df_melted['Date'].dt.month
                    df_melted['Qty'] = pd.to_numeric(df_melted['Qty'], errors='coerce').fillna(0)
                    
                    years = sorted(df_melted['Year'].dropna().unique(), reverse=True)
                    months = sorted(df_melted['Month'].dropna().unique())
                    
                    st.sidebar.markdown("### 📅 조회 기준 선택")
                    if not years:
                        st.sidebar.error("날짜 데이터가 없습니다.")
                    else:
                        target_year = st.sidebar.selectbox("조회 연도 (Target Year)", years)
                        target_month = st.sidebar.selectbox("조회 월 (Target Month)", months)
                        
                        if st.button("분석 실행하기"):
                            df_curr = df_melted[(df_melted['Year'] == target_year) & (df_melted['Month'] == target_month)]
                            df_prev = df_melted[(df_melted['Year'] == target_year - 1) & (df_melted['Month'] == target_month)]
                            
                            # 날짜 형식 동적 헤더 생성 (예: 2025.04)
                            prev_col_name = f"{target_year - 1}.{target_month:02d}"
                            curr_col_name = f"{target_year}.{target_month:02d}"
                            
                            def build_yoy(df_c, df_p, group_cols):
                                curr_agg = df_c.groupby(group_cols)['Qty'].sum().rename(curr_col_name)
                                prev_agg = df_p.groupby(group_cols)['Qty'].sum().rename(prev_col_name)
                                merged = pd.concat([prev_agg, curr_agg], axis=1).fillna(0)
                                merged['증감(Diff)'] = merged[curr_col_name] - merged[prev_col_name]
                                merged['증감율(YoY %)'] = np.where(merged[prev_col_name] == 0, 
                                                                np.where(merged[curr_col_name] > 0, 1.0, 0.0), 
                                                                merged['증감(Diff)'] / merged[prev_col_name])
                                return merged
                                
                            table1 = build_yoy(df_curr, df_prev, [hk_col, cn_col])
                            
                            total_curr = df_curr['Qty'].sum()
                            total_prev = df_prev['Qty'].sum()
                            total_diff = total_curr - total_prev
                            total_pct = (total_diff / total_prev) if total_prev != 0 else 0
                            
                            # 전체 실적 증감에 따른 동적 정렬 (플러스면 내림차순, 마이너스면 오름차순)
                            sort_ascending = True if total_diff < 0 else False
                            table1 = table1.sort_values('증감(Diff)', ascending=sort_ascending)
                            
                            st.markdown("---")
                            st.subheader(f"💡 전체 실적 요약: {total_curr:,.0f} (전년 동월 대비 **{total_diff:+,.0f}**, **{total_pct:+.1%}**)")
                            
                            if total_diff < 0:
                                st.error("📉 **분석 의견:** 전체 생산 실적이 전년 동월 대비 감소했습니다. 전년 동월 대비 가장 많이 감소한 항목(H/K, CN)은 다음과 같습니다.")
                                dec_df = table1[table1['증감(Diff)'] < 0]
                                for i, (idx, row) in enumerate(dec_df.iterrows()):
                                    hk, cn = idx
                                    st.write(f"{i+1}. **{hk} - {cn}** : {row['증감(Diff)']:,.0f} 감소 (전년비 {row['증감율(YoY %)']:.1%})")
                            elif total_diff > 0:
                                st.success("📈 **분석 의견:** 전체 생산 실적이 전년 동월 대비 증가했습니다. 전년 동월 대비 가장 많이 증가한 항목(H/K, CN)은 다음과 같습니다.")
                                inc_df = table1[table1['증감(Diff)'] > 0]
                                for i, (idx, row) in enumerate(inc_df.iterrows()):
                                    hk, cn = idx
                                    st.write(f"{i+1}. **{hk} - {cn}** : +{row['증감(Diff)']:,.0f} 증가 (전년비 {row['증감율(YoY %)']:.1%})")
                            else:
                                st.info("전년 동월 대비 전체 실적에 변동이 없습니다.")
                                
                            st.markdown("---")
                            st.subheader("1. H/K, CN 기준 전년 동월대비 실적")
                            format_dict = {prev_col_name: '{:,.0f}', curr_col_name: '{:,.0f}', '증감(Diff)': '{:,.0f}', '증감율(YoY %)': '{:.1%}'}
                            st.dataframe(table1.style.format(format_dict), use_container_width=True)
                            
                            if car_col in df_prod.columns:
                                table2 = build_yoy(df_curr, df_prev, [hk_col, cn_col, car_col])
                                table2 = table2.reset_index().sort_values(by=[hk_col, cn_col, '증감(Diff)'], ascending=[True, True, sort_ascending]).set_index([hk_col, cn_col, car_col])
                                
                                st.subheader("2. H/K, CN, Car code master 기준 전년 동월대비 실적")
                                st.dataframe(table2.style.format(format_dict), use_container_width=True)
                            else:
                                st.warning(f"'{car_col}' 컬럼이 없어서 상세 테이블은 생략되었습니다.")
                                
        except Exception as e:
            st.error(f"오류가 발생했습니다: {e}")
    else:
        st.info("👈 좌측 메뉴에서 '생산 실적 데이터'를 업로드하시면 실적 비교 리포트가 자동 생성됩니다.")
