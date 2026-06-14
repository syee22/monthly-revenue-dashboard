import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import uuid

# ==========================================
# 1. 페이지 설정 및 전역 변수 설정
# ==========================================
st.set_page_config(page_title="Sales Revenue - Monthly Report", layout="wide")

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
.report-table { border-collapse: collapse !important; font-family: 'Malgun Gothic', sans-serif; font-size: 12px; width: 100%; background-color: white; margin: 0 !important; border: 2px solid #002060 !important; }
.report-table tr { border-bottom: none !important; }
.report-table td, .report-table th { border-bottom: none !important; border-top: none !important; }
.report-table th, .report-table td { white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.report-table thead th { background-color: #002060 !important; color: white !important; border: 1px solid #8ea9db !important; text-align: center !important; padding: 4px 3px !important; font-weight: 600 !important; font-size: 11.5px !important; position: sticky; top: 0; z-index: 10; }
.report-table td { border: 1px solid #d9d9d9; text-align: center; padding: 4px; vertical-align: middle; }
.report-table .row_heading { color: #002060 !important; text-align: left !important; padding-left: 10px !important; border: 1px solid #d9d9d9 !important; vertical-align: middle !important; font-weight: bold !important; }
.report-table tr.total-row th, .report-table tr.total-row td { background-color: #ffffe0 !important; color: #002060 !important; font-weight: bold !important; border-top: 2px solid #8ea9db !important; border-bottom: 2px solid #8ea9db !important; }
.report-table tr.total-row-hyu th, .report-table tr.total-row-hyu td { background-color: #e6f2ff !important; color: #002060 !important; font-weight: bold !important; border-top: 2px solid #8ea9db !important; border-bottom: 2px solid #8ea9db !important; }
.report-table tr.total-row-kia th, .report-table tr.total-row-kia td { background-color: #ffe6e6 !important; color: #002060 !important; font-weight: bold !important; border-top: 2px solid #8ea9db !important; border-bottom: 2px solid #8ea9db !important; }
.report-table tr.total-row-direct th, .report-table tr.total-row-direct td { background-color: #99caff !important; color: #002060 !important; font-weight: bold !important; border-top: 2px solid #8ea9db !important; border-bottom: 2px solid #8ea9db !important; }
.report-table tr.total-row-comm th, .report-table tr.total-row-comm td { background-color: #d0d0d0 !important; color: #002060 !important; font-weight: bold !important; border-top: 2px solid #8ea9db !important; border-bottom: 2px solid #8ea9db !important; }
.report-table.biz-table .row_heading { text-align: center !important; padding-left: 4px !important; padding-right: 4px !important; }
</style>""", unsafe_allow_html=True)

# ==========================================
# 2. 동적 테두리 및 열 너비 제어
# ==========================================
def get_dynamic_column_css(table_id, df):
    """Con.과 SOP 열의 너비를 50px로 고정하는 CSS 생성"""
    css = ""
    # 컬럼 인덱스 찾기 (멀티인덱스 컬럼 처리)
    for i, col in enumerate(df.columns):
        col_name = col[1] if isinstance(col, tuple) else str(col)
        if col_name in ['Con.', 'SOP']:
            # index column 개수만큼 th 위치 더해주기 (+1은 nth-child가 1부터 시작해서)
            target_idx = df.index.nlevels + i + 1
            css += f"#{table_id} th:nth-child({target_idx}), #{table_id} td:nth-child({target_idx}) {{ min-width: 50px !important; max-width: 50px !important; width: 50px !important; }}"
    return f"<style>{css}</style>"

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
    if 0.95 <= val <= 1.0:
        return f'<span style="color: #000000; font-weight: bold; font-style: italic;">{pct_str} <span style="display: inline-block; width: 10px; height: 4px; background-color: #cc7a00; vertical-align: middle; margin-left: 5px;"></span></span>' 
    elif val > 1.0:
        return f'<span style="color: #2E86C1; font-weight: bold; font-style: italic;">{pct_str} ▲</span>'
    elif val > 0:
        return f'<span style="color: #c00000; font-weight: bold; font-style: italic;">{pct_str} ▼</span>'
    else:
        return f'<span style="font-style: italic;">{pct_str}</span>'

def apply_common_styles(styler, apply_hkmc_color=False, is_export=False):
    imp = "" if is_export else " !important"
    def style_row(row):
        row_str = str(row.name)
        base_style = ''
        if any(k in row_str for k in ['HYU_소계', 'DIRECT_Subtotal_숨김']): base_style = f'background-color: #e6f2ff{imp}; color: #002060; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;'
        elif 'KIA_소계' in row_str: base_style = f'background-color: #ffe6e6{imp}; color: #002060; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;'
        elif 'COMM_Subtotal_숨김' in row_str: base_style = f'background-color: #d0d0d0{imp}; color: #002060; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;'
        elif any(k in row_str for k in ['Unknown_Subtotal', 'GRAND_TOTAL', 'TTL', 'Total', '소계']): base_style = f'background-color: #ffffe0{imp}; color: #002060; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;'
        return [base_style] * len(row)
    styler.apply(style_row, axis=1)
    return styler

def optimize_html_headers(html_str, df):
    try:
        thead_start = html_str.find('<thead>')
        thead_end = html_str.find('</thead>')
        if thead_start == -1 or thead_end == -1: return html_str
        thead_html = html_str[thead_start:thead_end+8]
        tr_matches = list(re.finditer(r'<tr[^>]*>(.*?)</tr>', thead_html, re.IGNORECASE | re.DOTALL))
        if len(tr_matches) < 2: return html_str
        row0_inner, row1_inner = tr_matches[0].group(1), tr_matches[1].group(1)
        th_pattern = r'<th[^>]*>.*?</th>'
        ths0 = re.findall(th_pattern, row0_inner, re.IGNORECASE | re.DOTALL)
        ths1 = re.findall(th_pattern, row1_inner, re.IGNORECASE | re.DOTALL)
        num_indices = df.index.nlevels
        index_names = list(df.index.names)
        for i in range(num_indices):
            if i < len(ths0) and i < len(ths1):
                name = str(index_names[i]) if index_names[i] is not None else ""
                ths0[i] = f'<th rowspan="2" style="vertical-align: middle !important; text-align: center !important; background-color: #002060 !important; color: white !important; border: 1px solid #8ea9db !important; min-width: 80px;">{name}</th>'
                ths1[i] = ''
        return html_str[:thead_start] + f"<thead>\n<tr>{''.join(ths0)}</tr>\n<tr>{''.join(ths1)}</tr>\n</thead>" + html_str[thead_end+8:]
    except: return html_str

def post_process_html_styles(html_str):
    if '<tbody>' not in html_str: return html_str
    def process_row(match):
        row = match.group(0)
        if 'HYU_소계' in row: row = re.sub(r'^<tr', r'<tr class="total-row-hyu"', row.replace('HYU_소계', ''))
        elif 'KIA_소계' in row: row = re.sub(r'^<tr', r'<tr class="total-row-kia"', row.replace('KIA_소계', ''))
        elif 'DIRECT_Subtotal_숨김' in row: row = re.sub(r'^<tr', r'<tr class="total-row-direct"', row.replace('DIRECT_Subtotal_숨김', ''))
        elif 'COMM_Subtotal_숨김' in row: row = re.sub(r'^<tr', r'<tr class="total-row-comm"', row.replace('COMM_Subtotal_숨김', ''))
        elif 'GRAND_TOTAL_MERGE_START' in row:
            row = re.sub(r'^<tr', r'<tr class="total-row"', row)
            label = re.search(r'GRAND_TOTAL_MERGE_START(.*?)</th', row).group(1).strip()
            row = re.sub(r'<th[^>]*>GRAND_TOTAL_MERGE_START.*?</th>.*?</th>', f'<th colspan="4" style="text-align: left !important; padding-left: 15px !important; background-color: #ffffe0 !important; color: #002060 !important; font-weight: bold !important; border-top: 2px solid #8ea9db !important; border-bottom: 2px solid #8ea9db !important;">{label}</th>', row, flags=re.DOTALL)
        elif any(k in row for k in ['TTL', 'Total', '소계']): row = re.sub(r'^<tr', r'<tr class="total-row"', row)
        return row
    parts = html_str.split('<tbody>', 1)
    return parts[0] + '<tbody>' + re.sub(r'<tr[^>]*>.*?</tr>', process_row, parts[1], flags=re.DOTALL)

def to_excel_multiple(df_dict):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for name, df in df_dict.items():
            styler = df.style.format(lambda x: format_k_val(x) if isinstance(x, (int, float)) else x)
            styler = apply_common_styles(styler, apply_hkmc_color=True, is_export=True)
            styler.to_excel(writer, sheet_name=name[:31])
    return output.getvalue()

# ==========================================
# 3. 데이터 로드 및 렌더링
# ==========================================
uploaded_file = st.sidebar.file_uploader("엑셀 업로드", type=['xlsx', 'xls'])
if uploaded_file:
    # ... (데이터 로드 로직은 이전과 동일하므로 생략)
    raw_df = load_and_preprocess(uploaded_file)
    # ... (생략)
    # 렌더링 함수에만 아래 내용 추가
    def render_biz_html_table(df, phase_curr, apply_color=False):
        table_id = f"table_{uuid.uuid4().hex[:8]}"
        df_display = df.replace(0, '')
        styler = df_display.style.format({col: format_percentage_html if 'ACHI' in col[1] else format_k_val for col in df.columns}).set_table_attributes(f'class="report-table biz-table" id="{table_id}"')
        html_str = optimize_html_headers(styler.to_html(), df)
        html_str = post_process_html_styles(html_str)
        # 너비 제어 CSS + 테두리 CSS 추가
        return f'{get_dynamic_column_css(table_id, df)}{get_dynamic_highlight_css(table_id, df, phase_curr)}<div id="{table_id}" class="table-container">{html_str}</div>'
    
    # ... (나머지 로직은 동일)
