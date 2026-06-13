import streamlit as st
import pandas as pd
import numpy as np
import io
import re

# ==========================================
# 1. 페이지 설정 및 전역 변수 설정
# ==========================================
st.set_page_config(page_title="Sales Revenue - Monthly Report", layout="wide")

# [중요] 이 부분이 누락되면 NameError가 발생합니다!
MONTH_NAMES = {1:'Jan', 2:'Feb', 3:'Mar', 4:'Apr', 5:'May', 6:'Jun', 7:'Jul', 8:'Aug', 9:'Sep', 10:'Oct', 11:'Nov', 12:'Dec'}
BIZ_CONFIG = {"Power": "PE Biz", "Core": "Core Biz"}

# ==========================================
# CSS 주입 (total-row 클래스 추가)
# ==========================================
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
        color: #002060 !important;
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
    
    /* 기본 합계/소계 행 (연노랑) */
    .report-table tr.total-row th, 
    .report-table tr.total-row td {
        background-color: #ffffe0 !important;
        color: #002060 !important;
        font-weight: bold !important;
        border-top: 2px solid #8ea9db !important;
        border-bottom: 2px solid #8ea9db !important;
    }
    
    /* HYU 소계 행 (하늘색) */
    .report-table tr.total-row-hyu th, 
    .report-table tr.total-row-hyu td {
        background-color: #e6f2ff !important;
        color: #002060 !important;
        font-weight: bold !important;
        border-top: 2px solid #8ea9db !important;
        border-bottom: 2px solid #8ea9db !important;
    }

    /* KIA 소계 행 (분홍색) */
    .report-table tr.total-row-kia th, 
    .report-table tr.total-row-kia td {
        background-color: #ffe6e6 !important;
        color: #002060 !important;
        font-weight: bold !important;
        border-top: 2px solid #8ea9db !important;
        border-bottom: 2px solid #8ea9db !important;
    }
    
    /* DIRECT 소계 행 (연한 청색) */
    .report-table tr.total-row-direct th, 
    .report-table tr.total-row-direct td {
        background-color: #e6f2ff !important;
        color: #002060 !important;
        font-weight: bold !important;
        border-top: 2px solid #8ea9db !important;
        border-bottom: 2px solid #8ea9db !important;
    }

    /* COMM 소계 행 (연한 회색) */
    .report-table tr.total-row-comm th, 
    .report-table tr.total-row-comm td {
        background-color: #f2f2f2 !important;
        color: #002060 !important;
        font-weight: bold !important;
        border-top: 2px solid #8ea9db !important;
        border-bottom: 2px solid #8ea9db !important;
    }
    
    /* PE Biz 및 Core Biz 상세 테이블 인덱스 가운데 정렬 */
    .report-table.biz-table .row_heading {
        text-align: center !important;
        padding-left: 4px !important;
        padding-right: 4px !important;
    }

    /* Trend 테이블 마지막 열(당월) 강조 (진한 빨간색 테두리) */
    .report-table.trend-table th:last-child {
        border-top: 2px solid #c00000 !important;
        border-left: 2px solid #c00000 !important;
        border-right: 2px solid #c00000 !important;
    }
    .report-table.trend-table td:last-child {
        border-left: 2px solid #c00000 !important;
        border-right: 2px solid #c00000 !important;
    }
    .report-table.trend-table tbody tr:last-child td:last-child {
        border-bottom: 2px solid #c00000 !important;
    }
    </style>
""", unsafe_allow_html=True)

st.title(" Sales Revenue - Monthly Report")

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
    if 0.95 <= val <= 1.0:
        return f'<span style="color: #000000; font-weight: bold; font-style: italic;">{pct_str} <span style="display: inline-block; width: 10px; height: 4px; background-color: #cc7a00; vertical-align: middle; margin-left: 5px;"></span></span>' 
    elif val > 1.0:
        return f'<span style="color: #2E86C1; font-weight: bold; font-style: italic;">{pct_str} ▲</span>'
    elif val > 0:
        return f'<span style="color: #c00000; font-weight: bold; font-style: italic;">{pct_str} ▼</span>'
    else:
        return f'<span style="font-style: italic;">{pct_str}</span>'

def color_index_cells(v):
    val_str = str(v)
    if val_str == 'HYU': return 'background-color: #e6f2ff;'  # 하늘색
    if val_str == 'KIA': return 'background-color: #ffe6e6;'  # 분홍색
    if val_str == 'DIRECT': return 'background-color: #e6f2ff;' # 연한 청색
    if val_str == 'COMM': return 'background-color: #f2f2f2;' # 연한 회색
    return ''

def apply_common_styles(styler, apply_hkmc_color=False, is_export=False):
    # 엑셀 다운로드 시에는 !important 태그를 제외합니다.
    imp = "" if is_export else " !important"
    
    # 1. 엑셀 출력 및 셀 단위 스타일링 (데이터 영역)
    def style_row(row):
        row_str = str(row.name)
        if 'HYU_소계' in row_str:
            return [f'background-color: #e6f2ff{imp}; color: #002060; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;'] * len(row)
        elif 'KIA_소계' in row_str:
            return [f'background-color: #ffe6e6{imp}; color: #002060; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;'] * len(row)
        elif 'DIRECT_Subtotal_숨김' in row_str:
            return [f'background-color: #e6f2ff{imp}; color: #002060; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;'] * len(row)
        elif 'COMM_Subtotal_숨김' in row_str:
            return [f'background-color: #f2f2f2{imp}; color: #002060; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;'] * len(row)
        elif 'Unknown_Subtotal_숨김' in row_str:
            return [f'background-color: #ffffe0{imp}; color: #002060; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;'] * len(row)
        elif 'GRAND_TOTAL_MERGE' in row_str:
            return [f'background-color: #ffffe0{imp}; color: #002060; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;'] * len(row)
        elif any(keyword in row_str for keyword in ['TTL', 'Total', '소계']):
            return [f'background-color: #ffffe0{imp}; color: #002060; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;'] * len(row)
        return [''] * len(row)
    
    styler.apply(style_row, axis=1)
    
    # 2. 인덱스 영역 스타일링
    if hasattr(styler, 'apply_index'):
        def style_row_index(idx):
            res = []
            for label in idx:
                label_str = str(label)
                if 'HYU_소계' in label_str: 
                    res.append(f'background-color: #e6f2ff{imp}; color: #e6f2ff; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;')
                elif 'KIA_소계' in label_str: 
                    res.append(f'background-color: #ffe6e6{imp}; color: #ffe6e6; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;')
                elif 'DIRECT_Subtotal_숨김' in label_str:
                    res.append(f'background-color: #e6f2ff{imp}; color: #e6f2ff; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;')
                elif 'COMM_Subtotal_숨김' in label_str:
                    res.append(f'background-color: #f2f2f2{imp}; color: #f2f2f2; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;')
                elif 'Unknown_Subtotal_숨김' in label_str:
                    res.append(f'background-color: #ffffe0{imp}; color: #ffffe0; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;')
                elif 'GRAND_TOTAL_MERGE' in label_str:
                    res.append(f'background-color: #ffffe0{imp}; color: #002060; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;')
                elif any(k in label_str for k in ['TTL', 'Total', '소계']):
                    res.append(f'background-color: #ffffe0{imp}; color: #002060; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;')
                else: 
                    res.append('')
            return res
        styler.apply_index(style_row_index, axis=0)
    else:
        def highlight_total_index(val):
            val_str = str(val)
            if 'HYU_소계' in val_str:
                return f'background-color: #e6f2ff{imp}; color: #e6f2ff; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;'
            elif 'KIA_소계' in val_str:
                return f'background-color: #ffe6e6{imp}; color: #ffe6e6; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;'
            elif 'DIRECT_Subtotal_숨김' in val_str:
                return f'background-color: #e6f2ff{imp}; color: #e6f2ff; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;'
            elif 'COMM_Subtotal_숨김' in val_str:
                return f'background-color: #f2f2f2{imp}; color: #f2f2f2; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;'
            elif 'Unknown_Subtotal_숨김' in val_str:
                return f'background-color: #ffffe0{imp}; color: #ffffe0; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;'
            elif 'GRAND_TOTAL_MERGE' in val_str:
                return f'background-color: #ffffe0{imp}; color: #002060; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;'
            elif any(keyword in val_str for keyword in ['TTL', 'Total', '소계']):
                return f'background-color: #ffffe0{imp}; color: #002060; font-weight: bold; border-top: 2px solid #8ea9db; border-bottom: 2px solid #8ea9db;'
            return ''
        for i in range(styler.index.nlevels):
            styler.map_index(highlight_total_index, axis=0, level=i)
        
    # 3. 4번, 5번, 6번 및 Biz_Type 테이블에 기본 인덱스 셀 색상 적용
    if apply_hkmc_color:
        if hasattr(styler, 'map_index'):
            styler.map_index(color_index_cells, axis=0, level=0)
        elif hasattr(styler, 'applymap_index'):
            styler.applymap_index(color_index_cells, axis=0, level=0)
            
    return styler

def optimize_html_headers(html_str, df):
    """Pandas HTML 테이블의 빈 헤더를 배열 방식으로 안전하게 병합하여 중앙 정렬 구현"""
    try:
        thead_start = html_str.find('<thead>')
        thead_end = html_str.find('</thead>')
        if thead_start == -1 or thead_end == -1: return html_str
        
        thead_html = html_str[thead_start:thead_end+8]
        
        tr_matches = list(re.finditer(r'<tr[^>]*>(.*?)</tr>', thead_html, re.IGNORECASE | re.DOTALL))
        if len(tr_matches) < 2: return html_str
        
        row0_inner = tr_matches[0].group(1)
        row1_inner = tr_matches[1].group(1)
        
        th_pattern = r'<th[^>]*>.*?</th>'
        ths0 = re.findall(th_pattern, row0_inner, re.IGNORECASE | re.DOTALL)
        ths1 = re.findall(th_pattern, row1_inner, re.IGNORECASE | re.DOTALL)
        
        if hasattr(df, 'index') and hasattr(df.index, 'names'):
            index_names = [str(n) for n in df.index.names if n is not None and str(n).strip()]
            num_indices = len(index_names)
            
            for i in range(num_indices):
                if i < len(ths0) and i < len(ths1):
                    name = index_names[i]
                    ths0[i] = f'<th rowspan="2" style="vertical-align: middle !important; text-align: center !important; background-color: #002060 !important; color: white !important; border: 1px solid #8ea9db !important; min-width: 80px;">{name}</th>'
                    ths1[i] = ''
                    
        row0_new = "".join(ths0)
        row1_new = "".join(ths1)
        
        new_thead = f"<thead>\n<tr>{row0_new}</tr>\n<tr>{row1_new}</tr>\n</thead>"
        return html_str[:thead_start] + new_thead + html_str[thead_end+8:]
    except Exception:
        return html_str

def post_process_html_styles(html_str):
    """HTML 렌더링 시 브랜드를 추적하여 각 소계 행에 맞는 고유 CSS 클래스, 텍스트 삭제 및 병합 셀 처리"""
    if '<tbody>' not in html_str: return html_str
    
    def process_row(match):
        row = match.group(0)
        
        if 'HYU_소계' in row:
            row = row.replace('HYU_소계', '') # 텍스트 완전히 삭제 (빈칸 처리)
            row = re.sub(r'^<tr', r'<tr class="total-row-hyu"', row)
        elif 'KIA_소계' in row:
            row = row.replace('KIA_소계', '') 
            row = re.sub(r'^<tr', r'<tr class="total-row-kia"', row)
            
        # DIRECT 및 COMM 소계 처리
        elif 'DIRECT_Subtotal_숨김' in row:
            row = row.replace('DIRECT_Subtotal_숨김', '') 
            row = re.sub(r'^<tr', r'<tr class="total-row-direct"', row) # 연한 청색
        elif 'COMM_Subtotal_숨김' in row:
            row = row.replace('COMM_Subtotal_숨김', '') 
            row = re.sub(r'^<tr', r'<tr class="total-row-comm"', row) # 연한 회색
        elif 'Unknown_Subtotal_숨김' in row:
            row = row.replace('Unknown_Subtotal_숨김', '') 
            row = re.sub(r'^<tr', r'<tr class="total-row"', row) # 기존 노란색
            
        # Grand Total 병합 처리 로직
        elif 'GRAND_TOTAL_MERGE_START' in row:
            row = re.sub(r'^<tr', r'<tr class="total-row"', row)
            label_match = re.search(r'GRAND_TOTAL_MERGE_START(.*?)</th', row)
            if label_match:
                actual_label = label_match.group(1).strip()
                row = re.sub(
                    r'<th[^>]*>GRAND_TOTAL_MERGE_START.*?</th>\s*<th[^>]*>GRAND_TOTAL_MERGE_DEL</th>\s*<th[^>]*>GRAND_TOTAL_MERGE_DEL</th>\s*<th[^>]*>GRAND_TOTAL_MERGE_DEL</th>',
                    f'<th colspan="4" style="text-align: left !important; padding-left: 15px !important; background-color: #ffffe0 !important; color: #002060 !important; font-weight: bold !important; border-top: 2px solid #8ea9db !important; border-bottom: 2px solid #8ea9db !important;">{actual_label}</th>',
                    row,
                    flags=re.DOTALL
                )
                
        elif any(k in row for k in ['TTL', 'Total', '소계']):
            row = re.sub(r'^<tr', r'<tr class="total-row"', row)
            
        return row
        
    parts = html_str.split('<tbody>', 1)
    tbody_content = parts[1]
    tbody_content = re.sub(r'<tr[^>]*>.*?</tr>', process_row, tbody_content, flags=re.DOTALL)
    return parts[0] + '<tbody>' + tbody_content

def to_excel_multiple(df_dict):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet_name, original_df in df_dict.items():
            df = original_df.copy()
            
            # 엑셀 다운로드를 위해 병합용 내부 마커 텍스트를 깔끔하게 제거
            if isinstance(df.index, pd.MultiIndex):
                new_tuples = []
                for t in df.index:
                    new_t = list(t)
                    if isinstance(new_t[0], str) and 'GRAND_TOTAL_MERGE_START' in new_t[0]:
                        new_t[0] = new_t[0].replace('GRAND_TOTAL_MERGE_START', '')
                        for i in range(1, len(new_t)):
                            if new_t[i] == 'GRAND_TOTAL_MERGE_DEL':
                                new_t[i] = ''
                    new_tuples.append(tuple(new_t))
                df.index = pd.MultiIndex.from_tuples(new_tuples, names=df.index.names)
            
            styler = df.style.format(lambda x: format_k_val(x) if isinstance(x, (int, float)) else x)
            
            # 적용할 시트 목록 확장 (Biz_Type_Summary 포함)
            apply_color = sheet_name in ["PE_HKMC_Summary", "PE_Biz_Detailed", "Core_Biz", "Biz_Type_Summary"]
            styler = apply_common_styles(styler, apply_hkmc_color=apply_color, is_export=True)
            
            styler.to_excel(writer, sheet_name=sheet_name[:31])
            
            worksheet = writer.sheets[sheet_name[:31]]
            for i, col in enumerate(df.columns):
                worksheet.set_column(i+1, i+1, 15)
    return output.getvalue()

# ==========================================
# 3. 데이터 로드 및 전처리
# ==========================================
uploaded_file = st.sidebar.file_uploader("엑셀 데이터를 업로드하세요.", type=['xlsx', 'xls'])

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
            
            # 각 비즈니스 타입에 맞는 고유 숨김 식별자 사용
            results.append(pd.DataFrame([subtotal], index=pd.MultiIndex.from_tuples([(biz, f'{biz}_Subtotal_숨김')], names=['BIZ Type', 'KOx'])))
            
        final_df = pd.concat(results)
        
        # 소계용 숨김 식별자들을 제외하고 총계 계산
        grand_total = final_df[~final_df.index.get_level_values(1).str.contains('Subtotal_숨김', na=False)].sum(numeric_only=True)
        for p_name in phase_names:
            num = grand_total.get((p_name, 'ACT'), 0)
            den = grand_total.get((p_name, '26 FC1'), 0)
            grand_total[(p_name, 'ACHI %')] = num / den if den != 0 else 0
            
        grand_row = pd.DataFrame([grand_total], index=pd.MultiIndex.from_tuples([('TTL (K.€)', ' ')], names=['BIZ Type', 'KOx']))
        
        return pd.concat([final_df, grand_row])

    def get_biz_report(df, biz_type, year, month):
        if month == 1: prev_year, prev_month = year - 1, 12
        else: prev_year, prev_month = year, month - 1
        
        df_biz = df[(df['Business Type'].str.contains(biz_type, case=False, na=False)) & (df['Year'] == year)].copy()
        m_str, pm_str = MONTH_NAMES.get(month, f'{month}'), MONTH_NAMES.get(prev_month, f'{prev_month}')
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
            
            if brand != 'GM': 
                results.append(pd.DataFrame([subtotal], index=pd.MultiIndex.from_tuples([(brand, f'{brand}_소계', '', '')], names=['Cust. GR', 'Project', 'Con.', 'SOP'])))
            
        final_df = pd.concat(results)
        grand_total = final_df[~final_df.index.get_level_values(1).str.contains('소계', na=False)].sum(numeric_only=True)
        for p_name in phase_names:
            num = grand_total.get((p_name, 'ACT'), 0)
            den = grand_total.get((p_name, '26 FC1'), 0)
            grand_total[(p_name, 'ACHI %')] = num / den if den != 0 else 0
            
        grand_label = f'{BIZ_CONFIG.get(biz_type, biz_type)} Rev. TTL (K.€)'
        grand_row = pd.DataFrame(
            [grand_total], 
            index=pd.MultiIndex.from_tuples(
                [(f'GRAND_TOTAL_MERGE_START{grand_label}', 'GRAND_TOTAL_MERGE_DEL', 'GRAND_TOTAL_MERGE_DEL', 'GRAND_TOTAL_MERGE_DEL')], 
                names=['Cust. GR', 'Project', 'Con.', 'SOP']
            )
        )
        
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
        
        df_act = df[df['Desc.'] == 'ACT']
        
        pivot_data = {}
        for y, m in months:
            col_name = f"{MONTH_NAMES[m]}.{str(y)[-2:]}"
            temp_df = df_act[(df_act['Year'] == y) & (df_act['Month'] == m)]
            pivot_data[col_name] = temp_df.groupby('Group 2')['Rev. (€)'].sum()
            
        trend_df = pd.DataFrame(pivot_data)
        row_order = ['HYU', 'KIA', 'GM']
        trend_df = trend_df.reindex(row_order).fillna(0)
        trend_df.loc['TTL (K.€)'] = trend_df.sum(numeric_only=True)
        
        trend_df.index.name = None 
        trend_df.columns.name = None 
        
        trend_df.columns = [col.strip() for col in trend_df.columns.values]
        return trend_df

    # ==========================================
    # 5. 스타일링 및 렌더링
    # ==========================================
    def render_html_view(df, phase_curr, apply_color=False, title=None):
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
        
        html_str = styler.to_html()
        html_str = optimize_html_headers(html_str, df)
        html_str = post_process_html_styles(html_str)
        
        if title: 
            pass 
        return f'<div class="table-container">{html_str}</div>'

    def render_biz_html_table(df, apply_color=False, title=None):
        df_display = df.replace(0, '')
        format_dict = {col: format_percentage_html if 'ACHI' in col[1] else format_k_val for col in df.columns}
        styler = df_display.style.format(format_dict, na_rep='').set_table_attributes('class="report-table biz-table"')
        
        numeric_cols = get_numeric_cols(df)
        styler.set_properties(subset=numeric_cols, **{'text-align': 'right'})
        styler = apply_common_styles(styler, apply_hkmc_color=apply_color)
        
        html_str = styler.to_html()
        html_str = optimize_html_headers(html_str, df)
        html_str = post_process_html_styles(html_str)
        
        if title: 
            pass
        return f'<div class="table-container">{html_str}</div>'

    def render_trend_html_table(df, apply_color=False):
        df_display = df.replace(0, '')
        format_dict = {col: format_k_val for col in df.columns}
        # Trend 테이블에 전용 CSS 클래스 부여
        styler = df_display.style.format(format_dict, na_rep='').set_table_attributes('class="report-table trend-table"')
        
        styler.set_properties(**{'text-align': 'right'})
        styler = apply_common_styles(styler, apply_hkmc_color=apply_color)
        
        html_str = styler.to_html()
        html_str = post_process_html_styles(html_str)
        return f'<div class="table-container">{html_str}</div>'

    # ==========================================
    # 6. 화면 출력
    # ==========================================
    reports_to_download = {}
    st.subheader("📌 Sales Revenue Trend")
    df_trend = build_trend_report(raw_df, selected_year, selected_month)
    if not df_trend.empty:
        st.markdown(render_trend_html_table(df_trend, apply_color=False), unsafe_allow_html=True)
        reports_to_download["12M_Trend_Report"] = df_trend
        
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
    df_biz_type = get_biz_type_detailed_report(raw_df, selected_year, selected_month)
    if not df_biz_type.empty: 
        # DIRECT 및 COMM의 색상 처리를 위해 apply_color=True로 활성화합니다.
        st.markdown(render_html_view(df_biz_type, "", apply_color=True), unsafe_allow_html=True)
        reports_to_download["Biz_Type_Summary"] = df_biz_type

    st.subheader("📌 Sales Revenue: Power Electronics")
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
            st.markdown(render_html_view(df_pe_summary, c_col, apply_color=True), unsafe_allow_html=True)
            reports_to_download["PE_HKMC_Summary"] = df_pe_summary

    for filter_key, display_name in BIZ_CONFIG.items():
        st.subheader(f"📌 Sales Revenue: {display_name}")
        df_biz, _ = get_biz_report(raw_df, filter_key, selected_year, selected_month)
        if not df_biz.empty:
            st.markdown(render_biz_html_table(df_biz, apply_color=True), unsafe_allow_html=True)
            reports_to_download[f"{display_name}_Detailed"] = df_biz

    if reports_to_download:
        st.write("---")
        st.download_button(
            "📥 월간회의 자료용 엑셀 다운로드", 
            data=to_excel_multiple(reports_to_download), 
            file_name=f"Monthly_Closing_Report_{selected_year}_{selected_month:02d}.xlsx", 
            use_container_width=True
        )
else:
    st.info("👈 좌측 사이드바에서 엑셀 파일을 업로드하시면 요약 리포트가 자동 생성됩니다.")
