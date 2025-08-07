# 第一部分：所有導入
import streamlit as st
import pandas as pd
import numpy as np  # 只保留一次
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import altair as alt
from PIL import Image
import base64  
import re, os, io, warnings, traceback, subprocess, sys, time
import pathlib

def resource_path(rel):
    """開發階段與 PyInstaller 打包後皆能取資源檔"""
    if getattr(sys, "frozen", False):      # .exe 執行
        return pathlib.Path(sys._MEIPASS) / rel
    return pathlib.Path(__file__).parent / rel

# 設置頁面配置
st.set_page_config(
    page_title="IQC 效率管理系統",
    page_icon="IQC2.png",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 其他Streamlit命令和初始化
warnings.filterwarnings('ignore')

# 檢測並安裝所需套件
required_packages = ['xlsxwriter', 'openpyxl']
for package in required_packages:
    try:
        __import__(package)
    except ImportError:
        print(f"安裝所需套件: {package}")

# 添加現代無襯線字體設定
st.markdown("""
<style>
    /* IQC 效率管理系統設計更新 - 淺灰色設計方案 */

    /* ===== 1. 顏色變量定義 - 淺灰色配色方案 ===== */
    :root {
        /* 主要顏色 - 淺灰藍色系 */
        --primary-color: #78909C;      /* 主要顏色 - 淺灰藍 */
        --primary-dark: #607D8B;       /* 主要顏色稍深版 - 用於懸停效果 */
        --primary-light: #CFD8DC;      /* 主要顏色更淺版 - 用於背景、邊框等 */
        
        /* 中性色 - 更高對比度 */
        --neutral-dark: #37474F;       /* 深灰色 - 用於主要文字 */
        --neutral-medium: #90A4AE;     /* 中灰色 - 用於次要元素 */
        --neutral-light: #F5F7F9;      /* 極淺灰色 - 用於背景 */
        
        /* 強調色 - 更鮮明的色彩確保可識別性 */
        --accent-blue: #5C8AE6;        /* 淺藍色 - 取代過於暗沉的藍色 */
        --accent-green: #66BB6A;       /* 淺綠色 - 用於正面指標 */
        --accent-red: #EF5350;         /* 淺紅色 - 用於警告/錯誤 */
        --accent-amber: #FFB74D;       /* 琥珀色 - 用於提示、注意事項 */
        
        /* 字體與間距 */
        --font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    }

    /* ===== 2. 應用淺灰色配色方案 ===== */

    /* 全局元素 */
    html, body, [class*="css"], [class*="st-"] {
        font-family: var(--font-family) !important;
        color: var(--neutral-dark);
    }

    /* 主頁背景色調整為更淺的色調 */
    .main {
        background-color: var(--neutral-light);
        color: var(--neutral-dark);
    }

    /* 所有標題加深顏色以提高對比度 */
    h1, h2, h3, h4, h5, h6 {
        color: #2B3A42; /* 更深的顏色確保標題清晰可見 */
        font-weight: 600;
    }

    /* 段落與文字加深顏色 */
    p, div, span, label {
        color: #333333; /* 深灰色文字確保良好對比度 */
    }

    /* 標籤頁調整為淺色方案 */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        margin-bottom: 16px;
    }

    .stTabs [data-baseweb="tab"] {
        padding: 10px 16px;
        font-weight: 500;
        background-color: #ECEFF1; /* 更淺的背景色 */
        border-radius: 6px 6px 0 0;
        color: #455A64; /* 加深文字顏色 */
    }

    .stTabs [aria-selected="true"] {
        background-color: var(--primary-color) !important;
        color: white !important;
    }

    /* 卡片樣式統一，確保明亮背景與清晰文字 */
    div[data-testid="metric-container"] {
        background-color: white;
        border-radius: 10px;
        box-shadow: 0 2px 5px rgba(0,0,0,0.05);
        padding: 15px;
        margin-bottom: 20px;
        border: 1px solid #E0E0E0; /* 淺邊框提高卡片邊界清晰度 */
    }

    div[data-testid="metric-container"] label {
        font-weight: 500;
        color: #455A64; /* 加深標籤顏色 */
    }

    div[data-testid="metric-container"] [data-testid="stMetricValue"] {
        font-size: 24px;
        font-weight: bold;
        color: #37474F; /* 加深數值顏色確保清晰可見 */
    }

    /* 訊息框統一樣式 */
    div.stAlert {
        border-radius: 8px;
        padding: 16px;
        border-left-color: var(--primary-color);
        background-color: white; /* 確保背景為白色提高可讀性 */
    }

    div.stAlert[data-baseweb="notification"] {
        background-color: white;
    }

    /* ===== 3. 淺灰色按鈕樣式 ===== */

    /* 基本按鈕樣式 - 淺灰藍色 */
    .stButton>button {
        background-color: #90A4AE; /* 更淺的灰藍色 */
        color: white;
        border: none;
        border-radius: 6px;
        padding: 0.5rem 1rem;
        font-weight: 500;
        transition: all 0.2s;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }

    .stButton>button:hover {
        background-color: #78909C; /* 懸停時稍微深一點 */
        box-shadow: 0 2px 5px rgba(0,0,0,0.15);
        transform: translateY(-1px);
        color: white !important;
    }

    /* 處理資料按鈕樣式 */
    .stButton button[data-testid="baseButton-primary"] {
        background-color: #90A4AE !important; /* 與基本按鈕保持一致 */
        border: none !important;
        color: white !important;
        font-weight: 600 !important;
        border-radius: 6px !important;
    }

    /* 處理資料按鈕懸停樣式 */
    .stButton button[data-testid="baseButton-primary"]:hover {
        background-color: #78909C !important; /* 懸停時稍微深一點 */
        box-shadow: 0 2px 5px rgba(0,0,0,0.15);
    }

    /* 匯出按鈕樣式 */
    .export-button button {
        background-color: #78909C !important; /* 稍深一點區分不同功能 */
        border: none !important;
        color: white !important;
        font-weight: 500 !important;
        margin-top: 10px;
        border-radius: 6px !important;
    }

    /* 匯出按鈕懸停樣式 */
    .export-button button:hover {
        background-color: #607D8B !important; /* 懸停時更深 */
        box-shadow: 0 2px 5px rgba(0,0,0,0.15);
    }

    /* ===== 4. 側邊欄樣式優化 ===== */
    .css-6qob1r, .css-10oheav, .css-r698ls {
        background-color: white;
        font-family: var(--font-family) !important;
    }

    .sidebar .sidebar-content {
        background-color: white;
    }

    /* 側邊欄標題 */
    .sidebar .sidebar-content h3 {
        color: #455A64; /* 更深的顏色確保標題清晰可見 */
        font-weight: 600;
    }
    
    /* 側邊欄分隔線 */
    .sidebar .sidebar-content hr {
        border-color: #E0E0E0; /* 淺色分隔線 */
    }

    /* 檔案上傳區域標題 */
    .upload-title {
        color: #455A64; /* 加深顏色確保清晰可見 */
        font-weight: 600;
    }

    /* 進度條樣式優化 */
    .stProgress > div > div > div {
        background-color: #90A4AE !important; /* 使用稍深的灰藍色 */
    }

    /* ===== 5. 其他UI元素調整 ===== */

    /* 下拉選單樣式 */
    div[data-baseweb="select"] {
        border-radius: 6px;
        border-color: #E0E0E0; /* 淺邊框色 */
    }

    /* 下拉選單文字顏色加深 */
    div[data-baseweb="select"] span {
        color: #455A64 !important; /* 確保文字清晰可見 */
    }

    /* 日期選擇器樣式 */
    div[data-testid="stDateInput"] input {
        border-radius: 6px;
        padding: 10px;
        border-color: #E0E0E0; /* 淺邊框色 */
        color: #455A64; /* 加深文字顏色 */
    }

    /* 擴展區塊樣式 */
    .streamlit-expanderHeader {
        font-weight: 500;
        color: #455A64; /* 加深顏色確保文字清晰可見 */
        background-color: white;
        border-radius: 6px;
    }

    /* 日期篩選標題樣式 */
    .date-filter-title {
        color: #455A64 !important; /* 加深顏色確保標題清晰可見 */
        font-weight: 600 !important;
    }

    /* 調整資料表格樣式 */
    .dataframe {
        border-collapse: collapse;
        width: 100%;
    }

    .dataframe th {
        background-color: #ECEFF1; /* 更淺的背景色 */
        color: #455A64; /* 加深文字顏色 */
        padding: 10px 12px;
        font-weight: 500;
        text-align: left;
        border-bottom: 1px solid #E0E0E0;
    }

    .dataframe td {
        padding: 8px 12px;
        border-bottom: 1px solid #F5F7F9;
        color: #333333; /* 確保表格內文字清晰可見 */
    }

    .dataframe tr:hover {
        background-color: #F5F7F9;
    }

    /* ===== 6. 圖表色彩調整為更淺色調 ===== */

    /* Plotly圖表樣式 */
    .js-plotly-plot .plotly .modebar {
        right: 10px;
        top: 10px;
    }

    /* 折線圖中的線條顏色更明亮 */
    .js-plotly-plot .plotly .traces path.js-line {
        stroke: #5C8AE6 !important; /* 更淺明的藍色 */
    }

    /* 確保圖表背景色為白色 */
    .js-plotly-plot .plotly .plot-container {
        background-color: white !important;
    }

    /* 圖表標題顏色加深 */
    .js-plotly-plot .plotly .gtitle {
        fill: #333333 !important; /* 加深顏色確保標題清晰可見 */
    }
</style>
""", unsafe_allow_html=True)

# 初始化session_state中的關鍵變量 - 確保在使用前進行初始化
if 'show_cover' not in st.session_state:
    st.session_state['show_cover'] = True  # 初始顯示封面頁
if 'debug_info' not in st.session_state:
    st.session_state.debug_info = {'logs': []}
if 'log_level' not in st.session_state:
    st.session_state.log_level = "INFO"
if 'performance_mode' not in st.session_state:
    st.session_state.performance_mode = False
if 'iqc_report_data' not in st.session_state:
    st.session_state.iqc_report_data = None
if 'pcb_spec_data' not in st.session_state:
    st.session_state.pcb_spec_data = None
if 'pcb_standard_time_data' not in st.session_state:
    st.session_state.pcb_standard_time_data = None
if 'additional_tasks_data' not in st.session_state:
    st.session_state.additional_tasks_data = None
if 'processed_data' not in st.session_state:
    st.session_state.processed_data = None
if 'files_uploaded' not in st.session_state:
    st.session_state.files_uploaded = False
if 'processing_error' not in st.session_state:
    st.session_state.processing_error = None
if 'selected_categories' not in st.session_state:
    st.session_state.selected_categories = []


# 設定是否啟用調試模式
DEBUG_MODE = True

# 優化的日誌函數
def debug_log(message, data=None, level="DEBUG"):
    """優化版的日誌功能，添加日誌級別控制與性能模式"""
    
    # 如果處於高性能模式，只記錄WARNING及更高級別的日誌
    if st.session_state.get('performance_mode', False) and level not in ["WARNING", "ERROR"]:
        return
    
    # 定義日誌級別的優先順序
    log_levels = {
        "DEBUG": 1,
        "INFO": 2,
        "WARNING": 3,
        "ERROR": 4
    }
    
    # 獲取當前設定的日誌級別
    current_level = st.session_state.get('log_level', "DEBUG")
    
    # 只有當日誌級別優先順序大於等於當前設定時才記錄
    if log_levels.get(level, 1) >= log_levels.get(current_level, 1):
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_message = f"[{timestamp}][{level}] {message}"
        
        # 在高效能模式下，減少資料詳情
        if data is not None and not st.session_state.get('performance_mode', False):
            if isinstance(data, pd.DataFrame):
                log_message += f" - DataFrame shape: {data.shape}"
                if not data.empty and level == "DEBUG":
                    log_message += f", Columns: {list(data.columns)[:5]}..."
            else:
                try:
                    if level == "DEBUG":
                        log_message += f" - Data: {str(data)[:100]}..."
                    else:
                        log_message += f" - Data: {str(data)[:200]}"
                except:
                    log_message += f" - Data: (無法顯示)"
        
        if 'logs' not in st.session_state.debug_info:
            st.session_state.debug_info['logs'] = []
        
        # 限制最大日誌數量，避免內存過度使用
        max_logs = 500 if st.session_state.get('performance_mode', False) else 1000
        if len(st.session_state.debug_info['logs']) >= max_logs:
            st.session_state.debug_info['logs'] = st.session_state.debug_info['logs'][-max_logs+1:] + [log_message]
        else:
            st.session_state.debug_info['logs'].append(log_message)
        
        # 在高效能模式下，只輸出ERROR級別到控制台
        if level == "ERROR" or (not st.session_state.get('performance_mode', False) and level in ["WARNING", "ERROR"]):
            print(log_message)

FIELD_MAPPING = {
    # IQC Report 欄位映射
    'IQC_REPORT': {
        '檢驗員': ['檢驗員', 'Inspector', '檢驗人員', '操作員'],
        '檢驗日期': ['檢驗日期', 'Inspection Date', '日期', 'Date'],
        '料號': ['料號', 'Part No.', 'Part No', 'PartNo', 'Part Number'],
        '類別': ['類別', 'Category', 'Type'],
        '抽樣狀態': ['抽樣狀態', 'Sampling Status', 'Status'],
        '抽樣數量': ['抽樣數量', 'Sampling Qty', 'Sample Qty', 'Quantity'],
        'MRB': ['MRB', 'MRB狀態', 'MRB Status'],
        '檢驗標準工時': ['檢驗標準工時', 'Standard Time', '標準工時'],
        '檢驗耗時（調整後）': ['檢驗耗時（調整後）', '檢驗耗時(調整後)', '檢驗耗時', 'Actual Time', '實際耗時'],
        '檢驗開始時間': ['檢驗開始時間', 'AO', 'Start Time', '開始時間']
    },
    # PCB建檔明細欄位映射 - 壓合孔數從L欄改為N欄
    'PCB_SPECS': {
        '料號': ['料號', 'Part No.', 'Part No', 'PartNo', 'Part Number'],
        '壓合孔數': ['壓合孔數', 'Hole Count', '孔數'],  # 用於名稱映射
        '版長': ['版長', 'Length', '長度'],
        '版寬': ['版寬', 'Width', '寬度']
    },
    # PCB標準工時對應表欄位映射
    'PCB_STANDARD_TIME': {
        '面積範圍': ['面積範圍', 'Area Range', '面積'],
        '壓合總孔數': ['壓合總孔數', 'Total Hole Count', '孔數'],
        'PCB標準工時': ['PCB標準工時', 'Standard Time', '標準工時']
    },
    # IQC額外任務紀錄清單欄位映射
    'ADDITIONAL_TASKS': {
        '姓名': ['姓名', 'Name', '下拉式選單', '人員', 'B'],
        '日期': ['日期', 'Date', '手key', 'A'],
        '工作事項分類': ['工作事項分類', 'Task Type', '下拉式選單_1', '任務類型', 'C'],
        '用時(分鐘)': ['用時(分鐘)', '用時（分鐘）', 'Time(min)', '時間', '手key_3', 'H']
    }
}

# 在文件頂部添加這個常量（與其他常量一起）
EXCLUDED_INSPECTORS = ["謝芷馨", "蘇育珍", "許碧琦"]

# 添加檢驗員過濾函數
def filter_excluded_inspectors(df, inspector_column='處理後檢驗員'):
    """
    過濾掉不需要顯示在圖表中的檢驗員數據
    
    參數:
    df - DataFrame
    inspector_column - 檢驗員欄位名稱
    
    返回:
    過濾後的DataFrame
    """
    if df is None or df.empty:
        return df
        
    if inspector_column not in df.columns:
        return df
        
    # 過濾掉指定的檢驗員
    return df[~df[inspector_column].isin(EXCLUDED_INSPECTORS)].copy()

def update_field_mapping():
    global FIELD_MAPPING
    return FIELD_MAPPING

# 首先定义物料类别映射字典 - 在主函数开始前添加
MATERIAL_CATEGORY_MAPPING = {
    "ME-NC/NS/NH/NP/NT/NW/NR/ND/FC": ["NC", "NS", "NH", "NP", "NT", "NW", "NR", "ND", "FC"],
    "ME-WA/WE/WF/WU": ["WA", "WE", "WF", "WU"],
    "ME-XZ/XA/XV": ["XZ", "XA", "XV"],
    "EE-GB/GI": ["GB", "GI"],
    "EE-IC/IB/ID/IH/IL/IZ": ["IC", "IB", "ID", "IH", "IL", "IZ"],
    "EE-JB/JC/JD/JF/JL/JM/JR/JX": ["JB", "JC", "JD", "JF", "JL", "JM", "JR", "JX"],
    "EE-MP/MS/PC": ["MP", "MS", "PC"],
    "KM-BM/C-/9-/A-/GP": ["BM", "C-", "9-", "A-", "GP"],
    "PCB-QB": ["QB"]
}

@st.cache_data(ttl=3600, max_entries=50)
def get_field_value(row, field_name, mapping, default_value=None):
    """
    增強的安全獲取欄位值函數 - 根據映射表嘗試各種可能的欄位名稱
    
    參數:
    row: 資料行 (通常是dict格式)
    field_name: 欲獲取的欄位名稱
    mapping: 欄位映射字典
    default_value: 默認返回值
    
    返回:
    欄位值或默認值
    """
    # 檢查 mapping 中是否有這個欄位的映射
    if field_name in mapping:
        possible_field_names = mapping[field_name]
        
        # 如果映射是字符串，轉換為列表
        if isinstance(possible_field_names, str):
            possible_field_names = [possible_field_names]
        
        # 嘗試所有可能的欄位名稱
        for possible_name in possible_field_names:
            if possible_name in row:
                return row[possible_name]
        
        # 如果上面沒有找到，嘗試不區分大小寫的匹配
        for key in row.keys():
            if isinstance(key, str):
                for possible_name in possible_field_names:
                    if possible_name.lower() == key.lower():
                        return row[key]
    
    # 嘗試直接使用 field_name 作為鍵
    if field_name in row:
        return row[field_name]
    
    # 尋找可能的替代欄位 (模糊匹配)
    possible_keys = []
    for key in row.keys():
        if isinstance(key, str) and isinstance(field_name, str):
            if field_name.lower() in key.lower() or key.lower() in field_name.lower():
                possible_keys.append(key)
    
    if possible_keys:
        return row[possible_keys[0]]
    
    return default_value

@st.cache_data(ttl=3600, max_entries=100)
def parse_excel_date(date_val):
    """
    解析各種Excel日期格式，支持多種日期格式
    
    參數:
    date_val: 任意格式的日期值(數字、字符串、datetime對象等)
    
    返回:
    datetime對象或None(如果無法解析)
    """
    if pd.isna(date_val):
        return None
    
    try:
        # 如果是數字（Excel日期），轉換為datetime
        if isinstance(date_val, (int, float)):
            try:
                # Excel日期從1900年1月1日開始，但有個bug，會多算一天(1900年不是閏年)
                return datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(date_val) - 2)
            except:
                pass
        
        # 如果是字符串，嘗試多種格式
        if isinstance(date_val, str):
            # 嘗試各種日期格式
            formats = [
                '%Y-%m-%d', '%Y/%m/%d', '%m/%d/%Y', '%d/%m/%Y', 
                '%Y年%m月%d日', '%m-%d-%Y', '%d-%m-%Y',
                '%Y-%m-%d %H:%M:%S', '%Y/%m/%d %H:%M:%S',
                '%m/%d', '%Y-%m', '%m月%d日'
            ]
            
            for fmt in formats:
                try:
                    return datetime.strptime(date_val, fmt)
                except:
                    continue
            
            # 嘗試從日期字符串中提取數字
            import re
            date_match = re.search(r'(\d{1,4})[-/年](\d{1,2})[-/月](\d{1,2})', date_val)
            if date_match:
                try:
                    year, month, day = map(int, date_match.groups())
                    if year < 100:  # 兩位數年份
                        year += 2000
                    return datetime(year, month, day)
                except:
                    pass
        
        # 如果已經是datetime，直接返回
        if isinstance(date_val, datetime):
            return date_val
        
        # 最後嘗試pandas的通用解析
        return pd.to_datetime(date_val)
    
    except Exception as e:
        debug_log(f"日期解析錯誤: {e}", date_val)
        return None

@st.cache_data(ttl=3600, show_spinner=False)
def check_is_mrb(row, mrb_cache={}):
    """
    MRB檢測函數，支援單行和DataFrame處理
    """
    # 檢查輸入是否為DataFrame
    if isinstance(row, pd.DataFrame):
        # 創建結果列
        is_mrb = pd.Series(False, index=row.index)
        mrb_messages = pd.Series("無MRB標記", index=row.index)
        
        # 可能的MRB欄位
        mrb_columns = ['M', '異常問題匯總', 'Abnormal Summary']
        
        # 檢查所有可能的MRB欄位
        for col in mrb_columns:
            if col in row.columns:
                # 使用向量化操作而非逐行判斷
                col_mask = row[col].notna() & row[col].astype(str).str.strip().ne('')
                is_mrb = is_mrb | col_mask
                mrb_messages = mrb_messages.mask(col_mask, f"異常問題欄位({col})有內容")
        
        # 檢查Excel的M欄位位置 (第13個欄位)
        if len(row.columns) >= 13:
            m_col = row.columns[12]  # 第13個欄位 (0-based indexing)
            if m_col not in mrb_columns and m_col in row.columns:
                col_mask = row[m_col].notna() & row[m_col].astype(str).str.strip().ne('')
                is_mrb = is_mrb | col_mask
                mrb_messages = mrb_messages.mask(col_mask, f"M欄位({m_col})有內容")
        
        return pd.DataFrame({'是否為MRB': is_mrb, 'MRB訊息': mrb_messages})
    
    # 處理單行數據
    else:
        # 可能的MRB欄位
        mrb_columns = ['M', '異常問題匯總', 'Abnormal Summary']
        
        # 檢查所有可能的MRB欄位
        for col in mrb_columns:
            if col in row and pd.notna(row[col]) and str(row[col]).strip() != '':
                return (True, f"異常問題欄位({col})有內容: {row[col]}")
        
        # 嘗試直接檢查M欄位 (如果存在)
        if 'M' in row and pd.notna(row['M']) and str(row['M']).strip() != '':
            return (True, f"M欄位有內容: {row['M']}")
        
        return (False, "無MRB標記")

@st.cache_data(ttl=3600, max_entries=10, show_spinner=False)
def process_multiple_iqc_reports_optimized(files):
    try:
        debug_log(f"開始處理{len(files)}個IQC Report檔案")
        all_data_frames = []
        
        # 特殊檢驗員名稱對應字典
        special_inspectors = {
            'Cindy': '謝芷馨',
            'Joanne': '許碧琦',
            'Susu': '蘇育珍',
            'Wen': '許雅雯',
            'Flora': '毛凡甫',
            'ya-wen': '張雅雯'
        }
        
        for file_idx, file in enumerate(files):
            debug_log(f"處理第{file_idx+1}個IQC Report檔案: {file.name}")
            
            # 讀取Excel檔案
            xls = pd.ExcelFile(file)
            sheets = xls.sheet_names
            
            # 找到最可能的工作表
            target_sheet = next((sheet for sheet in sheets 
                                if any(keyword in sheet.lower() for keyword in 
                                      ['report', 'data', '資料', '報告', 'iqc'])), 
                              sheets[0])
            
            debug_log(f"使用工作表: {target_sheet}")
            df = pd.read_excel(file, sheet_name=target_sheet)
            debug_log(f"原始資料讀取完成，資料列數: {len(df)}")
            
            # 獲取欄位映射
            field_mapping = FIELD_MAPPING['IQC_REPORT']
            
            # 找出所有可能的M欄位（用於MRB判斷）
            m_column = df.columns[12] if len(df.columns) >= 13 else None
            possible_mrb_columns = [col for col in df.columns if 
                                   isinstance(col, str) and 
                                   ('異常問題' in col or 'Abnormal' in col or col.upper() == 'M')]
            if m_column and m_column not in possible_mrb_columns:
                possible_mrb_columns.append(m_column)
                
            debug_log(f"可能的MRB欄位: {possible_mrb_columns}")
            
            # 批量處理檢驗員名稱
            inspector_name_col = find_column_by_mapping(df, '檢驗員', field_mapping)
            if inspector_name_col:
                inspector_series = df[inspector_name_col].astype(str)
                df['處理後檢驗員'] = inspector_series
                
                # 提取括號內容
                bracket_pattern = r'\((.*?)\)'
                df['處理後檢驗員'] = df['處理後檢驗員'].str.extract(bracket_pattern, expand=False)
                
                # 檢查WYLZ標記
                df['包含WYLZ'] = df['處理後檢驗員'].str.contains('WYLZ', na=False)
                
                # 處理特殊案例
                df['處理後檢驗員'] = df['處理後檢驗員'].map(
                    lambda x: special_inspectors.get(x, x) if pd.notna(x) else x)
            else:
                df['處理後檢驗員'] = 'Unknown'
                df['包含WYLZ'] = False
            
            # 批量處理檢驗開始時間
            start_time_col = find_column_by_mapping(df, '檢驗開始時間', field_mapping)
            if start_time_col:
                df['檢驗開始時間'] = pd.to_datetime(df[start_time_col], errors='coerce')
            
            # ===== MRB狀態處理修正 =====
            # 重要：使用字符串類型進行存儲，避免後續轉換問題
            df['是否為MRB'] = "FALSE"
            df['MRB狀態'] = "Normal inspection"
            df['MRB訊息'] = "無MRB標記"
            df['MRB內容'] = None
            df['MRB加時'] = 0
            
            # 檢查所有可能的MRB欄位
            mrb_found = False
            for mrb_col in possible_mrb_columns:
                if mrb_col in df.columns:
                    # 創建掩碼標記非空值的MRB
                    mrb_mask = df[mrb_col].notna() & (df[mrb_col].astype(str).str.strip() != '')
                    if mrb_mask.any():
                        # 明確使用字符串"TRUE"而非布爾值True
                        df.loc[mrb_mask, '是否為MRB'] = "TRUE"
                        df.loc[mrb_mask, 'MRB狀態'] = "MRB"
                        df.loc[mrb_mask, 'MRB訊息'] = f"異常問題欄位({mrb_col})有內容"
                        df.loc[mrb_mask, 'MRB內容'] = df.loc[mrb_mask, mrb_col]
                        df.loc[mrb_mask, 'MRB加時'] = 30
                        mrb_found = True
                        debug_log(f"在欄位 {mrb_col} 找到 {mrb_mask.sum()} 筆MRB記錄")
            
            # 批量處理標準工時和檢驗耗時
            std_time_col = find_column_by_mapping(df, '檢驗標準工時', field_mapping)
            insp_time_col = find_column_by_mapping(df, '檢驗耗時（調整後）', field_mapping)
            
            # 向量化處理標準工時
            if std_time_col:
                df['處理後檢驗標準工時'] = pd.to_numeric(df[std_time_col], errors='coerce').fillna(0)
            else:
                df['處理後檢驗標準工時'] = 0
            
            # 向量化處理檢驗耗時
            if insp_time_col:
                df['檢驗耗時'] = pd.to_numeric(df[insp_time_col], errors='coerce').fillna(0)
            else:
                df['檢驗耗時'] = 0
            
            # 批量處理MRB加時 - 使用字符串比較
            mrb_mask = df['是否為MRB'] == "TRUE"
            df.loc[mrb_mask, '處理後檢驗標準工時'] += 30
            
            # 批量計算效率比值 - 向量化操作
            df['效率比值'] = 0
            
            # 處理不同情況
            zero_std_mask = df['處理後檢驗標準工時'] == 0
            zero_insp_mask = df['檢驗耗時'] <= 0.1
            
            # 標準工時為0的情況
            df.loc[zero_std_mask, '效率比值'] = 1
            
            # 檢驗耗時極小的情況
            df.loc[~zero_std_mask & zero_insp_mask, '效率比值'] = 0
            
            # 正常計算的情況
            normal_calc_mask = ~zero_std_mask & ~zero_insp_mask
            df.loc[normal_calc_mask, '效率比值'] = df.loc[normal_calc_mask, '處理後檢驗標準工時'] / df.loc[normal_calc_mask, '檢驗耗時']
            
            # 限制最大效率比值
            df.loc[df['效率比值'] > 20, '效率比值'] = 20
            
            # 批量處理其他欄位
            category_col = find_column_by_mapping(df, '類別', field_mapping)
            sample_status_col = find_column_by_mapping(df, '抽樣狀態', field_mapping)
            part_no_col = find_column_by_mapping(df, '料號', field_mapping)
            sample_qty_col = find_column_by_mapping(df, '抽樣數量', field_mapping)
            date_col = find_column_by_mapping(df, '檢驗日期', field_mapping)
            
            # 使用向量化操作處理各欄位
            df['類別'] = df[category_col] if category_col in df.columns else 'Unknown'
            df['抽樣狀態'] = df[sample_status_col] if sample_status_col in df.columns else ''
            df['料號'] = df[part_no_col] if part_no_col in df.columns else ''
            
            # 處理抽樣數量
            if sample_qty_col in df.columns:
                df['抽樣數量'] = pd.to_numeric(df[sample_qty_col], errors='coerce').fillna(1).astype(int)
            else:
                df['抽樣數量'] = 1
            
            # 處理檢驗日期
            if date_col in df.columns:
                df['檢驗日期'] = pd.to_datetime(df[date_col], errors='coerce')
            
            # 添加索引和檔案來源標記
            df['_index'] = range(len(df))
            df['檔案來源'] = file.name
            
            # 過濾掉抽樣狀態為 STS 的資料和包含 WYLZ 的資料
            filtered_df = df[(df['抽樣狀態'] != 'STS') & (~df['包含WYLZ'])]
            
            # 選取需要的欄位
            required_columns = [
                '處理後檢驗員', '處理後檢驗標準工時', '檢驗耗時', '效率比值', 
                '類別', '抽樣狀態', '料號', '抽樣數量', '檢驗日期', '檢驗開始時間',
                '包含WYLZ', '是否為MRB', 'MRB狀態', 'MRB訊息', 'MRB內容', 'MRB加時', '_index', '檔案來源'
            ]
            
            # 確保所有需要的欄位都存在
            for col in required_columns:
                if col not in filtered_df.columns:
                    filtered_df[col] = None
            
            # 檢查MRB狀態是否正確存在
            if mrb_found:
                mrb_check = filtered_df['是否為MRB'] == "TRUE"
                debug_log(f"過濾後仍有 {mrb_check.sum()} 筆MRB記錄", level="INFO")
            
            # 輸出表格前20行的MRB狀態統計以便調試
            status_counts = filtered_df.head(20)['MRB狀態'].value_counts()
            debug_log(f"頭20行MRB狀態統計: {status_counts.to_dict()}", level="INFO")
            
            all_data_frames.append(filtered_df[required_columns])
        
        # 合併所有資料框
        if all_data_frames:
            processed_df = pd.concat(all_data_frames, ignore_index=True)
            debug_log(f"所有IQC Report檔案處理完成，總資料列數: {len(processed_df)}")
            
            # 檢查合併後MRB狀態
            mrb_counts = processed_df['MRB狀態'].value_counts()
            debug_log(f"合併後MRB狀態統計: {mrb_counts.to_dict()}", level="INFO")
            
            # 檢查合併後結果的欄位名是否正確
            debug_log(f"合併後結果欄位名: {processed_df.columns.tolist()}")
            
            # 確保是否為MRB欄位為字符串類型，避免後續轉換問題
            processed_df['是否為MRB'] = processed_df['是否為MRB'].astype(str)
            
            # 最後檢查資料
            true_count = (processed_df['是否為MRB'] == "TRUE").sum()
            mrb_status_count = (processed_df['MRB狀態'] == "MRB").sum()
            
            debug_log(f"最終結果: 是否為MRB=TRUE的記錄數: {true_count}", level="INFO")
            debug_log(f"最終結果: MRB狀態=MRB的記錄數: {mrb_status_count}", level="INFO")
            
            return processed_df
        else:
            return pd.DataFrame()
    
    except Exception as e:
        error_msg = f"處理 IQC Report 時出錯: {str(e)}\n{traceback.format_exc()}"
        debug_log(error_msg)
        st.error(error_msg)
        raise e

# 輔助函數：根據映射尋找相應欄位
def find_column_by_mapping(df, field_name, mapping):
    """
    根據映射表尋找DataFrame中對應的欄位名稱
    
    參數:
    df - DataFrame
    field_name - 要尋找的欄位標準名稱
    mapping - 欄位映射字典
    
    返回:
    找到的欄位名稱或None
    """
    if field_name in mapping:
        possible_names = mapping[field_name]
        for name in possible_names:
            if name in df.columns:
                return name
                
        # 不區分大小寫嘗試
        for col in df.columns:
            if isinstance(col, str):
                for name in possible_names:
                    if name.lower() == col.lower():
                        return col
    return None

@st.cache_data(ttl=3600, max_entries=5, show_spinner=False)
def process_multiple_pcb_specs(files):
    try:
        debug_log(f"開始處理{len(files)}個PCB建檔明細檔案", level="INFO")
        all_data = []
        
        for file_idx, file in enumerate(files):
            debug_log(f"處理第{file_idx+1}個PCB建檔明細檔案: {file.name}", level="INFO")
            
            # 讀取Excel檔案，尋找目標工作表
            try:
                xls = pd.ExcelFile(file)
                sheets = xls.sheet_names
                
                # 智能工作表選擇 - 優先選擇包含關鍵字的工作表
                target_sheet = None
                
                # 按優先順序尋找工作表
                for keyword in ["建立規格_總表", "建立規格", "總表", "規格"]:
                    matches = [sheet for sheet in sheets if keyword in sheet]
                    if matches:
                        target_sheet = matches[0]
                        break
                
                # 如果沒找到，使用第一個工作表
                if not target_sheet:
                    target_sheet = sheets[0]
                
                debug_log(f"使用工作表: {target_sheet}", level="INFO")
                
                # 直接讀取資料，不進行列名處理
                df = pd.read_excel(file, sheet_name=target_sheet)
                
                # 快速定位關鍵欄位 - 不需要進行完整的列名轉換
                key_columns = {
                    'C': 'part_no',         # 料號 (C欄)
                    'N': 'hole_count',      # 壓合孔數 (N欄)
                    'L': 'hole_count_alt',  # 替代壓合孔數位置 (L欄)
                    'AB': 'length',         # 版長 (AB欄)
                    'AE': 'width'           # 版寬 (AE欄)
                }
                
                # 創建結果資料框 - 只保留必要欄位
                result_df = pd.DataFrame()
                
                # 提取料號 (C欄) - 必要欄位
                if 2 < df.shape[1]:  # 確保C欄存在
                    result_df['料號'] = df.iloc[:, 2].copy()
                    result_df['C'] = df.iloc[:, 2].copy()
                else:
                    debug_log("找不到C欄 (料號)，跳過此檔案", level="WARNING")
                    continue
                
                # 提取壓合孔數 (嘗試N欄，如果不存在則使用L欄)
                if 13 < df.shape[1]:  # N欄 (第14列)
                    result_df['壓合孔數'] = df.iloc[:, 13].copy()
                    result_df['N'] = df.iloc[:, 13].copy()
                    debug_log(f"使用N欄位獲取壓合孔數", level="INFO")
                elif 11 < df.shape[1]:  # L欄 (第12列)
                    result_df['壓合孔數'] = df.iloc[:, 11].copy()
                    result_df['N'] = df.iloc[:, 11].copy()
                    debug_log(f"N欄位不存在，使用L欄位獲取壓合孔數", level="INFO")
                else:
                    result_df['壓合孔數'] = 'NA'
                    result_df['N'] = 'NA'
                
                # 提取版長 (AB欄，第28列)
                if 27 < df.shape[1]:
                    result_df['版長'] = df.iloc[:, 27].copy()
                    result_df['AB'] = df.iloc[:, 27].copy()
                else:
                    result_df['版長'] = 0
                    result_df['AB'] = 0
                
                # 提取版寬 (AE欄，第31列)
                if 30 < df.shape[1]:
                    result_df['版寬'] = df.iloc[:, 30].copy()
                    result_df['AE'] = df.iloc[:, 30].copy()
                else:
                    result_df['版寬'] = 0
                    result_df['AE'] = 0
                
                # 添加檔案來源標記
                result_df['檔案來源'] = file.name
                
                # 過濾掉料號為空的資料 - 向量化操作
                result_df = result_df[result_df['料號'].notna()].copy()
                
                # 壓合孔數處理 - 向量化處理NA值
                result_df['壓合孔數'] = result_df['壓合孔數'].apply(
                    lambda x: 'NA' if pd.isna(x) or x == 'NA' else x
                )
                
                # 添加到總結果
                all_data.append(result_df)
                debug_log(f"第{file_idx+1}個檔案處理完成，資料列數: {len(result_df)}", level="INFO")
            
            except Exception as e:
                error_msg = f"處理檔案 {file.name} 時出錯: {str(e)}"
                debug_log(error_msg, level="ERROR")
                continue
        
        # 合併所有處理後的資料
        if all_data:
            processed_df = pd.concat(all_data, ignore_index=True)
            debug_log(f"所有PCB建檔明細檔案處理完成，總資料列數: {len(processed_df)}", level="INFO")
            
            # 檢查關鍵欄位
            key_fields = ['C', 'N', 'AB', 'AE']
            missing_cols = [col for col in key_fields if col not in processed_df.columns]
            
            if missing_cols:
                debug_log(f"警告: 缺少關鍵Excel欄位: {missing_cols}", level="WARNING")
            else:
                debug_log("成功提取所有關鍵Excel欄位", level="INFO")
                
                # 只顯示少量樣本數據，減少日誌量
                sample_size = min(5, len(processed_df))
                for i in range(sample_size):
                    debug_log(f"樣本 {i+1}: 料號={processed_df.iloc[i]['C']}, 壓合孔數={processed_df.iloc[i]['N']}, " +
                             f"版長={processed_df.iloc[i]['AB']}, 版寬={processed_df.iloc[i]['AE']}")
            
            return processed_df
        else:
            debug_log("沒有成功處理任何PCB建檔明細檔案", level="WARNING")
            return pd.DataFrame()
    
    except Exception as e:
        error_msg = f"處理 PCB建檔明細 時出錯: {str(e)}\n{traceback.format_exc()}"
        debug_log(error_msg, level="ERROR")
        st.error(error_msg)
        raise e

# 優化的面積範圍解析函數，從PCBDEBUG4_WORK整合
@st.cache_data(ttl=3600, max_entries=100)
def parse_area_range(area_range_str):
    """將面積範圍字符串解析為最小值和最大值"""
    try:
        area_range_str = str(area_range_str) if pd.notna(area_range_str) else ""
        
        min_area = 0
        max_area = float('inf')
        
        if "小於" in area_range_str:
            # 例如: "小於100000"
            max_area = float(''.join(filter(str.isdigit, area_range_str)))
        elif "超過" in area_range_str:
            # 例如: "超過250000"
            min_area = float(''.join(filter(str.isdigit, area_range_str)))
        elif "到" in area_range_str:
            # 例如: "100000到150000"
            parts = area_range_str.split("到")
            if len(parts) == 2:
                min_area = float(''.join(filter(str.isdigit, parts[0].strip())))
                max_area = float(''.join(filter(str.isdigit, parts[1].strip())))
        
        return min_area, max_area, area_range_str
    except Exception as e:
        debug_log(f"解析面積範圍時出錯: {e}, 原始值: {area_range_str}", level="ERROR")
        return 0, float('inf'), area_range_str

@st.cache_data(ttl=3600, max_entries=5, show_spinner=False)
def process_multiple_pcb_standard_times(files):
    try:
        debug_log(f"開始處理{len(files)}個PCB標準工時對應表檔案")
        all_data = []
        
        for file_idx, file in enumerate(files):
            debug_log(f"處理第{file_idx+1}個PCB標準工時對應表檔案: {file.name}")
            
            # 讀取Excel檔案
            xls = pd.ExcelFile(file)
            sheets = xls.sheet_names
            debug_log(f"Excel檔案包含以下工作表: {sheets}")
            
            # 使用第一個工作表
            sheet_name = sheets[0]
            debug_log(f"使用工作表: {sheet_name}")
            
            # 讀取所有欄位，不轉換列名
            df = pd.read_excel(file, sheet_name=sheet_name, header=0)
            debug_log(f"原始資料讀取完成，資料列數: {len(df)}")
            
            # 檢查並顯示一些欄位名稱進行調試
            debug_log(f"資料欄位名稱: {list(df.columns)[:10]}...")
            
            # 保留原始列索引 (A, B, C, ...)
            alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
            excel_cols = {}
            
            for i, col in enumerate(df.columns):
                # 生成Excel列名 (A, B, ..., Z, AA, AB, ...)
                if i < 26:
                    excel_col = alphabet[i]
                else:
                    excel_col = alphabet[i // 26 - 1] + alphabet[i % 26]
                
                excel_cols[col] = excel_col
                debug_log(f"列 {i}: {col} -> {excel_col}")
            
            # 將DataFrame的列重命名為Excel列名
            new_columns = {}
            for i, col in enumerate(df.columns):
                if i < 26:
                    new_columns[col] = alphabet[i]
                else:
                    new_columns[col] = alphabet[i // 26 - 1] + alphabet[i % 26]
            
            df = df.rename(columns=new_columns)
            
            # 準備處理後的數據
            processed_data = []
            
            # 將DataFrame轉換為字典列表
            records = df.to_dict('records')
            
            for idx, row in enumerate(records):
                # 只處理前幾筆進行調試
                if idx < 5:
                    debug_log(f"處理第 {idx+1} 筆資料: {row}")
                
                # 面積範圍處理 - 使用B欄
                area_range = row.get('B', '')
                
                # 抽樣數量處理 - 使用C欄
                sample_qty = row.get('C', '')
                
                # 壓合總孔數處理 - 使用D欄
                hole_count = row.get('D', '')
                
                # PCB標準工時處理 - 使用G欄 "檢驗工時_AI預測值 (1203版)"
                standard_time_field = row.get('G', 120)
                try:
                    standard_time = float(standard_time_field) if standard_time_field else 120
                    # 檢查標準工時單位，確保是分鐘
                    if standard_time < 10:  # 假設如果值小於10，可能是小時單位
                        debug_log(f"標準工時疑似為小時單位: {standard_time}，轉換為分鐘: {standard_time * 60}")
                        standard_time *= 60  # 轉換為分鐘
                except (ValueError, TypeError):
                    debug_log(f"無法轉換PCB標準工時值: {standard_time_field}，設置為120分鐘")
                    standard_time = 120
                
                processed_row = {
                    '面積範圍': area_range,
                    '抽樣數量': sample_qty,
                    '壓合總孔數': hole_count,
                    'PCB標準工時': standard_time,
                    'B': area_range,  # 保存原始Excel欄位數據
                    'C': sample_qty,
                    'D': hole_count,
                    'G': standard_time,
                    '檔案來源': file.name  # 添加檔案來源標記
                }
                
                # 合併原始資料和處理後的資料，保留所有Excel欄位
                for key, value in row.items():
                    if key not in processed_row:
                        processed_row[key] = value
                
                processed_data.append(processed_row)
            
            # 將當前檔案的處理結果添加到總結果中
            all_data.extend(processed_data)
            debug_log(f"第{file_idx+1}個檔案處理完成，累計資料筆數: {len(all_data)}")
        
        # 轉換為DataFrame
        processed_df = pd.DataFrame(all_data)
        debug_log(f"所有PCB標準工時對應表檔案處理完成，總資料列數: {len(processed_df)}")
        
        # 檢查是否成功提取了關鍵欄位
        if 'B' in processed_df.columns and 'D' in processed_df.columns and 'G' in processed_df.columns:
            debug_log("成功提取所有關鍵Excel欄位")
            
            # 顯示幾個樣本檢查數據
            for i in range(min(5, len(processed_df))):
                debug_log(f"樣本 {i+1}: 面積範圍={processed_df.iloc[i]['B']}, 壓合總孔數={processed_df.iloc[i]['D']}, " +
                         f"標準工時={processed_df.iloc[i]['G']}")
        else:
            missing_cols = []
            if 'B' not in processed_df.columns: missing_cols.append('B')
            if 'D' not in processed_df.columns: missing_cols.append('D')
            if 'G' not in processed_df.columns: missing_cols.append('G')
            debug_log(f"警告: 缺少關鍵Excel欄位: {missing_cols}")
        
        return processed_df
    
    except Exception as e:
        error_msg = f"處理 PCB標準工時對應表 時出錯: {str(e)}\n{traceback.format_exc()}"
        debug_log(error_msg)
        st.error(error_msg)
        raise e

@st.cache_data(ttl=3600, max_entries=5, show_spinner=False)
def process_multiple_additional_tasks(files):
    try:
        debug_log(f"開始處理{len(files)}個IQC額外任務紀錄清單檔案")
        all_data = []
        
        for file_idx, file in enumerate(files):
            debug_log(f"處理第{file_idx+1}個IQC額外任務紀錄清單檔案: {file.name}")
            
            # 读取Excel文件
            xls = pd.ExcelFile(file)
            sheets = xls.sheet_names
            debug_log(f"Excel檔案包含以下工作表: {sheets}")
            
            # 使用第一个工作表
            sheet_name = sheets[0]
            debug_log(f"使用工作表: {sheet_name}")
            
            # 嘗試不同的讀取方法
            try:
                # 先完全不指定header，得到原始數據
                raw_df = pd.read_excel(file, sheet_name=sheet_name, header=None)
                debug_log(f"原始數據前5行:\n{raw_df.head()}")
                
                # 嘗試偵測標題行 - 檢查前5行
                header_row = None
                for i in range(min(5, len(raw_df))):
                    row_str = ' '.join([str(x) for x in raw_df.iloc[i].values])
                    debug_log(f"第{i}行內容: {row_str}")
                    
                    # 如果該行包含關鍵字，可能是標題行
                    if '姓名' in row_str or '工作事項分類' in row_str or '用時' in row_str:
                        header_row = i
                        debug_log(f"偵測到第{i}行可能是標題行: {row_str}")
                        break
                
                # 使用偵測到的標題行或預設使用第0行
                if header_row is not None:
                    df = pd.read_excel(file, sheet_name=sheet_name, header=header_row)
                    debug_log(f"使用第{header_row}行作為標題")
                else:
                    df = pd.read_excel(file, sheet_name=sheet_name)
                    debug_log("使用預設標題行")
                
                debug_log(f"處理後資料欄位: {list(df.columns)}")
                
            except Exception as e:
                debug_log(f"標題偵測失敗，使用預設方式讀取: {e}")
                df = pd.read_excel(file, sheet_name=sheet_name)
            
            debug_log(f"原始資料讀取完成，資料列數: {len(df)}")
            
            # 使用更多的欄位名稱備選方案
            field_mapping = {
                '姓名': ['姓名', 'Name', '下拉式選單', '人員', 'B', '檢驗員', 'Inspector', '檢驗人員'],
                '日期': ['日期', 'Date', '手key', 'A', '檢驗日期', '任務日期', '記錄日期'],
                '工作事項分類': ['工作事項分類', 'Task Type', '下拉式選單_1', '任務類型', 'C', '工作項目', '事項分類'],
                '用時(分鐘)': ['用時(分鐘)', '用時（分鐘）', 'Time(min)', '時間', '手key_3', 'H', '工時', '分鐘']
            }
            
            # 準備處理後的數據
            processed_data = []
            
            # 轉換為字典列表
            records = df.to_dict('records')
            
            # 檢查是否使用Excel列標識（A, B, C...）
            excel_columns_map = {
                'A': '日期',
                'B': '姓名',
                'C': '工作事項分類',
                'H': '用時(分鐘)'
            }
            has_excel_columns = any(col in df.columns for col in ['A', 'B', 'C', 'H'])
            debug_log(f"是否含有Excel欄位標識: {has_excel_columns}")
            
            for idx, row in enumerate(records):
                # 只處理前幾筆進行調試
                if idx < 5:
                    debug_log(f"處理第 {idx+1} 筆資料: {row}")
                
                # 嘗試從不同來源獲取資料
                name = None
                date_value = None
                task_type = None
                time_value = None
                
                # 1. 直接使用欄位名稱
                for field, possible_names in field_mapping.items():
                    for possible_name in possible_names:
                        if possible_name in row and pd.notna(row[possible_name]):
                            if field == '姓名' and name is None:
                                name = row[possible_name]
                            elif field == '日期' and date_value is None:
                                date_value = row[possible_name]
                            elif field == '工作事項分類' and task_type is None:
                                task_type = row[possible_name]
                            elif field == '用時(分鐘)' and time_value is None:
                                time_value = row[possible_name]
                
                # 2. 使用Excel欄位標識
                if has_excel_columns:
                    if name is None and 'B' in row and pd.notna(row['B']):
                        name = row['B']
                    if date_value is None and 'A' in row and pd.notna(row['A']):
                        date_value = row['A']
                    if task_type is None and 'C' in row and pd.notna(row['C']):
                        task_type = row['C']
                    if time_value is None and 'H' in row and pd.notna(row['H']):
                        time_value = row['H']
                
                # 3. 使用數字索引（如果沒有名稱）
                if name is None and 1 in row and pd.notna(row[1]):
                    name = row[1]
                if date_value is None and 0 in row and pd.notna(row[0]):
                    date_value = row[0]
                if task_type is None and 2 in row and pd.notna(row[2]):
                    task_type = row[2]
                if time_value is None and 7 in row and pd.notna(row[7]):
                    time_value = row[7]
                
                # 檢查是否獲取到了有效資訊
                if name is None or pd.isna(name) or name == '姓名' or name == '下拉式選單':
                    continue
                    
                # 處理日期
                parsed_date = None
                if date_value is not None:
                    parsed_date = parse_excel_date(date_value)
                    if parsed_date is None:
                        debug_log(f"無法解析日期: {date_value}")
                        # 嘗試使用文件名稱中的日期
                        import re
                        date_match = re.search(r'(\d{1,2}[-/]\d{1,2})', file.name)
                        if date_match:
                            date_str = date_match.group(1)
                            try:
                                # 假設格式為月/日或月-日
                                current_year = datetime.now().year
                                if '/' in date_str:
                                    month, day = map(int, date_str.split('/'))
                                else:
                                    month, day = map(int, date_str.split('-'))
                                parsed_date = datetime(current_year, month, day)
                                debug_log(f"從檔名取得日期: {parsed_date}")
                            except:
                                debug_log("從檔名取得日期失敗")
                
                # 處理時間值
                try:
                    time_minutes = 0
                    if time_value is not None and pd.notna(time_value):
                        # 處理不同格式的時間值
                        if isinstance(time_value, str):
                            # 移除非數字字符
                            time_value = ''.join(c for c in time_value if c.isdigit() or c == '.')
                            if time_value:
                                time_minutes = float(time_value)
                        else:
                            time_minutes = float(time_value)
                except (ValueError, TypeError) as e:
                    debug_log(f"處理時間值出錯: {e}, 原始值: {time_value}")
                    time_minutes = 0
                
                # 整合資料
                processed_row = {
                    '姓名': name,
                    '日期': parsed_date,
                    '工作事項分類': task_type if task_type is not None else 'Other',
                    '用時(分鐘)': time_minutes,
                    '_index': idx,
                    '檔案來源': file.name
                }
                
                processed_data.append(processed_row)
                
                if idx < 5:
                    debug_log(f"處理結果: {processed_row}")
            
            # 將當前檔案的處理結果添加到總結果中
            all_data.extend(processed_data)
            debug_log(f"第{file_idx+1}個檔案處理完成，累計資料筆數: {len(all_data)}")
        
        # 轉換為DataFrame
        df_processed = pd.DataFrame(all_data)
        debug_log(f"所有IQC額外任務紀錄清單檔案處理完成，總資料列數: {len(df_processed)}")
        
        # 過濾掉無效的數據
        if not df_processed.empty:
            df_filtered = df_processed[(df_processed['姓名'] != 'Unknown') & 
                                      (df_processed['姓名'] != '姓名') & 
                                      (df_processed['姓名'] != '下拉式選單')]
            
            debug_log(f"過濾後資料列數: {len(df_filtered)}")
            
            # 確保所有日期都是datetime格式
            if '日期' in df_filtered.columns:
                df_filtered['日期'] = pd.to_datetime(df_filtered['日期'], errors='coerce')
                
                # 調試日期範圍
                if not df_filtered.empty:
                    min_date = df_filtered['日期'].min()
                    max_date = df_filtered['日期'].max()
                    debug_log(f"額外任務數據日期範圍: {min_date} 到 {max_date}")
            
            return df_filtered
        else:
            debug_log("處理後沒有有效資料")
            return pd.DataFrame()
    
    except Exception as e:
        error_msg = f"處理 IQC額外任務紀錄清單 時出錯: {str(e)}\n{traceback.format_exc()}"
        debug_log(error_msg)
        st.error(error_msg)
        return pd.DataFrame()  # 返回空DataFrame而不是拋出異常，避免中斷程序

def map_hole_count_to_range(hole_counts):
    """
    將實際壓合孔數映射到PCB標準工時對應表中的範圍值，添加記憶化以減少重複計算
    """
    try:
        # 處理空值或NA值
        if hole_counts is None or hole_counts == 'NA' or pd.isna(hole_counts):
            return 0
            
        # 確保孔數是數字
        try:
            hole_counts = float(hole_counts)
        except (ValueError, TypeError):
            debug_log(f"壓合孔數無法轉換為數字: '{hole_counts}'，設為0", level="WARNING")
            return 0
        
        # 使用查表法代替多個if判斷，提高效率
        ranges = [
            (0, 25, 0),
            (25, 75, 50),
            (75, 125, 100),
            (125, 175, 150),
            (175, 200, 200),
            (200, 250, 225),
            (250, 325, 300),
            (325, 450, 400),
            (450, 550, 500),
            (550, 750, 600),
            (750, 900, 800),
            (900, 1000, 1000),
            (1000, float('inf'), 1100)
        ]
        
        for min_val, max_val, mapped_val in ranges:
            if hole_counts <= max_val:
                return mapped_val
                
        return 1100  # 默認值
        
    except Exception as e:
        debug_log(f"映射壓合孔數時出錯: {e}", level="ERROR")
        return 0

# 完整替換calculate_pcb_standard_time函數中的MRB處理邏輯

def calculate_pcb_standard_time(iqc_df, pcb_specs_df, pcb_standard_time_df):
    """
    全面修正版的PCB標準工時計算函數，徹底修復MRB判斷和加時
    """
    try:
        debug_log("開始計算PCB標準工時", level="INFO")
        
        # 創建數據副本
        processed_df = iqc_df.copy()
        
        # 只處理QB類型的料號
        qb_indices = processed_df[processed_df['類別'] == 'QB'].index
        debug_log(f"發現QB類型料號數量: {len(qb_indices)}", level="INFO")
        
        # 如果沒有QB類型料號，則直接返回
        if len(qb_indices) == 0:
            debug_log("沒有發現QB類型料號，跳過PCB標準工時計算", level="INFO")
            return processed_df
        
        # 提取QB類型資料用於批量處理
        qb_df = processed_df.loc[qb_indices].copy()
        
        # 輸出欄位名稱，幫助調試
        debug_log(f"QB資料欄位: {qb_df.columns.tolist()}", level="INFO")
        
        # 首先，一次性確定所有MRB狀態，避免逐行判斷帶來的不一致
        debug_log("重新檢查所有QB記錄的MRB狀態", level="INFO")
        
        # 檢查是否已經有MRB狀態欄位
        if '是否為MRB' in qb_df.columns:
            # 統一轉換現有的MRB狀態為布爾值，確保一致性
            converted_mrb = []
            for i, row in qb_df.iterrows():
                mrb_value = row['是否為MRB']
                
                # 記錄原始值類型
                original_type = type(mrb_value).__name__
                converted = False  # 默認為False
                
                # 針對不同類型進行特定處理
                if isinstance(mrb_value, bool):
                    converted = mrb_value  # 已經是布爾值，不變
                elif isinstance(mrb_value, str):
                    # 字符串值，只有明確的true才算True
                    converted = mrb_value.upper() in ('TRUE', 'T', 'YES', 'Y', '1', 'MRB')
                elif isinstance(mrb_value, (int, float)):
                    # 數值，非零即True
                    converted = bool(mrb_value) 
                
                converted_mrb.append(converted)
                
                debug_log(f"料號 {row.get('料號', '')} MRB值轉換: {mrb_value}({original_type}) -> {converted}", level="DEBUG")
        else:
            # 如果沒有MRB狀態欄位，則使用MRB檢測函數
            debug_log("未找到MRB狀態欄位，執行MRB檢測", level="INFO")
            mrb_result = check_is_mrb(qb_df)
            converted_mrb = mrb_result['是否為MRB'].tolist()
        
        # 將轉換後的MRB狀態保存回DataFrame，確保一致性
        qb_df['是否為MRB'] = converted_mrb
        
        # 顯示MRB狀態分佈
        mrb_counts = pd.Series(converted_mrb).value_counts()
        debug_log(f"MRB狀態分佈: {mrb_counts.to_dict()}", level="INFO")
        
        # 1. 建立料號與PCB信息的對應關係
        debug_log("建立料號與PCB規格的對應關係", level="INFO")
        pcb_info = {}
        
        # 確定料號欄位
        part_no_col = None
        if 'C' in pcb_specs_df.columns:
            part_no_col = 'C'
        elif '料號' in pcb_specs_df.columns:
            part_no_col = '料號'
        
        # 確定其他欄位
        hole_count_col = 'N' if 'N' in pcb_specs_df.columns else ('壓合孔數' if '壓合孔數' in pcb_specs_df.columns else 'L')
        length_col = 'AB' if 'AB' in pcb_specs_df.columns else '版長'
        width_col = 'AE' if 'AE' in pcb_specs_df.columns else '版寬'
        
        # 建立PCB信息字典
        for _, row in pcb_specs_df.iterrows():
            if part_no_col in row and pd.notna(row[part_no_col]):
                part_no = str(row[part_no_col]).strip().upper()
                
                # 獲取壓合孔數
                hole_count = 0
                if hole_count_col in row and pd.notna(row[hole_count_col]):
                    hole_count_val = row[hole_count_col]
                    if str(hole_count_val).upper() == 'NA':
                        hole_count = 0
                    else:
                        try:
                            hole_count = float(hole_count_val)
                        except:
                            hole_count = 0
                
                # 獲取板長和板寬
                length = 0
                width = 0
                if length_col in row and pd.notna(row[length_col]):
                    try:
                        length = float(row[length_col])
                    except:
                        pass
                
                if width_col in row and pd.notna(row[width_col]):
                    try:
                        width = float(row[width_col])
                    except:
                        pass
                
                # 計算面積
                area = length * width
                
                # 映射壓合孔數
                mapped_hole_count = map_hole_count_to_range(hole_count)
                
                # 儲存PCB信息
                pcb_info[part_no] = {
                    'hole_count': hole_count,
                    'mapped_hole_count': mapped_hole_count,
                    'length': length,
                    'width': width,
                    'area': area
                }
        
        debug_log(f"已建立 {len(pcb_info)} 個料號的PCB信息", level="INFO")
        
        # 2. 建立面積範圍和標準工時對應
        debug_log("解析PCB標準工時對應表", level="INFO")
        area_ranges = []
        
        # 確認標準工時對應表中的關鍵欄位
        area_range_col = 'B' if 'B' in pcb_standard_time_df.columns else '面積範圍'
        hole_count_col = 'D' if 'D' in pcb_standard_time_df.columns else '壓合總孔數'
        std_time_col = 'G' if 'G' in pcb_standard_time_df.columns else 'PCB標準工時'
        
        # 解析所有面積範圍
        for idx, row in pcb_standard_time_df.iterrows():
            if area_range_col in row and pd.notna(row[area_range_col]):
                min_area, max_area, area_range_str = parse_area_range(row[area_range_col])
                
                # 獲取孔數和標準工時
                try:
                    hole_count = float(row[hole_count_col]) if hole_count_col in row and pd.notna(row[hole_count_col]) else None
                    std_time = float(row[std_time_col]) if std_time_col in row and pd.notna(row[std_time_col]) else 120
                except (ValueError, TypeError):
                    hole_count = None
                    std_time = 120
                
                area_ranges.append({
                    'min_area': min_area,
                    'max_area': max_area,
                    'range_str': area_range_str,
                    'hole_count': hole_count,
                    'std_time': std_time
                })
        
        debug_log(f"已解析 {len(area_ranges)} 個面積範圍", level="INFO")
        
        # 建立查找表
        area_hole_lookup = {}
        for ar in area_ranges:
            range_key = (ar['min_area'], ar['max_area'])
            if range_key not in area_hole_lookup:
                area_hole_lookup[range_key] = {}
            
            if ar['hole_count'] not in area_hole_lookup[range_key]:
                area_hole_lookup[range_key][ar['hole_count']] = ar['std_time']
        
        # 3. 處理每個QB料號
        debug_log("開始處理每個QB料號的標準工時", level="INFO")
        
        # 添加結果列
        qb_df['面積'] = 0
        qb_df['壓合孔數'] = 'NA'
        qb_df['映射壓合孔數'] = 0
        qb_df['匹配狀態'] = '未處理'
        qb_df['基礎標準工時'] = 120  # 默認值
        qb_df['MRB加時'] = 0
        qb_df['處理後檢驗標準工時'] = 120  # 默認值
        qb_df['匹配詳情'] = ''
        
        match_count = 0
        mrb_count = 0
        
        # 遍歷每個QB料號
        for idx, row in qb_df.iterrows():
            part_no = str(row['料號']).strip().upper()
            
            # 查找PCB信息
            if part_no in pcb_info:
                # 獲取PCB數據
                pcb_data = pcb_info[part_no]
                area = pcb_data['area']
                hole_count = pcb_data['hole_count']
                mapped_hole_count = pcb_data['mapped_hole_count']
                
                # 更新結果數據
                qb_df.at[idx, '面積'] = area
                qb_df.at[idx, '壓合孔數'] = hole_count
                qb_df.at[idx, '映射壓合孔數'] = mapped_hole_count
                
                # 查找匹配的面積範圍
                matched_range = None
                matched_range_str = None
                
                for ar in area_ranges:
                    if ar['min_area'] <= area < ar['max_area']:
                        matched_range = (ar['min_area'], ar['max_area'])
                        matched_range_str = ar['range_str']
                        break
                
                if matched_range is None:
                    # 面積範圍未匹配
                    qb_df.at[idx, '匹配狀態'] = '面積範圍未匹配'
                    qb_df.at[idx, '匹配詳情'] = f"面積 {area} 未找到匹配範圍"
                    
                    # 使用預設標準工時
                    base_std_time = 120
                    qb_df.at[idx, '基礎標準工時'] = base_std_time
                else:
                    # 面積範圍匹配，檢查孔數
                    if matched_range in area_hole_lookup and mapped_hole_count in area_hole_lookup[matched_range]:
                        # 孔數也匹配
                        base_std_time = area_hole_lookup[matched_range][mapped_hole_count]
                        
                        qb_df.at[idx, '基礎標準工時'] = base_std_time
                        qb_df.at[idx, '匹配狀態'] = '匹配成功'
                        qb_df.at[idx, '匹配詳情'] = f"面積: {matched_range_str}, 孔數: {mapped_hole_count}, 基礎標準工時: {base_std_time}"
                        match_count += 1
                    else:
                        # 面積匹配但孔數未匹配
                        base_std_time = 120
                        qb_df.at[idx, '基礎標準工時'] = base_std_time
                        qb_df.at[idx, '匹配狀態'] = '孔數未匹配'
                        qb_df.at[idx, '匹配詳情'] = f"面積範圍匹配: {matched_range_str}, 但壓合孔數 {mapped_hole_count} 未匹配"
            else:
                # 料號未找到
                base_std_time = 120
                qb_df.at[idx, '基礎標準工時'] = base_std_time
                qb_df.at[idx, '匹配狀態'] = '料號未找到'
                qb_df.at[idx, '匹配詳情'] = f"料號 {part_no} 在PCB建檔明細中未找到"
            
            # 處理MRB加時 - 徹底修正的邏輯
            is_mrb = qb_df.at[idx, '是否為MRB']  # 已統一轉換為布爾值
            
            mrb_add_time = 0
            if is_mrb:
                mrb_add_time = 30
                mrb_count += 1
                debug_log(f"料號 {part_no} 是MRB，添加30分鐘標準工時", level="DEBUG")
            
            # 更新MRB相關欄位
            qb_df.at[idx, 'MRB加時'] = mrb_add_time
            qb_df.at[idx, '處理後檢驗標準工時'] = base_std_time + mrb_add_time
        
        # 在返回前進行一次最終檢查
        debug_log("進行最終MRB計算檢查", level="INFO")
        for idx, row in qb_df.iterrows():
            # 檢查MRB狀態和加時是否一致
            is_mrb = bool(row['是否為MRB'])
            expected_add_time = 30 if is_mrb else 0
            actual_add_time = row['MRB加時']
            
            if expected_add_time != actual_add_time:
                debug_log(f"不一致警告：料號 {row['料號']} MRB狀態={is_mrb} 但加時={actual_add_time}", level="WARNING")
                # 修正不一致
                qb_df.at[idx, 'MRB加時'] = expected_add_time
                qb_df.at[idx, '處理後檢驗標準工時'] = row['基礎標準工時'] + expected_add_time
        
        # 更新到原始DataFrame
        for col in ['是否為MRB', '面積', '壓合孔數', '映射壓合孔數', '匹配狀態', 
                   '基礎標準工時', 'MRB加時', '處理後檢驗標準工時', '匹配詳情']:
            if col in qb_df.columns:
                processed_df.loc[qb_indices, col] = qb_df[col]
        
        # 輸出統計信息
        debug_log(f"PCB標準工時計算完成，共處理 {len(qb_df)} 筆QB類型料號，成功匹配 {match_count} 筆", level="INFO")
        debug_log(f"MRB狀態總數: {sum(converted_mrb)}, 加了MRB加時的記錄數: {mrb_count}", level="INFO")
        
        # 處理非QB類型物料的基礎標準工時
        # 對於非QB類型的物料，將處理後檢驗標準工時作為基礎標準工時，並考慮MRB加時
        non_qb_indices = processed_df[processed_df['類別'] != 'QB'].index
        if len(non_qb_indices) > 0:
            debug_log(f"處理 {len(non_qb_indices)} 筆非QB類型物料的基礎標準工時", level="INFO")
            
            # 先檢查是否已經有基礎標準工時欄位
            if '基礎標準工時' not in processed_df.columns:
                processed_df['基礎標準工時'] = None
            
            # 設置一般物料的基礎標準工時
            for idx in non_qb_indices:
                # 檢查MRB加時欄位是否存在
                mrb_time = 0
                if 'MRB加時' in processed_df.columns and pd.notna(processed_df.loc[idx, 'MRB加時']):
                    try:
                        mrb_time = float(processed_df.loc[idx, 'MRB加時'])
                    except:
                        mrb_time = 0
                
                # 從處理後檢驗標準工時中減去MRB加時得到基礎標準工時
                if '處理後檢驗標準工時' in processed_df.columns and pd.notna(processed_df.loc[idx, '處理後檢驗標準工時']):
                    try:
                        std_time = float(processed_df.loc[idx, '處理後檢驗標準工時'])
                        base_time = std_time - mrb_time
                        processed_df.loc[idx, '基礎標準工時'] = base_time
                    except:
                        # 如果轉換失敗，直接使用原值
                        processed_df.loc[idx, '基礎標準工時'] = processed_df.loc[idx, '處理後檢驗標準工時']

        return processed_df
    
    except Exception as e:
        error_msg = f"計算PCB標準工時時出錯: {str(e)}\n{traceback.format_exc()}"
        debug_log(error_msg, level="ERROR")
        return iqc_df  # 如果出錯，返回原始數據


@st.cache_data(ttl=1800, show_spinner=False)
def filter_by_date_range(df, start_date=None, end_date=None):
    """
    根據日期範圍過濾資料
    """
    if df is None or df.empty or (not start_date and not end_date):
        return df
    
    debug_log(f"根據日期範圍過濾: 開始日期={start_date}, 結束日期={end_date}")
    
    # 複製DataFrame以避免修改原始資料
    filtered_df = df.copy()
    
    # 確定日期欄位
    date_column = None
    for possible_column in ['檢驗日期', '日期']:
        if possible_column in filtered_df.columns:
            date_column = possible_column
            break
    
    if date_column is None:
        debug_log("找不到日期欄位，跳過日期過濾")
        return filtered_df
    
    debug_log(f"使用日期欄位: {date_column}")
    
    # 確保日期欄位是datetime類型
    try:
        # 先檢查類型
        debug_log(f"日期欄位類型: {filtered_df[date_column].dtype}")
        
        # 強制轉換日期欄位為datetime（不丟棄無法解析的值）
        filtered_df[date_column] = pd.to_datetime(filtered_df[date_column], errors='coerce')
        
        # 檢查轉換後有多少非NaT值
        valid_dates_count = filtered_df[date_column].notna().sum()
        debug_log(f"有效日期數量: {valid_dates_count}/{len(filtered_df)}")
        
        # 日期解析有問題時，打印一些樣本進行調試
        if valid_dates_count < len(filtered_df) * 0.5:
            debug_log("大部分日期轉換失敗，顯示部分原始值:")
            for i, val in enumerate(filtered_df[date_column].dropna().head()):
                debug_log(f"日期樣本{i+1}: {val}")
            
            # 顯示無法解析的日期樣本
            null_dates = filtered_df[filtered_df[date_column].isna()][date_column].head()
            debug_log(f"無法解析的日期樣本: {null_dates.index.tolist()}")
            
            # 再次嘗試使用自定義函數解析
            debug_log("嘗試使用自定義函數解析日期")
            filtered_df[date_column] = filtered_df[date_column].apply(lambda x: parse_excel_date(x))
        
        # 再次檢查有效日期數量
        valid_dates_count = filtered_df[date_column].notna().sum()
        debug_log(f"處理後有效日期數量: {valid_dates_count}/{len(filtered_df)}")
            
    except Exception as e:
        debug_log(f"轉換日期欄位時出錯: {e}")
        st.warning(f"轉換日期欄位時出錯，部分日期可能無法正確過濾")
        return filtered_df
    
    # 轉換開始和結束日期為datetime
    try:
        start_date_obj = pd.to_datetime(start_date) if start_date else None
        end_date_obj = pd.to_datetime(end_date) if end_date else None
        
        debug_log(f"轉換後的篩選日期: 開始={start_date_obj}, 結束={end_date_obj}")
        
        if end_date_obj:
            # 將結束日期設為當天的結束時間
            end_date_obj = end_date_obj.replace(hour=23, minute=59, second=59)
    except Exception as e:
        debug_log(f"轉換篩選日期時出錯: {e}")
        return filtered_df
    
    # 根據日期範圍過濾
    mask = pd.Series(True, index=filtered_df.index)
    
    if start_date_obj:
        debug_log(f"篩選大於等於 {start_date_obj} 的資料")
        mask = mask & (filtered_df[date_column] >= start_date_obj)
    
    if end_date_obj:
        debug_log(f"篩選小於等於 {end_date_obj} 的資料")
        mask = mask & (filtered_df[date_column] <= end_date_obj)
    
    result_df = filtered_df[mask]
    debug_log(f"過濾前資料列數: {len(filtered_df)}, 過濾後: {len(result_df)}")
    
    # 如果過濾後沒有資料，可能日期格式有問題
    if len(result_df) == 0 and len(filtered_df) > 0:
        debug_log("過濾後沒有資料，顯示日期最大最小值檢查:")
        min_date = filtered_df[date_column].min()
        max_date = filtered_df[date_column].max()
        debug_log(f"數據日期範圍: {min_date} 到 {max_date}")
        
        # 如果日期篩選導致沒有數據，返回原始數據並警告
        st.warning(f"日期篩選 ({start_date} 到 {end_date}) 沒有匹配的數據。數據實際日期範圍是 {min_date} 到 {max_date}。")
        return filtered_df
    
    return result_df

# 修正: 計算效率並剔除極值的函數，確保正確處理0%剔除情況
def calculate_efficiency_with_trimming(processed_df, trim_percentage=0):
    """
    計算檢驗效率並可選擇剔除一定百分比的極值數據
    
    參數:
    processed_df - 處理後的完整數據DataFrame
    trim_percentage - 要剔除的百分比（上下各剔除這個百分比的數據）
    
    返回:
    包含效率排名的DataFrame
    """
    debug_log(f"開始計算效率，剔除極值百分比: {trim_percentage}%")
    
    if processed_df is None or processed_df.empty:
        debug_log("無數據可供計算", level="WARNING")
        return pd.DataFrame()
    
    # 設定效率值的合理上限
    MAX_EFFICIENCY = 20
    
    # 如果不需要剔除極值，使用常規計算
    if trim_percentage <= 0:
        # 清除剔除效果相關狀態
        if 'using_trimmed_efficiency' in st.session_state:
            st.session_state.using_trimmed_efficiency = False
        if 'trimmed_avg_efficiency' in st.session_state:
            del st.session_state.trimmed_avg_efficiency
        if 'trimmed_record_stats' in st.session_state:
            del st.session_state.trimmed_record_stats
            
        # 使用標準計算函數
        return calculate_overall_efficiency(processed_df)
    
    # 準備存儲效率數據的字典
    inspector_efficiency = {}
    
    # 首先按檢驗員分組
    for inspector, inspector_data in processed_df.groupby('處理後檢驗員'):
        if pd.isna(inspector) or inspector == '':
            continue  # 跳過無效檢驗員

        # 初始化這個檢驗員的數據
        if inspector not in inspector_efficiency:
            inspector_efficiency[inspector] = {
                'total_standard_time': 0,
                'total_actual_time': 0,
                'record_count': 0,
                'category_counts': {},
                'raw_records': []
            }
        
        # 獲取這個檢驗員的所有記錄
        for _, row in inspector_data.iterrows():
            category = row['類別'] if pd.notna(row['類別']) else 'Unknown'
            
            # 處理標準工時
            try:
                standard_time = float(row['處理後檢驗標準工時']) if pd.notna(row['處理後檢驗標準工時']) else 0
            except (ValueError, TypeError):
                standard_time = 0
            
            # 處理實際檢驗時間
            try:
                actual_time = float(row['檢驗耗時']) if pd.notna(row['檢驗耗時']) else 0
            except (ValueError, TypeError):
                actual_time = 0
            
            # 確保實際時間不為0，避免除以零錯誤
            if actual_time < 0.1:
                actual_time = 0.1
            
            # 計算單筆記錄的效率比值
            efficiency_ratio = min(standard_time / actual_time, MAX_EFFICIENCY) if actual_time > 0 else 0
            
            # 儲存所有原始記錄
            inspector_efficiency[inspector]['raw_records'].append({
                'standard_time': standard_time,
                'actual_time': actual_time,
                'efficiency_ratio': efficiency_ratio,
                'category': category
            })
            
            # 更新類別統計
            if category not in inspector_efficiency[inspector]['category_counts']:
                inspector_efficiency[inspector]['category_counts'][category] = 0
            inspector_efficiency[inspector]['category_counts'][category] += 1
    
    # 現在對每個檢驗員的數據進行極值剔除
    overall_efficiency_data = []
    
    for inspector, data in inspector_efficiency.items():
        raw_records = data['raw_records']
        
        # 如果記錄過少，不進行剔除
        if len(raw_records) <= 2:
            # 簡單加總所有記錄
            total_standard_time = sum(record['standard_time'] for record in raw_records)
            total_actual_time = sum(record['actual_time'] for record in raw_records)
            record_count = len(raw_records)
        else:
            # 按效率比值排序
            sorted_records = sorted(raw_records, key=lambda x: x['efficiency_ratio'])
            
            # 計算要剔除的記錄數量
            trim_count = int(len(sorted_records) * (trim_percentage / 100))
            
            # 確保至少有一半的數據保留
            max_trim = len(sorted_records) // 2 - 1
            if trim_count * 2 > max_trim:
                trim_count = max_trim // 2
                debug_log(f"調整剔除數量為 {trim_count}，確保至少保留一半數據", level="INFO")
            
            # 剔除最低和最高的極值
            trimmed_records = sorted_records[trim_count:-trim_count] if trim_count > 0 else sorted_records
            
            # 計算剔除後的總和
            total_standard_time = sum(record['standard_time'] for record in trimmed_records)
            total_actual_time = sum(record['actual_time'] for record in trimmed_records)
            record_count = len(trimmed_records)
        
        # 確保實際時間不為0，避免除以零錯誤
        if total_actual_time < 0.1:
            total_actual_time = 0.1
        
        # 計算剔除極值後的效率
        efficiency = min(total_standard_time / total_actual_time, MAX_EFFICIENCY) if total_actual_time > 0 else 0
        
        overall_efficiency_data.append({
            'inspector': inspector,
            'efficiency': efficiency,
            'total_standard_time': total_standard_time,
            'total_actual_time': total_actual_time,
            'record_count': record_count,
            'category_counts': data['category_counts'],
            'original_record_count': len(raw_records),
            'trimmed_count': len(raw_records) - record_count if len(raw_records) > record_count else 0
        })
    
    # 按效率從高到低排序
    overall_efficiency_ranking = pd.DataFrame(overall_efficiency_data).sort_values(by='efficiency', ascending=False)
    debug_log(f"效率計算完成，共 {len(overall_efficiency_ranking)} 位檢驗員", level="INFO")
    
    # 計算並保存剔除後的平均效率
    avg_efficiency = overall_efficiency_ranking['efficiency'].mean()
    st.session_state.trimmed_avg_efficiency = avg_efficiency
    st.session_state.using_trimmed_efficiency = True  # 標記使用了剔除後的效率
    
    # 同時保存剔除的記錄數量信息
    total_original_records = sum(item['original_record_count'] for item in overall_efficiency_data)
    total_trimmed_records = sum(item['trimmed_count'] for item in overall_efficiency_data)
    st.session_state.trimmed_record_stats = {
        'total_original': total_original_records,
        'total_trimmed': total_trimmed_records,
        'trim_percentage': trim_percentage
    }
    
    return overall_efficiency_ranking

@st.cache_data(ttl=1800, show_spinner=False)
def calculate_efficiency_metrics(processed_df, start_date=None, end_date=None, selected_material_categories=None, merge_categories=False):
    """優化的效率指標計算函數"""
    debug_log("開始計算IQC檢驗效率指標", level="INFO")
    
    # 過濾資料，確保在日期範圍內
    filtered_df = filter_by_date_range(processed_df, start_date, end_date)
    
    # 如果指定了物料類別，進行過濾
    if selected_material_categories and len(selected_material_categories) > 0:
        filtered_df = filtered_df[filtered_df['類別'].isin(selected_material_categories)]
        debug_log(f"根據物料類別過濾: {selected_material_categories}, 過濾後資料行數: {len(filtered_df)}", level="INFO")
    
    if filtered_df is None or filtered_df.empty:
        debug_log("過濾後沒有資料，返回空結果", level="WARNING")
        return {
            'overall_efficiency_ranking': pd.DataFrame(),
            'category_efficiency_data': {}
        }
    
    # 設定效率值的合理上限
    MAX_EFFICIENCY = 20
    
    # 一次性計算所有所需的統計數據
    try:
        # 使用 pandas 1.0.0+ 的 named aggregation，一次性完成聚合操作
        aggregated_data = filtered_df.groupby(['處理後檢驗員', '類別']).agg(
            total_standard_time=('處理後檢驗標準工時', 'sum'),
            total_actual_time=('檢驗耗時', 'sum'),
            record_count=('處理後檢驗標準工時', 'count')  # 使用任意列計數
        ).reset_index()
        
        # 重命名列以避免衝突
        aggregated_data.columns = ['inspector', 'category', 'total_standard_time', 'total_actual_time', 'record_count']
    except Exception:
        # 兼容舊版 pandas
        stat_data = filtered_df.groupby(['處理後檢驗員', '類別']).agg({
            '處理後檢驗標準工時': 'sum',
            '檢驗耗時': 'sum'
        }).reset_index()
        
        record_counts = filtered_df.groupby(['處理後檢驗員', '類別']).size().reset_index(name='record_count')
        
        aggregated_data = pd.merge(stat_data, record_counts, on=['處理後檢驗員', '類別'])
        
        # 重命名列
        aggregated_data.columns = ['inspector', 'category', 'total_standard_time', 'total_actual_time', 'record_count']
    
    # 向量化計算效率比值，替代 apply+lambda
    aggregated_data['efficiency'] = (
        aggregated_data['total_standard_time'] / 
        aggregated_data['total_actual_time'].clip(lower=0.1)
    ).clip(upper=MAX_EFFICIENCY)
    
    # 1. 計算總效率 - 按檢驗員分組
    overall_efficiency = aggregated_data.groupby('inspector').agg({
        'total_standard_time': 'sum',
        'total_actual_time': 'sum',
        'record_count': 'sum'
    }).reset_index()
    
    # 向量化計算總效率
    overall_efficiency['efficiency'] = (
        overall_efficiency['total_standard_time'] / 
        overall_efficiency['total_actual_time'].clip(lower=0.1)
    ).clip(upper=MAX_EFFICIENCY)
    
    # 按效率從高到低排序
    overall_efficiency = overall_efficiency.sort_values(by='efficiency', ascending=False)
    
    # 2. 計算類別效率數據
    category_efficiency_data = {}
    
    # 如果需要合併類別
    if merge_categories and selected_material_categories:
        merged_category = "+".join(selected_material_categories)
        
        # 為合併類別創建新的DataFrame
        merged_data = aggregated_data.copy()
        merged_data['category'] = merged_category
        
        # 按檢驗員分組
        merged_grouped = merged_data.groupby('inspector').agg({
            'total_standard_time': 'sum',
            'total_actual_time': 'sum',
            'record_count': 'sum'
        }).reset_index()
        
        # 向量化計算效率
        merged_grouped['efficiency'] = (
            merged_grouped['total_standard_time'] / 
            merged_grouped['total_actual_time'].clip(lower=0.1)
        ).clip(upper=MAX_EFFICIENCY)
        
        # 排序並保存
        merged_grouped = merged_grouped.sort_values(by='efficiency', ascending=False)
        category_efficiency_data[merged_category] = merged_grouped.to_dict('records')
    else:
        # 對每個類別，按效率排序並保存
        for category, group in aggregated_data.groupby('category'):
            sorted_group = group.sort_values(by='efficiency', ascending=False)
            category_efficiency_data[category] = sorted_group.to_dict('records')
    
    debug_log(f"效率指標計算完成，共 {len(overall_efficiency)} 位檢驗員, {len(category_efficiency_data)} 個類別", level="INFO")
    
    return {
        'overall_efficiency_ranking': overall_efficiency,
        'category_efficiency_data': category_efficiency_data
    }

@st.cache_data(ttl=1800, show_spinner=False)
def calculate_workload_metrics(processed_df, additional_tasks_df, start_date=None, end_date=None):
    """優化的工作負載指標計算函數 - 使用標準檢驗工時而非實際檢驗時間"""
    debug_log("開始計算工作負載指標")
    
    # 過濾資料，確保在日期範圍內
    filtered_iqc_df = filter_by_date_range(processed_df, start_date, end_date)
    
    # 確保additional_tasks_df不是None
    if additional_tasks_df is None or additional_tasks_df.empty:
        debug_log("額外任務資料為空，只計算檢驗工作負載")
        filtered_tasks_df = pd.DataFrame()
    else:
        filtered_tasks_df = filter_by_date_range(additional_tasks_df, start_date, end_date)
    
    # 一次性初始化結果DataFrame
    result_data = []
    
    # 處理IQC報告數據 - 向量化操作
    if filtered_iqc_df is not None and not filtered_iqc_df.empty:
        debug_log(f"處理 {len(filtered_iqc_df)} 筆IQC報告資料")
        
        # 確保日期欄位是datetime類型並一次性處理
        filtered_iqc_df['檢驗日期'] = pd.to_datetime(filtered_iqc_df['檢驗日期'], errors='coerce')
        
        # 使用格式化的日期作為索引
        filtered_iqc_df['formatted_date'] = filtered_iqc_df['檢驗日期'].dt.strftime('%Y-%m-%d')
        
        # 使用groupby進行聚合計算 - 關鍵修改: 使用處理後檢驗標準工時替代檢驗耗時
        iqc_workload = filtered_iqc_df.groupby(['formatted_date', '處理後檢驗員']).agg(
            inspection_standard_time=('處理後檢驗標準工時', 'sum'),  # 修改這裡: 使用標準工時
            inspection_count=('處理後檢驗員', 'count')
        ).reset_index()
        
        # 添加工作時段分析 - 只為有檢驗開始時間的數據
        if '檢驗開始時間' in filtered_iqc_df.columns:
            # 計算每組的最早和最晚時間
            time_analysis = filtered_iqc_df.dropna(subset=['檢驗開始時間']).groupby(
                ['formatted_date', '處理後檢驗員']
            ).agg(
                earliest_time=('檢驗開始時間', lambda x: x.dt.hour.min() + x.dt.minute.min()/60 if not x.empty else None),
                latest_time=('檢驗開始時間', lambda x: x.dt.hour.max() + x.dt.minute.max()/60 if not x.empty else None)
            ).reset_index()
            
            # 合併時間分析結果
            iqc_workload = pd.merge(
                iqc_workload, 
                time_analysis, 
                on=['formatted_date', '處理後檢驗員'], 
                how='left'
            )
            
            # 向量化計算工作時段
            iqc_workload['work_period'] = iqc_workload.apply(
                lambda row: f"{row['earliest_time']:.1f}时 - {row['latest_time']:.1f}时 (跨{row['latest_time']-row['earliest_time']:.1f}小时)" 
                if pd.notna(row['earliest_time']) and pd.notna(row['latest_time']) and row['latest_time'] - row['earliest_time'] > 0
                else (f"{row['earliest_time']:.1f}时" if pd.notna(row['earliest_time']) else "無法分析"),
                axis=1
            )
        else:
            iqc_workload['work_period'] = "無法分析"
        
        # 初始化額外任務時間為0
        iqc_workload['additional_task_time'] = 0
        
        # 將結果添加到結果清單
        for _, row in iqc_workload.iterrows():
            result_data.append({
                'date': row['formatted_date'],
                'inspector': row['處理後檢驗員'],
                'inspection_standard_time': row['inspection_standard_time'],  # 修改這裡: 使用標準工時
                'additional_task_time': 0,  # 將在後續與額外任務合併
                'inspection_count': row['inspection_count'],
                'work_period': row['work_period']
            })
    
    # 處理額外任務數據 - 向量化操作
    if filtered_tasks_df is not None and not filtered_tasks_df.empty:
        debug_log(f"處理 {len(filtered_tasks_df)} 筆額外任務資料")
        
        # 確保日期欄位是datetime類型並一次性處理
        filtered_tasks_df['日期'] = pd.to_datetime(filtered_tasks_df['日期'], errors='coerce')
        
        # 使用格式化的日期作為索引
        filtered_tasks_df['formatted_date'] = filtered_tasks_df['日期'].dt.strftime('%Y-%m-%d')
        
        # 使用groupby進行聚合計算
        tasks_workload = filtered_tasks_df.groupby(['formatted_date', '姓名']).agg(
            additional_task_time=('用時(分鐘)', 'sum')
        ).reset_index()
        
        # 添加任務數據到結果，首先創建索引
        date_inspector_map = {}
        for i, item in enumerate(result_data):
            key = f"{item['date']}-{item['inspector']}"
            date_inspector_map[key] = i
        
        # 將額外任務時間合併到結果中
        for _, row in tasks_workload.iterrows():
            key = f"{row['formatted_date']}-{row['姓名']}"
            if key in date_inspector_map:
                # 更新現有記錄
                idx = date_inspector_map[key]
                result_data[idx]['additional_task_time'] = row['additional_task_time']
            else:
                # 添加新記錄
                result_data.append({
                    'date': row['formatted_date'],
                    'inspector': row['姓名'],
                    'inspection_standard_time': 0,  # 修改這裡: 使用標準工時
                    'additional_task_time': row['additional_task_time'],
                    'inspection_count': 0,
                    'work_period': "無工作時段"
                })
    
    # 轉換為DataFrame並計算總時間和工作負載指數
    workload_df = pd.DataFrame(result_data)
    
    if not workload_df.empty:
        # 向量化計算總時間和工作負載指數 - 修改計算邏輯
        workload_df['total_time'] = workload_df['inspection_standard_time'] + workload_df['additional_task_time']
        workload_df['workload_index'] = workload_df['total_time'] / 480  # 8小時 = 480分鐘
        
        # 按日期排序
        workload_df['date'] = pd.to_datetime(workload_df['date'])
        workload_df = workload_df.sort_values('date')
        workload_df['date'] = workload_df['date'].dt.strftime('%Y-%m-%d')  # 轉回字符串格式，方便繪圖
    
    debug_log(f"工作負載計算完成，共 {len(workload_df)} 筆資料")
    return workload_df

# 同样修改时间分配比例计算和额外任务监控函数
def calculate_time_allocation_metrics(processed_df, additional_tasks_df, start_date=None, end_date=None):
    debug_log("開始計算時間分配比例")
    
    # 過濾資料，確保在日期範圍內
    filtered_iqc_df = filter_by_date_range(processed_df, start_date, end_date)
    
    # 确保additional_tasks_df不是None
    if additional_tasks_df is None or additional_tasks_df.empty:
        debug_log("額外任務資料為空，只計算檢驗時間佔比")
        filtered_tasks_df = pd.DataFrame()  # 使用空DataFrame
    else:
        filtered_tasks_df = filter_by_date_range(additional_tasks_df, start_date, end_date)
    
    # 按檢驗員分組計算時間分配
    time_allocation_by_inspector = {}
    
    # 处理IQC报告数据
    if filtered_iqc_df is not None and not filtered_iqc_df.empty:
        debug_log(f"处理 {len(filtered_iqc_df)} 筆IQC报告资料")
        
        for _, row in filtered_iqc_df.iterrows():
            inspector = row['處理後檢驗員'] if pd.notna(row['處理後檢驗員']) else 'Unknown'
            
            if inspector not in time_allocation_by_inspector:
                time_allocation_by_inspector[inspector] = {
                    'inspector': inspector,
                    'inspection_time': 0,
                    'additional_task_time': 0,
                    'additional_task_details': {}
                }
            
            # 处理检验时间
            try:
                inspection_time = float(row['檢驗耗時']) if pd.notna(row['檢驗耗時']) else 0
            except (ValueError, TypeError):
                debug_log(f"无法转换检验耗时 {row['檢驗耗時']} 为浮点数，设为0")
                inspection_time = 0
            
            time_allocation_by_inspector[inspector]['inspection_time'] += inspection_time
    
    # 处理额外任务数据
    if filtered_tasks_df is not None and not filtered_tasks_df.empty:
        debug_log(f"处理 {len(filtered_tasks_df)} 筆额外任务资料")
        
        for _, row in filtered_tasks_df.iterrows():
            inspector = row['姓名'] if pd.notna(row['姓名']) else 'Unknown'
            task_type = row['工作事項分類'] if pd.notna(row['工作事項分類']) else 'Other'
            
            # 处理任务时间
            try:
                task_time = int(row['用時(分鐘)']) if pd.notna(row['用時(分鐘)']) else 0
            except (ValueError, TypeError):
                debug_log(f"无法转换用时 {row['用時(分鐘)']} 为整数，设为0")
                task_time = 0
            
            if inspector not in time_allocation_by_inspector:
                time_allocation_by_inspector[inspector] = {
                    'inspector': inspector,
                    'inspection_time': 0,
                    'additional_task_time': 0,
                    'additional_task_details': {}
                }
            
            time_allocation_by_inspector[inspector]['additional_task_time'] += task_time
            
            if task_type not in time_allocation_by_inspector[inspector]['additional_task_details']:
                time_allocation_by_inspector[inspector]['additional_task_details'][task_type] = 0
            
            time_allocation_by_inspector[inspector]['additional_task_details'][task_type] += task_time
    
    # 计算比例
    time_allocation_data = []
    
    for inspector, data in time_allocation_by_inspector.items():
        total_time = data['inspection_time'] + data['additional_task_time']
        inspection_ratio = data['inspection_time'] / total_time if total_time > 0 else 1.0  # 如果没有额外任务，检验时间占比为100%
        additional_task_ratio = data['additional_task_time'] / total_time if total_time > 0 else 0.0
        
        # 计算细项任务的比例
        task_detail_ratios = {}
        for task_type, time in data['additional_task_details'].items():
            task_detail_ratios[task_type] = time / total_time if total_time > 0 else 0
        
        time_allocation_data.append({
            'inspector': inspector,
            'inspection_time': data['inspection_time'],
            'additional_task_time': data['additional_task_time'],
            'total_time': total_time,
            'inspection_ratio': inspection_ratio,
            'additional_task_ratio': additional_task_ratio,
            'task_detail_ratios': task_detail_ratios
        })
    
    debug_log(f"时间分配比例计算完成，共 {len(time_allocation_data)} 位检验员")
    return pd.DataFrame(time_allocation_data)

def calculate_additional_tasks_metrics(additional_tasks_df, start_date=None, end_date=None):
    debug_log("开始计算额外任务监控数据")
    
    # 如果额外任务数据为空，返回空DataFrame
    if additional_tasks_df is None or additional_tasks_df.empty:
        debug_log("额外任务资料为空，返回空结果")
        return pd.DataFrame(columns=['inspector', 'task_type', 'total_time'])
    
    # 過濾資料，確保在日期範圍內
    filtered_tasks_df = filter_by_date_range(additional_tasks_df, start_date, end_date)
    
    # 如果沒有資料，返回空的DataFrame
    if filtered_tasks_df.empty:
        debug_log("过滤后没有额外任务资料")
        return pd.DataFrame(columns=['inspector', 'task_type', 'total_time'])
    
    debug_log(f"处理 {len(filtered_tasks_df)} 筆额外任务资料")
    
    # 按檢驗員和任務類型分組計算時間
    task_summary = {}
    
    for _, row in filtered_tasks_df.iterrows():
        inspector = row['姓名'] if pd.notna(row['姓名']) else 'Unknown'
        task_type = row['工作事項分類'] if pd.notna(row['工作事項分類']) else 'Other'
        
        # 处理任务时间
        try:
            task_time = int(row['用時(分鐘)']) if pd.notna(row['用時(分鐘)']) else 0
        except (ValueError, TypeError):
            debug_log(f"无法转换用时 {row['用時(分鐘)']} 为整数，设为0")
            task_time = 0
        
        if inspector not in task_summary:
            task_summary[inspector] = {}
        
        if task_type not in task_summary[inspector]:
            task_summary[inspector][task_type] = 0
        
        task_summary[inspector][task_type] += task_time
    
    # 轉換為DataFrame格式
    task_monitor_data = []
    
    for inspector, tasks in task_summary.items():
        for task_type, total_time in tasks.items():
            task_monitor_data.append({
                'inspector': inspector,
                'task_type': task_type,
                'total_time': total_time
            })
    
    debug_log(f"额外任务监控数据计算完成，共 {len(task_monitor_data)} 筆资料")
    return pd.DataFrame(task_monitor_data)

@st.cache_data(ttl=86400, max_entries=5, show_spinner=False)
def calculate_all_metrics(iqc_df, pcb_specs_df, pcb_standard_time_df, additional_tasks_df, start_date=None, end_date=None):
    try:
        debug_log("開始計算所有指標", level="INFO")
        start_time = time.time()  # 記錄開始時間
        
        # 檢查資料是否為空
        if iqc_df is None or iqc_df.empty:
            debug_log("IQC報告資料為空，無法計算指標", level="ERROR")
            st.error("IQC報告資料為空，請上傳有效的資料檔案。")
            return None
        
        # 使用優化後的PCB標準工時計算
        iqc_data_with_pcb_time = calculate_pcb_standard_time(iqc_df, pcb_specs_df, pcb_standard_time_df)
        debug_log(f"PCB標準工時計算完成，資料筆數: {len(iqc_data_with_pcb_time)}", level="INFO")
        pcb_time_end = time.time()
        debug_log(f"PCB標準工時計算用時: {pcb_time_end - start_time:.2f}秒", level="INFO")
        
        # 重要修改: 先進行日期篩選，確保後續所有計算都使用篩選後的數據
        filtered_data = filter_by_date_range(iqc_data_with_pcb_time, start_date, end_date)
        debug_log(f"日期篩選完成，從 {len(iqc_data_with_pcb_time)} 筆資料篩選出 {len(filtered_data)} 筆", level="INFO")
        
        # 使用日期篩選後的資料進行效率計算
        efficiency_start = time.time()
        efficiency_data = calculate_efficiency_metrics(
            filtered_data,  # 使用篩選後的資料 
            start_date, 
            end_date
        )
        efficiency_end = time.time()
        debug_log(f"效率指標計算用時: {efficiency_end - efficiency_start:.2f}秒", level="INFO")
        
        # 計算工作負載指標
        workload_start = time.time()
        if additional_tasks_df is None:
            debug_log("IQC額外任務紀錄清單数据不存在，將使用空數據計算相關指標", level="WARNING")
            additional_tasks_df = pd.DataFrame()
        
        # 使用篩選後的資料計算工作負載
        workload_data = calculate_workload_metrics(filtered_data, additional_tasks_df, start_date, end_date)
        workload_end = time.time()
        debug_log(f"工作負載指標計算用時: {workload_end - workload_start:.2f}秒", level="INFO")
        
        # 計算時間分配和額外任務指標
        time_allocation_start = time.time()
        time_allocation_data = calculate_time_allocation_metrics(filtered_data, additional_tasks_df, start_date, end_date)
        additional_tasks_monitor_data = calculate_additional_tasks_metrics(additional_tasks_df, start_date, end_date)
        time_allocation_end = time.time()
        debug_log(f"時間分配與額外任務指標計算用時: {time_allocation_end - time_allocation_start:.2f}秒", level="INFO")
        
        # 總計算時間
        total_time = time.time() - start_time
        debug_log(f"所有指標計算完成，總用時: {total_time:.2f}秒", level="INFO")
        
        # 存儲日期篩選信息到session_state，供後續使用
        st.session_state.filtered_start_date = start_date
        st.session_state.filtered_end_date = end_date
        
        # 創建返回的指標結果字典
        metrics = {
            'processed_data': filtered_data,  # 這裡改為使用日期篩選後的資料
            'efficiency_data': efficiency_data,
            'workload_data': workload_data,
            'time_allocation_data': time_allocation_data,
            'additional_tasks_monitor_data': additional_tasks_monitor_data
        }
        
        # 在返回結果前過濾指定檢驗員
        if metrics:
            # 過濾processed_data
            if 'processed_data' in metrics and metrics['processed_data'] is not None:
                metrics['processed_data'] = filter_excluded_inspectors(metrics['processed_data'])
                
            # 過濾efficiency_data中的overall_efficiency_ranking
            if 'efficiency_data' in metrics and 'overall_efficiency_ranking' in metrics['efficiency_data']:
                metrics['efficiency_data']['overall_efficiency_ranking'] = filter_excluded_inspectors(
                    metrics['efficiency_data']['overall_efficiency_ranking'], 
                    inspector_column='inspector'
                )
                
            # 過濾efficiency_data中的category_efficiency_data
            if 'efficiency_data' in metrics and 'category_efficiency_data' in metrics['efficiency_data']:
                category_data = metrics['efficiency_data']['category_efficiency_data']
                for category in category_data:
                    if isinstance(category_data[category], list):
                        # 如果是列表格式，過濾每個項目
                        category_data[category] = [
                            item for item in category_data[category] 
                            if item.get('inspector') not in EXCLUDED_INSPECTORS
                        ]
                
            # 過濾workload_data
            if 'workload_data' in metrics and metrics['workload_data'] is not None:
                metrics['workload_data'] = filter_excluded_inspectors(metrics['workload_data'], inspector_column='inspector')
                
            # 過濾time_allocation_data
            if 'time_allocation_data' in metrics and metrics['time_allocation_data'] is not None:
                metrics['time_allocation_data'] = filter_excluded_inspectors(metrics['time_allocation_data'], inspector_column='inspector')
                
            # 過濾additional_tasks_monitor_data
            if 'additional_tasks_monitor_data' in metrics and metrics['additional_tasks_monitor_data'] is not None:
                metrics['additional_tasks_monitor_data'] = filter_excluded_inspectors(metrics['additional_tasks_monitor_data'], inspector_column='inspector')
        
        return metrics
    
    except Exception as e:
        error_msg = f"計算指標時出錯: {str(e)}\n{traceback.format_exc()}"
        debug_log(error_msg, level="ERROR")
        st.error(error_msg)
        st.session_state.processing_error = error_msg
        return None

# 計算總效率指標的函數 (不受物料類別篩選影響)
def calculate_overall_efficiency(filtered_df):
    if filtered_df is None or filtered_df.empty:
        debug_log("沒有數據，返回空DataFrame")
        return pd.DataFrame()
    
    # 準備存儲效率數據的字典
    inspector_efficiency = {}
    
    # 設定效率值的合理上限
    MAX_EFFICIENCY = 20
    
    # 計算每個IQC人員的效率
    for _, row in filtered_df.iterrows():
        inspector = row['處理後檢驗員'] if pd.notna(row['處理後檢驗員']) else 'Unknown'
        category = row['類別'] if pd.notna(row['類別']) else 'Unknown'
        
        # 處理標準工時
        try:
            standard_time = float(row['處理後檢驗標準工時']) if pd.notna(row['處理後檢驗標準工時']) else 0
        except (ValueError, TypeError):
            debug_log(f"無法轉換標準工時 {row['處理後檢驗標準工時']} 為浮點數，設為0")
            standard_time = 0
        
        # 處理實際檢驗時間
        try:
            actual_time = float(row['檢驗耗時']) if pd.notna(row['檢驗耗時']) else 0
        except (ValueError, TypeError):
            debug_log(f"無法轉換檢驗耗時 {row['檢驗耗時']} 為浮點數，設為0")
            actual_time = 0
        
        # 檢查實際耗時是否過小，避免除以零或極小值
        if actual_time < 0.1:
            debug_log(f"發現極小實際耗時: {actual_time}, 檢驗員: {inspector}, 料號: {row['料號'] if '料號' in row else 'Unknown'}")
            actual_time = 0.1  # 設定最小值避免除以零
        
        # 計算效率比值並限制上限
        efficiency_ratio = min(standard_time / actual_time, MAX_EFFICIENCY) if actual_time > 0 else 0
        
        # 總效率計算
        if inspector not in inspector_efficiency:
            inspector_efficiency[inspector] = {
                'total_standard_time': 0,
                'total_actual_time': 0,
                'records': []
            }
        
        inspector_efficiency[inspector]['total_standard_time'] += standard_time
        inspector_efficiency[inspector]['total_actual_time'] += actual_time
        inspector_efficiency[inspector]['records'].append({
            'efficiency_ratio': efficiency_ratio,
            'standard_time': standard_time,
            'actual_time': actual_time,
            'category': category
        })
    
    # 計算總效率並排名
    overall_efficiency_data = []
    
    for inspector, data in inspector_efficiency.items():
        # 避免除以零，並限制最大效率值
        if data['total_actual_time'] > 0:
            efficiency = min(data['total_standard_time'] / data['total_actual_time'], MAX_EFFICIENCY)
        else:
            efficiency = 0
            debug_log(f"警告: 檢驗員 {inspector} 的總實際耗時為零")
        
        # 計算各類別的筆數
        category_counts = {}
        for record in data['records']:
            category = record['category']
            if category not in category_counts:
                category_counts[category] = 0
            category_counts[category] += 1
        
        overall_efficiency_data.append({
            'inspector': inspector,
            'efficiency': efficiency,
            'total_standard_time': data['total_standard_time'],
            'total_actual_time': data['total_actual_time'],
            'record_count': len(data['records']),
            'category_counts': category_counts
        })
    
    # 按效率從高到低排序
    overall_efficiency_ranking = pd.DataFrame(overall_efficiency_data).sort_values(by='efficiency', ascending=False)
    debug_log(f"總效率排名計算完成，共 {len(overall_efficiency_ranking)} 位檢驗員")
    
    return overall_efficiency_ranking

# 計算初始的物料類別效率（用於初始加載，不篩選）
def calculate_category_efficiency_initial(filtered_df, categories=None):
    if filtered_df is None or filtered_df.empty:
        return {}
    
    # 準備存儲效率數據的字典
    category_efficiency_data = {}
    
    # 設定效率值的合理上限
    MAX_EFFICIENCY = 20
    
    # 計算每個IQC人員在各物料類別的效率
    for _, row in filtered_df.iterrows():
        inspector = row['處理後檢驗員'] if pd.notna(row['處理後檢驗員']) else 'Unknown'
        category = row['類別'] if pd.notna(row['類別']) else 'Unknown'
        
        # 如果指定了類別且當前類別不在指定類別中，則跳過
        if categories and category not in categories:
            continue
        
        # 處理標準工時和實際耗時
        try:
            standard_time = float(row['處理後檢驗標準工時']) if pd.notna(row['處理後檢驗標準工時']) else 0
        except (ValueError, TypeError):
            standard_time = 0
        
        try:
            actual_time = float(row['檢驗耗時']) if pd.notna(row['檢驗耗時']) else 0
        except (ValueError, TypeError):
            actual_time = 0
        
        # 檢查實際耗時是否過小
        if actual_time < 0.1:
            actual_time = 0.1
        
        # 物料類別效率計算
        key = f"{category}-{inspector}"
        if key not in category_efficiency_data:
            category_efficiency_data[key] = {
                'category': category,
                'inspector': inspector,
                'total_standard_time': 0,
                'total_actual_time': 0,
                'record_count': 0
            }
        
        category_efficiency_data[key]['total_standard_time'] += standard_time
        category_efficiency_data[key]['total_actual_time'] += actual_time
        category_efficiency_data[key]['record_count'] += 1
    
    # 按類別和檢驗員分組處理
    result = {}
    
    for key, data in category_efficiency_data.items():
        category = data['category']
        inspector = data['inspector']
        
        # 計算效率
        if data['total_actual_time'] > 0:
            category_efficiency = min(data['total_standard_time'] / data['total_actual_time'], MAX_EFFICIENCY)
        else:
            category_efficiency = 0
        
        if category not in result:
            result[category] = []
        
        result[category].append({
            'inspector': inspector,
            'efficiency': category_efficiency,
            'record_count': data['record_count'],
            'total_standard_time': data['total_standard_time'],
            'total_actual_time': data['total_actual_time']
        })
    
    # 為每個類別排序
    for category in result:
        result[category] = sorted(
            result[category],
            key=lambda x: x['efficiency'],
            reverse=True
        )
    
    return result

def calculate_mrb_rates(processed_df, start_date=None, end_date=None):
    """
    計算每個IQC人員的MRB率 - 安全的方式
    """
    # 篩選日期範圍
    filtered_df = filter_by_date_range(processed_df, start_date, end_date)
    
    if filtered_df is None or filtered_df.empty:
        return pd.DataFrame(columns=['檢驗員', 'MRB數量', '總檢驗筆數', 'MRB率'])
    
    # 初始化MRB掩碼
    mrb_mask = pd.Series(False, index=filtered_df.index)
    
    # 安全檢查所有可能的MRB欄位
    mrb_columns = ['M', '異常問題匯總', 'MRB', 'MRB狀態', 'MRB Status', '是否為MRB']
    
    for col in mrb_columns:
        if col in filtered_df.columns:
            if col == '是否為MRB':
                # 直接使用布爾值
                mrb_mask = mrb_mask | filtered_df[col]
            else:
                # 檢查非空值
                col_mask = filtered_df[col].notna() & (filtered_df[col].astype(str).str.strip() != '')
                mrb_mask = mrb_mask | col_mask
    
    # 使用標記
    filtered_df['MRB標記'] = mrb_mask
    
    # 按檢驗員分組計算MRB率
    mrb_stats = []
    
    for inspector, group in filtered_df.groupby('處理後檢驗員'):
        total_count = len(group)
        mrb_count = group['MRB標記'].sum()
        mrb_rate = mrb_count / total_count if total_count > 0 else 0
        
        mrb_stats.append({
            '檢驗員': inspector,
            'MRB數量': mrb_count,
            '總檢驗筆數': total_count,
            'MRB率': mrb_rate
        })
    
    # 轉換為DataFrame並按MRB率排序
    mrb_rate_df = pd.DataFrame(mrb_stats)
    mrb_rate_df = mrb_rate_df.sort_values('MRB率', ascending=False)
    
    return mrb_rate_df

def render_mrb_rate_chart(efficiency_data, processed_data=None):
    """
    使用Altair顯示IQC人員MRB率圖表，確保懸浮提示正確顯示
    """
    if processed_data is None or processed_data.empty:
        st.error("沒有可用的數據來計算MRB率")
        return
    
    st.header("IQC人員MRB率")
    st.markdown("此圖表顯示各IQC人員在選定時間區段內開立MRB的比例")
    
    # 計算MRB率
    mrb_rate_df = calculate_mrb_rates(processed_data)
    
    if mrb_rate_df.empty:
        st.info("沒有可用的MRB率數據")
        return
    
    # 使用Altair渲染圖表
    chart = render_inspector_mrb_rate_chart(mrb_rate_df)
    
    # 顯示圖表
    st.altair_chart(chart, use_container_width=True)
    
    # 顯示詳細數據
    with st.expander("查看MRB率詳細數據"):
        detail_df = mrb_rate_df.copy()
        detail_df['MRB率'] = detail_df['MRB率'].apply(lambda x: f"{x:.2%}")  # 格式化為百分比
        st.dataframe(detail_df, use_container_width=True)

def render_efficiency_dashboard(efficiency_data, processed_data=None):
    if efficiency_data is None:
        st.error("沒有可用的效率數據，請確保上傳了正確的檔案格式。")
        return
    
    # 過濾處理後的數據 - 排除指定檢驗員
    if processed_data is not None and not processed_data.empty:
        processed_data = filter_excluded_inspectors(processed_data)
    
    # 過濾整體效率排名數據 - 排除指定檢驗員
    if 'overall_efficiency_ranking' in efficiency_data:
        overall_efficiency_ranking = efficiency_data['overall_efficiency_ranking']
        if 'inspector' in overall_efficiency_ranking.columns:
            overall_efficiency_ranking = overall_efficiency_ranking[~overall_efficiency_ranking['inspector'].isin(EXCLUDED_INSPECTORS)].copy()
            # 更新回 efficiency_data 字典
            temp_efficiency_data = efficiency_data.copy()
            temp_efficiency_data['overall_efficiency_ranking'] = overall_efficiency_ranking
            efficiency_data = temp_efficiency_data
    
    # 過濾類別效率數據 - 排除指定檢驗員
    if 'category_efficiency_data' in efficiency_data:
        category_data = efficiency_data['category_efficiency_data']
        filtered_category_data = {}
        
        for category, data in category_data.items():
            if isinstance(data, list):
                # 過濾列表格式的數據
                filtered_data = [item for item in data if 'inspector' in item and item['inspector'] not in EXCLUDED_INSPECTORS]
                filtered_category_data[category] = filtered_data
            elif isinstance(data, pd.DataFrame):
                # 過濾DataFrame格式的數據
                filtered_data = data[~data['inspector'].isin(EXCLUDED_INSPECTORS)].copy()
                filtered_category_data[category] = filtered_data
            else:
                # 保留其他格式的數據
                filtered_category_data[category] = data
        
        # 更新回 efficiency_data 字典
        temp_efficiency_data = efficiency_data.copy()
        temp_efficiency_data['category_efficiency_data'] = filtered_category_data
        efficiency_data = temp_efficiency_data
    
    st.header("📈IQC檢驗效率監控")
    
    st.markdown("""
    <h3 style="font-size: 1.3rem; font-weight: 600; margin-bottom: 1rem; display: flex; align-items: center;">
        排除極值設定
        <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512" width="22" height="22" style="margin-right: 8px;">
            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512"><!--!Font Awesome Free 6.7.2 by @fontawesome - https://fontawesome.com License - https://fontawesome.com/license/free Copyright 2025 Fonticons, Inc.--><path fill="#9baac5" d="M495.9 166.6c3.2 8.7 .5 18.4-6.4 24.6l-43.3 39.4c1.1 8.3 1.7 16.8 1.7 25.4s-.6 17.1-1.7 25.4l43.3 39.4c6.9 6.2 9.6 15.9 6.4 24.6c-4.4 11.9-9.7 23.3-15.8 34.3l-4.7 8.1c-6.6 11-14 21.4-22.1 31.2c-5.9 7.2-15.7 9.6-24.5 6.8l-55.7-17.7c-13.4 10.3-28.2 18.9-44 25.4l-12.5 57.1c-2 9.1-9 16.3-18.2 17.8c-13.8 2.3-28 3.5-42.5 3.5s-28.7-1.2-42.5-3.5c-9.2-1.5-16.2-8.7-18.2-17.8l-12.5-57.1c-15.8-6.5-30.6-15.1-44-25.4L83.1 425.9c-8.8 2.8-18.6 .3-24.5-6.8c-8.1-9.8-15.5-20.2-22.1-31.2l-4.7-8.1c-6.1-11-11.4-22.4-15.8-34.3c-3.2-8.7-.5-18.4 6.4-24.6l43.3-39.4C64.6 273.1 64 264.6 64 256s.6-17.1 1.7-25.4L22.4 191.2c-6.9-6.2-9.6-15.9-6.4-24.6c4.4-11.9 9.7-23.3 15.8-34.3l4.7-8.1c6.6-11 14-21.4 22.1-31.2c5.9-7.2 15.7-9.6 24.5-6.8l55.7 17.7c13.4-10.3 28.2-18.9 44-25.4l12.5-57.1c2-9.1 9-16.3 18.2-17.8C227.3 1.2 241.5 0 256 0s28.7 1.2 42.5 3.5c9.2 1.5 16.2 8.7 18.2 17.8l12.5 57.1c15.8 6.5 30.6 15.1 44 25.4l55.7-17.7c8.8-2.8 18.6-.3 24.5 6.8c8.1 9.8 15.5 20.2 22.1 31.2l4.7 8.1c6.1 11 11.4 22.4 15.8 34.3zM256 336a80 80 0 1 0 0-160 80 80 0 1 0 0 160z"/></svg>
        </svg>
    </h3>
    """, unsafe_allow_html=True)
    
    # 使用兩列佈局使控制項更緊湊
    col1, col2 = st.columns([3, 1])
    
    with col1:
        trim_percentage = st.slider(
            label="極值百分比",
            min_value=0.0,
            max_value=25.0,
            value=st.session_state.get('trim_percentage', 0.0),  # 使用儲存的值
            step=1.0,
            help="設定要剔除的極值百分比，例如設定10%會從每個檢驗員的數據中剔除效率最高的5%和最低的5%"
        )
    
    with col2:
        apply_trim = st.button("應用", use_container_width=True)
    
    # 當用戶點擊"應用"按鈕時，重新計算效率數據
    if apply_trim or ('trim_percentage' in st.session_state and trim_percentage != st.session_state.trim_percentage):
        # 更新設定值
        old_trim = st.session_state.get('trim_percentage', 0.0)
        st.session_state.trim_percentage = trim_percentage
        
        debug_log(f"極值設定從 {old_trim}% 變更為 {trim_percentage}%", level="INFO")
        
        # 使用保存的設定值
        current_trim = st.session_state.trim_percentage
        
        if current_trim > 0:
            # 如果有進行極值剔除，顯示提示訊息
            st.info(f"已從每位檢驗員的資料中剔除效率最高的 {current_trim/2:.1f}% 和最低的 {current_trim/2:.1f}%")
        else:
            # 如果設為0，清除之前剔除效果
            if 'using_trimmed_efficiency' in st.session_state:
                st.session_state.using_trimmed_efficiency = False
            if 'trimmed_avg_efficiency' in st.session_state:
                del st.session_state.trimmed_avg_efficiency
            if 'trimmed_record_stats' in st.session_state:
                del st.session_state.trimmed_record_stats
        
        # 重新計算效率數據 - 確保使用正確的日期篩選資料
        if processed_data is not None and not processed_data.empty:
            with st.spinner("重新計算效率中..."):
                # 檢查資料是否已經過日期篩選
                if 'start_date' in st.session_state or 'end_date' in st.session_state:
                    start_date = st.session_state.get('start_date')
                    end_date = st.session_state.get('end_date')
                    debug_log(f"重新計算效率時使用日期篩選: {start_date} 到 {end_date}", level="INFO")
                
                # 確認使用的數據是否正確
                debug_log(f"使用資料行數: {len(processed_data)}", level="INFO")
                
                # 使用極值剔除方法重新計算效率
                overall_efficiency_ranking = calculate_efficiency_with_trimming(
                    processed_data,  # 這裡是關鍵，確保使用的是日期篩選後的數據
                    current_trim
                )
                
                # 將新計算的效率排名保存回session_state
                temp_efficiency_data = st.session_state.efficiency_data.copy()
                temp_efficiency_data['overall_efficiency_ranking'] = overall_efficiency_ranking
                st.session_state.efficiency_data = temp_efficiency_data
                
                # 使用 st.rerun() 重新加載頁面以顯示新的關鍵指標
                st.rerun()
        else:
            # 如果沒有原始資料，使用現有效率數據
            overall_efficiency_ranking = efficiency_data['overall_efficiency_ranking']
    else:
        # 使用原始效率數據或上次計算的效率數據
        overall_efficiency_ranking = efficiency_data['overall_efficiency_ranking']
    
    # 显示IQC人员总检验效率排名
    if overall_efficiency_ranking.empty:
        st.error("沒有可用的總效率數據，請確保上傳了正確的檔案格式或檢查過濾條件。")
        return
    
    # 設置標題 - 帶有極值剔除信息
    title_suffix = ""
    if 'trim_percentage' in st.session_state and st.session_state.trim_percentage > 0:
        title_suffix = f" (剔除極值 {st.session_state.trim_percentage}%)"
        
    st.header(f"IQC人員總檢驗效率排名🏆{title_suffix}")
    
    # 設置最大效率顯示值，避免異常值影響圖表比例
    max_display_efficiency = min(5, overall_efficiency_ranking['efficiency'].max() * 1.1)
    
    # 手動設置顏色並添加標籤
    colors = []
    for eff in overall_efficiency_ranking['efficiency']:
        if eff < 1:
            # 效率小於1使用紅色系
            red_intensity = max(100, 255 - int((1 - eff) * 155))
            colors.append(f'rgba({red_intensity}, 0, 0, 0.8)')
        else:
            # 效率大於等於1使用藍色系
            blue_intensity = max(100, 255 - int((eff - 1) * 85))
            colors.append(f'rgba(0, 0, {blue_intensity}, 0.8)')
    
    # 創建基本圖表
    fig = go.Figure()
    
    # 添加條形圖
    fig.add_trace(go.Bar(
        x=overall_efficiency_ranking['inspector'],
        y=overall_efficiency_ranking['efficiency'],
        marker_color=colors
    ))
    
    # 添加參考線
    fig.add_shape(
        type="line",
        x0=-0.5,
        y0=1,
        x1=len(overall_efficiency_ranking)-0.5,
        y1=1,
        line=dict(
                color="#90A4AE",  # 淺灰藍色 - 與主題一致
                width=1.5,        # 適中線寬
                dash=None         # 實線而非虛線
            )
    )
    
    # 添加顏色標籤 - 使用隱藏的散點圖
    fig.add_trace(
        go.Scatter(
            x=[None], y=[None],
            mode='markers',
            marker=dict(
                colorscale=[
                    [0, "rgb(180, 0, 0)"],    # 深紅色 (效率為0)
                    [0.25, "rgb(255, 0, 0)"],  # 紅色 (效率為0.5)
                    [0.5, "rgb(255, 150, 150)"], # 淺紅色 (效率為1)
                    [0.5, "rgb(150, 150, 255)"], # 淺藍色 (效率為1)
                    [0.75, "rgb(0, 0, 255)"],   # 藍色 (效率為1.5)
                    [1, "rgb(0, 0, 150)"]      # 深藍色 (效率為2+)
                ],
                showscale=True,
                cmin=0,
                cmax=2,
                colorbar=dict(
                    title="檢驗效率",
                    tickvals=[0, 0.5, 1.0, 1.5, 2.0],
                    ticktext=["0", "0.5", "1.0", "1.5", "2.0"],
                    len=0.5,
                    thickness=15,
                    y=0.5,
                    yanchor="middle",
                    x=1.02,
                    xanchor="left",
                    outlinewidth=0,  # 移除框線
                    borderwidth=0    # 移除邊框
                )
            ),
            hoverinfo='none',
            showlegend=False
        )
    )
    
    # 設置標題和軸標籤
    fig.update_layout(
        title=f"IQC人員總檢驗效率排名",
        xaxis_title="檢驗員",
        yaxis_title="檢驗效率",
        height=500
    )
    
    st.plotly_chart(fig, use_container_width=True, key="overall_efficiency_chart")
    
    # 顯示效率數據明細
    with st.expander("檢視效率數據明細"):
        detail_df = overall_efficiency_ranking[['inspector', 'efficiency', 'total_standard_time', 'total_actual_time', 'record_count']].copy()
        
        # 如果有進行極值剔除，添加原始記錄數量和剔除數量
        if 'original_record_count' in overall_efficiency_ranking.columns and 'trimmed_count' in overall_efficiency_ranking.columns:
            detail_df['original_record_count'] = overall_efficiency_ranking['original_record_count']
            detail_df['trimmed_count'] = overall_efficiency_ranking['trimmed_count']
            detail_df.columns = ['檢驗員', '效率', '標準工時總和(分鐘)', '實際耗時總和(分鐘)', '使用記錄數', '原始記錄數', '剔除記錄數']
        else:
            detail_df.columns = ['檢驗員', '效率', '標準工時總和(分鐘)', '實際耗時總和(分鐘)', '紀錄筆數']
        
        detail_df['效率'] = detail_df['效率'].round(2)
        detail_df['標準工時總和(分鐘)'] = detail_df['標準工時總和(分鐘)'].round(0)
        detail_df['實際耗時總和(分鐘)'] = detail_df['實際耗時總和(分鐘)'].round(0)
        
        # 根据效率值添加背景色
        def highlight_efficiency(val):
            if val < 1:
                # 效率<1，使用红色渐变
                red_intensity = max(70, 255 - int((1 - val) * 180))
                return f'background-color: rgba({red_intensity}, 0, 0, 0.3)'
            else:
                # 效率>=1，使用蓝色渐变
                blue_intensity = max(70, 255 - int((val - 1) * 50))
                return f'background-color: rgba(0, 0, {blue_intensity}, 0.3)'
        
        # 应用样式
        styled_df = detail_df.style.applymap(
            highlight_efficiency, 
            subset=['效率']
        )
        
        st.dataframe(styled_df, use_container_width=True)
    
    # 物料類別效率排名
    st.header("物料類別效率排名🗃️ ")
    
    # 創建一個更緊湊的3列布局，調整比例使第一列和第二列靠近
    col1, col2, col3 = st.columns([1.5, 3, 1.5])
    
    # 第一列放置分析模式選擇
    with col1:
        st.write("選擇分析模式")
        selection_mode = st.radio(
            label="",  # 空標題
            options=["大類別", "子類別"],
            index=0,
            key="category_mode_selector",
            horizontal=True,
            label_visibility="collapsed"  # 隱藏標籤
        )
    
    # 第二列放置物料類別選擇
    with col2:
        if selection_mode == "大類別":
            # 類別模式：選擇大類別
            st.write("選擇要顯示的物料類別")
            selected_categories = st.selectbox(
                label="",  # 空標題
                options=list(MATERIAL_CATEGORY_MAPPING.keys()),
                key="material_main_category_selector",
                label_visibility="collapsed"  # 隱藏標籤
            )
            
            # 如果選擇了類別，獲取該類別下的所有子類別
            if selected_categories:
                selected_subcategories = MATERIAL_CATEGORY_MAPPING.get(selected_categories, [])
            else:
                selected_subcategories = []
            
            # 類別模式強制合併顯示
            merge_categories = True
        else:
            # 子類別模式
            available_subcategories = []
            if processed_data is not None and not processed_data.empty:
                available_subcategories = sorted(processed_data['類別'].dropna().unique().tolist())
            
            st.write("選擇要顯示的物料子類別")
            selected_subcategories = st.multiselect(
                label="",  # 空標題
                options=available_subcategories,
                default=[],
                key="material_subcategory_selector",
                label_visibility="collapsed"  # 隱藏標籤
            )
            selected_categories = None  # 不使用大類別
            
            # 重要修改：根據選擇的子類別數量決定是否合併顯示
            # 如果選擇了多個子類別，則自動合併顯示
            merge_categories = len(selected_subcategories) > 1
    
    # 第三列放置應用選擇按鈕，確保與下拉選單同高
    with col3:
        st.write("　")  # 添加空白行，使按鈕與下拉選單對齊
        # 應用選擇按鈕
        apply_button = st.button(
            label="應用選擇", 
            key="apply_category_selection",
            use_container_width=True
        )
        
        if apply_button:
            if selection_mode == "大類別":
                if not selected_subcategories:
                    st.warning("所選類別沒有可用的子類別")
                else:
                    with st.spinner("計算中..."):
                        # 使用大類別下的所有子類別計算
                        new_category_data = calculate_category_efficiency(
                            processed_data,
                            selected_subcategories,
                            True  # 類別模式強制合併顯示
                        )
                        
                        if new_category_data:
                            # 只更新類別效率數據
                            st.session_state.category_efficiency_data = new_category_data
                            st.session_state.selected_category_mode = "category"
                            st.session_state.selected_main_category = selected_categories
                            st.session_state.selected_subcategories = selected_subcategories
                            st.session_state.merge_categories = True  # 確保合併顯示
                            st.session_state.has_applied_selection = True
                            st.rerun()
            else:
                if not selected_subcategories:
                    st.warning("請至少選擇一個物料子類別")
                else:
                    with st.spinner("計算中..."):
                        # 使用選擇的子類別計算，根據子類別數量決定是否合併顯示
                        new_category_data = calculate_category_efficiency(
                            processed_data,
                            selected_subcategories,
                            merge_categories  # 多於一個子類別時自動合併
                        )
                        
                        if new_category_data:
                            # 只更新類別效率數據
                            st.session_state.category_efficiency_data = new_category_data
                            st.session_state.selected_category_mode = "subcategory"
                            st.session_state.selected_subcategories = selected_subcategories
                            st.session_state.merge_categories = merge_categories  # 存儲合併狀態
                            st.session_state.has_applied_selection = True
                            st.rerun()
    
    # 顯示物料類別效率排名結果
    category_efficiency_data = efficiency_data['category_efficiency_data']
    if 'category_efficiency_data' in st.session_state:
        category_efficiency_data = st.session_state.category_efficiency_data
    
    # 只有在用戶應用了選擇後才顯示效率圖表
    if 'has_applied_selection' in st.session_state and st.session_state.has_applied_selection:
        # 獲取應用選擇後保存的參數
        selected_subcategories = st.session_state.get('selected_subcategories', [])
        merge_categories = st.session_state.get('merge_categories', True)
        selected_category_mode = st.session_state.get('selected_category_mode', "subcategory")
        
        # 顯示效率分析結果
        if selected_subcategories:
            if selected_category_mode == "category" or merge_categories:
                # 合併顯示模式
                if selected_category_mode == "category" and 'selected_main_category' in st.session_state:
                    merged_category = st.session_state.selected_main_category  # 使用大類別名稱
                else:
                    merged_category = "+".join(selected_subcategories)
                
                # 查找可能的鍵名
                merged_category_key = merged_category
                if merged_category not in category_efficiency_data and "+".join(selected_subcategories) in category_efficiency_data:
                    merged_category_key = "+".join(selected_subcategories)
                
                if merged_category_key in category_efficiency_data:
                    data = category_efficiency_data[merged_category_key]
                    
                    df = pd.DataFrame(data)
                    
                    # 設置最大效率顯示值，避免異常值影響圖表比例
                    max_display_value = min(5, df['efficiency'].max() * 1.1) if not df.empty else 5
                    
                    # 使用與總效率相同的配色方案
                    fig = px.bar(
                        df,
                        x='inspector',
                        y='efficiency',
                        labels={'inspector': '檢驗員', 'efficiency': '檢驗效率'},
                        title=f"{merged_category} 類別效率排名",  # 加入所選類別名稱
                        range_y=[0, max_display_value],  # 限制Y軸範圍
                        color='efficiency',
                        color_continuous_scale="RdBu_r",  # 使用內建紅藍刻度
                        range_color=[0, 2.0]  # 設置顏色範圍，使1在中間位置
                    )

                    # 添加水平参考线 (效率=1)
                    fig.add_shape(
                        type="line",
                        x0=-0.5,
                        y0=1,
                        x1=len(df)-0.5,
                        y1=1,
                        line=dict(
                            color="#90A4AE",  # 淺灰藍色 - 與主題一致
                            width=1.5,        # 適中線寬
                            dash=None         # 實線而非虛線
                        )
                    )
                    
                    # 添加與總效率相同的顏色標籤設置
                    fig.update_layout(
                        xaxis_title="檢驗員",
                        yaxis_title="檢驗效率",
                        height=400,
                        coloraxis_colorbar=dict(
                            title="檢驗效率標籤",
                            tickvals=[0, 0.5, 1.0, 1.5, 2.0],
                            ticktext=["0", "0.5", "1.0", "1.5", "2.0"],
                            len=0.5,  # 調整顏色條長度
                            thickness=15,  # 調整顏色條粗細
                            y=0.5,  # 居中放置
                            yanchor="middle",
                            x=1.02,  # 調整右側間距
                            xanchor="left",
                            outlinewidth=0,  # 移除框線
                            borderwidth=0    # 移除邊框
                        )
                    )
                    
                    st.plotly_chart(fig, use_container_width=True, key=f"merged_category_chart")
                    
                    # 顯示類別效率數據明細
                    with st.expander(f"檢視類別效率數據明細"):
                        display_category_details(df)
                else:
                    st.info(f"沒有 {merged_category} 的合併數據。")
            else:
                # 單獨顯示模式 - 仅在子类别模式下使用
                for i, category in enumerate(selected_subcategories):
                    if category in category_efficiency_data:
                        data = category_efficiency_data[category]
                        
                        if not data:
                            st.write("沒有此類別的數據")
                            continue
                        
                        df = pd.DataFrame(data)
                        
                        # 設置最大效率顯示值，避免異常值影響圖表比例
                        max_display_value = min(5, df['efficiency'].max() * 1.1) if not df.empty else 5
                        
                        # 使用與總效率相同的配色方案
                        fig = px.bar(
                            df,
                            x='inspector',
                            y='efficiency',
                            labels={'inspector': '檢驗員', 'efficiency': '檢驗效率'},
                            title=f"{category} 類別效率排名",  # 加入當前類別名稱
                            range_y=[0, max_display_value],  # 限制Y軸範圍
                            color='efficiency',
                            color_continuous_scale="RdBu_r",  # 使用內建紅藍刻度
                            range_color=[0, 2.0]  # 設置顏色範圍，使1在中間位置
                        )
                        
                        # 添加水平参考线 (效率=1)
                        fig.add_shape(
                            type="line",
                            x0=-0.5,
                            y0=1,
                            x1=len(df)-0.5,
                            y1=1,
                            line=dict(
                                color="#90A4AE",  # 淺灰藍色 - 與主題一致
                                width=1.5,        # 適中線寬
                                dash=None         # 實線而非虛線
                            )
                        )
                        
                        # 添加與總效率相同的顏色標籤設置
                        fig.update_layout(
                            xaxis_title="檢驗員",
                            yaxis_title="檢驗效率",
                            height=400,
                            coloraxis_colorbar=dict(
                                title="檢驗效率",
                                tickvals=[0, 0.5, 1.0, 1.5, 2.0],
                                ticktext=["0", "0.5", "1.0", "1.5", "2.0"],
                                len=0.5,  # 調整顏色條長度
                                thickness=15,  # 調整顏色條粗細
                                y=0.5,  # 居中放置
                                yanchor="middle",
                                x=1.02,  # 調整右側間距
                                xanchor="left",
                                outlinewidth=0,  # 移除框線
                                borderwidth=0    # 移除邊框
                            )
                        )
                        
                        st.plotly_chart(fig, use_container_width=True, key=f"category_chart_{i}")
                        
                        # 顯示類別效率數據明細
                        with st.expander(f"檢視 {category} 類別效率數據明細"):
                            display_category_details(df)
                    else:
                        st.info(f"沒有 {category} 的數據。")
        else:
            st.info("請選擇物料類別")
    else:
        # 如果用户还没有应用选择，不显示效率图表
        st.markdown("""
            <div style="background-color:#f0f2f5; border-left:4px solid #78909C; border-radius:4px; padding:16px; margin-top:20px; display:flex; align-items:center;">
                <div>
                    <p style="margin:0; color:#37474F; font-weight:400;">"請選擇分析模式和物料類別，然後點擊「應用選擇」按鈕來查看效率排名。"</p>
                </div>
            </div>
            """, unsafe_allow_html=True)

def render_category_efficiency_chart(df, category_name, max_display_value=5):
    """渲染物料類別效率圖表，包含顏色標籤"""
    
    # 手動設置顏色
    colors = []
    for eff in df['efficiency']:
        if eff < 1:
            # 效率小於1使用紅色系
            red_intensity = max(100, 255 - int((1 - eff) * 155))
            colors.append(f'rgba({red_intensity}, 0, 0, 0.8)')
        else:
            # 效率大於等於1使用藍色系
            blue_intensity = max(100, 255 - int((eff - 1) * 85))
            colors.append(f'rgba(0, 0, {blue_intensity}, 0.8)')
    
    # 創建基本圖表
    fig = go.Figure()
    
    # 添加條形圖
    fig.add_trace(go.Bar(
        x=df['inspector'],
        y=df['efficiency'],
        marker_color=colors
    ))
    
    # 添加參考線
    fig.add_shape(
        type="line",
        x0=-0.5,
        y0=1,
        x1=len(df)-0.5,
        y1=1,
        line=dict(color="black", width=2, dash="dash"),
    )
    
    # 添加顏色標籤 - 使用隱藏的散點圖
    fig.add_trace(
        go.Scatter(
            x=[None], y=[None],
            mode='markers',
            marker=dict(
                colorscale=[
                    [0, "rgb(180, 0, 0)"],    # 深紅色 (效率為0)
                    [0.25, "rgb(255, 0, 0)"],  # 紅色 (效率為0.5)
                    [0.5, "rgb(255, 150, 150)"], # 淺紅色 (效率為1)
                    [0.5, "rgb(150, 150, 255)"], # 淺藍色 (效率為1)
                    [0.75, "rgb(0, 0, 255)"],   # 藍色 (效率為1.5)
                    [1, "rgb(0, 0, 150)"]      # 深藍色 (效率為2+)
                ],
                showscale=True,
                cmin=0,
                cmax=2,
                colorbar=dict(
                    title="檢驗效率",
                    tickvals=[0, 0.5, 1.0, 1.5, 2.0],
                    ticktext=["0", "0.5", "1.0", "1.5", "2.0"],
                    len=0.5,
                    thickness=15,
                    y=0.5,
                    yanchor="middle",
                    x=1.02,
                    xanchor="left",
                    outlinewidth=0,  # 移除框線
                    borderwidth=0    # 移除邊框
                )
            ),
            hoverinfo='none',
            showlegend=False
        )
    )
    
    # 設置標題和軸標籤
    fig.update_layout(
        title=f"{category_name} 物料類別效率排名",
        xaxis_title="檢驗員",
        yaxis_title="檢驗效率",
        height=400
    )
    
    return fig


# 修改物料類別效率分析圖表，添加顏色圖例
def display_category_efficiency_chart(category_efficiency_data, merged_category_key, selected_subcategories):
    """
    显示物料类别效率图表，并添加颜色图例
    """
    if merged_category_key in category_efficiency_data:
        data = category_efficiency_data[merged_category_key]
        
        df = pd.DataFrame(data)
        
        # 设置最大效率显示值，避免异常值影响图表比例
        max_display_value = min(5, df['efficiency'].max() * 1.1) if not df.empty else 5
        
        # 创建两列布局，一列用于图表，一列用于颜色图例
        col1, col2 = st.columns([4, 1])
        
        with col1:
            # 绘制物料类别效率排名图 - 使用红蓝色阶
            fig = px.bar(
                df,
                x='inspector',
                y='efficiency',
                labels={'inspector': '檢驗員', 'efficiency': '檢驗效率'},
                title=f"{merged_category_key} 物料類別效率排名",
                range_y=[0, max_display_value]  # 限制Y轴范围
            )
            
            # 修改颜色设置为红蓝色阶
            colors = []
            for row in df.itertuples():
                if row.efficiency < 1:
                    # 效率小于1使用红色渐变
                    red_intensity = max(50, 255 - int((1 - row.efficiency) * 200))
                    colors.append(f'rgba({red_intensity}, 0, 0, 0.8)')
                else:
                    # 效率大于等于1使用蓝色渐变
                    blue_intensity = max(50, 255 - int((row.efficiency - 1) * 50))
                    colors.append(f'rgba(0, 0, {blue_intensity}, 0.8)')
            
            # 一次性设置所有颜色
            fig.update_traces(marker_color=colors)
            
            # 添加水平参考线 (效率=1)
            fig.add_shape(
                type="line",
                x0=-0.5,
                y0=1,
                x1=len(df)-0.5,
                y1=1,
                line=dict(
                    color="#90A4AE",  # 淺灰藍色 - 與主題一致
                    width=1.5,        # 適中線寬
                    dash=None         # 實線而非虛線
                )
            )
            
            fig.update_layout(
                xaxis_title="檢驗員",
                yaxis_title="檢驗效率",
                height=400
            )
            
            st.plotly_chart(fig, use_container_width=True)
        
        with col2:
            # 添加颜色图例
            st.markdown("### 類別效率標籤")
            
            # 创建颜色图例图表
            color_scale_data = pd.DataFrame({
                'value': [0, 0.1, 0.2, 0.3, 0.4, 0.5],
                'color': ['darkred', 'red', 'lightred', 'lightblue', 'blue', 'darkblue']
            })
            
            fig_colorbar = px.bar(
                color_scale_data,
                x=['效率']*len(color_scale_data),
                y='value',
                color='value',
                color_continuous_scale=[[0, "rgb(150, 0, 0)"],  # 深红色
                                       [0.25, "rgb(255, 0, 0)"],  # 红色
                                       [0.5, "rgb(255, 100, 100)"],  # 浅红色
                                       [0.5, "rgb(100, 100, 255)"],  # 浅蓝色
                                       [0.75, "rgb(0, 0, 255)"],  # 蓝色
                                       [1, "rgb(0, 0, 150)"]],  # 深蓝色
                labels={'value': '效率值'},
                range_color=[0, 0.5]
            )
            
            fig_colorbar.update_layout(
                showlegend=False,
                coloraxis_colorbar=dict(
                    title="類別效率",
                    tickvals=[0, 0.1, 0.2, 0.3, 0.4, 0.5],
                    ticktext=["0", "0.1", "0.2", "0.3", "0.4", "0.5"],
                    len=0.8,
                    thickness=15,
                    x=0.8,
                    y=0.5
                ),
                margin=dict(l=0, r=0, t=40, b=0),
                height=300
            )
            
            # 隐藏坐标轴
            fig_colorbar.update_xaxes(visible=False)
            fig_colorbar.update_yaxes(visible=False)
            
            # 只显示颜色条
            fig_colorbar.update_traces(opacity=0)
            
            st.plotly_chart(fig_colorbar, use_container_width=True)
            
            # 添加效率值解释文字
            st.markdown("""
            **類別效率說明:**
            - **> 0.3**: 高效率（深藍色）
            - **> 0.2**: 良好效率（藍色）
            - **< 0.2**: 低效率（紅色）
            """)
        
        # 显示类别效率数据明细
        with st.expander(f"檢視 {merged_category_key} 類別效率數據明細"):
            display_category_details(df)
    else:
        st.info(f"沒有 {merged_category_key} 的合併數據。")

# 顯示類別效率數據明細的輔助函數
def display_category_details(detail_df):
    # 確保所有必要的列都存在
    if 'inspector' in detail_df.columns and 'efficiency' in detail_df.columns and 'record_count' in detail_df.columns:
        columns_to_display = ['inspector', 'efficiency', 'record_count']
        column_names = ['檢驗員', '效率', '紀錄筆數']
        
        # 如果有標準工時和實際耗時的資訊，也顯示出來
        if 'total_standard_time' in detail_df.columns and 'total_actual_time' in detail_df.columns:
            columns_to_display.extend(['total_standard_time', 'total_actual_time'])
            column_names.extend(['標準工時總和(分鐘)', '實際耗時總和(分鐘)'])
        
        display_df = detail_df[columns_to_display].copy()
        display_df.columns = column_names
        
        # 格式化數值
        if '效率' in display_df.columns:
            display_df['效率'] = display_df['效率'].round(2)
        if '標準工時總和(分鐘)' in display_df.columns:
            display_df['標準工時總和(分鐘)'] = display_df['標準工時總和(分鐘)'].round(0)
        if '實際耗時總和(分鐘)' in display_df.columns:
            display_df['實際耗時總和(分鐘)'] = display_df['實際耗時總和(分鐘)'].round(0)
        
        # 根據效率值添加背景色
        def highlight_efficiency(val):
            if val < 1:
                # 效率<1，使用紅色漸變
                red_intensity = max(70, 255 - int((1 - val) * 180))
                return f'background-color: rgba({red_intensity}, 0, 0, 0.3)'
            else:
                # 效率>=1，使用藍色漸變
                blue_intensity = max(70, 255 - int((val - 1) * 50))
                return f'background-color: rgba(0, 0, {blue_intensity}, 0.3)'
        
        # 應用樣式
        styled_df = display_df.style.applymap(
            highlight_efficiency, 
            subset=['效率']
        )
        
        st.dataframe(styled_df, use_container_width=True)
    else:
        st.write("數據格式不正確，無法顯示詳細資訊")

# 計算物料類別效率的獨立函數 (不影響總效率)
def calculate_category_efficiency(processed_df, selected_categories, merge_categories):
    if processed_df is None or processed_df.empty or not selected_categories:
        return {}
    
    debug_log(f"計算物料類別效率: 選擇類別={selected_categories}, 合併={merge_categories}")
    
    # 過濾出所選物料類別的數據
    filtered_df = processed_df[processed_df['類別'].isin(selected_categories)]
    
    if filtered_df.empty:
        debug_log("過濾後沒有數據")
        return {}
    
    # 準備存儲效率數據的字典
    category_efficiency_data = {}
    
    # 設定效率值的合理上限
    MAX_EFFICIENCY = 20
    
    # 計算每個IQC人員在各物料類別的效率
    for _, row in filtered_df.iterrows():
        inspector = row['處理後檢驗員'] if pd.notna(row['處理後檢驗員']) else 'Unknown'
        category = row['類別'] if pd.notna(row['類別']) else 'Unknown'
        
        # 如果合併類別，則將類別設為合併後的名稱
        if merge_categories:
            category = "+".join(selected_categories)
        
        # 處理標準工時和實際耗時
        try:
            standard_time = float(row['處理後檢驗標準工時']) if pd.notna(row['處理後檢驗標準工時']) else 0
        except (ValueError, TypeError):
            standard_time = 0
        
        try:
            actual_time = float(row['檢驗耗時']) if pd.notna(row['檢驗耗時']) else 0
        except (ValueError, TypeError):
            actual_time = 0
        
        # 檢查實際耗時是否過小
        if actual_time < 0.1:
            actual_time = 0.1
        
        # 物料類別效率計算
        key = f"{category}-{inspector}"
        if key not in category_efficiency_data:
            category_efficiency_data[key] = {
                'category': category,
                'inspector': inspector,
                'total_standard_time': 0,
                'total_actual_time': 0,
                'record_count': 0
            }
        
        category_efficiency_data[key]['total_standard_time'] += standard_time
        category_efficiency_data[key]['total_actual_time'] += actual_time
        category_efficiency_data[key]['record_count'] += 1
    
    # 按類別和檢驗員分組處理
    result = {}
    
    for key, data in category_efficiency_data.items():
        category = data['category']
        inspector = data['inspector']
        
        # 計算效率
        if data['total_actual_time'] > 0:
            category_efficiency = min(data['total_standard_time'] / data['total_actual_time'], MAX_EFFICIENCY)
        else:
            category_efficiency = 0
        
        if category not in result:
            result[category] = []
        
        result[category].append({
            'inspector': inspector,
            'efficiency': category_efficiency,
            'record_count': data['record_count'],
            'total_standard_time': data['total_standard_time'],
            'total_actual_time': data['total_actual_time']
        })
    
    # 為每個類別排序
    for category in result:
        result[category] = sorted(
            result[category],
            key=lambda x: x['efficiency'],
            reverse=True
        )
        debug_log(f"類別 {category} 效率排名計算完成，共 {len(result[category])} 位檢驗員")
    
    return result

def calculate_mrb_statistics(processed_data, start_date=None, end_date=None):
    """
    計算MRB統計數據，統一基於'MRB加時 > 0'判斷MRB狀態
    """
    if processed_data is None or processed_data.empty:
        return None
        
    # 過濾日期範圍
    filtered_data = filter_by_date_range(processed_data, start_date, end_date)
    
    if filtered_data.empty:
        return None
    
    # 確保有MRB加時欄位
    if 'MRB加時' not in filtered_data.columns:
        return None
    
    # 統一使用MRB加時判斷MRB狀態
    is_mrb = filtered_data['MRB加時'] > 0
    
    # 計算MRB統計
    total_count = len(filtered_data)
    mrb_count = is_mrb.sum()
    mrb_rate = mrb_count / total_count if total_count > 0 else 0
    
    return {
        'mrb_rate': mrb_rate,
        'mrb_count': mrb_count,
        'total_count': total_count,
        'filtered_data': filtered_data  # 保存過濾後的數據供後續分析
    }

def render_mrb_analysis_dashboard(processed_data, start_date=None, end_date=None):
    """顯示MRB分析儀表板，基於MRB加時判斷MRB狀態，確保每次都重新計算結果"""
    if processed_data is None or processed_data.empty:
        st.error("沒有可用的數據來進行MRB分析")
        return
    
    # 過濾處理後的數據 - 排除指定檢驗員
    processed_data = filter_excluded_inspectors(processed_data)
    
    if processed_data.empty:
        st.warning("過濾後沒有可用的數據來進行MRB分析。")
        return

    st.header("🔍MRB分析")
    
    # 重要修改：總是重新計算MRB統計數據，不依賴session state
    debug_log(f"重新計算MRB統計數據，日期範圍: {start_date} - {end_date}")
    
    # 過濾日期範圍
    filtered_data = filter_by_date_range(processed_data, start_date, end_date)
    
    if filtered_data is None or filtered_data.empty:
        st.warning(f"所選時間區段 ({start_date} 到 {end_date}) 內沒有數據")
        return
    
    # 計算MRB統計
    total_count = len(filtered_data)
    
    # 使用MRB加時>0來判斷MRB狀態
    if 'MRB加時' in filtered_data.columns:
        mrb_count = (filtered_data['MRB加時'] > 0).sum()
    elif '是否為MRB' in filtered_data.columns:
        # 備用方法：使用是否為MRB欄位
        if filtered_data['是否為MRB'].dtype == 'object':
            mrb_count = filtered_data['是否為MRB'].apply(
                lambda x: True if str(x).upper() in ['TRUE', 'MRB'] else False
            ).sum()
        else:
            mrb_count = filtered_data['是否為MRB'].sum()
    else:
        st.warning("數據中缺少MRB相關欄位，無法計算MRB率")
        return
    
    mrb_rate = mrb_count / total_count if total_count > 0 else 0
    
    # 將計算結果保存到session state以供其他頁面使用
    st.session_state.mrb_analysis_results = {
        'mrb_rate': mrb_rate,
        'mrb_count': mrb_count,
        'total_count': total_count,
        'filtered_data': filtered_data
    }
    
    # 顯示MRB統計
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("總體MRB率", f"{mrb_rate:.2%}")
    with col2:
        st.metric("MRB數量", f"{mrb_count}")
    with col3:
        st.metric("總檢驗筆數", f"{total_count}")
    
    # 輸出調試信息
    # debug_log(f"MRB分析結果: 率={mrb_rate:.2%}, 數量={mrb_count}, 總數={total_count}", level="INFO")
    
    # 每日MRB趨勢
    st.subheader("每日MRB趨勢📈 ")
    
    # 確保日期欄位
    date_column = '檢驗日期' if '檢驗日期' in filtered_data.columns else '日期'
    if date_column in filtered_data.columns:
        # 將日期轉換為datetime
        filtered_data[date_column] = pd.to_datetime(filtered_data[date_column])
        
        # 按日期分組計算MRB率
        daily_stats = []
        for date, group in filtered_data.groupby(pd.Grouper(key=date_column, freq='D')):
            if len(group) > 0:  # 跳過沒有數據的日期
                # 使用MRB加時>0判斷MRB
                if 'MRB加時' in group.columns:
                    group_mrb_count = (group['MRB加時'] > 0).sum()
                elif '是否為MRB' in group.columns:
                    group_mrb_count = group['是否為MRB'].sum()
                else:
                    continue
                    
                group_total = len(group)
                group_mrb_rate = group_mrb_count / group_total
                
                daily_stats.append({
                    '日期': date,
                    'MRB率': group_mrb_rate,
                    'MRB數量': group_mrb_count,
                    '總筆數': group_total
                })
        
        # 繪製趨勢圖
        if daily_stats:
            daily_df = pd.DataFrame(daily_stats)
            
            fig = px.line(
                daily_df,
                x='日期',
                y='MRB率',
                markers=True,
                labels={'日期': '日期', 'MRB率': 'MRB率'},
                title="每日MRB率趨勢",
                height=400
            )
            
            # 添加MRB數量作為懸浮提示，使用自定義格式
            fig.update_traces(
                hovertemplate='日期: %{x}<br>MRB率: %{y:.2%}<br>MRB數量: %{text}<br>總筆數: %{customdata}',
                text=daily_df['MRB數量'],
                customdata=daily_df['總筆數']
            )
            
            # 設置Y軸為百分比格式
            fig.update_layout(
                yaxis=dict(
                    tickformat=".0%",
                    title="MRB率"
                ),
                xaxis_title="日期"
            )
            
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("沒有足夠的數據來繪製趨勢圖")
    
    # IQC人員MRB率
    st.subheader("IQC人員MRB率📈")
    
    # 按檢驗員分組計算MRB率
    inspector_column = '處理後檢驗員' if '處理後檢驗員' in filtered_data.columns else '檢驗員'
    if inspector_column in filtered_data.columns:
        inspector_stats = []
        for inspector, group in filtered_data.groupby(inspector_column):
            inspector_total = len(group)
            # 使用MRB加時>0判斷MRB
            if 'MRB加時' in group.columns:
                inspector_mrb_count = (group['MRB加時'] > 0).sum()
            elif '是否為MRB' in group.columns:
                inspector_mrb_count = group['是否為MRB'].sum()
            else:
                continue
                
            inspector_mrb_rate = inspector_mrb_count / inspector_total if inspector_total > 0 else 0
            
            inspector_stats.append({
                '檢驗員': inspector,
                'MRB數量': inspector_mrb_count,
                '總檢驗筆數': inspector_total,
                'MRB率': inspector_mrb_rate
            })
        
        # 繪製條形圖
        if inspector_stats:
            inspector_df = pd.DataFrame(inspector_stats)
            inspector_df = inspector_df.sort_values('MRB率', ascending=False)
            
            fig = px.bar(
                inspector_df,
                x='檢驗員',
                y='MRB率',
                text=inspector_df['MRB率'].apply(lambda x: f"{x:.1%}"),
                labels={'檢驗員': '檢驗員', 'MRB率': 'MRB率 (%)'},
                title="IQC人員MRB率",
                height=500,
                color='MRB率',
                color_continuous_scale=[
                    [0, "rgb(220, 220, 220)"],  # 淺灰色 (低MRB率)
                    [0.05, "rgb(255, 200, 200)"],  # 淺紅色
                    [0.1, "rgb(255, 100, 100)"],  # 中紅色
                    [0.2, "rgb(200, 0, 0)"],  # 深紅色 (高MRB率)
                    [1, "rgb(128, 0, 0)"]  # 極深紅色 (極高MRB率)
                ],
                range_color=[0, 0.2]  # 設置顏色範圍，0-20%
            )
            
            # 添加自定義懸浮提示模板
            fig.update_traces(
                hovertemplate='<b>檢驗員: %{x}</b><br>' +
                            'MRB率: %{y:.2%}<br>' +  # 格式化為百分比
                            'MRB數量: %{customdata[0]}<br>' +
                            '總檢驗筆數: %{customdata[1]}',
                customdata=inspector_df[['MRB數量', '總檢驗筆數']].values
            )
            
            fig.update_layout(
                yaxis=dict(
                    tickformat=".0%",
                    title="MRB率"
                ),
                xaxis_title="檢驗員"
            )
            
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("沒有足夠的數據來顯示IQC人員MRB率")
    
    # 物料類別MRB率
    if '類別' in filtered_data.columns:
        st.subheader("物料類別MRB率📈")
        
        category_stats = []
        for category, group in filtered_data.groupby('類別'):
            category_total = len(group)
            # 使用MRB加時>0判斷MRB
            if 'MRB加時' in group.columns:
                category_mrb_count = (group['MRB加時'] > 0).sum()
            elif '是否為MRB' in group.columns:
                category_mrb_count = group['是否為MRB'].sum()
            else:
                continue
                
            category_mrb_rate = category_mrb_count / category_total if category_total > 0 else 0
            
            category_stats.append({
                '物料類別': category,
                'MRB數量': category_mrb_count,
                '總檢驗筆數': category_total,
                'MRB率': category_mrb_rate
            })
        
        # 繪製條形圖
        if category_stats:
            category_df = pd.DataFrame(category_stats)
            category_df = category_df.sort_values('MRB率', ascending=False)
            
            fig = px.bar(
                category_df,
                x='物料類別',
                y='MRB率',
                text=category_df['MRB率'].apply(lambda x: f"{x:.1%}"),
                labels={'物料類別': '物料類別', 'MRB率': 'MRB率 (%)'},
                title="物料類別MRB率",
                height=500,
                color='MRB率',
                color_continuous_scale=[
                    [0, "rgb(220, 220, 220)"],  # 淺灰色 (低MRB率)
                    [0.1, "rgb(255, 200, 200)"],  # 淺紅色
                    [0.2, "rgb(255, 100, 100)"],  # 中紅色
                    [0.5, "rgb(200, 0, 0)"],  # 深紅色 (高MRB率)
                    [1, "rgb(128, 0, 0)"]  # 極深紅色 (極高MRB率)
                ],
                range_color=[0, 0.5]  # 設置顏色範圍，0-50%
            )
            
            # 添加自定義懸浮提示模板
            fig.update_traces(
                hovertemplate='<b>物料類別: %{x}</b><br>' +
                            'MRB率: %{y:.2%}<br>' +  # 格式化為百分比
                            'MRB數量: %{customdata[0]}<br>' +
                            '總檢驗筆數: %{customdata[1]}',
                customdata=category_df[['MRB數量', '總檢驗筆數']].values
            )
            
            fig.update_layout(
                yaxis=dict(
                    tickformat=".0%",
                    title="MRB率"
                ),
                xaxis_title="物料類別"
            )
            
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("沒有足夠的數據來顯示物料類別MRB率")

def calculate_overall_mrb_rate(data):
    """
    計算並顯示總體MRB率
    """
    total_count = len(data)
    # 確保將MRB欄位轉換為布林值
    if data['是否為MRB'].dtype == 'object':
        mrb_count = data['是否為MRB'].apply(
            lambda x: True if isinstance(x, str) and x.lower() in ['true', 'mrb'] else 
                       (True if x is True else False)
        ).sum()
    else:
        mrb_count = data['是否為MRB'].sum()
    
    mrb_rate = mrb_count / total_count if total_count > 0 else 0
    
    # 使用st.metric顯示總體MRB率
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("總體MRB率", f"{mrb_rate:.2%}")
    with col2:
        st.metric("MRB數量", f"{mrb_count}")
    with col3:
        st.metric("總檢驗筆數", f"{total_count}")
    
    return mrb_rate, mrb_count, total_count

def plot_daily_mrb_trend(data):
    """
    繪製每日MRB趨勢折線圖
    """
    st.subheader("每日MRB趨勢")
    
    # 確保日期欄位
    date_column = '檢驗日期' if '檢驗日期' in data.columns else '日期'
    if date_column not in data.columns:
        st.warning("數據中缺少日期欄位，無法繪製趨勢圖")
        return
    
    # 將日期轉換為datetime
    data[date_column] = pd.to_datetime(data[date_column])
    
    # 按日期分組計算MRB率
    daily_stats = []
    for date, group in data.groupby(pd.Grouper(key=date_column, freq='D')):
        if len(group) > 0:  # 跳過沒有數據的日期
            mrb_count = group['是否為MRB'].sum()
            total_count = len(group)
            mrb_rate = mrb_count / total_count
            daily_stats.append({
                '日期': date,
                'MRB率': mrb_rate,
                'MRB數量': mrb_count,
                '總筆數': total_count
            })
    
    if not daily_stats:
        st.info("所選時間區段內沒有每日MRB數據")
        return
    
    daily_df = pd.DataFrame(daily_stats)
    
    # 繪製折線圖
    fig = px.line(
        daily_df,
        x='日期',
        y='MRB率',
        markers=True,
        labels={'日期': '日期', 'MRB率': 'MRB率'},
        title="每日MRB率趨勢",
        height=400
    )
    
    # 添加MRB數量作為懸浮提示
    fig.update_traces(
        hovertemplate='日期: %{x}<br>MRB率: %{y:.2%}<br>MRB數量: %{customdata[0]}<br>總筆數: %{customdata[1]}',
        customdata=daily_df[['MRB數量', '總筆數']].values
    )
    
    # 設置Y軸為百分比格式
    fig.update_layout(
        yaxis=dict(
            tickformat=".0%",
            title="MRB率"
        ),
        xaxis_title="日期"
    )
    
    st.plotly_chart(fig, use_container_width=True)

def show_inspector_mrb_rates(data):
    """
    顯示IQC人員MRB率
    """
    st.subheader("IQC人員MRB率")
    
    # 確保有檢驗員欄位
    inspector_column = '處理後檢驗員' if '處理後檢驗員' in data.columns else '檢驗員'
    if inspector_column not in data.columns:
        st.warning("數據中缺少檢驗員欄位，無法顯示人員MRB率")
        return
    
    # 計算每位檢驗員的MRB率
    mrb_stats = []
    for inspector, group in data.groupby(inspector_column):
        total_count = len(group)
        mrb_count = group['是否為MRB'].sum()
        mrb_rate = mrb_count / total_count if total_count > 0 else 0
        
        mrb_stats.append({
            '檢驗員': inspector,
            'MRB數量': mrb_count,
            '總檢驗筆數': total_count,
            'MRB率': mrb_rate
        })
    
    # 轉換為DataFrame並按MRB率排序
    mrb_rate_df = pd.DataFrame(mrb_stats)
    mrb_rate_df = mrb_rate_df.sort_values('MRB率', ascending=False)
    
    # 創建MRB率條形圖
    fig = px.bar(
        mrb_rate_df,
        x='檢驗員',
        y='MRB率',
        text=mrb_rate_df['MRB率'].apply(lambda x: f"{x:.1%}"),
        labels={'檢驗員': '檢驗員', 'MRB率': 'MRB率 (%)'},
        title="IQC人員MRB率 - MRB數量/總檢驗筆數",
        height=500,
        custom_data=mrb_rate_df[['MRB數量', '總檢驗筆數']].values  # 添加自定義數據用於懸浮提示
    )
    
    # 使用色階顯示不同的MRB率 - 較高的MRB率用較深的紅色表示
    colors = []
    for rate in mrb_rate_df['MRB率']:
        # 根據MRB率設置顏色深淺
        red_intensity = min(255, int(100 + rate * 400))  # 調整紅色的深淺
        colors.append(f'rgba({red_intensity}, 50, 50, 0.8)')
    
    # 添加自定義懸浮提示模板
    fig.update_traces(
        hovertemplate='<b>檢驗員: %{x}</b><br>' +
                     'MRB率: %{y:.2%}<br>' +  # 格式化為百分比
                     'MRB數量: %{customdata[0]}<br>' +
                     '總檢驗筆數: %{customdata[1]}',
        marker_color=colors, 
        textposition='outside'
    )
    
    # 添加百分比標籤
    fig.update_layout(
        yaxis=dict(
            tickformat=".0%",  # Y軸顯示為百分比
            title="MRB率"
        ),
        xaxis_title="檢驗員",
        margin=dict(t=80, b=50),
        uniformtext_minsize=10,
        uniformtext_mode='hide'
    )
    
    st.plotly_chart(fig, use_container_width=True)
    
    # 顯示詳細數據
    with st.expander("查看IQC人員MRB率詳細數據"):
        detail_df = mrb_rate_df.copy()
        detail_df['MRB率'] = detail_df['MRB率'].apply(lambda x: f"{x:.2%}")  # 格式化為百分比
        st.dataframe(detail_df, use_container_width=True)

def show_material_category_mrb_rates(data):
    """
    顯示物料類別MRB率
    """
    st.subheader("物料類別MRB率")
    
    # 確保有類別欄位
    if '類別' not in data.columns:
        st.warning("數據中缺少'類別'欄位，無法顯示物料類別MRB率")
        return
    
    # 計算每個物料類別的MRB率
    category_stats = []
    for category, group in data.groupby('類別'):
        if pd.notna(category) and category != '':  # 排除空類別
            total_count = len(group)
            mrb_count = group['是否為MRB'].sum()
            mrb_rate = mrb_count / total_count if total_count > 0 else 0
            
            category_stats.append({
                '物料類別': category,
                'MRB數量': mrb_count,
                '總檢驗筆數': total_count,
                'MRB率': mrb_rate
            })
    
    # 轉換為DataFrame並按MRB率排序
    cat_mrb_rate_df = pd.DataFrame(category_stats)
    cat_mrb_rate_df = cat_mrb_rate_df.sort_values('MRB率', ascending=False)
    
    # 創建MRB率條形圖
    fig = px.bar(
        cat_mrb_rate_df,
        x='物料類別',
        y='MRB率',
        text=cat_mrb_rate_df['MRB率'].apply(lambda x: f"{x:.1%}"),
        labels={'物料類別': '物料類別', 'MRB率': 'MRB率 (%)'},
        title="物料類別MRB率 - MRB數量/總檢驗筆數",
        height=500,
        color='MRB率',
        color_continuous_scale=[ ... ],  # 保留原有顏色設定
        custom_data=['物料類別', 'MRB數量', '總檢驗筆數']  # 添加自定義數據用於懸浮提示
    )

    # 然後添加自定義懸浮提示模板
    fig.update_traces(
        hovertemplate='<b>物料類別: %{customdata[0]}</b><br>' +
                    'MRB率: %{y:.2%}<br>' +
                    'MRB數量: %{customdata[1]}<br>' +
                    '總檢驗筆數: %{customdata[2]}',
        textposition='outside'
    )
    
    # 使用色階顯示不同的MRB率 - 較高的MRB率用較深的紅色表示
    colors = []
    for rate in cat_mrb_rate_df['MRB率']:
        # 根據MRB率設置顏色深淺
        red_intensity = min(255, int(100 + rate * 400))  # 調整紅色的深淺
        colors.append(f'rgba({red_intensity}, 50, 50, 0.8)')
    
    fig.update_traces(marker_color=colors, textposition='outside')
    
    # 添加百分比標籤
    fig.update_layout(
        yaxis=dict(
            tickformat=".0%",  # Y軸顯示為百分比
            title="MRB率"
        ),
        xaxis_title="物料類別",
        margin=dict(t=80, b=50),
        uniformtext_minsize=10,
        uniformtext_mode='hide'
    )
    
    st.plotly_chart(fig, use_container_width=True)
    
    # 顯示詳細數據
    with st.expander("查看物料類別MRB率詳細數據"):
        detail_df = cat_mrb_rate_df.copy()
        detail_df['MRB率'] = detail_df['MRB率'].apply(lambda x: f"{x:.2%}")  # 格式化為百分比
        st.dataframe(detail_df, use_container_width=True)

def render_workload_dashboard(workload_data):
    if workload_data is None or workload_data.empty:
        st.error("沒有可用的工作負載數據，請確保上傳了正確的檔案格式或檢查過濾條件。")
        return
    
    # 過濾工作負載數據 - 排除指定檢驗員
    workload_data = filter_excluded_inspectors(workload_data, inspector_column='inspector')
    
    if workload_data.empty:
        st.warning("過濾後沒有可用的工作負載數據。")
        return
    
    st.header("⏱️ IQC人員工作負載監控")
    
    # 使用熱力圖顯示工作負載指數趨勢
    st.subheader("IQC人員工作負載指數趨勢 (使用標準檢驗工時計算)")
    
    # 將數據整理成熱力圖所需格式
    # 將日期轉為日期時間類型
    workload_data['date'] = pd.to_datetime(workload_data['date'])
    
    # 檢查日期範圍，如果超過2個月則以週為單位顯示
    date_range = (workload_data['date'].max() - workload_data['date'].min()).days
    use_weekly_view = date_range > 60  # 超過60天（約2個月）使用週視圖
    
    if use_weekly_view:
        # 創建週標識，格式為"年-週號"
        workload_data['week'] = workload_data['date'].dt.strftime('%Y-%U')
        # 取每週的第一天作為標籤
        week_start_dates = workload_data.groupby('week')['date'].min().reset_index()
        week_mapping = dict(zip(week_start_dates['week'], week_start_dates['date']))
        
        # 使用週分組創建透視表
        pivot_data = workload_data.groupby(['inspector', 'week'])['workload_index'].mean().unstack(fill_value=0)
        
        # 按週開始日期對列進行排序
        week_order = sorted(week_mapping.items(), key=lambda x: x[1])
        ordered_weeks = [w[0] for w in week_order]
        pivot_data = pivot_data[ordered_weeks]
        
        # 創建X軸標籤（週開始日期的月/日格式）
        x_labels = [week_mapping[week].strftime('%m/%d') for week in ordered_weeks]
    else:
        # 正常日期視圖的透視表
        pivot_data = workload_data.pivot_table(
            index='inspector', 
            columns='date', 
            values='workload_index',
            aggfunc='mean'  # 如果同一天有多個值，取平均值
        ).fillna(0)
        
        # 使用日期作為X軸標籤
        x_labels = pivot_data.columns.strftime('%m/%d')
    
    # 重新排序行，按平均工作負載降序
    row_means = pivot_data.mean(axis=1)
    pivot_data = pivot_data.loc[row_means.sort_values(ascending=False).index]
    
    # 生成熱力圖
    fig = px.imshow(
        pivot_data,
        labels=dict(x="日期", y="檢驗員", color="工作負載指數"),
        x=x_labels,  # 使用適當的x軸標籤
        y=pivot_data.index,
        color_continuous_scale=[
            [0, 'rgb(220, 220, 220)'],     # 低負載: 淺灰色
            [0.5, 'rgb(144, 238, 144)'],   # 中等負載: 淺綠色
            [0.7, 'rgb(255, 165, 0)'],     # 高負載: 橙色
            [1, 'rgb(220, 20, 60)']        # 極高負載: 紅色
        ],
        range_color=[0, 1.5],  # 設置色彩範圍，1.0是標準工作負載
        aspect="auto",         # 自動調整長寬比
        height=500
    )
    
    # 優化佈局
    fig.update_layout(
        coloraxis_colorbar=dict(
            title="工作負載指數",
            tickvals=[0, 0.5, 1.0, 1.5],
            ticktext=["0", "0.5", "1.0", "1.5+"]
        ),
        margin=dict(l=50, r=50, t=50, b=50)
    )
    
    st.plotly_chart(fig, use_container_width=True)
    
    # 顯示各檢驗員的平均工作負載
    st.subheader("各檢驗員平均工作負載⏳ ")
    
    avg_workload = workload_data.groupby('inspector')['workload_index'].mean().reset_index()
    avg_workload.columns = ['檢驗員', '平均工作負載指數']
    avg_workload['平均工作負載指數'] = avg_workload['平均工作負載指數'].round(2)
    avg_workload = avg_workload.sort_values('平均工作負載指數', ascending=False)
    
    # 最大工作負載值，用于设置颜色范围
    max_workload = max(5, avg_workload['平均工作負載指數'].max() * 1.1)
    
    # 創建條形圖，正確對應顏色和標籤順序
    fig_avg = px.bar(
        avg_workload,
        x='檢驗員',
        y='平均工作負載指數',
        title="各檢驗員平均工作負載指數 (使用標準檢驗工時計算)",  # 更新說明
        color='平均工作負載指數',
        color_continuous_scale=[
            [0, 'rgb(220, 220, 220)'],      # 低負載: 淺灰色
            [0.33, 'rgb(144, 238, 144)'],   # 中等負載: 淺綠色
            [0.67, 'rgb(255, 165, 0)'],     # 高負載: 橙色
            [1.0, 'rgb(220, 20, 60)']       # 極高負載: 紅色
        ],
        range_color=[0, 1.5]
    )
    
    # 添加參考線 (工作負載為1的水平線)
    fig_avg.add_shape(
        type="line",
        x0=-0.5,
        y0=1,
        x1=len(avg_workload)-0.5,
        y1=1,
        line=dict(color="black", width=2, dash="dash"),
    )
    
    # 修正顏色標籤與文字對應
    fig_avg.update_layout(
        xaxis_title="檢驗員",
        yaxis_title="平均工作負載指數",
        height=500,
        coloraxis_colorbar=dict(
            title="工作負載標籤",
            tickvals=[0, 0.5, 1.0, 1.5],
            ticktext=["0", "0.5", "1.0", "1.5+"],
            len=0.5,           # 調整色標長度
            thickness=25,      # 色標粗細
            y=0.5,             # 居中放置
            yanchor="middle",
            x=1.05,            # 調整右側間距
            xanchor="left",
            title_font=dict(size=14),
            tickfont=dict(size=12)
        )
    )
    
    st.plotly_chart(fig_avg, use_container_width=True)
    
    # 新增: 添加檢驗負載分析
    st.subheader("各檢驗員平均檢驗負載⏳ ")

    # 計算每個檢驗員的總天數、檢驗時間和額外任務時間
    inspector_summary = workload_data.groupby('inspector').agg(
        day_count=('date', 'nunique'),
        total_inspection_time=('inspection_standard_time', 'sum'),
        total_additional_time=('additional_task_time', 'sum')
    ).reset_index()

    # 計算可用時間（480分鐘 * 天數 - 額外任務時間）
    inspector_summary['total_available_time'] = inspector_summary['day_count'] * 480 - inspector_summary['total_additional_time']
    # 確保最小可用時間，避免除以零
    inspector_summary['total_available_time'] = inspector_summary['total_available_time'].apply(lambda x: max(x, 0.1))

    # 計算檢驗負載指數
    inspector_summary['inspection_load_index'] = inspector_summary['total_inspection_time'] / inspector_summary['total_available_time']

    # 格式化顯示
    avg_inspection_load = inspector_summary[['inspector', 'inspection_load_index']].copy()
    avg_inspection_load.columns = ['檢驗員', '平均檢驗負載指數']
    avg_inspection_load['平均檢驗負載指數'] = avg_inspection_load['平均檢驗負載指數'].round(2)
    avg_inspection_load = avg_inspection_load.sort_values('平均檢驗負載指數', ascending=False)

    # 設定顏色範圍最大值，避免異常值影響圖表比例
    max_inspection_load = max(5, avg_inspection_load['平均檢驗負載指數'].max() * 1.1)

    # 創建條形圖
    fig_inspection_load = px.bar(
        avg_inspection_load,
        x='檢驗員',
        y='平均檢驗負載指數',
        title="各檢驗員平均檢驗負載指數 (檢驗時間÷可用時間)",
        color='平均檢驗負載指數',
        color_continuous_scale=[
            [0, 'rgb(220, 220, 220)'],      # 低負載: 淺灰色
            [0.33, 'rgb(144, 238, 144)'],   # 中等負載: 淺綠色
            [0.67, 'rgb(255, 165, 0)'],     # 高負載: 橙色
            [1.0, 'rgb(220, 20, 60)']       # 極高負載: 紅色
        ],
        range_color=[0, 1.5]
    )

    # 添加參考線 (負載為1的水平線)
    fig_inspection_load.add_shape(
        type="line",
        x0=-0.5,
        y0=1,
        x1=len(avg_inspection_load)-0.5,
        y1=1,
        line=dict(color="black", width=2, dash="dash"),
    )

    # 設置圖表佈局
    fig_inspection_load.update_layout(
        xaxis_title="檢驗員",
        yaxis_title="平均檢驗負載指數",
        height=500,
        coloraxis_colorbar=dict(
            title="檢驗負載標籤",
            tickvals=[0, 0.5, 1.0, 1.5],
            ticktext=["0", "0.5", "1.0", "1.5+"],
            len=0.5,
            thickness=25,
            y=0.5,
            yanchor="middle",
            x=1.05,
            xanchor="left",
            title_font=dict(size=14),
            tickfont=dict(size=12)
        )
    )

    st.plotly_chart(fig_inspection_load, use_container_width=True)

    # 檢驗次數分析 - 改為物料大類別疊圖並改進懸停標籤
    if 'inspection_count' in workload_data.columns and 'processed_data' in st.session_state and not st.session_state.processed_data.empty:
        st.subheader("檢驗次數分析與物料類別分布🔢 ")
        
        # 獲取原始處理後的數據，用於物料類別分析
        processed_df = st.session_state.processed_data.copy()
        
        # 確保有類別欄位
        if '類別' in processed_df.columns:
            # 創建物料類別映射 - 將子類別映射到大類別
            category_to_main = {}
            for main_category, subcategories in MATERIAL_CATEGORY_MAPPING.items():
                for subcategory in subcategories:
                    category_to_main[subcategory] = main_category
            
            # 添加大類別欄位到原始數據
            processed_df['大類別'] = processed_df['類別'].apply(
                lambda x: category_to_main.get(x, '其他') if pd.notna(x) else '未分類'
            )
            
            # 獲取所有大類別列表（排序以保持一致的颜色）
            main_categories = sorted(MATERIAL_CATEGORY_MAPPING.keys())
            main_categories.append('其他')  # 添加"其他"類別用於未映射的物料
            
            # 計算每位檢驗員在每個大類別的檢驗次數
            inspector_category_counts = {}
            
            # 按檢驗員和日期分組
            for inspector, date_group in processed_df.groupby(['處理後檢驗員', '檢驗日期']):
                inspector_name = inspector[0]  # 檢驗員名稱
                
                if inspector_name not in inspector_category_counts:
                    inspector_category_counts[inspector_name] = {
                        '總次數': 0,
                        '日期數': set(),
                        '類別計數': {cat: 0 for cat in main_categories}
                    }
                
                # 記錄不重複的日期
                inspector_category_counts[inspector_name]['日期數'].add(inspector[1])
                
                # 按大類別統計當天的檢驗數量
                day_category_counts = date_group['大類別'].value_counts().to_dict()
                
                for category, count in day_category_counts.items():
                    if category in inspector_category_counts[inspector_name]['類別計數']:
                        inspector_category_counts[inspector_name]['類別計數'][category] += count
                    else:
                        inspector_category_counts[inspector_name]['類別計數']['其他'] += count
                    
                    inspector_category_counts[inspector_name]['總次數'] += count
            
            # 轉換為DataFrame用於繪圖
            stacked_data = []
            
            for inspector, data in inspector_category_counts.items():
                num_days = len(data['日期數'])
                if num_days > 0:  # 確保有工作天數
                    # 先計算所有類別的檢驗次數
                    category_counts = {}
                    for category in main_categories:
                        avg_count = data['類別計數'].get(category, 0) / num_days
                        if avg_count > 0:  # 只添加有檢驗次數的記錄
                            category_counts[category] = avg_count
                    
                    # 對類別進行排序 - 按照檢驗次數從高到低
                    sorted_categories = sorted(category_counts.items(), key=lambda x: x[1], reverse=True)
                    
                    # 創建排序信息字符串 - 用於懸停顯示
                    rank_info = "<br>".join([f"{i+1}. {cat}: {count:.1f}" 
                                             for i, (cat, count) in enumerate(sorted_categories)])
                    
                    # 計算總平均檢驗次數
                    total_avg_count = data['總次數'] / num_days
                    
                    # 添加每個類別的數據
                    for category, avg_count in category_counts.items():
                        stacked_data.append({
                            '檢驗員': inspector,
                            '物料大類別': category,
                            '平均檢驗次數': avg_count,
                            '總平均檢驗次數': total_avg_count,
                            '類別排名信息': rank_info  # 添加排名信息
                        })
            
            # 轉換為DataFrame
            stacked_df = pd.DataFrame(stacked_data)
            
            if not stacked_df.empty:
                # 按總平均檢驗次數排序
                inspectors_order = stacked_df.groupby('檢驗員')['總平均檢驗次數'].first().sort_values(ascending=False).index.tolist()
                
                # 在DataFrame中設置排序順序
                stacked_df['檢驗員'] = pd.Categorical(
                    stacked_df['檢驗員'], 
                    categories=inspectors_order, 
                    ordered=True
                )
                
                # 繪製疊圖 - 添加自定義懸停數據
                fig = px.bar(
                    stacked_df.sort_values('檢驗員'),
                    x='檢驗員',
                    y='平均檢驗次數',
                    color='物料大類別',
                    title="各檢驗員平均檢驗次數與物料類別分布",
                    labels={'檢驗員': '檢驗員', '平均檢驗次數': '平均檢驗次數', '物料大類別': '物料大類別'},
                    color_discrete_sequence=px.colors.qualitative.Pastel,
                    custom_data=['檢驗員', '物料大類別', '平均檢驗次數', '總平均檢驗次數', '類別排名信息']  # 修改自定義數據
                )
                
                # 自定義懸停模板 - 按照要求的格式修改
                fig.update_traces(
                    hovertemplate="<b>%{customdata[0]}</b><br>" +  # 檢驗員名稱
                                  "<b>%{customdata[1]}</b> 檢驗次數: %{customdata[2]:.1f}<br><br>" +  # 所選類別及其檢驗次數
                                  "平均檢驗次數: %{customdata[3]:.1f}<br><br>" +  # 平均檢驗次數
                                  "<b>類別檢驗次數排名:</b><br>" +  # 排名標題
                                  "%{customdata[4]}",  # 排名詳情
                    hoverlabel=dict(
                        bgcolor="white",
                        font_size=12,
                        font_family="Arial",
                        bordercolor="gray"  # 添加邊框顏色使標籤更明顯
                    )
                )
                
                # 添加數字標籤到每個檢驗員的總數
                for inspector in inspectors_order:
                    total = stacked_df[stacked_df['檢驗員'] == inspector]['平均檢驗次數'].sum()
                    fig.add_annotation(
                        x=inspector,
                        y=total,
                        text=f"{total:.1f}",
                        showarrow=False,
                        yshift=10,
                        font=dict(size=12, color="black")
                    )
                
                fig.update_layout(
                    xaxis_title="檢驗員",
                    yaxis_title="平均檢驗次數",
                    legend_title="物料大類別",
                    height=500,
                    barmode='stack',
                    hoverdistance=100,  # 提高懸停靈敏度
                    hovermode="closest"  # 懸停模式設置為最近點
                )
                
                st.plotly_chart(fig, use_container_width=True)
                
                # 顯示詳細數據
                with st.expander("查看物料類別檢驗次數詳細數據"):
                    # 按檢驗員和物料大類別透視
                    pivot_data = stacked_df.pivot_table(
                        values='平均檢驗次數',
                        index='檢驗員',
                        columns='物料大類別',
                        aggfunc='sum',
                        fill_value=0
                    )
                    
                    # 添加總計列
                    pivot_data['總平均檢驗次數'] = pivot_data.sum(axis=1)
                    
                    # 對每個值進行四捨五入
                    for col in pivot_data.columns:
                        pivot_data[col] = pivot_data[col].round(1)
                    
                    # 按總次數排序
                    pivot_data = pivot_data.sort_values('總平均檢驗次數', ascending=False)
                    
                    st.dataframe(pivot_data, use_container_width=True)
            else:
                st.info("沒有足夠的數據來顯示物料類別分析")
        else:
            # 如果沒有類別數據，顯示簡單的檢驗次數分析
            avg_inspections = workload_data.groupby('inspector')['inspection_count'].mean().reset_index()
            avg_inspections.columns = ['檢驗員', '平均檢驗次數']
            avg_inspections['平均檢驗次數'] = avg_inspections['平均檢驗次數'].round(1)
            avg_inspections = avg_inspections.sort_values('平均檢驗次數', ascending=False)
            
            fig_insp = px.bar(
                avg_inspections,
                x='檢驗員',
                y='平均檢驗次數',
                color='平均檢驗次數',
                color_continuous_scale='Blues',
                title="各檢驗員平均檢驗次數")
            
            fig_insp.update_layout(
                xaxis_title="檢驗員",
                yaxis_title="平均檢驗次數",
                height=400
            )
            
            st.plotly_chart(fig_insp, use_container_width=True)
    
    # 顯示工作負載詳細資料
    st.subheader("工作負載詳細資料📃")
    
    with st.expander("查看工作負載詳細資料"):
        # 處理數據用於顯示
        display_df = workload_data.copy()
        display_df['inspection_standard_time'] = display_df['inspection_standard_time'].round(2)  # 修改欄位名稱
        display_df['additional_task_time'] = display_df['additional_task_time'].round(2)
        display_df['total_time'] = display_df['total_time'].round(2)
        display_df['workload_index'] = display_df['workload_index'].round(2)
        
        # 重命名欄位
        rename_dict = {
            'date': '日期',
            'inspector': '檢驗員',
            'inspection_standard_time': '標準檢驗工時(分鐘)',  # 修改欄位名稱
            'additional_task_time': '額外任務時間(分鐘)',
            'total_time': '總時間(分鐘)',
            'workload_index': '工作負載指數'
        }
        
        # 添加新欄位的重命名
        if 'work_period' in display_df.columns:
            rename_dict['work_period'] = '工作時段'
        if 'inspection_count' in display_df.columns:
            rename_dict['inspection_count'] = '檢驗次數'
        
        display_df = display_df.rename(columns=rename_dict)
        
        st.dataframe(display_df, use_container_width=True)
    
    # 新增: 顯示檢驗負載詳細資料
    st.subheader("檢驗負載詳細資料📊")
    
    with st.expander("查看檢驗負載詳細資料"):
        # 計算每個檢驗員的詳細檢驗負載數據
        inspection_load_details = workload_data.groupby('inspector').agg(
            day_count=('date', 'nunique'),
            total_inspection_time=('inspection_standard_time', 'sum'),
            total_additional_time=('additional_task_time', 'sum')
        ).reset_index()
        
        # 計算可用時間和檢驗負載指數
        inspection_load_details['total_work_time'] = inspection_load_details['day_count'] * 480
        inspection_load_details['available_time'] = inspection_load_details['total_work_time'] - inspection_load_details['total_additional_time']
        # 確保最小可用時間，避免除以零
        inspection_load_details['available_time'] = inspection_load_details['available_time'].apply(lambda x: max(x, 0.1))
        inspection_load_details['inspection_load_index'] = inspection_load_details['total_inspection_time'] / inspection_load_details['available_time']
        
        # 格式化顯示
        inspection_load_details = inspection_load_details.rename(columns={
            'inspector': '檢驗員',
            'day_count': '工作天數',
            'total_inspection_time': '總檢驗時間(分鐘)',
            'total_additional_time': '總額外任務時間(分鐘)',
            'total_work_time': '總工作時間(分鐘)',
            'available_time': '可用時間(分鐘)',
            'inspection_load_index': '檢驗負載指數'
        })
        
        # 四捨五入小數位數
        for col in ['總檢驗時間(分鐘)', '總額外任務時間(分鐘)', '總工作時間(分鐘)', '可用時間(分鐘)']:
            inspection_load_details[col] = inspection_load_details[col].round(1)
        
        inspection_load_details['檢驗負載指數'] = inspection_load_details['檢驗負載指數'].round(2)
        
        # 添加計算公式說明列
        inspection_load_details['計算公式'] = inspection_load_details.apply(
            lambda row: f"檢驗負載 = {row['總檢驗時間(分鐘)']} ÷ ({row['工作天數']} × 480 - {row['總額外任務時間(分鐘)']} = {row['可用時間(分鐘)']}) = {row['檢驗負載指數']}", 
            axis=1
        )
        
        # 排序
        inspection_load_details = inspection_load_details.sort_values('檢驗負載指數', ascending=False)
        
        # 顯示數據表格
        st.dataframe(inspection_load_details, use_container_width=True)
        
        # 添加計算邏輯說明
        st.markdown("""
        **檢驗負載指數計算邏輯說明：**
        - 總工作時間 = 工作天數 × 480分鐘 (8小時工作日)
        - 可用時間 = 總工作時間 - 總額外任務時間
        - 檢驗負載指數 = 總檢驗時間 ÷ 可用時間
        
        當檢驗負載指數 > 1 時，表示檢驗工作量已經超過了扣除額外任務後的可用時間。
        """)

# 更通用的辅助函数，用于确保颜色刻度中的所有数值都是Python内置float类型
def safe_color_scale(scale_points, max_value):
    """
    确保颜色刻度数组中的所有数值都是Python内置float类型
    
    参数:
    scale_points - 颜色刻度点列表，例如 [[0, "red"], [0.5, "blue"], [1, "green"]]
    max_value - 用于归一化的最大值
    
    返回:
    安全的颜色刻度点列表
    """
    safe_scale = []
    
    for point in scale_points:
        if isinstance(point, list) and len(point) == 2:
            # 如果第一个元素是数值，确保它是Python float类型
            value = point[0]
            if isinstance(value, (int, float)) or (hasattr(value, 'dtype') and np.issubdtype(value.dtype, np.number)):
                if value != 0 and value != 1:
                    # 如果值是相对于max_value的分数，确保除法结果是Python float
                    if '/' in str(value):
                        # 这是一个表达式，需要计算
                        parts = str(value).split('/')
                        if len(parts) == 2:
                            num = float(parts[0])
                            denom = float(parts[1])
                            if denom == 0:
                                value = 0.0  # 避免除以零
                            else:
                                value = float(num / denom)
                        else:
                            value = float(value)
                    else:
                        # 直接转换为float
                        value = float(value)
            
            safe_scale.append([value, point[1]])
        else:
            # 如果格式不对，保留原样
            safe_scale.append(point)
    
    return safe_scale

def render_inspector_mrb_rate_chart(mrb_rate_df):
    """
    使用Altair渲染IQC人員MRB率圖表，確保懸浮提示正確顯示
    """
    import altair as alt
    
    # 確保MRB率是浮點數
    mrb_rate_df = mrb_rate_df.copy()
    mrb_rate_df['MRB率'] = mrb_rate_df['MRB率'].astype(float)
    
    # 添加格式化後的MRB率文本列，用於顯示在條形上
    mrb_rate_df['MRB率顯示'] = mrb_rate_df['MRB率'].apply(lambda x: f"{x:.1%}")
    
    # 設置顏色範圍
    # 注意：Altair需要設置一個離散的顏色映射，我們將使用一個條件子句
    color_scale = alt.Color('MRB率:Q',
                           scale=alt.Scale(domain=[0, 0.05, 0.1, 0.2, 0.3, 1],
                                          range=['#DCDCDC', '#FFCCCC', '#FF6666', '#CC0000', '#800000', '#400000']))
    
    # 創建基本柱狀圖
    base = alt.Chart(mrb_rate_df).encode(
        x=alt.X('檢驗員:N', sort='-y', title='檢驗員'),
        y=alt.Y('MRB率:Q', axis=alt.Axis(format='%'), title='MRB率'),
        color=color_scale,
        tooltip=[
            alt.Tooltip('檢驗員:N', title='檢驗員'),
            alt.Tooltip('MRB率:Q', title='MRB率', format='.2%'),
            alt.Tooltip('MRB數量:Q', title='MRB數量'),
            alt.Tooltip('總檢驗筆數:Q', title='總檢驗筆數')
        ]
    )
    
    # 創建柱狀圖
    bars = base.mark_bar()
    
    # 添加百分比標籤
    text = base.mark_text(
        align='center',
        baseline='bottom',
        dy=-5,  # 向上偏移
        color='black'
    ).encode(
        text='MRB率顯示:N'
    )
    
    # 組合圖表
    chart = (bars + text).properties(
        title='IQC人員MRB率 - MRB數量/總檢驗筆數',
        width=600,
        height=400
    ).configure_axis(
        labelFontSize=12,
        titleFontSize=14
    ).configure_title(
        fontSize=16,
        anchor='start'
    )
    
    return chart

def render_category_mrb_rate_chart(cat_mrb_rate_df):
    """
    使用Altair渲染物料類別MRB率圖表，確保懸浮提示正確顯示
    """
    import altair as alt
    
    # 確保MRB率是浮點數
    cat_mrb_rate_df = cat_mrb_rate_df.copy()
    cat_mrb_rate_df['MRB率'] = cat_mrb_rate_df['MRB率'].astype(float)
    
    # 添加格式化後的MRB率文本列
    cat_mrb_rate_df['MRB率顯示'] = cat_mrb_rate_df['MRB率'].apply(lambda x: f"{x:.1%}")
    
    # 設置顏色範圍 - 針對物料類別可能有更高的MRB率
    color_scale = alt.Color('MRB率:Q',
                           scale=alt.Scale(domain=[0, 0.1, 0.2, 0.3, 0.5, 1],
                                          range=['#DCDCDC', '#FFCCCC', '#FF6666', '#CC0000', '#800000', '#400000']))
    
    # 創建基本柱狀圖
    base = alt.Chart(cat_mrb_rate_df).encode(
        x=alt.X('物料類別:N', sort='-y', title='物料類別'),
        y=alt.Y('MRB率:Q', axis=alt.Axis(format='%'), title='MRB率'),
        color=color_scale,
        tooltip=[
            alt.Tooltip('物料類別:N', title='物料類別'),
            alt.Tooltip('MRB率:Q', title='MRB率', format='.2%'),
            alt.Tooltip('MRB數量:Q', title='MRB數量'),
            alt.Tooltip('總檢驗筆數:Q', title='總檢驗筆數')
        ]
    )
    
    # 創建柱狀圖
    bars = base.mark_bar()
    
    # 添加百分比標籤
    text = base.mark_text(
        align='center',
        baseline='bottom',
        dy=-5,
        color='black'
    ).encode(
        text='MRB率顯示:N'
    )
    
    # 組合圖表
    chart = (bars + text).properties(
        title='物料類別MRB率 - MRB數量/總檢驗筆數',
        width=600,
        height=400
    ).configure_axis(
        labelFontSize=12,
        titleFontSize=14
    ).configure_title(
        fontSize=16,
        anchor='start'
    )
    
    return chart


# 4. 修改時間分配概覽，使其默認摺疊
def render_time_allocation_dashboard(time_allocation_data):
    if time_allocation_data is None or time_allocation_data.empty:
        st.error("沒有可用的時間分配數據，請確保上傳了正確的檔案格式或檢查過濾條件。")
        return
    
    # 過濾時間分配數據 - 排除指定檢驗員
    time_allocation_data = filter_excluded_inspectors(time_allocation_data, inspector_column='inspector')
    
    if time_allocation_data.empty:
        st.warning("過濾後沒有可用的時間分配數據。")
        return
    
    st.header("⚖️ IQC人員時間分配監控")
    
    # 檢查資料結構，確認有必要的欄位
    required_cols = ['inspector', 'inspection_ratio', 'additional_task_ratio', 'task_detail_ratios']
    for col in required_cols:
        if col not in time_allocation_data.columns:
            st.error(f"時間分配數據缺少必要欄位: {col}")
            return
    
    # 1. 首先顯示額外任務時間比例排名 (由高到低)
    st.subheader("額外任務時間比例排名🔖")
    
    # 計算和準備數據
    additional_task_ranking = time_allocation_data[['inspector', 'additional_task_ratio', 'total_time']].copy()
    additional_task_ranking = additional_task_ranking.sort_values('additional_task_ratio', ascending=False)
    additional_task_ranking.columns = ['檢驗員', '額外任務時間比例', '總時間(分鐘)']
    
    # 創建條形圖
    fig_ranking = px.bar(
        additional_task_ranking,
        x='檢驗員',
        y='額外任務時間比例',
        color='額外任務時間比例',
        color_continuous_scale='RdBu_r',  # 使用藍到紅的色彩方案，比例高的顯示為紅色
        title="各檢驗員額外任務時間比例排名"
    )
    
    # 設置百分比格式
    fig_ranking.update_layout(
        yaxis=dict(
            tickformat=".0%",
            title="額外任務時間比例"
        ),
        xaxis_title="檢驗員",
        height=400
    )
    
    # 添加文字標籤
    fig_ranking.update_traces(
        texttemplate='%{y:.1%}',
        textposition='outside'
    )
    
    st.plotly_chart(fig_ranking, use_container_width=True)
    
    # 2. 讓使用者選擇人員後才顯示個人時間分配比例
    st.subheader("個人時間分配比例詳情👥 ")
    
    # 取得所有檢驗員列表
    inspectors = sorted(time_allocation_data['inspector'].unique().tolist())
    
    # 選擇檢驗員 - 預設不選擇任何人
    selected_inspector = st.selectbox(
        "選擇檢驗員查看詳細時間分配",
        options=["請選擇檢驗員..."] + inspectors,
        key="time_allocation_inspector_selector"
    )
    
    # 只有當用戶選擇了一個檢驗員（不是預設選項）時才顯示詳情
    if selected_inspector != "請選擇檢驗員...":
        inspector_data = time_allocation_data[time_allocation_data['inspector'] == selected_inspector].iloc[0]
        
        # 創建餅圖數據
        pie_data = pd.DataFrame([
            {'category': '檢驗時間', 'ratio': inspector_data['inspection_ratio']},
            {'category': '額外任務時間', 'ratio': inspector_data['additional_task_ratio']}
        ])
        
        fig = px.pie(
            pie_data,
            values='ratio',
            names='category',
            title=f"{selected_inspector} 時間分配比例",
            color_discrete_sequence=px.colors.qualitative.Set3
        )
        
        fig.update_traces(
            textinfo='percent+label',
            hoverinfo='label+percent'
        )
        
        st.plotly_chart(fig, use_container_width=True)
        
        # 顯示額外任務細項分配
        st.write("額外任務細項分配：")
        
        if isinstance(inspector_data['task_detail_ratios'], dict) and inspector_data['task_detail_ratios']:
            task_details = []
            
            for task_type, ratio in inspector_data['task_detail_ratios'].items():
                task_details.append({
                    'task_type': task_type,
                    'ratio': ratio
                })
            
            task_df = pd.DataFrame(task_details)
            
            if not task_df.empty:
                # 顯示額外任務細項餅圖
                fig_detail = px.pie(
                    task_df,
                    values='ratio',
                    names='task_type',
                    title="額外任務細項分配",
                    color_discrete_sequence=px.colors.qualitative.Pastel
                )
                
                fig_detail.update_traces(
                    textinfo='percent+label',
                    hoverinfo='label+percent'
                )
                
                st.plotly_chart(fig_detail, use_container_width=True)
                
                # 顯示額外任務詳細數據
                with st.expander(f"{selected_inspector} 額外任務詳細數據"):
                    detail_df = task_df.copy()
                    detail_df['ratio'] = (detail_df['ratio'] * 100).round(1).astype(str) + '%'
                    detail_df.columns = ['任務類型', '佔總時間比例']
                    st.dataframe(detail_df, use_container_width=True)
            else:
                st.write("沒有額外任務記錄")
        else:
            st.write("沒有額外任務記錄")
    else:
        st.markdown("""
            <div style="background-color:#f0f2f5;  border-radius:4px; padding:16px; margin-top:20px; display:flex; align-items:center;">
                <div>
                    <p style="margin:0; color:#37474F; font-weight:400;">👆 請選擇一位檢驗員以查看其時間分配詳情。</p>
                </div>
            </div>
            """, unsafe_allow_html=True)
    
    # 3. 將所有檢驗員時間分配概覽放在最後顯示，且默認摺疊
    st.subheader("所有檢驗員時間分配概覽👥 ")
    
    # 使用expander且默認為collapsed (expanded=False)
    with st.expander("點擊展開查看所有檢驗員時間分配概覽", expanded=False):
        overview_data = []
        for _, row in time_allocation_data.iterrows():
            overview_data.append({
                '檢驗員': row['inspector'],
                '檢驗時間比例': row['inspection_ratio'],
                '額外任務時間比例': row['additional_task_ratio'],
                '總時間(分鐘)': row['total_time']
            })
        
        overview_df = pd.DataFrame(overview_data)
        overview_df['檢驗時間比例'] = (overview_df['檢驗時間比例'] * 100).round(1).astype(str) + '%'
        overview_df['額外任務時間比例'] = (overview_df['額外任務時間比例'] * 100).round(1).astype(str) + '%'
        overview_df['總時間(分鐘)'] = overview_df['總時間(分鐘)'].round(0).astype(int)
        
        # 依額外任務時間比例排序
        overview_df['排序值'] = overview_df['額外任務時間比例'].str.rstrip('%').astype(float)
        overview_df = overview_df.sort_values('排序值', ascending=False)
        overview_df = overview_df.drop(columns=['排序值'])
        
        st.dataframe(overview_df, use_container_width=True)

# 2. 修改額外任務監控儀表板 - 預設不選擇檢驗員 & 調整平均時間計算方式
def render_additional_tasks_dashboard(additional_tasks_monitor_data):
    if additional_tasks_monitor_data is None or additional_tasks_monitor_data.empty:
        st.error("沒有可用的額外任務數據，請確保上傳了正確的檔案格式或檢查過濾條件。")
        return
    
    # 過濾額外任務數據 - 排除指定檢驗員
    additional_tasks_monitor_data = filter_excluded_inspectors(additional_tasks_monitor_data, inspector_column='inspector')
    
    if additional_tasks_monitor_data.empty:
        st.warning("過濾後沒有可用的額外任務數據。")
        return

    st.header("📋IQC額外任務監控")
    
    # 顯示所有額外任務概覽
    st.subheader("所有額外任務時間統計📝")
    
    # 按任務類型分組統計
    task_summary = additional_tasks_monitor_data.groupby('task_type')['total_time'].sum().reset_index()
    task_summary.columns = ['任務類型', '總時間(分鐘)']
    task_summary = task_summary.sort_values('總時間(分鐘)', ascending=False)
    
    fig_task_summary = px.bar(
        task_summary,
        x='任務類型',
        y='總時間(分鐘)',
        title="額外任務類型時間統計",
        color='總時間(分鐘)'
    )
    
    fig_task_summary.update_layout(
        xaxis_title="任務類型",
        yaxis_title="總時間(分鐘)",
        height=400
    )
    
    st.plotly_chart(fig_task_summary, use_container_width=True)
    
    # 需要匯總每個檢驗員每個任務的實際發生天數
    # 首先將數據轉換為DataFrame，確保有日期欄位
    raw_data = additional_tasks_monitor_data.copy()
    
    # 按檢驗員和任務類型聚合，計算資料列數量（即該任務的工作天數）和總時間
    try:
        # 獲取日期欄位 - 首先嘗試直接從原始資料中獲取
        date_column = None
        if 'date' in raw_data.columns:
            date_column = 'date'
        elif '日期' in raw_data.columns:
            date_column = '日期'
        
        # 如果沒有找到日期欄位，使用檔案來源分組
        if date_column is not None:
            # 確保日期欄位是datetime類型
            raw_data[date_column] = pd.to_datetime(raw_data[date_column], errors='coerce')
            
            # 統計每個檢驗員每種任務的實際工作天數和總時間
            task_days = raw_data.groupby(['inspector', 'task_type', date_column])['total_time'].sum().reset_index()
            task_stats = task_days.groupby(['inspector', 'task_type']).agg(
                task_days=pd.NamedAgg(column=date_column, aggfunc='count'),
                total_time=pd.NamedAgg(column='total_time', aggfunc='sum')
            ).reset_index()
        else:
            # 如果沒有日期欄位，使用檔案來源作為替代分組依據
            if '檔案來源' in raw_data.columns:
                task_days = raw_data.groupby(['inspector', 'task_type', '檔案來源'])['total_time'].sum().reset_index()
                task_stats = task_days.groupby(['inspector', 'task_type']).agg(
                    task_days=pd.NamedAgg(column='檔案來源', aggfunc='count'),
                    total_time=pd.NamedAgg(column='total_time', aggfunc='sum')
                ).reset_index()
            else:
                # 如果無法通過日期或檔案來源分組，則每個任務假設為1天
                task_stats = raw_data.groupby(['inspector', 'task_type']).agg(
                    total_time=pd.NamedAgg(column='total_time', aggfunc='sum')
                ).reset_index()
                task_stats['task_days'] = 1  # 假設每個任務只有1天
        
        # 計算每個任務在實際工作天的平均時間
        task_stats['每次平均時間(分鐘)'] = (task_stats['total_time'] / task_stats['task_days']).round(1)
                
        # 1. 提供檢驗員選擇介面 - 預設不選擇任何人員
        inspectors = sorted(additional_tasks_monitor_data['inspector'].unique())
        selected_inspector = st.selectbox(
            "選擇檢驗員查看詳細額外任務時間分配",
            options=["請選擇檢驗員..."] + list(inspectors),
            key="additional_tasks_inspector_selector"
        )
        
        # 顯示選定檢驗員的額外任務時間分配
        if selected_inspector != "請選擇檢驗員...":
            st.subheader(f"{selected_inspector} 額外任務時間分配")
            
            # 篩選該檢驗員的數據
            inspector_data = task_stats[task_stats['inspector'] == selected_inspector].copy()
            
            # 按每次平均時間排序
            inspector_data = inspector_data.sort_values('每次平均時間(分鐘)', ascending=False)
            
            # 2. 將額外任務時間換算為每次平均分鐘並依此排序顯示
            if not inspector_data.empty:
                fig = px.bar(
                    inspector_data,
                    x='task_type',
                    y='每次平均時間(分鐘)',
                    labels={'task_type': '任務類型', '每次平均時間(分鐘)': '每次平均時間 (分鐘)'},
                    color='每次平均時間(分鐘)',
                    title=f"{selected_inspector} 額外任務每次平均時間分配"
                )
                
                fig.update_layout(
                    xaxis_title="任務類型",
                    yaxis_title="每次平均時間 (分鐘)",
                    height=400,
                    showlegend=False
                )
                
                # 添加文字標籤
                fig.update_traces(
                    texttemplate='%{y:.1f}',
                    textposition='outside'
                )
                
                st.plotly_chart(fig, use_container_width=True)
                
                # 顯示詳細數據
                with st.expander(f"檢視 {selected_inspector} 額外任務詳細數據"):
                    detail_df = inspector_data[['task_type', 'total_time', 'task_days', '每次平均時間(分鐘)']].copy()
                    detail_df.columns = ['任務類型', '總時間(分鐘)', '任務發生天數', '每次平均時間(分鐘)']
                    st.dataframe(detail_df, use_container_width=True)
            else:
                st.info(f"{selected_inspector} 沒有額外任務記錄")
        else:
            st.markdown("""
            <div style="background-color:#f0f2f5; border-radius:4px; padding:16px; margin-top:20px; display:flex; align-items:center;">
                <div>
                    <p style="margin:0; color:#37474F; font-weight:400;">👆 請選擇一位檢驗員以查看其額外任務時間分配。</p>
                </div>
            </div>
            """, unsafe_allow_html=True)
        
        # 顯示所有檢驗員的額外任務平均時間統計（按每次平均時間排序）
        st.subheader("所有檢驗員每次額外任務時間統計📝")
        
        # 按檢驗員分組
        inspector_summary = task_stats.groupby('inspector').agg(
            total_time=pd.NamedAgg(column='total_time', aggfunc='sum'),
            avg_task_time=pd.NamedAgg(column='每次平均時間(分鐘)', aggfunc=lambda x: (x * task_stats.loc[x.index, 'task_days']).sum() / task_stats.loc[x.index, 'task_days'].sum())
        ).reset_index()
        
        inspector_summary.columns = ['檢驗員', '總時間(分鐘)', '加權每次平均時間(分鐘)']
        inspector_summary['加權每次平均時間(分鐘)'] = inspector_summary['加權每次平均時間(分鐘)'].round(1)
        inspector_summary = inspector_summary.sort_values('加權每次平均時間(分鐘)', ascending=False)
        
        # 創建按每次平均時間排序的條形圖
        fig_inspector_avg = px.bar(
            inspector_summary,
            x='檢驗員',
            y='加權每次平均時間(分鐘)',
            title="各檢驗員額外任務每次平均時間",
            color='加權每次平均時間(分鐘)',
            color_continuous_scale='Viridis'
        )
        
        fig_inspector_avg.update_layout(
            xaxis_title="檢驗員",
            yaxis_title="每次平均時間 (分鐘)",
            height=400
        )
        
        # 添加文字標籤
        fig_inspector_avg.update_traces(
            texttemplate='%{y:.1f}',
            textposition='outside'
        )
        
        st.plotly_chart(fig_inspector_avg, use_container_width=True)
        
        # 顯示詳細數據表格
        with st.expander("查看所有檢驗員額外任務時間詳細數據"):
            st.dataframe(inspector_summary, use_container_width=True)
            
    except Exception as e:
        st.error(f"計算每次平均時間時發生錯誤: {str(e)}")
        st.write("使用原始數據顯示：")
        
        # 如果有錯誤，退回到簡單的統計
        inspector_task_summary = additional_tasks_monitor_data.groupby(
            ['inspector', 'task_type'])['total_time'].sum().reset_index()
        
        # 計算每個檢驗員的總任務時間
        inspector_summary = inspector_task_summary.groupby('inspector')['total_time'].sum().reset_index()
        inspector_summary.columns = ['檢驗員', '總時間(分鐘)']
        inspector_summary = inspector_summary.sort_values('總時間(分鐘)', ascending=False)
        
        # 顯示簡單統計
        st.dataframe(inspector_summary, use_container_width=True)

# 顯示處理狀態
def render_status():
    if st.session_state.processing_error:
        st.error(f"錯誤: {st.session_state.processing_error}")
    
# 顯示調試信息
# def render_debug_info():
    #if DEBUG_MODE and 'debug_info' in st.session_state:
        #with st.expander("調試信息", expanded=False):
            #if 'logs' in st.session_state.debug_info:
                #logs = st.session_state.debug_info['logs']
                #for log in logs[-50:]:  # 只顯示最後50條日誌
                    #st.text(log)

def render_calculation_details():
    """顯示所有計算詳情，方便檢查計算過程是否正確"""
    st.header("計算詳情檢查")
    st.markdown("此頁面顯示所有計算中間結果，方便檢查計算過程是否正確。")
    
    if not st.session_state.files_uploaded:
        st.info("請先上傳並處理檔案")
        return
    
    tabs = st.tabs([
        "原始處理資料",
        "PCB標準工時計算",
        "檢驗效率計算",
        "工作負載計算",
        "時間分配計算",
        "額外任務統計",
        "MRB率分析",  # 新增MRB率分析頁籤
    ])
    
    # 原始處理資料
    with tabs[0]:
        st.subheader("處理後的原始資料")
        if st.session_state.processed_data is not None and not st.session_state.processed_data.empty:
            # 新增篩選功能
            inspector_filter = st.selectbox(
                "選擇檢驗員",
                options=["全部"] + sorted(st.session_state.processed_data["處理後檢驗員"].unique().tolist()),
                index=0
            )
            
            category_filter = st.selectbox(
                "選擇物料類別",
                options=["全部"] + sorted(st.session_state.processed_data["類別"].unique().tolist()),
                index=0
            )
            
            df_filtered = st.session_state.processed_data.copy()
            
            if inspector_filter != "全部":
                df_filtered = df_filtered[df_filtered["處理後檢驗員"] == inspector_filter]
            
            if category_filter != "全部":
                df_filtered = df_filtered[df_filtered["類別"] == category_filter]
            
            # 顯示關鍵欄位
            key_columns = [
                '處理後檢驗員', '料號', '類別', '抽樣數量', '檢驗日期',
                '處理後檢驗標準工時', '檢驗耗時', '效率比值', '抽樣狀態'
            ]
            
            # 確保所有列都存在
            existing_columns = [col for col in key_columns if col in df_filtered.columns]
            
            st.dataframe(df_filtered[existing_columns])
            
            st.metric("總筆數", len(df_filtered))
            
            if len(df_filtered) > 0:
                st.metric("平均效率比值", round(df_filtered["效率比值"].mean(), 2))
            
        else:
            st.info("沒有可用的處理資料")
    
    # PCB標準工時計算
    with tabs[1]:
        st.subheader("PCB標準工時計算詳情")
        if st.session_state.processed_data is not None and not st.session_state.processed_data.empty:
            # 只顯示QB類型的料號
            qb_data = st.session_state.processed_data[st.session_state.processed_data["類別"] == "QB"].copy()
            
            if not qb_data.empty:
                # 顯示PCB標準工時計算關鍵欄位
                pcb_columns = [
                    '料號', '抽樣數量', '處理後檢驗標準工時', '檢驗耗時', '效率比值'
                ]
                
                # 確保所有列都存在
                existing_pcb_columns = [col for col in pcb_columns if col in qb_data.columns]
                
                st.dataframe(qb_data[existing_pcb_columns])
                
                if st.session_state.pcb_spec_data is not None and not st.session_state.pcb_spec_data.empty:
                    st.subheader("PCB建檔明細")
                    st.dataframe(st.session_state.pcb_spec_data.head(20))
                
                if st.session_state.pcb_standard_time_data is not None and not st.session_state.pcb_standard_time_data.empty:
                    st.subheader("PCB標準工時對應表")
                    st.dataframe(st.session_state.pcb_standard_time_data)
            else:
                st.info("沒有發現QB類型料號")
        else:
            st.info("沒有可用的處理資料")
    
    # 檢驗效率計算
    with tabs[2]:
        st.subheader("檢驗效率計算詳情")
        
        if 'efficiency_data' in st.session_state:
            # 總效率數據
            st.subheader("總效率排名")
            if 'overall_efficiency_ranking' in st.session_state.efficiency_data:
                overall_eff = st.session_state.efficiency_data['overall_efficiency_ranking'].copy()
                if not overall_eff.empty:
                    # 重命名欄位
                    overall_eff = overall_eff.rename(columns={
                        'inspector': '檢驗員',
                        'efficiency': '效率指標',
                        'total_standard_time': '標準工時總和(分鐘)',
                        'total_actual_time': '實際耗時總和(分鐘)',
                        'record_count': '記錄筆數'
                    })
                    st.dataframe(overall_eff)
            
            # 類別效率數據
            st.subheader("各物料類別效率")
            if 'category_efficiency_data' in st.session_state.efficiency_data:
                category_data = st.session_state.efficiency_data['category_efficiency_data']
                
                categories = list(category_data.keys())
                if categories:
                    selected_category = st.selectbox("選擇類別", options=categories)
                    
                    if selected_category and selected_category in category_data:
                        cat_data = category_data[selected_category]
                        if cat_data:
                            cat_df = pd.DataFrame(cat_data)
                            # 重命名欄位
                            cat_df = cat_df.rename(columns={
                                'inspector': '檢驗員',
                                'efficiency': '效率指標',
                                'record_count': '記錄筆數',
                                'total_standard_time': '標準工時總和(分鐘)',
                                'total_actual_time': '實際耗時總和(分鐘)'
                            })
                            st.dataframe(cat_df)
                else:
                    st.info("沒有可用的類別效率數據")
        else:
            st.info("沒有可用的效率計算數據")
    
    # 工作負載計算
    with tabs[3]:
        st.subheader("工作負載計算詳情")
        
        if 'workload_data' in st.session_state and not st.session_state.workload_data.empty:
            workload_df = st.session_state.workload_data.copy()
            
            # 重命名欄位
            workload_df = workload_df.rename(columns={
                'date': '日期',
                'inspector': '檢驗員',
                'inspection_time': '檢驗時間(分鐘)',
                'additional_task_time': '額外任務時間(分鐘)',
                'total_time': '總時間(分鐘)',
                'workload_index': '工作負載指數',
                'work_period': '工作時段',
                'inspection_count': '檢驗次數'
            })
            
            # 檢驗員篩選
            inspectors = ["全部"] + sorted(workload_df["檢驗員"].unique().tolist())
            selected_inspector = st.selectbox("選擇檢驗員", options=inspectors, key="workload_inspector")
            
            filtered_workload = workload_df
            if selected_inspector != "全部":
                filtered_workload = workload_df[workload_df["檢驗員"] == selected_inspector]
            
            st.dataframe(filtered_workload)
            
            # 計算並顯示統計資訊
            if not filtered_workload.empty:
                avg_workload = filtered_workload["工作負載指數"].mean()
                max_workload = filtered_workload["工作負載指數"].max()
                
                col1, col2 = st.columns(2)
                with col1:
                    st.metric("平均工作負載指數", round(avg_workload, 2))
                with col2:
                    st.metric("最高工作負載指數", round(max_workload, 2))
        else:
            st.info("沒有可用的工作負載數據")
    
    # 時間分配計算
    with tabs[4]:
        st.subheader("時間分配計算詳情")
        
        if 'time_allocation_data' in st.session_state and not st.session_state.time_allocation_data.empty:
            time_allocation_df = st.session_state.time_allocation_data.copy()
            
            # 這個數據框架含有字典類型的列，需要特殊處理
            if 'task_detail_ratios' in time_allocation_df.columns:
                # 顯示原始數據（含字典）
                st.subheader("原始時間分配數據")
                
                # 將字典列轉換為字符串以顯示
                time_allocation_df['task_detail_ratios'] = time_allocation_df['task_detail_ratios'].apply(
                    lambda x: str(x) if isinstance(x, dict) else str(x)
                )
                
                # 重命名欄位
                renamed_df = time_allocation_df.rename(columns={
                    'inspector': '檢驗員',
                    'inspection_time': '檢驗時間(分鐘)',
                    'additional_task_time': '額外任務時間(分鐘)',
                    'total_time': '總時間(分鐘)',
                    'inspection_ratio': '檢驗時間比例',
                    'additional_task_ratio': '額外任務時間比例',
                    'task_detail_ratios': '任務細項比例(字典)'
                })
                
                st.dataframe(renamed_df)
                
                # 展開任務細項數據
                st.subheader("展開後的任務細項數據")
                expanded_df = time_allocation_df.copy()
                
                # 為每個檢驗員創建展開的任務細項
                all_task_types = set()
                for idx, row in time_allocation_df.iterrows():
                    if isinstance(row['task_detail_ratios'], str):
                        try:
                            # 嘗試將字符串轉回字典
                            task_dict = eval(row['task_detail_ratios'])
                            if isinstance(task_dict, dict):
                                for task_type, ratio in task_dict.items():
                                    all_task_types.add(task_type)
                                    expanded_df.at[idx, f'任務比例_{task_type}'] = ratio
                        except:
                            pass
                
                # 重命名基本欄位
                base_columns = {
                    'inspector': '檢驗員',
                    'inspection_time': '檢驗時間(分鐘)',
                    'additional_task_time': '額外任務時間(分鐘)',
                    'total_time': '總時間(分鐘)',
                    'inspection_ratio': '檢驗時間比例',
                    'additional_task_ratio': '額外任務時間比例'
                }
                
                # 只選擇存在的列進行重命名
                valid_columns = [col for col in base_columns.keys() if col in expanded_df.columns]
                for col in valid_columns:
                    expanded_df.rename(columns={col: base_columns[col]}, inplace=True)
                
                # 刪除原始字典列
                if 'task_detail_ratios' in expanded_df.columns:
                    expanded_df = expanded_df.drop(columns=['task_detail_ratios'])
                
                task_columns = [col for col in expanded_df.columns if col.startswith('任務比例_')]
                display_columns = ['檢驗員', '檢驗時間(分鐘)', '額外任務時間(分鐘)', 
                                   '總時間(分鐘)', '檢驗時間比例', '額外任務時間比例'] + task_columns
                
                # 只選擇存在的列
                valid_display_columns = [col for col in display_columns if col in expanded_df.columns]
                st.dataframe(expanded_df[valid_display_columns])
            else:
                # 如果沒有任務細項，直接顯示基本時間分配
                renamed_df = time_allocation_df.rename(columns={
                    'inspector': '檢驗員',
                    'inspection_time': '檢驗時間(分鐘)',
                    'additional_task_time': '額外任務時間(分鐘)',
                    'total_time': '總時間(分鐘)',
                    'inspection_ratio': '檢驗時間比例',
                    'additional_task_ratio': '額外任務時間比例'
                })
                
                st.dataframe(renamed_df)
        else:
            st.info("沒有可用的時間分配數據")
    
    # 額外任務統計
    with tabs[5]:
        st.subheader("額外任務統計詳情")
        
        if 'additional_tasks_monitor_data' in st.session_state and not st.session_state.additional_tasks_monitor_data.empty:
            tasks_df = st.session_state.additional_tasks_monitor_data.copy()
            
            # 重命名欄位
            tasks_df = tasks_df.rename(columns={
                'inspector': '檢驗員',
                'task_type': '任務類型',
                'total_time': '總時間(分鐘)'
            })
            
            # 檢驗員篩選
            inspectors = ["全部"] + sorted(tasks_df["檢驗員"].unique().tolist())
            selected_inspector = st.selectbox("選擇檢驗員", options=inspectors, key="tasks_inspector")
            
            filtered_tasks = tasks_df
            if selected_inspector != "全部":
                filtered_tasks = tasks_df[tasks_df["檢驗員"] == selected_inspector]
            
            st.dataframe(filtered_tasks)
            
            # 按任務類型分組統計
            task_summary = tasks_df.groupby('任務類型')['總時間(分鐘)'].sum().reset_index()
            task_summary = task_summary.sort_values('總時間(分鐘)', ascending=False)
            
            st.subheader("任務類型統計")
            st.dataframe(task_summary)
            
            # 任務分布圖
            st.subheader("任務時間分布")
            fig = px.pie(
                task_summary,
                values='總時間(分鐘)',
                names='任務類型',
                title="額外任務時間分布"
            )
            st.plotly_chart(fig, use_container_width=True)
            
            # 按檢驗員分組統計
            inspector_summary = tasks_df.groupby('檢驗員')['總時間(分鐘)'].sum().reset_index()
            inspector_summary = inspector_summary.sort_values('總時間(分鐘)', ascending=False)
            
            st.subheader("檢驗員額外任務時間統計")
            st.dataframe(inspector_summary)
        else:
            st.info("沒有可用的額外任務數據")

    # 添加MRB率分析頁籤
    with tabs[6]:
        st.subheader("MRB率分析")
        
        if 'processed_data' not in st.session_state or st.session_state.processed_data is None:
            st.error("沒有處理後的數據可供分析")
            return
        
        df = st.session_state.processed_data
        
        # 確保有"是否為MRB"欄位
        if '是否為MRB' not in df.columns:
            st.error("數據中缺少'是否為MRB'欄位，無法分析MRB率")
            return
        
        # 計算總體MRB率
        total_count = len(df)
        mrb_count = df['是否為MRB'].sum() if df['是否為MRB'].dtype != 'object' else df['是否為MRB'].apply(
            lambda x: True if str(x).upper() == 'TRUE' or x == 'MRB' else False
        ).sum()
        
        overall_mrb_rate = mrb_count / total_count if total_count > 0 else 0
        
        # 顯示總體MRB率
        st.metric("總體MRB率", f"{overall_mrb_rate:.2%}", f"{mrb_count} MRB / {total_count} 總筆數")
        
        # 按檢驗員分組計算MRB率
        mrb_rate_df = calculate_mrb_rates(df)
        
        # 顯示每個檢驗員的MRB率
        st.subheader("各檢驗員MRB率")
        
        # 格式化MRB率為百分比
        display_df = mrb_rate_df.copy()
        display_df['MRB率'] = display_df['MRB率'].apply(lambda x: f"{x:.2%}")
        
        # 添加色彩標記，高MRB率顯示為紅色
        def highlight_high_mrb_rate(s):
            if s.name == 'MRB率':
                return ['background-color: rgba(255,50,50,0.2)' if float(x.strip('%'))/100 > overall_mrb_rate * 1.5 else '' for x in s]
            return [''] * len(s)
        
        # 應用樣式並顯示
        styled_df = display_df.style.apply(highlight_high_mrb_rate)
        st.dataframe(styled_df, use_container_width=True)
        
        # 按物料類別分析MRB率
        st.subheader("按物料類別分析MRB率")
        
        if '類別' in df.columns:
            category_stats = []
            
            for category, group in df.groupby('類別'):
                cat_total = len(group)
                cat_mrb = group['是否為MRB'].sum() if group['是否為MRB'].dtype != 'object' else group['是否為MRB'].apply(
                    lambda x: True if str(x).upper() == 'TRUE' or x == 'MRB' else False
                ).sum()
                cat_mrb_rate = cat_mrb / cat_total if cat_total > 0 else 0
                
                category_stats.append({
                    '物料類別': category,
                    'MRB數量': cat_mrb,
                    '總檢驗筆數': cat_total,
                    'MRB率': cat_mrb_rate
                })
            
            cat_df = pd.DataFrame(category_stats)
            cat_df = cat_df.sort_values('MRB率', ascending=False)
            
            # 格式化MRB率為百分比
            cat_df['MRB率'] = cat_df['MRB率'].apply(lambda x: f"{x:.2%}")
            
            st.dataframe(cat_df, use_container_width=True)
            
            # 創建物料類別MRB率條形圖
            cat_df_for_chart = pd.DataFrame(category_stats)
            fig = px.bar(
                cat_df_for_chart.sort_values('MRB率', ascending=False),
                x='物料類別',
                y='MRB率',
                text=cat_df_for_chart['MRB率'].apply(lambda x: f"{x:.1%}"),
                labels={'物料類別': '物料類別', 'MRB率': 'MRB率 (%)'},
                title="物料類別MRB率",
                color='MRB率',
                color_continuous_scale="Reds"
            )
            
            fig.update_layout(yaxis_tickformat=".0%")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("數據中缺少'類別'欄位，無法按物料類別分析MRB率")

def export_processed_data():
    """
    匯出處理後的所有相關資料，並確保中文正確顯示和MRB狀態正確顯示
    """
    try:
        # 檢查是否有可用數據
        if st.session_state.processed_data is None:
            st.error("沒有可用的處理數據，請先上傳和處理文件")
            return

        # 创建匯出目錄
        import os
        export_dir = "iqc_export_data"
        if not os.path.exists(export_dir):
            os.makedirs(export_dir)
        
        # 1. 處理後的原始資料並修正MRB狀態
        processed_df = st.session_state.processed_data.copy()
        
        # 修正MRB狀態 - 使用MRB加時判斷
        processed_df['是否為MRB'] = processed_df['MRB加時'] > 0
        processed_df['是否為MRB'] = processed_df['是否為MRB'].apply(lambda x: "TRUE" if x else "FALSE")
        processed_df['MRB狀態'] = processed_df['MRB加時'].apply(lambda x: "MRB" if x > 0 else "Normal inspection")
        
        # 確保MRB訊息也與狀態一致
        processed_df.loc[processed_df['MRB加時'] > 0, 'MRB訊息'] = "有MRB標記"
        
        # 顯示MRB狀態統計供測試
        mrb_true_count = (processed_df['是否為MRB'] == "TRUE").sum()
        mrb_status_mrb_count = (processed_df['MRB狀態'] == "MRB").sum()
        mrb_time_count = (processed_df['MRB加時'] > 0).sum()
        
        debug_log(f"匯出前MRB狀態檢查: 是否為MRB=TRUE: {mrb_true_count}, MRB狀態=MRB: {mrb_status_mrb_count}, MRB加時>0: {mrb_time_count}", level="INFO")
        
        # 将布尔值转换为文字显示
        processed_df['MRB狀態'] = processed_df['MRB加時'].apply(
            lambda x: "MRB" if x > 0 else "Normal inspection"
        )
        
        # 确保所有必要列字段都存在
        essential_columns = [
            '處理後檢驗員', '料號', '類別', '抽樣數量', '檢驗日期',
            '處理後檢驗標準工時', '檢驗耗時', '效率比值', '抽樣狀態', 
            '檢驗開始時間', '是否為MRB', 'MRB狀態', 'MRB訊息'
        ]
        
        # 添加PCB相關欄位，如果存在的話
        pcb_columns = ['基礎標準工時', 'MRB加時']
        for col in pcb_columns:
            if col in processed_df.columns:
                essential_columns.append(col)
        
        # 如果有M列字段，也保留它以便参考
        if 'M' in processed_df.columns:
            essential_columns.append('M')
        
        # 确保所有必要列字段都存在，如果不存在则跳过
        valid_columns = [col for col in essential_columns if col in processed_df.columns]
        filtered_df = processed_df[valid_columns].copy()

        # 2. 效率数据
        efficiency_df = pd.DataFrame()
        if 'efficiency_data' in st.session_state and 'overall_efficiency_ranking' in st.session_state.efficiency_data:
            efficiency_df = st.session_state.efficiency_data['overall_efficiency_ranking'].copy()
            # 重命名列字段，使其更易理解
            if 'inspector' in efficiency_df.columns:
                efficiency_df.rename(columns={
                    'inspector': '檢驗員',
                    'efficiency': '效率指標',
                    'total_standard_time': '標準工時總和(分鐘)',
                    'total_actual_time': '實際耗時總和(分鐘)',
                    'record_count': '記錄筆數'
                }, inplace=True)

        # 3. 物料类别效率数据
        category_efficiency_dfs = {}
        if 'efficiency_data' in st.session_state and 'category_efficiency_data' in st.session_state.efficiency_data:
            category_data = st.session_state.efficiency_data['category_efficiency_data']
            for category, data in category_data.items():
                if data:  # 确保有数据
                    cat_df = pd.DataFrame(data)
                    # 重命名列字段
                    if 'inspector' in cat_df.columns:
                        cat_df.rename(columns={
                            'inspector': '檢驗員',
                            'efficiency': '效率指標',
                            'record_count': '記錄筆數',
                            'total_standard_time': '標準工時總和(分鐘)',
                            'total_actual_time': '實際耗時總和(分鐘)'
                        }, inplace=True)
                    category_efficiency_dfs[category] = cat_df

        # 4. 工作负载数据
        workload_df = pd.DataFrame()
        if 'workload_data' in st.session_state and not st.session_state.workload_data.empty:
            workload_df = st.session_state.workload_data.copy()
            # 重命名列字段
            workload_df.rename(columns={
                'date': '日期',
                'inspector': '檢驗員',
                'inspection_time': '檢驗時間(分鐘)',
                'additional_task_time': '額外任務時間(分鐘)',
                'total_time': '總時間(分鐘)',
                'workload_index': '工作負載指數',
                'work_period': '工作時段',
                'inspection_count': '檢驗次數'
            }, inplace=True)

        # 5. 时间分配数据
        time_allocation_df = pd.DataFrame()
        if 'time_allocation_data' in st.session_state and not st.session_state.time_allocation_data.empty:
            time_allocation_df = st.session_state.time_allocation_data.copy()
            # 这个数据框架含有字典类型的列，需要特殊处理
            if 'task_detail_ratios' in time_allocation_df.columns:
                # 将字典类型的列展开成多个列
                for idx, row in time_allocation_df.iterrows():
                    if isinstance(row['task_detail_ratios'], dict):
                        for task_type, ratio in row['task_detail_ratios'].items():
                            col_name = f'任務__{task_type}'
                            time_allocation_df.at[idx, col_name] = ratio
                
                # 删除原始字典列
                time_allocation_df = time_allocation_df.drop(columns=['task_detail_ratios'])
            
            # 重命名列字段
            new_columns = {
                'inspector': '檢驗員',
                'inspection_time': '檢驗時間(分鐘)',
                'additional_task_time': '額外任務時間(分鐘)',
                'total_time': '總時間(分鐘)',
                'inspection_ratio': '檢驗時間比例',
                'additional_task_ratio': '額外任務時間比例'
            }
            # 只重命名存在的列
            valid_renames = {k: v for k, v in new_columns.items() if k in time_allocation_df.columns}
            time_allocation_df.rename(columns=valid_renames, inplace=True)

        # 6. 额外任务数据
        additional_tasks_df = pd.DataFrame()
        if 'additional_tasks_monitor_data' in st.session_state and not st.session_state.additional_tasks_monitor_data.empty:
            additional_tasks_df = st.session_state.additional_tasks_monitor_data.copy()
            # 重命名列字段
            additional_tasks_df.rename(columns={
                'inspector': '檢驗員',
                'task_type': '任務類型',
                'total_time': '總時間(分鐘)'
            }, inplace=True)

        # 创建一个Excel文件，包含所有数据表
        excel_buffer = io.BytesIO()
        
        with pd.ExcelWriter(excel_buffer, engine='xlsxwriter', engine_kwargs={'options': {'encoding': 'utf-8'}}) as writer:
            # 1. 处理后的原始数据
            filtered_df.to_excel(writer, sheet_name='處理後資料', index=False)
            
            # 2. 效率数据
            if not efficiency_df.empty:
                efficiency_df.to_excel(writer, sheet_name='整體效率排名', index=False)
            
            # 3. 物料类别效率数据
            for category, df in category_efficiency_dfs.items():
                # 确保工作表名称有效（最多31个字符）
                sheet_name = f"類別效率_{category}"
                if len(sheet_name) > 31:
                    sheet_name = sheet_name[:28] + "..."
                df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            # 4. 工作负载数据
            if not workload_df.empty:
                workload_df.to_excel(writer, sheet_name='工作負載數據', index=False)
            
            # 5. 时间分配数据
            if not time_allocation_df.empty:
                time_allocation_df.to_excel(writer, sheet_name='時間分配數據', index=False)
            
            # 6. 额外任务数据
            if not additional_tasks_df.empty:
                additional_tasks_df.to_excel(writer, sheet_name='額外任務數據', index=False)
            
            # 7. MRB 统计数据 - 完全移除比例計算
            # 完全删除原来的MRB统计代码
            if '是否為MRB' in filtered_df.columns:
                # 直接删除这部分，不再计算MRB比例
                pass

            # 如果需要將所有MRB相關資料列匯出，可以使用以下代碼
            # 创建一个包含所有需要的列的副本
            if 'MRB狀態' in filtered_df.columns and 'MRB訊息' in filtered_df.columns:
                # 只保留四個MRB相關欄位
                mrb_cols = ['MRB狀態', 'MRB訊息']
                
                # 確認基礎標準工時和MRB加時欄位是否存在
                if '基礎標準工時' in filtered_df.columns:
                    mrb_cols.append('基礎標準工時')
                if 'MRB加時' in filtered_df.columns:
                    mrb_cols.append('MRB加時')
                
                # 將選定的MRB欄位匯出到Excel中的"MRB數據"工作表
                try:
                    if len(mrb_cols) > 0:
                        # 只選擇需要的列
                        mrb_data = filtered_df[mrb_cols]
                        # 寫入Excel
                        mrb_data.to_excel(writer, sheet_name='MRB數據', index=False)
                        debug_log(f"成功匯出 {len(mrb_data)} 筆MRB相關資料")
                except Exception as e:
                    debug_log(f"匯出MRB資料時發生錯誤: {str(e)}", level="ERROR")
                                    
            # 獲取workbook和worksheet进行格式设置
            workbook = writer.book
            worksheet = writer.sheets['處理後資料']
            
            # 设置条件格式以高亮MRB记录
            if '是否為MRB' in filtered_df.columns:
                # 找到MRB列的索引
                mrb_col_idx = list(filtered_df.columns).index('是否為MRB')
                col_letter = chr(65 + mrb_col_idx)  # 将列索引转为Excel列字母
                
                # 添加条件格式来高亮MRB记录 - 修改為比較字串"TRUE"
                worksheet.conditional_format(f"{col_letter}2:{col_letter}{len(filtered_df)+1}", {
                    'type': 'cell',
                    'criteria': 'equal to',
                    'value': '"TRUE"',  # 修改這裡，使用字串值
                    'format': workbook.add_format({'bg_color': '#FFC7CE'})
                })
            
            # 設置列寬，特別是MRB相關列
            worksheet.set_column('A:Z', 15)  # 默認列寬
            
            # 為MRB消息列設置更寬的列寬
            if 'MRB訊息' in filtered_df.columns:
                mrb_msg_col_idx = list(filtered_df.columns).index('MRB訊息')
                col_letter = chr(65 + mrb_msg_col_idx)
                worksheet.set_column(f"{col_letter}:{col_letter}", 40)  # 更寬的列寬
        
        excel_data = excel_buffer.getvalue()
        
        # 下載按鈕
        st.download_button(
            label="下載完整Excel報告 (整體計算後結果)",
            data=excel_data,
            file_name="IQC完整數據報告.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        
        debug_log("匯出功能執行完成")
        
    except Exception as e:
        error_msg = f"匯出資料時出錯: {str(e)}\n{traceback.format_exc()}"
        debug_log(error_msg)
        st.error(error_msg)

# 3. 添加一個調試函數，用於檢查MRB狀態
def debug_mrb_status():
    """
    檢查並顯示MRB狀態的詳細信息
    """
    if 'processed_data' not in st.session_state or st.session_state.processed_data is None:
        st.error("沒有處理後的數據可供檢查")
        return
    
    df = st.session_state.processed_data.copy()
    
    # 檢查數據中的列名
    st.write("### 數據欄位名稱")
    st.write(list(df.columns))
    
    # 檢查M欄位的情況
    if 'M' in df.columns:
        st.write("### M欄位情況")
        non_empty_m = df[df['M'].notna()]
        st.write(f"M欄位共有 {len(non_empty_m)} 個非空值")
        
        if not non_empty_m.empty:
            st.write("M欄位非空值示例:")
            st.dataframe(non_empty_m[['料號', 'M', '是否為MRB']].head(10))
    
    # 檢查MRB狀態
    if '是否為MRB' in df.columns:
        st.write("### MRB狀態統計")
        mrb_count = df['是否為MRB'].sum()
        st.write(f"數據中共有 {mrb_count} 筆MRB記錄")
        
        if mrb_count > 0:
            st.write("MRB記錄示例:")
            st.dataframe(df[df['是否為MRB'] == True][['料號', 'M', '是否為MRB', '處理後檢驗標準工時']].head(10))
    
    # 重新檢查每一行，確認MRB判斷邏輯
    st.write("### 重新檢查MRB邏輯")
    
    # 取樣檢查
    sample_size = min(100, len(df))
    sample_df = df.sample(sample_size) if len(df) > 100 else df
    
    results = []
    for idx, row in sample_df.iterrows():
        is_mrb, msg = check_is_mrb(row)
        results.append({
            "索引": idx,
            "料號": row.get('料號', ''),
            "M欄位": row.get('M', ''),
            "原始MRB狀態": row.get('是否為MRB', None),
            "重新檢查結果": is_mrb,
            "檢查訊息": msg,
            "是否一致": row.get('是否為MRB', None) == is_mrb
        })
    
    results_df = pd.DataFrame(results)
    
    # 顯示不一致的結果
    inconsistent = results_df[results_df['是否一致'] == False]
    if not inconsistent.empty:
        st.write(f"發現 {len(inconsistent)} 個MRB狀態不一致的記錄:")
        st.dataframe(inconsistent)
    else:
        st.write("所有抽樣記錄的MRB狀態一致")
    
    # 提供更新選項
    if st.button("使用增強邏輯更新所有MRB狀態"):
        mrb_count = 0
        for idx, row in df.iterrows():
            is_mrb, msg = check_is_mrb(row)
            df.at[idx, '是否為MRB'] = is_mrb
            if is_mrb:
                mrb_count += 1
        
        st.session_state.processed_data = df
        st.success(f"已更新MRB狀態，共發現 {mrb_count} 筆MRB記錄")
        st.rerun()

def create_sidebar():
    """
    側邊欄設計 - 保留原始拖放界面，同時隱藏自定義文件列表
    """
    with st.sidebar:
        # Logo和標題區域 - 使用圖片和黑色漸層背景
        import base64
        
        # 讀取圖片並轉換為base64 - 添加错误处理
        def get_base64_of_bin_file(bin_file):
            try:
                with open(bin_file, 'rb') as f:
                    data = f.read()
                return base64.b64encode(data).decode()
            except FileNotFoundError:
                # 如果找不到图片文件，返回空字符串并显示文本标题
                st.warning(f"无法找到图片文件: {bin_file}")
                st.title("IQC 效率管理系統")
                return ""
            except Exception as e:
                st.error(f"读取图片出错: {e}")
                return ""
        
        # 尝试获取logo的base64编码 - 尝试多个可能的文件路径
        img_paths = ["IQC1.png", "./IQC1.png", "../IQC1.png", "IQC2.png", "IQC.png"]
        img_base64 = ""
        
        for img_path in img_paths:
            try:
                if os.path.exists(img_path):
                    img_base64 = get_base64_of_bin_file(img_path)
                    if img_base64:
                        break
            except:
                continue
        
        # 顯示帶有黑色漸層背景的logo
        st.markdown(f"""
        <div style="text-align: center; margin-bottom: 30px;">
            <div style="background: linear-gradient(to bottom, #333333, #000000, #1A1A1A); 
                        padding: 15px; 
                        border-radius: 8px; 
                        margin-bottom: 10px;
                        box-shadow: 0 4px 6px rgba(0,0,0,0.2);">
                <img src="data:image/png;base64,{img_base64}" style="max-width: 90%; height: auto;">
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        # 檔案上傳區域標題
        st.markdown("""
        <div class="upload-title" style="font-size: 20px; font-weight: 600;">
            📂 資料上傳
        </div>
        """, unsafe_allow_html=True)

        # 使用原始上傳器（不修改其外觀）
        uploaded_files = st.file_uploader("", type=['xlsx', 'xls'], accept_multiple_files=True, 
                    key="excel_files_uploader", label_visibility="collapsed")
        
        # 如果有上傳的文件，只顯示處理按鈕（不顯示自定義文件列表）
        if uploaded_files:
            # 處理按鈕 - 使用藍色背景和粗體文字
            process_button = st.button(
                "處理資料", 
                key="process_data_button", 
                use_container_width=True,
                type="primary"  # 使用Streamlit原生的primary類型獲得藍色背景
            )
            
            if process_button:
                process_files_button_click(uploaded_files, st.session_state.get('start_date'), st.session_state.get('end_date'))
        
        # 視覺分隔線
        st.markdown("<hr style='margin: 25px 0; border: none; height: 1px; background-color: #eee;'>", unsafe_allow_html=True)
        
        st.markdown("""
        <div style="display: flex; align-items: center; margin-bottom: 8px;">
            <div class="upload-title" style="font-size: 20px; font-weight: 600;">
                📅 日期篩選
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        # 使用兩欄布局使日期選擇更緊湊
        col1, col2 = st.columns(2)
        with col1:
            st.write("開始日期")
            start_date = st.date_input("", value=None, key="start_date", label_visibility="collapsed")
        with col2:
            st.write("結束日期")
            end_date = st.date_input("", value=None, key="end_date", label_visibility="collapsed")
        
        # 檔案已處理後的選項
        if 'files_uploaded' in st.session_state and st.session_state.files_uploaded:
            # 應用日期篩選按鈕
            filter_button = st.button(
                "應用日期篩選",
                key="apply_filter_button",
                use_container_width=True
            )
            
            if filter_button:
                with st.spinner("應用日期篩選中..."):
                    try:
                        debug_log(f"應用日期篩選：從 {start_date} 到 {end_date}", level="INFO")
                        
                        # 關鍵修改：清除所有與分析相關的緩存結果
                        cache_keys = [
                            'mrb_analysis_results',       # MRB分析結果
                            'trimmed_avg_efficiency',     # 剔除極值後的效率
                            'trimmed_record_stats',       # 剔除記錄統計
                            'category_efficiency_data',   # 類別效率數據
                            'has_applied_selection'       # 應用選擇標記
                        ]
                        
                        for key in cache_keys:
                            if key in st.session_state:
                                del st.session_state[key]
                                debug_log(f"已清除緩存：{key}", level="INFO")
                        
                        # 重新計算所有指標
                        metrics = calculate_all_metrics(
                            st.session_state.iqc_report_data,
                            st.session_state.pcb_spec_data,
                            st.session_state.pcb_standard_time_data,
                            st.session_state.additional_tasks_data,
                            start_date,
                            end_date
                        )
                        
                        if metrics:
                            # 更新所有數據
                            st.session_state.processed_data = metrics['processed_data']
                            st.session_state.efficiency_data = metrics['efficiency_data']
                            st.session_state.workload_data = metrics['workload_data']
                            st.session_state.time_allocation_data = metrics['time_allocation_data']
                            st.session_state.additional_tasks_monitor_data = metrics['additional_tasks_monitor_data']
                            
                            # 立即進行一次頁面刷新以確保所有數據都被更新
                            debug_log("日期篩選已應用，所有數據已更新", level="INFO")
                            st.success(f"已成功應用日期篩選：{start_date} 到 {end_date}")
                            
                            # 可選：強制重新運行以確保UI更新
                            st.rerun()
                    
                    except Exception as e:
                        error_msg = f"應用篩選時出錯: {str(e)}\n{traceback.format_exc()}"
                        debug_log(error_msg, level="ERROR")
                        st.error(error_msg)
                        st.session_state.processing_error = error_msg
            
            # 添加匯出功能按鈕
            st.markdown("<div class='export-button'>", unsafe_allow_html=True)
            export_button = st.button(
                "💾 下載匯出Excel報表",
                key="export_data_button",
                use_container_width=True
            )
            st.markdown("</div>", unsafe_allow_html=True)
            
            if export_button:
                with st.spinner("準備匯出資料中..."):
                    export_processed_data()
        
        # 修復系統設定折疊選項的CSS樣式 - 針對 Streamlit Cloud 優化
        st.markdown("""
        <style>
        /* 針對 Streamlit Cloud 的系統設定優化 */
        .streamlit-expanderHeader {
            background-color: white !important;
            padding: 12px !important;
            border-radius: 8px !important;
            margin-bottom: 8px !important;
            color: #455A64 !important;
            font-weight: 600 !important;
            font-size: 16px !important;
            line-height: 1.5 !important;
            border: 1px solid #E0E0E0 !important;
        }

        .streamlit-expanderContent {
            background-color: white !important;
            padding: 20px !important;
            border-radius: 8px !important;
            margin-top: 8px !important;
            border: 1px solid #E0E0E0 !important;
        }

        /* 強制修復 Streamlit Cloud 中的選擇框樣式 */
        .streamlit-expanderContent .stSelectbox {
            margin-bottom: 20px !important;
        }

        .streamlit-expanderContent .stSelectbox label {
            color: #333 !important;
            font-size: 16px !important;
            font-weight: 600 !important;
            margin-bottom: 8px !important;
            display: block !important;
            line-height: 1.4 !important;
        }

        .streamlit-expanderContent .stSelectbox > div {
            margin-top: 8px !important;
            margin-bottom: 20px !important;
        }

        /* 強制修復複選框樣式 */
        .streamlit-expanderContent .stCheckbox {
            margin-bottom: 20px !important;
            margin-top: 15px !important;
        }

        .streamlit-expanderContent .stCheckbox label {
            color: #333 !important;
            font-size: 16px !important;
            font-weight: 500 !important;
            line-height: 1.5 !important;
            display: flex !important;
            align-items: center !important;
        }

        /* 強制修復按鈕樣式 */
        .streamlit-expanderContent .stButton {
            margin-top: 15px !important;
            margin-bottom: 10px !important;
        }

        .streamlit-expanderContent .stButton button {
            background-color: #90A4AE !important;
            color: white !important;
            border: none !important;
            border-radius: 6px !important;
            padding: 10px 20px !important;
            font-size: 14px !important;
            font-weight: 500 !important;
            width: 100% !important;
            margin-top: 5px !important;
        }

        .streamlit-expanderContent .stButton button:hover {
            background-color: #78909C !important;
            transform: translateY(-1px) !important;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1) !important;
        }

        /* 添加間距分隔符 */
        .settings-divider {
            height: 1px;
            background-color: #E0E0E0;
            margin: 20px 0;
            border: none;
        }

        /* 確保文字不會重疊 */
        .streamlit-expanderContent > div {
            margin-bottom: 15px !important;
            clear: both !important;
        }

        /* 針對 Streamlit Cloud 的額外修正 */
        .streamlit-expanderContent .element-container {
            margin-bottom: 15px !important;
        }

        .streamlit-expanderContent .row-widget {
            margin-bottom: 15px !important;
        }
        </style>
        """, unsafe_allow_html=True)
        
        # 系統設定折疊選項 - V23 針對 Streamlit Cloud 優化版本
        with st.expander("⚙️ 系統設定", expanded=False):
            # 日誌級別設定 - 使用 HTML 標籤避免重疊
            st.markdown('<p style="font-weight: 600; color: #333; margin-bottom: 8px; font-size: 16px;">日誌級別</p>', unsafe_allow_html=True)
            
            log_levels = ["DEBUG", "INFO", "WARNING", "ERROR"]
            selected_log_level = st.selectbox(
                "",
                options=log_levels,
                index=log_levels.index(st.session_state.get('log_level', "DEBUG")),
                key="log_level_select_v23",
                label_visibility="collapsed"
            )
            
            if selected_log_level != st.session_state.get('log_level', "DEBUG"):
                st.session_state.log_level = selected_log_level
                st.success(f"✅ 日誌級別已設為 {selected_log_level}")
            
            # 間距分隔符
            st.markdown('<hr class="settings-divider">', unsafe_allow_html=True)
            
            # 性能模式設定
            st.markdown('<p style="font-weight: 600; color: #333; margin-bottom: 8px; font-size: 16px;">性能設定</p>', unsafe_allow_html=True)
            
            performance_mode = st.checkbox(
                "啟用高性能模式",
                value=st.session_state.get('performance_mode', False),
                help="啟用後將減少日誌輸出，提高計算速度，但不會顯示詳細日誌",
                key="performance_mode_check_v23"
            )
            
            if performance_mode != st.session_state.get('performance_mode', False):
                st.session_state.performance_mode = performance_mode
                if performance_mode:
                    st.session_state.log_level = "WARNING"
                    st.info("已啟用高性能模式，日誌級別自動設為WARNING")
                else:
                    st.session_state.log_level = "DEBUG"
                    st.info("已停用高性能模式，日誌級別自動設為DEBUG")
            
            # 間距分隔符
            st.markdown('<hr class="settings-divider">', unsafe_allow_html=True)
            
            # 系統操作按鈕
            st.markdown('<p style="font-weight: 600; color: #333; margin-bottom: 8px; font-size: 16px;">系統操作</p>', unsafe_allow_html=True)
            
            # 清理日誌的按鈕
            if st.button("清理日誌", key="clear_logs_btn_v23", use_container_width=True):
                if 'logs' in st.session_state.debug_info:
                    st.session_state.debug_info['logs'] = []
                st.success("已清理所有日誌")
            
            # 重設所有數據的按鈕
            if st.button("重設所有數據緩存", key="reset_cache_btn_v23", use_container_width=True, help="清除所有已計算的結果緩存，強制重新計算"):
                # 清除所有與分析相關的緩存
                cache_keys = [
                    'mrb_analysis_results',
                    'trimmed_avg_efficiency',
                    'trimmed_record_stats',
                    'category_efficiency_data',
                    'has_applied_selection',
                    'selected_material_categories',
                    'selected_category_mode',
                    'selected_main_category',
                    'selected_subcategories',
                    'merge_categories'
                ]
                
                for key in cache_keys:
                    if key in st.session_state:
                        del st.session_state[key]
                
                st.success("已清除所有數據緩存，將在下次更新時重新計算")


# 改进后的文件分类函数
def classify_files(files):
    """
    根据文件名和内容智能分类上传的文件
    返回四组文件：IQC Report, PCB建檔明細, PCB標準工時對應表, IQC額外任務紀錄清單
    """
    debug_log(f"开始分类 {len(files)} 个文件")
    
    iqc_report_files = []
    pcb_specs_files = []
    pcb_standard_time_files = []
    additional_tasks_files = []
    
    # 打印所有上传的文件名，便于调试
    file_names = [f.name for f in files]
    debug_log(f"所有上传文件: {file_names}")
    
    # 更完整的关键词匹配规则
    iqc_report_keywords = ['iqc', 'report', '報告', '檢驗報告', '檢驗', 'inspector']
    pcb_specs_keywords = ['pcb', '建檔', '明細', 'spec', '建立規格']
    pcb_std_time_keywords = ['標準工時', 'standard', 'time', '對應表', '工時']
    additional_tasks_keywords = ['額外', '任務', 'task', '清單', '紀錄', '工作事項']
    
    # 精确匹配特定文件名
    pcb_std_time_exact = ['pcb標準工時對應表.xlsx', 'pcb標準工時對應表.xls']
    additional_tasks_exact = ['iqc額外任務紀錄清單.xlsx', 'iqc額外任務紀錄清單.xls']
    
    for file in files:
        # 保存原始文件指针位置
        file_position = file.tell()
        
        filename = file.name
        filename_lower = filename.lower()
        file_classified = False
        
        # 1. 首先尝试精确匹配文件名
        if any(exact_name.lower() == filename_lower for exact_name in pcb_std_time_exact):
            pcb_standard_time_files.append(file)
            debug_log(f"文件 {filename} 通过精确匹配被识别为 PCB標準工時對應表")
            file_classified = True
        
        elif any(exact_name.lower() == filename_lower for exact_name in additional_tasks_exact):
            additional_tasks_files.append(file)
            debug_log(f"文件 {filename} 通过精确匹配被识别为 IQC額外任務紀錄清單")
            file_classified = True
            
        # 2. 如果没有精确匹配成功，尝试关键词匹配
        elif not file_classified:
            # PCB标准工时关键词优先级高于其他类型
            if any(keyword in filename_lower for keyword in pcb_std_time_keywords):
                pcb_standard_time_files.append(file)
                debug_log(f"文件 {filename} 通过关键词被识别为 PCB標準工時對應表")
                file_classified = True
                
            elif any(keyword in filename_lower for keyword in additional_tasks_keywords):
                additional_tasks_files.append(file)
                debug_log(f"文件 {filename} 通过关键词被识别为 IQC額外任務紀錄清單")
                file_classified = True
                
            elif any(keyword in filename_lower for keyword in pcb_specs_keywords):
                pcb_specs_files.append(file)
                debug_log(f"文件 {filename} 通过关键词被识别为 PCB建檔明細")
                file_classified = True
                
            elif any(keyword in filename_lower for keyword in iqc_report_keywords) or 'iqc report' in filename_lower:
                iqc_report_files.append(file)
                debug_log(f"文件 {filename} 通过关键词被识别为 IQC Report")
                file_classified = True
        
        # 3. 如果仍然未识别，尝试读取内容
        if not file_classified:
            try:
                # 重置文件指针
                file.seek(0)
                
                # 读取Excel文件的表头来识别文件类型
                df = pd.read_excel(file, nrows=5)
                columns = [str(col).lower() for col in df.columns]
                debug_log(f"文件 {filename} 的表头: {columns[:10]}")
                
                # 判断是否为PCB标准工时对应表
                if any(keyword in ','.join(columns) for keyword in ['面積範圍', '面积范围', '壓合總孔數', '压合总孔数', 'pcb標準工時']):
                    pcb_standard_time_files.append(file)
                    debug_log(f"文件 {filename} 通过内容被识别为 PCB標準工時對應表")
                
                # 判断是否为IQC额外任务记录清单
                elif any(keyword in ','.join(columns) for keyword in ['姓名', '用時(分鐘)', '用时(分钟)', '工作事項分類', '工作事项分类']):
                    additional_tasks_files.append(file)
                    debug_log(f"文件 {filename} 通过内容被识别为 IQC額外任務紀錄清單")
                
                # 判断是否为PCB建档明细
                elif any(keyword in ','.join(columns) for keyword in ['料號', '料号', '壓合孔數', '压合孔数', '版長', '版长', '版寬', '版宽']):
                    pcb_specs_files.append(file)
                    debug_log(f"文件 {filename} 通过内容被识别为 PCB建檔明細")
                
                # 判断是否为IQC Report
                elif any(keyword in ','.join(columns) for keyword in ['inspector', '檢驗員', '检验员', '檢驗人員', '检验人员', 'mrb']):
                    iqc_report_files.append(file)
                    debug_log(f"文件 {filename} 通过内容被识别为 IQC Report")
                
                else:
                    # 如果仍然无法识别，根据特定线索进一步判断
                    if '工時' in filename_lower or '工时' in filename_lower or 'time' in filename_lower:
                        pcb_standard_time_files.append(file)
                        debug_log(f"文件 {filename} 通过额外规则被识别为 PCB標準工時對應表")
                    elif '任務' in filename_lower or '任务' in filename_lower or 'task' in filename_lower:
                        additional_tasks_files.append(file)
                        debug_log(f"文件 {filename} 通过额外规则被识别为 IQC額外任務紀錄清單")
                    else:
                        # 最后的默认分类
                        iqc_report_files.append(file)
                        debug_log(f"文件 {filename} 无法确定类型，默认识别为 IQC Report")
                
            except Exception as e:
                debug_log(f"尝试读取文件 {filename} 内容时出错: {str(e)}")
                # 如果读取失败，尝试最后一次通过文件名判断
                if any(term in filename_lower for term in ['工時', '工时', 'time']):
                    pcb_standard_time_files.append(file)
                    debug_log(f"文件 {filename} 读取失败，通过文件名判断为 PCB標準工時對應表")
                elif any(term in filename_lower for term in ['任務', '任务', 'task']):
                    additional_tasks_files.append(file)
                    debug_log(f"文件 {filename} 读取失败，通过文件名判断为 IQC額外任務紀錄清單")
                else:
                    # 如果还是不能确定，默认为IQC Report
                    iqc_report_files.append(file)
                    debug_log(f"文件 {filename} 读取失败，默认识别为 IQC Report")
            
        # 重置文件指针回原位置
        file.seek(file_position)
    
    # 检查特定类型文件是否缺失，如果用户明确上传了文件但未被正确识别，强制分类
    if not pcb_standard_time_files and any('工時' in f.name.lower() or '工时' in f.name.lower() or 'time' in f.name.lower() for f in files):
        for file in files:
            if ('工時' in file.name.lower() or '工时' in file.name.lower() or 'time' in file.name.lower()) and file not in pcb_standard_time_files:
                pcb_standard_time_files.append(file)
                debug_log(f"强制将文件 {file.name} 识别为 PCB標準工時對應表")
                # 从其他类别中移除(如果存在)
                if file in iqc_report_files: iqc_report_files.remove(file)
                if file in pcb_specs_files: pcb_specs_files.remove(file)
                if file in additional_tasks_files: additional_tasks_files.remove(file)
    
    if not additional_tasks_files and any('任務' in f.name.lower() or '任务' in f.name.lower() or 'task' in f.name.lower() for f in files):
        for file in files:
            if ('任務' in file.name.lower() or '任务' in file.name.lower() or 'task' in file.name.lower()) and file not in additional_tasks_files:
                additional_tasks_files.append(file)
                debug_log(f"强制将文件 {file.name} 识别为 IQC額外任務紀錄清單")
                # 从其他类别中移除(如果存在)
                if file in iqc_report_files: iqc_report_files.remove(file)
                if file in pcb_specs_files: pcb_specs_files.remove(file)
                if file in pcb_standard_time_files: pcb_standard_time_files.remove(file)
    
    # 再次检查，如果仍然缺少特定类型，尝试从未分类文件或多余的IQC Report文件中找出可能的匹配
    remaining_files = [f for f in files if f not in iqc_report_files and f not in pcb_specs_files and 
                       f not in pcb_standard_time_files and f not in additional_tasks_files]
    
    if not pcb_standard_time_files and remaining_files:
        # 从剩余文件中添加第一个作为PCB标准工时对应表
        pcb_standard_time_files.append(remaining_files[0])
        debug_log(f"未找到PCB標準工時對應表，强制将文件 {remaining_files[0].name} 识别为此类型")
        remaining_files.pop(0)
    
    if not additional_tasks_files and remaining_files:
        # 从剩余文件中添加第一个作为IQC额外任务记录清单
        additional_tasks_files.append(remaining_files[0])
        debug_log(f"未找到IQC額外任務紀錄清單，强制将文件 {remaining_files[0].name} 识别为此类型")
        remaining_files.pop(0)
    
    # 最后一次检查，如果仍然缺少，且IQC Report有多个，则将其中一个重新分类
    if not pcb_standard_time_files and len(iqc_report_files) > 1:
        file = iqc_report_files.pop() # 移除最后一个IQC Report
        pcb_standard_time_files.append(file)
        debug_log(f"未找到PCB標準工時對應表，从IQC Report中重新分类文件 {file.name}")
    
    if not additional_tasks_files and len(iqc_report_files) > 1:
        file = iqc_report_files.pop() # 移除最后一个IQC Report
        additional_tasks_files.append(file)
        debug_log(f"未找到IQC額外任務紀錄清單，从IQC Report中重新分类文件 {file.name}")
    
    # 输出最终分类结果
    debug_log(f"文件分类完成: IQC Report({len(iqc_report_files)}), PCB建檔明細({len(pcb_specs_files)}), " +
             f"PCB標準工時對應表({len(pcb_standard_time_files)}), IQC額外任務紀錄清單({len(additional_tasks_files)})")
    
    debug_log(f"PCB標準工時對應表: {[f.name for f in pcb_standard_time_files]}")
    debug_log(f"IQC額外任務紀錄清單: {[f.name for f in additional_tasks_files]}")
    
    return iqc_report_files, pcb_specs_files, pcb_standard_time_files, additional_tasks_files

# 自定義進度條函數
def custom_progress_bar():
    """創建自定義進度條，返回更新函數和容器"""
    progress_container = st.empty()
    
    # 初始化進度條為0%
    progress_container.markdown("""
    <div style="margin: 10px 0;">
        <div style="background-color: #E0E0E0; height: 6px; border-radius: 3px; width: 100%;">
            <div style="background-color: #90A4AE; height: 6px; width: 0%; border-radius: 3px; transition: width 0.3s ease;"></div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # 更新進度的函數
    def update_progress(percent):
        # 確保百分比在0-100之間
        percent = max(0, min(100, percent))
        progress_container.markdown(f"""
        <div style="margin: 10px 0;">
            <div style="background-color: #E0E0E0; height: 6px; border-radius: 3px; width: 100%;">
                <div style="background-color: #90A4AE; height: 6px; width: {percent}%; border-radius: 3px; transition: width 0.3s ease;"></div>
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    return update_progress, progress_container

# 自定義進度圈函數
def custom_spinner(text="正在處理資料，請稍候..."):
    """創建自定義進度圈(spinner)"""
    spinner_container = st.empty()
    spinner_container.markdown(f"""
    <div style="display: flex; align-items: center; margin: 10px 0; color: #37474F;">
        <div class="custom-spinner" style="border: 3px solid #f3f3f3; border-top: 3px solid #90A4AE; border-radius: 50%; width: 16px; height: 16px; margin-right: 10px; animation: spin 1s linear infinite;"></div>
        {text}
    </div>
    <style>
    @keyframes spin {{
        0% {{ transform: rotate(0deg); }}
        100% {{ transform: rotate(360deg); }}
    }}
    </style>
    """, unsafe_allow_html=True)
    return spinner_container

# 函數2: 文件上傳和日期篩選處理 (處理按鈕點擊)
def process_files_button_click(uploaded_files, start_date, end_date):
    st.session_state.processing_error = None
    st.session_state.debug_info = {'logs': []}
    
    start_time = time.time()  # 記錄處理開始時間
    
    # 使用自定義進度圈替代原來的st.spinner
    spinner = custom_spinner("正在處理資料，請稍候...")
    
    try:
        # 自動識別分類檔案
        iqc_report_files, pcb_specs_files, pcb_standard_time_files, additional_tasks_files = classify_files(uploaded_files)
        
        # 輸出分類結果到日誌
        debug_log(f"分類結果: IQC Report({len(iqc_report_files)}), PCB建檔明細({len(pcb_specs_files)}), " +
                  f"PCB標準工時對應表({len(pcb_standard_time_files)}), IQC額外任務紀錄清單({len(additional_tasks_files)})", level="INFO")
        
        # 顯示分類結果
        with st.expander("檔案分類結果", expanded=False):
            st.write("IQC Report: " + ", ".join([f.name for f in iqc_report_files]))
            st.write("PCB建檔明細: " + ", ".join([f.name for f in pcb_specs_files]))
            st.write("PCB標準工時對應表: " + ", ".join([f.name for f in pcb_standard_time_files]))
            st.write("IQC額外任務紀錄清單: " + ", ".join([f.name for f in additional_tasks_files]))
        
        # 創建自定義進度條，替代原本的st.progress()
        update_progress, progress_container = custom_progress_bar()
        debug_log("開始處理檔案", level="INFO")
        
        # 優化數據處理，使用性能模式設定
        high_performance = st.session_state.get('performance_mode', False)
        
        # 逐步處理各檔案類型，更新進度條
        update_progress(10)
        iqc_report_data = process_multiple_iqc_reports_optimized(
            iqc_report_files
        ) if iqc_report_files else None

        update_progress(40)
        pcb_spec_data = process_multiple_pcb_specs(
            pcb_specs_files
        ) if pcb_specs_files else None

        update_progress(60)
        pcb_standard_time_data = process_multiple_pcb_standard_times(
            pcb_standard_time_files
        ) if pcb_standard_time_files else None

        update_progress(80)
        additional_tasks_data = process_multiple_additional_tasks(
            additional_tasks_files
        ) if additional_tasks_files else None
        
        # 檢查是否所有必要數據都已處理
        if iqc_report_data is None:
            st.error("無法處理IQC Report數據，請檢查上傳的檔案")
            # 清除進度條和spinner
            progress_container.empty()
            spinner.empty()
            return False
        
        # 存儲處理後的數據
        st.session_state.iqc_report_data = iqc_report_data
        st.session_state.pcb_spec_data = pcb_spec_data
        st.session_state.pcb_standard_time_data = pcb_standard_time_data
        st.session_state.additional_tasks_data = additional_tasks_data
        
        # 清除任何之前的極值設定和日期篩選狀態
        for key in ['trim_percentage', 'using_trimmed_efficiency', 'trimmed_avg_efficiency', 'trimmed_record_stats']:
            if key in st.session_state:
                del st.session_state[key]
        
        debug_log("檔案處理完成，計算指標", level="INFO")
        
        # 使用優化後的函數計算指標
        metrics = calculate_all_metrics(
            iqc_report_data,
            pcb_spec_data,
            pcb_standard_time_data,
            additional_tasks_data,
            start_date,
            end_date
        )
        
        update_progress(100)
        
        if metrics:
            st.session_state.processed_data = metrics['processed_data']
            st.session_state.efficiency_data = metrics['efficiency_data']
            st.session_state.workload_data = metrics['workload_data']
            st.session_state.time_allocation_data = metrics['time_allocation_data']
            st.session_state.additional_tasks_monitor_data = metrics['additional_tasks_monitor_data']
            st.session_state.files_uploaded = True
            
            # 確保初始沒有選擇的物料類別
            st.session_state.selected_material_categories = []
            
            # 保存日期篩選狀態
            st.session_state.start_date = start_date
            st.session_state.end_date = end_date
            
            end_time = time.time()  # 記錄處理結束時間
            processing_time = end_time - start_time
            debug_log(f"指標計算完成，處理時間: {processing_time:.2f}秒", level="INFO")
            
            # 清除進度條和spinner
            progress_container.empty()
            spinner.empty()
            
            # 使用自定義樣式顯示成功訊息，而不是使用st.success()
            st.markdown(f"""
            <div style="display: flex; align-items: center; color: #37474F; padding: 10px 0; 
                font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; 
                font-size: 14px; font-weight: 400; letter-spacing: normal; line-height: 1.6;">
                <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" 
                    stroke="#90A4AE" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="margin-right: 8px;">
                    <circle cx="12" cy="12" r="10"></circle>
                    <polyline points="16 10 11 15 8 12"></polyline>
                </svg>
                資料處理完成！用時 {processing_time:.2f} 秒
            </div>
            """, unsafe_allow_html=True)
            
            return True
        
        # 清除進度條和spinner
        progress_container.empty()
        spinner.empty()
        return False
            
    except Exception as e:
        error_msg = f"處理檔案時出錯: {str(e)}\n{traceback.format_exc()}"
        debug_log(error_msg, level="ERROR")
        st.error(error_msg)
        st.session_state.processing_error = error_msg
        
        # 清除進度條和spinner
        if 'progress_container' in locals():
            progress_container.empty()
        if 'spinner' in locals():
            spinner.empty()
            
        return False

def render_settings_panel():
    """
    渲染設定面板，讓用戶可以調整程式行為
    """
    st.sidebar.header("系統設定")
    
    # 日誌級別設定
    log_levels = ["DEBUG", "INFO", "WARNING", "ERROR"]
    selected_log_level = st.sidebar.selectbox(
        "日誌級別",
        options=log_levels,
        index=log_levels.index(st.session_state.get('log_level', "DEBUG"))
    )
    
    if selected_log_level != st.session_state.get('log_level', "DEBUG"):
        st.session_state.log_level = selected_log_level
        st.sidebar.success(f"日誌級別已設為 {selected_log_level}")
    
    # 在設置面板中添加PCB標準工時調試功能
    st.sidebar.header("PCB標準工時調試")
    if st.sidebar.button("檢查PCB標準工時對應"):
        if 'processed_data' in st.session_state and st.session_state.processed_data is not None:
            df = st.session_state.processed_data.copy()
            
            # 篩選QB類型料號
            qb_df = df[df['類別'] == 'QB']
            
            if not qb_df.empty:
                # 顯示QB料號的標準工時
                st.write(f"### 發現 {len(qb_df)} 筆QB類型料號")
                
                # 計算基本統計資訊
                avg_std_time = qb_df['處理後檢驗標準工時'].mean()
                min_std_time = qb_df['處理後檢驗標準工時'].min()
                max_std_time = qb_df['處理後檢驗標準工時'].max()
                
                st.write(f"平均標準工時: {avg_std_time:.2f}分鐘")
                st.write(f"最小標準工時: {min_std_time:.2f}分鐘")
                st.write(f"最大標準工時: {max_std_time:.2f}分鐘")
                
                # 顯示標準工時分佈
                st.write("### 標準工時分佈")
                std_time_counts = qb_df['處理後檢驗標準工時'].value_counts().reset_index()
                std_time_counts.columns = ['標準工時', '料號數量']
                std_time_counts = std_time_counts.sort_values('標準工時')
                
                # 創建分佈圖
                fig = px.bar(
                    std_time_counts,
                    x='標準工時',
                    y='料號數量',
                    title="PCB標準工時分佈"
                )
                st.plotly_chart(fig, use_container_width=True)
                
                # 顯示詳細清單
                st.write("### QB料號與標準工時詳細清單")
                display_df = qb_df[['料號', '處理後檢驗標準工時', '是否為MRB']].copy()
                display_df.columns = ['料號', '標準工時(分鐘)', 'MRB狀態']
                st.dataframe(display_df.sort_values('標準工時(分鐘)'), use_container_width=True)
            else:
                st.write("未發現QB類型料號")
        else:
            st.write("沒有可用數據，請先上傳並處理文件")

    # 性能模式設定
    performance_mode = st.sidebar.checkbox(
        "啟用高性能模式",
        value=st.session_state.get('performance_mode', False),
        help="啟用後將減少日誌輸出，提高計算速度，但不會顯示詳細日誌"
    )
    
    if performance_mode != st.session_state.get('performance_mode', False):
        st.session_state.performance_mode = performance_mode
        if performance_mode:
            st.session_state.log_level = "WARNING"  # 高性能模式下將日誌級別設為WARNING
            st.sidebar.info("已啟用高性能模式，日誌級別自動設為WARNING")
        else:
            st.session_state.log_level = "DEBUG"
            st.sidebar.info("已停用高性能模式，日誌級別自動設為DEBUG")
    
    # 添加清理日誌的按鈕
    if st.sidebar.button("清理日誌"):
        if 'logs' in st.session_state.debug_info:
            st.session_state.debug_info['logs'] = []
        st.sidebar.success("已清理所有日誌")

def get_base64_of_bin_file(bin_file):
    """
    將二進制文件轉換為base64編碼的字符串
    
    參數:
    bin_file (str): 二進制文件的路徑
    
    返回:
    str: base64編碼的字符串
    """
    try:
        with open(bin_file, 'rb') as f:
            data = f.read()
        return base64.b64encode(data).decode()
    except Exception as e:
        print(f"讀取圖片失敗: {e}")
        return ""

def show_cover_page():
    """
    封面頁面 - 使用纯CSS优化按钮外观和悬停效果
    """
    # 检查URL参数
    params = st.experimental_get_query_params()
    if "enter_clicked" in params:
        st.session_state.show_cover = False
        st.rerun()
    
    # 隐藏所有Streamlit默认元素并应用自定义样式
    st.markdown("""
        <style>
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        header {visibility: hidden;}
        .stDeployButton {display:none;}
        .stToolbar {display:none;}
        .stAlert {display:none !important;}
        
        /* 基本页面设置 */
        .stApp {
            background-color: black;
            overflow: hidden !important;
        }
        
        /* 移除所有内边距 */
        .main .block-container {
            padding: 0 !important;
            max-width: 100% !important;
        }
        
        /* 最大化增强按钮样式清晰度 */
        .stButton {
            position: fixed !important;
            bottom: 50px !important;
            left: 50% !important;
            transform: translateX(-50%) !important;
            z-index: 1000 !important;
            width: 250px !important;
        }
        
        .stButton button {
            background-color: rgba(0, 0, 0, 0.5) !important; /* 半透明黑色 */
            color: #FFFFFF !important; /* 纯白色文字 */
            border: 2px solid white !important; /* 白色边框 */
            width: 100% !important;
            height: 50px !important; /* 固定高度 */
            border-radius: 6px !important;
            font-weight: 10000 !important; /* 极粗字体 */
            font-size: 20px !important; /* 更大字体 */
            letter-spacing: 4px !important; /* 更大字间距 */
            font-family: Arial, Helvetica, sans-serif !important; /* 无衬线字体 */
            text-rendering: geometricPrecision !important; /* 提高文字渲染精度 */
            -webkit-font-smoothing: antialiased !important;
            -moz-osx-font-smoothing: grayscale !important;
            padding: 0 !important; /* 移除内边距干扰 */
            transition: all 0.3s ease !important; /* 平滑过渡效果 */
        }
        
        /* 实际悬停效果 */
        .stButton button:hover {
            background-color: rgba(255, 255, 255, 0.9) !important; /* 变为半透明白色 */
            transform: translateY(-3px) !important; /* 上浮效果 */
            box-shadow: 0 0 20px rgba(255, 255, 255, 0.5) !important; /* 发光效果 */
        }
        
        /* 确保页面不出现滚动条 */
        html, body {
            overflow: hidden !important;
            height: 100% !important;
        }
                
        /* 隐藏侧边栏和侧边栏里的返回按钮 */
        [data-testid="stSidebar"] {
            display: none !important;
        }
                
        </style>
    """, unsafe_allow_html=True)
    
    # 尝试获取图片并转换为base64
    try:
        import base64
        from PIL import Image
        import io
        import os
        
        # 尝试多个可能的图片路径
        img_paths = ["IQC3.png", "./IQC3.png", "../IQC3.png"]
        img = None
        
        for path in img_paths:
            if os.path.exists(path):
                try:
                    img = Image.open(path)  # 修正：使用實際的路徑而不是固定的 "cover.png"
                    successful_path = path
                    break
                except Exception as e:
                    print(f"無法開啟圖片 {path}: {e}")
                    continue
        
        if img:
            # 转换为base64
            buffered = io.BytesIO()
            img.save(buffered, format="PNG")
            img_str = base64.b64encode(buffered.getvalue()).decode()
            
            # 使用HTML直接控制图片显示
            st.markdown(f"""
                <div style="display: flex; justify-content: center; align-items: center; 
                            height: 85vh; width: 100%; margin: 0; padding: 0;">
                    <img src="data:image/png;base64,{img_str}" 
                         style="width: auto; height: auto; max-width: 65%; max-height: 80vh; 
                                min-width: 400px; display: block; margin: 0 auto;">
                </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown("""
                <div style="display: flex; justify-content: center; align-items: center; 
                            height: 80vh; width: 100%;">
                    <div style="width: 400px; height: 250px; background-color: #111; 
                                display: flex; justify-content: center; align-items: center;">
                        <p style="color: #666; font-size: 18px;">图片未找到</p>
                    </div>
                </div>
            """, unsafe_allow_html=True)
    except Exception as e:
        st.markdown(f"""
            <div style="display: flex; justify-content: center; align-items: center; 
                        height: 80vh; width: 100%;">
                <div style="width: 400px; height: 250px; background-color: #111; 
                            display: flex; justify-content: center; align-items: center;">
                    <p style="color: #666; font-size: 18px;">加载错误: {str(e)}</p>
                </div>
            </div>
        """, unsafe_allow_html=True)
    
    # 添加ENTER按钮
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if st.button("ENTER", key="enter_button", use_container_width=True):
            st.session_state.show_cover = False
            st.rerun()

def main():

    # 添加標題和描述
    st.title("IQC 效率管理系統")
    st.markdown("透過數據量化分析，分析IQC檢驗效率、工作負載、時間管理分配，從而協助提升IQC效能與品質水平。")
    # 建立側邊欄
    create_sidebar()
    
    # 主頁面內容    
    # 顯示處理狀態
    render_status()
    
    # 確保初始沒有選擇物料類別
    if 'selected_material_categories' not in st.session_state:
        st.session_state.selected_material_categories = []
    
    # 如果檔案已上傳，顯示儀表板
    if st.session_state.files_uploaded:
        # 預先計算MRB統計數據，確保一致性
        if 'processed_data' in st.session_state and st.session_state.processed_data is not None:
            if 'mrb_analysis_results' not in st.session_state:
                mrb_results = calculate_mrb_statistics(
                    st.session_state.processed_data,
                    st.session_state.get('start_date'),
                    st.session_state.get('end_date')
                )
                if mrb_results:
                    st.session_state.mrb_analysis_results = mrb_results
        # 在處理數據前設定高性能模式提醒        
        # 創建摘要指標區域 - 優雅的三欄佈局
        if 'efficiency_data' in st.session_state and 'overall_efficiency_ranking' in st.session_state.efficiency_data:
            overall_eff = st.session_state.efficiency_data['overall_efficiency_ranking']
            if not overall_eff.empty:
                avg_efficiency = overall_eff['efficiency'].mean()
                
                # 三欄佈局顯示關鍵指標
                st.markdown("""
                <div style="margin-bottom: 25px;">
                    <h3 style="font-size: 20px; margin-bottom: 15px;">關鍵指標概覽 🎯</h3>
                </div>
                """, unsafe_allow_html=True)
                
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    # 效率指標卡片 - 藍色主題
                    delta_val = None
                    if avg_efficiency >= 1:
                        delta_val = "高於標準"
                    else:
                        delta_val = "低於標準"
                    
                    st.metric(
                        "整體平均效率", 
                        f"{avg_efficiency:.2f}", 
                        delta=delta_val,
                        delta_color="normal",
                        help="所有檢驗員的平均效率指標"
                    )
                
                # 重要改變：從MRB分析結果中獲取MRB率
                # 如果已經計算過MRB分析，則優先使用MRB分析頁面中的計算結果
                if 'mrb_analysis_results' in st.session_state:
                    mrb_results = st.session_state.mrb_analysis_results
                    mrb_rate = mrb_results['mrb_rate']
                    mrb_count = mrb_results['mrb_count']
                    total_count = mrb_results['total_count']
                    
                    with col2:
                        # MRB率指標卡片 - 紅色系
                        st.metric(
                            "整體MRB率", 
                            f"{mrb_rate:.2%}", 
                            delta=f"{mrb_count} 筆 MRB",
                            delta_color="off",
                            help="所有檢驗項目的MRB比率"
                        )
                # 如果沒有MRB分析結果，則使用原始方法計算
                elif 'processed_data' in st.session_state and st.session_state.processed_data is not None:
                    df = st.session_state.processed_data
                    if '是否為MRB' in df.columns:
                        # 計算總體MRB率
                        total_count = len(df)
                        if df['是否為MRB'].dtype != 'object':
                            mrb_count = df['是否為MRB'].sum()
                        else:
                            mrb_count = df['是否為MRB'].apply(
                                lambda x: True if isinstance(x, str) and x.upper() in ['TRUE', 'MRB'] else 
                                       (True if x is True else False)
                            ).sum()
                        
                        overall_mrb_rate = mrb_count / total_count if total_count > 0 else 0
                        
                        with col2:
                            # MRB率指標卡片 - 紅色系
                            st.metric(
                                "整體MRB率", 
                                f"{overall_mrb_rate:.2%}", 
                                delta=f"{mrb_count} 筆 MRB",
                                delta_color="off",
                                help="所有檢驗項目的MRB比率"
                            )
                
                # 如果有工作負載數據
                if 'workload_data' in st.session_state and st.session_state.workload_data is not None:
                    workload_df = st.session_state.workload_data
                    if 'workload_index' in workload_df.columns:
                        avg_workload = workload_df['workload_index'].mean()
                        
                        with col3:
                            # 工作負載指標卡片 - 綠色系
                            delta_color = "normal"
                            if avg_workload > 1.2:
                                delta_val = "工作量高"
                                delta_color = "inverse"
                            elif avg_workload < 0.8:
                                delta_val = "工作量低"
                                delta_color = "normal"
                            else:
                                delta_val = "適中"
                                delta_color = "off"
                                
                            st.metric(
                                "平均工作負載", 
                                f"{avg_workload:.2f}", 
                                delta=delta_val,
                                delta_color=delta_color,
                                help="所有檢驗員的平均工作負載指數 (8小時=1.0)"
                            )
                
                # 添加分隔線
                st.markdown("<hr style='margin: 30px 0 20px 0; border: none; height: 1px; background-color: #E2E8F0;'>", unsafe_allow_html=True)
        
        # 修改部分: 增加選項卡的響應式設計
        st.markdown("""
        <style>
        /* 使所有標籤頁自適應寬度 */
        button[data-baseweb="tab"] {
            min-width: 0 !important; /* 移除最小寬度限制 */
            flex-grow: 1 !important;
            flex-basis: 0 !important;
            padding: 8px 4px !important; /* 減少內邊距 */
            margin: 0 2px !important; /* 減少標籤間距 */
            background-color: #78909C !important;
            color: black !important;
            border-radius: 4px !important;
            font-size: 13px !important; /* 減小字體大小 */
            font-weight: normal !important;
            text-align: center !important;
            white-space: nowrap !important; /* 防止文字換行 */
            overflow: hidden !important; /* 隱藏溢出文字 */
            text-overflow: ellipsis !important; /* 使用省略號表示溢出 */
        }

        /* 選中的標籤頁樣式 */
        button[data-baseweb="tab"][aria-selected="true"] {
            background-color: #607D8B !important;
            color: white !important;
            font-weight: bold !important;
        }

        /* 標籤頁容器樣式 */
        div[data-testid="stHorizontalBlock"] [data-testid="stTabsSwimlane"] {
            gap: 2px !important; /* 減少間隙 */
            flex-wrap: nowrap !important; /* 避免換行 */
            overflow-x: auto !important; /* 允許水平滾動 */
            scrollbar-width: none !important; /* 隱藏滾動條 (Firefox) */
        }

        /* 隱藏滾動條 (Chrome, Safari) */
        div[data-testid="stHorizontalBlock"] [data-testid="stTabsSwimlane"]::-webkit-scrollbar {
            display: none !important;
        }

        /* 標籤組樣式改進 */
        [data-testid="stTabs"] {
            width: 100% !important;
        }

        [data-testid="stTabs"] > div:first-child {
            overflow-x: auto !important;
            flex-wrap: nowrap !important;
            scrollbar-width: none !important;
        }

        [data-testid="stTabs"] > div:first-child::-webkit-scrollbar {
            display: none !important;
        }

        /* 移除可能導致所有標籤文字都為白色的選擇器 */
        [data-testid="stTabs"] [data-baseweb="tab"] * {
            color: inherit !important;
        }

        /* 媒體查詢 - 在小螢幕上進一步調整樣式 */
        @media (max-width: 768px) {
            button[data-baseweb="tab"] {
                padding: 6px 2px !important;
                font-size: 11px !important;
                margin: 0 1px !important;
            }
        }
        </style>
        """, unsafe_allow_html=True)

        # 創建頁籤 - 使用圖標提升視覺識別度
        tab1, tab2, tab3, tab4, tab5= st.tabs([
            "📈 檢驗效率監控",
            "🔍 MRB 分析",
            "⏱️ 工作負載監控", 
            "⚖️ 時間分配監控", 
            "📋 額外任務監控",
        ])
        
        with tab1:
            # 使用優化後的效率計算結果
            render_efficiency_dashboard(st.session_state.efficiency_data, st.session_state.processed_data)
        
        with tab2:
            # 顯示新的MRB分析儀表板
            render_mrb_analysis_dashboard(st.session_state.processed_data, 
                                         st.session_state.get('start_date'), 
                                         st.session_state.get('end_date'))
        
        with tab3:
            render_workload_dashboard(st.session_state.workload_data)
        
        with tab4:
            render_time_allocation_dashboard(st.session_state.time_allocation_data)
        
        with tab5:
            render_additional_tasks_dashboard(st.session_state.additional_tasks_monitor_data)
            
    else:        
        # 將數字圓圈改為灰色調風格
        # 標題
        st.markdown("## 📋 使用步驟")

        # 使用Streamlit columns佈局控制寬度
        # 創建兩列，左列佔70%（放卡片），右列佔30%（空白）
        left_col, right_col = st.columns([7, 3])

        with left_col:
            # 步驟1 - 將圓圈顏色改為灰色調 (#78909C)
            st.markdown("""
            <div style="background:#f8f9fa; border-radius:10px; padding:15px; margin-bottom:15px; display:flex; align-items:flex-start;">
            <div style="flex-shrink:0; margin-right:15px;">
                <div style="background:#78909C; color:white; border-radius:50%; width:36px; height:36px; text-align:center; line-height:36px; font-weight:bold;">1</div>
            </div>
            <div style="flex-grow:1;">
                <div style="font-weight:bold; margin-bottom:5px;">上傳Excel檔案</div>
                <div style="color:#666; font-size:14px;">在左側資料上傳區域拖放或選擇檔案</div>
            </div>
            <div style="flex-shrink:0;">
            </div>
            </div>
            """, unsafe_allow_html=True)
            
            # 步驟2
            st.markdown("""
            <div style="background:#f8f9fa; border-radius:10px; padding:15px; margin-bottom:15px; display:flex; align-items:flex-start;">
            <div style="flex-shrink:0; margin-right:15px;">
                <div style="background:#78909C; color:white; border-radius:50%; width:36px; height:36px; text-align:center; line-height:36px; font-weight:bold;">2</div>
            </div>
            <div style="flex-grow:1;">
                <div style="font-weight:bold; margin-bottom:5px;">處理資料</div>
                <div style="color:#666; font-size:14px;">點擊「處理資料」按鈕開始分析</div>
            </div>
            <div style="flex-shrink:0;">
            </div>
            </div>
            """, unsafe_allow_html=True)
            
            # 步驟3
            st.markdown("""
            <div style="background:#f8f9fa; border-radius:10px; padding:15px; margin-bottom:15px; display:flex; align-items:flex-start;">
            <div style="flex-shrink:0; margin-right:15px;">
                <div style="background:#78909C; color:white; border-radius:50%; width:36px; height:36px; text-align:center; line-height:36px; font-weight:bold;">3</div>
            </div>
            <div style="flex-grow:1;">
                <div style="font-weight:bold; margin-bottom:5px;">篩選資料</div>
                <div style="color:#666; font-size:14px;">選擇日期範圍進行資料篩選</div>
            </div>
            <div style="flex-shrink:0;">
            </div>
            </div>
            """, unsafe_allow_html=True)
            
            # 步驟4
            st.markdown("""
            <div style="background:#f8f9fa; border-radius:10px; padding:15px; margin-bottom:15px; display:flex; align-items:flex-start;">
            <div style="flex-shrink:0; margin-right:15px;">
                <div style="background:#78909C; color:white; border-radius:50%; width:36px; height:36px; text-align:center; line-height:36px; font-weight:bold;">4</div>
            </div>
            <div style="flex-grow:1;">
                <div style="font-weight:bold; margin-bottom:5px;">查看分析結果</div>
                <div style="color:#666; font-size:14px;">切換標籤頁查看不同的分析圖表</div>
            </div>
            <div style="flex-shrink:0;">
            </div>
            </div>
            """, unsafe_allow_html=True)
            
            # 步驟5
            st.markdown("""
            <div style="background:#f8f9fa; border-radius:10px; padding:15px; margin-bottom:15px; display:flex; align-items:flex-start;">
            <div style="flex-shrink:0; margin-right:15px;">
                <div style="background:#78909C; color:white; border-radius:50%; width:36px; height:36px; text-align:center; line-height:36px; font-weight:bold;">5</div>
            </div>
            <div style="flex-grow:1;">
                <div style="font-weight:bold; margin-bottom:5px;">匯出報告</div>
                <div style="color:#666; font-size:14px;">下載分析結果為Excel檔案</div>
            </div>
            <div style="flex-shrink:0;">
            </div>
            </div>
            """, unsafe_allow_html=True)
            
            # 自定義提示框，改用灰色調而非藍色的 st.info()
            st.markdown("""
            <div style="background-color:#f0f2f5;  border-radius:4px; padding:16px; margin-top:20px; display:flex; align-items:center;">
                <div style="margin-right:12px; font-size:20px;">👈</div>
                <div>
                    <p style="margin:0; color:#37474F; font-weight:400;">請在左側上傳Excel檔案以開始分析，支援IQC Report、PCB建檔明細、標準工時對應表及額外任務紀錄</p>
                </div>
            </div>
            """, unsafe_allow_html=True)

        # 右列保持空白，作為寬度限制
        with right_col:
            # 保持空白，不顯示任何內容
            pass
            
    # 顯示調試信息
    #if not st.session_state.get('performance_mode', False):
        #render_debug_info()
    
    # 添加頁尾
    st.markdown("""
    <div style="text-align: center; margin-top: 60px; padding: 20px; color: #718096; font-size: 15px; border-top: 1px solid #EDF2F7;">
        <p>IQC 效率管理系統 | © 2025 SQM Joye</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    # 确保session_state中有show_cover
    if 'show_cover' not in st.session_state:
        st.session_state.show_cover = True
        
    # 根据session_state显示封面或主程序
    if st.session_state.show_cover:
        show_cover_page()
    else:
        # 测试模式：提供返回封面的按钮
        #if st.sidebar.button("返回封面頁"):
            #st.session_state.show_cover = True
            #st.rerun()
        
        # 运行主程序
        main()

