import pandas as pd
import streamlit as st
from datetime import datetime, timedelta
import re
from collections import defaultdict

# ======================
# 1. 核心解析逻辑 (保持稳定)
# ======================
def excel_to_date_serial(excel_num):
    if not isinstance(excel_num, (int, float)) or excel_num < 40000:
        return None
    if excel_num < 61: excel_num -= 1
    base_date = datetime(1899, 12, 30)
    try:
        target_date = base_date + timedelta(days=float(excel_num))
        if 2020 <= target_date.year <= 2030:
            return target_date.strftime("%Y-%m-%d")
    except: pass
    return None

def robust_date_parser(cell_value):
    if pd.isna(cell_value): return None
    if isinstance(cell_value, (int, float)):
        return excel_to_date_serial(cell_value)
    if isinstance(cell_value, str):
        text = cell_value.strip().replace(" ", "")
        cn_match = re.search(r'(\d{1,2})月(\d{1,2})日', text)
        if cn_match:
            try:
                month, day = int(cn_match.group(1)), int(cn_match.group(2))
                year = 2026 
                return datetime(year, month, day).strftime("%Y-%m-%d")
            except: pass
        for fmt in ["%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y年%m月%d日"]:
            try:
                dt = datetime.strptime(text, fmt)
                if dt.year < 100: dt = dt.replace(year=dt.year + 2000)
                return dt.strftime("%Y-%m-%d")
            except: continue
    return None

def find_all_date_rows(df, min_dates=2):
    date_rows, date_maps = [], []
    for row_idx in range(len(df)):
        row = df.iloc[row_idx]
        date_map = {col_idx: robust_date_parser(cell) for col_idx, cell in enumerate(row) if robust_date_parser(cell)}
        if len(date_map) >= min_dates:
            date_rows.append(row_idx); date_maps.append(date_map)
    if not date_rows and min_dates > 1:
        return find_all_date_rows(df, min_dates=1)
    return date_rows, date_maps

def extract_tasks_by_date_blocks(df, date_rows, date_maps):
    tasks = defaultdict(str)
    sorted_blocks = sorted(zip(date_rows, date_maps), key=lambda x: x[0])
    for i, (date_row, date_map) in enumerate(sorted_blocks):
        start = date_row + 1
        end = len(df) if i == len(sorted_blocks)-1 else sorted_blocks[i+1][0]
        for col_idx, date_str in date_map.items():
            content = [str(df.iloc[r, col_idx]).strip() for r in range(start, end) 
                      if r < len(df) and col_idx < len(df.columns) and pd.notna(df.iloc[r, col_idx]) and str(df.iloc[r, col_idx]).strip() not in ["", "nan"]]
            if content:
                tasks[date_str] += ("\n\n" if tasks[date_str] else "") + "\n\n".join(content)
    return dict(tasks)

def parse_excel_file(file_obj, filename=""):
    try:
        df = pd.read_excel(file_obj, header=None, engine='openpyxl')
        df = df.dropna(how='all', axis=0).dropna(how='all', axis=1)
        if df.empty: return None, {}, ""
        student_name = filename.split('.')[0].replace("学习计划", "").replace("同学", "").strip()
        date_rows, date_maps = find_all_date_rows(df, min_dates=2)
        if not date_rows: return student_name, {}, "未找到日期"
        tasks = extract_tasks_by_date_blocks(df, date_rows, date_maps)
        return student_name, tasks, "Success"
    except Exception as e: return None, {}, str(e)

# ======================
# 2. 响应式 UI 设计 (Apple Theme - Dual Mode Support)
# ======================
st.set_page_config(page_title="Plan", page_icon="🍎", layout="wide")

# CSS 适配 Light/Dark 模式
st.markdown("""
    <style>
    /* 核心文字与背景适配 */
    html, body, [data-testid="stAppViewContainer"] {
        font-family: "SF Pro Display", -apple-system, BlinkMacSystemFont, sans-serif;
    }
    
    /* 标题样式：跟随主题颜色 */
    h1 {
        font-weight: 700 !important;
        letter-spacing: -0.03em !important;
        padding-bottom: 0px !important;
    }
    
    /* 侧边栏微调 */
    [data-testid="stSidebar"] {
        border-right: 1px solid rgba(128, 128, 128, 0.1);
    }

    /* 日期标签：Apple 经典的次级文本感 */
    .date-label {
        color: #86868b;
        font-size: 0.85rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        margin-top: 2rem;
    }

    /* 针对 Dark Mode 的 Code Block 深度美化 */
    /* 移除背景色，改用透明玻璃感，使复制按钮对比度更高 */
    div[data-testid="stCodeBlock"] {
        border-radius: 18px !important;
        border: 1px solid rgba(128, 128, 128, 0.2) !important;
        background-color: rgba(128, 128, 128, 0.05) !important;
        padding: 10px !important;
    }
    
    code {
        font-family: "SF Pro Text", ui-monospace, sans-serif !important;
        font-size: 1.05rem !important;
        line-height: 1.7 !important;
    }

    /* 按钮：统一为 Apple 蓝 */
    .stButton>button {
        border-radius: 20px;
        padding: 0.4rem 1.5rem;
        background-color: #0071e3;
        color: white !important;
        border: none;
    }

    /* 隐藏部分 UI 干扰项 */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)

if 'all_students' not in st.session_state: st.session_state.all_students = {}

# --- 侧边栏 ---
with st.sidebar:
    st.markdown("### 计划管理")
    uploaded_files = st.file_uploader("上传 Excel", type=["xlsx"], accept_multiple_files=True, label_visibility="collapsed")
    
    if st.button("更新解析内容", type="primary", use_container_width=True):
        if uploaded_files:
            for f in uploaded_files:
                name, tasks, _ = parse_excel_file(f, f.name)
                if name and tasks:
                    if name in st.session_state.all_students:
                        for d, t in tasks.items():
                            st.session_state.all_students[name][d] = t
                    else:
                        st.session_state.all_students[name] = tasks
            st.rerun()
    
    if st.button("重置系统", use_container_width=True):
        st.session_state.all_students = {}
        st.rerun()

# --- 主界面 ---
if not st.session_state.all_students:
    st.markdown("<div style='height: 35vh'></div>", unsafe_allow_html=True)
    st.markdown("<h1 style='text-align: center; opacity: 0.3;'>Learning Plan</h1>", unsafe_allow_html=True)
else:
    # 顶部导航选择区 (自动适配宽屏)
    col_search, col_name, col_date = st.columns([1, 1, 1])
    
    with col_search:
        q = st.text_input("快速搜索学生", placeholder="键入姓名...")
    
    all_n = sorted(st.session_state.all_students.keys())
    filtered_n = [n for n in all_n if q.lower() in n.lower()] if q else all_n

    with col_name:
        sel_name = st.selectbox("学生姓名", filtered_n if filtered_n else ["无匹配项"])

    with col_date:
        if sel_name in st.session_state.all_students:
            dates = sorted(st.session_state.all_students[sel_name].keys(), reverse=True)
            today = datetime.now().strftime("%Y-%m-%d")
            idx = dates.index(today) if today in dates else 0
            sel_date = st.selectbox("计划日期", dates, index=idx)
        else:
            sel_date = None

    # 内容展示
    if sel_name and sel_date and sel_name != "无匹配项":
        task_text = st.session_state.all_students[sel_name].get(sel_date, "")
        
        # 布局层次
        st.markdown(f"<div class='date-label'>{sel_date}</div>", unsafe_allow_html=True)
        st.markdown(f"<h1>{sel_name}</h1>", unsafe_allow_html=True)
        
        # 核心交互区：一键复制与展示
        st.markdown("---")
        st.caption("提示：鼠标悬停在下方区域，点击右上角图标即可一键复制。")
        
        # 使用 st.code 并去掉语言着色，以保持纯文本美感
        st.code(task_text, language=None)

        # 辅助功能
        st.markdown("<div style='margin-top: 1rem;'></div>", unsafe_allow_html=True)
        with st.expander("辅助操作"):
            st.download_button("导出为 .txt 文件", task_text, file_name=f"{sel_name}_{sel_date}.txt")
            st.text_area("若复制按钮失效，请手动在此复制", value=task_text, height=200)
