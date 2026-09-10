import os
import re
import sqlite3
import time
import pandas as pd
import openpyxl
import streamlit as st
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from io import BytesIO

# 注册中文字体，避免PDF乱码
pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))

# ==================== 1. 页面配置 ====================
st.set_page_config(
    page_title="智能网联汽车跨国数据合规平台",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# 初始化 Session State (账号记忆与登录状态)
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "user_identity" not in st.session_state:
    st.session_state.user_identity = None
if "auth_mode" not in st.session_state:
    st.session_state.auth_mode = "login"  # login / register
if "login_type" not in st.session_state:
    st.session_state.login_type = "account"  # account / phone
if "mock_code" not in st.session_state:
    st.session_state.mock_code = ""
if "code_send_time" not in st.session_state:
    st.session_state.code_send_time = 0
if "nav_choice" not in st.session_state:
    st.session_state.nav_choice = "首页"
if "show_terms_page" not in st.session_state:
    st.session_state.show_terms_page = False
if "selected_case" not in st.session_state:
    st.session_state.selected_case = None
if "selected_laws" not in st.session_state:
    st.session_state.selected_laws = []
if "highlighted_case" not in st.session_state:
    st.session_state.highlighted_case = None

# 模拟用户数据库 (支持注册记忆)
if "user_db" not in st.session_state:
    st.session_state.user_db = {
        "admin": "123456",
        "13800138000": "123456"
    }

# ==================== 2. 全局 CSS 样式与 UI 设计系统 (白底黑线风格) ====================
NEWSPRINT_CSS = """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500&family=Playfair+Display:ital,wght@0,400;0,600;0,700;0,900;1,400&family=Lora:ital,wght@0,400;0,600;1,400&display=swap');
    
    :root {
        --bg-base: #F9F9F7;
        --bg-surface: #FFFFFF;
        --text-primary: #111111;
        --text-muted: #666666;
        --border-color: #111111;
        --accent-red: #CC0000;
        --divider-grey: #E5E5E0;
    }
    
    /* 页面基础背景 */
    [data-testid="stAppViewContainer"] {
        background-color: var(--bg-base) !important;
        background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='4' height='4' viewBox='0 0 4 4'%3E%3Cpath fill='%23111111' fill-opacity='0.04' d='M1 3h1v1H1V3zm2-2h1v1H3V1z'%3E%3C/path%3E%3C/svg%3E") !important;
        color: var(--text-primary) !important;
    }
    
    [data-testid="stHeader"] {
        background-color: transparent !important;
    }
    
    /* 全局字体定义 */
    html, body, p, label, li, .law-content {
        font-family: 'Lora', Georgia, serif;
        color: var(--text-primary);
    }
    
    /* 标题统合 */
    h1, h2, h3, h4 {
        font-family: 'Playfair Display', 'Times New Roman', serif !important;
        color: #111111 !important;
        font-weight: 900 !important;
        letter-spacing: -0.03em;
        border-bottom: 2px solid #111111;
        padding-bottom: 8px;
        margin-bottom: 20px;
    }
    
    /* 绝对零圆角卡片 */
    .sharp-card, div[data-testid="stExpander"], .term-card, .timeline-card, .header-card {
        background-color: #FFFFFF !important;
        border: 1px solid #111111 !important;
        border-radius: 0px !important;
        box-shadow: none !important;
        transition: all 150ms cubic-bezier(0, 0, 0.2, 1);
        padding: 24px;
        margin-bottom: 20px;
    }
    .hard-shadow-hover:hover {
        box-shadow: 4px 4px 0px 0px #111111 !important;
        transform: translate(-2px, -2px);
    }
    
    /* 案例卡片高亮选中效果 */
    .case-card-selected {
        border: 2px solid #CC0000 !important;
        box-shadow: 4px 4px 0px 0px #CC0000 !important;
        background-color: #FFFFFF !important;
    }
    
    /* Expander 折叠面板样式：全部改为白底黑框 */
    div[data-testid="stExpander"] { padding: 0 !important; background-color: #FFFFFF !important; }
    div[data-testid="stExpander"] summary {
        padding: 16px 20px;
        background-color: #FFFFFF !important;
        border: 1px solid #111111 !important;
        border-radius: 0px !important;
        color: #111111 !important;
    }
    div[data-testid="stExpander"] summary span[data-testid="stExpanderToggleIcon"] {
        font-size: 0px !important;
        color: transparent !important;
        width: 0px !important;
        display: none !important;
    }
    div[data-testid="stExpander"] summary:hover {
        background-color: #F0F0EB !important;
        color: #111111 !important;
    }
    div[data-testid="stExpander"] summary p {
        font-family: 'Playfair Display', serif !important;
        font-weight: 700;
        color: #111111 !important;
        margin: 0 !important;
    }
    
    /* 隐藏默认侧边栏 */
    [data-testid="stSidebar"] {
        display: none !important;
    }
    
    /* 顶端导航链接 (白底黑框) */
    .nav-tabs-container {
        display: flex;
        border: 1px solid #111111;
        background-color: #FFFFFF;
        margin-bottom: 25px;
    }
    .nav-tab-item {
        flex: 1;
        text-align: center;
        border-right: 1px solid #111111;
    }
    .nav-tab-item:last-child {
        border-right: none;
    }
    .nav-tab-item button {
        background-color: #FFFFFF !important;
        color: #111111 !important;
        border: 1px solid #111111 !important;
        border-radius: 0px !important;
        font-family: 'Inter', sans-serif !important;
        font-weight: 700 !important;
        font-size: 0.9rem !important;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        box-shadow: none !important;
        width: 100% !important;
        padding: 12px 0px !important;
        margin: 0px !important;
        text-align: center;
        transition: all 100ms ease !important;
    }
    .nav-tab-item button:hover {
        background-color: #F0F0EB !important;
        color: #111111 !important;
    }
    .nav-tab-active button {
        background-color: #FFFFFF !important;
        color: #CC0000 !important;
        border: 2px solid #111111 !important;
        font-weight: 900 !important;
    }
    
    /* 目录列表样式 (白底黑框) */
    .case-dir-container {
        border: 1px solid #111111;
        background-color: #FFFFFF;
        margin-bottom: 25px;
    }
    .case-dir-item {
        border-bottom: 1px solid #111111;
    }
    .case-dir-item:last-child {
        border-bottom: none;
    }
    .case-dir-item button {
        background-color: #FFFFFF !important;
        color: #111111 !important;
        border: none !important;
        border-radius: 0px !important;
        font-family: 'Lora', Georgia, serif !important;
        font-weight: 600 !important;
        font-size: 0.95rem !important;
        box-shadow: none !important;
        width: 100% !important;
        padding: 10px 16px !important;
        margin: 0px !important;
        text-align: left !important;
        transition: all 100ms ease !important;
    }
    .case-dir-item button:hover {
        background-color: #F0F0EB !important;
        color: #111111 !important;
    }
    /* 术语按钮样式 */
    .inline-term-btn button {
        background-color: #FFFFFF !important;
        color: #111111 !important;
        border: 1px solid #111111 !important;
        border-radius: 0px !important;
        font-family: 'Inter', sans-serif !important;
        font-weight: 600 !important;
        font-size: 0.85rem !important;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        box-shadow: none !important;
        padding: 8px 16px !important;
        white-space: nowrap;
    }
    .inline-term-btn button:hover {
        background-color: #F0F0EB !important;
        color: #111111 !important;
        border: 1px solid #111111 !important;
    }
    
    /* 报纸风格标签 (白底黑边) */
    .law-tag {
        display: inline-block;
        background-color: #FFFFFF;
        color: #111111;
        padding: 4px 10px;
        border-radius: 0px !important;
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.75rem;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        margin-bottom: 4px;
        margin-right: 6px;
        border: 1px solid #111111;
        font-weight: bold;
    }
    
    /* 条款排版容器 */
    .law-content {
        background-color: #FFFFFF;
        border: 1px solid #111111;
        border-left: 6px solid #111111 !important;
        padding: 20px;
        color: #111111;
        line-height: 1.8;
        font-size: 1rem;
        text-align: justify;
        border-radius: 0px !important;
    }
    
    /* 术语卡片 */
    .term-card {
        border-left: 6px solid var(--accent-red) !important;
        background-color: #FFFFFF !important;
    }
    .term-source {
        font-family: 'JetBrains Mono', monospace;
        color: #666666;
        font-size: 0.8rem;
        margin-top: 15px;
        text-align: right;
    }
    
    /* 时间轴 */
    .timeline-container {
        position: relative;
        padding-left: 30px;
        margin: 30px 0;
        border-left: 2px solid #111111;
    }
    .timeline-item {
        position: relative;
        margin-bottom: 35px;
    }
    .timeline-node {
        position: absolute;
        left: -37px;
        top: 18px;
        width: 12px;
        height: 12px;
        border-radius: 0px !important;
        background-color: #FFFFFF;
        border: 3px solid #111111;
    }
    
    /* 输入框样式 (白底黑边) */
    div[data-testid="stTextInput"] input, div[data-testid="stSelectbox"] div[data-baseweb="select"] {
        background-color: #FFFFFF !important;
        border: 1px solid #111111 !important;
        color: #111111 !important;
        font-family: 'JetBrains Mono', monospace !important;
        border-radius: 0px !important;
        box-shadow: none !important;
    }
    div[data-testid="stTextInput"] input::placeholder {
        color: #666666 !important;
        opacity: 1 !important;
    }
    div[data-testid="stTextInput"] input:focus {
        background-color: #FFFFFF !important;
        color: #000000 !important;
        border: 2px solid #111111 !important;
    }
    div[data-testid="stTextInput"] label, div[data-testid="stSelectbox"] label {
        font-family: 'Inter', sans-serif !important;
        font-weight: 700 !important;
        text-transform: uppercase;
        font-size: 0.8rem;
        color: #111111 !important;
    }
    
    /* Tabs 选项卡样式修改：彻底取消黑色，改为白底黑框 */
    button[data-baseweb="tab"] {
        background-color: #FFFFFF !important;
        color: #111111 !important;
        border: 1px solid #111111 !important;
        border-radius: 0px !important;
        font-weight: 600 !important;
        margin-right: 4px !important;
    }
    button[data-baseweb="tab"][aria-selected="true"] {
        background-color: #FFFFFF !important;
        color: #CC0000 !important;
        border: 2px solid #111111 !important;
        font-weight: bold !important;
    }
    
    /* 按钮通用重置：白底黑边框黑字 */
    div.stButton > button {
        border-radius: 0px !important;
        border: 1px solid #111111 !important;
        font-family: 'Inter', sans-serif !important;
        font-weight: 700 !important;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        background-color: #FFFFFF !important;
        color: #111111 !important;
        box-shadow: none !important;
        transition: all 100ms ease !important;
    }
    div.stButton > button:hover {
        background-color: #F0F0EB !important;
        color: #111111 !important;
        border: 1px solid #111111 !important;
        box-shadow: 2px 2px 0px 0px #111111 !important;
    }
    
    /* “还没有账号？”文本样式 */
    .register-hover-text {
        font-size: 0.85rem;
        line-height: 2.2;
        color: #111111;
        cursor: pointer;
        transition: color 200ms ease, font-weight 200ms ease;
    }
    .register-hover-text:hover {
        color: #CC0000 !important;
        font-weight: 700;
        text-decoration: underline;
    }
    
    /* 关于我们悬停变红效果 */
    .about-feedback-hover {
        color: #333333;
        transition: color 200ms ease;
    }
    .about-feedback-hover:hover {
        color: #CC0000 !important;
        font-weight: bold;
    }

    /* 报头元数据 */
    .newsprint-masthead {
        border-top: 3px solid #111111;
        border-bottom: 1px solid #111111;
        padding: 8px 0;
        margin-bottom: 24px;
        display: flex;
        justify-content: space-between;
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.75rem;
        text-transform: uppercase;
        letter-spacing: 0.1em;
    }
    /* 顶部标题栏样式 */
    .top-header-bar {
        display: flex;
        justify-content: space-between;
        align-items: center;
        border-bottom: 2px solid #111111;
        padding-bottom: 12px;
        margin-bottom: 20px;
    }
    .top-header-title {
        font-family: 'Playfair Display', serif;
        font-size: 1.4rem;
        font-weight: 900;
        letter-spacing: -0.02em;
    }
</style>
"""
st.markdown(NEWSPRINT_CSS, unsafe_allow_html=True)

DB_FILE = "car_compliance.db"

# 密码校验逻辑：允许大小写字母、数字，以及仅限 “_”、“@”、“*” 三种特殊字符
def is_valid_password(pwd):
    pattern = r"^[A-Za-z0-9_@*]+$"
    return bool(re.match(pattern, pwd))

# ==================== 3. 登录与注册模块 (包含大小写与特殊符号识别) ====================
def render_auth_page():
    # 顶部品牌标识
    st.markdown(
        """
        <div class="top-header-bar">
            <div class="top-header-title">智能网联汽车跨国数据合规平台</div>
            <div style="font-family: 'JetBrains Mono', monospace; font-size: 0.8rem; text-transform: uppercase;">系统身份认证</div>
        </div>
        """,
        unsafe_allow_html=True
    )
    
    st.write("")
    st.write("")
    
    # 三列居中布局
    col1, col2, col3 = st.columns([1, 1.3, 1])
    
    with col2:
        st.markdown(
            """
            <div class="sharp-card" style="border-top: 4px solid #111111; padding: 30px; background-color: #FFFFFF;">
            """,
            unsafe_allow_html=True
        )
        
        if st.session_state.auth_mode == "login":
            st.markdown("<h2 style='text-align: center; border-bottom: 1px solid #111; padding-bottom: 10px; margin-bottom: 20px; font-size: 1.8rem;'>用户登录</h2>", unsafe_allow_html=True)
            
            # 登录方式切换
            tab_acc, tab_phone = st.tabs(["账号密码登录", "手机号一键登录"])
            
            with tab_acc:
                st.write("")
                acc_input = st.text_input("账号 / 用户名", key="login_acc_input", placeholder="输入用户名...")
                pwd_input = st.text_input("密码", type="password", key="login_pwd_input", placeholder="输入密码...")
                st.write("")
                if st.button("立即登录", key="btn_login_acc", use_container_width=True):
                    if not acc_input or not pwd_input:
                        st.error("请输入账号和密码")
                    elif acc_input in st.session_state.user_db and st.session_state.user_db[acc_input] == pwd_input:
                        st.session_state.authenticated = True
                        st.session_state.user_identity = acc_input
                        st.rerun()
                    else:
                        st.error("账号或密码错误（请注意区分字母大小写）")
            
            with tab_phone:
                st.write("")
                phone_input = st.text_input("手机号码", key="login_phone_input", placeholder="输入11位手机号...")
                
                code_col1, code_col2 = st.columns([1.5, 1])
                with code_col1:
                    code_input = st.text_input("验证码", key="login_code_input", placeholder="6位验证码...")
                with code_col2:
                    st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
                    now = time.time()
                    time_passed = now - st.session_state.code_send_time
                    if time_passed < 60:
                        st.button(f"{int(60 - time_passed)}s 后重发", disabled=True, key="btn_send_code_disabled", use_container_width=True)
                    else:
                        if st.button("获取验证码", key="btn_send_code", use_container_width=True):
                            if not re.match(r"^1[3-9]\d{9}$", phone_input):
                                st.error("请输入有效的手机号码")
                            else:
                                import random
                                st.session_state.mock_code = str(random.randint(100000, 999999))
                                st.session_state.code_send_time = time.time()
                                st.rerun()
                
                if st.session_state.mock_code and (time.time() - st.session_state.code_send_time < 60):
                    st.info(f"验证码已发送（测试提示：{st.session_state.mock_code}）")
                
                st.write("")
                if st.button("手机号一键登录", key="btn_login_phone", use_container_width=True):
                    if not phone_input or not code_input:
                        st.error("请输入手机号和验证码")
                    elif code_input == st.session_state.mock_code and st.session_state.mock_code != "":
                        st.session_state.authenticated = True
                        st.session_state.user_identity = phone_input
                        st.session_state.mock_code = ""
                        st.rerun()
                    else:
                        st.error("验证码不正确或已失效")
            
            st.divider()
            switch_col1, switch_col2 = st.columns([1.2, 1])
            with switch_col1:
                st.markdown("<p class='register-hover-text'>还没有账号？</p>", unsafe_allow_html=True)
            with switch_col2:
                if st.button("注册账号", key="goto_register_btn", use_container_width=True):
                    st.session_state.auth_mode = "register"
                    st.rerun()
        else:
            st.markdown("<h2 style='text-align: center; border-bottom: 1px solid #111; padding-bottom: 10px; margin-bottom: 20px; font-size: 1.8rem;'>注册账号</h2>", unsafe_allow_html=True)
            reg_user = st.text_input("设置用户名 / 手机号", key="reg_user_input", placeholder="请输入用户名或手机号...")
            reg_pwd = st.text_input("设置密码 (仅支持“_”“@”“*”三种特殊字符)", type="password", key="reg_pwd_input", placeholder="请输入包含大小写、数字及允许字符的密码...")
            reg_pwd_confirm = st.text_input("确认密码", type="password", key="reg_pwd_confirm_input", placeholder="请再次输入密码...")
            st.write("")
            if st.button("完成注册并登录", key="btn_register_submit", use_container_width=True):
                if not reg_user or not reg_pwd:
                    st.error("用户名和密码不能为空")
                elif not is_valid_password(reg_pwd):
                    st.error("密码包含不支持的字符！仅支持大小写字母、数字及“_”“@”“*”三种特殊字符。")
                elif reg_pwd != reg_pwd_confirm:
                    st.error("两次输入的密码不一致")
                elif reg_user in st.session_state.user_db:
                    st.error("该账号已被注册")
                else:
                    # 记住账号与密码，并直接登录
                    st.session_state.user_db[reg_user] = reg_pwd
                    st.session_state.authenticated = True
                    st.session_state.user_identity = reg_user
                    st.session_state.auth_mode = "login"
                    st.success("注册成功！正在登录...")
                    time.sleep(0.5)
                    st.rerun()
            
            st.divider()
            if st.button("返回登录界面", key="goto_login_btn", use_container_width=True):
                st.session_state.auth_mode = "login"
                st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

# 未通过认证则拦截并渲染登录页
if not st.session_state.authenticated:
    render_auth_page()
    st.stop()

# ==================== 4. 核心处理与数据库函数 ====================
def extract_sort_key(text):
    match_cn = re.search(r"第([零一二三四五六七八九十百0-9]+)条", text)
    if match_cn:
        num_str = match_cn.group(1)
        mapping = {"一":1, "二":2, "三":3, "四":4, "五":5, "六":6, "七":7, "八":8, "九":9, "十":10,
                   "十一":11, "十二":12, "十三":13, "十四":14, "十五":15, "十六":16, "十七":17, "十八":18, "十九":19, "二十":20,
                   "二十一":21, "二十二":22, "二十三":23, "二十四":24, "二十五":25, "二十六":26, "二十七":27, "二十八":28, "二十九":29, "三十":30,
                   "三十一":31, "三十二":32, "三十三":33, "三十四":34, "三十五":35, "三十六":36, "三十七":37, "三十八":38, "三十九":39, "四十":40}
        if num_str in mapping: return mapping[num_str]
        try: return int(num_str)
        except ValueError: pass
    match_en = re.search(r"Article\s+(\d+)", text, re.IGNORECASE)
    if match_en:
        try: return int(match_en.group(1))
        except ValueError: pass
    return 999

def extract_article_number(text):
    """提取法条文本中的条文编号（如第xx条、Article xx等）"""
    if not text:
        return ""
    match_cn = re.search(r"第[零一二三四五六七八九十百千0-9]+条", text)
    if match_cn:
        return match_cn.group(0)
    match_en = re.search(r"(Article\s+\d+|Art\.\s*\d+|Recital\s+\d+|Section\s+\d+)", text, re.IGNORECASE)
    if match_en:
        return match_en.group(0)
    return ""

def parse_fine_amount(text):
    """提取“1、罚款：”背后的数字大小用于排序"""
    if not text or text == "（暂无内容）":
        return -1.0
    match = re.search(r"1[、:]\s*罚款[：:]\s*([0-9.]+)\s*([万亿]*)\s*(欧元|人民币|美元|英镑)?", text)
    if match:
        num = float(match.group(1))
        unit = match.group(2)
        if unit == '亿':
            num *= 100000000
        elif unit == '万':
            num *= 10000
        return num
    return -1.0

def get_clean_cell_text(cell):
    if cell.value is None or str(cell.value).strip() == "nan":
        return ""
    is_rich = hasattr(cell, 'value') and isinstance(cell.value, openpyxl.cell.rich_text.CellRichText)
    if is_rich:
        full_text = "".join([str(rt.text) for rt in cell.value])
    else:
        full_text = str(cell.value)
    full_text = re.sub(r'(?im)^\s*svg\s*$', '', full_text)
    return full_text.strip()

def clean_scenario_cell(val):
    if pd.isna(val) or val is None:
        return ""
    s = str(val).strip()
    if s.lower() in ["nan", "none", "null", "unnamed: 0", "unnamed: 1", "unnamed: 2"]:
        return ""
    return s

@st.cache_data
def init_database_from_excel():
    current_dir = os.path.dirname(os.path.abspath(__file__))
    possible_names = [
        "合规平台条文整理（修改4.0）_4.xlsx",
        "合规平台条文整理（修改4.0）.xlsx",
        "合规平台条文整理（修改1.0）.xlsx",
        "合规平台条文整理 (1).xlsx"
    ]
    excel_path = None
    for fname in possible_names:
        p = os.path.join(current_dir, fname)
        if os.path.exists(p):
            excel_path = p
            break
    if not excel_path:
        for fname in os.listdir(current_dir):
            if fname.endswith(".xlsx") and "合规平台" in fname:
                excel_path = os.path.join(current_dir, fname)
                break
    if not excel_path or not os.path.exists(excel_path):
        return False
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS compliance_laws (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            region TEXT,
            category TEXT,
            law_title TEXT,
            sub_cat_0 TEXT,
            sub_cat_1 TEXT,
            content TEXT,
            sort_order INTEGER
        )
    """)
    cursor.execute("DELETE FROM compliance_laws")
    wb = openpyxl.load_workbook(excel_path, data_only=False)
    sheet_name = wb.sheetnames[0]
    ws = wb[sheet_name]
    df_raw = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)
    categories_row = df_raw.iloc[0]
    titles_row = df_raw.iloc[1]
    
    col0_raw = df_raw.iloc[:, 0] if len(df_raw.columns) > 0 else pd.Series([""] * len(df_raw))
    col1_raw = df_raw.iloc[:, 1] if len(df_raw.columns) > 1 else pd.Series([""] * len(df_raw))
    col2_raw = df_raw.iloc[:, 2] if len(df_raw.columns) > 2 else pd.Series([""] * len(df_raw))
    laws_dict = {}
    laws_order = []
    all_law_columns = []
    for col_idx in range(3, len(df_raw.columns)):
        cat_raw = str(categories_row.iloc[col_idx]).strip()
        law_title = str(titles_row.iloc[col_idx]).strip()
        if not law_title or law_title == "nan":
            continue
        if "-" in cat_raw:
            parts = cat_raw.split("-", 1)
            region = parts[0].strip()
            category = parts[1].strip()
        elif "—" in cat_raw:
            parts = cat_raw.split("—", 1)
            region = parts[0].strip()
            category = parts[1].strip()
        elif "–" in cat_raw:
            parts = cat_raw.split("–", 1)
            region = parts[0].strip()
            category = parts[1].strip()
        else:
            if "欧盟" in cat_raw:
                region = "欧盟"
                category = cat_raw.replace("欧盟", "").strip() or "通用模块"
            elif "美国" in cat_raw:
                region = "美国"
                category = cat_raw.replace("美国", "").strip() or "通用模块"
            elif "中国" in cat_raw:
                region = "中国"
                category = cat_raw.replace("中国", "").strip() or "通用模块"
            else:
                region = "中国"
                category = cat_raw if cat_raw and cat_raw != "nan" else "通用模块"
        all_law_columns.append((region, category, law_title))
        for row_idx in range(2, len(df_raw)):
            cell_obj = ws.cell(row=row_idx + 1, column=col_idx + 1)
            content_str = get_clean_cell_text(cell_obj)
            
            if content_str and content_str != "nan":
                s0 = clean_scenario_cell(col0_raw.iloc[row_idx])
                s1 = clean_scenario_cell(col1_raw.iloc[row_idx])
                s2 = clean_scenario_cell(col2_raw.iloc[row_idx]) if len(df_raw.columns) > 2 else ""
                
                sc_parts = [x for x in [s0, s1, s2] if x]
                scenario_text = " -> ".join(sc_parts)
                sort_val = extract_sort_key(content_str)
                
                key = (region, category, law_title, content_str)
                if key not in laws_dict:
                    laws_dict[key] = {
                        "region": region,
                        "category": category,
                        "law_title": law_title,
                        "content": content_str,
                        "sort_order": sort_val,
                        "scenarios": []
                    }
                    laws_order.append(key)
                
                if scenario_text and scenario_text not in laws_dict[key]["scenarios"]:
                    laws_dict[key]["scenarios"].append(scenario_text)
    processed_laws = set()
    for key in laws_order:
        item = laws_dict[key]
        combined_scenarios = " | ".join(item["scenarios"])
        cursor.execute(
            "INSERT INTO compliance_laws (region, category, law_title, sub_cat_0, sub_cat_1, content, sort_order) VALUES (?, ?, ?, ?, ?, ?, ?)",
            (item["region"], item["category"], item["law_title"], combined_scenarios, "", item["content"], item["sort_order"])
        )
        processed_laws.add((item["region"], item["category"], item["law_title"]))
    for (r, c, lt) in set(all_law_columns):
        if (r, c, lt) not in processed_laws:
            cursor.execute(
                "INSERT INTO compliance_laws (region, category, law_title, sub_cat_0, sub_cat_1, content, sort_order) VALUES (?, ?, ?, ?, ?, ?, ?)",
                (r, c, lt, "", "", "（该法规条文正在整理补充中，敬请期待...）", 999)
            )
            
    conn.commit()
    conn.close()
    return True

success_db = init_database_from_excel()

# ==================== 5. 顶端栏与导航选项卡 ====================
top_bar_left, top_bar_right = st.columns([3, 1])
with top_bar_left:
    st.markdown(f"<div class='top-header-title'>智能网联汽车跨国数据合规平台</div>", unsafe_allow_html=True)
with top_bar_right:
    user_disp_col, logout_btn_col = st.columns([1.5, 1])
    with user_disp_col:
        st.markdown(f"<div style='font-family: JetBrains Mono, monospace; font-size: 0.8rem; text-align: right; padding-top: 6px;'>用户: {st.session_state.user_identity}</div>", unsafe_allow_html=True)
    with logout_btn_col:
        if st.button("退出登录", key="btn_logout", use_container_width=True):
            st.session_state.authenticated = False
            st.session_state.user_identity = None
            st.rerun()
st.write("")
nav_items = [
    ("首页", "首页"),
    ("法律库", "法律库"),
    ("出境全流程时间轴", "出境全流程时间轴"),
    ("案例库", "案例库"),
    ("关于我们", "关于我们")
]
top_cols = st.columns(5)
for idx, (label, choice_key) in enumerate(nav_items):
    with top_cols[idx]:
        is_active = (st.session_state.nav_choice == choice_key)
        active_style_class = "nav-tab-active" if is_active else ""
        st.markdown(f'<div class="nav-tab-item {active_style_class}" style="border:none; margin-bottom:15px;">', unsafe_allow_html=True)
        if st.button(label, key=f"nav_btn_{idx}", use_container_width=True):
            st.session_state.nav_choice = choice_key
            st.session_state.show_terms_page = False
            st.session_state.selected_case = None
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

# ==================== 6. 主内容页面展示逻辑 ====================
if st.session_state.show_terms_page:
    st.markdown("<h2 style='text-align: center; border-bottom: 3px solid #111;'>术语解释总结全库专栏</h2>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; font-family: Lora, serif; color: #666666;'>展示完整的汽车数据及出境合规术语释义，还原现代纸媒专栏的严谨审慎与清晰结构。</p>", unsafe_allow_html=True)
    
    term_keyword = st.text_input("检索术语关键字 (如：个人信息、重要数据、GDPR...)", key="standalone_term_search", placeholder="在此输入关键字进行检索...")
    
    current_dir = os.path.dirname(os.path.abspath(__file__))
    term_excel_path = os.path.join(current_dir, "术语解释总结.xlsx")
    
    if os.path.exists(term_excel_path):
        try:
            wb = openpyxl.load_workbook(term_excel_path, data_only=True)
            ws = wb.active
            law_names = []
            for cell in ws[1]:
                law_names.append(str(cell.value).strip() if cell.value else "")
                
            terms_list = []
            for row in ws.iter_rows(min_row=2):
                for c_idx, cell in enumerate(row):
                    if cell.value:
                        cell_str = str(cell.value).strip()
                        if not cell_str or cell_str.lower() == "nan":
                            continue
                        source_law = law_names[c_idx] if c_idx < len(law_names) and law_names[c_idx] else "未知法规"
                        
                        parts = cell_str.split(":", 1)
                        if len(parts) == 2:
                            term_name = parts[0].strip()
                            definition = parts[1].strip()
                        else:
                            term_name = cell_str
                            definition = cell_str
                            
                        color = None
                        if cell.fill and cell.fill.start_color:
                            color_val = cell.fill.start_color.index
                            if color_val and str(color_val) != '00000000':
                                color = str(color_val)
                                
                        terms_list.append({
                            "term_name": term_name,
                            "definition": definition,
                            "source": source_law,
                            "color": color,
                            "original_full": cell_str
                        })
                        
            final_results = []
            if term_keyword:
                term_keyword_lower = term_keyword.lower()
                matched_colors = set()
                direct_match_indices = set()
                for i, t in enumerate(terms_list):
                    if term_keyword_lower in t["term_name"].lower():
                        direct_match_indices.add(i)
                        if t["color"]:
                            matched_colors.add(t["color"])
                for i, t in enumerate(terms_list):
                    if i in direct_match_indices or (t["color"] and t["color"] in matched_colors):
                        final_results.append(t)
                st.markdown(f"**检索到相关术语/条文共计：{len(final_results)} 条**")
            else:
                final_results = terms_list
                st.markdown(f"**当前库内完整术语/条文共计：{len(final_results)} 条**")
                
            for t_item in final_results:
                if t_item['term_name'] != t_item['definition']:
                    def_html = f"<b>{t_item['term_name']}：</b>" + t_item['definition']
                else:
                    def_html = t_item['original_full']
                
                def_html = def_html.replace('\n', '<br>')
                source_text = f"（来源：《{t_item['source']}》）"
                
                html_str = (
                    f'<div class="term-card hard-shadow-hover">'
                    f'<div>{def_html}</div>'
                    f'<div class="term-source">{source_text}</div>'
                    f'</div>'
                )
                st.markdown(html_str, unsafe_allow_html=True)
                
        except Exception as e:
            st.error(f"加载术语表异常: {e}")
    else:
        st.warning("未检测到 `术语解释总结.xlsx` 文件，请确认已上传至同一目录。")
else:
    if st.session_state.nav_choice == "首页":
        st.markdown(
            """
            <div class="newsprint-masthead">
                <span>VOL. I NO. 01</span>
                <span>AUTOMOTIVE DATA COMPLIANCE REVIEW</span>
                <span>GLOBAL EDITION</span>
            </div>
            """,
            unsafe_allow_html=True
        )
        st.markdown(
            """
            <div class="sharp-card" style="border-top: 4px solid #111;">
                <h1 style='margin-top:0; border-bottom:none; font-size: 2.8rem;'>智能网联汽车跨国数据合规平台 - 首页</h1>
                <p style='font-family: Lora, serif; font-size: 1.1rem; line-height: 1.6; color: #333333; margin-bottom: 0;'>
                    智能网联汽车车外实景影像数据跨境流动的双向合规路径研究——以中国与欧盟为例<br>
教育部大学生创新训练计划项目 · 聚焦中欧数据跨境规则冲突与合规路径
                </p>
            </div>
            """, 
            unsafe_allow_html=True
        )
        st.markdown(
            """
            <div style="font-family: Lora, serif; font-size: 1.1rem; line-height: 1.8; color: #333333; margin: 10px 0 30px 0;">
                <b>项目概况：</b><br>
                本项目聚焦智能网联汽车出海欧盟时，车外实景影像数据跨境流动的双向合规困境。中国《汽车数据出境安全指引》与欧盟GDPR在数据定性、出境路径及执法机制上存在显著冲突，导致企业面临高昂合规成本及法律风险。研究采用功能主义比较法，通过规范分析、企业访谈及案例实证，揭示中欧规制差异，并构建分场景双向合规操作框架及数字化风险识别平台。
            </div>
            """,
            unsafe_allow_html=True
        )
    elif st.session_state.nav_choice == "关于我们":
        st.markdown(
            """
            <div class="newsprint-masthead">
                <span>VOL. I NO. 01</span>
                <span>AUTOMOTIVE DATA COMPLIANCE REVIEW</span>
                <span>GLOBAL EDITION</span>
            </div>
            """,
            unsafe_allow_html=True
        )
        st.markdown(
            """
            <h1 style='margin-top:0; border-bottom:none; font-size: 2.8rem;'>关于我们</h1>
            <p style='font-family: Lora, serif; font-size: 1.1rem; line-height: 1.8; color: #333333; margin: 10px 0 30px 0;'>
                汽车数据观察室，是由华东政法大学国际金融法律学院、法律学院、经济法学院、传播学院、商学院五大学院本科生组建的跨学科研究团队。我们聚焦中国智能网联汽车出海欧盟过程中，车外实景影像数据跨境流动面临的中欧法律规制冲突，通过比较法研究、案例实证与企业深度访谈，探索兼顾数据安全与产业发展的双向合规路径，为中国汽车产业的全球化进程提供学术支撑与实践参考。<br><br>
                <span class="about-feedback-hover">遇到问题/提供反馈意见请在公众号@汽车数据观察室后台留言</span>
            </p>
            """, 
            unsafe_allow_html=True
        )
    elif st.session_state.nav_choice == "案例库":
        st.markdown(
            """
            <div class="newsprint-masthead">
                <span>VOL. I NO. 01</span>
                <span>AUTOMOTIVE DATA COMPLIANCE REVIEW</span>
                <span>GLOBAL EDITION</span>
            </div>
            """,
            unsafe_allow_html=True
        )
        
        current_dir = os.path.dirname(os.path.abspath(__file__))
        case_excel_path = os.path.join(current_dir, "案例库.xlsx")
        
        if os.path.exists(case_excel_path):
            try:
                df_case = pd.read_excel(case_excel_path, header=None)
                cases_data = []
                for col_idx in range(1, len(df_case.columns)):
                    case_name = str(df_case.iloc[0, col_idx]).strip()
                    if not case_name or case_name == "nan":
                        continue
                    sections = {}
                    for row_idx in range(1, len(df_case)):
                        sec_title = str(df_case.iloc[row_idx, 0]).strip()
                        sec_content = str(df_case.iloc[row_idx, col_idx]).strip()
                        if sec_title and sec_title != "nan":
                            sections[sec_title] = sec_content if sec_content != "nan" else "（暂无内容）"
                    
                    fine_val = parse_fine_amount(sections.get("处罚结果", ""))
                    cases_data.append({
                        "case_name": case_name,
                        "sections": sections,
                        "fine_amount": fine_val
                    })
                
                cases_data.sort(key=lambda x: x["fine_amount"], reverse=True)
                    
                if st.session_state.selected_case is None:
                    st.markdown(
                        """
                        <div class="sharp-card" style="border-top: 4px solid #111;">
                            <h1 style='margin-top:0; border-bottom:none; font-size: 2.4rem;'>合规典型案例库</h1>
                            <p style='font-family: Lora, serif; font-size: 1rem; line-height: 1.5; color: #333333; margin-bottom: 0;'>
                                典型案例库，收录全球数据合规与跨境执法案件。点击目录案例名称，可跳转至对应案例，查看案件基本信息、事实梳理、法律分析、处罚结果、合规启示和原始资料链接。
                            </p>
                        </div>
                        """, 
                        unsafe_allow_html=True
                    )
                    st.write("")
                    
                    st.markdown("### 案例快速检索目录")
                    st.markdown("<p style='font-size: 0.85rem; color: #666;'>点击下方案例名称可自动定位至对应案例并进行框选高亮：</p>", unsafe_allow_html=True)
                    
                    st.markdown('<div class="case-dir-container">', unsafe_allow_html=True)
                    for i, c_item in enumerate(cases_data):
                        c_name = c_item["case_name"]
                        case_anchor_id = f"case_card_{i}"
                        
                        st.markdown('<div class="case-dir-item">', unsafe_allow_html=True)
                        if st.button(c_name, key=f"dir_btn_{i}", use_container_width=True):
                            st.session_state.highlighted_case = c_name
                            js_code = f"""
                            <script>
                                var element = parent.document.getElementById('{case_anchor_id}');
                                if(element) {{
                                    element.scrollIntoView({{behavior: 'smooth', block: 'center'}});
                                }}
                            </script>
                            """
                            st.components.v1.html(js_code, height=0, width=0)
                        st.markdown('</div>', unsafe_allow_html=True)
                    st.markdown('</div>', unsafe_allow_html=True)
                    
                    st.divider()
                    for i, c_item in enumerate(cases_data):
                        c_name = c_item["case_name"]
                        c_region = c_item["sections"].get("地域", "（暂无）")
                        c_info = c_item["sections"].get("案件基本信息", "每起案例从六个角度拆解：案件背景、事实梳理、GDPR或国内法核心条款、监管逻辑、处罚裁决，以及对出海的启示。")
                        
                        is_highlighted = (st.session_state.highlighted_case == c_name)
                        card_class = "case-card-selected" if is_highlighted else "hard-shadow-hover"
                        case_anchor_id = f"case_card_{i}"
                        
                        st.markdown(
                            f"""
                            <div id="{case_anchor_id}" class="sharp-card {card_class}" style="border-left: 6px solid #111; padding: 20px 24px; margin-bottom: 16px;">
                                <div style="display: flex; justify-content: space-between; align-items: center;">
                                    <h3 style="margin-top: 0; margin-bottom: 10px; font-family: Playfair Display, serif; font-size: 1.5rem; border-bottom: none;">
                                        {c_name}
                                    </h3>
                                    <span class="law-tag">辖区: {c_region}</span>
                                </div>
                                <p style="font-family: Lora, serif; color: #333; margin-bottom: 15px; font-size: 0.95rem; line-height: 1.6;">
                                    <b>【案件基本信息】</b>：{c_info}
                                </p>
                            </div>
                            """,
                            unsafe_allow_html=True
                        )
                        if st.button(f"查看完整案例报告 ->", key=f"case_btn_{i}", use_container_width=True):
                            st.session_state.selected_case = c_name
                            st.rerun()
                else:
                    active_case = next((c for c in cases_data if c["case_name"] == st.session_state.selected_case), None)
                    if st.button("<- 返回案例库列表", key="back_to_cases"):
                        st.session_state.selected_case = None
                        st.rerun()
                        
                    if active_case:
                        st.markdown(f"<h1 style='margin-top: 10px; font-size: 2.5rem;'>{active_case['case_name']}</h1>", unsafe_allow_html=True)
                        sections_order = ["地域", "案件基本信息", "案件基本情况", "法律分析", "处罚结果", "合规启示", "相关资料"]
                        for sec_title in sections_order:
                            if sec_title in active_case["sections"]:
                                content_val = active_case["sections"][sec_title]
                                st.markdown(f"<h3 style='font-family: Playfair Display, serif; margin-top: 25px; border-bottom: 1px solid #111;'>{sec_title}</h3>", unsafe_allow_html=True)
                                safe_content_val = content_val.replace('\n', '<br>')
                                st.markdown(f'<div class="law-content" style="white-space: pre-wrap;">{safe_content_val}</div>', unsafe_allow_html=True)
                    else:
                        st.warning("未找到该案例详情。")
            except Exception as e:
                st.error(f"加载案例库表格异常: {e}")
        else:
            st.warning("未检测到 `案例库.xlsx` 文件，请确认已上传至同一目录。")
            
    else:
        st.markdown(
            """
            <div class="newsprint-masthead">
                <span>VOL. I NO. 01</span>
                <span>AUTOMOTIVE DATA COMPLIANCE REVIEW</span>
                <span>GLOBAL EDITION</span>
            </div>
            """,
            unsafe_allow_html=True
        )
        
        title_col, btn_col = st.columns([4, 1])
        with title_col:
            st.markdown(
                """
                <div class="sharp-card" style="border-top: 4px solid #111; margin-bottom: 0; height: 100%;">
                    <h1 style='margin-top:0; border-bottom:none; font-size: 2.4rem;'>智能网联汽车跨国数据合规平台</h1>
                    <p style='font-family: Lora, serif; font-size: 1rem; line-height: 1.5; color: #333333; margin-bottom: 0;'>
                        <b>中国、欧盟、美国</b>三大法域的车外实景影像及关键汽车数据合规要求汇总。<br>
                        按法域分类整理，方便对照查阅，帮您在做跨境合规时快速找到需要的规则。
                    </p>
                </div>
                """, 
                unsafe_allow_html=True
            )
        with btn_col:
            st.markdown(
                """
                <div class="sharp-card" style="border-top: 4px solid #111; margin-bottom: 0; display: flex; align-items: center; justify-content: center; height: 100%;">
                """,
                unsafe_allow_html=True
            )
            st.markdown('<div class="inline-term-btn">', unsafe_allow_html=True)
            if st.button("术语解释总结", key="inline_terms_btn", use_container_width=True):
                st.session_state.show_terms_page = True
                st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)
            
        st.write("")
        
        if not success_db:
            st.error("主数据加载失败！请确保对应的 Excel 文件与本项目代码在同一目录下。")
        else:
            conn = sqlite3.connect(DB_FILE)
            
            if st.session_state.nav_choice == "法律库":
                filter_col1, filter_col2 = st.columns(2)
                with filter_col1:
                    selected_region = st.selectbox("司法辖区", ["全部", "中国", "欧盟", "美国"])
                with filter_col2:
                    if selected_region == "全部":
                        categories_df = pd.read_sql("SELECT DISTINCT category FROM compliance_laws", conn)
                    else:
                        categories_df = pd.read_sql("SELECT DISTINCT category FROM compliance_laws WHERE region = ?", conn, params=(selected_region,))
                    categories = ["全部"] + categories_df["category"].tolist()
                    selected_category = st.selectbox("合规模块", categories)
                    
                keyword = st.text_input("搜索", placeholder="如：数据出境、GDPR...")
                
                export_col1, export_col2 = st.columns([2, 1])
                with export_col1:
                    st.markdown(f"### 已选择 {len(st.session_state.selected_laws)} 条法条")
                    if st.session_state.selected_laws:
                        st.markdown("**自动生成引用格式：**")
                        for law in st.session_state.selected_laws:
                            art_num = extract_article_number(law.get("content", ""))
                            citation_title = f"{law['law_title']} {art_num}".strip() if art_num else law["law_title"]
                            st.write(citation_title)
                with export_col2:
                    st.markdown("### 操作")
                    generate_pdf_clicked = st.button("生成所选法条 PDF")
                    pdf_download_container = st.container()
                    
                query = "SELECT region, category, law_title, sub_cat_0, sub_cat_1, content FROM compliance_laws"
                conditions = []
                params = []
                if selected_region != "全部":
                    conditions.append("region = ?")
                    params.append(selected_region)
                if selected_category != "全部":
                    conditions.append("category = ?")
                    params.append(selected_category)
                if keyword:
                    wildcard = f"%{keyword}%"
                    conditions.append("(content LIKE ? OR law_title LIKE ? OR category LIKE ? OR sub_cat_0 LIKE ? OR sub_cat_1 LIKE ?)")
                    params.extend([wildcard]*5)
                    
                if conditions:
                    query += " WHERE " + " AND ".join(conditions)
                query += " ORDER BY region, category, sort_order"
                
                module_df = pd.read_sql(query, conn, params=tuple(params))
                
                if keyword:
                    st.markdown(f"**检索结果**：包含 <span style='border:1px solid #111; color:#111; font-weight:bold; padding:2px 6px;'>“{keyword}”</span> 的内容共 **{len(module_df)}** 条", unsafe_allow_html=True)
                else:
                    st.markdown(f"**检索条件**：辖区 [{selected_region}] | 模块 [{selected_category}] -> 共计检索到 **{len(module_df)}** 条内容")
                    
                current_ids = set(module_df.index.tolist())
                st.session_state.selected_laws = [
                    x for x in st.session_state.selected_laws
                    if x["db_index"] in current_ids
                ]
                
                grouped = module_df.groupby(["region", "category", "law_title"], sort=False)
                for (region_name, cat_name, law_title), group in grouped:
                    expander_label = f"【{region_name}】 {law_title} ({len(group)} 条)"
                    with st.expander(expander_label, expanded=False):
                        st.markdown(f"<h4 style='font-family: Playfair Display, serif;'>{law_title}</h4>", unsafe_allow_html=True)
                        st.caption(f"归属辖区：{region_name} | 模块：{cat_name}")
                        
                        for idx, row in group.iterrows():
                            sc0 = row["sub_cat_0"]
                            law_id = int(idx)
                            law_text = row["content"]
                            
                            checked = st.checkbox(
                                f"选择该条文",
                                key=f"law_checkbox_{law_id}"
                            )
                            
                            law_item = {
                                "db_index": law_id,
                                "law_title": law_title,
                                "content": law_text
                            }
                            if checked:
                                if not any(x["db_index"] == law_id for x in st.session_state.selected_laws):
                                    st.session_state.selected_laws.append(law_item)
                            else:
                                st.session_state.selected_laws = [
                                    x for x in st.session_state.selected_laws
                                    if x["db_index"] != law_id
                                ]
                                
                            tags_html = ""
                            if sc0:
                                tags = [t.strip() for t in sc0.split("|") if t.strip()]
                                tags_str = "".join([f'<span class="law-tag">{t}</span>' for t in tags])
                                tags_html = f'<div style="margin-bottom:10px;">{tags_str}</div>'
                                
                            content_text = law_text
                            content_text = re.sub(r'(?im)^\s*svg\s*$', '', content_text)
                            
                            if keyword:
                                content_text = content_text.replace(
                                    keyword,
                                    f"<span style='border:1px solid #111;color:#CC0000;font-weight:bold;'>{keyword}</span>"
                                )
                                
                            content_text = content_text.replace('\n', '<br>')
                            
                            html_str = (
                                f'<div class="law-content" style="margin-bottom:20px;white-space:normal;">'
                                f'{tags_html}'
                                f'<div style="white-space:pre-wrap;line-height:1.8;">{content_text}</div>'
                                f'</div>'
                            )
                            st.markdown(html_str, unsafe_allow_html=True)
                            
                st.divider()
                
                if generate_pdf_clicked:
                    buffer = BytesIO()
                    pdf = SimpleDocTemplate(buffer, pagesize=A4)
                    styles = getSampleStyleSheet()
                    
                    title_style = ParagraphStyle(
                        "ChineseTitle", parent=styles["Title"], fontName="STSong-Light"
                    )
                    heading_style = ParagraphStyle(
                        "ChineseHeading", parent=styles["Heading3"], fontName="STSong-Light"
                    )
                    body_style = ParagraphStyle(
                        "ChineseBody", parent=styles["BodyText"], fontName="STSong-Light", leading=18
                    )
                    
                    elements = []
                    elements.append(Paragraph("企业合规自查法条清单", title_style))
                    elements.append(Spacer(1, 12))
                    
                    for law in st.session_state.selected_laws:
                        art_num = extract_article_number(law.get("content", ""))
                        citation_title = f"{law['law_title']} {art_num}".strip() if art_num else law["law_title"]
                        elements.append(Paragraph(citation_title, heading_style))
                        clean_pdf_text = re.sub(r'(?im)^\s*svg\s*$', '', law["content"])
                        safe_content = clean_pdf_text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\n", "<br/>")
                        elements.append(Paragraph(safe_content, body_style))
                        elements.append(Spacer(1, 12))
                        
                    elements.append(Paragraph("提示：本清单用于企业合规检索与参考，不构成法律意见。", body_style))
                    pdf.build(elements)
                    buffer.seek(0)
                    
                    with pdf_download_container:
                        st.download_button(
                            "下载PDF",
                            data=buffer,
                            file_name="企业合规自查法条清单.pdf",
                            mime="application/pdf"
                        )
                        
            elif st.session_state.nav_choice == "出境全流程时间轴":
                st.markdown("### 数据出境全流程纵向时间轴")
                st.markdown("我们将数据出境的合规流程拆成三个阶段：出境前的准备与评估、出境中的实施与传输、出境后的合规监督。按这个顺序梳理，您能更清楚每一步该做什么。")
                
                all_laws_df = pd.read_sql("SELECT region, category, law_title, sub_cat_0, sub_cat_1, content FROM compliance_laws", conn)
                
                timeline_phases = [
                    {"title": "Phase 1：出境前准备与评估 (Data Mapping & Assessment)", 
                     "desc": "完成数据资产梳理、分类分级，执行数据出境安全评估、标准合同签署或个人信息保护认证。"},
                    {"title": "Phase 2：出境中实施与传输 (Secure Transmission & Protection)", 
                     "desc": "车内处理、默认不收集、脱敏等原则，以及跨境传输链路安全和技术保护措施。"},
                    {"title": "Phase 3：出境后合规监督 (Post-transfer Monitoring & Audit)", 
                     "desc": "建立持续合规审计机制、安全事件应急响应与境外接收方权益保障监督。"}
                ]
                
                phase_tabs = st.tabs([f"{p['title'].split(' ')[0]} {p['title'].split(' ')[1]}" for p in timeline_phases])
                
                for i, p_info in enumerate(timeline_phases):
                    with phase_tabs[i]:
                        st.markdown(f"<h4 style='font-family: Playfair Display, serif;'>{p_info['title']}</h4>", unsafe_allow_html=True)
                        st.info(p_info['desc'])
                        
                        if i == 0:
                            phase_df = all_laws_df[all_laws_df["category"].str.contains("分类|出境|安全评估|标准合同|认证", na=False) | all_laws_df["law_title"].str.contains("评估|办法|条例|规定", na=False)]
                        elif i == 1:
                            phase_df = all_laws_df[all_laws_df["content"].str.contains("传输|出境|向境外|接收|加密|安全保护", na=False)]
                            if phase_df.empty: phase_df = all_laws_df.iloc[3:7]
                        else:
                            phase_df = all_laws_df[all_laws_df["content"].str.contains("监督|审计|评估|报告|应急|处置", na=False)]
                            if phase_df.empty: phase_df = all_laws_df.iloc[7:]
                            
                        st.markdown('<div class="timeline-container">', unsafe_allow_html=True)
                        for _, row in phase_df.iterrows():
                            region_n = row["region"]
                            law_t = row["law_title"]
                            sc0 = row["sub_cat_0"]
                            content = row["content"]
                            tag_str = f"[{region_n}] {sc0}" if sc0 else f"[{region_n}]"
                            
                            content_safe = content
                            content_safe = re.sub(r'(?im)^\s*svg\s*$', '', content_safe)
                            content_safe = content_safe.replace('\n', '<br>')
                            
                            timeline_card_html = (
                                f'<div class="timeline-item">'
                                f'<div class="timeline-node"></div>'
                                f'<div class="timeline-card hard-shadow-hover">'
                                f'<span class="law-tag">{tag_str}</span>'
                                f'<h4 style="margin-top: 5px; color: #111111; font-family: Playfair Display, serif; border-bottom: none;">{law_t}</h4>'
                                f'<div class="law-content" style="margin-bottom: 0; white-space: pre-wrap;">{content_safe}</div>'
                                f'</div>'
                                f'</div>'
                            )
                            st.markdown(timeline_card_html, unsafe_allow_html=True)
                        st.markdown('</div>', unsafe_allow_html=True)
            conn.close()
