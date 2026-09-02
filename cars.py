import os
import re
import sqlite3
import pandas as pd
import openpyxl
import streamlit as st

# ==================== 1. 页面配置 ====================
st.set_page_config(
    page_title="智能网联汽车与跨国数据合规检索平台 | Newsprint Edition",
    page_icon="📰", 
    layout="wide",
    initial_sidebar_state="collapsed"
)

# 初始化 Session State，用于控制术语界面的显示/隐藏与顶部导航状态
if "show_terms_page" not in st.session_state:
    st.session_state.show_terms_page = False

if "nav_choice" not in st.session_state:
    st.session_state.nav_choice = "法律库"

def toggle_terms_page():
    st.session_state.show_terms_page = not st.session_state.show_terms_page

# ==================== 2. 全局 CSS 样式与 UI 设计系统 (Newsprint 风格重构) ====================
NEWSPRINT_CSS = """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500&family=Playfair+Display:ital,wght@0,400;0,600;0,700;0,900;1,400&family=Lora:ital,wght@0,400;0,600;1,400&display=swap');

    :root {
        --bg-base: #F9F9F7;
        --bg-surface: #F9F9F7;
        --text-primary: #111111;
        --text-muted: #666666;
        --border-color: #111111;
        --accent-red: #CC0000;
        --divider-grey: #E5E5E0;
    }

    /* 页面基础背景与网纹 */
    [data-testid="stAppViewContainer"] {
        background-color: var(--bg-base) !important;
        background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='4' height='4' viewBox='0 0 4 4'%3E%3Cpath fill='%23111111' fill-opacity='0.04' d='M1 3h1v1H1V3zm2-2h1v1H3V1z'%3E%3C/path%3E%3C/svg%3E") !important;
        color: var(--text-primary) !important;
    }
    
    [data-testid="stHeader"] {
        background-color: transparent !important;
    }

    /* 全局字体定义：避开 span 标签，防止干扰内部图标 */
    html, body, p, label, li, .law-content {
        font-family: 'Lora', Georgia, serif;
        color: var(--text-primary);
    }

    /* 标题统合为 Playfair Display 衬线大标题 */
    h1, h2, h3, h4 {
        font-family: 'Playfair Display', 'Times New Roman', serif !important;
        color: #111111 !important;
        font-weight: 900 !important;
        letter-spacing: -0.03em;
        border-bottom: 2px solid #111111;
        padding-bottom: 8px;
        margin-bottom: 20px;
    }

    /* 绝对零圆角与硬阴影悬浮效果 */
    .sharp-card, div[data-testid="stExpander"], .term-card, .timeline-card, .header-card {
        background-color: #F9F9F7 !important;
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

    /* Expander 严格零圆角和报纸风折叠头，并彻底隐藏 arrow 图标防止文字重叠 */
    div[data-testid="stExpander"] { padding: 0 !important; }
    div[data-testid="stExpander"] summary {
        padding: 16px 20px;
        background-color: #F9F9F7;
        border: 1px solid #111111;
        border-radius: 0px !important;
    }
    div[data-testid="stExpander"] summary span[data-testid="stExpanderToggleIcon"] {
        font-size: 0px !important;
        color: transparent !important;
        width: 0px !important;
        display: none !important;
    }
    div[data-testid="stExpander"] summary:hover {
        background-color: #111111 !important;
        color: #F9F9F7 !important;
    }
    div[data-testid="stExpander"] summary p {
        font-family: 'Playfair Display', serif !important;
        font-weight: 700;
        color: #111111 !important;
        margin: 0 !important;
    }
    div[data-testid="stExpander"] summary:hover p {
        color: #F9F9F7 !important;
    }

    /* 隐藏默认侧边栏 */
    [data-testid="stSidebar"] {
        display: none !important;
    }

    /* 顶部导航栏按钮统一样式 */
    .stButton button {
        background-color: #F9F9F7 !important;
        color: #111111 !important;
        border: 1px solid #111111 !important;
        border-radius: 0px !important;
        font-family: 'Inter', sans-serif !important;
        font-weight: 600 !important;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        box-shadow: none !important;
        width: 100%;
        transition: background-color 100ms ease, color 100ms ease !important;
    }
    .stButton button:hover {
        background-color: #111111 !important;
        color: #F9F9F7 !important;
        border: 1px solid #111111 !important;
        box-shadow: 3px 3px 0px 0px #111111 !important;
    }

    /* 报纸风格标签 */
    .law-tag {
        display: inline-block;
        background-color: #111111;
        color: #F9F9F7;
        padding: 4px 10px;
        border-radius: 0px !important;
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.75rem;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        margin-bottom: 12px;
        border: 1px solid #111111;
    }

    /* 详细条款排版：衬线体、两端对齐、左侧重黑边 */
    .law-content {
        background-color: #ffffff;
        border: 1px solid #111111;
        border-left: 6px solid #111111 !important;
        padding: 20px;
        color: #111111;
        line-height: 1.8;
        font-size: 1rem;
        text-align: justify;
        white-space: pre-wrap;
        border-radius: 0px !important;
    }

    /* 术语卡片与时间轴微调 */
    .term-card {
        border-left: 6px solid var(--accent-red) !important;
    }
    .term-source {
        font-family: 'JetBrains Mono', monospace;
        color: #666666;
        font-size: 0.8rem;
        margin-top: 15px;
        text-align: right;
    }

    /* 纵向时间轴样式 */
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
        background-color: #F9F9F7;
        border: 3px solid #111111;
    }

    /* 控件设计：无圆角、底部双黑线输入框 */
    div[data-testid="stTextInput"] input {
        background-color: transparent !important;
        border: none !important;
        border-bottom: 2px solid #111111 !important;
        color: #111111 !important;
        font-family: 'JetBrains Mono', monospace !important;
        border-radius: 0px !important;
        box-shadow: none !important;
    }
    div[data-testid="stTextInput"] input:focus {
        background-color: #E5E5E0 !important;
    }
    div[data-testid="stTextInput"] label {
        font-family: 'Inter', sans-serif !important;
        font-weight: 700 !important;
        text-transform: uppercase;
        font-size: 0.8rem;
    }

    /* 顶栏报头元数据排版 */
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
</style>
"""
st.markdown(NEWSPRINT_CSS, unsafe_allow_html=True)

DB_FILE = "car_compliance.db"

# ==================== 3. 核心处理与数据库函数 ====================
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

def get_clean_cell_text(cell):
    if not cell.value or str(cell.value).strip() == "nan":
        return ""
    
    is_rich = hasattr(cell, 'value') and isinstance(cell.value, openpyxl.cell.rich_text.CellRichText)
    if is_rich:
        full_text = "".join([str(rt.text) for rt in cell.value])
    else:
        full_text = str(cell.value)
    
    return full_text.strip()

@st.cache_data
def init_database_from_excel():
    current_dir = os.path.dirname(os.path.abspath(__file__))
    excel_path = os.path.join(current_dir, "合规平台条文整理（修改1.0）.xlsx")
    if not os.path.exists(excel_path):
        excel_path = os.path.join(current_dir, "合规平台条文整理 (1).xlsx")

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

    if not os.path.exists(excel_path):
        conn.close()
        return False

    wb = openpyxl.load_workbook(excel_path, data_only=False)
    sheet_name = wb.sheetnames[0]
    ws = wb[sheet_name]

    df_raw = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)
    cursor.execute("DELETE FROM compliance_laws")

    categories_row = df_raw.iloc[0]
    titles_row = df_raw.iloc[1]

    col0_raw = df_raw.iloc[:, 0] if len(df_raw.columns) > 0 else pd.Series([""] * len(df_raw))
    col1_raw = df_raw.iloc[:, 1] if len(df_raw.columns) > 1 else pd.Series([""] * len(df_raw))
    col2_raw = df_raw.iloc[:, 2] if len(df_raw.columns) > 2 else pd.Series([""] * len(df_raw))

    for col_idx in range(3, len(df_raw.columns)):
        cat_name = str(categories_row.iloc[col_idx]).strip()
        law_title = str(titles_row.iloc[col_idx]).strip()
        if not law_title or law_title == "nan": 
            continue

        region = "中国"
        if "欧盟" in cat_name: region = "欧盟"
        elif "美国" in cat_name: region = "美国"

        category = cat_name if cat_name and cat_name != "nan" else "通用效力模块"

        has_content = False
        
        for row_idx in range(2, len(df_raw)):
            cell_obj = ws.cell(row=row_idx + 1, column=col_idx + 1)
            cell_val = cell_obj.value
            
            s0 = str(col0_raw.iloc[row_idx]).strip()
            s1 = str(col1_raw.iloc[row_idx]).strip()
            s2 = str(col2_raw.iloc[row_idx]).strip() if len(df_raw.columns) > 2 else ""

            sub_c0 = s0 if s0 and s0 != "nan" else ""
            sub_c1 = " / ".join([x for x in [s1, s2] if x and x != "nan"])

            if cell_val is not None and str(cell_val).strip() != "nan":
                content_str = get_clean_cell_text(cell_obj)
                if content_str and content_str != "nan":
                    has_content = True
                    sort_val = extract_sort_key(content_str)
                    
                    cursor.execute(
                        """SELECT COUNT(1) FROM compliance_laws 
                           WHERE region=? AND category=? AND law_title=? AND sub_cat_0=? AND sub_cat_1=? AND content=?""",
                        (region, category, law_title, sub_c0, sub_c1, content_str)
                    )
                    exists = cursor.fetchone()[0]
                    
                    if exists == 0:
                        cursor.execute(
                            "INSERT INTO compliance_laws (region, category, law_title, sub_cat_0, sub_cat_1, content, sort_order) VALUES (?, ?, ?, ?, ?, ?, ?)",
                            (region, category, law_title, sub_c0, sub_c1, content_str, sort_val)
                        )

        if not has_content:
            cursor.execute(
                "INSERT INTO compliance_laws (region, category, law_title, sub_cat_0, sub_cat_1, content, sort_order) VALUES (?, ?, ?, ?, ?, ?, ?)",
                (region, category, law_title, "暂无分类", "暂无指引", "（该法规条文正在整理补充中，敬请期待...）", 999)
            )

    conn.commit()
    conn.close()
    return True

success_db = init_database_from_excel()


# ==================== 4. 页面最顶端：术语解释按钮与导航栏集成 ====================
top_col1, top_col2 = st.columns([1, 4])
with top_col1:
    if st.button("📖 术语解释总结" if not st.session_state.show_terms_page else "🔙 返回主检索系统", key="top_terms_btn"):
        toggle_terms_page()
        st.rerun()

st.markdown("---")

# 顶部导航栏：从左到右依次为 首页 - 法律库 - 出境全流程时间轴 - 关于我们
nav_cols = st.columns(4)
with nav_cols[0]:
    if st.button("🏠 首页", key="nav_home"):
        st.session_state.nav_choice = "首页"
        st.session_state.show_terms_page = False
        st.rerun()
with nav_cols[1]:
    if st.button("📑 法律库", key="nav_law"):
        st.session_state.nav_choice = "法律库"
        st.session_state.show_terms_page = False
        st.rerun()
with nav_cols[2]:
    if st.button("⏱️ 出境全流程时间轴", key="nav_timeline"):
        st.session_state.nav_choice = "出境全流程时间轴"
        st.session_state.show_terms_page = False
        st.rerun()
with nav_cols[3]:
    if st.button("ℹ️ 关于我们", key="nav_about"):
        st.session_state.nav_choice = "关于我们"
        st.session_state.show_terms_page = False
        st.rerun()

st.markdown("---")


# ==================== 5. 页面展示逻辑 ====================

if st.session_state.show_terms_page:
    st.markdown("<h2 style='text-align: center; border-bottom: 3px solid #111;'>📖 术语解释总结全库专栏</h2>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; font-family: Lora, serif; color: #666666;'>展示完整的汽车数据及出境合规术语释义，还原现代纸媒专栏的严谨审慎与清晰结构。</p>", unsafe_allow_html=True)
    st.markdown("---")

    term_keyword = st.text_input("🔍 检索术语关键字 (如：个人信息、重要数据、GDPR...)", key="standalone_term_search", placeholder="在此输入关键字进行检索...")

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

                source_text = f"（来源：《{t_item['source']}》）"

                st.markdown(
                    f"""
                    <div class="term-card hard-shadow-hover">
                        <div>{def_html}</div>
                        <div class="term-source">{source_text}</div>
                    </div>
                    """, 
                    unsafe_allow_html=True
                )

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
                    欢迎来到首页。此处保持空白，后续可根据需要自由添加内容。
                </p>
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
            <div class="sharp-card" style="border-top: 4px solid #111;">
                <h1 style='margin-top:0; border-bottom:none; font-size: 2.8rem;'>关于我们</h1>
                <p style='font-family: Lora, serif; font-size: 1.1rem; line-height: 1.6; color: #333333; margin-bottom: 0;'>
                    此处保持空白，后续可根据需要自由添加平台介绍及团队信息。
                </p>
            </div>
            """, 
            unsafe_allow_html=True
        )

    else:
        # Newsprint 报头结构信息
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
                <h1 style='margin-top:0; border-bottom:none; font-size: 2.8rem;'>智能网联汽车跨国数据合规平台</h1>
                <p style='font-family: Lora, serif; font-size: 1.1rem; line-height: 1.6; color: #333333; margin-bottom: 0;'>
                    全面汇总 <b>中国、欧盟、美国</b> 三大核心司法辖区车外实景影像与关键汽车数据合规指引。<br>
                    秉持绝对清晰的网格架构与严谨审慎的编辑标准，为出境合规提供权威、可靠的决策支撑。
                </p>
            </div>
            """, 
            unsafe_allow_html=True
        )

        if not success_db:
            st.error("主数据加载失败！请确保对应的 Excel 文件与本项目代码在同一目录下。")
        else:
            conn = sqlite3.connect(DB_FILE)

            if st.session_state.nav_choice == "法律库":
                # 将原来的侧边栏筛选控件放在主界面上方
                filter_col1, filter_col2 = st.columns(2)
                with filter_col1:
                    selected_region = st.selectbox("🌐 司法辖区", ["全部", "中国", "欧盟", "美国"])
                with filter_col2:
                    if selected_region == "全部":
                        categories_df = pd.read_sql("SELECT DISTINCT category FROM compliance_laws", conn)
                    else:
                        categories_df = pd.read_sql("SELECT DISTINCT category FROM compliance_laws WHERE region = ?", conn, params=(selected_region,))
                    categories = ["全部"] + categories_df["category"].tolist()
                    selected_category = st.selectbox("📁 合规模块", categories)

                st.markdown("---")

                # 穿透式法规检索搜索框直接放到这一行下面
                keyword = st.text_input("🔍 穿透式法规检索关键词", placeholder="如：数据出境、GDPR...")

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
                    st.markdown(f"**检索结果**：包含 <span style='background-color:#111; color:#F9F9F7; font-weight:bold; padding:2px 6px;'>“{keyword}”</span> 的内容共 **{len(module_df)}** 条", unsafe_allow_html=True)
                else:
                    st.markdown(f"**检索条件**：辖区 [{selected_region}] &nbsp;|&nbsp; 模块 [{selected_category}] &nbsp;➔&nbsp; 共计检索到 **{len(module_df)}** 条内容")
                st.write("")

                grouped = module_df.groupby("law_title")

                for law_title, group in grouped:
                    region_name = group.iloc[0]["region"]
                    cat_name = group.iloc[0]["category"]
                    expander_label = f"📌 【{region_name}】 {law_title} ({len(group)} 条)"

                    with st.expander(expander_label, expanded=False):
                        st.markdown(f"<h4 style='font-family: Playfair Display, serif;'>{law_title}</h4>", unsafe_allow_html=True)
                        st.caption(f"归属辖区：{region_name} | 模块：{cat_name}")

                        for idx, row in group.reset_index().iterrows():
                            sc0, sc1 = row["sub_cat_0"], row["sub_cat_1"]
                            if sc0 or sc1:
                                tag_content = f"{sc0}" + (f" ➔ {sc1}" if sc1 else "")
                                st.markdown(f'<span class="law-tag">💡 {tag_content}</span>', unsafe_allow_html=True)

                            content_text = row["content"]
                            if keyword:
                                content_text = content_text.replace(keyword, f"<span style='background-color:#111; color:#F9F9F7; font-weight:bold; padding:0 2px;'>{keyword}</span>")
                            st.markdown(f'<div class="law-content">{content_text}</div>', unsafe_allow_html=True)

            elif st.session_state.nav_choice == "出境全流程时间轴":
                st.markdown("### ⏱️ 数据出境全流程纵向时间轴")
                st.markdown("通过合规生命周期节点（**Phase 1：出境前准备与评估** ➔ **Phase 2：出境中实施与传输** ➔ **Phase 3：出境后合规监督**），直观展现合规实操全景。")
                st.write("")

                all_laws_df = pd.read_sql("SELECT region, category, law_title, sub_cat_0, sub_cat_1, content FROM compliance_laws", conn)

                timeline_phases = [
                    {"title": "Phase 1：出境前准备与评估 (Data Mapping & Assessment)", 
                     "desc": "完成数据资产梳理、分类分级，执行数据出境安全评估、标准合同签署或个人信息保护认证。"},
                    {"title": "Phase 2：出境中实施与传输 (Secure Transmission & Protection)", 
                     "desc": "在符合车内处理、默认不收集、脱敏等原则下，落实跨境传输链路安全及技术保护措施。"},
                    {"title": "Phase 3：出境后合规监督 (Post-transfer Monitoring & Audit)", 
                     "desc": "建立持续合规审计机制、安全事件应急响应与境外接收方权益保障监督。"}
                ]

                phase_tabs = st.tabs([f"📌 {p['title'].split(' ')[0]} {p['title'].split(' ')[1]}" for p in timeline_phases])

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
                            sc0, sc1 = row["sub_cat_0"], row["sub_cat_1"]
                            content = row["content"]

                            tag_str = f"[{region_n}] " + (f"{sc0} ➔ {sc1}" if (sc0 or sc1) else "")

                            timeline_card_html = f"""
                            <div class="timeline-item">
                                <div class="timeline-node"></div>
                                <div class="timeline-card hard-shadow-hover">
                                    <span class="law-tag">{tag_str}</span>
                                    <h4 style="margin-top: 5px; color: #111111; font-family: Playfair Display, serif; border-bottom: none;">{law_t}</h4>
                                    <div class="law-content" style="margin-bottom: 0;">{content}</div>
                                </div>
                            </div>
                            """
                            st.markdown(timeline_card_html, unsafe_allow_html=True)
                        st.markdown('</div>', unsafe_allow_html=True)

            conn.close()
