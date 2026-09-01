import os
import re
import sqlite3
import pandas as pd
import openpyxl
import streamlit as st

# ==================== 1. 页面配置 ====================
st.set_page_config(
    page_title="智能网联汽车与跨国数据合规检索平台",
    page_icon="⚖️", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# 初始化 Session State，用于控制术语界面的显示/隐藏
if "show_terms_page" not in st.session_state:
    st.session_state.show_terms_page = False

def toggle_terms_page():
    st.session_state.show_terms_page = not st.session_state.show_terms_page

# ==================== 2. 全局 CSS 样式与 UI 设计系统 (蓝白色调) ====================
CUSTOM_CSS = """
<style>
    /* ================= 设计系统变量 (蓝白色调) ================= */
    :root {
        --bg-base: #f4f8fc;
        --bg-surface: #ffffff;
        --text-primary: #1e293b;
        --text-muted: #64748b;
        --border-subtle: #e2e8f0;
        --border-light: #cbd5e1;
        --accent-color: #2563eb;
        --accent-glow: rgba(37, 99, 235, 0.15);
        --easing-primary: cubic-bezier(0.16, 1, 0.3, 1);
        --easing-hover: ease-out;
    }

    /* 页面基础背景 */
    [data-testid="stAppViewContainer"] {
        background-color: var(--bg-base) !important;
        color: var(--text-primary) !important;
    }
    
    [data-testid="stHeader"] {
        background-color: transparent !important;
    }

    /* 全局字体与文字颜色 */
    html, body, [class*="css"], p, span, div, label, li {
        font-family: "PingFang SC", "Microsoft YaHei", "Helvetica Neue", sans-serif;
        color: var(--text-primary);
    }

    /* ================= 动画与入场效果 ================= */
    @keyframes fadeUp {
        0% { opacity: 0; transform: translateY(16px); }
        100% { opacity: 1; transform: translateY(0); }
    }

    /* 标题统一样式 */
    h1, h2, h3, h4 {
        color: #1d4ed8 !important;
        font-weight: 600 !important;
        letter-spacing: 0.3px;
    }

    /* ================= 核心组件样式 (Cards & Expanders) ================= */
    .header-card, div[data-testid="stExpander"], .term-card, .timeline-card {
        background-color: var(--bg-surface) !important;
        border: 1px solid var(--border-subtle) !important;
        border-radius: 10px;
        box-shadow: 0 4px 12px rgba(37, 99, 235, 0.04);
        transition: transform 300ms var(--easing-primary), box-shadow 300ms var(--easing-primary), border-color 300ms var(--easing-primary);
        animation: fadeUp 500ms var(--easing-primary) forwards;
        margin-bottom: 20px;
        padding: 20px;
    }

    /* 悬浮微交互 */
    .header-card:hover, div[data-testid="stExpander"]:hover, .term-card:hover, .timeline-card:hover {
        transform: translateY(-2px);
        border-color: var(--accent-color) !important;
        box-shadow: 0 8px 20px rgba(37, 99, 235, 0.08);
    }

    /* Expander 内部定制 */
    div[data-testid="stExpander"] { padding: 0 !important; }
    div[data-testid="stExpander"] summary {
        padding: 12px 18px;
        background-color: #f8fafc;
        border-radius: 8px;
    }
    div[data-testid="stExpander"] summary p {
        font-weight: 600;
        color: #1e3a8a !important;
    }

    /* 侧边栏卡片化与清爽色调 */
    [data-testid="stSidebar"] {
        background-color: #ffffff !important;
        border-right: 1px solid var(--border-subtle) !important;
    }

    /* 分割线 */
    hr {
        border: none !important;
        height: 1px;
        background: linear-gradient(90deg, transparent, #cbd5e1, transparent) !important;
        margin: 2rem 0;
    }

    /* ================= 标签与内容展示 ================= */
    .law-tag {
        display: inline-block;
        background-color: #eff6ff;
        color: #1d4ed8;
        padding: 4px 12px;
        border-radius: 6px;
        font-size: 0.85em;
        font-weight: 600;
        margin-bottom: 12px;
        border: 1px solid #dbeafe;
    }

    .law-content {
        background-color: #f8fafc;
        border-left: 3px solid #2563eb;
        padding: 15px 20px;
        color: #334155;
        line-height: 1.8;
        font-size: 0.95em;
        text-align: justify;
        white-space: pre-wrap;
        border-radius: 0 8px 8px 0;
    }

    /* 术语解释专属卡片样式 */
    .term-card {
        border-left: 3px solid #f59e0b !important;
    }
    .term-source {
        color: #64748b;
        font-size: 0.85em;
        margin-top: 15px;
        text-align: right;
        font-weight: 500;
    }

    /* ================= 纵向时间轴 ================= */
    .timeline-container {
        position: relative;
        padding-left: 30px;
        margin: 30px 0;
        border-left: 2px solid #93c5fd;
    }
    .timeline-item {
        position: relative;
        margin-bottom: 35px;
        animation: fadeUp 500ms var(--easing-primary) forwards;
    }
    .timeline-node {
        position: absolute;
        left: -37px;
        top: 18px;
        width: 12px;
        height: 12px;
        border-radius: 50%;
        background-color: #ffffff;
        border: 3px solid #2563eb;
        box-shadow: 0 0 0 4px rgba(37, 99, 235, 0.1);
    }

    /* ================= 交互控件 (Inputs & Buttons) ================= */
    div[data-testid="stTextInput"] input {
        background-color: #ffffff !important;
        border: 1px solid #cbd5e1 !important;
        color: #1e293b !important;
        font-weight: 500 !important;
        border-radius: 8px !important;
        transition: all 200ms var(--easing-primary);
    }
    div[data-testid="stTextInput"] input:focus {
        border-color: #2563eb !important;
        box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.15) !important;
    }
    div[data-testid="stTextInput"] label {
        font-weight: 600 !important;
        color: #1e3a8a !important;
    }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

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

    # 表格结构定义：
    # 第 0 行 (iloc[0]): 辖区 + 模块
    # 第 1 行 (iloc[1]): 法律名称
    # 前三列 (col 0, 1, 2): 适用场景标签 (sub_cat)
    # 从第 3 列开始 (col_idx >= 3) 为具体的法律条款列，每个单元格独立切分
    categories_row = df_raw.iloc[0]
    titles_row = df_raw.iloc[1]

    # 读取前三列作为适用场景依据
    col0_raw = df_raw.iloc[:, 0] if len(df_raw.columns) > 0 else pd.Series([""] * len(df_raw))
    col1_raw = df_raw.iloc[:, 1] if len(df_raw.columns) > 1 else pd.Series([""] * len(df_raw))
    col2_raw = df_raw.iloc[:, 2] if len(df_raw.columns) > 2 else pd.Series([""] * len(df_raw))

    # 从第 3 列（索引 3）开始，逐列扫描
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
        # 从第 3 行开始（索引从 2 开始），对该列的每一个单元格进行独立识别与入库
        for row_idx in range(2, len(df_raw)):
            cell_obj = ws.cell(row=row_idx + 1, column=col_idx + 1)
            cell_val = cell_obj.value
            
            s0 = str(col0_raw.iloc[row_idx]).strip()
            s1 = str(col1_raw.iloc[row_idx]).strip()
            s2 = str(col2_raw.iloc[row_idx]).strip() if len(df_raw.columns) > 2 else ""

            # 对应前三列作为适用场景标签（若为空则保留空）
            sub_c0 = s0 if s0 and s0 != "nan" else ""
            sub_c1 = " / ".join([x for x in [s1, s2] if x and x != "nan"])

            # 严格识别每一个非空单元格，做到“一个格子对应一条记录”
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


# ==================== 4. 侧边栏及页面路由 ====================
st.sidebar.markdown("### 📖 术语库导航")
if st.sidebar.button("进入【术语解释总结】页面 ➔" if not st.session_state.show_terms_page else "🔙 退出术语界面", on_click=toggle_terms_page, type="primary"):
    pass

st.sidebar.markdown("---")
st.sidebar.markdown("### 🧭 主系统功能")
nav_mode = st.sidebar.radio(
    "切换功能模块", 
    ["📑 体系化法律库", "🔎 穿透式法规检索", "⏱️ 出境全流程时间轴"],
    disabled=st.session_state.show_terms_page 
)


# ==================== 5. 页面展示逻辑 ====================

if st.session_state.show_terms_page:
    st.button("🔙 返回主合规平台", on_click=toggle_terms_page)
    st.markdown("<h2 style='text-align: center; color: #d97706 !important;'>📖 术语解释总结全库展示</h2>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #64748b;'>展示完整的术语释义。支持基于首个英文冒号前关键词的模糊搜索与多国近似词自动联动。</p>", unsafe_allow_html=True)
    st.markdown("---")

    term_keyword = st.text_input("🔍 输入术语关键词 (如：个人信息、sell、重要数据...)", key="standalone_term_search", placeholder="在此输入关键字进行检索...")

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
                    <div class="term-card">
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
    st.markdown(
        """
        <div class="header-card">
            <h1 style='margin-top:0;'>⚖️ 智能网联汽车跨国数据合规平台</h1>
            <p style='color:#64748b; font-size:1.05em; margin-bottom:0;'>
                本系统集成 <b>中国、欧盟、美国</b> 三大核心司法辖区的合规指引，支持模块化导航与多维精准检索。<br>
                致力于为车企出境数据合规提供一站式法律支撑。
            </p>
        </div>
        """, 
        unsafe_allow_html=True
    )

    if not success_db:
        st.error("主数据加载失败！请确保对应的 Excel 文件与本项目代码在同一目录下。")
    else:
        conn = sqlite3.connect(DB_FILE)

        if nav_mode == "📑 体系化法律库":
            selected_region = st.sidebar.selectbox("🌐 司法辖区", ["全部", "中国", "欧盟", "美国"])

            if selected_region == "全部":
                categories_df = pd.read_sql("SELECT DISTINCT category FROM compliance_laws", conn)
            else:
                categories_df = pd.read_sql("SELECT DISTINCT category FROM compliance_laws WHERE region = ?", conn, params=(selected_region,))

            categories = ["全部"] + categories_df["category"].tolist()
            selected_category = st.sidebar.selectbox("📁 合规模块", categories)

            query = "SELECT region, category, law_title, sub_cat_0, sub_cat_1, content FROM compliance_laws"
            conditions = []
            params = []

            if selected_region != "全部":
                conditions.append("region = ?")
                params.append(selected_region)
            if selected_category != "全部":
                conditions.append("category = ?")
                params.append(selected_category)

            if conditions:
                query += " WHERE " + " AND ".join(conditions)

            query += " ORDER BY region, category, sort_order"
            module_df = pd.read_sql(query, conn, params=tuple(params))

            st.markdown(f"**检索条件**：辖区 [{selected_region}] &nbsp;|&nbsp; 模块 [{selected_category}] &nbsp;➔&nbsp; 共计检索到 **{len(module_df)}** 条内容")
            st.write("")

            grouped = module_df.groupby("law_title")

            for law_title, group in grouped:
                region_name = group.iloc[0]["region"]
                cat_name = group.iloc[0]["category"]
                expander_label = f"📌 【{region_name}】 {law_title} ({len(group)} 条)"

                with st.expander(expander_label, expanded=False):
                    st.markdown(f"#### {law_title}")
                    st.caption(f"归属辖区：{region_name} | 模块：{cat_name}")

                    for idx, row in group.reset_index().iterrows():
                        sc0, sc1 = row["sub_cat_0"], row["sub_cat_1"]
                        if sc0 or sc1:
                            tag_content = f"{sc0}" + (f" ➔ {sc1}" if sc1 else "")
                            st.markdown(f'<span class="law-tag">💡 {tag_content}</span>', unsafe_allow_html=True)

                        st.markdown(f'<div class="law-content">{row["content"]}</div>', unsafe_allow_html=True)

        elif nav_mode == "🔎 穿透式法规检索":
            keyword = st.sidebar.text_input("🔍 输入检索关键词", placeholder="如：数据出境、GDPR...")
            st.sidebar.caption("支持模糊搜索法规条款, 标签或分类维度。")

            if keyword:
                wildcard = f"%{keyword}%"
                search_query = """
                    SELECT region, category, law_title, sub_cat_0, sub_cat_1, content 
                    FROM compliance_laws 
                    WHERE content LIKE ? OR law_title LIKE ? OR category LIKE ? OR sub_cat_0 LIKE ? OR sub_cat_1 LIKE ?
                    ORDER BY region, category, sort_order
                """
                results_df = pd.read_sql(search_query, conn, params=(wildcard,)*5)

                st.markdown(f"**检索结果**：包含 <span style='color:#2563eb; font-weight:bold;'>“{keyword}”</span> 的内容共 **{len(results_df)}** 条", unsafe_allow_html=True)
                st.write("")

                grouped_search = results_df.groupby("law_title")
                for law_title, group in grouped_search:
                    region_name, cat_name = group.iloc[0]["region"], group.iloc[0]["category"]

                    with st.expander(f"📌 【{region_name}】 {law_title}"):
                        st.markdown(f"#### {law_title}")
                        for idx, row in group.reset_index().iterrows():
                            sc0, sc1 = row["sub_cat_0"], row["sub_cat_1"]
                            if sc0 or sc1:
                                tag_content = f"{sc0}" + (f" ➔ {sc1}" if sc1 else "")
                                st.markdown(f'<span class="law-tag">💡 {tag_content}</span>', unsafe_allow_html=True)

                            highlighted_content = row["content"].replace(keyword, f"<span style='background-color:#dbeafe; color:#1d4ed8; font-weight:bold; padding:0 2px; border-radius:2px;'>{keyword}</span>")
                            st.markdown(f'<div class="law-content">{highlighted_content}</div>', unsafe_allow_html=True)
            else:
                st.info("👈 请在左侧侧边栏输入关键词以获取检索结果。")

        elif nav_mode == "⏱️ 出境全流程时间轴":
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
                    st.markdown(f"#### {p_info['title']}")
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
                            <div class="timeline-card">
                                <span class="law-tag">{tag_str}</span>
                                <h4 style="margin-top: 5px; color: #1d4ed8;">{law_t}</h4>
                                <div class="law-content" style="margin-bottom: 0;">{content}</div>
                            </div>
                        </div>
                        """
                        st.markdown(timeline_card_html, unsafe_allow_html=True)
                    st.markdown('</div>', unsafe_allow_html=True)

        conn.close()
