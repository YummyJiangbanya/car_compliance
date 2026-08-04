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
)

# ==================== 2. 注入北大法宝风格与时间轴专属 CSS ====================
CUSTOM_CSS = """
<style>
    /* 全局字体 */
    html, body, [class*="css"] {
        font-family: "PingFang SC", "Microsoft YaHei", "Helvetica Neue", sans-serif;
    }
    
    /* 背景色 */
    [data-testid="stAppViewContainer"] {
        background-color: #f0f2f6; 
    }
    
    /* 顶部标题区卡片化 */
    .header-card {
        background-color: #ffffff;
        padding: 25px 30px;
        border-radius: 10px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.05);
        border-top: 5px solid #1a5276;
        margin-bottom: 25px;
    }
    
    h1, h2, h3 {
        color: #1a5276 !important; 
        font-weight: 600 !important;
    }
    
    [data-testid="stSidebar"] {
        background-color: #ffffff;
        border-right: 1px solid #e0e0e0;
        box-shadow: 2px 0 10px rgba(0,0,0,0.02);
    }
    
    div[data-testid="stExpander"] {
        background-color: #ffffff;
        border: 1px solid #e6e9f0;
        border-radius: 8px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.04);
        margin-bottom: 15px;
    }
    div[data-testid="stExpander"] summary {
        background-color: #fafbfc;
        color: #2c3e50;
        font-weight: 600;
        padding: 10px 15px;
    }
    
    .law-tag {
        display: inline-block;
        background-color: #e8f0fe;
        color: #1a5276;
        padding: 4px 12px;
        border-radius: 15px;
        font-size: 0.85em;
        font-weight: 600;
        margin-bottom: 10px;
        border: 1px solid #c6dafc;
    }
    
    .law-content {
        background-color: #fafafa;
        border-left: 4px solid #1a5276;
        padding: 15px 20px;
        color: #444444;
        line-height: 1.8;
        font-size: 0.95em;
        margin-bottom: 10px;
        text-align: justify;
    }

    /* 术语解释专属卡片样式 */
    .term-card {
        background-color: #fdfefe;
        padding: 20px;
        border-radius: 8px;
        margin-bottom: 15px;
        border-left: 4px solid #e67e22;
        box-shadow: 0 2px 5px rgba(0,0,0,0.05);
        color: #333;
        line-height: 1.7;
        font-size: 0.95em;
    }
    
    .term-source {
        color: #7f8c8d;
        font-size: 0.95em;
        margin-top: 15px;
        text-align: right;
        font-weight: bold;
    }

    /* 纵向时间轴美化样式 */
    .timeline-container {
        position: relative;
        padding-left: 30px;
        margin-top: 20px;
        margin-bottom: 20px;
        border-left: 3px solid #1a5276;
    }
    .timeline-item {
        position: relative;
        margin-bottom: 30px;
    }
    .timeline-node {
        position: absolute;
        left: -37.5px;
        top: 0px;
        width: 18px;
        height: 18px;
        border-radius: 50%;
        background-color: #1a5276;
        border: 3px solid #ffffff;
        box-shadow: 0 0 0 2px #1a5276;
    }
    .timeline-card {
        background-color: #ffffff;
        padding: 20px;
        border-radius: 8px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.04);
        border: 1px solid #e1e8ed;
    }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

DB_FILE = "car_compliance.db"

# ==================== 3. 核心处理函数 ====================
def parse_and_split_content(cell_text):
    if not cell_text or str(cell_text).strip() == "nan":
        return []
    text = str(cell_text).strip()
    pattern = r"(?=(?:Article\s+\d+|第[零一二三四五六七八九十百0-9]+条|Step\s+\d+))"
    parts = re.split(pattern, text)
    cleaned_parts = [p.strip() for p in parts if p.strip()]
    return [text] if not cleaned_parts else cleaned_parts

def extract_sort_key(text):
    match_cn = re.search(r"第([零一二三四五六七八九十百0-9]+)条", text)
    if match_cn:
        num_str = match_cn.group(1)
        mapping = {"一":1, "二":2, "三":3, "四":4, "五":5, "六":6, "七":7, "八":8, "九":9, "十":10,
                   "十一":11, "十二":12, "十三":13, "十四":14, "十五":15, "十六":16, "十七":17, "十八":18, "十九":19, "二十":20}
        if num_str in mapping: return mapping[num_str]
        try: return int(num_str)
        except ValueError: pass

    match_en = re.search(r"Article\s+(\d+)", text, re.IGNORECASE)
    if match_en:
        try: return int(match_en.group(1))
        except ValueError: pass
    return 999

def init_database_from_excel():
    current_dir = os.path.dirname(os.path.abspath(__file__))
    excel_path = os.path.join(current_dir, "合规平台条文整理.xlsx")
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

    xl_file = pd.ExcelFile(excel_path)
    target_sheet = xl_file.sheet_names[0]
    df_raw = pd.read_excel(excel_path, sheet_name=target_sheet, header=None)
    
    cursor.execute("DELETE FROM compliance_laws")

    categories_row = df_raw.iloc[0].ffill()
    titles_row = df_raw.iloc[1]
    col0_ffill = df_raw.iloc[:, 0].ffill()
    col1_ffill = df_raw.iloc[:, 1].ffill()

    for col_idx in range(3, len(df_raw.columns)):
        cat_name = str(categories_row.iloc[col_idx]).strip()
        law_title = str(titles_row.iloc[col_idx]).strip()
        if not law_title or law_title == "nan": continue

        region = "中国"
        if "欧盟" in cat_name: region = "欧盟"
        elif "美国" in cat_name: region = "美国"

        category = cat_name if cat_name and cat_name != "nan" else "通用效力模块"

        has_content = False
        for row_idx in range(2, len(df_raw)):
            cell_val = df_raw.iloc[row_idx, col_idx]
            s0 = str(col0_ffill.iloc[row_idx]).strip()
            s1 = str(col1_ffill.iloc[row_idx]).strip()
            sub_c0 = s0 if s0 and s0 != "nan" else ""
            sub_c1 = s1 if s1 and s1 != "nan" else ""

            if pd.notna(cell_val):
                split_contents = parse_and_split_content(cell_val)
                for content_str in split_contents:
                    if content_str and content_str != "nan":
                        has_content = True
                        sort_val = extract_sort_key(content_str)
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

success = init_database_from_excel()


# ==================== 4. 页面前端展示 ====================
# 顶部 Banner
st.markdown(
    """
    <div class="header-card">
        <h1 style='margin-top:0;'>⚖️ 智能网联汽车跨国数据合规平台</h1>
        <p style='color:#555; font-size:1.05em; margin-bottom:0;'>
        本系统集成 <b>中国、欧盟、美国</b> 三大核心司法辖区的合规指引，支持模块化导航与多维精准检索。<br>
        致力于为车企出境数据合规提供一站式法律支撑。
        </p>
    </div>
    """, 
    unsafe_allow_html=True
)

if not success:
    st.error("数据加载失败！请确保 `合规平台条文整理.xlsx` 与本项目代码在同一目录下。")
else:
    # --- 主界面顶部独立显眼的术语解释搜索区 ---
    with st.expander("📖 术语解释总结速查库 (点击展开检索)", expanded=False):
        st.markdown("基于项目术语对照表，支持中英文模糊检索与同义词联动。")
        term_keyword = st.text_input("输入术语关键词，例如：个人信息、sell、重要数据...", key="main_term_search")
        
        current_dir = os.path.dirname(os.path.abspath(__file__))
        # 【严格遵循要求4】此处无论如何固定使用 "术语解释总结.xlsx"
        term_excel_path = os.path.join(current_dir, "术语解释总结.xlsx")
        
        if os.path.exists(term_excel_path):
            try:
                # 使用 openpyxl 读取，以提取单元格高亮颜色
                wb = openpyxl.load_workbook(term_excel_path, data_only=True)
                ws = wb.active
                
                # 第一行（openpyxl中为第1行）作为所属法规名称
                law_names = []
                for cell in ws[1]:
                    law_names.append(str(cell.value).strip() if cell.value else "")
                
                terms_list = []
                # 从第二行开始遍历数据
                for row in ws.iter_rows(min_row=2):
                    for c_idx, cell in enumerate(row):
                        if cell.value:
                            cell_str = str(cell.value).strip()
                            if not cell_str or cell_str.lower() == "nan":
                                continue
                            
                            source_law = law_names[c_idx] if c_idx < len(law_names) and law_names[c_idx] else "未知法规"
                            
                            # 【要求1】按照第一个英文冒号切割。冒号前为术语名词，冒号后为定义。
                            parts = cell_str.split(":", 1)
                            if len(parts) == 2:
                                term_name = parts[0].strip()
                                definition = parts[1].strip()
                            else:
                                term_name = cell_str
                                definition = cell_str
                                
                            # 【要求3】获取单元格高亮颜色（用于近似词检索）
                            color = None
                            if cell.fill and cell.fill.start_color:
                                color_val = cell.fill.start_color.index
                                if color_val and str(color_val) != '00000000':
                                    color = str(color_val)
                                    
                            terms_list.append({
                                "term_name": term_name,
                                "definition": definition,
                                "source": source_law,
                                "color": color
                            })
                            
                if term_keyword:
                    # 【要求2】英文检索部分不区分大小写
                    term_keyword_lower = term_keyword.lower()
                    
                    matched_colors = set()
                    direct_match_indices = set()
                    
                    # 第一轮：通过“第一个冒号前”的关键词进行模糊比对
                    for i, t in enumerate(terms_list):
                        if term_keyword_lower in t["term_name"].lower():
                            direct_match_indices.add(i)
                            if t["color"]:
                                matched_colors.add(t["color"])
                                
                    # 第二轮：凡是带有相同高亮颜色（即被分类为同义词）的条文，一并加入检索结果
                    final_results = []
                    for i, t in enumerate(terms_list):
                        if i in direct_match_indices or (t["color"] and t["color"] in matched_colors):
                            final_results.append(t)
                            
                    st.markdown(f"**检索到相关术语/条文 ({len(final_results)} 条)**:")
                    
                    for t_item in final_results:
                        # 确保原文内的所有换行和标点符号完整保留
                        def_html = t_item['definition'].replace('\n', '<br>')
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
                else:
                    st.info("👈 请在上方输入框输入关键词检索。")
            except Exception as e:
                st.error(f"加载术语表异常: {e}")
        else:
            st.warning("未检测到 `术语解释总结.xlsx` 文件，请确认已上传至同一目录。")


    # --- 左侧边栏导航 ---
    st.sidebar.markdown("### 🧭 系统导航")
    nav_mode = st.sidebar.radio(
        "切换功能模块", ["📑 体系化法律库", "🔎 穿透式法规检索", "⏱️ 出境全流程时间轴"]
    )
    st.sidebar.markdown("---")

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
        keyword = st.sidebar.text_input("输入检索关键词", placeholder="如：数据出境、GDPR...")
        st.sidebar.caption("支持模糊搜索法规条款、标签或分类维度。")

        if keyword:
            wildcard = f"%{keyword}%"
            search_query = """
                SELECT region, category, law_title, sub_cat_0, sub_cat_1, content 
                FROM compliance_laws 
                WHERE content LIKE ? OR law_title LIKE ? OR category LIKE ? OR sub_cat_0 LIKE ? OR sub_cat_1 LIKE ?
                ORDER BY region, category, sort_order
            """
            results_df = pd.read_sql(search_query, conn, params=(wildcard,)*5)

            st.markdown(f"**检索结果**：包含 <span style='color:#c0392b; font-weight:bold;'>“{keyword}”</span> 的内容共 **{len(results_df)}** 条", unsafe_allow_html=True)
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
                        
                        highlighted_content = row["content"].replace(keyword, f"<span style='background-color:#ffeaa7; font-weight:bold;'>{keyword}</span>")
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
                            <h4 style="margin-top: 5px; color: #1a5276;">{law_t}</h4>
                            <div class="law-content" style="margin-bottom: 0;">{content}</div>
                        </div>
                    </div>
                    """
                    st.markdown(timeline_card_html, unsafe_allow_html=True)
                st.markdown('</div>', unsafe_allow_html=True)

    conn.close()
