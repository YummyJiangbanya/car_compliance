import os
import re
import sqlite3
import pandas as pd
import streamlit as st

# ==================== 1. 页面配置 ====================
st.set_page_config(
    page_title="智能网联汽车与跨国数据合规检索平台",
    page_icon="⚖️", 
    layout="wide",
)

# ==================== 2. 注入北大法宝风格 CSS ====================
CUSTOM_CSS = """
<style>
    /* 全局字体 */
    html, body, [class*="css"] {
        font-family: "PingFang SC", "Microsoft YaHei", "Helvetica Neue", sans-serif;
    }
    
    /* 更改整个应用的背景色为浅灰蓝色，打破大面积留白 */
    [data-testid="stAppViewContainer"] {
        background-color: #f0f2f6; 
    }
    
    /* 顶部标题区卡片化 */
    .header-card {
        background-color: #ffffff;
        padding: 25px 30px;
        border-radius: 10px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.05);
        border-top: 5px solid #1a5276; /* 北大蓝强调色 */
        margin-bottom: 25px;
    }
    
    /* 标题颜色 */
    h1, h2, h3 {
        color: #1a5276 !important; 
        font-weight: 600 !important;
    }
    
    /* 侧边栏样式优化 */
    [data-testid="stSidebar"] {
        background-color: #ffffff;
        border-right: 1px solid #e0e0e0;
        box-shadow: 2px 0 10px rgba(0,0,0,0.02);
    }
    
    /* Expander (折叠面板) 拟物卡片化 */
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
    
    /* 自定义分类标签 */
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
    
    /* 法条正文阅读框 */
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

    # 新增：出境全流程专属数据表
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS timeline_laws (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            phase TEXT,
            sub_phase TEXT,
            law_title TEXT,
            content TEXT
        )
    """)

    if not os.path.exists(excel_path):
        conn.close()
        return False

    df_raw = pd.read_excel(excel_path, sheet_name=0, header=None)
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

    # 尝试加载第二张表（或适配含有出境全流程的Sheet），如果您的出境全流程放在了特定的Sheet中，可在此自动读取
    try:
        excel_sheets = pd.ExcelFile(excel_path).sheet_names
        target_sheet = None
        for sheet in excel_sheets:
            if "流程" in sheet or "时间轴" in sheet or "出境" in sheet:
                target_sheet = sheet
                break
        
        if target_sheet:
            df_timeline = pd.read_excel(excel_path, sheet_name=target_sheet)
            cursor.execute("DELETE FROM timeline_laws")
            for _, row in df_timeline.iterrows():
                # 兼容表格字段：阶段、环节/子阶段、法规名称/标题、完整条款内容
                phase = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else "未分类"
                sub_phase = str(row.iloc[1]).strip() if len(row) > 1 and pd.notna(row.iloc[1]) else ""
                law_title = str(row.iloc[2]).strip() if len(row) > 2 and pd.notna(row.iloc[2]) else "相关合规要求"
                content = str(row.iloc[3]).strip() if len(row) > 3 and pd.notna(row.iloc[3]) else str(row.iloc[-1])
                
                cursor.execute(
                    "INSERT INTO timeline_laws (phase, sub_phase, law_title, content) VALUES (?, ?, ?, ?)",
                    (phase, sub_phase, law_title, content)
                )
    except Exception:
        pass

    conn.commit()
    conn.close()
    return True

success = init_database_from_excel()

# ==================== 4. 页面前端展示 ====================
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
    # --- 侧边栏设计 ---
    st.sidebar.markdown("### 🧭 系统导航")
    
    nav_mode = st.sidebar.radio(
        "切换功能模块", ["📑 体系化法律库", "🔎 穿透式法规检索", "⏱️ 出境全流程"]
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

    elif nav_mode == "⏱️ 出境全流程":
        st.markdown("### ⏱️ 数据出境全流程合规指引")
        st.markdown("按照数据出境的生命周期（**出境前、出境中、出境后**），为您全景展示各环节对应的合规要求、具体条款及操作指引。")
        st.write("")

        # 从数据库中读取时间轴/出境全流程数据
        timeline_df = pd.read_sql("SELECT phase, sub_phase, law_title, content FROM timeline_laws", conn)

        if timeline_df.empty:
            st.warning("⚠️ 暂未检测到“出境全流程”的明细数据。请检查您的Excel表格中是否包含对应维度的内容，或确认数据已被成功导入数据库。")
        else:
            # 设定固定的三个主阶段标签页
            phases = ["出境前", "出境中", "出境后"]
            tabs = st.tabs([f"📍 {p}" for p in phases])

            for i, phase_name in enumerate(phases):
                with tabs[i]:
                    # 过滤出当前阶段的数据
                    phase_data = timeline_df[timeline_df["phase"].str.contains(phase_name, na=False)]

                    if phase_data.empty:
                        st.info(f"暂无【{phase_name}】阶段的相关合规条文数据。")
                    else:
                        # 按照子阶段或法规名称进行分组展示，提升层次感
                        grouped_phase = phase_data.groupby("sub_phase" if "sub_phase" in phase_data.columns else "law_title")
                        
                        for sub_key, group in grouped_phase:
                            display_title = sub_key if sub_key and sub_key != "nan" else "核心合规指引"
                            st.markdown(f"#### 📌 {display_title}")
                            
                            for _, row in group.iterrows():
                                law_t = row.get("law_title", "")
                                content = row.get("content", "")
                                
                                if law_t and law_t != "nan" and law_t != display_title:
                                    st.markdown(f"**📜 法规来源/条款**：{law_t}")
                                
                                st.markdown(f'<div class="law-content">{content}</div>', unsafe_allow_html=True)
                            st.markdown("---")

    conn.close()
