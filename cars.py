import os
import re
import sqlite3
import pandas as pd
import streamlit as st

# ==================== 1. 页面配置 ====================
st.set_page_config(
    page_title="智能网联汽车与跨国数据合规检索平台",
    page_icon="🏛️", # 换成更有学术/法律感的图标
    layout="wide",
)

# ==================== 2. 注入知网风格 CSS ====================
# 参考知网的学术蓝与简约灰白配色
CUSTOM_CSS = """
<style>
    /* 全局字体与背景 */
    html, body, [class*="css"] {
        font-family: "PingFang SC", "Microsoft YaHei", "Helvetica Neue", Helvetica, Arial, sans-serif;
    }
    
    /* 标题颜色 (知网主题蓝) */
    h1, h2, h3, h4 {
        color: #2b579a !important; 
        font-weight: 600 !important;
    }
    
    /* 侧边栏样式优化 */
    [data-testid="stSidebar"] {
        background-color: #f8f9fa;
        border-right: 1px solid #e0e0e0;
    }
    
    /* Expander (折叠面板) 拟物卡片化 */
    div[data-testid="stExpander"] {
        background-color: #ffffff;
        border: 1px solid #e6e9f0;
        border-radius: 8px;
        box-shadow: 0 2px 5px rgba(0,0,0,0.02);
        margin-bottom: 15px;
    }
    div[data-testid="stExpander"] summary {
        background-color: #fbfbfb;
        color: #333333;
        font-weight: bold;
    }
    
    /* 自定义分类标签 (类似知网的关键词 Tag) */
    .cnki-tag {
        display: inline-block;
        background-color: #e8f0fe;
        color: #1a73e8;
        padding: 4px 10px;
        border-radius: 4px;
        font-size: 0.85em;
        font-weight: bold;
        margin-bottom: 8px;
        border: 1px solid #d2e3fc;
    }
    
    /* 法条正文阅读框 */
    .law-content {
        background-color: #fdfdfd;
        border-left: 4px solid #2b579a;
        padding: 12px 18px;
        color: #4a4a4a;
        line-height: 1.7;
        font-size: 0.95em;
        margin-bottom: 10px;
        text-align: justify;
    }
    
    /* 顶部引导语样式 */
    .hero-text {
        font-size: 1.1em;
        color: #555;
        background: #f4f6f9;
        padding: 15px;
        border-radius: 5px;
        border-left: 5px solid #2b579a;
        margin-bottom: 25px;
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

    conn.commit()
    conn.close()
    return True

success = init_database_from_excel()

# ==================== 4. 页面前端展示 ====================
st.title("🏛️ 智能网联汽车跨国数据合规检索平台")
st.markdown(
    '<div class="hero-text">'
    '<b>欢迎使用本平台。</b><br>'
    '本系统集成 <b>中国、欧盟、美国</b> 三大核心司法辖区的完整法律法规、行政法规、部门规章及行业指南，为您提供基于合规场景的模块化导航与多维精准检索支持。'
    '</div>', 
    unsafe_allow_html=True
)

if not success:
    st.error("数据加载失败！请确保 `合规平台条文整理.xlsx` 与本项目代码在同一目录下。")
else:
    # --- 侧边栏设计 ---
    st.sidebar.image("https://upload.wikimedia.org/wikipedia/commons/thumb/a/a2/Scale_of_justice_2.svg/512px-Scale_of_justice_2.svg.png", width=80) 
    st.sidebar.header("检索导航")
    nav_mode = st.sidebar.radio(
        "选择浏览模式", ["📚 体系化模块浏览", "🔍 全文精准检索"]
    )
    st.sidebar.markdown("---")

    conn = sqlite3.connect(DB_FILE)

    if nav_mode == "📚 体系化模块浏览":
        selected_region = st.sidebar.selectbox("📍 司法辖区", ["全部", "中国", "欧盟", "美国"])

        if selected_region == "全部":
            categories_df = pd.read_sql("SELECT DISTINCT category FROM compliance_laws", conn)
        else:
            categories_df = pd.read_sql("SELECT DISTINCT category FROM compliance_laws WHERE region = ?", conn, params=(selected_region,))

        categories = ["全部"] + categories_df["category"].tolist()
        selected_category = st.sidebar.selectbox("📂 效力模块", categories)

        # 动态查询构造
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

        # 头部统计与结果展示
        st.markdown(f"#### 检索结果概览")
        st.caption(f"当前筛选：辖区 [{selected_region}] | 模块 [{selected_category}] —— 共计找到 **{len(module_df)}** 个合规条文片段")
        st.markdown("---")

        grouped = module_df.groupby("law_title")

        for law_title, group in grouped:
            region_name = group.iloc[0]["region"]
            cat_name = group.iloc[0]["category"]
            expander_label = f"📖 【{region_name} - {cat_name}】 {law_title} ({len(group)} 条)"

            with st.expander(expander_label, expanded=False):
                st.markdown(f"#### {law_title}")
                st.caption(f"**归属辖区：** {region_name} &nbsp;|&nbsp; **模块层级：** {cat_name}")
                
                for idx, row in group.reset_index().iterrows():
                    sc0, sc1 = row["sub_cat_0"], row["sub_cat_1"]
                    if sc0 or sc1:
                        tag_content = f"{sc0}" + (f" ➔ {sc1}" if sc1 else "")
                        # 使用 HTML 渲染自定义标签
                        st.markdown(f'<span class="cnki-tag">🏷️ 场景维度：{tag_content}</span>', unsafe_allow_html=True)
                    
                    # 使用 HTML 渲染知网文献式的正文框，替代原先丑陋的 st.text 等宽字体
                    st.markdown(f'<div class="law-content">{row["content"]}</div>', unsafe_allow_html=True)
                    st.write("") # 留白

    else:
        keyword = st.sidebar.text_input("输入检索关键词", placeholder="如：数据出境、GDPR、敏感信息...")
        st.sidebar.caption("支持模糊搜索相关法规条款或分类维度。")

        if keyword:
            wildcard = f"%{keyword}%"
            search_query = """
                SELECT region, category, law_title, sub_cat_0, sub_cat_1, content 
                FROM compliance_laws 
                WHERE content LIKE ? OR law_title LIKE ? OR category LIKE ? OR sub_cat_0 LIKE ? OR sub_cat_1 LIKE ?
                ORDER BY region, category, sort_order
            """
            results_df = pd.read_sql(search_query, conn, params=(wildcard,)*5)

            st.markdown(f"#### 🔍 关键词匹配：<span style='color:#d93025'>“{keyword}”</span>", unsafe_allow_html=True)
            st.caption(f"为您精准匹配到 **{len(results_df)}** 条相关合规内容：")
            st.markdown("---")

            grouped_search = results_df.groupby("law_title")
            for law_title, group in grouped_search:
                region_name, cat_name = group.iloc[0]["region"], group.iloc[0]["category"]
                
                with st.expander(f"📖 【{region_name} - {cat_name}】 {law_title}"):
                    st.markdown(f"#### {law_title}")
                    for idx, row in group.reset_index().iterrows():
                        sc0, sc1 = row["sub_cat_0"], row["sub_cat_1"]
                        if sc0 or sc1:
                            tag_content = f"{sc0}" + (f" ➔ {sc1}" if sc1 else "")
                            st.markdown(f'<span class="cnki-tag">🏷️ 场景维度：{tag_content}</span>', unsafe_allow_html=True)
                        
                        # 在检索时高亮关键字 (增加阅读体验)
                        highlighted_content = row["content"].replace(keyword, f"<span style='background-color:#ffeb3b; font-weight:bold;'>{keyword}</span>")
                        st.markdown(f'<div class="law-content">{highlighted_content}</div>', unsafe_allow_html=True)
                        st.write("")
        else:
            st.info("👈 请在左侧侧边栏输入关键词以获取检索结果。")

    conn.close()
