import os
import re
import sqlite3
import pandas as pd
import streamlit as st

# 设置页面配置
st.set_page_config(
    page_title="智能网联汽车与跨国数据合规检索平台",
    page_icon="🚗",
    layout="wide",
)

DB_FILE = "car_compliance.db"


def extract_article_number(text):
  """提取文本中的条款编号（如“第二条”、“第14条”），用于精准排序"""
  match = re.search(r"第([零一二三四五六七八九十百0-9]+)条", text)
  if match:
    num_str = match.group(1)
    mapping = {
        "一": 1,
        "二": 2,
        "三": 3,
        "四": 4,
        "五": 5,
        "六": 6,
        "七": 7,
        "八": 8,
        "九": 9,
        "十": 10,
        "十一": 11,
        "十二": 12,
        "十三": 13,
        "十四": 14,
        "十五": 15,
        "十六": 16,
        "十七": 17,
        "十八": 18,
        "十九": 19,
        "二十": 20,
    }
    if num_str in mapping:
      return mapping[num_str]
    try:
      return int(num_str)
    except ValueError:
      return 999
  return 999


def init_database_from_excel():
  """自动从同级目录的 Excel 初始化 SQLite 数据库"""
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
            content TEXT,
            sort_order INTEGER
        )
    """)

  if not os.path.exists(excel_path):
    conn.close()
    return False

  df_raw = pd.read_excel(excel_path, sheet_name=0, header=None)
  cursor.execute("DELETE FROM compliance_laws")

  region_mapping = {
      3: ("中国", "法律"),
      4: ("中国", "法律"),
      5: ("中国", "法律"),
      6: ("中国", "行政法规"),
      7: ("中国", "行政法规"),
      8: ("中国", "部门规章"),
      9: ("中国", "部门规章"),
      10: ("中国", "部门规章"),
      11: ("中国", "规范性文件"),
      12: ("中国", "行业特别规定"),
      13: ("中国", "配套操作指引"),
      14: ("中国", "配套操作指引"),
      15: ("欧盟", "欧盟-条例"),
      16: ("欧盟", "欧盟-次级立法"),
      17: ("欧盟", "欧盟-次级立法"),
      18: ("欧盟", "欧盟-次级立法"),
      19: ("欧盟", "欧盟-次级立法"),
      20: ("欧盟", "欧盟-次级立法"),
      21: ("欧盟", "欧盟-指南/建议"),
      22: ("欧盟", "欧盟-指南/建议"),
      23: ("美国", "美国-国家安全层面"),
      24: ("美国", "美国-联邦层面"),
      25: ("美国", "加州"),
  }

  for col_idx, (region, category) in region_mapping.items():
    if col_idx < len(df_raw.columns):
      law_title = str(df_raw.iloc[0, col_idx]).strip()
      if not law_title or law_title == "nan":
        continue

      for row_idx in range(1, len(df_raw)):
        cell_val = df_raw.iloc[row_idx, col_idx]
        if pd.notna(cell_val):
          content_str = str(cell_val).strip()
          if content_str and content_str != "nan":
            sort_val = extract_article_number(content_str)
            cursor.execute(
                """
                        INSERT INTO compliance_laws (region, category, law_title, content, sort_order)
                        VALUES (?, ?, ?, ?, ?)
                    """,
                (region, category, law_title, content_str, sort_val),
            )

  conn.commit()
  conn.close()
  return True


success = init_database_from_excel()

# ==================== 页面前端设计 ====================
st.title("🚗 智能网联汽车与跨国数据合规检索平台")
st.markdown(
    "> 本平台集成 **中国、欧盟、美国** 三大核心司法辖区的完整法律法规、行政法规、部门规章及行业指南，"
    "支持多维地理与效力模块化导航，法规名称提纲挈领置顶，下属条文按序号自动排序展示。"
)

if not success:
  current_dir = os.path.dirname(os.path.abspath(__file__))
  st.error(
      f"找不到数据文件！请确保 `合规平台条文整理.xlsx` 已经上传并与 `cars.py` 放在同一目录下（当前查找路径："
      f" {os.path.join(current_dir, '合规平台条文整理.xlsx')}）。"
  )
else:
  st.sidebar.header("⚖️ 北大法宝风格法规导航")
  nav_mode = st.sidebar.radio(
      "选择浏览模式", ["📚 按地理与效力模块浏览", "🔍 全文关键词精准检索"]
  )

  conn = sqlite3.connect(DB_FILE)

  if nav_mode == "📚 按地理与效力模块浏览":
    st.sidebar.markdown("---")
    regions = ["全部", "中国", "欧盟", "美国"]
    selected_region = st.sidebar.selectbox("📍 选择司法辖区 (地理层面)", regions)

    if selected_region == "全部":
      categories_df = pd.read_sql(
          "SELECT DISTINCT category FROM compliance_laws", conn
      )
    else:
      categories_df = pd.read_sql(
          "SELECT DISTINCT category FROM compliance_laws WHERE region = ?",
          conn,
          params=(selected_region,),
      )

    categories = ["全部"] + categories_df["category"].tolist()
    selected_category = st.sidebar.selectbox("📂 选择同类效力模块", categories)

    st.sidebar.markdown("---")

    # 查询筛选结果
    if selected_region == "全部" and selected_category == "全部":
      module_data_query = (
          "SELECT region, category, law_title, content FROM compliance_laws"
          " ORDER BY region, category, sort_order"
      )
      module_data_params = ()
    elif selected_region == "全部":
      module_data_query = (
          "SELECT region, category, law_title, content FROM compliance_laws"
          " WHERE category = ? ORDER BY region, sort_order"
      )
      module_data_params = (selected_category,)
    elif selected_category == "全部":
      module_data_query = (
          "SELECT region, category, law_title, content FROM compliance_laws"
          " WHERE region = ? ORDER BY category, sort_order"
      )
      module_data_params = (selected_region,)
    else:
      module_data_query = (
          "SELECT region, category, law_title, content FROM compliance_laws"
          " WHERE region = ? AND category = ? ORDER BY sort_order"
      )
      module_data_params = (selected_region, selected_category)

    module_df = pd.read_sql(
        module_data_query, conn, params=module_data_params
    )

    st.subheader(
        f"📂 当前检索模块：辖区 [{selected_region}] | 模块 [{selected_category}]"
        f" （共找到 {len(module_df)} 条相关条文）"
    )
    st.markdown("---")

    # 按法规名称（《》名称作为提纲主标题）分组展示
    grouped = module_df.groupby("law_title")

    for law_title, group in grouped:
      region_name = group.iloc[0]["region"]
      cat_name = group.iloc[0]["category"]

      # 将法规名称（大标题）作为折叠面板的核心置顶标题，下属所有条文在内部按顺序展开
      expander_label = (
          f"📜 【{region_name} - {cat_name}】 {law_title} （包含"
          f" {len(group)} 项细则条款）"
      )

      with st.expander(expander_label, expanded=True):
        st.markdown(
            f"### 📌 法规全称：**{law_title}**"
        )  # 再次在内部醒目强调法律法规名称作为提纲
        st.markdown(
            f"**所属司法辖区：** {region_name}  |  **效力层级模块：** {cat_name}"
        )
        st.markdown("---")
        st.markdown("**具体条文内容（已按序号正序排列）：**")

        for idx, row in group.reset_index().iterrows():
          st.text(row["content"])
          st.markdown("---")

  else:
    st.sidebar.markdown("---")
    keyword = st.sidebar.text_input(
        "请输入要检索的关键词（如：出境、分类分级、敏感个人信息、GDPR等）"
    )

    if keyword:
      search_query = """
                SELECT region, category, law_title, content 
                FROM compliance_laws 
                WHERE content LIKE ? OR law_title LIKE ? OR category LIKE ?
                ORDER BY region, category, sort_order
            """
      wildcard = f"%{keyword}%"
      results_df = pd.read_sql(
          search_query, conn, params=(wildcard, wildcard, wildcard)
      )

      st.subheader(f"🔍 关键词 “{keyword}” 的检索结果")
      st.markdown(f"为您匹配到 **{len(results_df)}** 条相关合规内容：")

      grouped_search = results_df.groupby("law_title")
      for law_title, group in grouped_search:
        region_name = group.iloc[0]["region"]
        cat_name = group.iloc[0]["category"]
        with st.expander(f"【{region_name} - {cat_name}】 {law_title}"):
          st.markdown(f"### 📌 法规全称：**{law_title}**")
          st.markdown(f"**所属辖区：** {region_name} | **分类模块：** {cat_name}")
          st.markdown("---")
          for idx, row in group.reset_index().iterrows():
            st.text(row["content"])
            st.markdown("---")
    else:
      st.info("👈 请在左侧侧边栏输入关键词开始进行全文精准检索。")

  conn.close()
