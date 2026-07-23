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


def parse_and_split_content(cell_text):
  """针对欧盟及长文本单元格，按 Article / 条 / 细则关键词进行智能拆分"""
  if not cell_text or str(cell_text).strip() == "nan":
    return []

  text = str(cell_text).strip()
  pattern = r"(?=(?:Article\s+\d+|第[零一二三四五六七八九十百0-9]+条|Step\s+\d+))"
  parts = re.split(pattern, text)

  cleaned_parts = [p.strip() for p in parts if p.strip()]
  if not cleaned_parts:
    return [text]
  return cleaned_parts


def extract_sort_key(text):
  """提取条款编号用于正序排序"""
  match_cn = re.search(r"第([零一二三四五六七八九十百0-9]+)条", text)
  if match_cn:
    num_str = match_cn.group(1)
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
      pass

  match_en = re.search(r"Article\s+(\d+)", text, re.IGNORECASE)
  if match_en:
    try:
      return int(match_en.group(1))
    except ValueError:
      pass

  return 999


def init_database_from_excel():
  """自动从同级目录的 Excel 初始化 SQLite 数据库，同时记录横向分类与纵向分类标签"""
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

  # 动态前向填充横向分类（第0行）和纵向分类（第0列、第1列）
  categories_row = df_raw.iloc[0].ffill()
  titles_row = df_raw.iloc[1]

  col0_ffill = df_raw.iloc[:, 0].ffill()
  col1_ffill = df_raw.iloc[:, 1].ffill()

  # 遍历从第 3 列开始的所有法规列
  for col_idx in range(3, len(df_raw.columns)):
    cat_name = str(categories_row.iloc[col_idx]).strip()
    law_title = str(titles_row.iloc[col_idx]).strip()

    if not law_title or law_title == "nan":
      continue

    region = "中国"
    if "欧盟" in cat_name:
      region = "欧盟"
    elif "美国" in cat_name:
      region = "美国"

    category = cat_name if cat_name and cat_name != "nan" else "通用效力模块"

    has_content = False
    for row_idx in range(2, len(df_raw)):
      cell_val = df_raw.iloc[row_idx, col_idx]

      # 获取纵向分类标签（第0列和第1列对应的纵向指引）
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
                """
                        INSERT INTO compliance_laws (region, category, law_title, sub_cat_0, sub_cat_1, content, sort_order)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                    """,
                (
                    region,
                    category,
                    law_title,
                    sub_c0,
                    sub_c1,
                    content_str,
                    sort_val,
                ),
            )

    if not has_content:
      cursor.execute(
          """
                INSERT INTO compliance_laws (region, category, law_title, sub_cat_0, sub_cat_1, content, sort_order)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
          (
              region,
              category,
              law_title,
              "暂无分类",
              "暂无指引",
              "（该法规条文正在整理补充中，敬请期待...）",
              999,
          ),
      )

  conn.commit()
  conn.close()
  return True


success = init_database_from_excel()

# ==================== 页面前端设计 ====================
st.title("🚗 智能网联汽车与跨国数据合规检索平台")
st.markdown(
    "> 本平台集成 **中国、欧盟、美国** 三大核心司法辖区的完整法律法规、行政法规、部门规章及行业指南，"
    "支持北大法宝风格模块化导航与多维精准检索。"
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

    if selected_region == "全部" and selected_category == "全部":
      module_data_query = (
          "SELECT region, category, law_title, sub_cat_0, sub_cat_1, content"
          " FROM compliance_laws ORDER BY region, category, sort_order"
      )
      module_data_params = ()
    elif selected_region == "全部":
      module_data_query = (
          "SELECT region, category, law_title, sub_cat_0, sub_cat_1, content"
          " FROM compliance_laws WHERE category = ? ORDER BY region, sort_order"
      )
      module_data_params = (selected_category,)
    elif selected_category == "全部":
      module_data_query = (
          "SELECT region, category, law_title, sub_cat_0, sub_cat_1, content"
          " FROM compliance_laws WHERE region = ? ORDER BY category, sort_order"
      )
      module_data_params = (selected_region,)
    else:
      module_data_query = (
          "SELECT region, category, law_title, sub_cat_0, sub_cat_1, content"
          " FROM compliance_laws WHERE region = ? AND category = ? ORDER BY"
          " sort_order"
      )
      module_data_params = (selected_region, selected_category)

    module_df = pd.read_sql(
        module_data_query, conn, params=module_data_params
    )

    st.subheader(
        f"📂 当前检索模块：辖区 [{selected_region}] | 模块 [{selected_category}]"
        f" （共找到 {len(module_df)} 部/条合规文件）"
    )
    st.markdown("---")

    grouped = module_df.groupby("law_title")

    for law_title, group in grouped:
      region_name = group.iloc[0]["region"]
      cat_name = group.iloc[0]["category"]

      expander_label = (
          f"📜 【{region_name} - {cat_name}】 {law_title} （包含"
          f" {len(group)} 项条款）"
      )

      with st.expander(expander_label, expanded=True):
        st.markdown(f"### 📌 法规全称：**{law_title}**")
        st.markdown(
            f"**所属司法辖区：** {region_name}  |  **效力层级模块：** {cat_name}"
        )
        st.markdown("---")

        for idx, row in group.reset_index().iterrows():
          sc0 = row["sub_cat_0"]
          sc1 = row["sub_cat_1"]
          # 醒目展示纵向分类标签指引
          if sc0 or sc1:
            tag_str = (
                f"🏷️ **分类指引维度：** `{sc0}`"
                + (f" ➔ `{sc1}`" if sc1 else "")
            )
            st.markdown(tag_str)
          st.text(row["content"])
          st.markdown("---")

  else:
    st.sidebar.markdown("---")
    keyword = st.sidebar.text_input(
        "请输入要检索的关键词（如：出境、数据定性、敏感个人信息、GDPR等）"
    )

    if keyword:
      search_query = """
                SELECT region, category, law_title, sub_cat_0, sub_cat_1, content 
                FROM compliance_laws 
                WHERE content LIKE ? OR law_title LIKE ? OR category LIKE ? OR sub_cat_0 LIKE ? OR sub_cat_1 LIKE ?
                ORDER BY region, category, sort_order
            """
      wildcard = f"%{keyword}%"
      results_df = pd.read_sql(
          search_query,
          conn,
          params=(wildcard, wildcard, wildcard, wildcard, wildcard),
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
            sc0 = row["sub_cat_0"]
            sc1 = row["sub_cat_1"]
            if sc0 or sc1:
              tag_str = (
                  f"🏷️ **分类指引维度：** `{sc0}`"
                  + (f" ➔ `{sc1}`" if sc1 else "")
              )
              st.markdown(tag_str)
            st.text(row["content"])
            st.markdown("---")
    else:
      st.info("👈 请在左侧侧边栏输入关键词开始进行全文精准检索。")

  conn.close()
