import os
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


def init_database_from_excel():
  """自动从同级目录的 Excel 初始化 SQLite 数据库，确保所有完整信息不丢失"""
  current_dir = os.path.dirname(os.path.abspath(__file__))
  excel_path = os.path.join(current_dir, "合规平台条文整理.xlsx")

  conn = sqlite3.connect(DB_FILE)
  cursor = conn.cursor()

  # 创建合规条目表
  cursor.execute("""
        CREATE TABLE IF NOT EXISTS compliance_laws (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            region TEXT,
            category TEXT,
            law_title TEXT,
            content TEXT
        )
    """)

  if not os.path.exists(excel_path):
    conn.close()
    return False

  # 读取原始表格
  df_raw = pd.read_excel(excel_path, sheet_name=0, header=None)

  # 清空旧数据以防重复插入
  cursor.execute("DELETE FROM compliance_laws")

  # 定义列对应的地理大区与分类映射
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
            cursor.execute(
                """
                        INSERT INTO compliance_laws (region, category, law_title, content)
                        VALUES (?, ?, ?, ?)
                    """,
                (region, category, law_title, content_str),
            )

  conn.commit()
  conn.close()
  return True


# 每次运行强制初始化并加载数据
success = init_database_from_excel()

# ==================== 页面前端设计 ====================
st.title("🚗 智能网联汽车与跨国数据合规检索平台")
st.markdown(
    "> 本平台集成 **中国、欧盟、美国** 三大核心司法辖区的完整法律法规、行政法规、部门规章及行业指南，"
    "支持多维地理层级跳转与全文本检索，提供真实的法律条文与核心摘要。"
)

if not success:
  current_dir = os.path.dirname(os.path.abspath(__file__))
  st.error(
      f"找不到数据文件！请确保 `合规平台条文整理.xlsx` 已经上传并与 `cars.py` 放在同一目录下（当前查找路径："
      f" {os.path.join(current_dir, '合规平台条文整理.xlsx')}）。"
  )
else:
  # 侧边栏：北大法宝风格的地理与层级导航模块
  st.sidebar.header("⚖️ 合规数据库导航")
  nav_mode = st.sidebar.radio(
      "选择浏览模式", ["🌍 地理层面与效力层级导航", "🔍 全文关键词精准检索"]
  )

  conn = sqlite3.connect(DB_FILE)

  if nav_mode == "🌍 地理层面与效力层级导航":
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
    selected_category = st.sidebar.selectbox("📂 选择效力层级/分类", categories)

    query_str = "SELECT DISTINCT law_title FROM compliance_laws WHERE 1=1"
    params = []
    if selected_region != "全部":
      query_str += " AND region = ?"
      params.append(selected_region)
    if selected_category != "全部":
      query_str += " AND category = ?"
      params.append(selected_category)

    laws_df = pd.read_sql(query_str, conn, params=params)
    law_titles = ["全部"] + laws_df["law_title"].tolist()
    selected_law = st.sidebar.selectbox("📖 选择具体法律文件", law_titles)

    main_query = (
        "SELECT region, category, law_title, content FROM compliance_laws"
        " WHERE 1=1"
    )
    main_params = []
    if selected_region != "全部":
      main_query += " AND region = ?"
      main_params.append(selected_region)
    if selected_category != "全部":
      main_query += " AND category = ?"
      main_params.append(selected_category)
    if selected_law != "全部":
      main_query += " AND law_title = ?"
      main_params.append(selected_law)

    results_df = pd.read_sql(main_query, conn, params=main_params)

    st.subheader(
        f"📊 检索结果展示 (当前筛选: 辖区 [{selected_region}] | 层级"
        f" [{selected_category}] | 文件 [{selected_law}])"
    )
    st.markdown(f"共找到 **{len(results_df)}** 条相关合规条文/核心内容：")

    for idx, row in results_df.iterrows():
      with st.expander(
          f"【{row['region']} - {row['category']}】 {row['law_title']} (条文摘要"
          f" {idx+1})"
      ):
        st.markdown(f"**所属辖区：** {row['region']}")
        st.markdown(f"**效力分类：** {row['category']}")
        st.markdown(f"**法律名称：** **{row['law_title']}**")
        st.markdown("---")
        st.markdown("**条文与内容详情：**")
        st.text(row["content"])

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
            """
      wildcard = f"%{keyword}%"
      results_df = pd.read_sql(
          search_query, conn, params=(wildcard, wildcard, wildcard)
      )

      st.subheader(f"🔍 关键词 “{keyword}” 的检索结果")
      st.markdown(f"为您匹配到 **{len(results_df)}** 条相关合规内容：")

      for idx, row in results_df.iterrows():
        with st.expander(
            f"【{row['region']} - {row['category']}】 {row['law_title']}"
        ):
          st.markdown(
              f"**所属辖区：** {row['region']} | **分类：** {row['category']}"
          )
          st.markdown(f"**法律全称：** **{row['law_title']}**")
          st.markdown("---")
          st.markdown("**详细内容：**")
          st.text(row["content"])
    else:
      st.info("👈 请在左侧侧边栏输入关键词开始进行全文精准检索。")

  conn.close()
