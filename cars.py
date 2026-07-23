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

  # 向右扫描合并表头（处理 Excel 跨列单元格标题只有左侧有文字的情况）
  current_category_name = ""
  current_region_name = ""

  # 根据表格结构精准映射列索引
  # 遍历每一列，如果第一行有大类名称则更新，如果是 nan 则继承左侧的分类
  column_meta = {}
  active_region = "中国"
  active_cat = "法律"

  for col_idx in range(len(df_raw.columns)):
    val = df_raw.iloc[0, col_idx]
    if pd.notna(val):
      title_str = str(val).strip()
      if "-" in title_str:
        parts = title_str.split("-", 1)
        active_region = parts[0].strip()
        active_cat = parts[1].strip()
      else:
        active_cat = title_str
    # 建立映射
    column_meta[col_idx] = (active_region, active_cat)

  # 手工校准标准列映射以确保万无一失
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

      # 收集该列下所有的非空单元格内容作为条文
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
    "全面对标北大法宝多维分类与目录导航模式，支持点击具体法规名称无损跳转与全文展示。"
)

if not success:
  current_dir = os.path.dirname(os.path.abspath(__file__))
  st.error(
      f"找不到数据文件！请确保 `合规平台条文整理.xlsx` 已经上传并与 `cars.py` 放在同一目录下（当前查找路径："
      f" {os.path.join(current_dir, '合规平台条文整理.xlsx')}）。"
  )
else:
  # 侧边栏：北大法宝风格的树状与模块化导航
  st.sidebar.header("⚖️ 北大法宝风格法规导航")
  nav_mode = st.sidebar.radio(
      "选择浏览模式", ["📚 按地理与效力模块浏览", "🔍 全文关键词精准检索"]
  )

  conn = sqlite3.connect(DB_FILE)

  if nav_mode == "📚 按地理与效力模块浏览":
    st.sidebar.markdown("---")
    regions = ["全部", "中国", "欧盟", "美国"]
    selected_region = st.sidebar.selectbox("📍 选择司法辖区 (地理层面)", regions)

    # 动态获取对应的效力层级/分类（同类规章归到同一个模块）
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

    categories = categories_df["category"].tolist()
    selected_category = st.sidebar.selectbox("📂 选择同类效力模块", categories)

    # 获取该模块下的所有具体法规名称列表（供点击跳转）
    laws_df = pd.read_sql(
        "SELECT DISTINCT law_title FROM compliance_laws WHERE region = ? AND"
        " category = ?",
        conn,
        params=(
            selected_region if selected_region != "All" else regions[1],
            selected_category,
        )
        if selected_region != "全部"
        else ("中国", selected_category),  # 简化默认处理
    )

    # 如果选了全部区域，则按区域和分类组合查询
    if selected_region == "全部":
      laws_df = pd.read_sql(
          "SELECT DISTINCT region, law_title FROM compliance_laws WHERE"
          " category = ?",
          conn,
          params=(selected_category,),
      )
      law_options = [
          f"[{row['region']}] {row['law_title']}"
          for _, row in laws_df.iterrows()
      ]
    else:
      laws_df = pd.read_sql(
          "SELECT DISTINCT law_title FROM compliance_laws WHERE region = ? AND"
          " category = ?",
          conn,
          params=(selected_region, selected_category),
      )
      law_options = laws_df["law_title"].tolist()

    st.sidebar.markdown("---")
    st.sidebar.markdown(f"📖 **当前模块: {selected_category}**")
    st.sidebar.markdown("请点击下方法规名称直接跳转查看详情：")

    # 侧边栏生成法规快速跳转按钮或选择框
    if law_options:
      selected_law_jump = st.sidebar.radio("选择要跳转的法规文件", law_options)
    else:
      selected_law_jump = None

    # 主界面展示：按同类模块归类，并高亮展示用户点击跳转的法规
    st.subheader(
        f"📂 当前模块归档：{selected_category} （共包含"
        f" {len(law_options)} 部法规文件）"
    )
    st.markdown(
        "---"
    )

    # 获取当前分类下的所有具体法规及完整内容
    if selected_region == "全部":
      module_data_query = (
          "SELECT region, category, law_title, content FROM compliance_laws"
          " WHERE category = ?"
      )
      module_data_params = (selected_category,)
    else:
      module_data_query = (
          "SELECT region, category, law_title, content FROM compliance_laws"
          " WHERE region = ? AND category = ?"
      )
      module_data_params = (selected_region, selected_category)

    module_df = pd.read_sql(
        module_data_query, conn, params=module_data_params
    )

    # 按法规名称分组展示，实现“点击/定位跳转”效果
    grouped = module_df.groupby("law_title")

    for law_title, group in grouped:
      # 判断是否是用户在侧边栏选中的那一部法规，如果是则默认展开并醒目标记
      is_target = False
      if selected_law_jump:
        if law_title in selected_law_jump:
          is_target = True

      expander_label = (
          f"📜 【点击跳转/展开】{law_title} （包含 {len(group)}"
          " 项核心条款与细则）"
      )
      with st.expander(expander_label, expanded=is_target):
        st.markdown(
            f"**所属辖区：** {group.iloc[0]['region']}  |  **效力层级模块：**"
            f" **{selected_category}**"
        )
        st.markdown(f"**法规全称：** **{law_title}**")
        st.markdown("---")
        st.markdown("### 📋 完整法规条文与内容（未做任何省略）:")

        # 循环展示该法规在表格中对应的每一个具体条文/单元格内容
        for idx, row in group.reset_index().iterrows():
          st.markdown(f"**【条款 / 细则 {idx+1}】**")
          st.text(row["content"])
          st.markdown("")

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
              f"**所属辖区：** {row['region']} | **分类模块：**"
              f" {row['category']}"
          )
          st.markdown(f"**法规全称：** **{row['law_title']}**")
          st.markdown("---")
          st.markdown("**详细条文内容：**")
          st.text(row["content"])
    else:
      st.info("👈 请在左侧侧边栏输入关键词开始进行全文精准检索。")

  conn.close()
