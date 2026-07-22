import os
import pandas as pd
import sqlite3
import streamlit as st

# 数据库文件名字
DB_FILE = "car_compliance.db"


# 自动解析并加载 Excel 表格初始化数据库（适配同学们的最细分类）
def init_database_from_excel():
  excel_path = r"E:\大创\cars\合规平台条文整理.xlsx"
  if not os.path.exists(excel_path):
    st.error(
        f"找不到数据文件 {excel_path}，请确保该 Excel 文件与 cars.py 放在同级目录下！"
    )
    return

  # 读取 Excel（无表头模式，方便我们精准按行列坐标解析）
  df_raw = pd.read_excel(excel_path, sheet_name=0, header=None)

  # 解析表头（第 0 行是法域/法规类别大类，第 1 行是具体法规名称）
  category_row = df_raw.iloc[0]
  law_name_row = df_raw.iloc[1]

  parsed_records = []

  # 从第 2 行开始是具体的分类标签与法条内容
  for r_idx in range(2, len(df_raw)):
    row_data = df_raw.iloc[r_idx]
    # Col 0 和 Col 1 通常是主题大类和细分标签（如：中国出境 -> 数据分类/出境路径选择）
    theme_col = str(row_data.iloc[0]).strip() if pd.notna(row_data.iloc[0]) else ""
    sub_tag = str(row_data.iloc[1]).strip() if pd.notna(row_data.iloc[1]) else ""

    # 遍历每一列（代表不同的法规）
    for c_idx in range(3, len(df_raw.columns)):
      cell_value = row_data.iloc[c_idx]
      if pd.notna(cell_value) and str(cell_value).strip() != "":
        # 获取大类（通过向前填充处理合并单元格，如果为空则用上一行的主题）
        jurisdiction_category = (
            str(category_row.iloc[c_idx]).strip()
            if pd.notna(category_row.iloc[c_idx])
            else ""
        )
        law_title = (
            str(law_name_row.iloc[c_idx]).strip()
            if pd.notna(law_name_row.iloc[c_idx])
            else f"法规_Col{c_idx}"
        )

        # 简单的法域归类识别
        jurisdiction = "中国"
        if "欧盟" in jurisdiction_category or "GDPR" in law_title:
          jurisdiction = "欧盟"
        elif "美国" in jurisdiction_category or "CCPA" in law_title:
          jurisdiction = "美国"

        parsed_records.append({
            "jurisdiction_category": jurisdiction_category,  # 如：中国-法律、中国-行政法规、欧盟-条例
            "jurisdiction": jurisdiction,  # 中国 / 欧盟 / 美国
            "law_title": law_title,  # 具体法规名称（如：中华人民共和国网络安全法）
            "theme": theme_col,  # 综合主题
            "sub_tag": sub_tag,  # 同学们整理的细分分类标签（如：数据分类、出境路径选择等）
            "full_law_content": str(cell_value).strip(),  # 法条全文
        })

  df_final = pd.DataFrame(parsed_records)

  # 写入 SQLite 数据库
  conn = sqlite3.connect(DB_FILE)
  df_final.to_sql("rules_cases_v2", conn, if_exists="replace", index=False)
  conn.close()


# 每次运行或重启时自动从 Excel 重建数据库，确保与同学们的最新分类保持同步
init_database_from_excel()

# ==================== Streamlit 网页前端界面 ====================
st.set_page_config(
    page_title="智能网联汽车数据跨境双向合规平台", layout="wide"
)

st.title("🚗 智能网联汽车数据跨境双向合规检索平台（北大法宝版）")
st.markdown(
    "基于最新细分法条整理表格：输入关键词检索相关法规与法条，点击展开即可查看**原原本本的法条条文全文**。"
)

# 侧边栏：多维度检索与分类过滤
st.sidebar.header("🔍 多维合规检索与过滤")

# 1. 法域筛选
jurisdiction_list = ["全部", "中国", "欧盟", "美国"]
selected_jurisdiction = st.sidebar.selectbox("选择法域", jurisdiction_list)

# 2. 动态获取对应法域下的法规名称供用户进一步精准点选
conn = sqlite3.connect(DB_FILE)
if selected_jurisdiction == "全部":
  laws_df = pd.read_sql("SELECT DISTINCT law_title FROM rules_cases_v2", conn)
else:
  laws_df = pd.read_sql(
      "SELECT DISTINCT law_title FROM rules_cases_v2 WHERE jurisdiction = ?",
      conn,
      params=(selected_jurisdiction,),
  )
conn.close()

law_options = ["全部法规"] + laws_df["law_title"].tolist()
selected_law = st.sidebar.selectbox("选择具体法律法规", law_options)

# 主页面搜索框
search_keyword = st.text_input(
    "请输入关键词（支持全文模糊搜索，例如：出境、重要数据、人脸、安全评估、GDPR等...）：",
    placeholder="在此输入任意关键词进行检索...",
)


# 核心：多条件联合查询函数
def search_database(keyword, jurisdiction, law_title):
  conn = sqlite3.connect(DB_FILE)
  query = "SELECT jurisdiction_category, jurisdiction, law_title, theme, sub_tag, full_law_content FROM rules_cases_v2 WHERE 1=1"
  params = []

  if jurisdiction != "全部":
    query += " AND jurisdiction = ?"
    params.append(jurisdiction)

  if law_title != "全部法规":
    query += " AND law_title = ?"
    params.append(law_title)

  if keyword:
    query += (
        " AND (full_law_content LIKE ? OR law_title LIKE ? OR theme LIKE ? OR"
        " sub_tag LIKE ?)"
    )
    like_pattern = f"%{keyword}%"
    params.extend([like_pattern, like_pattern, like_pattern, like_pattern])

  df_result = pd.read_sql(query, conn, params=params)
  conn.close()
  return df_result


# 执行查询
result_df = search_database(
    search_keyword, selected_jurisdiction, selected_law
)

st.divider()
st.subheader(f"📋 检索结果 (共找到 {len(result_df)} 条相关法条记录)")

if not result_df.empty:
  for index, row in result_df.iterrows():
    with st.container():
      col1, col2 = st.columns([1, 4])
      with col1:
        st.markdown(f"**法规分类:** `{row['jurisdiction_category']}`")
        st.markdown(f"**法域:** `{row['jurisdiction']}`")
        if row["sub_tag"]:
          st.markdown(f"**细分标签:** `{row['sub_tag']}`")
      with col2:
        # 标题显示具体法规名称
        st.markdown(f"### 📌 {row['law_title']}")
        if row["theme"]:
          st.caption(f"当前主题大类：{row['theme']}")

        # 北大法宝核心交互：点击展开查看原原本本的法条内容
        with st.expander(
            "📖 点击展开：查看该分类下的【完整法定条文原文】",
            expanded=(index == 0),
        ):
          st.text(row["full_law_content"])

      st.divider()
else:
  st.warning("没有找到符合条件的记录，请尝试更换关键词或放宽筛选条件。")
