import os
import pandas as pd
import sqlite3
import streamlit as st

# 1. 数据库初始化文件名字
DB_FILE = "car_compliance.db"


# 自动生成测试数据库（包含法条和案例，并用 category 字段区分）
def init_sample_database():
  if not os.path.exists(DB_FILE):
    data = {
        "id": ["CN_01", "EU_01", "CASE_01"],
        "category": ["中国法条", "欧盟法条", "典型案例"],
        "jurisdiction": ["中国", "欧盟", "欧盟/跨国"],
        "data_type": ["车外实景影像/人脸/车牌", "个人数据/生物识别", "车外摄像头/人脸记录"],
        "scenario": ["联网运行/出境", "算法训练/出境", "内部测试/日常运营"],
        "rule_title": [
            "《汽车数据出境安全指引(2026版)》",
            "GDPR 第9条 & 汽车场景",
            "TikTok / 大众汽车车载摄像头案",
        ],
        "core_content": [
            (
                "包含人脸信息、车牌信息的车外视频、图像数据属于重要数据，出境"
                "需严格履行安全评估程序。"
            ),
            (
                "行人面部特征、车牌号属于个人数据与生物识别数据。原则上禁止处理，"
                "除非获得明确同意或符合法定豁免。"
            ),
            (
                "企业因未能正确告知测试驾驶员车载摄像头正在记录、未充分评估风险及"
                "缺乏合法性基础，遭监管机构重罚。"
            ),
        ],
        "compliance_action": [
            "属于重要数据，禁止直接无序出境，须申报数据出境安全评估。",
            (
                "触发GDPR禁令，需严格脱敏或获得数据主体明确授权，且面临严格的"
                "TIA（传输影响评估）。"
            ),
            (
                "必须落实透明度义务，强化告知同意与DPIA（数据保护影响评估），"
                "避免高额罚款。"
            ),
        ],
    }
    df = pd.DataFrame(data)
    conn = sqlite3.connect(DB_FILE)
    df.to_sql("rules_cases", conn, if_exists="replace", index=False)
    conn.close()


# 初始化数据库
init_sample_database()

# ==================== Streamlit 网页前端界面 ====================
st.set_page_config(
    page_title="智能网联汽车数据跨境双向合规平台", layout="wide"
)

st.title("🚗 智能网联汽车车外实景影像数据跨境双向合规检索平台")
st.markdown(
    "针对车企出海痛点，内置中国法、欧盟GDPR及典型案例。支持**纯文字字段与模糊搜索**，助您一键检索合规要求。"
)

# 侧边栏：模块选择与法域筛选
st.sidebar.header("🔍 检索与分类选项")

# 1. 模块/类别选择（实现“法条和案例分成两个模块”的需求）
category_filter = st.sidebar.selectbox(
    "选择检索模块", ["全部展示", "法条专区", "案例专区"]
)

# 2. 法域筛选
jurisdiction_filter = st.sidebar.selectbox(
    "选择法域", ["全部", "中国", "欧盟", "欧盟/跨国"]
)

# 主页面搜索框：优化提示文字
search_keyword = st.text_input(
    "请输入关键词（支持模糊搜索，例如：车外、人脸、GDPR、安全评估、罚款等...）：",
    placeholder="在此输入任意关键词进行检索...",
)


# 核心：多条件联合查询（处理模块、法域、模糊搜索）
def search_database(keyword, category, jurisdiction):
  conn = sqlite3.connect(DB_FILE)

  query = "SELECT category, jurisdiction, data_type, scenario, rule_title, core_content, compliance_action FROM rules_cases WHERE 1=1"
  params = []

  # 模块分类筛选
  if category == "法条专区":
    query += " AND category LIKE '%法条%'"
  elif category == "案例专区":
    query += " AND category LIKE '%案例%'"

  # 法域筛选
  if jurisdiction != "全部":
    query += " AND jurisdiction = ?"
    params.append(jurisdiction)

  # 模糊搜索关键词匹配
  if keyword:
    query += (
        " AND (core_content LIKE ? OR rule_title LIKE ? OR data_type LIKE ? OR"
        " scenario LIKE ?)"
    )
    like_pattern = f"%{keyword}%"
    params.extend([like_pattern, like_pattern, like_pattern, like_pattern])

  df = pd.read_sql(query, conn, params=params)
  conn.close()
  return df


# 执行查询并展示结果
result_df = search_database(
    search_keyword, category_filter, jurisdiction_filter
)

st.divider()
st.subheader(
    f"📋 检索结果 (当前模块：【{category_filter}】 | 共找到 {len(result_df)}"
    " 条相关记录)"
)

if not result_df.empty:
  for index, row in result_df.iterrows():
    with st.container():
      col1, col2 = st.columns([1, 4])
      with col1:
        st.markdown(f"**模块:** `{row['category']}`")
        st.markdown(f"**法域:** `{row['jurisdiction']}`")
        st.markdown(f"**数据类型:** {row['data_type']}")
        st.markdown(f"**场景:** {row['scenario']}")
      with col2:
        st.markdown(f"### 📌 {row['rule_title']}")
        st.info(f"**核心摘要（模糊匹配列）:** {row['core_content']}")
        st.success(f"**合规应对动作:** {row['compliance_action']}")
      st.divider()
else:
  st.warning("没有找到符合条件的记录，请尝试更换关键词或切换检索模块。")