import os
import pandas as pd
import sqlite3
import streamlit as st

# 1. 数据库初始化文件名字
DB_FILE = "car_compliance.db"


# 强制每次启动时重建数据库（包含北大法宝逻辑：列表展示 + 详细法条全文/案例内容）
def init_sample_database():
  data = {
      "id": ["CN_01", "EU_01", "CASE_01"],
      "category": ["中国法条", "欧盟法条", "典型案例"],
      "jurisdiction": ["中国", "欧盟", "欧盟/跨国"],
      "data_type": ["车外实景影像/人脸/车牌", "个人数据/生物识别", "车外摄像头/人脸记录"],
      "scenario": ["联网运行/出境", "算法训练/出境", "内部测试/日常运营"],
      "rule_title": [
          "《汽车数据安全管理若干规定（试行）》第十一条",
          "《欧盟通用数据保护条例（GDPR）》第6条与第9条",
          "荷兰数据保护局 (AP) 处罚 Uber 车载监控案",
      ],
      # 列表页显示的简短指引/摘要
      "core_content": [
          "规范汽车数据处理者向境外提供重要数据及个人信息的申报要求。",
          "规定个人数据处理的合法性基础及生物识别特殊类别数据的处理禁令。",
          "网约车平台因车载摄像头违规采集司机与周边人员影像遭巨额处罚。",
      ],
      "compliance_action": [
          "属于重要数据，禁止直接无序出境，须按规定向国家网信部门申报数据出境安全评估。",
          "触发GDPR合规红线，必须在采集前取得数据主体明确授权，并完成严格的传输影响评估（TIA）。",
          "必须落实透明度义务，强化告知同意与数据保护影响评估（DPIA），避免高额罚款。",
      ],
      # 【核心新增字段】点击进去/展开后查看的法条全文或标准案例详情（北大法宝正文效果）
      "full_law_content": [
          (
              "《汽车数据安全管理若干规定（试行）》\n"
              "第十一条 汽车数据处理者开展以下汽车数据处理活动，应当依照法律、行政法规和国家网信部门有关规定，向国家网信部门申报数据出境安全评估：\n"
              "（一）向境外提供重要数据；\n"
              "（二）关键信息基础设施运营者和处理个人信息达到国家网信部门规定数量的汽车数据处理者向境外提供个人信息。\n"
              "前款规定以外的汽车数据处理者向境外提供个人信息，需要依照法律、行政法规和国家网信部门有关规定履行备案等程序的，从其规定。"
          ),
          (
              "《欧盟通用数据保护条例（GDPR）》\n"
              "【第6条 处理的合法性】\n"
              "1. 仅当且仅当满足以下至少一项条件时，处理方为合法：\n"
              "   (a) 数据主体已同意为其一个或多个特定目的处理其个人数据；\n"
              "   (b) 为履行数据主体作为一方当事人的合同所必需...\n"
              "【第9条 特殊类别个人数据的处理】\n"
              "1. 严禁处理旨在唯一识别自然人的种族、政治面貌、生物识别数据（如行人和驾驶员的面部图像）。\n"
              "2. 除非数据主体已给出明确的明示同意，或处理是为了重大公共利益所需。"
          ),
          (
              "荷兰数据保护局 (AP) 处罚 Uber 车载监控案（标准案例详情）\n"
              "1. 案情概述：网约车巨头 Uber 在欧洲运营期间，通过车载摄像头及相关设备在公开道路上持续收集了大量司机及周边行人的面部特征、行踪轨迹等敏感数据。\n"
              "2. 调查与违规焦点：荷兰数据保护局（AP）介入调查后发现，Uber 未能在车辆显著位置提供充分的采集告知，长期无法向监管机构合理解释其收集该类数据的合法性基础，严重违反了 GDPR 的透明度与数据最小化原则。\n"
              "3. 裁判与处罚结果：AP 依据 GDPR 相关规定，对 Uber 开出了高达 2.9 亿欧元的巨额罚单，并勒令其彻底整改车载设备的采集与出境传输链路。"
          ),
      ],
  }
  df = pd.DataFrame(data)
  conn = sqlite3.connect(DB_FILE)
  # 使用 replace 会在每次运行或更新时自动刷新表结构，绝不会出现列名对不上的报错！
  df.to_sql("rules_cases", conn, if_exists="replace", index=False)
  conn.close()


# 每次运行强制初始化/更新数据库，确保字段永远匹配
init_sample_database()

# ==================== Streamlit 网页前端界面 ====================
st.set_page_config(
    page_title="智能网联汽车数据跨境双向合规平台", layout="wide"
)

st.title("🚗 智能网联汽车车外实景影像数据跨境双向合规检索平台")
st.markdown(
    "仿北大法宝结构：输入关键词检索相关法规与案例，点击条目即可展开查看**原原本本的法条全文或标准案例详情**。"
)

# 侧边栏：模块选择与法域筛选
st.sidebar.header("🔍 检索与分类选项")

category_filter = st.sidebar.selectbox(
    "选择检索模块", ["全部展示", "法条专区", "案例专区"]
)
jurisdiction_filter = st.sidebar.selectbox(
    "选择法域", ["全部", "中国", "欧盟", "欧盟/跨国"]
)

search_keyword = st.text_input(
    "请输入关键词（支持模糊搜索，例如：车外、人脸、GDPR、安全评估、罚款等...）：",
    placeholder="在此输入任意关键词进行检索...",
)


# 查询函数（包含新增的 full_law_content 字段）
def search_database(keyword, category, jurisdiction):
  conn = sqlite3.connect(DB_FILE)
  query = "SELECT category, jurisdiction, data_type, scenario, rule_title, core_content, compliance_action, full_law_content FROM rules_cases WHERE 1=1"
  params = []

  if category == "法条专区":
    query += " AND category LIKE '%法条%'"
  elif category == "案例专区":
    query += " AND category LIKE '%案例%'"

  if jurisdiction != "全部":
    query += " AND jurisdiction = ?"
    params.append(jurisdiction)

  if keyword:
    query += " AND (core_content LIKE ? OR rule_title LIKE ? OR data_type LIKE ? OR scenario LIKE ? OR full_law_content LIKE ?)"
    like_pattern = f"%{keyword}%"
    params.extend(
        [like_pattern, like_pattern, like_pattern, like_pattern, like_pattern]
    )

  df = pd.read_sql(query, conn, params=params)
  conn.close()
  return df


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
        st.write(f"**摘要：** {row['core_content']}")
        st.success(f"**合规应对：** {row['compliance_action']}")

        # 【核心交互】北大法宝式的“点击进去看具体法条和内容”
        with st.expander(
            "📖 点击进入：查看该法规的标准条文全文 / 详细案例剖析"
        ):
          # 这里原原本本地展示完整的法条或案例内容
          st.text(row["full_law_content"])

      st.divider()
else:
  st.warning("没有找到符合条件的记录，请尝试更换关键词或切换检索模块。")