import os
import pandas as pd
import sqlite3
import streamlit as st

# 1. 数据库初始化文件名字
DB_FILE = "car_compliance.db"


# 自动生成测试数据库（包含真实存在的中国法规、欧盟GDPR条款及真实合规案例）
def init_sample_database():
  if not os.path.exists(DB_FILE):
    data = {
        "id": ["CN_01", "EU_01", "CASE_01"],
        "category": ["中国法条", "欧盟法条", "典型案例"],
        "jurisdiction": ["中国", "欧盟", "欧盟/跨国"],
        "data_type": ["车外实景影像/人脸/车牌", "个人数据/生物识别", "车外摄像头/人脸记录"],
        "scenario": ["联网运行/出境", "算法训练/出境", "内部测试/日常运营"],
        "rule_title": [
            "《汽车数据安全管理规定（试行）》第十一条",
            "GDPR 第 6 条（处理的合法性）与 第 9 条",
            "荷兰数据保护局 (AP) 处罚 Uber 车载监控案",
        ],
        "core_content": [
            "汽车数据处理者开展以下汽车数据处理活动，应当依照法律、行政法规和国家网信部门有关规定，向国家网信部门申报数据出境安全评估：向境外提供重要数据...",
            "个人数据的处理必须具备合法性基础（如明确同意、合同履行等）；涉及面部等特殊类别（生物识别）数据时原则上禁止处理，除非符合严格豁免条件。",
            "网约车平台因长期在车内/车外测试中未充分履行告知与合法性评估，且过度收集司机与周边人员影像，遭欧盟监管机构重罚。",
        ],
        "compliance_action": [
            "属于重要数据，禁止直接无序出境，须按规定向国家网信部门申报数据出境安全评估。",
            "触发GDPR合规红线，必须在采集前取得数据主体明确授权，并完成严格的传输影响评估（TIA）。",
            "必须落实透明度义务，强化告知同意与数据保护影响评估（DPIA），避免高额罚款。",
        ],
        "standard_detail": [
            (
                "【标准法条全文】\n"
                "《汽车数据安全管理规定（试行）》第十一条：\n"
                "汽车数据处理者开展以下汽车数据处理活动，应当依照法律、行政法规和国家网信部门有关规定，"
                "向国家网信部门申报数据出境安全评估：\n"
                "（一）向境外提供重要数据；\n"
                "（二）关键信息基础设施运营者和处理个人信息达到国家网信部门规定数量的汽车数据处理者向境外提供个人信息。\n"
                "【合规要点】涉及车外人脸、车牌等重要数据出境的，必须依法履行安全评估程序。"
            ),
            (
                "【标准法条全文】\n"
                "《欧盟通用数据保护条例（GDPR）》第 6 条与第 9 条：\n"
                "- 第 6 条：个人数据的处理必须满足“数据主体同意”、“为履行合同所必需”等合法性基础。\n"
                "- 第 9 条：严禁处理旨在唯一识别自然人的生物识别数据（如行人和驾驶员的面部图像），除非数据主体给予了明确的明示同意（Explicit consent）。"
            ),
            (
                "【标准案例概括】\n"
                "- 【案情背景】荷兰数据保护局（AP）对网约车巨头 Uber 进行调查，发现其在欧洲运营期间，通过车载摄像头及相关设备收集了大量司机及周边人员的面部特征、行踪等敏感数据。\n"
                "- 【违规痛点】违反了 GDPR 关于透明度、数据最小化以及生物识别数据处理的严格限制。\n"
                "- 【处罚结果】荷兰 AP 依据 GDPR 规定，对 Uber 开出了高达数千万欧元的巨额罚单。"
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
    "针对车企出海痛点，内置中国法、欧盟GDPR及典型案例。支持**纯文字字段与模糊搜索**，点击检索结果可展开查阅**标准法条全文与标准案例概括**。"
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


def search_database(keyword, category, jurisdiction):
  conn = sqlite3.connect(DB_FILE)
  query = "SELECT category, jurisdiction, data_type, scenario, rule_title, core_content, compliance_action, standard_detail FROM rules_cases WHERE 1=1"
  params = []

  if category == "法条专区":
    query += " AND category LIKE '%法条%'"
  elif category == "案例专区":
    query += " AND category LIKE '%案例%'"

  if jurisdiction != "全部":
    query += " AND jurisdiction = ?"
    params.append(jurisdiction)

  if keyword:
    query += " AND (core_content LIKE ? OR rule_title LIKE ? OR data_type LIKE ? OR scenario LIKE ? OR standard_detail LIKE ?)"
    like_pattern = f"%{keyword}%"
    params.extend(
        [like_pattern, like_pattern, like_pattern, like_pattern, like_pattern]
    )

  # 使用 pandas 读取查询结果
  result_df = pd.read_sql(query, conn, params=params)
  conn.close()
  return result_df


# 执行检索
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
        st.write(f"**核心摘要：** {row['core_content']}")
        st.success(f"**合规应对：** {row['compliance_action']}")

        # 点击展开查看标准法条全文 / 标准案例概括
        with st.expander("📖 点击查看【标准法条全文 / 详细案例概括】"):
          st.markdown(row["standard_detail"])

      st.divider()
else:
  st.warning("没有找到符合条件的记录，请尝试更换关键词或切换检索模块。")
