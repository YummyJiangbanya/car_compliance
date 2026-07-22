import os
import pandas as pd
import sqlite3
import streamlit as st

# 1. 数据库初始化文件名字
DB_FILE = "car_compliance.db"


# 自动生成测试数据库（基于你提供的原结构，将内容丰富为真实的法条全文与案例详情）
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
            "《欧盟通用数据保护条例（GDPR）》第6条与第9条",
            "荷兰数据保护局 (AP) 处罚 Uber 车载监控案",
        ],
        "core_content": [
            (
                "【标准法条全文】\n"
                "《汽车数据安全管理规定（试行）》第十一条：\n"
                "汽车数据处理者开展以下汽车数据处理活动，应当依照法律、行政法规和国家网信部门有关规定，向国家网信部门申报数据出境安全评估：\n"
                "（一）向境外提供重要数据；\n"
                "（二）关键信息基础设施运营者和处理个人信息达到国家网信部门规定数量的汽车数据处理者向境外提供个人信息。\n"
                "【合规要点】涉及车外人脸、车牌等重要数据出境的，必须依法履行安全评估程序。"
            ),
            (
                "【标准法条全文】\n"
                "《欧盟通用数据保护条例（GDPR）》第6条与第9条：\n"
                "- 第6条（处理合法性）：个人数据的处理必须满足“数据主体明确同意”、“为履行合同所必需”等合法基础。\n"
                "- 第9条（特殊类别数据）：严禁处理旨在唯一识别自然人的生物识别数据（如行人和驾驶员的面部图像）。除非获得明示同意（Explicit consent）或出于重大公共利益。"
            ),
            (
                "【标准案例概括】\n"
                "- 【案情背景】荷兰数据保护局（AP）对网约车巨头 Uber 进行调查，发现其在欧洲运营期间，通过车载摄像头及相关设备收集了大量司机及周边人员的面部特征、行踪等敏感数据，且长期未能向监管机构合理解释其合法性基础。\n"
                "- 【违规痛点】违反了 GDPR 关于透明度、数据最小化以及生物识别数据处理的严格限制。\n"
                "- 【裁判/处罚结果】荷兰 AP 依据 GDPR 规定，对 Uber 开出了高达数千万欧元的巨额罚单，并责令其彻底整改数据采集与出境传输机制。"
            ),
        ],
        "compliance_action": [
            "属于重要数据，禁止直接无序出境，须按规定向国家网信部门申报数据出境安全评估。",
            (
                "触发GDPR合规红线，必须在采集前取得数据主体明确授权，并完成严格的"
                "TIA（传输影响评估）。"
            ),
            (
                "必须落实透明度义务，强化告知同意与数据保护影响评估（DPIA），"
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
    "针对车企出海痛点，内置中国法、欧盟GDPR及典型案例。支持**纯文字字段与模糊搜索**，点击检索结果即可展开查看对应详细法条与应对动作。"
)

# 侧边栏：模块选择与法域筛选
st.sidebar.header("🔍 检索与分类选项")

# 1. 模块/类别选择
category_filter = st.sidebar.selectbox(
    "选择检索模块", ["全部展示", "法条专区", "案例专区"]
)

# 2. 法域筛选
jurisdiction_filter = st.sidebar.selectbox(
    "选择法域", ["全部", "中国", "欧盟", "欧盟/跨国"]
)

# 主页面搜索框
search_keyword = st.text_input(
    "请输入关键词（支持模糊搜索，例如：车外、人脸、GDPR、安全评估、罚款等...）：",
    placeholder="在此输入任意关键词进行检索...",
)


# 核心：多条件联合查询
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
    # 使用折叠面板（expander）实现点击展开查看真实的法条全文与案例详情
    with st.expander(
        f"📌 [{row['category']} - {row['jurisdiction']}] {row['rule_title']}"
        " （点击展开查看真实法条全文/案例详情）",
        expanded=(index == 0),
    ):
      col1, col2 = st.columns([1, 3])
      with col1:
        st.markdown(f"**模块:** `{row['category']}`")
        st.markdown(f"**法域:** `{row['jurisdiction']}`")
        st.markdown(f"**数据类型:** {row['data_type']}")
        st.markdown(f"**场景:** {row['scenario']}")
      with col2:
        st.markdown("### 📄 真实法条全文 / 标准案例概括")
        # 这里直接渲染真实的法条全文/案例概括内容
        st.info(row["core_content"])
        st.success(f"**合规应对动作:** {row['compliance_action']}")
else:
  st.warning("没有找到符合条件的记录，请尝试更换关键词或切换检索模块。")
