# -*- coding: utf-8 -*-
"""
校园志愿时长汇总助手（路径B：Streamlit + 通义千问 Tool Calling）
====================================================================
- 通过 OpenAI 兼容接口调用阿里云百炼 DashScope 的通义千问模型
- 自行设计 6 个业务工具，由模型真实发起 Tool Call（非 if 关键词判断）
  1. query_volunteer_data    按起止日期（+项目名 / 校内校外）查询志愿记录
  2. calc_total_hours        按姓名分组统计总时长、校内时长、校外时长
  3. calc_hours_by_type      把本人志愿时长拆成 校内 / 校外 / 总计 三个口径
  4. export_volunteer_excel  导出汇总 Excel（网页内置下载按钮）
  5. query_open_activities   按全国任意省 / 市 / 区县查询校内与校外活动
  6. add_volunteer_record    为已绑定的个人账号录入新的志愿活动与时长
- 数据层 geo.py：34 省 / 368 市 / 3078 区县 + 全国 1309 所高校（数据由公开数据集生成）
- 加分项：个人账号绑定 / 学校绑定 / 日志分区展示 / 表单新增记录 / 强异常容错 / 自动柱状图
运行：streamlit run app.py
"""

import html
import io
import json
import os
import time
import traceback
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st
from dotenv import load_dotenv
from openai import OpenAI

import geo              # 本地数据层：省市区 / 全国高校 / 活动模板

load_dotenv()

# ============================ 基础配置 ============================
st.set_page_config(page_title="校园志愿时长汇总助手", page_icon="⏱️", layout="wide")

MODEL_NAME = os.getenv("DASHSCOPE_MODEL", "qwen-plus")
FALLBACK_MODEL = os.getenv("DASHSCOPE_FALLBACK_MODEL", "qwen-turbo")  # 主模型持续限流时的备用模型
API_KEY = os.getenv("DASHSCOPE_API_KEY", "")
BASE_URL = os.getenv(
    "DASHSCOPE_BASE_URL",
    "https://dashscope.aliyuncs.com/compatible-mode/v1",
)

SYSTEM_PROMPT = (
    "你是校园志愿服务时长管理助手。用户会用自然语言提出查询、统计、导出、查找活动等需求，"
    "你必须优先调用提供的工具完成业务操作，禁止编造数据。"
    "调用工具前，先根据对话上下文把用户的相对时间（如'本月'、'上周'）换算成 YYYY-MM-DD 格式。"
    "工具返回 status 为 error 时，请把原因友好地转述给用户并给出修正建议；"
    "返回 status 为 empty 时，提示该条件下没有查到数据。"
    "统计前必须先调用查询工具；导出前必须先完成统计。"
    "调用统计和导出工具时不要自己回传数据数组（省略 records / summary 参数即可），"
    "工具会自动使用最近一次查询或统计的结果。"
    "查询记录时注意口径：用户说“我/我的/本人”或没有说明范围时，scope 传 me，只查当前登录账号本人的记录；"
    "用户明确说“每个人/全员/全班/所有同学/大家”时，scope 传 all。"
    "若按本人口径查询，回答时说明“这是你本人的记录”，并提示如需全员记录可以直说。"
    "志愿记录分为「校内」与「校外」两类：用户问校内、校外分别多少小时，或问总共多少小时，"
    "调用 calc_hours_by_type 获取校内 / 校外 / 总计三个口径；也可用 query_volunteer_data 的 kind 参数只看某一类。"
    "当用户询问最近有什么志愿活动、校内活动、某地校外活动时，调用 query_open_activities；"
    "该工具支持按 scope 区分校内/校外，按 province / city / district 指定全国任意省市区县"
    "（例如用户说“成都武侯区”“南京鼓楼区”，就把地名分别填进 city 与 district），按 keyword 过滤；"
    "用户没有说地名时不要自己填 province/city/district，工具会使用页面当前选择的地区。"
    "回答活动类问题时，请按日期列出要点，说明主办方、时间、时长、剩余名额与所在区县；"
    "只有工具返回的 sort_by 为 distance（说明所选地区就在校区所在城市）时，才可以说“按距校区由近到远排序”，"
    "跨市时不要提距离。"
    "并在末尾提示用户可在活动卡片下方展开查看详情、通过报名链接前往官方平台报名。"
    "当用户要登记自己新参加的志愿活动（如'帮我记一笔：社区清洁 3 小时'）时，调用 add_volunteer_record；"
    "如果用户说明了是校内还是校外活动，把 kind 传成「校内」或「校外」，没说就省略让工具自动判断；"
    "姓名与学号由绑定的个人账号自动带入，不要向用户索要姓名学号；缺少活动名称或时长时先追问。"
    "若工具返回未绑定账号或未绑定学校，请提示用户在左侧对应面板绑定。"
    "回答请简洁、使用中文。"
)

# ============================ 模拟数据 ============================
# 每条记录带「类型」：校内 / 校外——用于把志愿时长拆成校内、校外与总计三个口径
DEFAULT_DATA = [
    {"姓名": "张三", "学号": "2023001", "志愿项目": "图书馆图书整理", "类型": "校内",
     "日期": "2026-09-12", "时长(小时)": 3.0, "地点": "金海校区图书馆", "备注": "院系组织"},
    {"姓名": "张三", "学号": "2023001", "志愿项目": "迎新志愿服务", "类型": "校内",
     "日期": "2026-09-20", "时长(小时)": 5.0, "地点": "校门广场", "备注": ""},
    {"姓名": "张三", "学号": "2023001", "志愿项目": "校园运动会保障", "类型": "校内",
     "日期": "2026-10-15", "时长(小时)": 4.5, "地点": "东区田径场", "备注": ""},
    {"姓名": "张三", "学号": "2023001", "志愿项目": "地铁站客流引导", "类型": "校外",
     "日期": "2026-10-22", "时长(小时)": 3.0, "地点": "12 号线金海路站", "备注": "晚高峰"},
    {"姓名": "李四", "学号": "2023002", "志愿项目": "图书馆图书整理", "类型": "校内",
     "日期": "2026-09-13", "时长(小时)": 2.5, "地点": "金海校区图书馆", "备注": ""},
    {"姓名": "李四", "学号": "2023002", "志愿项目": "社区敬老院慰问", "类型": "校外",
     "日期": "2026-10-01", "时长(小时)": 6.0, "地点": "浦东新区金桥敬老院", "备注": "国庆场"},
    {"姓名": "王五", "学号": "2023003", "志愿项目": "迎新志愿服务", "类型": "校内",
     "日期": "2026-09-20", "时长(小时)": 4.0, "地点": "校门广场", "备注": ""},
    {"姓名": "王五", "学号": "2023003", "志愿项目": "校园运动会保障", "类型": "校内",
     "日期": "2026-10-15", "时长(小时)": 3.5, "地点": "东区田径场", "备注": ""},
    {"姓名": "王五", "学号": "2023003", "志愿项目": "实验室开放日引导", "类型": "校内",
     "日期": "2026-10-18", "时长(小时)": 2.0, "地点": "实验楼 A 座", "备注": ""},
    {"姓名": "王五", "学号": "2023003", "志愿项目": "徐汇滨江文明劝导", "类型": "校外",
     "日期": "2026-10-25", "时长(小时)": 4.0, "地点": "徐汇滨江步道", "备注": ""},
    {"姓名": "赵六", "学号": "2023004", "志愿项目": "社区敬老院慰问", "类型": "校外",
     "日期": "2026-10-01", "时长(小时)": 5.5, "地点": "浦东新区金桥敬老院", "备注": ""},
    {"姓名": "赵六", "学号": "2023004", "志愿项目": "图书馆图书整理", "类型": "校内",
     "日期": "2026-10-08", "时长(小时)": 3.0, "地点": "金海校区图书馆", "备注": ""},
]
RECORD_COLUMNS = ["姓名", "学号", "志愿项目", "类型", "日期", "时长(小时)", "地点", "备注"]
KIND_ON, KIND_OFF = "校内", "校外"

# ============================ 个人账号 ============================
# 可绑定的演示账号（与 DEFAULT_DATA 中的志愿者对应），也支持自己新建账号
DEMO_ACCOUNTS = [
    {"姓名": "张三", "学号": "2023001", "学院": "信息科学与技术学院", "专业": "计算机科学与技术 2301 班"},
    {"姓名": "李四", "学号": "2023002", "学院": "管理学院", "专业": "工商管理 2302 班"},
    {"姓名": "王五", "学号": "2023003", "学院": "外国语学院", "专业": "英语 2301 班"},
    {"姓名": "赵六", "学号": "2023004", "学院": "护理学院", "专业": "护理学 2303 班"},
]

# ============================ 学校与志愿活动数据 ============================
# 数据来源：regions.json（34 省 / 368 市 / 3078 区县）与 universities.json（全国 1308 所本科高校），
# 由 tools/build_geo_data.py 从公开数据集生成；学校对象统一由 geo.py 提供。
SCHOOLS = geo.SCHOOLS

# 报名入口（官方平台）
LINK_SANDA = geo.LINK_SANDA        # 上海杉达学院官网（校团委通知 / 志愿招募）
LINK_SH_VOL = geo.LINK_SH_VOL      # 上海志愿者网（全市志愿项目报名）
LINK_CN_VOL = geo.LINK_CN_VOL      # 中国志愿服务网（全国志愿项目报名）
LINK_SSTM = geo.LINK_SSTM          # 上海科技馆 / 天文馆志愿者
LINK_LIB = geo.LINK_LIB            # 上海图书馆志愿者
LINK_ART = geo.LINK_ART            # 中华艺术宫
LINK_BLOOD = geo.LINK_BLOOD        # 上海市血液中心


def haversine_km(lat1, lng1, lat2, lng2) -> float:
    return geo.haversine_km(lat1, lng1, lat2, lng2)


def region_activities(prov: str, city: str, area: str) -> list:
    """按所选省市区生成校外活动池（同一地区结果稳定）"""
    return geo.build_off_campus(prov, city, area)


def school_activities(school: dict) -> list:
    """按绑定学校生成校内活动池"""
    return geo.build_on_campus(school)


def tick_feed(force: bool = False):
    """模拟活动实时更新：随时间推进报名人数、刷新“更新于 N 分钟”，并定期推送新活动"""
    import random
    from datetime import datetime as _dt

    now = _dt.now()
    last = st.session_state.get("feed_tick_at")
    # 距上次更新不足 25 秒且非手动刷新，则跳过（避免频繁跳动）
    if not force and last and (now - last).total_seconds() < 25:
        return
    st.session_state.feed_tick_at = now

    for key in ("activities_on", "activities_off"):
        for a in st.session_state.get(key, []):
            a["更新时间"] = a.get("更新时间", 0) + random.randint(0, 3)
            if a["已报名"] < a["名额"] and random.random() < 0.55:
                a["已报名"] = min(a["名额"], a["已报名"] + random.randint(1, 2))

    # 每累计两次轮询推送一条新活动：从更长的模板序列里取尚未展示的那条
    st.session_state.feed_polls = st.session_state.get("feed_polls", 0) + 1
    if st.session_state.feed_polls % 2 == 0:
        prov = st.session_state.get("act_prov", "")
        city = st.session_state.get("act_city", "")
        area = st.session_state.get("act_area", "")
        shown = {a["标题"] for a in st.session_state.get("activities_off", [])}
        if prov and area:
            for a in geo.build_off_campus(prov, city, area, count=14):
                if a["标题"] not in shown:
                    a = dict(a)
                    a["更新时间"] = 1
                    st.session_state.activities_off.insert(0, a)
                    st.session_state.new_push = f"{a['标题']}（{a['日期']}）"
                    break
        else:
            shown_on = {a["标题"] for a in st.session_state.get("activities_on", [])}
            school = st.session_state.get("bound_school")
            if school:
                for a in geo.build_on_campus(school, count=12):
                    if a["标题"] not in shown_on:
                        a = dict(a)
                        a["更新时间"] = 1
                        st.session_state.activities_on.insert(0, a)
                        st.session_state.new_push = f"{a['标题']}（{a['日期']}）"
                        break

TOOL_ICON = {
    "query_volunteer_data": "🔍",
    "calc_total_hours": "🧮",
    "calc_hours_by_type": "⚖️",
    "export_volunteer_excel": "📤",
    "query_open_activities": "🏫",
    "add_volunteer_record": "📝",
}
STATUS_COLOR = {"success": "#22C55E", "error": "#EF4444", "empty": "#F59E0B"}
STATUS_LABEL = {"success": "成功", "error": "异常", "empty": "空结果"}

# 运行日志：记录每一次模型请求与异常，便于排查真实报错
LOG_FILE = Path(__file__).parent / "logs" / "app.log"


def log_event(kind: str, detail: str):
    """把关键事件追加写入本地日志文件（失败不影响主流程）"""
    try:
        LOG_FILE.parent.mkdir(exist_ok=True)
        with LOG_FILE.open("a", encoding="utf-8") as f:
            f.write(f"{datetime.now():%Y-%m-%d %H:%M:%S} [{kind}] {detail}\n")
    except Exception:
        pass


def read_recent_logs(n: int = 30) -> str:
    try:
        lines = LOG_FILE.read_text(encoding="utf-8").splitlines()
        return "\n".join(lines[-n:]) or "(暂无日志)"
    except FileNotFoundError:
        return "(暂无日志)"
    except Exception as e:
        return f"(读取日志失败: {e})"


def init_state():
    """初始化会话状态（模拟数据表、日志、聊天记录、导出文件）"""
    if "volunteer_df" not in st.session_state:
        st.session_state.volunteer_df = pd.DataFrame(DEFAULT_DATA, columns=RECORD_COLUMNS)
    else:
        st.session_state.volunteer_df = migrate_df(st.session_state.volunteer_df)
    if "tool_logs" not in st.session_state:
        st.session_state.tool_logs = []          # 工具调用日志
    if "history" not in st.session_state:
        st.session_state.history = []            # 聊天记录 [(role, text)]
    if "excel_file" not in st.session_state:
        st.session_state.excel_file = None       # (文件名, bytes)
    if "last_summary" not in st.session_state:
        st.session_state.last_summary = None     # 最近一次统计结果（用于柱状图）
    if "last_owner" not in st.session_state:
        st.session_state.last_owner = ""         # 最近一次查询的记录归属（本人 / 全员）
    if "toast" not in st.session_state:
        st.session_state.toast = None            # 一次性提示
    # —— 个人账号绑定 ——
    if "bound_account" not in st.session_state:
        st.session_state.bound_account = dict(DEMO_ACCOUNTS[0])   # 默认绑定演示账号“张三”
    if "acct_new" not in st.session_state:
        st.session_state.acct_new = False          # 是否正在新建账号
    if "show_all_rows" not in st.session_state:
        st.session_state.show_all_rows = False     # 数据表是否显示全部人
    # —— 学校绑定（默认：上海杉达学院 金海校区）——
    if "bound_school" not in st.session_state:
        st.session_state.bound_school = geo.default_school()
    school = st.session_state.get("bound_school") or geo.default_school()
    # —— 活动地区（省 / 市 / 区县三级联动）——
    if "act_prov" not in st.session_state:
        st.session_state.act_prov = school.get("省份") or (geo.provinces()[0] if geo.provinces() else "")
    if "act_city" not in st.session_state:
        st.session_state.act_city = school.get("城市") or ""
    if "act_area" not in st.session_state:
        st.session_state.act_area = school.get("区") or ""
    if "act_scope" not in st.session_state:
        st.session_state.act_scope = "全部"        # 全部 / 校内 / 校外
    if "act_kw" not in st.session_state:
        st.session_state.act_kw = ""               # 活动关键字
    # —— 活动池 ——
    if "activities_on" not in st.session_state:
        st.session_state.activities_on = school_activities(school)
    if "activities_off" not in st.session_state:
        st.session_state.activities_off = region_activities(
            st.session_state.act_prov, st.session_state.act_city, st.session_state.act_area)
    if "feed_tick_at" not in st.session_state:
        st.session_state.feed_tick_at = None     # 上次更新时间
    if "feed_polls" not in st.session_state:
        st.session_state.feed_polls = 0
    if "new_push" not in st.session_state:
        st.session_state.new_push = None         # 新推送提示
    if "show_my_records" not in st.session_state:
        st.session_state.show_my_records = False
    if "detail_kind" not in st.session_state:
        st.session_state.detail_kind = "全部"     # 明细面板口径：全部 / 校内 / 校外


def migrate_df(df):
    """兼容旧会话：补齐「类型 / 地点 / 备注」列，并按项目名推断校内或校外"""
    df = df.copy()
    if "类型" not in df.columns:
        on_words = ("图书馆", "迎新", "运动会", "校园", "实验室", "校运会", "宿舍")
        df["类型"] = df["志愿项目"].map(lambda x: KIND_ON if any(w in str(x) for w in on_words) else KIND_OFF)
    for col in ("地点", "备注"):
        if col not in df.columns:
            df[col] = ""
    df["类型"] = df["类型"].fillna(KIND_OFF).replace({"": KIND_OFF})
    for col in ("地点", "备注"):
        df[col] = df[col].fillna("")
    return df[RECORD_COLUMNS]


# ============================ 工具实现 ============================
def _parse_date(text):
    """严格校验日期格式，非法返回 None（对应容错设计）"""
    try:
        return datetime.strptime(str(text).strip(), "%Y-%m-%d").date()
    except (ValueError, TypeError):
        return None


def query_volunteer_data(start_date: str, end_date: str, project_name: str = "",
                         scope: str = "me", kind: str = "") -> dict:
    """工具1：按起止日期查询志愿记录。
    scope=me 只查当前绑定账号本人的记录（默认，与页面记录口径保持一致）；
    scope=all 查全部成员记录。project_name 为可选的项目名称筛选关键字；
    kind 可选「校内 / 校外」，用于只看某一类记录。"""
    d1, d2 = _parse_date(start_date), _parse_date(end_date)
    if d1 is None or d2 is None:
        return {"status": "error", "message": "日期格式不正确，请使用 YYYY-MM-DD 格式，例如 2026-09-01。"}
    if d2 < d1:
        return {"status": "error", "message": f"结束日期({end_date})早于起始日期({start_date})，请调整后重新查询。"}

    acc = get_bound_account()
    want_all = str(scope or "me").strip().lower() in ("all", "全员", "everyone", "team")

    df = st.session_state.volunteer_df.copy()
    if want_all:
        owner_label = "全员"
    elif acc:
        # 只查本人：切换账号后这里会自然跟着变
        df = df[df["姓名"] == acc["姓名"]]
        owner_label = f"{acc['姓名']}（本人）"
    else:
        owner_label = "全员（未绑定账号）"

    df["日期date"] = df["日期"].map(_parse_date)
    mask = (df["日期date"] >= d1) & (df["日期date"] <= d2)
    if project_name:
        mask &= df["志愿项目"].str.contains(str(project_name).strip(), na=False)
    kind_txt = str(kind or "").strip()
    kind_sel = ""
    if kind_txt in ("校内", "on", "on_campus", "校园"):
        kind_sel = KIND_ON
    elif kind_txt in ("校外", "off", "off_campus", "社会"):
        kind_sel = KIND_OFF
    if kind_sel:
        mask &= df["类型"] == kind_sel
    result = df[mask].drop(columns=["日期date"])

    if result.empty:
        tip = f"{start_date} 至 {end_date}" + (f"（项目含“{project_name}”）" if project_name else "")
        return {
            "status": "empty",
            "scope": "all" if want_all else "me",
            "owner": owner_label,
            "message": f"{owner_label}在 {tip} 时间范围内没有{kind_sel or ''}志愿记录。",
            "records": [],
        }

    split = {
        KIND_ON: float(round(result.loc[result["类型"] == KIND_ON, "时长(小时)"].sum(), 2)),
        KIND_OFF: float(round(result.loc[result["类型"] == KIND_OFF, "时长(小时)"].sum(), 2)),
    }
    return {
        "status": "success",
        "scope": "all" if want_all else "me",
        "owner": owner_label,
        "count": int(len(result)),
        "分类合计(小时)": {**split, "合计": float(round(split[KIND_ON] + split[KIND_OFF], 2))},
        "message": f"{owner_label}共查询到 {len(result)} 条志愿记录"
                   f"（校内 {split[KIND_ON]:g} 小时、校外 {split[KIND_OFF]:g} 小时）。",
        "records": result.to_dict(orient="records"),
    }


def calc_total_hours(records: list = None) -> dict:
    """工具2：按姓名分组统计累计总时长与参与次数。
    records 可省略——省略时自动使用最近一次查询到的记录，
    避免让模型把大数组原样回传（参数 JSON 易被截断导致调用失败）"""
    if not records:
        records = st.session_state.get("last_records") or []
    if not records:
        return {"status": "empty", "message": "没有可统计的志愿记录，请先调用 query_volunteer_data 查询。"}
    try:
        df = pd.DataFrame(records)
        if "姓名" not in df.columns or "时长(小时)" not in df.columns:
            return {"status": "error", "message": "传入的记录缺少“姓名”或“时长(小时)”字段，无法统计。"}
        df["时长(小时)"] = pd.to_numeric(df["时长(小时)"], errors="coerce").fillna(0)
        if "类型" not in df.columns:
            df["类型"] = KIND_ON
        grouped = (
            df.groupby("姓名", as_index=False)["时长(小时)"]
            .sum()
            .sort_values("时长(小时)", ascending=False)
        )
        counts = df.groupby("姓名").size().reset_index(name="参与次数")
        on_sum = df[df["类型"] == KIND_ON].groupby("姓名")["时长(小时)"].sum()
        off_sum = df[df["类型"] == KIND_OFF].groupby("姓名")["时长(小时)"].sum()
        grouped = grouped.merge(counts, on="姓名")
        grouped["时长(小时)"] = grouped["时长(小时)"].round(2)

        summary = []
        for _, r in grouped.iterrows():
            summary.append({
                "姓名": r["姓名"],
                "总时长(小时)": float(r["时长(小时)"]),
                "校内(小时)": float(round(on_sum.get(r["姓名"], 0.0), 2)),
                "校外(小时)": float(round(off_sum.get(r["姓名"], 0.0), 2)),
                "参与次数": int(r["参与次数"]),
            })
        on_total = float(round(df.loc[df["类型"] == KIND_ON, "时长(小时)"].sum(), 2))
        off_total = float(round(df.loc[df["类型"] == KIND_OFF, "时长(小时)"].sum(), 2))
        return {
            "status": "success",
            "message": f"已统计 {len(summary)} 位同学的累计志愿时长"
                       f"（校内合计 {on_total:g} 小时、校外合计 {off_total:g} 小时）。",
            "summary": summary,
            "总计(小时)": float(round(df["时长(小时)"].sum(), 2)),
            "校内合计(小时)": on_total,
            "校外合计(小时)": off_total,
        }
    except Exception as e:  # 任何统计异常都不允许程序崩溃
        return {"status": "error", "message": f"统计失败：{e}"}


def calc_hours_by_type(scope: str = "me", start_date: str = "", end_date: str = "") -> dict:
    """工具（新增）：把志愿时长拆成「校内 / 校外 / 总计」三个口径。
    默认统计当前账号本人全部记录；也可传入起止日期限定范围，或 scope=all 看全员合计。"""
    acc = get_bound_account()
    want_all = str(scope or "me").strip().lower() in ("all", "全员", "everyone", "team")
    if not want_all and not acc:
        return {"status": "error", "message": "尚未绑定个人账号，请先在左侧「个人账号」绑定后再统计。"}

    df = st.session_state.volunteer_df.copy()
    if not want_all:
        df = df[df["姓名"] == acc["姓名"]]
    owner_label = "全员" if want_all else f"{acc['姓名']}（本人）"

    period = ""
    if start_date or end_date:
        d1 = _parse_date(start_date) or date.min
        d2 = _parse_date(end_date) or date.max
        if start_date and _parse_date(start_date) is None:
            return {"status": "error", "message": f"起始日期「{start_date}」格式不正确，请用 YYYY-MM-DD。"}
        if end_date and _parse_date(end_date) is None:
            return {"status": "error", "message": f"结束日期「{end_date}」格式不正确，请用 YYYY-MM-DD。"}
        df["_d"] = df["日期"].map(_parse_date)
        df = df[(df["_d"] >= d1) & (df["_d"] <= d2)].drop(columns=["_d"])
        period = f"{start_date or '最早'} 至 {end_date or '今天'}"

    if df.empty:
        return {"status": "empty", "owner": owner_label,
                "message": f"{owner_label}在{period or '全部时间'}内没有志愿记录。"}

    df["时长(小时)"] = pd.to_numeric(df["时长(小时)"], errors="coerce").fillna(0)
    on_df = df[df["类型"] == KIND_ON]
    off_df = df[df["类型"] == KIND_OFF]
    on_h = float(round(on_df["时长(小时)"].sum(), 2))
    off_h = float(round(off_df["时长(小时)"].sum(), 2))
    total = float(round(on_h + off_h, 2))
    ratio = f"校内 {on_h / total * 100:.0f}% · 校外 {off_h / total * 100:.0f}%" if total else "—"

    def _projects(sub):
        if sub.empty:
            return []
        g = sub.groupby("志愿项目")["时长(小时)"].sum().sort_values(ascending=False)
        return [{"项目": k, "小时": float(round(v, 2))} for k, v in g.items()]

    return {
        "status": "success",
        "owner": owner_label,
        "期间": period or "全部时间",
        "校内(小时)": on_h,
        "校外(小时)": off_h,
        "总计(小时)": total,
        "校内条数": int(len(on_df)),
        "校外条数": int(len(off_df)),
        "占比": ratio,
        "校内项目": _projects(on_df),
        "校外项目": _projects(off_df),
        "message": (f"{owner_label}累计志愿 {total:g} 小时："
                    f"校内 {on_h:g} 小时（{len(on_df)} 条）、校外 {off_h:g} 小时（{len(off_df)} 条），"
                    f"{ratio}。"),
    }


def export_volunteer_excel(summary: list = None) -> dict:
    """工具3：把统计结果导出为 Excel，写入会话状态供页面下载按钮使用。
    summary 可省略——省略时自动使用最近一次统计结果"""
    if not summary:
        summary = st.session_state.get("last_summary") or []
    if not summary:
        return {"status": "empty", "message": "没有可导出的统计数据，请先完成查询与统计。"}
    try:
        df = pd.DataFrame(summary)
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="志愿时长汇总")
        excel_bytes = buf.getvalue()
        # 文件名带上记录归属，换账号后导出的表一眼能区分
        owner = str(st.session_state.get("last_owner") or "").split("（")[0]
        tag = f"_{owner}" if owner else ""
        filename = f"志愿时长汇总{tag}_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        st.session_state.excel_file = (filename, excel_bytes)  # 供页面下载按钮使用
        return {
            "status": "success",
            "message": f"汇总表已生成：{filename}（共 {len(df)} 行），可在页面下载区点击按钮下载。",
            "filename": filename,
        }
    except Exception as e:
        return {"status": "error", "message": f"导出失败：{e}"}


def _norm_prov(text: str) -> str:
    """把「上海 / 上海市 / 上海直辖市」这类写法统一成 regions.json 里的省名"""
    t = str(text or "").strip()
    if not t:
        return ""
    if t in geo.REGIONS:
        return t
    for p in geo.provinces():
        if p.startswith(t) or t.startswith(p[:2]):
            return p
    return t


def _norm_city(prov: str, text: str) -> str:
    t = str(text or "").strip()
    if not t:
        return ""
    cs = geo.cities(prov)
    if t in cs:
        return t
    for c in cs:
        if c.startswith(t) or t.startswith(c) or c[:2] == t[:2]:
            return c
    return t


def _norm_area(prov: str, city: str, text: str) -> str:
    t = str(text or "").strip()
    if not t:
        return ""
    ars = geo.areas(prov, city)
    if t in ars:
        return t
    for a in ars:
        if a.startswith(t) or t.startswith(a):
            return a
    return t


def _find_prov_by_city(city: str) -> str:
    """只给了城市名时，反查它属于哪个省级行政区（如「成都」→「四川省」）"""
    t = str(city or "").strip()
    if not t:
        return ""
    hits = [p for p in geo.provinces()
            if any(c == t or c.startswith(t) or t.startswith(c[:2]) for c in geo.cities(p))]
    return hits[0] if len(hits) == 1 else (hits[0] if hits else "")


def _find_by_area(area: str) -> tuple:
    """只给了区县名时，反查 (省, 市)（如「武侯区」→ (四川省, 成都市)）"""
    t = str(area or "").strip()
    if not t:
        return "", ""
    for p in geo.provinces():
        for c in geo.cities(p):
            if any(a == t or a.startswith(t) for a in geo.areas(p, c)):
                return p, c
    return "", ""


def resolve_region(province: str = "", city: str = "", district: str = "") -> dict:
    """把模型/界面传来的地区解析成 (省, 市, 区)，非法时给出可选项。

    模型经常只给「成都」「武侯区」而漏掉省份，这里做反查补全，避免多跑一轮工具。
    """
    prov = _norm_prov(province)
    raw_city, raw_area = str(city or "").strip(), str(district or "").strip()

    # 省份缺省时，先按区县反查，再按城市反查
    if not prov and raw_area:
        prov = _find_by_area(raw_area)[0]
    if not prov and raw_city:
        prov = _find_prov_by_city(raw_city)
    if not prov:
        prov = st.session_state.get("act_prov", "")
    if not prov or prov not in geo.REGIONS:
        return {"status": "error",
                "message": f"未识别省份「{province}」。可选省级行政区：{'、'.join(geo.provinces())}。"}

    city = _norm_city(prov, raw_city)
    if not city and raw_area:                       # 只给区县：反查所属地级市
        city = _find_by_area(raw_area)[1]
    if not city and prov == st.session_state.get("act_prov"):
        city = st.session_state.get("act_city", "")
    if not city:
        city = geo.cities(prov)[0]
    if city not in geo.cities(prov):
        return {"status": "error",
                "message": f"「{prov}」下没有「{raw_city}」。可选城市：{'、'.join(geo.cities(prov)[:12])}……"}

    area = _norm_area(prov, city, raw_area)
    if not area and (prov, city) == (st.session_state.get("act_prov"), st.session_state.get("act_city")):
        area = st.session_state.get("act_area", "")
    if area and area not in geo.areas(prov, city):
        return {"status": "error",
                "message": f"「{city}」下没有「{raw_area}」。可选区县：{'、'.join(geo.areas(prov, city)[:12])}……"}
    return {"status": "ok", "province": prov, "city": city, "area": area}


def query_open_activities(scope: str = "all", province: str = "", city: str = "",
                          district: str = "", keyword: str = "") -> dict:
    """工具4：查询志愿活动。
    scope=on_campus 查绑定学校的校内活动；off_campus 查指定地区的校外活动（同城按距校区距离排序，跨市按日期排序）；
    all 两者都查。province / city / district 支持全国任意省市区县，为空时用页面当前选择。"""
    school = st.session_state.get("bound_school")
    kw = str(keyword or "").strip()
    scope_raw = str(scope or "all").strip().lower()
    if scope_raw in ("on_campus", "on", "校内", "campus"):
        want_on, want_off = True, False
    elif scope_raw in ("off_campus", "off", "校外", "社会", "nearby"):
        want_on, want_off = False, True
    else:
        want_on, want_off = True, True

    if want_on and not school:
        return {"status": "error",
                "message": "尚未绑定学校，无法查询校内活动。请先在左侧「学校绑定」中选择并绑定学校。"}

    reg = resolve_region(province, city, district)
    if reg["status"] != "ok":
        return {"status": "error", "message": reg["message"]}
    prov, cty, area = reg["province"], reg["city"], reg["area"]

    def _match(act):
        if not kw:
            return True
        blob = " ".join([str(act.get("标题", "")), str(act.get("主办", "")),
                         " ".join(act.get("标签", [])), str(act.get("分类", ""))])
        return kw in blob

    def _dist(a):
        if not school or not school.get("lat") or not a.get("lat"):
            return None
        return geo.haversine_km(school["lat"], school["lng"], a["lat"], a["lng"])

    on_list, off_list = [], []
    # 所选地区是否就在校区所在城市——只有同城时「按距离排序」才有意义
    same_city_as_school = bool(school) and cty == geo.full_city(
        school.get("省份", ""), school.get("城市", ""))
    if want_on:
        for a in st.session_state.get("activities_on") or school_activities(school):
            if _match(a):
                on_list.append({
                    "标题": a["标题"], "主办": a["主办"], "地点": a["地点"],
                    "日期": a["日期"], "时间段": a["时间段"], "时长(小时)": a["时长"],
                    "名额": a["名额"], "已报名": a["已报名"],
                    "剩余名额": max(0, a["名额"] - a["已报名"]), "标签": a["标签"],
                    "类型": KIND_ON, "报名链接": a.get("报名链接", LINK_SANDA),
                })
        on_list.sort(key=lambda x: x["日期"])
    if want_off:
        # 只有当前页面选中的就是同一地区时才复用缓存池，否则按工具参数重建，
        # 否则模型问「成都武侯区」会读到页面上残留的上海活动池，被过滤成空结果。
        same_region = (prov, cty, area) == (
            st.session_state.get("act_prov"),
            st.session_state.get("act_city"),
            st.session_state.get("act_area"),
        )
        off_pool = (st.session_state.get("activities_off") if same_region else None) \
            or region_activities(prov, cty, area)
        for a in off_pool:
            if not _match(a):
                continue
            item = {
                "标题": a["标题"], "主办": a["主办"], "地址": a["地址"],
                "省份": a.get("省份", prov), "城市": a.get("城市", cty), "区域": a.get("区域", area),
                "日期": a["日期"], "时间段": a["时间段"], "时长(小时)": a["时长"],
                "名额": a["名额"], "已报名": a["已报名"],
                "剩余名额": max(0, a["名额"] - a["已报名"]),
                "标签": a["标签"], "类型": KIND_OFF,
                "报名链接": a.get("报名链接", LINK_SH_VOL),
            }
            if same_city_as_school:      # 跨城时距离只作参考，不参与排序
                d = _dist(a)
                if d is not None:
                    item["距离km"] = round(d, 1)
            off_list.append(item)
        if same_city_as_school:
            off_list.sort(key=lambda x: (x.get("距离km", 99999), x["日期"]))
        else:
            off_list.sort(key=lambda x: x["日期"])

    area_txt = geo.region_label(prov, cty, area)
    if not on_list and not off_list:
        return {"status": "empty",
                "message": f"「{area_txt}」范围内没有找到符合条件的活动（关键词：{kw or '无'}）。"
                           "可以换一个区县或换关键词再试。"}

    parts = []
    if want_on:
        parts.append(f"校内 {len(on_list)} 个活动")
    if want_off:
        seg = f"{area_txt} 校外 {len(off_list)} 个活动"
        if not area and off_list:
            seg += f"，覆盖 {len({a['区域'] for a in off_list})} 个区县"
        parts.append(seg)
    msg = "；".join(parts) + "。"
    if want_on and want_off:
        msg += "校内活动来自你绑定的学校，校外活动按所在地区推送。"
    if off_list:
        if same_city_as_school:
            msg += "校外活动按距校区距离由近到远排序，每个活动都附有详情与官方报名链接。"
        elif school:
            msg += (f"所选地区不在你校区所在城市（{school.get('城市', '')}），"
                    "已按日期排序；每个活动都附有详情与官方报名链接。")
        else:
            msg += "每个活动都附有详情与官方报名链接。"
    if not school:
        msg += "（未绑定学校，暂无法计算距离，绑定后可按距离排序。）"
    return {
        "status": "success",
        "school": ({"名称": school["名称"], "城市": school["城市"], "校区": school.get("校区", "")}
                   if school else None),
        "area": area_txt,
        "province": prov, "city": cty, "district": area,
        "sort_by": "distance" if (off_list and same_city_as_school) else "date",
        "on_campus": on_list,
        "off_campus": off_list,
        "message": msg,
    }



def get_bound_account():
    """当前绑定的个人账号，未绑定返回 None"""
    return st.session_state.get("bound_account")


def records_scope() -> str:
    """当前记录展示口径：绑定账号后默认只看本人（me），勾选“显示全部”或未绑定时为全员（all）"""
    if not get_bound_account():
        return "all"
    return "all" if st.session_state.get("show_all_rows") else "me"


def view_df():
    """按当前口径返回记录表——切换个人账号后，这里会自动跟着切到新账号"""
    df = st.session_state.volunteer_df
    acc = get_bound_account()
    if acc and not st.session_state.get("show_all_rows"):
        return df[df["姓名"] == acc["姓名"]]
    return df


def reset_record_views(keep_logs: bool = False):
    """记录口径变了（换账号 / 新写入记录）就清掉旧的查询结果、统计、导出文件，
    避免页面还显示上一个账号的数据"""
    st.session_state.last_records = None
    st.session_state.last_summary = None
    st.session_state.last_owner = ""
    st.session_state.excel_file = None
    st.session_state.last_failed_input = None
    if not keep_logs:
        st.session_state.history = []
        st.session_state.tool_logs = []


def switch_account(acc, tip: str = ""):
    """切换 / 绑定 / 退出登录的统一入口：账号与记录一起切换"""
    old = get_bound_account()
    st.session_state.bound_account = dict(acc) if acc else None
    st.session_state.show_all_rows = False         # 回到“只看本人”
    st.session_state.pop("show_all_toggle", None)  # 复选框回到默认（只看本人）
    reset_record_views()                           # 清空上一个账号的结果与对话
    if acc:
        name = acc["姓名"]
        if old and old.get("姓名") == name:
            st.session_state.toast = ("success", f"当前已是 {name} 的账号。")
        else:
            suffix = f"，{tip}" if tip else ""
            st.session_state.toast = (
                "success",
                f"已切换为 {name}（{acc['学号']}）{suffix}，志愿记录已同步切换为该账号的数据。",
            )
    else:
        st.session_state.toast = (
            "success",
            "已退出个人账号，记录视图已切回全员；重新绑定后录入与查询会跟随账号。",
        )


def add_volunteer_record(project_name: str, date: str = "", hours: float = 0.0,
                         place: str = "", note: str = "", kind: str = "") -> dict:
    """工具5：为当前绑定的个人账号录入一条新的志愿记录（只需活动名称与时长）。
    kind 用于区分「校内 / 校外」，省略时按活动名称与地点自动判断。"""
    acc = get_bound_account()
    if not acc:
        return {
            "status": "error",
            "message": "尚未绑定个人账号。请先在左侧「个人账号」中选择或新建账号并绑定，再录入志愿记录。",
        }

    proj = str(project_name or "").strip()
    if not proj:
        return {"status": "error", "message": "缺少志愿活动名称，请补充「做了什么活动」。"}

    # 日期：未提供则默认今天；提供了但格式非法则报错
    date_txt = str(date or "").strip()
    if not date_txt:
        d = datetime.now().date()
    else:
        d = _parse_date(date_txt)
        if d is None:
            return {"status": "error",
                    "message": f"日期「{date_txt}」格式不正确，请使用 YYYY-MM-DD，例如 2026-10-27。"}
    if d > datetime.now().date():
        return {"status": "error",
                "message": f"服务日期 {d.strftime('%Y-%m-%d')} 还没到，只能登记已经完成的志愿活动。"}

    try:
        h = float(hours)
    except (TypeError, ValueError):
        return {"status": "error", "message": "时长必须是数字，单位为小时，例如 3 或 2.5。"}
    if h <= 0:
        return {"status": "error", "message": "时长必须大于 0 小时。"}
    if h > 24:
        return {"status": "error", "message": f"单次志愿时长 {h:g} 小时明显超出合理范围（一般不超过 24），请核对后重填。"}
    h = round(h * 2) / 2  # 统一到 0.5 小时刻度

    date_str = d.strftime("%Y-%m-%d")
    df = st.session_state.volunteer_df
    dup = df[
        (df["姓名"] == acc["姓名"])
        & (df["志愿项目"] == proj)
        & (df["日期"] == date_str)
    ]
    if not dup.empty:
        return {
            "status": "error",
            "message": f"「{proj}」在 {date_str} 已经有一条记录了（{dup.iloc[0]['时长(小时)']:g} 小时），"
                       f"如确需修改请先重置数据或换个日期。",
        }

    kind_raw = str(kind or "").strip()
    if kind_raw in (KIND_ON, "on", "on_campus", "校园", "校内活动"):
        kind_val = KIND_ON
    elif kind_raw in (KIND_OFF, "off", "off_campus", "社会", "校外活动"):
        kind_val = KIND_OFF
    else:   # 未指定时按关键字自动判断，默认校外
        blob = proj + str(place or "")
        on_words = ("图书馆", "迎新", "运动会", "校园", "校内", "实验室", "校运会", "宿舍", "教学楼")
        kind_val = KIND_ON if any(w in blob for w in on_words) else KIND_OFF

    new_row = {
        "姓名": acc["姓名"],
        "学号": acc["学号"],
        "志愿项目": proj,
        "类型": kind_val,
        "日期": date_str,
        "时长(小时)": h,
        "地点": str(place or "").strip(),
        "备注": str(note or "").strip(),
    }
    st.session_state.volunteer_df = pd.concat(
        [df, pd.DataFrame([new_row], columns=RECORD_COLUMNS)], ignore_index=True)

    mine = st.session_state.volunteer_df[st.session_state.volunteer_df["姓名"] == acc["姓名"]]
    total = float(mine["时长(小时)"].sum())
    on_h = float(mine.loc[mine["类型"] == KIND_ON, "时长(小时)"].sum())
    off_h = float(mine.loc[mine["类型"] == KIND_OFF, "时长(小时)"].sum())
    # 数据变了，之前查到的结果与统计已过期，一并清掉，避免和新数据对不上
    st.session_state.last_records = None
    st.session_state.last_summary = None
    st.session_state.last_owner = ""
    st.session_state.excel_file = None
    extra = "；".join([x for x in [f"地点 {place.strip()}" if str(place).strip() else "",
                                   f"备注 {note.strip()}" if str(note).strip() else ""] if x])
    return {
        "status": "success",
        "record": new_row,
        "personal_total": total,
        "personal_count": int(len(mine)),
        "personal_on_campus": round(on_h, 2),
        "personal_off_campus": round(off_h, 2),
        "message": (f"已为 {acc['姓名']}（学号 {acc['学号']}）录入：{date_str} {proj} {h:g} 小时（{kind_val}）"
                    + (f"（{extra}）" if extra else "")
                    + f"。累计志愿 {total:.1f} 小时（校内 {on_h:g} + 校外 {off_h:g}），共 {len(mine)} 条记录。"),
    }


TOOL_IMPL = {
    "query_volunteer_data": query_volunteer_data,
    "calc_total_hours": calc_total_hours,
    "calc_hours_by_type": calc_hours_by_type,
    "export_volunteer_excel": export_volunteer_excel,
    "query_open_activities": query_open_activities,
    "add_volunteer_record": add_volunteer_record,
}

# 工具 Schema：交给通义千问，由模型自主决定何时调用（真实 Tool Calling）
TOOLS_SCHEMA = [
    {
        "type": "function",
        "function": {
            "name": "query_volunteer_data",
            "description": "按起止日期查询校园志愿记录；可用 project_name 按项目名筛选，用 scope 指定本人或全员，用 kind 只看校内或校外。",
            "parameters": {
                "type": "object",
                "properties": {
                    "start_date": {"type": "string", "description": "起始日期，格式 YYYY-MM-DD"},
                    "end_date": {"type": "string", "description": "结束日期，格式 YYYY-MM-DD"},
                    "project_name": {"type": "string", "description": "可选。志愿项目名称关键字，不需要时省略"},
                    "scope": {
                        "type": "string",
                        "enum": ["me", "all"],
                        "description": "查询口径：me 只查当前登录账号本人的记录（默认）；all 查全部成员。用户说“每个人/全员/全班”时传 all",
                    },
                    "kind": {
                        "type": "string",
                        "enum": ["校内", "校外"],
                        "description": "可选。只看某一类记录：校内 / 校外；不传则两类都查",
                    },
                },
                "required": ["start_date", "end_date"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "calc_total_hours",
            "description": "对最近查询到的记录按姓名分组，统计每人累计总时长、校内时长、校外时长与参与次数。调用前需先执行 query_volunteer_data；通常不传 records，工具自动使用最近一次查询结果。",
            "parameters": {
                "type": "object",
                "properties": {
                    "records": {
                        "type": "array",
                        "description": "可选。要统计的记录数组；一般省略，省略时自动使用最近一次查询结果",
                        "items": {"type": "object"},
                    }
                },
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "calc_hours_by_type",
            "description": "把志愿时长拆成校内 / 校外 / 总计三个口径并给出占比与各项目明细。用户问“我校内多少小时”“校内外分别多少”“总时长多少”时调用；不需要先查询，工具直接读当前账号数据。",
            "parameters": {
                "type": "object",
                "properties": {
                    "scope": {
                        "type": "string",
                        "enum": ["me", "all"],
                        "description": "统计口径：me 当前账号本人（默认） / all 全员合计",
                    },
                    "start_date": {"type": "string", "description": "可选。起始日期 YYYY-MM-DD，只看某段时间"},
                    "end_date": {"type": "string", "description": "可选。结束日期 YYYY-MM-DD"},
                },
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "export_volunteer_excel",
            "description": "把统计结果导出为 Excel 汇总表。调用前需先执行 calc_total_hours；通常不需要传 summary，工具会自动使用最近一次统计结果。",
            "parameters": {
                "type": "object",
                "properties": {
                    "summary": {
                        "type": "array",
                        "description": "可选。统计结果数组；一般省略，省略时自动使用最近一次统计结果",
                        "items": {"type": "object"},
                    }
                },
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "query_open_activities",
            "description": "查询志愿活动：绑定学校的校内活动，以及全国任意省市区县的校外志愿活动（含地址、区县、报名链接）。用户提到某个省份/城市/区县（如“成都武侯区有什么志愿活动”）时，把地名填进 province / city / district；只提到城市或区县也可以，省份会自动反查。",
            "parameters": {
                "type": "object",
                "properties": {
                    "scope": {
                        "type": "string",
                        "enum": ["all", "on_campus", "off_campus"],
                        "description": "查询范围：all 全部 / on_campus 仅校内 / off_campus 仅校外",
                    },
                    "province": {
                        "type": "string",
                        "description": "可选。省级行政区，如 上海市、四川省、广东省；不传则用页面当前选择的地区",
                    },
                    "city": {
                        "type": "string",
                        "description": "可选。城市，如 成都市、南京市；不传则用页面当前选择的城市",
                    },
                    "district": {
                        "type": "string",
                        "description": "可选。区县，如 武侯区、浦东新区；不传则用页面当前选择的区县",
                    },
                    "keyword": {
                        "type": "string",
                        "description": "可选。按标题、主办方、标签或分类过滤，例如 科普、助老、环保、支教",
                    },
                },
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "add_volunteer_record",
            "description": "为当前已绑定的个人账号新增一条志愿记录，只需要活动名称与时长。用户说“帮我记一笔/登记/录入”自己参加的志愿活动时调用。姓名与学号由账号自动带入，不要向用户索要。",
            "parameters": {
                "type": "object",
                "properties": {
                    "project_name": {
                        "type": "string",
                        "description": "志愿活动名称，必填，例如 社区垃圾分类督导",
                    },
                    "date": {
                        "type": "string",
                        "description": "服务日期 YYYY-MM-DD，可选；用户说“今天/昨天”请先换算成具体日期，省略则默认今天",
                    },
                    "hours": {
                        "type": "number",
                        "description": "本次志愿时长（小时），必填，例如 3 或 2.5",
                    },
                    "place": {
                        "type": "string",
                        "description": "可选。服务地点，例如 浦东新区金桥镇社区服务中心",
                    },
                    "kind": {
                        "type": "string",
                        "enum": ["校内", "校外"],
                        "description": "可选。这条记录属于校内还是校外志愿服务；用户说“在学校里/校内的活动”传校内，社区、场馆、地铁等社会活动传校外",
                    },
                    "note": {
                        "type": "string",
                        "description": "可选。备注，例如 含往返交通、由院系统一组织",
                    },
                },
                "required": ["project_name", "hours"],
            },
        },
    },
]



def add_log(tool_name: str, args: dict, result: dict):
    """记录一条工具调用日志（日志分区展示）"""
    st.session_state.tool_logs.append(
        {
            "time": datetime.now().strftime("%H:%M:%S"),
            "tool": tool_name,
            "args": args,
            "status": result.get("status", "unknown"),
            "brief": result.get("message", ""),
        }
    )


def execute_tool(name: str, args: dict) -> dict:
    """统一工具调度入口：执行工具并写日志"""
    fn = TOOL_IMPL.get(name)
    if fn is None:
        result = {"status": "error", "message": f"未知工具：{name}"}
    else:
        try:
            result = fn(**args)
        except TypeError as e:
            result = {"status": "error", "message": f"工具参数不匹配：{e}"}
        except Exception as e:
            result = {"status": "error", "message": f"工具执行异常：{e}"}
    add_log(name, args, result)
    return result


# ============================ Agent 主流程 ============================
def friendly_model_error(e) -> str:
    """把底层异常翻译成用户能看懂、可行动的提示（强化容错）"""
    s = str(e)
    if "429" in s or "limit_requests" in s or "Too Many Requests" in s:
        return (
            "⚠️ 接口限流（429）：当前业务空间的请求频率/配额较低，已自动重试仍失败。\n"
            "请稍等半分钟再试，或到阿里云百炼控制台调高该空间的限流/配额配置。"
        )
    if "401" in s or "InvalidApiKey" in s or "invalid_api_key" in s:
        return "⚠️ API Key 无效（401）：请检查 .env 中的 DASHSCOPE_API_KEY 是否正确、是否已失效。"
    if "403" in s or "AccessDenied" in s:
        return (
            f"⚠️ 模型无访问权限（403）：当前业务空间未开通 {MODEL_NAME}。\n"
            "请到百炼控制台开通模型授权，或把 .env 的 DASHSCOPE_MODEL 换成已开通的模型后重试。"
        )
    if "400" in s and ("InvalidParameter" in s or "invalid_parameter" in s):
        return (
            "⚠️ 工具参数格式异常（400）：模型生成的调用参数不合法，本次已自动拦截。\n"
            "请重新提问一次，或把需求说得更简洁一些。"
        )
    if "timed out" in s.lower() or "timeout" in s.lower():
        return "⚠️ 请求超时：网络波动或模型响应较慢，请稍后重试。"
    if "Connection" in s or "connection" in s.lower() or "getaddrinfo" in s:
        return "⚠️ 网络连接失败：请检查本机网络（或代理设置）后重试。"
    return f"⚠️ 调用模型失败：{e}"


def _is_retryable(s: str) -> bool:
    """判断异常是否值得自动重试（限流/超时/网关错误）"""
    return (
        "429" in s or "limit_requests" in s or "Too Many Requests" in s
        or "timed out" in s.lower() or "timeout" in s.lower()
        or "502" in s or "503" in s
    )


def create_completion_with_retry(client, status=None, **kwargs):
    """带指数退避的自动重试：应对业务空间限流(429)与瞬时超时。
    主模型重试 4 次（2/4/8/12 秒）仍失败时，自动切换备用模型再试一次。"""
    delays = [2, 4, 8, 12]
    last_err = None
    for attempt in range(len(delays) + 1):
        try:
            resp = client.chat.completions.create(**kwargs)
            log_event("MODEL_OK", f"{kwargs.get('model')} 第{attempt + 1}次尝试成功")
            return resp
        except Exception as e:
            last_err = e
            log_event(
                "MODEL_ERR",
                f"{kwargs.get('model')} 第{attempt + 1}次尝试失败 | {type(e).__name__} | {str(e)[:400]}",
            )
            if _is_retryable(str(e)) and attempt < len(delays):
                if status is not None:
                    status.update(
                        label=f"⏳ 接口限流/超时，{delays[attempt]} 秒后自动重试"
                        f"（第 {attempt + 1}/4 次）...", expanded=True, state="running"
                    )
                time.sleep(delays[attempt])
                if status is not None:
                    status.update(label="🤖 正在重新请求模型...", expanded=True, state="running")
                continue
            break
    # 主模型多次受限 → 切换备用模型做最后尝试
    if kwargs.get("model") != FALLBACK_MODEL and last_err is not None and _is_retryable(str(last_err)):
        if status is not None:
            status.update(
                label=f"🔁 {kwargs.get('model')} 持续受限，切换备用模型 {FALLBACK_MODEL}...",
                expanded=True, state="running",
            )
        try:
            return client.chat.completions.create(**{**kwargs, "model": FALLBACK_MODEL})
        except Exception:
            raise last_err
    raise last_err


def run_agent(user_input: str, status=None):
    """带工具调用循环的 Agent：模型自主决定调用哪个工具，最多循环 6 轮防失控。
    status: st.status 容器，用于实时展示"思考 → 逐个工具调用 → 完成"的处理过程。"""
    client = OpenAI(api_key=API_KEY, base_url=BASE_URL, timeout=60.0, max_retries=2)
    st.session_state.history.append(("user", user_input))
    st.session_state.last_failed_input = None  # 本次处理成功与否在结束时确定

    # 账号上下文：让模型知道"我"是谁，录入时不必再问姓名学号
    sys_content = SYSTEM_PROMPT
    today = datetime.now()
    sys_content += (
        f"\n今天是 {today.strftime('%Y-%m-%d')}（{today.strftime('%A')}）。"
        "用户说“今天/昨天/前天/本周/上周/本月”等相对时间时，必须基于这一天换算成具体日期，不要凭记忆推断年份。"
    )
    acc = st.session_state.get("bound_account")
    if acc:
        sys_content += (
            f"\n当前登录的个人账号：{acc['姓名']}（学号 {acc['学号']}"
            + (f"，{acc['学院']}" if acc.get("学院") else "")
            + "）。用户说“我”时即指这位同学，录入记录时工具会自动带入该账号。"
        )
    else:
        sys_content += "\n当前尚未绑定个人账号，录入类请求需先提示用户去左侧绑定。"

    messages = [{"role": "system", "content": sys_content}]
    for role, text in st.session_state.history:
        messages.append({"role": role, "content": text})

    try:
        for _ in range(6):
            resp = create_completion_with_retry(
                client, status=status,
                model=MODEL_NAME, messages=messages,
                tools=TOOLS_SCHEMA, temperature=0.3,
            )
            if not resp.choices:
                raise RuntimeError("模型返回为空，请重试。")
            msg = resp.choices[0].message

            if msg.tool_calls:  # 模型发起了真实 Tool Call
                if status is not None:
                    status.update(label="🛠️ 正在执行工具...", expanded=True, state="running")
                # 先规范化参数：不合法的 JSON 不再原样回传（否则下一轮请求会 400）
                normalized_calls, parsed_calls = [], []
                for tc in msg.tool_calls:
                    raw = tc.function.arguments or ""
                    try:
                        args = json.loads(raw) if raw.strip() else {}
                        ok = isinstance(args, dict)
                    except Exception:
                        args, ok = {}, False
                    normalized_calls.append(
                        {
                            "id": tc.id,
                            "type": "function",
                            "function": {
                                "name": tc.function.name,
                                "arguments": json.dumps(args, ensure_ascii=False) if ok else "{}",
                            },
                        }
                    )
                    parsed_calls.append((tc.function.name, args, ok, tc.id, raw))

                messages.append(
                    {
                        "role": "assistant",
                        "content": msg.content or "",
                        "tool_calls": normalized_calls,
                    }
                )
                for name, args, ok, call_id, raw in parsed_calls:
                    if status is not None:
                        status.update(
                            label=f"🛠️ 执行工具：{name}", expanded=True, state="running"
                        )
                    if not ok:
                        result = {
                            "status": "error",
                            "message": "模型生成的工具参数不是合法 JSON（通常是传入了过大的数据），本次调用已跳过，请省略 records/summary 参数重试。",
                        }
                        log_event("TOOL_ARGS_INVALID", f"{name} | {raw[:200]}")
                        add_log(name, {"raw": raw[:200]}, result)
                    else:
                        result = execute_tool(name, args)
                    # 查询/统计成功后，同步刷新页面的可视化
                    if name == "query_volunteer_data" and result.get("status") == "success":
                        st.session_state.last_records = result["records"]
                        st.session_state.last_owner = result.get("owner", "")
                    if name == "calc_total_hours" and result.get("status") == "success":
                        st.session_state.last_summary = result["summary"]
                    if status is not None:
                        status.write(
                            f"{TOOL_ICON.get(name, '🔧')} `{name}` →"
                            f" {STATUS_LABEL.get(result.get('status', 'unknown'), result.get('status'))}：{result.get('message', '')}"
                        )
                    messages.append(
                        {
                            "role": "tool",
                            "tool_call_id": call_id,
                            "content": json.dumps(result, ensure_ascii=False),
                        }
                    )
                continue

            reply = msg.content or "（模型未返回内容）"
            st.session_state.history.append(("assistant", reply))
            if status is not None:
                status.update(label="✅ 处理完成", state="complete")
            return
        # 连续 6 轮都在调工具仍未给出回答
        st.session_state.history.append(
            ("assistant", "⚠️ 工具调用轮次超出上限，任务未完成，请把需求拆分得更具体一些再试。")
        )
        st.session_state.last_failed_input = user_input
        if status is not None:
            status.update(label="⚠️ 轮次超限", state="error")
    except Exception as e:
        st.session_state.history.append(("assistant", friendly_model_error(e)))
        st.session_state.last_failed_input = user_input  # 记录失败请求，供一键重试
        log_event("AGENT_FAIL", f"{type(e).__name__}: {e}\n{traceback.format_exc()[-800:]}")
        if status is not None:
            status.update(label="❌ 调用失败，可在对话区点击重试", state="error")


# ============================ 全局样式 ============================
GLOBAL_CSS = """
<style>
  /* ---------- 全局 ---------- */
  .stApp { background: #F5F7FA; font-family: "Microsoft YaHei", "PingFang SC", sans-serif; }
  #MainMenu, footer { visibility: hidden; }
  h1, h2, h3 { font-family: "Microsoft YaHei", "PingFang SC", sans-serif !important; }
  section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #1E3A5F 0%, #2B4C7E 100%);
    border-right: none;
  }
  section[data-testid="stSidebar"] * { color: #E6EEF8 !important; }
  section[data-testid="stSidebar"] .stButton > button {
    background: rgba(255,255,255,.12); color: #E6EEF8 !important;
    border: 1px solid rgba(255,255,255,.25); border-radius: 10px; width: 100%;
  }
  section[data-testid="stSidebar"] .stButton > button:hover {
    background: rgba(255,255,255,.22); border-color: rgba(255,255,255,.5);
  }
  section[data-testid="stSidebar"] input, section[data-testid="stSidebar"] .stDateInput input {
    background: #FFFFFF; color: #111827 !important;
    -webkit-text-fill-color: #111827 !important;
    border: 1px solid #CBD5E0; border-radius: 8px;
  }
  section[data-testid="stSidebar"] label { color: #A8C3E8 !important; font-size: 13px; }
  section[data-testid="stSidebar"] hr { border-color: rgba(255,255,255,.15); }
  section[data-testid="stSidebar"] [data-testid="stForm"] {
    background: rgba(255,255,255,.06); border: 1px solid rgba(255,255,255,.12);
    border-radius: 14px; padding: 14px;
  }

  /* ---------- Hero 横幅 ---------- */
  .hero {
    background: linear-gradient(135deg, #1E3A5F 0%, #2B6CB0 55%, #4299E1 100%);
    border-radius: 18px; padding: 26px 32px; margin-bottom: 18px;
    color: #fff; position: relative; overflow: hidden;
    box-shadow: 0 8px 24px rgba(43,108,176,.25);
  }
  .hero::after {
    content: "⏱"; position: absolute; right: 26px; top: 50%; transform: translateY(-50%);
    font-size: 92px; opacity: .16;
  }
  .hero-title { font-size: 30px; font-weight: 800; letter-spacing: 1px; }
  .hero-sub { font-size: 14px; color: #CFE3F7; margin-top: 6px; }
  .badge {
    display: inline-block; background: rgba(255,255,255,.16); border: 1px solid rgba(255,255,255,.35);
    border-radius: 999px; padding: 3px 12px; font-size: 12px; margin: 12px 6px 0 0; color: #fff;
  }

  /* ---------- 指标卡片 ---------- */
  .stat-card {
    background: #fff; border: 1px solid #E6ECF3; border-radius: 14px;
    padding: 16px 18px; box-shadow: 0 2px 8px rgba(45,55,72,.05);
    height: 100%;
  }
  .stat-top { display: flex; align-items: center; gap: 8px; }
  .stat-icon {
    width: 34px; height: 34px; border-radius: 10px; display: inline-flex;
    align-items: center; justify-content: center; font-size: 17px;
  }
  .stat-label { font-size: 13px; color: #718096; }
  .stat-value { font-size: 26px; font-weight: 800; color: #1E3A5F; margin-top: 8px; }
  .stat-unit { font-size: 12px; color: #A0AEC0; font-weight: 400; }
  .stat-hint { font-size: 11px; color: #2B6CB0; margin-top: 8px; font-weight: 600; }
  /* 总时长卡：最醒目的一张 —— 深蓝渐变 + 更大字号 + 校内/校外占比条 */
  .stat-card.hero {
    background: linear-gradient(135deg, #16305A 0%, #2563EB 62%, #3B82F6 100%);
    border: none; color: #fff;
    box-shadow: 0 12px 28px rgba(37,99,235,.30);
    position: relative; overflow: hidden;
  }
  .stat-card.hero::after {
    content: ""; position: absolute; right: -2.2rem; top: -2.2rem;
    width: 8rem; height: 8rem; border-radius: 50%; background: rgba(255,255,255,.08);
  }
  .stat-card.hero .stat-label { color: #D7E4FF; font-weight: 700; letter-spacing: .02em; }
  .stat-card.hero .stat-icon { background: rgba(255,255,255,.18); }
  .stat-card.hero .stat-value { color: #fff; font-size: 42px; line-height: 1.05; letter-spacing: -.02em; }
  .stat-card.hero .stat-unit { color: #C7DAFF; font-size: 15px; font-weight: 600; }
  .stat-card.hero .stat-hint { color: #C7DAFF; font-weight: 600; }
  .hero-bar { display: flex; height: 8px; border-radius: 99px; overflow: hidden; margin-top: 10px; background: rgba(255,255,255,.22); }
  .hero-bar i { display: block; height: 100%; }
  .hero-bar .on { background: #7DD3FC; }
  .hero-bar .off { background: #FBBF24; }
  .hero-legend { display: flex; gap: 12px; font-size: 11px; color: #D7E4FF; margin-top: 6px; font-weight: 600; }
  .hero-legend b { color: #fff; }
  .kind-dot { display: inline-block; width: 7px; height: 7px; border-radius: 50%; margin-right: 4px; vertical-align: middle; }
  .rec-card { cursor: pointer; transition: box-shadow .2s ease, transform .2s ease; }
  .rec-card:hover { box-shadow: 0 6px 18px rgba(43,108,176,.18); transform: translateY(-2px); }
  /* 明细面板：总计 / 校内 / 校外三段小卡 */
  .split-summary { display: flex; gap: 10px; flex-wrap: wrap; margin: 4px 0 8px; }
  .split-summary > div {
    flex: 1; min-width: 130px; background: #fff; border: 1px solid #E6ECF3;
    border-radius: 10px; padding: 8px 12px; font-size: 12px; color: #4A5568;
  }
  .split-summary b { font-size: 18px; color: #1E3A5F; margin: 0 3px; }
  .split-summary .on { border-left: 4px solid #0EA5E9; }
  .split-summary .off { border-left: 4px solid #F59E0B; }
  /* 活动中心筛选条 */
  .filter-bar {
    background: linear-gradient(180deg,#F7FAFF,#EEF4FF); border: 1px solid #DDE8FA;
    border-radius: 12px; padding: 10px 14px 2px; margin: 2px 0 6px;
  }
  .crumb {
    display: flex; align-items: center; gap: 8px; flex-wrap: wrap;
    background: #fff; border: 1px solid #E3ECFA; border-radius: 10px;
    padding: 8px 12px; font-size: 12.5px; color: #334155; margin-bottom: 8px;
  }
  .crumb b { color: #1D4ED8; }
  .crumb .res { margin-left: auto; font-weight: 700; color: #0F766E; }
  .crumb .chip-scope {
    padding: 2px 8px; border-radius: 99px; font-size: 11.5px; font-weight: 700;
  }
  .crumb .chip-scope.on { background: #E0F2FE; color: #0369A1; }
  .crumb .chip-scope.off { background: #FEF3C7; color: #B45309; }
  .crumb .chip-scope.all { background: #EDE9FE; color: #6D28D9; }
  .sec-head {
    display: flex; align-items: baseline; gap: 8px; font-size: 13.5px; font-weight: 800;
    margin: 10px 0 8px; padding-left: 9px; border-left: 4px solid #38A169; color: #1A202C;
  }
  .sec-head.off { border-left-color: #DD6B20; }
  .sec-head span { font-size: 11.5px; font-weight: 600; color: #718096; }
  .kind-chip { font-size: 10px; font-weight: 800; border-radius: 4px; padding: 1px 6px; }
  .kind-chip.on { background: #E0F2FE; color: #0369A1; }
  .kind-chip.off { background: #FEF3C7; color: #B45309; }
  /* 「志愿记录总数」整卡可点：透明按钮覆盖在卡片上层，浏览器需支持 :has */
  div[data-testid="stHorizontalBlock"]:has(.rec-card) > div[data-testid="stColumn"]:first-child {
    position: relative;
  }
  /* 中间层不能成为定位包含块，否则按钮会跑到卡片下方 */
  div[data-testid="stHorizontalBlock"]:has(.rec-card) > div[data-testid="stColumn"]:first-child div[data-testid="stVerticalBlock"],
  div[data-testid="stHorizontalBlock"]:has(.rec-card) > div[data-testid="stColumn"]:first-child div[data-testid="stElementContainer"] {
    position: static;
  }
  div[data-testid="stHorizontalBlock"]:has(.rec-card) > div[data-testid="stColumn"]:first-child div[data-testid="stButton"] {
    position: absolute; top: 0; left: 0; width: 100%; height: 100%; z-index: 5;
  }
  div[data-testid="stHorizontalBlock"]:has(.rec-card) > div[data-testid="stColumn"]:first-child div[data-testid="stButton"] > div {
    width: 100%; height: 100%;
  }
  div[data-testid="stHorizontalBlock"]:has(.rec-card) > div[data-testid="stColumn"]:first-child div[data-testid="stButton"] button {
    width: 100%; height: 100%; opacity: 0; background: transparent; border: none; cursor: pointer;
  }

  /* ---------- 面板 ---------- */
  .panel-title { display: flex; align-items: center; gap: 8px; margin-bottom: 10px; }
  .panel-title .dot { width: 8px; height: 8px; border-radius: 50%; }
  .panel-title span.txt { font-size: 16px; font-weight: 700; color: #2D3748; }
  .panel-title span.hint { font-size: 12px; color: #A0AEC0; font-weight: 400; margin-left: 4px; }
  .empty-box {
    border: 1.5px dashed #CBD5E0; border-radius: 12px; padding: 22px 12px;
    text-align: center; color: #A0AEC0; font-size: 13px; background: #FAFBFC;
  }
  .empty-box .big { font-size: 26px; display: block; margin-bottom: 6px; }

  /* ---------- 日志卡片 ---------- */
  .log-card {
    border-radius: 12px; padding: 10px 14px; margin-bottom: 8px;
    background: #fff; border: 1px solid #E6ECF3;
    box-shadow: 0 1px 4px rgba(45,55,72,.05); font-size: 13px;
  }
  .log-head { display: flex; align-items: center; gap: 8px; }
  .log-tool { font-weight: 700; color: #1E3A5F; font-size: 13.5px; }
  .log-pill {
    font-size: 11px; font-weight: 700; border-radius: 999px;
    padding: 1px 9px; color: #fff;
  }
  .log-time { color: #A0AEC0; font-size: 11.5px; margin-left: auto; }
  .log-line { color: #4A5568; margin-top: 5px; word-break: break-all; }
  .log-line b { color: #2D3748; }

  /* ---------- 对话气泡 ---------- */
  [data-testid="stChatMessage"] {
    border-radius: 14px; border: 1px solid #E6ECF3; background: #fff;
    border-left: 4px solid #4299E1;
    box-shadow: 0 2px 10px rgba(43,108,176,.08);
    padding-top: 12px; padding-bottom: 12px; margin-bottom: 14px;
  }
  [data-testid="stChatMessage"] img { border-radius: 10px; }
  [data-testid="stChatMessage"] p,
  [data-testid="stChatMessage"] li { font-size: 15.3px; line-height: 1.78; color: #1A202C; }
  [data-testid="stChatMessage"] strong { color: #1E3A5F; }

  /* 回答标签 */
  .ans-tag {
    display: inline-block; background: #EBF8FF; color: #2B6CB0;
    border-radius: 999px; padding: 3px 13px; font-size: 12px; font-weight: 800;
    letter-spacing: .5px; margin-bottom: 4px;
  }

  /* ---------- 用户提问气泡：醒目高亮 ---------- */
  .user-bubble {
    background: linear-gradient(135deg, #2B6CB0 0%, #4299E1 100%);
    color: #FFFFFF; border-radius: 14px; padding: 14px 18px; margin: 10px 0 14px;
    font-size: 16.5px; font-weight: 700; line-height: 1.55;
    box-shadow: 0 4px 14px rgba(43, 108, 176, .30);
  }
  .user-bubble .qtag {
    display: inline-block; background: rgba(255,255,255,.22);
    border-radius: 999px; padding: 2px 12px; font-size: 12px; font-weight: 700;
    margin-bottom: 7px; letter-spacing: 1px;
  }

  /* ---------- 错误回复：红色高亮卡片 ---------- */
  .err-bubble {
    background: #FFF5F5; border: 1px solid #FED7D7; border-left: 5px solid #EF4444;
    color: #742A2A; border-radius: 12px; padding: 12px 16px; margin: 8px 0 14px;
    font-size: 14.5px; font-weight: 600; line-height: 1.6;
  }

  /* ---------- 输入控件统一：白底 + 黑字（任何主题下都看得清） ---------- */
  [data-testid="stChatInput"] {
    background: #FFFFFF !important;
    border: 1.5px solid #CBD5E0 !important;
    border-radius: 14px !important;
    box-shadow: 0 2px 10px rgba(45,55,72,.06);
  }
  [data-testid="stChatInput"] > div,
  [data-testid="stChatInput"] [data-baseweb="textarea"],
  [data-testid="stChatInput"] [data-baseweb="base-input"],
  [data-testid="stChatInput"] textarea {
    background: #FFFFFF !important;
    border-radius: 12px !important;
  }
  [data-testid="stChatInput"] textarea {
    color: #111827 !important;
    -webkit-text-fill-color: #111827 !important;
    font-size: 15.5px !important;
    font-weight: 600 !important;
    caret-color: #2B6CB0 !important;
  }
  [data-testid="stChatInput"] textarea::placeholder {
    color: #9CA3AF !important;
    -webkit-text-fill-color: #9CA3AF !important;
    font-weight: 400 !important;
  }
  [data-testid="stChatInput"] button { background: #FFFFFF !important; }
  [data-testid="stChatInput"] button svg { fill: #2B6CB0 !important; color: #2B6CB0 !important; }

  /* 所有文本 / 数字 / 日期输入框：白底黑字，不留任何歧义 */
  input, textarea, [data-baseweb="input"], [data-baseweb="input"] > div, [data-baseweb="textarea"] {
    background-color: #FFFFFF !important;
    color: #111827 !important;
    -webkit-text-fill-color: #111827 !important;
    caret-color: #2B6CB0 !important;
  }
  input::placeholder, textarea::placeholder {
    color: #9CA3AF !important;
    -webkit-text-fill-color: #9CA3AF !important;
  }
  section[data-testid="stSidebar"] input,
  section[data-testid="stSidebar"] textarea,
  section[data-testid="stSidebar"] [data-baseweb="input"] > div {
    background-color: #FFFFFF !important;
    color: #111827 !important;
    -webkit-text-fill-color: #111827 !important;
    font-weight: 600 !important;
  }
  section[data-testid="stSidebar"] input::placeholder {
    color: #9CA3AF !important;
    -webkit-text-fill-color: #9CA3AF !important;
  }
  /* 侧栏日期/数字框内的图标按钮：白底上改深色，避免看不见 */
  section[data-testid="stSidebar"] [data-testid="stNumberInput"] button,
  section[data-testid="stSidebar"] [data-testid="stNumberInput"] button svg,
  section[data-testid="stSidebar"] [data-testid="stDateInput"] svg,
  section[data-testid="stSidebar"] [data-testid="stDateInput"] button {
    color: #111827 !important;
    fill: #111827 !important;
  }

  /* ---------- 按钮统一 ---------- */
  .stButton > button {
    border-radius: 10px; border: 1px solid #CBD5E0; background: #fff;
    color: #2B6CB0; font-weight: 600; font-size: 13px;
    box-shadow: 0 1px 3px rgba(45,55,72,.06);
  }
  .stButton > button:hover { border-color: #4299E1; color: #2B6CB0; background: #EBF8FF; }
  div[data-testid="stDownloadButton"] > button {
    background: linear-gradient(135deg, #2B6CB0, #4299E1) !important;
    color: #fff !important; border: none; border-radius: 10px; font-weight: 700;
    box-shadow: 0 4px 12px rgba(66,153,225,.35); width: 100%;
  }
  [data-testid="stFormSubmitButton"] > button {
    background: linear-gradient(135deg, #2B6CB0, #4299E1);
    color: #fff; border: none; border-radius: 10px; font-weight: 700; width: 100%;
  }
  [data-testid="stDataFrame"] { border-radius: 12px; overflow: hidden; border: 1px solid #E6ECF3; }

  /* ---------- 快捷指令 chips ---------- */
  .chip {
    display: inline-block; background: #fff; border: 1px solid #CBD5E0;
    border-radius: 999px; padding: 5px 14px; font-size: 12.5px; color: #2B6CB0;
    margin-right: 6px; margin-bottom: 4px;
  }

  /* ---------- 学校绑定状态条 ---------- */
  .school-strip {
    display: flex; align-items: center; justify-content: space-between; gap: 16px;
    background: linear-gradient(120deg, #0F2B46 0%, #1B4E6E 60%, #16697A 100%);
    border-radius: 14px; padding: 16px 22px; margin: 6px 0 16px; color: #fff;
    box-shadow: 0 6px 20px rgba(15, 43, 70, .22);
  }
  .school-strip.unbound {
    background: #FFF8E6; border: 1.5px dashed #E2B93B; box-shadow: none;
  }
  .school-strip.unbound .ss-label { color: #B7791F; }
  .school-strip.unbound .ss-name { color: #7B5804; }
  .school-strip.unbound .ss-sub { color: #A17A1E; }
  .ss-left { display: flex; flex-direction: column; gap: 3px; }
  .ss-label { font-size: 11px; letter-spacing: 2px; color: #9FC6E0; font-weight: 700; }
  .ss-name { font-size: 19px; font-weight: 800; letter-spacing: .4px; }
  .ss-sub { font-size: 12px; color: #B9D4E6; }
  .ss-right { display: flex; gap: 26px; }
  .ss-right > div { text-align: center; }
  .ss-right b {
    display: block; font-size: 19px; font-weight: 800; color: #fff; font-family: Consolas, monospace;
  }
  .ss-right span { font-size: 11px; color: #A9C8DC; }

  .school-chip {
    background: rgba(255,255,255,.14); border: 1px solid rgba(255,255,255,.28);
    border-radius: 10px; padding: 7px 10px; margin-top: 6px;
    font-size: 12.5px; font-weight: 700; color: #fff;
  }
  .school-chip span { display: block; font-size: 11px; font-weight: 400; color: #A8C3E8; margin-top: 2px; }

  /* ---------- 个人账号 ---------- */
  .acct-strip {
    display: flex; align-items: center; justify-content: space-between; gap: 16px;
    background: linear-gradient(120deg, #FFFFFF 0%, #EFF6FF 100%);
    border: 1px solid #DCE7F5; border-left: 4px solid #2B6CB0;
    border-radius: 14px; padding: 12px 22px; margin: 4px 0 10px;
    box-shadow: 0 3px 12px rgba(43,108,176,.09);
  }
  .acct-strip.unbound {
    background: #FFFDF5; border: 1.5px dashed #E2B93B; box-shadow: none;
  }
  .acct-strip.unbound .as-label { color: #B7791F; }
  .acct-strip.unbound .as-name { color: #7B5804; font-size: 15px; }
  .acct-strip.unbound .as-sub { color: #A17A1E; }
  .as-left { display: flex; align-items: center; gap: 12px; }
  .as-avatar {
    width: 42px; height: 42px; flex: 0 0 42px; border-radius: 50%;
    background: linear-gradient(135deg, #2B6CB0 0%, #63B3ED 100%);
    color: #fff; font-size: 18px; font-weight: 800;
    display: flex; align-items: center; justify-content: center;
    box-shadow: 0 2px 8px rgba(43,108,176,.28);
  }
  .acct-strip.unbound .as-avatar {
    background: #F5E6C0; color: #8A6D1F; box-shadow: none;
  }
  .as-txt { display: flex; flex-direction: column; gap: 2px; }
  .as-label { font-size: 10.5px; letter-spacing: 2px; font-weight: 700; color: #5B7FA6; }
  .as-name { font-size: 18px; font-weight: 800; color: #17324D; letter-spacing: .3px; }
  .as-name em {
    font-style: normal; font-size: 11.5px; font-weight: 600; color: #2B6CB0;
    background: #E4EFFC; border-radius: 99px; padding: 1px 9px; margin-left: 8px;
    vertical-align: middle;
  }
  .as-sub { font-size: 11.5px; color: #6B7F94; }
  .as-right { display: flex; gap: 26px; }
  .as-right > div { text-align: center; }
  .as-right b {
    display: block; font-size: 18px; font-weight: 800; color: #1F5C99;
    font-family: Consolas, "Courier New", monospace;
  }
  .as-right span { font-size: 11px; color: #7C8FA3; }

  .acct-chip {
    display: flex; gap: 10px; align-items: flex-start;
    background: rgba(255,255,255,.12); border: 1px solid rgba(255,255,255,.26);
    border-radius: 12px; padding: 9px 11px; margin-top: 6px;
  }
  .ac-avatar {
    width: 32px; height: 32px; flex: 0 0 32px; border-radius: 50%;
    background: linear-gradient(135deg, #4299E1, #63B3ED); color: #fff;
    font-size: 14px; font-weight: 800;
    display: flex; align-items: center; justify-content: center;
  }
  .ac-name { font-size: 13px; font-weight: 800; color: #fff; }
  .ac-id { font-size: 10px; font-weight: 500; color: #A8C3E8; margin-left: 6px; }
  .ac-sub { font-size: 10.5px; color: #9CB6D4; margin-top: 2px; line-height: 1.35; }
  .ac-kpi { font-size: 11px; color: #C6DBEF; margin-top: 3px; }
  .ac-kpi b { color: #fff; font-family: Consolas, monospace; }

  /* ---------- 活动卡片 ---------- */
  .act-card {
    background: #fff; border: 1px solid #E6ECF3; border-left: 4px solid #38A169;
    border-radius: 12px; padding: 12px 16px; margin-bottom: 10px;
    box-shadow: 0 2px 8px rgba(45,55,72,.05);
    transition: transform .18s ease, box-shadow .18s ease;
  }
  .act-card:hover { transform: translateY(-2px); box-shadow: 0 8px 18px rgba(45,55,72,.10); }
  .act-card.off { border-left-color: #DD6B20; }
  .act-head { display: flex; align-items: center; gap: 8px; }
  .act-title { font-size: 14.5px; font-weight: 800; color: #1A202C; }
  .act-hours {
    margin-left: auto; font-family: Consolas, monospace; font-size: 13px; font-weight: 800;
    color: #2B6CB0; background: #EBF8FF; border-radius: 6px; padding: 1px 8px;
  }
  .badge-new {
    font-size: 10px; font-weight: 800; letter-spacing: .6px; color: #fff;
    background: linear-gradient(135deg, #E53E3E, #F56565); border-radius: 4px; padding: 1px 6px;
  }
  .badge-full {
    font-size: 10px; font-weight: 800; color: #718096; background: #EDF2F7;
    border-radius: 4px; padding: 1px 6px;
  }
  .act-meta { font-size: 12px; color: #718096; margin-top: 4px; }
  .act-grid {
    display: flex; flex-wrap: wrap; gap: 6px 16px; margin-top: 7px;
    font-size: 12px; color: #4A5568;
  }
  .act-dist { color: #C05621; font-weight: 700; }
  .act-fresh { color: #A0AEC0; }
  .act-bar { height: 5px; background: #EDF2F7; border-radius: 99px; margin-top: 9px; overflow: hidden; }
  .act-bar i { display: block; height: 100%; background: linear-gradient(90deg, #4299E1, #38B2AC); border-radius: 99px; }
  .act-tags { margin-top: 8px; }
  .tagx {
    display: inline-block; font-size: 11px; color: #2C7A7B; background: #E6FFFA;
    border-radius: 99px; padding: 1px 9px; margin-right: 5px;
  }
  .act-region {
    display: inline-block; font-size: 10.5px; font-weight: 800; color: #B7791F;
    background: #FEFCBF; border-radius: 4px; padding: 1px 6px;
  }
  .act-hint { font-size: 11.5px; color: #A0AEC0; margin-top: 6px; }
  .det-box { background: #F7FAFC; border: 1px solid #E6ECF3; border-radius: 10px; padding: 10px 14px; }
  .det-row { display: flex; gap: 10px; padding: 5px 0; border-bottom: 1px dashed #E6ECF3; font-size: 12.5px; }
  .det-row:last-child { border-bottom: none; }
  .det-k {
    flex: none; width: 62px; font-weight: 800; color: #2B6CB0;
    letter-spacing: .5px; font-size: 12px;
  }
  .det-v { color: #4A5568; line-height: 1.6; }
</style>
"""


def render_header():
    st.markdown(GLOBAL_CSS, unsafe_allow_html=True)
    st.markdown(
        f"""
        <div class="hero">
          <div class="hero-title">⏱️ 校园志愿时长汇总助手</div>
          <div class="hero-sub">绑定个人账号，一句话完成：登记志愿 · 查询记录 · 统计时长 · 导出 Excel · 发现活动</div>
          <div>
            <span class="badge">🤖 {MODEL_NAME}</span>
            <span class="badge">🛠️ 6 个业务工具</span>
            <span class="badge">👤 个人账号绑定</span>
            <span class="badge">🧾 工具调用日志</span>
            <span class="badge">📊 自动可视化</span>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def my_records_df():
    """当前绑定账号本人的全部志愿记录（按日期由近到远；未绑定时为空表）"""
    df = st.session_state.volunteer_df
    acc = get_bound_account()
    if acc:
        return df[df["姓名"] == acc["姓名"]].sort_values("日期", ascending=False)
    return df.iloc[0:0]


def render_stats():
    """顶部 KPI：① 醒目的累计总时长 → ② 校内 → ③ 校外 → ④ 记录总数（整卡可点开明细）。
    四张卡都只跟随当前绑定账号本人的记录，录入或切换账号后同一轮渲染即更新。"""
    mine = my_records_df()
    acc = get_bound_account()
    n = len(mine)
    total = on_h = off_h = 0.0
    on_n = off_n = 0
    if n:
        hours = pd.to_numeric(mine["时长(小时)"], errors="coerce").fillna(0)
        total = float(hours.sum())
        on_mask = mine["类型"] == KIND_ON
        on_h = float(hours[on_mask].sum())
        off_h = float(hours[~on_mask].sum())
        on_n = int(on_mask.sum())
        off_n = int((~on_mask).sum())
    on_pct = (on_h / total * 100) if total else 0.0
    off_pct = 100 - on_pct if total else 0.0

    # 顺序：① 累计总时长（最醒目，占最大宽度）→ ② 校内 → ③ 校外 → ④ 记录总数
    c1, c2, c3, c4 = st.columns([1.28, 1.0, 1.0, 1.0], gap="small")

    with c1:
        st.markdown(
            f"""
            <div class="stat-card hero">
              <div class="stat-top">
                <span class="stat-icon">⏱</span>
                <span class="stat-label">{"我的累计志愿总时长" if acc else "累计志愿总时长"}</span>
              </div>
              <div class="stat-value">{total:.1f} <span class="stat-unit">小时</span></div>
              <div class="hero-bar">
                <i class="on" style="width:{on_pct:.1f}%"></i><i class="off" style="width:{off_pct:.1f}%"></i>
              </div>
              <div class="hero-legend">
                <span><span class="kind-dot" style="background:#7DD3FC"></span>校内 <b>{on_h:.1f}h</b></span>
                <span><span class="kind-dot" style="background:#FBBF24"></span>校外 <b>{off_h:.1f}h</b></span>
              </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with c2:
        st.markdown(
            f"""
            <div class="stat-card">
              <div class="stat-top">
                <span class="stat-icon" style="background:#E6F6FF;">🏫</span>
                <span class="stat-label">校内志愿时长</span>
              </div>
              <div class="stat-value">{on_h:.1f} <span class="stat-unit">小时</span></div>
              <div class="stat-hint" style="color:#0369A1;">{on_n} 条 · 占 {on_pct:.0f}%</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with c3:
        st.markdown(
            f"""
            <div class="stat-card">
              <div class="stat-top">
                <span class="stat-icon" style="background:#FFF7E6;">🌆</span>
                <span class="stat-label">校外志愿时长</span>
              </div>
              <div class="stat-value">{off_h:.1f} <span class="stat-unit">小时</span></div>
              <div class="stat-hint" style="color:#B45309;">{off_n} 条 · 占 {off_pct:.0f}%</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with c4:
        rec_label = f"{acc['姓名']}的记录" if acc else "志愿记录总数"
        rec_hint = (
            '<div class="stat-hint">👁 点本卡查看每一项记录</div>' if acc
            else '<div class="stat-hint" style="color:#A0AEC0;">绑定账号后显示本人记录</div>'
        )
        st.markdown(
            f"""
            <div class="stat-card rec-card">
              <div class="stat-top">
                <span class="stat-icon" style="background:#EBF8FF;">📋</span>
                <span class="stat-label">{rec_label}</span>
              </div>
              <div class="stat-value">{n} <span class="stat-unit">条</span></div>
              {rec_hint}
            </div>
            """,
            unsafe_allow_html=True,
        )
        # 透明热区按钮：让整张卡片可点击（老浏览器下退化为普通按钮，功能不受影响）
        clicked = st.button("🔍 查看我的志愿记录明细", key="open_my_records", use_container_width=True)

    if clicked:  # 点卡片 = 展开本人记录明细（再点一次收起）
        st.session_state.show_my_records = not st.session_state.get("show_my_records")
        st.rerun()

    if not acc:
        st.caption("⚠️ 尚未绑定个人账号：在左侧「个人账号」选择或新建并绑定后，四张卡片即显示本人数据。")
    elif st.session_state.get("show_all_rows"):
        st.caption("左侧数据表当前按「全员」口径展示；顶部四张卡片始终显示本人数据，与账号同步切换。")



def render_my_records_detail():
    """点开「志愿记录总数」卡片后的明细面板：可按全部 / 校内 / 校外切换查看"""
    acc = get_bound_account()
    if not acc:
        return
    mine = my_records_df()
    _panel_title("📋", "#2B6CB0", f"{acc['姓名']}的志愿记录明细",
                 "每一项记录 · 按日期由近到远 · 可切换校内 / 校外")

    if mine.empty:
        st.markdown(
            '<div class="empty-box"><span class="big">📝</span>'
            "你名下还没有志愿记录<br>在左侧「录入我的志愿记录」或直接对话登记第一条吧</div>",
            unsafe_allow_html=True,
        )
        if st.button("🔼 收起明细", key="close_my_records", use_container_width=True):
            st.session_state.show_my_records = False
            st.rerun()
        return

    hours = pd.to_numeric(mine["时长(小时)"], errors="coerce").fillna(0)
    on_mask = mine["类型"] == KIND_ON
    total, on_h, off_h = float(hours.sum()), float(hours[on_mask].sum()), float(hours[~on_mask].sum())

    pick = st.radio(
        "查看范围", ["全部", KIND_ON, KIND_OFF],
        index=["全部", KIND_ON, KIND_OFF].index(st.session_state.get("detail_kind", "全部")),
        horizontal=True, key="detail_kind_radio", label_visibility="collapsed",
    )
    st.session_state.detail_kind = pick
    view = mine if pick == "全部" else mine[mine["类型"] == pick]

    st.markdown(
        f"""
        <div class="split-summary">
          <div><span>总计</span><b>{total:.1f}</b> 小时 · {len(mine)} 条</div>
          <div class="on"><span>🏫 校内</span><b>{on_h:.1f}</b> 小时 · {int(on_mask.sum())} 条</div>
          <div class="off"><span>🌆 校外</span><b>{off_h:.1f}</b> 小时 · {int((~on_mask).sum())} 条</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.caption(f"当前显示：{pick} · {len(view)} 条 · 合计 "
               f"{float(pd.to_numeric(view['时长(小时)'], errors='coerce').fillna(0).sum()):.1f} 小时")
    st.dataframe(
        view[["日期", "志愿项目", "类型", "时长(小时)", "地点", "备注"]],
        use_container_width=True, hide_index=True, height=250,
    )
    if st.button("🔼 收起明细", key="close_my_records", use_container_width=True):
        st.session_state.show_my_records = False
        st.rerun()



def _panel_title(icon, dot_color, text, hint=""):
    h = f'<span class="hint">{hint}</span>' if hint else ""
    st.markdown(
        f'<div class="panel-title"><span class="dot" style="background:{dot_color};"></span>'
        f'<span class="txt">{icon} {text}</span>{h}</div>',
        unsafe_allow_html=True,
    )


def render_tool_logs():
    """日志分区展示——工具调用记录单独高亮显示，与对话内容区分"""
    _panel_title("🧾", "#F59E0B", "工具调用日志", "Tool Call 流水 · 最新在前")
    logs = st.session_state.tool_logs
    if not logs:
        st.markdown(
            '<div class="empty-box"><span class="big">🛠️</span>'
            "暂无工具调用记录<br>对助手说一句话，模型发起 Tool Call 时会在这里高亮展示</div>",
            unsafe_allow_html=True,
        )
        return
    html = []
    for log in reversed(logs):  # 最新在前
        color = STATUS_COLOR.get(log["status"], "#718096")
        label = STATUS_LABEL.get(log["status"], log["status"])
        icon = TOOL_ICON.get(log["tool"], "🔧")
        html.append(
            f"""
            <div class="log-card" style="border-left: 4px solid {color};">
              <div class="log-head">
                <span>{icon}</span>
                <span class="log-tool">{log['tool']}</span>
                <span class="log-pill" style="background:{color};">{label}</span>
                <span class="log-time">{log['time']}</span>
              </div>
              <div class="log-line"><b>入参</b>　{log['args']}</div>
              <div class="log-line"><b>结果</b>　{log['brief']}</div>
            </div>
            """
        )
    st.markdown("\n".join(html), unsafe_allow_html=True)
    st.caption(f"共 {len(logs)} 次调用 · 🟢 成功 · 🟡 空结果 · 🔴 异常")


def render_chart():
    """查询统计后自动生成柱状图"""
    _panel_title("📊", "#4299E1", "志愿时长可视化", "统计成功后自动生成")
    if not st.session_state.get("last_summary"):
        st.markdown(
            '<div class="empty-box"><span class="big">📈</span>'
            "还没有统计数据<br>对助手说“统计每个人累计志愿总时长”试试</div>",
            unsafe_allow_html=True,
        )
        return
    df = pd.DataFrame(st.session_state.last_summary)
    if "校内(小时)" in df.columns and "校外(小时)" in df.columns:
        # 统计结果带校内外拆分：画堆叠柱，一眼看出每人校内 / 校外各占多少
        long = df.melt(id_vars="姓名", value_vars=["校内(小时)", "校外(小时)"],
                       var_name="类型", value_name="小时")
        long["类型"] = long["类型"].str.replace("(小时)", "", regex=False)
        fig = px.bar(
            long, x="姓名", y="小时", color="类型", text="小时", barmode="stack",
            color_discrete_map={"校内": "#38BDF8", "校外": "#FBBF24"},
        )
        fig.update_traces(texttemplate="%{y:g}", textposition="inside")
        fig.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0, title=""))
    else:
        fig = px.bar(
            df, x="姓名", y="总时长(小时)", color="姓名", text="总时长(小时)",
            color_discrete_sequence=["#4299E1", "#48BB78", "#ED8936", "#9F7AEA", "#38B2AC", "#F6AD55"],
        )
        fig.update_traces(texttemplate="%{y}", textposition="outside")
    fig.update_layout(
        showlegend=bool("类型" in df.columns or "校内(小时)" in df.columns),
        height=330,
        margin=dict(l=10, r=10, t=24, b=10),
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Microsoft YaHei", size=13, color="#4A5568"),
        yaxis=dict(gridcolor="#EDF2F7", title=""),
        xaxis=dict(title=""),
    )
    st.plotly_chart(fig, use_container_width=True)


def render_download():
    _panel_title("📥", "#38A169", "Excel 导出下载", "导出成功后在此一键保存")
    if st.session_state.excel_file:
        fname, fbytes = st.session_state.excel_file
        st.download_button(
            label=f"⬇️ 下载 {fname}",
            data=fbytes,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    else:
        st.markdown(
            '<div class="empty-box"><span class="big">🗂️</span>'
            "还没有导出过汇总表<br>对助手说“帮我导出汇总 Excel”试试</div>",
            unsafe_allow_html=True,
        )


# ============================ 个人账号绑定 ============================
def account_stats(acc):
    """返回 (我的记录数, 我的累计时长, 校内时长, 校外时长)"""
    df = st.session_state.volunteer_df
    mine = df[df["姓名"] == acc["姓名"]]
    hours = pd.to_numeric(mine["时长(小时)"], errors="coerce").fillna(0)
    return (len(mine), float(hours.sum()),
            float(hours[mine["类型"] == KIND_ON].sum()),
            float(hours[mine["类型"] != KIND_ON].sum()))


def _do_create_account():
    """新建账号表单的回调：校验后写入 bound_account，并让选择器跳到新账号"""
    n_name = (st.session_state.get("new_acct_name") or "").strip()
    n_sid = (st.session_state.get("new_acct_sid") or "").strip()
    n_col = (st.session_state.get("new_acct_col") or "").strip()
    if not n_name or not n_sid:
        st.session_state.toast = ("error", "姓名和学号都要填写，才能绑定个人账号。")
        return
    if any(a["学号"] == n_sid and a["姓名"] != n_name for a in DEMO_ACCOUNTS):
        st.session_state.toast = ("error", f"学号 {n_sid} 已被其他账号占用，请核对。")
        return
    st.session_state.pop("acct_picker", None)   # 让选择器按 index 重新定位到新账号
    st.session_state.new_acct_name = ""
    st.session_state.new_acct_sid = ""
    st.session_state.new_acct_col = ""
    switch_account(
        {"姓名": n_name, "学号": n_sid, "学院": n_col, "专业": ""},
        tip="之后录入只需填活动名称与时长",
    )


def render_account_bind():
    """侧栏：绑定 / 切换 / 新建 / 退出个人账号"""
    st.markdown("**👤 个人账号**")
    acc = st.session_state.get("bound_account")
    names = [a["姓名"] for a in DEMO_ACCOUNTS]
    options = list(names)
    if acc and acc["姓名"] not in names:
        options.append(f"{acc['姓名']}（我的账号）")
    options.append("➕ 新建我的账号")

    cur_label = None
    if acc:
        cur_label = acc["姓名"] if acc["姓名"] in names else f"{acc['姓名']}（我的账号）"
    idx = options.index(cur_label) if cur_label in options else len(options) - 1
    pick = st.selectbox("选择账号", options, index=idx, key="acct_picker")

    if pick == "➕ 新建我的账号":
        with st.form("new_acct_form"):
            st.text_input("姓名", placeholder="例如 陈小明", key="new_acct_name")
            st.text_input("学号", placeholder="例如 2024010234", key="new_acct_sid")
            st.text_input("学院 / 专业（可选）", placeholder="例如 信息科学与技术学院", key="new_acct_col")
            st.form_submit_button("绑定这个账号", use_container_width=True,
                                  on_click=_do_create_account)
    else:
        c1, c2 = st.columns(2)
        if c1.button("绑定 / 切换", use_container_width=True, key="acct_bind_btn"):
            demo = next((a for a in DEMO_ACCOUNTS if a["姓名"] == pick), None)
            switch_account(demo or acc)
            st.rerun()
        if c2.button("退出登录", use_container_width=True, key="acct_logout_btn"):
            switch_account(None)
            st.rerun()

    acc = st.session_state.get("bound_account")
    if not acc:
        st.caption("⚠️ 未绑定账号：无法录入自己的志愿记录")
        return
    cnt, hrs, on_h, off_h = account_stats(acc)
    sub = " · ".join([x for x in [acc.get("学院", ""), acc.get("专业", "")] if x]) or "未填写院系信息"
    st.markdown(
        f"""
        <div class="acct-chip">
          <div class="ac-avatar">{acc["姓名"][:1]}</div>
          <div class="ac-main">
            <div class="ac-name">{acc["姓名"]}<span class="ac-id">学号 {acc["学号"]}</span></div>
            <div class="ac-sub">{sub}</div>
            <div class="ac-kpi">已录 <b>{cnt}</b> 条 · 总 <b>{hrs:.1f}</b>h · 校内 <b>{on_h:.1f}</b>h · 校外 <b>{off_h:.1f}</b>h</div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    scope_txt = "仅本人" if records_scope() == "me" else "全员"
    st.caption(f"记录口径：{scope_txt} · 切换账号后，数据表 / KPI / 查询结果会一起切换")


def render_account_strip():
    """顶部：个人账号状态条（未绑定时给出引导）"""
    acc = st.session_state.get("bound_account")
    if not acc:
        st.markdown(
            """
            <div class="acct-strip unbound">
              <div class="as-left">
                <span class="as-avatar">?</span>
                <div class="as-txt">
                  <span class="as-label">未绑定个人账号</span>
                  <span class="as-name">绑定后录入志愿记录不必再填姓名学号</span>
                  <span class="as-sub">入口：左侧边栏「个人账号」→ 选择账号或新建我的账号 → 绑定</span>
                </div>
              </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        return

    cnt, hrs, on_h, off_h = account_stats(acc)
    df = st.session_state.volunteer_df
    sub = " · ".join([x for x in [acc.get("学院", ""), acc.get("专业", "")] if x]) or "未填写院系信息"
    st.markdown(
        f"""
        <div class="acct-strip">
          <div class="as-left">
            <span class="as-avatar">{acc["姓名"][:1]}</span>
            <div class="as-txt">
              <span class="as-label">当前账号</span>
              <span class="as-name">{acc["姓名"]}<em>学号 {acc["学号"]}</em></span>
              <span class="as-sub">{sub}</span>
            </div>
          </div>
          <div class="as-right">
            <div><b>{cnt}</b><span>我的记录</span></div>
            <div><b>{hrs:.1f}</b><span>我的总时长</span></div>
            <div><b>{on_h:.1f}</b><span>校内时长</span></div>
            <div><b>{off_h:.1f}</b><span>校外时长</span></div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ============================ 学校绑定 / 活动中心 ============================
def render_school_bind():
    """侧栏：按「省份 → 城市 → 学校」三级选择绑定学校（收录全国 1309 所本科高校）"""
    st.markdown("**🏫 学校绑定**")
    cur = st.session_state.get("bound_school")
    provs = geo.school_provinces()

    cur_prov = (cur or {}).get("省份") if (cur or {}).get("省份") in provs else provs[0]
    prov = st.selectbox("① 省份 / 直辖市", provs, index=provs.index(cur_prov), key="sch_prov")
    if st.session_state.get("_sch_prov_prev") != prov:
        st.session_state["_sch_prov_prev"] = prov
        for k in ("sch_city", "sch_pick"):
            st.session_state.pop(k, None)

    cities = ["全部城市"] + geo.school_cities(prov)
    cur_city = (cur or {}).get("城市") if (cur or {}).get("城市") in cities else "全部城市"
    city = st.selectbox("② 城市", cities, index=cities.index(cur_city), key="sch_city")
    city_arg = "" if city == "全部城市" else city

    kw = st.text_input("③ 搜索学校（可留空）", value="", key="sch_kw",
                       placeholder="输入校名关键字，如 复旦 / 杉达 / 师范").strip()
    pool = geo.schools_in(prov, city_arg, kw)
    if not pool:
        st.caption(f"没有匹配「{kw}」的学校，换个关键字或改选「全部城市」。")
        return

    names = [s["名称"] for s in pool]
    idx = 0
    if cur and cur.get("名称") in names:
        idx = names.index(cur["名称"])
    pick = st.selectbox(f"④ 选择学校（共 {len(names)} 所）", names, index=idx, key="sch_pick")

    c1, c2 = st.columns(2)
    if c1.button("绑定 / 切换", use_container_width=True, key="bind_btn"):
        school = dict(geo.SCHOOL_BY_NAME[pick])
        st.session_state.bound_school = school
        st.session_state.activities_on = school_activities(school)
        # 活动地区跟着学校走，避免“绑了成都的学校却还在看上海的校外活动”
        st.session_state.act_prov = school.get("省份") or provs[0]
        st.session_state.act_city = school.get("城市") or ""
        st.session_state.act_area = school.get("区") or ""
        for k in ("sel_prov", "sel_city", "sel_area", "sch_pick", "sch_city"):
            st.session_state.pop(k, None)
        st.session_state.pool_key = None
        st.session_state.feed_tick_at = None
        st.session_state.toast = (
            "success",
            f"已绑定 {school['名称']}（{school.get('省份','')}·{school.get('城市','')}），"
            "活动地区已同步切换。",
        )
        st.rerun()
    if c2.button("解除绑定", use_container_width=True, key="unbind_btn"):
        st.session_state.bound_school = None
        st.session_state.activities_on = []
        st.session_state.toast = ("success", "已解除学校绑定。")
        st.rerun()

    if cur:
        tags = "".join(f" · {t}" for t in cur.get("标签", [])[:2])
        st.markdown(
            f'<div class="school-chip">✅ {cur["名称"]}'
            f'<span>{cur.get("省份","")} · {cur.get("城市","")} · {cur.get("区","")}{tags}</span>'
            f'<span>{cur.get("地址","")}</span></div>',
            unsafe_allow_html=True,
        )
        if cur.get("官网"):
            st.caption(f"学校官网：{cur['官网']}")
    else:
        st.caption("⚠️ 未绑定：校内活动不可用，校外活动仍可按地区浏览")



def render_school_strip():
    """顶部：学校 + 当前活动地区状态条（一眼看清「绑了哪所学校、在看哪个地区」）"""
    cur = st.session_state.get("bound_school")
    prov = st.session_state.get("act_prov", "")
    city = st.session_state.get("act_city", "")
    area = st.session_state.get("act_area", "")
    region_txt = geo.region_label(prov, city, area) if prov else "未选择地区"
    if not area and city:
        region_txt = f"{city}（全市）"

    on_n = len(st.session_state.get("activities_on") or [])
    off_list = st.session_state.get("activities_off") or []
    off_n = len(off_list)
    areas_n = len({a.get("区域") for a in off_list if a.get("区域")})
    dists = ([geo.haversine_km(cur["lat"], cur["lng"], a["lat"], a["lng"])
              for a in off_list if a.get("lat") and cur.get("lat")] if cur else [])
    near = f"{min(dists):.1f} km" if dists else "—"

    if not cur:
        st.markdown(
            f"""
            <div class="school-strip unbound">
              <div class="ss-left">
                <span class="ss-label">未绑定学校</span>
                <span class="ss-name">当前浏览：{region_txt} 的校外活动</span>
                <span class="ss-sub">左侧「学校绑定」→ 省份 → 城市 → 学校 → 绑定，绑定后即可看校内活动与距离</span>
              </div>
              <div class="ss-right">
                <div><b>{off_n}</b><span>校外活动</span></div>
                <div><b>{areas_n}</b><span>覆盖区县</span></div>
              </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        return

    st.markdown(
        f"""
        <div class="school-strip">
          <div class="ss-left">
            <span class="ss-label">已绑定学校</span>
            <span class="ss-name">{cur["名称"]}</span>
            <span class="ss-sub">{cur.get("省份","")} · {cur.get("城市","")} · {cur.get("区","")} · {cur.get("地址","")}</span>
          </div>
          <div class="ss-right">
            <div><b>{region_txt}</b><span>当前活动地区</span></div>
            <div><b>{on_n}</b><span>校内活动</span></div>
            <div><b>{off_n}</b><span>校外活动</span></div>
            <div><b>{near}</b><span>最近校外</span></div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )



def render_activity_detail(a: dict, kind: str, dist: float = None):
    """活动详情（展开后显示）+ 报名链接"""
    d = a.get("详情", {}) or {}
    region = a.get("区域", "")
    rows = [
        ("活动介绍", d.get("介绍", "主办方暂未提供详细介绍，可点击下方链接查看官方发布。")),
        ("招募对象", d.get("招募对象", "以主办方公告为准")),
        ("集合信息", d.get("集合", (a.get("地点") or a.get("地址", "")) )),
        ("联系人", d.get("联系人", "见报名页面")),
        ("注意事项", d.get("注意", "请遵守主办方现场安排")),
    ]
    body = "".join(
        f'<div class="det-row"><span class="det-k">{k}</span><span class="det-v">{v}</span></div>'
        for k, v in rows
    )
    meta = f'{a["日期"]}　{a["时间段"]}　{a["时长"]:g} 小时'
    if dist is not None:
        meta += f'　·　距校区 {dist:.1f} km　·　{region}'
    st.markdown(
        f'<div class="det-box">'
        f'<div class="det-row"><span class="det-k">时间地点</span>'
        f'<span class="det-v"><b>{meta}</b><br>{a.get("地址") or a.get("地点", "")}</span></div>'
        f'{body}</div>',
        unsafe_allow_html=True,
    )

    c1, c2 = st.columns([1.1, 2])
    with c1:
        st.link_button("🔗 前往官方平台报名", a.get("报名链接", LINK_SH_VOL), use_container_width=True)
    with c2:
        st.caption(
            f"报名入口：{a.get('报名链接', LINK_SH_VOL)}　"
            "（名额与时长以主办方确认为准；未满员时点击链接即可进入官方报名页）"
        )


def render_activity_card(a: dict, kind: str, dist: float = None):
    """单个活动卡片（校内 kind=on / 校外 kind=off）"""
    left = max(0, a["名额"] - a["已报名"])
    pct = int(a["已报名"] / a["名额"] * 100) if a.get("名额") else 0
    is_new = a.get("更新时间", 99) <= 2
    tags = "".join(f'<span class="tagx">{t}</span>' for t in a.get("标签", []))
    place = a.get("地点") or a.get("地址", "")
    place_label = "地点" if kind == "on" else "地址"
    dist_html = f'<span class="act-dist">📍 距校区 {dist:.1f} km</span>' if dist is not None else ""
    new_html = '<span class="badge-new">NEW</span>' if is_new else ""
    full_html = '<span class="badge-full">已满</span>' if left == 0 else ""
    region_html = (
        f'<span class="act-region">{a["区域"]}</span>' if kind == "off" and a.get("区域") else ""
    )
    kind_chip = (f'<span class="kind-chip {"on" if kind == "on" else "off"}">'
                 f'{"🏫 校内" if kind == "on" else "🌆 校外"}</span>')
    st.markdown(
        f"""
        <div class="act-card {kind}">
          <div class="act-head">
            <span class="act-title">{a["标题"]}</span>{kind_chip}{new_html}{full_html}{region_html}
            <span class="act-hours">{a["时长"]:g} h</span>
          </div>
          <div class="act-meta">{a["主办"]} · {place_label}：{place}</div>
          <div class="act-grid">
            <span>🗓 {a["日期"]}　{a["时间段"]}</span>{dist_html}
            <span>👥 已报名 {a["已报名"]}/{a["名额"]}（余 {left}）</span>
            <span class="act-fresh">更新于 {a.get("更新时间", 0)} 分钟前</span>
          </div>
          <div class="act-bar"><i style="width:{pct}%"></i></div>
          <div class="act-tags">{tags}</div>
          <div class="act-hint">👇 展开下方「活动详情与报名」可查看介绍、集合地点、联系人，并跳转官方报名入口</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    with st.expander(f"📄 活动详情与报名 · {a['标题']}", expanded=False):
        render_activity_detail(a, kind, dist)


def _ensure_region_pool():
    """地区变化后重建校外活动池（只在地区真的变了时才重建）"""
    key = (st.session_state.act_prov, st.session_state.act_city, st.session_state.act_area)
    if st.session_state.get("pool_key") != key or not st.session_state.get("activities_off"):
        st.session_state.activities_off = region_activities(*key)
        st.session_state.pool_key = key
        st.session_state.new_push = None


def render_activities():
    """活动中心：全国省 / 市 / 区县三级筛选 + 校内校外一键切换"""
    _panel_title("🏫", "#0D9488", "志愿活动中心",
                 "全国省市 + 区县筛选 · 校内 / 校外分开展示 · 详情可展开 · 一键报名")
    school = st.session_state.get("bound_school")
    provs = geo.provinces()

    # ---------------- 筛选条：省份 → 城市 → 区县 ----------------
    f1, f2, f3, f4 = st.columns([1.15, 1.15, 1.15, 0.9])
    cur_prov = st.session_state.act_prov if st.session_state.act_prov in provs else provs[0]
    prov = f1.selectbox("省份 / 直辖市 / 自治区", provs,
                        index=provs.index(cur_prov), key="sel_prov")
    if prov != st.session_state.act_prov:
        st.session_state.act_prov = prov
        st.session_state.act_city = ""
        st.session_state.act_area = ""
        for k in ("sel_city", "sel_area"):
            st.session_state.pop(k, None)
        st.rerun()

    city_list = geo.cities(prov)
    cur_city = st.session_state.act_city if st.session_state.act_city in city_list else city_list[0]
    city = f2.selectbox("城市", city_list, index=city_list.index(cur_city), key="sel_city")
    if city != st.session_state.act_city:
        st.session_state.act_city = city
        st.session_state.act_area = ""
        st.session_state.pop("sel_area", None)
        st.rerun()

    area_opts = ["全部区县"] + geo.areas(prov, city)
    cur_area = st.session_state.act_area if st.session_state.act_area in area_opts else "全部区县"
    area_pick = f3.selectbox("区县", area_opts, index=area_opts.index(cur_area), key="sel_area")
    area = "" if area_pick == "全部区县" else area_pick
    if area != st.session_state.act_area:
        st.session_state.act_area = area
        st.rerun()
    if f4.button("🔄 刷新活动", use_container_width=True, key="refresh_feed"):
        tick_feed(force=True)
        st.session_state.toast = ("success", "活动列表已更新。")
        st.rerun()

    _ensure_region_pool()

    # ---------------- 第二行：范围切换 + 关键字 ----------------
    g1, g2, g3 = st.columns([1.9, 2.1, 0.8])
    scope = g1.radio("活动范围", ["全部", "校内", "校外"],
                     index=["全部", "校内", "校外"].index(st.session_state.act_scope),
                     horizontal=True, key="sel_scope")
    st.session_state.act_scope = scope
    kw = g2.text_input("关键字", value=st.session_state.act_kw, key="act_kw_input",
                       placeholder="按标题 / 主办 / 标签过滤，如 科普 · 助老 · 环保 · 支教").strip()
    st.session_state.act_kw = kw
    if g3.button("↩️ 回到本校", use_container_width=True, key="reset_region"):
        if school:
            st.session_state.act_prov = school.get("省份") or provs[0]
            st.session_state.act_city = school.get("城市") or ""
            st.session_state.act_area = school.get("区") or ""
        st.session_state.act_scope = "全部"
        st.session_state.act_kw = ""
        for k in ("sel_prov", "sel_city", "sel_area", "sel_scope", "act_kw_input"):
            st.session_state.pop(k, None)
        st.session_state.pool_key = None
        st.rerun()

    want_on = scope in ("全部", "校内")
    want_off = scope in ("全部", "校外")

    def _hit(a):
        if not kw:
            return True
        return kw in (a["标题"] + a["主办"] + "".join(a.get("标签", [])) + a.get("分类", ""))

    on_items = [a for a in (st.session_state.activities_on if want_on else []) if _hit(a)]
    on_items.sort(key=lambda x: x["日期"])
    # 只有所选地区与校区同城时才按距离排序；跨市距离没有参考价值，改按日期
    school_city = geo.full_city(school.get("省份", ""), school.get("城市", "")) if school else ""
    same_city = bool(school_city) and school_city == city
    off_rows = []
    if want_off:
        for a in st.session_state.activities_off:
            if not _hit(a):
                continue
            d = (geo.haversine_km(school["lat"], school["lng"], a["lat"], a["lng"])
                 if same_city and school.get("lat") and a.get("lat") else None)
            off_rows.append((d, a))
        off_rows.sort(key=lambda x: (x[0] if x[0] is not None else 99999, x[1]["日期"]))

    total_n = len(on_items) + len(off_rows)
    scope_cls = {"全部": "all", "校内": "on", "校外": "off"}[scope]
    scope_ico = {"全部": "📋", "校内": "🏫", "校外": "🌆"}[scope]
    area_txt = geo.region_label(prov, city, area) if area else f"{city}（全市）"
    kw_txt = f" · 关键字「{kw}」" if kw else ""
    st.markdown(
        f"""
        <div class="crumb">
          <span class="chip-scope {scope_cls}">{scope_ico} {scope}活动</span>
          <span>📍 <b>{area_txt}</b></span>
          <span>{kw_txt or '· 不限关键字'}</span>
          <span class="res">共 {total_n} 个 · 校内 {len(on_items)} / 校外 {len(off_rows)}</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    last = st.session_state.get("feed_tick_at")
    st.caption(
        "最近更新：" + (last.strftime("%H:%M:%S") if last else "—")
        + ("　·　校内外分开统计，校外按距校区由近到远" if same_city
           else "　·　校内外分开统计，跨市活动按日期排列")
        + "　·　活动为模拟数据，报名以主办方公告为准"
    )
    if st.session_state.get("new_push"):
        st.info(f"📬 新活动推送：{st.session_state.new_push}")
        st.session_state.new_push = None

    # ---------------- 校内 ----------------
    if want_on:
        st.markdown(
            f'<div class="sec-head on">🏫 校内活动 <span>{len(on_items)} 个'
            + (f" · 来自 {school['名称']}" if school else "")
            + "</span></div>",
            unsafe_allow_html=True,
        )
        if not school:
            st.markdown(
                '<div class="empty-box"><span class="big">🏫</span>'
                "尚未绑定学校，看不到校内活动<br>请在左侧「学校绑定」按省份选择学校并绑定</div>",
                unsafe_allow_html=True,
            )
        elif not on_items:
            st.markdown(
                '<div class="empty-box"><span class="big">🔍</span>'
                f"校内没有匹配「{kw}」的活动<br>清空关键字或切到「校外」看看</div>",
                unsafe_allow_html=True,
            )
        for a in on_items:
            render_activity_card(a, "on")

    # ---------------- 校外 ----------------
    if want_off:
        if same_city:
            head_extra = f" · 距 {school['名称'].split('（')[0]} 由近到远"
        elif school:
            head_extra = f" · 不在校区所在城市，已按日期排序"
        else:
            head_extra = " · 绑定学校后可显示距离"
        st.markdown(
            f'<div class="sec-head off">🌆 校外活动 <span>{len(off_rows)} 个{head_extra}</span></div>',
            unsafe_allow_html=True,
        )
        if not off_rows:
            st.markdown(
                '<div class="empty-box"><span class="big">🌆</span>'
                f"{area_txt} 没有匹配的活动<br>可换一个区县或清空关键字</div>",
                unsafe_allow_html=True,
            )
        else:
            near, far = off_rows[0][0], next((d for d, _ in reversed(off_rows) if d is not None), None)
            if near is not None and far is not None:
                st.caption(f"距校区 {near:.1f} ~ {far:.1f} km　·　共 {len(off_rows)} 个活动")
            elif school:
                st.caption(f"{area_txt} 不在 {school_city or school.get('城市', '')}，"
                           f"跨市活动不计算距离，按日期排列　·　共 {len(off_rows)} 个活动")
        for d, a in off_rows:
            render_activity_card(a, "off", d)



def render_sidebar():
    with st.sidebar:
        st.markdown(
            """
            <div style="text-align:center; padding: 4px 0 10px;">
              <div style="font-size: 30px;">🫡</div>
              <div style="font-size:17px; font-weight:800;">志愿者数据管理台</div>
              <div style="font-size:12px; color:#A8C3E8; margin-top:2px;">Volunteer Admin Console</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        # ---- 个人账号（置顶） ----
        render_account_bind()
        st.divider()

        # ---- 学校绑定 ----
        render_school_bind()
        st.divider()

        # ---- 数据表 ----
        acc = st.session_state.get("bound_account")
        show_all = st.toggle(
            "显示全部成员的记录",
            value=st.session_state.show_all_rows,
            key="show_all_toggle",
        )
        st.session_state.show_all_rows = show_all
        df_view = view_df()          # 切换账号后，这里自动变成新账号的记录
        if acc and not show_all:
            st.markdown(f"**📋 我的志愿记录**（{len(df_view)} 条）")
        else:
            st.markdown(f"**📋 全部志愿记录**（{len(df_view)} 条）")
        st.dataframe(df_view, use_container_width=True, height=210)

        # ---- 录入我的志愿记录（姓名/学号由账号带入） ----
        st.markdown("**➕ 录入我的志愿记录**")
        if not acc:
            st.caption("⚠️ 请先在上方「个人账号」绑定账号，绑定后这里只需填活动名称与时长。")
        else:
            st.caption(f"将记入：{acc['姓名']} · {acc['学号']}")
            with st.form("add_form", clear_on_submit=True):
                kind_pick = st.radio(
                    "志愿类型", [KIND_OFF, KIND_ON], horizontal=True,
                    help="校内：学校组织的岗位（图书馆、迎新、校运会等）；校外：社区、场馆、地铁等社会服务",
                )
                project = st.text_input("志愿活动名称", placeholder="例如 社区垃圾分类督导")
                d = st.date_input("服务日期", value=datetime.now().date())
                hours = st.number_input(
                    "志愿时长(小时)", min_value=0.5, max_value=24.0, step=0.5, value=3.0
                )
                place = st.text_input("服务地点（可选）", placeholder="例如 浦东新区金桥镇")
                submitted = st.form_submit_button("✔️ 记入我的志愿时长", use_container_width=True)
            if submitted:
                res = add_volunteer_record(
                    project_name=project,
                    date=d.strftime("%Y-%m-%d"),
                    hours=hours,
                    place=place,
                    kind=kind_pick,
                )
                add_log("add_volunteer_record",
                        {"project_name": project, "date": d.strftime("%Y-%m-%d"),
                         "hours": hours, "kind": kind_pick},
                        res)
                st.session_state.toast = (
                    "success" if res.get("status") == "success" else "error",
                    res.get("message", ""),
                )
                st.rerun()

        # ---- 工具栏 ----
        st.markdown("**🧹 会话管理**")
        c1, c2 = st.columns(2)
        if c1.button("清空对话", use_container_width=True):
            st.session_state.history = []
            st.session_state.last_records = None
            st.session_state.last_summary = None
            st.session_state.last_owner = ""
            st.session_state.toast = ("success", "对话与结果已清空。")
            st.rerun()
        if c2.button("重置数据", use_container_width=True):
            st.session_state.volunteer_df = pd.DataFrame(DEFAULT_DATA, columns=RECORD_COLUMNS)
            st.session_state.tool_logs = []
            st.session_state.history = []
            st.session_state.last_summary = None
            st.session_state.last_records = None
            st.session_state.last_owner = ""
            st.session_state.excel_file = None
            st.session_state.toast = ("success", "已重置为初始模拟数据。")
            st.rerun()

        st.divider()
        st.markdown("**🔌 接口自检**")
        if st.button("测试模型连接", use_container_width=True):
            with st.spinner("正在测试..."):
                try:
                    t0 = time.time()
                    cli = OpenAI(api_key=API_KEY, base_url=BASE_URL, timeout=30, max_retries=0)
                    r = cli.chat.completions.create(
                        model=MODEL_NAME,
                        messages=[{"role": "user", "content": "ping"}],
                        max_tokens=4,
                    )
                    cost = time.time() - t0
                    st.session_state.diag_result = (
                        "success",
                        f"✅ 接口正常 · {MODEL_NAME} · 耗时 {cost:.2f}s",
                    )
                    log_event("SELFTEST_OK", f"{MODEL_NAME} {cost:.2f}s")
                except Exception as e:
                    st.session_state.diag_result = (
                        "error",
                        f"❌ {type(e).__name__}: {str(e)[:220]}",
                    )
                    log_event("SELFTEST_FAIL", f"{type(e).__name__}: {e}")
        if st.session_state.get("diag_result"):
            kind, msg = st.session_state.diag_result
            (st.success if kind == "success" else st.error)(msg)

        with st.expander("📜 运行日志（最近 30 条）"):
            st.code(read_recent_logs(30), language="text")

        st.divider()
        st.caption(f"🤖 模型：{MODEL_NAME} · DashScope 兼容接口")


def render_results():
    """结果明细面板：把查询到的记录与统计结果用表格清楚展示"""
    _panel_title("📋", "#38A169", "结果明细", "查询与统计的数据表格")
    recs = st.session_state.get("last_records")
    summ = st.session_state.get("last_summary")
    if not recs and not summ:
        st.markdown(
            '<div class="empty-box"><span class="big">🧾</span>'
            "还没有结果<br>提问后这里会显示查询明细与统计汇总表</div>",
            unsafe_allow_html=True,
        )
        return

    tab1, tab2 = st.tabs(["🔍 查询明细", "🧮 统计汇总"])
    with tab1:
        if recs:
            df = pd.DataFrame(recs)
            total = pd.to_numeric(df["时长(小时)"], errors="coerce").fillna(0).sum()
            owner = st.session_state.get("last_owner") or ""
            head = f"{owner} · " if owner else ""
            st.caption(f"{head}共 {len(df)} 条记录 · 合计 {total:.1f} 小时")
            st.dataframe(df, use_container_width=True, height=210)
        else:
            st.caption("本次对话还没有查询结果，先问一次“查询某时间段的志愿记录”。")
    with tab2:
        if summ:
            df2 = pd.DataFrame(summ)
            df2.insert(0, "排名", range(1, len(df2) + 1))
            st.caption(
                f"共 {len(df2)} 人 · 总时长 {df2['总时长(小时)'].sum():.1f} 小时"
            )
            st.dataframe(df2, use_container_width=True, height=210)
        else:
            st.caption("本次对话还没有统计结果，先问一次“统计每个人的志愿总时长”。")


def render_chat():
    _panel_title("💬", "#2B6CB0", "智能对话", "说一句话，Agent 自动串联工具")

    # 快捷指令
    st.markdown(
        '<span class="chip">⚖️ 校内 / 校外各多少小时</span>'
        '<span class="chip">🕘 查询我10月的记录</span>'
        '<span class="chip">📊 统计每个人的总时长</span>'
        '<span class="chip">📍 查某地志愿活动</span>',
        unsafe_allow_html=True,
    )
    c1, c2, c3, c4 = st.columns(4)
    if c1.button("⚖️ 校内外时长", use_container_width=True):
        st.session_state.pending_input = "帮我看看我校内和校外的志愿时长分别是多少，总共多少小时"
    if c2.button("🕘 我的记录", use_container_width=True):
        st.session_state.pending_input = "帮我查询我 2026-10-01 到 2026-10-31 的志愿记录"
    if c3.button("📊 全员统计", use_container_width=True):
        st.session_state.pending_input = "统计 2026-09-01 到 2026-10-31 每个人（全员）的累计志愿总时长，并导出Excel"
    if c4.button("📍 查活动", use_container_width=True):
        st.session_state.pending_input = "学校最近有什么校内志愿活动？"

    # 历史对话：用户提问高亮编号显示，助手回复普通气泡，错误回复红色高亮
    if not st.session_state.history:
        st.markdown(
            '<div class="empty-box"><span class="big">👋</span>'
            "还没有对话<br>试试：“帮我统计 9 月以来每个人的志愿总时长，并导出 Excel”</div>",
            unsafe_allow_html=True,
        )
    q_num = 0
    for role, text in st.session_state.history:
        if role == "user":
            q_num += 1
            safe = html.escape(text).replace("\n", "<br>")
            st.markdown(
                f'<div class="user-bubble"><span class="qtag">❓ 问题 {q_num}</span>{safe}</div>',
                unsafe_allow_html=True,
            )
        else:
            if text.startswith("⚠️"):
                safe = html.escape(text).replace("\n", "<br>")
                st.markdown(f'<div class="err-bubble">{safe}</div>', unsafe_allow_html=True)
            else:
                with st.chat_message("assistant"):
                    st.markdown(
                        f'<div class="ans-tag">🤖 回答 {q_num or 1}</div>',
                        unsafe_allow_html=True,
                    )
                    st.markdown(text)

    # 失败请求一键重试
    if st.session_state.get("last_failed_input"):
        if st.button("🔄 重试上一次失败的请求", use_container_width=True):
            # 清理末尾的失败提示与原提问，保持对话记录干净
            h = st.session_state.history
            if h and h[-1][0] == "assistant" and h[-1][1].startswith("⚠️"):
                h.pop()
            if h and h[-1][0] == "user" and h[-1][1] == st.session_state.last_failed_input:
                h.pop()
            st.session_state.pending_input = st.session_state.last_failed_input
            st.session_state.last_failed_input = None
            st.rerun()


def main():
    init_state()

    # 一次性提示
    if st.session_state.toast:
        kind, msg = st.session_state.toast
        st.session_state.toast = None
        (st.success if kind == "success" else st.error)(msg)

    if not API_KEY:
        render_header()
        st.error("未检测到 DASHSCOPE_API_KEY，请复制 .env.example 为 .env 并填入你的 API Key 后重启。")
        st.stop()

    render_header()
    render_sidebar()          # 先渲染侧栏，账号/口径在本次运行就已确定
    render_stats()            # 两张 KPI 卡：始终为本人数据，随录入/切换账号实时更新
    if st.session_state.get("show_my_records"):
        render_my_records_detail()   # 点开「志愿记录总数」后的逐条明细
    render_account_strip()
    render_school_strip()

    # 双栏：左侧对话，右侧「工具日志 / 结果与导出」双页签
    left, right = st.columns([0.56, 0.44], gap="medium")
    with left:
        render_chat()
    with right:
        tab_log, tab_result = st.tabs(["🧾 工具日志", "📊 结果与导出"])
        with tab_log:
            render_tool_logs()
        with tab_result:
            render_results()
            st.divider()
            render_chart()
            st.divider()
            render_download()

    st.divider()
    render_activities()

    user_input = st.chat_input("💬 输入需求，例如：统计9月以来每个人的志愿总时长并导出Excel")
    pending = st.session_state.pop("pending_input", None) or user_input
    if pending:
        with st.chat_message("user"):
            st.markdown(pending)
        # 实时处理状态：思考 → 每个工具调用 → 完成/失败，全程可见
        with st.status(f"🤖 {MODEL_NAME} 正在思考...", expanded=True) as status:
            run_agent(pending, status)
        st.rerun()


if __name__ == "__main__":
    main()
