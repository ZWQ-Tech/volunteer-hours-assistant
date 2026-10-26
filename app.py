# -*- coding: utf-8 -*-
"""
校园志愿时长汇总助手（路径B：Streamlit + 通义千问 Tool Calling）
====================================================================
- 通过 OpenAI 兼容接口调用阿里云百炼 DashScope 的通义千问模型
- 自主新增 3 个业务工具，由模型真实发起 Tool Call（非 if 关键词判断）
  1. query_volunteer_data    按起止日期（+可选项目名）查询志愿记录
  2. calc_total_hours        按姓名分组统计累计志愿总时长
  3. export_volunteer_excel  导出汇总 Excel（网页内置下载按钮）
- 加分项：日志分区展示 / 表单新增记录 / 强异常容错 / 查询后自动生成柱状图
运行：streamlit run app.py
"""

import html
import io
import json
import os
import time
import traceback
from datetime import datetime
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st
from dotenv import load_dotenv
from openai import OpenAI

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
    "你是校园志愿服务时长管理助手。用户会用自然语言提出查询、统计、导出等需求，"
    "你必须优先调用提供的工具完成业务操作，禁止编造数据。"
    "调用工具前，先根据对话上下文把用户的相对时间（如'本月'、'上周'）换算成 YYYY-MM-DD 格式。"
    "工具返回 status 为 error 时，请把原因友好地转述给用户并给出修正建议；"
    "返回 status 为 empty 时，提示该时间段没有查到记录。"
    "统计前必须先调用查询工具；导出前必须先完成统计。"
    "调用统计和导出工具时不要自己回传数据数组（省略 records / summary 参数即可），"
    "工具会自动使用最近一次查询或统计的结果。回答请简洁、使用中文。"
)

# ============================ 模拟数据 ============================
DEFAULT_DATA = [
    {"姓名": "张三", "学号": "2023001", "志愿项目": "图书馆图书整理", "日期": "2026-09-12", "时长(小时)": 3.0},
    {"姓名": "张三", "学号": "2023001", "志愿项目": "迎新志愿服务", "日期": "2026-09-20", "时长(小时)": 5.0},
    {"姓名": "张三", "学号": "2023001", "志愿项目": "校园运动会保障", "日期": "2026-10-15", "时长(小时)": 4.5},
    {"姓名": "李四", "学号": "2023002", "志愿项目": "图书馆图书整理", "日期": "2026-09-13", "时长(小时)": 2.5},
    {"姓名": "李四", "学号": "2023002", "志愿项目": "社区敬老院慰问", "日期": "2026-10-01", "时长(小时)": 6.0},
    {"姓名": "王五", "学号": "2023003", "志愿项目": "迎新志愿服务", "日期": "2026-09-20", "时长(小时)": 4.0},
    {"姓名": "王五", "学号": "2023003", "志愿项目": "校园运动会保障", "日期": "2026-10-15", "时长(小时)": 3.5},
    {"姓名": "王五", "学号": "2023003", "志愿项目": "实验室开放日引导", "日期": "2026-10-18", "时长(小时)": 2.0},
    {"姓名": "赵六", "学号": "2023004", "志愿项目": "社区敬老院慰问", "日期": "2026-10-01", "时长(小时)": 5.5},
    {"姓名": "赵六", "学号": "2023004", "志愿项目": "图书馆图书整理", "日期": "2026-10-08", "时长(小时)": 3.0},
]

TOOL_ICON = {
    "query_volunteer_data": "🔍",
    "calc_total_hours": "🧮",
    "export_volunteer_excel": "📤",
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
        st.session_state.volunteer_df = pd.DataFrame(DEFAULT_DATA)
    if "tool_logs" not in st.session_state:
        st.session_state.tool_logs = []          # 工具调用日志
    if "history" not in st.session_state:
        st.session_state.history = []            # 聊天记录 [(role, text)]
    if "excel_file" not in st.session_state:
        st.session_state.excel_file = None       # (文件名, bytes)
    if "last_summary" not in st.session_state:
        st.session_state.last_summary = None     # 最近一次统计结果（用于柱状图）
    if "toast" not in st.session_state:
        st.session_state.toast = None            # 一次性提示


# ============================ 工具实现 ============================
def _parse_date(text):
    """严格校验日期格式，非法返回 None（对应容错设计）"""
    try:
        return datetime.strptime(str(text).strip(), "%Y-%m-%d").date()
    except (ValueError, TypeError):
        return None


def query_volunteer_data(start_date: str, end_date: str, project_name: str = "") -> dict:
    """工具1：按起止日期查询志愿记录，project_name 为可选的项目名称筛选关键字"""
    d1, d2 = _parse_date(start_date), _parse_date(end_date)
    if d1 is None or d2 is None:
        return {"status": "error", "message": "日期格式不正确，请使用 YYYY-MM-DD 格式，例如 2026-09-01。"}
    if d2 < d1:
        return {"status": "error", "message": f"结束日期({end_date})早于起始日期({start_date})，请调整后重新查询。"}

    df = st.session_state.volunteer_df.copy()
    df["日期date"] = df["日期"].map(_parse_date)
    mask = (df["日期date"] >= d1) & (df["日期date"] <= d2)
    if project_name:
        mask &= df["志愿项目"].str.contains(str(project_name).strip(), na=False)
    result = df[mask].drop(columns=["日期date"])

    if result.empty:
        tip = f"{start_date} 至 {end_date}" + (f"（项目含“{project_name}”）" if project_name else "")
        return {"status": "empty", "message": f"在 {tip} 时间范围内没有查询到志愿记录。", "records": []}

    return {
        "status": "success",
        "count": int(len(result)),
        "message": f"共查询到 {len(result)} 条志愿记录。",
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
        grouped = (
            df.groupby("姓名", as_index=False)["时长(小时)"]
            .sum()
            .sort_values("时长(小时)", ascending=False)
        )
        counts = df.groupby("姓名").size().reset_index(name="参与次数")
        grouped = grouped.merge(counts, on="姓名")
        grouped["时长(小时)"] = grouped["时长(小时)"].round(2)

        summary = [
            {"姓名": r["姓名"], "总时长(小时)": float(r["时长(小时)"]), "参与次数": int(r["参与次数"])}
            for _, r in grouped.iterrows()
        ]
        return {
            "status": "success",
            "message": f"已统计 {len(summary)} 位同学的累计志愿时长。",
            "summary": summary,
            "总计(小时)": float(round(df["时长(小时)"].sum(), 2)),
        }
    except Exception as e:  # 任何统计异常都不允许程序崩溃
        return {"status": "error", "message": f"统计失败：{e}"}


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
        filename = f"志愿时长汇总_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        st.session_state.excel_file = (filename, excel_bytes)  # 供页面下载按钮使用
        return {
            "status": "success",
            "message": f"汇总表已生成：{filename}（共 {len(df)} 行），可在页面下载区点击按钮下载。",
            "filename": filename,
        }
    except Exception as e:
        return {"status": "error", "message": f"导出失败：{e}"}


TOOL_IMPL = {
    "query_volunteer_data": query_volunteer_data,
    "calc_total_hours": calc_total_hours,
    "export_volunteer_excel": export_volunteer_excel,
}

# 工具 Schema：交给通义千问，由模型自主决定何时调用（真实 Tool Calling）
TOOLS_SCHEMA = [
    {
        "type": "function",
        "function": {
            "name": "query_volunteer_data",
            "description": "按起止日期查询校园志愿记录；可通过可选参数 project_name 按志愿项目名称关键字筛选。",
            "parameters": {
                "type": "object",
                "properties": {
                    "start_date": {"type": "string", "description": "起始日期，格式 YYYY-MM-DD"},
                    "end_date": {"type": "string", "description": "结束日期，格式 YYYY-MM-DD"},
                    "project_name": {"type": "string", "description": "可选。志愿项目名称关键字，不需要时省略"},
                },
                "required": ["start_date", "end_date"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "calc_total_hours",
            "description": "对最近查询到的志愿记录按姓名分组，统计每人累计志愿总时长与参与次数。调用前需先执行 query_volunteer_data；通常不需要传 records，工具会自动使用最近一次查询结果。",
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

    messages = [{"role": "system", "content": SYSTEM_PROMPT}]
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
</style>
"""


def render_header():
    st.markdown(GLOBAL_CSS, unsafe_allow_html=True)
    st.markdown(
        f"""
        <div class="hero">
          <div class="hero-title">⏱️ 校园志愿时长汇总助手</div>
          <div class="hero-sub">自然语言一句直达：查询记录 · 统计时长 · 导出 Excel —— 由通义千问 Tool Calling 真实驱动</div>
          <div>
            <span class="badge">🤖 {MODEL_NAME}</span>
            <span class="badge">🛠️ 3 个业务工具</span>
            <span class="badge">🧾 工具调用日志</span>
            <span class="badge">📊 自动可视化</span>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_stats():
    df = st.session_state.volunteer_df
    cards = [
        ("📋", "#EBF8FF", "#2B6CB0", "志愿记录总数", f"{len(df)}", "条"),
        ("👥", "#F0FFF4", "#38A169", "参与志愿者", f"{df['姓名'].nunique()}", "人"),
        ("⏳", "#FFFAF0", "#DD6B20", "累计志愿时长", f"{df['时长(小时)'].sum():.1f}", "小时"),
        ("🎯", "#FAF5FF", "#805AD5", "覆盖志愿项目", f"{df['志愿项目'].nunique()}", "个"),
    ]
    cols = st.columns(4)
    for col, (icon, bg, fg, label, value, unit) in zip(cols, cards):
        with col:
            st.markdown(
                f"""
                <div class="stat-card">
                  <div class="stat-top">
                    <span class="stat-icon" style="background:{bg};">{icon}</span>
                    <span class="stat-label">{label}</span>
                  </div>
                  <div class="stat-value">{value} <span class="stat-unit">{unit}</span></div>
                </div>
                """,
                unsafe_allow_html=True,
            )


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
    fig = px.bar(
        df,
        x="姓名",
        y="总时长(小时)",
        color="姓名",
        text="总时长(小时)",
        color_discrete_sequence=["#4299E1", "#48BB78", "#ED8936", "#9F7AEA", "#38B2AC", "#F6AD55"],
    )
    fig.update_traces(texttemplate="%{y}", textposition="outside")
    fig.update_layout(
        showlegend=False,
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

        # ---- 数据表 ----
        st.markdown("**📋 当前数据表**（模拟数据 · 可追加）")
        st.dataframe(st.session_state.volunteer_df, use_container_width=True, height=230)

        # ---- 表单新增记录 ----
        st.markdown("**➕ 新增志愿记录**")
        with st.form("add_form", clear_on_submit=True):
            name = st.text_input("姓名")
            sid = st.text_input("学号")
            project = st.text_input("志愿项目")
            d = st.date_input("日期")
            hours = st.number_input("时长(小时)", min_value=0.5, max_value=24.0, step=0.5)
            submitted = st.form_submit_button("✔️ 提交记录")
        if submitted:
            if not name.strip() or not project.strip():
                st.session_state.toast = ("error", "姓名和志愿项目不能为空。")
            elif hours <= 0:
                st.session_state.toast = ("error", "时长必须大于 0。")
            else:
                new_row = {
                    "姓名": name.strip(),
                    "学号": sid.strip() or "-",
                    "志愿项目": project.strip(),
                    "日期": d.strftime("%Y-%m-%d"),
                    "时长(小时)": float(hours),
                }
                st.session_state.volunteer_df = pd.concat(
                    [st.session_state.volunteer_df, pd.DataFrame([new_row])],
                    ignore_index=True,
                )
                st.session_state.toast = ("success", f"已新增：{name.strip()} · {project.strip()} {hours} 小时")

        # ---- 工具栏 ----
        st.markdown("**🧹 会话管理**")
        c1, c2 = st.columns(2)
        if c1.button("清空对话", use_container_width=True):
            st.session_state.history = []
            st.session_state.last_records = None
            st.session_state.last_summary = None
            st.session_state.toast = ("success", "对话与结果已清空。")
            st.rerun()
        if c2.button("重置数据", use_container_width=True):
            st.session_state.volunteer_df = pd.DataFrame(DEFAULT_DATA)
            st.session_state.tool_logs = []
            st.session_state.history = []
            st.session_state.last_summary = None
            st.session_state.last_records = None
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
            st.caption(f"共 {len(df)} 条记录 · 合计 {total:.1f} 小时")
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
        '<span class="chip">🕘 查询10月志愿记录</span>'
        '<span class="chip">📊 统计每人总时长</span>'
        '<span class="chip">📥 导出汇总Excel</span>',
        unsafe_allow_html=True,
    )
    c1, c2, c3 = st.columns(3)
    if c1.button("🕘 查询10月记录", use_container_width=True):
        st.session_state.pending_input = "帮我查询 2026-10-01 到 2026-10-31 的志愿记录"
    if c2.button("📊 统计总时长", use_container_width=True):
        st.session_state.pending_input = "统计 2026-09-01 到 2026-10-31 每个人的累计志愿总时长"
    if c3.button("📥 导出Excel", use_container_width=True):
        st.session_state.pending_input = "把 2026-09-01 到 2026-10-31 的志愿时长汇总导出成 Excel"

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
    render_stats()
    render_sidebar()

    # 双栏布局：左侧对话区，右侧日志 + 可视化 + 导出
    left, right = st.columns([0.58, 0.42], gap="medium")
    with left:
        render_chat()
    with right:
        render_tool_logs()
        st.divider()
        render_results()
        st.divider()
        render_chart()
        st.divider()
        render_download()

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
