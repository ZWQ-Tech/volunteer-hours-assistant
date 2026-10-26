---
marp: true
theme: default
paginate: true
size: 16:9
style: |
  section {
    background: #FBFAF7;
    color: #33414F;
    font-family: "Microsoft YaHei", "PingFang SC", sans-serif;
    padding: 46px 58px;
    border-top: 4px solid #1F5FA8;
  }
  h1 { color: #16202A; font-size: 40px; font-weight: 700; }
  h2 { color: #16202A; font-size: 26px; font-weight: 700; border-bottom: 2px solid #DCE2E7; padding-bottom: 8px; }
  h3 { color: #1F5FA8; font-size: 18px; }
  strong { color: #16202A; }
  table { font-size: 15px; }
  th { color: #74838F; font-size: 13px; border-bottom: 2px solid #DCE2E7; }
  td { border-bottom: 1px solid #EAEEF1; }
  blockquote { border-left: 4px solid #1F5FA8; background: #EAF1F8; padding: 8px 16px; color: #1B3B5C; }
  code { background: #EEF2F5; color: #1B4F7E; border-radius: 3px; padding: 1px 5px; font-size: 14px; }
  ul, ol { font-size: 16px; }
---

<!-- _paginate: false -->

# 校园志愿时长汇总助手

### 路径 B：在普通应用中加入 Agent 能力

基于 Streamlit 与通义千问 Tool Calling 的志愿时长查询、统计与导出工具。
用户以自然语言提出需求，模型调用本地工具完成数据处理并生成 Excel 汇总表。

**姓名 / 学号**：张三 2023001　李四 2023002　｜　**课程**：《开源AI技术实践》
**仓库**：github.com/ZWQ-Tech/volunteer-hours-assistant

---

## 01　背景：数据散、口径乱、报表重

**现状问题**
- 记录分散在各班 Excel，字段命名不统一
- 期末核算需人工筛选、逐人相加，易漏算
- 每次汇总都要重做一份表，报表格式反复返工
- 负责人不熟悉公式，操作门槛高

**需求提炼**
把“按时间段筛选 → 按人汇总 → 出报表”这条固定链路做成一句话可完成的操作，同时保留可核对的原始明细。

**一个学期的样本规模（模拟数据）**：10 条记录 · 4 名志愿者 · 6 个项目 · 累计 39.0 小时

---

## 02　方案：用 Agent 承载一条固定业务链路

模型只做两件事：理解用户要什么、决定按什么顺序调用哪个工具。
所有计算与文件生成在本地完成，结果可由原始记录核对。

| 能力 | 做法 | 输出 |
| --- | --- | --- |
| 条件查询 | 按起止日期检索，可用项目名进一步筛选 | 结构化记录数组 |
| 分组统计 | 按姓名分组累加时长，统计参与次数 | 降序排名 + 合计 |
| 导出报表 | openpyxl 写出 xlsx，挂到页面下载区 | 带时间戳的文件名 |

> 范围界定：面向校团委与院系志愿负责人，单机运行。

---

## 03　总体架构：五层结构

```
A 交互层      Streamlit 页面：对话、日志、表格、图表、下载
                ↓ 用户输入文本
B 会话层      Agent 循环：messages + tools 协议，工具调用上限 6 轮
                ↓ chat.completions.create(model, messages, tools)
C 模型层      通义千问 qwen-plus（DashScope 兼容接口）
                ↓ 返回 tool_calls（工具名 + JSON 参数）
D 工具层      query_volunteer_data / calc_total_hours / export_volunteer_excel
                ↓ pandas 读写 · openpyxl 落盘
E 数据层      内存数据表（10 条模拟记录）+ 最近查询/统计结果缓存
```

**关键约束**：模型不接触原始数据，只传参数；数据计算全部在本地完成。

---

## 04　运行机制：一次请求的完整时序

| # | 环节 | 动作 |
| --- | --- | --- |
| 01 | 用户 → 界面 | 输入需求，提交到 `st.chat_input` |
| 02 | Agent → 模型 | 拼装 system + 历史消息 + tools，发起请求 |
| 03 | 模型 → Agent | 返回 `tool_calls`，指定查询工具与日期参数 |
| 04 | Agent → 工具 | 执行筛选，返回 `{"status":"success","count":10}` |
| 05 | Agent → 模型 | 以 `role=tool` 回填，模型发起统计调用 |
| 06 | 工具内部 | 分组累加：张三 12.5 / 王五 9.5 / 李四 8.5 / 赵六 8.5 |
| 07 | Agent → 模型 | 回填统计结果，模型发起导出调用 |
| 08 | 工具 → 页面 | 写出 xlsx，下载区出现按钮 |
| 09 | 模型 → 用户 | 生成结论并结束，日志保留三次调用 |

全流程约 3 次模型请求，限流时自动退避重试。

---

## 05　工具设计：入参与返回

| 工具 | 入参 | 返回 |
| --- | --- | --- |
| `query_volunteer_data` | `start_date` 必填<br>`end_date` 必填<br>`project_name` 可选 | `{status, count, message, records[]}` |
| `calc_total_hours` | `records` 可选（省略取最近查询结果） | `{status, summary[{姓名,总时长,参与次数}], 总计}` |
| `export_volunteer_excel` | `summary` 可选（省略取最近统计结果） | `{status, filename, message}` |

**三点设计取舍**
- **可选参数**：用户可以不提项目名，模型也能自行补默认行为
- **数据不回传**：统计与导出省略数组参数，避免长 JSON 被截断导致 400
- **状态三值**：`success / empty / error` 让模型分别给出结果、提示或修正建议

> 参数由模型生成，工具内部再做一次日期与类型校验，不信任外部输入。

---

## 06　功能演示：明细、统计、报表一次完成

**实际工具调用记录**

```
21:58:12  query_volunteer_data     success
          入参 start=2026-09-01 end=2026-10-31
          结果 共查询到 10 条志愿记录

21:58:13  calc_total_hours         success
          入参 (省略 records，取上次查询结果)
          结果 已统计 4 位同学，总计 39.0 小时

21:58:14  export_volunteer_excel   success
          结果 志愿时长汇总_20261026_215814.xlsx
```

**统计输出**：张三 12.5h（3 次）· 王五 9.5h（3 次）· 李四 8.5h（2 次）· 赵六 8.5h（2 次）· 合计 39.0h

<!-- 截图：主界面 + 链式调用日志区 -->

---

## 07　所用工具：技术栈与选型

| 组件 | 版本 | 承担的工作 | 选型理由 |
| --- | --- | --- | --- |
| Python | 3.13 | 运行环境 | 课程统一环境，生态完整 |
| Streamlit | ≥ 1.36 | 页面、交互、会话状态 | 纯 Python 出可用界面 |
| openai SDK | ≥ 1.40 | 调 DashScope 兼容接口 | 协议通用，便于换模型 |
| qwen-plus | — | 意图理解与工具决策 | 中文稳定，支持 Function Calling |
| pandas | ≥ 2.0 | 过滤、分组、汇总 | 几行完成统计，结果可复核 |
| plotly | ≥ 5.20 | 柱状图 | 交互式图表，易嵌入 |
| openpyxl | ≥ 3.1 | 生成 xlsx | 可直接用于台账归档 |
| python-dotenv | ≥ 1.0 | 密钥加载 | 密钥与代码分离 |
| Streamlit AppTest | 内置 | 端到端回归测试 | 无浏览器验证工具链 |

**运行**：`pip install -r requirements.txt` → `copy .env.example .env` → `streamlit run app.py`

---

## 08　与课堂实验的关系

**沿用课堂 demo**

| 来源 | 复用内容 |
| --- | --- |
| demo_1（Streamlit 界面） | 页面骨架、`st.chat_input` 交互、`st.session_state` 用法 |
| demo_2（接口与工具调用） | `OpenAI(api_key, base_url)` 写法、`tools` 协议结构、结果回填循环 |

**本项目新增**
- 三个业务工具：入参、返回结构、异常分支均自行设计
- 业务数据层：模拟表、条件筛选、按姓名分组、Excel 生成
- 稳定性处理：限流退避、备用模型、参数格式拦截、失败可重试
- 界面与可观测：双栏布局、日志分区、结果表格、运行日志
- 测试与交付：AppTest 验证、requirements、.env 模板、仓库提交

> 沿用的是“怎么把模型接进来”，业务逻辑部分全部重写。

---

## 09　数据来源、隐私与使用边界

**数据来源**
- 全部为课程自建模拟数据，无任何真实学生信息
- 姓名、学号均为虚构占位（张三 / 2023001）
- 数据仅存在于本机内存，不写入第三方服务

**隐私与密钥**
- 接口密钥存 `.env`，由 `.gitignore` 排除
- 代码、仓库、幻灯片中均不含密钥

**使用边界**
- 定位为课程演示原型，不作为正式认定依据
- 真实场景需补：登录鉴权、数据脱敏、操作留痕
- 统计结果须以学校志愿管理系统原始记录为准

**学术诚信**
- 工具调用由模型真实发起，未用关键词判断伪造能力；数据由工具真实计算

---

## 10　已知不足与后续计划

**当前不足**

| 问题 | 影响 |
| --- | --- |
| 数据存于内存 | 刷新或重启即恢复样例，无法长期积累 |
| 单用户运行 | 无登录与权限，多人使用会互串状态 |
| 筛选维度有限 | 仅日期与项目名，无法按姓名或时长区间查询 |
| 相对时间依赖模型 | “上个月”等表述偶有偏差，需补具体日期 |
| 无导出留痕 | 多次导出只保留最新一份 |

**后续计划**
- 数据持久化：接入 SQLite，按学期归档
- 权限与流程：登录鉴权 + 提交、审核、归档，导出需确认
- 工具扩展：个人明细、月度趋势、未达标提醒
- 部署：容器化，支持多人访问

---

## 11　小结

**完成的工作**
- 三个自主设计的业务工具
- 基于真实 Function Calling 的调用链
- 可演示的界面、日志与结果展示
- 仓库交付：代码、文档、幻灯片

**遇到的问题**
- 模型回传大数组导致参数截断（已改为工具端取缓存）
- 业务空间限流造成调用中断（已加退避重试与备用模型）
- 编辑器主题与页面配色冲突（已固定浅色主题）

**收获**
- Function Calling 的分工边界：模型决策，代码执行
- 工具参数设计对稳定性的影响大于提示词
- 异常与边界处理决定能否用于真实业务
