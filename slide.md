---
marp: true
theme: default
paginate: true
size: 16:9
style: |
  section {
    background: #FFFFFF;
    color: #3B4A5A;
    font-family: "Microsoft YaHei", "PingFang SC", sans-serif;
    padding: 42px 54px;
    border-top: 5px solid #2563EB;
    text-align: center;
  }
  h1 { color: #0F172A; font-size: 40px; font-weight: 800; }
  h2 { color: #0F172A; font-size: 26px; font-weight: 800; }
  h2 em { font-style: normal; color: #2563EB; }
  h3 { color: #2563EB; font-size: 17px; }
  strong, b { color: #0F172A; }
  table { font-size: 14px; text-align: left; }
  th { color: #7B8A99; font-size: 12px; letter-spacing: .08em; border-bottom: 2px solid #E4EAF1; }
  td { border-bottom: 1px solid #E4EAF1; }
  blockquote { border-left: 4px solid #2563EB; background: #F4F8FF; padding: 10px 16px; color: #1B3550; text-align: left; }
  code { background: #EFF3F7; color: #1B4F7E; border-radius: 3px; padding: 1px 5px; font-size: 13px; }
  ul, ol { font-size: 15px; text-align: left; display: inline-block; }
  p { font-size: 15px; }
  section.cover { background: #0B1220; border-top: none; }
  section.cover h1 { color: #fff; }
  section.cover h2, section.cover p, section.cover strong { color: #B9C6DC; }
---

<!-- _paginate: false -->
<!-- _class: cover -->

# 校园志愿时长汇总助手

### Streamlit × 通义千问 Tool Calling

用一句话完成「查询 · 统计 · 导出报表」，并推送身边的志愿活动

**姓名 / 学号**：张三 2023001 · 李四 2023002
**课程**：《开源AI技术实践》期末汇报 · 路径 B
**仓库**：github.com/ZWQ-Tech/volunteer-hours-assistant

---

## 01 · 背景：数据散、口径乱、报表重

**一个学期的样本规模**：10 条记录 · 4 名志愿者 · 6 个项目 · 累计 39.0 小时

| 手工流程 | 本项目的目标 |
| --- | --- |
| ① 收集各班表 → ② 统一字段 → ③ 按日期筛 → ④ 逐人累加 → ⑤ 重做汇总表 | 说清时间范围与需求，工具自动执行 |
| 任一步改数据，后面全部重来 | 每个数字可回溯到原始记录，报表可直接归档 |

> 目标：一句话拿到可核对的汇总表。

---

## 02 · 方案：把重复劳动交给 Agent 执行

模型只做两件事：理解要什么、决定调用哪个工具；计算与出表全部在本地完成。

| 能力 | 做法 | 输出 |
| --- | --- | --- |
| 条件查询 | 按起止日期检索，可用项目名筛选 | 结构化记录数组 |
| 分组统计 | 按姓名分组累加，统计参与次数 | 4 人排名，合计 39.0 小时 |
| 导出报表 | 写 xlsx 并挂到页面 | 志愿时长汇总_日期_时分.xlsx |

**链路一 · 时长统计**：查询 → 统计 → 导出，一句话串起三个工具
**链路二 · 活动发现**：绑定学校 → `query_open_activities` 一次调用返回结果，校内看实时刷新、校外按距离排队

**范围界定**：面向校团委与院系志愿负责人，单机运行，两个链路的业务数据均为课程自建模拟数据。

---

## 03 · 架构：五层结构，模型决策 / 代码执行

```
A 交互层   Streamlit 页面：对话 · 日志 · 结果表格 · 柱状图 · 下载 · 活动中心
            ↓ 用户输入文本
B 会话层   messages + st.session_state + tools 协议，调用循环上限 6 轮
            ↓ chat.completions.create(model, messages, tools)
C 模型层   通义千问 qwen-plus（DashScope 兼容接口）：意图理解与工具选择
            ↓ 返回 tool_calls（工具名 + JSON 参数）
D 工具层   查询 / 统计 / 导出 / 活动查询，执行后以 JSON 回填，形成下一轮上下文
            ↓ pandas 筛选分组 · openpyxl 写出文件 · Haversine 算距离
E 数据层   内存数据表（10 条记录）+ 学校与活动池 + 最近结果缓存
```

**关键约束**：模型不接触原始数据，只传参数；计算与文件生成全部在本地完成。

---

## 04 · 运行机制：一句话触发三次工具调用

| # | 环节 | 动作 |
| --- | --- | --- |
| 01 | 用户提交 | 需求进入 `st.chat_input` |
| 02 | 拼装上下文 | system + 历史消息 + tools |
| 03 | 模型返回 | `tool_calls`，指定查询工具与日期 |
| 04 | 执行查询 | 返回 `{"status":"success","count":10}` |
| 05 | 回填结果 | `role=tool`，模型发起统计调用 |
| 06 | 统计工具 | 分组累加 12.5 / 9.5 / 8.5 / 8.5 |
| 07 | 再次回填 | 模型发起导出调用 |
| 08 | 写出 xlsx | 下载区立即出现按钮 |
| 09 | 模型生成 | 中文结论，本轮结束 |

> 全流程约 3 次模型请求；限流或超时自动退避重试，失败可一键重发。
> 若问的是活动（如"附近 3 公里有环保活动吗"），走单轮分支：直接调用 `query_open_activities` 并返回距离排序结果。

---

## 05 · 工具设计：四个工具的入参、返回与四点取舍

| 工具 | 入参 | 返回 |
| --- | --- | --- |
| `query_volunteer_data` | `start_date` 必填 · `end_date` 必填 · `project_name` 可选 | `{status, count, message, records[]}` |
| `calc_total_hours` | `records` 可选（省略取最近查询结果） | `{status, summary[{姓名,总时长,参与次数}], 总计}` |
| `export_volunteer_excel` | `summary` 可选（省略取最近统计结果） | `{status, filename, message}` |
| `query_open_activities` | `scope`（all/on_campus/off_campus）· `radius_km` · `keyword`，全可选 | `{status, school, on_campus[], off_campus[]}` |

**取舍一 · 可选参数**：用户不必说全条件，模型可采用默认行为
**取舍二 · 数据不回传**：省略数组参数，避免长 JSON 被截断导致 400
**取舍三 · 状态三值**：`success / empty / error` 分别对应结果、提示、修正建议
**取舍四 · 距离不由模型算**：半径与排序交给 Haversine 公式，防止模型编造距离

---

## 06 · 功能演示：提问 · 日志 · 结果 · 活动中心

**界面四块区域**

- 左：对话区（用户提问高亮编号，助手回答卡片）
- 右上：工具调用日志（工具名 + 入参 + 状态 + 结果摘要）
- 右下：结果明细表格 · 柱状图 · 导出下载按钮
- 底部：活动中心（校内 / 校外两页签，校外按距离升序，可调推送半径）

**实际运行记录**

```
query_volunteer_data    成功  入参 2026-09-01 ~ 2026-10-31  结果 共 10 条记录
calc_total_hours        成功  入参（省略 records）          结果 4 人 · 39.0 小时
export_volunteer_excel  成功  结果 志愿时长汇总_215814.xlsx
query_open_activities   成功  入参 scope=off_campus, radius_km=5  结果 4 项，最近 0.3 km
```

<!-- 此处放实际界面截图 -->

---

## 07 · 结果：统计自动成图，差距一眼可见

| 姓名 | 总时长 | 次数 | 占比 |
| --- | --- | --- | --- |
| 张三 | 12.5 | 3 | 32.1% |
| 王五 | 9.5 | 3 | 24.4% |
| 李四 | 8.5 | 2 | 21.8% |
| 赵六 | 8.5 | 2 | 21.8% |
| **合计** | **39.0** | **10** | **100%** |

统计成功后页面自动绘制柱状图，与明细表、导出文件逐行一致。

---

## 08 · 所用工具：9 个组件各司其职

| 组件 | 承担的工作 |
| --- | --- |
| Python 3.13 | 运行环境 |
| Streamlit ≥ 1.36 | 页面、交互、会话状态 |
| openai SDK ≥ 1.40 | 调用 DashScope 兼容接口 |
| qwen-plus | 意图理解与工具决策 |
| pandas ≥ 2.0 | 筛选、分组、汇总 |
| plotly ≥ 5.20 | 交互式柱状图 |
| openpyxl ≥ 3.1 | 生成 xlsx 报表 |
| python-dotenv | 密钥与代码分离 |
| AppTest / Git | 端到端测试 · 版本管理 |

**启动**：`pip install -r requirements.txt` → `copy .env.example .env` → `streamlit run app.py`

---

## 09 · 与课堂实验的关系：沿用的是接入方式

| 来源 | 复用内容 |
| --- | --- |
| demo_1（Streamlit 界面） | 页面骨架、`st.chat_input` 交互、`st.session_state` 用法 |
| demo_2（接口与工具调用） | `OpenAI(api_key, base_url)` 写法、`tools` 协议结构、结果回填循环 |

**本项目新增**
- 四个业务工具：入参、返回结构、异常分支均自行设计（含活动查询与距离排序）
- 业务数据层：模拟表、条件筛选、按姓名分组、Excel 生成、学校与活动池
- 稳定性处理：限流退避重试、备用模型、参数格式拦截
- 界面与可观测：双栏布局、日志分区、结果表格、活动中心、运行日志
- 测试与交付：AppTest 验证、依赖清单、环境模板、仓库提交

---

## 10 · 合规说明：数据来源、隐私与使用边界

**数据来源**：全部为课程自建模拟数据，姓名学号为虚构占位；仅存于本机内存，不写入第三方服务
**隐私与密钥**：密钥存 `.env` 并由 gitignore 排除，代码、仓库、幻灯片中均不含真实 Key
**使用边界**：课程演示原型，不作正式认定依据；真实场景需补登录鉴权、数据脱敏、操作留痕
**学术诚信**：工具调用由模型真实发起，未用关键词判断伪造能力；数据由工具真实计算
**位置信息**：仅用学校公开的校区经纬度估算直线距离，不采集、不存储用户实时定位

> 统计结果最终以学校志愿管理系统的原始记录为准。

---

## 11 · 反思与计划

**当前不足**
- 数据存于内存，重启即恢复样例
- 单用户运行，无登录与多用户隔离
- 筛选维度有限，仅日期与项目名
- 活动为本地模拟数据，非校方实时接口
- 校外距离为直线估算，未接路网导航
- 相对时间依赖模型换算，偶有偏差
- 多次导出只保留最新一份，无留痕

**后续计划**
- 接入 SQLite，按学期归档数据
- 登录鉴权 + 提交、审核、归档三步流程
- 扩展工具：个人明细、月度趋势、未达标提醒
- 对接校团委活动源与地图路线 API
- 容器化部署，支持多人同时访问

**现场演示动线**：① 绑定学校 → ② 看活动中心 → ③ 一句统计 → ④ 导出下载

**谢谢，欢迎提问**
