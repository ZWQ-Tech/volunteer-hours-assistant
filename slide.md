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
  pre { text-align: left; background: #0E1A26; border-radius: 6px; padding: 12px 16px; }
  pre code { background: transparent; color: #CBD8E2; font-size: 12.5px; line-height: 1.6; }
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

用一句话完成「登记 · 查询 · 统计 · 导出报表」，并推送上海全市的志愿活动

**5 个业务工具** · 个人账号绑定 · 学校绑定与全市活动推送 · 网页内导出 Excel

**姓名 / 学号**：张三 2023001 · 李四 2023002
**课程**：《开源AI技术实践》期末汇报 · 路径 B
**仓库**：github.com/ZWQ-Tech/volunteer-hours-assistant

---

## 01 · 背景：数据散、活动散、报表重

**一个学期的样本规模**：10 条记录 · 4 名志愿者 · 6 个项目 · 累计 39.0 小时

| 痛点一 · 手工流程 | 痛点二 · 活动分散 | 本项目目标 |
| --- | --- | --- |
| ① 收表 → ② 统一字段 → ③ 按日期筛 → ④ 逐人累加 → ⑤ 重做汇总 | 活动散落在班群、公众号、海报里，报名链接与截止时间各不相同 | 说清时间范围与需求，工具自动执行 |
| 任一步改数据，后面全部重来 | 想凑满时长，得先自己筛一遍信息 | 每个数字可回溯到原始记录 |

**为什么这件事适合交给 Agent**

- 规则明确 → 能工具化：筛选、分组、求和都是确定性操作
- 结果可验 → 不怕幻觉：每个数字都能回到原始记录
- 入口自然 → 适合对话：学生只会说“帮我统计一下”，不写公式

> 规模说明：真实院系一学期通常是数百条记录；本 demo 用 10 条压缩演示，保证一次答辩跑完全部链路。

---

## 02 · 方案：把重复劳动交给 Agent 执行

模型只做两件事：理解要什么、决定调用哪个工具；计算与出表全部在本地完成。

| 能力 | 做法 | 输出 |
| --- | --- | --- |
| 条件查询 | 按起止日期检索，可用项目名筛选 | 结构化记录数组 |
| 分组统计 | 按姓名分组累加，统计参与次数 | 4 人排名，合计 39.0 小时 |
| 导出报表 | 写 xlsx 并挂到页面 | 志愿时长汇总_日期_时分.xlsx |
| 记录登记 | 只填活动名称与时长，身份取自绑定账号 | 写入记录并返回本人累计 |

**链路一 · 时长统计**：查询 → 统计 → 导出，一句话串起三个工具
**链路二 · 活动发现**：绑定学校 → `query_open_activities` 一次调用返回结果，校内实时刷新、上海全市校外按距离排队并可点开报名
**链路三 · 活动登记**：绑定个人账号 → 只说"做了什么、几小时"即可写入，写完立刻可查询、统计、导出

**范围界定**：面向在校学生与院系志愿负责人，单机运行，三条链路的业务数据均为课程自建模拟数据。

---

## 03 · 架构：五层结构，模型决策 / 代码执行

```
A 交互层   Streamlit 页面：对话 · 日志 · 结果表格 · 柱状图 · 下载 · 活动中心
            ↓ 用户输入文本
B 会话层   messages + st.session_state + tools 协议，调用循环上限 6 轮
            ↓ chat.completions.create(model, messages, tools)
C 模型层   通义千问 qwen-plus（DashScope 兼容接口）：意图理解与工具选择
            ↓ 返回 tool_calls（工具名 + JSON 参数）
D 工具层   查询 / 统计 / 导出 / 活动查询 / 记录录入，执行后以 JSON 回填
            ↓ pandas 筛选分组 · openpyxl 写出文件 · Haversine 算距离
E 数据层   内存数据表（10 条记录）+ 学校与活动池 + 最近结果缓存
```

| 层 | app.py 中的关键函数 |
| --- | --- |
| A 交互层 | `main()` · `render_chat()` · `render_stats()` |
| B 会话层 | `run_agent()` · 拼装 messages 与 TOOLS_SCHEMA |
| C 模型层 | `create_completion_with_retry()` |
| D 工具层 | `execute_tool()` + 5 个业务函数 |
| E 数据层 | `init_state()` · `st.session_state.volunteer_df` |

**关键约束**：模型不接触原始数据，只传参数；计算与文件生成全部在本地完成。

---

## 04 · 代码讲解 ①：核心循环，模型只决定调谁

**`run_agent()` 一次请求最多 6 轮：模型给工具名与参数，本地执行后把结果回填，再让模型继续判断**

```python
for _ in range(6):                        # ① 轮次上限，防失控
    resp = client.chat.completions.create(
        model=MODEL_NAME, messages=messages,
        tools=TOOLS_SCHEMA, temperature=0.3)
    msg = resp.choices[0].message

    if msg.tool_calls:                    # ② 模型决定调哪个工具
        messages.append({"role": "assistant",
                         "tool_calls": normalized_calls})
        for name, args, call_id in parsed_calls:
            result = execute_tool(name, args)      # ③ 本地真实执行
            messages.append({"role": "tool",       # ④ 结果回填
                "tool_call_id": call_id,
                "content": json.dumps(result, ensure_ascii=False)})
        continue                          # ⑤ 带着结果再问一次模型

    history.append(("assistant", msg.content))  # ⑥ 自然语言收尾
    return
```

- ① **轮次上限 6**：防止模型反复调工具停不下来，超限给出明确提示并可重试
- ② **决策权在模型**：代码里没有 `if "统计" in text` 这类关键词分支
- ③ **执行权在代码**：`execute_tool` 是本地真实函数，pandas 筛选、openpyxl 写文件
- ④ **回填形成闭环**：`role=tool` 带 `tool_call_id`，模型下一轮才能"看到"真实结果
- ⑤ **六轮仍未收敛** → 提示拆分需求，不无限重试；参数 JSON 截断时先规范化再回填，避免 400

> 三个数字：**6 轮**调用上限 · 典型链路 **3 次**工具调用 · 代码中 **0 处**关键词分支。

---

## 05 · 代码讲解 ②：工具 = JSON Schema + 校验关

**Schema 决定模型"能不能填对"，校验决定"填错了会不会炸"——两侧都要设计**

```python
{ "type": "function", "function": {
  "name": "add_volunteer_record",
  "description": "为当前绑定账号登记一条志愿记录",
  "parameters": {
    "type": "object",
    "properties": {
      "project_name": {"type": "string"},
      "hours": {"type": "number"},
      "date": {"type": "string", "description": "缺省=今天"}
    },
    "required": ["project_name", "hours"]
  }}}
```

```python
if d > date.today():  return {"status": "error", ...}   # 未来日期
if not 0 < h <= 24:   return {"status": "error", ...}   # 时长区间
if dup.shape[0]:      return {"status": "error", ...}   # 同人同日重复
```

- **必填只留两个**：活动名 + 时长是用户一定会说的；日期、地点可不给
- **身份不进参数**：姓名学号取自 `bound_account`，模型无权填写，杜绝张冠李戴
- **三类硬校验**：日期格式与未来日期、时长 0<h≤24（统一 0.5 刻度）、同人同日重复
- **错误说人话**：`status=error` + 具体原因，模型据此解释并请求修正，页面不崩溃
- **返回结构**：`{status, record, personal_total}`，模型直接回答"本人累计 16.0 小时"

> 参数设计的检验标准：模型第一次就能填对；填错时程序能拦住并说清原因。

---

## 06 · 代码讲解 ③：换账号，卡片 / 明细 / 对话一起换

**`switch_account()` 是唯一切换入口：先换身份，再清缓存；顶部卡片始终是"本人数据的实时派生"**

```python
def switch_account(acc):
    bound_account = dict(acc) if acc else None
    show_all_rows = False          # 回到只看本人
    reset_record_views()           # 清空查询 / 统计 / 导出 / 对话

def render_stats():                # 两张卡 = 本人数据
    mine = my_records_df()         # df[df.姓名 == 绑定账号]
    n, hrs = len(mine), mine["时长(小时)"].sum()
    if clicked():                  # 点卡片 → 展开明细
        show_my_records = True
        st.rerun()
```

| 切换前 · 张三 | 切换后 · 李四 | 再录一条 · 张三 |
| --- | --- | --- |
| 卡片 3 条 / 12.5h · 明细 3 行 · 对话 3 轮 | 卡片 2 条 / 8.5h · 明细 2 行 · 对话已清空 | 卡片立刻 4 条 / 16.0h，明细同步多一行 |

- **唯一入口**：绑定、切换、退出都走这一个函数，避免多处各写 session_state 导致不同步
- **缓存一起清**：`last_records / last_summary / excel_file / history` 全清，防止拿上一账号的结果回答
- **卡片是派生值**：不额外存统计，直接由 df 过滤求和，所以永远实时、不会与记录打架
- **点开即下钻**：点「志愿记录总数」整卡展开逐条明细（日期/项目/时长/地点/备注）

> Streamlit 每次交互都会重跑整个脚本，`st.session_state` 是唯一事实来源；多处各写一份就会出现"换了一半"的不一致。

---

## 07 · 运行机制：一句话触发三次工具调用

| # | 环节 | 动作 |
| --- | --- | --- |
| 01 | 用户提交 | 需求进入 `st.chat_input` |
| 02 | 拼装上下文 | system（含今天日期、当前账号）+ 历史消息 + tools |
| 03 | 模型返回 | `tool_calls`，指定查询工具与日期 |
| 04 | 执行查询 | 返回 `{"status":"success","count":10}` |
| 05 | 回填结果 | `role=tool`，模型发起统计调用 |
| 06 | 统计工具 | 分组累加 12.5 / 9.5 / 8.5 / 8.5 |
| 07 | 再次回填 | 模型发起导出调用 |
| 08 | 写出 xlsx | 下载区立即出现按钮 |
| 09 | 模型生成 | 中文结论，本轮结束 |

**messages 在三次请求中如何增长**

```
system → user → assistant(tool_calls: query) → tool(records)
       → assistant(tool_calls: calc) → tool(summary)
       → assistant(tool_calls: export) → tool(filename) → assistant(最终结论)
```

> 全流程约 3 次模型请求；限流或超时自动退避重试，失败可一键重发。
> 活动查询 / 记录登记走单轮分支：一次工具调用即可返回。

---

## 08 · 工具设计：五个工具的入参、返回与五点取舍

| 工具 | 入参 | 返回 |
| --- | --- | --- |
| `query_volunteer_data` | `start_date` 必填 · `end_date` 必填 · `project_name` 可选 · `scope`（me 本人 / all 全员，默认 me） | `{status, scope, owner, count, records[]}` |
| `calc_total_hours` | `records` 可选（省略取最近查询结果） | `{status, summary[{姓名,总时长,参与次数}], 总计}` |
| `export_volunteer_excel` | `summary` 可选（省略取最近统计结果） | `{status, filename, message}` |
| `query_open_activities` | `scope`（all/on_campus/off_campus）· `district`（上海各区，不传=全市）· `keyword`，全可选 | `{status, school, area, on_campus[], off_campus[]}` |
| `add_volunteer_record` | `project_name` 必填 · `hours` 必填 · `date` / `place` / `note` 可选 | `{status, record, personal_total, personal_count}` |

**取舍一 · 可选参数**：用户不必说全条件，模型可采用默认行为
**取舍二 · 数据不回传**：省略数组参数，避免长 JSON 被截断导致 400
**取舍三 · 状态三值**：`success / empty / error` 分别对应结果、提示、修正建议
**取舍四 · 距离不由模型算**：距离排序交给 Haversine 公式，区域由参数限定，防止模型编造
**取舍五 · 身份不进参数**：姓名学号取自绑定账号，模型无需询问也无权填写，避免张冠李戴

---

## 09 · 界面与功能：账号 · 对话 · 日志 · 活动中心

**界面四块区域**

- 顶部：两张 KPI 卡——「我的志愿记录总数」（点击展开逐条明细）与「我的累计志愿时长」，只显示本人数据、随录入实时更新
- 左：对话区（用户提问高亮编号，助手回答卡片）
- 右上：工具调用日志（工具名 + 入参 + 状态 + 结果摘要）
- 右下：结果明细表格 · 柱状图 · 导出下载按钮
- 切换账号时，卡片 / 明细 / 数据表 / 查询结果 / 导出文件同步切换，旧账号结果与对话会清空
- 底部：活动中心（校内 / 上海全市校外两页签，可筛区、可搜索，卡片可展开看详情并跳转报名）

**实际运行记录**

```
query_volunteer_data    成功  入参 2026-09-01 ~ 2026-10-31 · scope=all  结果 共 10 条记录
calc_total_hours        成功  入参（省略 records）          结果 4 人 · 39.0 小时
export_volunteer_excel  成功  结果 志愿时长汇总_215814.xlsx
add_volunteer_record    成功  入参 社区清洁 · 3.5h          结果 本人累计 16.0 小时
query_open_activities   成功  入参 scope=off_campus, district=浦东新区  结果 4 项，最近 2.4 km
```

<!-- 此处放实际界面截图：顶部两张 KPI 卡 + 点开后的记录明细面板 -->

---

## 10 · 结果：统计自动成图，差距一眼可见

| 姓名 | 总时长 | 次数 | 占比 |
| --- | --- | --- | --- |
| 张三 | 12.5 | 3 | 32.1% |
| 王五 | 9.5 | 3 | 24.4% |
| 李四 | 8.5 | 2 | 21.8% |
| 赵六 | 8.5 | 2 | 21.8% |
| **合计** | **39.0** | **10** | **100%** |

```python
df = pd.DataFrame(st.session_state.last_summary)
fig = px.bar(df, x="姓名", y="总时长(小时)", color="姓名", text="总时长(小时)")
fig.update_traces(texttemplate="%{y}", textposition="outside")
st.plotly_chart(fig, use_container_width=True)
```

**同源**：图表 = 表格 = Excel，三者都来自 `last_summary`，不会"图对表不对"
**空结果不画图**：查询无数据只给友好提示，页面不出现空图或报错
**颜色与人一一对应**：固定色板，同一人每次颜色一致

---

## 11 · 所用工具：9 个组件各司其职

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

## 12 · 与课堂实验的关系：沿用的是接入方式

| 来源 | 复用内容 |
| --- | --- |
| demo_1（Streamlit 界面） | 页面骨架、`st.chat_input` 交互、`st.session_state` 用法 |
| demo_2（接口与工具调用） | `OpenAI(api_key, base_url)` 写法、`tools` 协议结构、结果回填循环 |

**本项目新增**
- 五个业务工具：入参、返回结构、异常分支均自行设计（含活动查询、距离排序与记录录入）
- 业务数据层：模拟表、条件筛选、按姓名分组、Excel 生成、学校与活动池
- 顶部数据卡：两张卡只显示本人数据，点「志愿记录总数」可下钻逐条明细
- 稳定性处理：限流退避重试、备用模型、参数格式拦截
- 界面与可观测：双栏布局、日志分区、结果表格、活动中心、运行日志
- 测试与交付：AppTest 验证、依赖清单、环境模板、仓库提交

**考核要求 → 本项目落点**

| 考核要求 | 本项目落点 |
| --- | --- |
| 真实 Tool Calling，不是 if 关键词判断 | `run_agent()` 循环 + `tool_calls` 分派，代码中无关键词分支 |
| 自主新增业务工具 | 5 个工具（查询 / 统计 / 导出 / 活动 / 录入），入参与返回自设计 |
| 参数设计说明 | 必填最少 + 三值状态 + 身份不进参数 |
| 异常容错与边界 | 日期格式 / 未来日期 / 时长区间 / 重复记录校验，空结果友好提示 |
| 加分项：日志分区、网页下载、表单、柱状图 | 全部实现并可视化 |
| 文档与交付 | README + 环境模板 + AppTest + 本幻灯片 |

---

## 13 · 合规说明：数据来源、隐私与使用边界

**数据来源**：全部为课程自建模拟数据，姓名学号为虚构占位；仅存于本机内存，不写入第三方服务
**隐私与密钥**：密钥存 `.env` 并由 gitignore 排除，代码、仓库、幻灯片中均不含真实 Key
**使用边界**：课程演示原型，不作正式认定依据；真实场景需补登录鉴权、数据脱敏、操作留痕
**学术诚信**：工具调用由模型真实发起，未用关键词判断伪造能力；数据由工具真实计算
**位置信息**：仅用学校公开的校区经纬度估算直线距离，不采集、不存储用户实时定位

> 统计结果最终以学校志愿管理系统的原始记录为准。

---

## 14 · 反思与计划

**当前不足**
- 数据存于内存，重启即恢复样例
- 单用户运行，无登录与多用户隔离
- 明细面板只读，暂不支持页面内改 / 删
- 筛选维度有限，仅日期与项目名
- 活动为本地模拟数据，非校方实时接口
- 校外距离为直线估算，未接路网导航
- 相对时间依赖模型换算，偶有偏差
- 多次导出只保留最新一份，无留痕

**后续计划**
- 接入 SQLite，按学期归档数据
- 登录鉴权 + 提交、审核、归档三步流程
- 记录支持在线修改、撤销与操作留痕
- 扩展工具：月度趋势、未达标提醒、按项目分类
- 对接校团委活动源与地图路线 API
- 容器化部署，支持多人同时访问

**现场演示动线**：① 绑定账号 → ② 一句登记 → ③ 点开明细 → ④ 统计并导出 → ⑤ 活动中心

> 一句话总结：模型负责"听懂 + 编排"，代码负责"计算 + 落盘 + 校验"。

**谢谢，欢迎提问**
