# -*- coding: utf-8 -*-
"""
校园志愿时长汇总助手（路径B：Streamlit + 通义千问 Tool Calling）
====================================================================
- 通过 OpenAI 兼容接口调用阿里云百炼 DashScope 的通义千问模型
- 自行设计 5 个业务工具，由模型真实发起 Tool Call（非 if 关键词判断）
  1. query_volunteer_data    按起止日期（+可选项目名）查询志愿记录
  2. calc_total_hours        按姓名分组统计累计志愿总时长
  3. export_volunteer_excel  导出汇总 Excel（网页内置下载按钮）
  4. query_open_activities   查询校内活动与上海全市校外活动（含详情与报名链接）
  5. add_volunteer_record    为已绑定的个人账号录入新的志愿活动与时长
- 加分项：个人账号绑定 / 学校绑定 / 日志分区展示 / 表单新增记录 / 强异常容错 / 自动柱状图
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
    "工具会自动使用最近一次查询或统计的结果。"
    "查询记录时注意口径：用户说“我/我的/本人”或没有说明范围时，scope 传 me，只查当前登录账号本人的记录；"
    "用户明确说“每个人/全员/全班/所有同学/大家”时，scope 传 all，查询全部成员记录。"
    "若按本人口径查询，回答时说明“这是你本人的记录”，并提示如需全员记录可以直说。"
    "若上一轮是按本人口径查询、本轮用户却要统计或导出全员，请先用 scope=all 重新查询一次再统计。"
    "当用户询问最近有什么志愿活动、校内活动、校外/上海全市活动时，调用 query_open_activities；"
    "该工具支持按 scope 区分校内/校外、按 district 限定上海某个区、按 keyword 过滤；"
    "若该工具返回未绑定学校，请提示用户在页面左侧「学校绑定」中选择并绑定学校。"
    "回答活动类问题时，请按日期或距离列出要点，说明主办方、时间、时长、剩余名额与所在区，"
    "并在末尾提示用户可在活动卡片下方展开查看详情、通过报名链接前往官方平台报名。"
    "当用户要登记自己新参加的志愿活动（如'帮我记一笔：社区清洁 3 小时'）时，调用 add_volunteer_record，"
    "姓名与学号由绑定的个人账号自动带入，不要向用户索要姓名学号；缺少活动名称或时长时先追问。"
    "若该工具返回未绑定账号，请提示用户在左侧「个人账号」中绑定或新建账号。"
    "回答请简洁、使用中文。"
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

# ============================ 个人账号 ============================
# 可绑定的演示账号（与 DEFAULT_DATA 中的志愿者对应），也支持自己新建账号
DEMO_ACCOUNTS = [
    {"姓名": "张三", "学号": "2023001", "学院": "信息科学与技术学院", "专业": "计算机科学与技术 2301 班"},
    {"姓名": "李四", "学号": "2023002", "学院": "管理学院", "专业": "工商管理 2302 班"},
    {"姓名": "王五", "学号": "2023003", "学院": "外国语学院", "专业": "英语 2301 班"},
    {"姓名": "赵六", "学号": "2023004", "学院": "护理学院", "专业": "护理学 2303 班"},
]

# ============================ 学校与志愿活动数据 ============================
# 可绑定学校（真实校区与坐标，用于计算校外活动距离）
SCHOOLS = [
    {"名称": "上海杉达学院（金海校区）", "城市": "上海", "校区": "金海校区",
     "地址": "浦东新区金海路 2727 号", "lat": 31.268682, "lng": 121.659202,
     "官网": "https://www.sandau.edu.cn", "查号": "021-50210894"},
    {"名称": "上海杉达学院（嘉善校区）", "城市": "嘉兴", "校区": "嘉善校区",
     "地址": "浙江省嘉善县人民大道 505 号", "lat": 30.828400, "lng": 120.928300,
     "官网": "https://www.sandau.edu.cn", "查号": "0573-84239202"},
    {"名称": "复旦大学（邯郸校区）", "城市": "上海", "校区": "邯郸校区",
     "地址": "杨浦区邯郸路 220 号", "lat": 31.298900, "lng": 121.503300,
     "官网": "https://www.fudan.edu.cn", "查号": "021-65642222"},
    {"名称": "上海交通大学（闵行校区）", "城市": "上海", "校区": "闵行校区",
     "地址": "闵行区东川路 800 号", "lat": 31.024300, "lng": 121.433800,
     "官网": "https://www.sjtu.edu.cn", "查号": "021-54740000"},
    {"名称": "同济大学（四平路校区）", "城市": "上海", "校区": "四平路校区",
     "地址": "杨浦区四平路 1239 号", "lat": 31.282900, "lng": 121.506000,
     "官网": "https://www.tongji.edu.cn", "查号": "021-65982200"},
]

# 上海各区（校外活动区域筛选）
SHANGHAI_DISTRICTS = [
    "浦东新区", "黄浦区", "徐汇区", "长宁区", "静安区", "普陀区",
    "虹口区", "杨浦区", "闵行区", "宝山区", "嘉定区", "松江区", "青浦区", "奉贤区",
]

# 报名入口（官方平台）
LINK_SANDA = "https://www.sandau.edu.cn"          # 上海杉达学院官网（校团委通知/志愿招募）
LINK_SH_VOL = "https://www.volunteer.sh.cn"       # 上海志愿者网（全市志愿项目报名）
LINK_SSTM = "https://www.sstm.org.cn"             # 上海科技馆 / 天文馆志愿者
LINK_LIB = "https://www.library.sh.cn"            # 上海图书馆志愿者
LINK_ART = "https://www.artmuseumonline.org"      # 中华艺术宫
LINK_BLOOD = "https://www.sbc.org.cn"             # 上海市血液中心

# 校内活动（模拟“实时更新”：报名人数会随时间增长）
ON_CAMPUS = [
    {"id": "C101", "标题": "图书馆书库整理与导读", "主办": "杉达学院图书馆", "地点": "金海校区图书馆 3F 书库",
     "日期": "2026-10-29", "时间段": "14:00-17:00", "时长": 3.0, "名额": 25, "已报名": 16,
     "标签": ["图书管理", "室内"], "更新时间": 4, "报名链接": LINK_SANDA,
     "详情": {"介绍": "协助馆员完成书库错架整理、新书上架与阅览区导读，服务对象为全校师生。",
              "招募对象": "在校生，专业不限；需能连续站立 2 小时以上",
              "集合": "图书馆 1F 服务台，提前 10 分钟签到并领取志愿马甲",
              "联系人": "图书馆读者服务部 王老师（021-20262610）",
              "注意": "馆内保持安静；贵重物品随身携带；服务结束后现场签退认定时长"}},
    {"id": "C102", "标题": "校运会赛事保障志愿者", "主办": "校体育运动委员会", "地点": "金海校区东区田径场",
     "日期": "2026-10-31", "时间段": "08:00-12:00", "时长": 4.0, "名额": 40, "已报名": 34,
     "标签": ["大型活动", "户外"], "更新时间": 9, "报名链接": LINK_SANDA,
     "详情": {"介绍": "负责检录引导、器材搬运、赛道秩序维护与终点搀扶，保障校田径运动会顺利进行。",
              "招募对象": "在校生，身体健康；有急救证者优先",
              "集合": "东区田径场主席台西侧志愿者服务点",
              "联系人": "校体委办公室 李老师（021-20262610 转 3）",
              "注意": "穿运动鞋、自备防晒；雨天照常举行，请携带雨具"}},
    {"id": "C103", "标题": "校园开放日招生咨询引导", "主办": "招生办公室", "地点": "金海校区南门报到大厅",
     "日期": "2026-11-07", "时间段": "09:00-16:00", "时长": 6.0, "名额": 30, "已报名": 12,
     "标签": ["引导", "对外接待"], "更新时间": 21, "报名链接": LINK_SANDA,
     "详情": {"介绍": "为前来参加校园开放日的高中生及家长提供路线引导、资料发放与常见问题解答。",
              "招募对象": "在校生，普通话标准，熟悉校园布局",
              "集合": "南门报到大厅志愿者签到处",
              "联系人": "招生办 张老师（021-50210897）",
              "注意": "提前参加 11-06 晚 30 分钟线上培训；统一着校徽志愿者服"}},
    {"id": "C104", "标题": "无偿献血现场服务", "主办": "校医院 · 上海市血液中心", "地点": "大学生活动中心 1 号厅",
     "日期": "2026-11-03", "时间段": "10:00-16:00", "时长": 6.0, "名额": 30, "已报名": 13,
     "标签": ["公益", "医疗协助"], "更新时间": 33, "报名链接": LINK_BLOOD,
     "详情": {"介绍": "配合血液中心医护人员完成登记填表、排队引导、糖水发放与献血后观察陪护。",
              "招募对象": "在校生，无传染性疾病；本人不参与献血",
              "集合": "大学生活动中心 1 号厅入口",
              "联系人": "校医院保健科（021-20262610 转 6）",
              "注意": "服务期间佩戴口罩；不得代替医护人员做任何医学判断"}},
    {"id": "C105", "标题": "“书香杉达”读书月布展", "主办": "校图书馆 · 学生会", "地点": "图书馆 1F 中庭",
     "日期": "2026-11-05", "时间段": "13:00-17:00", "时长": 4.0, "名额": 15, "已报名": 5,
     "标签": ["布展", "文化活动"], "更新时间": 47, "报名链接": LINK_SANDA,
     "详情": {"介绍": "参与读书月主题展板搭建、书目推荐卡制作与现场秩序维护。",
              "招募对象": "在校生，有手工/设计经验者优先",
              "集合": "图书馆 1F 中庭",
              "联系人": "图书馆办公室 陈老师",
              "注意": "使用美工刀等工具时注意安全；展后共同撤展"}},
    {"id": "C106", "标题": "校园垃圾分类督导", "主办": "后勤保卫处", "地点": "学生公寓 3 号楼前投放点",
     "日期": "2026-11-02", "时间段": "07:30-09:30", "时长": 2.0, "名额": 18, "已报名": 7,
     "标签": ["环保", "校园治理"], "更新时间": 62, "报名链接": LINK_SANDA,
     "详情": {"介绍": "在早间投放高峰时段引导同学正确分类，记录混投情况并反馈后勤。",
              "招募对象": "在校生，住在金海校区者优先",
              "集合": "3 号楼前垃圾分类投放点",
              "联系人": "后勤保卫处 环卫科",
              "注意": "佩戴手套与袖标；遇到不配合情况只做劝导，不与同学争执"}},
]

# 校外活动（覆盖上海全市各区，带经纬度，按绑定学校计算距离）
OFF_CAMPUS = [
    {"id": "O201", "标题": "地铁 12 号线金海路站客流引导", "主办": "上海地铁志愿服务总队",
     "地址": "浦东新区 · 金海路站站厅", "区域": "浦东新区",
     "日期": "2026-10-28", "时间段": "16:30-19:30", "时长": 3.0, "名额": 20, "已报名": 12,
     "标签": ["交通引导", "城市服务"], "lat": 31.262000, "lng": 121.635000, "更新时间": 3,
     "报名链接": LINK_SH_VOL,
     "详情": {"介绍": "晚高峰在闸机口与换乘通道引导客流、协助老人与携带大件行李乘客进出站。",
              "招募对象": "18 岁以上，能听懂普通话与基本上海话",
              "集合": "金海路站站长室（12 号线站厅层）",
              "联系人": "地铁志愿总队 值班站长",
              "注意": "不得进入轨行区；站内禁用自拍杆；需提前 15 分钟到岗接受安全交底"}},
    {"id": "O202", "标题": "上海科技馆展厅科普引导", "主办": "上海科技馆志愿者服务基地",
     "地址": "浦东新区 · 世纪大道 2000 号", "区域": "浦东新区",
     "日期": "2026-11-01", "时间段": "09:00-17:00", "时长": 8.0, "名额": 30, "已报名": 22,
     "标签": ["科普", "场馆服务"], "lat": 31.217600, "lng": 121.541900, "更新时间": 11,
     "报名链接": LINK_SSTM,
     "详情": {"介绍": "在常设展区提供路线指引、展项原理解说与秩序维护，是科技馆长期运行的品牌志愿岗。",
              "招募对象": "在校大学生，理工科背景优先；需通过馆方线上岗前测试",
              "集合": "科技馆 1 号门志愿者之家",
              "联系人": "科技馆志愿者基地（021-68622000）",
              "注意": "全天岗，午餐由馆方提供；须穿平底鞋，馆内不可饮食"}},
    {"id": "O203", "标题": "上海图书馆东馆“书香伴读”", "主办": "上海图书馆（东馆）",
     "地址": "浦东新区 · 合欢路 300 号", "区域": "浦东新区",
     "日期": "2026-11-04", "时间段": "13:30-17:30", "时长": 4.0, "名额": 15, "已报名": 9,
     "标签": ["伴读", "图书管理"], "lat": 31.222600, "lng": 121.550300, "更新时间": 18,
     "报名链接": LINK_LIB,
     "详情": {"介绍": "在少儿阅览区陪伴小读者阅读、整理归还图书并协助自助借还机使用。",
              "招募对象": "在校生，喜爱与儿童沟通",
              "集合": "东馆 1F 少儿区服务台",
              "联系人": "上图读者服务中心",
              "注意": "需保持安静；不与小读者单独离开阅览区"}},
    {"id": "O204", "标题": "中华艺术宫观众服务", "主办": "中华艺术宫（上海美术馆）",
     "地址": "浦东新区 · 上南路 205 号", "区域": "浦东新区",
     "日期": "2026-11-06", "时间段": "10:00-16:00", "时长": 6.0, "名额": 20, "已报名": 7,
     "标签": ["美术", "讲解"], "lat": 31.183700, "lng": 121.489600, "更新时间": 26,
     "报名链接": LINK_ART,
     "详情": {"介绍": "为参观者提供导览路线咨询、重点展品讲解与文明观展提醒。",
              "招募对象": "在校生，对美术史有兴趣；需参加 1 次线上培训",
              "集合": "艺术宫 0 米层观众服务中心",
              "联系人": "艺术宫公共教育部",
              "注意": "展厅内禁止触摸展品；讲解口径以馆方提供文稿为准"}},
    {"id": "O205", "标题": "徐汇滨江“文明游园”劝导", "主办": "徐汇区新时代文明实践中心",
     "地址": "徐汇区 · 龙腾大道西岸沿线", "区域": "徐汇区",
     "日期": "2026-11-08", "时间段": "14:00-18:00", "时长": 4.0, "名额": 25, "已报名": 10,
     "标签": ["环保", "城市治理"], "lat": 31.188000, "lng": 121.459000, "更新时间": 35,
     "报名链接": LINK_SH_VOL,
     "详情": {"介绍": "沿滨江步道劝导遛狗牵绳、垃圾分类投放，并为游客提供问路指引。",
              "招募对象": "在校生，能连续步行 3 公里以上",
              "集合": "西岸艺术中心 A 馆前广场志愿驿站",
              "联系人": "徐汇区志愿服务中心",
              "注意": "统一发放马甲与垃圾夹；劝导以礼为先，不强制"}},
    {"id": "O206", "标题": "南京东路步行街文明引导", "主办": "黄浦区志愿者协会",
     "地址": "黄浦区 · 南京东路世纪广场", "区域": "黄浦区",
     "日期": "2026-11-09", "时间段": "15:00-19:00", "时长": 4.0, "名额": 30, "已报名": 18,
     "标签": ["交通引导", "大客流"], "lat": 31.237000, "lng": 121.475000, "更新时间": 41,
     "报名链接": LINK_SH_VOL,
     "详情": {"介绍": "节假日大客流时段在步行街值守，提供问路咨询、协助疏导人流与应急联络。",
              "招募对象": "在校生，表达清晰，能应对大客流环境",
              "集合": "世纪广场志愿服务亭",
              "联系人": "黄浦区志愿者协会",
              "注意": "严禁与游客发生冲突；发现异常情况立即报告现场民警"}},
    {"id": "O207", "标题": "静安寺街道“为老送餐”", "主办": "静安寺街道社区志愿服务中心",
     "地址": "静安区 · 愚园路街区", "区域": "静安区",
     "日期": "2026-11-10", "时间段": "10:30-13:30", "时长": 3.0, "名额": 12, "已报名": 6,
     "标签": ["助老", "社区"], "lat": 31.225000, "lng": 121.445000, "更新时间": 52,
     "报名链接": LINK_SH_VOL,
     "详情": {"介绍": "为社区高龄独居老人配送午餐，并做简单的问候与居家安全观察记录。",
              "招募对象": "在校生，耐心细致，会讲上海话者优先",
              "集合": "静安寺街道社区食堂",
              "联系人": "街道为老服务中心",
              "注意": "不进入老人卧室；不代收钱款；异常情况当日反馈社工"}},
    {"id": "O208", "标题": "杨浦滨江人民城市建设展示馆讲解", "主办": "杨浦区滨江志愿服务队",
     "地址": "杨浦区 · 安浦路 461 号", "区域": "杨浦区",
     "日期": "2026-11-11", "时间段": "09:30-14:30", "时长": 5.0, "名额": 10, "已报名": 4,
     "标签": ["讲解", "城市历史"], "lat": 31.262000, "lng": 121.538000, "更新时间": 58,
     "报名链接": LINK_SH_VOL,
     "详情": {"介绍": "为参观团队提供展馆讲解，讲述杨浦滨江从工业锈带到生活秀带的转型历程。",
              "招募对象": "在校生，普通话二级甲等以上优先；需背熟讲解稿",
              "集合": "展示馆 1F 序厅",
              "联系人": "滨江志愿服务队",
              "注意": "讲解稿以馆方最新版本为准；每场讲解不超过 40 分钟"}},
    {"id": "O209", "标题": "上海市血液中心献血屋服务", "主办": "上海市血液中心",
     "地址": "长宁区 · 虹桥路 1191 号", "区域": "长宁区",
     "日期": "2026-11-12", "时间段": "09:00-13:00", "时长": 4.0, "名额": 15, "已报名": 8,
     "标签": ["公益", "医疗协助"], "lat": 31.199000, "lng": 121.408000, "更新时间": 66,
     "报名链接": LINK_BLOOD,
     "详情": {"介绍": "在献血屋协助登记、引导体检流程、发放纪念品与献血后观察陪护。",
              "招募对象": "在校生，身体健康；有志愿服务经历者优先",
              "集合": "血液中心献血屋前台",
              "联系人": "血液中心社会服务部",
              "注意": "不接触血液标本；严格保护献血者隐私，不得拍照"}},
    {"id": "O210", "标题": "顾村公园“赏樱季”文明引导", "主办": "宝山区绿化市容局志愿服务队",
     "地址": "宝山区 · 沪太路 4788 号", "区域": "宝山区",
     "日期": "2026-11-14", "时间段": "08:30-14:30", "时长": 6.0, "名额": 40, "已报名": 26,
     "标签": ["环保", "大客流"], "lat": 31.348000, "lng": 121.370000, "更新时间": 72,
     "报名链接": LINK_SH_VOL,
     "详情": {"介绍": "在主要赏花点位维持游园秩序、劝导不文明行为、协助走失儿童求助。",
              "招募对象": "在校生，能适应户外长时间站立",
              "集合": "顾村公园 1 号门游客中心",
              "联系人": "宝山区绿化市容局志愿队",
              "注意": "园内禁用扩音器；垃圾随身带走；雨天活动顺延"}},
    {"id": "O211", "标题": "松江大学城“爱心寒托班”助教", "主办": "松江区团委 · 社区少工委",
     "地址": "松江区 · 文汇路街道社区学校", "区域": "松江区",
     "日期": "2026-11-16", "时间段": "08:30-16:30", "时长": 8.0, "名额": 18, "已报名": 5,
     "标签": ["支教", "青少年"], "lat": 31.038000, "lng": 121.220000, "更新时间": 80,
     "报名链接": LINK_SH_VOL,
     "详情": {"介绍": "协助班主任管理课堂纪律、辅导作业、组织手工与体育活动。",
              "招募对象": "在校生，有家教或带班经验者优先；需提供健康承诺书",
              "集合": "文汇路街道社区学校一楼大厅",
              "联系人": "松江区团委学校部",
              "注意": "全程两名志愿者在场；不得单独带学生离开教学区"}},
    {"id": "O212", "标题": "普陀“半马苏河”驿站便民服务", "主办": "普陀区长寿路街道志愿服务中心",
     "地址": "普陀区 · 苏州河长寿段驿站", "区域": "普陀区",
     "日期": "2026-11-15", "时间段": "09:00-12:00", "时长": 3.0, "名额": 14, "已报名": 3,
     "标签": ["便民", "社区"], "lat": 31.240000, "lng": 121.425000, "更新时间": 88,
     "报名链接": LINK_SH_VOL,
     "详情": {"介绍": "在滨河驿站提供饮水、充电、应急药品借用与路线咨询服务。",
              "招募对象": "在校生，性格开朗，能使用基础办公设备",
              "集合": "半马苏河驿站（长寿路段）",
              "联系人": "长寿路街道志愿服务中心",
              "注意": "药品仅登记发放，不做用药建议；每日盘点物资"}},
]

# 定时注入的新活动（模拟推送到达）
NEW_POOL = [
    {"kind": "on", "id": "C107", "标题": "新生心理健康讲座协助", "主办": "心理健康教育中心", "地点": "金海校区学术报告厅",
     "日期": "2026-11-06", "时间段": "18:30-21:00", "时长": 2.5, "名额": 8, "已报名": 1,
     "标签": ["讲座协助", "室内"], "更新时间": 1, "报名链接": LINK_SANDA,
     "详情": {"介绍": "负责入场签到、座位引导与讲座后问卷回收。", "招募对象": "在校生",
              "集合": "学术报告厅前厅", "联系人": "心理健康教育中心", "注意": "讲座内容涉及同学隐私，不得录音录像"}},
    {"kind": "off", "id": "O213", "标题": "虹口鲁迅公园“无障碍陪伴”", "主办": "虹口区残疾人联合会",
     "地址": "虹口区 · 四川北路 2288 号", "区域": "虹口区",
     "日期": "2026-11-13", "时间段": "13:30-16:30", "时长": 3.0, "名额": 12, "已报名": 9,
     "标签": ["助残", "公园"], "lat": 31.273000, "lng": 121.478000, "更新时间": 1,
     "报名链接": LINK_SH_VOL,
     "详情": {"介绍": "陪同视障朋友游园，提供口述影像与路线引导。", "招募对象": "在校生，有耐心",
              "集合": "鲁迅公园南门", "联系人": "虹口区残联", "注意": "先询问再搀扶；不擅自替对方做决定"}},
    {"kind": "on", "id": "C108", "标题": "校园反诈宣传摊位", "主办": "保卫处 · 学生会", "地点": "学生食堂西门广场",
     "日期": "2026-11-08", "时间段": "11:00-13:00", "时长": 2.0, "名额": 10, "已报名": 2,
     "标签": ["宣传", "安全"], "更新时间": 1, "报名链接": LINK_SANDA,
     "详情": {"介绍": "发放反诈手册、指导同学安装国家反诈中心 APP。", "招募对象": "在校生",
              "集合": "食堂西门广场", "联系人": "保卫处", "注意": "不代同学操作手机银行"}},
    {"kind": "off", "id": "O214", "标题": "上海天文馆星空讲解志愿岗", "主办": "上海天文馆",
     "地址": "浦东新区 · 临港大道 380 号", "区域": "浦东新区",
     "日期": "2026-11-18", "时间段": "09:00-17:00", "时长": 8.0, "名额": 20, "已报名": 16,
     "标签": ["科普", "讲解"], "lat": 30.910400, "lng": 121.925000, "更新时间": 1,
     "报名链接": LINK_SSTM,
     "详情": {"介绍": "在家园、宇宙、征程三大展区提供展项讲解与观展引导。", "招募对象": "在校生，天文爱好者优先",
              "集合": "天文馆 1 号门志愿者之家", "联系人": "天文馆志愿者基地", "注意": "需提前完成馆方线上培训并考核"}},
]



def haversine_km(lat1, lng1, lat2, lng2) -> float:
    """两坐标间球面距离（公里），用于计算校外活动与学校的距离"""
    from math import radians, sin, cos, asin, sqrt

    lat1, lng1, lat2, lng2 = map(radians, [lat1, lng1, lat2, lng2])
    dlat, dlng = lat2 - lat1, lng2 - lng1
    a = sin(dlat / 2) ** 2 + cos(lat1) * cos(lat2) * sin(dlng / 2) ** 2
    return round(2 * 6371 * asin(sqrt(a)), 1)


def tick_feed(force: bool = False):
    """模拟活动实时更新：随时间推进报名人数、刷新“更新于 N 分钟”，并定时注入新活动"""
    import random
    from datetime import datetime as _dt

    now = _dt.now()
    last = st.session_state.get("feed_tick_at")
    # 距上次更新不足 25 秒且非手动刷新，则跳过（避免频繁跳动）
    if not force and last and (now - last).total_seconds() < 25:
        return
    st.session_state.feed_tick_at = now

    for key in ("activities_on", "activities_off"):
        for a in st.session_state[key]:
            a["更新时间"] = a.get("更新时间", 0) + random.randint(0, 3)
            if a["已报名"] < a["名额"] and random.random() < 0.55:
                a["已报名"] = min(a["名额"], a["已报名"] + random.randint(1, 2))

    # 每累计两次轮询注入一条新活动，模拟推送到达
    st.session_state.feed_polls = st.session_state.get("feed_polls", 0) + 1
    if st.session_state.feed_polls % 2 == 0:
        used = st.session_state.get("feed_pool_used", 0)
        if used < len(NEW_POOL):
            item = dict(NEW_POOL[used])
            item["更新时间"] = 1
            st.session_state["activities_" + ("on" if item["kind"] == "on" else "off")].insert(0, item)
            st.session_state.feed_pool_used = used + 1
            st.session_state.new_push = item["标题"]

TOOL_ICON = {
    "query_volunteer_data": "🔍",
    "calc_total_hours": "🧮",
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
        st.session_state.volunteer_df = pd.DataFrame(DEFAULT_DATA)
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
    # —— 学校绑定与活动推送 ——
    if "bound_school" not in st.session_state:
        # 默认绑定上海杉达学院（金海校区），打开即可看到校内活动与上海全市推送
        st.session_state.bound_school = dict(SCHOOLS[0])
    if "activities_on" not in st.session_state:
        st.session_state.activities_on = [dict(a) for a in ON_CAMPUS]
    if "activities_off" not in st.session_state:
        st.session_state.activities_off = [dict(a) for a in OFF_CAMPUS]
    if "feed_pool_used" not in st.session_state:
        st.session_state.feed_pool_used = 0      # 已注入的新活动数量
    if "feed_tick_at" not in st.session_state:
        st.session_state.feed_tick_at = None     # 上次更新时间
    if "feed_polls" not in st.session_state:
        st.session_state.feed_polls = 0
    if "push_district" not in st.session_state:
        st.session_state.push_district = "全部区域"   # 校外活动区域筛选
    if "push_keyword" not in st.session_state:
        st.session_state.push_keyword = ""           # 校外活动关键字
    if "new_push" not in st.session_state:
        st.session_state.new_push = None         # 新推送提示


# ============================ 工具实现 ============================
def _parse_date(text):
    """严格校验日期格式，非法返回 None（对应容错设计）"""
    try:
        return datetime.strptime(str(text).strip(), "%Y-%m-%d").date()
    except (ValueError, TypeError):
        return None


def query_volunteer_data(start_date: str, end_date: str, project_name: str = "",
                         scope: str = "me") -> dict:
    """工具1：按起止日期查询志愿记录。
    scope=me 只查当前绑定账号本人的记录（默认，与页面记录口径保持一致）；
    scope=all 查全部成员记录。project_name 为可选的项目名称筛选关键字。"""
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
    result = df[mask].drop(columns=["日期date"])

    if result.empty:
        tip = f"{start_date} 至 {end_date}" + (f"（项目含“{project_name}”）" if project_name else "")
        return {
            "status": "empty",
            "scope": "all" if want_all else "me",
            "owner": owner_label,
            "message": f"{owner_label}在 {tip} 时间范围内没有志愿记录。",
            "records": [],
        }

    return {
        "status": "success",
        "scope": "all" if want_all else "me",
        "owner": owner_label,
        "count": int(len(result)),
        "message": f"{owner_label}共查询到 {len(result)} 条志愿记录。",
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


def query_open_activities(scope: str = "all", district: str = "", keyword: str = "") -> dict:
    """工具4：查询校内正在招募的志愿活动，以及上海全市的校外活动（按距离由近到远）"""
    school = st.session_state.get("bound_school")
    if not school:
        return {
            "status": "error",
            "message": "尚未绑定学校。请先在左侧「学校绑定」中选择学校并点击绑定，再查询活动。",
        }

    kw = str(keyword or "").strip()
    scope_raw = str(scope or "all").lower()
    if scope_raw in ("on_campus", "on", "校内", "campus"):
        want_on, want_off = True, False
    elif scope_raw in ("off_campus", "off", "校外", "nearby", "上海", "shanghai"):
        want_on, want_off = False, True
    else:
        want_on, want_off = True, True

    # 区域：空 / "all" / "全部" 视为不限（覆盖上海全市）
    dist_raw = str(district or "").strip()
    district_name = "" if dist_raw.lower() in ("", "all", "全部", "全部区域", "上海市", "上海") else dist_raw
    if district_name and not district_name.endswith("区"):
        district_name += "区"
    if district_name and district_name not in SHANGHAI_DISTRICTS:
        return {
            "status": "error",
            "message": f"未识别区域「{dist_raw}」。上海市可选区域：{'、'.join(SHANGHAI_DISTRICTS)}；"
                       "也可以不传 district，默认推送上海全市的活动。",
        }

    def _match(act):
        if not kw:
            return True
        blob = " ".join(
            [str(act.get("标题", "")), str(act.get("主办", "")), " ".join(act.get("标签", []))]
        )
        return kw in blob

    on_list, off_list = [], []
    if want_on:
        for a in st.session_state.activities_on:
            if _match(a):
                on_list.append({
                    "标题": a["标题"], "主办": a["主办"], "地点": a["地点"],
                    "日期": a["日期"], "时间段": a["时间段"], "时长(小时)": a["时长"],
                    "名额": a["名额"], "已报名": a["已报名"],
                    "剩余名额": max(0, a["名额"] - a["已报名"]), "标签": a["标签"],
                    "报名链接": a.get("报名链接", LINK_SANDA),
                })
        on_list.sort(key=lambda x: x["日期"])
    if want_off:
        for a in st.session_state.activities_off:
            if district_name and a.get("区域") != district_name:
                continue
            dist = haversine_km(school["lat"], school["lng"], a["lat"], a["lng"])
            if _match(a):
                off_list.append({
                    "标题": a["标题"], "主办": a["主办"], "地址": a["地址"],
                    "区域": a.get("区域", ""), "距离km": dist, "日期": a["日期"],
                    "时间段": a["时间段"], "时长(小时)": a["时长"],
                    "名额": a["名额"], "已报名": a["已报名"],
                    "剩余名额": max(0, a["名额"] - a["已报名"]),
                    "标签": a["标签"], "报名链接": a.get("报名链接", LINK_SH_VOL),
                })
        off_list.sort(key=lambda x: x["距离km"])

    area_txt = district_name or "上海全市"
    if not on_list and not off_list:
        return {
            "status": "empty",
            "message": f"在「{school['名称']}」的{area_txt}范围内没有找到符合条件的活动"
                       f"（关键词：{kw or '无'}）。可以换关键词或改为不限区域再试。",
        }

    cover = len({a["区域"] for a in off_list})
    parts = []
    if want_on:
        parts.append(f"校内 {len(on_list)} 个活动")
    if want_off:
        seg = f"{area_txt}校外 {len(off_list)} 个活动"
        if district_name == "" and off_list:
            seg += f"，覆盖 {cover} 个区"
        parts.append(seg)
    msg = "；".join(parts) + "。"
    if want_off and off_list:
        msg += "校外已按距校区的距离由近到远排序，每个活动都附有详情与官方报名链接。"
    return {
        "status": "success",
        "school": {"名称": school["名称"], "城市": school["城市"], "校区": school["校区"]},
        "area": area_txt,
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
                         place: str = "", note: str = "") -> dict:
    """工具5：为当前绑定的个人账号录入一条新的志愿记录（只需活动名称与时长）"""
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

    new_row = {
        "姓名": acc["姓名"],
        "学号": acc["学号"],
        "志愿项目": proj,
        "日期": date_str,
        "时长(小时)": h,
    }
    st.session_state.volunteer_df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)

    mine = st.session_state.volunteer_df[st.session_state.volunteer_df["姓名"] == acc["姓名"]]
    total = float(mine["时长(小时)"].sum())
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
        "message": (f"已为 {acc['姓名']}（学号 {acc['学号']}）录入：{date_str} {proj} {h:g} 小时"
                    + (f"（{extra}）" if extra else "")
                    + f"。该同学累计志愿 {total:.1f} 小时，共 {len(mine)} 条记录。"),
    }


TOOL_IMPL = {
    "query_volunteer_data": query_volunteer_data,
    "calc_total_hours": calc_total_hours,
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
            "description": "按起止日期查询校园志愿记录；可通过可选参数 project_name 按志愿项目名称关键字筛选，通过 scope 指定查本人还是全员。",
            "parameters": {
                "type": "object",
                "properties": {
                    "start_date": {"type": "string", "description": "起始日期，格式 YYYY-MM-DD"},
                    "end_date": {"type": "string", "description": "结束日期，格式 YYYY-MM-DD"},
                    "project_name": {"type": "string", "description": "可选。志愿项目名称关键字，不需要时省略"},
                    "scope": {
                        "type": "string",
                        "enum": ["me", "all"],
                        "description": "查询口径：me 只查当前登录账号本人的记录（默认）；all 查全部成员记录。用户说“每个人/全员/全班/所有同学”时传 all",
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
    {
        "type": "function",
        "function": {
            "name": "query_open_activities",
            "description": "查询志愿活动：绑定学校的校内活动列表，以及上海全市的校外志愿活动（按距校区距离由近到远，含详细地址、所属区、报名链接）。需要用户已绑定学校；未绑定时提示先去绑定。",
            "parameters": {
                "type": "object",
                "properties": {
                    "scope": {
                        "type": "string",
                        "enum": ["all", "on_campus", "off_campus"],
                        "description": "查询范围：all 全部 / on_campus 仅校内 / off_campus 仅校外（上海全市）",
                    },
                    "district": {
                        "type": "string",
                        "description": "可选。限定上海某个区，如 浦东新区、徐汇区；不传则返回上海全市的校外活动",
                    },
                    "keyword": {
                        "type": "string",
                        "description": "可选。按标题、主办方或标签过滤，例如 科普、助老、环保",
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
  .rec-card { cursor: pointer; transition: box-shadow .2s ease, transform .2s ease; }
  .rec-card:hover { box-shadow: 0 6px 18px rgba(43,108,176,.18); transform: translateY(-2px); }
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
            <span class="badge">🛠️ 5 个业务工具</span>
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
    """顶部 KPI：只保留两张卡，始终跟随当前绑定账号的本人记录——
    录入新记录、切换账号后，这里的数字都会即时更新；不受「显示全部成员」开关影响"""
    mine = my_records_df()
    acc = get_bound_account()
    n = len(mine)
    hrs = float(mine["时长(小时)"].sum()) if n else 0.0

    rec_label = f"{acc['姓名']}的志愿记录" if acc else "志愿记录总数"
    rec_hint = (
        '<div class="stat-hint">👁 点击本卡片，查看我的每一项志愿记录</div>' if acc
        else '<div class="stat-hint" style="color:#A0AEC0;">绑定账号后这里会显示本人的记录</div>'
    )
    hrs_hint = (
        f'<div class="stat-hint" style="color:#DD6B20;">共 {n} 条记录累计 · 随录入实时更新</div>' if acc
        else '<div class="stat-hint" style="color:#A0AEC0;">绑定账号后显示本人的累计时长</div>'
    )

    cols = st.columns(2, gap="medium")
    with cols[0]:
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
        clicked = st.button(
            "🔍 查看我的志愿记录明细", key="open_my_records", use_container_width=True,
        )
    with cols[1]:
        st.markdown(
            f"""
            <div class="stat-card">
              <div class="stat-top">
                <span class="stat-icon" style="background:#FFFAF0;">⏳</span>
                <span class="stat-label">{"我的累计志愿时长" if acc else "累计志愿时长"}</span>
              </div>
              <div class="stat-value">{hrs:.1f} <span class="stat-unit">小时</span></div>
              {hrs_hint}
            </div>
            """,
            unsafe_allow_html=True,
        )

    if clicked:  # 点卡片 = 展开本人记录明细（再点一次收起）
        st.session_state.show_my_records = not st.session_state.get("show_my_records")
        st.rerun()

    if not acc:
        st.caption("⚠️ 尚未绑定个人账号：在左侧「个人账号」选择或新建并绑定后，这两张卡片即显示本人数据。")
    elif st.session_state.get("show_all_rows"):
        st.caption("左侧数据表当前按「全员」口径展示；顶部两张卡片始终显示本人数据，与账号同步切换。")


def render_my_records_detail():
    """点开「志愿记录总数」卡片后的明细面板：本人每一项志愿记录逐条可查"""
    acc = get_bound_account()
    if not acc:
        return
    mine = my_records_df()
    _panel_title("📋", "#2B6CB0", f"{acc['姓名']}的志愿记录明细", "每一项记录 · 按日期由近到远")
    if mine.empty:
        st.markdown(
            '<div class="empty-box"><span class="big">📝</span>'
            "你名下还没有志愿记录<br>在左侧「录入我的志愿记录」或直接对话登记第一条吧</div>",
            unsafe_allow_html=True,
        )
    else:
        total = float(mine["时长(小时)"].sum())
        latest = mine.iloc[0]
        st.caption(
            f"共 {len(mine)} 条 · 合计 {total:.1f} 小时 · 最近一次："
            f"{latest['日期']} {latest['志愿项目']} {latest['时长(小时)']:g} 小时"
        )
        st.dataframe(mine, use_container_width=True, hide_index=True)
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


# ============================ 个人账号绑定 ============================
def account_stats(acc):
    """返回 (我的记录数, 我的累计时长)"""
    df = st.session_state.volunteer_df
    mine = df[df["姓名"] == acc["姓名"]]
    return len(mine), float(mine["时长(小时)"].sum()) if len(mine) else 0.0


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
    cnt, hrs = account_stats(acc)
    sub = " · ".join([x for x in [acc.get("学院", ""), acc.get("专业", "")] if x]) or "未填写院系信息"
    st.markdown(
        f"""
        <div class="acct-chip">
          <div class="ac-avatar">{acc["姓名"][:1]}</div>
          <div class="ac-main">
            <div class="ac-name">{acc["姓名"]}<span class="ac-id">学号 {acc["学号"]}</span></div>
            <div class="ac-sub">{sub}</div>
            <div class="ac-kpi">已录 <b>{cnt}</b> 条 · 累计 <b>{hrs:.1f}</b> 小时</div>
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

    cnt, hrs = account_stats(acc)
    df = st.session_state.volunteer_df
    avg = hrs / cnt if cnt else 0.0
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
            <div><b>{hrs:.1f}</b><span>我的累计时长</span></div>
            <div><b>{avg:.1f}</b><span>单次平均时长</span></div>
            <div><b>{len(df)}</b><span>全员记录</span></div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ============================ 学校绑定 / 活动中心 ============================
def render_school_bind():
    """侧栏：绑定、切换、解除学校"""
    st.markdown("**🏫 学校绑定**")
    names = [s["名称"] for s in SCHOOLS]
    cur = st.session_state.get("bound_school")
    idx = names.index(cur["名称"]) if cur and cur["名称"] in names else 0
    pick = st.selectbox("选择学校", names, index=idx, key="school_picker")

    c1, c2 = st.columns(2)
    if c1.button("绑定 / 切换", use_container_width=True, key="bind_btn"):
        school = next(s for s in SCHOOLS if s["名称"] == pick)
        st.session_state.bound_school = dict(school)
        st.session_state.toast = (
            "success",
            f"已绑定 {school['名称']}（{school['城市']} · {school['校区']}）",
        )
        st.rerun()
    if c2.button("解除绑定", use_container_width=True, key="unbind_btn"):
        st.session_state.bound_school = None
        st.session_state.toast = ("success", "已解除学校绑定。")
        st.rerun()

    if cur:
        addr = cur.get("地址", "")
        st.markdown(
            f'<div class="school-chip">✅ {cur["名称"]}'
            f'<span>{cur["城市"]} · {cur["校区"]}{" · " + addr if addr else ""} · {cur["lat"]:.3f}, {cur["lng"]:.3f}</span></div>',
            unsafe_allow_html=True,
        )
        if cur.get("官网"):
            st.caption(f"学校官网：{cur['官网']}")
    else:
        st.caption("⚠️ 未绑定：校内活动与校外推送不可用")


def render_school_strip():
    """顶部：学校绑定状态条"""
    cur = st.session_state.get("bound_school")
    if not cur:
        st.markdown(
            """
            <div class="school-strip unbound">
              <div class="ss-left">
                <span class="ss-label">未绑定学校</span>
                <span class="ss-name">绑定后可查看校内实时活动与上海全市校外推送</span>
                <span class="ss-sub">入口：左侧边栏「学校绑定」→ 选择学校 → 点击「绑定 / 切换」</span>
              </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        return

    on_n = len(st.session_state.activities_on)
    off_n = len(st.session_state.activities_off)
    dists = [
        haversine_km(cur["lat"], cur["lng"], a["lat"], a["lng"])
        for a in st.session_state.activities_off
    ]
    near = f"{min(dists):.1f} km" if dists else "—"
    areas = len({a.get("区域") for a in st.session_state.activities_off if a.get("区域")})
    st.markdown(
        f"""
        <div class="school-strip">
          <div class="ss-left">
            <span class="ss-label">已绑定学校</span>
            <span class="ss-name">{cur["名称"]}</span>
            <span class="ss-sub">{cur["城市"]} · {cur["校区"]} · {cur.get("地址", "")} · 坐标 {cur["lat"]:.4f}, {cur["lng"]:.4f}</span>
          </div>
          <div class="ss-right">
            <div><b>{on_n}</b><span>校内活动</span></div>
            <div><b>{off_n}</b><span>上海全市校外</span></div>
            <div><b>{areas}</b><span>覆盖区数</span></div>
            <div><b>{near}</b><span>最近校外活动</span></div>
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
    st.markdown(
        f"""
        <div class="act-card {kind}">
          <div class="act-head">
            <span class="act-title">{a["标题"]}</span>{new_html}{full_html}{region_html}
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


def render_activities():
    """活动中心：校内实时活动 + 上海全市校外推送（含详情与报名链接）"""
    _panel_title("🏫", "#0D9488", "志愿活动中心", "校内实时更新 · 上海全市校外推送 · 详情可展开 · 一键报名")
    cur = st.session_state.get("bound_school")
    if not cur:
        st.markdown(
            '<div class="empty-box"><span class="big">🏫</span>'
            "尚未绑定学校<br>请先在左侧边栏「学校绑定」中选择并绑定学校</div>",
            unsafe_allow_html=True,
        )
        return

    c1, c2, c3 = st.columns([0.9, 1.3, 2.0])
    auto = c2.toggle("自动刷新", value=True, key="feed_auto")
    if auto:
        tick_feed()  # 时间驱动的模拟实时更新（25 秒节流）
    if c1.button("🔄 刷新活动", use_container_width=True, key="refresh_feed"):
        tick_feed(force=True)
        st.session_state.toast = ("success", "活动列表已更新。")
        st.rerun()
    last = st.session_state.get("feed_tick_at")
    c3.caption(
        "最近更新：" + (last.strftime("%H:%M:%S") if last else "—") + "　（活动数据为模拟实时流）"
    )

    if st.session_state.get("new_push"):
        st.info(f"📬 新活动推送：{st.session_state.new_push}")
        st.session_state.new_push = None

    on_tab, off_tab = st.tabs(
        [
            f"🏫 校内活动（{len(st.session_state.activities_on)}）",
            f"🌆 上海全市校外（{len(st.session_state.activities_off)}）",
        ]
    )
    with on_tab:
        kw_on = st.text_input(
            "校内活动搜索", value="", key="act_kw_on",
            placeholder="输入关键字过滤，如 图书馆 / 运动会 / 环保",
        ).strip()
        items = sorted(st.session_state.activities_on, key=lambda x: x["日期"])
        if kw_on:
            items = [
                a for a in items
                if kw_on in (a["标题"] + a["主办"] + "".join(a.get("标签", [])))
            ]
            st.caption(f"匹配「{kw_on}」的校内活动 {len(items)} 个")
        if not items:
            st.markdown(
                '<div class="empty-box"><span class="big">🏫</span>'
                "没有匹配的校内活动<br>换个关键字试试</div>",
                unsafe_allow_html=True,
            )
        for a in items:
            render_activity_card(a, "on")
    with off_tab:
        c1, c2 = st.columns([1, 1.6])
        areas = ["全部区域"] + [
            d for d in SHANGHAI_DISTRICTS
            if any(a.get("区域") == d for a in st.session_state.activities_off)
        ]
        cur_area = st.session_state.get("push_district", "全部区域")
        area = c1.selectbox(
            "区域筛选（上海全市）", areas,
            index=areas.index(cur_area) if cur_area in areas else 0,
            key="area_picker",
        )
        st.session_state.push_district = area
        kw = c2.text_input(
            "校外活动搜索", value=st.session_state.get("push_keyword", ""),
            key="act_kw_off", placeholder="输入关键字过滤，如 科普 / 助老 / 环保",
        ).strip()
        st.session_state.push_keyword = kw

        rows = []
        for a in st.session_state.activities_off:
            if area != "全部区域" and a.get("区域") != area:
                continue
            if kw and kw not in (a["标题"] + a["主办"] + "".join(a.get("标签", []))):
                continue
            rows.append((haversine_km(cur["lat"], cur["lng"], a["lat"], a["lng"]), a))
        rows.sort(key=lambda x: x[0])

        if not rows:
            st.markdown(
                '<div class="empty-box"><span class="big">🌆</span>'
                "当前筛选条件下没有校外活动<br>可切换为「全部区域」或清空关键字</div>",
                unsafe_allow_html=True,
            )
        else:
            cover = len({a.get("区域") for _, a in rows})
            farthest = rows[-1][0]
            st.caption(
                f"{area}　共 {len(rows)} 个活动，覆盖 {cover} 个区，"
                f"距校区 {rows[0][0]:.1f} ~ {farthest:.1f} km，按距离由近到远排列"
            )
        for d, a in rows:
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
                )
                add_log("add_volunteer_record",
                        {"project_name": project, "date": d.strftime("%Y-%m-%d"), "hours": hours},
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
            st.session_state.volunteer_df = pd.DataFrame(DEFAULT_DATA)
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
        '<span class="chip">🕘 查询我10月的记录</span>'
        '<span class="chip">📊 统计每个人的总时长</span>'
        '<span class="chip">📥 导出汇总Excel</span>',
        unsafe_allow_html=True,
    )
    c1, c2, c3 = st.columns(3)
    if c1.button("🕘 查询我的记录", use_container_width=True):
        st.session_state.pending_input = "帮我查询我 2026-10-01 到 2026-10-31 的志愿记录"
    if c2.button("📊 统计总时长", use_container_width=True):
        st.session_state.pending_input = "统计 2026-09-01 到 2026-10-31 每个人（全员）的累计志愿总时长"
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
