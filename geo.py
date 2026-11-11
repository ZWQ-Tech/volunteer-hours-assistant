# -*- coding: utf-8 -*-
"""地区、高校与志愿活动数据层。

- regions.json        34 个省级单位 / 368 个城市 / 3078 个区县（含城市中心坐标）
- universities.json   全国 1308 所本科高校（校名 / 省份 / 城市 / 坐标 / 官网）

本模块只负责“有什么数据、怎么生成演示活动”，不含任何界面代码，
app.py 里的 5+1 个工具与页面都从这里取数据，保证口径一致。
"""
import json
import zlib
from datetime import date, timedelta
from pathlib import Path

DATA_DIR = Path(__file__).parent

LINK_CN_VOL = "https://www.chinavolunteer.cn"      # 中国志愿服务网（全国）
LINK_SH_VOL = "https://www.volunteer.sh.cn"        # 上海志愿者网
LINK_SANDA = "https://www.sandau.edu.cn"           # 上海杉达学院
LINK_SSTM = "https://www.sstm.org.cn"              # 上海科技馆 / 天文馆
LINK_LIB = "https://www.library.sh.cn"             # 上海图书馆
LINK_ART = "https://www.artmuseumonline.org"       # 中华艺术宫
LINK_BLOOD = "https://www.sbc.org.cn"              # 上海市血液中心


def _load(name: str, default):
    p = DATA_DIR / name
    try:
        with p.open(encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:      # 数据文件缺失时退化为空，不影响启动
        return default


REGIONS = _load("regions.json", {})
UNIVERSITIES = _load("universities.json", [])

# 少数高校在公开数据集里只有分校区坐标，这里做人工校正（主校区）
COORD_FIX = {
    "复旦大学": (31.2989, 121.5033, "上海市杨浦区邯郸路220号", "邯郸校区"),
    "上海交通大学": (31.0243, 121.4338, "上海市闵行区东川路800号", "闵行校区"),
    "同济大学": (31.2829, 121.5060, "上海市杨浦区四平路1239号", "四平路校区"),
    "清华大学": (40.0030, 116.3260, "北京市海淀区清华园1号", "主校区"),
    "北京大学": (39.9925, 116.3057, "北京市海淀区颐和园路5号", "燕园校区"),
    "上海杉达学院": (31.2687, 121.6592, "上海市浦东新区金海路2727号", "金海校区"),
    "四川大学": (30.6300, 104.0830, "四川省成都市武侯区一环路南一段24号", "望江校区"),
    "中山大学": (23.0965, 113.2988, "广东省广州市海珠区新港西路135号", "南校园"),
    "浙江大学": (30.3040, 120.0870, "浙江省杭州市西湖区余杭塘路866号", "紫金港校区"),
}

# 上海杉达学院有两个校区，单独展开（演示默认学校）
EXTRA_CAMPUS = [
    {"名称": "上海杉达学院（嘉善校区）", "省份": "浙江省", "城市": "嘉兴",
     "区": "嘉善县", "地址": "浙江省嘉兴市嘉善县人民大道505号",
     "lat": 30.8284, "lng": 120.9283, "官网": "https://www.sandau.edu.cn",
     "校区": "嘉善校区", "层次": "本科", "985": False, "211": False},
]


def haversine_km(lat1, lng1, lat2, lng2) -> float:
    """两坐标间球面距离（公里）"""
    from math import radians, sin, cos, asin, sqrt

    lat1, lng1, lat2, lng2 = map(radians, [lat1, lng1, lat2, lng2])
    dlat, dlng = lat2 - lat1, lng2 - lng1
    a = sin(dlat / 2) ** 2 + cos(lat1) * cos(lat2) * sin(dlng / 2) ** 2
    return round(2 * 6371 * asin(sqrt(a)), 1)


# ============================ 地区 ============================
def provinces() -> list:
    return list(REGIONS.keys())


def cities(prov: str) -> list:
    return list(REGIONS.get(prov, {}).keys())


def areas(prov: str, city: str) -> list:
    return list(REGIONS.get(prov, {}).get(city, {}).get("areas", []))


def city_coord(prov: str, city: str):
    info = REGIONS.get(prov, {}).get(city, {})
    if info.get("lat") and info.get("lng"):
        return float(info["lat"]), float(info["lng"])
    return None


def full_city(prov: str, city: str) -> str:
    """把高校数据里的城市简称（成都）对齐到行政区划里的全称（成都市）"""
    if not city:
        return city
    for c in cities(prov):
        if c == city:
            return c
        if c.startswith(city) or city.rstrip("市") == c.rstrip("市"):
            return c
    return city


def full_area(prov: str, city: str, area: str) -> str:
    """把高校数据里的区县名对齐到行政区划里的写法"""
    if not area:
        return area
    for a in areas(prov, city):
        if a == area or a.startswith(area) or area.startswith(a):
            return a
    return area


def region_label(prov: str, city: str = "", area: str = "") -> str:
    parts = [x for x in (prov, city if city and city != prov else "", area) if x]
    return " · ".join(parts) if parts else "全国"


# ============================ 高校 ============================
def _school_from_uni(u: dict) -> dict:
    name = u["名称"]
    lat, lng = u.get("lat"), u.get("lng")
    addr = u.get("地址") or ""
    campus = "主校区"
    if name in COORD_FIX:
        lat, lng, addr, campus = COORD_FIX[name]
    if not lat or not lng:
        c = city_coord(u["省份"], u["城市"])
        if c:
            lat, lng = c
    if not addr:
        addr = f"{u['省份']}{u['城市']}{u.get('区', '')}".strip()
    tags = []
    if u.get("985"):
        tags.append("985")
    elif u.get("211"):
        tags.append("211")
    tags.append(u.get("层次") or "本科")
    city_full = full_city(u["省份"], u["城市"])
    area_full = full_area(u["省份"], city_full, u.get("区", ""))
    # 地址里若出现更准确的区县名（如人工校正后的主校区地址），以地址为准
    if addr:
        for _a in areas(u["省份"], city_full):
            if _a in addr:
                area_full = _a
                break
    return {
        "名称": name, "省份": u["省份"], "城市": city_full, "区": area_full,
        "校区": campus, "地址": addr,
        "lat": round(float(lat), 4) if lat else None,
        "lng": round(float(lng), 4) if lng else None,
        "官网": u.get("官网", ""), "标签": tags,
    }


SCHOOLS = [_school_from_uni(u) for u in UNIVERSITIES]
if not any(s["名称"] == "上海杉达学院" for s in SCHOOLS):
    SCHOOLS.insert(0, {
        "名称": "上海杉达学院（金海校区）", "省份": "上海市", "城市": "上海市",
        "区": "浦东新区", "校区": "金海校区", "地址": "上海市浦东新区金海路2727号",
        "lat": 31.2687, "lng": 121.6592, "官网": LINK_SANDA, "标签": ["本科"],
    })
else:
    for s in SCHOOLS:
        if s["名称"] == "上海杉达学院":
            s["名称"] = "上海杉达学院（金海校区）"
            break
SCHOOLS.extend(EXTRA_CAMPUS)

# 统一城市 / 区县写法（高校数据用简称，行政区划用全称）
for _s in SCHOOLS:
    _s["城市"] = full_city(_s["省份"], _s["城市"])
    _s["区"] = full_area(_s["省份"], _s["城市"], _s.get("区", ""))

SCHOOL_BY_NAME = {s["名称"]: s for s in SCHOOLS}


def school_provinces() -> list:
    seen = []
    for s in SCHOOLS:
        if s["省份"] not in seen:
            seen.append(s["省份"])
    return seen or provinces()


def school_cities(prov: str) -> list:
    seen = []
    for s in SCHOOLS:
        if s["省份"] == prov and s["城市"] not in seen:
            seen.append(s["城市"])
    return seen


def schools_in(prov: str, city: str = "", keyword: str = "") -> list:
    out = [s for s in SCHOOLS if s["省份"] == prov and (not city or s["城市"] == city)]
    kw = (keyword or "").strip()
    if kw:
        out = [s for s in out if kw in s["名称"]]
    return out


# ============================ 活动模板 ============================
# 校外活动模板：{area}/{city}/{prov} 会被替换为所选地区，保证任何区县都有可看的内容
OFF_TEMPLATES = [
    {"cat": "图书管理", "title": "{area}图书馆“书香伴读”志愿服务", "org": "{area}图书馆",
     "place": "{area}图书馆总服务台", "hours": 4.0, "tags": ["伴读", "图书管理"], "w": 1.0,
     "intro": "在少儿与期刊阅览区协助整理归还图书、指导自助借还机使用，并为读者提供检索帮助。",
     "who": "在校生，耐心细致，能连续服务 3 小时以上", "meet": "图书馆一楼总服务台",
     "contact": "{area}图书馆读者服务部", "note": "馆内保持安静；不与未成年人单独离开阅览区"},
    {"cat": "科普", "title": "{city}科技馆展厅科普引导", "org": "{city}科学技术馆",
     "place": "{city}科技馆一层序厅", "hours": 6.0, "tags": ["科普", "场馆服务"], "w": 1.0,
     "intro": "在常设展区提供路线指引与展项原理解说，协助维持参观秩序。",
     "who": "在校大学生，理工科背景优先，需完成馆方线上岗前培训", "meet": "科技馆 1 号门志愿者之家",
     "contact": "{city}科技馆志愿者基地", "note": "全天岗，午餐由馆方提供；馆内禁止饮食"},
    {"cat": "助老", "title": "{area}社区“为老送餐”与陪伴", "org": "{area}街道社区志愿服务中心",
     "place": "{area}社区食堂及各小区", "hours": 3.0, "tags": ["助老", "社区"], "w": 1.1,
     "intro": "为社区高龄独居老人配送午餐，做简单问候与居家安全观察记录。",
     "who": "在校生，善于与老人沟通，会本地方言者优先", "meet": "{area}社区食堂门口",
     "contact": "街道为老服务中心", "note": "不进入老人卧室；不代收钱款；异常情况当日反馈社工"},
    {"cat": "环保", "title": "{area}生活垃圾分类督导", "org": "{area}城市管理志愿服务队",
     "place": "{area}各小区垃圾投放点", "hours": 2.5, "tags": ["环保", "社区治理"], "w": 1.1,
     "intro": "在投放高峰时段引导居民正确分类，记录混投情况并反馈街道。",
     "who": "在校生，能早起，佩戴手套与袖标上岗", "meet": "{area}城管志愿服务站",
     "contact": "{area}城市管理志愿服务队", "note": "只做劝导不与居民争执；雨天顺延"},
    {"cat": "交通引导", "title": "{city}地铁站客流引导", "org": "{city}轨道交通志愿服务总队",
     "place": "{city}轨道交通换乘枢纽站", "hours": 3.0, "tags": ["交通引导", "城市服务"], "w": 1.2,
     "intro": "晚高峰在闸机口与换乘通道引导客流，协助老人及携带大件行李的乘客进出站。",
     "who": "18 岁以上，能听懂普通话，服从站方调度", "meet": "车站站长室（站厅层）",
     "contact": "轨道交通志愿服务总队", "note": "不得进入轨行区；需提前 15 分钟到岗接受安全交底"},
    {"cat": "医疗协助", "title": "{area}中心医院导医协助", "org": "{area}人民医院志愿服务队",
     "place": "{area}人民医院门诊大厅", "hours": 4.0, "tags": ["医疗协助", "便民"], "w": 1.0,
     "intro": "在门诊大厅协助患者使用自助机挂号缴费、指引科室位置、维护候诊秩序。",
     "who": "在校生，沟通耐心；医学相关专业优先", "meet": "门诊大厅志愿服务岗",
     "contact": "医院团委 / 门诊部", "note": "不做任何医学判断、不解释检查结果与用药"},
    {"cat": "支教", "title": "{area}青少年活动中心课业辅导", "org": "{area}团委 · 青少年活动中心",
     "place": "{area}青少年活动中心", "hours": 4.0, "tags": ["支教", "青少年"], "w": 1.0,
     "intro": "为小学至初中学生提供作业辅导与兴趣小组协助，组织手工与体育活动。",
     "who": "在校生，有家教经验优先；需提供健康承诺", "meet": "青少年活动中心一楼大厅",
     "contact": "{area}团委学校部", "note": "全程两名志愿者在场；不得单独带学生离开教学区"},
    {"cat": "助残", "title": "{area}“无障碍陪伴”出行协助", "org": "{area}残疾人联合会",
     "place": "{area}城市公园及周边", "hours": 3.0, "tags": ["助残", "公园"], "w": 0.9,
     "intro": "陪同视障或行动不便的朋友游园、就医，提供口述影像与路线引导。",
     "who": "在校生，有耐心，愿意学习基础导盲技巧", "meet": "{area}残联服务大厅",
     "contact": "{area}残疾人联合会", "note": "先询问再搀扶；不擅自替对方做决定"},
    {"cat": "文化讲解", "title": "{city}博物馆志愿讲解", "org": "{city}博物馆",
     "place": "{city}博物馆观众服务中心", "hours": 5.0, "tags": ["讲解", "文化传承"], "w": 1.0,
     "intro": "为参观团队提供展线导览与重点展品讲解，协助文明观展提醒。",
     "who": "在校生，普通话标准；需背熟馆方讲解稿", "meet": "博物馆 0 米层观众服务中心",
     "contact": "博物馆公共教育部", "note": "讲解口径以馆方最新文稿为准；展厅内禁止触摸展品"},
    {"cat": "宣传", "title": "{area}新时代文明实践站志愿服务岗", "org": "{area}新时代文明实践中心",
     "place": "{area}新时代文明实践站", "hours": 3.0, "tags": ["宣传", "便民"], "w": 0.8,
     "intro": "在文明实践站提供办事指引、便民物品借用，并协助开展主题活动。",
     "who": "在校生，性格开朗，能使用基础办公设备", "meet": "{area}文明实践站前台",
     "contact": "{area}新时代文明实践中心", "note": "便民药品仅登记发放，不做用药建议"},
    {"cat": "大型活动", "title": "{city}马拉松赛会志愿服务", "org": "{city}体育局 · 赛事组委会",
     "place": "{city}赛道沿线补给点", "hours": 8.0, "tags": ["大型活动", "户外"], "w": 0.7,
     "intro": "负责补给点布置、饮水发放、赛道秩序维护与选手引导。",
     "who": "在校生，身体健康，能长时间户外站立", "meet": "组委会指定集结点（提前统一乘车）",
     "contact": "赛事组委会志愿者部", "note": "需参加赛前培训；严禁进入赛道干扰选手"},
    {"cat": "应急科普", "title": "{area}消防与急救知识宣传", "org": "{area}消防救援站 · 红十字会",
     "place": "{area}社区广场", "hours": 3.0, "tags": ["应急", "宣传"], "w": 0.8,
     "intro": "发放消防与急救知识手册，协助演示灭火器使用与心肺复苏操作。",
     "who": "在校生，已参加红十字会急救培训者优先", "meet": "{area}消防救援站",
     "contact": "{area}红十字会", "note": "演示内容以红十字会与应急管理部门公开教材为准"},
]

# 校内活动模板：{school} 会被替换为绑定学校名
ON_TEMPLATES = [
    {"title": "图书馆书库整理与导读", "org": "{school}图书馆", "place": "图书馆 3F 书库",
     "hours": 3.0, "tags": ["图书管理", "室内"],
     "intro": "协助馆员完成错架整理、新书上架与阅览区导读，服务对象为全校师生。",
     "who": "在校生，专业不限；能连续站立 2 小时以上", "meet": "图书馆 1F 服务台",
     "contact": "图书馆读者服务部", "note": "馆内保持安静；服务结束后现场签退认定时长"},
    {"title": "校运会赛事保障志愿者", "org": "校体育运动委员会", "place": "东区田径场",
     "hours": 4.0, "tags": ["大型活动", "户外"],
     "intro": "负责检录引导、器材搬运、赛道秩序维护与终点搀扶。",
     "who": "在校生，身体健康；有急救证者优先", "meet": "田径场主席台西侧服务点",
     "contact": "校体委办公室", "note": "穿运动鞋、自备防晒；雨天照常举行"},
    {"title": "校园开放日咨询引导", "org": "招生办公室", "place": "南门报到大厅",
     "hours": 6.0, "tags": ["引导", "对外接待"],
     "intro": "为前来参加开放日的高中生及家长提供路线引导、资料发放与答疑。",
     "who": "在校生，普通话标准，熟悉校园布局", "meet": "南门报到大厅签到处",
     "contact": "招生办公室", "note": "提前参加线上培训；统一着志愿者服"},
    {"title": "无偿献血现场服务", "org": "校医院 · 血液中心", "place": "大学生活动中心 1 号厅",
     "hours": 6.0, "tags": ["公益", "医疗协助"],
     "intro": "配合医护完成登记填表、排队引导、糖水发放与献血后观察陪护。",
     "who": "在校生，本人不参与献血", "meet": "大学生活动中心 1 号厅入口",
     "contact": "校医院保健科", "note": "服务期间佩戴口罩；不代替医护做医学判断"},
    {"title": "“书香校园”读书月布展", "org": "图书馆 · 学生会", "place": "图书馆 1F 中庭",
     "hours": 4.0, "tags": ["布展", "文化活动"],
     "intro": "参与读书月主题展板搭建、书目推荐卡制作与现场秩序维护。",
     "who": "在校生，有手工或设计经验者优先", "meet": "图书馆 1F 中庭",
     "contact": "图书馆办公室", "note": "使用美工刀等工具注意安全；展后共同撤展"},
    {"title": "校园垃圾分类督导", "org": "后勤保卫处", "place": "学生公寓投放点",
     "hours": 2.0, "tags": ["环保", "校园治理"],
     "intro": "在早间投放高峰引导同学正确分类，记录混投情况并反馈后勤。",
     "who": "在校生，住校者优先", "meet": "学生公寓 3 号楼前投放点",
     "contact": "后勤保卫处环卫科", "note": "佩戴手套与袖标；不与被劝导同学争执"},
    {"title": "校园反诈宣传摊位", "org": "保卫处 · 学生会", "place": "学生食堂西门广场",
     "hours": 2.0, "tags": ["宣传", "安全"],
     "intro": "发放反诈手册，指导同学安装国家反诈中心 APP，讲解常见骗局。",
     "who": "在校生，表达清晰", "meet": "食堂西门广场",
     "contact": "保卫处", "note": "不代同学操作手机银行与支付账户"},
    {"title": "新生迎新接站与行李搬运", "org": "学生工作处", "place": "校门广场 · 宿舍区",
     "hours": 5.0, "tags": ["迎新", "引导"],
     "intro": "在报到点引导新生完成注册流程，协助搬运行李并解答常见问题。",
     "who": "在校生，能搬运轻便行李，熟悉报到流程", "meet": "校门广场迎新服务台",
     "contact": "学生工作处", "note": "注意防暑；不代收任何费用"},
]


def _rng(seed_text: str):
    import random

    return random.Random(zlib.crc32(seed_text.encode("utf-8")))


def _link_for(prov: str, cat: str) -> str:
    if prov in ("上海市",):
        return {
            "科普": LINK_SSTM, "图书管理": LINK_LIB, "文化讲解": LINK_ART,
            "医疗协助": LINK_BLOOD,
        }.get(cat, LINK_SH_VOL)
    return LINK_CN_VOL


def build_off_campus(prov: str, city: str, area: str, count: int = 8) -> list:
    """按所选地区生成校外活动（同一地区每次结果一致，模拟“当地正在招募”的项目）"""
    area = area or city or prov
    coord = city_coord(prov, city) or city_coord(prov, cities(prov)[0] if cities(prov) else "")
    base_lat, base_lng = coord if coord else (31.23, 121.47)
    today = date.today()

    picked = [t for t in OFF_TEMPLATES]
    rng = _rng(f"{prov}{city}{area}")
    rng.shuffle(picked)
    picked.sort(key=lambda t: -t["w"])          # 权重高的更容易靠前
    picked = picked[:count]

    out = []
    for i, t in enumerate(picked):
        jitter_lat = (rng.random() - 0.5) * 0.09
        jitter_lng = (rng.random() - 0.5) * 0.11
        d = today + timedelta(days=rng.randint(1, 18))
        slot = rng.choice(["08:30-12:30", "09:00-17:00", "13:30-17:30", "14:00-18:00", "10:00-16:00"])
        quota = rng.choice([12, 15, 18, 20, 25, 30, 40])
        signed = rng.randint(max(1, quota // 5), max(2, quota - 2))
        fill = lambda s: s.format(area=area, city=city or area, prov=prov)  # noqa: E731
        out.append({
            "id": f"O{prov[:2]}{city[:2]}{i:02d}",
            "kind": "off",
            "标题": fill(t["title"]),
            "主办": fill(t["org"]),
            "地址": f"{area} · {fill(t['place'])}",
            "区域": area, "城市": city, "省份": prov,
            "日期": d.strftime("%Y-%m-%d"),
            "时间段": slot,
            "时长": t["hours"],
            "名额": quota, "已报名": signed,
            "标签": t["tags"],
            "分类": t["cat"],
            "lat": round(base_lat + jitter_lat, 4),
            "lng": round(base_lng + jitter_lng, 4),
            "更新时间": rng.randint(1, 50),
            "报名链接": _link_for(prov, t["cat"]),
            "详情": {
                "介绍": fill(t["intro"]), "招募对象": fill(t["who"]),
                "集合": fill(t["meet"]), "联系人": fill(t["contact"]),
                "注意": fill(t["note"]),
            },
        })
    out.sort(key=lambda a: a["日期"])
    return out


def build_on_campus(school: dict, count: int = 6) -> list:
    """按绑定学校生成校内活动"""
    name = school.get("名称", "学校").split("（")[0]
    campus = school.get("校区", "")
    today = date.today()
    rng = _rng(f"{school.get('名称')}on")
    picked = list(ON_TEMPLATES)
    rng.shuffle(picked)
    picked = picked[:count]
    out = []
    for i, t in enumerate(picked):
        d = today + timedelta(days=rng.randint(1, 14))
        quota = rng.choice([10, 15, 20, 25, 30, 40])
        out.append({
            "id": f"C{school.get('名称', '')[:2]}{i:02d}",
            "kind": "on",
            "标题": t["title"],
            "主办": t["org"].format(school=name),
            "地点": f"{campus}{t['place']}",
            "区域": school.get("区", ""), "城市": school.get("城市", ""), "省份": school.get("省份", ""),
            "日期": d.strftime("%Y-%m-%d"),
            "时间段": rng.choice(["08:00-12:00", "13:00-17:00", "14:00-17:00", "18:30-21:00"]),
            "时长": t["hours"],
            "名额": quota,
            "已报名": rng.randint(max(1, quota // 5), max(2, quota - 1)),
            "标签": t["tags"] + ["校内"],
            "分类": "校内",
            "lat": school.get("lat"), "lng": school.get("lng"),
            "更新时间": rng.randint(1, 50),
            "报名链接": school.get("官网") or LINK_SANDA,
            "详情": {
                "介绍": t["intro"], "招募对象": t["who"], "集合": f"{name}{campus}{t['meet']}",
                "联系人": f"{t['contact']}（{name}）", "注意": t["note"],
            },
        })
    out.sort(key=lambda a: a["日期"])
    return out


def default_school() -> dict:
    """默认绑定学校：上海杉达学院（金海校区）"""
    for s in SCHOOLS:
        if s["名称"].startswith("上海杉达学院（金海"):
            return dict(s)
    return dict(SCHOOLS[0]) if SCHOOLS else {}
