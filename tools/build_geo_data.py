# -*- coding: utf-8 -*-
"""生成项目所需的行政区划与高校数据（离线构建脚本，只在需要更新数据时运行）。

数据来源（均为公开数据，仅供课程演示使用）：
1) 省 / 市 / 区县三级行政区划
   modood/Administrative-divisions-of-China · dist/pca-code.json
   https://github.com/modood/Administrative-divisions-of-China
2) 全国本科高校名单（校名 / 省份 / 城市 / 层次 / 985-211 / 官网）
   luhuadong/open-edu-data · data/china-universities.csv
   https://github.com/luhuadong/open-edu-data
3) 高校经纬度（按校名匹配，用于计算校外活动距离）
   ZsTs119/china-university-database · public/大学.json
   https://github.com/ZsTs119/china-university-database

输出：regions.json（province -> city -> {areas, lat, lng}）、universities.json
用法：python tools/build_geo_data.py
"""
import csv
import io
import json
import os
import re
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
PROJECT = os.path.dirname(HERE)
CACHE = os.path.join(HERE, ".cache")

SRC_PCA = "https://cdn.jsdelivr.net/gh/modood/Administrative-divisions-of-China@master/dist/pca-code.json"
SRC_UNI = "https://cdn.jsdelivr.net/gh/luhuadong/open-edu-data@main/data/china-universities.csv"
SRC_GEO = "https://cdn.jsdelivr.net/gh/ZsTs119/china-university-database@main/public/%E5%A4%A7%E5%AD%A6.json"

# 行政区划数据只覆盖大陆 31 个省级单位，港澳台按官方口径补齐
EXTRA = {
    "香港特别行政区": {
        "香港岛": ["中西区", "湾仔区", "东区", "南区"],
        "九龙": ["油尖旺区", "深水埗区", "九龙城区", "黄大仙区", "观塘区"],
        "新界": ["荃湾区", "屯门区", "元朗区", "北区", "大埔区", "西贡区", "沙田区", "葵青区", "离岛区"],
    },
    "澳门特别行政区": {
        "澳门半岛": ["花地玛堂区", "圣安多尼堂区", "大堂区", "望德堂区", "风顺堂区"],
        "离岛": ["嘉模堂区", "圣方济各堂区", "路氹填海区"],
    },
    "台湾省": {
        "台北市": ["全市"], "新北市": ["全市"], "桃园市": ["全市"], "台中市": ["全市"],
        "台南市": ["全市"], "高雄市": ["全市"], "基隆市": ["全市"], "新竹市": ["全市"],
        "嘉义市": ["全市"], "新竹县": ["全县"], "苗栗县": ["全县"], "彰化县": ["全县"],
        "南投县": ["全县"], "云林县": ["全县"], "嘉义县": ["全县"], "屏东县": ["全县"],
        "宜兰县": ["全县"], "花莲县": ["全县"], "台东县": ["全县"], "澎湖县": ["全县"],
        "金门县": ["全县"], "连江县": ["全县"],
    },
}

SHORT_PROV = {
    "北京": "北京市", "天津": "天津市", "河北": "河北省", "山西": "山西省",
    "内蒙古": "内蒙古自治区", "辽宁": "辽宁省", "吉林": "吉林省", "黑龙江": "黑龙江省",
    "上海": "上海市", "江苏": "江苏省", "浙江": "浙江省", "安徽": "安徽省",
    "福建": "福建省", "江西": "江西省", "山东": "山东省", "河南": "河南省",
    "湖北": "湖北省", "湖南": "湖南省", "广东": "广东省", "广西": "广西壮族自治区",
    "海南": "海南省", "重庆": "重庆市", "四川": "四川省", "贵州": "贵州省",
    "云南": "云南省", "西藏": "西藏自治区", "陕西": "陕西省", "甘肃": "甘肃省",
    "青海": "青海省", "宁夏": "宁夏回族自治区", "新疆": "新疆维吾尔自治区",
}
CITY_SUFFIX = ("市", "地区", "自治州", "盟", "特别行政区")


def fetch(url, name):
    os.makedirs(CACHE, exist_ok=True)
    path = os.path.join(CACHE, name)
    if not os.path.exists(path):
        print(f"下载 {name} ...")
        with urllib.request.urlopen(url, timeout=90) as resp, open(path, "wb") as fh:
            fh.write(resp.read())
    return path


def short_city(name):
    for suf in CITY_SUFFIX:
        if name.endswith(suf) and len(name) > len(suf):
            return name[: -len(suf)]
    return name


def build_regions(pca_path):
    data = json.load(open(pca_path, encoding="utf-8"))
    regions = {}
    for prov in data:
        pname = prov["name"]
        cities = {}
        for city in prov.get("children") or []:
            cname = city["name"]
            if cname in ("市辖区", "县", "省直辖县级行政区划", "自治区直辖县级行政区划"):
                cname = pname          # 直辖市把“市辖区”替换成市名本身
            areas = [a["name"] for a in (city.get("children") or [])]
            if not areas:
                areas = [cname]
            cities[cname] = {"areas": areas, "lat": None, "lng": None}
        regions[pname] = cities
    for pname, cities in EXTRA.items():
        regions[pname] = {c: {"areas": a, "lat": None, "lng": None} for c, a in cities.items()}
    return regions


def build_universities(csv_path, geo_path):
    rows = list(csv.DictReader(io.open(csv_path, encoding="utf-8")))
    geo = json.load(open(geo_path, encoding="utf-8"))

    # 校名 -> 坐标（同名多条时优先取省份一致的那条）
    by_name = {}
    for g in geo:
        nm = (g.get("name") or "").strip()
        loc = g.get("location") or {}
        if not nm or not loc.get("lat"):
            continue
        by_name.setdefault(nm, []).append(
            {"lat": loc["lat"], "lng": loc["lng"], "prov": g.get("province", ""), "area": g.get("area", ""), "addr": g.get("address", "")}
        )

    out, miss = [], 0
    for r in rows:
        name = (r["name"] or "").strip()
        prov_short = (r["province"] or "").strip()
        prov = SHORT_PROV.get(prov_short, prov_short)
        city = (r["city"] or "").strip()
        cand = by_name.get(name)
        if not cand:
            # 精确匹配失败时退化为前缀 / 包含匹配（数据集里常带“（xx校区）”后缀）
            pref = [v for k, v in by_name.items() if k.startswith(name) and len(k) <= len(name) + 8]
            if not pref:
                pref = [v for k, v in by_name.items() if name in k and len(k) <= len(name) + 8]
            cand = [item for sub in pref for item in sub] or None
        lat = lng = None
        area = ""
        addr = ""
        if cand:
            pick = next((c for c in cand if prov_short in c["prov"]), cand[0])
            lat, lng, area, addr = pick["lat"], pick["lng"], pick["area"], pick.get("addr", "")
        else:
            miss += 1
        out.append({
            "名称": name, "省份": prov, "城市": city, "区": area, "地址": addr,
            "lat": lat, "lng": lng,
            "层次": r.get("level", "本科"),
            "985": r.get("project_985", "0") == "1",
            "211": r.get("project_211", "0") == "1",
            "官网": (r.get("website") or "").strip(),
        })
    print(f"高校 {len(out)} 所，其中 {miss} 所没有匹配到坐标")
    return out


def fill_coords(regions, unis):
    """用高校坐标推算每个城市的中心坐标（取该市高校坐标均值）"""
    from collections import defaultdict
    from statistics import median

    acc = defaultdict(list)
    for u in unis:
        if u["lat"] and u["lng"] and u["城市"]:
            acc[(u["省份"], u["城市"])].append((u["lat"], u["lng"]))
    prov_acc = defaultdict(list)
    for (p, _c), pts in acc.items():
        prov_acc[p].extend(pts)

    hit = 0
    for pname, cities in regions.items():
        for cname, info in cities.items():
            key = (pname, cname)
            pts = acc.get(key)
            if not pts:
                # 用“市”后缀/短名再试一次
                pts = None
                for (p2, c2), v in acc.items():
                    if p2 == pname and (short_city(cname) == c2 or cname.startswith(c2) or c2.startswith(short_city(cname))):
                        pts = v
                        break
            if pts:
                info["lat"] = round(median([x[0] for x in pts]), 4)
                info["lng"] = round(median([x[1] for x in pts]), 4)
                hit += 1
            elif prov_acc.get(pname):
                pts = prov_acc[pname]
                info["lat"] = round(median([x[0] for x in pts]), 4)
                info["lng"] = round(median([x[1] for x in pts]), 4)
    print(f"城市坐标：{hit} 个精确命中，其余按省份均值兜底")


def main():
    pca = fetch(SRC_PCA, "pca-code.json")
    uni_csv = fetch(SRC_UNI, "china-universities.csv")
    geo = fetch(SRC_GEO, "uni-geo.json")

    regions = build_regions(pca)
    unis = build_universities(uni_csv, geo)
    fill_coords(regions, unis)

    rp = os.path.join(PROJECT, "regions.json")
    up = os.path.join(PROJECT, "universities.json")
    json.dump(regions, open(rp, "w", encoding="utf-8"), ensure_ascii=False, separators=(",", ":"))
    json.dump(unis, open(up, "w", encoding="utf-8"), ensure_ascii=False, separators=(",", ":"))

    n_city = sum(len(v) for v in regions.values())
    n_area = sum(len(c["areas"]) for v in regions.values() for c in v.values())
    print(f"regions.json : {len(regions)} 省 / {n_city} 市 / {n_area} 区县 · {os.path.getsize(rp)//1024} KB")
    print(f"universities.json : {len(unis)} 所 · {os.path.getsize(up)//1024} KB")


if __name__ == "__main__":
    main()
