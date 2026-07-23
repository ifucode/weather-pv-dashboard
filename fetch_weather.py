# -*- coding: utf-8 -*-
"""
天气 + 光伏发电效率 每日抓取脚本
- 读取 city.xlsx 城市列表
- 抓取当天 06:00-18:00 逐小时天气 (Open-Meteo, 免费无需Key)
- 计算光伏发电效率 / 预计发电量
- 写入本地 data/ (JSON + data.js 供看板离线读取)
"""
import os
import sys
import json
import math
import datetime
import time

import requests

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
CITY_FILE = os.path.join(BASE_DIR, "city.xlsx")

# 主要城市经纬度 (缺省内置, 保证自动化稳定; 新城市走在线地理编码)
CITY_COORDS = {
    "湖州": (30.8945, 120.0937),
    "金华": (29.0793, 119.6479),
    "衢州": (28.9411, 118.8726),
    "温州": (27.9938, 120.6994),
    "上海": (31.2304, 121.4737),
    "无锡": (31.4912, 120.3119),
    "镇江": (32.2027, 119.4533),
    "淮安": (33.5012, 119.1510),
    "海安": (32.5486, 120.4730),
    "郴州": (25.7683, 113.0118),
}

# WMO weather_code -> 中文描述 + 图标
WMO = {
    0: ("晴", "☀️"), 1: ("多云间晴", "🌤️"), 2: ("多云", "⛅"), 3: ("阴", "☁️"),
    45: ("雾", "🌫️"), 48: ("霜雾", "🌫️"),
    51: ("小毛毛雨", "🌦️"), 53: ("毛毛雨", "🌦️"), 55: ("大毛毛雨", "🌧️"),
    56: ("冻毛毛雨", "🌧️"), 57: ("强冻毛毛雨", "🌧️"),
    61: ("小雨", "🌦️"), 63: ("中雨", "🌧️"), 65: ("大雨", "🌧️"),
    66: ("冻雨", "🌧️"), 67: ("强冻雨", "🌧️"),
    71: ("小雪", "🌨️"), 73: ("中雪", "🌨️"), 75: ("大雪", "❄️"), 77: ("雪粒", "🌨️"),
    80: ("阵雨", "🌦️"), 81: ("强阵雨", "🌧️"), 82: ("暴雨", "⛈️"),
    85: ("阵雪", "🌨️"), 86: ("强阵雪", "❄️"),
    95: ("雷阵雨", "⛈️"), 96: ("雷阵雨伴冰雹", "⛈️"), 99: ("强雷雨伴冰雹", "⛈️"),
}


def desc_of(code):
    return WMO.get(code, ("未知", "❓"))


def read_cities():
    """读取 city.xlsx -> [(省, 市), ...]"""
    import openpyxl
    wb = openpyxl.load_workbook(CITY_FILE, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    cities = []
    for r in rows[1:]:
        if not r or r[0] is None or (len(r) > 1 and r[1] is None):
            continue
        prov, city = str(r[0]).strip(), str(r[1]).strip()
        if city:
            cities.append((prov, city))
    return cities


def geocode(name):
    """在线地理编码 (内置未命中时兜底)"""
    try:
        r = requests.get(
            "https://geocoding-api.open-meteo.com/v1/search",
            params={"name": name, "count": 1, "language": "zh", "format": "json"},
            timeout=20,
        )
        js = r.json()
        if js.get("results"):
            res = js["results"][0]
            return (res["latitude"], res["longitude"])
    except Exception as e:
        print(f"  地理编码失败 {name}: {e}")
    return None


def solar_elevation(lat, lon, dt_local, tz_offset=8):
    """计算太阳高度角(度). dt_local 为当地时间(不含tz)"""
    day = dt_local.timetuple().tm_yday
    hour = dt_local.hour + dt_local.minute / 60.0
    decl = math.radians(23.45) * math.sin(math.radians(360.0 * (284 + day) / 365.0))
    # 时角: 以真太阳时近似 (按经度修正)
    solar_time = hour + (lon - tz_offset * 15.0) / 15.0
    ha = math.radians(15.0 * (solar_time - 12.0))
    latr = math.radians(lat)
    sin_elev = (math.sin(latr) * math.sin(decl) +
                math.cos(latr) * math.cos(decl) * math.cos(ha))
    return math.degrees(math.asin(max(-1.0, min(1.0, sin_elev))))


def clearsky_ghi(elev_deg):
    """Haurwitz 晴空GHI模型 (W/m2)"""
    if elev_deg <= 0:
        return 0.0
    sin_elev = math.sin(math.radians(elev_deg))
    return 1098.0 * sin_elev * math.exp(-0.059 / sin_elev)


def compute_pv(hours, lat, lon, date_str):
    """
    计算光伏指标.
    返回: 每小时PV效率 + 全天汇总
    """
    STC = 1000.0            # 标准辐照 W/m2
    TEMP_COEF = -0.0040     # 功率温度系数 /°C (晶硅约 -0.4%/°C)
    SYS_LOSS = 0.86         # 系统综合效率(逆变/线损/污渍等)

    hourly_pv = []
    energy_actual = 0.0     # kWh/kWp 等效
    energy_clear = 0.0      # 晴空理论 kWh/kWp
    for h in hours:
        ghi = h["ghi"] or 0.0
        tair = h["temp"] if h["temp"] is not None else 25.0
        # 电池温度 (NOCT 近似)
        tcell = tair + ghi / 800.0 * 20.0
        tfactor = 1.0 + TEMP_COEF * (tcell - 25.0)
        # 逐小时相对额定功率的出力比 (1kWp)
        p_ratio = max(0.0, (ghi / STC) * tfactor * SYS_LOSS)
        energy_actual += p_ratio  # 每步1小时 -> kWh/kWp

        # 晴空理论出力
        t = datetime.datetime.strptime(h["time"], "%Y-%m-%dT%H:%M")
        elev = solar_elevation(lat, lon, t)
        cghi = clearsky_ghi(elev)
        tcell_c = tair + cghi / 800.0 * 20.0
        p_ratio_clear = max(0.0, (cghi / STC) * (1.0 + TEMP_COEF * (tcell_c - 25.0)) * SYS_LOSS)
        energy_clear += p_ratio_clear

        hourly_pv.append(round(p_ratio, 3))

    # 天气影响效率 = 实际/晴空 (%) —— 反映云雨对发电的削减
    weather_eff = (energy_actual / energy_clear * 100.0) if energy_clear > 0 else 0.0
    weather_eff = min(100.0, weather_eff)

    if weather_eff >= 85:
        rating, rlevel = "优", "excellent"
    elif weather_eff >= 65:
        rating, rlevel = "良", "good"
    elif weather_eff >= 45:
        rating, rlevel = "中", "fair"
    elif weather_eff >= 25:
        rating, rlevel = "较差", "poor"
    else:
        rating, rlevel = "差", "bad"

    return {
        "hourly_pv_ratio": hourly_pv,
        "energy_kwh_per_kwp": round(energy_actual, 2),   # 预计发电量 (每kWp装机)
        "peak_sun_hours": round(energy_actual, 2),
        "weather_efficiency_pct": round(weather_eff, 1),  # 天气影响下发电效率
        "rating": rating,
        "rating_level": rlevel,
    }


def fetch_city_once(prov, city, date_str):
    """单次抓取一个城市; 失败抛异常由上层重试。"""
    coord = CITY_COORDS.get(city) or geocode(city)
    if not coord:
        print(f"  跳过 {city}: 无法获取经纬度")
        return None
    lat, lon = coord
    r = requests.get(
        "https://api.open-meteo.com/v1/forecast",
        params={
            "latitude": lat, "longitude": lon,
            "hourly": "temperature_2m,cloud_cover,precipitation,shortwave_radiation,wind_speed_10m,weather_code,relative_humidity_2m",
            "timezone": "Asia/Shanghai",
            "start_hour": f"{date_str}T06",
            "end_hour": f"{date_str}T18",
        },
        timeout=40,
    )
    r.raise_for_status()
    d = r.json()
    H = d.get("hourly", {})
    times = H.get("time", [])
    if not times:
        print(f"  {city}: 无数据")
        return None

    hours = []
    codes = []
    temps = []
    precip_sum = 0.0
    for i, t in enumerate(times):
        code = H["weather_code"][i]
        desc, icon = desc_of(code)
        temp = H["temperature_2m"][i]
        pr = H["precipitation"][i] or 0.0
        precip_sum += pr
        if temp is not None:
            temps.append(temp)
        codes.append(code)
        hours.append({
            "time": t,
            "hour": t.split("T")[1][:5],
            "temp": temp,
            "humidity": H["relative_humidity_2m"][i],
            "cloud": H["cloud_cover"][i],
            "precip": round(pr, 1),
            "ghi": H["shortwave_radiation"][i],
            "wind": H["wind_speed_10m"][i],
            "code": code,
            "desc": desc,
            "icon": icon,
        })

    # 全天主导天气 (取白天出现最多的天气码)
    dom_code = max(set(codes), key=codes.count)
    dom_desc, dom_icon = desc_of(dom_code)

    pv = compute_pv(hours, lat, lon, date_str)

    return {
        "province": prov,
        "city": city,
        "date": date_str,
        "lat": lat, "lon": lon,
        "temp_max": round(max(temps), 1) if temps else None,
        "temp_min": round(min(temps), 1) if temps else None,
        "precip_total": round(precip_sum, 1),
        "dominant_desc": dom_desc,
        "dominant_icon": dom_icon,
        "hours": hours,
        "pv": pv,
        "updated_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }


def fetch_city(prov, city, date_str, retries=3):
    """带重试的抓取; 全部失败返回 None(由 main 兜底填充)。"""
    last_err = None
    for attempt in range(1, retries + 1):
        try:
            rec = fetch_city_once(prov, city, date_str)
            if rec is not None:
                return rec
            return None  # 无数据(非网络错误), 不重试
        except Exception as e:
            last_err = e
            print(f"  尝试 {attempt}/{retries} 失败 {city}: {e}")
            if attempt < retries:
                time.sleep(2 * attempt)
    print(f"  {city} 全部重试失败: {last_err}")
    return None


def main():
    os.makedirs(DATA_DIR, exist_ok=True)
    date_str = (sys.argv[1] if len(sys.argv) > 1
                else datetime.date.today().isoformat())
    print(f"抓取日期: {date_str}")

    cities = read_cities()
    print(f"城市数量: {len(cities)}")

    records = []
    for prov, city in cities:
        print(f"-> {prov} {city}")
        try:
            rec = fetch_city(prov, city, date_str)
            if rec:
                records.append(rec)
                print(f"   OK  {rec['dominant_desc']}  发电效率 {rec['pv']['weather_efficiency_pct']}%  预计 {rec['pv']['energy_kwh_per_kwp']} kWh/kWp")
        except Exception as e:
            print(f"   失败: {e}")
        time.sleep(0.4)

    # 失败兜底: 抓不到的城市, 用历史中最近一天的同城市数据补上, 保证城市不缺失
    prev = {}
    for fn in os.listdir(DATA_DIR):
        if fn.endswith(".json") and fn != "all_data.json" and len(fn) == 15:
            try:
                with open(os.path.join(DATA_DIR, fn), encoding="utf-8") as f:
                    obj = json.load(f)
                prev[obj["date"]] = obj["records"]
            except Exception:
                pass
    got = {r["city"] for r in records}
    fallback_date = date_str if date_str in prev else (max(prev.keys()) if prev else None)
    if fallback_date:
        for rec in prev[fallback_date]:
            if rec["city"] not in got:
                rec = dict(rec)
                rec["stale"] = True
                rec["updated_at"] = f"{rec.get('updated_at','')} (兜底:沿用{fallback_date})"
                records.append(rec)
                print(f"   兜底填充 {rec['city']} (沿用 {fallback_date} 数据)")

    # 保存当日文件
    day_file = os.path.join(DATA_DIR, f"{date_str}.json")
    with open(day_file, "w", encoding="utf-8") as f:
        json.dump({"date": date_str, "records": records}, f, ensure_ascii=False, indent=2)

    # 汇总所有历史 -> all_data.json + data.js
    all_data = {}
    for fn in os.listdir(DATA_DIR):
        if fn.endswith(".json") and fn != "all_data.json" and len(fn) == 15:
            try:
                with open(os.path.join(DATA_DIR, fn), encoding="utf-8") as f:
                    obj = json.load(f)
                all_data[obj["date"]] = obj["records"]
            except Exception:
                pass

    with open(os.path.join(DATA_DIR, "all_data.json"), "w", encoding="utf-8") as f:
        json.dump(all_data, f, ensure_ascii=False, indent=2)

    with open(os.path.join(DATA_DIR, "data.js"), "w", encoding="utf-8") as f:
        f.write("window.WEATHER_DATA = ")
        json.dump(all_data, f, ensure_ascii=False)
        f.write(";\n")
        f.write(f"window.LATEST_DATE = {json.dumps(max(all_data.keys()))};\n")

    print(f"\n完成! 当日 {len(records)} 个城市, 历史共 {len(all_data)} 天。")


if __name__ == "__main__":
    main()
