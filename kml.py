import streamlit as st
import pandas as pd
import re
from geopy.geocoders import Nominatim
import simplekml
from collections import defaultdict
from io import BytesIO

# ========= 页面标题 =========
st.title("🗺️ Excel → KML 转换工具")

# ========= 基本列名配置 =========
COL_ADDRESS = "Address"
COL_NAME = "Name"
COL_PAPER = "Newspaper"
COL_NOTE = "Notes"
COL_WEEK = "Week"

# 单订阅图标
PAPER_ICONS = {
    "NP-KC": "http://maps.google.com/mapfiles/kml/paddle/red-circle.png",
    "FT-K": "http://maps.google.com/mapfiles/kml/paddle/blu-circle.png",
    "BR": "http://maps.google.com/mapfiles/kml/paddle/grn-circle.png",
    "Süddeutsche": "http://maps.google.com/mapfiles/kml/paddle/ylw-circle.png",
    "FAZ": "http://maps.google.com/mapfiles/kml/paddle/purple-circle.png",
}

DEFAULT_ICON = "http://maps.google.com/mapfiles/kml/paddle/wht-circle.png"
MULTI_SUB_ICON = "http://maps.google.com/mapfiles/kml/paddle/ylw-stars.png"

# ========= 星期映射 =========
WEEK_MAP = {
    "MO": 1,
    "TU": 2,
    "WE": 3,
    "TH": 4,
    "FR": 5,
    "SA": 6,
}

VOLL_ABO_DAYS = {1, 2, 3, 4, 5, 6}

# ========= 工具函数 =========
def extract_house_number(address):
    match = re.search(r"\b\d+[a-zA-Z]?\b", address)
    return match.group(0) if match else ""

def parse_week_days(week_str):
    week_str = str(week_str).strip()
    if week_str == "Vollabo":
        return VOLL_ABO_DAYS.copy()

    days = set()
    for part in week_str.split("+"):
        part = part.strip()
        if part in WEEK_MAP:
            days.add(WEEK_MAP[part])
    return days


def build_kml_bytes(rows):
    kml = simplekml.Kml()
    geolocator = Nominatim(user_agent="newspaper_app")
    locations = defaultdict(list)

    # geocode + 聚合
    for _, row in rows.iterrows():
        address = str(row[COL_ADDRESS])
        try:
            loc = geolocator.geocode(address)
            if not loc:
                st.warning(f"无法定位: {address}")
                continue

            key = (round(loc.longitude, 6), round(loc.latitude, 6))
            locations[key].append(row)
        except Exception:
            continue

    # 生成点
    for (lon, lat), items in locations.items():
        is_multi_today = len(items) > 1

        address = str(items[0][COL_ADDRESS])
        house_no = extract_house_number(address)

        if is_multi_today:
            name = f"{house_no}（{len(items)}份）"
        else:
            person = str(items[0][COL_NAME]).strip()
            paper = str(items[0][COL_PAPER]).strip()
            name = f"{house_no} {paper} + {person}"

        desc_lines = [f"地址：{address}", ""]
        for i, r in enumerate(items, 1):
            desc_lines.append(
                f"{i}. {r[COL_NAME]}｜{r[COL_PAPER]}｜{r[COL_WEEK]}｜{r[COL_NOTE]}"
            )

        p = kml.newpoint(
            name=name,
            description="\n".join(desc_lines),
            coords=[(lon, lat)]
        )

        style = simplekml.Style()
        if is_multi_today:
            style.iconstyle.icon.href = MULTI_SUB_ICON
            style.iconstyle.scale = 1.2
        else:
            style.iconstyle.icon.href = PAPER_ICONS.get(
                str(items[0][COL_PAPER]).strip(), DEFAULT_ICON
            )
            style.iconstyle.scale = 1.1

        p.style = style

    # 转为字节流
    buffer = BytesIO()
    buffer.write(kml.kml().encode("utf-8"))
    buffer.seek(0)

    return buffer


# ========= 文件上传 =========
uploaded_file = st.file_uploader("上传 Excel 文件", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    st.success("Excel 读取成功！")

    # 分组
    groups = {
        "monday": [],
        "weekday_234": [],
        "weekend_56": [],
    }

    for _, row in df.iterrows():
        days = parse_week_days(row[COL_WEEK])

        if 1 in days:
            groups["monday"].append(row)
        if days & {2, 3, 4}:
            groups["weekday_234"].append(row)
        if days & {5, 6}:
            groups["weekend_56"].append(row)

    if st.button("开始生成 KML"):
        st.info("正在地理编码地址，请稍等...")

        for group_name, rows in groups.items():
            if rows:
                rows_df = pd.DataFrame(rows)
                kml_bytes = build_kml_bytes(rows_df)

                st.download_button(
                    label=f"下载 {group_name}.kml",
                    data=kml_bytes,
                    file_name=f"{group_name}.kml",
                    mime="application/vnd.google-earth.kml+xml"
                )

        st.success("生成完成！")
