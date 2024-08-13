import pandas as pd
import pgeocode
from collections import defaultdict, Counter
import folium
import re

# 读取 Excel 文件
file_path = 'C:/Users/Benson/Desktop/dataAnalyze/myData.xlsx'
df = pd.read_excel(file_path)

# 从地址中提取邮政编码
def extract_postal_code(address):
    postal_codes = re.findall(r'\b\d{5}\b', address)
    return postal_codes[-1] if postal_codes else None

df['PostalCode'] = df['Address'].dropna().apply(extract_postal_code)
postal_codes = df['PostalCode'].dropna().tolist()

# 初始化 pgeocode 查询工具，设置为美国
nomi = pgeocode.Nominatim('us')

city_to_locations = defaultdict(lambda: None)  # 存储每个城市的位置
city_to_zipcodes = defaultdict(list)  # 用来存储每个城市的邮政编码列表
unprocessed_codes = []

# 遍历邮政编码列表，获取城市名和位置
for postal_code in postal_codes:
    try:
        location = nomi.query_postal_code(str(postal_code))
        if location is not None and not location.empty:
            if not pd.isnull(location.place_name):
                city = location.place_name
                if city_to_locations[city] is None:
                    city_to_locations[city] = (location.latitude, location.longitude)
                city_to_zipcodes[city].append(postal_code)
            else:
                unprocessed_codes.append(postal_code)
        else:
            unprocessed_codes.append(postal_code)
    except Exception as e:
        print(f"Failed to process postal code {postal_code}: {e}")
        unprocessed_codes.append(postal_code)

# 创建地图，中心设在美国中部
m = folium.Map(location=[39.8283, -98.5795], zoom_start=4)

# 添加标记
for city, location in city_to_locations.items():
    if location is not None:
        count = len(city_to_zipcodes[city])
        zipcodes = ', '.join(city_to_zipcodes[city])
        folium.Marker(
            location,
            popup=f"{city} ({count})\nZip Codes: {zipcodes}",
            tooltip=f"{city} ({count})"
        ).add_to(m)

# 保存地图为 HTML 文件
m.save("usa_city_map.html")

print("Map has been saved as usa_city_map.html")
