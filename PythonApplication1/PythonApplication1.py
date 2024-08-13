import pandas as pd                 # 用于读取Excel文件并处理“Address”列数据，返回一个列表
import pgeocode                     # 用于查询邮政编码对应的城市名和地理坐标，并处理相关的地理位置信息。
from collections import Counter     # 用于统计列表中元素的频率，此处计算城市名出现的次数。
import folium                       # 用于创建和显示以美国中心的交互式地图，并在每个城市位置添加标记显示名称和计数。
import re                           # 用于正则表达式匹配邮政编码

# 读取 Excel 文件
file_path = 'C:/Users/Benson/Desktop/dataAnalyze/myData.xlsx'
df = pd.read_excel(file_path)

# 从地址中提取邮政编码
def extract_postal_code(address):
    # 查找所有5位数的数字
    postal_codes = re.findall(r'\b\d{5}\b', address)
    # 返回最后一个匹配的邮政编码，如果存在多个
    return postal_codes[-1] if postal_codes else None

# 获取邮政编码列表
df['PostalCode'] = df['Address'].dropna().apply(extract_postal_code)
postal_codes = df['PostalCode'].dropna().tolist()

# 初始化 pgeocode 查询工具，设置为美国
nomi = pgeocode.Nominatim('us')

cities = []
locations = []
unprocessed_codes = []  # 存储未能成功处理的邮政编码

# 遍历邮政编码列表，获取城市名和位置
for postal_code in postal_codes:
    try:
        location = nomi.query_postal_code(str(postal_code))
        if location is not None and not location.empty:
            if not pd.isnull(location.place_name):
                cities.append(location.place_name)
                locations.append((location.latitude, location.longitude))
            else:
                unprocessed_codes.append(postal_code)  # 如果没有地点名，记录邮政编码
        else:
            unprocessed_codes.append(postal_code)  # 如果查询返回空，记录邮政编码
    except Exception as e:
        print(f"Failed to process postal code {postal_code}: {e}")
        unprocessed_codes.append(postal_code)  # 如果发生异常，记录邮政编码

# 统计每个城市的出现次数
city_count = Counter(cities)

# 创建地图，中心设在美国中部
m = folium.Map(location=[39.8283, -98.5795], zoom_start=4)

# 添加标记
for city, location in zip(cities, locations):       # 使用 zip() 函数同时迭代 cities 和 locations 列表，以便在地图上为每个城市添加标记。
    count = city_count[city]
    folium.Marker(
        location,
        popup=f"{city} ({count})",
        tooltip=f"{city} ({count})"
    ).add_to(m)

# 保存地图为 HTML 文件
m.save("usa_city_map.html")

print("Map has been saved as usa_city_map.html")
