import pandas as pd                 #用于读取Excel文件并处理“Address”列数据，返回一个列表
import pgeocode                     #用于查询邮政编码对应的城市名和地理坐标，并处理相关的地理位置信息。
from collections import Counter     #用于统计列表中元素的频率，此处计算城市名出现的次数。
import folium                       #用于创建和显示以美国中心的交互式地图，并在每个城市位置添加标记显示名称和计数。

# 读取 Excel 文件
file_path = 'C:/Users/Benson/Desktop/dataAnalyze/myData.xlsx'  
df = pd.read_excel(file_path)

# 获取邮政编码列表
postal_codes = df['Address'].dropna().tolist()

# 初始化 pgeocode 查询工具，设置为美国
nomi = pgeocode.Nominatim('us')

cities = []
locations = []

# 遍历邮政编码列表，获取城市名和位置
for postal_code in postal_codes:
    try:
        location = nomi.query_postal_code(str(postal_code)) #nomi 是一個 pgeocode.Nominatim 對象，設置為美國的郵政編碼查詢
        if location is not None and not location.empty:
            if not pd.isnull(location.place_name):
                cities.append(location.place_name)
                locations.append((location.latitude, location.longitude))
    except Exception as e:
        print(f"Failed to process postal code {postal_code}: {e}")

# 统计每个城市的出现次数
city_count = Counter(cities)

# 创建地图，中心设在美国中部
m = folium.Map(location=[39.8283, -98.5795], zoom_start=4)

# 添加标记
for city, location in zip(cities, locations):       #使用 zip() 函數同時迭代 cities 和 locations 列表，以便在地圖上為每個城市添加標記。
    count = city_count[city]
    folium.Marker(
        location,
        popup=f"{city} ({count})",
        tooltip=f"{city} ({count})"
    ).add_to(m)

# 保存地图为 HTML 文件
m.save("usa_city_map.html")

print("Map has been saved as usa_city_map.html")