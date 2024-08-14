import tkinter as tk                            # tkinter 是 Python 的一個標準 GUI (圖形用戶界面) 庫，用於創建桌面應用的窗口和控件
from tkinter import filedialog                  # 從 tkinter 引入 filedialog, 用於打開文件或保存文件
import re                                       # re模塊啟用正則表達式
import pandas as pd                             # pandas 是一个用於数据操作和分析的 Python 庫, 用於讀取 Excel 文件
import pgeocode                                 # pgeocode 是一個 Python 庫, 用于從郵政編碼中查詢地理信息
import folium                                   # folium 是一个基于 Leaflet.js 的 Python 库，用于通过代码创建并保存带有标记、弹出窗口等交互元素的地图
from branca.element import IFrame               # branca 是 folium 的一个工具库，允许创建各种 HTML 元素，以增强地图的交互性。
from collections import defaultdict             # Python Collection 中導入 defaultdict 
                                                # defaultdict: 一種字典類, 可自定義字典中的鍵提供默認值, 當訪問不存在的鍵時, 不會拋出 KeyError, 而是返回默認值

# 使用 re 模块的 findall 方法从地址中查找一个或多个5位数邮政编码，返回最后一个或 None。
# `r`: 告诉 Python 这是一个原始字符串，这意味着字符串内的转义字符将不被处理（例如，\不会被视为转义字符）。
# `\b`: 单词边界，确保数字序列位于单词的开始或结束。
# `\d{5}`: 匹配五个连续的数字（\d 代表 [0-9] 范围内的任意数字）。
# `\b`: 末尾的单词边界，确保数字序列后没有紧跟其他数字。
# `?=` : 正向前瞻(lookahead), 確保五位數字後是連字符或單詞邊界。
# `-`:表示數字緊接一個連字符
# `|`:邏輯或操作符號, 表示滿足前後任一條件即可
def extract_postal_code(address):
    #postal_codes = re.findall(r'\b\d{5}\b', address) 
    postal_codes = re.findall(r'\d{5}(?=-|\b)', address)
    return postal_codes[-1] if postal_codes else None

# 透過folium裡的branca工具庫, 引用IFrame控制顯示內容
# 參數:
#  content: 要顯示在彈窗中的html內容
#  width: 彈窗基本寬度
#  height: 彈窗基本高度 
def create_large_popup(content, width=250, height=100):
    iframe = IFrame(html=content, width=width+20, height=height+20)  # 加20為padding
    return folium.Popup(iframe, max_width=width+20)    #創件並返回包含上述IFrame的folium.Popup對象


def load_data():
    root = tk.Tk()  # 創建一個窗口實例
    root.withdraw()  # 不顯示主窗口
    file_path = filedialog.askopenfilename(title="選擇 Excel 文件", filetypes=[("Excel files", "*.xlsx *.xls")])    # 來自 tkinter 庫中的 filedialog 模塊, filedialog.askopenfilename() 彈出一個對話框，讓用戶選擇本地 Excel 文件。
    if not file_path:
        print("No file selected.")
        return

    df = pd.read_excel(file_path)   # 讀取Excel文件, 將其加載為一個 pandas 的 DataFrame 對象

    # 从地址中提取邮政编码
    df['PostalCode'] = df['Address'].dropna().apply(extract_postal_code) # 訪問 DataFrame 中的 Address 列, 用 .dropna() 刪除列中 Nan 行, 函數 extract_postal_code 提取出 ZipCode 放在一個新的列
    postal_codes = df['PostalCode'].dropna().tolist()   # 將上述列轉換為 Python 列表

    #  pgeocode庫初始化 Nominatim對象, 用來查询 ZipCode 返回相關地理信息(例如 城市名,經緯度, 州/省)，设置为美国
    nomi = pgeocode.Nominatim('us')

    city_to_locations = defaultdict(lambda: None)  #  創建一個 defaultdict (Python 的 collections 模塊中的一種特殊字典) 每個鍵對應程式名, 值則是該城市相關地理位置, 訪問不存在的鍵會自動生成默認值為None的項
    city_to_zipcodes = defaultdict(list)  # 同上, 但字典訪問未添加的程式會自動初始化一個空列表作為該程式的值
    unprocessed_codes = []

    # 遍历邮政编码列表，获取城市名和位置, 這邊要注意到有可能會有不同zipcode返回相同城市名 所以不能用CityName作為唯一鍵 我這邊選擇用 CityName + Zipcode
    for postal_code in postal_codes:
        try:
            location = nomi.query_postal_code(str(postal_code))     # 使用 pgeocode 库中的 Nominatim 类查询给定 Zipcode 的地理信息
            if location is not None and not location.empty:
                if not pd.isnull(location.place_name):
                    city = location.place_name
                    state = location.state_name
                    unique_key = f"{city}, {state}"
                    if city_to_locations[unique_key] is None:
                        city_to_locations[unique_key] = (location.latitude, location.longitude)
                    city_to_zipcodes[unique_key].append(postal_code)
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
            content = f"<div style='font-size:12px;'><strong>{city} ({count})</strong><br>Zip Codes: {zipcodes}</div>"
            
            # 設定寬度和高度
            popup = create_large_popup(content, width=250, height=100)  
            
            # 使用HTML格式化的tooltip文字
            tooltip_text = f"<div style='font-size:18px; font-weight:bold;'>{city} ({count})</div>"
            
            folium.Marker(
                location,
                popup=popup,
                tooltip=tooltip_text
            ).add_to(m)

    # filedialog.asksaveasfilename()保存地图为 HTML 文件
    map_file_path = filedialog.asksaveasfilename(title="保存地圖文件", filetypes=[("HTML files", "*.html")], defaultextension=".html")
    if map_file_path:
        m.save(map_file_path)
        print(f"Map has been saved as {map_file_path}")
    else:
        print("Map save cancelled.")

# if __name__ == "__main__": 是 Python 中的慣用法，用來判斷當前模組是否被直接執行。
# 如果是直接執行，__name__ 的值為 "__main__"，程式會執行這段區塊的代碼。
# 如果模組是被導入到其他模組中，__name__ 的值會是模組的名稱，這段代碼則不會被執行。
if __name__ == "__main__":
    load_data()

# 最後使用 PyInstaller 將 Python 腳本轉換為可執行文件 .exe:
# 1. 安裝 PyInstaller         pip install pyinstaller
# 2. 打包生成 .exe 文件        python -m PyInstaller --onefile --windowed C:\Users\Benson\Desktop\github-storage\zipCityMapTool\PythonApplication1\PythonApplication1.py  
#                            -- onefile：将所有依賴項打包到一个单独的可执行文件中。
#                            -- windowed：生成的可执行文件将不会打开命令行窗口（适用于 GUI 应用程序）。 
#                            -- ...PythonApplication1.py : 此Python腳本的絕對路徑