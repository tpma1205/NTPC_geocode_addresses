# 利用Google Geocoding API 將地址資料轉換為經緯度資料
import time

import pandas as pd
import requests


def geocode_address(address, api_key):
    url = f"https://maps.googleapis.com/maps/api/geocode/json?address={address}&key={api_key}"
    response = requests.get(url)
    if response.status_code == 200:
        result = response.json()
        if result["results"]:
            location = result["results"][0]["geometry"]["location"]
            return location["lat"], location["lng"], None
        else:
            return None, None, "No results found"
    else:
        return None, None, f"Error: {response.status_code}"


# 設定Google Maps API金鑰
api_key = "[ENCRYPTION_KEY]"

# 讀取Excel檔案
file_path = "C:\\Users\\User\\Downloads\\地址轉換經緯度.xlsx"
addresses = pd.read_excel(file_path)

# 初始化新的欄位
addresses["經度"] = None
addresses["緯度"] = None
addresses["錯誤原因"] = None

# 為每個地址獲取經緯度
for index, row in addresses.iterrows():
    lat, lng, error = geocode_address(row["S_ADDR2"], api_key)
    addresses.at[index, "經度"] = lng if not error else None
    addresses.at[index, "緯度"] = lat if not error else None
    addresses.at[index, "錯誤原因"] = error
    time.sleep(0.05)  # 增加延遲時間 避免過度調用

# 將結果寫入Excel檔案
output_file_path = "C:\\Users\\User\\Downloads\\地址轉換經緯度_結果.xlsx"
addresses.to_excel(output_file_path, index=False)
