# 利用Google Geocoding API 將地址資料轉換為經緯度資料
import argparse
import time

import pandas as pd
import requests

API_URL = "https://maps.googleapis.com/maps/api/geocode/json"
MAX_RETRIES = 3
BASE_DELAY = 1.0  # 指數退避的基礎延遲秒數


def geocode_address(
    address: str, api_key: str
) -> tuple[float | None, float | None, str | None]:
    """將單一地址轉換為經緯度，含重試機制。"""
    for attempt in range(MAX_RETRIES):
        response = requests.get(API_URL, params={"address": address, "key": api_key})

        if response.status_code != 200:
            return None, None, f"HTTP Error: {response.status_code}"

        data = response.json()
        status = data.get("status")

        if status == "OK" and data["results"]:
            location = data["results"][0]["geometry"]["location"]
            return location["lat"], location["lng"], None

        if status == "OVER_QUERY_LIMIT":
            delay = BASE_DELAY * (2**attempt)
            print(
                f"  ⚠ 超過查詢限制，{delay:.0f} 秒後重試 ({attempt + 1}/{MAX_RETRIES})..."
            )
            time.sleep(delay)
            continue

        # ZERO_RESULTS, REQUEST_DENIED, INVALID_REQUEST 等
        return None, None, f"{status}: {data.get('error_message', '無結果')}"

    return None, None, "OVER_QUERY_LIMIT: 已達最大重試次數"


def main():
    parser = argparse.ArgumentParser(description="地址批量轉經緯度工具")
    parser.add_argument(
        "--input",
        "-i",
        default=r"C:\Users\User\Downloads\地址轉換經緯度.xlsx",
        help="輸入 Excel 檔案路徑",
    )
    parser.add_argument(
        "--output",
        "-o",
        default=None,
        help="輸出 Excel 檔案路徑（預設為輸入檔名加 _結果）",
    )
    parser.add_argument("--key", "-k", default=None, help="Google Maps API 金鑰")
    parser.add_argument(
        "--column", "-c", default="S_ADDR2", help="地址欄位名稱（預設: S_ADDR2）"
    )
    args = parser.parse_args()

    # API 金鑰：命令列參數 > 這裡手動填入
    api_key = args.key or "在這裡填入你的API金鑰"
    if api_key == "在這裡填入你的API金鑰":
        print("❌ 請先設定 Google Maps API 金鑰！")
        print("   方式一：直接修改程式碼中的 api_key")
        print("   方式二：執行時加上 --key YOUR_API_KEY")
        return

    # 路徑處理
    input_path = args.input
    if args.output:
        output_path = args.output
    else:
        stem = input_path.rsplit(".", 1)[0]
        output_path = f"{stem}_結果.xlsx"

    # 讀取 Excel
    print(f"📂 讀取檔案: {input_path}")
    df = pd.read_excel(input_path)

    if args.column not in df.columns:
        print(f"❌ 找不到欄位 '{args.column}'，可用欄位: {list(df.columns)}")
        return

    total = len(df)
    print(f"📍 共 {total} 筆地址，開始轉換...\n")

    df["經度"] = None
    df["緯度"] = None
    df["錯誤原因"] = None

    success_count = 0
    for index, row in df.iterrows():
        address = row[args.column]
        lat, lng, error = geocode_address(address, api_key)

        df.at[index, "經度"] = lng
        df.at[index, "緯度"] = lat
        df.at[index, "錯誤原因"] = error

        i = index + 1
        if error:
            print(f"  [{i}/{total}] ✗ {address} → {error}")
        else:
            print(f"  [{i}/{total}] ✓ {address} → ({lat}, {lng})")
            success_count += 1

        time.sleep(0.05)

    # 寫入結果
    df.to_excel(output_path, index=False)
    print(f"\n✅ 完成！成功 {success_count}/{total} 筆")
    print(f"📄 結果已儲存至: {output_path}")


if __name__ == "__main__":
    main()
