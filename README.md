# 地址轉經緯度工具

利用 Google Maps Geocoding API，將 Excel 檔案中的地址批量轉換為經緯度座標。

## 功能特色

- 📍 批量地址轉經緯度
- 🔄 API 超限自動重試（指數退避）
- 📊 結果輸出為 Excel 檔案
- ❌ 錯誤原因逐筆記錄

## 前置需求

- Python 3.10+
- Google Maps Geocoding API 金鑰（[申請方式](https://developers.google.com/maps/documentation/geocoding/get-api-key)）

## 安裝

```bash
pip install pandas requests openpyxl
```

## 使用方式

### 基本用法

```bash
python geocode_addresses.py --key YOUR_API_KEY
```

預設讀取 `C:\Users\User\Downloads\地址轉換經緯度.xlsx`，輸出至同目錄下的 `地址轉換經緯度_結果.xlsx`。

### 完整參數

| 參數 | 縮寫 | 預設值 | 說明 |
|------|------|--------|------|
| `--input` | `-i` | `C:\Users\User\Downloads\地址轉換經緯度.xlsx` | 輸入 Excel 檔案路徑 |
| `--output` | `-o` | `{輸入檔名}_結果.xlsx` | 輸出 Excel 檔案路徑 |
| `--key` | `-k` | 無 | Google Maps API 金鑰 |
| `--column` | `-c` | `ADDR` | Excel 中的地址欄位名稱 |

### 範例

```bash
# 指定輸入輸出檔案與欄位名稱
python geocode_addresses.py -i data.xlsx -o result.xlsx -c 地址 -k YOUR_API_KEY
```

## 輸入格式

Excel 檔案需包含一個地址欄位（預設欄位名稱為 `ADDR`），例如：

| ADDR |
|------|
| 台北市信義區市府路1號 |
| 高雄市前鎮區成功二路25號 |

## 輸出格式

原始資料加上三個新欄位：

| ADDR | 經度 | 緯度 | 錯誤原因 |
|------|------|------|----------|
| 台北市信義區市府路1號 | 121.5654 | 25.0374 | |
| 無效地址 | | | ZERO_RESULTS: 無結果 |

## 注意事項

- Google Maps Geocoding API 為**付費服務**，請留意用量
- 每次 API 請求間隔 50ms，避免觸發頻率限制
- 超過查詢限制時會自動以指數退避重試（最多 3 次）
