# [越南雲端數據中台 / OVS-LX-VDO-01]
# [越南消金新增 Google Maps 地理空間數據 - Geocoding API]
# SA / SD 系統分析與設計文件

---

## 文件制定／修訂履歷

| 制定／修訂版次 | 制定／修訂日期 | 制定／修訂說明 | 作者 | 備註 |
|---|---|---|---|---|
| v0.1 | 2026/02/04 | 初版 | Denny | |
| v0.2 | 2026/02/12 | 調整並完善規格 | Tom | |

---

## 目錄

- [第一章、功能概述 (SA)](#第一章功能概述-sa)
  - [1.1 專案/功能說明](#11-專案功能說明)
  - [1.2 系統/服務流程](#12-系統服務流程)
- [第二章、API 規格 (SA)](#第二章api-規格-sa)
  - [2.1 基本資訊](#21-基本資訊)
  - [2.2 上行/請求 API 規格](#22-上行請求-api-規格)
  - [2.3 上行/請求 API 範例](#23-上行請求-api-範例)
  - [2.4 下行/回應 API 規格](#24-下行回應-api-規格)
  - [2.5 下行/回應 API 範例](#25-下行回應-api-範例)
  - [2.6 處理結果代碼](#26-處理結果代碼)
- [第三章、處理邏輯 (SA)](#第三章處理邏輯-sa)
  - [3.1 程式處理流程](#31-程式處理流程)
  - [3.2 檢核下行](#32-檢核下行)
  - [3.3 儲存下行 (GCS)](#33-儲存下行-gcs)
  - [3.4 儲存下行 (BigQuery)](#34-儲存下行-bigquery)
  - [3.5 處理結果代碼](#35-處理結果代碼)
- [第四章、雲端資源配置與設計 (SD)](#第四章雲端資源配置與設計-sd)
  - [4.1 GCP 資源申請表](#41-gcp-資源申請表)
  - [4.2 服務架構暨程式流程圖](#42-服務架構暨程式流程圖)
- [第五章、非功能性需求設計 (SD)](#第五章非功能性需求設計-sd)
- [第六章、附錄](#第六章附錄)

---

## 第一章、功能概述 (SA)

### 1.1 專案/功能說明

**API 設計目的與業務應用場景**

串接 Google Maps 地理空間數據服務，作為越南消金業務後續風險分析使用。  
本階段目標在於透過 Google Maps API 將地址轉換成經緯度定位、經緯度轉換成地址定位，進一步識別高風險、違約率偏高的用戶群所屬詳細地點，並以此作為後續驗證團夥詐欺行為的重要依據。

**關聯資料來源**

| 資料來源 | 說明 |
|---|---|
| Google Maps Geocoding API | 外部地理空間服務，提供地址 ↔ 座標轉換 |
| BigQuery - HES | `RAW_HES_DATASET.APPLICATION`、`RAW_HES_DATASET.CUSTOMER`，存有用戶申貸資訊 |
| BigQuery - VMB | `RAW_VMB_DATASET.APPLY_INFO`，存有申貸時的簽約經緯度 |
| BigQuery - EDEP | `RAW_EDEP_DATASET.GEOCODING`，本服務的發查結果儲存目的地 |
| GCS | `gs://ovslxvdo01-{env}-rawdata-api/GOOGLEMAPS/GEOCODING/`，原始 JSON 回應儲存位置 |

**服務性質**

內部使用（東南亞數據團隊開發人員與風險分析師）。

---

### 1.2 系統/服務流程

本服務執行三類地理空間發查，處理流程如下：

```
┌──────────────────────────────────────────────────────────────┐
│  Cloud Run Job (排程觸發，每日執行)                           │
│                                                              │
│  1. 從 BigQuery HES / VMB 取出未發查的客戶資料               │
│     ├─ 通訊地址 (CONTACT_ADDRESS)                            │
│     ├─ 戶籍地址 (RESIDENCE_ADDRESS)                          │
│     └─ 簽約經緯度 (CONTRACT_COORDINATES)                     │
│                                                              │
│  2. 呼叫 Google Maps Geocoding API                           │
│     ├─ address → latlng（地址轉座標）                        │
│     └─ latlng → address（座標轉地址）                        │
│                                                              │
│  3. 下行檢核（格式驗證、必填欄位補全）                        │
│                                                              │
│  4. 儲存結果                                                 │
│     ├─ GCS：原始 JSON 格式                                   │
│     └─ BigQuery：RAW_EDEP_DATASET.GEOCODING                  │
│                                                              │
│  5. 更新已發查紀錄（避免重複發查）                            │
└──────────────────────────────────────────────────────────────┘
```

---

## 第二章、API 規格 (SA)

### 2.1 基本資訊

| 項目 | 內容 |
|---|---|
| API 名稱 | Google Maps : Geocoding API |
| API URI | `https://maps.googleapis.com/maps/api/geocode/json?` |
| HTTP Method | GET |
| Content-Type | application/json |
| 驗證方式 | API Key（由 GCP Secret Manager 管理） |
| 功能描述 | 地址 ↔ 座標轉換、取得詳細地址定位 |
| API 官方文件 | [Google Maps Geocoding API 說明](https://developers.google.com/maps/documentation/geocoding/requests-geocoding?hl=zh-tw) |

---

### 2.2 上行/請求 API 規格

> **注意：** `address` 與 `latlng` 二擇一必填。

| LVL | 欄位名稱 | 資料型態 | 必填 | 最大長度 | 說明 |
|---|---|---|---|---|---|
| 1 | Content-Type | String | N | — | 預設為 `application/json` |
| 1 | key | String | Y | — | Google API 金鑰（由 Secret Manager 注入） |
| 1 | latlng | String | 條件必填 | — | 經緯度，格式為 `"latitude,longitude"` |
| 1 | address | String | 條件必填 | — | 要查詢的地址（URL encoded） |

---

### 2.3 上行/請求 API 範例

**輸入經緯度（座標 → 地址）**

```
GET https://maps.googleapis.com/maps/api/geocode/json?latlng={latlng}&key={API_KEY}
```

範例：
```
GET https://maps.googleapis.com/maps/api/geocode/json?latlng=10.8231,106.6297&key=YOUR_API_KEY
```

**輸入地址（地址 → 座標）**

```
GET https://maps.googleapis.com/maps/api/geocode/json?address={address}&key={API_KEY}
```

範例：
```
GET https://maps.googleapis.com/maps/api/geocode/json?address=268+Ly+Thuong+Kiet,+Ho+Chi+Minh&key=YOUR_API_KEY
```

---

### 2.4 下行/回應 API 規格

| LVL | 欄位名稱 | 資料型態 | 說明 |
|---|---|---|---|
| 1 | status | string | 請求狀態代碼（見 2.6） |
| 1 | place_id | string | 地點唯一識別碼 |
| 1 | formatted_address | string | 完整格式化地址 |
| 1 | types | list | 地點類型列表 |
| 1 | address_components | list | 地址組成元件列表 |
| 2 | long_name | string | 地址元件全名 |
| 2 | short_name | string | 地址元件縮寫 |
| 2 | types | list | 元件類型（如 `country`、`administrative_area_level_1`） |
| 1 | geometry | dict | 地理幾何資訊 |
| 2 | location | dict | 座標位置 |
| 3 | lat | float | 緯度 |
| 3 | lng | float | 經度 |
| 1 | plus_code | dict | Plus Code 資訊 |
| 2 | global_code | string | 全域 Plus Code |

---

### 2.5 下行/回應 API 範例

```json
[
    {
        "status": "OK",
        "place_id": "ChIJxxxxxxxxxxxxxxxx",
        "formatted_address": "268 Lý Thường Kiệt, Phường 14, Quận 10, Thành phố Hồ Chí Minh, Vietnam",
        "types": ["street_address"],
        "address_components": [
            {
                "long_name": "268",
                "short_name": "268",
                "types": ["street_number"]
            },
            {
                "long_name": "Lý Thường Kiệt",
                "short_name": "Lý Thường Kiệt",
                "types": ["route"]
            },
            {
                "long_name": "Phường 14",
                "short_name": "P. 14",
                "types": ["sublocality_level_1", "sublocality", "political"]
            },
            {
                "long_name": "Quận 10",
                "short_name": "Q. 10",
                "types": ["administrative_area_level_2", "political"]
            },
            {
                "long_name": "Thành phố Hồ Chí Minh",
                "short_name": "TP. HCM",
                "types": ["administrative_area_level_1", "political"]
            },
            {
                "long_name": "Vietnam",
                "short_name": "VN",
                "types": ["country", "political"]
            }
        ],
        "geometry": {
            "location": {
                "lat": 10.7734,
                "lng": 106.6669
            }
        },
        "plus_code": {
            "global_code": "7P28QPX7+QX"
        }
    }
]
```

---

### 2.6 處理結果代碼

| RETURNCODE | RETURNDESC | 回傳時機 | 建議處理方式 |
|---|---|---|---|
| `OK` | 請求成功，回傳有效結果 | 正常回應 | 繼續儲存下行 |
| `ZERO_RESULTS` | 查詢成功，但沒有符合的結果 | 地址無效或過於模糊 | 記錄 log，不儲存，標記為不可發查 |
| `OVER_QUERY_LIMIT` | 已超過每日或每秒的配額限制 | 超出 API 配額 | 指數退避重試（最多 3 次），超過後延至次日 |
| `REQUEST_DENIED` | 請求被拒絕，通常是 API Key 無效或權限不足 | API Key 異常 | 立即告警（PagerDuty / Slack），停止批次 |
| `INVALID_REQUEST` | 請求參數錯誤或缺少必要欄位 | 上行資料缺少 address 或 latlng | 記錄 log，跳過該筆，繼續下一筆 |
| `UNKNOWN_ERROR` | 伺服器端錯誤，請稍後重試 | Google 服務端臨時異常 | 固定間隔重試（30 秒），最多 3 次 |

---

## 第三章、處理邏輯 (SA)

### 3.1 程式處理流程

執行 Geocoding 批次作業時，判斷邏輯如下：

**Step 1 — 取得客戶基本資訊**

取出所有客戶資訊，以下欄位皆不能為 `NULL` 或字串 `"null"`（每筆皆取最新）：

| 資料類型 | 來源欄位 | 發查類型（GEO_TYPE） |
|---|---|---|
| 通訊地址 | `HES.customer.CURRENT_DETAILED_ADDRESS` | `CONTACT_ADDRESS` |
| 戶籍地址 | `HES.customer.PERMANENT_DETAILED_ADDRESS` | `RESIDENCE_ADDRESS` |
| 簽約經緯度 | `VMB.apply_info.LONGITUDE` / `VMB.apply_info.LATITUDE` | `CONTRACT_COORDINATES` |
| 申貸序號 | `HES.application.SERIAL_NUMBER` | — |
| 客戶 ID | `HES.customer.CUID` | — |

**Step 2 — 篩選未發查資料**

分別依照附錄一、二、三的 SQL 邏輯，篩選出尚未出現在 `RAW_EDEP_DATASET.GEOCODING` 的資料列進行發查。

**Step 3 — 呼叫 Google Maps Geocoding API**

- 通訊地址 / 戶籍地址：以 `address` 欄位發查
- 簽約經緯度：以 `latlng` 欄位發查（格式：`latitude,longitude`，取小數點後 4 位）

**Step 4 — 更新已發查資料**

依照附錄四、五、六的 SQL 邏輯，將已完成發查的資料寫回 `RAW_EDEP_DATASET.GEOCODING`。

---

### 3.2 檢核下行

可參閱 [附錄七 — 下行檢核程式碼](#附錄七-下行檢核程式碼)，並依實際情形調整。

主要驗證項目：
- `status` 為 `OK` 才進行儲存
- `geometry.location.lat` / `lng` 必須為合法的浮點數
- `formatted_address` 不得為空字串
- `address_components` 必須包含至少一個元件

---

### 3.3 儲存下行 (GCS)

每個 Geocoding 發查完的原始 JSON 回應，儲存至 GCS：

```
gs://ovslxvdo01-{env}-rawdata-api/GOOGLEMAPS/GEOCODING/{data_date}_{serial_number}.json
```

| 參數 | 說明 |
|---|---|
| `{env}` | 環境代碼，如 `dev`、`stg`、`prd` |
| `{data_date}` | 發查日期，格式 `YYYYMMDD` |
| `{serial_number}` | `HES.application.SERIAL_NUMBER` |

---

### 3.4 儲存下行 (BigQuery)

目標資料表：`RAW_EDEP_DATASET.GEOCODING`

| 欄位 | 資料來源 | 欄位說明 |
|---|---|---|
| DATA_DATE | BQ 預設值 `CURRENT_DATE()` | 發查日期 |
| CUID | `HES.customer.CUID` | 客戶唯一識別碼 |
| SERIAL_NUMBER | `HES.application.SERIAL_NUMBER` | 申貸序號 |
| CREATED_AT | `HES.application.CREATED_AT` | 申貸建立時間 |
| GEO_TYPE | — | 發查類型：`RESIDENCE_ADDRESS` / `CONTACT_ADDRESS` / `CONTRACT_COORDINATES` |
| REQUEST_LONGITUDE | `VMB.apply_info.LONGITUDE`（僅 CONTRACT_COORDINATES），取小數點後 4 位，其他為 NULL | 請求經度 |
| REQUEST_LATITUDE | `VMB.apply_info.LATITUDE`（僅 CONTRACT_COORDINATES），取小數點後 4 位，其他為 NULL | 請求緯度 |
| REQUEST_ADDRESS | 依 GEO_TYPE：RESIDENCE→`PERMANENT_DETAILED_ADDRESS`；CONTACT→`CURRENT_DETAILED_ADDRESS`；CONTRACT→NULL | 請求地址 |
| RESPONSE_PLACE_ID | 下行 `place_id`，取不到為 NULL | 回應地點 ID |
| RESPONSE_ADRESS | 下行 `formatted_address`，取不到為 NULL | 回應完整地址 |
| RESPONSE_GLOBAL_CODE | 下行 `plus_code.global_code`，取不到為 NULL | 回應 Plus Code |
| RESPONSE_PLACE_TYPES | 下行 `types`（LIST），以逗號合併為字串，取不到為 NULL | 回應地點類型 |
| RESPONSE_LONGITUDE | 下行 `geometry.location.lng`，取不到為 NULL | 回應經度 |
| RESPONSE_LATITUDE | 下行 `geometry.location.lat`，取不到為 NULL | 回應緯度 |
| RESPONSE_COUNTRY | `address_components` 中 `types` 含 `country` 的 `long_name`，取不到為 NULL | 回應國家 |
| RESPONSE_CITY | `address_components` 中 `types` 含 `administrative_area_level_1` 的 `long_name`，取不到為 NULL | 回應城市/省份 |
| RESPONSE_DISTRICT | `address_components` 中 `types` 含 `administrative_area_level_2` 的 `long_name`，取不到為 NULL | 回應縣市/區 |
| RESPONSE_WARD | `address_components` 中 `types` 含 `sublocality_level_1` 的 `long_name`，取不到為 NULL | 回應鄉鎮/街道 |
| RESPONSE_STREET | `address_components` 中 `types` 含 `route` 的 `long_name`，取不到為 NULL | 回應路名 |
| BQ_CREATED_TIME | BQ 預設值 `CURRENT_TIMESTAMP()` | 建立時間 |
| BQ_UPDATED_TIME | BQ 預設值 `CURRENT_TIMESTAMP()` | 最後更新時間 |

---

### 3.5 處理結果代碼

同 [2.6 處理結果代碼](#26-處理結果代碼)。

---

## 第四章、雲端資源配置與設計 (SD)

### 4.1 GCP 資源申請表

> **說明：** 以下為本服務所需的 GCP 資源規格，需向 GCP 管理員申請並確認環境設定。

#### Cloud Run Job

| 項目 | 規格 | 備註 |
|---|---|---|
| 服務名稱 | `geocoding-job` | 依環境加上 `-dev` / `-stg` / `-prd` 後綴 |
| 映像檔來源 | GCR / Artifact Registry | `{region}-docker.pkg.dev/{project}/geocoding/geocoding-job:{tag}` |
| 執行方式 | Cloud Run Job（排程觸發） | |
| 排程 | 每日 01:00 VNT（UTC+7）| Cloud Scheduler 設定 |
| 記憶體 | 2 GiB | 批次處理大量資料時的基本需求 |
| CPU | 2 vCPU | |
| 逾時時間 | 3600 秒（1 小時） | 視資料量可調整至 7200 秒 |
| 最大重試次數 | 3 | 失敗後自動重試 |
| 執行環境 | 第二代執行環境 | |
| VPC 連接 | 需連接至 Private VPC | 用於存取內部 BigQuery / GCS |

#### Secret Manager

| 金鑰名稱 | 說明 | 存取帳號 |
|---|---|---|
| `googlemaps-api-key` | Google Maps Geocoding API 金鑰 | `geocoding-sa@{project}.iam.gserviceaccount.com` |

#### Service Account

| 項目 | 規格 |
|---|---|
| 帳號名稱 | `geocoding-sa@{project}.iam.gserviceaccount.com` |
| 角色 | `roles/bigquery.dataEditor`（EDEP dataset）、`roles/bigquery.dataViewer`（HES、VMB dataset）、`roles/storage.objectCreator`（GCS bucket）、`roles/secretmanager.secretAccessor` |

#### BigQuery

| 項目 | 規格 | 備註 |
|---|---|---|
| Dataset | `RAW_EDEP_DATASET` | 已存在，需申請寫入權限 |
| Table | `GEOCODING` | 若不存在需依 3.4 欄位定義建立 |
| 資料分區 | 依 `DATA_DATE` 分區 | 建議使用 DATE 分區提升查詢效能 |
| 資料保留 | 依資料治理規範（建議 3 年） | |

#### GCS

| 項目 | 規格 | 備註 |
|---|---|---|
| Bucket | `ovslxvdo01-{env}-rawdata-api` | 已存在，需申請 objectCreator 權限 |
| 路徑 | `GOOGLEMAPS/GEOCODING/` | |
| 物件命名格式 | `{data_date}_{serial_number}.json` | |
| 存取控制 | Uniform bucket-level access | |
| 資料保留 | 依資料治理規範（建議 90 天後轉 Nearline） | |

---

### 4.2 服務架構暨程式流程圖

```
                        ┌──────────────────────────────────────┐
                        │        Cloud Scheduler（每日 01:00）  │
                        └──────────────────┬───────────────────┘
                                           │ 觸發
                                           ▼
                        ┌──────────────────────────────────────┐
                        │         Cloud Run Job                 │
                        │         geocoding-job                 │
                        │                                       │
  ┌─────────────┐  查詢  │  Step 1: 取未發查客戶清單             │
  │  BigQuery   │◄──────│  (HES application/customer, VMB)     │
  │  HES / VMB  │──────►│                                       │
  └─────────────┘  回傳  │  Step 2: 迴圈發查每筆資料             │
                        │         └─ 呼叫 Google Maps API      │
  ┌─────────────┐        │                                       │
  │Secret Mgr   │──────►│  API Key 注入                         │
  └─────────────┘        │                                       │
                        │  Step 3: 檢核下行                     │
  ┌─────────────┐  儲存  │                                       │
  │    GCS      │◄──────│  Step 4a: 儲存原始 JSON 至 GCS        │
  └─────────────┘        │                                       │
  ┌─────────────┐  寫入  │  Step 4b: 寫入結構化欄位至 BigQuery   │
  │  BigQuery   │◄──────│  (RAW_EDEP_DATASET.GEOCODING)         │
  │    EDEP     │        │                                       │
  └─────────────┘        └──────────────────────────────────────┘
                                           │
                                    呼叫外部API
                                           ▼
                        ┌──────────────────────────────────────┐
                        │     Google Maps Geocoding API         │
                        │  maps.googleapis.com/maps/api/geocode │
                        └──────────────────────────────────────┘
```

---

## 第五章、非功能性需求設計 (SD)

### 5.1 效能需求

| 項目 | 目標值 | 說明 |
|---|---|---|
| 每次作業處理筆數 | 無上限（依當日新增資料量） | 批次全量處理 |
| 單筆 API 呼叫回應時間 | ≤ 500 ms（P95） | 依 Google Maps SLA |
| 每日作業完成時間 | ≤ 2 小時 | 需於 03:00 VNT 前完成，供早盤風控分析使用 |
| API 呼叫速率 | ≤ 50 QPS | 遵循 Google Maps 免費配額上限，超出需調整 |

> **配額管理：** 每日 Geocoding API 呼叫上限依 GCP 專案設定，預設免費額度為 40,000 次/月（超出按量計費）。若每日資料量超過 1,300 筆，需評估升級為付費方案或申請配額提升。

> **效能可行性驗算：**
> - 作業時間限制：01:00 → 03:00 VNT = **2 小時 = 7,200 秒**
> - QPS 上限：50 次/秒
> - 理論最大處理量：50 QPS × 7,200 秒 = **360,000 筆/日**（已扣除網路延遲估算需降至 70% = 約 252,000 筆）
> - 若每日新增資料遠超此上限，需評估：a) 提高 QPS 配額（付費方案），或 b) 分批執行（多個 Cloud Run Job 並行）

---

### 5.2 可用性與可靠性

| 項目 | 目標值 | 說明 |
|---|---|---|
| 服務可用性 | ≥ 99.0%（月度計算） | Cloud Run Job 不常駐，無需 SLA 監控；可用性以「排程作業成功率」衡量 |
| 作業成功率 | ≥ 99.5%（每日） | 單日作業失敗率不超過 0.5% |
| 失敗重試機制 | Cloud Run Job 最多重試 3 次 | 適用於整個 Job 失敗的情況 |

---

### 5.3 錯誤處理策略

| 錯誤類型 | 重試策略 | 告警 | 備註 |
|---|---|---|---|
| `OVER_QUERY_LIMIT` | 指數退避（1s → 2s → 4s），最多 3 次 | 若連續 3 次失敗，送出告警 | 避免在每秒配額密集期發送 |
| `UNKNOWN_ERROR` | 固定間隔重試（30s），最多 3 次 | 若連續 3 次失敗，記錄 log 並跳過 | Google 服務端臨時問題 |
| `REQUEST_DENIED` | 不重試 | 立即告警（高優先級），停止整批作業 | API Key 失效或配額耗盡 |
| `INVALID_REQUEST` | 不重試 | 記錄 log，標記該筆為無效 | 上行資料問題，非 API 問題 |
| `ZERO_RESULTS` | 不重試 | 記錄至 log，不寫入 BigQuery | 地址無法解析 |
| 網路逾時 | 固定間隔重試（10s），最多 3 次 | 若連續 3 次失敗，記錄 log 並跳過 | 設定 API 呼叫 timeout = 10s |

---

### 5.4 快取與去重策略

為避免對同一地址或座標重複呼叫 Google Maps API（節省配額與費用），已在 SQL 邏輯中實作去重：

- **通訊地址 / 戶籍地址：** 若 `REQUEST_ADDRESS` 已存在於 `RAW_EDEP_DATASET.GEOCODING`，則跳過該筆（見附錄一、二 SQL `LEFT JOIN` 條件）
- **簽約經緯度：** 若 `REQUEST_LONGITUDE` + `REQUEST_LATITUDE` 組合已存在，則跳過（見附錄三 SQL）
- **SERIAL_NUMBER 去重：** 同一申貸序號若已有任一類型的發查紀錄，亦不重複發查

---

### 5.5 安全性

| 項目 | 措施 |
|---|---|
| API Key 管理 | 儲存於 GCP Secret Manager，不硬編碼於程式碼或環境變數 |
| 最小權限原則 | Service Account 僅授予必要的 BigQuery / GCS / Secret Manager 角色 |
| 資料傳輸加密 | 所有 Google Maps API 呼叫強制使用 HTTPS |
| 內網存取 | Cloud Run Job 透過 VPC Connector 存取 BigQuery / GCS，不開放公網 |
| API Key 輪換 | 建議每 90 天輪換一次，輪換期間使用雙金鑰策略避免服務中斷 |

---

### 5.6 監控與日誌

| 項目 | 工具 | 監控指標 |
|---|---|---|
| 作業成功/失敗 | Cloud Monitoring + Cloud Run Job 狀態 | Job 執行狀態、失敗次數 |
| API 呼叫量 | Cloud Monitoring | 每日 Geocoding API 呼叫次數、配額使用率 |
| 錯誤率 | Cloud Logging | `REQUEST_DENIED`、`OVER_QUERY_LIMIT` 錯誤頻率 |
| 資料寫入量 | BigQuery 查詢 | 每日新增至 `GEOCODING` 的資料筆數 |
| 告警通知 | Cloud Monitoring Alerts → Slack / Email | `REQUEST_DENIED` 立即告警；連續失敗超過閾值時告警 |

---

## 附錄零 — 術語表（Glossary）

| 術語 | 全名 / 說明 |
|---|---|
| **SASD** | System Analysis & System Design，系統分析與設計文件 |
| **HES** | Hire Enterprise System，越南消金核心貸款系統 |
| **VMB** | Vietnam Mobile Banking，越南行動銀行申貸系統 |
| **EDEP** | External Data Exchange Platform，外部數據交換平台（本服務的結果儲存位置） |
| **CUID** | Customer Unique ID，客戶唯一識別碼 |
| **SERIAL_NUMBER** | 申貸序號，來自 `HES.APPLICATION.SERIAL_NUMBER`，唯一識別每一筆申貸案件 |
| **GEO_TYPE** | Geocoding 發查類型：`CONTACT_ADDRESS`（通訊地址）/ `RESIDENCE_ADDRESS`（戶籍地址）/ `CONTRACT_COORDINATES`（簽約經緯度） |
| **Geocoding** | 地理編碼，將地址字串轉換為經緯度座標 |
| **Reverse Geocoding** | 反向地理編碼，將經緯度座標轉換為地址字串 |
| **QPS** | Queries Per Second，每秒查詢次數（API 呼叫速率上限單位） |
| **VNT** | Vietnam Time，越南標準時間（UTC+7） |
| **env** | 環境代碼：`dev`（開發）/ `stg`（測試）/ `prd`（正式） |
| **RESPONSE_ADRESS** | 文件保留原始欄位名稱（注意：少一個 d），與現有資料表 Schema 保持一致；如需修正請同步更新 BQ Schema |

---

## 附錄零 b — BigQuery Schema DDL

```sql
CREATE TABLE IF NOT EXISTS `RAW_EDEP_DATASET.GEOCODING` (
    DATA_DATE           DATE          NOT NULL OPTIONS(description="發查日期，BQ 預設 CURRENT_DATE()"),
    CUID                STRING        NOT NULL OPTIONS(description="客戶唯一識別碼"),
    SERIAL_NUMBER       STRING        NOT NULL OPTIONS(description="申貸序號"),
    CREATED_AT          TIMESTAMP              OPTIONS(description="申貸建立時間，來自 HES.APPLICATION.CREATED_AT"),
    GEO_TYPE            STRING        NOT NULL OPTIONS(description="發查類型: CONTACT_ADDRESS / RESIDENCE_ADDRESS / CONTRACT_COORDINATES"),
    REQUEST_LONGITUDE   FLOAT64                OPTIONS(description="請求經度，僅 CONTRACT_COORDINATES 有值，取小數點後4位"),
    REQUEST_LATITUDE    FLOAT64                OPTIONS(description="請求緯度，僅 CONTRACT_COORDINATES 有值，取小數點後4位"),
    REQUEST_ADDRESS     STRING                 OPTIONS(description="請求地址，地址類型發查時有值"),
    RESPONSE_PLACE_ID   STRING                 OPTIONS(description="Google Maps place_id"),
    RESPONSE_ADRESS     STRING                 OPTIONS(description="完整格式化地址（formatted_address）"),
    RESPONSE_GLOBAL_CODE STRING               OPTIONS(description="Plus Code global_code"),
    RESPONSE_PLACE_TYPES STRING               OPTIONS(description="types LIST，以逗號合併為字串"),
    RESPONSE_LONGITUDE  FLOAT64                OPTIONS(description="回應經度 geometry.location.lng"),
    RESPONSE_LATITUDE   FLOAT64                OPTIONS(description="回應緯度 geometry.location.lat"),
    RESPONSE_COUNTRY    STRING                 OPTIONS(description="國家名稱（address_components type=country）"),
    RESPONSE_CITY       STRING                 OPTIONS(description="城市/省份（administrative_area_level_1）"),
    RESPONSE_DISTRICT   STRING                 OPTIONS(description="縣市/區（administrative_area_level_2）"),
    RESPONSE_WARD       STRING                 OPTIONS(description="鄉鎮/街道（sublocality_level_1）"),
    RESPONSE_STREET     STRING                 OPTIONS(description="路名（route）"),
    BQ_CREATED_TIME     TIMESTAMP              OPTIONS(description="BQ 建立時間，預設 CURRENT_TIMESTAMP()"),
    BQ_UPDATED_TIME     TIMESTAMP              OPTIONS(description="BQ 更新時間，預設 CURRENT_TIMESTAMP()")
)
PARTITION BY DATA_DATE
OPTIONS(
    description="Google Maps Geocoding 發查結果資料表",
    partition_expiration_days=1095  -- 3 年保留
);
```

---

## 第六章、附錄

### 附錄一 — 通訊地址發查

```sql
-- 通訊地址：取出尚未發查的客戶通訊地址
WITH latest_hes_application AS (
    SELECT id,
           serial_number,
           created_at,
           customer_id,
      FROM `RAW_HES_DATASET.APPLICATION`
   QUALIFY ROW_NUMBER() OVER (
           PARTITION BY id
            ORDER BY TIMESTAMP(created_at) DESC,
                     SAFE_CAST(finished_at AS TIMESTAMP) DESC,
                     PARTITION_DATE DESC
           ) = 1
),
latest_hes_customer AS (
    SELECT id,
           cuid,
           current_detailed_address,
      FROM `RAW_HES_DATASET.CUSTOMER`
     WHERE UPPER(COALESCE(current_detailed_address, 'NULL')) != 'NULL'
   QUALIFY ROW_NUMBER() OVER (
           PARTITION BY cuid
            ORDER BY BQ_UPDATED_TIME DESC
           ) = 1
)
SELECT appl.serial_number,
       appl.created_at,
       cust.cuid,
       cust.current_detailed_address
  FROM latest_hes_application AS appl
  JOIN latest_hes_customer AS cust ON appl.customer_id = cust.id
 LEFT JOIN `RAW_EDEP_DATASET.GEOCODING` AS geo_coding
        ON appl.serial_number = geo_coding.serial_number
        OR cust.current_detailed_address = geo_coding.REQUEST_ADDRESS
 WHERE geo_coding.SERIAL_NUMBER IS NULL
   AND geo_coding.REQUEST_ADDRESS IS NULL
```

---

### 附錄二 — 戶籍地址發查

```sql
-- 戶籍地址：取出尚未發查的客戶戶籍地址
WITH latest_hes_application AS (
    SELECT id,
           serial_number,
           created_at,
           customer_id,
      FROM `RAW_HES_DATASET.APPLICATION`
   QUALIFY ROW_NUMBER() OVER (
           PARTITION BY id
            ORDER BY TIMESTAMP(created_at) DESC,
                     SAFE_CAST(finished_at AS TIMESTAMP) DESC,
                     PARTITION_DATE DESC
           ) = 1
),
latest_hes_customer AS (
    SELECT id,
           cuid,
           permanent_detailed_address,
      FROM `RAW_HES_DATASET.CUSTOMER`
     WHERE UPPER(COALESCE(permanent_detailed_address, 'NULL')) != 'NULL'
   QUALIFY ROW_NUMBER() OVER (
           PARTITION BY cuid
            ORDER BY BQ_UPDATED_TIME DESC
           ) = 1
)
SELECT appl.serial_number,
       appl.created_at,
       cust.cuid,
       cust.permanent_detailed_address
  FROM latest_hes_application AS appl
  JOIN latest_hes_customer AS cust ON appl.customer_id = cust.id
 LEFT JOIN `RAW_EDEP_DATASET.GEOCODING` AS geo_coding
        ON appl.serial_number = geo_coding.serial_number
        OR cust.permanent_detailed_address = geo_coding.REQUEST_ADDRESS
 WHERE geo_coding.SERIAL_NUMBER IS NULL
   AND geo_coding.REQUEST_ADDRESS IS NULL
```

---

### 附錄三 — 簽約經緯度發查

```sql
-- 簽約經緯度：取出尚未發查的客戶簽約座標
WITH latest_hes_application AS (
    SELECT id,
           serial_number,
           created_at,
           customer_id,
      FROM `RAW_HES_DATASET.APPLICATION`
   QUALIFY ROW_NUMBER() OVER (
           PARTITION BY id
            ORDER BY TIMESTAMP(created_at) DESC,
                     SAFE_CAST(finished_at AS TIMESTAMP) DESC,
                     PARTITION_DATE DESC
           ) = 1
),
latest_hes_customer AS (
    SELECT id,
           cuid,
      FROM `RAW_HES_DATASET.CUSTOMER`
   QUALIFY ROW_NUMBER() OVER (
           PARTITION BY cuid
            ORDER BY BQ_UPDATED_TIME DESC
           ) = 1
),
latest_vmb_apply_info AS (
    SELECT cuid,
           ROUND(longitude, 4) AS longitude,
           ROUND(latitude, 4) AS latitude,
      FROM `RAW_VMB_DATASET.APPLY_INFO`
     WHERE UPPER(COALESCE(longitude, 'NULL')) != 'NULL'
       AND UPPER(COALESCE(latitude, 'NULL')) != 'NULL'
   QUALIFY ROW_NUMBER() OVER (
           PARTITION BY cuid
            ORDER BY BQ_UPDATED_TIME DESC
           ) = 1
)
SELECT appl.serial_number,
       appl.created_at,
       cust.cuid,
       apply_info.longitude,
       apply_info.latitude
  FROM latest_hes_application AS appl
  JOIN latest_hes_customer AS cust ON appl.customer_id = cust.id
  JOIN latest_vmb_apply_info AS apply_info ON cust.cuid = apply_info.cuid
 LEFT JOIN `RAW_EDEP_DATASET.GEOCODING` AS geo_coding
        ON appl.serial_number = geo_coding.serial_number
        OR apply_info.longitude = geo_coding.REQUEST_LONGITUDE
        OR apply_info.latitude = geo_coding.REQUEST_LATITUDE
 WHERE geo_coding.SERIAL_NUMBER IS NULL
   AND geo_coding.REQUEST_LONGITUDE IS NULL
   AND geo_coding.REQUEST_LATITUDE IS NULL
```

---

### 附錄四 — 更新已發查通訊地址

```sql
-- 更新通訊地址資訊
INSERT INTO `RAW_EDEP_DATASET.GEOCODING` (
    CUID,
    SERIAL_NUMBER,
    CREATED_AT,
    GEO_TYPE,
    REQUEST_ADDRESS,
    RESPONSE_PLACE_ID,
    RESPONSE_ADRESS,
    RESPONSE_GLOBAL_CODE,
    RESPONSE_PLACE_TYPES,
    RESPONSE_LONGITUDE,
    RESPONSE_LATITUDE,
    RESPONSE_COUNTRY,
    RESPONSE_CITY,
    RESPONSE_DISTRICT,
    RESPONSE_WARD,
    RESPONSE_STREET
)
WITH latest_hes_application AS (
    SELECT id,
           serial_number,
           created_at,
           customer_id,
      FROM `RAW_HES_DATASET.APPLICATION`
   QUALIFY ROW_NUMBER() OVER (
           PARTITION BY id
            ORDER BY TIMESTAMP(created_at) DESC,
                     SAFE_CAST(finished_at AS TIMESTAMP) DESC,
                     PARTITION_DATE DESC
           ) = 1
),
latest_hes_customer AS (
    SELECT id,
           cuid,
           current_detailed_address,
      FROM `RAW_HES_DATASET.CUSTOMER`
     WHERE UPPER(COALESCE(current_detailed_address, 'NULL')) != 'NULL'
   QUALIFY ROW_NUMBER() OVER (
           PARTITION BY cuid
            ORDER BY BQ_UPDATED_TIME DESC
           ) = 1
),
geocoding_serial_number_list AS (
    SELECT DISTINCT serial_number
      FROM `RAW_EDEP_DATASET.GEOCODING`
     WHERE GEO_TYPE = 'CONTACT_ADDRESS'
),
distinct_address AS (
    SELECT DISTINCT REQUEST_ADDRESS,
           RESPONSE_PLACE_ID,
           RESPONSE_ADRESS,
           RESPONSE_GLOBAL_CODE,
           RESPONSE_PLACE_TYPES,
           RESPONSE_LONGITUDE,
           RESPONSE_LATITUDE,
           RESPONSE_COUNTRY,
           RESPONSE_CITY,
           RESPONSE_DISTRICT,
           RESPONSE_WARD,
           RESPONSE_STREET
      FROM `RAW_EDEP_DATASET.GEOCODING`
     WHERE GEO_TYPE = 'CONTACT_ADDRESS'
       AND REQUEST_ADDRESS IS NOT NULL
)
SELECT cust.cuid,
       appl.serial_number,
       appl.created_at,
       'CONTACT_ADDRESS'              AS GEO_TYPE,
       cust.current_detailed_address  AS REQUEST_ADDRESS,
       da.RESPONSE_PLACE_ID,
       da.RESPONSE_ADRESS,
       da.RESPONSE_GLOBAL_CODE,
       da.RESPONSE_PLACE_TYPES,
       da.RESPONSE_LONGITUDE,
       da.RESPONSE_LATITUDE,
       da.RESPONSE_COUNTRY,
       da.RESPONSE_CITY,
       da.RESPONSE_DISTRICT,
       da.RESPONSE_WARD,
       da.RESPONSE_STREET
  FROM latest_hes_application AS appl
  JOIN latest_hes_customer AS cust ON appl.customer_id = cust.id
  JOIN distinct_address AS da ON cust.current_detailed_address = da.REQUEST_ADDRESS
 WHERE appl.serial_number NOT IN (SELECT serial_number FROM geocoding_serial_number_list)
```

---

### 附錄五 — 更新已發查戶籍地址

> 邏輯同附錄四，將 `CONTACT_ADDRESS` 替換為 `RESIDENCE_ADDRESS`，欄位改為 `permanent_detailed_address`，可依此類推撰寫。

---

### 附錄六 — 更新已發查經緯度

> 邏輯同附錄四，將 `GEO_TYPE` 替換為 `CONTRACT_COORDINATES`，請求欄位改為 `longitude` / `latitude`，可依此類推撰寫。

---

### 附錄七 — 下行檢核程式碼

```python
def validate_geocoding_response(response: dict) -> bool:
    """
    驗證 Google Maps Geocoding API 下行回應是否有效。
    回傳 True 表示資料有效，可儲存；False 表示跳過。
    """
    status = response.get("status")

    if status != "OK":
        return False

    results = response.get("results", [])
    if not results:
        return False

    result = results[0]

    # 必填欄位驗證
    if not result.get("formatted_address"):
        return False

    geometry = result.get("geometry", {})
    location = geometry.get("location", {})
    lat = location.get("lat")
    lng = location.get("lng")

    if lat is None or lng is None:
        return False

    try:
        float(lat)
        float(lng)
    except (ValueError, TypeError):
        return False

    address_components = result.get("address_components", [])
    if not address_components:
        return False

    return True


def extract_address_component(address_components: list, target_type: str) -> str | None:
    """
    從 address_components 中取出指定類型的 long_name。
    """
    for component in address_components:
        if target_type in component.get("types", []):
            return component.get("long_name")
    return None
```
