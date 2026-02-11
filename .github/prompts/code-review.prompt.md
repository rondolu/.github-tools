---
name: code-review
description: Python / BigQuery Standard SQL / checkmarx
agent: Review
argument-hint: 貼上要審查的 Python/SQL 程式碼或 diff；可包含關鍵字或標籤：{報表}/報表/report/{資料匯出}/匯出/export/{載入}/load/etl
model: Grok Code Fast 1 (copilot)
---

# Coding Style & Security Standard（Python / BigQuery Standard SQL）

> 目的：只定義「coding style standard」與「Security Baseline」一致性規範，提升可讀性、可維護性、安全性與 code review 的可重複性。
>
> 非目標：不規範任何業務邏輯／業務規則；業務正確性請以測試案例與測試 review 為主。

---

## 規範適用範圍

### In Scope

- Python 程式碼（含 Flask/Web 應用程式與腳本）
- BigQuery Standard SQL（`.sql` 檔與 Python 內組裝的 query 字串）
- Code review 僅針對「style」與「security」可判定事項

### Out of Scope

- 業務規則（例如報表欄位型別、拆檔規則、特定流程狀態碼意義）
- 資料正確性（由測試案例與資料驗證流程保障）
- 效能/成本最佳化（除非影響可讀性或明顯違反 BigQuery 常見建議；若要納入請另立文件）

---

## 1) 官方參考標準

### Python

- PEP 8 — Style Guide for Python Code
- PEP 257 — Docstring Conventions
- PEP 484 / PEP 526 — Type Hints / Variable Annotations
- Python 3.11 官方文件：`typing`、例外處理、標準函式庫慣例

### Security (Based on Checkmarx & OWASP)

- OWASP Top 10 (Web/API Security)
- CWE/SANS Top 25
- Checkmarx Python Vulnerability Best Practices

### BigQuery SQL

- BigQuery Standard SQL 語法（避免 Legacy SQL）
- BigQuery 參數化查詢（`@param`）

> 註：本文件以官方標準為準；若 repo 現況與官方建議衝突，以「可落地」為優先，採漸進式修正。

---

## 2) 用語與規範強度

- **MUST**：必須遵守（新程式碼與修改行必須符合）
- **SHOULD**：建議遵守（新程式碼建議符合；舊程式碼可逐步調整）
- **MAY**：可選（依情境選擇）

> 本文件的立場：新/改動程式碼必須符合；舊程式碼可用「觸碰即修」策略逐步清理。

### 2.1 審查狀態Mapping（必用）

- **每條規則**都必須標註其 規則等級：MUST / SHOULD / MAY（以本文件條目上的 MUST/SHOULD/MAY 為準）。

- **每條規則狀態（Status）**判定：

    - 規則屬 **MUST** 且未滿足 → **Fail❌**
    - 規則屬 **SHOULD** 且未滿足 → **Warning⚠️**
    - 規則屬 **MAY** 且未滿足 → **Pass✅**（仍需在「詳細說明」註記未符合與原因）

- **整體（Overall）狀態**判定（取最高嚴重度）：

    - 任一 MUST Fail → Overall = **Fail❌**
    - 否則任一 SHOULD 未滿足 → Overall = **Pass⚠️**
    - 否則 → Overall = **Pass✅**
- **修正程式碼片段**：僅針對 **Fail❌（MUST 未滿足）** 的項目提供修正程式碼；Warning⚠️ 僅提供建議。

### 2.2 場景觸發（必用）

- 需從「使用者描述／PR 描述」中精確匹配關鍵字字串（完全相同，不允許部分匹配或推斷），以啟動對應場景規範。若未精確匹配，則不觸發場景檢查。
- 場景關鍵字定義見本文件第 7 節。

### 2.3 Coverage 與掃描範圍（必用）

- MUST：每次審查都要先宣告 `Denominator Type`，僅能為 `Folder` / `PR` / `FileList` / `N/A`。
- MUST：Denominator 必須可被可靠枚舉為 repo 相對路徑清單；不得用推測或抽樣。預設選擇規則：
    - 使用者提供變更檔案清單或可解析的 diff header → `PR`。
    - 指定資料夾或 glob → `Folder`。
    - 其他情境需人工列出檔案清單 → `FileList`。
    - 無法取得任何清單 → `N/A`（直接 Fail(Coverage)）。
- MUST：`Excluded` 不得計入 `N`；所有排除規則需在 `Scope/Exclude` 區塊輸出 `Exclude Rules`（glob / 類型 / 其他條件）與 `Exclude Reasons`（包含 reason_code + 描述），並提供 `Excluded Count`（可列前 N 筆 + 省略說明）。
- MUST：對納入分母的所有項目逐一枚舉，並標註 `Reviewed` 或 `Unreviewed`；同時輸出 `Total (N)`、`Reviewed (R)`、`Unreviewed (U)`，並自我驗證 `N = R + U`。若無法成立，Coverage 視為 `N/A`。
- MUST：Coverage 僅在 `R = N ≠ 0` 時才可標示為 100%；任何 `Coverage < 100%` 或 `Coverage = N/A` → `Fail(Coverage)`，並禁止輸出 `## Rules` / `## Fail diff` / `## CTA`（仍需輸出 `Scope/Exclude`、`Coverage`、`Overall` 與 Summary）。
- MUST：`Coverage Reason` 為必填欄位，內容需說明：分母類型、枚舉來源、排除規則、若 Coverage 未達 100% 則列出缺漏項與原因。
- MUST：清單過長時，可輸出「摘要」取代引出全部項目，但必須保留：`Total` / `Reviewed` / `Unreviewed` / `Excluded` / `N` / `Omitted`（省略數量）/ `Top N Reviewed Items` / `Missing/Excluded Reasons Summary`（原因碼建議使用：`NOT_PROVIDED`、`PATH_UNKNOWN`、`BINARY_OR_IMAGE`、`GENERATED`、`THIRD_PARTY`、`OUT_OF_SCOPE`、`TOO_LARGE`、`TIME_BUDGET`）。
- SHOULD：`Top N Reviewed Items` 列表包含 `item_id`（檔案或規則）、`status`、`one-line reason`；若 `Reviewed > N`，需額外輸出 `... omitted: <Omitted> items`。
- SHOULD：`Missing/Excluded Reasons Summary` 需將同原因碼聚合並說明其屬性（unreviewed 或 excluded）。
- MAY：為減少報告長度，可設定 N≤30；若需完整清單，可附加檔或指示使用者改用 `/apply-fix --list-all` 之類流程。

---

---

## 3) Step 1 — Python Style

### 3.1 Formatting（PEP 8）

- MUST：使用 **4 spaces** 縮排；**禁止 Tab**。
- MUST：行尾不得有多餘空白。
- SHOULD：除了字串以外，禁止使用Magic Values。
- MUST：判斷式比較None時統一用 is None / is not None。
- MUST：複合判斷式必須使用小括號 ()包裹明確邏輯。
- SHOULD：不使用負向條件判斷，如：<>、!= 等 (檢查空值例外)。
- SHOULD：單行長度遵循 PEP 8 風格。
- SHOULD：一個函式/方法只做一件事；若邏輯過長，優先抽成可測的輔助函式（這是可讀性要求，不是業務要求）。
- SHOULD：Method/Class 不超過 50 行；巢狀結構不超過 3 層。

### 3.2 命名（PEP 8）

- MUST：`ClassName` 使用 CapWords（UpperCamelCase）。
- MUST：函式/變數使用 `snake_case`。
- MUST：常數使用 `UPPER_SNAKE_CASE`。
- MUST：Private變數和Method以單底線 `_name` 表示。
- MUST：避免無意義縮寫（例如 `a`, `i`, `tmp`）；迴圈索引 `i/j` 僅限極短、語意清楚情境。

### 3.3 Imports（PEP 8）

- MUST：imports 分三段並以空行分隔：標準函式庫 / 第三方套件 / 專案內模組。
- MUST：避免 wildcard import（`from x import *`）。
- SHOULD：優先使用 absolute import；相對 import 只用於明確的 package 內部引用。

### 3.4 Docstring（PEP 257）

- SHOULD：每個 Method/Function 都要有 docstring（說明目的、Args、Returns）。
- MUST：對外介面（routes handler、service public method、library-like helper）docstring 內容需完整且易讀。
- MUST：docstring 描述「做什麼/為什麼」，避免重複程式碼細節。
- MUST：參數/回傳值的描述要與型別標註一致。

### 3.5 型別標註（PEP 484/526；Python 3.11 建議寫法）

- MUST：跨層邊界的函式簽章要有靜態型別標註（例如 routes ↔ services、services ↔ infrastructure）。
- SHOULD：資料結構以 `dict[str, Any]`、`list[dict[str, Any]]` 等 3.11 寫法為主。
- SHOULD：可選值使用 `T | None`（Python 3.10+）或 `Optional[T]` 擇一，但同 repo 需一致。
- MAY：在需要避免 import cycle 或減少執行期成本時使用 `from __future__ import annotations`。

### 3.6 例外處理（Style）

- MUST：禁止裸露 `except:`；至少捕捉 `Exception`。
- MUST：捕捉例外後若要重新拋出，使用 `raise X from e` 保留上下文。
- MUST：不要吞掉例外（`except Exception: pass`）。若真的要忽略，必須有註解說明原因與風險。
- MUST：Method 開頭需進行必要的例外/前置條件判斷（early validation），針對可能發生的錯誤進行對應例外處理。
- SHOULD：錯誤訊息要可行動（actionable），包含必要上下文（例如 resource id、參數摘要），但不要記錄敏感資料。

### 3.7 Logging（Style）

- MUST：禁止以 `print(...)` 做長期日誌；改用GCP Cloud Logging 的自訂 logger（custom-log）寫入日誌。
- MUST：所有 Error/Warning 需寫入GCP Cloud Logging 的自訂 logger（custom-log）寫入日誌。
- SHOULD：warning/error 需包含可追蹤的識別資訊（例如 request id/uuid），但不要包含 secret/token。
- MAY：log message 以「事件 + 關鍵欄位」形式撰寫（偏結構化），避免單純堆疊長字串。

---

## 4) Step 2 — BigQuery Standard SQL Style

### 4.1 語法與模式

- MUST：使用 **BigQuery Standard SQL**。
- MUST：在 Python 端明確 `use_legacy_sql = False`（若使用 BigQuery client）。

### 4.2 排版

- MUST：使用 SQL formatter（Tab size=4、Keywords=Upper、Indentation=Tabular, Right）。
- MUST：關鍵字大小寫一致（全大寫如：`SELECT`, `FROM`, `WHERE`, `JOIN`, `QUALIFY`）。
- MUST：資料型態大小寫一致(全大寫如: `STRING`, `INT64`, `FLOAT64`, `DATETIME`)。
- MUST：函數大小寫一致(全大寫如: `CAST()`, `DATE()`, `CURRENT_DATE()` )。
- SHOULD：運算式寬度為 50 個字元以內（可讀性優先）。
- SHOULD：查詢語句間的空白行為 1 行。
- SHOULD：一行一個欄位（`SELECT` list、`INSERT` 欄位清單）；逗號放行尾。

### 4.3 命名
- MUST：CTE/alias 使用 `snake_case`。
- SHOULD：alias 需可讀（例如 `customer`/`application`），避免 `a/b/c`（除非 query 很短且語意清楚）。

### 4.5 SELECT 欄位與 `SELECT *`

- SHOULD：最終輸出 `SELECT` 禁止 `SELECT *`。
- SHOULD：子查詢也避免 `SELECT *`；若是為了搭配 `QUALIFY ROW_NUMBER()` 取最新一筆，仍建議只取需要欄位。

### 4.6 JOIN 與條件可讀性

- MUST：所有 `JOIN` 必須有明確的 `ON` 條件。
- SHOULD：當在 `WHERE` 或 `JOIN` 條件中使用 `AND` / `OR` 運算子換行時，運算子應該放在下一行的行首。

### 4.7 參數化與日期時間

- MUST：可變條件使用參數（例如 `@partition_date`），避免字串拼接。
- SHOULD：日期/時間處理用 BigQuery 提供的型別/函式（例如 `DATE(...)`、`CURRENT_DATE()`）；若有固定需求（例如禁止 CURRENT_DATE）屬業務規則，請放到測試案例。

### 4.8 物件命名與反引號

- SHOULD：完整表格名稱使用反引號：`` `project.dataset.table` ``。
- MAY：在無歧義且不跨 dataset 的情境省略反引號，但同 repo 建議統一使用反引號以降低踩雷。

---

## 5) Step 3 — Security Standard (Python/Web)

> 基於 Checkmarx 掃描與 OWASP 標準。

### 5.1 錯誤處理與資訊洩露 (Information Exposure)
- MUST：**禁止直接回傳原始錯誤訊息**。API 或前端回應中包含 `str(e)`、Exception 物件或 Stack Trace的物件要用GCP Cloud Logging 的自訂 logger（custom-log）寫入日誌。
- MUST：禁止在 Log 或 UI 輸出敏感的內部結構資訊（如資料庫連線字串、資料表欄位定義）。
- SHOULD：使用統一的錯誤/警告回應機制，用GCP Cloud Logging 的自訂 logger（custom-log）寫入日誌。

### 5.2 日誌安全 (Logging Security)
- MUST：**防範 Log Forging**。寫入 Log 的變數若來自使用者輸入（如 `get_json`），改為使用環境變數傳遞（避免外部控制 API 請求內容）。
- MUST：禁止在 Log 紀錄中輸出 PII（如身分證、電話）、Session ID、Auth Token 或業務敏感資料。
- MUST：**防範不受控制的格式化字串**。禁止將外部輸入直接用於格式化字串（如 `f"Hello {external_input}"`）；應使用 `string.Template.safe_substitute()` 或其他安全方法處理。

### 5.3 網路連線與傳輸安全 (Transmission)
- MUST：**強制 SSL 驗證**。使用 HTTP Client（如 `requests`）時，嚴禁設定 `verify=False`。SSL 驗證參數由環境變數控制，確保一致性。
- MUST：**防範 SSRF**。若程式會根據輸入請求遠端網址，將相關參數改為使用環境變數傳遞（避免外部控制 API 請求內容）。
- MUST：Web 應用程式應設定 `Strict-Transport-Security` (HSTS) 標頭，預設使用Talisman (flask_talisman == 1.1.0)。

### 5.4 組態與憑證管理 (Secrets & Config)
- MUST：**停用除錯模式**。正式環境程式碼中，`app.run()` 的 `debug` 參數必須為 `False`。
- MUST：**禁止硬編碼憑證**。程式碼中不可包含寫死的密碼、API Key 或連線字串（如 `_Password` 變數）；應使用不具敏感含義且可被弱掃軟體接受的同義字，如 Credential_Code。
- MUST：Web 應用程式回應標頭應設定 `Content-Security-Policy` (CSP)，並限制所有來源為 none、禁止被 frame、禁止表單提交、禁止 base-uri。

### 5.5 寫法指引 (Coding guidelines)
- MUST：**對齊弱掃規範**。將 with open(...).read() 改為單行讀取（例如 `data = open(...).read()`），以符合 Checkmarx 的安全掃描規範。
- MUST：**防範不受控制的格式化字串**。禁止將使用者輸入直接用於格式化字串（如 `f"Hello {user_input}"`）；應使用 `string.Template.safe_substitute()` 或其他安全方法處理。
---

## 6) Step 4 — Code Review Checklist

> 使用方式：PR reviewer（人）與 AI reviewer 都只對照此清單，不討論業務規則；業務規則請看測試案例。

### 6.1 Python Style Checklist
- [ ] 縮排 4 spaces、無 Tab、無行尾空白
- [ ] 判斷式比較 None 統一用 `is None` / `is not None`
- [ ] 複合判斷式必須使用小括號 () 包裹明確邏輯
- [ ] 命名符合 PEP 8（class/function/const/private）
- [ ] 禁止 Magic Values（字串除外），改用常數
- [ ] imports 分組清楚、無 wildcard import、無循環依賴跡象
- [ ] 每個 Method/Function 有 docstring（目的/Args/Returns）
- [ ] 跨層邊界函式有型別標註；`Optional`/`| None` 用法一致
- [ ] 無 `print(...)` 作為日誌；改用自訂義logger
- [ ] 無裸露 `except:`；無吞例外；必要時 `raise ... from e`
- [ ] Error/Warning 必須寫入 logger，且有足夠上下文
- [ ] 函式/方法長度與巢狀結構合理（Method/Class 不超過 50 行；巢狀不超過 3 層）

### 6.2 Security Checklist (Checkmarx Compliance)
- [ ] **Error Handling**: 未直接回傳 `str(e)` 或 Stack Trace 給前端/API
- [ ] **Logging**: Log 內無 PII/Secrets；使用者輸入寫入 Log 前已淨化 (防 Log Forging)
- [ ] **Network**: HTTP 請求未設定 `verify=False` (SSL Bypass)
- [ ] **Network**: 針對使用者輸入的 URL 請求已做白名單檢查 (SSRF)
- [ ] **Config**: 正式環境 `debug=False`；無 Hardcoded Secrets (密碼/Key)
- [ ] **Validation**: 未將使用者輸入直接用於格式化字串 (Uncontrolled Format String)
- [ ] **Header**: (Web專用) 已設定 CSP 與 HSTS 標頭

### 6.3 BigQuery SQL Style Checklist
- [ ] 確認為 Standard SQL（不是 Legacy）
- [ ] 使用 SQL formatter（Tab size=4、Keywords=Upper、Indentation=Tabular, Right）
- [ ] 關鍵字大小寫、縮排、逗號位置一致
- [ ] CTE 命名 `snake_case` 且結構清楚（從 raw/prepare → transform → final）
- [ ] 最終輸出無 `SELECT *`；欄位清單明確
- [ ] JOIN/ON 條件清楚可讀；避免在 WHERE 裡暗藏 join 條件
- [ ] 參數使用 `@param`；避免字串拼接 query
- [ ] 完整表格名稱與反引號使用一致

---

## 7) 場景規範（依觸發條件啟動）

> 注意：此為「程式碼審查規則」的一部分。若未觸發場景，則不必檢查本節。

### 7.1 場景：報表 / 儀表板（Report / Dashboard）

觸發關鍵字：`報表`, `report`, `dashboard`, `視覺化`

- MUST：金額/金錢欄位必須 `CAST` 為 `FLOAT64`（當欄位語意明確為金額時）。
- MUST：SQL 必須使用傳入的日期參數（例如 `@run_date`），不可使用 `CURRENT_DATE()` / `CURRENT_DATETIME()` 作為報表日期。
- MUST：Task Number 必須寫在程式碼註解中（SQL 或 Python 入口皆可；禁止寫在 common.json）。
- MUST：必須包含拆檔機制（Split mechanism）。
- MUST：SQL 邏輯必須獨立為 `.sql` 檔案，不可嵌入 Python 長字串內。

### 7.2 場景：資料匯出 / 產檔（Data Export / File Generation）

觸發關鍵字：`{資料匯出}`, `資料匯出`, `匯出`, `export`, `output`, `產檔`

- MUST：程式碼必須處理檔案分割（例如限制檔案大小或列數）。
- MUST：金額欄位必須 `CAST` 為 `FLOAT64`（當欄位語意明確為金額時）。
- MUST：SQL 必須使用傳入的日期參數，不可使用 `CURRENT_DATE()` / `CURRENT_DATETIME()`。
- MUST：日期條件必須使用參數注入（不可用字串拼接）。
- MUST：SQL 語法必須獨立為 `.sql` 檔。

### 7.3 場景：資料載入 / ETL（Data Loading / ETL）

觸發關鍵字：`{載入}`, `載入`, `load`, `etl`, `ingest`

- MUST：當資料會被載入記憶體處理時，必須避免一次性全量載入：
	- 若存在 `partition_date` → MUST 依日期分批處理。
	- 若無分區 → MUST 使用 chunking（例如 `chunksize=10000` 或等效策略）以避免記憶體溢位（OOM）。
- MUST：輸入參數必須包含 `start_date` 與 `end_date`（若流程語意為區間處理）。

---

## 8) 對 Review Agent 的執行指示（必用）

### 8.1 修正程式碼片段（必用）

- **僅**針對狀態為 **Fail❌（MUST 未滿足）** 的項目提供修正程式碼片段。
- Warning⚠️ 與 Pass✅（含 MAY 未滿足）只提供建議，不貼大段重寫。
- 修正程式碼片段必須以 unified diff 格式輸出，並包在 ```diff 區塊內。