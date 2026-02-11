---
name: apply-fix
description: 套用 Code_review_result.md 的修正 diff
agent: Apply-Fix
argument-hint: 可選參數：--include-warning（預設只套用 Fail❌）
model: Grok Code Fast 1 (copilot)
---

# /apply-fix — 自動套用審查修正

## 目的

使用者在看完 [Code_review_result.md](../../Code_review_result.md) 後，可在不手動貼 diff 的情況下，要求 Copilot 自動套用修正。

## 使用方式

- 基本：輸入 `/apply-fix`
- 進階：輸入 `/apply-fix --include-warning`
- 指定來源：輸入 `/apply-fix --from .code-review-batches/batch-01.report.md`
- 組合：輸入 `/apply-fix --include-warning --from .code-review-batches/batch-01.report.md`

## 行為規範（必用）

- 觸發條件：**只要使用者輸入 `/apply-fix`** 即進入套用流程。
- 資料來源：
  - 預設：你**必須自行讀取**工作區的 [Code_review_result.md](../../Code_review_result.md)，並從中抽取所有 ` ```diff ` 區塊。
  - 若使用者提供 `--from <report_path>`：改讀取該報告檔（repo 相對路徑）。
- 套用範圍：
  - 預設只套用 **Fail❌** 對應的修正（通常位於 `## 修正程式碼片段`）。
  - 只有在使用者提供 `--include-warning` 時，才允許嘗試處理 Warning⚠️（若報告中沒有 Warning 的 diff，則不得自行產生大量重寫）。

## 安全限制（必用）

- 只允許套用到 repo 內的相對路徑檔案。
- 若 diff 中出現以下任一情況，必須拒絕該 diff block 並回報原因：
  - 絕對路徑（例如 `C:\\...`、`/etc/...`、`\\\\server\\share...`）
  - 路徑穿越（包含 `..`）
  - 目標檔案不存在（保守策略：不建立新檔）

> `--from` 的 `<report_path>` 同樣必須符合「repo 相對路徑」與禁止 `..` / 絕對路徑等限制；檔案不存在則必須拒絕。

## 找不到 diff 時的行為

- 若 [Code_review_result.md](../../Code_review_result.md) 不存在或沒有 ` ```diff ` 區塊：
  - 不修改任何檔案
  - 引導使用者先執行 `/code-review` 產生報告，或確認是否有 Fail❌

## 輸出格式

- 回覆必須包含：
  - 讀取到幾個 diff block
  - 套用成功/失敗的檔案清單與原因
  - 下一步建議（例如重新執行 `/code-review` 驗證是否已消除 Fail❌）
