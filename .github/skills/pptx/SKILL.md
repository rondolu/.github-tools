---
name: pptx
description: "Use this skill any time a .pptx file is involved in any way — as input, output, or both. This includes: creating slide decks, pitch decks, or presentations; reading, parsing, or extracting text from any .pptx file (even if the extracted content will be used elsewhere, like in an email or summary); editing, modifying, or updating existing presentations; combining or splitting slide files; working with templates, layouts, speaker notes, or comments. Trigger whenever the user mentions \"deck,\" \"slides,\" \"presentation,\" or references a .pptx filename, regardless of what they plan to do with the content afterward. If a .pptx file needs to be opened, created, or touched, use this skill."
license: Proprietary. LICENSE.txt has complete terms
---

# PPTX 技能

## 快速參考

| 任務 | 指引 |
|------|-------|
| 讀取／分析內容 | `python -m markitdown presentation.pptx` |
| 從範本編輯或建立 | 閱讀 [editing.md](editing.md) |
| 從頭建立 | 閱讀 [pptxgenjs.md](pptxgenjs.md) |
| 檢視／切換設計範本 | [fetch_template_from_ppt.yaml](fetch_template_from_ppt.yaml)（預設）· [cathay_template.yaml](cathay_template.yaml)（備用） |

---

## 讀取內容

```bash
# 文字擷取
python -m markitdown presentation.pptx

# 視覺概覽
python scripts/thumbnail.py presentation.pptx

# 原始 XML
python scripts/office/unpack.py presentation.pptx unpacked/
```

---

## 編輯工作流程

**完整細節請閱讀 [editing.md](editing.md)。**

1. 使用 `thumbnail.py` 分析範本
2. 解包 → 操作投影片 → 編輯內容 → 清理 → 重新打包

---

## 從頭建立

**完整細節請閱讀 [pptxgenjs.md](pptxgenjs.md)。**

當沒有可用的範本或參考簡報時使用此方式。

---

## 設計系統（必須遵守）

**在產生任何簡報之前，您必須讀取當前設計範本 YAML 並套用其規格。**
設計範本完整定義了顏色、字型、版面配置及每張投影片的元素規格——請勿使用 YAML 中未定義的值。

**預設範本**（在撰寫任何投影片程式碼前，請先讀取此檔案）：
```
.github/skills/pptx/fetch_template_from_ppt.yaml
```

YAML 結構如下所示：

| YAML 區段 | 定義內容 |
|--------------|-----------------|
| `design_system.color_palette` | 所有允許的十六進位色碼及使用規則——**不允許使用其他顏色** |
| `design_system.typography` | 字型家族、大小層級、字重（全部為粗體）及各文字角色的顏色 |
| `design_system.layout_rules` | 投影片尺寸（960×540pt = 10"×5.625"）、邊距、格線及 UI 元素樣式 |
| `layout_rules.recurring_elements` | **每張投影片都必須出現**的裝飾性元素 |
| `slide_templates` | 具有精確元素位置與樣式的具名版面配置模式 |
| `instructions_for_generation` | 必要規則：語調、禁止元素、色彩規範、間距 |

### 必要閱讀工作流程

在撰寫任何投影片程式碼或 XML 之前：
1. **讀取** `.github/skills/pptx/fetch_template_from_ppt.yaml`
2. **擷取**相關設計值（顏色、字型、版面規格）
3. **套用**這些值於 PptxGenJS 呼叫或 XML 編輯——YAML 到程式碼的對應指南請參閱 [pptxgenjs.md](pptxgenjs.md)

### 重複裝飾元素（每張投影片——無例外）

定義於 `layout_rules.recurring_elements`。這些元素**必須出現在每張投影片上**：

| 元素 | 位置 | 樣式 |
|---------|----------|-------|
| 頂部細長條 | x=0, y=0, 全寬, h≈5pt | 深海軍藍填充（`0D1F33`） |
| 底部三段式條 | y≈518pt, h≈22pt, 三等分 | 由左至右漸層為漸淺的深藍色 |
| 左側垂直強調線 | x=0, y≈5pt, h≈513pt, w=5–8pt | 電光藍 `186AFF` |
| 頁面／章節編號標籤 | 左上角，約 80×70pt 文字框 | 40pt 粗體 `186AFF`，兩位數編碼（01/02/…） |

### 投影片範本選擇

YAML 在 `slide_templates` 下定義了具名版面配置類型。請將每張投影片的內容對應到最接近的類型：

| 範本類型 | 最適合用於 |
|---------------|---------------|
| Cover / Title Slide（封面／標題投影片） | 開場／標題頁 |
| Agenda / Table of Contents（議程／目錄） | 章節概覽，附大型 KPI 數字 |
| Content - Horizontal Row（內容－水平列） | 2–3 個並列主題；每列含圖示與正文 |
| Content - Left Panel + Right Multi-Section（內容－左側面板＋右側多節） | 深度主題，含側邊欄背景資訊 |
| Content - Section Divider（內容－章節分隔） | 章節轉換／視覺換頁 |
| Content - Comparison Table（內容－對比表格） | 錯誤 vs 正確／之前 vs 之後 |
| Content - Flow Diagram（內容－流程圖） | 循序步驟或決策樹 |
| Content - Three-Card Comparison（內容－三卡片對比） | 並排規則或選項卡片 |
| Closing / Thank You（結尾／感謝） | 摘要／結語 |

閱讀 YAML 的 `slide_templates[].elements`，可取得每個範本精確的 x/y 位置、尺寸、字型及顏色。

### 替代範本

原始國泰企業設計系統的備份保存於：
```
.github/skills/pptx/cathay_template.yaml
```
若要切換主題，請改讀取該 YAML 並依照相同工作流程套用其值。

---

## 品質驗證（必要步驟）

**預設有問題存在。您的任務是找出它們。**

第一次的輸出幾乎從不正確。請以找 Bug 的心態進行品質驗證，而非確認步驟。如果第一次檢查就發現零問題，代表您看得不夠仔細。

### 內容品質驗證

```bash
python -m markitdown output.pptx
```

檢查是否有遺漏內容、錯字、順序錯誤。

**使用範本時，請檢查是否有殘留的預留位置文字：**

```bash
python -m markitdown output.pptx | grep -iE "xxxx|lorem|ipsum|this.*(page|slide).*layout"
```

若 grep 有回傳結果，請在宣告完成前修正。

### 視覺品質驗證

**⚠️ 請使用子代理（SUBAGENTS）**——即使只有 2-3 張投影片也一樣。您已盯著程式碼太久，只會看見您期望看到的，而非實際存在的內容。子代理擁有全新的視角。

將投影片轉換為圖片（參見[轉換為圖片](#轉換為圖片)），然後使用以下提示詞：

```
請視覺化檢查這些投影片。預設有問題存在——請找出它們。

需檢查的項目：
- 元素重疊（文字穿越圖形、線條穿越文字、堆疊元素）
- 文字溢出或在邊緣／框線處被截斷
- 裝飾線設計給單行文字，但標題換行為兩行
- 來源引用或頁腳與上方內容碰撞
- 元素間距過近（< 0.3" 間距）或卡片／節區幾乎相接
- 間距不均勻（一處有大片空白，另一處卻很擁擠）
- 距投影片邊緣邊距不足（< 0.5"）
- 欄或類似元素未對齊
- 低對比文字（例如淺灰色文字在米白色背景上）
- 低對比圖示（例如深色圖示在深色背景上，沒有對比圓圈）
- 文字框過窄導致過度換行
- 殘留的預留位置內容

針對每張投影片，列出問題或需注意之處，即使是細節也要列出。

請讀取並分析這些圖片：
1. /path/to/slide-01.jpg（預期內容：[簡短描述]）
2. /path/to/slide-02.jpg（預期內容：[簡短描述]）

回報所有發現的問題，包括細節問題。
```

### 驗證循環

1. 產生投影片 → 轉換為圖片 → 檢查
2. **列出發現的問題**（若未發現問題，請再更嚴格地重新檢查）
3. 修正問題
4. **重新驗證受影響的投影片**——修正一個問題往往會產生另一個問題
5. 重複執行，直到完整檢查一遍後不再出現新問題

**在完成至少一次修正後驗證循環之前，請勿宣告完成。**

---

## 轉換為圖片

將簡報轉換為個別投影片圖片以進行視覺檢查：

```bash
python scripts/office/soffice.py --headless --convert-to pdf output.pptx
pdftoppm -jpeg -r 150 output.pdf slide
```

這會建立 `slide-01.jpg`、`slide-02.jpg` 等檔案。

修正後重新輸出特定投影片：

```bash
pdftoppm -jpeg -r 150 -f N -l N output.pdf slide-fixed
```

---

## Dependencies

- `pip install "markitdown[pptx]"` - text extraction
- `pip install Pillow` - thumbnail grids
- `npm install -g pptxgenjs` - creating from scratch and PDF-to-PPTX conversion
- `pip install rapidocr_onnxruntime` - OCR for image-based PDFs (recommended, no system deps)
- `pip install pytesseract` - fallback OCR for image-based PDFs (requires Tesseract binary)
- LibreOffice (`soffice`) - PDF conversion (auto-configured for sandboxed environments via `scripts/office/soffice.py`)
- Poppler (`pdftoppm`, `pdfimages`) - PDF to images and embedded image extraction
