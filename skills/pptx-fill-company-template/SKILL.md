---
name: pptx-fill-company-template
description: 套版：把現有內容重做到公司簡報範本（.pptx/.potx）上。來源可以是舊簡報（月報、週報等），或文字類檔案（.md .txt .csv .xlsx .docx；PDF 先轉文字）。用範本的版面、主題色、字型重建每一頁。Use when asked to apply a company template to a deck or document, "套版", "把內容套到公司簡報範本", "換成公司範本".
---

# 套用公司範本（套版）

**重建，不要複製**：從來源抽出內容，用範本的版面建立「新的」投影片。直接複製來源投影片會把舊母片、背景一起帶過來。

本技能的腳本在 `<skills 資料夾>/pptx-fill-company-template/scripts/`，以下用 `DIR` 代表。分享時要給整個資料夾。

## 來源是簡報（pptx）：用 auto_build，不要自己寫整份腳本

```
python "DIR/auto_build.py" "來源.pptx" "範本.pptx" "輸出.pptx"
```

這一個指令會做完全部：
1. **分析每一頁**，寫出頁面計畫 `_work/plan.json`（每頁一筆：封面、目錄、表格、圖表、文字卡片、結尾…）。
2. **依計畫產生投影片**：文字和數字在產生時直接從來源讀，計畫裡不記內文，所以不會漏字、不會出現自己編的字。
   一頁對一頁；表格自動挑字級放進一頁（最小 8pt）；表格旁邊的說明放右側欄；「標籤：數字」做成數字卡片；
   每行都是「A：23 B：17 C：73%」的做成小表格；並排的區塊做成卡片；來源用顏色標示的字（紅色 ● 等）保留顏色。
3. **自動檢查**：漏字、內容跑到別頁、一頁拆成多頁、數字不見、自己加的文字、重疊、圖表數列；
   有 PowerPoint 時匯出每頁圖片和總覽圖，並做 PowerPoint 實際排版檢查。
4. 最後印出：`說明 …`（要轉告使用者的事）、`問題 …`（要處理的事）、一行 `結果：…`。

### 看結果
- `結果：自動檢查全部通過` → 用 read 看印出的**總覽圖**（`_work/out_png/overview_NN.png`，一張 6 頁）確認版面，就可以交付。
- `結果：還有 N 個問題` → **不可以交付**，照下面改 plan.json 後重跑。

### 改計畫（只改 plan.json，不要改 auto_build.py）
```
python "DIR/auto_build.py" "來源.pptx" "範本.pptx" "輸出.pptx" --plan "_work/plan.json"
```
plan.json 每頁可以改的欄位：

| 欄位 | 用途 |
|---|---|
| `kind` | `cover` 封面、`agenda` 目錄、`table` 表格頁、`chart` 圖表頁、`text` 文字卡片頁、`section` 章節頁、`closing` 結尾、`custom` 自己寫、`skip` 不做（要寫 `reason`，而且會被列成問題） |
| `title`、`subtitle` | 標題、副標（程式判斷錯時改） |
| `cols` | 文字卡片的欄數（`text` 頁） |
| `side` | `"none"`：表格頁不要右側欄（說明改放表格上下） |
| `side_width` | 右側欄寬度（吋，預設 3.4） |
| `script` | `kind: "custom"` 時，自己寫那一頁的 Python 檔（見下方） |

常見情況：

| 看到的問題 | 改法 |
|---|---|
| 卡片太擠或太空 | 改 `cols` |
| 標題抓錯（抓到章節名、圖例） | 改 `title`／`subtitle` |
| 表格頁的側欄不對 | `side: "none"`，或改 `side_width` |
| 某頁版型判斷錯 | 改 `kind`（例如 `text` 改成 `table`） |
| 怎麼改都不對 | 那一頁改 `kind: "custom"` 自己寫，其他頁維持自動 |

### custom 頁
plan.json：`{"src": 8, "kind": "custom", "title": "標題", "script": "_work/p8.py"}`（路徑相對於執行指令的資料夾）。
腳本裡可以直接用 `s`（已建立好、有標題的投影片）、`src_no`（來源頁碼）和 helpers 的所有函式：
```python
x, y, w, h = content_area(s)
card_grid(s, src_cards(src_no), cols=2, area=(x, y, w, h))
```
文字一律用 `src_title`／`src_cards`／`src_paras`／`src_tables`／`src_chart` 從來源讀，**不要手打**。函式說明見下方「helpers.py」。

### 回報給使用者
附上 auto_build 印出的所有「說明」（例如表格用了小字、結尾頁沒放來源的字），列出你改過計畫的頁面和原因。
還有「問題」沒解決時，不可以說完成；要逐條列出，請使用者決定。

## 來源不是簡報（md / txt / csv / xlsx / docx / pdf）
1. PDF 先用 `pdf` 技能抽文字存成 .md。
2. 看內容結構：`python "DIR/extract_source.py" 來源檔`（輸出 JSON）
3. 直接套版：`python "DIR/build_from_text.py" 來源檔 "範本.pptx" "輸出.pptx"`
   - 封面 → 目錄（3–8 個主章節時）→ 每一節 → 結尾頁（範本有才加）；短的含數字條列做成數字卡片；表格用 `table_pages`。
4. 這只是起點：文字來源沒有分頁，**由你決定**內容怎麼對應到投影片。要調整時複製一份 build_from_text.py 修改，
   用 `init(None, "範本.pptx")`，內容用 `from extract_source import extract; doc = extract(路徑)` 讀。
5. 匯出圖片檢查：`python "<skills 資料夾>/powerpoint-pptx/scripts/pptx_tool.py" slides 輸出.pptx -o _work/out_png`（先看總覽圖）。

## 圖片與 PowerPoint
- auto_build 和 `pptx_tool.py slides` 都會自己找 PowerPoint（32/64 位元，從登錄檔找）。**不要用 `where powerpnt` 判斷有沒有裝**：Office 不在 PATH。
- **看圖要省**：先看總覽圖，只有需要細節時才開單頁 `slide_NNN.png`，整個對話盡量 10 張以內。
  有些平台累積十幾張圖後會拒絕請求（例如 `input rejected by content safety screening`），要開新對話。
- 完全沒有 PowerPoint 也沒有 LibreOffice 時：問使用者可不可以安裝 LibreOffice（`winget install TheDocumentFoundation.LibreOffice`），
  同時用 `python -I "DIR/check_bounds.py" 輸出.pptx` 估算，回報時說明沒有做視覺檢查。

## helpers.py（custom 頁、非簡報來源時才需要；很長，不用整個讀）
- `init(來源, 範本)`：開檔並從範本主題挑顏色（`BLUE` 主色、`NAVY` 深色字、`CLOUD` 卡片底、`ZEBRA` 斑馬紋、`GRAY` 次要字）。
- `new_slide("cover"|"content"|"section"|"closing"|版面名稱, 標題)`、`set_subtitle(s, 文字, nth=0)`、`content_area(s)` → (x, y, w, h)
- 讀來源：`src_title(頁)`（標題，字大的在前）、`src_cards(頁)`（[(小標, [內容行])]）、`src_paras(頁)`（一行一筆）、
  `src_tables(頁)`＋`cell_lines(儲存格)`、`src_chart(頁)`（全部數列，可直接 `column_chart(s, **src_chart(頁))`）、`run_color(run)`
- 版型：`agenda`、`stat_cards(s, [(標籤, 數字, 單位)])`、`card_grid(s, cards, cols=2, area=...)`、
  `table_pages(標題, rows, first=s, max_pages=1, area=...)`、`put_table(s, rows, x, y, w, size=11)`、`column_chart`、`progress_bar`
- 基本元件：`box`、`txt`（會預估溢出）、`table`、`cell_set`、`style_header_row`、`zebra_body`
- 檢查：`save(輸出)` 會刪範本範例頁、存檔並印出 MISSING／WARN（自己寫整份腳本時用；auto_build 已內建）。

## 踩過的坑
- 不要編造內容：來源只有標題沒有內容的區塊，只放標題，不要補「無」或自己寫說明；要說明的寫在回報裡。
- 不要把某頁放不下的內容搬到別頁、另開一頁，或整個拿掉；來源一頁就做成一頁（真的放不下先問使用者）。
- 來源的錯字、裝飾字、範本已經有的字（結尾頁的 Thank you）不要硬加；在回報裡說明。
- 模組別取名 `h`（迴圈變數常叫 h 會蓋掉），用 `import helpers as H`。
- 表格空白儲存格預設 18pt 會把列撐高：用 `cell_set`（已處理）。
- 組合圖要用 `src_chart()`，自己讀 `chart.plots[0]` 會漏掉折線。
