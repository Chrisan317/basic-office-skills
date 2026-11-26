---
name: pptx-fill-company-template
description: 套版：把現有內容重做到公司簡報範本（.pptx/.potx）上。來源可以是舊簡報（月報、週報等），或文字類檔案（.md .txt .csv .xlsx .docx；PDF 先轉文字）。用範本的版面、主題色、字型重建每一頁。Use when asked to apply a company template to a deck or document, "套版", "把內容套到公司簡報範本", "換成公司範本".
---

# 套用公司範本（套版）

**重建，不要複製**：先從來源抽出內容，再用範本的版面（layout）建立「新的」投影片。
直接複製來源投影片會把舊的母片、背景一起帶過來，永遠不會像範本。

本技能的腳本在 `<skills 資料夾>/pptx-fill-company-template/scripts/`，以下用 `DIR` 代表。
讀取、匯出圖片、轉 PDF 用 `powerpoint-pptx` 技能的 `pptx_tool.py`（以下用 `PPT` 代表）。
分享本技能時要給**整個資料夾**（SKILL.md + scripts/）。

## 檔案
- `helpers.py`：共用函式庫，直接 import 不要改。
  - `init(來源, 範本)`：開檔，並**從範本的佈景主題挑顏色**，印出挑選結果（例如 `BLUE=0055CC（accent4）`）。
    主色不一定是 accent1：有些範本的 accent1 是很淺的底色，所以會挑「白底上看得清楚、有彩度」的 accent 當主色。
    挑錯時在 init 之後自己指定：`H.BLUE = BLUE = "1F4E79"`（兩個都要改：函式預設值看 `H.BLUE`，你自己的程式看 `BLUE`），或 `init(..., theme=False)` 用預設色。
  - 顏色名稱：`BLUE` 主色（標題字、數字、表頭）、`NAVY` 深色字、`CLOUD` 淺色卡片底、`ZEBRA` 表格斑馬紋、`ICE`/`CYAN` 圖表輔色、`GRAY` 次要文字、`BLACK`、`WHITE`、`RED`
  - `new_slide("cover" | "content" | "section" | "closing" | 版面名稱, 標題)`：自動找範本中合適的版面
  - `set_subtitle(s, 文字, nth=0)`：填副標預留位置（封面常有兩個：nth=0 講者、nth=1 日期）。內容頁會自動把副標移到標題下方、靠左
  - `content_area(slide)`：依投影片大小、標題、已填的副標，回傳可放內容的 (x, y, w, h)，16:9 / 4:3 都適用
  - 基本元件：`box`（圓角卡片）、`txt`（文字框，會預估是否溢出）、`table`、`cell_set`、`style_header_row`、`zebra_body`、`set_border`、`fit`、`norm`、`src_table`、`cell_lines`、`walk_shapes`（會進入群組）、`bar_line_combo`、`save`（會刪掉範本的範例頁）
  - **常用版型**（見下節）：`agenda`、`stat_cards`、`card_grid`、`table_pages`、`column_chart`、`progress_bar`
- `extract_source.py`：把 md / txt / csv / xlsx / docx 轉成統一結構（章節 → 段落 / 條列 / 表格），不摘要、不編造。
- `build_from_text.py`：非簡報來源的通用套版，可直接用，也可以複製後調整。
- `check_bounds.py`：沒辦法輸出圖片時，用來估算有沒有東西超出投影片。

## 常用版型（helpers.py，位置預設用 content_area）
| 函式 | 用途 |
|---|---|
| `agenda(s, ["項目1", "項目2", ...])` | 目錄頁：編號＋項目卡片，5 項內一列，更多兩列 |
| `stat_cards(s, [("生效", 570, "筆"), ("失效", 647, "筆")], y=None)` | 一排數字卡片，回傳下緣 y（下一個元件從這裡往下排） |
| `card_grid(s, [("標題", ["項目", ("主文字", "灰色說明")]), ...], cols=2)` | 2×2 等格狀卡片 |
| `table_pages("標題", rows, first=s, top=y, bold_last=True)` | 表格：自動挑字級、依換行估列高、分頁時各頁平均、多頁標題加 (1/n)；rows 的格子可以是字串或字串 list（多段落） |
| `column_chart(s, 類別, [("數列", 值), ...], line=("合計", 值))` | 原生直條圖，可加一條折線（組合圖） |
| `progress_bar(s, x, y, w, 75.5)` | 達成率進度條 |

## 來源不是簡報（md / txt / csv / xlsx / docx / pdf）
1. PDF 先用 `pdf` 技能抽文字存成 .md。
2. 看內容結構：`python "DIR/extract_source.py" 來源檔`（輸出 JSON）
3. 直接套版：`python "DIR/build_from_text.py" 來源檔 "範本.pptx" "輸出.pptx"`
   - 封面（標題＋短副標）→ 目錄（3–8 個主章節時）→ 每一節 → 結尾頁（範本有才加）
   - 2–4 個短、含數字的條列（「新客戶 12 家」）→ 數字卡片
   - 段落/條列 → 卡片，字級依數量調整，太多分兩欄、再多分頁
   - 表格 → `table_pages`；表格前的短說明放在表格上方；小型數值表另加一頁直條圖（合計欄畫成折線）
4. 這只是起點。正式交付時複製一份調整：文字來源沒有分頁，**由你決定**內容怎麼對應到投影片。
   自己寫腳本時用 `init(None, "範本.pptx")`（沒有來源簡報），內容用 `from extract_source import extract; doc = extract(路徑)` 讀。

## 來源是簡報（pptx）
1. **把來源當資料讀**：
   - `python "PPT" read 來源.pptx` 得到每頁文字、表格、備忘稿
   - 需要位置資訊時用 python-pptx 列出每個 shape 的位置（**群組要遞迴進去**，用 `walk_shapes`）。位置可以判斷哪個標籤配哪個文字框。
   - 表格讀儲存格段落（`cell_lines`）；圖表讀 `chart.plots`（**全部 plot**，不只 `plots[0]`）的 categories 和 series.values；用顏色表示的狀態要讀 run 的顏色。
2. **讀範本**：`python "PPT" layouts 範本.pptx` 看版面名稱、預留位置、尺寸。範本裡的範例頁通常只是示範或樣式說明，參考配色後由 `save()` 自動刪掉。
3. **每頁對應一個版型**：
   | 來源頁 | 做法 |
   |---|---|
   | 封面 | `new_slide("cover", 標題)` + `set_subtitle(s, 講者)` + `set_subtitle(s, 日期, nth=1)` |
   | 目錄/大綱 | `agenda` |
   | 幾個重點數字＋明細 | `stat_cards` 在上，`table_pages(first=s, top=下緣+0.3)` 在下 |
   | 表格 | `table_pages`；來源只有一頁的總表就保留一頁（`sizes=(10, 9)` 小字） |
   | 多個文字區塊 | `card_grid` |
   | 圖表 | `column_chart`（數值用程式從來源圖表讀） |
   | 達成率 | `stat_cards` + `progress_bar` |
   | 結尾 | `new_slide("closing")` |
   其他內容用 `box` + `txt` 自己排；顏色用主題色（標題主色、內文黑、次要灰、淺色卡片底）。
4. **寫腳本建立**：在 `_work` 資料夾寫自己的對應腳本，數值從來源用程式讀，不要手打：
   ```python
   import sys; sys.path.insert(0, r"DIR")
   import helpers as H
   from helpers import *
   src, prs = init("來源.pptx", "範本.pptx")       # 會印出挑到的主題色
   s = new_slide("cover", "月報"); set_subtitle(s, "部門"); set_subtitle(s, "2026/10", nth=1)
   s = new_slide("content", "業績摘要")
   y = stat_cards(s, [("營收", "1,234", "萬"), ("達成率", "104", "%")])
   t0 = src_table(3)                                # 來源第 3 頁的第一個表格
   rows = [[cell_lines(c) for c in r.cells] for r in t0.rows]
   table_pages("業績摘要", rows, first=s, top=y + 0.3, bold_last=True)
   s = new_slide("content", "趨勢")
   column_chart(s, ["1月", "2月"], [("A", [3, 5]), ("B", [2, 4])], line=("合計", [5, 9]))
   save("輸出.pptx")
   ```
5. **最後才刪範本範例頁**：`save()` 會處理。
6. **輸出圖片逐頁檢查**（見下），修正後再輸出一次。最後 `python "PPT" read 輸出.pptx` 和來源比對有沒有漏字。

## 一定要看投影片長什麼樣子
`python "PPT" slides 輸出.pptx` → 每頁 PNG，用 read 工具逐頁看。
這個指令會自動處理：依序試 pywin32 → PowerShell（Office 是 32 位元時先用 32 位元 PowerShell）→ LibreOffice，並印出用了哪一種。**不要自己猜 PowerPoint 的路徑。**

都沒有時（沒有 PowerPoint 也沒有 LibreOffice）：
- 不要停在「無法輸出圖片」。告訴使用者，問可不可以安裝 LibreOffice（`winget install TheDocumentFoundation.LibreOffice`）。
- 同時跑 `python -I "DIR/check_bounds.py" 輸出.pptx` 估算有沒有東西超出投影片。
- 最後回報時清楚說明「沒有做視覺檢查」，並列出最擠的頁面（大表格）請使用者先看。

沒做過上面任一種檢查，**不可以**說投影片沒問題：表格溢出常常只有輸出圖片才看得出來。

## 踩過的坑
- **模組別取名 `h`**：`import helpers as h` 之後，迴圈裡的 `for h, sb in ...`、`h = 高度` 會蓋掉模組，`h.CLOUD` 就變成錯誤。用 `import helpers as H`，或直接用 `BLUE`、`CLOUD`（`init()` 會更新它們）。
- **主題色要看 init() 印出的結果**：如果主色看起來太淺或不像公司色，自己指定 `H.BLUE` 等，再往下做。
- **空白儲存格預設 18pt**，會默默把每一列撐高（31 列的表格就這樣超出投影片）。每個段落都要有和表格字級相同的 `a:endParaRPr sz=...`（`cell_set` 已處理）。
- 來源只有一頁總表，就保留一頁（小字、合併標題、照原本的側欄）。要拆頁先問使用者。
- 組合圖：python-pptx 讀 `chart.plots` 才會拿到全部，`LinePlot` 裡的數列很容易漏。重建用 `column_chart(..., line=...)`。
- 預留位置會繼承置中對齊：副標用 `set_subtitle`（內容頁自動靠左），或給 `align=PP_ALIGN.LEFT`。
- 不同範本的副標預留位置 idx 不同：用 `set_subtitle()`，不要寫死 idx。
- 密集表格（20 列以上）要分頁，或用 9.5–10pt；`table_pages` 會依每格文字換行數估列高（中文字約 1 個字級寬、英數約 0.55）。
- `cell.merge()` 會把被合併儲存格的文字串起來：合併前先清空非起點儲存格。
- 表格樣式：tableStyleId 設 `{2D5ABB26-0587-4C30-8999-92F81FD0307C}`（No Style, No Grid），儲存格自己填色；框線放在 `tcPr` 的填色元素之前（`table`、`set_border` 已處理）。
- 怪的來源結構整理成乾淨的（例如一個標題跨 3 個子欄，合成一欄）；兩行的標題放在同一格的兩個段落。
- **不要編造內容。** 來源空白的段落保留標題並註明空白，並告訴使用者。標題明顯錯字可以修，但要回報。
- 只用顏色表示的圖例/狀態是推測：要跟使用者說明你怎麼對應。
- 不要加裝飾用的色條；用淺色圓角卡片。同類型頁面的標題位置、大小保持一致。

## 最後回報給使用者
列出：每頁的對應方式、所有判斷（拆頁、合併欄位、顏色意義的推測、空白段落）、哪幾頁最需要先檢查。
