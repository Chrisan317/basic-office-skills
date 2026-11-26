---
name: pptx-fill-company-template
description: 套版：把現有內容重做到公司簡報範本（.pptx/.potx）上。來源可以是舊簡報（月報、週報等），或文字類檔案（.md .txt .csv .xlsx .docx；PDF 先轉文字）。用範本的版面、主題色、字型重建每一頁。Use when asked to apply a company template to a deck or document, "套版", "把內容套到公司簡報範本", "換成公司範本".
---

# 套用公司範本（套版）

**重建，不要複製**：先從來源抽出內容，再用範本的版面（layout）建立「新的」投影片。
直接複製來源投影片會把舊的母片、背景一起帶過來，永遠不會像範本。

本技能的腳本在 `<skills 資料夾>/pptx-fill-company-template/scripts/`，以下用 `DIR` 代表。
讀取、匯出圖片、轉 PDF 用 `powerpoint-pptx` 技能的 `pptx_tool.py`（以下用 `PPT` 代表）。

## 檔案
- `helpers.py`：共用函式庫，直接 import 不要改。
  - `init(來源, 範本)`：開檔，並**自動從範本的佈景主題讀主色**（之後用 `h.BLUE`、`h.CLOUD` 取顏色）
  - `new_slide("cover" | "content" | "section" | "closing" | 版面名稱, 標題)`：自動找範本中合適的版面
  - `content_area(slide)`：依投影片大小與標題位置，回傳可放內容的 (x, y, w, h)，16:9 / 4:3 都適用
  - `set_subtitle`、`box`（圓角卡片）、`txt`（文字框，會預估是否溢出）、`table`、`cell_set`、`style_header_row`、`zebra_body`、`set_border`、`fit`、`norm`、`src_table`、`cell_lines`、`walk_shapes`（會進入群組）、`bar_line_combo`（長條＋折線組合圖）、`save`（會刪掉範本的範例頁）
- `extract_source.py`：把 md / txt / csv / xlsx / docx 轉成統一結構（章節 → 段落 / 條列 / 表格），不摘要、不編造。
- `build_from_text.py`：非簡報來源的通用套版，可直接用，也可以複製後調整。
- `check_bounds.py`：沒辦法輸出圖片時，用來估算有沒有東西超出投影片。

## 來源不是簡報（md / txt / csv / xlsx / docx / pdf）
1. PDF 先用 `pdf` 技能抽文字存成 .md。
2. 看內容結構：`python "DIR/extract_source.py" 來源檔`（輸出 JSON）
3. 直接套版：`python "DIR/build_from_text.py" 來源檔 "範本.pptx" "輸出.pptx"`
   封面（標題）→ 每個章節一頁（條列/段落做成卡片，表格自動分頁並平均分配）→ 結尾頁（範本有才加）。
4. 這只是起點。正式交付時複製一份調整：幾個重點數字用數字卡片、數值表格改成原生圖表（`add_chart` + `bar_line_combo`）、長清單分兩欄、長文件加章節頁。
   文字來源沒有分頁，**由你決定**內容怎麼對應到投影片。

## 來源是簡報（pptx）
1. **把來源當資料讀**：
   - `python "PPT" read 來源.pptx` 得到每頁文字、表格、備忘稿
   - 需要位置資訊時用 python-pptx 列出每個 shape 的位置（**群組要遞迴進去**，用 `walk_shapes`）。位置可以判斷哪個標籤配哪個文字框。
   - 表格讀儲存格段落；圖表讀 `chart.plots`（**全部 plot**，不只 `plots[0]`）的 categories 和 series.values；用顏色表示的狀態要讀 run 的顏色。
2. **讀範本**：`python "PPT" layouts 範本.pptx` 看版面名稱、預留位置、尺寸。範本裡的範例頁通常只是示範或樣式說明，參考配色後由 `save()` 自動刪掉。
3. **每頁對應一個版面**：表格/卡片頁用「只有標題」的版面（`new_slide("content")`），頭尾用 cover / closing。標題、副標用預留位置；其他內容用 shape 加上，顏色用範本主題色（標題主色、內文黑、次要灰、淺色卡片底）。
4. **寫腳本建立**：在 `_work` 資料夾寫自己的對應腳本，數值從來源用程式讀，不要手打：
   ```python
   import sys; sys.path.insert(0, r"DIR")
   import helpers as h
   from helpers import *
   src, prs = init("來源.pptx", "範本.pptx")
   s = new_slide("cover", "月報"); set_subtitle(s, "2026/10")
   s = new_slide("content", "業績摘要")
   x, y, w, hh = content_area(s)
   box(s, x, y, w / 2 - 0.1, 1.4, h.CLOUD)
   txt(s, x + 0.15, y + 0.1, 3, 0.8, [dict(text="1,234", size=36, bold=True, color=h.BLUE)])
   t0 = src_table(3)          # 來源第 3 頁的第一個表格
   save("輸出.pptx")
   ```
5. **最後才刪範本範例頁**：`save()` 會處理。
6. **輸出圖片逐頁檢查**（見下），修正後再輸出一次。最後 `python "PPT" read 輸出.pptx` 和來源比對有沒有漏字。

## 一定要看投影片長什麼樣子
`python "PPT" slides 輸出.pptx` → 每頁 PNG，用 read 工具逐頁看。
這個指令會自動處理：PowerPoint 在 32 或 64 位元路徑、COM 失敗時改用 32 位元 PowerShell、沒有 PowerPoint 時改用 LibreOffice。**不要自己猜 PowerPoint 的路徑。**

都沒有時（沒有 PowerPoint 也沒有 LibreOffice）：
- 不要停在「無法輸出圖片」。告訴使用者，問可不可以安裝 LibreOffice（`winget install TheDocumentFoundation.LibreOffice`）。
- 同時跑 `python -I "DIR/check_bounds.py" 輸出.pptx` 估算有沒有東西超出投影片。
- 最後回報時清楚說明「沒有做視覺檢查」，並列出最擠的頁面（大表格）請使用者先看。

沒做過上面任一種檢查，**不可以**說投影片沒問題：表格溢出常常只有輸出圖片才看得出來。

## 踩過的坑
- **空白儲存格預設 18pt**，會默默把每一列撐高（31 列的表格就這樣超出投影片）。每個段落都要有和表格字級相同的 `a:endParaRPr sz=...`（`cell_set` 已處理）。
- 來源只有一頁總表，就保留一頁（小字、合併標題、照原本的側欄）。要拆頁先問使用者。
- 組合圖：python-pptx 讀 `chart.plots` 才會拿到全部，`LinePlot` 裡的數列很容易漏。重建長條＋折線用 `bar_line_combo`。
- 預留位置會繼承置中對齊：副標要明確設定段落對齊。
- 不同範本的副標預留位置 idx 不同：用 `set_subtitle()`，不要寫死 idx。
- 密集表格（20 列以上）要分頁，或用 9.5–10pt、列高約 0.23–0.31 吋；用行數估列高（中文字約 1 個字級寬、英數約 0.55），確認 y + 總高度 < 投影片高度。
- `cell.merge()` 會把被合併儲存格的文字串起來：合併前先清空非起點儲存格。
- 表格樣式：tableStyleId 設 `{2D5ABB26-0587-4C30-8999-92F81FD0307C}`（No Style, No Grid），儲存格自己填色；框線放在 `tcPr` 的填色元素之前（`set_border` 已處理）。
- 怪的來源結構整理成乾淨的（例如一個標題跨 3 個子欄，合成一欄）；兩行的標題放在同一格的兩個段落。
- **不要編造內容。** 來源空白的段落保留標題並註明空白，並告訴使用者。標題明顯錯字可以修，但要回報。
- 只用顏色表示的圖例/狀態是推測：要跟使用者說明你怎麼對應。
- 不要加裝飾用的色條；用淺色圓角卡片。同類型頁面的標題位置、大小保持一致。

## 最後回報給使用者
列出：每頁的對應方式、所有判斷（拆頁、合併欄位、顏色意義的推測、空白段落）、哪幾頁最需要先檢查。
