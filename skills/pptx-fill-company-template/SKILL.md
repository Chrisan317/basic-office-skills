---
name: pptx-fill-company-template
description: 套版：把現有內容重做到公司簡報範本（.pptx/.potx）上。來源可以是舊簡報（月報、週報等），或文字類檔案（.md .txt .csv .xlsx .docx；PDF 先轉文字）。用範本的版面、主題色、字型重建每一頁。Use when asked to apply a company template to a deck or document, "套版", "把內容套到公司簡報範本", "換成公司範本".
---

# 套用公司範本（套版）

**重建，不要複製**：先從來源抽出內容，再用範本的版面（layout）建立「新的」投影片。
直接複製來源投影片會把舊的母片、背景一起帶過來，永遠不會像範本。

本技能的腳本在 `<skills 資料夾>/pptx-fill-company-template/scripts/`，以下用 `DIR` 代表。
讀取、匯出圖片、轉 PDF 用 `powerpoint-pptx` 技能的 `pptx_tool.py`（檔名是**底線**；以下用 `PPT` 代表）。
它在本技能資料夾的**同一層**：`<skills 資料夾>/powerpoint-pptx/scripts/pptx_tool.py`，`init()` 也會印出完整路徑（`helpers.PPT`）。**不要搜尋、不要自己另寫一套讀簡報的程式。**
`helpers.py` 很長，**不用整個讀**：要用的函式和參數都列在下面。
分享本技能時要給**整個資料夾**（SKILL.md + scripts/）。

## 檔案
- `helpers.py`：共用函式庫，直接 import 不要改。
  - `init(來源, 範本)`：開檔，並**從範本的佈景主題挑顏色**，印出挑選結果（例如 `BLUE=0055CC（accent4）`）。
    主色不一定是 accent1：有些範本的 accent1 是很淺的底色，所以會挑「白底上看得清楚、有彩度」的 accent 當主色。
    挑錯時在 init 之後自己指定：`H.BLUE = BLUE = "1F4E79"`（兩個都要改：函式預設值看 `H.BLUE`，你自己的程式看 `BLUE`），或 `init(..., theme=False)` 用預設色。
  - 顏色名稱：`BLUE` 主色（標題字、數字、表頭）、`NAVY` 深色字、`CLOUD` 淺色卡片底、`ZEBRA` 表格斑馬紋、`ICE`/`CYAN` 圖表輔色、`GRAY` 次要文字、`BLACK`、`WHITE`、`RED`
  - `new_slide("cover" | "content" | "section" | "closing" | 版面名稱, 標題)`：自動找範本中合適的版面
  - `set_subtitle(s, 文字, nth=0)`：填副標預留位置（封面常有兩個：nth=0 講者、nth=1 日期）。內容頁會自動把副標移到標題下方、靠左；版面沒有副標預留位置時，會在標題下方補一個文字框
  - `content_area(slide)`：依投影片大小、標題、已填的副標，回傳可放內容的 (x, y, w, h)，16:9 / 4:3 都適用
  - **讀來源**（數值、文字一律用這些讀，不要手打、不要自己寫解析）：
    - `src_title(頁)`：來源頁的標題文字 list（字最大的在前），放進 `new_slide` 的標題，**不要再放進內文**
    - `src_cards(頁)`：把散放的文字框依位置分組成 `[(小標, [內容行...]), ...]`（並排卡片、左右欄、日期＋說明），直接 `card_grid(s, src_cards(頁), cols=…)`
    - `src_paras(頁)`：該頁所有文字框的文字，**一行一筆**、由上而下（同一個文字框裡的兩行會拆開）；`pos=True` 附位置。只用來看內容或抓數字，不要整串倒進一個文字框
    - `src_tables(頁)`：該頁**所有**表格（一頁常常不只一個）；`cell_lines(儲存格)` 讀儲存格文字，一行一筆（段落和 Shift+Enter 換行都會拆開）
    - `run_color(run)`：文字顏色（`'FF0000'`、`'accent1'`、`'tx1'`…，沒設定是 None）。用顏色表示狀態（● 紅＝延期）時用
    - `src_chart(頁)`：圖表的**全部**數列（含組合圖的折線），直接 `column_chart(s, **src_chart(頁))`
  - `save(輸出)`：刪掉範本範例頁、存檔，並自動檢查，印出：
    - `MISSING 來源第 N 頁：…`：來源有、輸出沒有的字
    - `WARN 第 N 頁：A 和 B 重疊`／`超出投影片`：文字疊到表格、其他文字、範本的圖案（例如結尾頁的 Thank You）
    - `WARN 來源第 N 頁圖表…`：圖表數列少了、折線變長條、數值和來源不同（手打）
    - `WARN 來源第 N 頁的字放到了輸出第 M 頁`：內容被搬到別頁（例如把放不下的側欄塞到下一頁）
    - `WARN 來源第 N 頁變成輸出第 A、B 頁`：來源一頁被拆成好幾頁（例如另開一頁「說明」放側欄）
    - `WARN 來源第 N 頁的數字在輸出第 M 頁找不到`：數字被濃縮或省略（MISSING 是整份比對，數字常在別頁碰巧出現）
    - `WARN 第 N 頁有來源沒有的文字`：自己寫的解釋、結論、「來源沒有提供」之類的說明
    - `WARN 第 N 頁：來源…是 K 個分開的區塊，輸出變成一整串清單`、`內文又出現一次標題`：文字頁沒排版
    - `WARN table … 拆成 N 頁`、`overflow?`、`stat_cards…`、`card_grid…`
  - 基本元件：`box`（圓角卡片）、`txt`（文字框，會預估是否溢出）、`table`、`cell_set`、`style_header_row`、`zebra_body`、`set_border`、`fit`、`norm`、`src_table`、`walk_shapes`（會進入群組）、`bar_line_combo`
  - **常用版型**（見下節）：`agenda`、`stat_cards`、`card_grid`、`table_pages`、`put_table`、`column_chart`、`progress_bar`
- `extract_source.py`：把 md / txt / csv / xlsx / docx 轉成統一結構（章節 → 段落 / 條列 / 表格），不摘要、不編造。
- `build_from_text.py`：非簡報來源的通用套版，可直接用，也可以複製後調整。
- `check_bounds.py`：沒辦法輸出圖片時，用來估算有沒有東西超出投影片。

## 常用版型（helpers.py，位置預設用 content_area）
| 函式 | 用途 |
|---|---|
| `agenda(s, ["項目1", "項目2", ...])` | 目錄頁：編號＋項目卡片，5 項內一列，更多兩列 |
| `stat_cards(s, [("有效", 120, "筆"), ("到期", 35, "筆")], y=None)` | 一排數字卡片，回傳下緣 y（下一個元件從這裡往下排）。數值只放數字，一張卡片一個值 |
| `card_grid(s, [("標題", ["項目", ("主文字", "灰色說明")]), ...], cols=2)` | 2×2 等格狀卡片；內容多時自動縮小字級 |
| `table_pages("標題", rows, first=s, top=y, bold_last=True, max_pages=1)` | 表格：自動挑字級、依換行估列高；rows 的格子可以是字串或字串 list（多段落）。**來源是一頁的表格一定加 `max_pages=1`**：會縮到最小 8pt、列高壓緊來放進一頁。不加時放不下會自動分頁（標題加 (1/n)） |
| `put_table(s, rows, x, y, w, size=11)` | 把表格放在指定位置、不分頁，回傳 (table, 下緣 y)：一頁兩個表格、表格旁邊有側欄時用（側欄寬度先扣掉） |
| `column_chart(s, 類別, [("數列", 值), ...], line=("合計", 值))` | 原生直條圖，可加折線（組合圖，多條用 list）；來源是圖表時用 `column_chart(s, **src_chart(頁))` |
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
   - `python "PPT" read 來源.pptx -o _work/來源_內容.md` 得到每頁文字、表格、備忘稿（加 `-o` 放進 `_work`，不要放在來源檔旁邊）
   - 寫腳本時每頁先印出 `src_title(頁)`、`src_cards(頁)`、`src_tables(頁)` 看清楚有哪些東西，再決定版型。`src_paras(頁, pos=True)` 附位置，可以判斷哪個標籤配哪個值。
   - 圖表一律用 `src_chart(頁)`（自己讀 `chart.plots[0]` 會漏掉折線）；用顏色表示的狀態要讀 run 的顏色。
2. **讀範本**：`python "PPT" layouts 範本.pptx` 看版面名稱、預留位置、尺寸。範本裡的範例頁通常只是示範或樣式說明，參考配色後由 `save()` 自動刪掉。
3. **每頁對應一個版型**：
   | 來源頁 | 做法 |
   |---|---|
   | 封面 | `new_slide("cover", 標題)` + `set_subtitle(s, 講者)` + `set_subtitle(s, 日期, nth=1)` |
   | 目錄/大綱 | `agenda` |
   | 幾個重點數字＋明細 | `stat_cards` 在上，`table_pages(first=s, top=下緣+0.3, max_pages=1)` 在下 |
   | 表格 | `table_pages(..., max_pages=1)`：來源一頁就保留一頁；8pt 還放不下時 save() 會警告，先問使用者要不要拆頁 |
   | 一頁兩個表格、表格＋側欄 | `put_table` 放在指定位置（或 `table_pages(..., area=(x, y, w, h))`） |
   | 多個文字區塊、標籤＋內容（日期/講者、負責人/地點、左右兩欄…） | `card_grid(s, src_cards(頁), cols=區塊數或 2)`；標題用 `src_title(頁)[0]` |
   | 只有一段條列（沒有小標） | `card_grid(s, [("", src_cards(頁)[0][1])], cols=1)`，或 `txt`；條列不要包含標題 |
   | 圖表 | `column_chart(s, **src_chart(頁))` |
   | 達成率 | `stat_cards` + `progress_bar` |
   | 結尾 | `new_slide("closing")` |
   其他內容用 `box` + `txt` 自己排；顏色用主題色（標題主色、內文黑、次要灰、淺色卡片底）。
   **不要把整頁文字一行一行倒進一個 txt**：那只是換了背景，沒有排版。先分出「標籤／數值／說明」再選版型。
4. **寫腳本建立**：在 `_work` 資料夾寫自己的對應腳本，數值從來源用程式讀，不要手打：
   ```python
   import sys; sys.path.insert(0, r"DIR")
   import helpers as H
   from helpers import *
   src, prs = init("來源.pptx", "範本.pptx")       # 會印出挑到的主題色
   s = new_slide("cover", "月報"); set_subtitle(s, "部門"); set_subtitle(s, "2026/10", nth=1)
   s = new_slide("content", "業績摘要")
   print(src_paras(3))                              # 先看：['營收: 1,234萬', '達成率: 104%', ...]
   y = stat_cards(s, [("營收", "1,234", "萬"), ("達成率", "104", "%")])
   ts = src_tables(3)                               # 一頁可能有好幾個表格，每個都要放
   rows = lambda t: [[cell_lines(c) for c in r.cells] for r in t.rows]
   x, _, w, _ = content_area(s)
   tb, y2 = put_table(s, rows(ts[0]), x, y + 0.3, w, bold_last=True)
   for t in ts[1:]:                                 # 同一頁的其他表格接著往下放
       tb, y2 = put_table(s, rows(t), x, y2 + 0.3, w)
   s = new_slide("content", "明細")
   table_pages("明細", rows(src_tables(4)[0]), first=s, max_pages=1)   # 大表格，來源一頁就保留一頁
   s = new_slide("content", "趨勢")
   column_chart(s, **src_chart(4))                  # 來源第 4 頁的圖表，全部數列
   s = new_slide("content", src_title(5)[0])        # 文字頁：標題用來源的標題
   card_grid(s, src_cards(5), cols=3)               # 並排的區塊 → 卡片（小標＋內容）
   save("輸出.pptx")                                # 會印出 MISSING / WARN
   ```
5. **處理 save() 印出的每一條**：
   - `MISSING 來源第 N 頁：…` 是來源有、輸出沒有的字（常見：漏掉第二個表格、說明文字、季度數字）。補上，或在回報裡說明為什麼刻意省略。
     **來源的錯字、裝飾字、範本已經有的字（結尾頁的 Thank you、公司名）不要硬加回去**，在回報裡說明就好。不要為了消掉 MISSING 在投影片上疊文字。
     **也不要把字搬到別頁**（例如把表格頁的標題、圖例、側欄塞到後面的統計頁）：每頁的內容要留在對應的那一頁。
   - `WARN 來源第 N 頁的字放到了輸出第 M 頁`：把那些字放回原本那頁。表格佔滿放不下側欄時，用 `put_table(s, rows, x, y, w)` 把表格寬度縮小（例如 8.3 吋），右邊放側欄；來源頁上的大標題放進標題，圖例放副標。
   - `WARN 來源第 N 頁變成輸出第 A、B 頁`：合回一頁。表格＋說明：`put_table` 把表格寬度縮到約 8.3 吋，右邊 `box`＋`txt` 放說明；圖例放副標。另開「說明」頁也算搬頁。
   - `WARN 來源第 N 頁的數字…找不到`：把那些數字放回去。例如來源每一行是「A：23 B：17 比率：73%」，卡片只放了比率，就在卡片下方用 `txt` 或小表格列出每一行。
   - `WARN 第 N 頁有來源沒有的文字`：刪掉。投影片上只放來源的內容；來源某區塊只有標題沒有內容，就只放標題，說明寫在給使用者的回報裡。
   - `WARN 第 N 頁：來源…是 K 個分開的區塊，輸出變成一整串清單`：改成 `card_grid(s, src_cards(來源頁), cols=…)`。
   - `WARN 第 N 頁：內文又出現一次標題`：內文用 `src_cards()`（不含標題），或從清單拿掉 `src_title()` 的字。
   - `WARN … 重疊`、`超出投影片`：改位置或大小。總覽圖太小看不出來，**有這種警告的頁面要打開單頁圖確認**。
   - `WARN 來源第 N 頁圖表…`：改用 `column_chart(s, **src_chart(N))`。
   - `WARN table … 拆成 N 頁`：來源一頁的表格加 `max_pages=1`。
   - `WARN stat_cards 數值…不像單一數字`：把兩行當成一個值了，用 `src_paras()` 一行一行讀。
   - `WARN overflow?`、`card_grid 卡片…沒有內容`：版面或解析有問題，修好再存。
6. **輸出圖片檢查**（見下），修正後再輸出一次，直到沒有 WARN、MISSING 都補上或說明過、圖片看起來沒問題。
   save() 還有 WARN 時，**不可以**說「全部檢查通過」。

## 一定要看投影片長什麼樣子
`python "PPT" slides 輸出.pptx -o _work/out_png` → 每頁一張 `slide_NNN.png`，另外合成**總覽圖** `overview_NN.png`（一張 6 頁，左上角紅底是頁碼）。
- **先看總覽圖**（17 頁只要 3 張），找出可疑的頁面，只有需要看細節時才開單頁圖。
- **整個對話的看圖張數越少越好（建議 10 張以內）**：圖片會一直留在對話裡、每次都重送。
  有些平台累積十幾張圖後會拒絕請求（例如 NVIDIA 的 `input rejected by content safety screening`），之後同一個對話怎麼重試都會失敗，只能開新對話。
- 修改後重新匯出，只看改過的那幾頁或新的總覽圖，不要每頁重看。
這個指令會自動處理：依序試 pywin32 → PowerShell（Office 是 32 位元時先用 32 位元 PowerShell）→ LibreOffice，並印出用了哪一種。
- **不要自己找 PowerPoint**：`where powerpnt`、`Get-Command powerpnt` 一定找不到（Office 不在 PATH，32 位元裝在 `Program Files (x86)`）。**有沒有裝，以 `slides` 指令的結果為準。**
- 用 PowerPoint 匯出時，最後會印出「**PowerPoint 實際排版檢查**」：PowerPoint 排版後真的超出投影片或文字框的東西（例如表格列被撐高、跑出投影片底部）。有 `WARN` 就要修，這比看圖準。

`slides` 指令印出所有方法都失敗時（沒有 PowerPoint 也沒有 LibreOffice）：
- 不要停在「無法輸出圖片」。告訴使用者，問可不可以安裝 LibreOffice（`winget install TheDocumentFoundation.LibreOffice`）。
- 同時跑 `python -I "DIR/check_bounds.py" 輸出.pptx` 估算有沒有東西超出投影片。
- 最後回報時清楚說明「沒有做視覺檢查」，並列出最擠的頁面（大表格）請使用者先看。

沒做過上面任一種檢查，**不可以**說投影片沒問題：表格溢出常常只有輸出圖片才看得出來。

## 踩過的坑
- **模組別取名 `h`**：`import helpers as h` 之後，迴圈裡的 `for h, sb in ...`、`h = 高度` 會蓋掉模組，`h.CLOUD` 就變成錯誤。用 `import helpers as H`，或直接用 `BLUE`、`CLOUD`（`init()` 會更新它們）。
- **主題色要看 init() 印出的結果**：如果主色看起來太淺或不像公司色，自己指定 `H.BLUE` 等，再往下做。
- **空白儲存格預設 18pt**，會默默把每一列撐高（31 列的表格就這樣超出投影片）。每個段落都要有和表格字級相同的 `a:endParaRPr sz=...`（`cell_set` 已處理）。
- 來源只有一頁總表，就保留一頁：`table_pages(..., max_pages=1)`（小字、合併標題、照原本的側欄）。要拆頁先問使用者。
- 組合圖：自己讀 `chart.plots[0]` 會漏掉 `LinePlot` 裡的數列。用 `src_chart()`，它也會跳過沒有值的隱藏數列。
- 同一個文字框常有好幾行（「有效: 120筆」換行「到期: 35筆」）：用 `text_frame.text` 一次讀會黏在一起，用 `src_paras()`。
- 預留位置會繼承置中對齊：副標用 `set_subtitle`（內容頁自動靠左），或給 `align=PP_ALIGN.LEFT`。
- 不同範本的副標預留位置 idx 不同：用 `set_subtitle()`，不要寫死 idx。
- 密集表格（20 列以上）：來源一頁就用 `max_pages=1`（最小 8pt）；`table_pages` 會依每格文字換行數估列高（中文字約 1 個字級寬、英數約 0.55）。
- `cell.merge()` 會把被合併儲存格的文字串起來：合併前先清空非起點儲存格。
- 表格樣式：tableStyleId 設 `{2D5ABB26-0587-4C30-8999-92F81FD0307C}`（No Style, No Grid），儲存格自己填色；框線放在 `tcPr` 的填色元素之前（`table`、`set_border` 已處理）。
- 怪的來源結構整理成乾淨的（例如一個標題跨 3 個子欄，合成一欄）；兩行的標題放在同一格的兩個段落。
- **不要編造內容。** 來源空白的段落保留標題並註明空白，並告訴使用者。標題明顯錯字可以修，但要回報。
- 只用顏色表示的圖例/狀態是推測：要跟使用者說明你怎麼對應。
- 不要加裝飾用的色條；用淺色圓角卡片。同類型頁面的標題位置、大小保持一致。

## 最後回報給使用者
列出：每頁的對應方式、所有判斷（拆頁、合併欄位、顏色意義的推測、空白段落）、哪幾頁最需要先檢查。
