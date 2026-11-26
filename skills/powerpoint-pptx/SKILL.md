---
name: powerpoint-pptx
description: 讀取、建立、修改 PowerPoint 簡報（.pptx）：讀每頁內容與備忘稿、用公司範本做簡報、加表格圖表圖片、轉 PDF。Use for any task involving PowerPoint - reading slides, creating a presentation or deck, editing slides, or converting pptx to PDF.
---

# PowerPoint 技能

現成腳本：本技能資料夾的 `scripts/pptx_tool.py`（缺套件會自動安裝），**直接用，不要重寫**。
路徑通常是 `<skills 資料夾>/powerpoint-pptx/scripts/pptx_tool.py`（OpenCode 是 `~/.config/opencode/skills`，Claude Code 是 `~/.claude/skills`，通用位置是 `~/.agents/skills`），以下用 `TOOL` 代表。

## 讀取
`python "TOOL" read "簡報.pptx"` → 輸出 `簡報_內容.md`（每頁標題、文字、表格、備忘稿），再用 read 工具讀。
需要看版面設計：`python "TOOL" slides "簡報.pptx"` 每頁輸出成 PNG，再用 read 工具看圖。

## 套用公司範本（套版）
要把現有內容（簡報、Word、Excel、Markdown…）重做到公司範本上，改用 `pptx-fill-company-template` 技能。
本技能處理一般的讀取、建立、修改。

## 建立簡報
1. **有公司範本就一定用範本**。先看範本有哪些版面：`python "TOOL" layouts "範本.pptx"`
2. 先規劃大綱（每頁標題 + 3~5 個重點），給使用者確認後再做。
3. 用 python-pptx 產生：
```python
from pptx import Presentation
from pptx.util import Pt, Cm
prs = Presentation("範本.pptx")        # 沒範本就 Presentation()，並設 16:9：
# prs.slide_width, prs.slide_height = Cm(33.867), Cm(19.05)
s = prs.slides.add_slide(prs.slide_layouts[1])   # 1 通常是「標題及內容」
s.shapes.title.text = "標題"
body = s.placeholders[1].text_frame
body.text = "第一點"
p = body.add_paragraph(); p.text = "子項目"; p.level = 1
s.notes_slide.notes_text_frame.text = "講者備忘稿"
s.shapes.add_picture("圖.png", Cm(2), Cm(5), width=Cm(15))
prs.save("輸出.pptx")
```
- 中文字型：設定文字後 `for r in p.runs: r.font.name = "微軟正黑體"`。
- 表格：`s.shapes.add_table(列, 欄, Cm(2), Cm(4), Cm(29), Cm(8)).table`，`table.cell(r, c).text = "..."`
- 圖表：用 `pptx.chart.data.CategoryChartData` + `s.shapes.add_chart(...)`，或用 matplotlib 畫圖再插入圖片。
- 每頁文字不要太多：一頁最多 6 行重點，每行不超過 20 字。

## 修改現有簡報
用 `Presentation("原檔.pptx")` 開啟、修改、另存 `原檔_修改.pptx`。
取代文字要逐個 run 改（`for r in p.runs: r.text = r.text.replace(...)`），才不會掉格式。

## 轉 PDF、匯出圖片、找 PowerPoint
- 轉 PDF：`python "TOOL" topdf "簡報.pptx"`
- 每頁匯出 PNG：`python "TOOL" slides "簡報.pptx"`
- 找 Office / LibreOffice 裝在哪：`python "TOOL" office`
- **不要自己猜 PowerPoint 的路徑。** 32 位元 Office 在 `C:\Program Files (x86)\...`，64 位元在 `C:\Program Files\...`，每台電腦不一樣。
  腳本會從登錄檔找，PowerPoint 自動化（COM）失敗時也會自動改用 32 位元 PowerShell 再試，最後才用 LibreOffice。
- 自己寫程式要用這些功能時：`sys.path.insert(0, r"DIR")` 再 `from office_convert import to_pdf, export_slides, find_app`（DIR = 本技能的 scripts 資料夾）。

## 完成後驗證
對輸出檔跑 `read` 檢查每頁內容，再用 `slides` 匯出圖片，看有沒有文字超出版面或重疊。
