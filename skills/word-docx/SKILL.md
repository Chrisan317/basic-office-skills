---
name: word-docx
description: 讀取、建立、修改 Word 文件（.docx / .doc）：讀內容與表格、套範本產生文件、批次取代文字、轉 PDF。Use for any task involving a Word document - read, summarize, create, edit, fill a template, find-and-replace, or convert docx to PDF.
---

# Word 技能

現成腳本在本技能資料夾的 `scripts/`（缺套件會自動安裝），**直接用，不要重寫**。
路徑 = 載入技能時顯示的 base directory + `/scripts/...`，
通常是 `<skills 資料夾>/word-docx/scripts/`（OpenCode 是 `~/.config/opencode/skills`，Claude Code 是 `~/.claude/skills`，通用位置是 `~/.agents/skills`）。以下用 `DIR` 代表。

## 讀取
`python "DIR/docx_tool.py" read "檔案.docx"` → 輸出 `檔案_內容.md`（標題、清單、表格依原順序），再用 read 工具讀。
- 舊格式 `.doc`：先 `python "DIR/docx_tool.py" fromdoc "檔案.doc"` 轉成 .docx。
- 需要看版面/圖片：先 `topdf` 轉 PDF，再用 `pdf` 技能的 render 轉圖片看。

## 修改現有文件
- 批次取代文字（保留格式，含表格、頁首頁尾）：
  `python "DIR/docx_tool.py" replace "原檔.docx" -o "原檔_修改.docx" --pair "舊=>新" --pair "2025年=>2026年"`
- 其他修改自己寫 Python（python-docx）。**用原檔載入再另存新檔**，不要重新建立文件，才不會掉格式。
- 填範本：如果範本裡有 `{{客戶名稱}}` 這類標記，直接用 replace；沒有標記就先 read 看結構，再用程式定位段落/表格儲存格填入。

## 建立新文件
用 `docx_helpers.py`，它已處理好**中文字型**（只設 `run.font.name` 中文不會變，要設 eastAsia）：
```python
import sys; sys.path.insert(0, r"DIR")
from docx_helpers import new_document, add_table, set_run_font
doc = new_document(font="微軟正黑體", size=12)   # 或 new_document(template="公司範本.docx")
doc.add_heading("報告標題", level=1)
doc.add_paragraph("內文……")
add_table(doc, ["項目", "數量", "金額"], [["A", 2, 1000], ["B", 1, 500]])
doc.add_page_break()
doc.save("輸出.docx")
```
- 有公司範本就用範本，格式最一致。
- 圖片：`doc.add_picture("圖.png", width=docx.shared.Cm(15))`
- 頁首：`doc.sections[0].header.paragraphs[0].text = "公司名稱"`

## 轉 PDF
`python "DIR/docx_tool.py" topdf "檔案.docx"`（Windows 用 Word 轉，沒有 Word 改用 LibreOffice）。

## 完成後驗證
對輸出檔再跑一次 `read`，確認內容正確、沒有遺漏、中文沒有變亂碼。
