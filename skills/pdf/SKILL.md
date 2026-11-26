---
name: pdf
description: 讀取任何 PDF（機票、合約、發票、說明書、論文等，不論大小）的預設方法：抽取文字與表格、掃描檔辨識(OCR)、轉成圖片查看、合併、分割、旋轉、產生 PDF。read 工具無法開 PDF，一律用本技能。Default way to read or process any PDF file. 唯一例外：使用者要「審查/檢查測試報告」時，改用測試報告審查的專門技能。
---

# PDF 技能

本技能資料夾內有現成腳本 `scripts/pdf_tool.py`（缺套件會自動安裝），**直接執行，不要重寫**。
腳本路徑 = 載入技能時顯示的 base directory + `/scripts/pdf_tool.py`，
通常是 `<skills 資料夾>/pdf/scripts/pdf_tool.py`（OpenCode 是 `~/.config/opencode/skills`，Claude Code 是 `~/.claude/skills`，通用位置是 `~/.agents/skills`）。以下用 `TOOL` 代表這個路徑。

## 讀 PDF 的標準流程
1. 先看概況（頁數、是否掃描檔、目錄）：
   `python "TOOL" info "檔案.pdf"`
2. 有文字層 → 抽取文字與表格到 Markdown 檔，再用 read 工具讀那個 .md：
   `python "TOOL" extract "檔案.pdf"`（輸出 `檔案_內容.md`）
   - 只要部分頁：`--pages 1-3,7`
   - 很長的 PDF（>30 頁）：先讀目錄，再分段抽需要的頁，不要一次全讀。
3. 沒有文字層（掃描檔、圖片）或排版/圖表需要「看」：
   - 方法 A（優先，你是視覺模型）：轉成圖片，再用 read 工具開 PNG 直接看
     `python "TOOL" render "檔案.pdf" --pages 1-3`
   - 方法 B：OCR 成文字（第一次會自動安裝 rapidocr，支援中文）
     `python "TOOL" ocr "檔案.pdf" --pages 1-3`
   - 兩種方法互為備案：一個不行就換另一個。

## 其他操作
- 合併：`python "TOOL" merge "輸出.pdf" "a.pdf" "b.pdf"`
- 分割/抽頁：`python "TOOL" split "檔案.pdf" --pages 3-5 -o "輸出.pdf"`
- 旋轉：`python "TOOL" rotate "檔案.pdf" --angle 90 --pages 2 -o "輸出.pdf"`
- 有密碼的 PDF：問使用者密碼，然後在 Python 裡 `doc.authenticate("密碼")`。

## 產生 PDF
- 最穩的方式：先做成 Word（載入 `word-docx` 技能），再轉 PDF。
- 簡單文字/表格也可用 PyMuPDF 直接產生，但中文要指定字型，否則會變方框：
  ```python
  import pymupdf
  doc = pymupdf.open(); page = doc.new_page()
  page.insert_font(fontname="jh", fontfile=r"C:\Windows\Fonts\msjh.ttc")
  page.insert_text((72, 72), "中文內容", fontname="jh", fontsize=12)
  doc.save("輸出.pdf")
  ```

## 腳本不夠用時
自己寫 Python，用 `import pymupdf`（不要用舊的 `import fitz`）。常用 API：
`page.get_text("text"/"blocks"/"dict")`、`page.find_tables()`、`page.get_images()`、
`page.search_for("關鍵字")`、`page.add_highlight_annot(rect)`、`page.insert_text()`。
填寫 PDF 表單：`for w in page.widgets(): w.field_value = "..."; w.update()`。
