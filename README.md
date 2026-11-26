# basic-office-skills

給 AI agent 用的基本辦公文書技能包：讀寫 PDF、Word、Excel、PowerPoint，撰寫 Email、會議記錄、週報、簽呈。

採用通用的 [Agent Skills](https://agentskills.io) 格式（一個資料夾一個 `SKILL.md`），OpenCode、Claude Code 等支援 skills 的 agent 都能用。
特別適合開源/較小的模型（例如 NVIDIA Nemotron）：每個技能附**現成、測試過的 Python 腳本**，模型直接執行，不用每次重寫程式；
全域規則 `AGENTS.md` 讓模型遇到困難先想辦法（寫程式、裝套件），不會直接說「沒有工具、做不到」。

## 內容

| 檔案 | 用途 |
|---|---|
| `AGENTS.md` | 全域規則：先找技能、沒有就自己寫 Python、試 3 種方法才能說做不到、各種檔案的正確讀法、Windows 中文編碼、不覆蓋原檔、完成要驗證 |
| `skills/pdf` | 抽文字＋表格、掃描檔 OCR（rapidocr，支援中文）、轉圖片給視覺模型看、合併/分割/旋轉 |
| `skills/word-docx` | 讀成 Markdown、保留格式的批次取代、建立文件（已處理中文 eastAsia 字型）、轉 PDF、.doc 轉 .docx |
| `skills/excel-xlsx` | 看結構、pandas 分析、做報表、加公式與格式、修改不洗掉原格式、CSV 自動判斷編碼（utf-8/cp950/big5） |
| `skills/powerpoint-pptx` | 讀每頁內容與備忘稿、列出範本版面、用 python-pptx 建簡報、轉 PDF |
| `skills/office-writing` | 台灣職場文書格式：Email、會議記錄、週報、簽呈、公告、摘要 |
| `skills/make-skill` | 說「把這個做法存成技能」，agent 會自己照規格建立新技能 |

## 安裝

把 `skills` 裡的資料夾複製到你的 agent 讀取技能的位置，`AGENTS.md` 放到全域規則的位置：

| Agent | 技能放這裡 | 全域規則 |
|---|---|---|
| OpenCode | `~/.config/opencode/skills/` | `~/.config/opencode/AGENTS.md` |
| Claude Code | `~/.claude/skills/` | 把 `AGENTS.md` 內容併入 `~/.claude/CLAUDE.md` |
| 其他支援 Agent Skills 的工具 | `~/.agents/skills/`（多數工具會讀）或該工具文件指定的位置 | 依該工具的規則檔（多半是 `AGENTS.md`） |

Windows 的 `~` 是 `C:\Users\<你的帳號>`。

> 已經有自己的全域規則檔時，請把內容合併，不要直接覆蓋。
> `AGENTS.md` 的部分內容（例如「這個模型不能用 read 直接開 PDF」）是針對沒有原生 PDF 支援的模型寫的，用其他模型時可視情況調整。

需要 Python 3.10 以上。建議先裝好套件（不裝也行，腳本缺套件時會自動安裝）：

```
python -m pip install pymupdf python-docx openpyxl python-pptx pandas tabulate xlrd
```

Windows 上要把 Word/PPT 轉 PDF，另外需要 `pywin32` 和 Microsoft Office，或改裝 LibreOffice。

裝好後**重新啟動 agent**，應該能在技能清單（多數工具是輸入 `/`）看到 `pdf`、`word-docx` 等技能。

## 使用範例

- 幫我摘要 合約.pdf 的重點
- 把 報價單.docx 裡的「2025年」全部改成「2026年」，另存新檔
- 統計 銷售.xlsx 每個業務每個月的業績，做成新的報表
- 根據 會議筆記.txt 幫我整理成會議記錄
- （教它新做法後）把這個做法存成技能

腳本也可以不透過 agent 直接用，例如：

```
python skills/pdf/scripts/pdf_tool.py extract 合約.pdf
python skills/excel-xlsx/scripts/xlsx_tool.py info 銷售.xlsx
```

## 搭配視覺模型（OpenCode 範例）

`pdf` 技能會優先把掃描檔頁面轉成 PNG，讓模型直接看圖。OpenCode 需要在模型設定宣告支援圖片輸入，例如：

```jsonc
"models": {
  "nvidia/nemotron-3.5-super-vl-120b-a12b": {
    "name": "Nemotron 3.5 Super VL 120B",
    "capabilities": { "tools": true, "input": ["text", "image"], "output": ["text"] }
  }
}
```

模型不支援圖片時，改用 `ocr` 指令轉成文字即可。

## 授權

MIT
