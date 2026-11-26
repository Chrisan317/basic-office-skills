---
name: excel-xlsx
description: 讀取、分析、建立、修改 Excel 試算表與 CSV（.xlsx .xls .xlsm .csv）：看內容、統計加總、樞紐分析、比對兩份資料、做報表、畫圖表、加公式與格式。Use for any task involving Excel or CSV files - reading, analysis, pivot, comparing data, creating reports, charts, formulas, formatting.
---

# Excel / CSV 技能

現成腳本：本技能資料夾的 `scripts/xlsx_tool.py`（缺套件會自動安裝），**直接用，不要重寫**。
路徑通常是 `<skills 資料夾>/excel-xlsx/scripts/xlsx_tool.py`（OpenCode 是 `~/.config/opencode/skills`，Claude Code 是 `~/.claude/skills`，通用位置是 `~/.agents/skills`），以下用 `TOOL` 代表。

## 第一步：一定先看結構
`python "TOOL" info "檔案.xlsx"` → 每個工作表的列數、欄名、型別、前 5 列。
- 欄名很多是 `Unnamed` → 標題列不在第一列，讀取時用 `header=2` 或 `skiprows=` 調整。
- 要看完整內容：`python "TOOL" read "檔案.xlsx" --sheet "工作表1"`（輸出 .md 再用 read 工具讀）
- 要看公式：`python "TOOL" formulas "檔案.xlsx"`
- CSV 中文亂碼：腳本會自動試 utf-8 / cp950 / big5。

## 分析資料：一律用 pandas 寫程式算，不要用眼睛看表格心算
```python
import pandas as pd
df = pd.read_excel("檔案.xlsx", sheet_name="工作表1")
df.groupby("部門")["金額"].sum()                       # 分組加總
pd.pivot_table(df, index="客戶", columns="月份", values="金額", aggfunc="sum", fill_value=0)  # 樞紐
pd.merge(a, b, on="料號", how="outer", indicator=True)  # 比對兩份資料，_merge 欄看差異
```
- 日期欄：`pd.to_datetime(df["日期"], errors="coerce")`；民國年要自己 +1911 轉換。
- 數字欄有逗號或文字：`pd.to_numeric(df["金額"].astype(str).str.replace(",", ""), errors="coerce")`

## 產生新報表（要好看就用 openpyxl 加格式）
```python
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
with pd.ExcelWriter("報表.xlsx", engine="openpyxl") as w:
    df.to_excel(w, sheet_name="明細", index=False)
    ws = w.sheets["明細"]
    for c in ws[1]:                                     # 標題列
        c.font = Font(bold=True, name="微軟正黑體"); c.fill = PatternFill("solid", fgColor="D9E2F3")
    for col in ws.columns:                              # 自動欄寬（中文算 2 寬）
        width = max(sum(2 if ord(ch) > 127 else 1 for ch in str(c.value or "")) for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(width + 2, 50)
    ws.freeze_panes = "A2"; ws.auto_filter.ref = ws.dimensions
```
- 合計列盡量用 **Excel 公式**（例如 `ws["D20"] = "=SUM(D2:D19)"`），使用者改數字會自動更新。
- 金額格式：`cell.number_format = '#,##0'`；百分比：`'0.0%'`
- 注意：用 Python 寫入的公式，在 Excel 打開存檔前**沒有計算結果**（pandas 讀回會是空值）。驗證數字時用 pandas 另外算一次來核對。

## 修改現有檔案（保留原本格式）
- 用 `openpyxl.load_workbook("原檔.xlsx")` 改儲存格，再 `wb.save("原檔_修改.xlsx")`。
- **不要**用 pandas 讀再寫回原檔，會把格式、公式、其他工作表全部洗掉。
- **不要**用 `data_only=True` 開檔後存檔，公式會全部變成固定值。
- .xlsm（有巨集）：`load_workbook(..., keep_vba=True)`，存檔也要 .xlsm。
- .xls 舊格式只能讀（需 xlrd），要修改就另存 .xlsx。

## 圖表
- 要放在 Excel 裡：`openpyxl.chart`（BarChart、LineChart、PieChart）。
- 要做成圖片放進 Word/PPT：matplotlib，**中文字型要設定**否則變方框：
```python
import matplotlib; matplotlib.use("Agg")
import matplotlib.pyplot as plt
plt.rcParams["font.sans-serif"] = ["Microsoft JhengHei"]; plt.rcParams["axes.unicode_minus"] = False
```

## 轉 PDF
`python "<skills 資料夾>/excel-xlsx/scripts/office_convert.py" topdf "報表.xlsx"`
（會自己找 32/64 位元 Excel，沒有 Excel 改用 LibreOffice。不要自己猜 Excel 的安裝路徑。）

## 完成後驗證
重新讀回輸出檔，確認列數、加總金額和原始資料對得上，並在回報中列出關鍵數字。
