"""Excel / CSV 工具：快速了解檔案結構、轉 Markdown、轉 CSV。

用法：
  python xlsx_tool.py info  檔案.xlsx                 （每個工作表的大小、欄名、前 5 列）
  python xlsx_tool.py read  檔案.xlsx [--sheet 名稱] [--rows 200] [-o 輸出.md]
  python xlsx_tool.py formulas 檔案.xlsx [--sheet 名稱]  （列出公式）
  python xlsx_tool.py tocsv 檔案.xlsx [--sheet 名稱] [-o 輸出.csv]
支援 .xlsx .xlsm .xls .csv
"""
import argparse
import os
import subprocess
import sys

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass


def ensure(mod, pip_name):
    try:
        return __import__(mod)
    except ImportError:
        print(f"[安裝套件] {pip_name} ...", file=sys.stderr)
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", pip_name])
        return __import__(mod)


pd = ensure("pandas", "pandas")
ensure("openpyxl", "openpyxl")
ensure("tabulate", "tabulate")


def read_csv_any(path):
    """台灣常見 CSV 編碼都試一次。"""
    for enc in ("utf-8-sig", "cp950", "big5", "utf-16", "latin1"):
        try:
            return pd.read_csv(path, encoding=enc), enc
        except (UnicodeDecodeError, UnicodeError):
            continue
    raise RuntimeError("無法判斷 CSV 編碼")


def load_all(path, sheet=None, header=0):
    ext = os.path.splitext(path)[1].lower()
    if ext in (".csv", ".txt"):
        df, enc = read_csv_any(path)
        print(f"(CSV 編碼: {enc})", file=sys.stderr)
        return {"CSV": df}
    if ext == ".xls":
        ensure("xlrd", "xlrd")
    return pd.read_excel(path, sheet_name=sheet or None, header=header)


def as_dict(res, sheet):
    return res if isinstance(res, dict) else {sheet: res}


def cmd_info(a):
    sheets = as_dict(load_all(a.file, a.sheet), a.sheet)
    print(f"檔案: {a.file}，工作表數: {len(sheets)}")
    for name, df in sheets.items():
        print(f"\n=== 工作表「{name}」 {df.shape[0]} 列 x {df.shape[1]} 欄 ===")
        print("欄位: " + ", ".join(f"{c}({t})" for c, t in zip(df.columns, df.dtypes.astype(str))))
        unnamed = sum(str(c).startswith("Unnamed") for c in df.columns)
        if unnamed > len(df.columns) / 2:
            print("注意: 很多欄名是 Unnamed，標題列可能不在第 1 列 → 用 pandas 的 header= 或 skiprows= 調整")
        print(df.head(5).to_markdown(index=False))


def cmd_read(a):
    sheets = as_dict(load_all(a.file, a.sheet), a.sheet)
    out = a.output or os.path.splitext(a.file)[0] + "_內容.md"
    lines = [f"# {os.path.basename(a.file)}", ""]
    for name, df in sheets.items():
        lines.append(f"\n## 工作表：{name}（{df.shape[0]} 列 x {df.shape[1]} 欄）\n")
        part = df.head(a.rows)
        lines.append(part.to_markdown(index=False))
        if len(df) > a.rows:
            lines.append(f"\n（只顯示前 {a.rows} 列，共 {len(df)} 列。分析全部資料請用 pandas 寫程式計算）")
    with open(out, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print(f"已輸出: {out}")


def cmd_formulas(a):
    import openpyxl
    wb = openpyxl.load_workbook(a.file, data_only=False)
    names = [a.sheet] if a.sheet else wb.sheetnames
    n = 0
    for name in names:
        ws = wb[name]
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    print(f"{name}!{c.coordinate}: {c.value}")
                    n += 1
                    if n >= 500:
                        print("（超過 500 個公式，只列前 500）")
                        return
    print(f"共 {n} 個公式")


def cmd_tocsv(a):
    sheets = as_dict(load_all(a.file, a.sheet), a.sheet)
    base = os.path.splitext(a.file)[0]
    for name, df in sheets.items():
        out = a.output if (a.output and len(sheets) == 1) else f"{base}_{name}.csv"
        df.to_csv(out, index=False, encoding="utf-8-sig")  # utf-8-sig 讓 Excel 開中文不亂碼
        print(f"已輸出: {out}")


def main():
    p = argparse.ArgumentParser(description="Excel 工具")
    sub = p.add_subparsers(dest="cmd", required=True)
    s = sub.add_parser("info"); s.add_argument("file"); s.add_argument("--sheet"); s.set_defaults(fn=cmd_info)
    s = sub.add_parser("read"); s.add_argument("file"); s.add_argument("--sheet")
    s.add_argument("--rows", type=int, default=200); s.add_argument("-o", "--output"); s.set_defaults(fn=cmd_read)
    s = sub.add_parser("formulas"); s.add_argument("file"); s.add_argument("--sheet"); s.set_defaults(fn=cmd_formulas)
    s = sub.add_parser("tocsv"); s.add_argument("file"); s.add_argument("--sheet")
    s.add_argument("-o", "--output"); s.set_defaults(fn=cmd_tocsv)
    a = p.parse_args()
    a.fn(a)


if __name__ == "__main__":
    main()
