"""Word 工具：讀取、取代文字、轉 PDF。

用法：
  python docx_tool.py read    檔案.docx [-o 輸出.md]
  python docx_tool.py replace 檔案.docx -o 輸出.docx --pair "舊文字=>新文字" [--pair ...]
  python docx_tool.py topdf   檔案.docx [-o 輸出.pdf]
  python docx_tool.py fromdoc 舊檔.doc        （.doc 轉 .docx，需要 Word 或 LibreOffice）
"""
import argparse
import os
import shutil
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


docx = ensure("docx", "python-docx")
from docx.table import Table  # noqa: E402
from docx.text.paragraph import Paragraph  # noqa: E402


def iter_block_items(parent):
    """依文件順序走訪段落與表格。"""
    body = parent.element.body if hasattr(parent, "element") and hasattr(parent.element, "body") else parent._element
    for child in body.iterchildren():
        tag = child.tag.split("}")[-1]
        if tag == "p":
            yield Paragraph(child, parent)
        elif tag == "tbl":
            yield Table(child, parent)


def table_to_md(table):
    rows = []
    for r in table.rows:
        cells = []
        for c in r.cells:
            cells.append(c.text.replace("\n", "<br>").replace("|", "\\|").strip())
        rows.append(cells)
    if not rows:
        return ""
    width = max(len(r) for r in rows)
    rows = [r + [""] * (width - len(r)) for r in rows]
    out = ["| " + " | ".join(rows[0]) + " |", "|" + "---|" * width]
    out += ["| " + " | ".join(r) + " |" for r in rows[1:]]
    return "\n".join(out)


def para_to_md(p):
    text = p.text.strip()
    if not text:
        return ""
    style = (p.style.name or "").lower() if p.style is not None else ""
    if style.startswith("heading") or style.startswith("標題"):
        digits = "".join(ch for ch in style if ch.isdigit())
        level = int(digits) if digits else 1
        return "#" * min(level + 1, 6) + " " + text
    if style == "title":
        return "# " + text
    if "list" in style or "清單" in style:
        return "- " + text
    return text


def cmd_read(a):
    d = docx.Document(a.file)
    out = a.output or os.path.splitext(a.file)[0] + "_內容.md"
    lines = [f"<!-- 來源: {os.path.basename(a.file)} -->", ""]
    n_tables = 0
    for block in iter_block_items(d):
        if isinstance(block, Paragraph):
            md = para_to_md(block)
            if md:
                lines.append(md)
                lines.append("")
        else:
            n_tables += 1
            lines.append(f"**[表格 {n_tables}]**\n")
            lines.append(table_to_md(block))
            lines.append("")
    # 頁首頁尾
    hf = []
    for s in d.sections:
        for part, name in ((s.header, "頁首"), (s.footer, "頁尾")):
            t = "\n".join(p.text for p in part.paragraphs if p.text.strip())
            if t and t not in hf:
                hf.append(t)
                lines.append(f"<!-- {name}: {t} -->")
    with open(out, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print(f"已輸出: {out}")
    print(f"段落 {len(d.paragraphs)}、表格 {n_tables}、圖片 {len(d.inline_shapes)}")


def replace_in_paragraph(p, old, new):
    if old not in p.text:
        return 0
    count = p.text.count(old)
    # 情況 1：文字完整落在單一 run 內 → 保留格式
    done = False
    for r in p.runs:
        if old in r.text:
            r.text = r.text.replace(old, new)
            done = True
    if done and old not in p.text:
        return count
    # 情況 2：文字被拆在多個 run → 合併到第一個 run（保留第一個 run 的格式）
    full = p.text.replace(old, new)
    runs = p.runs
    if runs:
        runs[0].text = full
        for r in runs[1:]:
            r.text = ""
    return count


def all_paragraphs(d):
    def from_container(c):
        for p in c.paragraphs:
            yield p
        for t in c.tables:
            for row in t.rows:
                for cell in row.cells:
                    yield from from_container(cell)
    yield from from_container(d)
    for s in d.sections:
        for part in (s.header, s.footer, s.first_page_header, s.first_page_footer):
            yield from from_container(part)


def cmd_replace(a):
    if os.path.abspath(a.file) == os.path.abspath(a.output) and not a.overwrite:
        sys.exit("輸出檔和原檔相同，為避免覆蓋請換檔名（或加 --overwrite）")
    d = docx.Document(a.file)
    pairs = []
    for pair in a.pair:
        if "=>" not in pair:
            sys.exit(f"格式錯誤: {pair}，應為 舊=>新")
        pairs.append(pair.split("=>", 1))
    seen = set()
    totals = {old: 0 for old, _ in pairs}
    for p in all_paragraphs(d):
        if id(p._p) in seen:  # 合併儲存格會重複
            continue
        seen.add(id(p._p))
        for old, new in pairs:
            totals[old] += replace_in_paragraph(p, old, new)
    d.save(a.output)
    for old, n in totals.items():
        print(f"「{old}」取代 {n} 處")
    print(f"已輸出: {a.output}")


sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import office_convert  # noqa: E402  找 Office（32/64 位元）與轉檔


def cmd_topdf(a):
    dst = a.output or os.path.splitext(a.file)[0] + ".pdf"
    how = office_convert.to_pdf(a.file, dst)
    print(f"已用 {how} 轉成: {os.path.abspath(dst)}")


def cmd_fromdoc(a):
    """.doc → .docx：pywin32 → PowerShell（64/32 位元）→ LibreOffice"""
    src = os.path.abspath(a.file)
    dst = os.path.abspath(a.output or os.path.splitext(a.file)[0] + ".docx")
    errors = []
    if sys.platform == "win32":
        try:
            ensure("win32com", "pywin32")
            import win32com.client
            w = win32com.client.DispatchEx("Word.Application")
            try:
                d = w.Documents.Open(src, ReadOnly=True)
                d.SaveAs2(dst, FileFormat=16)
                d.Close(False)
            finally:
                w.Quit()
            print(f"已用 Word 轉成: {dst}")
            return
        except Exception as e:
            errors.append(f"pywin32: {e}")
        try:
            q = office_convert._ps_quote
            office_convert._run_ps("$ErrorActionPreference='Stop';$a=New-Object -ComObject Word.Application;"
                                   f"$x=$a.Documents.Open({q(src)},$false,$true);$x.SaveAs2({q(dst)},16);"
                                   "$x.Close($false);$a.Quit()")
            print(f"已用 Word（PowerShell）轉成: {dst}")
            return
        except Exception as e:
            errors.append(f"PowerShell: {e}")
    so = office_convert.find_app("soffice.exe")
    if so:
        outdir = os.path.dirname(dst)
        subprocess.check_call([so, "--headless", "--convert-to", "docx", "--outdir", outdir, src])
        produced = os.path.join(outdir, os.path.splitext(os.path.basename(src))[0] + ".docx")
        if os.path.abspath(produced) != dst:
            shutil.move(produced, dst)
        print(f"已用 LibreOffice 轉成: {dst}")
        return
    sys.exit("轉檔失敗：\n" + "\n".join(errors) + "\n需要 Microsoft Word 或 LibreOffice")


def main():
    p = argparse.ArgumentParser(description="Word 工具")
    sub = p.add_subparsers(dest="cmd", required=True)
    s = sub.add_parser("read"); s.add_argument("file"); s.add_argument("-o", "--output"); s.set_defaults(fn=cmd_read)
    s = sub.add_parser("replace"); s.add_argument("file"); s.add_argument("-o", "--output", required=True)
    s.add_argument("--pair", action="append", required=True); s.add_argument("--overwrite", action="store_true")
    s.set_defaults(fn=cmd_replace)
    s = sub.add_parser("topdf"); s.add_argument("file"); s.add_argument("-o", "--output"); s.set_defaults(fn=cmd_topdf)
    s = sub.add_parser("fromdoc"); s.add_argument("file"); s.add_argument("-o", "--output"); s.set_defaults(fn=cmd_fromdoc)
    a = p.parse_args()
    try:
        a.fn(a)
    except Exception as e:
        sys.exit(str(e))


if __name__ == "__main__":
    main()
