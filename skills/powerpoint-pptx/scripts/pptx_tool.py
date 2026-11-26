"""PowerPoint 工具：讀取內容、列出版面配置、轉 PDF。

用法：
  python pptx_tool.py read    簡報.pptx [-o 輸出.md]      （每頁文字、表格、備忘稿）
  python pptx_tool.py layouts 簡報或範本.pptx              （列出可用的版面配置與預留位置）
  python pptx_tool.py topdf   簡報.pptx [-o 輸出.pdf]      （需要 PowerPoint 或 LibreOffice）
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


ensure("pptx", "python-pptx")
from pptx import Presentation  # noqa: E402
from pptx.util import Emu  # noqa: E402


def shape_lines(shape, depth=0):
    out = []
    if shape.shape_type == 6 and hasattr(shape, "shapes"):  # 群組
        for s in shape.shapes:
            out += shape_lines(s, depth)
        return out
    if getattr(shape, "has_table", False) and shape.has_table:
        rows = [[c.text.replace("\n", " ").replace("|", "\\|") for c in r.cells] for r in shape.table.rows]
        if rows:
            out.append("| " + " | ".join(rows[0]) + " |")
            out.append("|" + "---|" * len(rows[0]))
            out += ["| " + " | ".join(r) + " |" for r in rows[1:]]
        return out
    if getattr(shape, "has_chart", False) and shape.has_chart:
        out.append(f"[圖表: {shape.chart.chart_type}]")
        return out
    if shape.shape_type == 13:
        out.append(f"[圖片: {shape.name}]")
        return out
    if getattr(shape, "has_text_frame", False) and shape.has_text_frame:
        for p in shape.text_frame.paragraphs:
            t = "".join(r.text for r in p.runs).strip()
            if t:
                out.append("  " * p.level + "- " + t)
    return out


def cmd_read(a):
    prs = Presentation(a.file)
    out = a.output or os.path.splitext(a.file)[0] + "_內容.md"
    lines = [f"# {os.path.basename(a.file)}（{len(prs.slides)} 頁）", ""]
    for i, slide in enumerate(prs.slides, 1):
        title = slide.shapes.title.text.strip() if slide.shapes.title is not None and slide.shapes.title.has_text_frame else ""
        lines.append(f"\n## 第 {i} 頁：{title}  〔版面: {slide.slide_layout.name}〕\n")
        for shape in slide.shapes:
            if slide.shapes.title is not None and shape.shape_id == slide.shapes.title.shape_id:
                continue
            lines += shape_lines(shape)
        if slide.has_notes_slide:
            notes = slide.notes_slide.notes_text_frame.text.strip()
            if notes:
                lines.append(f"\n> 備忘稿：{notes}")
    with open(out, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print(f"已輸出: {out}")


def cmd_layouts(a):
    prs = Presentation(a.file)
    print(f"投影片尺寸: {Emu(prs.slide_width).cm:.2f} x {Emu(prs.slide_height).cm:.2f} cm")
    for i, layout in enumerate(prs.slide_layouts):
        phs = ", ".join(f"idx{ph.placeholder_format.idx}={ph.placeholder_format.type}" for ph in layout.placeholders)
        print(f"[{i}] {layout.name}: {phs}")


def cmd_topdf(a):
    dst = a.output or os.path.splitext(a.file)[0] + ".pdf"
    errors = []
    if sys.platform == "win32":
        try:
            ensure("win32com", "pywin32")
            import win32com.client
            app = win32com.client.DispatchEx("PowerPoint.Application")
            try:
                pres = app.Presentations.Open(os.path.abspath(a.file), ReadOnly=True, WithWindow=False)
                pres.SaveAs(os.path.abspath(dst), 32)  # 32 = PDF
                pres.Close()
            finally:
                app.Quit()
            print(f"已用 PowerPoint 轉成: {dst}")
            return
        except Exception as e:
            errors.append(f"PowerPoint: {e}")
    soffice = None
    for c in ("soffice", r"C:\Program Files\LibreOffice\program\soffice.exe"):
        if shutil.which(c) or os.path.exists(c):
            soffice = shutil.which(c) or c
            break
    if soffice:
        outdir = os.path.dirname(os.path.abspath(dst))
        subprocess.check_call([soffice, "--headless", "--convert-to", "pdf", "--outdir", outdir, a.file])
        produced = os.path.join(outdir, os.path.splitext(os.path.basename(a.file))[0] + ".pdf")
        if os.path.abspath(produced) != os.path.abspath(dst):
            shutil.move(produced, dst)
        print(f"已用 LibreOffice 轉成: {dst}")
        return
    errors.append("LibreOffice: 找不到")
    sys.exit("轉檔失敗：\n" + "\n".join(errors))


def main():
    p = argparse.ArgumentParser(description="PowerPoint 工具")
    sub = p.add_subparsers(dest="cmd", required=True)
    s = sub.add_parser("read"); s.add_argument("file"); s.add_argument("-o", "--output"); s.set_defaults(fn=cmd_read)
    s = sub.add_parser("layouts"); s.add_argument("file"); s.set_defaults(fn=cmd_layouts)
    s = sub.add_parser("topdf"); s.add_argument("file"); s.add_argument("-o", "--output"); s.set_defaults(fn=cmd_topdf)
    a = p.parse_args()
    a.fn(a)


if __name__ == "__main__":
    main()
