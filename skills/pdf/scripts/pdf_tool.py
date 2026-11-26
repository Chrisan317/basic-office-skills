"""PDF 工具：讀取、轉圖片、OCR、合併、分割、旋轉。

用法：
  python pdf_tool.py info    檔案.pdf
  python pdf_tool.py extract 檔案.pdf [-o 輸出.md] [--pages 1-3,5]
  python pdf_tool.py render  檔案.pdf [-o 資料夾] [--pages 1-3] [--dpi 150]
  python pdf_tool.py ocr     檔案.pdf [-o 輸出.md] [--pages 1-3]
  python pdf_tool.py merge   輸出.pdf a.pdf b.pdf ...
  python pdf_tool.py split   檔案.pdf --pages 1-3 -o 輸出.pdf
  python pdf_tool.py rotate  檔案.pdf --angle 90 [--pages 2] -o 輸出.pdf
頁碼一律從 1 開始。
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


def ensure(pkg, pip_name=None):
    try:
        return __import__(pkg)
    except ImportError:
        print(f"[安裝套件] {pip_name or pkg} ...", file=sys.stderr)
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", pip_name or pkg])
        return __import__(pkg)


pymupdf = ensure("pymupdf", "pymupdf")


def parse_pages(spec, total):
    """'1-3,5' -> [0,1,2,4]（0 起算）。None = 全部。"""
    if not spec:
        return list(range(total))
    result = []
    for part in spec.split(","):
        part = part.strip()
        if "-" in part:
            a, b = part.split("-", 1)
            a = int(a) if a else 1
            b = int(b) if b else total
            result.extend(range(a - 1, min(b, total)))
        elif part:
            n = int(part)
            if 1 <= n <= total:
                result.append(n - 1)
    return result


def default_out(path, suffix):
    base, _ = os.path.splitext(path)
    return base + suffix


def cmd_info(a):
    doc = pymupdf.open(a.file)
    print(f"檔案: {a.file}")
    print(f"頁數: {doc.page_count}")
    print(f"加密: {doc.is_encrypted}")
    meta = {k: v for k, v in (doc.metadata or {}).items() if v}
    for k, v in meta.items():
        print(f"{k}: {v}")
    # 判斷是否為掃描檔（幾乎沒有文字）
    sample = parse_pages(None, doc.page_count)[:20]
    no_text = [i + 1 for i in sample if len(doc[i].get_text().strip()) < 20]
    if len(no_text) == len(sample):
        print("判斷: 幾乎沒有文字層，可能是掃描檔/圖片 PDF → 用 render 轉圖片看，或用 ocr")
    elif no_text:
        print(f"判斷: 大部分有文字層，可用 extract；但這些頁幾乎沒文字（可能是圖片）: {no_text}")
    else:
        print("判斷: 有文字層，可直接用 extract")
    toc = doc.get_toc()
    if toc:
        print("目錄:")
        for lvl, title, page in toc[:50]:
            print(f"{'  ' * (lvl - 1)}- {title} (p.{page})")


def cmd_extract(a):
    doc = pymupdf.open(a.file)
    pages = parse_pages(a.pages, doc.page_count)
    out = a.output or default_out(a.file, "_內容.md")
    lines = [f"# {os.path.basename(a.file)}", ""]
    empty_pages = []
    for i in pages:
        page = doc[i]
        lines.append(f"\n## 第 {i + 1} 頁\n")
        text = page.get_text("text", sort=True).strip()
        if not text:
            empty_pages.append(i + 1)
            lines.append("（此頁沒有文字層，可能是圖片，需用 render 或 ocr）")
        else:
            lines.append(text)
        if not a.no_tables:
            try:
                tables = page.find_tables()
                for t_idx, t in enumerate(tables.tables, 1):
                    md = t.to_markdown()
                    if md.strip():
                        lines.append(f"\n### 表格 {i + 1}-{t_idx}\n")
                        lines.append(md)
            except Exception as e:  # 表格偵測失敗不影響文字
                lines.append(f"\n（表格偵測失敗: {e}）")
    with open(out, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print(f"已輸出: {out}")
    print(f"共 {len(pages)} 頁，約 {sum(len(l) for l in lines)} 字")
    if empty_pages:
        print(f"沒有文字的頁: {empty_pages} → 用 render 轉成圖片查看，或用 ocr")


def cmd_render(a):
    doc = pymupdf.open(a.file)
    pages = parse_pages(a.pages, doc.page_count)
    outdir = a.output or default_out(a.file, "_頁面圖")
    os.makedirs(outdir, exist_ok=True)
    for i in pages:
        pix = doc[i].get_pixmap(dpi=a.dpi)
        p = os.path.join(outdir, f"page_{i + 1:03d}.png")
        pix.save(p)
        print(p)
    print(f"已輸出 {len(pages)} 張圖到: {outdir}")


def cmd_ocr(a):
    try:
        from rapidocr import RapidOCR  # 新版套件
    except ImportError:
        try:
            from rapidocr_onnxruntime import RapidOCR  # 舊版套件
        except ImportError:
            print("[安裝 OCR 套件] rapidocr onnxruntime ...", file=sys.stderr)
            subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "rapidocr", "onnxruntime"])
            from rapidocr import RapidOCR
    engine = RapidOCR()
    doc = pymupdf.open(a.file)
    pages = parse_pages(a.pages, doc.page_count)
    out = a.output or default_out(a.file, "_OCR.md")
    lines = [f"# {os.path.basename(a.file)}（OCR）", ""]
    for i in pages:
        png = doc[i].get_pixmap(dpi=200).tobytes("png")
        res = engine(png)
        # 新舊版回傳格式不同
        if hasattr(res, "txts"):
            texts = list(res.txts or [])
        else:
            texts = [r[1] for r in (res[0] or [])]
        lines.append(f"\n## 第 {i + 1} 頁\n")
        lines.append("\n".join(texts) if texts else "（辨識不到文字）")
        print(f"第 {i + 1} 頁: {len(texts)} 行", file=sys.stderr)
    with open(out, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print(f"已輸出: {out}")


def cmd_merge(a):
    out = pymupdf.open()
    for f in a.inputs:
        with pymupdf.open(f) as d:
            out.insert_pdf(d)
    out.save(a.output_file)
    print(f"已合併 {len(a.inputs)} 個檔案 → {a.output_file}（{out.page_count} 頁）")


def cmd_split(a):
    doc = pymupdf.open(a.file)
    pages = parse_pages(a.pages, doc.page_count)
    out = pymupdf.open()
    for i in pages:
        out.insert_pdf(doc, from_page=i, to_page=i)
    target = a.output or default_out(a.file, f"_p{a.pages}.pdf")
    out.save(target)
    print(f"已輸出 {len(pages)} 頁 → {target}")


def cmd_rotate(a):
    doc = pymupdf.open(a.file)
    for i in parse_pages(a.pages, doc.page_count):
        doc[i].set_rotation((doc[i].rotation + a.angle) % 360)
    target = a.output or default_out(a.file, "_旋轉.pdf")
    doc.save(target)
    print(f"已輸出 → {target}")


def main():
    p = argparse.ArgumentParser(description="PDF 工具")
    sub = p.add_subparsers(dest="cmd", required=True)

    s = sub.add_parser("info"); s.add_argument("file"); s.set_defaults(fn=cmd_info)

    s = sub.add_parser("extract"); s.add_argument("file")
    s.add_argument("-o", "--output"); s.add_argument("--pages")
    s.add_argument("--no-tables", action="store_true"); s.set_defaults(fn=cmd_extract)

    s = sub.add_parser("render"); s.add_argument("file")
    s.add_argument("-o", "--output"); s.add_argument("--pages")
    s.add_argument("--dpi", type=int, default=150); s.set_defaults(fn=cmd_render)

    s = sub.add_parser("ocr"); s.add_argument("file")
    s.add_argument("-o", "--output"); s.add_argument("--pages"); s.set_defaults(fn=cmd_ocr)

    s = sub.add_parser("merge"); s.add_argument("output_file"); s.add_argument("inputs", nargs="+")
    s.set_defaults(fn=cmd_merge)

    s = sub.add_parser("split"); s.add_argument("file"); s.add_argument("--pages", required=True)
    s.add_argument("-o", "--output"); s.set_defaults(fn=cmd_split)

    s = sub.add_parser("rotate"); s.add_argument("file"); s.add_argument("--angle", type=int, default=90)
    s.add_argument("--pages"); s.add_argument("-o", "--output"); s.set_defaults(fn=cmd_rotate)

    a = p.parse_args()
    a.fn(a)


if __name__ == "__main__":
    main()
