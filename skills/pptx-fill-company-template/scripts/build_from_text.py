"""把非 pptx 的來源（md/txt/csv/xlsx/docx）套到任何公司範本上。這是起點，不是成品：
真正交付時複製一份，依內容調整每一節要怎麼呈現。

用法：python build_from_text.py SOURCE TEMPLATE.pptx OUT.pptx
規則：封面（文件標題）→ 每一節一頁（條列/段落做成卡片，表格自動分頁）→ 結尾頁（範本有的話）
版面自動偵測（cover / content / closing），顏色取自範本的佈景主題，尺寸依投影片大小計算（16:9、4:3 都可）。
"""
import math
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import helpers as h  # noqa: E402
from helpers import (C, box, est_h, fit, new_slide, save, set_subtitle, drop_empty_placeholders,  # noqa: E402
                     content_area, style_header_row, table, text_w, txt, zebra_body, init)
from extract_source import extract  # noqa: E402

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

SRC, TPL, OUT = sys.argv[1:4]
doc = extract(SRC)
init(None, TPL)

# ---- 封面：文件標題 + 標題區第一段短文字當副標 ----
s = new_slide("cover")
if s.shapes.title is not None:
    s.shapes.title.text = doc["title"]
first = next((b["text"] for sec in doc["sections"] for b in sec["blocks"]
              if b["type"] == "paragraph" and len(b["text"]) <= 40), "")
if first:
    set_subtitle(s, first)
drop_empty_placeholders(s)


def is_num(t):
    return bool(re.fullmatch(r"[-+]?[\d,.]+%?", t.strip()))


def table_slides(title, rows):
    ncol = max(len(r) for r in rows)
    rows = [(r + [""] * ncol)[:ncol] for r in rows]
    size = 12 if len(rows) <= 10 else 11 if len(rows) <= 16 else 10
    probe = new_slide("content", title)            # 先開一頁量出內容區
    x, y, cw, ch = content_area(probe)
    w = [max(0.7, min(max(text_w(r[c], size) for r in rows) / 72 + 0.3, 5.0)) for c in range(ncol)]
    w = fit(w, cw)

    def row_h(r):
        lines = max(max(1, math.ceil(text_w(v, size) / ((w[c] - 0.12) * 72))) for c, v in enumerate(r))
        return max(0.34, lines * size * 1.25 / 72 + 0.1)

    head_h = row_h(rows[0]) + 0.05
    pages, cur, used = [], [], head_h
    for r in rows[1:]:
        h_ = row_h(r)
        if cur and used + h_ > ch:
            pages.append(cur)
            cur, used = [], head_h
        cur.append((r, h_))
        used += h_
    pages.append(cur)
    if len(pages) > 1:                             # 平均分配，避免最後一頁只剩一兩列
        allrows = [x_ for pg in pages for x_ in pg]
        per = math.ceil(len(allrows) / len(pages))
        pages = [allrows[k:k + per] for k in range(0, len(allrows), per)]
    for i, pg in enumerate(pages, 1):
        t_title = title + (f" ({i}/{len(pages)})" if len(pages) > 1 else "")
        s = probe if i == 1 else new_slide("content", t_title)
        if i == 1 and s.shapes.title is not None:
            s.shapes.title.text = t_title
        t = table(s, x, y, w, len(pg) + 1, [head_h] + [h_ for _, h_ in pg], name="資料表")
        style_header_row(t, 0, rows[0], size)
        zebra_body(t, 1, [r for r, _ in pg], size - 1 if size > 10 else size,
                   {c: C for c in range(ncol) if all(is_num(r[c]) for r, _ in pg if r[c])})


def text_slide(title, blocks):
    items = []
    for b in blocks:
        if b["type"] == "paragraph":
            items.append((b["text"], False))
        elif b["type"] == "bullets":
            items += [(t, True) for t in b["items"]]
    if not items:
        return
    s = new_slide("content", title)
    x0, y0, aw, ah = content_area(s)
    two = len(items) > 7
    cols = [items[: (len(items) + 1) // 2], items[(len(items) + 1) // 2:]] if two else [items]
    gap = 0.23
    cw = aw if not two else (aw - gap) / 2
    for i, col in enumerate(cols):
        x = x0 + i * (cw + gap)
        paras = [dict(text=("• " if bul else "") + t, size=16 if not two else 14, color=h.BLACK, after=10)
                 for t, bul in col]
        need = est_h([(p["text"], p["size"], p["after"]) for p in paras], cw - 0.7) + 0.4
        ch = min(ah, max(1.4, need + 0.3))         # 卡片高度跟著內容
        box(s, x, y0, cw, ch, h.CLOUD)
        txt(s, x + 0.3, y0 + 0.15, cw - 0.6, ch - 0.3, paras, name=f"內容{i + 1}")


for sec in doc["sections"]:
    head = sec["heading"] or doc["title"]
    if sec["level"] == 1 and not any(b["type"] == "table" for b in sec["blocks"]) and \
       all(b["type"] == "paragraph" and len(b["text"]) <= 40 for b in sec["blocks"]):
        continue                                   # 標題區，已放在封面
    text_slide(head, [b for b in sec["blocks"] if b["type"] != "table"])
    for b in sec["blocks"]:
        if b["type"] == "table" and len(b["rows"]) > 1:
            table_slides(head, b["rows"])

end = new_slide("closing")                         # 範本沒有結尾版面就略過
if end is not None:
    drop_empty_placeholders(end)
save(OUT)
