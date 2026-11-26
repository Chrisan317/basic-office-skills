"""Fallback when nothing can render the deck: flag shapes/tables that leave the slide.
usage: python -I check_bounds.py deck.pptx
Table height is estimated from row heights AND wrapped text (CJK=1em, Latin~0.55em), because
PowerPoint grows rows to fit text. Empty cells with no endParaRPr size count as 18pt (the real-world cause of overflow)."""
import sys, math
from pptx import Presentation
E = 914400
def tw(s, sz): return sum(sz if ord(c) > 0x2E80 else sz * .55 for c in s)
p = Presentation(sys.argv[1]); W, H = p.slide_width / E, p.slide_height / E
bad = 0
for n, s in enumerate(p.slides, 1):
    for sh in s.shapes:
        if sh.left is None: continue
        x, y, w, h = sh.left / E, sh.top / E, sh.width / E, sh.height / E
        if sh.has_table:
            t = sh.table; tot = 0
            for r in t.rows:
                need = r.height / E
                for ci, c in enumerate(r.cells):
                    cw = t.columns[ci].width / E - (c.margin_left + c.margin_right) / E
                    hh = (c.margin_top + c.margin_bottom) / E
                    for para in c.text_frame.paragraphs:
                        szs = [ru.font.size.pt for ru in para.runs if ru.font.size]
                        sz = szs[0] if szs else 18
                        txt = para.text
                        if not para.runs:
                            e = para._p.find('{http://schemas.openxmlformats.org/drawingml/2006/main}endParaRPr')
                            sz = int(e.get('sz')) / 100 if e is not None and e.get('sz') else 18
                        lines = max(1, math.ceil(tw(txt, sz) / max(cw * 72, 1)))
                        hh += lines * sz * 1.2 / 72
                    need = max(need, hh)
                tot += need
            h = tot
        if x < -0.01 or y < -0.01 or x + w > W + 0.01 or y + h > H + 0.01:
            bad += 1; print(f"slide {n}: '{sh.name}' ends at x={x+w:.2f} y={y+h:.2f} (slide {W:.2f}x{H:.2f})")
print("OK: nothing outside the slide" if not bad else f"{bad} problem(s)")
