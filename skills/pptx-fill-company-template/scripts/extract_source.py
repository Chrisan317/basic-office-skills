"""Normalize a NON-pptx source (.md .txt .csv .xlsx .docx) into one structure so a build script can map it onto template slides.

usage:  python extract_source.py report.md            -> prints JSON
        from extract_source import extract; doc = extract("report.xlsx")

doc = {"file": str, "title": str, "sections": [ {"heading": str, "level": int, "blocks": [
          {"type": "paragraph", "text": str},
          {"type": "bullets",   "items": [str, ...]},
          {"type": "table",     "rows": [[str, ...], ...]}   # rows[0] is the header row
       ]} ]}
Content before the first heading goes in a section with heading "". Nothing is invented or summarized here;
deciding which section becomes which slide is the build script's job. PDF: use the pdf skill to get text, save it as .txt/.md, then run this."""
import sys, os, re, csv, json

def _blocks_from_lines(lines):
    """Markdown-ish text -> list of sections."""
    secs = [{"heading": "", "level": 0, "blocks": []}]
    para, bul, tab = [], [], []
    def flush():
        nonlocal para, bul, tab
        b = secs[-1]["blocks"]
        if para: b.append({"type": "paragraph", "text": " ".join(para)}); para = []
        if bul: b.append({"type": "bullets", "items": bul}); bul = []
        if tab: b.append({"type": "table", "rows": tab}); tab = []
    for raw in lines:
        ln = raw.rstrip()
        m = re.match(r"^(#{1,6})\s+(.*)$", ln)
        if m:
            flush(); secs.append({"heading": m.group(2).strip(), "level": len(m.group(1)), "blocks": []}); continue
        if ln.lstrip().startswith("|"):
            cells = [c.strip() for c in ln.strip().strip("|").split("|")]
            if all(re.fullmatch(r":?-{2,}:?", c) for c in cells if c): continue   # |---|---| separator
            if para or bul: flush()
            tab.append(cells); continue
        m = re.match(r"^\s*(?:[-*+•]|\d+[.)])\s+(.*)$", ln)
        if m:
            if para or tab: flush()
            bul.append(m.group(1).strip()); continue
        if not ln.strip(): flush(); continue
        if bul or tab: flush()
        para.append(ln.strip())
    flush()
    return [s for s in secs if s["heading"] or s["blocks"]]

def _from_text(path):
    return _blocks_from_lines(open(path, encoding="utf-8-sig").read().splitlines())

def _from_csv(path):
    with open(path, encoding="utf-8-sig", newline="") as f:
        rows = [[c.strip() for c in r] for r in csv.reader(f) if any(c.strip() for c in r)]
    return [{"heading": os.path.splitext(os.path.basename(path))[0], "level": 1, "blocks": [{"type": "table", "rows": rows}]}]

def _from_xlsx(path):
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True, read_only=True)
    out = []
    for ws in wb.worksheets:
        rows = []
        for r in ws.iter_rows(values_only=True):
            cells = ["" if v is None else (f"{v:g}" if isinstance(v, float) else str(v)).strip() for v in r]
            if any(cells): rows.append(cells)
        if not rows: continue
        w = max(len(r) for r in rows)
        while w and all(not r[w - 1] if len(r) >= w else True for r in rows): w -= 1   # trim empty right columns
        rows = [(r + [""] * w)[:w] for r in rows]
        out.append({"heading": ws.title, "level": 1, "blocks": [{"type": "table", "rows": rows}]})
    return out

def _from_docx(path):
    import docx
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    d = docx.Document(path)
    secs = [{"heading": "", "level": 0, "blocks": []}]
    bul = []
    def flush():
        nonlocal bul
        if bul: secs[-1]["blocks"].append({"type": "bullets", "items": bul}); bul = []
    for el in d.element.body.iterchildren():
        tag = el.tag.split("}")[1]
        if tag == "p":
            p = Paragraph(el, d); t = p.text.strip(); st = (p.style.name or "") if p.style is not None else ""
            if not t: flush(); continue
            m = re.match(r"(?:Heading|標題)\s*(\d)", st) or (re.match(r"Title", st) and re.match(r"(1)", "1"))
            if m:
                flush(); secs.append({"heading": t, "level": int(m.group(1)), "blocks": []})
            elif "List" in st or "清單" in st:
                bul.append(t)
            else:
                flush(); secs[-1]["blocks"].append({"type": "paragraph", "text": t})
        elif tag == "tbl":
            flush(); tb = Table(el, d)
            rows = [[c.text.strip().replace("\n", " ") for c in r.cells] for r in tb.rows]
            secs[-1]["blocks"].append({"type": "table", "rows": rows})
    flush()
    return [s for s in secs if s["heading"] or s["blocks"]]

def extract(path):
    ext = os.path.splitext(path)[1].lower()
    fn = {".md": _from_text, ".markdown": _from_text, ".txt": _from_text, ".csv": _from_csv,
          ".xlsx": _from_xlsx, ".xlsm": _from_xlsx, ".docx": _from_docx}.get(ext)
    if fn is None: raise SystemExit(f"unsupported source type {ext}: convert to .md/.txt/.csv/.xlsx/.docx first (PDF: use the pdf skill to get text)")
    secs = fn(path)
    title = next((s["heading"] for s in secs if s["heading"] and s["level"] <= 1), "") or os.path.splitext(os.path.basename(path))[0]
    return {"file": os.path.basename(path), "title": title, "sections": secs}

if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    print(json.dumps(extract(sys.argv[1]), ensure_ascii=False, indent=1))
