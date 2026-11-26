"""找 Office / LibreOffice、用它們轉檔。處理 32 位元與 64 位元 Office 的差異。

不要猜 Office 的安裝路徑：
  64 位元 Office → C:\\Program Files\\Microsoft Office\\root\\Office16\\
  32 位元 Office → C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\
本模組會從登錄檔（App Paths，64/32 位元兩邊都查）和兩個 Program Files 資料夾找。

用法（命令列）：
  python office_convert.py find                       列出找到的 Word / Excel / PowerPoint / LibreOffice
  python office_convert.py topdf 檔案.pptx|docx|xlsx [-o 輸出.pdf]
  python office_convert.py slides 簡報.pptx [-o 圖片資料夾]   每頁輸出成 PNG
在程式裡：
  from office_convert import to_pdf, export_slides, find_app
"""
import argparse
import glob
import os
import shutil
import subprocess
import sys
import tempfile

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

APPS = {"word": "WINWORD.EXE", "excel": "EXCEL.EXE", "powerpoint": "POWERPNT.EXE", "libreoffice": "soffice.exe"}
PROGIDS = {"word": "Word.Application", "excel": "Excel.Application", "powerpoint": "PowerPoint.Application"}


def find_app(exe):
    """回傳執行檔完整路徑，找不到回傳 None。exe 例如 'POWERPNT.EXE'、'soffice.exe'。"""
    cands = []
    if sys.platform == "win32":
        import winreg
        for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            for view in (winreg.KEY_WOW64_64KEY, winreg.KEY_WOW64_32KEY):
                try:
                    k = winreg.OpenKey(hive, rf"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\{exe}",
                                       0, winreg.KEY_READ | view)
                    v, _ = winreg.QueryValueEx(k, None)
                    cands.append(os.path.expandvars(v.strip('"')))
                except OSError:
                    pass
        bases = {os.environ.get(n) for n in ("ProgramW6432", "ProgramFiles", "ProgramFiles(x86)")}
        bases |= {r"C:\Program Files", r"C:\Program Files (x86)"}
        for b in filter(None, bases):
            cands += glob.glob(os.path.join(b, "Microsoft Office", "root", "Office*", exe))
            cands += glob.glob(os.path.join(b, "Microsoft Office", "Office*", exe))
            cands += glob.glob(os.path.join(b, "LibreOffice*", "program", exe))
    w = shutil.which(exe) or shutil.which(os.path.splitext(exe)[0])
    if w:
        cands.append(w)
    for c in cands:
        if c and os.path.isfile(c):
            return c
    return None


def describe():
    lines = []
    for name, exe in APPS.items():
        p = find_app(exe)
        bit = ""
        if p:
            bit = "（32 位元）" if "(x86)" in p else "（64 位元）" if sys.platform == "win32" and "Program Files" in p else ""
        lines.append(f"{name:12s}: {p or '找不到'} {bit}")
    return "\n".join(lines)


def _powershells():
    """先用一般（64 位元）PowerShell，再用 32 位元 PowerShell。COM 在其中一邊壞掉時另一邊通常可以。"""
    root = os.environ.get("SystemRoot", r"C:\Windows")
    out = []
    for sub in ("Sysnative", "System32", "SysWOW64"):  # Sysnative：32 位元 Python 呼叫 64 位元 PowerShell
        p = os.path.join(root, sub, "WindowsPowerShell", "v1.0", "powershell.exe")
        if os.path.isfile(p) and p not in out:
            out.append(p)
    return out


def _ps_quote(s):
    return "'" + s.replace("'", "''") + "'"


def _run_ps(script):
    errors = []
    for ps in _powershells():
        r = subprocess.run([ps, "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
                           capture_output=True, text=True, encoding="utf-8", errors="replace")
        if r.returncode == 0:
            return ps
        errors.append(f"{ps}: {(r.stderr or r.stdout).strip()[:300]}")
    raise RuntimeError("\n".join(errors) or "找不到 PowerShell")


def _kind(path):
    ext = os.path.splitext(path)[1].lower()
    if ext in (".ppt", ".pptx", ".pptm", ".potx"):
        return "powerpoint"
    if ext in (".doc", ".docx", ".docm", ".rtf", ".dotx"):
        return "word"
    if ext in (".xls", ".xlsx", ".xlsm"):
        return "excel"
    raise ValueError(f"不支援的檔案類型: {ext}")


def _com_pdf_pywin32(kind, src, dst):
    import win32com.client  # noqa: 需要 pywin32
    app = win32com.client.DispatchEx(PROGIDS[kind])
    try:
        if kind == "word":
            app.Visible = False
            d = app.Documents.Open(src, ReadOnly=True)
            d.SaveAs2(dst, FileFormat=17)
            d.Close(False)
        elif kind == "powerpoint":
            p = app.Presentations.Open(src, ReadOnly=True, WithWindow=False)
            p.SaveAs(dst, 32)
            p.Close()
        else:
            app.Visible = False
            app.DisplayAlerts = False
            wb = app.Workbooks.Open(src, ReadOnly=True)
            wb.ExportAsFixedFormat(0, dst)
            wb.Close(False)
    finally:
        app.Quit()


def _com_pdf_ps(kind, src, dst):
    s, d = _ps_quote(src), _ps_quote(dst)
    if kind == "word":
        body = (f"$a=New-Object -ComObject Word.Application;$a.Visible=$false;"
                f"$x=$a.Documents.Open({s},$false,$true);$x.SaveAs2({d},17);$x.Close($false);$a.Quit()")
    elif kind == "powerpoint":
        body = (f"$a=New-Object -ComObject PowerPoint.Application;"
                f"$x=$a.Presentations.Open({s},-1,0,0);$x.SaveAs({d},32);$x.Close();$a.Quit()")
    else:
        body = (f"$a=New-Object -ComObject Excel.Application;$a.Visible=$false;$a.DisplayAlerts=$false;"
                f"$x=$a.Workbooks.Open({s},0,$true);$x.ExportAsFixedFormat(0,{d});$x.Close($false);$a.Quit()")
    return _run_ps("$ErrorActionPreference='Stop';" + body)


def _soffice_pdf(src, dst):
    so = find_app("soffice.exe")
    if not so:
        raise RuntimeError("找不到 LibreOffice")
    outdir = tempfile.mkdtemp(prefix="lo_")
    subprocess.run([so, "--headless", "--convert-to", "pdf", "--outdir", outdir, src],
                   check=True, capture_output=True)
    produced = os.path.join(outdir, os.path.splitext(os.path.basename(src))[0] + ".pdf")
    shutil.move(produced, dst)


def to_pdf(src, dst=None):
    """轉 PDF。依序試：pywin32 COM → PowerShell COM（64 位元，再 32 位元）→ LibreOffice。回傳使用的方法。"""
    src = os.path.abspath(src)
    dst = os.path.abspath(dst or os.path.splitext(src)[0] + ".pdf")
    kind = _kind(src)
    errors = []
    if sys.platform == "win32":
        try:
            _com_pdf_pywin32(kind, src, dst)
            return "Office（pywin32）"
        except Exception as e:
            errors.append(f"pywin32: {e}")
        try:
            ps = _com_pdf_ps(kind, src, dst)
            return f"Office（{'32' if 'SysWOW64' in ps else '64'} 位元 PowerShell）"
        except Exception as e:
            errors.append(f"PowerShell: {e}")
    try:
        _soffice_pdf(src, dst)
        return "LibreOffice"
    except Exception as e:
        errors.append(f"LibreOffice: {e}")
    raise RuntimeError("轉 PDF 失敗：\n" + "\n".join(errors) +
                       "\n需要 Microsoft Office 或 LibreOffice（winget install TheDocumentFoundation.LibreOffice）")


def export_slides(pptx, outdir=None, width=1600):
    """每頁簡報輸出成 PNG，回傳圖片路徑清單。PowerPoint COM 優先，不行就轉 PDF 再用 PyMuPDF 轉圖。"""
    pptx = os.path.abspath(pptx)
    outdir = os.path.abspath(outdir or os.path.splitext(pptx)[0] + "_頁面圖")
    os.makedirs(outdir, exist_ok=True)
    for f in glob.glob(os.path.join(outdir, "slide_*.png")):
        os.remove(f)
    if sys.platform == "win32":
        h = int(width * 9 / 16)
        script = ("$ErrorActionPreference='Stop';$a=New-Object -ComObject PowerPoint.Application;"
                  f"$x=$a.Presentations.Open({_ps_quote(pptx)},-1,0,0);"
                  "for($i=1;$i -le $x.Slides.Count;$i++){"
                  f"$x.Slides.Item($i).Export((Join-Path {_ps_quote(outdir)} ('slide_{{0:D3}}.png' -f $i)),'PNG',{width},{h})}};"
                  "$x.Close();$a.Quit()")
        try:
            _run_ps(script)
            files = sorted(glob.glob(os.path.join(outdir, "slide_*.png")))
            if files:
                return files
        except Exception as e:
            print(f"[PowerPoint 匯出失敗，改用 PDF 轉圖] {e}", file=sys.stderr)
    pdf = os.path.join(tempfile.mkdtemp(prefix="slides_"), "deck.pdf")
    to_pdf(pptx, pdf)
    try:
        import pymupdf
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "pymupdf"])
        import pymupdf
    files = []
    with pymupdf.open(pdf) as doc:
        for i, page in enumerate(doc, 1):
            p = os.path.join(outdir, f"slide_{i:03d}.png")
            page.get_pixmap(dpi=110).save(p)
            files.append(p)
    return files


def main():
    ap = argparse.ArgumentParser(description="Office 尋找與轉檔")
    sub = ap.add_subparsers(dest="cmd", required=True)
    sub.add_parser("find")
    s = sub.add_parser("topdf"); s.add_argument("file"); s.add_argument("-o", "--output")
    s = sub.add_parser("slides"); s.add_argument("file"); s.add_argument("-o", "--output")
    a = ap.parse_args()
    if a.cmd == "find":
        print(describe())
    elif a.cmd == "topdf":
        dst = a.output or os.path.splitext(a.file)[0] + ".pdf"
        how = to_pdf(a.file, dst)
        print(f"已用 {how} 轉成: {os.path.abspath(dst)}")
    else:
        files = export_slides(a.file, a.output)
        print(f"已輸出 {len(files)} 張: {os.path.dirname(files[0]) if files else ''}")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        sys.exit(str(e))
