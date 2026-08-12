#!/usr/bin/env python3
"""Convert a presentation (pptx/ppt) to PDF and report its page count.

Usage:
    python convert_to_pdf.py <input.pptx|.ppt|.pdf> <output.pdf>

Conversion backends, tried in order:
  1. LibreOffice headless (`soffice`) — cross-platform.
  2. PowerPoint COM via PowerShell — Windows with Office installed.

If the input is already a PDF it is copied as-is. Prints JSON:
    {"pdf": "<path>", "pages": N, "backend": "soffice|powerpoint|copy"}
Exit codes: 0 ok · 2 no backend available · 3 conversion failed.
"""
import json
import os
import shutil
import subprocess
import sys
import tempfile


def page_count(pdf_path: str) -> int:
    from pypdf import PdfReader
    return len(PdfReader(pdf_path).pages)


def try_soffice(src: str, dst: str) -> bool:
    exe = shutil.which("soffice")
    if not exe and os.name == "nt":
        for cand in (r"C:\Program Files\LibreOffice\program\soffice.exe",
                     r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"):
            if os.path.exists(cand):
                exe = cand
                break
    if not exe:
        return False
    outdir = os.path.dirname(os.path.abspath(dst)) or "."
    res = subprocess.run(
        [exe, "--headless", "--convert-to", "pdf", "--outdir", outdir, src],
        capture_output=True, text=True, timeout=600,
    )
    produced = os.path.join(
        outdir, os.path.splitext(os.path.basename(src))[0] + ".pdf")
    if res.returncode != 0 or not os.path.exists(produced):
        print(f"soffice failed: {res.stderr.strip()[:400]}", file=sys.stderr)
        return False
    if os.path.abspath(produced) != os.path.abspath(dst):
        shutil.move(produced, dst)
    return True


PS_TEMPLATE = r"""
$ErrorActionPreference = 'Stop'
$pp = New-Object -ComObject PowerPoint.Application
try {
  $pres = $pp.Presentations.Open($env:VTR_SRC, $true, $false, $false)
  $pres.SaveAs($env:VTR_DST, 32)  # 32 = ppSaveAsPDF
  $pres.Close()
} finally {
  $pp.Quit()
}
"""


def try_powerpoint(src: str, dst: str) -> bool:
    if os.name != "nt":
        return False
    with tempfile.NamedTemporaryFile(
            mode="w", suffix=".ps1", delete=False, encoding="utf-8") as f:
        f.write(PS_TEMPLATE)
        script = f.name
    try:
        env = dict(os.environ,
                   VTR_SRC=os.path.abspath(src), VTR_DST=os.path.abspath(dst))
        res = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
             "-File", script],
            capture_output=True, text=True, timeout=600, env=env,
        )
        if res.returncode != 0 or not os.path.exists(dst):
            print(f"PowerPoint COM failed: {res.stderr.strip()[:400]}",
                  file=sys.stderr)
            return False
        return True
    finally:
        os.unlink(script)


def main() -> None:
    if len(sys.argv) != 3:
        print(__doc__, file=sys.stderr)
        sys.exit(1)
    src, dst = sys.argv[1], sys.argv[2]
    if not os.path.exists(src):
        print(f"ERROR: input not found: {src}", file=sys.stderr)
        sys.exit(3)

    backend = None
    if src.lower().endswith(".pdf"):
        if os.path.abspath(src) != os.path.abspath(dst):
            shutil.copyfile(src, dst)
        backend = "copy"
    elif try_soffice(src, dst):
        backend = "soffice"
    elif try_powerpoint(src, dst):
        backend = "powerpoint"
    else:
        print("ERROR: no conversion backend. Install LibreOffice "
              "(https://libreoffice.org) or Microsoft PowerPoint.",
              file=sys.stderr)
        sys.exit(2)

    try:
        pages = page_count(dst)
    except Exception as e:  # produced PDF unreadable -> conversion failed
        print(f"ERROR: produced PDF unreadable: {e}", file=sys.stderr)
        sys.exit(3)
    print(json.dumps({"pdf": dst, "pages": pages, "backend": backend}))


if __name__ == "__main__":
    main()
