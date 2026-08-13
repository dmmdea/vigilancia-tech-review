#!/usr/bin/env python3
"""Convert a presentation (pptx/ppt) to PDF and report its page count.

Usage:
    python convert_to_pdf.py <input.pptx|.ppt|.pdf> <output.pdf>

Conversion backends, tried in order:
  1. LibreOffice headless (`soffice`) — cross-platform.
  2. PowerPoint COM via PowerShell — Windows with Office installed.

If the input is already a PDF it is copied as-is. Prints JSON:
    {"pdf": "<path>", "pages": N, "backend": "soffice|powerpoint|copy"}
Exit codes: 0 ok · 2 environment problem (no backend installed / pypdf missing —
stop the whole run) · 3 THIS file failed to convert (record as NO REVISADO and
continue with the other files).
"""
import json
import os
import shutil
import subprocess
import sys
import tempfile

try:
    from pypdf import PdfReader
except ImportError:
    print("ERROR: pypdf is not installed. Run: pip install pypdf", file=sys.stderr)
    sys.exit(2)


def find_soffice() -> str | None:
    exe = shutil.which("soffice")
    if not exe and os.name == "nt":
        for cand in (r"C:\Program Files\LibreOffice\program\soffice.exe",
                     r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"):
            if os.path.exists(cand):
                return cand
    return exe


def powerpoint_available() -> bool:
    if os.name != "nt":
        return False
    try:
        res = subprocess.run(
            ["powershell", "-NoProfile", "-Command",
             "try { $p = New-Object -ComObject PowerPoint.Application; $p.Quit(); "
             "exit 0 } catch { exit 1 }"],
            capture_output=True, timeout=120,
        )
    except (subprocess.TimeoutExpired, OSError):
        return False
    return res.returncode == 0


def run_soffice(exe: str, src: str, dst: str) -> bool:
    outdir = os.path.dirname(os.path.abspath(dst)) or "."
    os.makedirs(outdir, exist_ok=True)
    produced = os.path.join(
        outdir, os.path.splitext(os.path.basename(src))[0] + ".pdf")
    # Remove any stale artifact so a failed run can't be validated against it.
    if os.path.exists(produced) and os.path.abspath(produced) != os.path.abspath(src):
        os.unlink(produced)
    try:
        res = subprocess.run(
            [exe, "--headless", "--convert-to", "pdf", "--outdir", outdir, src],
            capture_output=True, text=True, errors="replace", timeout=600,
        )
    except subprocess.TimeoutExpired:
        print("soffice timed out after 600s.", file=sys.stderr)
        return False
    if res.returncode != 0 or not os.path.exists(produced):
        print(f"soffice failed (rc={res.returncode}): "
              f"{(res.stderr or res.stdout).strip()[:400]}", file=sys.stderr)
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


def run_powerpoint(src: str, dst: str) -> bool:
    os.makedirs(os.path.dirname(os.path.abspath(dst)) or ".", exist_ok=True)
    with tempfile.NamedTemporaryFile(
            mode="w", suffix=".ps1", delete=False, encoding="utf-8") as f:
        f.write(PS_TEMPLATE)
        script = f.name
    try:
        env = dict(os.environ,
                   VTR_SRC=os.path.abspath(src), VTR_DST=os.path.abspath(dst))
        try:
            res = subprocess.run(
                ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
                 "-File", script],
                capture_output=True, text=True, errors="replace",
                timeout=600, env=env,
            )
        except subprocess.TimeoutExpired:
            print("PowerPoint COM timed out after 600s.", file=sys.stderr)
            return False
        if res.returncode != 0 or not os.path.exists(dst):
            print(f"PowerPoint COM failed (rc={res.returncode}): "
                  f"{(res.stderr or res.stdout).strip()[:400]}", file=sys.stderr)
            return False
        return True
    finally:
        os.unlink(script)


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    if len(sys.argv) != 3:
        print(__doc__, file=sys.stderr)
        sys.exit(1)
    src, dst = sys.argv[1], sys.argv[2]
    if not os.path.exists(src):
        print(f"ERROR: input not found: {src}", file=sys.stderr)
        sys.exit(3)

    backend = None
    if src.lower().endswith(".pdf"):
        os.makedirs(os.path.dirname(os.path.abspath(dst)) or ".", exist_ok=True)
        if os.path.abspath(src) != os.path.abspath(dst):
            shutil.copyfile(src, dst)
        backend = "copy"
    else:
        # "Nothing installed" (exit 2, environment, stop the run) must never be
        # conflated with "this one file failed" (exit 3, per-file). The COM
        # probe is expensive (~7s, launches PowerPoint), so probe it lazily —
        # only when soffice is absent or failed on this file.
        soffice = find_soffice()
        if soffice and run_soffice(soffice, src, dst):
            backend = "soffice"
        else:
            has_ppt = powerpoint_available()
            if has_ppt and run_powerpoint(src, dst):
                backend = "powerpoint"
            elif not soffice and not has_ppt:
                print("ERROR: no conversion backend installed. Install "
                      "LibreOffice (https://libreoffice.org) or Microsoft "
                      "PowerPoint.", file=sys.stderr)
                sys.exit(2)
            else:
                print(f"ERROR: conversion failed for this file: {src}",
                      file=sys.stderr)
                sys.exit(3)

    try:
        pages = len(PdfReader(dst).pages)
    except Exception as e:  # produced PDF unreadable -> conversion failed
        print(f"ERROR: produced PDF unreadable: {e}", file=sys.stderr)
        sys.exit(3)
    if pages == 0:
        print(f"ERROR: produced PDF has 0 pages ({dst}) — conversion failed.",
              file=sys.stderr)
        sys.exit(3)
    print(json.dumps({"pdf": dst, "pages": pages, "backend": backend}))


if __name__ == "__main__":
    main()
