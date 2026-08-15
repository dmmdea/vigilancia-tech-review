#!/usr/bin/env python3
"""Chequeo de entorno en un solo comando — córrelo ANTES de cualquier corrida.

    python scripts/preflight.py [--rasterize] [--json]

Verifica todo lo que la corrida necesitará y lo reporta en español, con un
veredicto claro por ítem (OK / FALTA / OPCIONAL). Diseñado para que el
operador sepa en 10 segundos si puede correr el skill completo, en lugar de
descubrir un backend faltante a mitad de la revisión de 75 estudiantes.

Exit 0 = todo lo obligatorio presente · 1 = falta algo obligatorio.
--rasterize: además exige pymupdf (harnesses sin visión de páginas PDF).
--json: emite el resultado como JSON (para orquestadores).
"""
import importlib.util
import json
import os
import shutil
import subprocess
import sys

for _s in (sys.stdout, sys.stderr):
    if _s in (sys.__stdout__, sys.__stderr__) and hasattr(_s, "reconfigure"):
        _s.reconfigure(encoding="utf-8")


def have_module(name):
    return importlib.util.find_spec(name) is not None


def find_exe(*cands):
    for c in cands:
        if os.path.sep in c or (os.name == "nt" and ":" in c):
            if os.path.exists(c):
                return c
        else:
            w = shutil.which(c)
            if w:
                return w
    return None


def com_available(prog):
    if os.name != "nt":
        return False
    try:
        r = subprocess.run(
            ["powershell", "-NoProfile", "-Command",
             f"try {{ $o = New-Object -ComObject {prog}; $o.Quit(); exit 0 }} "
             "catch { exit 1 }"],
            capture_output=True, timeout=90)
        return r.returncode == 0
    except (subprocess.TimeoutExpired, OSError):
        return False


def main():
    need_raster = "--rasterize" in sys.argv[1:]
    as_json = "--json" in sys.argv[1:]
    checks = []

    def check(name, ok, required, hint):
        checks.append({"item": name, "ok": bool(ok), "required": required,
                       "hint": hint})

    check("Python >= 3.10", sys.version_info >= (3, 10), True,
          "instala Python 3.10 o superior")
    check("openpyxl (Excel)", have_module("openpyxl"), True,
          "pip install -r requirements.txt")
    check("pypdf (conteo de páginas)", have_module("pypdf"), True,
          "pip install -r requirements.txt")
    check("pymupdf (rasterizar PDFs)",
          have_module("pymupdf") or have_module("fitz"), need_raster,
          "pip install pymupdf — obligatorio en harnesses sin visión de PDF")

    soffice = find_exe("soffice",
                       r"C:\Program Files\LibreOffice\program\soffice.exe")
    ppt = com_available("PowerPoint.Application") if not soffice else False
    check("diapositivas→PDF (LibreOffice o PowerPoint)", bool(soffice or ppt),
          True, "instala LibreOffice (libreoffice.org) o Microsoft PowerPoint")
    word = com_available("Word.Application") if not soffice else False
    check("docx→PDF (Word o LibreOffice)", bool(word or soffice), True,
          "instala LibreOffice o Microsoft Word — sin esto las entregas .docx "
          "no se pueden revisar (¡regla de equidad!)")
    chrome = find_exe(
        "chrome", "msedge",
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe")
    check("html→PDF (Chrome/Edge headless)", bool(chrome), False,
          "instala Chrome o Edge — sin esto las entregas .html no se revisan")
    check("ffmpeg (fotogramas de video)", bool(find_exe("ffmpeg")), False,
          "instala ffmpeg — sin esto los videos se revisan sin fotogramas")

    missing = [c for c in checks if c["required"] and not c["ok"]]
    if as_json:
        print(json.dumps({"ok": not missing, "checks": checks},
                         ensure_ascii=False, indent=2))
    else:
        for c in checks:
            mark = "OK     " if c["ok"] else (
                "FALTA  " if c["required"] else "OPCIONAL faltante")
            print(f"[{mark}] {c['item']}")
            if not c["ok"]:
                print(f"          → {c['hint']}")
        print()
        if missing:
            print(f"NO LISTO: {len(missing)} requisito(s) obligatorio(s) "
                  "faltante(s) — corrige lo de arriba antes de correr el skill.")
        else:
            opt = [c for c in checks if not c["ok"] and not c["required"]]
            print("LISTO para correr."
                  + (f" ({len(opt)} opcional(es) faltante(s) — algunos "
                     "formatos tendrán revisión degradada.)" if opt else ""))
    sys.exit(1 if missing else 0)


if __name__ == "__main__":
    main()
