#!/usr/bin/env python3
"""Rasterize a PDF into one PNG per page — the portability fallback for
harnesses whose file-reading tool cannot render PDF pages visually.

Claude Code's Read tool renders PDFs page-by-page, so it never needs this.
Codex CLI and Antigravity read files as text/images but do not paginate PDFs
visually — on those harnesses the orchestrator runs this script and hands the
reviewer the PNG paths instead of the PDF path (the reviewer's vision then
sees exactly the same pages).

Usage:
    python pdf_to_images.py <input.pdf> <outdir> [--dpi=140] [--max-pages=60]

Prints JSON: {"pages": N, "images": ["<outdir>/p001.png", ...]}
Exit codes: 0 ok · 1 usage · 2 environment problem (pymupdf missing — stop
and install) · 3 THIS file failed (corrupt/encrypted PDF — per-file failure).

Requires PyMuPDF: pip install pymupdf
"""
import json
import os
import sys

for _s in (sys.stdout, sys.stderr):
    if _s in (sys.__stdout__, sys.__stderr__) and hasattr(_s, "reconfigure"):
        _s.reconfigure(encoding="utf-8")

try:
    import pymupdf as fitz  # modern name; the old 'import fitz' prints a
                            # deprecation warning ON STDOUT that corrupts the
                            # JSON contract for callers that parse our output
except ImportError:
    try:
        import fitz  # very old PyMuPDF without the pymupdf alias
    except ImportError:
        print("ERROR: PyMuPDF no está instalado. Ejecuta: pip install pymupdf",
              file=sys.stderr)
        sys.exit(2)


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    if len(args) != 2:
        print(__doc__, file=sys.stderr)
        sys.exit(1)
    src, outdir = args
    dpi = 140
    max_pages = 60
    for a in sys.argv[1:]:
        if a.startswith("--dpi="):
            dpi = int(a.split("=", 1)[1])
        elif a.startswith("--max-pages="):
            max_pages = int(a.split("=", 1)[1])

    if not os.path.exists(src):
        print(f"ERROR: no existe: {src}", file=sys.stderr)
        sys.exit(3)
    os.makedirs(outdir, exist_ok=True)

    try:
        doc = fitz.open(src)
        if doc.needs_pass:
            print("ERROR: PDF protegido con contraseña.", file=sys.stderr)
            sys.exit(3)
        n = min(len(doc), max_pages)
        if n == 0:
            print("ERROR: PDF con 0 páginas.", file=sys.stderr)
            sys.exit(3)
        images = []
        zoom = dpi / 72.0
        mat = fitz.Matrix(zoom, zoom)
        for i in range(n):
            out = os.path.join(outdir, f"p{i + 1:03d}.png")
            doc[i].get_pixmap(matrix=mat).save(out)
            images.append(os.path.abspath(out))
        truncated = len(doc) > max_pages
    except Exception as e:
        print(f"ERROR: no se pudo rasterizar: {e}", file=sys.stderr)
        sys.exit(3)

    result = {"pages": len(images), "images": images}
    if truncated:
        result["truncated_from"] = len(doc)
        print(f"AVISO: PDF de {len(doc)} páginas truncado a {max_pages} — "
              "revisa el resto manualmente o sube --max-pages.",
              file=sys.stderr)
    print(json.dumps(result, ensure_ascii=False))


if __name__ == "__main__":
    main()
