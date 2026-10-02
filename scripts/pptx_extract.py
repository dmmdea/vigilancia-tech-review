#!/usr/bin/env python3
"""Rebuild a .pptx slide by slide when no renderer will open it.

Some valid packages are refused by PowerPoint (programmatic generators,
re-zipped decks — two real final-round submissions failed every retry with
"corrupted and unreadable" / "could not open the file", although the zip and
every XML part were intact). A renderer failure must never leave a student
unreviewed, so this module reads the OOXML package with the standard library:

  * slide ORDER from ppt/presentation.xml (sldIdLst) — part names are not order
  * each slide's text (text boxes, tables, SmartArt) and its speaker notes
  * chart data (series, categories, values) as text — charts are often the PoC
  * the images placed on each slide, at full resolution
  * docProps/thumbnail.jpeg (first slide only) when present

Same output shape as keynote_extract.extract(), so the bundle builder treats
both as one "rebuilt deck".

Usage (debug): python pptx_extract.py <file.pptx> <outdir>
"""
import json
import os
import posixpath
import re
import shutil
import sys
import zipfile
import xml.etree.ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from keynote_extract import _write, VIEW_EXT  # noqa: E402  same media rules

for _s in (sys.stdout, sys.stderr):
    if _s in (sys.__stdout__, sys.__stderr__) and hasattr(_s, "reconfigure"):
        _s.reconfigure(encoding="utf-8")

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}
R_ID = "{%s}id" % NS["r"]


def _rels(z, part):
    """{rId: (type_suffix, target_part)} for a part, targets resolved."""
    d, b = posixpath.split(part)
    rp = posixpath.join(d, "_rels", b + ".rels")
    out = {}
    if rp not in z.namelist():
        return out
    root = ET.fromstring(z.read(rp))
    for rel in root.findall("rel:Relationship", NS):
        if rel.get("TargetMode") == "External":
            continue
        tgt = rel.get("Target", "")
        full = (tgt.lstrip("/") if tgt.startswith("/")
                else posixpath.normpath(posixpath.join(d, tgt)))
        out[rel.get("Id")] = (rel.get("Type", "").rsplit("/", 1)[-1], full)
    return out


def _paragraphs(xml_bytes):
    """Text of every a:p, one line each (tables and SmartArt included)."""
    try:
        root = ET.fromstring(xml_bytes)
    except ET.ParseError:
        return []
    lines = []
    for p in root.iter("{%s}p" % NS["a"]):
        t = "".join(x.text or "" for x in p.iter("{%s}t" % NS["a"])).strip()
        if t:
            lines.append(t)
    return lines


def _chart_text(xml_bytes):
    try:
        root = ET.fromstring(xml_bytes)
    except ET.ParseError:
        return ""
    out = []
    title = " ".join(t.text or "" for t in root.iter("{%s}t" % NS["a"])).strip()
    if title:
        out.append(f"Gráfico: {title}")
    for ser in root.iter("{%s}ser" % NS["c"]):
        def vals(tag):
            node = ser.find(f"c:{tag}", NS)
            return ([v.text for v in node.iter("{%s}v" % NS["c"])]
                    if node is not None else [])
        name = " ".join(vals("tx")) or "serie"
        cats = vals("cat") or vals("xVal")
        nums = vals("val") or vals("yVal")
        pairs = (", ".join(f"{c}: {v}" for c, v in zip(cats, nums))
                 if cats else ", ".join(nums))
        out.append(f"  {name} → {pairs}")
    return "\n".join(out)


def extract(src, outdir, ffmpeg=None):
    """Raises on an unreadable package (not a zip / no presentation.xml) and
    on one with no slide content at all."""
    outdir = os.path.abspath(outdir)
    os.makedirs(outdir, exist_ok=True)
    with zipfile.ZipFile(src) as z:
        names = set(z.namelist())
        pres = "ppt/presentation.xml"
        if pres not in names:
            raise ValueError("no es un .pptx (falta ppt/presentation.xml)")
        prels = _rels(z, pres)
        root = ET.fromstring(z.read(pres))
        order = [prels[s.get(R_ID)][1]
                 for s in root.findall("p:sldIdLst/p:sldId", NS)
                 if s.get(R_ID) in prels]
        order_exact = bool(order)
        if not order:              # fall back to part-name order
            order = sorted((n for n in names
                            if re.match(r"ppt/slides/slide\d+\.xml$", n)),
                           key=lambda n: int(re.findall(r"\d+", n)[-1]))
        # one file per media part, listed on EVERY slide that uses it: a deck
        # often reuses one image (a logo, a recurring screenshot) and each
        # slide's evidence must stay complete
        written = {}
        slides = []
        for k, part in enumerate(order, 1):
            if part not in names:
                continue
            lines = _paragraphs(z.read(part))
            images, extra = [], []
            for rid, (typ, tgt) in _rels(z, part).items():
                if tgt not in names:
                    continue
                if typ == "image":
                    ext = os.path.splitext(tgt)[1].lower()
                    if ext not in VIEW_EXT:
                        continue
                    if tgt not in written:
                        written[tgt] = _write(z, tgt, os.path.join(
                            outdir, f"m{len(written) + 1:03d}{ext}"), ffmpeg)
                    if written[tgt] not in images:
                        images.append(written[tgt])
                elif typ == "chart":
                    ct = _chart_text(z.read(tgt))
                    if ct:
                        extra.append(ct)
                elif typ in ("diagramData",):
                    extra.extend(_paragraphs(z.read(tgt)))
                elif typ == "notesSlide":
                    notes = [ln for ln in _paragraphs(z.read(tgt))
                             if not ln.isdigit()]   # drop the page number
                    if notes:
                        extra.append("Notas del orador: " + " / ".join(notes))
            text = "\n".join(lines + extra)
            slides.append({"n": len(slides) + 1, "thumb": None,
                           "text": text, "images": images})
        if slides and "docProps/thumbnail.jpeg" in names:
            slides[0]["thumb"] = _write(
                z, "docProps/thumbnail.jpeg",
                os.path.join(outdir, "s01_miniatura.jpeg"))
    if not any(s["text"] or s["images"] for s in slides):
        raise ValueError("el .pptx no contiene ninguna lámina legible")
    text_path = os.path.join(outdir, "texto_laminas.txt")
    with open(text_path, "w", encoding="utf-8") as f:
        for s in slides:
            f.write(f"=== Lámina {s['n']} ===\n{s['text'] or '(sin texto)'}\n\n")
    note = ("PowerPoint no pudo abrir el archivo: presentación reconstruida "
            "lámina por lámina (texto, gráficos como datos e imágenes de cada "
            "lámina; sin vista del diseño)")
    if not order_exact:
        note += " | AVISO: orden de láminas APROXIMADO"
    return {"slides": slides, "movies": [], "loose_images": [],
            "text_path": os.path.abspath(text_path),
            "order_exact": order_exact, "note": note}


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print(__doc__, file=sys.stderr)
        sys.exit(1)
    res = extract(sys.argv[1], sys.argv[2], ffmpeg=shutil.which("ffmpeg"))
    print(json.dumps({k: v for k, v in res.items() if k != "slides"},
                     ensure_ascii=False, indent=1))
    for s in res["slides"]:
        print(f"lámina {s['n']}: {len(s['text'])} car., "
              f"{len(s['images'])} imagen(es)")
