#!/usr/bin/env python3
"""Rebuild an Apple Keynote (.key) deck for review — no Keynote, no LibreOffice.

A .key file is a slide deck, so the fairness rule applies in full: the format
is never a reason to skip a student. Windows machines usually have neither
Keynote nor a LibreOffice build that reads it, and PowerPoint cannot open it.

Modern Keynote files (Keynote 6+, 2013 onward) are zip archives:
  * Index/Document.iwa       — the slide tree (presentation ORDER)
  * Index/Slide-<id>.iwa     — one per slide: its text boxes, notes and the
                               ids of the media placed on it
  * Data/st-<uuid>-<id>.jpg  — a small thumbnail Keynote keeps of every slide
  * Data/*                   — embedded images and movies at full resolution
The .iwa streams are protobuf messages framed in raw-snappy chunks. This
module decodes them with the standard library only (no protobuf schemas, no
third-party snappy), using the per-object reference lists Keynote records
(TSP.ArchiveInfo / MessageInfo object_references + data_references).

Slide ORDER comes from the slide tree, never from file numbering: the number
in Slide-<id>.iwa is creation order, and a slide inserted at the front of a
real submission carried the HIGHEST id. When the tree cannot be read the
order falls back to creation order and the result says so (order_exact).

Older Keynote '09 files (index.apxl[.gz], XML) fall back to the XML text plus
every embedded image, with no per-slide split.

Output (extract()): ordered slides with their text, thumbnail and full-size
images; a text file with every slide's text; embedded movies extracted to
disk (the caller routes them through the normal video path).

Usage (debug): python keynote_extract.py <file.key> <outdir>
"""
import gzip
import json
import os
import re
import shutil
import subprocess
import sys
import zipfile

for _s in (sys.stdout, sys.stderr):
    if _s in (sys.__stdout__, sys.__stderr__) and hasattr(_s, "reconfigure"):
        _s.reconfigure(encoding="utf-8")

IMAGE_EXT = {".png", ".jpg", ".jpeg", ".gif", ".tif", ".tiff", ".bmp", ".webp"}
HEIC_EXT = {".heic", ".heif"}      # iPhone photos: converted to JPEG via ffmpeg
VIEW_EXT = IMAGE_EXT | HEIC_EXT | {".pdf"}   # opened visually by the reviewer
MOVIE_EXT = {".mp4", ".mov", ".m4v"}
# Keynote's own furniture, not the student's content: template fills,
# bullets, placeholder art, master/slide thumbnails, low-res duplicates.
SKIP_DATA = re.compile(r"(^|/)(PresetImageFill|bullet_|PlaceholderImage|mt-|st-)"
                       r"|-small-\d+\.", re.IGNORECASE)
TEXT_STORAGE = 2001        # TSWP.StorageArchive: field 3 holds the text
OBJECT_CHAR = "￼"     # inline-attachment marker inside text storages
MAX_SIDE = 2400            # downscale huge photos so any image reader opens them


# ---- iwa decoding ------------------------------------------------------------
def _snappy(buf):
    """Raw (unframed) snappy decompression."""
    pos = 0
    while buf[pos] & 0x80:          # skip the uncompressed-length varint
        pos += 1
    pos += 1
    out = bytearray()
    n = len(buf)
    while pos < n:
        tag = buf[pos]
        pos += 1
        kind = tag & 3
        if kind == 0:               # literal
            ln = tag >> 2
            if ln < 60:
                ln += 1
            else:
                nb = ln - 59
                ln = int.from_bytes(buf[pos:pos + nb], "little") + 1
                pos += nb
            out += buf[pos:pos + ln]
            pos += ln
            continue
        if kind == 1:
            ln = ((tag >> 2) & 7) + 4
            off = ((tag >> 5) << 8) | buf[pos]
            pos += 1
        elif kind == 2:
            ln = (tag >> 2) + 1
            off = int.from_bytes(buf[pos:pos + 2], "little")
            pos += 2
        else:
            ln = (tag >> 2) + 1
            off = int.from_bytes(buf[pos:pos + 4], "little")
            pos += 4
        if off <= 0 or off > len(out):
            raise ValueError("snappy: copia fuera de rango")
        start = len(out) - off
        for i in range(ln):         # copies may overlap their own output
            out.append(out[start + i])
    return bytes(out)


def _iwa(data):
    """IWA chunk framing: 1-byte type (0) + 3-byte LE length + snappy body."""
    pos, out = 0, bytearray()
    while pos < len(data):
        if data[pos] != 0:
            raise ValueError(f"bloque iwa con cabecera {data[pos]}")
        ln = int.from_bytes(data[pos + 1:pos + 4], "little")
        pos += 4
        out += _snappy(data[pos:pos + ln])
        pos += ln
    return bytes(out)


def _varint(b, p):
    v, s = 0, 0
    while True:
        x = b[p]
        p += 1
        v |= (x & 0x7F) << s
        if x < 0x80:
            return v, p
        s += 7


def _fields(b):
    """Schema-less protobuf walk -> [(field, wiretype, value)]."""
    p, out = 0, []
    while p < len(b):
        key, p = _varint(b, p)
        f, wt = key >> 3, key & 7
        if wt == 0:
            v, p = _varint(b, p)
        elif wt == 1:
            v, p = b[p:p + 8], p + 8
        elif wt == 2:
            ln, p = _varint(b, p)
            v, p = b[p:p + ln], p + ln
        elif wt == 5:
            v, p = b[p:p + 4], p + 4
        else:
            raise ValueError(f"wiretype {wt} no soportado")
        out.append((f, wt, v))
    return out


def _packed(b):
    p, out = 0, []
    while p < len(b):
        v, p = _varint(b, p)
        out.append(v)
    return out


def _objects(stream):
    """{identifier: [(type, payload, object_refs, data_refs)]}"""
    objs = {}
    p = 0
    while p < len(stream):
        ln, p = _varint(stream, p)
        info = _fields(stream[p:p + ln])
        p += ln
        ident = next((v for f, wt, v in info if f == 1 and wt == 0), None)
        msgs = []
        for f, wt, v in info:
            if f != 2 or wt != 2:
                continue
            mi = _fields(v)
            typ = next((x for g, w, x in mi if g == 1 and w == 0), None)
            mlen = next((x for g, w, x in mi if g == 3 and w == 0), 0)
            orefs, drefs = [], []
            for g, w, x in mi:
                if g in (5, 6):
                    vals = _packed(x) if w == 2 else [x]
                    (orefs if g == 5 else drefs).extend(vals)
            msgs.append((typ, stream[p:p + mlen], orefs, drefs))
            p += mlen
        objs[ident] = msgs
    return objs


def _nested_ids(b, depth=0):
    """field-1 varints of nested messages, in payload order (TSP.Reference
    shape). Used only to recover the slide tree's child ORDER."""
    out = []
    try:
        fl = _fields(b)
    except (ValueError, IndexError):
        return out
    for f, wt, v in fl:
        if wt != 2 or depth > 6:
            continue
        try:
            sub = _fields(v)
        except (ValueError, IndexError):
            continue
        if sub and sub[0][0] == 1 and sub[0][1] == 0:
            out.append(sub[0][2])
        out.extend(_nested_ids(v, depth + 1))
    return out


def _storage_texts(payload):
    texts = []
    try:
        fl = _fields(payload)
    except (ValueError, IndexError):
        return texts
    for f, wt, v in fl:
        if f == 3 and wt == 2:
            try:
                s = v.decode("utf-8")
            except UnicodeDecodeError:
                continue
            s = s.replace(OBJECT_CHAR, "").replace(" ", "\n").strip()
            if s:
                texts.append(s)
    return texts


# ---- media helpers -----------------------------------------------------------
def _dims(path):
    """(width, height) from a PNG/JPEG header, else None."""
    try:
        with open(path, "rb") as f:
            head = f.read(512 * 1024)
    except OSError:
        return None
    if head[:8] == b"\x89PNG\r\n\x1a\n" and len(head) >= 24:
        return (int.from_bytes(head[16:20], "big"),
                int.from_bytes(head[20:24], "big"))
    if head[:2] == b"\xff\xd8":
        j = 2
        while j + 9 < len(head):
            if head[j] != 0xFF:
                j += 1
                continue
            if head[j + 1] in (0xC0, 0xC1, 0xC2):
                return (int.from_bytes(head[j + 7:j + 9], "big"),
                        int.from_bytes(head[j + 5:j + 7], "big"))
            j += 2 + int.from_bytes(head[j + 2:j + 4], "big")
    return None


def _ffmpeg_image(ffmpeg, src, dst, scale):
    args = [ffmpeg, "-y", "-v", "error", "-i", src]
    if scale:
        args += ["-vf", f"scale=w='min({MAX_SIDE},iw)':h='min({MAX_SIDE},ih)'"
                        ":force_original_aspect_ratio=decrease"]
    if dst.lower().endswith((".jpg", ".jpeg")):
        args += ["-q:v", "2"]
    try:
        r = subprocess.run(args + ["-frames:v", "1", dst],
                           capture_output=True, timeout=120)
    except (OSError, subprocess.SubprocessError):
        return False
    return r.returncode == 0 and os.path.exists(dst) and os.path.getsize(dst) > 0


def _write(z, name, dst, ffmpeg=None):
    """Extract one Data/ member. Photos larger than MAX_SIDE are scaled down
    in their OWN format (a PNG screenshot stays PNG: its text is the PoC
    evidence); HEIC becomes JPEG. Returns the path to hand the reviewer."""
    with z.open(name) as src, open(dst, "wb") as out:
        shutil.copyfileobj(src, out)
    ext = os.path.splitext(dst)[1].lower()
    if ext in HEIC_EXT:
        jpg = os.path.splitext(dst)[0] + ".jpg"
        if ffmpeg and _ffmpeg_image(ffmpeg, dst, jpg, scale=True):
            os.remove(dst)
            return jpg
        return dst            # left as-is; the bundle notes it may not open
    dims = _dims(dst) if (ffmpeg and ext in IMAGE_EXT) else None
    if not dims or max(dims) <= MAX_SIDE:
        return dst
    tmp = os.path.splitext(dst)[0] + "_tmp" + ext
    if _ffmpeg_image(ffmpeg, dst, tmp, scale=True):
        os.replace(tmp, dst)
    elif os.path.exists(tmp):
        os.remove(tmp)
    return dst


# ---- extraction ----------------------------------------------------------------
def extract(src, outdir, ffmpeg=None):
    """Rebuild the deck into `outdir`. Returns a dict:
       slides: [{n, thumb, text, images}], movies: [{path, name, slide}],
       text_path, order_exact, note.   Raises on an unreadable archive."""
    os.makedirs(outdir, exist_ok=True)
    with zipfile.ZipFile(src) as z:
        names = z.namelist()
        if "Index/Document.iwa" not in names:
            return _extract_legacy(z, names, outdir, ffmpeg)
        data_by_id = {}
        for n in names:
            m = re.search(r"-(\d+)\.[A-Za-z0-9]+$", n)
            if n.startswith("Data/") and m:
                data_by_id[int(m.group(1))] = n
        slide_files = {}
        for n in names:
            m = re.match(r"Index/Slide-(\d+)(?:-\d+)?\.iwa$", n)
            if m:
                slide_files[int(m.group(1))] = n
        doc = _objects(_iwa(z.read("Index/Document.iwa")))

        # slide node -> its slide (exactly one slide ref) + its thumbnail
        node_slide, node_thumb = {}, {}
        for oid, msgs in doc.items():
            for _t, _payload, orefs, drefs in msgs:
                hits = [r for r in orefs if r in slide_files]
                if len(hits) == 1:
                    node_slide[oid] = hits[0]
                    th = [data_by_id[d] for d in drefs
                          if d in data_by_id and "/st-" in data_by_id[d]]
                    node_thumb[oid] = th[0] if th else None
        # the container that lists the most slide nodes, in payload order
        best = []
        for _oid, msgs in doc.items():
            for _t, payload, _o, _d in msgs:
                seq = []
                for i in _nested_ids(payload):
                    if i in node_slide and i not in seq:
                        seq.append(i)
                if len(seq) > len(best):
                    best = seq
        order_exact = bool(best) and len(best) == len(node_slide) > 0
        if node_slide:
            ordered = best + sorted((n for n in node_slide if n not in best),
                                    key=lambda n: node_slide[n])
            plan = [(node_slide[n], node_thumb.get(n)) for n in ordered]
        else:
            plan = [(sid, None) for sid in sorted(slide_files)]

        slides, movies, used = [], [], set()
        for k, (sid, thumb) in enumerate(plan, 1):
            objs = _objects(_iwa(z.read(slide_files[sid])))
            texts, drefs = [], []
            for _oid, msgs in objs.items():
                for t, payload, _o, dr in msgs:
                    drefs.extend(dr)
                    if t == TEXT_STORAGE:
                        texts.extend(_storage_texts(payload))
            entry = {"n": k, "thumb": None, "text": "\n".join(texts),
                     "images": []}
            if thumb:
                entry["thumb"] = _write(
                    z, thumb, os.path.join(outdir, f"s{k:02d}_miniatura.jpg"))
            seen = []
            for d in drefs:
                name = data_by_id.get(d)
                if not name or name in seen or SKIP_DATA.search(name):
                    continue
                seen.append(name)
                ext = os.path.splitext(name)[1].lower()
                if ext in VIEW_EXT:
                    entry["images"].append(_write(
                        z, name, os.path.join(
                            outdir, f"s{k:02d}_img{len(entry['images']) + 1}{ext}"),
                        ffmpeg))
                    used.add(name)
                elif ext in MOVIE_EXT:
                    dst = os.path.join(outdir, f"s{k:02d}_mov{len(movies) + 1}{ext}")
                    _write(z, name, dst)
                    movies.append({"path": dst, "name": os.path.basename(name),
                                   "slide": k})
                    used.add(name)
            slides.append(entry)

        # media the deck holds but no slide references (removed from a slide,
        # or referenced from somewhere this reader does not follow): still the
        # student's material, so list it instead of dropping it
        loose = []
        for name in sorted(set(data_by_id.values()) - used):
            if SKIP_DATA.search(name):
                continue
            ext = os.path.splitext(name)[1].lower()
            if ext in VIEW_EXT:
                loose.append(_write(z, name, os.path.join(
                    outdir, f"extra_img{len(loose) + 1}{ext}"), ffmpeg))
            elif ext in MOVIE_EXT:
                dst = os.path.join(outdir, f"extra_mov{len(movies) + 1}{ext}")
                _write(z, name, dst)
                movies.append({"path": dst, "name": os.path.basename(name),
                               "slide": None})

    text_path = os.path.join(outdir, "texto_laminas.txt")
    with open(text_path, "w", encoding="utf-8") as f:
        for s in slides:
            f.write(f"=== Lámina {s['n']} ===\n{s['text'] or '(sin texto)'}\n\n")
    note = ("presentación Keynote reconstruida lámina por lámina "
            "(texto, miniatura e imágenes de cada lámina)")
    if not order_exact:
        note += " | AVISO: orden de láminas APROXIMADO (orden de creación)"
    if loose:
        note += (f" | {len(loose)} imagen(es) del archivo sin lámina "
                 "identificada, listadas aparte")
    return {"slides": slides, "movies": movies, "loose_images": loose,
            "text_path": os.path.abspath(text_path),
            "order_exact": order_exact, "note": note}


def _extract_legacy(z, names, outdir, ffmpeg):
    """Keynote '09 (index.apxl XML): text + every image, no slide split."""
    xml_name = next((n for n in names if n in ("index.apxl", "index.apxl.gz")),
                    None)
    if xml_name is None:
        raise ValueError("no es un Keynote reconocible (falta Index/Document.iwa "
                         "e index.apxl)")
    raw = z.read(xml_name)
    if xml_name.endswith(".gz"):
        raw = gzip.decompress(raw)
    text = re.sub(r"<[^>]+>", " ", raw.decode("utf-8", errors="replace"))
    text = re.sub(r"[ \t]+", " ", text)
    text = "\n".join(ln.strip() for ln in text.splitlines() if ln.strip())
    images, movies = [], []
    for n in names:
        ext = os.path.splitext(n)[1].lower()
        if n.endswith("/") or SKIP_DATA.search(n):
            continue
        if ext in VIEW_EXT:
            images.append(_write(z, n, os.path.join(
                outdir, f"img{len(images) + 1}{ext}"), ffmpeg))
        elif ext in MOVIE_EXT:
            dst = os.path.join(outdir, f"mov{len(movies) + 1}{ext}")
            _write(z, n, dst)
            movies.append({"path": dst, "name": os.path.basename(n),
                           "slide": None})
    text_path = os.path.join(outdir, "texto_laminas.txt")
    with open(text_path, "w", encoding="utf-8") as f:
        f.write(text or "(sin texto)")
    return {"slides": [{"n": 1, "thumb": None, "text": text, "images": images}],
            "movies": movies, "loose_images": [],
            "text_path": os.path.abspath(text_path), "order_exact": False,
            "note": "Keynote antiguo (formato XML): texto completo e imágenes, "
                    "SIN división por láminas"}


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print(__doc__, file=sys.stderr)
        sys.exit(1)
    res = extract(sys.argv[1], sys.argv[2], ffmpeg=shutil.which("ffmpeg"))
    print(json.dumps({k: v for k, v in res.items() if k != "slides"},
                     ensure_ascii=False, indent=1))
    for s in res["slides"]:
        print(f"lámina {s['n']}: {len(s['text'])} car. de texto, "
              f"miniatura={'sí' if s['thumb'] else 'no'}, "
              f"{len(s['images'])} imagen(es)")
