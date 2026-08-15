#!/usr/bin/env python3
"""List a LOCAL submissions folder (Google Drive Desktop mount, Canvas bulk
download, or any plain folder) and classify what to review.

Use this instead of drive_list.py + drive_download.py when the teaching team
already has the submissions on disk — e.g. a Drive Desktop mount under
G:\\.shortcut-targets-by-id\\... or an unzipped Canvas bulk download.

Two supported shapes, auto-detected:
  * NESTED  (Canvas): one subfolder per student, named
              "<studentId>-<assignmentId> - Student Name - <timestamp>"
  * FLAT:   files sitting directly in the folder.

Writes into <outdir>:
  listing.json        — drive_list.py-compatible listing of the top level
  sub-listings/*.json — one per student subfolder (nested shape only)
  manifest.json       — every file with its path/size
  review_plan.json    — per folder: which file is the primary submission,
                        which are supplementary evidence, which folders are
                        superseded duplicate submissions

Usage:
    python local_list.py <folder> <outdir>

Exit codes: 0 ok · 2 folder missing/unreadable · 3 nothing found.

WINDOWS LONG PATHS: student folders + filenames routinely push past the 260
char MAX_PATH limit. Every filesystem call here goes through longpath(), which
applies the \\\\?\\ extended-length prefix — without it os.path.exists() and
open() silently fail on the deepest submissions and students get dropped.
"""
import hashlib
import json
import os
import re
import sys

for _s in (sys.stdout, sys.stderr):
    if _s in (sys.__stdout__, sys.__stderr__) and hasattr(_s, "reconfigure"):
        _s.reconfigure(encoding="utf-8")

REVIEWABLE_EXT = {".pptx", ".ppt", ".pdf", ".odp"}
# NOTE: never put a legitimate submission format here (an earlier version
# skipped "index.html" — a real student submission shape; silently dropped).
SKIP_NAMES = {"desktop.ini", ".ds_store", "thumbs.db"}

CANVAS_RE = re.compile(
    r"^(\d+-\d+)\s*-\s*(.+?)\s*-\s*(\d{1,2} de \w+ de \d{4} \d{1,2}[_:]\d{2})$")
MESES = {"enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5,
         "junio": 6, "julio": 7, "agosto": 8, "septiembre": 9,
         "octubre": 10, "noviembre": 11, "diciembre": 12}


def longpath(p: str) -> str:
    """Win32 extended-length prefix — REQUIRED past 260 chars."""
    if os.name != "nt":
        return p
    ap = os.path.abspath(p)
    if ap.startswith("\\\\?\\"):
        return ap
    if ap.startswith("\\\\"):
        return "\\\\?\\UNC\\" + ap[2:]
    return "\\\\?\\" + ap


def sid(*parts: str) -> str:
    return "local-" + hashlib.sha256("||".join(parts).encode("utf-8")).hexdigest()[:16]


def wopen(path, mode="w"):
    """Open for writing through the long-path prefix. The outdir the caller
    passes can itself be deep (session scratchpads often are), and these are
    plain Python writes, so the prefix is always safe here."""
    return open(longpath(path), mode, encoding="utf-8")


def listdir(p):
    return sorted(os.listdir(longpath(p)))


def isdir(p):
    return os.path.isdir(longpath(p))


def getsize(p):
    """Size, or (None, reason). On a Drive Desktop mount an OSError here often
    predicts a later copy failure — keep the cause for diagnostics."""
    try:
        return os.path.getsize(longpath(p)), None
    except OSError as e:
        return None, str(e)


# 'nueva'/'última' are ordinary Spanish adjectives ("última milla", "nueva
# era") — they only count as version markers ADJACENT to a version noun
# ("última versión", "versión nueva", "entrega final"). Bare final/definitiva/
# corregida remain markers (reviewed tradeoff: strongly version-shaped).
VERSION_MARK = re.compile(
    r"(?:^|[\s_\-.(])("
    r"v(?:ersi[oó]n)?\s*(\d+)"
    r"|(?:versi[oó]n|entrega)\s+(?:final|nueva|[uú]ltima|definitiva)"
    r"|(?:[uú]ltima|nueva)\s+(?:versi[oó]n|entrega)"
    r"|final(?:isim[ao])?|definitiv[ao]|corregid[ao]"
    r"|\((\d+)\))(?:[\s_\-.)]|$)",
    re.IGNORECASE)


def norm_person(name: str) -> str:
    """Accent/case/space-insensitive person key — for FLAGGING possible
    duplicates only, never for auto-superseding: two distinct students can
    legitimately share a name, and silently discarding one student's work on
    a name match would be the worst possible bug in a grading tool."""
    import unicodedata
    n = unicodedata.normalize("NFKD", name or "")
    n = "".join(c for c in n if not unicodedata.combining(c))
    return " ".join(n.lower().split())


def version_rank(fname: str):
    """Sortable version signal from a filename, or None when the name carries
    no version marker. 'final/definitiva/última'-type words BEAT numbered
    versions (they mean the last one by definition: v1 < v2 < FINAL);
    among numbered, higher wins."""
    best = None
    for m in VERSION_MARK.finditer(fname):
        num = m.group(2) or m.group(3)
        r = (1, int(num)) if num else (2, 0)   # keyword-final outranks vN
        if best is None or r > best:
            best = r
    return best


def strip_versions(stem: str) -> str:
    return " ".join(VERSION_MARK.sub(" ", stem).lower().split())


def parse_ts(ts: str):
    """Returns a sortable tuple, or None when the timestamp cannot be parsed.
    Callers must treat None as 'cannot decide order' and SAY SO — a silent
    fallback here once risked grading the WRONG duplicate as authoritative."""
    try:
        parts = ts.replace("_", ":").split(" de ")
        day = int(parts[0])
        month = MESES[parts[1].lower()]
        rest = parts[2].split(" ")
        year = int(rest[0])
        hh, mm = rest[1].split(":")
        return (year, month, day, int(hh), int(mm))
    except Exception:
        return None


def collect_files(folder, notes=None, depth=1):
    """Files inside `folder`, descending ONE level into subfolders (students
    often keep screenshots in an "evidencias/" subfolder). Anything deeper
    gets a visible note — never a silent skip; a silently ignored folder is
    indistinguishable from "the student didn't submit", the pilot's worst
    failure class.

    Deliberately tests "not a directory" rather than os.path.isfile(): on a
    Google Drive Desktop mount isfile() can report False for a real file, and
    a listdir entry that is not a directory is a file for our purposes.
    """
    out = []
    for name in listdir(folder):
        if name.lower() in SKIP_NAMES:
            continue
        full = os.path.join(folder, name)
        if isdir(full):
            if depth > 0:
                sub = collect_files(full, notes=notes, depth=depth - 1)
                for fr in sub:
                    fr["name"] = f"{name}/{fr['name']}"
                out.extend(sub)
            elif notes is not None:
                notes.append(f"subcarpeta no explorada (muy profunda): {name} "
                             "— revisar manualmente")
            continue
        size, size_err = getsize(full)
        rec = {"name": name, "ext": os.path.splitext(name)[1].lower(),
               "path": full, "size": size}
        if size_err:
            rec["size_error"] = size_err
        out.append(rec)
    return out


def main():
    if len(sys.argv) != 3:
        print(__doc__, file=sys.stderr)
        sys.exit(1)
    base, outdir = sys.argv[1], sys.argv[2]
    if not isdir(base):
        print(f"ERROR: no es una carpeta: {base}", file=sys.stderr)
        sys.exit(2)
    os.makedirs(longpath(outdir), exist_ok=True)
    sub_dir = os.path.join(outdir, "sub-listings")
    os.makedirs(longpath(sub_dir), exist_ok=True)

    entries = listdir(base)
    subfolders = [e for e in entries if isdir(os.path.join(base, e))]

    main_listing = {"folder_id": sid("ROOT", base),
                    "folder_name": os.path.basename(base.rstrip("\\/")),
                    "entries": []}
    manifest = {"folder_path": base, "shape": "", "students": []}

    if subfolders:
        manifest["shape"] = "nested"
        groups = {}
        for name in subfolders:
            m = CANVAS_RE.match(name)
            groups.setdefault(m.group(1) if m else name, []).append(name)

        for name in subfolders:
            folder = os.path.join(base, name)
            fid = sid("FOLDER", name)
            main_listing["entries"].append(
                {"id": fid, "name": name, "kind": "folder", "href": ""})
            folder_notes = []
            files = collect_files(folder, notes=folder_notes)
            sub_entries = []
            for fr in files:
                fr["id"] = sid("FILE", fid, fr["name"])
                sub_entries.append({
                    "id": fr["id"], "name": fr["name"],
                    "kind": "file" if fr["ext"] in REVIEWABLE_EXT else "other",
                    "href": ""})
            with wopen(os.path.join(sub_dir, f"{fid}.json")) as f:
                json.dump({"folder_id": fid, "folder_name": name,
                           "entries": sub_entries}, f,
                          ensure_ascii=False, indent=2)
            m = CANVAS_RE.match(name)
            manifest["students"].append({
                "folder_name": name, "folder_id": fid,
                "canvas_key": m.group(1) if m else None,
                "student_name": m.group(2) if m else name,
                "timestamp": m.group(3) if m else "",
                "collect_notes": folder_notes,
                "files": files})
        dup = {k: v for k, v in groups.items() if len(v) > 1}

        # F2 fix: loose files at the TOP level alongside student folders are
        # submissions too (a stray "Entrega-Juan.pptx" next to Canvas folders)
        # — each becomes its own flat-style entry, never silently dropped.
        loose = collect_files(base, notes=None, depth=0)
        for fr in loose:
            fr["id"] = sid("FILE", main_listing["folder_id"], fr["name"])
            main_listing["entries"].append({
                "id": fr["id"], "name": fr["name"],
                "kind": "file" if fr["ext"] in REVIEWABLE_EXT else "other",
                "href": ""})
            manifest["students"].append({
                "folder_name": fr["name"],
                "folder_id": sid("SOLO", fr["name"]),
                "canvas_key": None,
                "student_name": os.path.splitext(fr["name"])[0],
                "timestamp": "", "collect_notes": [],
                "files": [fr]})
            print(f"NOTE archivo suelto en el nivel superior (junto a las "
                  f"carpetas de estudiantes): {fr['name']} — se trata como "
                  "entrega propia; confirmar de quién es", file=sys.stderr)
    else:
        manifest["shape"] = "flat"
        files = collect_files(base)
        fid = main_listing["folder_id"]
        for fr in files:
            fr["id"] = sid("FILE", fid, fr["name"])
            main_listing["entries"].append({
                "id": fr["id"], "name": fr["name"],
                "kind": "file" if fr["ext"] in REVIEWABLE_EXT else "other",
                "href": ""})
            manifest["students"].append({
                "folder_name": fr["name"], "folder_id": sid("SOLO", fr["name"]),
                "canvas_key": None,
                "student_name": os.path.splitext(fr["name"])[0],
                "timestamp": "", "files": [fr]})
        dup = {}

    if not manifest["students"]:
        print("ERROR: no se encontraron entregas.", file=sys.stderr)
        sys.exit(3)

    by_folder = {s["folder_name"]: s for s in manifest["students"]}
    superseded = {}
    ts_warnings = set()
    for _key, names in dup.items():
        parsed = {n: parse_ts(by_folder[n]["timestamp"]) for n in names}
        bad = [n for n, t in parsed.items() if t is None]
        if bad:
            # cannot prove which submission is final — order best-effort but
            # SAY SO loudly and on every affected plan entry
            for n in names:
                ts_warnings.add(n)
            print("NOTE grupo duplicado con timestamp no interpretable "
                  f"({bad}) — verificar manualmente cuál entrega es la final",
                  file=sys.stderr)
        ordered = sorted(names, key=lambda n: parsed[n] or (0, 0, 0, 0, 0))
        for n in ordered[:-1]:
            superseded[n] = ordered[-1]

    plan = {"folders": []}
    for s in manifest["students"]:
        name = s["folder_name"]
        files = s["files"]
        e = {"folder_name": name, "folder_id": s["folder_id"],
             "canvas_key": s["canvas_key"], "student_name": s["student_name"],
             "timestamp": s["timestamp"],
             "notes": list(s.get("collect_notes") or [])}

        if name in ts_warnings:
            e["notes"].append("ORDEN DE ENTREGAS INCIERTO: timestamp no "
                              "interpretable en el grupo duplicado — verificar "
                              "manualmente cuál entrega es la final")
        if name in superseded:
            auth = superseded[name]
            e.update(status="no_deck", no_deck_files=files,
                     superseded_by=auth,
                     superseded_by_id=by_folder[auth]["folder_id"],
                     no_deck_reason=("entrega duplicada del mismo estudiante — "
                                     "reemplazada por un envío posterior en la "
                                     f"carpeta '{auth}'."))
            plan["folders"].append(e)
            continue

        if not files:
            e.update(status="no_deck", no_deck_files=[],
                     no_deck_reason="carpeta sin archivos entregados.")
            plan["folders"].append(e)
            continue

        reviewable = [fr for fr in files if fr["ext"] in REVIEWABLE_EXT]

        # R15: WITHIN-folder versioned duplicates (v1/v2/FINAL/(2)...) of the
        # SAME document — grade the most recent version, mark the earlier
        # ones REEMPLAZADA (never "evidence", never silently dropped).
        superseded_files = []
        by_base = {}
        for fr in reviewable:
            stem = os.path.splitext(fr["name"])[0].strip()
            by_base.setdefault(strip_versions(stem), []).append(fr)
        surviving = []
        for _vbase, group in by_base.items():
            # group by FULL stem: an export pair (Deck.pdf + Deck.pptx) shares
            # one stem and is NEVER a version relation — it survives together.
            stems = {}
            for fr in group:
                stems.setdefault(os.path.splitext(fr["name"])[0].strip(),
                                 []).append(fr)
            ranks = {st: version_rank(st) for st in stems}
            marked_stems = [st for st, vr in ranks.items() if vr is not None]
            unmarked = [st for st, vr in ranks.items() if vr is None]
            has_final_kw = any(vr and vr[0] == 2 for vr in ranks.values())
            # unambiguous version relation: everything is marked, OR the
            # marked side includes a final-class keyword that outranks the
            # plain name. Unmarked + numbered-only (Deck + Deck v1) is
            # AMBIGUOUS — the plain file is often the real final export;
            # never supersede on ambiguity, a human decides.
            unambiguous = (len(stems) >= 2 and marked_stems
                           and (not unmarked or has_final_kw))
            if len(stems) >= 2 and marked_stems and not unambiguous:
                e["notes"].append(
                    "POSIBLES VERSIONES EN LA MISMA CARPETA (ambiguo: archivo "
                    f"sin marcador junto a versión numerada: {sorted(stems)}) "
                    "— NO se descartó ninguna; confirmar cuál es la final.")
            if unambiguous:
                def skey(item):
                    st, frs = item
                    vr = version_rank(st)
                    return (vr or (-1, -1),
                            max(fr.get("size") or 0 for fr in frs))
                ordered_stems = sorted(stems.items(), key=skey)
                win_stem, win_files = ordered_stems[-1]
                surviving.extend(win_files)
                n_old = 0
                for st, frs in ordered_stems[:-1]:
                    for fr in frs:
                        fr = dict(fr)
                        fr["superseded_by_file"] = win_files[0]["name"]
                        superseded_files.append(fr)
                        n_old += 1
                e["notes"].append(
                    "VERSIONES EN LA MISMA CARPETA: se calificó "
                    f"'{win_stem}' (marcador de versión más reciente); "
                    f"{n_old} archivo(s) de versiones anteriores quedan como "
                    "REEMPLAZADA — confirmar con un humano.")
            else:
                surviving.extend(group)
        reviewable = surviving

        by_stem = {}
        for fr in reviewable:
            by_stem.setdefault(os.path.splitext(fr["name"])[0].strip().lower(),
                               []).append(fr)

        deck = None
        source_of = []
        leftovers = []
        for _stem, group in by_stem.items():
            pdf = next((g for g in group if g["ext"] == ".pdf"), None)
            if len(group) > 1 and pdf and deck is None:
                deck, source_of = pdf, [g for g in group if g is not pdf]
            else:
                leftovers.extend(group)

        if deck is None and leftovers:
            deck = leftovers.pop(0)
        elif deck is not None:
            leftovers = [fr for fr in leftovers if fr is not deck]

        if deck is None:
            # No slides/PDF at all. This is NOT "unreviewable": prepare_materials
            # renders docx/html/images/video into reviewable form. Pick the
            # richest file as the primary and let the reviewer judge content.
            e.update(status="no_deck", no_deck_files=files,
                     no_deck_reason=("entrega sin diapositivas (formatos: "
                                     f"{sorted({f['ext'] for f in files})}) — "
                                     "se revisa con la ruta multiformato."))
            plan["folders"].append(e)
            continue

        if leftovers:
            e["notes"].append(
                "MULTIPLES ARCHIVOS REVISABLES: se eligió "
                f"'{deck['name']}' como entrega principal por heurística de "
                "nombre — confirmar con un humano.")
        e.update(status="review", deck=deck, deck_source_of=source_of,
                 superseded_files=superseded_files,
                 evidence=[fr for fr in files if fr["ext"] not in REVIEWABLE_EXT
                           and fr["name"] not in
                           {sf["name"] for sf in superseded_files}]
                          + leftovers)
        plan["folders"].append(e)

    # R15: cross-folder duplicate FLAGGING by normalized student name — only
    # a flag, NEVER auto-supersede (distinct students can share a name; a
    # name-based auto-discard would silently destroy one student's work).
    by_person = {}
    for e in plan["folders"]:
        if e.get("canvas_key"):
            continue          # canvas-keyed folders already grouped reliably
        by_person.setdefault(norm_person(e["student_name"]), []).append(e)
    for person, entries in by_person.items():
        if person and len(entries) > 1:
            names = [x["folder_name"] for x in entries]
            for x in entries:
                x["notes"].append(
                    "POSIBLE ENTREGA DUPLICADA (mismo nombre de estudiante en "
                    f"varias carpetas sin clave Canvas: {names}) — verificar "
                    "manualmente si es la misma persona; NO se descartó "
                    "ninguna automáticamente.")
            print(f"NOTE posible duplicado por nombre: {person} -> {names}",
                  file=sys.stderr)

    for fname, obj in (("listing.json", main_listing),
                       ("manifest.json", manifest),
                       ("review_plan.json", plan)):
        with wopen(os.path.join(outdir, fname)) as f:
            json.dump(obj, f, ensure_ascii=False, indent=2)

    n_files = sum(len(s["files"]) for s in manifest["students"])
    n_rev = sum(1 for e in plan["folders"] if e["status"] == "review")
    n_no = sum(1 for e in plan["folders"] if e["status"] == "no_deck")
    print(json.dumps({"shape": manifest["shape"],
                      "folders": len(manifest["students"]),
                      "files": n_files, "with_deck": n_rev,
                      "without_deck": n_no,
                      "duplicate_groups": len(dup)}, ensure_ascii=False))
    for e in plan["folders"]:
        for n in e["notes"]:
            print(f"NOTE {e['student_name']}: {n}", file=sys.stderr)


if __name__ == "__main__":
    main()
