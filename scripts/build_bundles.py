#!/usr/bin/env python3
"""Build bundles.json: one multi-format review target per student who needs
one — plus the generic resubmission carry-forward rule.

A student needs the bundle (multi-format) review path when ANY of:
  a) they submitted no slide deck at all (docx / png / html / video / ...),
  b) they submitted a deck PLUS supplementary evidence files (the evidence is
     usually the proof behind the 50% PoC score — a deck-only review silently
     under-grades them),
  c) their AUTHORITATIVE (latest) submission is accompanied by an earlier,
     superseded submission that contains evidence KINDS the final one lacks
     (e.g. the first upload carried the PoV video, the resubmission only the
     PDF). Grading the latest while ignoring evidence the student clearly
     produced is unfair — the earlier evidence rides along, flagged.

Reads:  <workdir>/review_plan.json  (from local_list.py / your listing step)
        <workdir>/materials.json    (from prepare_materials.py)
        <workdir>/convert_results.json  (optional; pass-1 deck conversions,
                                          {folder_id: {pdf, pages, ...}})
Writes: <workdir>/bundles.json — for each target: student identity, a
        `materials_block` (explicit per-item open-and-read instructions), the
        expected material labels, and any carried-forward evidence records.

Usage:
    python build_bundles.py <workdir> [--images]

--images: emit page-image instructions instead of PDF instructions for any
material that carries a `page_images` list (produced by pdf_to_images.py) —
for harnesses whose file reader cannot render PDF pages.
"""
import json
import os
import sys

for _s in (sys.stdout, sys.stderr):
    if _s in (sys.__stdout__, sys.__stderr__) and hasattr(_s, "reconfigure"):
        _s.reconfigure(encoding="utf-8")


def load(work, name, required=True):
    p = os.path.join(work, name)
    if not os.path.exists(p):
        if required:
            print(f"ERROR: falta {p}", file=sys.stderr)
            sys.exit(2)
        return None
    with open(p, encoding="utf-8") as f:
        return json.load(f)


def describe(items, use_images):
    """Render explicit per-item reading instructions. Nothing implicit."""
    lines = []
    labels = []
    for n, it in enumerate(items, 1):
        k = it["kind"]
        label = it.get("label", "(sin nombre)")
        carried = it.get("carried_forward_note")
        suffix = f"\n   Nota: {it['note']}" if it.get("note") else ""
        if carried:
            suffix += f"\n   Nota: {carried}"
        if k == "pdf" and use_images and it.get("page_images"):
            imgs = "\n".join(f"     - `{p}`" for p in it["page_images"])
            lines.append(
                f"{n}. [PÁGINAS COMO IMÁGENES] «{label}» — "
                f"{len(it['page_images'])} páginas rasterizadas.\n"
                f"   Abre y MIRA cada imagen (cada una es una página):\n{imgs}"
                + suffix)
        elif k == "pdf":
            pages = it.get("pages")
            lines.append(
                f"{n}. [PDF] «{label}» — {pages} páginas.\n"
                f"   Ábrelo con tu herramienta de lectura en `{it['path']}` "
                f"y LEE LAS {pages} PÁGINAS (visualmente, por páginas)."
                + suffix)
        elif k == "image":
            lines.append(
                f"{n}. [IMAGEN] «{label}»\n"
                f"   Ábrela en `{it['path']}` y MÍRALA — evalúala como "
                f"evidencia (o como la entrega completa si es lo único)."
                + suffix)
        elif k == "text":
            lines.append(
                f"{n}. [TEXTO/DATOS] «{label}»\n"
                f"   Léelo completo en `{it['path']}`." + suffix)
        elif k == "video":
            frames = it.get("frames", [])
            dur = it.get("duration_sec")
            tp = it.get("transcript_path")
            flist = "\n".join(f"     - `{p}`" for p in frames)
            dtxt = f" — duración {dur:.0f}s aprox" if dur else ""
            lines.append(
                f"{n}. [VIDEO] «{label}»{dtxt}.\n"
                f"   No puedes reproducir el video, pero SÍ tienes sus "
                f"fotogramas{' y su transcripción' if tp else ''}:\n"
                f"   a) Fotogramas (ábrelos TODOS; son imágenes):\n{flist}\n"
                + (f"   b) Transcripción del audio: `{tp}`\n" if tp else "")
                + "   Evalúa el video con esa evidencia combinada." + suffix)
        elif k == "error":
            lines.append(
                f"{n}. [NO LEGIBLE] «{label}» — {it.get('note', '')}\n"
                f"   Menciónalo en evidence_notes como material no legible.")
        labels.append(label)
    return "\n".join(lines), labels


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    if len(args) != 1:
        print(__doc__, file=sys.stderr)
        sys.exit(1)
    work = args[0]
    use_images = "--images" in sys.argv[1:]

    plan = load(work, "review_plan.json")
    materials = load(work, "materials.json")
    conv = load(work, "convert_results.json", required=False) or {}

    by_fid = {e["folder_id"]: e for e in plan["folders"]}

    # superseded folder -> authoritative folder (same canvas key)
    superseded_of = {}
    for e in plan["folders"]:
        if e["status"] == "no_deck" and "duplicada" in e.get("no_deck_reason", ""):
            target_name = e["no_deck_reason"].split("'")
            target = target_name[1] if len(target_name) >= 2 else None
            auth = next((x for x in plan["folders"]
                         if x["folder_name"] == target), None)
            if auth:
                superseded_of[e["folder_id"]] = auth["folder_id"]

    def kinds_of(fid):
        m = materials.get(fid)
        return {it["kind"] for it in m["items"]} if m else set()

    bundles = []
    for e in plan["folders"]:
        fid = e["folder_id"]
        if fid in superseded_of:            # superseded folders get no grade
            continue

        m = materials.get(fid)
        items = list(m["items"]) if m else []

        # (c) carry forward evidence kinds from a superseded twin
        carried = []
        for sup_fid, auth_fid in superseded_of.items():
            if auth_fid != fid:
                continue
            have = kinds_of(fid) or ({"pdf"} if fid in conv else set())
            sup_m = materials.get(sup_fid)
            if not sup_m:
                continue
            for it in sup_m["items"]:
                if it["kind"] in ("video", "image", "text") \
                        and it["kind"] not in have:
                    it = dict(it)
                    it["carried_forward_note"] = (
                        "evidencia del envío ANTERIOR del estudiante (entrega "
                        "duplicada); se incluye para no perjudicarlo — flag "
                        "EVIDENCIA DE ENVIO ANTERIOR INCLUIDA")
                    it["carried_from"] = sup_fid
                    carried.append(it)

        needs_bundle = bool(
            (e["status"] == "no_deck" and items)          # (a) no deck at all
            or (e["status"] == "review" and e.get("evidence"))  # (b) deck+extras
            or carried)                                    # (c) carry-forward
        if not needs_bundle:
            continue

        # ensure the deck itself is listed when materials.json lacks it
        if e["status"] == "review" and not any(i["kind"] == "pdf" for i in items):
            c = conv.get(fid)
            if c and c.get("ok", True) and c.get("pdf"):
                items.insert(0, {"kind": "pdf",
                                 "path": os.path.abspath(c["pdf"]),
                                 "label": e["deck"]["name"],
                                 "pages": c.get("pages")})
        items = items + carried
        if not items:
            continue

        block, labels = describe(items, use_images)
        bundles.append({
            "folder_id": fid,
            "student_name": e["student_name"],
            "canvas_key": e.get("canvas_key"),
            "primary_label": labels[0] if labels else "",
            "materials_block": block,
            "material_labels": labels,
            "materials_expected": len(labels),
            "was_no_deck": e["status"] == "no_deck",
            "carried_forward": [
                {"from_folder_id": it.get("carried_from"),
                 "label": it.get("label")} for it in carried],
        })

    with open(os.path.join(work, "bundles.json"), "w", encoding="utf-8") as f:
        json.dump({"bundles": bundles}, f, ensure_ascii=False, indent=2)

    print(f"bundles: {len(bundles)}")
    for b in bundles:
        cf = f" +{len(b['carried_forward'])} arrastrado(s)" if b["carried_forward"] else ""
        print(f"  {b['student_name'][:34]:34s} items={b['materials_expected']}"
              f" nodeck={b['was_no_deck']}{cf}")


if __name__ == "__main__":
    main()
