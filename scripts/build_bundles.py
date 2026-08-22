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
    python build_bundles.py <workdir> [--images] [--all]

--images: emit page-image instructions instead of PDF instructions for any
material that carries a `page_images` list (populated by
`prepare_materials.py --rasterize`) — for harnesses whose file reader cannot
render PDF pages.
--all: emit a bundle for EVERY student, including plain single-deck ones —
required on non-PDF-vision harnesses (Codex, Antigravity), where the simple
deck template cannot be used because its {{pdf_path}} would hand the
reviewer a PDF it cannot read. Use together with --images.

Exits 2 when any student the plan requires a bundle for has no materials, or
a duplicate submission's authoritative folder cannot be resolved — both mean
a student would silently lose their review; fix the inputs, never proceed.
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
        if it.get("truncated_from") and use_images:
            # only meaningful in image mode: in PDF mode the reviewer holds
            # the COMPLETE file, nothing is truncated for them (round-3)
            avail = len(it.get("page_images") or [])
            suffix += (f"\n   AVISO: material truncado — solo {avail} de "
                       f"{it['truncated_from']} páginas disponibles; el resto "
                       "NO fue revisado (flag REVISAR MANUALMENTE).")
        if k == "pdf" and use_images and it.get("page_images"):
            imgs = "\n".join(f"     - `{p}`" for p in it["page_images"])
            lines.append(
                f"{n}. [PÁGINAS COMO IMÁGENES] «{label}» — "
                f"{len(it['page_images'])} páginas rasterizadas.\n"
                f"   Abre y MIRA cada imagen (cada una es una página):\n{imgs}"
                + suffix)
        elif k == "pdf":
            pages = it.get("pages")
            if use_images:
                # --images with no page_images = the operator skipped
                # --rasterize: on this harness the reviewer CANNOT read a
                # PDF — degrade LOUDLY, never silently
                print(f"AVISO: --images activo pero «{label}» no tiene "
                      "page_images (¿faltó prepare_materials --rasterize?) — "
                      "el revisor probablemente NO pueda leer este PDF.",
                      file=sys.stderr)
                suffix += ("\n   AVISO: este PDF NO fue rasterizado; si tu "
                           "herramienta no puede leer PDFs por páginas, "
                           "repórtalo con el flag EVIDENCIA NO LEGIBLE en "
                           "lugar de adivinar.")
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
        elif k == "audio":
            tp = it.get("transcript_path")
            dur = it.get("duration_sec")
            size = it.get("size_bytes")
            dtxt = (f" — duración {dur:.0f}s aprox" if dur
                    else " — duración desconocida")
            # 0 bytes is KNOWN-broken, not unknown: say so instead of
            # rendering it the same as a size we could not read.
            if size is None:
                stxt = ", tamaño desconocido.\n"
            elif size == 0:
                stxt = (", ARCHIVO VACÍO (0 bytes): la subida falló, no hay "
                        "nada que evaluar en él — repórtalo con el flag "
                        "EVIDENCIA NO LEGIBLE.\n")
            else:
                stxt = f", {size / 1e6:.1f} MB.\n"
            head = f"{n}. [AUDIO] «{label}»{dtxt}" + stxt
            if size == 0:
                # nothing to read and nothing to credit: a 0-byte upload is
                # broken, so the two-readings choice below must not apply
                lines.append(
                    head
                    + "   El archivo está VACÍO: no contiene audio de ningún "
                      "tipo, así que NO puede sustentar la PoC ni contar "
                      "como evidencia propia (tampoco si el resto de la "
                      "entrega afirma que es la salida de la herramienta). "
                      "Repórtalo con EVIDENCIA NO LEGIBLE y dilo en "
                      "evidence_notes."
                    + suffix)
            elif tp:
                lines.append(
                    head
                    + "   No puedes escuchar el audio, pero SÍ tienes su "
                      f"transcripción COMPLETA: `{tp}`\n"
                      "   Léela entera y evalúa con ella. Si la entrega es "
                      "hablada, la ficha técnica y la PoV están en lo que "
                      "DICE, no en diapositivas: no la castigues por no ser "
                      "un .pptx.\n"
                      "   Si la transcripción viene vacía o sin sentido, el "
                      "audio probablemente NO es habla (ver abajo)."
                    + suffix)
            else:
                lines.append(
                    head
                    + "   NO hay transcripción y no puedes escucharlo. Dos "
                      "lecturas, y debes elegir la que la evidencia sostenga:\n"
                      "   a) Si el resto de la entrega indica que este archivo "
                      "es la SALIDA de la herramienta evaluada (música, voz "
                      "o audio GENERADO por ella), entonces el archivo ES "
                      "evidencia propia: su existencia, formato y duración "
                      "cuentan a favor de la PoC. Dilo así en la "
                      "justificación, citando dónde lo sustenta el resto del "
                      "material.\n"
                      "   b) Si parece una presentación HABLADA (el estudiante "
                      "narra su trabajo), NO puedes calificar su contenido: "
                      "marca el flag EVIDENCIA NO LEGIBLE y dilo en "
                      "evidence_notes.\n"
                      "   En ningún caso inventes lo que se oye."
                    + suffix)
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
    include_all = "--all" in sys.argv[1:]

    plan = load(work, "review_plan.json")
    materials = load(work, "materials.json")
    conv = load(work, "convert_results.json", required=False) or {}

    by_fid = {e["folder_id"]: e for e in plan["folders"]}

    # superseded folder -> authoritative folder (same canvas key).
    # Primary source: the MACHINE field superseded_by_id emitted by
    # local_list.py. Prose parsing is only a legacy fallback and warns loudly:
    # a folder name containing an apostrophe breaks a quote-split silently,
    # and a silent miss here means carried-forward evidence never reaches the
    # reviewer AND the superseded folder gets a wasted duplicate review.
    superseded_of = {}
    unresolved_superseded = []
    for e in plan["folders"]:
        if e["status"] != "no_deck" or "duplicada" not in e.get("no_deck_reason", ""):
            continue
        auth_id = e.get("superseded_by_id")
        if not auth_id:
            target_name = e["no_deck_reason"].split("'")
            target = target_name[1] if len(target_name) >= 2 else None
            auth = next((x for x in plan["folders"]
                         if x["folder_name"] == target), None)
            auth_id = auth["folder_id"] if auth else None
            print(f"AVISO: plan sin campo superseded_by_id para "
                  f"'{e['folder_name']}' — usando análisis de texto "
                  f"({'resuelto' if auth_id else 'NO RESUELTO'})",
                  file=sys.stderr)
        if auth_id:
            superseded_of[e["folder_id"]] = auth_id
        else:
            unresolved_superseded.append(e["folder_name"])
            print(f"ERROR: entrega duplicada '{e['folder_name']}' sin carpeta "
                  "autoritativa resoluble — el arrastre de evidencia NO "
                  "ocurrirá y la carpeta podría revisarse dos veces. "
                  "Corrige review_plan.json antes de continuar.",
                  file=sys.stderr)

    def kinds_of(fid):
        m = materials.get(fid)
        return {it["kind"] for it in m["items"]} if m else set()

    bundles = []
    missing_materials = []
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

        should_have_materials = (
            (e["status"] == "no_deck" and e.get("no_deck_files"))
            or (e["status"] == "review" and e.get("evidence"))
            # --all makes EVERY review-status student bundle-dependent, so a
            # missing materials entry is just as blocking for them
            or (include_all and e["status"] == "review"))
        if should_have_materials and m is None:
            print(f"ERROR: review_plan requiere bundle para "
                  f"'{e['student_name']}' pero materials.json no contiene su "
                  f"carpeta ({fid}) — ¿corrió prepare_materials.py sobre "
                  "ella? Este estudiante NO tendrá revisión hasta corregirlo.",
                  file=sys.stderr)
            missing_materials.append(e["student_name"])

        needs_bundle = bool(
            (e["status"] == "no_deck" and items)          # (a) no deck at all
            or (e["status"] == "review" and e.get("evidence"))  # (b) deck+extras
            or carried                                     # (c) carry-forward
            or (include_all and e["status"] == "review"))  # (d) --all
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
        was_no_deck = e["status"] == "no_deck"
        # the {{no_deck_note}} text the bundle template requires — produced
        # HERE so every orchestrator gets identical fairness framing
        if was_no_deck:
            note = ("**IMPORTANTE — este estudiante NO entregó una "
                    "presentación de diapositivas.** Entregó el/los "
                    "material(es) listados abajo. Evalúalo con la MISMA "
                    "rúbrica y con justicia: el formato distinto NO es motivo "
                    "de castigo automático ni de descalificación. Juzga el "
                    "CONTENIDO. En 'comunicación' puedes considerar si el "
                    "material funciona como una presentación de 3 minutos, "
                    "pero no penalices el mero hecho de no ser un .pptx.")
        elif len(labels) > 1:
            note = ("**Este estudiante entregó una presentación MÁS material "
                    "complementario.** Debes revisar TODO antes de calificar: "
                    "la evidencia complementaria suele ser precisamente la "
                    "prueba propia que sustenta la PoC.")
        else:
            note = ""
        bundles.append({
            "folder_id": fid,
            "student_name": e["student_name"],
            "canvas_key": e.get("canvas_key"),
            "primary_label": labels[0] if labels else "",
            "materials_block": block,
            "material_labels": labels,
            "materials_expected": len(labels),
            "pdf_pages_total": sum(it.get("pages") or 0 for it in items
                                   if it["kind"] == "pdf"),
            # citable units = everything a spot-checker could ask the
            # reviewer to cite: PDF pages (only when the count is a REAL
            # number — unreadable is unknown, not zero) + each image + each
            # video frame + each text item. The 1-page spot-check exemption
            # keys on THIS, never on pdf pages alone (a png/video/docx
            # bundle has 0 pdf pages and plenty to cite).
            "citable_units": sum(
                (it.get("pages") if isinstance(it.get("pages"), int) else 0)
                if it["kind"] == "pdf"
                else len(it.get("frames") or []) if it["kind"] == "video"
                else 1 if it["kind"] in ("image", "text")
                # audio counts ONLY once transcribed: nobody can be asked to
                # cite a line from a file no reviewer could hear
                else 1 if (it["kind"] == "audio"
                           and it.get("transcript_path"))
                else 0
                for it in items),
            "was_no_deck": was_no_deck,
            "no_deck_note": note,
            "carried_forward": [
                {"from_folder_id": it.get("carried_from"),
                 "label": it.get("label")} for it in carried],
        })

    with open(os.path.join(work, "bundles.json"), "w", encoding="utf-8") as f:
        json.dump({"bundles": bundles}, f, ensure_ascii=False, indent=2)

    print(f"bundles: {len(bundles)}"
          + (f"  |  SIN MATERIALES (bloqueante): {len(missing_materials)}"
             if missing_materials else ""))
    if missing_materials or unresolved_superseded:
        # a student would silently lose their review — refuse to hand the
        # orchestrator a bundles.json that looks complete
        sys.exit(2)
    for b in bundles:
        cf = f" +{len(b['carried_forward'])} arrastrado(s)" if b["carried_forward"] else ""
        print(f"  {b['student_name'][:34]:34s} items={b['materials_expected']}"
              f" nodeck={b['was_no_deck']}{cf}")


if __name__ == "__main__":
    main()
