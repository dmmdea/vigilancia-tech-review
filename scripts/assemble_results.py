#!/usr/bin/env python3
"""Assemble the final results.json for make_excel.py from the review passes.

Inputs (in <workdir>):
  review_plan.json       — from local_list.py (or your listing step)
  bundles.json           — from build_bundles.py (defines which students used
                            the multi-format path + any carried-forward files)
  deck_reviews.json      — pass 1: {"reviewed": [{"folder_id", "result",
                            "problems": [...], "spotcheck": {...}|null}, ...]}
  bundle_reviews.json    — pass 2, same shape (optional)

Pass 2 SUPERSEDES pass 1 for any folder it covers (it saw strictly more).

Row model: every submitted file gets exactly one Excel row. Within a
student's folder, ONE row carries the grade; the rest get NO REVISADO rows
whose reason states they WERE read inside that student's integral review —
never silently dropped, never double-counted.

Every graded row is run through validate_review.normalize_review() — the
same gate the orchestrator applied per-review — so nothing out-of-format
reaches the Excel even if an orchestrator skipped the per-review call.
`observations` is folded into evidence_notes (make_excel renders that).

SAME-TOOL RECONCILIATION (cross-student consistency): after assembly, rows
are grouped by normalized tool name; if two students' verified dates for the
same tool differ by more than one month, BOTH rows get VERIFICAR FECHA plus
an explanatory note. Reviewers verify independently, so this is the only
place disagreement becomes visible.

Usage:
    python assemble_results.py <workdir> <run-date> [--folder-desc="..."]
"""
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from validate_review import normalize_review  # single source of truth

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


def month_index(datestr):
    m = re.match(r"^(\d{4})-(\d{2})", datestr or "")
    return int(m.group(1)) * 12 + int(m.group(2)) if m else None


def norm_tool(t):
    """First TWO significant words — deliberately coarse-but-conservative.

    Four words under-groups fatally: 'ChatGPT Work (modo agente...)' and
    'ChatGPT Work (agente de OpenAI...)' produce different keys and the
    reconciliation gate can never fire (caught by mutation test, 2026-08-15).
    Two words groups same-product rows while keeping distinct products apart
    ('google flow' vs 'google notebooklm'). Word-order variants of the same
    product still slip through — this check is best-effort, not a guarantee.
    """
    t = (t or "").lower()
    t = re.sub(r"[^a-z0-9ñáéíóú ]+", " ", t)
    # Spanish connectors and vendor prefixes carry no product identity —
    # without dropping them, 'Copilot para Excel' and 'Copilot para Word'
    # both key as 'copilot para' and distinct products get cross-flagged.
    stop = {"para", "con", "del", "los", "las", "una", "uno", "modo", "tipo",
            "microsoft", "google", "openai", "adobe", "365", "app",
            "plataforma", "función", "funcion", "herramienta"}
    words = [w for w in t.split() if len(w) > 2 and w not in stop][:2]
    return " ".join(words)


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    if len(args) != 2:
        print(__doc__, file=sys.stderr)
        sys.exit(1)
    work, run_date = args
    folder_desc = "(carpeta local de entregas)"
    for a in sys.argv[1:]:
        if a.startswith("--folder-desc="):
            folder_desc = a.split("=", 1)[1]

    plan = load(work, "review_plan.json")
    # optional: adversarial date re-verification verdicts (R19). Shape:
    # {"checks": [{"row_id", "verdict", "older_date", "older_evidence_url",
    #              "older_capability", "notes"}]}
    date_checks = (load(work, "date_checks.json", required=False)
                   or {}).get("checks", [])
    bundles = (load(work, "bundles.json", required=False) or {}).get("bundles", [])
    deck_pass = (load(work, "deck_reviews.json", required=False) or {}).get("reviewed", [])
    bundle_pass = (load(work, "bundle_reviews.json", required=False) or {}).get("reviewed", [])

    deck_by_fid = {i["folder_id"]: i for i in deck_pass}
    bundle_by_fid = {i["folder_id"]: i for i in bundle_pass}
    bundle_def = {b["folder_id"]: b for b in bundles}

    # a review whose folder_id matches nothing in the plan is a paid-for
    # review about to be silently discarded (LLM-mangled hash id, stale
    # inputs) — say so before it vanishes
    plan_fids = {e["folder_id"] for e in plan["folders"]}
    for label, ids in (("deck", deck_by_fid), ("bundle", bundle_by_fid)):
        for orphan in set(ids) - plan_fids:
            print(f"AVISO: revisión {label} con folder_id desconocido "
                  f"'{orphan}' — no corresponde a ninguna carpeta del plan y "
                  "será ignorada (¿id mal copiado por el revisor?).",
                  file=sys.stderr)

    # carried-forward: source folder -> (authoritative fid, labels)
    carried_from = {}
    for b in bundles:
        for cf in b.get("carried_forward", []):
            if cf.get("from_folder_id"):
                carried_from.setdefault(cf["from_folder_id"], []).append(
                    (b["folder_id"], cf.get("label")))

    results = []

    def no_rev(fr, reason, flags=None, student="", status="no_revisado",
               canvas_key=None):
        # every row carries the student's name — a NO REVISADO row whose
        # Estudiante column is blank makes the Excel unusable for TA triage.
        # status may also be:
        #   "revisado_anexo" — the file WAS read inside the student's
        #       integral review (Excel: REVISADO (ANEXO)); R16 class feedback:
        #       these previously showed as NO REVISADO and read as skipped.
        #   "reemplazada"    — an older version/duplicate superseded by a
        #       newer submission (Excel: REEMPLAZADA).
        results.append({"id": fr["id"], "file": fr["name"],
                        "student": student, "canvas_key": canvas_key,
                        "status": status, "status_reason": reason,
                        "disqualified": False, "flags": list(flags or [])})

    def graded(fr, r, extra_flags=None, extra_notes=None, student_name=None,
               canvas_key=None):
        row = dict(r)
        row["id"] = fr["id"]
        row["canvas_key"] = canvas_key
        row["file"] = fr["name"]              # always the ORIGINAL filename
        row["status"] = "revisado"
        problems, _moved = normalize_review(row)
        # AFTER normalization: the normalizer can blank a malformed student
        # field — a graded, top-5-eligible row must never lose its name.
        if not (row.get("student") or "").strip() and student_name:
            row["student"] = student_name     # Canvas folder always has it
        row.setdefault("flags", [])
        if problems:
            # should have been caught per-review; keep the grade but surface it
            if "REVISAR MANUALMENTE" not in row["flags"]:
                row["flags"].append("REVISAR MANUALMENTE")
            # a date problem specifically disarms/poisons the DQ filter —
            # mark it with the flag TAs filter on for date doubts
            if any("fecha" in p.lower() or "launch_date" in p for p in problems):
                if "VERIFICAR FECHA" not in row["flags"]:
                    row["flags"].append("VERIFICAR FECHA")
            extra_notes = ((extra_notes + " | ") if extra_notes else "") + \
                "Validación en ensamblaje: " + "; ".join(problems)
        for f in (extra_flags or []):
            if f not in row["flags"]:
                row["flags"].append(f)
        notes = row.get("evidence_notes", "") or ""
        obs = row.pop("observations", "")
        for extra in (extra_notes, ("OBS: " + obs) if obs else ""):
            if extra:
                notes = (notes + " | " if notes else "") + extra
        row["evidence_notes"] = notes
        results.append(row)

    def notes_annotations(e):
        """Map plan-level warnings onto the graded row: flags + note text.
        (Hunter finding B: stderr NOTEs never reached the Excel, so the
        deliverable asserted confidence the pipeline had disclaimed.)"""
        flags, texts = [], []
        for n in e.get("notes") or []:
            texts.append(n)
            up = n.upper()
            if "POSIBLE ENTREGA DUPLICADA" in up:
                flags.append("ENTREGA DUPLICADA")
            if ("ORDEN DE ENTREGAS INCIERTO" in up or "VERSIONES" in up
                    or "MULTIPLES ARCHIVOS" in up or "SUBCARPETA" in up
                    or "AMBIGUO" in up):
                flags.append("REVISAR MANUALMENTE")
        return flags, (" | ".join(texts) if texts else None)

    def spotcheck_annotations(item):
        """1-page submissions cannot cite 'slide N' — a citation-rule
        rejection there is not evidence of fabrication."""
        sc = item.get("spotcheck")
        if not sc:
            return [], None
        notes = sc.get("notes", "")
        if "plausible" not in sc:
            # an unrecognized verdict shape must read as UNKNOWN — branding a
            # student a suspected fabricator over a schema drift is the worst
            # possible false positive
            return [], ("Spot-check con formato de veredicto no reconocido — "
                        f"tratado como NO CONCLUYENTE. Detalle: {notes[:300]}")
        if sc.get("plausible"):
            return [], f"Spot-check de honestidad: OK. {notes[:400]}"
        pages_total = item.get("pages_total")
        if pages_total is None:
            pages_total = (item.get("result") or {}).get("pages_total")
        if pages_total == 1:
            return ([], "Spot-check: entrega de UNA página — no existe número "
                        "de slide que citar; el verificador evaluó fidelidad "
                        f"de contenido. Detalle: {notes[:400]}")
        return (["SPOT-CHECK FALLIDO", "REVISAR MANUALMENTE"],
                f"Spot-check de honestidad FALLIDO: {notes[:400]}")

    for e in plan["folders"]:
        fid = e["folder_id"]
        is_review = e["status"] == "review"
        files = ([e["deck"]] + e.get("deck_source_of", []) + e.get("evidence", [])
                 if is_review else list(e.get("no_deck_files", [])))
        superseded = (not is_review
                      and bool(e.get("superseded_by_id")
                               or "duplicada" in e.get("no_deck_reason", "")))

        if superseded:
            carried_labels = {lbl for _t, lbl in carried_from.get(fid, [])}
            for fr in files:
                if fr["name"] in carried_labels:
                    no_rev(fr, "evidencia de este envío anterior SÍ revisada "
                               "dentro de la revisión integral de la entrega "
                               "final del estudiante — la nota está en esa fila.",
                           ["EVIDENCIA DE ENVIO ANTERIOR INCLUIDA"],
                           student=e["student_name"], status="revisado_anexo")
                else:
                    no_rev(fr, e["no_deck_reason"], [],
                           student=e["student_name"], status="reemplazada")
            continue

        # R15: older versions superseded WITHIN the folder — visible rows,
        # graded row pointed at, never dropped (they are in the listings, so
        # the make_excel completeness gate requires them anyway)
        for fr in e.get("superseded_files", []):
            no_rev(fr, "versión anterior de "
                       f"'{fr.get('superseded_by_file', 'la entrega final')}' "
                       "en la misma carpeta — se calificó la versión más "
                       "reciente; la nota está en esa fila.",
                   student=e["student_name"], status="reemplazada")

        bundle = bundle_by_fid.get(fid)
        deck = deck_by_fid.get(fid)

        if bundle:                             # pass 2 wins when present
            bdef = bundle_def.get(fid, {})
            primary_name = bdef.get("primary_label", "")
            primary = next((fr for fr in files if fr["name"] == primary_name),
                           None)
            if primary is None and files:
                # e.g. a .zip submission whose bundle labels are inner files —
                # attach to the first plan file but SAY SO, never silently
                print(f"AVISO: primary_label '{primary_name}' no coincide con "
                      f"ningún archivo del plan para '{e['student_name']}' — "
                      f"la nota se ancla a '{files[0]['name']}'.",
                      file=sys.stderr)
                primary = files[0]
            if primary is None:
                # a paid-for review with no plan files to attach it to means
                # the inputs are inconsistent (stale bundles.json vs a
                # regenerated plan) — NEVER a silent drop
                print(f"ERROR: revisión de bundle para {fid} "
                      f"('{e['student_name']}') sin archivos en el plan — "
                      "entradas inconsistentes; fila sintética generada.",
                      file=sys.stderr)
                no_rev({"id": fid, "name": f"({e['student_name']} — "
                                            "inconsistencia de entradas)"},
                       "revisión existente pero sin archivos en review_plan "
                       "— regenerar bundles.json y re-ensamblar.",
                       ["REVISAR MANUALMENTE"], student=e["student_name"])
                continue
            if bundle.get("problems"):
                no_rev(primary, "revisión incompleta tras reintento — "
                       + "; ".join(bundle["problems"]), ["REVISAR MANUALMENTE"],
                       student=e["student_name"])
            else:
                r = dict(bundle["result"])
                reviewed_list = r.pop("materials_reviewed", [])
                flags = []
                note = (f"Revisión integral de "
                        f"{bdef.get('materials_expected', len(reviewed_list))} "
                        f"material(es): {', '.join(reviewed_list)[:400]}.")
                if bdef.get("was_no_deck"):
                    flags.append("ENTREGA SIN PPT")
                    note += (" Sin diapositivas; se evaluó el material "
                             "entregado con la misma rúbrica.")
                if bdef.get("carried_forward"):
                    flags.append("EVIDENCIA DE ENVIO ANTERIOR INCLUIDA")
                # 1-page exemption ONLY when the bundle genuinely has a
                # single citable unit (one 1-page PDF, one image). Keying on
                # pdf pages alone silently disarmed SPOT-CHECK FALLIDO for
                # every png/video/docx bundle (round-3 critical finding).
                bundle_sc = dict(bundle)
                cu = bdef.get("citable_units")
                if (bdef.get("materials_expected") == 1
                        and isinstance(cu, int) and cu <= 1):
                    bundle_sc["pages_total"] = 1
                sf, sn = spotcheck_annotations(bundle_sc)
                flags += sf
                if sn:
                    note += " | " + sn
                nf, nn = notes_annotations(e)
                flags += nf
                if nn:
                    note += " | AVISOS DEL INVENTARIO: " + nn
                graded(primary, r, flags, note, e["student_name"],
                       canvas_key=e.get("canvas_key") or "")
            # "SÍ fue leído" may only be asserted for files the bundle
            # actually LISTED — anything else is a false certification that
            # suppresses the human follow-up (finding 4, 2026-08-15 review).
            listed = set(bdef.get("material_labels") or [])
            # a .zip container whose EXTRACTED children were listed (labels
            # "container :: inner") counts as read — the contents are the
            # submission, the container is just packaging
            container_prefixes = {lbl.split(" :: ")[0] for lbl in listed
                                  if " :: " in lbl}
            listed |= container_prefixes
            review_ok = not bundle.get("problems")
            for fr in files:
                if fr is primary:
                    continue
                if fr["name"] in listed and review_ok:
                    no_rev(fr, "leído como anexo dentro de la revisión "
                               f"integral de '{primary['name']}' — la nota "
                               "del estudiante está en esa fila.",
                           student=e["student_name"], status="revisado_anexo")
                elif fr["name"] in listed:
                    no_rev(fr, "material listado en una revisión integral que "
                               "quedó incompleta — revisar manualmente junto "
                               f"con '{primary['name']}'.",
                           ["REVISAR MANUALMENTE"], student=e["student_name"])
                else:
                    no_rev(fr, "material adjunto NO incluido en la revisión "
                               "integral (no llegó al bundle) — revisar "
                               "manualmente.", ["REVISAR MANUALMENTE"],
                           student=e["student_name"])
            continue

        if deck:                               # pass 1 (deck-only)
            if deck.get("problems"):
                no_rev(e["deck"], "revisión incompleta tras reintento — "
                       + "; ".join(deck["problems"]), ["REVISAR MANUALMENTE"],
                       student=e["student_name"])
            else:
                flags, note = spotcheck_annotations(deck)
                nf, nn = notes_annotations(e)
                flags = list(flags) + nf
                if nn:
                    note = ((note + " | ") if note else "") +                         "AVISOS DEL INVENTARIO: " + nn
                graded(e["deck"], deck["result"], flags, note,
                       e["student_name"],
                       canvas_key=e.get("canvas_key") or "")
            for fr in e.get("deck_source_of", []):
                no_rev(fr, "archivo fuente (mismo contenido) de la "
                           f"presentación ya revisada '{e['deck']['name']}' — "
                           "la nota está en esa fila.",
                       student=e["student_name"], status="revisado_anexo")
            for fr in e.get("evidence", []):
                no_rev(fr, "material adjunto no incluido en la revisión — "
                           "revisar manualmente.", ["REVISAR MANUALMENTE"],
                       student=e["student_name"])
            continue

        if not files:
            # an empty folder must still be VISIBLE to the TAs — absence of a
            # row is indistinguishable from a pipeline drop (the pilot's
            # MAX_PATH bug looked exactly like this)
            no_rev({"id": fid, "name": "(carpeta sin archivos)"},
                   f"la carpeta de {e['student_name']} está vacía — verificar "
                   "en Canvas si el estudiante entregó.",
                   ["REVISAR MANUALMENTE"], student=e["student_name"])
            continue
        for fr in files:
            no_rev(fr, "no se pudo revisar (sin resultado de ninguna pasada).",
                   ["REVISAR MANUALMENTE"], student=e["student_name"])

    # ---- REEMPLAZADA cross-check: the replacement must actually be graded.
    # (Hunter finding F2: a row must never certify "la nota está en esa fila"
    # when the replacing submission produced no grade — e.g. the student's
    # newer Canvas folder was empty.)
    plan_by_fid = {e["folder_id"]: e for e in plan["folders"]}

    def folder_file_ids(e):
        ids = set()
        if e.get("deck"):
            ids.add(e["deck"]["id"])
        for key in ("deck_source_of", "evidence", "no_deck_files",
                    "superseded_files"):
            ids |= {fr["id"] for fr in e.get(key) or []}
        return ids

    graded_ids = {row["id"] for row in results if row["status"] == "revisado"}

    # in-folder version supersedes: a REEMPLAZADA row certifies that ITS
    # winner got graded — per FAMILY, not per folder: a folder can hold
    # several version families (Deck v1/final + Anexo v1/final) and only
    # the deck family's winner is ever graded, so "any graded row in the
    # folder" falsely certified the annex family (convergence finding 7).
    def _stem(n):
        return os.path.splitext(n)[0].strip().lower()

    for e in plan["folders"]:
        if not e.get("superseded_files"):
            continue
        ids_by_stem = {}
        all_files = ([e["deck"]] if e.get("deck") else [])
        for key in ("deck_source_of", "evidence", "no_deck_files",
                    "superseded_files"):
            all_files += e.get(key) or []
        for fr in all_files:
            ids_by_stem.setdefault(_stem(fr["name"]), set()).add(fr["id"])
        winner_of = {fr["id"]: fr.get("superseded_by_file")
                     for fr in e["superseded_files"]}
        for row in results:
            if row["status"] != "reemplazada" or row["id"] not in winner_of:
                continue
            winner = winner_of[row["id"]]
            winner_ids = ids_by_stem.get(_stem(winner), set()) if winner \
                else set()
            if winner_ids & graded_ids:
                continue
            if "REVISAR MANUALMENTE" not in row["flags"]:
                row["flags"].append("REVISAR MANUALMENTE")
            row["status_reason"] += (
                " | AVISO: la versión más reciente de ESTE documento NO "
                "quedó calificada en esta corrida — revisar ESTA versión "
                "manualmente.")

    for e in plan["folders"]:
        target = e.get("superseded_by_id")
        if not target:
            continue
        te = plan_by_fid.get(target)
        target_graded = bool(te and (folder_file_ids(te) & graded_ids))
        if not target_graded:
            own_ids = folder_file_ids(e)
            for row in results:
                if row["status"] == "reemplazada" and row["id"] in own_ids:
                    if "REVISAR MANUALMENTE" not in row["flags"]:
                        row["flags"].append("REVISAR MANUALMENTE")
                    row["status_reason"] += (
                        " | AVISO: la entrega que la reemplaza NO quedó "
                        "calificada en esta corrida — revisar ESTA versión "
                        "manualmente.")

    # ---- R19: apply adversarial date-check verdicts as FLAGS --------------
    # The checker can prove a capability is older than the accepted date; the
    # pipeline surfaces that loudly but never silently re-grades or DQs —
    # humans decide (same never-edit-verdicts rule as everywhere else).
    import unicodedata

    def _norm_verdict(v):
        v = unicodedata.normalize("NFKD", str(v or "").strip().lower())
        return "".join(ch for ch in v if not unicodedata.combining(ch))

    VALID_VERDICTS = {"confirmada", "mas_vieja", "no_concluyente"}
    row_ids = {r.get("id") for r in results}
    checks_by_id = {}
    for c in date_checks:
        rid = c.get("row_id")
        if not rid or rid not in row_ids:
            print(f"AVISO: date_check con row_id desconocido {rid!r} — "
                  "veredicto DESCARTADO; corrige el id y re-ensambla.",
                  file=sys.stderr)
            continue
        v = _norm_verdict(c.get("verdict"))
        if v not in VALID_VERDICTS:
            print(f"AVISO: date_check {rid} con verdict no reconocido "
                  f"{c.get('verdict')!r} — tratado como no_concluyente.",
                  file=sys.stderr)
            v = "no_concluyente"
        if v == "mas_vieja" and not (c.get("older_date")
                                     and c.get("older_evidence_url")):
            print(f"AVISO: date_check {rid} dice mas_vieja SIN older_date/"
                  "older_evidence_url — degradado a no_concluyente.",
                  file=sys.stderr)
            v = "no_concluyente"
        c = dict(c, verdict=v)
        prev = checks_by_id.get(rid)
        if prev:
            # duplicate: the WORST verdict survives (mas_vieja > no_concl > conf)
            rank = {"mas_vieja": 2, "no_concluyente": 1, "confirmada": 0}
            print(f"AVISO: date_check duplicado para {rid} — se conserva el "
                  "veredicto más severo.", file=sys.stderr)
            if rank[v] <= rank[prev["verdict"]]:
                continue
        checks_by_id[rid] = c
    n_older = 0
    applied_ids = set()   # count CHECKS applied, not rows — two rows sharing
    for row in results:   # an id must not double-count (convergence minor)
        c = checks_by_id.get(row.get("id"))
        if not c:
            continue
        applied_ids.add(row["id"])
        v = c.get("verdict")
        if v == "mas_vieja":
            n_older += 1
            for fl in ("VERIFICAR FECHA", "DISCREPANCIA FECHA",
                       "REVISAR MANUALMENTE"):
                if fl not in row["flags"]:
                    row["flags"].append(fl)
            extra = (f"VERIFICACIÓN ADVERSARIAL DE FECHA: la capacidad "
                     f"demostrada ya existía desde {c.get('older_date', '?')} "
                     f"({c.get('older_capability', '')}) — evidencia: "
                     f"{c.get('older_evidence_url', '')}. La fila puede estar "
                     "por FUERA de la ventana de 4 meses; decisión humana "
                     "requerida.")
            key = "evidence_notes" if row["status"] == "revisado" else "status_reason"
            row[key] = ((row.get(key) or "") + " | " + extra).strip(" |")
        elif v == "no_concluyente":
            if "VERIFICAR FECHA" not in row["flags"]:
                row["flags"].append("VERIFICAR FECHA")
            key = "evidence_notes" if row["status"] == "revisado" else "status_reason"
            row[key] = ((row.get(key) or "") + " | Verificación adversarial "
                        "de fecha NO concluyente: "
                        + (c.get("notes") or "")).strip(" |")
    if date_checks:
        print(f"date_checks: {len(date_checks)} recibidos, {len(applied_ids)} "
              f"aplicados ({n_older} con evidencia de fecha más vieja)")

    # ---- same-tool cross-student reconciliation ---------------------------
    groups = {}
    for row in results:
        if row["status"] != "revisado":
            continue
        key = norm_tool(row.get("tool"))
        if key:
            groups.setdefault(key, []).append(row)
    n_flagged = 0
    for key, rows in groups.items():
        dated = [(row, month_index(row.get("verified_launch_date")))
                 for row in rows]
        dated = [(row, mi) for row, mi in dated if mi is not None]
        if len(dated) < 2:
            continue
        months = [mi for _row, mi in dated]
        if max(months) - min(months) > 1:
            for row, _mi in dated:
                if "VERIFICAR FECHA" not in row["flags"]:
                    row["flags"].append("VERIFICAR FECHA")
                    n_flagged += 1
                notes = row.get("evidence_notes", "") or ""
                row["evidence_notes"] = (notes + " | " if notes else "") + (
                    "Reconciliación entre estudiantes: otras entregas sobre "
                    "la misma herramienta reportan fechas verificadas que "
                    "difieren en más de un mes — puede ser una función "
                    "distinta o un error de verificación; confirmar.")

    out = {"run_date": run_date, "folder_url": folder_desc, "results": results}
    with open(os.path.join(work, "results.json"), "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)

    n_rev = sum(1 for r in results if r["status"] == "revisado")
    n_dq = sum(1 for r in results
               if r["status"] == "revisado" and r.get("disqualified"))
    n_no = sum(1 for r in results if r["status"] == "no_revisado")
    print(f"results.json: {len(results)} filas | revisado={n_rev} "
          f"(DQ={n_dq}) | no_revisado={n_no} | "
          f"reconciliación agregó VERIFICAR FECHA a {n_flagged} fila(s)")


if __name__ == "__main__":
    main()
