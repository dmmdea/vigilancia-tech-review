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

    def no_rev(fr, reason, flags=None, student="", status="no_revisado"):
        # every row carries the student's name — a NO REVISADO row whose
        # Estudiante column is blank makes the Excel unusable for TA triage.
        # status may also be:
        #   "revisado_anexo" — the file WAS read inside the student's
        #       integral review (Excel: REVISADO (ANEXO)); R16 class feedback:
        #       these previously showed as NO REVISADO and read as skipped.
        #   "reemplazada"    — an older version/duplicate superseded by a
        #       newer submission (Excel: REEMPLAZADA).
        results.append({"id": fr["id"], "file": fr["name"],
                        "student": student,
                        "status": status, "status_reason": reason,
                        "disqualified": False, "flags": list(flags or [])})

    def graded(fr, r, extra_flags=None, extra_notes=None, student_name=None):
        row = dict(r)
        row["id"] = fr["id"]
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
                      and "duplicada" in e.get("no_deck_reason", ""))

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
                graded(primary, r, flags, note, e["student_name"])
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
                graded(e["deck"], deck["result"], flags, note,
                       e["student_name"])
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
