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
    t = (t or "").lower()
    t = re.sub(r"[^a-z0-9ñáéíóú ]+", " ", t)
    words = [w for w in t.split() if len(w) > 2][:4]
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

    # carried-forward: source folder -> (authoritative fid, labels)
    carried_from = {}
    for b in bundles:
        for cf in b.get("carried_forward", []):
            if cf.get("from_folder_id"):
                carried_from.setdefault(cf["from_folder_id"], []).append(
                    (b["folder_id"], cf.get("label")))

    results = []

    def no_rev(fr, reason, flags=None):
        results.append({"id": fr["id"], "file": fr["name"],
                        "status": "no_revisado", "status_reason": reason,
                        "disqualified": False, "flags": list(flags or [])})

    def graded(fr, r, extra_flags=None, extra_notes=None, student_name=None):
        row = dict(r)
        row["id"] = fr["id"]
        row["file"] = fr["name"]              # always the ORIGINAL filename
        row["status"] = "revisado"
        if not (row.get("student") or "").strip() and student_name:
            row["student"] = student_name     # Canvas folder always has it
        problems, _moved = normalize_review(row)
        row.setdefault("flags", [])
        if problems:
            # should have been caught per-review; keep the grade but surface it
            if "REVISAR MANUALMENTE" not in row["flags"]:
                row["flags"].append("REVISAR MANUALMENTE")
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
        if sc.get("plausible"):
            return [], f"Spot-check de honestidad: OK. {notes[:400]}"
        if item.get("pages_total") == 1:
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
                           ["ENTREGA DUPLICADA",
                            "EVIDENCIA DE ENVIO ANTERIOR INCLUIDA"])
                else:
                    no_rev(fr, e["no_deck_reason"], ["ENTREGA DUPLICADA"])
            continue

        bundle = bundle_by_fid.get(fid)
        deck = deck_by_fid.get(fid)

        if bundle:                             # pass 2 wins when present
            bdef = bundle_def.get(fid, {})
            primary_name = bdef.get("primary_label", "")
            primary = next((fr for fr in files if fr["name"] == primary_name),
                           files[0] if files else None)
            if primary is None:
                continue
            if bundle.get("problems"):
                no_rev(primary, "revisión incompleta tras reintento — "
                       + "; ".join(bundle["problems"]), ["REVISAR MANUALMENTE"])
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
                sf, sn = spotcheck_annotations(bundle)
                flags += sf
                if sn:
                    note += " | " + sn
                graded(primary, r, flags, note, e["student_name"])
            for fr in files:
                if fr is primary:
                    continue
                no_rev(fr, "material adjunto del estudiante; SÍ fue leído "
                           "dentro de la revisión integral de "
                           f"'{primary['name']}' — la nota está en esa fila.")
            continue

        if deck:                               # pass 1 (deck-only)
            if deck.get("problems"):
                no_rev(e["deck"], "revisión incompleta tras reintento — "
                       + "; ".join(deck["problems"]), ["REVISAR MANUALMENTE"])
            else:
                flags, note = spotcheck_annotations(deck)
                graded(e["deck"], deck["result"], flags, note,
                       e["student_name"])
            for fr in e.get("deck_source_of", []):
                no_rev(fr, "archivo fuente de la MISMA presentación ya "
                           f"revisada ('{e['deck']['name']}', mismo nombre "
                           "base) — la nota está en esa fila.")
            for fr in e.get("evidence", []):
                no_rev(fr, "material adjunto no incluido en la revisión — "
                           "revisar manualmente.", ["REVISAR MANUALMENTE"])
            continue

        for fr in files:
            no_rev(fr, "no se pudo revisar (sin resultado de ninguna pasada).",
                   ["REVISAR MANUALMENTE"])

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
