#!/usr/bin/env python3
"""Fairness gate + field-integrity validator for ONE reviewer result JSON.

Harness-agnostic: whatever agent framework produced the review (Claude Code
subagent, Codex spawn_agent child, Antigravity invoke_subagent, or a
sequential main-loop review), run its JSON through this script. It is the
single source of truth for "is this review acceptable?" — orchestrators must
not re-implement these checks ad hoc, and assemble_results.py imports
`normalize_review` from here so assembly applies the exact same rules.

Usage:
    python validate_review.py <review.json>
        [--expect-pages=N]              # deck review: pages_read must equal N
        [--expect-materials=lbl1|lbl2]  # bundle review: labels that must all
                                        # appear in materials_reviewed ('|'-sep)
        [--normalized-out=<path>]       # write the cleaned/normalized review

Checks (each one occurred in real output in the 2026-08-14 pilot run):
  * scores in [1.0, 5.0]; justifications non-empty.
  * pages_read == expected pages / materials_reviewed covers every label.
  * confident (alta/media) verification requires numeric age_months; "baja"
    requires age_months null.
  * DATE FIELDS must be "YYYY-MM-DD", "YYYY-MM", or "" — reviewers otherwise
    stuff whole sentences into them and the Excel gets prose in date columns
    (5 rows in the pilot). A parseable date buried in prose is EXTRACTED
    (prose moved to observations); no date at all -> "" + problem.
  * student must be a bare name, optionally "(código NNNN)" — commentary is
    stripped to observations (8 rows in the pilot).
  * tool capped at 70 chars (detail moved to observations; 45 rows).
  * FLAGS are normalized to the canonical vocabulary below; free-form flags
    move into `observations` (the pilot produced 104 distinct flags, 99 used
    once — unusable for TA filtering). Content is never lost, only
    relocated. Scores and justifications are NEVER edited.

Canonical flags (the only ones that stay in `flags`):
  VERIFICAR FECHA · DISCREPANCIA FECHA · ENTREGA SIN PPT · ENTREGA DUPLICADA
  REVISAR MANUALMENTE · SIN EVIDENCIA PROPIA · IMPACTO NO CUANTIFICADO
  HERRAMIENTA GENERAL - FUNCION ESPECIFICA · SPOT-CHECK FALLIDO
  EVIDENCIA NO LEGIBLE · EVIDENCIA DE ENVIO ANTERIOR INCLUIDA

CLI output: JSON to stdout {"ok": bool, "problems": [...],
"moved_to_observations": [...]}. Exit 0 = acceptable (possibly after
normalization) · 1 = usage · 2 = review must be retried/rejected.
"""
import json
import re
import sys

for _s in (sys.stdout, sys.stderr):
    if _s in (sys.__stdout__, sys.__stderr__) and hasattr(_s, "reconfigure"):
        _s.reconfigure(encoding="utf-8")

CANONICAL_FLAGS = (
    "VERIFICAR FECHA", "DISCREPANCIA FECHA", "ENTREGA SIN PPT",
    "ENTREGA DUPLICADA", "REVISAR MANUALMENTE", "SIN EVIDENCIA PROPIA",
    "IMPACTO NO CUANTIFICADO", "HERRAMIENTA GENERAL - FUNCION ESPECIFICA",
    "SPOT-CHECK FALLIDO", "EVIDENCIA NO LEGIBLE",
    "EVIDENCIA DE ENVIO ANTERIOR INCLUIDA",
)
DATE_OK = re.compile(r"^(\d{4}-\d{2}(-\d{2})?)?$")
DATE_FIND = re.compile(r"\d{4}-\d{2}(-\d{2})?")
STUDENT_OK = re.compile(r"^[^()]{0,60}(\(c[oó]digo [\w]+\))?$", re.IGNORECASE)


def normalize_review(r, expect_pages=None, expect_materials=None):
    """Validate + normalize IN PLACE. Returns (problems, moved) lists.

    problems -> the review must be retried/rejected (hard failures; nothing
    is auto-fixed for these). moved -> fields whose out-of-format content was
    relocated to `observations` (soft fixes; content preserved).
    """
    problems = []
    moved = []
    obs = [r.get("observations")] if r.get("observations") else []

    # --- scores & justifications (hard failures, never auto-fixed) ---------
    s = r.get("scores") or {}
    for k in ("poc", "impacto", "comunicacion"):
        v = s.get(k)
        if not isinstance(v, (int, float)) or not 1.0 <= v <= 5.0:
            problems.append(f"scores.{k} inválido: {v!r}")
    j = r.get("justification") or {}
    for k in ("poc", "impacto", "comunicacion"):
        if not (j.get(k) or "").strip():
            problems.append(f"justification.{k} vacía")

    # --- coverage ----------------------------------------------------------
    if expect_pages is not None and r.get("pages_read") != expect_pages:
        problems.append(f"pages_read ({r.get('pages_read')}) != páginas "
                        f"totales ({expect_pages})")
    if expect_materials is not None:
        seen = set(r.get("materials_reviewed") or [])
        missing = [m for m in expect_materials if m not in seen]
        if missing:
            problems.append(f"materials_reviewed no cubre: {missing}")

    # --- verification coherence -------------------------------------------
    conf = r.get("verification_confidence")
    age = r.get("age_months")
    if conf in ("alta", "media") and not isinstance(age, (int, float)):
        problems.append("confianza alta/media pero age_months no es numérico "
                        "(desactiva silenciosamente el filtro DQ)")
    if conf == "baja" and age is not None:
        problems.append("confianza baja pero age_months no es null "
                        "(un número no verificado presentado como dato)")

    # --- date fields: extract-or-empty, prose to observations --------------
    for f_ in ("declared_launch_date", "verified_launch_date"):
        v = (r.get(f_) or "").strip()
        if DATE_OK.match(v):
            continue
        m = DATE_FIND.search(v)
        obs.append(f"{f_} original del revisor: {v}")
        moved.append(f_)
        r[f_] = m.group(0) if m else ""
        if not m:
            problems.append(f"{f_} sin fecha reconocible: {v[:60]!r}")

    # --- student field: name (+ código) only -------------------------------
    st = (r.get("student") or "").strip()
    if st and not STUDENT_OK.match(st):
        obs.append(f"campo student original del revisor: {st}")
        moved.append("student")
        m = re.match(r"^([^()]+)", st)
        r["student"] = (m.group(1).strip() if m else "")[:60]

    # --- tool length --------------------------------------------------------
    tool = (r.get("tool") or "").strip()
    if len(tool) > 70:
        obs.append(f"descripción completa de la herramienta: {tool}")
        moved.append("tool")
        r["tool"] = tool[:67] + "..."

    # --- flag vocabulary ----------------------------------------------------
    kept = []
    for f_ in r.get("flags") or []:
        f_ = str(f_).strip()
        base = f_.split(":")[0].strip().upper()
        canon = next((c for c in CANONICAL_FLAGS
                      if base == c or base.startswith(c)), None)
        if canon:
            if canon not in kept:
                kept.append(canon)
            if f_.upper() != canon:          # keep the reviewer's detail
                obs.append(f"flag detallado: {f_}")
        else:
            obs.append(f"observación del revisor: {f_}")
            moved.append(f"flag:{f_[:40]}")
    r["flags"] = kept
    if obs:
        r["observations"] = " | ".join(x for x in obs if x)

    return problems, moved


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    if len(args) != 1:
        print(__doc__, file=sys.stderr)
        sys.exit(1)
    expect_pages = None
    expect_materials = None
    out_path = None
    for a in sys.argv[1:]:
        if a.startswith("--expect-pages="):
            expect_pages = int(a.split("=", 1)[1])
        elif a.startswith("--expect-materials="):
            expect_materials = [m for m in a.split("=", 1)[1].split("|") if m]
        elif a.startswith("--normalized-out="):
            out_path = a.split("=", 1)[1]

    try:
        with open(args[0], encoding="utf-8") as f:
            r = json.load(f)
    except (OSError, json.JSONDecodeError) as e:
        print(json.dumps({"ok": False, "problems": [f"JSON ilegible: {e}"],
                          "moved_to_observations": []}, ensure_ascii=False))
        sys.exit(2)

    problems, moved = normalize_review(r, expect_pages, expect_materials)
    if out_path:
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(r, f, ensure_ascii=False, indent=2)
    print(json.dumps({"ok": not problems, "problems": problems,
                      "moved_to_observations": moved}, ensure_ascii=False))
    sys.exit(0 if not problems else 2)


if __name__ == "__main__":
    main()
