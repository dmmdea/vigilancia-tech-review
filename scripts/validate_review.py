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
DATE_OK = re.compile(r"^(\d{4}-(0[1-9]|1[0-2])(-(0[1-9]|[12]\d|3[01]))?)?$")
DATE_FIND = re.compile(r"\d{4}-(?:0[1-9]|1[0-2])(?:-(?:0[1-9]|[12]\d|3[01]))?")
# Spanish long-form and dd/mm/yyyy dates are EXPECTED reviewer output (the
# skill's own language rule mandates Spanish) — they must normalize, not
# hard-fail the gate.
MESES = {"enero": "01", "febrero": "02", "marzo": "03", "abril": "04",
         "mayo": "05", "junio": "06", "julio": "07", "agosto": "08",
         "septiembre": "09", "octubre": "10", "noviembre": "11",
         "diciembre": "12"}
DATE_ES = re.compile(r"(\d{1,2})\s+de\s+(" + "|".join(MESES) + r")\s+(?:de\s+)?(\d{4})",
                     re.IGNORECASE)
DATE_DMY = re.compile(r"\b(\d{1,2})/(\d{1,2})/(\d{4})\b")
STUDENT_OK = re.compile(r"^[^()]{0,60}(\(c[oó]digo [\w]+\))?$", re.IGNORECASE)


def find_dates(text):
    """All distinct ISO-normalizable dates in `text`, in order of appearance."""
    found = []
    for m in DATE_FIND.finditer(text):
        found.append(m.group(0))
    for m in DATE_ES.finditer(text):
        found.append(f"{m.group(3)}-{MESES[m.group(2).lower()]}-{int(m.group(1)):02d}")
    for m in DATE_DMY.finditer(text):
        d, mo, y = int(m.group(1)), int(m.group(2)), m.group(3)
        if 1 <= mo <= 12 and 1 <= d <= 31:
            found.append(f"{y}-{mo:02d}-{d:02d}")
    seen = []
    for f in found:
        if f not in seen:
            seen.append(f)
    return seen

# Reviewer wordings that MEAN a canonical flag but don't prefix-match it —
# without these aliases a flag with real canonical intent silently drops out
# of every TA flag-filtered view (relocation to observations loses the
# flag's FUNCTION even though the text survives).
FLAG_ALIASES = {
    "DISCREPANCIA DE FECHA": "DISCREPANCIA FECHA",
    "DISCREPANCIA EN FECHA": "DISCREPANCIA FECHA",
    "ENTREGA SIN DIAPOSITIVAS": "ENTREGA SIN PPT",
    "SIN PPT": "ENTREGA SIN PPT",
    "SIN DIAPOSITIVAS": "ENTREGA SIN PPT",
    "FORMATO NO-PPT": "ENTREGA SIN PPT",
    "SIN EVIDENCIA": "SIN EVIDENCIA PROPIA",
    "IMPACTO SIN CUANTIFICAR": "IMPACTO NO CUANTIFICADO",
    "SIN CUANTIFICACION DE IMPACTO": "IMPACTO NO CUANTIFICADO",
    "REVISION MANUAL": "REVISAR MANUALMENTE",
    "VERIFICAR LA FECHA": "VERIFICAR FECHA",
}


def normalize_review(r, expect_pages=None, expect_materials=None):
    """Validate + normalize IN PLACE. Returns (problems, moved) lists.

    problems -> the review must be retried/rejected (hard failures; nothing
    is auto-fixed for these). moved -> fields whose out-of-format content was
    relocated to `observations` (soft fixes; content preserved).
    """
    problems = []
    moved = []
    # LLM-drifted shapes must not crash the gate (a TypeError inside assembly
    # loses every student's row at once): coerce before touching anything.
    o = r.get("observations")
    if isinstance(o, list):
        r["observations"] = " | ".join(str(x) for x in o)
    elif o is not None and not isinstance(o, str):
        r["observations"] = str(o)
    f0 = r.get("flags")
    if isinstance(f0, str):
        r["flags"] = [f0]
    elif f0 is not None and not isinstance(f0, list):
        r["flags"] = [str(f0)]
    obs = [r.get("observations")] if r.get("observations") else []

    # --- scores & justifications (hard failures, never auto-fixed) ---------
    s = r.get("scores") or {}
    for k in ("poc", "impacto", "comunicacion"):
        v = s.get(k)
        # isinstance(True, int) is True in Python — a reviewer emitting
        # "poc": true must NOT pass as a real 1.0 grade.
        if (isinstance(v, bool) or not isinstance(v, (int, float))
                or not 1.0 <= v <= 5.0):
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
        def bare(x):
            return str(x).strip().strip("«»\"'")
        seen = {bare(m) for m in (r.get("materials_reviewed") or [])}
        missing = [m for m in expect_materials if bare(m) not in seen]
        if missing:
            problems.append(f"materials_reviewed no cubre: {missing}")

    # --- verification coherence -------------------------------------------
    conf = r.get("verification_confidence")
    age = r.get("age_months")
    if isinstance(age, bool):
        age = None
    if conf not in ("alta", "media", "baja"):
        # an unrecognized value would escape BOTH coherence branches and
        # silently disarm the DQ filter for the row
        problems.append(f"verification_confidence inválida: {conf!r} "
                        "(solo alta|media|baja)")
    if conf in ("alta", "media") and not isinstance(age, (int, float)):
        problems.append("confianza alta/media pero age_months no es numérico "
                        "(desactiva silenciosamente el filtro DQ)")
    if conf == "baja" and age is not None:
        problems.append("confianza baja pero age_months no es null "
                        "(un número no verificado presentado como dato)")

    # --- DQ coherence: the exclusion decision itself gets a mechanical gate
    dq = bool(r.get("disqualified"))
    if dq and conf == "baja":
        problems.append("disqualified=true con confianza baja — la regla "
                        "prohíbe descalificar sin verificación confiable")
    if (not dq and conf in ("alta", "media")
            and isinstance(age, (int, float)) and age > 4.5):
        problems.append(f"age_months={age} (> 4.5) con confianza {conf} pero "
                        "disqualified=false — el filtro de exclusión exige DQ")
    if (dq and isinstance(age, (int, float)) and age <= 4.0
            and not (r.get("dq_reason") or "").strip()):
        problems.append("disqualified=true con age_months <= 4.0 y sin "
                        "dq_reason — justifica la descalificación")

    # --- date fields: extract-or-empty, prose to observations --------------
    for f_ in ("declared_launch_date", "verified_launch_date"):
        v = (r.get(f_) or "").strip()
        if DATE_OK.match(v):
            continue
        candidates = find_dates(v)
        obs.append(f"{f_} original del revisor: {v}")
        moved.append(f_)
        if len(candidates) == 1:
            # A reconstructed date is an INFERENCE, not a verification — it
            # must carry VERIFICAR FECHA or a "consultado el ..." access date
            # silently becomes a launch date and drives the DQ filter.
            r[f_] = candidates[0]
            r.setdefault("flags", [])
            if "VERIFICAR FECHA" not in r["flags"]:
                r["flags"].append("VERIFICAR FECHA")
        else:
            r[f_] = ""
            if not candidates:
                problems.append(f"{f_} sin fecha reconocible: {v[:60]!r}")
            else:
                problems.append(f"{f_} con {len(candidates)} fechas distintas "
                                f"({candidates}) — entrega UNA sola fecha en "
                                "formato YYYY-MM-DD")

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
    def _flag_match(base, name):
        """Prefix match with a WORD BOUNDARY: 'VERIFICAR FECHAS COINCIDEN'
        must NOT collapse to VERIFICAR FECHA (meaning inversion)."""
        if base == name:
            return True
        if base.startswith(name):
            nxt = base[len(name):len(name) + 1]
            return not nxt.isalnum()
        return False

    kept = []
    for f_ in r.get("flags") or []:
        f_ = str(f_).strip()
        base = f_.split(":")[0].split("(")[0].strip().upper()
        canon = next((c for c in CANONICAL_FLAGS if _flag_match(base, c)), None)
        if canon is None:
            canon = next((v for a, v in FLAG_ALIASES.items()
                          if _flag_match(base, a)), None)
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
