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
        [--require-extended]            # enforce indicio_ia (1-5) +
                                        # retroalimentacion 2+2 (fresh reviews)
        [--anchor-dates=d1|d2|...]      # the student's allowed fechas ancla
                                        # (historial.json anclas_permitidas)

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
  * RETROALIMENTACION (--require-extended): exactly two `fortalezas` and two
    `por_mejorar`, each ONE short phrase (hard cap 20 words; the prompt asks
    for 15). Teaching team, 2026-10-01: "dos cosas buenas y dos por mejorar y
    ya… que no escriba mil cosas" — the previous free-form 2-4 sentences ran
    71 words on average and up to 184. Rejected inside a phrase: first person
    of the reviewer, praise formulas, em dashes, scores, "para la próxima",
    and chained ideas (semicolons / line breaks).
  * FECHA ANCLA (--anchor-dates): the tool's age is measured from the first
    round in which the student presented that same tool, or from their own
    submission date for a new tool — never from the run date. `fecha_ancla`
    must be one of the allowed dates, the verified launch cannot be later
    than it, and `age_months` must match the two dates (±0.15 months and
    never across the 4.0 cutoff; ±0.6 when the launch date is month-precision).

Canonical flags (the only ones that stay in `flags`):
  VERIFICAR FECHA · DISCREPANCIA FECHA · ENTREGA SIN PPT · ENTREGA DUPLICADA
  REVISAR MANUALMENTE · SIN EVIDENCIA PROPIA · IMPACTO NO CUANTIFICADO
  HERRAMIENTA GENERAL - FUNCION ESPECIFICA · SPOT-CHECK FALLIDO
  EVIDENCIA NO LEGIBLE · EVIDENCIA DE ENVIO ANTERIOR INCLUIDA
  ENTREGA SIN CAMBIOS (assembler only)

CLI output: JSON to stdout {"ok": bool, "problems": [...],
"moved_to_observations": [...]}. Exit 0 = acceptable (possibly after
normalization) · 1 = usage · 2 = review must be retried/rejected.
"""
import json
import re
import sys
import unicodedata
from datetime import date

for _s in (sys.stdout, sys.stderr):
    if _s in (sys.__stdout__, sys.__stderr__) and hasattr(_s, "reconfigure"):
        _s.reconfigure(encoding="utf-8")

CANONICAL_FLAGS = (
    "VERIFICAR FECHA", "DISCREPANCIA FECHA", "ENTREGA SIN PPT",
    "ENTREGA DUPLICADA", "REVISAR MANUALMENTE", "SIN EVIDENCIA PROPIA",
    "IMPACTO NO CUANTIFICADO", "HERRAMIENTA GENERAL - FUNCION ESPECIFICA",
    "SPOT-CHECK FALLIDO", "EVIDENCIA NO LEGIBLE",
    "EVIDENCIA DE ENVIO ANTERIOR INCLUIDA",
    "ENTREGA SIN CAMBIOS",      # assembler (R30): same submission as a past round
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

# --- retroalimentación (R25) -------------------------------------------------
RETRO_PARTS = (("fortalezas", "fortaleza"), ("por_mejorar", "por mejorar"))
RETRO_MAX_WORDS = 20          # hard cap; the prompt asks for 15
# A list marker needs whitespace after it: "2.5 horas por reporte" and
# "-5% de error" are content, not bullets (an earlier pattern turned them
# into "5 horas" and "5% de error" — text that is pasted to the student).
RETRO_BULLET = re.compile(r"^\s*(?:[•*·▪●]+\s*|-\s+|\d{1,2}[.)]\s+)")
AGE_TOL_DAY = 0.15            # one-decimal rounding + calendar vs 30.44 days
AGE_TOL_MONTH = 0.6           # launch known only to the month (15th assumed)
# How the monitores do NOT write (vtr voice notes): the reviewer narrating
# itself, praise formulas, em dashes, grades inside the text, a "next time"
# closer that means nothing on a final submission. Only unambiguous forms:
# accented first-person preterites, so an imperative like "revise" passes.
RETRO_BANNED = (
    (re.compile(r"\b(revisé|leí|encontré|noté|observé|considero|me parece|"
                r"creo que|en mi opinión|veo que|vi que)\b", re.IGNORECASE),
     "primera persona del revisor"),
    (re.compile(r"\b(?:muy\s+)?(?:buen|excelente|gran)\s+trabajo\b|"
                r"\bfelicitaciones\b|\bfelicito\b", re.IGNORECASE),
     "fórmula de elogio"),
    (re.compile(r"\bpara la pr[oó]xima\b", re.IGNORECASE),
     "cierre 'para la próxima'"),
    # a GRADE, not any number: "4.5/5", "nota de 3,8", "calificación 4.0".
    # "Prueba sobre 5 facturas", "acierta 4/5 casos" or "puntaje de leads"
    # are content and must pass (each false hit costs a retry, two cost the
    # student a NO REVISADO).
    (re.compile(r"\b[1-5][.,]\d{1,2}\s*/\s*5\b|"
                r"\b(?:nota|calificaci[oó]n)\s+(?:de\s+|final\s+|"
                r"r[uú]brica\s+)?[1-5](?:[.,]\d{1,2})\b", re.IGNORECASE),
     "nota o puntaje dentro del texto"),
    (re.compile("—"), "raya (—)"),
)


def _as_date(s):
    """'YYYY-MM-DD' -> (date, 'dia'); 'YYYY-MM' -> (15th of month, 'mes')."""
    m = re.match(r"^(\d{4})-(\d{2})(?:-(\d{2}))?$", (s or "").strip())
    if not m:
        return None, None
    try:
        return (date(int(m.group(1)), int(m.group(2)),
                     int(m.group(3)) if m.group(3) else 15),
                "dia" if m.group(3) else "mes")
    except ValueError:
        return None, None


def months_between(launch, anchor):
    """Age in months, the reviewers' convention: days / 30.44."""
    return (anchor - launch).days / 30.4375


def _normalize_retro(r):
    """Soft fixes only: list shape, bullets/numbering, surrounding spaces.
    Returns the problems a strict (fresh-review) gate would raise."""
    problems = []
    rt = r.get("retroalimentacion")
    if not isinstance(rt, dict):
        return ["retroalimentacion ausente o con otra forma (requerido: "
                "{\"fortalezas\": [2 frases], \"por_mejorar\": [2 frases]})"]
    seen = []
    for key, label in RETRO_PARTS:
        items = rt.get(key)
        if isinstance(items, str):
            items = [items]
        if not isinstance(items, list):
            items = []
        clean = []
        for s in items:
            s = RETRO_BULLET.sub("", str(s or "")).strip()
            if s:
                clean.append(s)
        rt[key] = clean
        if len(clean) != 2:
            problems.append(f"retroalimentacion.{key}: {len(clean)} frase(s) "
                            "— deben ser EXACTAMENTE 2")
        for i, s in enumerate(clean, 1):
            n = len(s.split())
            if n > RETRO_MAX_WORDS:
                problems.append(f"{label} {i}: {n} palabras — una frase corta "
                                f"(máximo 15; tope duro {RETRO_MAX_WORDS})")
            if ";" in s or "\n" in s:
                problems.append(f"{label} {i}: una sola idea por frase (sin "
                                "punto y coma ni saltos de línea)")
            for rx, what in RETRO_BANNED:
                m = rx.search(s)
                if m:
                    problems.append(f"{label} {i}: {what} ({m.group(0)!r})")
            k = s.casefold().rstrip(".")
            if k in seen:
                problems.append(f"{label} {i}: frase repetida")
            seen.append(k)
    r["retroalimentacion"] = rt
    return problems


def find_dates(text):
    """All distinct ISO-normalizable dates in `text` (grouped by pattern:
    ISO first, then Spanish long-form, then dd/mm/yyyy)."""
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
    # a month-precision rendering of a date whose full form is also present is
    # the SAME date, not a second candidate (e.g. "2026-07" + "2026-07-16")
    full = [f for f in seen if len(f) == 10]
    seen = [f for f in seen
            if len(f) == 10 or not any(d.startswith(f) for d in full)]
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
    "SIN EVIDENCIA": "SIN EVIDENCIA PROPIA",  # safe again: aliases match EXACT
    "IMPACTO SIN CUANTIFICAR": "IMPACTO NO CUANTIFICADO",
    "SIN CUANTIFICACION DE IMPACTO": "IMPACTO NO CUANTIFICADO",
    "REVISION MANUAL": "REVISAR MANUALMENTE",
    "VERIFICAR LA FECHA": "VERIFICAR FECHA",
}


def normalize_review(r, expect_pages=None, expect_materials=None,
                     require_extended=False, anchor_dates=None):
    """Validate + normalize IN PLACE. Returns (problems, moved) lists.

    problems -> the review must be retried/rejected (hard failures; nothing
    is auto-fixed for these). moved -> fields whose out-of-format content was
    relocated to `observations` (soft fixes; content preserved).

    require_extended: enforce the class-feedback fields (`indicio_ia`
    integer 1-5 and the 2+2 `retroalimentacion`) as HARD requirements —
    orchestrators pass this on fresh per-review calls; assembly re-validation
    leaves it off so pre-extension review sets (which carry the older
    free-form `feedback_sugerido`) keep working.

    anchor_dates: the student's allowed `fecha_ancla` values; when given, the
    anchor and the age arithmetic are checked (None = pre-anchor review set).
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
    elif not isinstance(f0, list):
        # covers None AND any other drifted scalar — a null here once crashed
        # the entire assembly (TypeError on membership test)
        r["flags"] = [] if f0 is None else [str(f0)]
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
            # Unicode-normalize: a filename coming out of a .zip is often
            # NFD ("tecnolo" + combining acute) while the reviewer echoes
            # NFC. Byte-comparing them fails a CORRECT review of any
            # accented Spanish filename, and a repeated failure downgrades
            # the student to NO REVISADO (2026-08-21 run).
            return unicodedata.normalize(
                "NFC", str(x).strip().strip("«»\"'"))
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

    # --- extended fields (class-1 feedback, R17/R18) ------------------------
    ia = r.get("indicio_ia")
    if isinstance(ia, bool) or (ia is not None
                                and not isinstance(ia, int)):
        # floats/strings drift here; a non-integer signal is unusable for
        # TA filtering — relocate and clear rather than guessing
        obs.append(f"indicio_ia original del revisor: {ia!r}")
        moved.append("indicio_ia")
        r["indicio_ia"] = None
        ia = None
    if isinstance(ia, int) and not 1 <= ia <= 5:
        problems.append(f"indicio_ia fuera de rango 1-5: {ia}")
    retro_problems = (_normalize_retro(r) if "retroalimentacion" in r
                      or require_extended else [])
    if require_extended:
        if not isinstance(ia, int):
            problems.append("indicio_ia ausente o no entero (requerido: 1-5)")
        problems.extend(retro_problems)

    # --- fecha ancla (R26) ----------------------------------------------------
    # shape first, always: a list/number here once reached make_excel raw and
    # would have crashed the workbook for every student
    for k_ in ("fecha_ancla", "ancla_motivo"):
        v_ = r.get(k_)
        if isinstance(v_, list):
            r[k_] = " ".join(str(x) for x in v_).strip()
        elif v_ is not None and not isinstance(v_, str):
            r[k_] = str(v_)
    if anchor_dates is not None:
        fa = str(r.get("fecha_ancla") or "").strip()
        r["fecha_ancla"] = fa
        if fa not in anchor_dates:
            problems.append(f"fecha_ancla {fa!r} no es una de las permitidas "
                            f"({' | '.join(anchor_dates)}) — la edad se mide "
                            "desde la primera ronda con la misma herramienta "
                            "o desde la fecha de entrega, nunca otra fecha")
        if not str(r.get("ancla_motivo") or "").strip():
            problems.append("ancla_motivo vacío — di 'misma herramienta desde "
                            "<ronda>' o 'herramienta nueva en esta entrega'")
        anchor, _ = _as_date(fa)
        launch, prec = _as_date(r.get("verified_launch_date"))
        if anchor and launch and conf in ("alta", "media"):
            # month precision: the launch month may equal the anchor month
            later = ((launch.year, launch.month) > (anchor.year, anchor.month)
                     if prec == "mes" else launch > anchor)
            if later:
                problems.append(
                    f"verified_launch_date {r.get('verified_launch_date')} es "
                    f"POSTERIOR a fecha_ancla {fa}: lo que se lanzó después de "
                    "esa ronda no puede ser lo que el estudiante presentó en "
                    "ella — usa la ronda correcta o la fecha de entrega")
            elif isinstance(age, (int, float)):
                expected = months_between(launch, anchor)
                tol = AGE_TOL_DAY if prec == "dia" else AGE_TOL_MONTH
                if abs(age - expected) > tol:
                    problems.append(
                        f"age_months={age} no coincide con fecha_ancla − "
                        f"fecha verificada ({expected:.1f} meses; días/30.44)")
                elif (age > 4.0) != (expected > 4.0) and prec == "dia":
                    # inside the tolerance but on the other side of the
                    # cutoff: exactly the run-date error R26 removes
                    problems.append(
                        f"age_months={age} y la edad calculada desde "
                        f"fecha_ancla ({expected:.2f}) quedan a lados "
                        "distintos del corte de 4.0 meses — recalcula con "
                        "la fecha ancla")

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
            # aliases match EXACTLY — prefix-matching short aliases like
            # "SIN EVIDENCIA" turned "SIN EVIDENCIA DE IMPACTO" into a
            # PoC-weighted accusation of no OWN evidence (round-2 finding)
            canon = FLAG_ALIASES.get(base)
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
    anchor_dates = None
    out_path = None
    require_extended = "--require-extended" in sys.argv[1:]
    KNOWN = ("--expect-pages=", "--expect-materials=", "--normalized-out=",
             "--require-extended", "--anchor-dates=")
    for a in sys.argv[1:]:
        if a.startswith("--") and not a.startswith(KNOWN):
            # A typo'd flag must never silently DISABLE a fairness check —
            # an ignored --pages-total once let 38 deck reviews through with
            # no page-coverage check at all (2026-08-21 run).
            print(json.dumps({"ok": False, "problems": [
                f"opción desconocida: {a} — ¿quisiste decir "
                "--expect-pages=N o --expect-materials=lbl1|lbl2?"],
                "moved_to_observations": []}, ensure_ascii=False))
            sys.exit(1)
        if a.startswith("--expect-pages="):
            expect_pages = int(a.split("=", 1)[1])
        elif a.startswith("--expect-materials="):
            expect_materials = [m for m in a.split("=", 1)[1].split("|") if m]
        elif a.startswith("--normalized-out="):
            out_path = a.split("=", 1)[1]
        elif a.startswith("--anchor-dates="):
            anchor_dates = [d.strip() for d in a.split("=", 1)[1].split("|")
                            if d.strip()]
            if not anchor_dates or not all(_as_date(d)[1] == "dia"
                                           for d in anchor_dates):
                # an empty/garbled list would reject every anchor — or, worse,
                # a typo'd one would silently skip the check
                print(json.dumps({"ok": False, "problems": [
                    f"--anchor-dates inválido: {a!r} (YYYY-MM-DD separadas "
                    "por |)"], "moved_to_observations": []},
                    ensure_ascii=False))
                sys.exit(1)

    try:
        with open(args[0], encoding="utf-8") as f:
            r = json.load(f)
    except (OSError, json.JSONDecodeError) as e:
        print(json.dumps({"ok": False, "problems": [f"JSON ilegible: {e}"],
                          "moved_to_observations": []}, ensure_ascii=False))
        sys.exit(2)

    problems, moved = normalize_review(r, expect_pages, expect_materials,
                                       require_extended=require_extended,
                                       anchor_dates=anchor_dates)
    if out_path:
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(r, f, ensure_ascii=False, indent=2)
    print(json.dumps({"ok": not problems, "problems": problems,
                      "moved_to_observations": moved}, ensure_ascii=False))
    sys.exit(0 if not problems else 2)


if __name__ == "__main__":
    main()
