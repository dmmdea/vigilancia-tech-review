#!/usr/bin/env python3
"""Build the three-sheet results Excel (Ranking, Detalle, Meta) from the orchestrator's results JSON.

Usage:
    python make_excel.py <results.json> <output.xlsx> [--listing=<listing.json>]... [--dq-policy=penalty:0.5:3.0]

The final grade is computed HERE (single source of truth):
    rubric = 0.50*poc + 0.25*impacto + 0.25*comunicacion, rounded HALF-UP to
    0.01 in exact decimal arithmetic — what a TA gets recomputing it with
    Excel's ROUND (binary-float round() sent an exact 2.995 to 2.99, which the
    3.0 floor below then failed to protect)
    a row that broke the date/general-tool rule (`disqualified` in the
    review) -> Nota final per --dq-policy (Nota rúbrica always shown):
        penalty:X[:F]  (DEFAULT, penalty:0.5:3.0) the row is NOT excluded: final
                 = rúbrica − X (floor 1.0); with F, a rúbrica >= F never ends
                 below F — breaking the date rule alone never turns a pass into
                 a fail. Ranked with everyone, top-5 eligible. Operator decision
                 2026-10-01 (final round): "should just penalize the grade a bit
                 instead of disqualifying" -> −0.5, protected at 3.0.
        cap:N    final = min(rúbrica, N) — TA-team feedback 2026-08-21:
                 "flexibiliza... les ponga 3 o algo así y que deje la nota"
        fixed:N  final = N regardless of the rubric
        rubric   final = rúbrica (no penalty; the DQ stays visible as a flag)
        legacy   final = 1.0 (the original automatic-1.0 rule)
    Every mode except penalty EXCLUDES the row from the ranking (it sorts
    after the valid rows and cannot be top-5).
    unchanged resubmission (`nota_piso`, set by assemble_results.py when the
    file is the same as a delivered earlier round) -> Nota final is never
    below that round's grade (operator rule 2026-10-01); flag NOTA PROTEGIDA
    when the floor raised it, column "Nota mínima (sin cambios)" always shown.
A reviewed row with missing/invalid/out-of-range scores is DOWNGRADED to
"no_revisado" with a visible reason — it never renders as a normal graded row
and never aborts the workbook (fairness: every submission stays visible).

Row identity is the Drive file id (`id` field) when present, so duplicate
filenames cannot corrupt the top-5. Ranking order: reviewed rows (by final
desc; under an excluding policy, rule-breaking rows after the rest) -> not
reviewed. The top 5 ranked rows are starred and highlighted.

--listing (repeatable): the drive_list.py JSON(s) for the folder and any
subfolders. When given, EVERY listed entry needs a results row (reviewed or
NO REVISADO) — except folders that were themselves passed as a --listing.
Any entry without a row fails the build (exit 2) naming the absent files —
the mechanical backstop for "nothing is silently dropped".
Exit codes: 0 ok · 1 usage (bad args, unknown option, unreadable results
JSON) · 2 empty results / duplicate id / unreadable or invalid listing /
completeness violation.
"""
import json
import sys
from decimal import ROUND_HALF_UP, Decimal

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# Windows defaults std streams to cp1252; names and messages are non-ASCII.
for _s in (sys.stdout, sys.stderr):
    # Only touch the process's own console streams — never a stream an
    # importing caller substituted (reconfiguring theirs corrupts their file).
    if _s in (sys.__stdout__, sys.__stderr__) and hasattr(_s, "reconfigure"):
        _s.reconfigure(encoding="utf-8")

HEADER_FILL = PatternFill("solid", fgColor="1F2937")
HEADER_FONT = Font(bold=True, color="FFFFFF")
TOP5_FILL = PatternFill("solid", fgColor="FDE68A")
DQ_FILL = PatternFill("solid", fgColor="FECACA")
PEN_FILL = PatternFill("solid", fgColor="FED7AA")      # penalizada, no excluida
NOREV_FILL = PatternFill("solid", fgColor="E5E7EB")
ANEXO_FILL = PatternFill("solid", fgColor="D1FAE5")    # leído en la revisión
REEMP_FILL = PatternFill("solid", fgColor="EDE9FE")    # versión reemplazada
THIN = Border(*[Side(style="thin", color="D1D5DB")] * 4)

DEFAULT_DQ_POLICY = ("penalty", (0.5, 3.0))
DQ_POLICY = DEFAULT_DQ_POLICY   # overridden by --dq-policy; see docstring


def q2(x) -> float:
    """Round half-up to 0.01 in exact decimal arithmetic (Excel's ROUND)."""
    return float(Decimal(str(x)).quantize(Decimal("0.01"), ROUND_HALF_UP))


def parse_dq_policy(text: str):
    """'penalty:0.5[:3.0]' | 'cap:3.0' | 'fixed:3.0' | 'rubric' | 'legacy'
    -> (mode, value); penalty's value is (deduction, protected_floor|None)."""
    mode, _, val = text.strip().lower().partition(":")
    if mode in ("rubric", "legacy"):
        if val:
            raise ValueError(f"'{mode}' no lleva valor")
        return (mode, 1.0 if mode == "legacy" else None)
    if mode == "penalty":
        x, _, f = val.partition(":")
        try:
            x = float(x)
            f = float(f) if f else None
        except ValueError:
            raise ValueError("'penalty' requiere un descuento y, opcional, un "
                             "piso protegido: p. ej. penalty:0.5:3.0")
        if not 0.0 < x <= 4.0:
            raise ValueError("el descuento debe estar entre 0 y 4.0")
        if f is not None and not 1.0 <= f <= 5.0:
            raise ValueError("el piso protegido debe estar entre 1.0 y 5.0")
        return (mode, (x, f))
    if mode in ("cap", "fixed"):
        try:
            v = float(val)
        except ValueError:
            raise ValueError(f"'{mode}' requiere un número, p. ej. {mode}:3.0")
        if not 1.0 <= v <= 5.0:
            raise ValueError("el valor debe estar entre 1.0 y 5.0")
        return (mode, v)
    raise ValueError("modos válidos: penalty:X[:F], cap:N, fixed:N, rubric, legacy")


def penalizes() -> bool:
    """True when breaking the date/general-tool rule costs points but does
    NOT exclude the row from the ranking."""
    return DQ_POLICY[0] == "penalty"


def apply_dq_policy(rubric: float) -> float:
    mode, v = DQ_POLICY
    if mode == "penalty":
        x, floor = v
        final = max(Decimal("1.0"), Decimal(str(rubric)) - Decimal(str(x)))
        if floor is not None and rubric >= floor:
            final = max(final, Decimal(str(floor)))   # never fails a pass
        return q2(final)
    if mode == "cap":
        return q2(min(rubric, v))
    if mode == "fixed":
        return v
    if mode == "legacy":
        return 1.0
    return rubric            # 'rubric'


def dq_policy_text() -> str:
    mode, v = DQ_POLICY
    if mode == "penalty":
        x, floor = v
        return (f"no se descalifica: nota final = nota rúbrica − {x:.1f}"
                + (f"; si la nota rúbrica es ≥ {floor:.1f}, la penalización nunca "
                   f"la deja por debajo de {floor:.1f}" if floor is not None else "")
                + "; la fila compite en el ranking como las demás")
    return {"cap": f"nota final = mín(nota rúbrica, {v}) — conserva la nota de "
                   "la rúbrica y la limita",
            "fixed": f"nota final = {v} fija",
            "rubric": "nota final = nota rúbrica (sin penalización; la DQ "
                      "queda como señal)",
            "legacy": "nota final = 1.0 automática"}[mode]


RANK_COLS = [
    ("Posición", 9), ("Top 5", 7), ("Archivo", 38), ("Estudiante", 24),
    ("Clave", 12),
    ("Herramienta", 26), ("Fecha lanz. declarada", 15),
    ("Fecha lanz. verificada", 15), ("Fuente verificación", 40),
    ("Confianza", 10), ("Fecha ancla", 24), ("Edad (meses)", 9),
    ("¿Descalificado?", 13),
    ("Razón DQ", 30), ("PoC (50%)", 10), ("Impacto (25%)", 10),
    ("Comunicación (25%)", 12), ("Nota rúbrica", 10), ("Nota final", 10),
    ("Nota mínima (sin cambios)", 12),
    ("Retroalimentación", 62),
    ("Indicio IA (1-5)", 10),
    ("Huella", 34),
    ("Flags revisión humana", 28),
    ("Estado", 16), ("Detalle estado", 30),
]
GRADE_COLS = {"PoC (50%)", "Impacto (25%)", "Comunicación (25%)",
              "Nota rúbrica", "Nota final", "Nota mínima (sin cambios)"}

DETAIL_COLS = [
    ("Archivo", 38), ("Herramienta", 24), ("Slides leídas / total", 12),
    ("Justificación PoC", 60), ("Justificación Impacto", 60),
    ("Justificación Comunicación", 60),
    ("Retroalimentación para el estudiante", 60),
    ("Evidencia indicio IA", 50),
    ("Notas de evidencia", 60),
]


def floor_of(r: dict):
    """The unchanged-submission floor, rounded — only when the assembler set
    it from the delivered history (piso_rondas is its provenance); a bare
    nota_piso from anywhere else is ignored."""
    piso = r.get("nota_piso")
    if (not r.get("piso_rondas") or isinstance(piso, bool)
            or not isinstance(piso, (int, float)) or not 1.0 <= piso <= 5.0):
        return None
    return q2(piso)


def retro_text(r: dict) -> str:
    """The paste-ready student feedback: two strengths, two to improve.
    Review sets from before R25 carry a free-form `feedback_sugerido`."""
    rt = r.get("retroalimentacion")
    if isinstance(rt, dict):
        parts = []
        for key, title in (("fortalezas", "Fortalezas"),
                           ("por_mejorar", "Por mejorar")):
            items = [str(s).strip() for s in (rt.get(key) or [])
                     if str(s or "").strip()]
            if items:
                parts.append(title + ":\n" + "\n".join(f"• {s}" for s in items))
        if parts:
            return "\n".join(parts)
    return r.get("feedback_sugerido", "") or ""


def ancla_text(r: dict) -> str:
    def s(v):
        return (" ".join(map(str, v)) if isinstance(v, list)
                else str(v or "")).strip()
    fa = s(r.get("fecha_ancla"))
    if not fa:
        return ""
    motivo = s(r.get("ancla_motivo"))
    return f"{fa} · {motivo}" if motivo else fa


def normalize(r: dict) -> None:
    """Validate one result row in place: set r['_final'] or downgrade to no_revisado."""
    # flags is the one field this module appends to and joins — coerce any
    # LLM-drifted shape (null, bare string, list of dicts) to a list of strings.
    f = r.get("flags")
    r["flags"] = ([str(x) for x in f] if isinstance(f, list)
                  else [str(f)] if f else [])
    # Mechanical backstops for the launch-date rules (prompt-only otherwise):
    age = r.get("age_months")
    conf = r.get("verification_confidence")
    if isinstance(age, (int, float)):
        if 3.5 <= age <= 4.5 and "VERIFICAR FECHA" not in r["flags"]:
            r["flags"].append("VERIFICAR FECHA")
    elif conf in ("alta", "media") and r.get("status") == "revisado":
        # confident verification must yield a numeric age, else the DQ filter
        # silently never fires
        if "REVISAR MANUALMENTE" not in r["flags"]:
            r["flags"].append("REVISAR MANUALMENTE")
        if "EDAD SIN CALCULAR" not in r["flags"]:
            r["flags"].append("EDAD SIN CALCULAR")
    status = r.get("status")
    if status not in ("revisado", "no_revisado", "revisado_anexo",
                      "reemplazada"):
        r["status"] = "no_revisado"
        r.setdefault("status_reason", "estado ausente o inválido en results.json")
        r["_final"] = None
        return
    r["_rubric"] = None
    if status in ("no_revisado", "revisado_anexo", "reemplazada"):
        r["_final"] = None
        return
    dq = bool(r.get("disqualified"))
    s = r.get("scores") or {}
    try:
        poc = float(s["poc"])
        imp = float(s["impacto"])
        com = float(s["comunicacion"])
    except (KeyError, TypeError, ValueError):
        # regardless of policy: a DQ row with no usable scores was never
        # actually reviewed — 'fixed'/'legacy' must not fabricate a grade
        # for it (silent-failure review, 2026-08-21); humans assign it
        r["status"] = "no_revisado"
        r["status_reason"] = "puntajes ausentes o inválidos en la revisión"
        r["flags"].append("REVISAR MANUALMENTE")
        r["_final"] = None
        return
    if any(not 1.0 <= v <= 5.0 for v in (poc, imp, com)):
        r["status"] = "no_revisado"
        r["status_reason"] = (f"puntaje fuera de rango 1.0-5.0 "
                              f"(poc={poc}, impacto={imp}, comunicacion={com})")
        r["flags"].append("REVISAR MANUALMENTE")
        r["_final"] = None
        return
    # exact decimal, half-up: the TA's Excel ROUND gives the same number
    rubric = q2(Decimal("0.50") * Decimal(str(poc)) + Decimal("0.25") * Decimal(str(imp))
                + Decimal("0.25") * Decimal(str(com)))
    r["_rubric"] = rubric
    r["_final"] = apply_dq_policy(rubric) if dq else rubric
    # R30: the same submission as a delivered earlier round never scores
    # lower than it did then (assemble_results.py records the floor)
    piso = floor_of(r)
    if piso is not None and r["_final"] < piso:
        r["_final"] = piso
        if "NOTA PROTEGIDA" not in r["flags"]:
            r["flags"].append("NOTA PROTEGIDA")


def excluded(r: dict) -> bool:
    """Out of the ranking: a DQ under every policy except penalty."""
    return bool(r.get("disqualified")) and not penalizes()


def sort_key(r: dict):
    status = r.get("status")
    final = r.get("_final")
    if status == "revisado":
        group = 1 if excluded(r) else 0
    else:
        group = {"revisado_anexo": 2, "reemplazada": 3}.get(status, 4)
    return (group, -(final if final is not None else 0.0))


def check_completeness(results: list, listing_paths: list) -> None:
    """EVERY listed entry must be accounted for: non-folder entries need a
    results row (reviewed or NO REVISADO); folder entries need either their own
    --listing (they were explored) or a results row (declared unexplored)."""
    expected = {}
    explored_folders = set()
    entries = []
    for p in listing_paths:
        try:
            with open(p, encoding="utf-8") as f:
                listing = json.load(f)
        except (OSError, json.JSONDecodeError) as e:
            print(f"ERROR: cannot read listing {p}: {e}", file=sys.stderr)
            sys.exit(2)
        if (not isinstance(listing, dict)
                or not isinstance(listing.get("entries"), list)
                or not listing.get("folder_id")
                or not all(isinstance(e, dict) for e in listing["entries"])):
            print(f"ERROR: {p} is not a drive_list.py listing "
                  "(needs folder_id and a list of entry objects).",
                  file=sys.stderr)
            sys.exit(2)
        explored_folders.add(listing["folder_id"])
        entries.extend(listing["entries"])
    for e in entries:
        if e.get("kind") == "folder" and e.get("id") in explored_folders:
            continue
        expected[e.get("id")] = e.get("name", "(sin nombre)")
    have = {r.get("id") for r in results if r.get("id")}
    missing = {i: n for i, n in expected.items() if i not in have}
    if missing:
        print("ERROR: results.json is missing rows for these listed entries "
              "(nothing may be silently dropped — every entry gets a row, "
              "reviewed or NO REVISADO):", file=sys.stderr)
        for i, n in missing.items():
            print(f"  - {n} (id {i})", file=sys.stderr)
        sys.exit(2)


def style_header(ws, cols):
    for i, (name, width) in enumerate(cols, 1):
        c = ws.cell(row=1, column=i, value=name)
        c.fill, c.font = HEADER_FILL, HEADER_FONT
        c.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.freeze_panes = "A2"


def main() -> None:
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    listings = [a.split("=", 1)[1] for a in sys.argv[1:]
                if a.startswith("--listing=")]
    global DQ_POLICY
    DQ_POLICY = DEFAULT_DQ_POLICY   # a run without the flag must get the default
    for a in sys.argv[1:]:
        if a.startswith("--dq-policy="):
            try:
                DQ_POLICY = parse_dq_policy(a.split("=", 1)[1])
            except ValueError as e:
                print(f"ERROR: --dq-policy inválida: {e}", file=sys.stderr)
                sys.exit(1)
    unknown = [a for a in sys.argv[1:]
               if a.startswith("--") and not a.startswith(("--listing=",
                                                             "--dq-policy="))]
    if unknown:
        # A typo'd flag must never silently disarm the completeness gate.
        print(f"ERROR: unknown option(s): {' '.join(unknown)}", file=sys.stderr)
        sys.exit(1)
    if len(args) != 2:
        print(__doc__, file=sys.stderr)
        sys.exit(1)
    try:
        with open(args[0], encoding="utf-8") as f:
            data = json.load(f)
    except (OSError, json.JSONDecodeError) as e:
        print(f"ERROR: cannot read results JSON {args[0]}: {e}", file=sys.stderr)
        sys.exit(1)

    results = data.get("results", [])
    if not results:
        print("ERROR: results.json contains no results.", file=sys.stderr)
        sys.exit(2)
    if listings:
        check_completeness(results, listings)

    seen_ids = {}
    for idx, r in enumerate(results):
        rid = r.get("id")
        if rid:
            if rid in seen_ids:
                print(f"ERROR: duplicate id '{rid}' in results.json (rows "
                      f"{seen_ids[rid]} and {idx}) — one Drive file must have "
                      "exactly one row.", file=sys.stderr)
                sys.exit(2)
            seen_ids[rid] = idx
        r["_uid"] = rid or f"__row{idx}"
        normalize(r)
    results.sort(key=sort_key)

    rankable = [r for r in results
                if r.get("status") == "revisado"
                and not excluded(r) and r["_final"] is not None]
    top5_uids = {r["_uid"] for r in rankable[:5]}
    rankable_uids = {r["_uid"] for r in rankable}
    # Flag a tie only when it actually straddles the top-5 cut (position 5 vs 6);
    # a tie wholly inside the top 5 contests nothing.
    if len(rankable) > 5 and rankable[5]["_final"] == rankable[4]["_final"]:
        cut = rankable[4]["_final"]
        for r in rankable:
            if r["_final"] == cut:
                r["flags"].append("EMPATE TOP5")

    wb = Workbook()
    ws = wb.active
    ws.title = "Ranking"
    # under penalty the DQ columns describe a deduction, not an exclusion —
    # a TA reading "¿Descalificado? SÍ" next to a 4.1 would be misled
    cols = [(("¿Penalizada?" if n == "¿Descalificado?" else
              "Razón penalización" if n == "Razón DQ" else n), w)
            for n, w in RANK_COLS] if penalizes() else RANK_COLS
    style_header(ws, cols)

    pos = 0
    for i, r in enumerate(results, start=2):
        reviewed = r.get("status") == "revisado"
        dq = bool(r.get("disqualified"))
        in_top5 = r["_uid"] in top5_uids
        if r["_uid"] in rankable_uids:
            pos += 1
            position = pos
        else:
            position = ""
        s = r.get("scores") or {}
        status = r.get("status")
        estado = {"revisado": "REVISADO",
                  "revisado_anexo": "REVISADO (ANEXO)",
                  "reemplazada": "REEMPLAZADA"}.get(status, "NO REVISADO")
        row = [
            position,
            "⭐ TOP 5" if in_top5 else "",
            r.get("file", ""), r.get("student", ""),
            r.get("canvas_key", ""), r.get("tool", ""),
            r.get("declared_launch_date", ""), r.get("verified_launch_date", ""),
            r.get("verification_source", ""), r.get("verification_confidence", ""),
            ancla_text(r) if reviewed else "",
            r.get("age_months", ""),
            ("SÍ" if dq else "NO") if reviewed else "",
            r.get("dq_reason", ""),
            s.get("poc", ""), s.get("impacto", ""), s.get("comunicacion", ""),
            r.get("_rubric") if r.get("_rubric") is not None else "",
            r["_final"] if r["_final"] is not None else "",
            # shown for unreviewed main files too: the human grader needs it
            floor_of(r) if floor_of(r) is not None else "",
            retro_text(r) if reviewed else "",
            r.get("indicio_ia", ""),
            r.get("huella", "") or "",
            ", ".join(r.get("flags", [])),
            estado,
            r.get("status_reason", ""),
        ]
        fill = (TOP5_FILL if in_top5
                else (PEN_FILL if penalizes() else DQ_FILL) if (reviewed and dq)
                else ANEXO_FILL if status == "revisado_anexo"
                else REEMP_FILL if status == "reemplazada"
                else NOREV_FILL if not reviewed else None)
        for j, v in enumerate(row, 1):
            c = ws.cell(row=i, column=j, value=v)
            c.border = THIN
            if fill:
                c.fill = fill
            # by header, not by index: a column inserted upstream once would
            # have silently moved the 0.00 format onto the wrong cells
            if RANK_COLS[j - 1][0] in GRADE_COLS:
                c.number_format = "0.00"
            elif RANK_COLS[j - 1][0] == "Retroalimentación":
                c.alignment = Alignment(wrap_text=True, vertical="top")

    ws2 = wb.create_sheet("Detalle")
    style_header(ws2, DETAIL_COLS)
    for i, r in enumerate(results, start=2):
        j = r.get("justification") or {}
        row = [
            r.get("file", ""), r.get("tool", ""),
            f'{r.get("pages_read", "?")} / {r.get("pages_total", "?")}',
            j.get("poc", ""), j.get("impacto", ""), j.get("comunicacion", ""),
            retro_text(r),
            r.get("indicio_ia_evidencia", ""),
            r.get("evidence_notes", ""),
        ]
        for k, v in enumerate(row, 1):
            c = ws2.cell(row=i, column=k, value=v)
            c.border = THIN
            c.alignment = Alignment(wrap_text=True, vertical="top")

    meta = wb.create_sheet("Meta")
    meta["A1"], meta["B1"] = "Fecha de corrida", data.get("run_date", "")
    meta["A2"], meta["B2"] = "Carpeta Drive", data.get("folder_url", "")
    meta["A3"], meta["B3"] = "Regla de corte", ("> 4 meses entre el lanzamiento verificado y la fecha ancla, o "
                                                "herramienta general sin función específica reciente → "
                                                + ("PENALIZADA" if penalizes() else "DESCALIFICADA")
                                                + ". Fecha ancla = la PRIMERA ronda en que el "
                                                "estudiante presentó la misma herramienta (llevarla de una "
                                                "ronda a otra es válido); si la herramienta es nueva, su fecha "
                                                "de entrega; nunca la fecha de corrida. "
                                                f"Política de nota: {dq_policy_text()}; "
                                                "zona 3.5–4.5 meses lleva flag VERIFICAR FECHA")
    meta["A4"], meta["B4"] = "Ponderación", ("PoC 50% · Impacto 25% · Comunicación 25% = Nota rúbrica "
                                             "(redondeo a 0.01 hacia arriba en el medio, como ROUND de "
                                             "Excel); Nota final = Nota rúbrica salvo incumplimiento de la "
                                             "regla de fecha (ver Regla de corte), y nunca menor que la "
                                             "Nota mínima de una entrega sin cambios")
    meta["A6"], meta["B6"] = "Indicio IA (1-5)", ("señal ADVISORY de uso de IA sin filtro sobre el material "
                                                  "entregado (1=curado a mano, 5=volcado sin filtrar); NUNCA "
                                                  "es componente de la nota")
    meta["A7"], meta["B7"] = "Estados", ("REVISADO=fila calificada · REVISADO (ANEXO)=archivo leído dentro "
                                          "de la revisión integral del estudiante · REEMPLAZADA=versión "
                                          "anterior superada por una entrega más reciente · NO REVISADO=requiere humano")
    meta["A5"], meta["B5"] = "Generado por", ("vigilancia-tech-review (revisores de IA con "
                                              "contexto limpio; revisión humana "
                                              "requerida para nota oficial)")
    meta["A8"], meta["B8"] = "Retroalimentación", ("dos fortalezas y dos por mejorar, en frases "
                                                   "cortas, lista para copiar al estudiante "
                                                   "(pedido del equipo docente)")
    meta["A9"], meta["B9"] = "Nota mínima (sin cambios)", (
        "si la entrega es la MISMA que la de una ronda anterior (huella de contenido "
        "idéntica, o un archivo re-guardado con el mismo contenido, con la evidencia en "
        "las notas), la nota final nunca es menor que la nota que recibió en esa ronda. "
        "Flag ENTREGA SIN CAMBIOS; NOTA PROTEGIDA cuando ese mínimo subió la nota")
    meta.column_dimensions["A"].width = 20
    meta.column_dimensions["B"].width = 80

    wb.save(args[1])
    print(f"OK {args[1]} ({len(results)} filas, top5={len(top5_uids)})")


if __name__ == "__main__":
    main()
