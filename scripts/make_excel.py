#!/usr/bin/env python3
"""Build the two-sheet results Excel from the orchestrator's results JSON.

Usage:
    python make_excel.py <results.json> <output.xlsx> [--listing=<listing.json>]...

The final grade is computed HERE (single source of truth):
    final = 0.50*poc + 0.25*impacto + 0.25*comunicacion   (rounded to 0.01)
    disqualified -> final = 1.0
A reviewed row with missing/invalid/out-of-range scores is DOWNGRADED to
"no_revisado" with a visible reason — it never renders as a normal graded row
and never aborts the workbook (fairness: every submission stays visible).

Row identity is the Drive file id (`id` field) when present, so duplicate
filenames cannot corrupt the top-5. Ranking order: reviewed & not DQ (by final
desc) -> DQ -> not reviewed. Top 5 non-DQ rows are starred and highlighted.

--listing (repeatable): the drive_list.py JSON(s) for the folder and any
subfolders. When given, every reviewable entry (pptx/ppt/pdf/gslides) missing
from results.json fails the build (exit 2) listing the absent files — the
mechanical backstop for "nothing is silently dropped".
Exit codes: 0 ok · 1 usage · 2 empty results / completeness violation.
"""
import json
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

HEADER_FILL = PatternFill("solid", fgColor="1F2937")
HEADER_FONT = Font(bold=True, color="FFFFFF")
TOP5_FILL = PatternFill("solid", fgColor="FDE68A")
DQ_FILL = PatternFill("solid", fgColor="FECACA")
NOREV_FILL = PatternFill("solid", fgColor="E5E7EB")
THIN = Border(*[Side(style="thin", color="D1D5DB")] * 4)

DECK_EXTS = (".pptx", ".ppt", ".pdf")

RANK_COLS = [
    ("Posición", 9), ("Top 5", 7), ("Archivo", 38), ("Estudiante", 24),
    ("Herramienta", 26), ("Fecha lanz. declarada", 15),
    ("Fecha lanz. verificada", 15), ("Fuente verificación", 40),
    ("Confianza", 10), ("Edad (meses)", 9), ("¿Descalificado?", 13),
    ("Razón DQ", 30), ("PoC (50%)", 10), ("Impacto (25%)", 10),
    ("Comunicación (25%)", 12), ("Nota final", 10), ("Flags revisión humana", 28),
    ("Estado", 12), ("Detalle estado", 30),
]

DETAIL_COLS = [
    ("Archivo", 38), ("Herramienta", 24), ("Slides leídas / total", 12),
    ("Justificación PoC", 60), ("Justificación Impacto", 60),
    ("Justificación Comunicación", 60), ("Notas de evidencia", 60),
]


def normalize(r: dict) -> None:
    """Validate one result row in place: set r['_final'] or downgrade to no_revisado."""
    r.setdefault("flags", [])
    status = r.get("status")
    if status not in ("revisado", "no_revisado"):
        r["status"] = "no_revisado"
        r.setdefault("status_reason", "estado ausente o inválido en results.json")
        r["_final"] = None
        return
    if status == "no_revisado":
        r["_final"] = None
        return
    if r.get("disqualified"):
        r["_final"] = 1.0
        return
    s = r.get("scores") or {}
    try:
        poc = float(s["poc"])
        imp = float(s["impacto"])
        com = float(s["comunicacion"])
    except (KeyError, TypeError, ValueError):
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
    r["_final"] = round(0.50 * poc + 0.25 * imp + 0.25 * com, 2)


def sort_key(r: dict):
    reviewed = r.get("status") == "revisado"
    dq = bool(r.get("disqualified"))
    final = r.get("_final")
    group = 0 if (reviewed and not dq) else (1 if reviewed else 2)
    return (group, -(final if final is not None else 0.0))


def reviewable_names(listing: dict) -> dict:
    """id -> name for entries the pipeline must account for."""
    out = {}
    for e in listing.get("entries", []):
        kind, name = e.get("kind"), e.get("name", "")
        if kind == "gslides" or (kind == "file"
                                 and name.lower().endswith(DECK_EXTS)):
            out[e.get("id")] = name
    return out


def check_completeness(results: list, listing_paths: list) -> None:
    expected = {}
    for p in listing_paths:
        with open(p, encoding="utf-8") as f:
            expected.update(reviewable_names(json.load(f)))
    have = {r.get("id") for r in results if r.get("id")}
    missing = {i: n for i, n in expected.items() if i not in have}
    if missing:
        print("ERROR: results.json is missing rows for these reviewable files "
              "(nothing may be silently dropped):", file=sys.stderr)
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
    sys.stdout.reconfigure(encoding="utf-8")
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    listings = [a.split("=", 1)[1] for a in sys.argv[1:]
                if a.startswith("--listing=")]
    if len(args) != 2:
        print(__doc__, file=sys.stderr)
        sys.exit(1)
    with open(args[0], encoding="utf-8") as f:
        data = json.load(f)

    results = data.get("results", [])
    if not results:
        print("ERROR: results.json contains no results.", file=sys.stderr)
        sys.exit(2)
    if listings:
        check_completeness(results, listings)

    for idx, r in enumerate(results):
        r["_uid"] = r.get("id") or f"__row{idx}"
        normalize(r)
    results.sort(key=sort_key)

    rankable = [r for r in results
                if r.get("status") == "revisado"
                and not r.get("disqualified") and r["_final"] is not None]
    top5_uids = {r["_uid"] for r in rankable[:5]}
    rankable_ids = {id(r) for r in rankable}
    if len(rankable) > 5 and rankable[4]["_final"] == rankable[5]["_final"]:
        rankable[4]["flags"].append("EMPATE TOP5")
        rankable[5]["flags"].append("EMPATE TOP5")

    wb = Workbook()
    ws = wb.active
    ws.title = "Ranking"
    style_header(ws, RANK_COLS)

    pos = 0
    for i, r in enumerate(results, start=2):
        reviewed = r.get("status") == "revisado"
        dq = bool(r.get("disqualified"))
        in_top5 = r["_uid"] in top5_uids
        if id(r) in rankable_ids:
            pos += 1
            position = pos
        else:
            position = ""
        s = r.get("scores") or {}
        row = [
            position,
            "⭐ TOP 5" if in_top5 else "",
            r.get("file", ""), r.get("student", ""), r.get("tool", ""),
            r.get("declared_launch_date", ""), r.get("verified_launch_date", ""),
            r.get("verification_source", ""), r.get("verification_confidence", ""),
            r.get("age_months", ""),
            ("SÍ" if dq else "NO") if reviewed else "",
            r.get("dq_reason", ""),
            s.get("poc", ""), s.get("impacto", ""), s.get("comunicacion", ""),
            r["_final"] if r["_final"] is not None else "",
            ", ".join(r.get("flags", [])),
            "REVISADO" if reviewed else "NO REVISADO",
            r.get("status_reason", ""),
        ]
        fill = (TOP5_FILL if in_top5
                else DQ_FILL if (reviewed and dq)
                else NOREV_FILL if not reviewed else None)
        for j, v in enumerate(row, 1):
            c = ws.cell(row=i, column=j, value=v)
            c.border = THIN
            if fill:
                c.fill = fill
            if j in (13, 14, 15, 16):
                c.number_format = "0.00"

    ws2 = wb.create_sheet("Detalle")
    style_header(ws2, DETAIL_COLS)
    for i, r in enumerate(results, start=2):
        j = r.get("justification") or {}
        row = [
            r.get("file", ""), r.get("tool", ""),
            f'{r.get("pages_read", "?")} / {r.get("pages_total", "?")}',
            j.get("poc", ""), j.get("impacto", ""), j.get("comunicacion", ""),
            r.get("evidence_notes", ""),
        ]
        for k, v in enumerate(row, 1):
            c = ws2.cell(row=i, column=k, value=v)
            c.border = THIN
            c.alignment = Alignment(wrap_text=True, vertical="top")

    meta = wb.create_sheet("Meta")
    meta["A1"], meta["B1"] = "Fecha de corrida", data.get("run_date", "")
    meta["A2"], meta["B2"] = "Carpeta Drive", data.get("folder_url", "")
    meta["A3"], meta["B3"] = "Regla de corte", ("> 4 meses desde lanzamiento verificado → 1.0; "
                                                "zona 3.5–4.5 meses lleva flag VERIFICAR FECHA")
    meta["A4"], meta["B4"] = "Ponderación", "PoC 50% · Impacto 25% · Comunicación 25%"
    meta["A5"], meta["B5"] = "Generado por", ("vigilancia-tech-review (sub-agentes Sonnet; "
                                              "revisión humana requerida para nota oficial)")
    meta.column_dimensions["A"].width = 20
    meta.column_dimensions["B"].width = 80

    wb.save(args[1])
    print(f"OK {args[1]} ({len(results)} filas, top5={len(top5_uids)})")


if __name__ == "__main__":
    main()
