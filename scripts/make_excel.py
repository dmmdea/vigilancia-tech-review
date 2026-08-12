#!/usr/bin/env python3
"""Build the two-sheet results Excel from the orchestrator's results JSON.

Usage:
    python make_excel.py <results.json> <output.xlsx>

The final grade is computed HERE (single source of truth):
    final = 0.50*poc + 0.25*impacto + 0.25*comunicacion   (rounded to 0.01)
    disqualified -> final = 1.0
Ranking order: reviewed & not DQ (by final desc) -> DQ -> not reviewed.
Top 5 non-DQ rows are starred and highlighted.
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


def compute_final(r: dict) -> float | None:
    if r.get("disqualified"):
        return 1.0
    s = r.get("scores") or {}
    try:
        poc, imp, com = float(s["poc"]), float(s["impacto"]), float(s["comunicacion"])
    except (KeyError, TypeError, ValueError):
        return None
    for v in (poc, imp, com):
        if not 1.0 <= v <= 5.0:
            raise ValueError(f"score out of range 1.0-5.0 in {r.get('file')}: {v}")
    return round(0.50 * poc + 0.25 * imp + 0.25 * com, 2)


def sort_key(r: dict):
    reviewed = r.get("status", "revisado") == "revisado"
    dq = bool(r.get("disqualified"))
    final = r.get("_final")
    group = 0 if (reviewed and not dq) else (1 if reviewed else 2)
    return (group, -(final if final is not None else 0.0))


def style_header(ws, cols):
    for i, (name, width) in enumerate(cols, 1):
        c = ws.cell(row=1, column=i, value=name)
        c.fill, c.font = HEADER_FILL, HEADER_FONT
        c.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.freeze_panes = "A2"


def main() -> None:
    if len(sys.argv) != 3:
        print(__doc__, file=sys.stderr)
        sys.exit(1)
    with open(sys.argv[1], encoding="utf-8") as f:
        data = json.load(f)

    results = data.get("results", [])
    if not results:
        print("ERROR: results.json contains no results.", file=sys.stderr)
        sys.exit(2)

    for r in results:
        r["_final"] = compute_final(r)
    results.sort(key=sort_key)

    rankable = [r for r in results
                if r.get("status", "revisado") == "revisado"
                and not r.get("disqualified") and r["_final"] is not None]
    top5_files = {r["file"] for r in rankable[:5]}

    wb = Workbook()
    ws = wb.active
    ws.title = "Ranking"
    style_header(ws, RANK_COLS)

    pos = 0
    for i, r in enumerate(results, start=2):
        reviewed = r.get("status", "revisado") == "revisado"
        dq = bool(r.get("disqualified"))
        if reviewed and not dq and r["_final"] is not None:
            pos += 1
            position = pos
        else:
            position = ""
        s = r.get("scores") or {}
        row = [
            position,
            "⭐ TOP 5" if r["file"] in top5_files else "",
            r.get("file", ""), r.get("student", ""), r.get("tool", ""),
            r.get("declared_launch_date", ""), r.get("verified_launch_date", ""),
            r.get("verification_source", ""), r.get("verification_confidence", ""),
            r.get("age_months", ""),
            "SÍ" if dq else ("" if not reviewed else "NO"),
            r.get("dq_reason", ""),
            s.get("poc", ""), s.get("impacto", ""), s.get("comunicacion", ""),
            r["_final"] if r["_final"] is not None else "",
            ", ".join(r.get("flags", [])),
            "REVISADO" if reviewed else "NO REVISADO",
            r.get("status_reason", ""),
        ]
        fill = (TOP5_FILL if r["file"] in top5_files
                else DQ_FILL if dq
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
    just_rows = sorted(results, key=sort_key)
    for i, r in enumerate(just_rows, start=2):
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
    meta["A3"], meta["B3"] = "Regla de corte", "> 4 meses desde lanzamiento verificado → 1.0"
    meta["A4"], meta["B4"] = "Ponderación", "PoC 50% · Impacto 25% · Comunicación 25%"
    meta["A5"], meta["B5"] = "Generado por", "vigilancia-tech-review (sub-agentes Sonnet; revisión humana requerida para nota oficial)"
    meta.column_dimensions["A"].width = 20
    meta.column_dimensions["B"].width = 80

    wb.save(sys.argv[2])
    print(f"OK {sys.argv[2]} ({len(results)} filas, top5={len(top5_files)})")


if __name__ == "__main__":
    main()
