#!/usr/bin/env python3
"""Maestro multi-ronda: acumula las rondas de entregas en UN workbook sin
borrar jamás el histórico de rondas anteriores.

    python merge_rounds.py <maestro.xlsx> <ronda.xlsx> --round="Semana 2"

- Toma el Excel de la ronda recién generada por make_excel.py (Ranking,
  Detalle, Meta) y lo incorpora al maestro como hojas
  "Ranking - <ronda>", "Detalle - <ronda>", "Meta - <ronda>".
- Si el maestro no existe, lo crea.
- Re-entregar la MISMA ronda reemplaza SOLO sus tres hojas (refresh);
  las hojas de cualquier otra ronda no se tocan nunca.
- Mantiene/reconstruye la hoja "Histórico": nota final por estudiante por
  ronda (una fila por estudiante, una columna por ronda, orden de llegada),
  reconstruida desde las hojas Ranking presentes — nunca desde memoria.

Exit codes: 0 ok · 1 uso · 2 entrada ilegible / hoja Ranking de la ronda sin
las columnas esperadas.

MAX_PATH: escribe primero a un temporal corto y copia al destino con prefijo
largo — el maestro vive en una carpeta Drive profunda.
"""
import os
import shutil
import sys
import tempfile
from copy import copy

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

for _s in (sys.stdout, sys.stderr):
    if _s in (sys.__stdout__, sys.__stderr__) and hasattr(_s, "reconfigure"):
        _s.reconfigure(encoding="utf-8")

HEADER_FILL = PatternFill("solid", fgColor="1F2937")
HEADER_FONT = Font(bold=True, color="FFFFFF")


def longpath(p):
    if os.name != "nt":
        return p
    ap = os.path.abspath(p)
    if ap.startswith("\\\\?\\"):
        return ap
    if ap.startswith("\\\\"):
        return "\\\\?\\UNC\\" + ap[2:]
    return "\\\\?\\" + ap


def copy_sheet(src_ws, dst_wb, title):
    """Cell-by-cell copy (values + basic style) — openpyxl cannot move sheets
    across workbooks natively."""
    ws = dst_wb.create_sheet(title=title[:31])
    for row in src_ws.iter_rows():
        for cell in row:
            c = ws.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                c.font = copy(cell.font)
                c.fill = copy(cell.fill)
                c.border = copy(cell.border)
                c.alignment = copy(cell.alignment)
                c.number_format = cell.number_format
    for col, dim in src_ws.column_dimensions.items():
        if dim.width:
            ws.column_dimensions[col].width = dim.width
    if src_ws.freeze_panes:
        ws.freeze_panes = src_ws.freeze_panes
    return ws


def rebuild_historico(wb):
    """Nota final por estudiante por ronda, SIEMPRE reconstruida desde las
    hojas Ranking presentes (la verdad vive en las hojas, no en un estado)."""
    if "Histórico" in wb.sheetnames:
        del wb["Histórico"]
    rounds = [n[len("Ranking - "):] for n in wb.sheetnames
              if n.startswith("Ranking - ")]
    data = {}          # student -> {round: final}
    order = []
    for rnd in rounds:
        ws = wb[f"Ranking - {rnd}"]
        hdr = [c.value for c in ws[1]]
        try:
            i_est = hdr.index("Estudiante")
            i_fin = hdr.index("Nota final")
            i_st = hdr.index("Estado")
        except ValueError:
            continue
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[i_st] != "REVISADO":
                continue
            est = (row[i_est] or "").strip()
            if not est:
                continue
            if est not in data:
                data[est] = {}
                order.append(est)
            data[est][rnd] = row[i_fin]
    ws = wb.create_sheet("Histórico", 0)
    hdr = ["Estudiante"] + [f"Nota {r}" for r in rounds]
    for j, h in enumerate(hdr, 1):
        c = ws.cell(row=1, column=j, value=h)
        c.fill, c.font = HEADER_FILL, HEADER_FONT
        c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.column_dimensions["A"].width = 30
    for j in range(2, len(hdr) + 1):
        ws.column_dimensions[chr(64 + j) if j <= 26 else "A"].width = 14
    for i, est in enumerate(sorted(order, key=str.casefold), 2):
        ws.cell(row=i, column=1, value=est)
        for j, rnd in enumerate(rounds, 2):
            v = data[est].get(rnd)
            if v is not None:
                c = ws.cell(row=i, column=j, value=v)
                c.number_format = "0.00"
    ws.freeze_panes = "B2"
    return len(order), rounds


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    rnd = None
    for a in sys.argv[1:]:
        if a.startswith("--round="):
            rnd = a.split("=", 1)[1].strip()
    if len(args) != 2 or not rnd:
        print(__doc__, file=sys.stderr)
        sys.exit(1)
    master_path, round_path = args

    try:
        src = load_workbook(longpath(round_path))
    except Exception as e:
        print(f"ERROR: no se pudo leer el Excel de la ronda: {e}",
              file=sys.stderr)
        sys.exit(2)
    for needed in ("Ranking", "Detalle", "Meta"):
        if needed not in src.sheetnames:
            print(f"ERROR: al Excel de la ronda le falta la hoja '{needed}'.",
                  file=sys.stderr)
            sys.exit(2)
    hdr = [c.value for c in src["Ranking"][1]]
    for col in ("Estudiante", "Nota final", "Estado"):
        if col not in hdr:
            print(f"ERROR: la hoja Ranking de la ronda no tiene la columna "
                  f"'{col}' — ¿se generó con make_excel.py?", file=sys.stderr)
            sys.exit(2)

    if os.path.exists(longpath(master_path)):
        wb = load_workbook(longpath(master_path))
    else:
        wb = Workbook()
        wb.remove(wb.active)

    # refresh = replace ONLY this round's sheets; other rounds untouched
    replaced = False
    for kind in ("Ranking", "Detalle", "Meta"):
        title = f"{kind} - {rnd}"[:31]
        if title in wb.sheetnames:
            del wb[title]
            replaced = True
        copy_sheet(src[kind], wb, title)

    n_students, rounds = rebuild_historico(wb)

    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    tmp.close()
    wb.save(tmp.name)
    shutil.copyfile(tmp.name, longpath(master_path))
    os.unlink(tmp.name)
    print(f"OK maestro '{os.path.basename(master_path)}': ronda '{rnd}' "
          f"{'reemplazada' if replaced else 'agregada'} · rondas presentes: "
          f"{rounds} · estudiantes en Histórico: {n_students}")


if __name__ == "__main__":
    main()
