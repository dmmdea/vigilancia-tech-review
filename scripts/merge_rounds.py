#!/usr/bin/env python3
"""Maestro multi-ronda: acumula las rondas de entregas en UN workbook sin
borrar jamás el histórico de rondas anteriores.

    python merge_rounds.py <maestro.xlsx> <ronda.xlsx> --round="Semana 2"

- Incorpora el Excel de la ronda (Ranking/Detalle/Meta de make_excel.py) al
  maestro como un trío de hojas propio de la ronda.
- Los títulos de hoja son DETERMINISTAS y a prueba de colisiones: Excel corta
  los títulos a 31 caracteres, así que dos rondas con nombres largos
  parecidos colapsarían en la misma hoja y una BORRARÍA a la otra (defecto
  real encontrado por revisión). Los nombres largos llevan un sufijo hash del
  nombre completo, y la hoja oculta "_Rondas" registra título↔nombre completo
  + orden de llegada. Si un título calculado ya pertenece a OTRA ronda según
  el registro, el script se niega (exit 2) en lugar de borrar historia.
- Re-entregar la MISMA ronda reemplaza SOLO su trío (refresh); las hojas de
  cualquier otra ronda no se tocan nunca.
- Hoja "Histórico": nota final por estudiante por ronda, en orden de llegada
  de las rondas. Se agrupa por la columna estable **Clave** (clave Canvas /
  id de carpeta) — el nombre visible del estudiante varía entre revisores
  (acentos, abreviaciones, códigos) y partiría el histórico; la clave no.
  El nombre mostrado es el de la ronda más reciente donde aparece.

Exit codes: 0 ok · 1 uso (args/round inválido) · 2 entrada ilegible, columnas
faltantes, o conflicto de títulos que borraría otra ronda.

MAX_PATH: escribe a un temporal corto y copia al destino con prefijo largo.
"""
import hashlib
import os
import re
import shutil
import sys
import tempfile
from copy import copy

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

for _s in (sys.stdout, sys.stderr):
    if _s in (sys.__stdout__, sys.__stderr__) and hasattr(_s, "reconfigure"):
        _s.reconfigure(encoding="utf-8")

HEADER_FILL = PatternFill("solid", fgColor="1F2937")
HEADER_FONT = Font(bold=True, color="FFFFFF")
KINDS = ("Ranking", "Detalle", "Meta")
REGISTRY = "_Rondas"
ILLEGAL = re.compile(r"[\\/?*\[\]:]")


def longpath(p):
    if os.name != "nt":
        return p
    ap = os.path.abspath(p)
    if ap.startswith("\\\\?\\"):
        return ap
    if ap.startswith("\\\\"):
        return "\\\\?\\UNC\\" + ap[2:]
    return "\\\\?\\" + ap


def sheet_title(kind, rnd):
    """Deterministic, <=31 chars, collision-proof across DIFFERENT rounds:
    long round names get a 4-hex hash of the FULL name, so truncation can
    never make two rounds share a title."""
    base = f"{kind} - {rnd}"
    if len(base) <= 31:
        return base
    h = hashlib.sha256(rnd.encode("utf-8")).hexdigest()[:4]
    keep = 31 - len(kind) - 3 - 5          # "kind - " + "~hash"
    return f"{kind} - {rnd[:keep]}~{h}"


def load_registry(wb):
    """[(round_full_name, {kind: title})] in arrival order."""
    if REGISTRY not in wb.sheetnames:
        return []
    out = []
    for row in wb[REGISTRY].iter_rows(min_row=2, values_only=True):
        if row and row[0]:
            out.append((row[0], {k: t for k, t in zip(KINDS, row[1:4]) if t}))
    return out


def save_registry(wb, entries):
    if REGISTRY in wb.sheetnames:
        del wb[REGISTRY]
    ws = wb.create_sheet(REGISTRY)
    ws.sheet_state = "hidden"
    ws.append(["Ronda", *KINDS])
    for name, titles in entries:
        ws.append([name] + [titles.get(k, "") for k in KINDS])


def copy_sheet(src_ws, dst_wb, title):
    ws = dst_wb.create_sheet(title=title)
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


def rebuild_historico(wb, registry):
    """Grade per student per round, arrival-ordered, keyed on the stable
    'Clave' column (falls back to the visible name for pre-Clave rounds)."""
    if "Histórico" in wb.sheetnames:
        del wb["Histórico"]
    rounds = [name for name, _t in registry]
    data = {}          # key -> {"name": display, "grades": {round: final}}
    order = []
    for name, titles in registry:
        t = titles.get("Ranking")
        if not t or t not in wb.sheetnames:
            continue
        ws = wb[t]
        hdr = [c.value for c in ws[1]]
        try:
            i_est = hdr.index("Estudiante")
            i_fin = hdr.index("Nota final")
            i_st = hdr.index("Estado")
        except ValueError:
            continue
        i_key = hdr.index("Clave") if "Clave" in hdr else None
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[i_st] != "REVISADO":
                continue
            est = (row[i_est] or "").strip()
            key = ((row[i_key] or "").strip() if i_key is not None else "") or est
            if not key:
                continue
            if key not in data:
                data[key] = {"name": est, "grades": {}}
                order.append(key)
            data[key]["grades"][name] = row[i_fin]
            if est:
                data[key]["name"] = est     # latest round's display name wins
    ws = wb.create_sheet("Histórico", 0)
    hdr = ["Estudiante"] + [f"Nota {r}" for r in rounds]
    for j, h in enumerate(hdr, 1):
        c = ws.cell(row=1, column=j, value=h)
        c.fill, c.font = HEADER_FILL, HEADER_FONT
        c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.column_dimensions["A"].width = 30
    for j in range(2, len(hdr) + 1):
        ws.column_dimensions[get_column_letter(j)].width = 14
    ordered = sorted(order, key=lambda k: str.casefold(data[k]["name"]))
    for i, key in enumerate(ordered, 2):
        ws.cell(row=i, column=1, value=data[key]["name"])
        for j, rname in enumerate(rounds, 2):
            v = data[key]["grades"].get(rname)
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
    if ILLEGAL.search(rnd):
        print(f"ERROR: el nombre de ronda '{rnd}' contiene caracteres no "
              "permitidos en títulos de hoja de Excel (\\ / ? * [ ] :) — "
              "usa otro nombre.", file=sys.stderr)
        sys.exit(1)
    master_path, round_path = args

    try:
        src = load_workbook(longpath(round_path))
    except Exception as e:
        print(f"ERROR: no se pudo leer el Excel de la ronda: {e}",
              file=sys.stderr)
        sys.exit(2)
    for needed in KINDS:
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
    if "Clave" not in hdr:
        print("AVISO: la ronda no trae columna 'Clave' — el Histórico usará "
              "el nombre visible del estudiante, que puede variar entre "
              "rondas.", file=sys.stderr)

    if os.path.exists(longpath(master_path)):
        wb = load_workbook(longpath(master_path))
    else:
        wb = Workbook()
        wb.remove(wb.active)

    registry = load_registry(wb)
    titles = {k: sheet_title(k, rnd) for k in KINDS}

    # refuse to touch a title that the registry attributes to ANOTHER round —
    # deleting it would destroy that round's history
    for other_name, other_titles in registry:
        if other_name == rnd:
            continue
        clash = set(titles.values()) & set(other_titles.values())
        if clash:
            print(f"ERROR: el título de hoja {sorted(clash)} ya pertenece a "
                  f"la ronda '{other_name}' — nombres de ronda demasiado "
                  "parecidos; elige un nombre distinto. NO se borró nada.",
                  file=sys.stderr)
            sys.exit(2)

    replaced = any(name == rnd for name, _t in registry)
    for kind in KINDS:
        t = titles[kind]
        if t in wb.sheetnames:
            del wb[t]
        copy_sheet(src[kind], wb, t)

    if replaced:
        registry = [(n, titles if n == rnd else t) for n, t in registry]
    else:
        registry.append((rnd, titles))
    save_registry(wb, registry)

    n_students, rounds = rebuild_historico(wb, registry)

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
