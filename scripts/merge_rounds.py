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
- Hoja "Seguimiento": quién NO entregó en la ronda más reciente, quién
  volvió a mandar EXACTAMENTE el mismo archivo (no actualizó su caso) y cómo
  viene su nota, para que el equipo docente pueda buscar a esos estudiantes.
  Un estudiante solo puede reportarse como "SIN ENTREGA" si aparece en alguna
  ronda ANTERIOR: de otro modo no hay evidencia de que exista.
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
import unicodedata
from copy import copy

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
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
    keep = max(31 - len(kind) - 3 - 5, 1)  # "kind - " + "~hash"; never <=0
    return f"{kind} - {rnd[:keep]}~{h}"


def load_registry(wb):
    """[(round_full_name, {kind: title})] in arrival order.

    A master written before the registry existed has no _Rondas sheet but
    DOES have per-round sheets — backfill from the 'Ranking - <ronda>'
    titles so the first post-upgrade merge doesn't rebuild a Histórico
    that silently drops every prior round (convergence finding 2)."""
    if REGISTRY not in wb.sheetnames:
        out = []
        for title in wb.sheetnames:
            if not title.startswith("Ranking - "):
                continue
            # only a REAL round sheet backfills — a TA's hand-added
            # "Ranking - consolidado" must not become a phantom round
            # (round-2 finding 2)
            hdr = [c.value for c in wb[title][1]]
            if not {"Estudiante", "Nota final", "Estado"} <= set(hdr):
                print(f"AVISO: la hoja '{title}' parece de ronda pero no "
                      "tiene las columnas de make_excel — se ignora en la "
                      "reconstrucción del registro.", file=sys.stderr)
                continue
            rnd = title[len("Ranking - "):]
            # the pre-registry version TRUNCATED titles to 31 chars, and
            # each kind prefix has a different length — probe by prefix,
            # not by exact recomposition (round-2 finding 1, Meta orphan)
            titles = {}
            for k in KINDS:
                want = f"{k} - {rnd}"
                # exact match wins OUTRIGHT; a prefix hit only counts when
                # UNIQUE — a disjunctive next() bound whichever round came
                # first in tab order and could hand this entry ANOTHER
                # round's sheets (round-3 finding 1: silent grade loss)
                if want in wb.sheetnames:
                    titles[k] = want
                    continue
                cands = [t for t in wb.sheetnames if t.startswith(want)]
                if len(cands) == 1:
                    titles[k] = cands[0]
                elif cands:
                    print(f"AVISO: varias hojas coinciden con '{want}' "
                          f"({cands}) — ambigua, se omite del registro "
                          "reconstruido.", file=sys.stderr)
            out.append((rnd, titles))
        if out:
            print(f"AVISO: maestro sin hoja de registro '{REGISTRY}' — "
                  f"{len(out)} ronda(s) reconstruida(s) desde los títulos "
                  "de hoja existentes (posiblemente truncados a 31 chars "
                  "por la versión anterior).", file=sys.stderr)
        return out
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



THIN = Border(*[Side(style="thin", color="D1D5DB")] * 4)
SEG_FILL_SIN = PatternFill("solid", fgColor="FECACA")     # sin entrega
SEG_FILL_REP = PatternFill("solid", fgColor="FEF3C7")     # entrega repetida
SEG_FILL_OK = PatternFill("solid", fgColor="D1FAE5")      # entregó
SEG_FILL_DUDA = PatternFill("solid", fgColor="E5E7EB")    # sin confirmar
NOTA_MINIMA = 3.0      # nota aprobatoria del curso


def _sin_archivo(nombre):
    """True when the row is a placeholder the assembler synthesised because
    the folder held NO file at all.

    assemble_results.py emits these wrapped in parentheses and without an
    extension ("(carpeta sin archivos)", "(<estudiante> — inconsistencia de
    entradas)"). They must NOT count as a delivery: an empty Canvas folder is
    precisely a student who did not hand anything in, and counting it would
    clear the most common real case this sheet exists to catch. An unreadable
    FILE is different and does count — the student did the work.
    """
    n = str(nombre or "").strip()
    return n.startswith("(") and n.endswith(")")


def _norm_archivo(name):
    """Compare submissions by file name, accent-normalised and case-folded.

    NFC matters: a name that came out of a .zip is decomposed while the same
    name typed by a reviewer is composed, and a byte comparison would call
    two identical submissions different (and thus miss a repeat).
    """
    return unicodedata.normalize("NFC", str(name or "")).strip().casefold()


def build_seguimiento(wb, data, order, rounds, keyed_rounds=None,
                     huella_rounds=None):
    """Who did NOT hand in something new in the most recent round.

    Answers the teaching team's question: "¿qué estudiantes no han hecho
    nuevos envíos?", so a student who stops updating their case while doing
    badly can be contacted. Three states for the latest round:

      ENTREGÓ            appears in the round with a submission of its own
      ENTREGA REPETIDA   appears, and the submitted file is provably the one
                         already handed in before: identical content
                         fingerprint (or, when no round carries one, the same
                         file name — reported as "MISMO NOMBRE" instead,
                         because a student may legitimately keep the filename
                         while rewriting the deck)
      SIN ENTREGA        absent from the round, having submitted before
      SIN CONFIRMAR      absent, BUT the comparison crosses rounds that do not
                         share the stable 'Clave' identifier, so the absence
                         cannot be told apart from an identifier mismatch

    Two things this sheet refuses to do, because both would send a teacher to
    reproach a student who did the work:
      * assert an absence across rounds keyed differently (a pre-'Clave'
        round matched by display name splits students; on the real Semana 2
        ↔ Semana 3 master, 25 of 52 apparent absences were that split, and
        three of them had in fact submitted);
      * call a resubmission "repeated" on filename alone.
    A student who has never appeared in ANY round cannot be listed either:
    this workbook holds submissions, not the class roster, so their absence
    is unknown, not proven. That gap is stated on the sheet instead of guessed.
    """
    keyed_rounds = set(keyed_rounds or ())
    huella_rounds = set(huella_rounds or ())
    if "Seguimiento" in wb.sheetnames:
        del wb["Seguimiento"]
    if not rounds:
        return 0
    actual = rounds[-1]
    previas = rounds[:-1]

    filas = []
    for key in order:
        d = data[key]
        rows = d.get("rows") or {}
        vistos = d.get("vistos") or set()
        if actual not in rows and not any(r in rows for r in previas) \
                and actual not in vistos:
            continue                      # never seen anywhere: no evidence
        entregadas = [r for r in rounds if r in rows]
        # consecutive missed rounds ending at the most recent one
        seguidas = 0
        for r in reversed(rounds):
            if r in rows:
                break
            seguidas += 1
        notas = [(r, d["grades"][r]) for r in rounds if r in d["grades"]
                 and isinstance(d["grades"][r], (int, float))]
        ultima = notas[-1][1] if notas else None
        promedio = round(sum(n for _r, n in notas) / len(notas), 2) if notas else None

        # an absence is only assertable between rounds where THIS student
        # was identified by Clave; a name-keyed earlier row cannot be told
        # apart from an identifier mismatch
        carpeta_vacia = False
        comparables = [r for r in previas
                       if r in rows and rows[r].get("con_clave")]
        if actual in rows:
            cur = rows[actual]
            hue = (cur.get("huella") or "").strip()
            antes_h = {(rows[r].get("huella") or "").strip() for r in previas
                       if r in rows}
            antes_h.discard("")
            arch = _norm_archivo(cur.get("archivo"))
            antes_a = {_norm_archivo(rows[r].get("archivo")) for r in previas
                       if r in rows}
            antes_a.discard("")
            if hue and hue in antes_h:
                estado = "ENTREGA REPETIDA"      # identical content: proven
            elif hue and antes_h:
                # both sides carry a fingerprint and they differ: the file
                # DID change. Falling back to the filename here would accuse
                # a student who kept the name and rewrote the deck.
                estado = "ENTREGÓ"
            elif arch and arch in antes_a:
                estado = "MISMO NOMBRE"          # no fingerprint: verify
            else:
                estado = "ENTREGÓ"
        elif actual in vistos:
            # DIRECT evidence: the student has a folder in this round and it
            # holds no file. No cross-round identity matching is needed, so
            # this is confirmed no matter how the rounds are keyed.
            estado = "SIN ENTREGA"
            carpeta_vacia = True
        elif actual in keyed_rounds and comparables:
            estado = "SIN ENTREGA"
        else:
            estado = "SIN CONFIRMAR"

        alerta = []
        if estado == "SIN ENTREGA" and carpeta_vacia:
            alerta.append(f"su carpeta de {actual} está vacía: no entregó "
                          "nada esta ronda")
        elif estado == "SIN ENTREGA":
            alerta.append(f"no entregó en {actual}"
                          + (f" ({seguidas} rondas seguidas)" if seguidas > 1 else ""))
        elif estado == "SIN CONFIRMAR":
            falta = "la ronda actual no trae columna 'Clave'" \
                if actual not in keyed_rounds \
                else "sus rondas anteriores no traen columna 'Clave'"
            alerta.append("no aparece en " + actual + ", pero " + falta
                          + ", así que la ausencia NO está confirmada: puede "
                            "ser un desajuste de identificador. Verificar "
                            "antes de escribirle")
        elif estado == "ENTREGA REPETIDA":
            alerta.append("volvió a enviar un archivo de contenido idéntico: "
                          "el caso no avanzó")
        elif estado == "MISMO NOMBRE":
            alerta.append("reenvió un archivo con el mismo nombre que ya "
                          "había entregado; puede haberlo actualizado por "
                          "dentro, verificar el archivo antes de escribirle")
        if ultima is not None and ultima < NOTA_MINIMA:
            alerta.append(f"última nota {ultima:.2f}, por debajo de {NOTA_MINIMA:.1f}")
        filas.append({
            "key": key if isinstance(key, str) else f"{key[0]} ({key[1]})",
            "name": d["name"], "estado": estado,
            "entregadas": f"{len(entregadas)} de {len(rounds)}",
            "seguidas": seguidas, "ultima": ultima, "promedio": promedio,
            "ronda_ultima": notas[-1][0] if notas else "",
            "alerta": " · ".join(alerta),
        })

    # A split identity shows up twice: the old name-keyed half (SIN
    # CONFIRMAR) and the current Clave-keyed row (ENTREGÓ). Say so on the
    # doubtful row so nobody chases a student who did submit. This only
    # annotates; it never merges the rows or moves a grade.
    def _norm_nombre(x):
        x = unicodedata.normalize("NFKD", str(x or ""))
        x = "".join(ch for ch in x if not unicodedata.combining(ch)).casefold()
        x = re.sub(r"\(.*?\)", " ", x)
        return " ".join(sorted(w for w in re.split(r"[^a-z]+", x) if len(w) > 2))

    entregaron = {}
    for f in filas:
        if f["estado"] in ("ENTREGÓ", "MISMO NOMBRE", "ENTREGA REPETIDA"):
            entregaron.setdefault(_norm_nombre(f["name"]), []).append(f["key"])
    for f in filas:
        if f["estado"] != "SIN CONFIRMAR":
            continue
        otras = entregaron.get(_norm_nombre(f["name"]))
        if otras:
            f["alerta"] += (" · OJO: hay otra fila con el mismo nombre que SÍ "
                            f"entregó en {actual} (clave {otras[0]}); lo más "
                            "probable es que sean la misma persona y que el "
                            "identificador no cruce entre rondas")

    # the ones the team has to act on first
    prio = {"SIN ENTREGA": 0, "ENTREGA REPETIDA": 1, "MISMO NOMBRE": 2,
            "SIN CONFIRMAR": 3, "ENTREGÓ": 4}
    filas.sort(key=lambda f: (prio[f["estado"]],
                              f["ultima"] if f["ultima"] is not None else 9,
                              str.casefold(f["name"])))

    ws = wb.create_sheet("Seguimiento", 1)
    cols = [("Clave", 12), ("Estudiante", 30), (f"Estado en {actual}"[:31], 18),
            ("Rondas con entrega", 14), ("Rondas seguidas sin entregar", 14),
            ("Última nota", 11), ("Ronda de esa nota", 16), ("Promedio", 10),
            ("Por qué aparece aquí", 58)]
    for j, (h, w) in enumerate(cols, 1):
        c = ws.cell(row=1, column=j, value=h)
        c.fill, c.font = HEADER_FILL, HEADER_FONT
        c.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[get_column_letter(j)].width = w
    for i, f in enumerate(filas, 2):
        vals = [f["key"], f["name"], f["estado"], f["entregadas"],
                f["seguidas"] or "", f["ultima"], f["ronda_ultima"],
                f["promedio"], f["alerta"]]
        fill = {"SIN ENTREGA": SEG_FILL_SIN,
                "ENTREGA REPETIDA": SEG_FILL_REP,
                "MISMO NOMBRE": SEG_FILL_REP,
                "SIN CONFIRMAR": SEG_FILL_DUDA}.get(f["estado"], SEG_FILL_OK)
        for j, v in enumerate(vals, 1):
            c = ws.cell(row=i, column=j, value=v)
            c.border = THIN
            c.alignment = Alignment(wrap_text=(j == 9), vertical="top")
            if j in (6, 8) and isinstance(v, (int, float)):
                c.number_format = "0.00"
            if j == 3:
                c.fill = fill
    n = len(filas) + 2
    ws.cell(row=n + 1, column=1, value=(
        f"Cómo leer esta hoja. SIN ENTREGA: no aparece en '{actual}' y sus "
        "rondas anteriores comparten el identificador estable 'Clave', así "
        "que la ausencia sí está confirmada. SIN CONFIRMAR: no aparece, pero "
        "la comparación cruza rondas sin 'Clave' y la ausencia puede ser un "
        "desajuste de identificador, no una falta de entrega; verificar antes "
        "de escribirle. ENTREGA REPETIDA: el archivo entregado tiene "
        "contenido idéntico a uno anterior (huella de contenido), es decir el "
        "caso no avanzó. MISMO NOMBRE: coincide el nombre del archivo pero no "
        "hay huella para comparar el contenido; pudo haberlo actualizado por "
        "dentro. Un estudiante que NUNCA ha entregado no puede aparecer aquí, "
        "porque este libro solo contiene entregas: para esos hay que cruzar "
        "con el listado oficial del curso."))
    ws.cell(row=n + 1, column=1).alignment = Alignment(wrap_text=True,
                                                       vertical="top")
    ws.merge_cells(start_row=n + 1, start_column=1, end_row=n + 3,
                   end_column=9)
    ws.freeze_panes = "C2"
    cuenta = {}
    for f in filas:
        cuenta[f["estado"]] = cuenta.get(f["estado"], 0) + 1
    print("Seguimiento (ronda '" + actual + "'): "
          + ", ".join(f"{v} {k}" for k, v in sorted(cuenta.items())))
    if cuenta.get("SIN CONFIRMAR"):
        print(f"AVISO: {cuenta['SIN CONFIRMAR']} estudiante(s) quedaron en "
              "SIN CONFIRMAR porque la comparación cruza rondas sin columna "
              "'Clave'. NO son ausencias probadas; regenera las rondas viejas "
              "con el make_excel actual para poder confirmarlas.",
              file=sys.stderr)
    return len(filas)

def rebuild_historico(wb, registry):
    """Grade per student per round, arrival-ordered, keyed on the stable
    'Clave' column (falls back to the visible name for pre-Clave rounds)."""
    if "Histórico" in wb.sheetnames:
        del wb["Histórico"]
    rounds = [name for name, _t in registry]
    # key -> {"name": display, "grades": {round: final},
    #         "rows": {round: {"estado", "archivo"}}}
    data = {}
    order = []
    keyed, name_keyed = [], []
    keyed_rounds, huella_rounds = set(), set()
    for name, titles in registry:
        t = titles.get("Ranking")
        if not t or t not in wb.sheetnames:
            # a REGISTERED round whose sheet is gone is registry/workbook
            # divergence, not "no data" — its Histórico column will be
            # blank; say so instead of silently succeeding (finding 4)
            print(f"AVISO: la ronda registrada '{name}' no tiene su hoja "
                  f"'{t or '(sin título de Ranking registrado)'}' en el maestro (¿renombrada o borrada a mano?) — "
                  "su columna del Histórico quedará vacía.", file=sys.stderr)
            continue
        ws = wb[t]
        hdr = [c.value for c in ws[1]]
        try:
            i_est = hdr.index("Estudiante")
            i_fin = hdr.index("Nota final")
            i_st = hdr.index("Estado")
        except ValueError:
            print(f"AVISO: la hoja '{t}' de la ronda '{name}' no tiene las "
                  "columnas esperadas (Estudiante/Nota final/Estado) — su "
                  "columna del Histórico quedará vacía.", file=sys.stderr)
            continue
        i_key = hdr.index("Clave") if "Clave" in hdr else None
        i_arch = hdr.index("Archivo") if "Archivo" in hdr else None
        i_hue = hdr.index("Huella") if "Huella" in hdr else None
        if i_key is not None:
            keyed_rounds.add(name)
        if i_hue is not None:
            huella_rounds.add(name)
        (keyed if i_key is not None else name_keyed).append(name)
        # PRESENCE pass: a student who submitted something unreadable still
        # SUBMITTED. Reporting them as "sin entrega" would send a teacher to
        # scold someone who did the work, so presence is recorded from every
        # row of the round, not only the graded ones.
        # ONE pass: the identity of a row is decided once and used for both
        # its grade and its delivery record. Two passes could disagree — a
        # duplicate split into a tuple key by the grading side while the
        # presence side kept the plain key left one student invisible to the
        # Seguimiento sheet and cleared another who had stopped submitting.
        for row in ws.iter_rows(min_row=2, values_only=True):
            est = (row[i_est] or "").strip()
            key = ((row[i_key] or "").strip()
                   if i_key is not None else "") or est
            if not key:
                continue
            estado = row[i_st]
            if estado == "REEMPLAZADA":
                # a superseded duplicate is neither a grade nor a delivery,
                # and must not create a student row of its own
                continue
            # did THIS row carry a real Clave value? A round can have the
            # column while individual rows leave it empty (the partially
            # keyed Semana 2), and those rows fall back to the display name
            con_clave = bool(i_key is not None and (row[i_key] or "").strip())
            arch = (row[i_arch] or "").strip() if i_arch is not None else ""
            hue = (row[i_hue] or "").strip() if i_hue is not None else ""

            dup_label = None
            if estado == "REVISADO" and key in data \
                    and name in data[key]["grades"]:
                # two REVISADO rows in the SAME round collapsing onto one
                # key (blank Clave + shared display name) would silently
                # drop a grade — disambiguate with a TUPLE key, which can
                # never equal a real string Clave or display name (round-3
                # finding 3: an "X (2)" string key merged the duplicate
                # into a DIFFERENT real student), and label the extra row
                # so the two aren't visually identical.
                print(f"AVISO: dos filas REVISADO de la ronda '{name}' "
                      f"comparten la clave '{key}' — se separan en filas "
                      "distintas del Histórico; verificar la identidad de "
                      "esos estudiantes.", file=sys.stderr)
                base, n = key, 2
                while (base, n) in data and name in data[(base, n)]["grades"]:
                    n += 1
                key = (base, n)
                dup_label = f"{est} (fila duplicada {n})"
            if key not in data:
                data[key] = {"name": dup_label or est, "grades": {},
                             "rows": {}, "vistos": set()}
                order.append(key)
            data[key].setdefault("vistos", set())

            # Canvas made this student a folder, so we know they exist in
            # the round even when the folder turned out to be empty.
            data[key]["vistos"].add(name)
            if not _sin_archivo(arch):
                prev = data[key]["rows"].get(name)
                # the graded row wins the slot; an annex only fills a gap
                if prev is None or (estado == "REVISADO"
                                    and prev.get("estado") != "REVISADO"):
                    data[key]["rows"][name] = {
                        "estado": estado, "archivo": arch, "huella": hue,
                        "con_clave": con_clave}

            if estado == "REVISADO":
                data[key]["grades"][name] = row[i_fin]
                if dup_label:
                    data[key]["name"] = dup_label
                elif est:
                    data[key]["name"] = est   # latest round's name wins
    if keyed and name_keyed:
        # mixed keying splits every student whose Clave-keyed and name-keyed
        # rows don't coincide — visible only here, so warn here (finding 5)
        print("AVISO: el Histórico mezcla rondas CON columna 'Clave' "
              f"({', '.join(keyed)}) y SIN ella ({', '.join(name_keyed)}) — "
              "un mismo estudiante puede aparecer partido en dos filas. "
              "Regenera las rondas viejas con el make_excel actual para "
              "unificar.", file=sys.stderr)
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
    build_seguimiento(wb, data, order, rounds, keyed_rounds, huella_rounds)
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
    if rnd.endswith("'"):
        # Excel forbids sheet titles ENDING with an apostrophe (the title
        # always starts with the kind prefix, so a leading one is fine)
        print(f"ERROR: el nombre de ronda '{rnd}' no puede terminar con "
              "apóstrofo (regla de Excel para títulos de hoja).",
              file=sys.stderr)
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

    # Resolve round IDENTITY against reconstructed (pre-registry) names,
    # which are 31-char TRUNCATED remnants of the real round name
    # (round-2 findings 1 and 3):
    #  * hash-form titles round-trip exactly (the hash pins the full name)
    #    → same round: ADOPT rnd as the entry's name so refresh works;
    #  * a truncated-remnant prefix match is AMBIGUOUS (same round being
    #    refreshed, or a genuinely different round sharing the prefix) —
    #    proceeding would either duplicate the week or clobber another:
    #    refuse with both exits spelled out.
    if not any(name == rnd for name, _t in registry):
        resolved = []
        adopted = False
        adopt_old_titles = []
        for name, btitles in registry:
            rk = btitles.get("Ranking", "")
            if (not adopted and name != rnd and rk
                    and rk == sheet_title("Ranking", rnd)):
                # EXACT title match only: the reconstructed hash-form name
                # round-trips byte-identical. Casefolding here would make a
                # case-variant round name silently adopt-and-refresh instead
                # of refusing (the F3 guarantee). At most ONE entry may
                # adopt (round-3 finding 2: a duplicate adopt erased a
                # round and doubled the Histórico column).
                adopted = True
                adopt_old_titles = list(btitles.values())
                name = rnd
                print(f"AVISO: la ronda reconstruida '{rk[10:]}' es la "
                      f"misma que '{rnd}' (título idéntico) — se refresca.",
                      file=sys.stderr)
            elif (name != rnd and len(rk) == 31 and "~" not in rk
                    and rnd.startswith(name)):
                # Before refusing, use the Meta title to decide what the
                # Ranking truncation hid (Meta keeps 3 more name chars):
                # a complete Meta name, or an extended prefix rnd does NOT
                # match, PROVES this is a different round (round-3
                # finding 4: a legitimate new round was refused, and the
                # guidance would have overwritten the old round's grades).
                mt = btitles.get("Meta", "")
                mname = mt[len("Meta - "):] if mt else ""
                if mt and len(mt) < 31 and mname == rnd and not adopted:
                    # complete Meta name EQUALS rnd → PROVES the same round
                    # (round-4 finding: assuming "complete → different"
                    # dead-ended a 22/23-char round's own refresh and the
                    # clash message's advice then duplicated the week)
                    adopted = True
                    adopt_old_titles = list(btitles.values())
                    name = rnd
                    print(f"AVISO: la ronda reconstruida '{rk[10:]}' es la "
                          f"misma que '{rnd}' (nombre completo en la hoja "
                          "Meta) — se refresca.", file=sys.stderr)
                elif mt and len(mt) < 31:
                    pass          # complete name != rnd → different round
                elif mname and not rnd.startswith(mname):
                    pass          # extended prefix disproves identity
                else:
                    print(f"ERROR: el maestro contiene una ronda "
                          f"reconstruida '{name}' (título truncado por la "
                          "versión anterior) que coincide con el inicio de "
                          f"'{rnd}'. No es posible saber mecánicamente si "
                          "es la misma ronda. Si ES la misma, re-mergéala "
                          f"con --round=\"{name}\"; si es una ronda "
                          "DISTINTA, usa un nombre que no empiece igual, o "
                          "regenera el maestro desde cero con los Excel "
                          "por ronda. NO se modificó nada.", file=sys.stderr)
                    sys.exit(2)
            resolved.append((name, btitles))
        registry = resolved
        # the adopted entry's OLD sheets (plain-truncated titles) must be
        # replaced, not orphaned, when the new titles differ (Meta-adopt)
        for old_t in adopt_old_titles:
            if old_t and old_t.casefold() not in                     {t.casefold() for t in
                     (sheet_title(k, rnd) for k in KINDS)}:
                if old_t in wb.sheetnames:
                    del wb[old_t]
    if sum(1 for name, _t in registry if name == rnd) > 1:
        print(f"ERROR: el registro contiene la ronda '{rnd}' más de una "
              "vez — registro corrupto; repara la hoja _Rondas antes de "
              "continuar. NO se modificó nada.", file=sys.stderr)
        sys.exit(2)

    titles = {k: sheet_title(k, rnd) for k in KINDS}

    # refuse to touch a title that the registry attributes to ANOTHER round —
    # deleting it would destroy that round's history. Excel sheet titles are
    # CASE-INSENSITIVE-unique (openpyxl silently renames on collision), so
    # every comparison here must casefold (convergence finding 3).
    ours_cf = {t.casefold() for t in titles.values()}
    for other_name, other_titles in registry:
        if other_name == rnd:
            continue
        clash = ours_cf & {t.casefold() for t in other_titles.values()}
        if clash:
            print(f"ERROR: el título de hoja {sorted(clash)} ya pertenece a "
                  f"la ronda '{other_name}' — nombres de ronda demasiado "
                  "parecidos; elige un nombre distinto. NO se borró nada.",
                  file=sys.stderr)
            sys.exit(2)

    replaced = any(name == rnd for name, _t in registry)
    for kind in KINDS:
        t = titles[kind]
        for existing in list(wb.sheetnames):
            if existing.casefold() == t.casefold():
                del wb[existing]
        ws = copy_sheet(src[kind], wb, t)
        if ws.title != t:
            # openpyxl renamed the sheet — the registry would point at a
            # title that doesn't exist and the round would silently vanish
            print(f"ERROR: openpyxl renombró la hoja '{t}' a '{ws.title}' "
                  "(colisión de títulos no detectada) — abortando sin "
                  "guardar.", file=sys.stderr)
            sys.exit(2)

    if replaced:
        registry = [(n, titles if n == rnd else t) for n, t in registry]
    else:
        registry.append((rnd, titles))
    save_registry(wb, registry)

    n_students, rounds = rebuild_historico(wb, registry)

    # never truncate the irreplaceable master in place: stage next to it,
    # then atomically swap (os.replace) so a crash mid-write leaves the
    # previous master intact
    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    tmp.close()
    staged = longpath(master_path + ".tmp~")
    try:
        wb.save(tmp.name)
        shutil.copyfile(tmp.name, staged)
        os.replace(staged, longpath(master_path))
    except OSError as exc:
        # the staged copy holds student grades in the Drive-synced delivery
        # folder — never leave it behind; and a write failure is exit 2,
        # not the exit-1 "bad arguments" contract (round-2 finding 4)
        print(f"ERROR: no se pudo escribir el maestro: {exc}\n"
              "¿Está el archivo abierto en Excel? Ciérralo y reintenta. "
              "El maestro anterior quedó intacto.", file=sys.stderr)
        sys.exit(2)
    finally:
        for leftover in (tmp.name, staged):
            try:
                os.unlink(leftover)
            except FileNotFoundError:
                pass
            except OSError:
                if leftover == staged:
                    # the staged copy carries student grades in the
                    # Drive-synced folder — never leave it behind silently
                    print(f"AVISO: no se pudo borrar el temporal "
                          f"'{master_path}.tmp~' — bórralo a mano "
                          "(contiene notas de estudiantes).",
                          file=sys.stderr)
    print(f"OK maestro '{os.path.basename(master_path)}': ronda '{rnd}' "
          f"{'reemplazada' if replaced else 'agregada'} · rondas presentes: "
          f"{rounds} · estudiantes en Histórico: {n_students}")


if __name__ == "__main__":
    main()
