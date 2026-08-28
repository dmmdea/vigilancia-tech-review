#!/usr/bin/env python3
"""Rellena la columna 'Huella' de una ronda YA generada, a partir de los
archivos originales que todavía estén en disco.

Para qué: la huella de contenido es lo que permite demostrar que un
estudiante reenvió EXACTAMENTE el mismo archivo en vez de actualizar su
caso. `local_list.py` la calcula al listar, así que las rondas generadas
antes de esa versión no la traen y la hoja 'Seguimiento' tiene que degradar
a "MISMO NOMBRE (verificar)". Si los archivos de esa ronda siguen en disco
—la carpeta de Canvas, o el .zip de la descarga masiva— no hay que esperar
a la siguiente ronda: se recalcula la huella y la comparación queda exacta
desde ya.

    python backfill_huella.py <libro.xlsx> --hoja="Ranking - Semana 3" \\
        --desde="<carpeta o .zip>" [--desde=...] [--aplicar]

Sin `--aplicar` no escribe nada: informa qué haría (ensayo).

`--asignacion=<idTarea>` es OBLIGATORIA cuando la fuente contiene más de una
tarea de Canvas. Una misma carpeta puede guardar la Semana 2 y la Semana 3
(los mismos estudiantes, distinto idTarea), y tomar la huella de la ronda
equivocada INVENTA una "entrega repetida" que nunca ocurrió. Comprobado en
los datos reales: sin este filtro, 14 estudiantes aparecían acusados de
reenviar el mismo archivo usando la huella de un archivo de OTRA ronda. Si la
fuente mezcla tareas y no se indica cuál, el script se niega (exit 2).

CÓMO EMPAREJA (y por qué se niega a adivinar): la clave de un archivo es
(clave del estudiante, nombre del archivo). La clave del estudiante sale del
nombre de la carpeta de Canvas — "<idEstudiante>-<idTarea> - Nombre - fecha"
— igual que en local_list.py. Un nombre de archivo suelto NO alcanza: dos
estudiantes distintos entregan "presentacion.pptx" el mismo día. Cuando la
fila del Excel no tiene Clave, solo se rellena si ese nombre de archivo es
ÚNICO en toda la fuente; si no, se reporta como ambigua y se deja vacía.
Una huella equivocada acusaría a un estudiante de no haber actualizado su
trabajo, así que aquí no se adivina.

Exit: 0 ok · 1 uso · 2 el libro o la hoja no se pueden leer.
"""
import io
import os
import re
import sys
import unicodedata
import zipfile

from openpyxl import load_workbook

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from local_list import huella, longpath          # misma huella que la corrida

for _s in (sys.stdout, sys.stderr):
    if _s in (sys.__stdout__, sys.__stderr__) and hasattr(_s, "reconfigure"):
        _s.reconfigure(encoding="utf-8")

CARPETA_CANVAS = re.compile(r"^(\d+)-(\d+)\s*-\s*")
SKIP = {"desktop.ini", ".ds_store", "thumbs.db"}


def nfc(x):
    return unicodedata.normalize("NFC", str(x or "").strip())


def _indexar_directorio(raiz, por_clave, por_nombre, asignacion=None,
                        vistas=None):
    for d in sorted(os.listdir(longpath(raiz))):
        ruta = os.path.join(raiz, d)
        if not os.path.isdir(longpath(ruta)):
            continue
        m = CARPETA_CANVAS.match(d)
        if m and vistas is not None:
            vistas.add(m.group(2))
        if m and asignacion and m.group(2) != asignacion:
            continue          # carpeta de OTRA tarea de Canvas
        clave = m.group(1) if m else None
        for base, _dirs, archivos in os.walk(longpath(ruta)):
            for a in archivos:
                if a.lower() in SKIP:
                    continue
                p = os.path.join(base, a)
                try:
                    h = huella(p, os.path.getsize(p))
                except OSError:
                    h = None
                if not h:
                    continue
                if clave:
                    por_clave.setdefault((clave, nfc(a)), set()).add(h)
                por_nombre.setdefault(nfc(a), set()).add(h)


def _indexar_zip(ruta, por_clave, por_nombre, asignacion=None, vistas=None):
    import tempfile
    with zipfile.ZipFile(longpath(ruta)) as z:
        for info in z.infolist():
            if info.is_dir():
                continue
            partes = info.filename.split("/")
            nombre = partes[-1]
            if not nombre or nombre.lower() in SKIP:
                continue
            m = CARPETA_CANVAS.match(partes[0]) if len(partes) > 1 else None
            if m and vistas is not None:
                vistas.add(m.group(2))
            if m and asignacion and m.group(2) != asignacion:
                continue      # carpeta de OTRA tarea de Canvas
            clave = m.group(1) if m else None
            # la huella se calcula sobre el archivo real, así que se extrae a
            # un temporal y se borra enseguida (un .zip de clase pesa cientos
            # de MB y no cabe entero en disco dos veces)
            with tempfile.NamedTemporaryFile(delete=False) as tmp:
                destino = tmp.name
            try:
                with z.open(info) as src, open(destino, "wb") as dst:
                    while True:
                        b = src.read(1 << 20)
                        if not b:
                            break
                        dst.write(b)
                h = huella(destino, os.path.getsize(destino))
            finally:
                try:
                    os.unlink(destino)
                except OSError:
                    pass
            if not h:
                continue
            if clave:
                por_clave.setdefault((clave, nfc(nombre)), set()).add(h)
            por_nombre.setdefault(nfc(nombre), set()).add(h)


def construir_indice(fuentes, asignacion=None):
    por_clave, por_nombre, vistas = {}, {}, set()
    for f in fuentes:
        if not os.path.exists(longpath(f)):
            print(f"ERROR: no existe la fuente '{f}'.", file=sys.stderr)
            sys.exit(2)
        if f.lower().endswith(".zip"):
            _indexar_zip(f, por_clave, por_nombre, asignacion, vistas)
        elif os.path.isdir(longpath(f)):
            _indexar_directorio(f, por_clave, por_nombre, asignacion, vistas)
        else:
            print(f"ERROR: '{f}' no es carpeta ni .zip.", file=sys.stderr)
            sys.exit(2)
    if len(vistas) > 1 and not asignacion:
        print(f"ERROR: la fuente mezcla {len(vistas)} tareas de Canvas "
              f"({sorted(vistas)}). Tomar la huella de la ronda equivocada "
              "inventaría una 'entrega repetida' que nunca ocurrió. Repite "
              "indicando cuál es con --asignacion=<idTarea>.", file=sys.stderr)
        sys.exit(2)
    return por_clave, por_nombre


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    hoja = None
    fuentes = []
    aplicar = "--aplicar" in sys.argv[1:]
    asignacion = None
    forzar = "--forzar" in sys.argv[1:]
    for a in sys.argv[1:]:
        if a.startswith("--hoja="):
            hoja = a.split("=", 1)[1]
        elif a.startswith("--desde="):
            fuentes.append(a.split("=", 1)[1])
        elif a.startswith("--asignacion="):
            asignacion = a.split("=", 1)[1].strip()
        elif (a.startswith("--") and a not in ("--aplicar", "--forzar")
                and not a.startswith("--asignacion=")):
            print(f"ERROR: opción desconocida {a}", file=sys.stderr)
            sys.exit(1)
    if len(args) != 1 or not hoja or not fuentes:
        print(__doc__, file=sys.stderr)
        sys.exit(1)
    libro = args[0]

    try:
        wb = load_workbook(longpath(libro))
    except Exception as e:
        print(f"ERROR: no se pudo abrir '{libro}': {e}", file=sys.stderr)
        sys.exit(2)
    if hoja not in wb.sheetnames:
        print(f"ERROR: el libro no tiene la hoja '{hoja}'. Hojas: "
              f"{wb.sheetnames}", file=sys.stderr)
        sys.exit(2)
    ws = wb[hoja]
    hdr = [c.value for c in ws[1]]
    if "Archivo" not in hdr:
        print(f"ERROR: la hoja '{hoja}' no tiene columna 'Archivo'.",
              file=sys.stderr)
        sys.exit(2)
    i_arch = hdr.index("Archivo") + 1
    i_clave = hdr.index("Clave") + 1 if "Clave" in hdr else None
    if "Huella" in hdr:
        i_hue = hdr.index("Huella") + 1
    else:
        # una ronda anterior a la columna: se agrega al final
        i_hue = len(hdr) + 1
        c = ws.cell(row=1, column=i_hue, value="Huella")
        base = ws.cell(row=1, column=1)
        from copy import copy as _copy
        c.font, c.fill, c.alignment = (_copy(base.font), _copy(base.fill),
                                       _copy(base.alignment))
        print(f"NOTA: la hoja no tenía columna 'Huella'; se agrega al final.")

    por_clave, por_nombre = construir_indice(fuentes, asignacion)
    print(f"fuentes indexadas: {len(por_clave)} pares (clave, archivo), "
          f"{len(por_nombre)} nombres distintos")

    puestas = ya = sin_fuente = ambiguas = 0
    por_origen = {"clave+archivo": 0, "archivo único": 0}
    detalle_amb, detalle_sin = [], []
    for fila in range(2, ws.max_row + 1):
        arch = ws.cell(row=fila, column=i_arch).value
        if not arch:
            continue
        actual = ws.cell(row=fila, column=i_hue).value
        if actual and not forzar:
            ya += 1
            continue
        nombre = nfc(arch)
        clave = (str(ws.cell(row=fila, column=i_clave).value or "").strip()
                 if i_clave else "")
        cand = por_clave.get((clave, nombre)) if clave else None
        origen = "clave+archivo"
        if not cand:
            # sin Clave en la fila: solo vale si el nombre es único en la fuente
            cand = por_nombre.get(nombre)
            origen = "archivo único"
        if not cand:
            sin_fuente += 1
            detalle_sin.append(str(arch)[:52])
            continue
        if len(cand) > 1:
            ambiguas += 1
            detalle_amb.append(f"{str(arch)[:44]} ({len(cand)} contenidos)")
            continue
        if aplicar:
            ws.cell(row=fila, column=i_hue, value=next(iter(cand)))
        puestas += 1
        por_origen[origen] += 1

    print(f"filas: {puestas} con huella {'puesta' if aplicar else 'lista'} "
          f"[por clave+archivo: {por_origen['clave+archivo']}, "
          f"por nombre único: {por_origen['archivo único']}] · "
          f"{ya} ya tenían · {sin_fuente} sin archivo en la fuente · "
          f"{ambiguas} ambiguas")
    for d in detalle_amb[:8]:
        print(f"  AMBIGUA (se deja vacía): {d}", file=sys.stderr)
    for d in detalle_sin[:8]:
        print(f"  sin fuente: {d}", file=sys.stderr)
    if aplicar and puestas:
        tmp = libro + ".tmp~"
        wb.save(longpath(tmp))
        os.replace(longpath(tmp), longpath(libro))
        print(f"OK escrito '{os.path.basename(libro)}'")
    elif not aplicar:
        print("ENSAYO: no se escribió nada. Repite con --aplicar.")


if __name__ == "__main__":
    main()
