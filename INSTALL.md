# Instalación

Guía para el equipo docente del MBA de Uniandes. El skill es **agnóstico del
harness**: el mismo `SKILL.md` y los mismos scripts corren en Claude Code,
Codex (GPT) y Antigravity (Gemini). Lo único que cambia es cómo cada uno
despacha a los revisores; eso está en `references/<tu-plataforma>.md`.

> English speakers: this is the install guide; every section below has the
> exact commands. The skill itself and all teacher-facing output are in
> Spanish because the course is taught in Spanish.

---

## 1. Requisitos (iguales en las tres plataformas)

| Qué | Para qué | Cómo verificar |
|---|---|---|
| **Python 3.10+** | todos los scripts | `python --version` |
| **`pip install -r requirements.txt`** | `openpyxl` (Excel), `pypdf` (contar páginas), `pymupdf` (rasterizar) | ver abajo |
| **LibreOffice** *o* **Microsoft Office** | `.pptx/.docx` → PDF | `soffice --version` |
| **Chrome o Edge** | `.html` → PDF | ya viene en Windows |
| **ffmpeg** | fotogramas de video, duración de audio | `ffmpeg -version` |
| *(opcional)* un **whisper** local | transcribir audio y video hablado | ver §5 |

```bash
git clone https://github.com/dmmdea/vigilancia-tech-review.git
cd vigilancia-tech-review
pip install -r requirements.txt
python scripts/preflight.py     # dice OK/FALTA por cada dependencia
```

`preflight.py` es la única verificación que hay que creerle: revisa los
backends de verdad, no la documentación. **No arranques una corrida hasta
que imprima `LISTO para correr`.**

### Windows: leer esto antes que nada
Las carpetas de Canvas más los nombres de archivo de los estudiantes pasan
de los 260 caracteres del `MAX_PATH` de Windows con facilidad. Los scripts ya
usan el prefijo `\\?\` internamente, pero **los conversores externos (Chrome,
ffmpeg, Office COM) no lo aceptan**. Por eso:

- Deja el `--matroot` corto: `--matroot=C:\Temp\vtr-mat` (el predeterminado ya
  es corto; no lo apuntes al escritorio ni a una carpeta de OneDrive).
- Escribe el Excel con un nombre corto (`res.xlsx`) y cópialo al nombre final
  al entregar.

---

## 2. Claude Code

El skill se instala como carpeta de skill y Claude lo invoca solo cuando la
tarea encaja.

```bash
# Windows (PowerShell o Git Bash)
git clone https://github.com/dmmdea/vigilancia-tech-review.git ^
  "$env:USERPROFILE\.claude\skills\vigilancia-tech-review"

# macOS / Linux
git clone https://github.com/dmmdea/vigilancia-tech-review.git \
  ~/.claude/skills/vigilancia-tech-review
```

Reinicia Claude Code y verifica que aparece:

```
/skills            # debe listar vigilancia-tech-review
```

Para usarlo, basta pedirlo en lenguaje natural con la carpeta de entregas:

```
Revisa las entregas de Vigilancia Tecnológica que están en
G:\...\Vigilancia Tecnologica\Semana 4
```

Adaptador: [`references/claude-code.md`](references/claude-code.md). Claude
Code lee PDFs por páginas de forma nativa, así que **no** hace falta
`--rasterize`.

**Actualizar:** `git -C ~/.claude/skills/vigilancia-tech-review pull`

---

## 3. Codex (GPT)

Codex no tiene carpeta de skills: se le pasa el `SKILL.md` como instrucciones
y él ejecuta los scripts.

```bash
git clone https://github.com/dmmdea/vigilancia-tech-review.git ~/skills/vigilancia-tech-review
cd ~/skills/vigilancia-tech-review
pip install -r requirements.txt
python scripts/preflight.py
```

Arranca la corrida dándole el SKILL.md y la carpeta de entregas:

```bash
codex exec -C ~/skills/vigilancia-tech-review \
  "Sigue SKILL.md al pie de la letra para revisar las entregas en
   '<ruta de la carpeta de Canvas>'. Lee también references/codex.md."
```

Cada revisor es un `codex exec` nuevo (contexto limpio). Ese comando está
verificado en codex-cli 0.147.0 y vive en
[`references/codex.md`](references/codex.md):

```bash
codex exec -s read-only --skip-git-repo-check -C "<matroot>" \
  -m "<modelo con visión de tu allowlist>" -c tools.web_search=true \
  --output-schema "<skill>/templates/review_schema.json" \
  -o "<work>/reviews/<folder_id>.json" - < "<work>/prompts/<folder_id>.md"
```

Tres cosas que sí importan en Codex:
1. **Rasteriza**: `prepare_materials.py <work> --rasterize` y luego
   `build_bundles.py <work> --images --all`. Sin eso el revisor recibe rutas
   de PDF que no puede abrir.
2. **Fija el modelo con `-m`** y confirma que las búsquedas web ocurrieron de
   verdad (el log imprime `web search: ...` y `verification_source` debe ser
   una URL real). No confíes solo en la bandera.
3. **Nunca uses `--ephemeral`** en una corrida de calificación: borra la
   sesión de `~/.codex/sessions/`, que es el único rastro de auditoría si un
   estudiante reclama su nota.

---

## 4. Antigravity (Gemini)

```bash
git clone https://github.com/dmmdea/vigilancia-tech-review.git ~/skills/vigilancia-tech-review
cd ~/skills/vigilancia-tech-review && pip install -r requirements.txt && python scripts/preflight.py
```

Antigravity 2.8.1 **no trae CLI**: las revisiones se despachan desde la
conversación del hub. Pega el contenido de `SKILL.md` (o pídele que lo lea) y
dale **rutas absolutas** tanto de los materiales como del JSON de salida —
"esta carpeta" es ambiguo dentro del hub.

Como en Codex, hay que rasterizar (`--rasterize` + `--images --all`), porque
las herramientas de archivo de Antigravity leen texto, no páginas de PDF.

**Trampa de arranque:** si lanzas Antigravity desde un proceso que tenga
`ELECTRON_RUN_AS_NODE=1` en el entorno (cualquier extensión de VS Code, y
algunos shells de CI), el ejecutable sale con código 0 al instante, sin
ventana y sin línea de log. Limpia esa variable antes de lanzarlo.

Adaptador: [`references/antigravity.md`](references/antigravity.md).

---

## 5. Audio y video hablado (opcional pero recomendado)

Un `.mp3` no se puede "mirar": si contiene una presentación hablada, el
revisor **no puede calificarla** sin transcripción y la marca
`EVIDENCIA NO LEGIBLE`. Si el audio **no** es habla (música o sonido generado
por la herramienta que el estudiante evalúa), no hace falta transcribirlo: el
archivo cuenta como evidencia propia y el bundle ya se lo explica al revisor.

`prepare_materials.py` imprime al final la lista de audios y videos sin
transcripción, con su identificador indexado. Transcríbelos con cualquier
whisper local y vuelve a correr:

```bash
python scripts/prepare_materials.py <work> \
  --transcript="<folder_id>#1=C:\ruta\transcripcion.txt"
```

El índice `#1` es **obligatorio cuando la carpeta tiene varios medios** (un
estudiante entregó un deck y tres pistas .mp3): sin él el script se niega a
adivinar a cuál pertenece la transcripción, en vez de ponerle palabras a un
archivo que nadie leyó.

---

## 6. Entrega de resultados

La entrega va al **maestro multi-ronda**, que nunca borra rondas anteriores:

```bash
python scripts/merge_rounds.py \
  "<carpeta Drive>\Resultados-Vigilancia-Tecnologica-MAESTRO.xlsx" \
  "<work>\res.xlsx" --round="Semana 4"
```

Eso agrega el trío de hojas de la ronda y reconstruye dos hojas transversales:

- **`Histórico`** — nota final por estudiante por ronda.
- **`Seguimiento`** — quién **no** entregó nada nuevo, quién reenvió un
  archivo idéntico, y la nota al lado, para buscar a quien deja de actualizar
  su caso y va mal. Lee la leyenda al pie de esa hoja: distingue una ausencia
  confirmada (`SIN ENTREGA`) de una que solo cruza rondas con identificadores
  distintos (`SIN CONFIRMAR`, que **no** es prueba de nada).

La política de nota para descalificadas se elige al generar el Excel:
`--dq-policy=cap:3.0` (predeterminada) conserva visible la nota de rúbrica y
limita la final a 3.0.

---

## 7. Reglas que no se negocian

Están en `SKILL.md`, pero conviene que todo el equipo las tenga presentes:

- **Todo archivo entregado aparece en el Excel**, calificado o con una razón
  escrita. Nada se descarta en silencio.
- **La nota oficial la pone un humano.** El skill produce candidatos con
  evidencia citada y una lista corta; no es un veredicto.
- **Las fechas de lanzamiento se verifican por búsqueda web**, nunca se le
  cree a la diapositiva. Manda la capacidad DEMOSTRADA, no la etiqueta de
  versión.
- **Los archivos de los estudiantes y los resultados no entran a ningún
  repositorio** ni a artefactos públicos.

## 8. Problemas frecuentes

| Síntoma | Causa y arreglo |
|---|---|
| `canvas_key` vacío para todos | El formato de fecha del export de Canvas cambió. Ya se soportan `11_58` y `1152`; si aparece otro, avisa: una clave vacía deja sin `Clave` toda la ronda y apaga la detección de duplicados. |
| Un revisor correcto no pasa la compuerta | Nombres de archivo con tilde: el .zip los entrega en NFD y el modelo responde en NFC. Ya se normaliza; si vuelve a pasar, reporta el nombre exacto. |
| `--only` sale con código 2 | Estás fusionando contra un `materials.json` que no existe. Corre primero **sin** `--only`. |
| El Excel sale sin la columna `Huella` | Es una ronda vieja generada antes de esta versión. No es error: el `Seguimiento` degrada a `MISMO NOMBRE` y pide verificar a mano. |
| Una opción mal escrita | Los scripts **rechazan** banderas desconocidas en vez de ignorarlas. Lee el mensaje: te dice cuál era la correcta. |
