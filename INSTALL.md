# Instalación

Guía para el equipo docente del MBA de Uniandes. El skill es **agnóstico del
harness**: el mismo `SKILL.md` y los mismos scripts corren en Claude Code,
Codex (GPT) y Antigravity (Gemini). Lo único que cambia es cómo cada uno
despacha a los revisores; eso está en `references/<tu-plataforma>.md`.

Hay dos caminos. **El rápido es el primero: pegarle un prompt a tu agente y
que él lo haga.** El manual está más abajo por si algo falla o prefieres ver
cada paso.

---

## A. Que lo instale el agente (recomendado)

Copia y pega el bloque de tu plataforma. El agente clona el repo, instala lo
que falte (Python 3.10+, ffmpeg, LibreOffice, librerías) y termina
verificando con `preflight.py`.

### A.1 Claude Code

```
Instala el skill "vigilancia-tech-review" en esta máquina:

1. Clónalo desde https://github.com/dmmdea/vigilancia-tech-review en la
   carpeta de skills de Claude Code: en Windows
   %USERPROFILE%\.claude\skills\vigilancia-tech-review, en macOS o Linux
   ~/.claude/skills/vigilancia-tech-review. Si ya existe, haz git pull.
2. Entra a esa carpeta e instala las dependencias del sistema y de Python
   corriendo el instalador que trae el repo: en Windows
   "powershell -ExecutionPolicy Bypass -File install.ps1", en macOS o Linux
   "bash install.sh". El instalador se puede correr varias veces sin
   problema y no toca lo que ya esté instalado.
3. Cuando termine, corre "python scripts/preflight.py" y muéstrame la
   salida. No des la instalación por buena hasta que imprima
   "LISTO para correr".
4. Dime qué quedó pendiente, si algo quedó pendiente.

Al terminar, reinicia Claude Code para que el skill aparezca en /skills.
```

### A.2 Codex (GPT)

```
Instala el skill "vigilancia-tech-review" en esta máquina:

1. Clona https://github.com/dmmdea/vigilancia-tech-review en
   ~/skills/vigilancia-tech-review (si ya existe, git pull).
2. Entra a la carpeta y corre el instalador del repo: en Windows
   "powershell -ExecutionPolicy Bypass -File install.ps1", en macOS o Linux
   "bash install.sh".
3. Corre "python scripts/preflight.py" y muéstrame la salida completa. Debe
   decir "LISTO para correr".
4. Lee references/codex.md y resúmeme en tres líneas cómo se despacha un
   revisor en Codex, para confirmar que lo tienes claro antes de la primera
   corrida.
```

### A.3 Antigravity (Gemini)

```
Instala el skill "vigilancia-tech-review" en esta máquina:

1. Clona https://github.com/dmmdea/vigilancia-tech-review en
   ~/skills/vigilancia-tech-review (si ya existe, git pull).
2. Entra a la carpeta y corre el instalador del repo: en Windows
   "powershell -ExecutionPolicy Bypass -File install.ps1", en macOS o Linux
   "bash install.sh".
3. Corre "python scripts/preflight.py" y muéstrame la salida completa. Debe
   decir "LISTO para correr".
4. Lee references/antigravity.md y confírmame que entendiste dos cosas: que
   hay que rasterizar los PDFs antes de repartir a los revisores, y que a
   cada revisor se le dan RUTAS ABSOLUTAS.
```

### A.4 Y para correr una ronda, en cualquiera de los tres

```
Usa el skill vigilancia-tech-review para revisar las entregas de
Vigilancia Tecnológica que están en <RUTA DE LA CARPETA O DEL .ZIP>.
Sigue SKILL.md al pie de la letra, incluida la verificación adversarial de
fechas, y entrégame el Excel de la ronda y el maestro actualizado.
La ronda es "<Semana N>".
```

---

## B. Instalación manual

### B.1 Un solo comando

```bash
git clone https://github.com/dmmdea/vigilancia-tech-review.git
cd vigilancia-tech-review

# Windows
powershell -ExecutionPolicy Bypass -File install.ps1
# macOS / Linux
bash install.sh
```

El instalador revisa qué falta y lo instala (Python 3.10+, ffmpeg,
LibreOffice si no hay Office, y las librerías de Python), y termina corriendo
`preflight.py`. **No des la instalación por buena hasta ver
`LISTO para correr`.** Se puede volver a correr sin miedo: no toca lo que ya
esté instalado.

Si prefieres omitir LibreOffice porque ya tienes Microsoft Office:
`install.ps1 -SinLibreOffice` · `SIN_LIBREOFFICE=1 bash install.sh`

### B.2 Qué instala, y para qué

| Qué | Para qué | Cómo verificar |
|---|---|---|
| **Python 3.10+** | todos los scripts | `python --version` |
| `openpyxl`, `pypdf`, `pymupdf` | Excel, contar páginas, rasterizar | `pip list` |
| **LibreOffice** *o* **Microsoft Office** | `.pptx/.docx` → PDF | `soffice --version` |
| **Chrome o Edge** | `.html` → PDF | ya viene en Windows |
| **ffmpeg** | fotogramas de video, duración de audio | `ffmpeg -version` |
| *(opcional)* un **whisper** local | transcribir audio y video hablado | §5 |

`preflight.py` es la única verificación en la que hay que confiar: prueba los
backends de verdad, no la documentación.

### Windows: leer esto antes que nada
Las carpetas de Canvas más los nombres de archivo de los estudiantes pasan
de los 260 caracteres del `MAX_PATH` de Windows con facilidad. Los scripts ya
usan el prefijo `\\?\` internamente, pero **los conversores externos (Chrome,
ffmpeg, Office COM) no lo aceptan**. Por eso:

- Deja el `--matroot` corto: `--matroot=C:\Temp\vtr-mat` (el predeterminado ya
  es corto; no lo apuntes al escritorio ni a una carpeta de OneDrive).
- Escribe el Excel con un nombre corto (`res.xlsx`) y cópialo al nombre final
  al entregar.

### Si el instalador no puede con algo
Lo dice al final, en la lista "FALTA POR RESOLVER", y sale con código 1. Las
dos causas normales:
- **`winget` no está disponible** (Windows viejo o sin App Installer):
  instala Python, ffmpeg y LibreOffice a mano y vuelve a correr el script.
- **Instaló Python pero la terminal sigue sin verlo**: el `PATH` de una
  terminal ya abierta no cambia. Cierra la terminal, ábrela de nuevo y
  vuelve a correr el instalador.

---

## 2. Claude Code

El skill se instala como carpeta de skill y Claude lo invoca solo cuando la
tarea encaja.

```bash
# Windows (PowerShell)
git clone https://github.com/dmmdea/vigilancia-tech-review.git `
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
bash install.sh          # o install.ps1 en Windows
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
cd ~/skills/vigilancia-tech-review && bash install.sh
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

La política de nota cuando se incumple la regla de fecha se elige al generar el
Excel: `--dq-policy=penalty:0.5:3.0` (predeterminada) descuenta 0.5 a la nota de
rúbrica sin bajar de 3.0 a quien aprueba; `--dq-policy=cap:3.0` vuelve a la regla
anterior (descalificar y limitar la final a 3.0).

### 6.1 Rondas viejas sin huella de contenido

La columna `Huella` es la que permite demostrar que un estudiante reenvió el
mismo archivo en vez de actualizar su caso. Las rondas generadas antes de que
existiera esa columna no la traen, y la hoja `Seguimiento` tiene que degradar
a `MISMO NOMBRE (verificar)`. Si los archivos de esa ronda **siguen en
disco**, no hay que esperar a la siguiente:

```bash
python scripts/backfill_huella.py "<libro.xlsx>" \
  --hoja="Ranking - Semana 3" \
  --desde="<carpeta de Canvas o el .zip de la descarga>" \
  --asignacion=<idTarea> --aplicar
```

Sin `--aplicar` es un ensayo: dice qué haría y no escribe nada.

**`--asignacion` importa de verdad.** Una misma carpeta puede guardar dos
tareas de Canvas (los mismos estudiantes, distinto `idTarea`), y tomar la
huella de la ronda equivocada **inventa** una "entrega repetida" que nunca
ocurrió. Pasó en los datos reales: sin ese filtro, 14 estudiantes salían
acusados de reenviar el mismo archivo usando la huella de un archivo de otra
ronda. Si la fuente mezcla tareas y no le dices cuál, el script se niega.

El `idTarea` es la segunda mitad del nombre de la carpeta de Canvas:
`12345-467275 - Nombre Apellido - ...` → `--asignacion=467275`.

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
| El Excel sale sin la columna `Huella` | Es una ronda vieja generada antes de esta versión. Rellénala con `backfill_huella.py` (§6.1) si los archivos siguen en disco; si no, el `Seguimiento` degrada a `MISMO NOMBRE` y pide verificar a mano. |
| `backfill_huella.py` se niega diciendo que la fuente mezcla tareas | Es a propósito: pásale `--asignacion=<idTarea>` (§6.1). |
| Una opción mal escrita | Los scripts **rechazan** banderas desconocidas en vez de ignorarlas. Lee el mensaje: te dice cuál era la correcta. |
