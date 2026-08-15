# vigilancia-tech-review (documentación en español)

Skill agéntico que revisa, califica y clasifica las presentaciones estudiantiles de la
actividad del MBA de Uniandes **"Vigilancia Tecnológica: IA de vanguardia"** — desde
una carpeta compartida de Google Drive por enlace, o desde una carpeta local (montaje
de Google Drive Desktop / descarga masiva de Canvas). Sin credenciales de Google.

Funciona en cualquier harness agéntico con visión: **Claude Code (Claude)**,
**Codex CLI (GPT)** y **Antigravity (Gemini)** tienen adaptadores dedicados en
`references/`; cualquier otro harness sigue la lista de capacidades de `SKILL.md`.

## Qué hace, por estudiante

1. **Inventaría** cada entrega (`scripts/local_list.py` o `scripts/drive_list.py`) —
   detecta carpetas por estudiante (nomenclatura Canvas con fechas en español),
   entregas duplicadas (gana la más reciente) y archivos de evidencia adicionales.
2. **Convierte todo formato a algo revisable** (`scripts/prepare_materials.py`):
   pptx/docx/html → PDF; imágenes tal cual; xlsx → volcado de celdas; video →
   fotogramas + transcripción; zip → extraído y re-enrutado. **El formato nunca es
   motivo para saltarse a un estudiante ni para descalificarlo.**
3. **Despacha un revisor de IA con contexto limpio por estudiante** que LEE todas las
   páginas visualmente, verifica por web la fecha real de lanzamiento de la
   herramienta, aplica el filtro de exclusión (>4 meses → 1.0), y califica
   PoC (50%) · Impacto (25%) · Comunicación (25%) citando slides concretas.
4. **Valida cada revisión mecánicamente** (`scripts/validate_review.py`): cobertura de
   páginas/materiales, rangos de nota, formatos estrictos de campos, vocabulario
   cerrado de flags. Un fallo se reintenta UNA vez; si persiste, la fila queda
   NO REVISADO con motivo. Hay spot-checks de honestidad en ambas rutas de revisión.
5. **Ensambla y reconcilia** (`scripts/assemble_results.py`): una fila por archivo
   entregado, una fila calificada por estudiante, y reconciliación entre estudiantes
   que presentaron la misma herramienta (fechas divergentes → VERIFICAR FECHA).
6. **Genera el Excel** (`scripts/make_excel.py`): hojas **Ranking** (ordenado, top-5
   con estrella, DQ en rojo, no-revisado en gris), **Detalle** (justificaciones con
   citas de slides) y **Meta**. La nota final se calcula en un único lugar.

## Reglas duras

- Todo archivo entregado aparece en el Excel — calificado, descalificado o
  NO REVISADO con motivo real. Nada se omite en silencio.
- Las fechas de lanzamiento se verifican por búsqueda web; nunca se confía solo en lo
  declarado. Verificación de baja confianza → nunca descalifica, flag VERIFICAR FECHA.
- Banda fronteriza 3.5–4.5 meses → siempre flag VERIFICAR FECHA.
- En una re-entrega se califica la ÚLTIMA versión, pero la evidencia adjunta solo al
  envío anterior acompaña la revisión (flag EVIDENCIA DE ENVIO ANTERIOR INCLUIDA).
- Los puntajes del revisor nunca se editan; la normalización de campos solo reubica
  contenido fuera de formato. Ante dudas, flag — deciden los humanos.
- Los archivos de estudiantes y los resultados nunca entran a un repositorio git ni a
  artefactos públicos.
- **Todo artefacto de cara a docentes/estudiantes sale en español.**
- **La nota oficial requiere revisión humana** — el skill produce una preselección con
  evidencia citada, no un veredicto.

## Requisitos

- Python 3.10+ con `openpyxl` y `pypdf` (`pip install -r requirements.txt`);
  `pymupdf` solo para harnesses sin visión nativa de páginas PDF.
- Backend de conversión: LibreOffice (`soffice`) o Microsoft Office (COM en Windows);
  Chrome/Edge headless para HTML; `ffmpeg` para videos.
- Para origen Drive: carpeta compartida como "cualquiera con el enlace".

## Nota para Windows (MAX_PATH)

Los nombres de carpeta de Canvas + nombres de archivo superan con frecuencia los 260
caracteres. Los scripts incluidos aplican el prefijo `\\?\` en toda E/S propia y
mantienen cortas las rutas generadas (raíz corta `%TEMP%\vtr-mat`), porque Chrome,
ffmpeg y COM de Office no aceptan el prefijo. Si escribes un helper propio, replica el
tratamiento `longpath()` — sin él, la E/S falla en silencio archivo por archivo y
parece que "el estudiante no entregó".
