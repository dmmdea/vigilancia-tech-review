# vigilancia-tech-review v2 — run plan (2026-08-15)

Nightshift run ledger for the `feat/harness-agnostic-v2` branch. Every request
from Daniel in this session is captured here; nothing ships until each item is
DONE or explicitly deferred with a reason.

## Daniel's requests (verbatim intent, in order received)

| # | Request | Status |
|---|---|---|
| R1 | Adversarial review of the skill's first production run; upgrades driven by findings | findings confirmed (below); fixes in progress |
| R2 | Make the skill agent/harness/model **agnostic** | in progress |
| R3 | Optimize for **Gemini in Antigravity** and **GPT in Codex**, like it is for Claude in Claude Code | in progress (references/) |
| R4 | Results Excel now lives in the **parent folder** `Vigilancia Tecnologica` (not Semana 2) — deliver there | pending (delivery step) |
| R5 | When everything is done, **apply the needed corrections to the output Excel** (regenerate with fixes) | pending (after ship) |
| R6 | Build under **nightshift** discipline | active |
| R7 | Full **clean-ship** ritual (worktree, semgrep, named specialist review, evidence, PR, ledger) | active — branch `feat/harness-agnostic-v2` in own worktree |
| R8 | **Full Spanish support** — the class is dictated in Spanish in every sense | pending (see item G) |
| R9 | Keep all requests saved in a plan — miss nothing | this file |
| R10 | Full documentation checkpoint — impeccable, beautiful docs | pending (item N) |
| R11 | Repo branded **Universidad de los Andes** | pending (item O) |
| R12 | Custom license: public repo, made as part of TA job for Uniandes; only Uniandes teachers/TAs/personnel may modify/improve; others download/test for academic purposes only. Recommend license type + adjust | pending (item P) |
| R13 | **Merge and deploy when all is green**; quality bar: reliable to a Uniandes-MBA standard — authorization granted in-conversation | active |
| R14 | Skill seamless and **frictionless for the operator** | preflight.py + quickstart (item R) |

Standing constraints from earlier in the session: every student reviewed
regardless of file format (fairness rule); vision-first, convert only what the
file reader can't open; student files never in git/public artifacts.

## Adversarial-review findings (data-verified against the 2026-08-14 run)

CONFIRMED defects → fixes:
- F1. Loose reviewer schema → prose in date fields (5 rows), polluted student
  fields (8), tool >70ch (45). → `validate_review.py` strict formats +
  normalization; templates §6 strict field rules. DONE (pending review)
- F2. Flag sprawl: 104 distinct flags, 99 used once. → controlled vocabulary
  (11 canonical) + `observations` field. DONE (pending review)
- F3. Pass-2 bundle reviews had ZERO spot-checks (incl. the #1-ranked
  student). → SKILL.md: spot-check every ~4th review across BOTH passes. TODO
- F4. Orchestration glue lived only in session scratchpad (incl. hardcoded
  folder IDs for one student's resubmission). → `build_bundles.py` (generic
  carry-forward rule) + `assemble_results.py` in the skill. DONE (pending review)
- F5. No same-tool cross-student reconciliation (dates happened to agree this
  run; nothing enforced it). → reconciliation pass in `assemble_results.py`
  (>1 month divergence → VERIFICAR FECHA both rows). DONE (pending review)
- F6. Spot-check false positive on 1-page decks (no "slide N" to cite). →
  1-page rule in assembler + templates. DONE (pending review)

VERIFIED non-problems (recorded so nobody "fixes" them):
- No pass-1 vs pass-2 scoring drift (means 3.617 vs 3.638, non-DQ rows).
- DQ rule application: 11/11 correct; border band 6/6 flagged; baja-confidence
  2/2 correct.
- Same-tool date verification agreed across students this run (emergent).

## Work items

- [x] A. Port pilot-run fixes into repo worktree (local_list, prepare_materials,
       SKILL.md fairness rule, bundle template)
- [x] B. New scripts: validate_review (module+CLI), pdf_to_images,
       build_bundles (generic carry-forward), assemble_results (reconciliation)
- [x] C. Templates: strict formats, controlled flags, observations,
       harness-neutral wording
- [x] D. references/claude-code.md · references/antigravity.md ·
       references/codex.md (R3)
- [x] E. SKILL.md: capability-based core (R2) — dispatch/validation/spot-check
       sections harness-neutral; platform table pointing at references/;
       sequential-fallback for harnesses without subagents; F3 fix
- [x] F. requirements.txt: add optional pymupdf note
- [x] G. Spanish support (R8): audit every human-facing string is Spanish
       (Excel ✓ headers/flags/reasons; templates ✓; script stderr partially
       English → fix); README.es.md; SKILL.md rule "todo artefacto de cara al
       docente/estudiante sale en español"; verify Spanish-locale parsing
       (Canvas months ✓, accents ✓, ñ ✓)
- [x] H. Review gate (3 rounds: semgrep clean on changed files; round-1 two specialists 39 findings all fixed; round-2 13 findings all fixed w/ red tests; round-3 convergence in flight): semgrep (PYTHONUTF8=1, exit-code discipline) →
       code-reviewer + silent-failure-hunter (validator/assembler have many
       fallback paths)
- [x] I. Evidence (pipeline reproduced pilot exactly + 9 mutation gates red/green + rasterize smoke on real PDF): run the full pipeline against the real Semana 2 folder in
       the worktree; byte-compare listing/manifest vs pilot; mutation-test the
       new gates (validator date/flag rules, reconciliation)
- [ ] J. Ship: PR on dmmdea/vigilancia-tech-review (account check per write);
       hand back or merge per authorization
- [ ] K. Sync installed copy at ~/.claude/skills/vigilancia-tech-review
- [ ] L. R4+R5: regenerate corrected Excel from pilot reviews via the NEW
       pipeline (normalization + reconciliation applied), deliver to parent
       folder `Vigilancia Tecnologica`, remove stale copy if any
- [ ] M. clean-ship ledger line + final report (inventory: branch, worktree,
       PR, deliveries)
- [x] N. Documentation checkpoint (R10): README.md + README.es.md rewritten for
       v2 (architecture, scripts table, per-platform quickstart, mermaid flow),
       consistent with SKILL.md; badges; TOC
- [x] O. Uniandes branding (R11): title block, course context (Reto Integrador
       1 – Tecnología de Información, MBA, Facultad de Administración), tasteful
       text branding in both READMEs
- [x] P. License (R12): source-available custom license — public to view/
       download/test for academic non-commercial purposes; modification
       restricted to Uniandes teachers/TAs/staff; bilingual LICENSE.md; README
       license sections; flag to Daniel that work-for-hire ownership should be
       confirmed with Uniandes legal (not legal advice)
- [x] R. Operator frictionlessness (preflight.py live 8/8 OK) (R14): scripts/preflight.py one-command
       environment check (deps, backends, Spanish verdicts, exit 0/1); wired
       as step 0 of SKILL.md
- [x] Q. Evidence addendum: mutation tests red/green log (reconciliation bug
       found by mutation 1 and fixed: norm_tool 4→2 words)

## Round 2 — feedback de la primera clase (R15-R18, 2026-08-15)

Recibidos tras la primera clase con el skill en uso real. Se ejecutan DESPUÉS
de cerrar lo pendiente (test Antigravity + ship de feat/live-harness-validation).

| # | Request (verbatim intent) | Diseño previsto |
|---|---|---|
| R15 | Gestionar las dobles entregas para que no se pierdan en la revisión y se tome la entrega más reciente (también cubre carpeta duplicada) | Endurecer detección de duplicados más allá del canvas_key: nombre de estudiante normalizado entre carpetas; VERSIONES dentro de una misma carpeta (v1/FINAL → gana la más reciente, la anterior queda superseded, no "evidencia"); Excel inequívoco sobre cuál fila es la calificada |
| R16 | Entregas con múltiples archivos: los archivos extra deben aparecer como SÍ revisados en el reporte | Nuevo estado "REVISADO (ANEXO)" (color propio) en make_excel + assembler para los archivos leídos dentro de la revisión integral — hoy dicen NO REVISADO con razón explicativa y se lee mal |
| R17 | Filtro adicional SOLO sobre la ppt/pdf entregada: indicio de qué tan obvio es el uso sin filtro de IA (los estudiantes deben revisar/filtrar/mejorar manualmente) | Nuevo campo del revisor `indicio_ia` (1-5 + evidencia citada: frases de chatbot sin editar, plantilla genérica, artefactos de generación, estructura stock) — columna advisory en el Excel, NUNCA componente de la nota |
| R18 | Columna adicional con feedback "sugerido" por estudiante (base interna del equipo docente; útil tenerla a mano aunque solo se entregue a quienes fallaron) | Nuevo campo `feedback_sugerido` (2-4 frases constructivas en español) en templates + schema + validador + columna en el Excel |

## Round 2b — feedback adicional (R19-R20, 2026-08-15)

| # | Request | Diseño previsto |
|---|---|---|
| R19 | Un top-5 llevaba >4 meses de lanzamiento: el gate de fechas y la verificación NO bastaron — mejorarlo seriamente. NO tocar el ranking ya entregado; corregir para la próxima corrida | Causa raíz identificada (fila Notion 3.5: el revisor DETECTÓ que la funcionalidad demostrada era de feb-2026 (5.7 meses) pero verificó la fecha de la ETIQUETA declarada y no descalificó). Fix doble: (a) regla endurecida en templates — la fecha que gobierna es SIEMPRE la de la funcionalidad DEMOSTRADA; un rebranding/versión nueva de una función vieja hereda la fecha vieja; (b) paso mecánico nuevo: verificación adversarial independiente de fechas sobre TODO candidato top-N y borderline antes de entregar (template dedicado, framing de refutación: "demuestra que esta función es MÁS VIEJA de lo declarado"), desacuerdo → VERIFICAR FECHA + nota, nunca silencioso |
| R20 | Uso multi-ronda: refrescar la salida con rondas adicionales de entregas SIN borrar el histórico de rondas anteriores | `scripts/merge_rounds.py`: workbook maestro con hojas por ronda (Ranking-<ronda>, Detalle-<ronda>, Meta-<ronda>) + hoja "Histórico" (nota final por estudiante por ronda); re-entrega de una ronda REEMPLAZA solo sus propias hojas, jamás las de otras rondas; SKILL.md entrega round-aware |

Nota: la re-entrega del Excel S2 con los nuevos estados queda CANCELADA por
instrucción de Daniel ("no modificar el ranking ya entregado") — S2 entra al
maestro histórico tal cual en la próxima entrega.
