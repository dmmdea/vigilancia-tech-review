---
name: vigilancia-tech-review
description: Use when the teaching team needs to review, score, and rank MBA "Vigilancia Tecnológica" student presentations — from a link-shared Google Drive folder OR a local/Drive-Desktop folder of Canvas submissions. Reads EVERY submitted file whatever its format (pptx, pdf, docx, html, png, mp4, xlsx, zip) with one fresh-context AI reviewer per student, web-verifies each tool's launch date, disqualifies tools older than 4 months, scores PoC/Impacto/Comunicación (1.0–5.0, weighted 50/25/25), and produces a three-sheet Excel ranking with the top-5 candidates for human TA review. Harness-agnostic — adapters for Claude Code, Codex (GPT), and Antigravity (Gemini) in references/. Triggers: "revisar presentaciones vigilancia tecnológica", "calificar las ppt de los estudiantes", "ranking vigilancia tecnológica", "escoger las mejores presentaciones".
---

# vigilancia-tech-review

Reviews every student submission for the Uniandes MBA activity **Vigilancia Tecnológica:
IA de vanguardia** and delivers a ranked Excel for the human TAs. The skill never assigns
the official grade — it produces evidence-cited candidate scores and a top-5 shortlist.

**Rubric:** Prueba de concepto 50% · Análisis de impacto 25% · Comunicación 25%.
Exclusion filter: tool launched **more than 4 months** before the run date → automatic
1.0 (verified by web search, not by trusting the deck). General-purpose tools (ChatGPT,
Gemini, Copilot, Claude…) are invalid unless the subject is a specific recently launched
feature. Scale: 1.0–5.0.

## THE FAIRNESS RULE THAT OVERRIDES CONVENIENCE

**Every student who submitted something gets a real review.** File format is NEVER a
reason to skip a student, and NEVER a reason to disqualify. A `.docx`, a screenshot, an
HTML page, or a video is a submission — render it and grade the CONTENT on the same
rubric. "Formato no soportado" is a bug in this skill, not a student's problem.
The only legitimate NO REVISADO reasons are: a genuinely unreadable/corrupt file, a
superseded duplicate submission, or a file that was read as supplementary evidence
inside another row's review (say so explicitly in the reason).

## LANGUAGE RULE — the class runs in Spanish

Every artifact a teacher or student may see leaves this skill **in Spanish**: the Excel
(headers, flags, status reasons, notes), the reviewer justifications, and the run
summary you give the teaching team. Reviewer prompt templates are Spanish and stay
Spanish. Inputs arrive in Spanish too — filenames with accents/ñ, Canvas timestamps
with Spanish month names — and the bundled scripts already parse them; never "sanitize"
Spanish out of names. (Talking to the operator about the run follows the operator's
language; the deliverables themselves are Spanish.)

## Capability requirements (harness-agnostic)

This skill assumes an agentic harness that can:

1. **See images** (its file/image tool renders PNG/JPG visually).
2. **Read PDF pages visually** — natively (Claude Code's Read), or via the bundled
   rasterizer `scripts/pdf_to_images.py` (PyMuPDF) that turns each page into a PNG.
3. **Search the web** — for launch-date verification. If the session has no web
   search, every review runs with `verification_confidence: "baja"` (never DQ) and
   the operator is told dates were not verified.
4. **Run Python 3.10+** with `openpyxl` + `pypdf` (plus `pymupdf` when rasterizing).
5. **Dispatch fresh-context sub-reviewers** — one per student, a mid-tier
   vision-capable model. A harness with no subagents reviews **sequentially in its
   own loop**, dropping/summarizing the previous student's pages between students so
   grading stays independent.
6. **Drive conversion backends**: slides→PDF (LibreOffice or PowerPoint COM),
   docx→PDF (Word COM or LibreOffice), html→PDF (headless Chrome/Edge), video
   keyframes (ffmpeg). `prepare_materials.py` autodetects what is installed.

**Platform adapters** — read yours before dispatching anything:

| Harness | Adapter |
|---|---|
| Claude Code (Claude) | `references/claude-code.md` |
| Codex CLI (GPT) | `references/codex.md` |
| Antigravity (Gemini) | `references/antigravity.md` |
| Anything else | follow the capability list above; the adapters show the shape |

For a Drive source: folder shared as **"anyone with the link"**. No credentials.

## Inputs to collect from the user

1. **Where the submissions are** — a Drive folder URL, or a local path (Google Drive
   Desktop mount / unzipped Canvas bulk download). Ask if not given.
2. **Where to deliver the Excel** — a local Drive-synced folder path. Confirm it once;
   deliveries go there every run afterwards.

## Procedure

`SKILL_DIR` = this skill's directory. `WORK` = a fresh working directory.
`RUN_DATE` = today (YYYY-MM-DD). Student files stay out of any git repo.

### 0. Preflight — one command, before anything

```bash
python "$SKILL_DIR/scripts/preflight.py" [--rasterize]
```

Verifies every dependency and conversion backend with a Spanish verdict per
item; exit 1 = fix what it names before touching student files. Pass
`--rasterize` on harnesses without native PDF-page vision. Never discover a
missing backend halfway through 75 students.

> **Windows MAX_PATH — read before anything else.** Student folder names plus filenames
> routinely exceed 260 chars, and Python does NOT auto-apply the `\\?\` prefix:
> file I/O then fails *silently per-file*, which looks exactly like "the student didn't
> submit". The bundled scripts already handle this. If you write your own helper, use
> the same `longpath()` treatment, and keep every generated path SHORT — put converted
> materials under a short root (default `%TEMP%\vtr-mat`), never under a deep session
> scratchpad, because Chrome/ffmpeg/Office COM do not accept `\\?\` paths.

### 1. Inventory the submissions

**Local folder (Drive Desktop mount, Canvas bulk download):**

```bash
python "$SKILL_DIR/scripts/local_list.py" "<folder>" "$WORK"
```

Auto-detects NESTED (one subfolder per student, Canvas naming) vs FLAT. Writes
`listing.json`, `sub-listings/*.json`, `manifest.json`, and `review_plan.json`
(primary submission + supplementary evidence + superseded duplicates per student).
Exit 2 = folder unreadable · 3 = nothing found. Review any `NOTE ...` lines on
stderr — those are ambiguous primary-file picks a human should confirm.

**Link-shared Drive folder:**

```bash
python "$SKILL_DIR/scripts/drive_list.py" "<folder-url>" > "$WORK/listing.json"
```

Exit 2 → not public/network failed, stop. Exit 3 → Drive's layout changed; STOP, never
proceed on a partial listing. List each subfolder one level deep and include its files;
a folder nested deeper gets a `NO REVISADO` row naming it. Then download each entry:

```bash
python "$SKILL_DIR/scripts/drive_download.py" "<id>" "$WORK/decks/<name>__<id>.<ext>"
python "$SKILL_DIR/scripts/drive_download.py" "<id>" "$WORK/decks/<name>__<id>.pptx" --kind=gslides
```

Exit 2/3 → `NO REVISADO` row for THAT file with the stderr reason. **Exit 4 (quota)
→ STOP the whole run** and retry in 30–60 min. Backstop: 3 consecutive "HTML page with
no download form" failures = throttling, stop too.

**Then synthesize the plan from the downloads**: run
`python "$SKILL_DIR/scripts/local_list.py" "$WORK/decks" "$WORK"` over the download
folder (flat shape) — the rest of the pipeline (steps 2–6) consumes its
`review_plan.json` exactly as in the local-folder path.

### 2. Make every file readable — whatever its format

```bash
python "$SKILL_DIR/scripts/prepare_materials.py" "$WORK" [--matroot=<short dir>]
```

Renders each submitted file into something a reviewer can consume, converting ONLY
where the harness's file tool cannot open the format (never "convert" a PNG):

| Submitted | Becomes | Why |
|---|---|---|
| `.pdf` | as-is | vision (rasterize per adapter if needed) |
| `.pptx .ppt .odp` | PDF | slide decks aren't readable directly |
| `.docx .doc .rtf` | PDF (Word COM → LibreOffice) | keeps embedded screenshots = the PoC evidence |
| `.html .htm` | PDF (headless Chrome) | keeps the rendering |
| `.png .jpg …` | passthrough | vision reads images directly |
| `.xlsx .xls .csv` | `.txt` cell dump | data reads better as data |
| `.mp4 .mov …` | keyframes (+ transcript) | the reviewer's own vision judges the frames |
| `.zip` | extracted, contents re-routed | — |
| `.py .txt .md` | passthrough | text |

Writes `materials.json`. It prints any video lacking a transcript; transcribe those
(a local whisper is ideal — mechanical work) and re-run with
`--transcript=<folder_id>=<file.txt>`. Re-runs with `--only=` merge, never clobber.
Fix every `ERROR` item it reports before moving on — each one is a student at risk of
losing credit for evidence they did submit.

**Harness without native PDF-page vision** (see your adapter): run step 2 with
`--rasterize` — it invokes `pdf_to_images.py` on every produced PDF and records
`page_images` (+ `truncated_from` past the page cap) on each item — then pass
`--images --all` to `build_bundles.py` below so EVERY student (single-deck ones
included) gets image instructions. Without `--rasterize`, `--images` has nothing
to point at and silently degrades to PDF instructions the reviewer cannot follow.

If a separate slides→PDF conversion step produced a `convert_results.json`
(`{folder_id: {pdf, pages}}`), leave it in `$WORK` — `build_bundles.py` uses it
to include decks whose folders aren't in `materials.json`.

### 3. Build the review targets

```bash
python "$SKILL_DIR/scripts/build_bundles.py" "$WORK" [--images]
```

Emits `bundles.json`: one multi-format target per student who needs one — no slide
deck at all, deck + supplementary evidence, or evidence carried forward from a
superseded duplicate submission (generic rule: evidence *kinds* present in the earlier
upload but missing from the final one ride along, flagged
`EVIDENCIA DE ENVIO ANTERIOR INCLUIDA`). Students with a single plain deck use the
simple template instead.

### 4. Dispatch one reviewer per student — fresh context, then GATE

Templates: `templates/reviewer-prompt.md` (single deck) ·
`templates/bundle-reviewer-prompt.md` (multi-format; feed it each bundle's
`materials_block`). Dispatch per your platform adapter, ~4 concurrent.

**The fairness gate is mechanical and mandatory — run it on EVERY review:**

```bash
# deck reviews:
python "$SKILL_DIR/scripts/validate_review.py" "<review.json>" --expect-pages=N --normalized-out="<review.json>"
# bundle reviews ('|'-separated BARE labels — no «guillemets»):
python "$SKILL_DIR/scripts/validate_review.py" "<review.json>" --expect-materials="a|b|c" --normalized-out="<review.json>"
```

Pass `--require-extended` on these fresh per-review calls: it enforces the
class-feedback fields — `indicio_ia` (1-5 advisory AI-slop signal on the
delivered material; NEVER a grade component) and `feedback_sugerido` (2-4
Spanish sentences of draft student feedback for the teaching team).

Exit 2 → re-dispatch that student ONCE with the `problems` list appended. Still
failing → `NO REVISADO`, reason "revisión incompleta", flag for humans. The validator
also normalizes fields (strict dates, bare student names, ≤70-char tool, controlled
flag vocabulary; everything else moves to `observations`) — content is preserved,
scores and justifications are never edited. Note the flag vocabulary split:
reviewers emit the subset in the templates; `ENTREGA SIN PPT`, `ENTREGA
DUPLICADA`, `SPOT-CHECK FALLIDO` and `EVIDENCIA DE ENVIO ANTERIOR INCLUIDA`
are assigned by the assembler, and `EMPATE TOP5` / `EDAD SIN CALCULAR` by `make_excel.py`.

**Spot-check honesty on BOTH passes** — every ~4th review, deck AND bundle alike
(the pilot spot-checked only deck reviews; the #1-ranked student came from the
unchecked pass): a small fresh-context checker opens ONE cited page/material and
compares it against the justification. **A one-page submission cannot cite "slide
N"** — check content fidelity there instead of citation presence. A failed check →
flags `SPOT-CHECK FALLIDO` + `REVISAR MANUALMENTE`, never a silent score edit.

### 5. Assemble results

```bash
python "$SKILL_DIR/scripts/assemble_results.py" "$WORK" "<RUN_DATE>" \
    --folder-desc="<source description>"
```

Expects `deck_reviews.json` / `bundle_reviews.json` in `$WORK` (shape:
`{"reviewed": [{"folder_id", "result", "problems": [], "spotcheck": null | {"plausible": bool, "notes": str}}]}`
— a spotcheck object without a `plausible` key is treated as INCONCLUSIVE, never
as failed).
Produces `results.json`: one row per submitted FILE — one graded row per student;
files READ inside that student's integral review get `revisado_anexo`
(Excel: **REVISADO (ANEXO)**, green tint), superseded duplicates and older
in-folder versions get `reemplazada` (Excel: **REEMPLAZADA**, violet tint),
and only genuinely unreviewed material stays NO REVISADO. It re-applies the validator
normalization (defense in depth) and runs **same-tool reconciliation**: students whose
verified dates for the same tool differ by >1 month all get `VERIFICAR FECHA` — the
reviewers verified independently, so this is where disagreement becomes visible.

### 6. Generate the Excel

```bash
cd "$WORK" && python "$SKILL_DIR/scripts/make_excel.py" results.json res.xlsx \
  --listing=listing.json --listing=sub-listings/<id>.json ...
```

Pass EVERY listing JSON (main + each subfolder), with relative paths — 70+ absolute
`--listing` args can blow the command-length limit. Write to a SHORT name
(`res.xlsx`), then copy to the final `Resultados-Vigilancia-Tecnologica-<RUN_DATE>.xlsx`
at delivery (a long name inside a deep `$WORK` hits MAX_PATH). The script computes the
final grade (single source of truth: 0.50/0.25/0.25, DQ → 1.0) and fails (exit 2)
naming any listed entry with no row. Fix the missing rows; never work around the gate.
Sheets: **Ranking**, **Detalle**, **Meta**.

### 7. Deliver

Copy the Excel to the delivery folder the user confirmed (a local Drive-synced path —
it syncs on its own). Replace the previous run's file if one is there; never leave two
versions side by side.

### 8. Report

Summarize for the teaching team **in Spanish**: revisados / descalificados / no
revisados (con motivos), el top-5 con herramientas y notas finales, cada flag de
revisión humana, y la ubicación del Excel. Remind them:
**la nota oficial requiere revisión humana** — esto es una preselección, no un
veredicto.

## Hard rules

- Every submitted file appears in the Excel — graded, DQ'd, or NO REVISADO with a real
  reason. Nothing silently skipped.
- Format is never a reason to skip or disqualify a student (fairness rule above).
- Launch dates are verified by web search; the deck's claim alone is never trusted.
  Low-confidence verification → `age_months` stays null, flag `VERIFICAR FECHA`, never DQ.
- Border band 3.5–4.5 months always carries flag `VERIFICAR FECHA`.
- On a resubmission, grade the LATEST; evidence attached only to an earlier attempt
  rides along flagged (`build_bundles.py` does this automatically). Within one
  folder, version-marked files (v1/v2/FINAL/(2)) resolve to the most recent —
  older versions become REEMPLAZADA rows. Cross-folder duplicates supersede
  only on a matching Canvas key; same-name-different-folder is FLAGGED for a
  human, never auto-discarded (two students can share a name).
- `indicio_ia` is advisory only: it never changes scores and never
  disqualifies — it signals unfiltered AI dumping for the teaching team.
- Never edit a reviewer's scores or justifications. Field normalization relocates
  out-of-format content; it never changes verdicts. If a verdict looks off, flag it —
  humans decide.
- Student files and results stay out of any git repo and out of public artifacts.
- Do not fabricate scores for unreadable material — mark it NO REVISADO.
- Human-facing artifacts leave in Spanish (language rule above).
