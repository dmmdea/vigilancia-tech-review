---
name: vigilancia-tech-review
description: Use when the teaching team needs to review, score, and rank MBA "Vigilancia Tecnológica" student presentations — from a link-shared Google Drive folder OR a local/Drive-Desktop folder of Canvas submissions. Reads EVERY submitted file whatever its format (pptx, pdf, docx, html, png, mp4, xlsx, zip) with one Sonnet reviewer per student, web-verifies each tool's launch date, disqualifies tools older than 4 months, scores PoC/Impacto/Comunicación (1.0–5.0, weighted 50/25/25), and produces a two-sheet Excel ranking with the top-5 candidates for human TA review. Triggers: "revisar presentaciones vigilancia tecnológica", "calificar las ppt de los estudiantes", "ranking vigilancia tecnológica", "escoger las mejores presentaciones".
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

## Requirements

- Python 3.10+ with `openpyxl` and `pypdf` (`pip install openpyxl pypdf`).
- Slides→PDF: LibreOffice (`soffice`) **or** Microsoft PowerPoint (Windows COM).
- Optional but strongly recommended for full format coverage:
  - Microsoft Word (COM) or LibreOffice → `.docx` (best fidelity: keeps embedded
    screenshots, which are usually the student's PoC evidence)
  - Chrome/Edge (headless) → `.html`
  - `ffmpeg` → video keyframes
  - a local/whisper STT for video audio (optional; frames alone still work)
- For a Drive source: folder shared as **"anyone with the link"**. No credentials.

## Inputs to collect from the user

1. **Where the submissions are** — a Drive folder URL, or a local path (Google Drive
   Desktop mount / unzipped Canvas bulk download). Ask if not given.
2. Optional: a local Drive-synced path to drop the Excel into so it auto-syncs.

## Procedure

`SKILL_DIR` = this skill's directory. `WORK` = a fresh working directory.
`RUN_DATE` = today (YYYY-MM-DD). Student files stay out of any git repo.

> **Windows MAX_PATH — read before anything else.** Student folder names plus filenames
> routinely exceed 260 chars, and Python here does NOT auto-apply the `\\?\` prefix:
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
Exit 2 = folder unreadable · 3 = nothing found. Review any `NOTE ...` lines it prints
on stderr — those are ambiguous primary-file picks a human should confirm.

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

### 2. Make every file readable — whatever its format

```bash
python "$SKILL_DIR/scripts/prepare_materials.py" "$WORK" [--matroot=<short dir>]
```

Renders each submitted file into something a reviewer can actually consume, converting
ONLY where the Read tool cannot open the format (Read already renders PDFs and images
visually — never "convert" a PNG):

| Submitted | Becomes | Why |
|---|---|---|
| `.pdf` | as-is | vision |
| `.pptx .ppt .odp` | PDF | Read can't open slides |
| `.docx .doc .rtf` | PDF (Word COM → LibreOffice) | keeps embedded screenshots = the PoC evidence |
| `.html .htm` | PDF (headless Chrome) | keeps the rendering |
| `.png .jpg …` | passthrough | Read shows images directly |
| `.xlsx .xls .csv` | `.txt` cell dump | data reads better as data |
| `.mp4 .mov …` | keyframes (+ transcript) | reviewer's own vision judges the frames |
| `.zip` | extracted, contents re-routed | — |
| `.py .txt .md` | passthrough | text |

Writes `materials.json`. It prints any video lacking a transcript; transcribe those
(a local whisper is ideal — mechanical work) and re-run with
`--transcript=<folder_id>=<file.txt>`. Re-runs with `--only=` merge, never clobber.
Fix every `ERROR` item it reports before moving on — each one is a student at risk of
losing credit for evidence they did submit.

### 3. Dispatch one reviewer per student — Sonnet, parallel

Use a Workflow (or the Agent tool in batches of ~4). Two templates:

- `templates/reviewer-prompt.md` — a plain single-deck submission.
- `templates/bundle-reviewer-prompt.md` — a student with **no slide deck** or with
  **deck + supplementary evidence**. It lists every material with explicit per-item
  reading instructions and requires the reviewer to report `materials_reviewed`.

A student who has BOTH a deck and extra evidence must be reviewed with the bundle
template — the evidence is usually the proof behind the PoC score, and a deck-only
review silently under-grades them.

**Fairness gate — validate every returned JSON:**
- `pages_read == pages_total` (deck template) or `materials_reviewed` covers every
  listed material (bundle template).
- Scores within 1.0–5.0; justifications non-empty.
- If `verification_confidence` is alta/media, `age_months` must be numeric — a null age
  with confident verification silently disables the DQ filter.
- Spot-check honesty on a sample: verify one cited slide against the actual page.
  **A one-page submission cannot carry a "slide N" citation** — do not treat a missing
  slide number there as fabrication; check whether the content matches instead.
- On violation: re-dispatch that student ONCE with the specific problem. Still invalid →
  `NO REVISADO`, reason "revisión incompleta", flag for humans.
- Never edit a reviewer's scores. If a verdict looks off, add a flag — humans decide.

### 4. Assemble results

Build `$WORK/results.json`:

```json
{"run_date": "<RUN_DATE>", "folder_url": "<source>", "results": [ ... ]}
```

**One row per submitted FILE.** Within a student, exactly ONE row carries the grade
(`"status": "revisado"`); the student's other files get `"status": "no_revisado"` with a
reason that says they WERE read inside that student's review and points to the graded
row. Every row needs `"id"` (the listing's file id) and `"file"` (the ORIGINAL filename —
reviewers often echo a converted path). Do not compute final grades yourself.

### 5. Generate the Excel

```bash
python "$SKILL_DIR/scripts/make_excel.py" "$WORK/results.json" \
  "$WORK/Resultados-Vigilancia-Tecnologica-<RUN_DATE>.xlsx" \
  --listing="$WORK/listing.json" --listing="$WORK/sub-listings/<id>.json" ...
```

Pass EVERY listing JSON (main + each subfolder). Run it from inside `$WORK` with
relative paths — 70+ absolute `--listing` args can blow the command-length limit —
and write to a SHORT output name (e.g. `res.xlsx`), then copy it to the final
`Resultados-…-<RUN_DATE>.xlsx` name at delivery: openpyxl saves through plain Windows
I/O, so a long name inside a deep `$WORK` hits MAX_PATH and fails.
The script computes the final grade (single source of truth: 0.50/0.25/0.25, DQ → 1.0)
and fails (exit 2) naming any listed entry with no row. Fix the missing rows; never
work around the gate. Sheets: **Ranking**, **Detalle**, **Meta**.

### 6. Deliver

Find a local Drive Desktop mount of the same folder (`G:\.shortcut-targets-by-id\<id>\…`,
`G:\My Drive\…`, `~/Google Drive/…`, `/Volumes/GoogleDrive/…`, or a user-given path),
copy the Excel there, and say so. Otherwise give the local path for manual upload.

### 7. Report

Summarize: reviewed / disqualified / not reviewed (with reasons), the top-5 with tool
names and finals, every human-review flag, and the Excel location. Remind:
**the official grade requires human TA review** — this is a shortlist, not a verdict.

## Hard rules

- Every submitted file appears in the Excel — graded, DQ'd, or NO REVISADO with a real
  reason. Nothing silently skipped.
- Format is never a reason to skip or disqualify a student (see the fairness rule above).
- Launch dates are verified by web search; the deck's claim alone is never trusted.
  Low-confidence verification → `age_months` stays null, flag `VERIFICAR FECHA`, never DQ.
- Border band 3.5–4.5 months always carries flag `VERIFICAR FECHA`.
- On a resubmission, grade the LATEST but do not lose evidence the student attached to
  an earlier attempt — attach it and flag it for the human.
- Student files and results stay out of any git repo and out of public artifacts.
- Do not fabricate scores for unreadable material — mark it NO REVISADO.
