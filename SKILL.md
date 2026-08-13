---
name: vigilancia-tech-review
description: Use when the teaching team needs to review, score, and rank MBA "Vigilancia Tecnológica" student presentations from a shared Google Drive folder — reads every slide of every deck visually with one Sonnet sub-agent per deck, web-verifies each tool's launch date, disqualifies tools older than 4 months, scores PoC/Impacto/Comunicación (1.0–5.0, weighted 50/25/25), and produces a two-sheet Excel ranking with the top-5 candidates for human TA review. Triggers: "revisar presentaciones vigilancia tecnológica", "calificar las ppt de los estudiantes", "ranking vigilancia tecnológica", "escoger las mejores presentaciones".
---

# vigilancia-tech-review

Reviews every student deck in a link-shared Google Drive folder against the official
rubric of the Uniandes MBA activity **Vigilancia Tecnológica: IA de vanguardia**, then
delivers a ranked Excel for the human TAs. The skill never assigns the official grade —
it produces evidence-cited candidate scores and the top-5 shortlist for human review.

**Rubric (from the assignment):** Prueba de concepto 50% · Análisis de impacto 25% ·
Comunicación 25%. Exclusion filter: tool launched **more than 4 months** before the run
date → automatic 1.0 (verified by web search, not by trusting the deck). General-purpose
tools (ChatGPT, Gemini, Copilot, Claude…) are invalid unless the subject is a specific
recently launched feature. Scale: 1.0–5.0.

## Requirements

- Python 3.10+ with `openpyxl` and `pypdf` (`pip install openpyxl pypdf`).
- A PDF conversion backend: LibreOffice (`soffice`) or Microsoft PowerPoint (Windows).
- The Drive folder must be shared as **"anyone with the link"**. No Google credentials
  are needed; nothing is uploaded to Drive (the Excel is delivered locally).

## Inputs to collect from the user

1. **Drive folder URL** of the student submissions. If not given, ask for it.
2. Optional: a local Google Drive Desktop path for the same folder, if they want the
   Excel dropped there to auto-sync. Auto-detect first (step 7) before asking.

## Procedure

Let `SKILL_DIR` = this skill's directory, `WORK` = a fresh directory under the session
scratchpad (e.g. `<scratchpad>/vtr-<YYYYMMDD-HHMM>/`). Student files never go inside a
git repo. Let `RUN_DATE` = today (YYYY-MM-DD).

### 1. List the folder

```bash
python "$SKILL_DIR/scripts/drive_list.py" "<folder-url>" > "$WORK/listing.json"
```

- Exit 2 → tell the user the folder is not public (or the network failed) and stop.
- Exit 3 → Drive's page layout changed or an entry failed to parse; STOP the run and
  say so. NEVER proceed with a partial listing and never treat a parse failure as an
  empty folder.
- Keep entries with kind `file` whose name ends in `.pptx`/`.ppt`/`.pdf`
  (**case-insensitive** — `DECK.PPTX` counts), plus every `gslides` entry. Entries of
  kind `folder`: list each subfolder once (one level deep) and include its decks —
  students sometimes upload into per-group subfolders. A `folder` found INSIDE a
  subfolder is not explored: add a `NO REVISADO` row "subcarpeta anidada no explorada"
  naming it, so it is visible to the humans.
- Any other entry (docs, sheets, images, zips): do NOT review, but record it for the
  final report and add it to the Excel as `NO REVISADO` / "formato no soportado".
- Empty folder (exit 0, zero entries) → report "no submissions yet" and stop.

### 2. Download every deck

For each entry: build the local filename as `<safe-name>__<drive-id>.<ext>` (sanitized
name + the Drive file id). Drive allows duplicate filenames in one folder — the id
suffix is what prevents one student's deck from silently overwriting another's.

```bash
python "$SKILL_DIR/scripts/drive_download.py" "<id>" "$WORK/decks/<safe-name>__<id>.<ext>"          # kind=file
python "$SKILL_DIR/scripts/drive_download.py" "<id>" "$WORK/decks/<safe-name>__<id>.pptx" --kind=gslides
```

Exit 2 or 3 → `NO REVISADO` row for THAT file with the stderr error as
`status_reason` — never silently dropped. **Exit 4 (Drive quota/rate-limit page)
→ STOP the whole run**: Drive is throttling anonymous downloads; tell the user to
wait (~30–60 min) and rerun, rather than mass-marking the remaining students as
NO REVISADO.

### 3. Convert to PDF

```bash
python "$SKILL_DIR/scripts/convert_to_pdf.py" "$WORK/decks/<name>" "$WORK/pdf/<name>.pdf"
```

Capture `pages` from the JSON output. Exit 2 = environment problem (no backend
installed, or pypdf missing) → tell the user what to install and stop the whole run.
Exit 3 = THIS file failed (corrupt, password-protected, 0 pages) → `NO REVISADO` row
with the stderr reason, continue with the other files.

### 4. Dispatch one reviewer sub-agent per deck — Sonnet, parallel batches

Read `SKILL_DIR/templates/reviewer-prompt.md`, fill the `{{placeholders}}`
(`pdf_path`, `pages_total`, `run_date`, `original_filename`) and dispatch with the
Agent tool: `subagent_type: general-purpose`, **`model: sonnet`**, in batches of ~4
concurrent agents. Each returns a single JSON object.

**Fairness gate — validate every returned JSON:**
- `pages_read == pages_total`, scores within 1.0–5.0, non-empty justifications.
- If `verification_confidence` is alta/media, `age_months` must be a number —
  a null age with confident verification silently disables the DQ filter.
- Spot-check honesty: for at least one deck per batch, verify one slide citation from
  the justification against the actual PDF page (a lazy reviewer can echo
  `pages_total` without reading; a fabricated citation exposes it).
- On any violation: re-dispatch that deck ONCE with a note about what was invalid.
  Still invalid → `NO REVISADO` row, reason "revisión incompleta", flag for humans.
- Never edit a reviewer's scores. If a verdict looks off, add a flag — humans decide.

### 5. Assemble results

Build `$WORK/results.json`:

```json
{"run_date": "<RUN_DATE>", "folder_url": "<url>", "results": [ <one reviewer JSON per deck,
  plus for each failed/skipped file: {"id": "<drive-file-id>", "file": "...",
  "status": "no_revisado", "status_reason": "...", "disqualified": false, "flags": []}> ]}
```

Reviewed decks get `"status": "revisado"`. For EVERY row (reviewed or not) you add two
fields from your own listing metadata: `"id"` (the Drive file id — the identity key
for the top-5 and the completeness check) and `"file"` overwritten with the ORIGINAL
Drive filename (reviewers sometimes echo the PDF name instead). Do not compute final
grades yourself — `make_excel.py` does it (single source of truth: 0.50/0.25/0.25,
DQ → 1.0; it also downgrades any reviewed row with invalid scores to NO REVISADO).

### 6. Generate the Excel

```bash
python "$SKILL_DIR/scripts/make_excel.py" "$WORK/results.json" \
  "$WORK/Resultados-Vigilancia-Tecnologica-<RUN_DATE>.xlsx" \
  --listing="$WORK/listing.json" --listing="$WORK/sub-<name>.json"  # one per listed folder
```

Pass EVERY listing JSON you produced (main folder + each subfolder). The script fails
(exit 2) naming any listed entry that has no row in results.json — decks, unsupported
formats, and unexplored nested folders alike all need a row (reviewed or NO REVISADO).
This is the mechanical backstop for "nothing is silently dropped". Fix the missing
rows; never work around it.

Sheets: **Ranking** (sorted, top-5 starred, DQ red, no-revisado gray), **Detalle**
(slide-cited justifications), **Meta** (run parameters).

### 7. Deliver

Try to find a local Google Drive Desktop mount of the SAME folder id (generic patterns —
never assume one specific user's layout):

- `G:\.shortcut-targets-by-id\<folder-id>\...` and `G:\My Drive\...` (any drive letter)
- `~/Google Drive/...`, `~/GoogleDrive/...`, `/Volumes/GoogleDrive/...` (macOS)
- A path the user provided.

If found: copy the Excel there (it syncs to Drive automatically) and say so.
If not: give the local path and tell the user to upload it to the Drive folder manually
(anonymous upload to Drive is impossible by design of this skill — no credentials).

### 8. Report

Summarize in chat: decks reviewed / disqualified / not reviewed (with reasons),
the top-5 with tool names and finals, every human-review flag, and the Excel location.
Remind: **the official grade requires human TA review** — this is a shortlist, not a verdict.

## Hard rules

- Every file in the folder appears in the Excel — reviewed, DQ'd, or NO REVISADO with
  a reason. Nothing is silently skipped.
- Launch dates are verified by web search; the deck's claim alone is never trusted.
  Low-confidence verification → `age_months` stays null (estimate goes to
  evidence_notes, labeled as declared-date-based), flag `VERIFICAR FECHA`, never DQ.
- Border band 3.5–4.5 months always carries flag `VERIFICAR FECHA` (DQ applies above
  4.0 per the assignment rule, but a human confirms borderline DQs).
- Student files and results stay out of any git repo and out of public artifacts.
- Do not fabricate scores for unreadable decks — mark them NO REVISADO.
