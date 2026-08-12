# vigilancia-tech-review — Design Spec

Date: 2026-08-12 · Status: approved by Daniel (brainstorm session)

## Purpose

A Claude Code skill for the teaching team of the Uniandes MBA course *Reto Integrador 1 –
Tecnología de Información*. The course's transversal activity **Vigilancia Tecnológica**
(20% of the final grade) has students pick an AI tool launched in the last 3–4 months,
run a real proof-of-value with it, and present it in 3 minutes. Students upload their
decks to a shared Google Drive folder. The skill reviews **every deck, every slide**,
scores each per the official rubric, disqualifies stale tools, ranks the cohort, and
delivers a two-sheet Excel so the human TAs can pick/verify the top 5.

## Official rubric (from the assignment sheet)

| Criterion | Weight |
|---|---|
| Prueba de concepto (PoV, real evidence) | 50% |
| Análisis de impacto (personal + business productivity) | 25% |
| Comunicación (ficha técnica, clarity, 3-min fitness) | 25% |

**Exclusion filter:** tool launched more than 3–4 months ago → automatic grade 1.0,
other criteria do not apply. General-purpose tools (ChatGPT, Gemini, Copilot, Claude…)
are invalid unless the subject is a *specific recently launched feature*.

## Decisions (locked in brainstorm)

1. **Drive access = public link only.** No OAuth, no rclone, no API keys. Reading uses
   the anonymous `embeddedfolderview` endpoint + `drive.usercontent.google.com/download`.
   Verified working 2026-08-12 against the real class folder. Anonymous upload is
   impossible → the results Excel is written locally; if a local Google Drive Desktop
   mount of the folder is detected (generic patterns only, never a hardcoded personal
   path), the Excel is copied there so it syncs; otherwise a TA uploads it manually.
2. **Recency check = web-verified.** Each sub-agent web-searches the tool's real launch
   date (official announcement, changelog, credible press) and cites the source. The
   verified date governs. Cutoff: **>4 months** older than the run date → DQ (1.0).
   Inconclusive verification or gray zone (3–4 months) → **no auto-DQ**; row flagged
   `VERIFICAR FECHA` for human decision.
3. **Scale = Colombian 1.0–5.0** per criterion; final = weighted 50/25/25; DQ = 1.0 flat.
4. **Reading = full visual.** Every deck is converted to PDF and the sub-agent reads the
   PDF *visually* (Claude's Read tool renders PDF pages), so screenshots, charts and
   design are graded — not just extracted text. Completeness is enforced: the reviewer
   must report `pages_read`; the orchestrator compares it to the actual PDF page count
   and re-dispatches on mismatch.
5. **Sub-agents = Sonnet** (user-specified), one per deck, dispatched in parallel batches.

## Architecture

Deterministic plumbing in Python scripts; judgment in sub-agents.

```
SKILL.md                    orchestration instructions (Claude follows)
scripts/
  drive_list.py             anonymous folder listing → JSON (id, name, type)
  drive_download.py         anonymous download by id; Google Slides → export .pptx
  convert_to_pdf.py         pptx/ppt → pdf (soffice, else PowerPoint COM); reports page count
  make_excel.py             results JSON → two-sheet .xlsx (openpyxl)
templates/
  reviewer-prompt.md        the per-deck sub-agent prompt (Spanish rubric, JSON contract)
```

### Flow

1. User invokes skill with the Drive folder URL (or accepts the default from config).
2. `drive_list.py` lists the folder → files (pptx/pdf/Google Slides; ignores others, reports them).
3. `drive_download.py` fetches each file into a work dir.
4. `convert_to_pdf.py` converts each to PDF, emits `{file, pdf, pages}`.
5. Orchestrator dispatches one Sonnet sub-agent per deck (batches of ~4) with the
   reviewer prompt + PDF path + run date. Contract: JSON with ficha técnica fields,
   `pages_read`, per-criterion scores + slide-cited justifications, launch-date
   verification (date, source URL, confidence), DQ verdict + reason, human-review flags.
6. Orchestrator validates each JSON (pages_read == pages; scores in range) and re-runs
   incomplete reviews once before flagging them.
7. Weighted final score, DQ enforcement, ranking. `make_excel.py` builds
   `Resultados-Vigilancia-Tecnologica-<YYYY-MM-DD>.xlsx`:
   - **Ranking** sheet: archivo, herramienta, fecha declarada, fecha verificada + fuente,
     edad (meses), descalificado + razón, PoC, Impacto, Comunicación, nota final,
     posición, Top-5 ⭐, flags.
   - **Detalle** sheet: full per-criterion justification with slide references.
8. Output placement: try generic Drive Desktop mounts for the target folder id
   (`G:\My Drive\…`, `G:\.shortcut-targets-by-id\<id>\…`, `~\Google Drive\…`,
   configurable override); if found, copy the Excel there; else print the local path
   and instruct manual upload.

### Error handling

- Folder not public / empty → clear message, stop.
- Download/convert failure per file → recorded in the Excel as `NO REVISADO` + reason;
  never silently dropped (fairness: every submission accounted for).
- Web verification inconclusive → flag, not DQ.
- A failed probe (network error) is never treated as "tool doesn't exist".

### Non-goals

- No upload to Drive (anonymous impossible — accepted).
- No grading of the live 3-minute delivery; Comunicación is judged from the deck.
- No LMS (Bloque Neón) integration.

## Privacy

Student decks and results never enter the public repo. The repo ships only the skill.
Work dirs default to the OS temp area.
