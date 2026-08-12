# vigilancia-tech-review

Claude Code skill that reviews, scores, and ranks MBA **"Vigilancia Tecnológica: IA de
vanguardia"** student presentations straight from a link-shared Google Drive folder —
no Google credentials required.

For each deck it:

1. Downloads it anonymously from the shared folder (PPTX, PDF, or Google Slides).
2. Converts it to PDF and dispatches a **Sonnet sub-agent** that reads **every slide
   visually** (screenshots and charts included — the PoV evidence is usually images).
3. Extracts the ficha técnica and **web-verifies the tool's real launch date** against
   official announcements — decks' claims are never trusted blindly.
4. Applies the exclusion filter: tools launched **more than 4 months** before the run
   date → automatic 1.0 (gray zone / low-confidence dates are flagged for humans, never
   auto-disqualified).
5. Scores **Prueba de concepto (50%) · Análisis de impacto (25%) · Comunicación (25%)**
   on the 1.0–5.0 scale, with slide-cited justifications.
6. Builds a ranked, color-coded Excel (Ranking + Detalle + Meta sheets) with the
   **top-5 candidates starred** for human TA review.

The skill produces a *shortlist with evidence*, not official grades — the human
teaching team always makes the final call.

## Install

```bash
git clone https://github.com/dmmdea/vigilancia-tech-review.git
# Claude Code (personal skills):
cp -r vigilancia-tech-review ~/.claude/skills/vigilancia-tech-review
pip install openpyxl pypdf
```

You also need one PDF conversion backend: [LibreOffice](https://www.libreoffice.org/)
(any OS) or Microsoft PowerPoint (Windows).

## Use

In Claude Code:

> Revisa las presentaciones de vigilancia tecnológica en
> https://drive.google.com/drive/folders/XXXXXXXXXXXX y dame el ranking

Requirements on the folder: shared as **"anyone with the link"**. The results Excel is
written locally (and copied into a local Google Drive Desktop mount of the folder when
one is detected, so it syncs up automatically); otherwise upload it to the folder
manually — anonymous uploads to Drive are impossible, which is the price of the
zero-credentials design.

## Repo layout

```
SKILL.md                    orchestration instructions Claude follows
scripts/drive_list.py       anonymous listing of a public Drive folder
scripts/drive_download.py   anonymous download (binary or Google Slides export)
scripts/convert_to_pdf.py   PPTX→PDF via LibreOffice or PowerPoint COM + page count
scripts/make_excel.py       results JSON → two-sheet ranked Excel
templates/reviewer-prompt.md  the per-deck Sonnet reviewer prompt (Spanish)
docs/specs/                 design spec
```

## Privacy

Student decks and results never enter this repo and are never uploaded anywhere by the
skill; all work happens in a local scratch directory.
