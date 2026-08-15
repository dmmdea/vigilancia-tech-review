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
- [ ] H. Review gate: semgrep (PYTHONUTF8=1, exit-code discipline) →
       code-reviewer + silent-failure-hunter (validator/assembler have many
       fallback paths)
- [ ] I. Evidence: run the full pipeline against the real Semana 2 folder in
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
