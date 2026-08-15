# Platform adapter — GPT (Codex CLI)

**LIVE-VERIFIED 2026-08-15 on codex-cli 0.147.0** against real student
submissions (single-image bundle + 18-page rasterized PDF): image reading,
web search, schema-enforced JSON, and the validator gate all pass end-to-end.
Trust your actual tool list over this table when they disagree.

## Primary dispatch mode: one `codex exec` per student (verified)

A fresh `codex exec` process IS the clean-context reviewer — simpler and more
isolated than multi-agent spawning, and it works headless:

```bash
codex exec -s read-only --skip-git-repo-check --ephemeral \
  -C "<matroot>" \
  -c tools.web_search=true \
  --output-schema "<skill>/templates/review_schema.json" \
  -o "<work>/reviews/<folder_id>.json" \
  - < "<work>/prompts/<folder_id>.md"
```

Verified behavior (both tests, real data):
- **Reads PNG page images and student screenshots from disk and genuinely
  SEES them** (justifications transcribed on-image text and cited specific
  pages, "páginas 3–4"). 18 sequential page PNGs handled in one run. `-i
  <file>...` can also attach images up front if disk reads are restricted.
- **Web search works headless** with `-c tools.web_search=true` — the test
  found the official Google blog announcement and dated a feature to the day,
  matching an independent Claude reviewer's verification exactly.
- **`--output-schema` enforces the response shape** (OpenAI structured
  outputs): the final message is schema-valid JSON, written cleanly to the
  `-o` file — no fence-stripping, no parsing. Use the bundled
  `templates/review_schema.json`.
- Spanish in/out end-to-end; `verification_confidence`, DQ decision, and
  flags came back rule-coherent (a 13-month-old tool was correctly DQ'd with
  DISCREPANCIA FECHA).
- `-s read-only` suffices — the reviewer only reads materials and searches.

The retry-once path: re-run the same `codex exec` with the validator's
`problems` list appended to the prompt (a fresh process is fine — the review
must be independent anyway). Then ALWAYS run `scripts/validate_review.py`
(schema checks shape, not the fairness rules).

| Capability | Codex mechanism |
|---|---|
| Read PDF pages visually | **Rasterize first**: `prepare_materials.py <work> --rasterize`, then `build_bundles.py <work> --images --all`. Codex then reads the PNGs from disk (verified) — PDFs themselves are not rendered visually. |
| Sub-reviewer with clean context | one headless `codex exec` per student (above). Alternative: `[features] multi_agent = true` + `spawn_agent {fork_turns: "none"}` inside an interactive session; on 0.145+ role files attach via `agent_type`. |
| Model choice | `-m <model>` on exec; pick a mid-tier vision-capable preset from your allowlist. In multi-agent mode set BOTH `model` and `reasoning_effort` on every spawn. |
| Parallelism | run several `codex exec` processes concurrently (~4); each is fully isolated. In multi-agent mode use `wait_agent` in 5–10 min stretches + `followup_task` for retries. |
| Web search | `-c tools.web_search=true` (verified). If your environment blocks it, reviews must return `verification_confidence: "baja"` (age null, flag VERIFICAR FECHA) and the operator is told dates were NOT verified. |
| Structured JSON | `--output-schema` + `-o` (verified) — then `validate_review.py` for the fairness rules. |
| Conversion backends | run `prepare_materials.py` OUTSIDE the sandbox (orchestrator side); Office COM / Chrome / ffmpeg need real system access that `read-only` reviewers don't get and don't need. |
| Audio transcription | any local whisper CLI, orchestrator side; else frames alone, noting the missing transcript. |

Notes:
- Cross-model reality check from the live test: Codex and Claude verified
  DIFFERENT launch dates for the same product family once (a 2025 rollout vs
  a 2026 GA of the rebranded feature) — both defensible reads. The
  same-tool reconciliation pass + VERIFICAR FECHA exist precisely for this;
  the human decides.
- Keep orchestration state (plan, listing, reviews) in files, not
  conversation memory — context compaction is aggressive.
