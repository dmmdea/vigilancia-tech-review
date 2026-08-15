# Platform adapter — GPT (Codex CLI)

**LIVE-VERIFIED 2026-08-15 on codex-cli 0.147.0** against real student
submissions (single-image bundle + 18-page rasterized PDF): image reading,
web search, schema-enforced JSON, and the validator gate all pass end-to-end.
Trust your actual tool list over this table when they disagree.

## Primary dispatch mode: one `codex exec` per student (verified)

A fresh `codex exec` process IS the clean-context reviewer — simpler and more
isolated than multi-agent spawning, and it works headless:

```bash
codex exec -s read-only --skip-git-repo-check \
  -C "<matroot>" \
  -m "<a vision-capable model from YOUR allowlist>" \
  -c tools.web_search=true \
  --output-schema "<skill>/templates/review_schema.json" \
  -o "<work>/reviews/<folder_id>.json" \
  - < "<work>/prompts/<folder_id>.md"

# then, ALWAYS, the fairness gate (deck reviews use --expect-pages=N):
python "<skill>/scripts/validate_review.py" "<work>/reviews/<folder_id>.json" \
  --expect-materials="a|b|c" --normalized-out="<work>/reviews/<folder_id>.json"
```

Two honesty notes about "verified": (1) the test runs INHERITED
`~/.codex/config.toml` (they resolved to that machine's default model, and its
`web_search` setting was already on) — pin `-m` explicitly and CONFIRM
searches actually fired (the exec log prints `web search: <query>` lines;
`verification_source` must be a real URL) rather than trusting the flag
alone. (2) do NOT pass `--ephemeral` for grading runs: it deletes the session
record under `~/.codex/sessions/`, which is the only audit trail if a student
disputes a score. Keep sessions; clean them per your data-retention policy.

Verified behavior (both tests, real data; image reading additionally proven
by an unguessable-text PNG probe — codex transcribed content it could not
have guessed):
- **Reads PNG page images and student screenshots from disk and genuinely
  SEES them** (justifications transcribed on-image text and cited specific
  pages). An 18-page run reported pages_read: 18 with citations spanning the
  document — solid but self-reported; codex does not log image-read events,
  so spot-checks (SKILL.md step 4) remain the independent confirmation. `-i
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
`problems` list appended to the prompt — a FRESH process, in exec mode and in
multi-agent mode alike (independence beats continuity here; prefer this over
`followup_task` to the same child). Schema enforcement checks shape, not the
fairness rules — the validator call above is unconditional, per review, no
exceptions.

| Capability | Codex mechanism |
|---|---|
| Read PDF pages visually | **Rasterize first**: `prepare_materials.py <work> --rasterize`, then `build_bundles.py <work> --images --all`. Codex then reads the PNGs from disk (verified) — PDFs themselves are not rendered visually. |
| Sub-reviewer with clean context | one headless `codex exec` per student (above). Alternative: `[features] multi_agent = true` + `spawn_agent {fork_turns: "none"}` inside an interactive session; on 0.145+ role files attach via `agent_type`. |
| Model choice | `-m <model>` on exec; pick a mid-tier vision-capable preset from your allowlist. In multi-agent mode set BOTH `model` and `reasoning_effort` on every spawn. |
| Parallelism | run several `codex exec` processes concurrently (~4); each is fully isolated. In multi-agent mode use `wait_agent` in 5–10 min stretches; retries still go to a FRESH reviewer. |
| Web search | `-c tools.web_search=true` (verified). If your environment blocks it, reviews must return `verification_confidence: "baja"` (age null, flag VERIFICAR FECHA) and the operator is told dates were NOT verified. |
| Structured JSON | `--output-schema` + `-o` (verified) — then `validate_review.py` for the fairness rules. |
| Conversion backends | run `prepare_materials.py` OUTSIDE the sandbox (orchestrator side); Office COM / Chrome / ffmpeg need real system access that `read-only` reviewers don't get and don't need. |
| Audio transcription | any local whisper CLI, orchestrator side; else frames alone, noting the missing transcript. |

Notes:
- **The fairness gate is `validate_review.py`, not the harness** — run it on
  every review regardless of how the reviewer was spawned (same rule as every
  other adapter).
- `--output-schema` constrains INTERMEDIATE assistant messages too, not just
  the final one — schema-shaped placeholders with zeroed scores appear
  mid-stream. Harmless with `-o` (last message only), but never parse the
  `--json` event stream for the review.
- Cross-model reality check from the live test: Codex and Claude verified
  DIFFERENT launch dates for the same product family once (a 2025 rollout vs
  a 2026 GA of the rebranded feature) — both defensible reads. The
  same-tool reconciliation pass + VERIFICAR FECHA exist precisely for this;
  the human decides.
- Keep orchestration state (plan, listing, reviews) in files, not
  conversation memory — context compaction is aggressive.
