# Platform adapter — GPT (Codex CLI)

How this skill's capability requirements map onto Codex. Trust your actual
tool list over this table when they disagree — multi-agent tool names depend
on the multi-agent version your model preset selects.

| Capability | Codex mechanism |
|---|---|
| Read PDF pages visually | **Not native — ALWAYS rasterize first**: `python scripts/pdf_to_images.py <pdf> <outdir>` and give the reviewer the PNG list (`build_bundles.py --images` emits image instructions). Codex file reads are text; PDFs must become images to be SEEN. |
| View images | `view_image` on each PNG (pages, keyframes, screenshots) |
| Sub-reviewer with clean context | Requires `[features] multi_agent = true` in `~/.codex/config.toml`. `spawn_agent {fork_turns: "none"}` for a clean-context child; on Codex 0.145+ you may attach a role file via `agent_type`. |
| Mid-tier vision model | Set BOTH `model` AND `reasoning_effort` explicitly on every spawn (setting `model` alone silently resets the child's effort to that model's default). Pick a mid-tier vision-capable preset from your spawn allowlist — never copy a model name from this file. |
| Waiting on reviewers | `wait_agent` in bounded 5–10 min stretches (it is an event subscription — short polling buys nothing); `list_agents` to reconcile; `followup_task` to send the retry-once corrections to the SAME child. |
| No multi-agent available? | Review **sequentially in your own loop**, one student per iteration; between students, summarize-and-drop the previous student's pages from working context so grading stays independent. Slower but valid. |
| Web search | Codex `web_search` tool (must be enabled for the session). If unavailable, set `verification_confidence: "baja"` on every review, flag VERIFICAR FECHA, and tell the operator dates were NOT verified. |
| Structured JSON | No schema enforcement — instruct "SOLO JSON", parse, and run `scripts/validate_review.py`; exit 2 → one `followup_task` retry with the problems list. |
| Conversion backends | `shell` tool drives LibreOffice / Office COM / Chrome / ffmpeg; `prepare_materials.py` autodetects. Sandbox note: network-touching steps (web search aside) and some COM automation may need approval or a workspace-write profile — surface, don't silently skip. |
| Audio transcription | any local whisper CLI via `shell`; else frames alone, noting the missing transcript. |

Notes:
- The fairness gate is `validate_review.py`, not the harness: run it on every
  review regardless of how the reviewer was spawned.
- Keep the orchestration state (plan, listing, results) in files, not in
  conversation memory — Codex context compaction is aggressive.
