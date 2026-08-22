# Platform adapter — Claude (Claude Code)

How this skill's capability requirements map onto Claude Code. Trust your
actual tool list over this table when they disagree — harness tools evolve.

| Capability (SKILL.md speaks in these) | Claude Code mechanism |
|---|---|
| Read PDF pages visually | `Read` tool with `pages` parameter (max ~20 pages/call) — **native; `pdf_to_images.py` is NOT needed** |
| View images (PNG/JPG, video keyframes) | `Read` on the image path |
| Sub-reviewer with clean context | `Agent` tool (`subagent_type: general-purpose`), batches of ~4 — or the `Workflow` tool with a `pipeline()` over all students when multi-agent orchestration is authorized |
| Mid-tier vision model for reviewers | `model: sonnet` on the Agent/Workflow call |
| Web search (launch-date verification) | `WebSearch` tool (subagents inherit it) |
| Structured JSON back from reviewers | `Workflow` `schema:` option enforces it at the tool layer; with the plain `Agent` tool, instruct "SOLO JSON" and parse — then ALWAYS run `scripts/validate_review.py` either way (schema enforcement checks shape, not the fairness rules) |
| Retry-once on gate failure | re-dispatch with the validator's `problems` list appended to the prompt |
| Office/HTML/video conversion backends | `Bash`/`PowerShell` tools drive PowerPoint/Word COM (Windows), LibreOffice, Chrome headless, ffmpeg — `prepare_materials.py` autodetects |
| Audio transcription (video AND standalone audio) | `offload_transcribe` (local whisper, free) on the item's `video_path`/`audio_path`, then `--transcript=<folder_id>#<n>=<text_path>`. For video, frames alone still work; for **audio there is no fallback** — untranscribed speech is unreviewable. An empty/nonsense transcript usually means the audio is non-speech (the tool's generated output), which the bundle already handles. |

Notes:
- Windows: mind MAX_PATH (see SKILL.md) — the bundled scripts handle it; keep
  workdirs shallow anyway.
- Spot-checks: dispatch the checker as another small `Agent` call that Reads
  ONE cited page and compares it against the justification.
