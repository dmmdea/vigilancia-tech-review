# Platform adapter — Gemini (Antigravity)

How this skill's capability requirements map onto Antigravity. Trust your
actual tool list over this table when they disagree.

**Live-verified 2026-08-15** on Antigravity 2.8.1 (Windows hub desktop app,
model "Gemini 3.7 Flash High"): single-PNG review E2E — the agent opened the
local image by absolute path, ran real web search, and its date verification
matched Claude and Codex exactly (same tool, same launch date, same
confidence). Wrote the JSON to the requested absolute output path in ~2 min.
`validate_review.py` correctly rejected a first pass missing the extended
fields; one retry message citing the exact problems produced a fully valid
review (exit 0) — the re-dispatch-once rule works unchanged on this harness.

Operational notes from that run:
- **2.8.1 ships no CLI** — reviews are dispatched through the hub's agent
  conversation (paste the filled template) or via UI automation. Give the
  agent ABSOLUTE paths for both materials and the output JSON; "this folder"
  is ambiguous inside the hub.
- **Launcher gotcha:** if Antigravity is started from a process that has
  `ELECTRON_RUN_AS_NODE=1` in its environment (any VS Code extension host,
  some CI shells), the exe exits code 0 instantly with no window and no log
  line. Clear that variable before launching.

| Capability | Antigravity mechanism |
|---|---|
| Read PDF pages visually | Gemini is natively multimodal, but Antigravity's file tools read text — **rasterize first**: `python scripts/prepare_materials.py <work> --rasterize` (drives pdf_to_images.py per PDF), then `build_bundles.py <work> --images --all` so every student gets PNG-page instructions. If your session's file/attachment tool demonstrably renders PDFs page-by-page, use it and skip rasterizing. |
| View images | your image-viewing/file tool on each PNG (pages, keyframes) |
| Sub-reviewer with clean context | `invoke_subagent` with `TypeName: "self"` (full capability; the reviewer needs image viewing + web search). `research` type is read-only — fine for spot-checks. |
| Mid-tier vision model | Antigravity routes models per subagent type; if you cannot pin a model, note the actual model used in the run report so graders know. |
| Task tracking across 70+ students | Antigravity has NO todo tool — maintain a **task artifact**: `write_to_file` with `IsArtifact: true`, `ArtifactMetadata.ArtifactType: "task"`, one checklist line per student; update with `replace_file_content` as reviews land. Re-read it before each batch — it is the source of truth after context gets long. (`manage_task` manages background processes, not checklists.) |
| Web search | the browser/search subsystem. If unavailable in your session, set `verification_confidence: "baja"` everywhere, flag VERIFICAR FECHA, and tell the operator dates were NOT verified. |
| Structured JSON | No schema enforcement — instruct "SOLO JSON", parse, and run `scripts/validate_review.py`; exit 2 → one re-invoke with the problems list. |
| Conversion backends | `run_command` drives LibreOffice / Office COM / Chrome / ffmpeg; `prepare_materials.py` autodetects. |
| Audio transcription | any local whisper CLI via `run_command`; else frames alone, noting the missing transcript. |

Notes:
- The fairness gate is `validate_review.py`, not the harness: run it on every
  review regardless of how the reviewer was invoked.
- Batch ~4 concurrent subagents; more mostly contends on the conversion and
  web-search backends.
