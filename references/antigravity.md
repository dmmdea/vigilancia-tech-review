# Platform adapter — Gemini (Antigravity)

How this skill's capability requirements map onto Antigravity (`agy`). Trust
your actual tool list over this table when they disagree.

| Capability | Antigravity mechanism |
|---|---|
| Read PDF pages visually | Gemini is natively multimodal, but Antigravity's file tools read text — **rasterize first**: `python scripts/pdf_to_images.py <pdf> <outdir>`, then view the PNGs (`build_bundles.py --images` emits image instructions). If your session's file/attachment tool demonstrably renders PDFs page-by-page, use it and skip rasterizing. |
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
