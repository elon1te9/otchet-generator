# Workspace rules

- Start report work with `skills/generate-practice-report/SKILL.md` and `docs/setup.md`. Resolve missing or ambiguous inputs before drafting; save confirmed data in `config/student.yaml`. Use personal data and supervisors only from that config.
- Source priority: `references/official/` for requirements; `references/current-example/` for the current title and structure; `references/previous-reports/` for style only, never old tasks or claims.
- All rules in `docs/report-spec.md` and `docs/writing-style.md` remain mandatory for every report; examples cannot weaken them. Honor explicit user overrides.
- Inspect the real project and attempt build/run before writing. Never invent features, data, execution, tests or screenshots; distinguish source evidence from observed results. Successful checks and figures require actual execution evidence. If screenshots are unavailable, continue the DOCX using the blank-figure fallback in report-spec.
- Derive `temp/report-plan.json` from the selected current example and assignment; pass it to validators. Never reuse a previous report's section names.
- Reuse materials only when current `temp/setup.json` fingerprints and artifact provenance match. Do not use `temp/archive-original` as a new report input unless explicitly requested.
- Validate the final DOCX structurally, render every page and inspect each visually.
- For revisions, preserve the latest user-edited DOCX and protected pages; follow spec and writing-style.
- Keep temporary artifacts in `temp/`, final deliverables in `output/`. Use local skills and reusable tools; keep detailed procedures out of this file.
- Read selectively with rg, excluding .git, .vs, dependencies, build outputs, temp and output. Prefer current extractions; revisit sources for unresolved details. Use context reuse and compact tool output as described in setup, retaining full evidence and all final checks.
