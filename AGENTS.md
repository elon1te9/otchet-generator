# Workspace rules

- This workspace prepares an учебная практика report from a real software project.
- Never invent features, tests, launches, data, or screenshots.
- Inspect the project and attempt build/run before writing; distinguish source evidence from observed execution.
- `references/official/` has priority over every other source.
- `references/current-example/` defines the current title-page layout and report structure.
- `references/previous-reports/` defines writing style only; do not reuse old tasks or claims.
- Start report work with skills/generate-practice-report/SKILL.md and docs/setup.md. Ask for missing or ambiguous inputs before drafting; save confirmed data in config/student.yaml.
- Use personal data and supervisors only from the current confirmed config. Names and dates in examples are not confirmation.
- Preserve all accumulated formatting and prose rules in docs/report-spec.md and docs/writing-style.md for every user; examples supply structure and style, not permission to weaken these rules.
- Derive the current structure from the selected example and assignment; save it in temp/report-plan.json. Pass that plan to validators; do not reuse a previous report's section names.
- Only reuse extracted sources or evidence when temp/setup.json fingerprints still match the current inputs. Never use temp/archive-original as input for a new report unless explicitly requested.
- Figures must show functionality that was actually run and observed.
- If screenshots cannot be obtained, follow the blank-figure fallback in docs/report-spec.md and continue the DOCX.
- Do not claim successful execution or testing unless it was performed and evidenced.
- Validate the final DOCX structurally, render every page, and inspect it visually.
- Keep temporary artifacts under `temp/`; place only final deliverables in `output/`.
- Use the local Skills and reusable tools instead of expanding this file with detailed procedures.
- For revisions, preserve the latest user-edited DOCX and protected pages; follow the formatting and prose rules in docs/report-spec.md and docs/writing-style.md, including the user's explicit overrides.
- Read files selectively with rg; exclude .git, .vs, dependencies, build outputs and temp. Use extracted docs first; revisit source references only to verify unresolved details.

