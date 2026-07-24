---
name: experiment-report
description: Write Chinese university lab reports and course-design reports, or fit report content into WPS, Word, or docx templates. Use in Codex, OpenClaw, or other local agent workflows when the user asks for an experiment report, lab report, course design report, experiment summary, screenshot insertion, captions, PDF preview, or template filling from topics, requirements, code, data, tutorial pages, screenshots, or results.
---

## When this skill applies

- The user asks to write or complete an experiment report from zero.
- The user asks for a course design report that follows a fixed school-style structure.
- The user has a blank WPS, Word, or docx template and wants it filled.
- The user provides an experiment title, requirements, screenshots, code, data, outputs, or conclusions and wants a report draft or filled report.
- The user provides a tutorial article or CSDN page as the main procedural reference.

## Core workflow

1. Collect the minimum useful inputs:
   - course name
   - experiment name
   - template path or screenshots if formatting matters
   - experiment requirements or task description
   - actual steps, code, screenshots, outputs, data, or conclusions
   - whether the user wants a factual report or a clearly labeled sample version
   - If the user already provides enough facts to write the report body, do not stop to ask for optional metadata such as name, class, date, or template files.
   - For a DOCX or other template-controlled deliverable, if the user has not said whether a preferred template exists, ask once: `你有老师、学校或自己认可的优秀 DOCX/DOC 模板吗？有的话请上传，我会优先保持原格式；没有的话，我会从十套不含学校标识的内置模板中选择。`
   - Do not ask this template question when the user only wants report-body text. Do not ask it again after the user says no, chooses an internal template, or provides a template.
2. Write the full report content before touching template formatting.
3. If critical facts are missing, do not fabricate exact data, screenshots, or measurements.
4. If a local docx template exists and shell execution is available, run `scripts/analyze-docx-template.ps1 -TemplatePath <template.docx>` first. Use `scripts/extract-docx-template.ps1 -Path <template.docx>` when a readable outline is also useful. If the user has no template, run `python -m universal_report select-template --repo-root <repo> --report-type <type> --course <course>` or choose from the ten entries in `examples/report-templates/catalog.json`.
5. Treat the generated `TemplateStyleContract` and extracted structure as the source of truth. The contract includes effective style inheritance, page setup, title/body/caption roles, table geometry, headers, footers, sections, image placeholders, and structural risks.
6. If a template exists, adapt the finished content to the template order and field names.
7. If the user explicitly wants a filled local docx output and the template matches common report patterns, generate a field map with `scripts/generate-docx-field-map.ps1` and then run `scripts/apply-docx-field-map.ps1`.
   - Use label keys for normal blank-field filling.
   - Use `paragraphs` arrays for section-body content.
   - Use `mode: "after"` when the template keeps a fixed heading paragraph and the actual content should go into the following blank paragraph.
   - Use location keys such as `P2` or `T1R1C2` only when explicit overwrite is needed.
8. If the user also provides screenshots or experiment photos and wants them embedded into the final docx, prefer `scripts/generate-docx-image-map.ps1` on the filled copy and then run `scripts/insert-docx-images.ps1`.
9. Preserve user-template formatting by default. Run `scripts/format-docx-report-style.ps1` only when there is no user template, or when the user explicitly asks for normalization or a repository style profile.
10. Prefer content-first completion over fragile GUI wandering.
11. If screenshots are provided, treat them as factual evidence and layout assets.
12. If a tutorial page is provided, treat it as procedural reference and rewrite it into report-style Chinese instead of copying it.
13. When a local workflow should fetch tutorial references before generation, prefer `scripts/prepare-report-prompt.ps1` with `-ReferenceUrls` or `-ReferenceTextPaths`.
14. When the user provides local file paths in a direct chat workflow, inspect those files first if tool access is available; if a path cannot actually be opened, explicitly say which path was inaccessible instead of pretending it was read.
15. When the user asks for a final local `docx` result and the required paths are already present, prefer finishing the end-to-end local workflow over stopping at a body-only draft.
16. On Windows PowerShell, do not chain shell commands with `&&`; use separate executions or `;` so the command remains valid in legacy PowerShell hosts.
17. When intermediate JSON or text files contain Chinese paths, captions, or section names, prefer writing them through PowerShell with explicit UTF-8 encoding or through the bundled scripts; do not rely on generic editor-style writes that may corrupt non-ASCII content.
18. When direct chat already has a template path, screenshots, identity metadata, and either a finished report body or tutorial references, prefer the one-shot local wrapper `scripts/build-report-from-feishu.ps1` over ad-hoc multi-step shell orchestration.
19. If direct chat includes uploaded image attachments and the user also provides local image paths, use the uploaded images to understand the visible content and use the local image paths as the actual files for deterministic `docx` embedding.
20. If direct chat includes uploaded image attachments but no manual local image paths, check whether the runtime injected attachment note lines such as `[media attached ...]` into the prompt. If those lines contain usable image file paths, extract them and pass them into `-ImagePaths` for the local wrapper instead of stopping at body-only output.

## Template fidelity and visual standards

- A user-supplied DOCX template has the highest formatting priority. Default to `-TemplateStyleMode preserve`; do not replace its page setup, fonts, paragraph spacing, title hierarchy, tables, borders, headers, footers, sections, or caption style with repository defaults.
- A user upload always wins over a selected internal template. Do not blend the uploaded template with an internal design.
- When no user template exists, choose exactly one of the ten catalog entries: `neutral-classic-lab`, `neutral-bordered-lab`, `neutral-engineering-lab`, `neutral-course-design`, `neutral-modern-minimal`, `neutral-compact-header-lab`, `neutral-review-panel-lab`, `neutral-code-notebook-lab`, `neutral-data-analysis-lab`, or `neutral-project-dossier`. Use `-BuiltInTemplateId` for local wrappers when the choice is explicit.
- Prefer `neutral-code-notebook-lab` for programming, algorithm, code, test, and debug experiments; `neutral-data-analysis-lab` for measurement, statistics, raw-data, and error-analysis work; `neutral-compact-header-lab` for short weekly records; and `neutral-review-panel-lab` when teacher comments, score, or signature fields are required.
- Prefer `neutral-engineering-lab` for computer-network, operating-system, database, Web, Android, and software-engineering experiments. Prefer `neutral-project-dossier` for long-form system projects and `neutral-course-design` for ordinary course-design reports. Otherwise default to `neutral-classic-lab`; use the bordered or modern template when the requested visual direction calls for it.
- Internal templates must pass `python -m universal_report audit-template-catalog --repo-root <repo>`. They may not contain real school/college names, logos, watermarks, example student identities, embedded third-party media, or copied sample report content.
- Treat the internal DOCX files as original neutral reconstructions under the repository license. Public templates and GitHub projects are design references only; do not copy their branding, sample text, media, or restricted assets into the internal files.
- Before rendering, create or reuse a `TemplateStyleContract`. After rendering, run `scripts/validate-docx-format.ps1` and report the exact failed contract items instead of saying only that formatting is wrong.
- Use `-TemplateStyleMode normalize -StyleFinalDocx` only for the repository default template, demo output, or an explicit user request to unify styling.
- Default image layout is one image per line with a concrete caption below it. Preserve aspect ratio, keep the image and caption together when practical, and use row/grid layouts only when explicitly requested.
- Honor an exact requested image count. Deduplicate by content hash and record selected, rejected, and duplicate files plus selection reasons in `image-manifest.json`; never duplicate an image merely to reach the count.
- Course-design flowcharts, generated structure tables, large fixed image widths, and other profile enhancements are allowed only for a default/normalized template. Do not inject them automatically into an unrelated user template.
- Fast mode must still run structural, layout, and format validation. Strict mode must export DOCX to PDF, render every page, and run visual checks; if that chain cannot complete, return `needs-fix` rather than success.
- LibreOffice is the preferred PDF converter. WPS/Microsoft Word COM is opt-in only, must have a timeout, and may clean up only Office processes started by the current automation run.
- Keep these rules scoped to structured experiment and course-design reports. Do not claim support for arbitrary office documents whose template structures have not been analyzed.

## Writing rules

- Use clear Chinese suitable for university reports.
- Keep claims consistent with the provided requirements, code, data, screenshots, and outputs.
- Avoid empty filler and generic AI phrasing.
- If the user provides the course name or experiment name, write them explicitly into the final report instead of assuming the surrounding chat context is enough.
- If the user already supplied the experiment topic, environment, steps, results, and required headings, write the report immediately instead of asking for more materials.
- Ask follow-up questions only when missing facts would make the result materially wrong, or when the user explicitly wants template filling but no template is available.
- If the user only wants the report body, missing personal identity fields must not block generation.
- If direct chat is being used with local file paths, avoid optimistic assumptions about file access; either read the files or clearly state that file access was not available.
- When the experiment is software-related, include environment, implementation steps, results, analysis, and conclusion.
- When the template has fixed headings, preserve them exactly.
- If webpage instructions and user screenshots differ, trust the user screenshots and outputs.
- If both uploaded image attachments and local image paths are available, use the attachments as the semantic reference for what each image shows, but use the local paths for the final `docx` image insertion workflow.
- If uploaded image attachments are present without manual local paths, prefer the prompt-injected attachment paths from `[media attached ...]` notes as the `ImagePaths` input for `scripts/build-report-from-feishu.ps1` or `scripts/generate-docx-image-map.ps1`. If the runtime does not expose any real attachment path, say that clearly instead of pretending direct `docx` insertion succeeded.
- When screenshots are provided without explicit grouping, infer `实验环境`, `实验步骤`, `实验结果`, or `问题分析` from filenames and visible content, but do not invent unseen details.
- If a local workflow needs temporary JSON such as field maps or image maps and those files include Chinese text, write them in explicit UTF-8 and retry from that stage if parsing fails.

## Output modes

- Default: final report content with headings ready to paste.
- Template mode: exact field-to-content mapping in template order.
- Template mode can include block values such as `{"section-body": ["paragraph one", "paragraph two"]}` when the template body has multi-paragraph sections.
- Image mode: when screenshots should be embedded into a docx, prefer image specs or an image insertion map that can be passed to `scripts/generate-docx-image-map.ps1` or `scripts/insert-docx-images.ps1`. Stable section anchors such as `实验步骤` or `实验结果` are preferred over fragile paragraph numbers when the filled docx may add or move paragraphs.
- Completion mode: if the user explicitly asks to complete the template, first finish the content, then attempt template filling only if tooling is actually available.

## Optional helpers

- For local docx templates, run `scripts/extract-docx-template.ps1` to capture the actual field order before producing a field mapping.
- For local docx templates that should be machine-filled, run `scripts/generate-docx-field-map.ps1` after the report body is ready, then run `scripts/apply-docx-field-map.ps1`.
- For local screenshots or experiment photos that should be embedded into the filled docx, prefer `scripts/generate-docx-image-map.ps1` first and then run `scripts/insert-docx-images.ps1` on the already-filled copy.
- For a cleaner normalized copy, optionally run `scripts/format-docx-report-style.ps1` only when the user explicitly wants repository styling or no user template exists.
- Use `scripts/plan-report-content.ps1` to create a structured content plan before body rendering when materials come from multiple sources.
- Use `scripts/analyze-docx-template.ps1` and `scripts/validate-docx-format.ps1` for template-contract analysis and post-render fidelity checks.
- Use `scripts/run-visual-validation.ps1` for strict PDF and per-page preview validation.
- Use `python -m universal_report list-templates --repo-root <repo>` to show the five internal options, `select-template` to select one, and `audit-template-catalog` to enforce neutrality and provenance.
- For chat-driven local execution, prefer `scripts/build-report-from-feishu.ps1` so the wrapper can keep the final deliverable in the output root and move intermediate files into an `artifacts/` subdirectory.
- The image pipeline can resolve staged relative attachment paths such as `media/inbound/example.png` from OpenClaw-style prompts; when those paths appear in prompt-injected media notes, reuse them directly in `-ImagePaths`.
- When the template has fixed section headings plus blank paragraphs, prefer block mappings over flattening long body content into a single field.
- For public tutorial pages, prefer `scripts/fetch-web-article.ps1`; keep `scripts/fetch-csdn-article.ps1` as the compatibility wrapper for CSDN-specific workflows.
- When a tutorial page should flow directly into report generation, prefer `scripts/prepare-report-prompt.ps1` so the extracted reference text is appended to the final request deterministically.
- The helpers are optional. If they are unavailable, still finish the report from the information already provided.

## Read references as needed

- Read `references/common-structures.md` when choosing a report outline.
- Read `references/template-fit.md` when the user provides a WPS, Word, or docx template path or template screenshots.
- Read `references/image-handling.md` when the user provides experiment screenshots or process images.
